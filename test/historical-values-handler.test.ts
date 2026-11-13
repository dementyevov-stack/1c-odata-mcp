import { describe, expect, it, vi, afterEach } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

const ORG = "11111111-1111-4111-8111-111111111111";
const CP = "22222222-2222-4222-8222-222222222222";
const CTR = "33333333-3333-4333-8333-333333333333";
const WH = "44444444-4444-4444-8444-444444444444";
const NOM = "55555555-5555-4555-8555-555555555555";
const ACC = "11111111-1111-4111-8111-111111111111";

const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(cleanup.splice(0).map((fn) => fn().catch(() => undefined)));
});

async function setup(
  options: {
    documents?: Record<string, unknown>[];
    getError?: boolean;
    mutate?: (actual: Record<string, unknown>) => Record<string, unknown>;
  } = {},
) {
  const fake = fake1C(
    { READ_ONLY: "false", ODATA_WRITABLE: "true" },
    {
      data: (q: FakeQuery) =>
        q.base === "Document_ПоступлениеТоваровУслуг" ? (options.documents ?? []) : undefined,
    },
  );
  const original = fake.conn.client.getEntity.bind(fake.conn.client);
  vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
    if (path.startsWith("Catalog_ДоговорыКонтрагентов("))
      return {
        Ref_Key: CTR,
        Owner_Key: CP,
        Организация_Key: ORG,
        ВидДоговора: "СПоставщиком",
        DeletionMark: false,
      };
    if (path.startsWith("Document_ПоступлениеТоваровУслуг(") && !path.includes("$select")) {
      if (options.getError) throw new Error("GET failed");
      const actual = await original(path);
      const withMark = { DeletionMark: false, ...actual };
      return options.mutate ? options.mutate(withMark) : withMark;
    }
    return original(path);
  });
  const server = createServer(fake.ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "historical-values-handler-test", version: "1" });
  await client.connect(b);
  cleanup.push(async () => {
    await client.close().catch(() => undefined);
    await server.close().catch(() => undefined);
  });
  return { fake, client };
}

const call = async (client: Client, args: Record<string, unknown>) =>
  (await client.callTool({ name: "write.purchase.create_purchase", arguments: args })) as {
    isError?: boolean;
    structuredContent?: Record<string, unknown>;
  };
const postCount = (fake: ReturnType<typeof fake1C>): number =>
  fake.requests.filter((r) => r.method === "POST").length;

const args = {
  organization: ORG,
  counterpartyRef: CP,
  contractRef: CTR,
  warehouse: WH,
  date: "2026-11-01",
  number: "P-VAL-1",
  time: "12:34:56",
  incomingNumber: "IN-VAL-1",
  incomingDate: "2026-10-01",
  incomingTime: "09:08:07",
  documentTotal: 240,
  settlementAccountRef: ACC,
  advanceAccountRef: ACC,
  universalDocument: false,
  sumIncludesVat: false,
  operationKind: "ПокупкаКомиссия",
  lines: [
    {
      nomenclatureRef: NOM,
      quantity: 1,
      price: 100,
      amount: 100,
      vatRate: "НДС20",
      vatAmount: 0,
      inventoryAccountRef: ACC,
      vatAccountRef: ACC,
      warehouseEmpty: true,
    },
  ],
  serviceLines: [
    {
      nomenclatureRef: NOM,
      emptyQuantity: true,
      price: 100,
      amount: 100,
      vatRate: "НДС20",
      vatAmount: 20,
      costAccountRef: ACC,
      costAccountNuRef: ACC,
      vatAccountRef: ACC,
    },
  ],
};

describe("историческое поступление: точное представление и единая сверка", () => {
  it("различает omitted/empty/0/false и сохраняет время", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      Date: "2026-11-01T12:34:56",
      ДатаВходящегоДокумента: "2026-10-01T09:08:07",
      Posted: false,
      ЭтоУниверсальныйДокумент: false,
      СуммаВключаетНДС: false,
    });
    const goods = (payload["Товары"] as Array<Record<string, unknown>>)[0];
    expect(goods["СуммаНДС"]).toBe(0);
    expect(goods["Склад_Key"]).toBe("00000000-0000-0000-0000-000000000000");
    expect(goods).not.toHaveProperty("Содержание");
    const service = (payload["Услуги"] as Array<Record<string, unknown>>)[0];
    expect(service["Количество"]).toBeNull();
    expect(service).not.toHaveProperty("Содержание");
    expect(postCount(fake)).toBe(0);
  });

  it("успешно подтверждает обе таблицы", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    const created = await call(client, {
      ...args,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({ created: true, verification: "ok" });
    expect(postCount(fake)).toBe(1);
  });

  it("сервер подставил значение вместо явной пустоты — failed", async () => {
    const { client, fake } = await setup({
      mutate: (actual) => {
        const rows = actual["Товары"] as Array<Record<string, unknown>>;
        rows[0]["Склад_Key"] = WH;
        return actual;
      },
    });
    const preview = await call(client, { ...args, confirm: false });
    const created = await call(client, {
      ...args,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError).toBe(true);
    expect(JSON.stringify(created)).toContain("Склад_Key");
    expect(postCount(fake)).toBe(1);
  });

  it("пропало обязательное поле или строка — failed", async () => {
    for (const mutate of [
      (actual: Record<string, unknown>) => {
        delete (actual["Товары"] as Array<Record<string, unknown>>)[0]["Сумма"];
        return actual;
      },
      (actual: Record<string, unknown>) => {
        actual["Услуги"] = [];
        return actual;
      },
    ]) {
      const { client } = await setup({ mutate });
      const preview = await call(client, { ...args, confirm: false });
      const created = await call(client, {
        ...args,
        confirm: true,
        operationId: preview.structuredContent?.operationId,
      });
      expect(created.isError).toBe(true);
    }
  });

  it("GET упал после POST — failed и без второго POST", async () => {
    const { client, fake } = await setup({ getError: true });
    const preview = await call(client, { ...args, confirm: false });
    const created = await call(client, {
      ...args,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError).toBe(true);
    expect(created.structuredContent?.verification).toBe("failed");
    expect(postCount(fake)).toBe(1);
  });

  it("повтор с тем же operationId не создаёт второй документ", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    const id = preview.structuredContent?.operationId;
    const first = await call(client, { ...args, confirm: true, operationId: id });
    const second = await call(client, { ...args, confirm: true, operationId: id });
    expect(first.isError, JSON.stringify(first)).toBeFalsy();
    expect(second.isError, JSON.stringify(second)).toBeFalsy();
    expect(postCount(fake)).toBe(1);
  });
});
