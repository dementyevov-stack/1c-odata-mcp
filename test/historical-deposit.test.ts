import { describe, expect, it, vi, afterEach } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

const ORG = "11111111-1111-4111-8111-111111111111";
const ACC = "11111111-1111-4111-8111-111111111111";
const DDS = "22222222-2222-4222-8222-222222222222";
const EXISTING = "33333333-3333-4333-8333-333333333333";

const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(cleanup.splice(0).map((fn) => fn().catch(() => undefined)));
});

async function setup(options: { docs?: Record<string, unknown>[]; readback?: Record<string, unknown> } = {}) {
  const fake = fake1C(
    { READ_ONLY: "false", ODATA_WRITABLE: "true" },
    {
      data: (q: FakeQuery) =>
        q.base === "Document_ПоступлениеНаРасчетныйСчет" ? (options.docs ?? []) : undefined,
    },
  );
  const original = fake.conn.client.getEntity.bind(fake.conn.client);
  vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
    if (path.startsWith("Document_ПоступлениеНаРасчетныйСчет(") && !path.includes("$select")) {
      if (options.readback) return { DeletionMark: false, ...options.readback };
      const actual = await original(path);
      return { DeletionMark: false, ...actual };
    }
    return original(path);
  });
  const server = createServer(fake.ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "historical-deposit-test", version: "1" });
  await client.connect(b);
  cleanup.push(async () => {
    await client.close().catch(() => undefined);
    await server.close().catch(() => undefined);
  });
  return { fake, client };
}

const call = async (client: Client, args: Record<string, unknown>) =>
  (await client.callTool({ name: "write.money.create_payment", arguments: args })) as {
    isError?: boolean;
    structuredContent?: Record<string, unknown>;
  };
const postCount = (fake: ReturnType<typeof fake1C>): number =>
  fake.requests.filter((r) => r.method === "POST").length;

const args = {
  organization: ORG,
  amount: 100,
  orgBankAccount: ACC,
  number: "DEP-1",
  date: "2026-11-01",
  time: "12:34:56",
  incomingNumber: "IN-DEP-1",
  incomingDate: "2026-10-31",
  incomingTime: "09:08:07",
  bankAccountingAccountRef: ACC,
  settlementAccountRef: ACC,
  cashflowItemRef: DDS,
  rowCashflowItemRef: DDS,
  vatRate: "НДС20",
  usnReflection: "Не принимаются",
  purposeText: "Депозит",
  comment: "Исторический депозит",
};

const readback = {
  Number: "DEP-1",
  Date: "2026-11-01T12:34:56",
  Posted: false,
  ВидОперации: "Депозит",
  Организация_Key: ORG,
  СчетОрганизации_Key: ACC,
  СчетБанк_Key: ACC,
  СчетУчетаРасчетовСКонтрагентом_Key: ACC,
  НомерВходящегоДокумента: "IN-DEP-1",
  ДатаВходящегоДокумента: "2026-10-31T09:08:07",
  СуммаДокумента: 100,
  СтатьяДвиженияДенежныхСредств_Key: DDS,
  НазначениеПлатежа: "Депозит",
  Комментарий: "Исторический депозит",
  ОтражениеВУСН: "Не принимаются",
  РасшифровкаПлатежа: [
    {
      LineNumber: 1,
      СуммаПлатежа: 100,
      СтатьяДвиженияДенежныхСредств_Key: DDS,
      СтавкаНДС: "НДС20",
      СуммаНДС: 0,
    },
  ],
};

describe("write.money.create_payment: исторический Депозит", () => {
  it("точные реквизиты, одна строка и никаких плательщиков/аналитик", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      Number: "DEP-1",
      Date: "2026-11-01T12:34:56",
      ВидОперации: "Депозит",
      СчетБанк_Key: ACC,
      СчетУчетаРасчетовСКонтрагентом_Key: ACC,
      СтатьяДвиженияДенежныхСредств_Key: DDS,
      ОтражениеВУСН: "Не принимаются",
      СуммаДокумента: 100,
    });
    expect(payload).not.toHaveProperty("Контрагент");
    expect(payload).not.toHaveProperty("ДоговорКонтрагента_Key");
    const rows = payload["РасшифровкаПлатежа"] as Array<Record<string, unknown>>;
    expect(rows).toHaveLength(1);
    expect(rows[0]).toMatchObject({
      LineNumber: 1,
      СуммаПлатежа: 100,
      СтатьяДвиженияДенежныхСредств_Key: DDS,
      СтавкаНДС: "НДС20",
    });
    expect(postCount(fake)).toBe(0);
  });

  it("пустая ставка НДС передаётся явно и проверяется", async () => {
    const { client } = await setup({
      readback: {
        ...readback,
        РасшифровкаПлатежа: [
          {
            LineNumber: 1,
            СуммаПлатежа: 100,
            СтатьяДвиженияДенежныхСредств_Key: DDS,
            СтавкаНДС: "",
            СуммаНДС: 0,
          },
        ],
      },
    });
    const emptyArgs = { ...args, vatRate: "" };
    const preview = await call(client, { ...emptyArgs, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const rows = (preview.structuredContent?.payload as Record<string, unknown>)[
      "РасшифровкаПлатежа"
    ] as Array<Record<string, unknown>>;
    expect(rows[0]).toHaveProperty("СтавкаНДС", "");
    const created = await call(client, {
      ...emptyArgs,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
  });

  it("подставленная НДС20/БезНДС или непустые ссылки ломают readback", async () => {
    for (const bad of [
      { ...readback, РасшифровкаПлатежа: [{ ...readback.РасшифровкаПлатежа[0], СтавкаНДС: "БезНДС" }] },
      {
        ...readback,
        РасшифровкаПлатежа: [{ ...readback.РасшифровкаПлатежа[0], ДоговорКонтрагента_Key: ACC }],
      },
      { ...readback, ДоговорКонтрагента_Key: ACC },
    ]) {
      const { client } = await setup({ readback: bad });
      const preview = await call(client, { ...args, vatRate: "", confirm: false });
      const created = await call(client, {
        ...args,
        vatRate: "",
        confirm: true,
        operationId: preview.structuredContent?.operationId,
      });
      expect(created.isError).toBe(true);
    }
  });

  it("подтверждает и полностью проверяет readback", async () => {
    const { client, fake } = await setup({ readback });
    const preview = await call(client, { ...args, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    const created = await call(client, { ...args, confirm: true, operationId });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({ created: true, verification: "ok" });
    expect(created.structuredContent?.verifiedFields).toEqual(
      expect.arrayContaining(["Number", "СчетБанк_Key", "РасшифровкаПлатежа", "ОтражениеВУСН"]),
    );
    expect(postCount(fake)).toBe(1);
  });

  it("дубль по номеру/году/организации отклоняется до POST", async () => {
    const { client, fake } = await setup({
      docs: [{ Ref_Key: EXISTING, Number: "DEP-1", Date: "2026-01-01T00:00:00", Организация_Key: ORG }],
    });
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError).toBe(true);
    expect(JSON.stringify(preview)).toContain("уже занят");
    expect(postCount(fake)).toBe(0);
  });

  it("обычный режим требует контрагента и договор", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, {
      organization: ORG,
      amount: 100,
      orgBankAccount: ACC,
      operationKind: "ОплатаПокупателя",
      confirm: false,
    });
    expect(preview.isError).toBe(true);
    expect(JSON.stringify(preview)).toContain("counterpartyRef");
    expect(postCount(fake)).toBe(0);
  });
});
