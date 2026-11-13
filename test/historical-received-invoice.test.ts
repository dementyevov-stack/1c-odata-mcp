import { describe, expect, it, vi, afterEach } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

const ORG = "11111111-1111-4111-8111-111111111111";
const CP = "22222222-2222-4222-8222-222222222222";
const CTR = "33333333-3333-4333-8333-333333333333";
const BASE = "44444444-4444-4444-8444-444444444444";
const EXISTING = "55555555-5555-4555-8555-555555555555";

const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(cleanup.splice(0).map((fn) => fn().catch(() => undefined)));
});

async function setup(options: { invoices?: Record<string, unknown>[] } = {}) {
  const fake = fake1C(
    { READ_ONLY: "false", ODATA_WRITABLE: "true" },
    {
      data: (q: FakeQuery) =>
        q.base === "Document_СчетФактураПолученный" ? (options.invoices ?? []) : undefined,
    },
  );
  const original = fake.conn.client.getEntity.bind(fake.conn.client);
  vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
    if (path.startsWith("Document_ПоступлениеТоваровУслуг(")) {
      return {
        Ref_Key: BASE,
        Организация_Key: ORG,
        Контрагент_Key: CP,
        ДоговорКонтрагента_Key: CTR,
        СуммаДокумента: 120,
        Товары: [{ СуммаНДС: 20 }],
        Услуги: [],
      };
    }
    if (path.startsWith("Document_СчетФактураПолученный(") && !path.includes("$select")) {
      const actual = await original(path);
      return { DeletionMark: false, ...actual };
    }
    return original(path);
  });
  const server = createServer(fake.ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "historical-received-invoice-test", version: "1" });
  await client.connect(b);
  cleanup.push(async () => {
    await client.close().catch(() => undefined);
    await server.close().catch(() => undefined);
  });
  return { fake, client };
}

const call = async (client: Client, args: Record<string, unknown>) =>
  (await client.callTool({ name: "write.purchase.create_received_invoice", arguments: args })) as {
    isError?: boolean;
    structuredContent?: Record<string, unknown>;
  };
const postCount = (fake: ReturnType<typeof fake1C>): number =>
  fake.requests.filter((r) => r.method === "POST").length;

const args = {
  baseDocumentRef: BASE,
  number: "SF-1",
  date: "2026-11-01",
  time: "11:22:33",
  incomingNumber: "IN-SF-1",
  incomingDate: "2026-10-31",
  incomingTime: "10:11:12",
  receiptCode: 1,
  numberRepresentation: "№ IN-SF-1",
  operationCode: "01",
  vatDeductible: true,
};

describe("write.purchase.create_received_invoice: исторический opt-in", () => {
  it("сериализует отдельные номера, точные даты, код получения и наследует основание", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      Number: "SF-1",
      Date: "2026-11-01T11:22:33",
      ВидСчетаФактуры: "НаПоступление",
      НомерВходящегоДокумента: "IN-SF-1",
      ДатаВходящегоДокумента: "2026-10-31T10:11:12",
      КодСпособаПолучения: 1,
      ПредставлениеНомера: "№ IN-SF-1",
      КодВидаОперации: "01",
      НДСПредъявленКВычету: true,
      СуммаДокумента: 120,
      СуммаНДСДокумента: 20,
      ДокументОснование: BASE,
      ДокументОснование_Type: "StandardODATA.Document_ПоступлениеТоваровУслуг",
    });
    expect((payload["ДокументыОснования"] as Array<Record<string, unknown>>)[0]).toMatchObject({
      ДокументОснование: BASE,
      ДокументОснование_Type: "StandardODATA.Document_ПоступлениеТоваровУслуг",
    });
    expect(postCount(fake)).toBe(0);
  });

  it("подтверждает и проверяет все поля и тип основания", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    const created = await call(client, { ...args, confirm: true, operationId });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({ created: true, verification: "ok" });
    expect(created.structuredContent?.verifiedFields).toEqual(
      expect.arrayContaining([
        "Number",
        "ДатаВходящегоДокумента",
        "КодСпособаПолучения",
        "ДокументыОснования[1]",
      ]),
    );
    expect(postCount(fake)).toBe(1);
  });

  it("дубль по основанию или номеру отклоняется до POST", async () => {
    const { client, fake } = await setup({
      invoices: [
        {
          Ref_Key: EXISTING,
          Number: "OTHER",
          Date: "2026-01-01T00:00:00",
          Организация_Key: ORG,
          ДокументОснование: BASE,
        },
      ],
    });
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError).toBe(true);
    expect(JSON.stringify(preview)).toContain("уже существует");
    expect(postCount(fake)).toBe(0);
  });

  it("обычный вызов сохраняет прежний путь", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { baseDocumentRef: BASE, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).not.toHaveProperty("Number");
    expect(payload).toMatchObject({ ВидСчетаФактуры: "НаПоступление", СуммаДокумента: 120 });
    expect(postCount(fake)).toBe(0);
  });
});
