import { describe, expect, it, vi, afterEach } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

const ORG = "11111111-1111-4111-8111-111111111111";
const CP = "22222222-2222-4222-8222-222222222222";
const CTR = "33333333-3333-4333-8333-333333333333";
const WH = "44444444-4444-4444-8444-444444444444";
const NOM1 = "55555555-5555-4555-8555-555555555551";
const NOM2 = "55555555-5555-4555-8555-555555555552";
const UNIT = "66666666-6666-4666-8666-666666666666";
const COUNTRY = "77777777-7777-4777-8777-777777777777";
const GROUP = "88888888-8888-4888-8888-888888888888";
const ACC = "11111111-1111-4111-8111-111111111111";
const EXISTING = "99999999-9999-4999-8999-999999999999";
const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";

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
        q.base === "Document_РеализацияТоваровУслуг" ? (options.docs ?? []) : undefined,
    },
  );
  const original = fake.conn.client.getEntity.bind(fake.conn.client);
  vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
    if (path.startsWith("Catalog_ДоговорыКонтрагентов("))
      return {
        Ref_Key: CTR,
        Owner_Key: CP,
        Организация_Key: ORG,
        ВидДоговора: "СПокупателем",
        DeletionMark: false,
      };
    if (path.startsWith("Document_РеализацияТоваровУслуг(") && !path.includes("$select")) {
      if (options.readback) return { DeletionMark: false, ...options.readback };
      const actual = await original(path);
      return { DeletionMark: false, ...actual };
    }
    return original(path);
  });
  const server = createServer(fake.ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "historical-sale-shipment-test", version: "1" });
  await client.connect(b);
  cleanup.push(async () => {
    await client.close().catch(() => undefined);
    await server.close().catch(() => undefined);
  });
  return { fake, client };
}

const call = async (client: Client, args: Record<string, unknown>) =>
  (await client.callTool({ name: "write.sales.create_shipment", arguments: args })) as {
    isError?: boolean;
    structuredContent?: Record<string, unknown>;
  };
const postCount = (fake: ReturnType<typeof fake1C>): number =>
  fake.requests.filter((r) => r.method === "POST").length;

const line = (nomenclatureRef: string, amount: number, vatAmount: number) => ({
  nomenclatureRef,
  quantity: 1.5,
  price: 100,
  amount,
  vatRate: "НДС20",
  vatAmount,
  unitRef: UNIT,
  coefficient: 1,
  countryOfOriginRef: COUNTRY,
  accountingGroup: [{ ref: GROUP, type: "StandardODATA.Catalog_НоменклатурныеГруппы" }],
  inventoryAccountRef: ACC,
  incomeAccountRef: ACC,
  expenseAccountRef: ACC,
  vatAccountRef: ACC,
});

const args = {
  organization: ORG,
  counterpartyRef: CP,
  contractRef: CTR,
  warehouse: WH,
  date: "2026-11-01",
  time: "12:34:56",
  number: "SHIP-1",
  operationKind: "Товары",
  sumIncludesVat: false,
  documentTotal: 360,
  settlementAccountRef: ACC,
  advanceAccountRef: ACC,
  universalDocument: true,
  electronicDocumentType: "УПД",
  lines: [
    { ...line(NOM1, 100, 20), sourceLineId: "G-1" },
    { ...line(NOM2, 100, 20), sourceLineId: "G-2" },
  ],
  serviceLines: [
    {
      nomenclatureRef: NOM2,
      emptyQuantity: true,
      price: 100,
      amount: 100,
      vatRate: "НДС20",
      vatAmount: 20,
      content: "Услуга по договору",
      sourceLineId: "S-1",
      accountingGroup: [{ ref: GROUP, type: "StandardODATA.Catalog_НоменклатурныеГруппы" }],
      incomeAccountRef: ACC,
      expenseAccountRef: ACC,
      vatAccountRef: ACC,
    },
  ],
};

const readback = {
  Number: "SHIP-1",
  Date: "2026-11-01T12:34:56",
  Posted: false,
  ВидОперации: "Товары",
  Организация_Key: ORG,
  Контрагент_Key: CP,
  ДоговорКонтрагента_Key: CTR,
  Склад_Key: WH,
  БанковскийСчетОрганизации_Key: EMPTY_GUID,
  СчетУчетаРасчетовСКонтрагентом_Key: ACC,
  СчетУчетаРасчетовПоАвансам_Key: ACC,
  СуммаВключаетНДС: false,
  СуммаДокумента: 360,
  ЭтоУниверсальныйДокумент: true,
  ВидЭлектронногоДокумента: "УПД",
  Товары: [
    {
      LineNumber: 1,
      Номенклатура_Key: NOM1,
      Количество: 1.5,
      Цена: 100,
      Сумма: 100,
      СтавкаНДС: "НДС20",
      СуммаНДС: 20,
      ЕдиницаИзмерения_Key: UNIT,
      Коэффициент: 1,
      СтранаПроисхождения_Key: COUNTRY,
      Субконто: GROUP,
      Субконто_Type: "StandardODATA.Catalog_НоменклатурныеГруппы",
      ИдентификаторСтроки: "G-1",
      СчетУчета_Key: ACC,
      СчетДоходов_Key: ACC,
      СчетРасходов_Key: ACC,
      СчетУчетаНДСПоРеализации_Key: ACC,
    },
    {
      LineNumber: 2,
      Номенклатура_Key: NOM2,
      Количество: 1.5,
      Цена: 100,
      Сумма: 100,
      СтавкаНДС: "НДС20",
      СуммаНДС: 20,
      ЕдиницаИзмерения_Key: UNIT,
      Коэффициент: 1,
      СтранаПроисхождения_Key: COUNTRY,
      Субконто: GROUP,
      Субконто_Type: "StandardODATA.Catalog_НоменклатурныеГруппы",
      ИдентификаторСтроки: "G-2",
      СчетУчета_Key: ACC,
      СчетДоходов_Key: ACC,
      СчетРасходов_Key: ACC,
      СчетУчетаНДСПоРеализации_Key: ACC,
    },
  ],
  Услуги: [
    {
      LineNumber: 1,
      Номенклатура_Key: NOM2,
      Содержание: "Услуга по договору",
      Цена: 100,
      Сумма: 100,
      СтавкаНДС: "НДС20",
      СуммаНДС: 20,
      Субконто: GROUP,
      Субконто_Type: "StandardODATA.Catalog_НоменклатурныеГруппы",
      ИдентификаторСтроки: "S-1",
      СчетДоходов_Key: ACC,
      СчетРасходов_Key: ACC,
      СчетУчетаНДСПоРеализации_Key: ACC,
    },
  ],
};

describe("write.sales.create_shipment: исторические Товары+Услуги", () => {
  it("сериализует две таблицы, точный итог net+VAT и пустой банк", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      Number: "SHIP-1",
      Date: "2026-11-01T12:34:56",
      ВидОперации: "Товары",
      СуммаВключаетНДС: false,
      СуммаДокумента: 360,
      БанковскийСчетОрганизации_Key: EMPTY_GUID,
      ЭтоУниверсальныйДокумент: true,
      ВидЭлектронногоДокумента: "УПД",
    });
    expect((payload["Товары"] as unknown[]).length).toBe(2);
    expect((payload["Услуги"] as Array<Record<string, unknown>>)[0]).not.toHaveProperty("Количество");
    expect(postCount(fake)).toBe(0);
  });

  it("подтверждает и проверяет обе таблицы", async () => {
    const { client, fake } = await setup({ readback });
    const preview = await call(client, { ...args, confirm: false });
    const created = await call(client, {
      ...args,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({ created: true, verification: "ok" });
    expect(created.structuredContent?.verifiedFields).toEqual(expect.arrayContaining(["Товары", "Услуги"]));
    expect(postCount(fake)).toBe(1);
  });

  it("дубль по номеру/году/организации отклоняется до POST", async () => {
    const { client, fake } = await setup({
      docs: [{ Ref_Key: EXISTING, Number: "SHIP-1", Date: "2026-01-01T00:00:00", Организация_Key: ORG }],
    });
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError).toBe(true);
    expect(JSON.stringify(preview)).toContain("уже занят");
    expect(postCount(fake)).toBe(0);
  });

  it("обычный вызов без исторических полей сохраняет прежний путь", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, {
      organization: ORG,
      counterpartyRef: CP,
      contractRef: CTR,
      warehouse: WH,
      date: "2026-11-01",
      sumIncludesVat: true,
      lines: [{ nomenclatureRef: NOM1, quantity: 1, price: 100 }],
      confirm: false,
    });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    expect((preview.structuredContent?.payload as Record<string, unknown>)["Number"]).toBeUndefined();
    expect(postCount(fake)).toBe(0);
  });
});
