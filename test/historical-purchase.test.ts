import { describe, expect, it, vi, afterEach } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

/** Синтетическое историческое поступление товаров: черновик, точные реквизиты и readback. */
const ORG = "11111111-1111-4111-8111-111111111111";
const CP = "22222222-2222-4222-8222-222222222222";
const CTR = "33333333-3333-4333-8333-333333333333";
const WH = "44444444-4444-4444-8444-444444444444";
const NOM = "55555555-5555-4555-8555-555555555555";
const ACC = "11111111-1111-4111-8111-111111111111";
const DOC = "99999999-9999-4999-8999-999999999999";

const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(cleanup.splice(0).map((fn) => fn().catch(() => undefined)));
});

async function setup(
  options: { documents?: Record<string, unknown>[]; rowWarehouse?: string; contractKind?: string } = {},
) {
  const data = (q: FakeQuery): Record<string, unknown>[] | undefined =>
    q.base === "Document_ПоступлениеТоваровУслуг" ? (options.documents ?? []) : undefined;
  const fake = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" }, { data });
  const original = fake.conn.client.getEntity.bind(fake.conn.client);
  vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
    if (path.startsWith("Catalog_ДоговорыКонтрагентов("))
      return {
        Ref_Key: CTR,
        Owner_Key: CP,
        Организация_Key: ORG,
        ВидДоговора: options.contractKind ?? "СПоставщиком",
        DeletionMark: false,
      };
    if (path.startsWith("Document_ПоступлениеТоваровУслуг(") && !path.includes("$select")) {
      const actual = await original(path);
      const rows = (actual["Товары"] as Array<Record<string, unknown>>).map((row) => ({
        ...row,
        ...(options.rowWarehouse ? { Склад_Key: options.rowWarehouse } : {}),
      }));
      return { DeletionMark: false, ...actual, Товары: rows };
    }
    return original(path);
  });
  const server = createServer(fake.ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "historical-purchase-test", version: "1" });
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
  number: "P-1",
  time: "12:34:56",
  incomingNumber: "IN-1",
  incomingDate: "2026-10-01",
  incomingTime: "09:08:07",
  documentTotal: 120,
  settlementAccountRef: ACC,
  advanceAccountRef: ACC,
  universalDocument: true,
  sumIncludesVat: false,
  lines: [
    {
      nomenclatureRef: NOM,
      quantity: 1,
      price: 100,
      amount: 100,
      vatRate: "НДС20",
      vatAmount: 20,
      sourceLineId: "SRC-P-1",
      inventoryAccountRef: ACC,
      vatAccountRef: ACC,
      warehouseEmpty: true,
    },
  ],
};

describe("write.purchase.create_purchase: исторический opt-in", () => {
  it("сериализует точные поля, net/VAT/gross и пустой склад строки", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      Number: "P-1",
      Date: "2026-11-01T12:34:56",
      ВидОперации: "Товары",
      Posted: false,
      Склад_Key: WH,
      НомерВходящегоДокумента: "IN-1",
      ДатаВходящегоДокумента: "2026-10-01T09:08:07",
      СуммаВключаетНДС: false,
      СуммаДокумента: 120,
      ЭтоУниверсальныйДокумент: true,
      СчетУчетаРасчетовСКонтрагентом_Key: ACC,
      СчетУчетаРасчетовПоАвансам_Key: ACC,
    });
    const row = (payload["Товары"] as Array<Record<string, unknown>>)[0]!;
    expect(row).toMatchObject({
      Количество: 1,
      Сумма: 100,
      СуммаНДС: 20,
      СтавкаНДС: "НДС20",
      ИдентификаторСтроки: "SRC-P-1",
      СчетУчета_Key: ACC,
      СчетУчетаНДС_Key: ACC,
    });
    expect(row).not.toHaveProperty("Склад_Key");
    expect(postCount(fake)).toBe(0);
  });

  it("подтверждает и полностью проверяет readback", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    expect(typeof operationId).toBe("string");
    const created = await call(client, { ...args, confirm: true, operationId });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({ created: true, verification: "ok" });
    expect(created.structuredContent?.verifiedFields).toEqual(
      expect.arrayContaining([
        "Number",
        "Date",
        "ВидОперации",
        "Склад_Key",
        "НомерВходящегоДокумента",
        "ДатаВходящегоДокумента",
        "СуммаДокумента",
        "Товары[1].Сумма",
        "Товары[1].СуммаНДС",
        "Товары[1].ИдентификаторСтроки",
      ]),
    );
    expect(postCount(fake)).toBe(1);
  });

  it("склад строки, назначенный после записи, ломает verification", async () => {
    const { client } = await setup({ rowWarehouse: WH });
    const preview = await call(client, { ...args, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    const created = await call(client, { ...args, confirm: true, operationId });
    expect(created.isError).toBe(true);
    expect(JSON.stringify(created)).toContain("Склад_Key");
    expect(JSON.stringify(created)).toContain("Ref_Key");
  });

  it("договор СПоставщиком принимается, СПокупателем — отклоняется", async () => {
    const supplier = await setup();
    const ok = await call(supplier.client, { ...args, confirm: false });
    expect(ok.isError, JSON.stringify(ok)).toBeFalsy();

    const customer = await setup({ contractKind: "СПокупателем" });
    const bad = await call(customer.client, { ...args, confirm: false });
    expect(bad.isError).toBe(true);
    expect(JSON.stringify(bad)).toContain("СПокупателем");
    expect(postCount(customer.fake)).toBe(0);
  });

  it("time без date не подставляет текущую дату", async () => {
    const { client, fake } = await setup();
    const result = await call(client, { ...args, date: undefined, confirm: false });
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result)).toContain("date");
    expect(postCount(fake)).toBe(0);
  });

  it("дубль номера в том же году блокирует POST", async () => {
    const { client, fake } = await setup({
      documents: [{ Ref_Key: DOC, Number: "P-1", Date: "2026-04-01T00:00:00", Организация_Key: ORG }],
    });
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError).toBe(true);
    expect(JSON.stringify(preview)).toContain("уже занят");
    expect(postCount(fake)).toBe(0);
  });

  it("обычный вызов без исторических полей работает как раньше", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, {
      organization: ORG,
      counterpartyRef: CP,
      contractRef: CTR,
      warehouse: WH,
      date: "2026-11-01",
      sumIncludesVat: true,
      lines: [{ nomenclatureRef: NOM, quantity: 1, price: 100 }],
      confirm: false,
    });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({ СуммаДокумента: 100, СуммаВключаетНДС: true });
    expect(payload).not.toHaveProperty("Number");
    expect(postCount(fake)).toBe(0);
  });
});
