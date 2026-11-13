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
const UNIT = "66666666-6666-4666-8666-666666666666";
const COUNTRY = "77777777-7777-4777-8777-777777777777";
const DOC = "99999999-9999-4999-8999-999999999999";

const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(cleanup.splice(0).map((fn) => fn().catch(() => undefined)));
});

async function setup(
  options: {
    documents?: Record<string, unknown>[];
    rowWarehouse?: string;
    rowUnit?: string;
    rowCoefficient?: number;
    serviceQuantity?: number;
    contractKind?: string;
  } = {},
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
        ...(options.rowUnit !== undefined ? { ЕдиницаИзмерения_Key: options.rowUnit } : {}),
        ...(options.rowCoefficient !== undefined ? { Коэффициент: options.rowCoefficient } : {}),
      }));
      const services = Array.isArray(actual["Услуги"])
        ? (actual["Услуги"] as Array<Record<string, unknown>>).map((row) => ({
            ...row,
            ...(options.serviceQuantity !== undefined ? { Количество: options.serviceQuantity } : {}),
          }))
        : actual["Услуги"];
      return {
        DeletionMark: false,
        ...actual,
        Товары: rows,
        ...(services !== undefined ? { Услуги: services } : {}),
      };
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
      unitRef: UNIT,
      coefficient: 1,
      countryOfOriginRef: COUNTRY,
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
      ЕдиницаИзмерения_Key: UNIT,
      Коэффициент: 1,
      СтранаПроисхождения_Key: COUNTRY,
      Склад_Key: "00000000-0000-0000-0000-000000000000",
    });
    expect(postCount(fake)).toBe(0);
  });

  it("для Товары+Услуги требует явный operationKind и проверяет совместимость", async () => {
    const { client, fake } = await setup();
    const mixedBase = {
      ...args,
      sumIncludesVat: true,
      documentTotal: 300,
      serviceLines: [
        {
          nomenclatureRef: NOM,
          quantity: 1,
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
    const missing = await call(client, { ...mixedBase, confirm: false });
    expect(missing.isError).toBe(true);
    expect(JSON.stringify(missing)).toContain("operationKind");
    expect(postCount(fake)).toBe(0);

    const wrong = await call(client, { ...mixedBase, operationKind: "Товары", confirm: false });
    expect(wrong.isError).toBe(true);
    expect(JSON.stringify(wrong)).toContain("ПокупкаКомиссия");
  });

  it("смешанный документ: товарная и сервисная строки в отдельных таблицах", async () => {
    const { client } = await setup();
    const mixed = {
      ...args,
      operationKind: "ПокупкаКомиссия",
      sumIncludesVat: true,
      documentTotal: 300,
      lines: [
        {
          ...args.lines[0],
          quantity: 1.5,
          amount: 150,
          vatAmount: 30,
        },
      ],
      serviceLines: [
        {
          nomenclatureRef: NOM,
          quantity: 2,
          price: 50,
          amount: 100,
          vatRate: "НДС20",
          vatAmount: 20,
          content: "Услуга по договору",
          sourceLineId: "SRC-S-1",
          costAccountRef: ACC,
          costAccountNuRef: ACC,
          vatAccountRef: ACC,
          analytics: [{ ref: NOM, type: "StandardODATA.Catalog_Номенклатура" }],
          analyticsNu: [{ ref: NOM, type: "StandardODATA.Catalog_Номенклатура" }],
        },
      ],
    };
    const preview = await call(client, { ...mixed, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      ВидОперации: "ПокупкаКомиссия",
      СуммаВключаетНДС: true,
      СуммаДокумента: 300,
    });
    expect((payload["Товары"] as Array<Record<string, unknown>>)[0]).toMatchObject({
      Количество: 1.5,
      ЕдиницаИзмерения_Key: UNIT,
      Коэффициент: 1,
      СтранаПроисхождения_Key: COUNTRY,
      Склад_Key: "00000000-0000-0000-0000-000000000000",
    });
    expect((payload["Услуги"] as Array<Record<string, unknown>>)[0]).toMatchObject({
      Количество: 2,
      Сумма: 100,
      СуммаНДС: 20,
      СчетЗатрат_Key: ACC,
      СчетЗатратНУ_Key: ACC,
      СчетУчетаНДС_Key: ACC,
      ИдентификаторСтроки: "SRC-S-1",
      Субконто1: NOM,
      Субконто1_Type: "StandardODATA.Catalog_Номенклатура",
      СубконтоНУ1: NOM,
      СубконтоНУ1_Type: "StandardODATA.Catalog_Номенклатура",
    });
    const operationId = preview.structuredContent?.operationId;
    const created = await call(client, { ...mixed, confirm: true, operationId });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent?.verifiedFields).toEqual(
      expect.arrayContaining([
        "Товары[1].Склад_Key",
        "Услуги[1].СчетЗатрат_Key",
        "Услуги[1].СчетЗатратНУ_Key",
        "Услуги[1].Субконто1_Type",
        "Услуги[1].СубконтоНУ1_Type",
      ]),
    );
  });

  it("вид «Услуги» без таблицы «Товары» и с пустым складом шапки", async () => {
    const { client } = await setup();
    const serviceOnly = {
      ...args,
      lines: undefined,
      warehouse: undefined,
      sumIncludesVat: true,
      documentTotal: 120,
      serviceLines: [
        {
          nomenclatureRef: NOM,
          quantity: 1,
          price: 100,
          amount: 100,
          vatRate: "НДС20",
          vatAmount: 20,
          content: "Услуга",
          sourceLineId: "SRC-SVC",
          costAccountRef: ACC,
          costAccountNuRef: ACC,
          vatAccountRef: ACC,
          analytics: [{ ref: NOM, type: "StandardODATA.Catalog_Номенклатура" }],
          analyticsNu: [{ ref: NOM, type: "StandardODATA.Catalog_Номенклатура" }],
        },
      ],
    };
    const preview = await call(client, { ...serviceOnly, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      ВидОперации: "Услуги",
      Склад_Key: "00000000-0000-0000-0000-000000000000",
      СуммаДокумента: 120,
    });
    expect(payload).not.toHaveProperty("Товары");
    expect((payload["Услуги"] as Array<Record<string, unknown>>)[0]).toMatchObject({
      Сумма: 100,
      СуммаНДС: 20,
      ИдентификаторСтроки: "SRC-SVC",
    });
    const operationId = preview.structuredContent?.operationId;
    const created = await call(client, { ...serviceOnly, confirm: true, operationId });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
  });

  it("emptyQuantity услуги: без подстановки 1, с явными суммой и НДС", async () => {
    const { client, fake } = await setup();
    const emptyService = {
      ...args,
      lines: undefined,
      warehouse: undefined,
      sumIncludesVat: true,
      documentTotal: 120,
      serviceLines: [
        {
          nomenclatureRef: NOM,
          emptyQuantity: true,
          price: 100,
          amount: 100,
          vatRate: "НДС20",
          vatAmount: 20,
          content: "Услуга без количества",
          costAccountRef: ACC,
          costAccountNuRef: ACC,
          vatAccountRef: ACC,
        },
      ],
    };
    const preview = await call(client, { ...emptyService, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const row = (preview.structuredContent?.payload as Record<string, unknown>)["Услуги"] as Array<
      Record<string, unknown>
    >;
    expect(row[0]).not.toHaveProperty("Количество");
    expect(row[0]).toMatchObject({ Сумма: 100, СуммаНДС: 20 });
    expect(preview.structuredContent?.quantityVerification).toBe("pending_live_validation");
    const operationId = preview.structuredContent?.operationId;
    const created = await call(client, { ...emptyService, confirm: true, operationId });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({
      verification: "ok",
      quantityVerification: "pending_live_validation",
    });
    expect(postCount(fake)).toBe(1);
  });

  it("emptyQuantity услуги требует amount/vatAmount и не сочетается с quantity", async () => {
    const { client, fake } = await setup();
    const missing = await call(client, {
      ...args,
      lines: undefined,
      serviceLines: [{ nomenclatureRef: NOM, emptyQuantity: true, price: 100 }],
      confirm: false,
    });
    expect(missing.isError).toBe(true);
    const both = await call(client, {
      ...args,
      lines: undefined,
      serviceLines: [
        { nomenclatureRef: NOM, quantity: 1, emptyQuantity: true, price: 100, amount: 100, vatAmount: 20 },
      ],
      confirm: false,
    });
    expect(both.isError).toBe(true);
    expect(postCount(fake)).toBe(0);
  });

  it("непустое количество после записи пустой услуги ломает verification", async () => {
    const { client } = await setup({ serviceQuantity: 1 });
    const emptyService = {
      ...args,
      lines: undefined,
      warehouse: undefined,
      sumIncludesVat: true,
      documentTotal: 120,
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
    const preview = await call(client, { ...emptyService, confirm: false });
    const created = await call(client, {
      ...emptyService,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError, JSON.stringify(created)).toBe(true);
    expect(JSON.stringify(created)).toContain("Количество");
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
        "Товары[1].ЕдиницаИзмерения_Key",
        "Товары[1].Коэффициент",
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

  it("принимает только известный подходящий вид договора", async () => {
    const supplier = await setup();
    const ok = await call(supplier.client, { ...args, confirm: false });
    expect(ok.isError, JSON.stringify(ok)).toBeFalsy();

    const customer = await setup({ contractKind: "СПокупателем" });
    const bad = await call(customer.client, { ...args, confirm: false });
    expect(bad.isError).toBe(true);
    expect(JSON.stringify(bad)).toContain("СПокупателем");
    expect(postCount(customer.fake)).toBe(0);

    const empty = await setup({ contractKind: "" });
    const emptyResult = await call(empty.client, { ...args, confirm: false });
    expect(emptyResult.isError).toBe(true);
    expect(JSON.stringify(emptyResult)).toContain("не имеет вида");
    expect(postCount(empty.fake)).toBe(0);

    const generic = await setup({ contractKind: "Прочее" });
    const genericResult = await call(generic.client, { ...args, confirm: false });
    expect(genericResult.isError).toBe(true);
    expect(JSON.stringify(genericResult)).toContain("Прочее");
    const allowed = await call(generic.client, {
      ...args,
      allowGenericContractKind: true,
      confirm: false,
    });
    expect(allowed.isError, JSON.stringify(allowed)).toBeFalsy();
    expect(postCount(generic.fake)).toBe(0);
  });

  it("time без date не подставляет текущую дату", async () => {
    const { client, fake } = await setup();
    const result = await call(client, { ...args, date: undefined, confirm: false });
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result)).toContain("date");
    expect(postCount(fake)).toBe(0);
  });

  it("единица и коэффициент участвуют в GET: пустая единица/0 ломают verification", async () => {
    const { client } = await setup({ rowUnit: "", rowCoefficient: 0 });
    const preview = await call(client, { ...args, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    const created = await call(client, { ...args, confirm: true, operationId });
    expect(created.isError).toBe(true);
    expect(JSON.stringify(created)).toContain("ЕдиницаИзмерения_Key");
    expect(JSON.stringify(created)).toContain("Коэффициент");
    expect(JSON.stringify(created)).toContain("Ref_Key");
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
