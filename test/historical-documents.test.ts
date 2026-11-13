import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";
import { ODataError } from "../src/odata/errors.js";

/**
 * Исторические opt-in поля существующих инструментов: счёт покупателю и номенклатура.
 * Пустое количество и срок оплаты сознательно не поддержаны — проверяем отказ, а не подстановку 1.
 */

const ORG = "11111111-1111-4111-8111-111111111111";
const CP = "22222222-2222-4222-8222-222222222222";
const CTR = "33333333-3333-4333-8333-333333333333";
const ACC = "44444444-4444-4444-8444-444444444444";
const NOM = "55555555-5555-4555-8555-555555555555";
const GROUP = "66666666-6666-4666-8666-666666666666";
const KIND = "77777777-7777-4777-8777-777777777777";
const UNIT = "88888888-8888-4888-8888-888888888888";
const DOC = "99999999-9999-4999-8999-999999999999";

const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(cleanup.splice(0).map((fn) => fn().catch(() => undefined)));
});

interface Options {
  nomenclature?: Record<string, unknown>[];
  invoices?: Record<string, unknown>[];
  missingUnit?: boolean;
  badVatReadback?: boolean;
}

async function setup(options: Options = {}) {
  const dir = await mkdtemp(join(tmpdir(), "historical-docs-"));
  const data = (q: FakeQuery): Record<string, unknown>[] | undefined => {
    if (q.base === "Catalog_Номенклатура") return options.nomenclature ?? [];
    if (q.base === "Document_СчетНаОплатуПокупателю") return options.invoices ?? [];
    return undefined;
  };
  const fake = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" }, { data });
  const available = await fake.conn.available();
  vi.spyOn(fake.conn, "available").mockResolvedValue(
    new Set([
      ...available,
      "Catalog_ВидыНоменклатуры",
      "Catalog_КлассификаторЕдиницИзмерения",
      "Catalog_ЕдиницыИзмерения",
    ]),
  );
  const unitLookups: string[] = [];
  const originalGetEntity = fake.conn.client.getEntity.bind(fake.conn.client);
  vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
    if (
      options.missingUnit &&
      (path.startsWith("Catalog_ЕдиницыИзмерения(") ||
        path.startsWith("Catalog_КлассификаторЕдиницИзмерения("))
    )
      throw new ODataError({ kind: "not_found", message: "нет единицы" });
    if (
      options.badVatReadback &&
      path.startsWith("Document_СчетНаОплатуПокупателю(") &&
      !path.includes("$select")
    ) {
      const actual = await originalGetEntity(path);
      const rows = actual["Товары"] as Array<Record<string, unknown>>;
      return { ...actual, Товары: [{ ...rows[0], СуммаНДС: 0 }] };
    }
    if (path.startsWith("Catalog_ДоговорыКонтрагентов("))
      return {
        Ref_Key: CTR,
        Owner_Key: CP,
        Организация_Key: ORG,
        ВидДоговора: "СПокупателем",
        DeletionMark: false,
      };
    const known =
      /^Catalog_(ВидыНоменклатуры|ЕдиницыИзмерения|КлассификаторЕдиницИзмерения)\(guid'([^']+)'\)/.exec(path);
    if (known) {
      unitLookups.push(path);
      return { Ref_Key: known[2]!, DeletionMark: false };
    }
    return originalGetEntity(path);
  });
  const server = createServer(fake.ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "historical-test", version: "1" });
  await client.connect(b);
  cleanup.push(async () => {
    await client.close().catch(() => undefined);
    await server.close().catch(() => undefined);
    await rm(dir, { recursive: true, force: true });
  });
  return { fake, client, unitLookups };
}

const postCount = (fake: ReturnType<typeof fake1C>): number =>
  fake.requests.filter((r) => r.method === "POST").length;

const call = async (client: Client, name: string, args: Record<string, unknown>) =>
  (await client.callTool({ name, arguments: args })) as {
    isError?: boolean;
    structuredContent?: Record<string, unknown>;
    content?: Array<{ text?: string }>;
  };

async function confirmAfterPreview(
  client: Client,
  name: string,
  args: Record<string, unknown>,
): Promise<ReturnType<typeof call>> {
  const preview = await call(client, name, { ...args, confirm: false });
  expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
  const operationId = preview.structuredContent?.operationId;
  expect(typeof operationId).toBe("string");
  return call(client, name, { ...args, confirm: true, operationId });
}

describe("write.sales.create_invoice: исторические opt-in поля", () => {
  const invoiceArgs = {
    organization: ORG,
    counterpartyRef: CP,
    contractRef: CTR,
    orgBankAccount: ACC,
    date: "2026-11-01",
    number: "INV-1",
    time: "14:35:20",
    comment: "Исторический счёт",
    sumIncludesVat: true,
    lines: [
      {
        nomenclatureRef: NOM,
        quantity: 1,
        price: 1000,
        amount: 1000,
        vatRate: "НДС18",
        vatAmount: 180,
        content: "Услуга по договору",
      },
    ],
  };

  it("переносит номер/время, исторический НДС, сумму и содержание, затем проверяет readback", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, "write.sales.create_invoice", { ...invoiceArgs, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      Number: "INV-1",
      Date: "2026-11-01T14:35:20",
      Комментарий: "Исторический счёт",
      СуммаВключаетНДС: true,
      СуммаДокумента: 1000,
    });
    expect((payload["Товары"] as Array<Record<string, unknown>>)[0]).toMatchObject({
      Количество: 1,
      Сумма: 1000,
      СтавкаНДС: "НДС18",
      СуммаНДС: 180,
      Содержание: "Услуга по договору",
    });
    const created = await confirmAfterPreview(
      client,
      "write.sales.create_invoice",
      invoiceArgs as Record<string, unknown>,
    );
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({ created: true, verification: "ok" });
    expect(postCount(fake)).toBe(1);
  });

  it("sumIncludesVat=false: исторический итог документа включает net и НДС", async () => {
    const { client, fake } = await setup();
    const net = {
      ...invoiceArgs,
      sumIncludesVat: false,
      lines: [
        {
          nomenclatureRef: NOM,
          quantity: 1,
          price: 100,
          amount: 100,
          vatRate: "НДС20",
          vatAmount: 20,
          content: "Услуга net",
        },
      ],
    };
    const preview = await call(client, "write.sales.create_invoice", { ...net, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({ СуммаВключаетНДС: false, СуммаДокумента: 120 });
    expect((payload["Товары"] as Array<Record<string, unknown>>)[0]).toMatchObject({
      Сумма: 100,
      СуммаНДС: 20,
    });
    expect(postCount(fake)).toBe(0);
  });

  it("emptyQuantity: preview не сериализует Количество и помечает readback pending", async () => {
    const { client, fake } = await setup();
    const emptyArgs = {
      ...invoiceArgs,
      lines: [
        {
          nomenclatureRef: NOM,
          emptyQuantity: true,
          price: 1000,
          amount: 1000,
          vatRate: "НДС18",
          vatAmount: 180,
          content: "Услуга без количества",
        },
      ],
    };
    const preview = await call(client, "write.sales.create_invoice", { ...emptyArgs, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    const row = (payload["Товары"] as Array<Record<string, unknown>>)[0]!;
    expect(Object.prototype.hasOwnProperty.call(row, "Количество")).toBe(false);
    expect(row).toMatchObject({
      Цена: 1000,
      Сумма: 1000,
      СтавкаНДС: "НДС18",
      СуммаНДС: 180,
      Содержание: "Услуга без количества",
    });
    expect(preview.structuredContent?.quantityVerification).toBe("pending_live_validation");
    expect(postCount(fake)).toBe(0);
  });

  it("emptyQuantity требует amount/vatAmount и не сочетается с quantity", async () => {
    const { client, fake } = await setup();
    const missing = await call(client, "write.sales.create_invoice", {
      ...invoiceArgs,
      lines: [{ nomenclatureRef: NOM, emptyQuantity: true, price: 1000 }],
      confirm: false,
    });
    expect(missing.isError).toBe(true);
    const both = await call(client, "write.sales.create_invoice", {
      ...invoiceArgs,
      lines: [
        {
          nomenclatureRef: NOM,
          quantity: 1,
          emptyQuantity: true,
          price: 1000,
          amount: 1000,
          vatAmount: 180,
        },
      ],
      confirm: false,
    });
    expect(both.isError).toBe(true);
    expect(postCount(fake)).toBe(0);
  });

  it("ловит конфликт номера в том же году и разрешает тот же номер в другом", async () => {
    const { client, fake } = await setup({
      invoices: [{ Ref_Key: DOC, Number: "INV-1", Date: "2026-05-01T00:00:00", Организация_Key: ORG }],
    });
    const conflict = await call(client, "write.sales.create_invoice", { ...invoiceArgs, confirm: false });
    expect(conflict.isError).toBe(true);
    expect(JSON.stringify(conflict)).toContain("уже занят");
    expect(postCount(fake)).toBe(0);
    const otherYear = await call(client, "write.sales.create_invoice", {
      ...invoiceArgs,
      date: "2025-11-01",
      confirm: false,
    });
    expect(otherYear.isError, JSON.stringify(otherYear)).toBeFalsy();
  });

  it("не подставляет количество 1: пустое/нулевое количество отклоняется", async () => {
    const { client, fake } = await setup();
    const noQuantity = await call(client, "write.sales.create_invoice", {
      ...invoiceArgs,
      lines: [{ ...invoiceArgs.lines[0], quantity: undefined }],
      confirm: false,
    });
    expect(noQuantity.isError).toBe(true);
    const zero = await call(client, "write.sales.create_invoice", {
      ...invoiceArgs,
      lines: [{ ...invoiceArgs.lines[0], quantity: 0 }],
      confirm: false,
    });
    expect(zero.isError).toBe(true);
    expect(postCount(fake)).toBe(0);
  });

  it("расхождение readback по НДС помечает создание failed без повторного POST", async () => {
    const { client, fake } = await setup({ badVatReadback: true });
    const preview = await call(client, "write.sales.create_invoice", { ...invoiceArgs, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    expect(typeof operationId).toBe("string");
    const created = await call(client, "write.sales.create_invoice", {
      ...invoiceArgs,
      confirm: true,
      operationId,
    });
    expect(created.isError).toBe(true);
    expect(JSON.stringify(created)).toContain("СуммаНДС");
    expect(postCount(fake)).toBe(1);
  });
});

describe("write.catalog.create_nomenclature: исторические opt-in поля", () => {
  const nomArgs = {
    name: "Историческая услуга",
    fullName: "Историческая услуга (полное)",
    article: "ART-1",
    code: "HIST-001",
    groupRef: GROUP,
    serviceKindRef: KIND,
    unitRef: UNIT,
    vatKind: "Общая",
    isService: true,
  };

  it("переносит Code/вид/единицу/группу/ВидСтавкиНДС и проверяет readback", async () => {
    const { client, fake, unitLookups } = await setup();
    const preview = await call(client, "write.catalog.create_nomenclature", {
      ...nomArgs,
      confirm: false,
    });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    expect(preview.structuredContent?.payload).toMatchObject({
      Code: "HIST-001",
      НоменклатурнаяГруппа_Key: GROUP,
      ВидНоменклатуры_Key: KIND,
      ЕдиницаИзмерения_Key: UNIT,
      ВидСтавкиНДС: "Общая",
      Услуга: true,
    });
    const created = await confirmAfterPreview(client, "write.catalog.create_nomenclature", nomArgs);
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({ created: true, verification: "ok" });
    expect(
      unitLookups.some((p) => p.startsWith("Catalog_КлассификаторЕдиницИзмерения(")),
      "единица должна проверяться в актуальном каталоге BP",
    ).toBe(true);
    expect(postCount(fake)).toBe(1);
  });

  it("отказывает при занятом Code и не пишет", async () => {
    const { client, fake } = await setup({
      nomenclature: [{ Ref_Key: NOM, Code: "HIST-001", Description: "Уже есть" }],
    });
    const result = await call(client, "write.catalog.create_nomenclature", {
      ...nomArgs,
      confirm: false,
    });
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result)).toContain("уже занят");
    expect(postCount(fake)).toBe(0);
  });

  it("отказывает при отсутствующей единице измерения", async () => {
    const { client, fake } = await setup({ missingUnit: true });
    const result = await call(client, "write.catalog.create_nomenclature", {
      ...nomArgs,
      confirm: false,
    });
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result)).toContain("Единица измерения");
    expect(postCount(fake)).toBe(0);
  });
});
