import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

/** Синтетическая историческая реализация услуг: точные реквизиты без production-записи. */

const ORG = "11111111-1111-4111-8111-111111111111";
const CP = "22222222-2222-4222-8222-222222222222";
const CTR = "33333333-3333-4333-8333-333333333333";
const ACC = "11111111-1111-4111-8111-111111111111";
const NOM = "55555555-5555-4555-8555-555555555555";
const GROUP = "66666666-6666-4666-8666-666666666666";
const DOC = "99999999-9999-4999-8999-999999999999";
const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";

const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(cleanup.splice(0).map((fn) => fn().catch(() => undefined)));
});

interface Options {
  documents?: Record<string, unknown>[];
  missingUniversalField?: boolean;
  bankOwner?: string;
  autoBankAccount?: string;
  stringLineNumber?: boolean;
}

async function setup(options: Options = {}) {
  const dir = await mkdtemp(join(tmpdir(), "historical-service-"));
  const data = (q: FakeQuery): Record<string, unknown>[] | undefined =>
    q.base === "Document_РеализацияТоваровУслуг" ? (options.documents ?? []) : undefined;
  const fake = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" }, { data });
  const originalGetEntity = fake.conn.client.getEntity.bind(fake.conn.client);
  vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
    if (path.startsWith("Catalog_ДоговорыКонтрагентов("))
      return {
        Ref_Key: CTR,
        Owner_Key: CP,
        Организация_Key: ORG,
        ВидДоговора: "СПокупателем",
        DeletionMark: false,
      };
    if (path.startsWith("Catalog_НоменклатурныеГруппы(")) return { Ref_Key: GROUP, DeletionMark: false };
    if (
      (options.autoBankAccount || options.stringLineNumber) &&
      path.startsWith("Document_РеализацияТоваровУслуг(") &&
      !path.includes("$select")
    ) {
      const actual = await originalGetEntity(path);
      const rows = (actual["Услуги"] as Array<Record<string, unknown>>).map((row) => ({
        ...row,
        LineNumber: options.stringLineNumber ? String(row["LineNumber"]) : row["LineNumber"],
      }));
      return {
        ...actual,
        Posted: false,
        DeletionMark: false,
        ...(options.autoBankAccount ? { БанковскийСчетОрганизации_Key: options.autoBankAccount } : {}),
        Услуги: rows,
      };
    }
    return originalGetEntity(path);
  });
  if (options.missingUniversalField) {
    const meta = await fake.conn.getMetadata();
    const set = "Document_РеализацияТоваровУслуг";
    const doc = meta.entities.get(set)!;
    const entities = new Map(meta.entities);
    entities.set(set, {
      ...doc,
      properties: doc.properties.filter((p) => p.name !== "ЭтоУниверсальныйДокумент"),
    });
    vi.spyOn(fake.conn, "getMetadata").mockResolvedValue({ ...meta, entities });
  }
  const server = createServer(fake.ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "historical-service-test", version: "1" });
  await client.connect(b);
  cleanup.push(async () => {
    await client.close().catch(() => undefined);
    await server.close().catch(() => undefined);
    await rm(dir, { recursive: true, force: true });
  });
  return { fake, client };
}

const call = async (client: Client, name: string, args: Record<string, unknown>) =>
  (await client.callTool({ name, arguments: args })) as {
    isError?: boolean;
    structuredContent?: Record<string, unknown>;
  };

const postCount = (fake: ReturnType<typeof fake1C>): number =>
  fake.requests.filter((r) => r.method === "POST").length;

const actArgs = {
  organization: ORG,
  counterpartyRef: CP,
  contractRef: CTR,
  date: "2026-11-01",
  number: "ACT-1",
  time: "16:20:30",
  comment: "Исторический акт",
  documentTotal: 120,
  settlementAccountRef: ACC,
  advanceAccountRef: ACC,
  orgBankAccountEmpty: true,
  universalDocument: true,
  sumIncludesVat: false,
  lines: [
    {
      nomenclatureRef: NOM,
      emptyQuantity: true,
      price: 100,
      amount: 100,
      vatRate: "НДС18",
      vatAmount: 20,
      content: "Услуга по договору",
      sourceLineId: "SRC-LINE-1",
      accountingGroupRef: GROUP,
      accountingGroupType: "StandardODATA.Catalog_НоменклатурныеГруппы",
      incomeAccountRef: ACC,
      expenseAccountRef: ACC,
      vatAccountRef: ACC,
    },
  ],
};

describe("write.sales.create_act: исторический opt-in", () => {
  it("сериализует точные поля, пустое количество и пустой банковский счёт", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, "write.sales.create_act", { ...actArgs, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      Number: "ACT-1",
      Date: "2026-11-01T16:20:30",
      Комментарий: "Исторический акт",
      ЭтоУниверсальныйДокумент: true,
      СуммаВключаетНДС: false,
      СуммаДокумента: 120,
      СчетУчетаРасчетовСКонтрагентом_Key: ACC,
      СчетУчетаРасчетовПоАвансам_Key: ACC,
    });
    expect(payload).toMatchObject({ БанковскийСчетОрганизации_Key: EMPTY_GUID });
    expect(payload).not.toHaveProperty("ВидЭлектронногоДокумента");
    const row = (payload["Услуги"] as Array<Record<string, unknown>>)[0]!;
    expect(row).not.toHaveProperty("Количество");
    expect(row).toMatchObject({
      Сумма: 100,
      СуммаНДС: 20,
      СтавкаНДС: "НДС18",
      Содержание: "Услуга по договору",
      ИдентификаторСтроки: "SRC-LINE-1",
      Субконто: GROUP,
      Субконто_Type: "StandardODATA.Catalog_НоменклатурныеГруппы",
      СчетДоходов_Key: ACC,
      СчетРасходов_Key: ACC,
      СчетУчетаНДСПоРеализации_Key: ACC,
    });
    expect(preview.structuredContent?.quantityVerification).toBe("pending_live_validation");
    expect(postCount(fake)).toBe(0);
  });

  it("подтверждает исторический акт с readback и не отправляет пустое количество", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, "write.sales.create_act", { ...actArgs, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    expect(typeof operationId).toBe("string");
    const created = await call(client, "write.sales.create_act", {
      ...actArgs,
      confirm: true,
      operationId,
    });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({
      created: true,
      verification: "ok",
      quantityVerification: "pending_live_validation",
    });
    expect(created.structuredContent?.verifiedFields).toEqual(
      expect.arrayContaining(["Posted", "DeletionMark", "ВидОперации", "Услуги[1].Субконто_Type"]),
    );
    expect(postCount(fake)).toBe(1);
  });

  it("явный банковский счёт без empty проходит проверку владельца", async () => {
    const { client, fake } = await setup({ bankOwner: ORG });
    const args = { ...actArgs, orgBankAccountEmpty: false, orgBankAccount: ACC };
    const preview = await call(client, "write.sales.create_act", { ...args, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    expect(
      (preview.structuredContent?.payload as Record<string, unknown>)["БанковскийСчетОрганизации_Key"],
    ).toBe(ACC);
    expect(postCount(fake)).toBe(0);
  });

  it("независимый тип электронного документа не выводится из флага", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, "write.sales.create_act", {
      ...actArgs,
      electronicDocumentType: "УПД",
      confirm: false,
    });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      ЭтоУниверсальныйДокумент: true,
      ВидЭлектронногоДокумента: "УПД",
    });
    expect(postCount(fake)).toBe(0);
  });

  it("universalDocument=false сохраняется явно", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, "write.sales.create_act", {
      ...actArgs,
      universalDocument: false,
      confirm: false,
    });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({ ЭтоУниверсальныйДокумент: false });
    expect(payload).not.toHaveProperty("ВидЭлектронногоДокумента");
    expect(postCount(fake)).toBe(0);
  });

  it("автоназначение непустого счёта ломает readback: failed, Ref_Key, без второго POST", async () => {
    const { client, fake } = await setup({ autoBankAccount: ACC });
    const preview = await call(client, "write.sales.create_act", { ...actArgs, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    expect(typeof operationId).toBe("string");
    const created = await call(client, "write.sales.create_act", {
      ...actArgs,
      confirm: true,
      operationId,
    });
    expect(created.isError, JSON.stringify(created)).toBe(true);
    expect(JSON.stringify(created)).toContain("БанковскийСчетОрганизации_Key");
    expect(JSON.stringify(created)).toContain("Ref_Key");
    expect(created.structuredContent).toMatchObject({ verification: "failed" });
    expect(postCount(fake)).toBe(1);
  });

  it("строковый LineNumber в readback не ломает сверку", async () => {
    const { client, fake } = await setup({ stringLineNumber: true });
    const preview = await call(client, "write.sales.create_act", { ...actArgs, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    expect(typeof operationId).toBe("string");
    const created = await call(client, "write.sales.create_act", {
      ...actArgs,
      confirm: true,
      operationId,
    });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({ verification: "ok" });
    expect(postCount(fake)).toBe(1);
  });

  it("непубликованный флаг УПД блокирует подтверждение", async () => {
    const { client, fake } = await setup({ missingUniversalField: true });
    const preview = await call(client, "write.sales.create_act", { ...actArgs, confirm: false });
    expect(preview.isError).toBe(true);
    expect(JSON.stringify(preview)).toContain("ЭтоУниверсальныйДокумент");
    expect(postCount(fake)).toBe(0);
  });

  it("конфликт номера в том же году блокирует акт", async () => {
    const { client, fake } = await setup({
      documents: [{ Ref_Key: DOC, Number: "ACT-1", Date: "2026-05-01T00:00:00", Организация_Key: ORG }],
    });
    const preview = await call(client, "write.sales.create_act", { ...actArgs, confirm: false });
    expect(preview.isError).toBe(true);
    expect(JSON.stringify(preview)).toContain("уже занят");
    expect(postCount(fake)).toBe(0);
  });
});
