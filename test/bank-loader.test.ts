import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";
import { ODataError } from "../src/odata/errors.js";
import { BankLoadJournal, legacySourceHash } from "../src/bank-loader/journal.js";
import {
  ACCOUNT_FIND_TOOL,
  diagnose,
  ensureAccount,
  PAYMENT_TOOL,
  runPackage,
  type ToolClient,
} from "../src/bank-loader/runner.js";
import type { SourceDocument, TargetRequisites } from "../src/bank-loader/types.js";

const ORG = "11111111-1111-4111-8111-111111111111";
const CP = "22222222-2222-4222-8222-222222222222";
const CTR = "33333333-3333-4333-8333-333333333333";
const ACC = "44444444-4444-4444-8444-444444444444";
const CP_ACC = "55555555-5555-4555-8555-555555555555";
const BANK = "66666666-6666-4666-8666-666666666666";
const CUR = "77777777-7777-4777-8777-777777777777";
const DDS = "88888888-8888-4888-8888-888888888888";
const INV = "99999999-9999-4999-8999-999999999999";
const DOC = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const ACCOUNT_NUMBER = "40702810000000000001";

const target: TargetRequisites = {
  organizationInn: "7700000000",
  organizationKpp: "770000001",
  bankAccountNumber: ACCOUNT_NUMBER,
  bankBik: "044525225",
};

const source: SourceDocument = {
  sourceKey: "statement-2026-11-01-0001",
  date: "2026-11-01",
  number: "MY-1",
  time: "14:35:20",
  amount: 1180,
  payerInn: "7800000000",
  payerKpp: "780000001",
  payerName: "Синтетический покупатель",
  purpose: "Оплата по договору",
  documentNumber: "BANK-77",
  documentDate: "2026-10-31",
  contractNumber: "D-1",
  vatRate: "НДС18",
  vatAmount: 180,
  cashFlowItemRef: "88888888-8888-4888-8888-888888888888",
  customerInvoiceRef: "99999999-9999-4999-8999-999999999999",
};

const orgRow = {
  Ref_Key: ORG,
  Description: "Синтетическая организация",
  ИНН: target.organizationInn,
  КПП: target.organizationKpp,
  DeletionMark: false,
};

const cpRow = {
  Ref_Key: CP,
  Description: "Синтетический покупатель",
  ИНН: source.payerInn,
  КПП: source.payerKpp,
  DeletionMark: false,
  IsFolder: false,
};

const contractRow = {
  Ref_Key: CTR,
  Description: "Договор покупателя",
  Номер: "D-1",
  ВидДоговора: "СПокупателем",
  Owner_Key: CP,
  Организация_Key: ORG,
  DeletionMark: false,
  IsFolder: false,
};

const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(cleanup.splice(0).map((fn) => fn().catch(() => undefined)));
});

interface FixtureOptions {
  organizations?: Record<string, unknown>[];
  counterparties?: Record<string, unknown>[];
  contracts?: Record<string, unknown>[];
  bankAccounts?: Record<string, unknown>[];
  documents?: Record<string, unknown>[];
  cashFlowItems?: Record<string, unknown>[];
  invoices?: Record<string, unknown>[];
  banks?: Record<string, unknown>[];
  env?: Record<string, string>;
  dropPostResponse?: boolean;
  respectSelect?: boolean;
}

async function fixture(options: FixtureOptions = {}) {
  const journalDir = await mkdtemp(join(tmpdir(), "bank-loader-journal-"));
  const writeDir = await mkdtemp(join(tmpdir(), "bank-loader-write-"));
  const orgs = options.organizations ?? [orgRow];
  const counterparties = options.counterparties ?? [cpRow];
  const contracts = options.contracts ?? [contractRow];
  const accounts = options.bankAccounts ?? [
    {
      Ref_Key: ACC,
      Description: "Счёт организации",
      НомерСчета: ACCOUNT_NUMBER,
      Owner: ORG,
      Owner_Type: "StandardODATA.Catalog_Организации",
      DeletionMark: false,
    },
  ];
  const data = (q: FakeQuery): Record<string, unknown>[] | undefined => {
    if (q.base.startsWith("ChartOfAccounts_")) return undefined; // коды счетов отвечает сама подделка
    switch (q.base) {
      case "Catalog_Организации":
        return orgs;
      case "Catalog_Контрагенты":
        return counterparties;
      case "Catalog_ДоговорыКонтрагентов":
        return contracts;
      case "Catalog_БанковскиеСчета":
        return accounts;
      case "Catalog_Валюты":
        return [{ Ref_Key: CUR, Code: "643", Description: "Российский рубль" }];
      case "Catalog_Банки":
        return (
          options.banks ?? [
            { Ref_Key: BANK, Code: target.bankBik, Description: "Синтетический банк", DeletionMark: false },
          ]
        );
      case "Document_ПоступлениеНаРасчетныйСчет":
        return options.documents ?? [];
      case "Catalog_СтатьиДвиженияДенежныхСредств":
        return options.cashFlowItems ?? [{ Ref_Key: DDS, Description: "Статья ДДС", DeletionMark: false }];
      case "Document_СчетНаОплатуПокупателю":
        return options.invoices ?? [{ Ref_Key: INV, Description: "Счёт покупателю", DeletionMark: false }];
      case "InformationRegister_СчетаУчетаРасчетовСКонтрагентами":
        return [];
      default:
        return undefined;
    }
  };
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: writeDir,
      ...(options.env ?? {}),
    },
    {
      data,
      ...(options.dropPostResponse ? { dropPostResponse: true } : {}),
      ...(options.respectSelect ? { respectSelect: true } : {}),
    },
  );
  const server = createServer(fake.ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "bank-loader-test", version: "1" });
  await client.connect(b);
  const journal = new BankLoadJournal(journalDir);
  cleanup.push(async () => {
    await client.close().catch(() => undefined);
    await server.close().catch(() => undefined);
    await rm(journalDir, { recursive: true, force: true });
    await rm(writeDir, { recursive: true, force: true });
  });
  return { fake, client, journal, state: { organizations: orgs, counterparties, contracts, accounts } };
}

const postCount = (fake: ReturnType<typeof fake1C>): number =>
  fake.requests.filter((r) => r.method === "POST").length;

/** Оператор сначала видит и проверяет план, затем подтверждает его. */
async function prepareAndConfirm(
  client: Parameters<typeof runPackage>[0],
  target: TargetRequisites,
  source: SourceDocument,
  journal: BankLoadJournal,
  mode: "confirm" | "resume" = "confirm",
) {
  await runPackage(client, target, [source], journal, { confirm: false, mode: "prepare" });
  return runPackage(client, target, [source], journal, { confirm: true, mode });
}

const paymentArgs = (operationId?: string) => ({
  sourceKey: source.sourceKey,
  organizationInn: target.organizationInn,
  organizationKpp: target.organizationKpp,
  orgBankAccountNumber: target.bankAccountNumber,
  payerInn: source.payerInn,
  payerKpp: source.payerKpp,
  payerName: source.payerName,
  amount: source.amount,
  date: source.date,
  number: source.number,
  time: source.time,
  purpose: source.purpose,
  documentNumber: "BANK-77",
  documentDate: "2026-10-31",
  vatRate: "НДС18",
  vatAmount: 180,
  cashFlowItemRef: DDS,
  customerInvoiceRef: INV,
  contractNumber: "D-1",
  ...(operationId ? { operationId } : {}),
});

describe("операторская диагностика через MCP", () => {
  it("читает схему, находит организацию по ИНН/КПП и её счёт, не используя $top=1", async () => {
    const { client } = await fixture();
    const report = await diagnose(client, target);
    expect(report.ok, JSON.stringify(report)).toBe(true);
    expect(report.schema?.odataVersion).toBe("3.0");
    expect(report.organization?.ref).toBe(ORG);
    expect(report.account).toMatchObject({ found: true, ref: ACC });
  });

  it("при HTTP 500 на фильтре по ИНН переходит к ограниченному полному проходу", async () => {
    const { client, fake } = await fixture();
    const original = fake.conn.client.getCollection.bind(fake.conn.client);
    vi.spyOn(fake.conn.client, "getCollection").mockImplementation(async (path: string) => {
      if (path.includes("Catalog_Организации") && path.includes("$filter"))
        throw new ODataError({ kind: "server", status: 500, message: "Сервер 1С: 500" });
      return original(path);
    });
    const report = await diagnose(client, target);
    expect(report.ok, JSON.stringify(report)).toBe(true);
    expect(report.organizations?.strategy).toBe("scan");
    expect(report.organization?.ref).toBe(ORG);
  });

  it("отказывает при неоднозначных ИНН/КПП, не выбирая первую организацию", async () => {
    const { client } = await fixture({
      organizations: [orgRow, { ...orgRow, Ref_Key: CP, Description: "Дубль" }],
    });
    const report = await diagnose(client, target);
    expect(report.ok).toBe(false);
    expect(report.problems.join(" ")).toContain("подходит 2");
  });

  it("не выдаёт счёт контрагента за счёт организации", async () => {
    const { client } = await fixture({
      bankAccounts: [
        {
          Ref_Key: CP_ACC,
          Description: "Счёт контрагента",
          НомерСчета: ACCOUNT_NUMBER,
          Owner: CP,
          Owner_Type: "StandardODATA.Catalog_Контрагенты",
          DeletionMark: false,
        },
      ],
    });
    const found = (await client.callTool({
      name: ACCOUNT_FIND_TOOL,
      arguments: { organization: ORG, accountNumber: ACCOUNT_NUMBER },
    })) as { structuredContent?: Record<string, unknown> };
    expect(found.structuredContent).toMatchObject({
      found: false,
      ownerType: "StandardODATA.Catalog_Организации",
    });
    const report = await diagnose(client, target);
    expect(report.ok).toBe(false);
    expect(report.problems.join(" ")).toContain("нет банковского счёта");
  });
});

describe("подготовка оплаты покупателя", () => {
  it("MCP и загрузчик дают одинаковый план, включая явные НДС и входящий документ", async () => {
    const { client, fake, journal } = await fixture();
    const direct = (await client.callTool({
      name: PAYMENT_TOOL,
      arguments: paymentArgs(),
    })) as { isError?: boolean; structuredContent?: Record<string, unknown> };
    expect(direct.isError, JSON.stringify(direct)).toBeFalsy();
    const run = await runPackage(client, target, [source], journal, { confirm: false });
    expect(run.documents[0]?.action).toBe("dry_run");
    const plan = run.documents[0]?.plan as { payload?: Record<string, unknown>; resolved?: unknown };
    expect(plan.payload).toEqual(direct.structuredContent?.payload);
    expect(plan.resolved).toEqual(direct.structuredContent?.resolved);
    const row = (plan.payload?.["РасшифровкаПлатежа"] as Array<Record<string, unknown>>)[0]!;
    expect(row).toMatchObject({ СтавкаНДС: "НДС18", СуммаНДС: 180 });
    expect(plan.payload).toMatchObject({
      Number: "MY-1",
      Date: "2026-11-01T14:35:20",
      НомерВходящегоДокумента: "BANK-77",
      ДатаВходящегоДокумента: "2026-10-31T00:00:00",
      СтатьяДвиженияДенежныхСредств_Key: DDS,
    });
    const planRow = (plan.payload?.["РасшифровкаПлатежа"] as Array<Record<string, unknown>>)[0]!;
    expect(planRow).toMatchObject({ СчетНаОплату_Key: INV });
    expect(postCount(fake)).toBe(0);
  });

  it("prepare с confirm не является обходом: подтверждение без плана отклоняется", async () => {
    const { client, fake, journal } = await fixture();
    const run = await runPackage(client, target, [source], journal, { confirm: true, mode: "prepare" });
    expect(run.documents[0]?.action).toBe("blocked");
    expect(postCount(fake)).toBe(0);
  });

  it("load --confirm без заранее проверенного плана отказывает и не пишет", async () => {
    const { client, fake, journal } = await fixture();
    const run = await runPackage(client, target, [source], journal, { confirm: true });
    expect(run.documents[0]?.action).toBe("blocked");
    expect(run.documents[0]?.note).toContain("Нет проверенного плана");
    expect(postCount(fake)).toBe(0);
  });

  it("отказывает при неоднозначном контрагенте и ничего не пишет", async () => {
    const { client, fake, journal } = await fixture({
      counterparties: [cpRow, { ...cpRow, Ref_Key: CTR, Description: "Дубль" }],
    });
    const run = await runPackage(client, target, [source], journal, { confirm: false });
    expect(run.documents[0]?.action).toBe("rejected");
    expect(run.documents[0]?.note).toContain("несколько контрагентов");
    expect(postCount(fake)).toBe(0);
    expect(await journal.get(source.sourceKey)).toMatchObject({ state: "rejected" });
  });

  it("создаёт один документ, связывает sourceKey с Ref_Key и не дублирует при повторе", async () => {
    const { client, fake, journal } = await fixture();
    const first = await prepareAndConfirm(client, target, source, journal);
    expect(first.documents[0]?.action).toBe("created");
    expect(first.documents[0]?.verification).toBe("ok");
    const entry = await journal.get(source.sourceKey);
    expect(entry?.ref).toBe(first.documents[0]?.ref);
    expect(entry?.state).toBe("created");
    const posts = postCount(fake);

    const second = await runPackage(client, target, [source], journal, { confirm: true });
    expect(second.documents[0]?.action).toBe("already_loaded");
    expect(second.documents[0]?.ref).toBe(entry?.ref);
    expect(postCount(fake)).toBe(posts);
  });

  it("после потерянного ответа находит созданный документ и не делает второй POST", async () => {
    const { client, fake, journal } = await fixture({ dropPostResponse: true });
    const first = await prepareAndConfirm(client, target, source, journal);
    expect(first.documents[0]?.action).toBe("unknown");
    expect(postCount(fake)).toBe(1);
    expect(await journal.get(source.sourceKey)).toMatchObject({ state: "unknown" });

    const resumed = await runPackage(client, target, [source], journal, { confirm: true });
    expect(resumed.documents[0]?.action).toBe("reconciled");
    expect(resumed.documents[0]?.ref).toBeTruthy();
    expect(postCount(fake)).toBe(1);
    expect(await journal.get(source.sourceKey)).toMatchObject({ state: "created" });
  });

  it("брошенный transport timeout после POST не даёт новый operationId и дубль", async () => {
    const { client, fake, journal } = await fixture();
    const prepared = await runPackage(client, target, [source], journal, { confirm: false });
    const operationId = prepared.documents[0]?.operationId;
    expect(operationId).toBeTruthy();
    let threw = false;
    const flaky: ToolClient = {
      callTool: async (input) => {
        if (input.name === PAYMENT_TOOL && input.arguments?.["confirm"] === true && !threw) {
          await client.callTool(input as never);
          threw = true;
          throw new Error("MCP error -32001: Request timed out");
        }
        return client.callTool(input as never);
      },
    };
    const confirmed = await runPackage(flaky, target, [source], journal, { confirm: true });
    expect(confirmed.documents[0]?.action).toBe("unknown");
    expect(postCount(fake)).toBe(1);
    const unknownEntry = (await journal.get(source.sourceKey))!;
    expect(unknownEntry.state).toBe("unknown");
    expect(unknownEntry.operationId).toBe(operationId);

    const recovered = await runPackage(flaky, target, [source], journal, { confirm: false });
    expect(recovered.documents[0]?.action).toBe("reconciled");
    expect(recovered.documents[0]?.operationId).toBe(operationId);
    expect(postCount(fake)).toBe(1);
    expect((await journal.get(source.sourceKey))?.state).toBe("created");
  });

  it("потерянный POST: отличие только во времени документа ловится сверкой", async () => {
    const { client, fake, journal } = await fixture({ dropPostResponse: true });
    const first = await prepareAndConfirm(client, target, source, journal);
    expect(first.documents[0]?.action).toBe("unknown");
    const original = fake.conn.client.getEntity.bind(fake.conn.client);
    vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
      const actual = await original(path);
      if (path.startsWith("Document_ПоступлениеНаРасчетныйСчет(") && !path.includes("$select"))
        return { ...actual, Date: "2026-11-01T00:00:00" };
      return actual;
    });
    const resumed = await runPackage(client, target, [source], journal, {
      confirm: true,
      mode: "resume",
    });
    expect(resumed.documents[0]?.action).toBe("reconciled");
    expect(resumed.documents[0]?.verification).toBe("failed");
    expect(JSON.stringify(resumed.documents[0]?.mismatches)).toContain('"Date"');
    expect(postCount(fake)).toBe(1);
  });

  it("расхождение контрольного чтения по НДС: документ не дублируется, но помечается failed", async () => {
    const { client, fake, journal } = await fixture();
    const original = fake.conn.client.getEntity.bind(fake.conn.client);
    vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
      const actual = await original(path);
      if (path.startsWith("Document_ПоступлениеНаРасчетныйСчет(") && !path.includes("$select"))
        return {
          ...actual,
          РасшифровкаПлатежа: [
            {
              ...(actual["РасшифровкаПлатежа"] as Array<Record<string, unknown>>)[0],
              СтавкаНДС: "БезНДС",
              СуммаНДС: 0,
            },
          ],
        };
      return actual;
    });
    const prepared = await runPackage(client, target, [source], journal, { confirm: false });
    expect(prepared.documents[0]?.action).toBe("dry_run");
    const first = await runPackage(client, target, [source], journal, { confirm: true });
    expect(first.documents[0]?.action).toBe("created");
    expect(first.documents[0]?.verification).toBe("failed");
    expect(first.documents[0]?.mismatches).toBeTruthy();
    const posts = postCount(fake);
    const second = await runPackage(client, target, [source], journal, { confirm: true });
    expect(second.documents[0]?.action).toBe("already_loaded");
    expect(second.documents[0]?.verification).toBe("failed");
    expect(postCount(fake)).toBe(posts);
  });

  it("подтверждение требует совпадения плана с сохранённым отпечатком", async () => {
    const contracts = [{ ...contractRow }];
    const { client, fake, journal, state } = await fixture({ contracts });
    const prepared = await runPackage(client, target, [source], journal, { confirm: false });
    expect(prepared.documents[0]?.action).toBe("dry_run");
    // План изменился без изменения sourceKey: 1С вернёт другой Ref_Key договора.
    state.contracts[0]!.Ref_Key = "99999999-9999-4999-8999-999999999999";
    const confirmed = await runPackage(client, target, [source], journal, { confirm: true });
    expect(confirmed.documents[0]?.action).toBe("blocked");
    expect(confirmed.documents[0]?.note).toContain("не совпадает");
    expect(postCount(fake)).toBe(0);
  });

  it("смена цели после prepare сохраняет конфликт и не подтверждает старый план", async () => {
    const { client, fake, journal } = await fixture();
    await runPackage(client, target, [source], journal, { confirm: false });
    const otherTarget = { ...target, bankAccountNumber: "40702810000000000002" };
    const changed = await runPackage(client, otherTarget, [source], journal, { confirm: true });
    expect(changed.documents[0]?.action).toBe("blocked");
    expect(postCount(fake)).toBe(0);
  });

  it("resume без --confirm только сверяет отправленное и не подтверждает prepared", async () => {
    const { client, fake, journal } = await fixture();
    await runPackage(client, target, [source], journal, { confirm: false });
    const resumed = await runPackage(client, target, [source], journal, {
      confirm: false,
      mode: "resume",
    });
    expect(resumed.documents[0]?.action).toBe("prepared");
    expect(postCount(fake)).toBe(0);
    const extra = { ...source, sourceKey: "statement-new-row" };
    const skipped = await runPackage(client, target, [extra], journal, { confirm: false, mode: "resume" });
    expect(skipped.documents[0]?.action).toBe("skipped");
    expect(postCount(fake)).toBe(0);
    const confirmed = await runPackage(client, target, [source], journal, {
      confirm: true,
      mode: "resume",
    });
    expect(confirmed.documents[0]?.action).toBe("created");
    expect(postCount(fake)).toBe(1);
  });

  it("найденный после неизвестного POST документ сверяется по полям, а не только по Ref_Key", async () => {
    const { client, fake, journal } = await fixture({ dropPostResponse: true });
    const first = await prepareAndConfirm(client, target, source, journal);
    expect(first.documents[0]?.action).toBe("unknown");
    const original = fake.conn.client.getEntity.bind(fake.conn.client);
    vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
      const actual = await original(path);
      if (path.startsWith("Document_ПоступлениеНаРасчетныйСчет(") && !path.includes("$select")) {
        const row = (actual["РасшифровкаПлатежа"] as Array<Record<string, unknown>>)[0]!;
        return {
          ...actual,
          Date: "2020-01-01T00:00:00",
          Posted: true,
          РасшифровкаПлатежа: [
            {
              ...row,
              СуммаНДС: 0,
              СчетУчетаРасчетовСКонтрагентом_Key: "00000000-0000-0000-0000-000000000001",
            },
            { ...row, LineNumber: 2 },
          ],
        };
      }
      return actual;
    });
    const resumed = await runPackage(client, target, [source], journal, { confirm: true, mode: "resume" });
    expect(resumed.documents[0]?.action).toBe("reconciled");
    expect(resumed.documents[0]?.verification).toBe("failed");
    const mismatches = JSON.stringify(resumed.documents[0]?.mismatches);
    expect(mismatches).toContain("Date");
    expect(mismatches).toContain("Posted");
    expect(mismatches).toContain("РасшифровкаПлатежа");
    expect(mismatches).toContain("СчетУчетаРасчетовСКонтрагентом_Key");
    expect(postCount(fake)).toBe(1);
  });

  it("уже загруженный sourceKey с изменённым источником остаётся конфликтом", async () => {
    const { client, fake, journal } = await fixture();
    const first = await prepareAndConfirm(client, target, source, journal);
    expect(first.documents[0]?.action).toBe("created");
    const changed = { ...source, amount: 999 };
    const again = await runPackage(client, target, [changed], journal, { confirm: true });
    expect(again.documents[0]?.action).toBe("blocked");
    expect(again.documents[0]?.ref).toBe(first.documents[0]?.ref);
    expect(postCount(fake)).toBe(1);
  });

  it.each([
    ["number", "HIST-1"],
    ["time", "09:10:11"],
    ["cashFlowItemRef", DDS],
    ["customerInvoiceRef", INV],
  ])("добавление исторического поля %s к уже загруженному ключу — конфликт", async (field, value) => {
    const legacySource = {
      ...source,
      number: undefined,
      time: undefined,
      cashFlowItemRef: undefined,
      customerInvoiceRef: undefined,
    };
    const { client, fake, journal } = await fixture();
    const first = await prepareAndConfirm(client, target, legacySource, journal);
    expect(first.documents[0]?.action).toBe("created");
    const entry = (await journal.get(source.sourceKey))!;
    await journal.put({ ...entry, sourceHash: legacySourceHash(legacySource, target) });
    const changed = { ...legacySource, [field]: value } as SourceDocument;
    const again = await runPackage(client, target, [changed], journal, { confirm: true });
    expect(again.documents[0]?.action).toBe("blocked");
    expect(postCount(fake)).toBe(1);
  });

  it("второй пакет с тем же журналом получает отказ, а не теряет строки", async () => {
    const { client, fake, journal } = await fixture();
    const second = new BankLoadJournal(journal.directory);
    await journal.acquire();
    await expect(
      runPackage(client, target, [{ ...source, sourceKey: "second-package" }], second, { confirm: false }),
    ).rejects.toThrow("уже используется");
    expect(postCount(fake)).toBe(0);
    await journal.release();
  });

  it("существующий lock не снимается автоматически; восстановление — ручное", async () => {
    const dir = await mkdtemp(join(tmpdir(), "bank-loader-stale-"));
    cleanup.push(() => rm(dir, { recursive: true, force: true }));
    await writeFile(
      join(dir, "journal.lock"),
      JSON.stringify({ pid: 9999999, host: "other-host", startedAt: new Date(0).toISOString() }),
    );
    const journal = new BankLoadJournal(dir);
    await expect(journal.acquire()).rejects.toThrow("вручную");
    await rm(join(dir, "journal.lock"), { force: true });
    await journal.acquire();
    await journal.release();
  });

  it("release не удаляет чужой lock", async () => {
    const dir = await mkdtemp(join(tmpdir(), "bank-loader-foreign-lock-"));
    cleanup.push(() => rm(dir, { recursive: true, force: true }));
    const owner = new BankLoadJournal(dir);
    const contender = new BankLoadJournal(dir);
    await owner.acquire();
    await expect(contender.acquire()).rejects.toThrow("уже используется");
    await contender.release();
    await expect(new BankLoadJournal(dir).acquire()).rejects.toThrow("уже используется");
    await owner.release();
  });

  it("журнал предыдущей версии не считается изменённым источником", async () => {
    const legacySource = {
      ...source,
      number: undefined,
      time: undefined,
      cashFlowItemRef: undefined,
      customerInvoiceRef: undefined,
    };
    const { client, fake, journal } = await fixture();
    const first = await prepareAndConfirm(client, target, legacySource, journal);
    expect(first.documents[0]?.action).toBe("created");
    const entry = (await journal.get(source.sourceKey))!;
    await journal.put({ ...entry, sourceHash: legacySourceHash(legacySource, target) });
    const again = await runPackage(client, target, [legacySource], journal, { confirm: true });
    expect(again.documents[0]?.action).toBe("already_loaded");
    expect(postCount(fake)).toBe(1);
  });

  it("журнал предыдущей версии с неизвестным POST всё равно сверяется", async () => {
    const legacySource = {
      ...source,
      number: undefined,
      time: undefined,
      cashFlowItemRef: undefined,
      customerInvoiceRef: undefined,
    };
    const { client, fake, journal } = await fixture({ dropPostResponse: true });
    const first = await prepareAndConfirm(client, target, legacySource, journal);
    expect(first.documents[0]?.action).toBe("unknown");
    const entry = (await journal.get(source.sourceKey))!;
    await journal.put({ ...entry, sourceHash: legacySourceHash(legacySource, target) });
    const resumed = await runPackage(client, target, [legacySource], journal, {
      confirm: true,
      mode: "resume",
    });
    expect(resumed.documents[0]?.action).toBe("reconciled");
    expect(postCount(fake)).toBe(1);
  });

  it("Number-фильтр 500: ограниченный проход находит конфликт номера", async () => {
    const { client, fake, journal } = await fixture({
      documents: [{ Ref_Key: DOC, Number: "MY-1", Date: "2026-05-01T00:00:00", Организация_Key: ORG }],
    });
    const original = fake.conn.client.getCollection.bind(fake.conn.client);
    vi.spyOn(fake.conn.client, "getCollection").mockImplementation(async (path: string) => {
      if (path.startsWith("Document_ПоступлениеНаРасчетныйСчет?") && path.includes("$filter"))
        throw new ODataError({ kind: "server", status: 500, message: "фильтр Number запрещён" });
      return original(path);
    });
    const run = await runPackage(client, target, [source], journal, { confirm: false });
    expect(run.documents[0]?.action).toBe("rejected");
    expect(run.documents[0]?.note).toContain("уже занят");
    expect(postCount(fake)).toBe(0);
  });

  it("Number-фильтр 500 без конфликта: план готовится по ограниченному проходу", async () => {
    const { client, fake, journal } = await fixture();
    const original = fake.conn.client.getCollection.bind(fake.conn.client);
    vi.spyOn(fake.conn.client, "getCollection").mockImplementation(async (path: string) => {
      if (path.startsWith("Document_ПоступлениеНаРасчетныйСчет?") && path.includes("$filter"))
        throw new ODataError({ kind: "server", status: 500, message: "фильтр Number запрещён" });
      return original(path);
    });
    const run = await runPackage(client, target, [source], journal, { confirm: false });
    expect(run.documents[0]?.action).toBe("dry_run");
    expect(postCount(fake)).toBe(0);
  });

  it("фильтры контрагента и договора 500: работают ограниченные запасные проходы", async () => {
    const { client, fake, journal } = await fixture();
    const original = fake.conn.client.getCollection.bind(fake.conn.client);
    vi.spyOn(fake.conn.client, "getCollection").mockImplementation(async (path: string) => {
      if (
        (path.startsWith("Catalog_Контрагенты") || path.startsWith("Catalog_ДоговорыКонтрагентов")) &&
        path.includes("$filter")
      )
        throw new ODataError({ kind: "server", status: 500, message: "фильтр запрещён" });
      return original(path);
    });
    const run = await runPackage(client, target, [source], journal, { confirm: false });
    expect(run.documents[0]?.action, JSON.stringify(run.documents[0])).toBe("dry_run");
    expect(postCount(fake)).toBe(0);
  });

  it("конфликт номера документа за тот же год блокирует подтверждение", async () => {
    const { client, fake, journal } = await fixture({
      documents: [{ Ref_Key: DOC, Number: "MY-1", Date: "2026-05-01T00:00:00", Организация_Key: ORG }],
    });
    const run = await runPackage(client, target, [source], journal, { confirm: false });
    expect(run.documents[0]?.action).toBe("rejected");
    expect(run.documents[0]?.note).toContain("уже занят");
    expect(postCount(fake)).toBe(0);
  });

  it("несуществующая статья ДДС блокирует запись до предпросмотра", async () => {
    const { client, fake, journal } = await fixture();
    const original = fake.conn.client.getEntity.bind(fake.conn.client);
    vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
      if (path.startsWith("Catalog_СтатьиДвиженияДенежныхСредств("))
        throw new ODataError({ kind: "not_found", message: "нет статьи" });
      return original(path);
    });
    const run = await runPackage(client, target, [source], journal, { confirm: false });
    expect(run.documents[0]?.action).toBe("rejected");
    expect(run.documents[0]?.note).toContain("Статья ДДС");
    expect(postCount(fake)).toBe(0);
  });

  it("несуществующий счёт покупателю блокирует запись", async () => {
    const { client, fake, journal } = await fixture();
    const original = fake.conn.client.getEntity.bind(fake.conn.client);
    vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
      if (path.startsWith("Document_СчетНаОплатуПокупателю("))
        throw new ODataError({ kind: "not_found", message: "нет счёта" });
      return original(path);
    });
    const run = await runPackage(client, target, [source], journal, { confirm: false });
    expect(run.documents[0]?.action).toBe("rejected");
    expect(run.documents[0]?.note).toContain("Счёт покупателю");
    expect(postCount(fake)).toBe(0);
  });

  it("отсутствующее в схеме поле счёта покупателю блокирует запись", async () => {
    const { client, fake, journal } = await fixture();
    const meta = await fake.conn.getMetadata();
    const rowSet = "Document_ПоступлениеНаРасчетныйСчет_РасшифровкаПлатежа";
    const row = meta.entities.get(rowSet)!;
    const entities = new Map(meta.entities);
    entities.set(rowSet, { ...row, properties: row.properties.filter((p) => p.name !== "СчетНаОплату_Key") });
    vi.spyOn(fake.conn, "getMetadata").mockResolvedValue({ ...meta, entities });
    const run = await runPackage(client, target, [source], journal, { confirm: false });
    expect(run.documents[0]?.action).toBe("rejected");
    expect(run.documents[0]?.note).toContain("СчетНаОплату_Key");
    expect(postCount(fake)).toBe(0);
  });

  it("отсутствующая в схеме статья ДДС блокирует запись", async () => {
    const { client, fake, journal } = await fixture();
    const meta = await fake.conn.getMetadata();
    const docSet = "Document_ПоступлениеНаРасчетныйСчет";
    const doc = meta.entities.get(docSet)!;
    const entities = new Map(meta.entities);
    entities.set(docSet, {
      ...doc,
      properties: doc.properties.filter((p) => p.name !== "СтатьяДвиженияДенежныхСредств_Key"),
    });
    vi.spyOn(fake.conn, "getMetadata").mockResolvedValue({ ...meta, entities });
    const run = await runPackage(client, target, [source], journal, { confirm: false });
    expect(run.documents[0]?.action).toBe("rejected");
    expect(run.documents[0]?.note).toContain("СтатьяДвиженияДенежныхСредств_Key");
    expect(postCount(fake)).toBe(0);
  });

  it("явное поле, которого нет в схеме, останавливает запись с диагностикой", async () => {
    const { client, fake, journal } = await fixture();
    const meta = await fake.conn.getMetadata();
    const rowSet = "Document_ПоступлениеНаРасчетныйСчет_РасшифровкаПлатежа";
    const row = meta.entities.get(rowSet)!;
    const entities = new Map(meta.entities);
    entities.set(rowSet, { ...row, properties: row.properties.filter((p) => p.name !== "СтавкаНДС") });
    vi.spyOn(fake.conn, "getMetadata").mockResolvedValue({ ...meta, entities });
    const run = await runPackage(client, target, [source], journal, { confirm: false });
    expect(run.documents[0]?.action).toBe("rejected");
    expect(run.documents[0]?.note).toContain("СтавкаНДС");
    expect(postCount(fake)).toBe(0);
  });

  it("явная ставка без суммы НДС отклоняется до записи", async () => {
    const { client, fake, journal } = await fixture();
    const bad = { ...source, vatAmount: undefined };
    const run = await runPackage(client, target, [bad], journal, { confirm: false });
    expect(run.documents[0]?.action).toBe("rejected");
    expect(run.documents[0]?.note).toContain("и ставку");
    expect(postCount(fake)).toBe(0);
  });
});

describe("банковский счёт организации", () => {
  it("без счёта организации и без БИК явно сообщает блокер и минимальное исправление", async () => {
    const { client } = await fixture({ bankAccounts: [] });
    const report = await ensureAccount(client, { ...target, bankBik: undefined });
    expect(report.status).toBe("blocked");
    expect(report.blocker).toContain("bankBik");
    expect(report.minimalFix).toContain("target.bankBik");
  });

  it("при HTTP 500 на фильтрах банка и валюты переходит к ограниченному проходу", async () => {
    const { client, fake } = await fixture({ bankAccounts: [] });
    const original = fake.conn.client.getCollection.bind(fake.conn.client);
    vi.spyOn(fake.conn.client, "getCollection").mockImplementation(async (path: string) => {
      if (path.includes("$filter") && (path.startsWith("Catalog_Банки") || path.startsWith("Catalog_Валюты")))
        throw new ODataError({ kind: "server", status: 500, message: "фильтр запрещён публикацией" });
      return original(path);
    });
    const preview = await ensureAccount(client, target, { confirm: false });
    expect(preview.status, JSON.stringify(preview)).toBe("preview");
    expect(fake.requests.filter((r) => r.method === "POST")).toHaveLength(0);
    const bankQueries = fake.requests.filter((r) => r.resource === "Catalog_Банки").map((r) => r.query ?? "");
    const currencyQueries = fake.requests
      .filter((r) => r.resource === "Catalog_Валюты")
      .map((r) => r.query ?? "");
    expect(bankQueries.some((q) => q !== "" && !q.includes("$filter"))).toBe(true);
    expect(currencyQueries.some((q) => q !== "" && !q.includes("$filter"))).toBe(true);
  });

  it("несколько банков с одним БИК блокируют создание счёта", async () => {
    const { client } = await fixture({
      bankAccounts: [],
      banks: [
        { Ref_Key: BANK, Code: target.bankBik, Description: "Банк 1", DeletionMark: false },
        { Ref_Key: DDS, Code: target.bankBik, Description: "Банк 2", DeletionMark: false },
      ],
    });
    const preview = await ensureAccount(client, target, { confirm: false });
    expect(preview.status).toBe("refused");
    expect(preview.note).toContain("несколько банков");
  });

  it("находит банк по СВИФТБИК, если Code пуст", async () => {
    const { client, fake } = await fixture({
      bankAccounts: [],
      banks: [
        {
          Ref_Key: BANK,
          Code: "",
          СВИФТБИК: "SABRRUMM",
          Description: "Синтетический банк",
          DeletionMark: false,
        },
      ],
    });
    const preview = await ensureAccount(client, { ...target, bankBik: "SABRRUMM" }, { confirm: false });
    expect(preview.status, JSON.stringify(preview)).toBe("preview");
    expect(fake.requests.filter((r) => r.method === "POST")).toHaveLength(0);
  });

  it("при настоящем $select запрашивает Owner_Type и всё равно находит счёт организации", async () => {
    const { client, fake } = await fixture({ respectSelect: true });
    const found = (await client.callTool({
      name: ACCOUNT_FIND_TOOL,
      arguments: { organization: ORG, accountNumber: ACCOUNT_NUMBER },
    })) as { isError?: boolean; structuredContent?: Record<string, unknown> };
    expect(found.isError, JSON.stringify(found)).toBeFalsy();
    expect(found.structuredContent).toMatchObject({ found: true, account: { ref: ACC } });
    const query = fake.requests.find((r) => r.resource === "Catalog_БанковскиеСчета")?.query ?? "";
    expect(query).toContain("Owner_Type");
  });

  it("создаёт счёт именно организации, а повторно находит его без записи", async () => {
    const { client, fake } = await fixture({ bankAccounts: [] });
    const preview = await ensureAccount(client, target, { confirm: false });
    expect(preview.status, JSON.stringify(preview)).toBe("preview");
    expect(postCount(fake)).toBe(0);

    const created = await ensureAccount(client, target, { confirm: true });
    expect(created.status, JSON.stringify(created)).toBe("created");
    expect(created.ownerType).toBe("StandardODATA.Catalog_Организации");
    expect(postCount(fake)).toBe(1);

    const again = await ensureAccount(client, target, { confirm: true });
    expect(again.status).toBe("existing");
    expect(again.ref).toBe(created.ref);
    expect(postCount(fake)).toBe(1);
  });
});
