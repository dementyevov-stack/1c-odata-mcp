import { createHash } from "node:crypto";
import type { SourceDocument, TargetRequisites } from "./types.js";
import {
  hashSource,
  legacySourceHash,
  operationIdFor,
  type BankLoadEntry,
  BankLoadJournal,
} from "./journal.js";

/**
 * Загрузчик первого участка: говорит с MCP-сервером по stdio (в тестах — по in-memory
 * транспорту), поэтому подготовка и защита от дублей — те же, что у обычного вызова
 * write.money.create_payment_from_statement. Здесь только устойчивое сопоставление
 * sourceKey → Ref_Key, журнал и возобновление после неизвестного исхода.
 */

export const PAYMENT_TOOL = "write.money.create_payment_from_statement";
export const ACCOUNT_FIND_TOOL = "read.organization.find_bank_account";
export const ACCOUNT_CREATE_TOOL = "write.organization.create_bank_account";
export const STATUS_TOOL = "write.operation.status";
const ERROR_KIND_META = "odata-1c/errorKind";

/** Минимальная часть MCP-клиента, нужная загрузчику (Client ей соответствует). */
export interface ToolClient {
  callTool(input: { name: string; arguments?: Record<string, unknown> }): Promise<unknown>;
}

interface ToolResult {
  isError?: boolean;
  content?: Array<{ type?: string; text?: string }>;
  structuredContent?: Record<string, unknown>;
  _meta?: Record<string, unknown>;
}

const asResult = (value: unknown): ToolResult => value as ToolResult;
const textOf = (result: ToolResult): string =>
  (result.content ?? [])
    .filter((c) => c.type === "text" && typeof c.text === "string")
    .map((c) => c.text)
    .join("\n");
const dataOf = (result: ToolResult): Record<string, unknown> => result.structuredContent ?? {};
const kindOf = (result: ToolResult): string | undefined => {
  const value = result._meta?.[ERROR_KIND_META];
  return typeof value === "string" ? value : undefined;
};

export interface DiagnosticReport {
  ok: boolean;
  database?: string;
  schema?: { odataVersion?: string; entityCount?: number; baseUrl?: string };
  organization?: { requestedInn: string; requestedKpp?: string; ref?: string; name?: string };
  organizations?: { count: number; scanned?: number; strategy?: string };
  account?: {
    number: string;
    found: boolean;
    ref?: string;
    ownerType?: string;
    note?: string;
  };
  problems: string[];
  hints: string[];
}

const guarded = async (
  client: ToolClient,
  name: string,
  args: Record<string, unknown>,
): Promise<ToolResult> => {
  try {
    return asResult(await client.callTool({ name, arguments: args }));
  } catch (error) {
    return {
      isError: true,
      content: [{ type: "text", text: error instanceof Error ? error.message : String(error) }],
    };
  }
};

/** Реквизиты организации по точным ИНН/КПП через MCP; ровно одна или отказ. */
export async function resolveOrganizationForTarget(
  client: ToolClient,
  target: TargetRequisites,
  database?: string,
): Promise<{ ref: string; name: string }> {
  const result = await guarded(client, "read.system.list_organizations", {
    ...(database ? { database } : {}),
    inn: target.organizationInn,
    ...(target.organizationKpp ? { kpp: target.organizationKpp } : {}),
  });
  if (result.isError)
    throw new Error(`Организация не сопоставлена: ${textOf(result) || "сервер вернул отказ без текста"}`);
  const data = dataOf(result);
  const organizations = Array.isArray(data.organizations)
    ? (data.organizations as Array<Record<string, unknown>>)
    : [];
  if (organizations.length !== 1)
    throw new Error(
      organizations.length === 0
        ? `Организация с ИНН ${target.organizationInn}${target.organizationKpp ? `, КПП ${target.organizationKpp}` : ""} не найдена.`
        : `Под ИНН ${target.organizationInn}${target.organizationKpp ? `, КПП ${target.organizationKpp}` : ""} подходит несколько организаций — нужен точный КПП.`,
    );
  return { ref: String(organizations[0]!.ref), name: String(organizations[0]!.name ?? "") };
}

/** Операторская диагностика: авторизованные GET схемы, чтение организаций, счёт организации. */
export async function diagnose(
  client: ToolClient,
  target: TargetRequisites,
  database?: string,
): Promise<DiagnosticReport> {
  const problems: string[] = [];
  const hints: string[] = [];
  const health = await guarded(client, "read.system.health_check", { ...(database ? { database } : {}) });
  if (health.isError) problems.push(`GET схемы/проверка соединения: ${textOf(health) || "отказ без текста"}`);
  const healthData = dataOf(health);
  const orgResult = await guarded(client, "read.system.list_organizations", {
    ...(database ? { database } : {}),
    inn: target.organizationInn,
    ...(target.organizationKpp ? { kpp: target.organizationKpp } : {}),
  });
  const orgData = dataOf(orgResult);
  const organizations = Array.isArray(orgData.organizations)
    ? (orgData.organizations as Array<Record<string, unknown>>)
    : [];
  let organization: DiagnosticReport["organization"];
  if (orgResult.isError) {
    problems.push(`Чтение организаций: ${textOf(orgResult) || "отказ без текста"}`);
  } else if (organizations.length !== 1) {
    problems.push(
      organizations.length === 0
        ? "Организация с заданными ИНН/КПП не найдена."
        : `Под заданные ИНН/КПП подходит ${organizations.length} организаций — нужна точная сверка.`,
    );
  } else {
    organization = {
      requestedInn: target.organizationInn,
      ...(target.organizationKpp ? { requestedKpp: target.organizationKpp } : {}),
      ref: String(organizations[0]!.ref),
      name: String(organizations[0]!.name ?? ""),
    };
  }

  let account: DiagnosticReport["account"];
  if (organization?.ref) {
    const accountResult = await guarded(client, ACCOUNT_FIND_TOOL, {
      ...(database ? { database } : {}),
      organization: organization.ref,
      accountNumber: target.bankAccountNumber,
    });
    if (accountResult.isError) {
      problems.push(`Счёт организации: ${textOf(accountResult) || "отказ без текста"}`);
    } else {
      const accountData = dataOf(accountResult);
      const found = accountData.found === true;
      account = {
        number: target.bankAccountNumber,
        found,
        ...(typeof accountData.account === "object" && accountData.account
          ? { ref: String((accountData.account as Record<string, unknown>).ref ?? "") }
          : {}),
        ...(typeof accountData.ownerType === "string" ? { ownerType: accountData.ownerType } : {}),
        ...(typeof accountData.note === "string" ? { note: accountData.note } : {}),
      };
      if (!found) {
        problems.push("У организации нет банковского счёта с заданным номером.");
        hints.push(
          target.bankBik
            ? "Выполните `ensure-account --confirm`, чтобы создать счёт именно организации."
            : "Добавьте target.bankBik в локальную конфигурацию и выполните `ensure-account --confirm`: " +
                "счёт организации создаётся только с БИК банка.",
        );
      }
    }
  }

  return {
    ok: problems.length === 0,
    ...(typeof healthData.database === "string" ? { database: healthData.database } : {}),
    ...(!health.isError
      ? {
          schema: {
            ...(typeof healthData.odataVersion === "string" ? { odataVersion: healthData.odataVersion } : {}),
            ...(typeof healthData.entityCount === "number" ? { entityCount: healthData.entityCount } : {}),
            ...(typeof healthData.baseUrl === "string" ? { baseUrl: healthData.baseUrl } : {}),
          },
        }
      : {}),
    ...(organization ? { organization } : {}),
    ...(!orgResult.isError
      ? {
          organizations: {
            count: organizations.length,
            ...(typeof orgData.scanned === "number" ? { scanned: orgData.scanned } : {}),
            ...(typeof orgData.strategy === "string" ? { strategy: orgData.strategy } : {}),
          },
        }
      : {}),
    ...(account ? { account } : {}),
    problems,
    hints,
  };
}

export interface AccountReport {
  status: "existing" | "created" | "preview" | "blocked" | "refused";
  organization: { ref: string; name: string };
  accountNumber: string;
  ownerType?: string;
  ref?: string;
  blocker?: string;
  minimalFix?: string;
  note?: string;
  operationId?: string;
}

/** Сопоставить счёт организации; при отсутствии — создать (dry-run без confirm). */
export async function ensureAccount(
  client: ToolClient,
  target: TargetRequisites,
  options: { confirm: boolean; database?: string } = { confirm: false },
): Promise<AccountReport> {
  const organization = await resolveOrganizationForTarget(client, target, options.database);
  const found = await guarded(client, ACCOUNT_FIND_TOOL, {
    ...(options.database ? { database: options.database } : {}),
    organization: organization.ref,
    accountNumber: target.bankAccountNumber,
  });
  if (found.isError) throw new Error(`Счёт организации не проверен: ${textOf(found) || "отказ без текста"}`);
  const foundData = dataOf(found);
  if (foundData.found === true) {
    const account = (foundData.account ?? {}) as Record<string, unknown>;
    return {
      status: "existing",
      organization,
      accountNumber: target.bankAccountNumber,
      ownerType: typeof foundData.ownerType === "string" ? foundData.ownerType : undefined,
      ref: typeof account.ref === "string" ? account.ref : undefined,
      note: "Счёт организации уже есть; запись не выполнялась.",
    };
  }
  if (!target.bankBik)
    return {
      status: "blocked",
      organization,
      accountNumber: target.bankAccountNumber,
      blocker:
        "У организации нет этого счёта, а target.bankBik не задан. Без БИК создать счёт организации нельзя; " +
        "счёт контрагента вместо него не подставляется.",
      minimalFix:
        "Добавьте в локальную конфигурацию target.bankBik (9 цифр) и выполните ensure-account --confirm.",
    };
  const createArgs = (confirm: boolean, operationId?: string): Record<string, unknown> => ({
    ...(options.database ? { database: options.database } : {}),
    organizationRef: organization.ref,
    accountNumber: target.bankAccountNumber,
    bik: target.bankBik!,
    confirm,
    ...(operationId ? { operationId } : {}),
  });
  // Сначала всегда предпросмотр: инструмент сам сопоставит существующий счёт и вернёт operationId.
  const preview = await guarded(client, ACCOUNT_CREATE_TOOL, createArgs(false));
  const previewData = dataOf(preview);
  if (preview.isError)
    return {
      status: "refused",
      organization,
      accountNumber: target.bankAccountNumber,
      note: textOf(preview) || "Создание счёта организации отклонено.",
      ...(typeof previewData.operationId === "string" ? { operationId: previewData.operationId } : {}),
    };
  if (previewData.existing === true)
    return {
      status: "existing",
      organization,
      accountNumber: target.bankAccountNumber,
      ...(typeof previewData.ownerType === "string" ? { ownerType: previewData.ownerType } : {}),
      ...(typeof previewData.ref === "string" ? { ref: previewData.ref } : {}),
      note: "Счёт организации уже есть; запись не выполнялась.",
    };
  if (!options.confirm)
    return {
      status: "preview",
      organization,
      accountNumber: target.bankAccountNumber,
      ...(typeof previewData.ownerType === "string" ? { ownerType: previewData.ownerType } : {}),
      ...(typeof previewData.operationId === "string" ? { operationId: previewData.operationId } : {}),
      note:
        "Счёт организации будет создан с Owner = организация и Owner_Type организации. " +
        "Повторите с confirm, чтобы записать.",
    };
  if (typeof previewData.operationId !== "string")
    return {
      status: "refused",
      organization,
      accountNumber: target.bankAccountNumber,
      note: "Предпросмотр не вернул operationId; подтверждение невозможно.",
    };
  const result = await guarded(client, ACCOUNT_CREATE_TOOL, createArgs(true, previewData.operationId));
  const data = dataOf(result);
  if (result.isError)
    return {
      status: "refused",
      organization,
      accountNumber: target.bankAccountNumber,
      note: textOf(result) || "Создание счёта организации отклонено.",
      ...(typeof data.operationId === "string" ? { operationId: data.operationId } : {}),
    };
  if (data.existing === true)
    return {
      status: "existing",
      organization,
      accountNumber: target.bankAccountNumber,
      ...(typeof data.ownerType === "string" ? { ownerType: data.ownerType } : {}),
      ...(typeof data.ref === "string" ? { ref: data.ref } : {}),
      note: "Счёт организации уже есть; запись не выполнялась.",
    };
  if (data.created === true)
    return {
      status: "created",
      organization,
      accountNumber: target.bankAccountNumber,
      ...(typeof data.ownerType === "string" ? { ownerType: data.ownerType } : {}),
      ...(typeof data.ref === "string" ? { ref: data.ref } : {}),
      ...(typeof data.operationId === "string" ? { operationId: data.operationId } : {}),
      note: "Создан счёт организации; владелец проверен контрольным чтением.",
    };
  return {
    status: "refused",
    organization,
    accountNumber: target.bankAccountNumber,
    note: "Неожиданный ответ на создание счёта организации.",
  };
}

export type DocumentAction =
  | "dry_run"
  | "created"
  | "already_loaded"
  | "reconciled"
  | "rejected"
  | "unknown"
  | "blocked"
  | "failed"
  | "skipped"
  | "prepared";

export interface DocumentReport {
  sourceKey: string;
  action: DocumentAction;
  operationId?: string;
  ref?: string;
  number?: string;
  verification?: "ok" | "failed" | "journal_receipt";
  note?: string;
  errorKind?: string;
  plan?: Record<string, unknown>;
  mismatches?: unknown;
}

export interface RunReport {
  mode: "dry-run" | "confirm" | "resume";
  database?: string;
  summary: {
    total: number;
    dryRun: number;
    created: number;
    alreadyLoaded: number;
    reconciled: number;
    rejected: number;
    unknown: number;
    blocked: number;
    failed: number;
    skipped: number;
    prepared: number;
    verificationFailed: number;
  };
  documents: DocumentReport[];
}

const DEFINITIVE = new Set([
  "input",
  "refused",
  "auth",
  "bad_request",
  "not_found",
  "not_published",
  "busy",
  "overflow",
]);
const UNKNOWN_AFTER_SEND = new Set([
  "unknown",
  "timeout",
  "network",
  "server",
  "parse",
  "rate_limit",
  "internal",
]);
const looksUnknown = (text: string): boolean =>
  /Результат записи.*неизвестен|Повторная отправка заблокирована|ответ потерян/i.test(text);

const bounded = (text: string): string => (text.length > 500 ? `${text.slice(0, 500)}…` : text);

const argsFor = (
  source: SourceDocument,
  target: TargetRequisites,
  database: string | undefined,
  operationId: string,
  confirm: boolean,
): Record<string, unknown> => ({
  ...(database ? { database } : {}),
  sourceKey: source.sourceKey,
  organizationInn: target.organizationInn,
  ...(target.organizationKpp ? { organizationKpp: target.organizationKpp } : {}),
  orgBankAccountNumber: target.bankAccountNumber,
  payerInn: source.payerInn,
  ...(source.payerKpp ? { payerKpp: source.payerKpp } : {}),
  ...(source.payerName ? { payerName: source.payerName } : {}),
  ...(source.payerAccountNumber ? { payerAccountNumber: source.payerAccountNumber } : {}),
  amount: source.amount,
  date: source.date,
  ...(source.number ? { number: source.number } : {}),
  ...(source.time ? { time: source.time } : {}),
  ...(source.purpose ? { purpose: source.purpose } : {}),
  ...(source.documentNumber ? { documentNumber: source.documentNumber } : {}),
  ...(source.documentDate ? { documentDate: source.documentDate } : {}),
  ...(source.vatRate ? { vatRate: source.vatRate } : {}),
  ...(source.vatAmount !== undefined ? { vatAmount: source.vatAmount } : {}),
  ...(source.cashFlowItemRef ? { cashFlowItemRef: source.cashFlowItemRef } : {}),
  ...(source.customerInvoiceRef ? { customerInvoiceRef: source.customerInvoiceRef } : {}),
  ...(source.contractNumber ? { contractNumber: source.contractNumber } : {}),
  ...(source.contractRef ? { contractRef: source.contractRef } : {}),
  confirm,
  operationId,
});

const entryFor = (
  source: SourceDocument,
  sourceHash: string,
  operationId: string,
  attempt: number,
  state: BankLoadEntry["state"],
  extra: Partial<BankLoadEntry> = {},
): BankLoadEntry => ({
  version: 1,
  sourceKey: source.sourceKey,
  sourceHash,
  operationId,
  state,
  attempt,
  updatedAt: new Date().toISOString(),
  ...extra,
});

const PAYMENT_ENTITY_SET = "Document_ПоступлениеНаРасчетныйСчет";

const canonical = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === "object")
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
        .map(([key, nested]) => [key, canonical(nested)]),
    );
  return value;
};

/** Отпечаток проверенного плана: база, ссылки и payload, а не только исходная строка. */
const hashPlan = (value: Record<string, unknown>): string =>
  createHash("sha256")
    .update(JSON.stringify(canonical(value)))
    .digest("hex");

const SIGNIFICANT_HEADER = [
  "Date",
  "Posted",
  "Number",
  "Организация_Key",
  "СчетОрганизации_Key",
  "Контрагент",
  "Контрагент_Type",
  "ДоговорКонтрагента_Key",
  "СтатьяДвиженияДенежныхСредств_Key",
  "ВидОперации",
  "СуммаДокумента",
  "НазначениеПлатежа",
  "НомерВходящегоДокумента",
  "ДатаВходящегоДокумента",
  "СчетКонтрагента_Key",
] as const;
const SIGNIFICANT_ROW = [
  "ДоговорКонтрагента_Key",
  "СчетНаОплату_Key",
  "СчетУчетаРасчетовСКонтрагентом_Key",
  "СчетУчетаРасчетовПоАвансам_Key",
  "СпособПогашенияЗадолженности",
  "СуммаПлатежа",
  "СтавкаНДС",
  "СуммаНДС",
] as const;

function expectedFromPayload(payload: Record<string, unknown> | undefined): BankLoadEntry["expected"] {
  if (!payload) return undefined;
  const header = Object.fromEntries(
    SIGNIFICANT_HEADER.filter((field) => payload[field] !== undefined).map((field) => [
      field,
      payload[field],
    ]),
  );
  const row = (payload["РасшифровкаПлатежа"] as Array<Record<string, unknown>> | undefined)?.[0];
  if (!row) return { header, row: {} };
  return {
    header,
    row: Object.fromEntries(
      SIGNIFICANT_ROW.filter((field) => row[field] !== undefined).map((field) => [field, row[field]]),
    ),
  };
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
/** Допустимый Edm.DateTime 1С: дата, необязательное время, необязательные доли секунды. */
const DATE_RE = /^(\d{4}-\d{2}-\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?$/;

function dateParts(value: unknown): { date: string; time?: string } | undefined {
  const m = DATE_RE.exec(String(value ?? "").trim());
  if (!m) return undefined;
  return {
    date: m[1]!,
    ...(m[2] !== undefined ? { time: `${m[2]}:${m[3]}:${m[4] ?? "00"}` } : {}),
  };
}

/** Сравнение сохранённого значения с планом без типа Edm: GUID, даты, числа, строки. */
function samePlannedValue(expected: unknown, actual: unknown): boolean {
  if (expected === undefined) return true;
  if (typeof expected === "number") return Math.abs(Number(actual) - expected) < 1e-6;
  if (typeof expected === "boolean") return actual === expected;
  if (expected === null) return actual === null || actual === undefined || actual === "";
  const e = String(expected).trim();
  const a = String(actual ?? "").trim();
  if (!e) return !a;
  if (GUID_RE.test(e)) return a.toLowerCase() === e.toLowerCase();
  const expectedDate = dateParts(e);
  if (expectedDate) {
    const actualDate = dateParts(a);
    if (!actualDate) return false;
    // Date-only вход сверяем по дню; полный Date — по дате и точному времени.
    if (!expectedDate.time) return actualDate.date === expectedDate.date;
    return actualDate.date === expectedDate.date && (actualDate.time ?? "00:00:00") === expectedDate.time;
  }
  return a === e;
}

function verifyDocument(
  actual: Record<string, unknown> | undefined,
  expected: BankLoadEntry["expected"],
): { ok: boolean; mismatches: Array<{ field: string; expected: unknown; actual: unknown }> } {
  if (!expected || !actual) return { ok: false, mismatches: [] };
  const mismatches: Array<{ field: string; expected: unknown; actual: unknown }> = [];
  for (const [field, value] of Object.entries(expected.header))
    if (!samePlannedValue(value, actual[field]))
      mismatches.push({ field, expected: value, actual: actual[field] });
  const rows = actual["РасшифровкаПлатежа"];
  if (!Array.isArray(rows) || rows.length !== 1)
    mismatches.push({
      field: "РасшифровкаПлатежа",
      expected: "ровно одна строка",
      actual: Array.isArray(rows) ? `строк: ${rows.length}` : "нет табличной части",
    });
  const row = Array.isArray(rows) ? (rows[0] as Record<string, unknown> | undefined) : undefined;
  for (const [field, value] of Object.entries(expected.row))
    if (!samePlannedValue(value, row?.[field]))
      mismatches.push({ field: `РасшифровкаПлатежа.${field}`, expected: value, actual: row?.[field] });
  return { ok: mismatches.length === 0, mismatches };
}

/** Контрольное чтение после неизвестного POST: мало найти Ref_Key — поля и ТЧ должны совпасть с планом. */
async function verifyFoundDocument(
  client: ToolClient,
  database: string | undefined,
  entry: BankLoadEntry,
  ref: string,
): Promise<{ ok: boolean; mismatches?: unknown; note: string }> {
  const entitySet = entry.entitySet ?? PAYMENT_ENTITY_SET;
  if (!entry.expected)
    return {
      ok: false,
      note: "Документ найден, но план для сверки полей не сохранён — факт существования недостаточен. Проверьте в 1С вручную.",
    };
  const read = await guarded(client, "read.document.get_document", {
    ...(database ? { database } : {}),
    entitySet,
    ref,
  });
  if (read.isError)
    return { ok: false, note: `Документ найден, но перечитать его не удалось: ${bounded(textOf(read))}.` };
  const actual = dataOf(read);
  const checked = verifyDocument(actual, entry.expected);
  return checked.ok
    ? { ok: true, note: "Найденный документ сверен с планом: значимые поля и расшифровка платежа совпали." }
    : {
        ok: false,
        mismatches: checked.mismatches,
        note: `Документ найден, но поля не совпали с планом: ${checked.mismatches
          .map((m) => m.field)
          .join(", ")}.`,
      };
}

export type RunMode = "prepare" | "confirm" | "resume";

/**
 * Обработать одну строку. Возвращает отчёт и обновляет журнал; в 1С пишет только при
 * confirm=true и только после успешного предпросмотра.
 *
 * mode=resume без confirm — только сверка ранее отправленного: новые строки пропускаются,
 * подготовленные не подтверждаются. mode=confirm требует совпадения текущего плана с сохранённым
 * отпечатком (база, ссылки, payload).
 */
async function runOne(
  client: ToolClient,
  target: TargetRequisites,
  source: SourceDocument,
  journal: BankLoadJournal,
  options: { confirm: boolean; mode: RunMode; database?: string },
): Promise<DocumentReport> {
  const sourceHash = hashSource(source, target);
  const legacyHash = legacySourceHash(source, target);
  const existing = await journal.get(source.sourceKey);
  const hashMatches = existing?.sourceHash === sourceHash || existing?.sourceHash === legacyHash;
  const changed = existing !== undefined && !hashMatches;
  const mode = options.mode;

  if (mode === "resume" && !existing)
    return {
      sourceKey: source.sourceKey,
      action: "skipped",
      note: "resume не создаёт новые строки: сначала выполните prepare и подтвердите загрузку.",
    };

  if (changed && existing && ["created", "unknown", "submitted"].includes(existing.state))
    return {
      sourceKey: source.sourceKey,
      action: "blocked",
      operationId: existing.operationId,
      ...(existing.ref ? { ref: existing.ref } : {}),
      note:
        "sourceKey уже связан с записью в 1С, а данные источника или цель изменились. " +
        "Конфликт сохранён: не переиспользуйте ключ; сверьте документ в 1С вручную.",
    };

  if (existing?.state === "created")
    return {
      sourceKey: source.sourceKey,
      action: "already_loaded",
      operationId: existing.operationId,
      ...(existing.ref ? { ref: existing.ref } : {}),
      ...(existing.number ? { number: existing.number } : {}),
      ...(existing.verification ? { verification: existing.verification } : {}),
      note:
        existing.verification === "failed"
          ? "Строка уже загружена, но контрольное чтение ранее не подтвердило поля. Повторный POST не выполнялся."
          : "Строка уже загружена; повторный POST не выполнялся.",
    };

  const conflict = (): DocumentReport => ({
    sourceKey: source.sourceKey,
    action: "blocked",
    operationId: existing?.operationId,
    ...(existing?.ref ? { ref: existing.ref } : {}),
    note:
      "Источник или целевые реквизиты изменились после проверенного плана. " +
      "Подтверждение не отправлено; выполните prepare и проверьте новый план заново.",
  });

  if (mode !== "prepare" && changed && existing?.state === "prepared") return conflict();

  if (mode === "resume" && existing?.state === "prepared" && !options.confirm)
    return {
      sourceKey: source.sourceKey,
      action: "prepared",
      operationId: existing.operationId,
      note:
        "Есть подготовленный план. resume без --confirm не подтверждает его: " +
        "проверьте план и выполните load/resume с --confirm.",
    };

  if (existing?.state === "rejected") {
    if (mode !== "prepare")
      return {
        sourceKey: source.sourceKey,
        action: "rejected",
        operationId: existing.operationId,
        errorKind: existing.errorKind,
        note:
          "Операция ранее отклонена. Причина не исправлена автоматически: выполните prepare, " +
          "проверьте новый план и только затем подтверждайте.",
      };
  }

  let preparedState = existing;
  if (existing && (existing.state === "unknown" || existing.state === "submitted")) {
    const status = await guarded(client, STATUS_TOOL, {
      ...(options.database ? { database: options.database } : {}),
      operationId: existing.operationId,
    });
    if (status.isError)
      return {
        sourceKey: source.sourceKey,
        action: "blocked",
        operationId: existing.operationId,
        note: `Исход неизвестен; сверка статуса не удалась: ${bounded(textOf(status) || "отказ без текста")}. Повтор POST запрещён.`,
      };
    const statusData = dataOf(status);
    const statusName = String(statusData.status ?? "");
    if (
      (statusName === "succeeded" || statusName === "found_reconciled") &&
      typeof statusData.ref === "string"
    ) {
      const verified = await verifyFoundDocument(client, options.database, existing, statusData.ref);
      const note = `${verified.note} Повторный POST не выполнялся.`;
      await journal.put(
        entryFor(source, sourceHash, existing.operationId, existing.attempt, "created", {
          ref: statusData.ref,
          ...(typeof statusData.number === "string" ? { number: statusData.number } : {}),
          verification: verified.ok ? "ok" : "failed",
          verificationNote: verified.ok ? undefined : note,
          ...(existing.planHash ? { planHash: existing.planHash } : {}),
          ...(existing.planDatabase ? { planDatabase: existing.planDatabase } : {}),
          ...(existing.entitySet ? { entitySet: existing.entitySet } : {}),
          ...(existing.expected ? { expected: existing.expected } : {}),
        }),
      );
      return {
        sourceKey: source.sourceKey,
        action: "reconciled",
        operationId: existing.operationId,
        ref: statusData.ref,
        ...(typeof statusData.number === "string" ? { number: statusData.number } : {}),
        verification: verified.ok ? "ok" : "failed",
        ...(verified.mismatches ? { mismatches: verified.mismatches } : {}),
        note,
      };
    }
    if (statusName === "prepared") {
      if (mode === "resume" && !options.confirm)
        return {
          sourceKey: source.sourceKey,
          action: "prepared",
          operationId: existing.operationId,
          note: "Подтверждение не отправлялось; resume без --confirm его не повторяет.",
        };
      preparedState = { ...existing, state: "prepared" };
    } else if (statusName === "rejected") {
      await journal.put(
        entryFor(source, sourceHash, existing.operationId, existing.attempt, "rejected", {
          errorKind: "rejected",
          errorMessage: bounded(textOf(status) || "1С отклонила запись"),
        }),
      );
      return {
        sourceKey: source.sourceKey,
        action: "rejected",
        operationId: existing.operationId,
        note: "Операция отклонена; документ не создан. Выполните prepare для нового плана.",
      };
    } else if (statusName === "not_in_journal") {
      return {
        sourceKey: source.sourceKey,
        action: "blocked",
        operationId: existing.operationId,
        note:
          "Локальный журнал MCP не знает операцию, поэтому исход нельзя сверить. Повтор POST запрещён; " +
          "проверьте документ в 1С вручную.",
      };
    } else {
      return {
        sourceKey: source.sourceKey,
        action: "blocked",
        operationId: existing.operationId,
        note:
          `Исход записи неизвестен (статус ${statusName || "—"}). Повтор POST запрещён: ` +
          "сверьте документ в 1С вручную; новый sourceKey не нужен, пока операция не закрыта.",
      };
    }
  }

  if (options.confirm && preparedState?.state !== "prepared")
    return {
      sourceKey: source.sourceKey,
      action: "blocked",
      note:
        "Нет проверенного плана: подтверждение возможно только для строки, для которой ранее " +
        "выполнен prepare. В 1С ничего не записано.",
    };

  const reuse =
    preparedState?.state === "prepared" &&
    (preparedState.sourceHash === sourceHash || preparedState.sourceHash === legacyHash) &&
    preparedState.operationId !== undefined;
  const attempt = reuse && preparedState ? preparedState.attempt : (existing?.attempt ?? 0) + 1;
  const operationId =
    reuse && preparedState
      ? preparedState.operationId
      : operationIdFor(source.sourceKey, sourceHash, attempt);

  const preview = await guarded(
    client,
    PAYMENT_TOOL,
    argsFor(source, target, options.database, operationId, false),
  );
  if (preview.isError) {
    if (mode !== "prepare" && preparedState?.state === "prepared") {
      const text = bounded(textOf(preview) || "отказ без текста");
      return {
        sourceKey: source.sourceKey,
        action: "blocked",
        operationId,
        note: /не совпадает|не совпадают/i.test(text)
          ? `Текущий план не совпадает с проверенным отпечатком: ${text}. Подтверждение не отправлено; выполните prepare.`
          : `Текущий план не подготовлен: ${text}. Подтверждение не отправлено.`,
      };
    }
    await journal.put(
      entryFor(source, sourceHash, operationId, attempt, "rejected", {
        errorKind: kindOf(preview) ?? "refused",
        errorMessage: bounded(textOf(preview) || "предпросмотр отклонён"),
      }),
    );
    return {
      sourceKey: source.sourceKey,
      action: "rejected",
      operationId,
      errorKind: kindOf(preview) ?? "refused",
      note: bounded(textOf(preview) || "Предпросмотр отклонён; в 1С ничего не записано."),
    };
  }
  const previewData = dataOf(preview);
  const plan = {
    ...(previewData.database ? { database: previewData.database } : {}),
    ...(previewData.resolved ? { resolved: previewData.resolved } : {}),
    ...(previewData.payload ? { payload: previewData.payload } : {}),
    ...(previewData.verificationPlan ? { verificationPlan: previewData.verificationPlan } : {}),
    ...(previewData.notes ? { notes: previewData.notes } : {}),
  };
  const payload = previewData.payload as Record<string, unknown> | undefined;
  const planHash = hashPlan({
    database: previewData.database ?? options.database ?? "",
    resolved: previewData.resolved ?? null,
    payload: payload ?? null,
  });
  const expected = expectedFromPayload(payload);
  const entitySet = typeof previewData.willCreate === "string" ? previewData.willCreate : PAYMENT_ENTITY_SET;

  if (preparedState?.state === "prepared" && (mode === "confirm" || (mode === "resume" && options.confirm))) {
    if (!preparedState.planHash)
      return {
        sourceKey: source.sourceKey,
        action: "blocked",
        operationId,
        note: "Нет сохранённого отпечатка проверенного плана. Выполните prepare заново.",
      };
    if (preparedState.planHash !== planHash)
      return {
        sourceKey: source.sourceKey,
        action: "blocked",
        operationId,
        note:
          "Текущий план (база, ссылки или payload) не совпадает с проверенным ранее отпечатком. " +
          "Подтверждение не отправлено; выполните prepare и проверьте изменения.",
      };
  }

  await journal.put(
    entryFor(source, sourceHash, operationId, attempt, "prepared", {
      planHash,
      planDatabase: String(previewData.database ?? options.database ?? ""),
      entitySet,
      ...(expected ? { expected } : {}),
    }),
  );
  if (!options.confirm || (mode === "resume" && !options.confirm))
    return {
      sourceKey: source.sourceKey,
      action: "dry_run",
      operationId,
      plan: { ...plan, planHash },
      note: "Предпросмотр; в 1С ничего не записано.",
    };

  await journal.put(
    entryFor(source, sourceHash, operationId, attempt, "submitted", {
      planHash,
      planDatabase: String(previewData.database ?? options.database ?? ""),
      entitySet,
      ...(expected ? { expected } : {}),
    }),
  );
  const created = await guarded(
    client,
    PAYMENT_TOOL,
    argsFor(source, target, options.database, operationId, true),
  );
  const createdData = dataOf(created);
  const planFields = {
    planHash,
    planDatabase: String(previewData.database ?? options.database ?? ""),
    entitySet,
    ...(expected ? { expected } : {}),
  };
  if (created.isError) {
    // Документ создан, но контрольное чтение не подтвердило поля: фиксируем Ref_Key,
    // чтобы повтор не создал дубль, и отдаём оператору расхождения.
    if (createdData.created === true && typeof createdData.ref === "string") {
      const note = bounded(textOf(created) || "Контрольное чтение не подтвердило поля.");
      await journal.put(
        entryFor(source, sourceHash, operationId, attempt, "created", {
          ref: createdData.ref,
          verification: "failed",
          verificationNote: note,
          errorKind: kindOf(created) ?? "refused",
          errorMessage: note,
          ...planFields,
        }),
      );
      return {
        sourceKey: source.sourceKey,
        action: "created",
        operationId,
        ref: createdData.ref,
        verification: "failed",
        note,
        ...(createdData.mismatches ? { mismatches: createdData.mismatches } : {}),
      };
    }
    const kind = kindOf(created) ?? "";
    const definitive =
      DEFINITIVE.has(kind) || (!UNKNOWN_AFTER_SEND.has(kind) && !looksUnknown(textOf(created)) && !kind);
    const note = bounded(textOf(created) || "Запись отклонена.");
    await journal.put(
      entryFor(source, sourceHash, operationId, attempt, definitive ? "rejected" : "unknown", {
        errorKind: kind || (definitive ? "refused" : "unknown"),
        errorMessage: note,
        ...planFields,
      }),
    );
    return {
      sourceKey: source.sourceKey,
      action: definitive ? "rejected" : "unknown",
      operationId,
      errorKind: kind || undefined,
      note: definitive
        ? note
        : `${note} Исход неизвестен: повтор POST запрещён, при следующем запуске будет сверка write.operation.status.`,
    };
  }

  const ref = typeof createdData.ref === "string" ? createdData.ref : "";
  if (!ref) {
    await journal.put(
      entryFor(source, sourceHash, operationId, attempt, "unknown", {
        errorKind: "unknown",
        errorMessage: "1С приняла запрос, но ответ не содержит Ref_Key.",
        ...planFields,
      }),
    );
    return {
      sourceKey: source.sourceKey,
      action: "unknown",
      operationId,
      note: "1С приняла запрос, но ответ не содержит Ref_Key. Повтор POST запрещён; нужна ручная сверка.",
    };
  }
  const verification =
    createdData.verification === "ok"
      ? "ok"
      : createdData.verification === "failed"
        ? "failed"
        : createdData.replayed === true
          ? "journal_receipt"
          : "ok";
  await journal.put(
    entryFor(source, sourceHash, operationId, attempt, "created", {
      ref,
      ...(typeof createdData.number === "string" ? { number: createdData.number } : {}),
      verification,
      ...planFields,
    }),
  );
  return {
    sourceKey: source.sourceKey,
    action: "created",
    operationId,
    ref,
    ...(typeof createdData.number === "string" ? { number: createdData.number } : {}),
    verification,
    note:
      verification === "journal_receipt"
        ? "Операция уже была выполнена ранее; документ возвращён из журнала, повторный POST не отправлялся."
        : "Документ создан и проверен контрольным чтением.",
  };
}

/** Прогон пакета: dry-run по умолчанию, confirm — явное подтверждение оператора. */
export async function runPackage(
  client: ToolClient,
  target: TargetRequisites,
  sources: readonly SourceDocument[],
  journal: BankLoadJournal,
  options: { confirm: boolean; database?: string; mode?: RunMode },
): Promise<RunReport> {
  const mode = options.mode ?? (options.confirm ? "confirm" : "prepare");
  const documents: DocumentReport[] = [];
  for (const source of sources) {
    try {
      documents.push(await runOne(client, target, source, journal, { ...options, mode }));
    } catch (error) {
      documents.push({
        sourceKey: source.sourceKey,
        action: "failed",
        note: bounded(error instanceof Error ? error.message : String(error)),
      });
    }
  }
  const count = (action: DocumentAction): number => documents.filter((d) => d.action === action).length;
  return {
    mode: mode === "resume" ? "resume" : options.confirm ? "confirm" : "dry-run",
    ...(options.database ? { database: options.database } : {}),
    summary: {
      total: documents.length,
      dryRun: count("dry_run"),
      created: count("created"),
      alreadyLoaded: count("already_loaded"),
      reconciled: count("reconciled"),
      rejected: count("rejected"),
      unknown: count("unknown"),
      blocked: count("blocked"),
      failed: count("failed"),
      skipped: count("skipped"),
      prepared: count("prepared"),
      verificationFailed: documents.filter((d) => d.verification === "failed").length,
    },
    documents,
  };
}

/** Есть ли в отчёте что-то, что оператор должен разобрать до следующего запуска. */
export function reportHasProblems(report: RunReport): boolean {
  return (
    report.summary.rejected > 0 ||
    report.summary.unknown > 0 ||
    report.summary.blocked > 0 ||
    report.summary.failed > 0 ||
    report.summary.verificationFailed > 0
  );
}
