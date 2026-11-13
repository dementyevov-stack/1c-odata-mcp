import { createHash } from "node:crypto";
import { mkdir, open, readFile, rename, rm } from "node:fs/promises";
import { join } from "node:path";
import { randomUUID } from "node:crypto";
import { hostname } from "node:os";
import { InputError } from "../errors.js";
import type { SourceDocument, TargetRequisites } from "./types.js";

/**
 * Локальный журнал загрузчика: sourceKey → operationId/Ref_Key, состояние и значимые поля
 * проверенного плана. Полных реквизитов источника и ИНН здесь нет; журнал хранится локально
 * и не коммитится.
 */

export type BankLoadState = "prepared" | "submitted" | "created" | "unknown" | "rejected";

export interface BankLoadEntry {
  version: 1;
  sourceKey: string;
  /** SHA-256 от источника и целевых реквизитов: смена данных под тем же ключом видна. */
  sourceHash: string;
  operationId: string;
  state: BankLoadState;
  attempt: number;
  updatedAt: string;
  ref?: string;
  number?: string;
  errorKind?: string;
  errorMessage?: string;
  /** Контрольное чтение после создания: ok / failed / journal_receipt. */
  verification?: "ok" | "failed" | "journal_receipt";
  verificationNote?: string;
  /** Отпечаток проверенного оператором плана: база, ссылки и payload предпросмотра. */
  planHash?: string;
  /** База, в которой проверялся план. */
  planDatabase?: string;
  /** Объект, куда пишется документ: нужен для контрольного чтения после неизвестного исхода. */
  entitySet?: string;
  /** Значимые поля плана для проверки найденного документа, а не только факта его существования. */
  expected?: { header: Record<string, unknown>; row: Record<string, unknown> };
}

interface JournalFile {
  version: 1;
  entries: Record<string, BankLoadEntry>;
}

export interface JournalSummary {
  dir: string;
  total: number;
  byState: Record<BankLoadState, number>;
  entries: Array<Pick<BankLoadEntry, "sourceKey" | "state" | "operationId" | "updatedAt" | "ref" | "number">>;
}

const EMPTY: JournalFile = { version: 1, entries: {} };

export class BankLoadJournal {
  private readonly file: string;
  private readonly lockFile: string;
  private lockToken?: string;
  constructor(readonly directory: string) {
    this.file = join(directory, "source-journal.json");
    this.lockFile = join(directory, "journal.lock");
  }

  /**
   * Один писатель на каталог журнала. Создание lock-файла атомарно (`wx`); любой существующий
   * lock — отказ без автоматического удаления. Если процесс точно остановлен, оператор вручную
   * проверяет и удаляет journal.lock: гонка двух «уборщиков» не может снять чужой свежий lock.
   */
  async acquire(): Promise<void> {
    if (this.lockToken) return;
    await mkdir(this.directory, { recursive: true, mode: 0o700 });
    const token = randomUUID();
    try {
      const handle = await open(this.lockFile, "wx", 0o600);
      try {
        await handle.writeFile(
          JSON.stringify({
            token,
            pid: process.pid,
            host: hostname(),
            startedAt: new Date().toISOString(),
          }),
          "utf8",
        );
        await handle.sync();
      } finally {
        await handle.close();
      }
      this.lockToken = token;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
      const owner = await this.lockOwner();
      throw new InputError(
        `Журнал ${this.directory} уже используется` +
          (owner
            ? ` (pid ${String(owner.pid ?? "?")}, host ${String(owner.host ?? "?")}, с ${String(owner.startedAt ?? "?")})`
            : " (владелец lock-файла не прочитан)") +
          `. Автоматически lock не снимается. Если предыдущий процесс точно остановлен, проверьте это ` +
          `и удалите ${this.lockFile} вручную, затем повторите.`,
      );
    }
  }

  async release(): Promise<void> {
    if (!this.lockToken) return;
    const token = this.lockToken;
    this.lockToken = undefined;
    // Не удаляем чужой lock, если файл успели заменить: снимаем только со своим токеном.
    const owner = await this.lockOwner();
    if (owner?.token !== token) return;
    await rm(this.lockFile, { force: true }).catch(() => undefined);
  }

  private async lockOwner(): Promise<
    { token?: string; pid?: number; host?: string; startedAt?: string } | undefined
  > {
    try {
      return JSON.parse(await readFile(this.lockFile, "utf8")) as {
        token?: string;
        pid?: number;
        host?: string;
        startedAt?: string;
      };
    } catch {
      return undefined;
    }
  }

  async read(): Promise<JournalFile> {
    let text: string;
    try {
      text = await readFile(this.file, "utf8");
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return { ...EMPTY, entries: {} };
      throw error;
    }
    try {
      const parsed = JSON.parse(text) as JournalFile;
      if (parsed.version !== 1 || !parsed.entries || typeof parsed.entries !== "object")
        throw new Error("invalid journal");
      return parsed;
    } catch {
      throw new Error(
        `Журнал загрузчика ${this.file} повреждён. Повторная загрузка заблокирована: проверьте базу 1С и журнал вручную.`,
      );
    }
  }

  async get(sourceKey: string): Promise<BankLoadEntry | undefined> {
    return (await this.read()).entries[sourceKey];
  }

  async put(entry: BankLoadEntry): Promise<void> {
    const owned = !this.lockToken;
    if (owned) await this.acquire();
    try {
      const file = await this.read();
      file.entries[entry.sourceKey] = entry;
      await this.write(file);
    } finally {
      if (owned) await this.release();
    }
  }

  async summary(): Promise<JournalSummary> {
    const file = await this.read();
    const byState: Record<BankLoadState, number> = {
      prepared: 0,
      submitted: 0,
      created: 0,
      unknown: 0,
      rejected: 0,
    };
    const entries = Object.values(file.entries).sort((a, b) => a.sourceKey.localeCompare(b.sourceKey));
    for (const entry of entries) byState[entry.state] += 1;
    return {
      dir: this.directory,
      total: entries.length,
      byState,
      entries: entries.map((e) => ({
        sourceKey: e.sourceKey,
        state: e.state,
        operationId: e.operationId,
        updatedAt: e.updatedAt,
        ...(e.ref ? { ref: e.ref } : {}),
        ...(e.number ? { number: e.number } : {}),
        ...(e.verification ? { verification: e.verification } : {}),
      })),
    };
  }

  private async write(file: JournalFile): Promise<void> {
    await import("node:fs/promises").then(({ mkdir }) =>
      mkdir(this.directory, { recursive: true, mode: 0o700 }),
    );
    const temporaryPath = `${this.file}.${randomUUID()}.tmp`;
    const handle = await open(temporaryPath, "wx", 0o600);
    try {
      await handle.writeFile(JSON.stringify(file), "utf8");
      await handle.sync();
    } finally {
      await handle.close();
    }
    try {
      await rename(temporaryPath, this.file);
      const dir = await open(this.directory, "r").catch(() => undefined);
      if (dir) {
        try {
          await dir.sync().catch(() => undefined);
        } finally {
          await dir.close();
        }
      }
    } catch (error) {
      await rm(temporaryPath, { force: true }).catch(() => undefined);
      throw error;
    }
  }
}

/** Отпечаток источника и цели: хранится в журнале вместо самих реквизитов. */
/**
 * Отпечаток журнала предыдущей версии загрузчика (до number/time/ДДС/счёта покупателя).
 * Нужен, чтобы уже созданные строки не выглядели «изменённым источником» после обновления.
 */
export function legacySourceHash(source: SourceDocument, target: TargetRequisites): string {
  const stable = {
    sourceKey: source.sourceKey,
    date: source.date,
    amount: Math.round(source.amount * 100) / 100,
    payerInn: source.payerInn,
    payerKpp: source.payerKpp ?? "",
    payerName: source.payerName ?? "",
    payerAccountNumber: source.payerAccountNumber ?? "",
    purpose: source.purpose ?? "",
    documentNumber: source.documentNumber ?? "",
    documentDate: source.documentDate ?? "",
    vatRate: source.vatRate ?? "",
    vatAmount: source.vatAmount ?? null,
    contractNumber: source.contractNumber ?? "",
    contractRef: source.contractRef?.toLowerCase() ?? "",
    organizationInn: target.organizationInn,
    organizationKpp: target.organizationKpp ?? "",
    bankAccountNumber: target.bankAccountNumber,
    bankBik: target.bankBik ?? "",
  };
  return createHash("sha256").update(JSON.stringify(stable)).digest("hex");
}

/**
 * Отпечаток журнала версии 482449f: number/time/ДДС/счёт покупателя уже есть, а четырёх
 * новых opt-in полей (статья ДДС строки, счёт учёта банка, комментарий, ОтражениеВУСН) — нет.
 * Принимается как эквивалент только когда эти четыре поля отсутствуют.
 */
export function intermediateSourceHash(source: SourceDocument, target: TargetRequisites): string {
  const stable = {
    sourceKey: source.sourceKey,
    date: source.date,
    number: source.number ?? "",
    time: source.time ?? "",
    amount: Math.round(source.amount * 100) / 100,
    payerInn: source.payerInn,
    payerKpp: source.payerKpp ?? "",
    payerName: source.payerName ?? "",
    payerAccountNumber: source.payerAccountNumber ?? "",
    purpose: source.purpose ?? "",
    documentNumber: source.documentNumber ?? "",
    documentDate: source.documentDate ?? "",
    vatRate: source.vatRate ?? "",
    vatAmount: source.vatAmount ?? null,
    cashFlowItemRef: source.cashFlowItemRef?.toLowerCase() ?? "",
    customerInvoiceRef: source.customerInvoiceRef?.toLowerCase() ?? "",
    contractNumber: source.contractNumber ?? "",
    contractRef: source.contractRef?.toLowerCase() ?? "",
    organizationInn: target.organizationInn,
    organizationKpp: target.organizationKpp ?? "",
    bankAccountNumber: target.bankAccountNumber,
    bankBik: target.bankBik ?? "",
  };
  return createHash("sha256").update(JSON.stringify(stable)).digest("hex");
}

export function hashSource(source: SourceDocument, target: TargetRequisites): string {
  const stable = {
    sourceKey: source.sourceKey,
    date: source.date,
    number: source.number ?? "",
    time: source.time ?? "",
    amount: Math.round(source.amount * 100) / 100,
    payerInn: source.payerInn,
    payerKpp: source.payerKpp ?? "",
    payerName: source.payerName ?? "",
    payerAccountNumber: source.payerAccountNumber ?? "",
    purpose: source.purpose ?? "",
    documentNumber: source.documentNumber ?? "",
    documentDate: source.documentDate ?? "",
    vatRate: source.vatRate ?? "",
    vatAmount: source.vatAmount ?? null,
    cashFlowItemRef: source.cashFlowItemRef?.toLowerCase() ?? "",
    rowCashFlowItemRef: source.rowCashFlowItemRef?.toLowerCase() ?? "",
    customerInvoiceRef: source.customerInvoiceRef?.toLowerCase() ?? "",
    bankLedgerAccountRef: source.bankLedgerAccountRef?.toLowerCase() ?? "",
    comment: source.comment ?? "",
    incomeRecognition: source.incomeRecognition ?? "",
    contractNumber: source.contractNumber ?? "",
    contractRef: source.contractRef?.toLowerCase() ?? "",
    organizationInn: target.organizationInn,
    organizationKpp: target.organizationKpp ?? "",
    bankAccountNumber: target.bankAccountNumber,
    bankBik: target.bankBik ?? "",
  };
  return createHash("sha256").update(JSON.stringify(stable)).digest("hex");
}

/**
 * Детерминированный operationId (UUIDv5 от sourceKey и отпечатка): повторный запуск с теми же
 * данными идёт в тот же журнал операции MCP, а не создаёт вторую запись.
 */
export function operationIdFor(sourceKey: string, sourceHash: string, attempt = 1): string {
  const bytes = createHash("sha1")
    .update("1c-odata-mcp/bank-loader")
    .update("\n")
    .update(sourceKey)
    .update("\n")
    .update(sourceHash)
    .update("\n")
    .update(String(attempt))
    .digest();
  const b = Buffer.from(bytes.subarray(0, 16));
  b[6] = (b[6]! & 0x0f) | 0x50;
  b[8] = (b[8]! & 0x3f) | 0x80;
  const hex = b.toString("hex");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}
