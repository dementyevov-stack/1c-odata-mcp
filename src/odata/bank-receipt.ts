import type { Connection } from "../context.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
import { InputError } from "../errors.js";
import type { ODataEntity } from "../types/odata.js";
import { ODataError } from "./errors.js";
import { fetchAll } from "./pagination.js";
import { requireEntity } from "./publication.js";
import { and, cmp, odataGuid, odataString } from "./query.js";

/**
 * Разбор реквизитов строки банковской выписки для первого участка загрузчика БП:
 * организация, её банковский счёт, контрагент-покупатель и договор.
 *
 * Точная сверка локальная: фильтр по ИНН в 1С может ответить HTTP 500, поэтому он только
 * ускорение. Одна строка или $top=1 не доказывают единственность — читаем ограниченный
 * проход страницами и отказываем при неоднозначности или обрыве предела.
 */

/** Предел строк для точечного поиска реквизитов (не весь справочник в память). */
const REQUISITES_SCAN_MAX = 5_000;

interface ExactLookup {
  rows: ODataEntity[];
  scanned: number;
  truncated: boolean;
  strategy: "filter" | "scan";
  filterError?: string;
}

const isServer500 = (error: unknown): error is ODataError =>
  error instanceof ODataError && error.status === 500;

const fieldsOf = async (conn: Connection, set: string): Promise<Set<string>> =>
  new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name) ?? []);

const str = (v: unknown): string =>
  typeof v === "string" ? v : v === undefined || v === null ? "" : String(v);
const guid = (v: unknown): string => str(v).trim().toLowerCase();
const isReal = (v: unknown): boolean => {
  const g = guid(v);
  return g.length > 0 && g !== "00000000-0000-0000-0000-000000000000";
};
const trimInn = (v: unknown): string => str(v).trim();
const trimKpp = (v: unknown): string => str(v).trim();

const innDigits = (value: string): boolean => /^\d{10}$|^\d{12}$/.test(value);
const kppDigits = (value: string): boolean => /^\d{9}$/.test(value);

/**
 * Контрагент по точным ИНН/КПП. Удалённые и папки не подходят. Несколько совпадений —
 * отказ: запись документа на одного из однофамильцев была бы угадыванием.
 */
export async function resolveCounterpartyByRequisites(
  conn: Connection,
  p: { inn: string; kpp?: string | undefined },
): Promise<{ ref: string; name: string }> {
  const inn = p.inn.trim();
  const kpp = p.kpp?.trim() || undefined;
  if (!innDigits(inn)) throw new InputError("ИНН контрагента должен содержать 10 или 12 цифр.");
  if (kpp && !kppDigits(kpp)) throw new InputError("КПП контрагента должен содержать 9 цифр.");
  const set = await requireEntity(conn, CATALOGS.counterparties, "Справочник «Контрагенты»");
  const props = await fieldsOf(conn, set);
  if (!props.has("ИНН"))
    throw new InputError("Справочник «Контрагенты» не публикует поле «ИНН» — точный подбор невозможен.");
  if (kpp && !props.has("КПП"))
    throw new InputError(
      "Справочник «Контрагенты» не публикует поле «КПП» — точная сверка КПП невозможна; " +
        "уберите КПП из источника только осознанно или опубликуйте поле.",
    );
  const select = ["Ref_Key", "Description", "ИНН", "КПП", "DeletionMark", "IsFolder"].filter(
    (f) => f === "Ref_Key" || f === "Description" || props.has(f),
  );
  const maxRows = Math.min(conn.behavior.maxRows, REQUISITES_SCAN_MAX);
  const read = async (filter?: string): Promise<ExactLookup> => {
    const { rows, truncated } = await fetchAll(
      conn.client,
      set,
      { filter, select },
      conn.behavior.pageSize,
      maxRows,
    );
    return { rows, truncated, scanned: rows.length, strategy: filter ? "filter" : "scan" };
  };
  let found: ExactLookup;
  try {
    found = await read(
      and(
        cmp("ИНН", "eq", odataString(inn)),
        kpp && props.has("КПП") ? cmp("КПП", "eq", odataString(kpp)) : undefined,
      ),
    );
  } catch (error) {
    if (!isServer500(error)) throw error;
    // HTTP 500 на фильтре по ИНН: читаем страницы и сверяем реквизиты локально.
    found = await read();
    found = {
      ...found,
      strategy: "scan",
      filterError: error.message,
    };
  }
  if (found.truncated)
    throw new InputError(
      `Контрагенты: прочитано ${found.scanned} строк и достигнут предел — проход неполный. ` +
        "Нельзя утверждать, что подходящий контрагент один. Увеличьте ODATA_MAX_ROWS или уточните реквизиты.",
    );
  const exact = found.rows.filter(
    (r) => trimInn(r["ИНН"]) === inn && (kpp ? trimKpp(r["КПП"]) === kpp : true),
  );
  const active = exact.filter((r) => r["DeletionMark"] !== true && r["IsFolder"] !== true);
  if (active.length === 0) {
    if (exact.length)
      throw new InputError(
        `Контрагент с ИНН ${inn}${kpp ? `, КПП ${kpp}` : ""} найден, но помечен на удаление или является группой.`,
      );
    throw new InputError(
      `Контрагент с ИНН ${inn}${kpp ? `, КПП ${kpp}` : ""} не найден в справочнике «Контрагенты».`,
    );
  }
  if (active.length > 1)
    throw new InputError(
      `Под ИНН ${inn}${kpp ? `, КПП ${kpp}` : ""} подходит несколько контрагентов: ${active
        .map((r) => `${str(r["Description"])} (${str(r["Ref_Key"])})`)
        .join("; ")}. Уточните КПП или исправьте справочник.`,
    );
  const one = active[0]!;
  return { ref: guid(one["Ref_Key"]), name: str(one["Description"]) };
}

/** Договор покупателя: по Ref_Key, по точному номеру или единственный подходящий у контрагента. */
export async function resolveCustomerContract(
  conn: Connection,
  p: {
    orgKey: string;
    counterpartyRef: string;
    contractRef?: string | undefined;
    contractNumber?: string | undefined;
  },
): Promise<{ ref: string; number?: string; kind?: string; name?: string }> {
  const set = await requireEntity(conn, CATALOGS.contracts, "Справочник «Договоры контрагентов»");
  const props = await fieldsOf(conn, set);
  const select = [
    "Ref_Key",
    "Description",
    "Номер",
    "ВидДоговора",
    "Owner_Key",
    "Организация_Key",
    "DeletionMark",
    "IsFolder",
  ].filter((f) => f === "Ref_Key" || props.has(f));
  const orgKey = guid(p.orgKey);
  const cp = guid(p.counterpartyRef);
  const accepted = (r: ODataEntity): string | undefined => {
    if (r["DeletionMark"] === true) return "помечен на удаление";
    if (r["IsFolder"] === true) return "является группой";
    if (isReal(r["Owner_Key"]) && guid(r["Owner_Key"]) !== cp) return "принадлежит другому контрагенту";
    if (isReal(r["Организация_Key"]) && guid(r["Организация_Key"]) !== orgKey)
      return "относится к другой организации";
    return undefined;
  };
  const describe = (r: ODataEntity): string =>
    `${str(r["Description"]) || "(без названия)"}, номер ${str(r["Номер"]) || "—"} (${str(r["Ref_Key"])})`;

  if (p.contractRef) {
    const ref = guid(p.contractRef);
    const doc = await conn.client.getEntity(`${set}(${odataGuid(ref)})${selectQuery(select)}`);
    const why = accepted(doc);
    if (why) throw new InputError(`Договор ${ref} не подходит: ${why}.`);
    const number = str(doc["Номер"]) || undefined;
    if (p.contractNumber && number !== p.contractNumber.trim())
      throw new InputError(
        `Договор ${ref} имеет номер «${number ?? "—"}», а в источнике «${p.contractNumber}». Ничего не записано.`,
      );
    return {
      ref,
      ...(number ? { number } : {}),
      ...(str(doc["ВидДоговора"]) ? { kind: str(doc["ВидДоговора"]) } : {}),
      ...(str(doc["Description"]) ? { name: str(doc["Description"]) } : {}),
    };
  }

  let rows: ODataEntity[];
  let truncated: boolean;
  try {
    const read = await fetchAll<ODataEntity>(
      conn.client,
      set,
      {
        filter: and(
          cmp("Owner_Key", "eq", odataGuid(cp)),
          props.has("DeletionMark") ? cmp("DeletionMark", "eq", "false") : undefined,
        ),
        select,
      },
      conn.behavior.pageSize,
      Math.min(conn.behavior.maxRows, REQUISITES_SCAN_MAX),
    );
    rows = read.rows;
    truncated = read.truncated;
  } catch (error) {
    if (!isServer500(error)) throw error;
    // Фильтр по Owner_Key может быть запрещён публикацией: ограниченный проход без фильтра,
    // затем точная локальная сверка владельца, организации и удаления.
    const read = await fetchAll<ODataEntity>(conn.client, set, { select }, conn.behavior.pageSize, 5_000);
    rows = read.rows;
    truncated = read.truncated;
  }
  if (truncated)
    throw new InputError(
      `Договоры контрагента: прочитано ${rows.length} строк и достигнут предел — проход неполный. ` +
        "Укажите номер или Ref_Key договора, чтобы не выбирать по неполному списку.",
    );
  const usable = rows.filter((r) => guid(r["Owner_Key"]) === cp && !accepted(r));
  const number = p.contractNumber?.trim();
  if (number) {
    const hits = usable.filter((r) => str(r["Номер"]).trim() === number);
    if (hits.length === 0)
      throw new InputError(
        `Договор № ${number} у контрагента ${p.counterpartyRef} не найден. Укажите Ref_Key договора или исправьте номер.`,
      );
    if (hits.length > 1)
      throw new InputError(
        `Договоров № ${number} у контрагента несколько: ${hits.map(describe).join("; ")}. Укажите Ref_Key.`,
      );
    return {
      ref: guid(hits[0]!["Ref_Key"]),
      number,
      ...(str(hits[0]!["ВидДоговора"]) ? { kind: str(hits[0]!["ВидДоговора"]) } : {}),
      ...(str(hits[0]!["Description"]) ? { name: str(hits[0]!["Description"]) } : {}),
    };
  }

  // Без номера — только единственный подходящий договор покупателя/«Прочее».
  const suitable = usable.filter((r) => {
    const kind = str(r["ВидДоговора"]);
    return kind === "СПокупателем" || kind === "Прочее" || kind === "";
  });
  if (suitable.length === 0)
    throw new InputError(
      "У контрагента нет действующего договора покупателя. Укажите номер или Ref_Key договора, " +
        "либо создайте договор в 1С. Ничего не записано.",
    );
  if (suitable.length > 1)
    throw new InputError(
      `У контрагента несколько подходящих договоров: ${suitable.map(describe).join("; ")}. ` +
        "Укажите номер или Ref_Key договора: выбирать первый нельзя.",
    );
  const one = suitable[0]!;
  return {
    ref: guid(one["Ref_Key"]),
    ...(str(one["Номер"]) ? { number: str(one["Номер"]) } : {}),
    ...(str(one["ВидДоговора"]) ? { kind: str(one["ВидДоговора"]) } : {}),
    ...(str(one["Description"]) ? { name: str(one["Description"]) } : {}),
  };
}

/**
 * Банковский счёт владельца (организации или контрагента) по точному номеру. Фильтр по номеру —
 * ускорение; владелец проверяется локально (в 1С Owner — полиморфная ссылка, сравнивать её в
 * $filter нельзя).
 */
export async function findBankAccountsByOwnerAndNumber(
  conn: Connection,
  ownerKey: string,
  accountNumber: string,
  expectedOwnerType?: string,
): Promise<{ accounts: Array<{ ref: string; number: string; name?: string }>; truncated: boolean }> {
  const number = accountNumber.trim();
  if (!/^\d{20}$/.test(number)) throw new InputError("Номер банковского счёта должен содержать 20 цифр.");
  const set = await requireEntity(conn, CATALOGS.bankAccounts, "Справочник «Банковские счета»");
  const props = await fieldsOf(conn, set);
  if (!props.has("НомерСчета"))
    throw new InputError("Справочник «Банковские счета» не публикует поле «НомерСчета».");
  if (!props.has("Owner"))
    throw new InputError(
      "Справочник «Банковские счета» не публикует владельца «Owner» — нельзя отличить счёт организации.",
    );
  if (expectedOwnerType && !props.has("Owner_Type"))
    throw new InputError(
      "Справочник «Банковские счета» не публикует «Owner_Type» — нельзя отличить счёт организации от счёта контрагента.",
    );
  const select = ["Ref_Key", "Description", "НомерСчета", "Owner", "Owner_Type", "DeletionMark"].filter(
    (f) => f === "Ref_Key" || props.has(f),
  );
  const maxRows = Math.min(conn.behavior.maxRows, REQUISITES_SCAN_MAX);
  const read = async (filter?: string) => {
    const { rows, truncated } = await fetchAll(
      conn.client,
      set,
      { filter, select },
      conn.behavior.pageSize,
      maxRows,
    );
    return { rows, truncated };
  };
  let found: { rows: ODataEntity[]; truncated: boolean };
  try {
    found = await read(cmp("НомерСчета", "eq", odataString(number)));
  } catch (error) {
    if (!isServer500(error)) throw error;
    // HTTP 500 на точечном фильтре: ограниченный проход по всем счетам и локальная сверка.
    found = await read();
  }
  const owner = guid(ownerKey);
  const exact = found.rows.filter(
    (r) =>
      str(r["НомерСчета"]).trim() === number &&
      guid(r["Owner"]) === owner &&
      (!expectedOwnerType || str(r["Owner_Type"]) === expectedOwnerType) &&
      r["DeletionMark"] !== true,
  );
  return {
    accounts: exact.map((r) => ({
      ref: guid(r["Ref_Key"]),
      number,
      ...(str(r["Description"]) ? { name: str(r["Description"]) } : {}),
    })),
    truncated: found.truncated,
  };
}

export async function resolveBankAccountByOwnerAndNumber(
  conn: Connection,
  ownerKey: string,
  accountNumber: string,
  label: string,
  expectedOwnerType?: string,
): Promise<{ ref: string; number: string; name?: string }> {
  const found = await findBankAccountsByOwnerAndNumber(conn, ownerKey, accountNumber, expectedOwnerType);
  if (found.truncated)
    throw new InputError(
      "Банковские счета: достигнут предел строк, проход неполный. " +
        "Нельзя утверждать, что счёт с таким номером один. Увеличьте ODATA_MAX_ROWS.",
    );
  if (found.accounts.length === 0)
    throw new InputError(
      `У ${label} ${ownerKey} нет банковского счёта № ${accountNumber}. ` +
        "Проверьте владельца счёта: счёт контрагента не подходит вместо счёта организации.",
    );
  if (found.accounts.length > 1)
    throw new InputError(
      `Счёт № ${accountNumber} найден у ${label} несколько раз — исправьте справочник 1С.`,
    );
  return found.accounts[0]!;
}

export async function resolveOrgBankAccountByNumber(
  conn: Connection,
  orgKey: string,
  accountNumber: string,
): Promise<{ ref: string; number: string; name?: string }> {
  const orgSet = resolveEntity(CATALOGS.organizations, await conn.available());
  if (!orgSet)
    throw new InputError("Справочник «Организации» не опубликован — счёт организации не сопоставить.");
  return resolveBankAccountByOwnerAndNumber(
    conn,
    orgKey,
    accountNumber,
    "организации",
    `StandardODATA.${orgSet}`,
  );
}

export async function resolveCounterpartyBankAccountByNumber(
  conn: Connection,
  counterpartyRef: string,
  accountNumber: string,
): Promise<{ ref: string; number: string; name?: string }> {
  const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
  if (!cpSet)
    throw new InputError("Справочник «Контрагенты» не опубликован — счёт контрагента не сопоставить.");
  return resolveBankAccountByOwnerAndNumber(
    conn,
    counterpartyRef,
    accountNumber,
    "контрагента",
    `StandardODATA.${cpSet}`,
  );
}

/** Простой $select для одиночного GET (buildQuery не нужен: только поля, без фильтра). */
function selectQuery(fields: string[]): string {
  return fields.length ? `?$format=json&$select=${encodeURIComponent(fields.join(","))}` : "?$format=json";
}
