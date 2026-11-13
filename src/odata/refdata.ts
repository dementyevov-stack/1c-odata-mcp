import type { Connection } from "../context.js";
import { fetchAll } from "./pagination.js";
import { and, cmp, contains, odataGuid, odataString, or } from "./query.js";
import { resolveEntity } from "../config/mapping.js";
import { requireEntity } from "./publication.js";
import { InputError } from "../errors.js";
import { ODataError } from "./errors.js";
import { CATALOGS } from "../config/mapping.js";

/**
 * Резолверы справочных данных, нужных для записи «богатой» карточки контрагента:
 * виды контактной информации, банк по БИК, доп.реквизиты.
 */

export const KI_KINDS = ["Catalog_ВидыКонтактнойИнформации"] as const;
const BANKS = ["Catalog_Банки"] as const;
const ADD_PROPS = ["ChartOfCharacteristicTypes_ДополнительныеРеквизитыИСведения"] as const;

export interface ContactKind {
  key: string;
  тип: string;
}
export interface ContactKinds {
  phone?: ContactKind;
  email?: ContactKind;
  address?: ContactKind;
}

/** Виды КИ для контрагентов: телефон / email / (юридический) адрес. */
export async function contactKindsForCounterparties(conn: Connection): Promise<ContactKinds> {
  const set = resolveEntity(KI_KINDS, await conn.available());
  if (!set) return {};
  const { rows: folders } = await fetchAll(
    conn.client,
    set,
    {
      filter: and(contains("Description", "Контрагенты"), cmp("IsFolder", "eq", "true")),
      select: ["Ref_Key"],
    },
    1,
    1,
  );
  const folder = folders[0];
  if (!folder) return {};
  const { rows } = await fetchAll(
    conn.client,
    set,
    {
      filter: and(
        cmp("Parent_Key", "eq", odataGuid(String(folder["Ref_Key"]))),
        cmp("IsFolder", "eq", "false"),
      ),
      select: ["Ref_Key", "Description", "Тип"],
    },
    50,
    50,
  );
  const out: ContactKinds = {};
  for (const r of rows) {
    const t = String(r["Тип"]);
    const d = String(r["Description"]);
    const kind = { key: String(r["Ref_Key"]), тип: t };
    if (t === "Телефон" && !out.phone) out.phone = kind;
    else if (t === "АдресЭлектроннойПочты" && !out.email) out.email = kind;
    else if (t === "Адрес" && d.toLowerCase().includes("юридическ")) out.address = kind;
  }
  if (!out.address) {
    const a = rows.find((r) => String(r["Тип"]) === "Адрес");
    if (a) out.address = { key: String(a["Ref_Key"]), тип: "Адрес" };
  }
  return out;
}

const BANKS_SCAN_MAX = 5_000;
const isServer500 = (error: unknown): error is ODataError =>
  error instanceof ODataError && error.status === 500;
const fieldsOf = async (conn: Connection, set: string): Promise<Set<string>> =>
  new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name) ?? []);

/**
 * Банк по БИК или СВИФТБИК. В БП БИК — Code, а СВИФТБИК — отдельное поле; фильтр
 * по Code/СВИФТБИК у части публикаций отвечает 500. Тогда читаем ограниченный справочник
 * без фильтра и сверяем оба поля локально; неполный проход или >1 совпадение — отказ.
 */
export async function resolveBankByBik(
  conn: Connection,
  bik: string,
): Promise<{ ref: string; name: string }> {
  const query = bik.trim();
  if (!/^[A-Z0-9]{8,11}$/i.test(query))
    throw new InputError("БИК должен содержать 9 цифр, а СВИФТБИК — 8–11 латинских букв/цифр.");
  const set = await requireEntity(conn, BANKS, "Справочник «Банки»");
  const props = await fieldsOf(conn, set);
  if (!props.has("Code") && !props.has("СВИФТБИК"))
    throw new InputError("Справочник «Банки» не публикует ни «Code», ни «СВИФТБИК».");
  const select = ["Ref_Key", "Description", "Code", "СВИФТБИК", "IsFolder", "DeletionMark"].filter(
    (f) => f === "Ref_Key" || f === "Description" || props.has(f),
  );
  const matches = (row: ODataEntityLike, q: string): boolean =>
    (props.has("Code") && String(row["Code"] ?? "").trim() === q) ||
    (props.has("СВИФТБИК") &&
      String(row["СВИФТБИК"] ?? "")
        .trim()
        .toUpperCase() === q.toUpperCase());
  const active = (row: ODataEntityLike): boolean => row["DeletionMark"] !== true && row["IsFolder"] !== true;
  type ODataEntityLike = Record<string, unknown>;

  const conditions = [
    props.has("Code") ? cmp("Code", "eq", odataString(query.toUpperCase())) : undefined,
    props.has("СВИФТБИК") ? cmp("СВИФТБИК", "eq", odataString(query.toUpperCase())) : undefined,
  ].filter((c): c is string => Boolean(c));
  let rows: ODataEntityLike[];
  let truncated: boolean;
  try {
    const read = await fetchAll<ODataEntityLike>(
      conn.client,
      set,
      { filter: or(...conditions), select },
      100,
      100,
    );
    rows = read.rows;
    truncated = read.truncated;
  } catch (error) {
    if (!isServer500(error)) throw error;
    // Ограниченный полный проход: фильтр по Code/СВИФТБИК в этой публикации запрещён.
    const read = await fetchAll<ODataEntityLike>(
      conn.client,
      set,
      { select },
      conn.behavior.pageSize,
      Math.min(conn.behavior.maxRows, BANKS_SCAN_MAX),
    );
    rows = read.rows;
    truncated = read.truncated;
  }
  if (truncated)
    throw new InputError(
      "Справочник «Банки»: достигнут предел строк, проход неполный. Нельзя утверждать, " +
        "что банк с таким БИК/СВИФТБИК один. Увеличьте ODATA_MAX_ROWS.",
    );
  const exact = rows.filter((r) => matches(r, query));
  const usable = exact.filter(active);
  if (usable.length === 0) {
    if (exact.length)
      throw new InputError(
        `Банк с БИК/СВИФТБИК ${query} найден, но помечен на удаление или является группой.`,
      );
    throw new InputError(`Банк с БИК/СВИФТБИК ${query} не найден в справочнике «Банки».`);
  }
  if (usable.length > 1)
    throw new InputError(
      `Под БИК/СВИФТБИК ${query} подходит несколько банков: ${usable
        .map((r) => `${String(r["Description"] ?? "")} (${String(r["Ref_Key"] ?? "")})`)
        .join("; ")}. Уточните реквизит.`,
    );
  const one = usable[0]!;
  return { ref: String(one["Ref_Key"]), name: String(one["Description"] ?? "") };
}

/** Валюта организации по коду/названию; при 500 на фильтре — ограниченный локальный проход. */
export async function resolveCurrencyForAccount(
  conn: Connection,
  query: string,
): Promise<{ ref: string; code?: string; name?: string }> {
  const q = query.trim();
  if (!q) throw new InputError("Валюта: пустое значение.");
  const set = await requireEntity(conn, CATALOGS.currencies, "Справочник «Валюты»");
  const props = await fieldsOf(conn, set);
  if (!props.has("Code")) throw new InputError("Справочник «Валюты» не публикует поле «Code».");
  const select = ["Ref_Key", "Description", "Code", "DeletionMark"].filter(
    (f) => f === "Ref_Key" || props.has(f),
  );
  type ODataEntityLike = Record<string, unknown>;
  const matches = (row: ODataEntityLike): boolean =>
    String(row["Code"] ?? "").trim() === q ||
    String(row["Description"] ?? "")
      .trim()
      .toLowerCase() === q.toLowerCase();
  const readAll = async (filter?: string) =>
    fetchAll<ODataEntityLike>(
      conn.client,
      set,
      { ...(filter ? { filter } : {}), select },
      conditionsPageSize(conn),
      Math.min(conn.behavior.maxRows, 5_000),
    );
  let rows: ODataEntityLike[];
  let truncated: boolean;
  try {
    const read = await readAll(and(conditionFor(props, q), cmp("DeletionMark", "eq", "false")));
    rows = read.rows;
    truncated = read.truncated;
  } catch (error) {
    if (!isServer500(error)) throw error;
    const read = await readAll();
    rows = read.rows;
    truncated = read.truncated;
  }
  if (truncated)
    throw new InputError(
      "Справочник «Валюты»: достигнут предел строк, проход неполный. Уточните валюту кодом.",
    );
  const exact = rows.filter((r) => matches(r) && r["DeletionMark"] !== true);
  if (exact.length === 0) throw new InputError(`Валюта «${query}» не найдена.`);
  if (exact.length > 1)
    throw new InputError(`Под валюту «${query}» подходит несколько записей; укажите точный код.`);
  const one = exact[0]!;
  return {
    ref: String(one["Ref_Key"]),
    ...(String(one["Code"] ?? "") ? { code: String(one["Code"]) } : {}),
    ...(String(one["Description"] ?? "") ? { name: String(one["Description"]) } : {}),
  };
}

function conditionsPageSize(conn: Connection): number {
  return Math.max(1, Math.min(conn.behavior.pageSize, 100));
}

function conditionFor(props: Set<string>, q: string): string {
  const conditions = [
    props.has("Code") ? cmp("Code", "eq", odataString(q)) : undefined,
    props.has("Description") ? cmp("Description", "eq", odataString(q)) : undefined,
  ].filter((c): c is string => Boolean(c));
  return or(...conditions) || "true";
}

/** Ref доп.реквизита по точному наименованию (напр. «ОГРН»), либо undefined если не настроен. */
export async function resolveAdditionalProperty(conn: Connection, name: string): Promise<string | undefined> {
  const set = resolveEntity(ADD_PROPS, await conn.available());
  if (!set) return undefined;
  const { rows } = await fetchAll(
    conn.client,
    set,
    { filter: cmp("Description", "eq", odataString(name)), select: ["Ref_Key"] },
    1,
    1,
  );
  return rows[0] ? String(rows[0]["Ref_Key"]) : undefined;
}
