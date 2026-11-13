import type { Connection } from "../context.js";
import { fetchAll } from "./pagination.js";
import { CATALOGS } from "../config/mapping.js";
import { requireEntity } from "./publication.js";
import { InputError } from "../errors.js";
import { cached, lookupWithRefresh } from "./cache.js";
import { and, cmp, odataString } from "./query.js";
import { ODataError } from "./errors.js";

export interface Organization {
  ref: string;
  name: string;
  inn?: string;
  kpp?: string;
}

const orgSet = (conn: Connection): Promise<string> =>
  requireEntity(conn, CATALOGS.organizations, "Справочник «Организации»");

async function loadOrganizations(conn: Connection): Promise<Organization[]> {
  const set = await orgSet(conn);
  const props = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name) ?? []);
  const select = ["Ref_Key", "Description", "ИНН", ...(props.has("КПП") ? ["КПП"] : [])];
  const { rows } = await fetchAll(
    conn.client,
    set,
    { select, orderby: "Description" },
    conn.behavior.pageSize,
    conn.behavior.maxRows,
  );
  return rows.map((r) => ({
    ref: String(r["Ref_Key"] ?? ""),
    name: String(r["Description"] ?? ""),
    inn: r["ИНН"] ? String(r["ИНН"]) : undefined,
    kpp: r["КПП"] ? String(r["КПП"]) : undefined,
  }));
}

const ORGS_KEY = "organizations";
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

/** Все организации (юрлица) базы — из кэша подключения (cache.ts). */
export async function listOrganizations(conn: Connection): Promise<Organization[]> {
  return cached(conn, ORGS_KEY, () => loadOrganizations(conn));
}

/**
 * Резолвит организацию в Ref_Key: по части названия (без учёта регистра), по Ref_Key или по ИНН.
 * Бросает понятную ошибку, если ничего не найдено или найдено неоднозначно.
 */
export async function resolveOrganization(conn: Connection, query: string): Promise<Organization> {
  const q = query
    .trim()
    .replace(/^\{|\}$/g, "")
    .toLowerCase();
  // Ссылка из list_organizations или ИНН — точное совпадение; иначе — часть названия.
  const match = GUID.test(q)
    ? (o: Organization) => o.ref.toLowerCase() === q
    : /^\d{10,12}$/.test(q)
      ? (o: Organization) => o.inn === q
      : (o: Organization) => o.name.toLowerCase().includes(q);
  const found = await lookupWithRefresh(
    conn,
    ORGS_KEY,
    () => loadOrganizations(conn),
    (orgs) => orgs.filter(match),
    (r) => r.length === 0,
  );
  if (found.length === 0) {
    throw new InputError(`Организация "${query}" не найдена. Список — в list_organizations.`);
  }
  // Точное совпадение названия снимает неоднозначность («Ромашка» и «Ромашка-Юг»).
  const exact = found.filter((o) => o.name.toLowerCase() === q);
  if (found.length > 1 && exact.length !== 1) {
    const names = found.map((o) => o.name).join(", ");
    throw new InputError(`Под "${query}" подходит несколько организаций: ${names}. Уточните название.`);
  }
  return exact.length === 1 ? exact[0]! : found[0]!;
}

/**
 * Резолвит организацию по названию; без названия — авто-выбор, если в базе
 * ровно одна. Общий хелпер для всех инструментов с необязательным `organization`.
 */
export async function resolveOrgOrDefault(
  conn: Connection,
  organization: string | undefined,
): Promise<Organization> {
  if (organization) return resolveOrganization(conn, organization);
  const orgs = await listOrganizations(conn);
  if (orgs.length === 1) return orgs[0]!;
  throw new InputError(
    `В базе несколько организаций — укажите organization. Доступные: ${orgs.map((o) => o.name).join(", ")}`,
  );
}

/** Точный проход по организациям: сколько строк реально прочитано и был ли проход полным. */
export interface OrganizationLookup {
  organizations: Organization[];
  /** Сколько строк прочитано из 1С (до локальной сверки ИНН/КПП). */
  scanned: number;
  /** Проход оборван пределом строк: среди непрочитанных мог быть второй претендент. */
  truncated: boolean;
  /** Отбор на стороне 1С; scan — после HTTP 500 на фильтре по ИНН. */
  strategy: "filter" | "scan";
  /** Текст отказа 1С на фильтр (для отчёта оператора), не раскрывающий реквизиты. */
  filterError?: string;
}

const isFilterServerError = (error: unknown): boolean => error instanceof ODataError && error.status === 500;

const innDigits = (value: string): boolean => /^\d{10}$|^\d{12}$/.test(value);
const kppDigits = (value: string): boolean => /^\d{9}$/.test(value);

/**
 * Сколько строк коллекции читать при точечном поиске. Отдельный потолок нужен, чтобы
 * организация с десятками тысяч контрагентов не тянула весь справочник в память.
 */
const REQUISITES_SCAN_MAX = 5_000;

/**
 * Точный поиск организаций по ИНН и КПП.
 *
 * Фильтр по ИНН в 1С может ответить HTTP 500 (зависит от публикации/полей), поэтому это
 * ускорение, а не источник истины: при 500 читаем страницы до предела и сверяем ИНН/КПП
 * локально. $top=1 не используется: одна строка не доказывает единственность.
 */
export async function findOrganizationsByRequisites(
  conn: Connection,
  p: { inn: string; kpp?: string | undefined },
): Promise<OrganizationLookup> {
  const inn = p.inn.trim();
  const kpp = p.kpp?.trim() || undefined;
  if (!innDigits(inn)) throw new InputError("ИНН организации должен содержать 10 или 12 цифр.");
  if (kpp && !kppDigits(kpp)) throw new InputError("КПП организации должен содержать 9 цифр.");
  const set = await orgSet(conn);
  const props = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name) ?? []);
  if (!props.has("ИНН")) throw new InputError("Справочник «Организации» не публикует поле «ИНН».");
  const hasKpp = props.has("КПП");
  if (kpp && !hasKpp)
    throw new InputError(
      "Справочник «Организации» не публикует поле «КПП» — точная сверка КПП невозможна; " +
        "уберите КПП из цели только осознанно или опубликуйте поле.",
    );
  const select = ["Ref_Key", "Description", "ИНН", ...(hasKpp ? ["КПП"] : [])];
  let strategy: OrganizationLookup["strategy"] = "filter";
  let filterError: string | undefined;
  let rows;
  let truncated: boolean;
  try {
    ({ rows, truncated } = await fetchAll(
      conn.client,
      set,
      {
        filter: and(
          cmp("ИНН", "eq", odataString(inn)),
          kpp && hasKpp ? cmp("КПП", "eq", odataString(kpp)) : undefined,
        ),
        select,
      },
      conn.behavior.pageSize,
      Math.min(conn.behavior.maxRows, REQUISITES_SCAN_MAX),
    ));
  } catch (error) {
    if (!isFilterServerError(error)) throw error;
    // HTTP 500 на фильтре: не считаем, что организации нет, — читаем страницы и сверяем локально.
    strategy = "scan";
    filterError = error instanceof ODataError ? error.message : String(error);
    ({ rows, truncated } = await fetchAll(
      conn.client,
      set,
      { select },
      conn.behavior.pageSize,
      Math.min(conn.behavior.maxRows, REQUISITES_SCAN_MAX),
    ));
  }
  const organizations = rows
    .filter((r) => String(r["ИНН"] ?? "").trim() === inn)
    .filter((r) => (kpp ? String(r["КПП"] ?? "").trim() === kpp : true))
    .map((r) => ({
      ref: String(r["Ref_Key"] ?? ""),
      name: String(r["Description"] ?? ""),
      inn,
      ...(kpp ? { kpp } : {}),
    }));
  return {
    organizations,
    scanned: rows.length,
    truncated,
    strategy,
    ...(filterError ? { filterError } : {}),
  };
}

/**
 * Ровно одна организация по ИНН/КПП. Отказ при нуле, нескольких совпадениях или неполном
 * проходе: догадка о «первой подходящей» дала бы документ не того юрлица.
 */
export async function resolveOrganizationByRequisites(
  conn: Connection,
  p: { inn: string; kpp?: string | undefined },
): Promise<Organization> {
  const found = await findOrganizationsByRequisites(conn, p);
  const query = `ИНН ${p.inn}${p.kpp ? `, КПП ${p.kpp}` : ""}`;
  if (found.truncated)
    throw new InputError(
      `Организации: прочитано ${found.scanned} строк и достигнут предел — проход неполный. ` +
        "Нельзя утверждать, что подходящая организация одна. Увеличьте ODATA_MAX_ROWS или уточните КПП.",
    );
  if (found.organizations.length === 0)
    throw new InputError(`Организация с ${query} не найдена. Проверьте реквизиты и состав OData.`);
  if (found.organizations.length > 1)
    throw new InputError(
      `Под ${query} подходит несколько организаций: ${found.organizations
        .map((o) => `${o.name} (${o.ref})`)
        .join("; ")}. Уточните КПП или исправьте справочник.`,
    );
  return found.organizations[0]!;
}
