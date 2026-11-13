import {
  ALIGN_HINT,
  bankPostingBlocker,
  expectedDocumentVersion,
  guardedDocumentAction,
} from "../odata/existing-bank-settlement.js";
import { operationKind as configuredOperationKind, type AccountDefault } from "../config/accounting.js";
import type { CallToolResult, McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField, renderJson } from "./_shared.js";
import { CATALOGS, WRITE_DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { resolveOrgOrDefault, resolveOrganizationByRequisites } from "../odata/orgs.js";
import { ensurePublished, requireEntity } from "../odata/publication.js";
import {
  accountsByCode,
  accountsByRef,
  nomenclatureAccounts,
  resolveDefaultAccounts,
  type NomAccounts,
} from "../odata/accounting.js";
import { settlementAccountsFor, type SettlementSide } from "../odata/settlement-accounts.js";
import {
  findBankAccountsByOwnerAndNumber,
  resolveCounterpartyBankAccountByNumber,
  resolveCounterpartyByRequisites,
  resolveCustomerContract,
  resolveOrgBankAccountByNumber,
} from "../odata/bank-receipt.js";
import {
  contactKindsForCounterparties,
  resolveBankByBik,
  resolveCurrencyForAccount,
  resolveAdditionalProperty,
  type ContactKinds,
} from "../odata/refdata.js";
import { fetchAll } from "../odata/pagination.js";
import { and, buildQuery, cmp, contains, odataGuid, odataString } from "../odata/query.js";
import type { ODataEntity } from "../types/odata.js";
import {
  createResultSchema,
  patchResultSchema,
  markForDeletionResultSchema,
  postDocumentResultSchema,
  operationStatusSchema,
} from "../schemas/output.js";
import { currentWriteOperationId, fingerprintRecordFields } from "../odata/write-operation-context.js";
import { InputError } from "../errors.js";
import { ODataError } from "../odata/errors.js";
import { getDocumentPostings } from "./registers.js";
import { tryPreviewView } from "./preview-view.js";

/** Тип ссылки на номенклатуру в табличной части (полиморфная ссылка 1С). */
const NOMENCLATURE_TYPE = "StandardODATA.Catalog_Номенклатура";
const COUNTERPARTY_TYPE = "StandardODATA.Catalog_Контрагенты";

/** Строит строки табличной части «КонтактнаяИнформация» из телефона/email/адреса. */
function buildContactRows(
  kinds: ContactKinds,
  data: { phone?: string | undefined; email?: string | undefined; address?: string | undefined },
): Array<Record<string, unknown>> {
  const rows: Array<Record<string, unknown>> = [];
  if (data.phone && kinds.phone)
    rows.push({
      Тип: kinds.phone.тип,
      Вид_Key: kinds.phone.key,
      Представление: data.phone,
      НомерТелефона: data.phone,
    });
  if (data.email && kinds.email)
    rows.push({
      Тип: kinds.email.тип,
      Вид_Key: kinds.email.key,
      Представление: data.email,
      АдресЭП: data.email,
    });
  if (data.address && kinds.address)
    rows.push({ Тип: kinds.address.тип, Вид_Key: kinds.address.key, Представление: data.address });
  return rows.map((r, i) => ({ LineNumber: i + 1, ...r }));
}

/** Дополнительные поля карточки контрагента: контактная информация + ОГРН (доп.реквизит). */
async function counterpartyExtras(
  conn: Connection,
  data: {
    phone?: string | undefined;
    email?: string | undefined;
    address?: string | undefined;
    ogrn?: string | undefined;
  },
): Promise<{ fields: Record<string, unknown>; notes: string[] }> {
  const fields: Record<string, unknown> = {};
  const notes: string[] = [];
  if (data.phone || data.email || data.address) {
    const kinds = await contactKindsForCounterparties(conn);
    const rows = buildContactRows(kinds, data);
    if (rows.length) fields["КонтактнаяИнформация"] = rows;
    if (data.phone && !kinds.phone) notes.push("Вид «Телефон» не найден — телефон не записан.");
    if (data.email && !kinds.email) notes.push("Вид «Email» не найден — email не записан.");
    if (data.address && !kinds.address) notes.push("Вид «Адрес» не найден — адрес не записан.");
  }
  if (data.ogrn) {
    const prop = await resolveAdditionalProperty(conn, "ОГРН");
    if (prop)
      fields["ДополнительныеРеквизиты"] = [{ LineNumber: 1, Свойство_Key: prop, Значение: data.ogrn }];
    else
      notes.push(
        "ОГРН не записан: в этой базе нет дополнительного реквизита «ОГРН» для контрагентов. Заведите его в 1С (Администрирование → Дополнительные реквизиты), либо вносите ОГРН вручную.",
      );
  }
  return { fields, notes };
}

/** Виды договоров (Enum_ВидыДоговоровКонтрагентов). */
const CONTRACT_KINDS = ["СПокупателем", "СПоставщиком", "Прочее", "СКомиссионером", "СКомитентом"] as const;

/** Ставки НДС (Enum_СтавкиНДС), подмножество ходовых. */
const VAT_RATES = ["БезНДС", "НДС0", "НДС5", "НДС7", "НДС10", "НДС20", "НДС22"] as const;

/** Значения Enum_ВидыСтавокНДС из проверенной схемы БП 3.0. */
const VAT_KIND_VALUES = [
  "БезНДС",
  "Нулевая",
  "Общая",
  "ОбщаяРасчетная",
  "ОбщаяУСН",
  "ОбщаяУСНРасчетная",
  "Пониженная",
  "ПониженнаяРасчетная",
  "ПониженнаяУСН",
  "ПониженнаяУСНРасчетная",
] as const;

/**
 * Ставки НДС для переноса явных исторических значений из банковской выписки. Список — из
 * Enum_СтавкиНДС БП 3.0, включая ставки до 2019 года и расчётные; неизвестное значение
 * отклоняется до записи, а не отправляется в 1С наугад.
 */
/** Значения Enum_ОтражениеВУСН из проверенной схемы БП 3.0. */
const INCOME_RECOGNITION_VALUES = [
  "ВозвратРасхода",
  "НеПринимаются",
  "Принимаются",
  "Распределяются",
] as const;

const BANK_VAT_RATES = [
  "БезНДС",
  "НДС0",
  "НДС5",
  "НДС5_105",
  "НДС7",
  "НДС7_107",
  "НДС10",
  "НДС10_110",
  "НДС18",
  "НДС18_118",
  "НДС20",
  "НДС20_120",
  "НДС22",
  "НДС22_122",
] as const;

/**
 * Дата → формат 1С Edm.DateTime ('YYYY-MM-DDTHH:mm:ss', без зоны).
 * Берём ЛОКАЛЬНЫЕ компоненты (не toISOString/UTC): 1С хранит «настенную» дату без
 * зоны, а UTC-сдвиг в плюсовых поясах уводит полночь на предыдущий день.
 */
export function odataDate(d: Date): string {
  const p = (n: number): string => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

/** Резолвит организацию: по названию, либо авто, если в базе ровно одна. */
async function resolveOrg(
  conn: Connection,
  organization: string | undefined,
): Promise<{ key: string; name: string }> {
  const o = await resolveOrgOrDefault(conn, organization);
  return { key: o.ref, name: o.name };
}

/** Сколько совпадений по части названия читать, чтобы назвать варианты в отказе. */
const NAME_MATCH_LIMIT = 20;

/** Единственный объект; несколько — отказ с кодами и ссылками: первое дало бы документ с чужим складом или валютой. */
function singleMatch(rows: readonly ODataEntity[], query: string, label: string): ODataEntity {
  if (rows.length === 1) return rows[0] as ODataEntity;
  const variants = rows
    .slice(0, 10)
    .map((r) => {
      const code = r["Code"] ? `код ${String(r["Code"]).trim()}, ` : "";
      return `${String(r["Description"] ?? "")} (${code}${String(r["Ref_Key"])})`;
    })
    .join("; ");
  throw new InputError(
    `${label}: под «${query}» подходит несколько: ${variants}. Укажите Ref_Key, код или название точнее. Ничего не записано.`,
  );
}

const catalogProps = async (conn: Connection, set: string): Promise<Set<string>> =>
  new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name));

/** Рабочие объекты: помеченные на удаление и папки документ не должен получать. */
const activeOnly = (props: ReadonlySet<string>): string =>
  and(
    props.has("DeletionMark") ? cmp("DeletionMark", "eq", "false") : undefined,
    props.has("IsFolder") ? cmp("IsFolder", "eq", "false") : undefined,
  );

/**
 * Объект справочника по Ref_Key, коду или названию. Помеченные на удаление и папки не выбираются. Точное название
 * ищется отдельным запросом: среди совпадений по части оно могло бы не попасть в выборку и потеряться.
 */
async function findCatalogObject(
  conn: Connection,
  set: string,
  query: string,
  label: string,
  codeFirst: boolean,
): Promise<ODataEntity | undefined> {
  const props = await catalogProps(conn, set);
  const select = ["Ref_Key", "Description", ...(props.has("Code") ? ["Code"] : [])];
  const active = activeOnly(props);
  const read = async (condition: string) =>
    (
      await fetchAll<ODataEntity>(
        conn.client,
        set,
        { filter: and(condition, active || undefined), select },
        NAME_MATCH_LIMIT,
        NAME_MATCH_LIMIT,
      )
    ).rows;
  if (GUID_RE.test(query)) {
    const rows = await read(cmp("Ref_Key", "eq", odataGuid(query.replace(/[{}]/g, ""))));
    return rows[0];
  }
  // Пустой substringof в 1С находит всё, и единственный объект выбрался бы по строке из пробелов.
  if (!query) throw new InputError(`${label}: пустое значение. Ничего не записано.`);
  const byCode = () =>
    props.has("Code") ? read(cmp("Code", "eq", odataString(query))) : Promise.resolve([] as ODataEntity[]);
  const byName = () => read(cmp("Description", "eq", odataString(query)));
  for (const step of codeFirst ? [byCode, byName] : [byName, byCode]) {
    const rows = await step();
    if (rows.length) return singleMatch(rows, query, label);
    // Точное название у помеченного объекта: иначе молча выбрался бы другой, с таким кодом или похожим названием.
    if (step === byName) {
      const { rows: hidden } = await fetchAll<ODataEntity>(
        conn.client,
        set,
        { filter: cmp("Description", "eq", odataString(query)), select },
        1,
        1,
      );
      if (hidden.length) {
        throw new InputError(
          `${label}: «${query}» помечен на удаление или это папка. Укажите рабочий объект точным названием, кодом или Ref_Key. Ничего не записано.`,
        );
      }
    }
  }
  const partial = await read(contains("Description", query));
  if (!partial.length) return undefined;
  const wanted = query.toLowerCase();
  const exact = partial.filter(
    (r) =>
      String(r["Description"] ?? "")
        .trim()
        .toLowerCase() === wanted,
  );
  return singleMatch(exact.length ? exact : partial, query, label);
}

/** Резолвит склад по названию; если не задан и склад один — берёт его; иначе undefined. */
async function resolveWarehouse(conn: Connection, rawName: string | undefined): Promise<string | undefined> {
  const name = rawName?.trim();
  if (rawName !== undefined && !name) throw new InputError("Склад: пустое значение. Ничего не записано.");
  const set = resolveEntity(CATALOGS.warehouses, await conn.available());
  if (!set) {
    if (name)
      throw new InputError(
        `Справочник складов не опубликован в OData: склад "${name}" выбрать нельзя. Ничего не записано.`,
      );
    return undefined;
  }
  if (name) {
    const found = await findCatalogObject(conn, set, name, "Склад", false);
    if (!found) throw new InputError(`Склад "${name}" не найден (см. list_entities / справочник складов).`);
    return String(found["Ref_Key"]);
  }
  const active = activeOnly(await catalogProps(conn, set));
  const { rows } = await fetchAll(
    conn.client,
    set,
    { filter: active || undefined, select: ["Ref_Key"] },
    2,
    2,
  );
  return rows.length === 1 ? String((rows[0] as ODataEntity)["Ref_Key"]) : undefined;
}

/**
 * Резолвит элемент справочника по точному коду или части наименования.
 * Публикацию проверяет (requireEntity). Бросает Error, если не найдено.
 */
async function resolveCatalogItem(
  conn: Connection,
  candidates: readonly string[],
  label: string,
  query: string,
): Promise<{ ref: string; code?: string; name?: string }> {
  const set = await requireEntity(conn, candidates, label);
  const first = await findCatalogObject(conn, set, query.trim(), label, true);
  if (!first) throw new InputError(`${label}: «${query}» не найдено (по коду или наименованию).`);
  return {
    ref: String(first["Ref_Key"]),
    code: first["Code"] ? String(first["Code"]) : undefined,
    name: first["Description"] ? String(first["Description"]) : undefined,
  };
}

const GUID_RE = /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/;

/** Находит папку (группу, IsFolder=true) в иерархическом справочнике по части имени. */
async function resolveFolder(
  conn: Connection,
  entitySet: string,
  query: string,
): Promise<{ ref: string; name: string }> {
  const { rows } = await fetchAll(
    conn.client,
    entitySet,
    {
      filter: and(contains("Description", query), cmp("IsFolder", "eq", "true")),
      select: ["Ref_Key", "Description"],
    },
    10,
    10,
  );
  if (rows.length === 0) throw new InputError(`Папка «${query}» не найдена в ${entitySet}.`);
  if (rows.length > 1) {
    const names = rows.map((r) => String(r["Description"])).join(", ");
    throw new InputError(`Под «${query}» несколько папок: ${names}. Уточните название.`);
  }
  const r = rows[0] as ODataEntity;
  return { ref: String(r["Ref_Key"]), name: String(r["Description"] ?? "") };
}

/** Папка по ref (GUID) или по имени (резолв). undefined → без родителя. */
async function folderRefOf(
  conn: Connection,
  entitySet: string,
  folder: string | undefined,
): Promise<string | undefined> {
  if (!folder) return undefined;
  return GUID_RE.test(folder.trim())
    ? folder.replace(/[{}]/g, "")
    : (await resolveFolder(conn, entitySet, folder)).ref;
}

export interface GoodsLine {
  nomenclatureRef: string;
  quantity: number;
  price: number;
  vatRate: string;
  /** Содержание строки — печатается в счёте/УПД (для услуг обязательно по смыслу). */
  content?: string | undefined;
  /** Счёт доходов: код плана счетов («90.01.2») или готовый Ref_Key. */
  incomeAccount?: string | undefined;
  /** Счёт расходов/себестоимости: код («90.02.2») или Ref_Key. */
  expenseAccount?: string | undefined;
  /** Исходная строка документа (при правке существующего) — чтобы не терять её поля. */
  source?: Record<string, unknown> | undefined;
  /** Историческая сумма строки (если нельзя/не нужно считать quantity × price). */
  amount?: number | undefined;
  /** Явная историческая сумма НДС строки. */
  vatAmount?: number | undefined;
}
type LineAccountsFor = (line: GoodsLine) => Record<string, string>;

/** Счета в строке не заполняются (счёт поставщика, счёт покупателю — проводок нет). */
const noAccounts: LineAccountsFor = () => ({});

/** Округление денег до копеек — суммы складываем только через него. */
const roundMoney = (n: number): number => Math.round(n * 100) / 100;
const lineSum = (l: GoodsLine): number => roundMoney(l.quantity * l.price);
const rowsTotal = (rows: Array<Record<string, unknown>>): number =>
  Math.round(rows.reduce((s, r) => s + (r["Сумма"] as number), 0) * 100) / 100;

/**
 * Поля исходной строки, которые НЕ переносим: считаем заново (иначе после смены
 * количества/цены останется старая сумма НДС) либо это служебная ссылка.
 */
const NOT_CARRIED_ROW_FIELDS = ["LineNumber", "Сумма", "СуммаНДС", "СуммаСНДС", "Ref_Key"] as const;

/**
 * Слагаемые суммы НДС строки не изменились — значит прежняя сумма налога верна.
 * Пересчитать её мы не можем: формула зависит от СуммаВключаетНДС документа и
 * процента ставки, и угадывать тут нельзя. Но и терять нельзя — иначе у документа
 * с НДС строка уезжает в 1С с нулевым налогом.
 */
function vatInputsUnchanged(l: GoodsLine): boolean {
  const s = l.source;
  if (!s) return false;
  return (
    Number(s["Количество"]) === l.quantity &&
    Number(s["Цена"]) === l.price &&
    String(s["СтавкаНДС"] ?? "БезНДС") === l.vatRate &&
    // Сумму строки мы считаем как количество×цену. Если в исходной строке она была
    // другой (ручная скидка в 1С), то и прежний налог посчитан не от нашей суммы.
    Number(s["Сумма"]) === lineSum(l)
  );
}

/**
 * Прочие поля исходной строки (номенклатурная группа, ГТД, страна происхождения,
 * субконто) — переносим как есть: мы ими не управляем, но терять их нельзя.
 * СуммаНДС переносится только при неизменных количестве/цене/ставке.
 */
export function carriedRowFields(l: GoodsLine): Record<string, unknown> {
  if (!l.source) return {};
  const out: Record<string, unknown> = { ...l.source };
  for (const f of NOT_CARRIED_ROW_FIELDS) delete out[f];
  if (vatInputsUnchanged(l)) out["СуммаНДС"] = l.source["СуммаНДС"];
  return out;
}

/** Строки табличной части «Товары»/«Услуги» для поступления/реализации (Номенклатура_Key + счета). */
export function buildGoodsRows(
  lines: GoodsLine[],
  lineAccountsFor: LineAccountsFor,
): Array<Record<string, unknown>> {
  return lines.map((l, i) =>
    clean({
      ...carriedRowFields(l),
      LineNumber: i + 1,
      Номенклатура_Key: l.nomenclatureRef,
      // Только когда задано: иначе затёрли бы содержание, перенесённое из исходной строки.
      ...(l.content === undefined ? {} : { Содержание: l.content }),
      Количество: l.quantity,
      Цена: l.price,
      Сумма: lineSum(l),
      СтавкаНДС: l.vatRate,
      ...lineAccountsFor(l),
    }),
  );
}

type InvoiceLine = Omit<GoodsLine, "quantity"> & {
  quantity?: number | undefined;
  emptyQuantity?: boolean | undefined;
};

/** Строки «Товары» для счёта покупателю (полиморфная ссылка Номенклатура+_Type). */
export function buildInvoiceRows(lines: InvoiceLine[]): Array<Record<string, unknown>> {
  return lines.map((l, i) =>
    clean({
      ...carriedRowFields(l as GoodsLine),
      LineNumber: i + 1,
      Номенклатура: l.nomenclatureRef,
      Номенклатура_Type: NOMENCLATURE_TYPE,
      // Только когда задано: иначе затёрли бы содержание, перенесённое из исходной строки.
      ...(l.content === undefined ? {} : { Содержание: l.content }),
      // Opt-in пустого количества: поле не сериализуется вообще, сумма обязательна.
      ...(l.emptyQuantity ? {} : { Количество: l.quantity }),
      Цена: l.price,
      Сумма: l.amount ?? (l.emptyQuantity ? undefined : lineSum(l as GoodsLine)),
      СтавкаНДС: l.vatRate,
      ...(l.vatAmount !== undefined ? { СуммаНДС: l.vatAmount } : {}),
    }),
  );
}

/** Строит и создаёт/предпросматривает товарный документ (поступление/реализация). */
async function createGoodsDoc(
  conn: Connection,
  entitySet: string,
  p: {
    orgKey: string;
    counterpartyRef: string;
    contractRef?: string | undefined;
    warehouseKey?: string | undefined;
    date?: string | undefined;
    sumIncludesVat: boolean;
    lines: GoodsLine[];
    settlement?: string | undefined; // СчетУчетаРасчетовСКонтрагентом_Key (шапка)
    bankAccountKey?: string | undefined; // БанковскийСчетОрганизации_Key (есть у реализации)
    lineAccountsFor: LineAccountsFor; // счета строки
  },
  confirm: boolean,
) {
  const rows = buildGoodsRows(p.lines, p.lineAccountsFor);
  const payload = clean({
    Date: odataDate(p.date ? new Date(`${p.date}T00:00:00`) : new Date()),
    Posted: false,
    Организация_Key: p.orgKey,
    Контрагент_Key: p.counterpartyRef,
    ДоговорКонтрагента_Key: p.contractRef,
    Склад_Key: p.warehouseKey,
    СчетУчетаРасчетовСКонтрагентом_Key: p.settlement,
    БанковскийСчетОрганизации_Key: p.bankAccountKey,
    СуммаВключаетНДС: p.sumIncludesVat,
    СуммаДокумента: rowsTotal(rows),
    Товары: rows,
  });
  return createOrPreview(conn, entitySet, payload, confirm);
}

/**
 * Резолвит счета учёта по каждой номенклатуре: сперва из регистра «Счета учёта
 * номенклатуры», недостающее добивает дефолтами по кодам плана счетов.
 */
async function resolveLineAccounts(
  conn: Connection,
  orgKey: string,
  nomRefs: string[],
  defaults: NomAccounts,
): Promise<Map<string, NomAccounts>> {
  const map = new Map<string, NomAccounts>();
  for (const ref of new Set(nomRefs)) {
    const reg = await nomenclatureAccounts(conn, orgKey, ref);
    map.set(ref, {
      goods: reg?.goods ?? defaults.goods,
      incomingVat: reg?.incomingVat ?? defaults.incomingVat,
      outgoingVat: reg?.outgoingVat ?? defaults.outgoingVat,
      income: reg?.income ?? defaults.income,
      expense: reg?.expense ?? defaults.expense,
    });
  }
  return map;
}

/**
 * Делит заданные пользователем счета на готовые ссылки (GUID — берём как есть,
 * нормализуя фигурные скобки) и коды плана счетов, которые надо найти в базе.
 * Ключ карты и элементы codes — исходное значение после trim.
 */
export function splitAccountRefs(values: Array<string | undefined>): {
  refs: Map<string, string>;
  codes: string[];
} {
  const refs = new Map<string, string>();
  const codes = new Set<string>();
  for (const v of values) {
    const t = v?.trim();
    if (!t) continue;
    if (GUID_RE.test(t)) refs.set(t, t.replace(/[{}]/g, ""));
    else codes.add(t);
  }
  return { refs, codes: [...codes] };
}

/**
 * Резолвит заданные пользователем счета: код плана счетов («90.01.2») → Ref_Key,
 * готовый GUID — как есть. Все коды одним запросом; ключ карты — исходная строка
 * после trim. Неизвестный код — ошибка, чтобы не записать документ «мимо» счёта.
 */
async function resolveAccountOverrides(
  conn: Connection,
  values: Array<string | undefined>,
): Promise<Map<string, string>> {
  const { refs, codes } = splitAccountRefs(values);
  if (codes.length === 0) return refs;
  const found = await accountsByCode(conn, codes);
  for (const c of codes) {
    const ref = found.get(c);
    if (!ref)
      throw new InputError(
        `Счёт «${c}» не найден в плане счетов «Хозрасчётный». Укажите код точно как в 1С (напр. 90.01.2) или Ref_Key счёта.`,
      );
    refs.set(c, ref);
  }
  return refs;
}

/** Счета учёта, заданные на уровне документа (применяются ко всем строкам). */
interface DocAccounts {
  income?: string | undefined;
  expense?: string | undefined;
}

/**
 * Готовит счета учёта для товарного документа: счёт расчётов (шапка) и функцию
 * счетов строки по номенклатуре. kind различает поступление (Дт 41 Кт 60) и
 * реализацию (Дт 62 Кт 90, Дт 90 Кт 41). Приоритет счетов доходов/расходов:
 * строка → документ → регистр «Счета учёта номенклатуры» → коды по умолчанию.
 */
async function goodsAccounts(
  conn: Connection,
  orgKey: string,
  lines: GoodsLine[],
  kind: "purchase" | "shipment" | "service",
  docAccounts?: DocAccounts,
): Promise<{ settlement?: string | undefined; lineAccountsFor: LineAccountsFor }> {
  const accMap = await resolveLineAccounts(
    conn,
    orgKey,
    lines.map((l) => l.nomenclatureRef),
    {},
  );
  const overrides =
    kind === "purchase"
      ? new Map<string, string>()
      : await resolveAccountOverrides(conn, [
          docAccounts?.income,
          docAccounts?.expense,
          ...lines.flatMap((l) => [l.incomeAccount, l.expenseAccount]),
        ]);
  const ref = (v: string | undefined): string | undefined => overrides.get(v?.trim() ?? "");
  const needed = new Set<AccountDefault>(
    conn.cfg?.accounting?.strictDefaults ? [] : [kind === "purchase" ? "payable" : "receivable"],
  );
  for (const l of lines) {
    const a = accMap.get(l.nomenclatureRef);
    if (kind !== "service" && !a?.goods) needed.add("goods");
    if (kind === "purchase") {
      if (l.vatRate !== "БезНДС" && !a?.incomingVat) needed.add("incomingVat");
    } else {
      if (!(ref(l.incomeAccount) ?? ref(docAccounts?.income) ?? a?.income)) needed.add("income");
      if (!(ref(l.expenseAccount) ?? ref(docAccounts?.expense) ?? a?.expense)) needed.add("expense");
      if (l.vatRate !== "БезНДС" && !a?.outgoingVat) needed.add("outgoingVat");
    }
  }
  const defaults = await resolveDefaultAccounts(conn, [...needed]);
  return {
    settlement: defaults.get(kind === "purchase" ? "payable" : "receivable"),
    lineAccountsFor: (l) => {
      const a = accMap.get(l.nomenclatureRef);
      const goods = kind === "service" ? {} : { СчетУчета_Key: a?.goods ?? defaults.get("goods") };
      if (kind === "purchase")
        return clean({
          ...goods,
          ...(l.vatRate !== "БезНДС"
            ? { СчетУчетаНДС_Key: a?.incomingVat ?? defaults.get("incomingVat") }
            : {}),
        }) as Record<string, string>;
      return clean({
        ...goods,
        СчетДоходов_Key:
          ref(l.incomeAccount) ?? ref(docAccounts?.income) ?? a?.income ?? defaults.get("income"),
        СчетРасходов_Key:
          ref(l.expenseAccount) ?? ref(docAccounts?.expense) ?? a?.expense ?? defaults.get("expense"),
        ...(l.vatRate !== "БезНДС"
          ? { СчетУчетаНДСПоРеализации_Key: a?.outgoingVat ?? defaults.get("outgoingVat") }
          : {}),
      }) as Record<string, string>;
    },
  };
}

/**
 * Счёт учёта товаров (41.01) на строку — для складских документов без расчётов
 * с контрагентом (перемещение/оприходование/списание/инвентаризация).
 */
async function goodsOnlyAccounts(
  conn: Connection,
  orgKey: string,
  lines: GoodsLine[],
): Promise<(nomRef: string) => Record<string, string>> {
  const accMap = await resolveLineAccounts(
    conn,
    orgKey,
    lines.map((l) => l.nomenclatureRef),
    {},
  );
  const needsGoods = lines.some((l) => !accMap.get(l.nomenclatureRef)?.goods);
  const defaults = await resolveDefaultAccounts(conn, needsGoods ? ["goods"] : []);
  return (nomRef) =>
    clean({ СчетУчета_Key: accMap.get(nomRef)?.goods ?? defaults.get("goods") }) as Record<string, string>;
}

const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";
/** Непустая ссылка строкой; пустой GUID и пустая строка → undefined. */
const refOrUndefined = (v: unknown): string | undefined => {
  const s = typeof v === "string" ? v.replace(/[{}']/g, "") : "";
  return s && s !== EMPTY_GUID ? s : undefined;
};

/**
 * Разбирает строку ТЧ документа в GoodsLine, сохраняя исходную строку в source:
 * при пересборке ТЧ (add/remove/update_document_lines) содержание, счета и прочие
 * поля строки должны пережить правку, а не обнулиться.
 */
export function lineFromRow(r: Record<string, unknown>): GoodsLine {
  return {
    nomenclatureRef: String(r["Номенклатура_Key"] ?? r["Номенклатура"] ?? ""),
    quantity: Number(r["Количество"] ?? 0),
    price: Number(r["Цена"] ?? 0),
    vatRate: String(r["СтавкаНДС"] ?? "БезНДС"),
    content: (r["Содержание"] as string) || undefined,
    incomeAccount: refOrUndefined(r["СчетДоходов_Key"]),
    expenseAccount: refOrUndefined(r["СчетРасходов_Key"]),
    source: r,
  };
}

/** Табличная часть документа, в которой лежат позиции. */
export type Section = "Товары" | "Услуги";

/**
 * Определяет табличную часть документа. У реализации их две: «Товары» и «Услуги»,
 * и заполнена ровно одна — акт услуг (ВидОперации=Услуги) живёт в «Услугах».
 * Ориентируемся сперва на факт (где есть строки), затем на тип и вид операции —
 * у пустого документа строк ещё нет.
 */
export function pickSection(
  doc: { goods?: unknown[] | undefined; services?: unknown[] | undefined },
  hints: { servicesAct: boolean; operationKind?: unknown },
): Section {
  if (doc.services?.length) return "Услуги";
  if (doc.goods?.length) return "Товары";
  if (hints.servicesAct) return "Услуги";
  return hints.operationKind === "Услуги" ? "Услуги" : "Товары";
}

/** Строки табличной части документа (пустой массив, если ТЧ нет или она пуста). */
export function sectionRows(doc: Record<string, unknown>, section: Section): Array<Record<string, unknown>> {
  return (doc[section] as Array<Record<string, unknown>> | undefined) ?? [];
}

/** Вторая табличная часть — та, которую сейчас не правим. */
const otherSection = (s: Section): Section => (s === "Товары" ? "Услуги" : "Товары");

/**
 * Сумма «Сумма» по строкам. Документ может нести позиции в ОБЕИХ ТЧ (реализация
 * с товарами и услугами), поэтому итог документа = правимая ТЧ + вторая как есть:
 * считать только по одной значило бы занизить СуммаДокумента.
 */
export function rowsSum(rows: Array<Record<string, unknown>>): number {
  return roundMoney(rows.reduce((s, r) => s + Number(r["Сумма"] ?? 0), 0));
}

/**
 * Читает документ: организация, проведён ли, заполненная ТЧ и её строки.
 * otherTotal — сумма ВТОРОЙ табличной части: её мы не трогаем, но в
 * СуммаДокумента она входит (у реализации бывают заполнены и «Товары», и «Услуги»).
 */
async function getDocInfo(
  conn: Connection,
  entitySet: string,
  guid: string,
): Promise<{
  orgKey: string;
  posted: boolean;
  lines: GoodsLine[];
  section: Section;
  otherTotal: number;
}> {
  const doc = await conn.client.getEntity(`${entitySet}(guid'${guid}')${buildQuery({})}`);
  const section = pickSection(
    {
      goods: doc["Товары"] as unknown[] | undefined,
      services: doc["Услуги"] as unknown[] | undefined,
    },
    {
      servicesAct: (WRITE_DOCUMENTS.servicesAct as readonly string[]).includes(entitySet),
      operationKind: doc["ВидОперации"],
    },
  );
  return {
    orgKey: String(doc["Организация_Key"] ?? ""),
    posted: doc["Posted"] === true,
    lines: sectionRows(doc, section).map(lineFromRow),
    section,
    otherTotal: rowsSum(sectionRows(doc, otherSection(section))),
  };
}

/**
 * Реквизиты шапки, которые в копию не переносим: они принадлежат исходному
 * документу. Number отсутствует намеренно — 1С автонумерует новый документ,
 * только если Number не передан вовсе.
 */
const NOT_COPIED_DOC_FIELDS = [
  "Ref_Key",
  "Number",
  "Posted",
  "DeletionMark",
  "DataVersion",
  "СсылочныйИдентификатор",
] as const;

/**
 * Поля строк ТЧ, которые не копируем: идентификаторы строк связывают строку с
 * исходным документом и его «потомками» — в копии они должны быть своими.
 */
const NOT_COPIED_ROW_FIELDS = ["Ref_Key", "ИдентификаторСтроки", "ИдентификаторРодительскойСтроки"] as const;

/**
 * Готовит прочитанный документ к записи как нового: убирает служебные поля 1С,
 * технические ключи OData (odata.metadata, *@navigationLinkUrl) и идентификаторы
 * строк табличных частей. Остальное переносится как есть — в этом и смысл копии.
 */
export function stripForCopy(entity: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(entity)) {
    if (key.startsWith("odata.") || key.includes("@")) continue;
    if ((NOT_COPIED_DOC_FIELDS as readonly string[]).includes(key)) continue;
    if (Array.isArray(value)) {
      out[key] = value.map((row) => {
        if (!row || typeof row !== "object") return row;
        const copy: Record<string, unknown> = { ...(row as Record<string, unknown>) };
        for (const f of NOT_COPIED_ROW_FIELDS) delete copy[f];
        return copy;
      });
      continue;
    }
    out[key] = value;
  }
  return out;
}

/** Переопределение полей одной строки табличной части по её номеру. */
export interface LineOverride {
  lineNumber: number;
  fields: Record<string, unknown>;
}

/**
 * Применяет переопределения к строкам ТЧ по номеру строки. Номер, которого нет
 * в документе, — ошибка ввода: молча проигнорировать правку суммы нельзя.
 */
export function applyLineOverrides(
  rows: Array<Record<string, unknown>>,
  overrides: LineOverride[],
): Array<Record<string, unknown>> {
  const byNumber = new Map(overrides.map((o) => [o.lineNumber, o.fields]));
  const missing = overrides.filter((o) => o.lineNumber < 1 || o.lineNumber > rows.length);
  if (missing.length) {
    throw new InputError(
      `В документе ${rows.length} строк(и), а правка задана для строки ${missing
        .map((m) => m.lineNumber)
        .join(", ")}.`,
    );
  }
  return rows.map((row, i) => ({ ...row, ...(byNumber.get(i + 1) ?? {}) }));
}

/**
 * Сумма документа по строкам ТЧ. undefined, если хотя бы у одной строки нет
 * числовой «Суммы» — тогда пересчитывать нельзя и сумму оставляем исходную.
 */
export function totalFromRows(rows: Array<Record<string, unknown>>): number | undefined {
  let sum = 0;
  for (const row of rows) {
    const v = row["Сумма"];
    const n = typeof v === "number" ? v : typeof v === "string" ? Number(v) : NaN;
    if (!Number.isFinite(n)) return undefined;
    sum += n;
  }
  return Math.round(sum * 100) / 100;
}

/** Читает документ-основание счёта-фактуры: организация, контрагент, договор, суммы, НДС. */
async function invoiceBasis(
  conn: Connection,
  entitySet: string,
  ref: string,
): Promise<{ org: string; counterparty: string; contract?: string; total: number; vat: number }> {
  const guid = ref.replace(/[{}']/g, "");
  const doc = await conn.client.getEntity(`${entitySet}(guid'${guid}')${buildQuery({})}`);
  // НДС собираем по ОБЕИМ табличным частям: акт услуг (create_act) живёт в
  // «Услугах», и чтение одних «Товаров» давало счёт-фактуру с нулевым НДС.
  const rows = [...sectionRows(doc, "Товары"), ...sectionRows(doc, "Услуги")];
  const vat = Math.round(rows.reduce((s, r) => s + Number(r["СуммаНДС"] ?? 0), 0) * 100) / 100;
  return {
    org: String(doc["Организация_Key"] ?? ""),
    counterparty: String(doc["Контрагент_Key"] ?? doc["Контрагент"] ?? ""),
    contract: (doc["ДоговорКонтрагента_Key"] as string) || undefined,
    total: Number(doc["СуммаДокумента"] ?? 0),
    vat,
  };
}

/**
 * Собирает строки табличной части под тип документа (счёт/поступление/реализация/акт).
 * section — куда пойдут строки: для «Услуг» счёт учёта 41 не ставится (склада нет).
 */
async function buildSectionRows(
  conn: Connection,
  entitySet: string,
  orgKey: string,
  lines: GoodsLine[],
  section: Section,
): Promise<{ rows: Array<Record<string, unknown>>; total: number } | undefined> {
  const inList = (arr: readonly string[]): boolean => arr.includes(entitySet);
  if (inList(WRITE_DOCUMENTS.customerInvoice)) {
    const rows = buildInvoiceRows(lines);
    return { rows, total: rowsTotal(rows) };
  }
  if (inList(WRITE_DOCUMENTS.servicesAct)) {
    const { lineAccountsFor } = await goodsAccounts(conn, orgKey, lines, "service");
    const rows = buildGoodsRows(lines, lineAccountsFor);
    return { rows, total: rowsTotal(rows) };
  }
  // Реализация/поступление и возвраты: возврат покупателя считается как реализация
  // (Дт 90/62), возврат поставщику — как поступление (Дт 41/19).
  if (
    inList(WRITE_DOCUMENTS.purchases) ||
    inList(WRITE_DOCUMENTS.sales) ||
    inList(WRITE_DOCUMENTS.returnFromCustomer) ||
    inList(WRITE_DOCUMENTS.returnToSupplier)
  ) {
    const sale = inList(WRITE_DOCUMENTS.sales) || inList(WRITE_DOCUMENTS.returnFromCustomer);
    const kind = sale ? (section === "Услуги" ? "service" : "shipment") : "purchase";
    const { lineAccountsFor } = await goodsAccounts(conn, orgKey, lines, kind);
    const rows = buildGoodsRows(lines, lineAccountsFor);
    return { rows, total: rowsTotal(rows) };
  }
  if (
    inList(WRITE_DOCUMENTS.surplus) ||
    inList(WRITE_DOCUMENTS.writeoff) ||
    inList(WRITE_DOCUMENTS.transfer)
  ) {
    const lineAcc = await goodsOnlyAccounts(conn, orgKey, lines);
    const withSum = !inList(WRITE_DOCUMENTS.transfer); // у перемещения в ТЧ нет суммы
    const rows = lines.map((l, i) =>
      clean({
        ...carriedRowFields(l),
        LineNumber: i + 1,
        Номенклатура_Key: l.nomenclatureRef,
        ...(l.content === undefined ? {} : { Содержание: l.content }),
        Количество: l.quantity,
        Цена: l.price,
        ...(withSum ? { Сумма: lineSum(l) } : {}),
        ...lineAcc(l.nomenclatureRef),
      }),
    );
    return { rows, total: rowsTotal(rows) };
  }
  return undefined;
}

/**
 * Банковский счёт организации. Приоритет: явный Ref_Key → поиск по названию/номеру →
 * ОСНОВНОЙ счёт организации → первый по алфавиту. Раньше без названия брался
 * произвольный счёт (rows[0] без сортировки) — на счёте и в акте печатались
 * реквизиты не того банка.
 */
export async function resolveOrgBankAccount(
  conn: Connection,
  orgKey: string,
  query: string | undefined,
): Promise<string | undefined> {
  const q = query?.trim();
  if (q && GUID_RE.test(q)) return q.replace(/[{}]/g, "");
  if (!q) {
    const main = await mainOrgBankAccount(conn, orgKey);
    if (main) return main;
    // Подбор не удался — не критично: 1С подставит счёт организации по умолчанию. Из неполного
    // списка «первый по алфавиту» был бы случайным счётом, поэтому тогда не подставляем ничего.
    const list = await orgBankAccounts(conn, orgKey).catch(() => undefined);
    return list && !list.truncated ? list.accounts[0]?.ref : undefined;
  }
  // Ошибка чтения — не «счёта нет»: пусть дойдёт до пользователя как есть (#54).
  const { accounts, truncated } = await orgBankAccounts(conn, orgKey);
  const needle = q.toLowerCase();
  const hits = accounts.filter((a) => a.name.toLowerCase().includes(needle) || a.number.includes(q));
  // Список счетов обрезан: нужного может не быть в прочитанной части, а единственное совпадение —
  // не гарантия единственности. Без явного Ref_Key принимаем только точный номер счёта.
  if (truncated) {
    const exact = hits.filter((a) => a.number === q);
    if (exact.length === 1) return exact[0]!.ref;
    throw new InputError(
      `Банковский счёт организации «${q}»: в базе больше ${ORG_BANK_ACCOUNTS_MAX} банковских счетов, ` +
        "прочитана только часть, поэтому по названию счёт не подобрать. Укажите полный номер счёта или его Ref_Key.",
    );
  }
  if (hits.length === 1) return hits[0]?.ref;
  if (hits.length === 0) {
    const known = accounts.map((a) => a.name).join("; ");
    throw new InputError(
      `Банковский счёт организации «${q}» не найден.` +
        (known ? ` Есть: ${known}.` : " У организации нет счетов."),
    );
  }
  // Несколько совпадений: основной среди них — берём его, иначе просим уточнить.
  const main = await mainOrgBankAccount(conn, orgKey);
  const preferred = hits.find((a) => a.ref === main);
  if (preferred) return preferred.ref;
  throw new InputError(
    `Под «${q}» подходит несколько счетов организации: ${hits.map((a) => a.name).join("; ")}. Уточните номер счёта.`,
  );
}

/**
 * Банковские счета организации списком.
 *
 * Владелец счёта — полиморфная ссылка Owner (+ Owner_Type). Сравнивать саму Owner
 * в $filter 1С не даёт: «Нельзя сравнивать поля неограниченной длины и поля
 * несовместимых типов» (HTTP 500). Зато Owner_Type — обычная строка, ею и сужаем
 * выборку до счетов организаций, а конкретного владельца сверяем на клиенте.
 * Отбор по названию тоже клиентский: substringof по Description 1С отклоняет
 * по той же причине.
 */
/** Сколько банковских счетов (всех организаций) читаем для подбора по названию. */
const ORG_BANK_ACCOUNTS_MAX = 1000;

async function orgBankAccounts(
  conn: Connection,
  orgKey: string,
): Promise<{ accounts: Array<{ ref: string; name: string; number: string }>; truncated: boolean }> {
  const available = await conn.available();
  const set = resolveEntity(CATALOGS.bankAccounts, available);
  if (!set) return { accounts: [], truncated: false };
  const orgSet = resolveEntity(CATALOGS.organizations, available);
  const filter = and(
    orgSet ? cmp("Owner_Type", "eq", odataString(`StandardODATA.${orgSet}`)) : undefined,
    cmp("DeletionMark", "eq", "false"),
  );
  const { rows, truncated } = await fetchAll(
    conn.client,
    set,
    { filter, select: ["Ref_Key", "Description", "НомерСчета", "Owner"] },
    conn.behavior.pageSize,
    ORG_BANK_ACCOUNTS_MAX,
  );
  const accounts = rows
    .filter((r) => String(r["Owner"] ?? "") === orgKey)
    .map((r) => ({
      ref: String(r["Ref_Key"] ?? ""),
      name: String(r["Description"] ?? ""),
      number: String(r["НомерСчета"] ?? ""),
    }))
    .sort((a, b) => a.name.localeCompare(b.name, "ru"));
  return { accounts, truncated };
}

/** Основной банковский счёт организации (реквизит ОсновнойБанковскийСчет). */
async function mainOrgBankAccount(conn: Connection, orgKey: string): Promise<string | undefined> {
  const orgSet = resolveEntity(CATALOGS.organizations, await conn.available());
  if (!orgSet) return undefined;
  try {
    const org = await conn.client.getEntity(
      `${orgSet}(guid'${orgKey.replace(/[{}']/g, "")}')${buildQuery({ select: ["ОсновнойБанковскийСчет_Key"] })}`,
    );
    return refOrUndefined(org["ОсновнойБанковскийСчет_Key"]);
  } catch {
    return undefined;
  }
}

/** Валюта документа по умолчанию: код 643 (рубль), иначе — единственная в справочнике. */
async function resolveDefaultCurrency(conn: Connection): Promise<string | undefined> {
  const set = resolveEntity(CATALOGS.currencies, await conn.available());
  if (!set) return undefined;
  try {
    const byCode = await fetchAll(
      conn.client,
      set,
      { filter: cmp("Code", "eq", odataString("643")), select: ["Ref_Key"] },
      2,
      2,
    );
    if (byCode.rows[0]) return String(byCode.rows[0]["Ref_Key"]);
    const all = await fetchAll(conn.client, set, { select: ["Ref_Key"] }, 2, 2);
    return all.rows.length === 1 ? String((all.rows[0] as ODataEntity)["Ref_Key"]) : undefined;
  } catch {
    return undefined;
  }
}

const confirmField = z
  .boolean()
  .default(false)
  .describe(
    "false (по умолчанию) — только предпросмотр (dry-run), запись НЕ выполняется. " +
      "true — выполнить создание. Сначала всегда показывайте dry-run и получайте согласие пользователя; " +
      "при подтверждении передавайте operationId из результата предпросмотра и те же аргументы. После таймаута " +
      "повторяйте подтверждение с тем же id, не создавая новый предпросмотр.",
  );

const contentField = z
  .string()
  .optional()
  .describe(
    "Содержание строки (реквизит «Содержание») — именно этот текст печатается в счёте и УПД. " +
      "Для услуг заполняйте обязательно, напр. «Услуги согласно приложению № 4 от 01.10.2025 г. " +
      "к договору № 87 от 01.12.2023 г. за июль 2026 г.». 1С сама это поле НЕ заполняет.",
  );

const incomeAccountField = z
  .string()
  .optional()
  .describe(
    "Счёт доходов — код плана счетов как в 1С (напр. «90.01.1» или «90.01.2») либо Ref_Key счёта. " +
      "Без указания берётся из регистра «Счета учёта номенклатуры», иначе 90.01.1.",
  );

const expenseAccountField = z
  .string()
  .optional()
  .describe(
    "Счёт расходов (себестоимости) — код как в 1С (напр. «90.02.1» или «90.02.2») либо Ref_Key. " +
      "Без указания берётся из регистра «Счета учёта номенклатуры», иначе 90.02.1.",
  );

const orgBankAccountField = z
  .string()
  .optional()
  .describe(
    "Банковский счёт организации — печатается в документе как реквизиты для оплаты. " +
      "Название/номер счёта (напр. «Сбербанк» или «40802…») либо Ref_Key. " +
      "Без указания берётся ОСНОВНОЙ счёт организации.",
  );

/** Базовые поля позиции документа (без счетов учёта). */
const baseLineShape = {
  nomenclatureRef: z.string().describe("Ref_Key номенклатуры"),
  quantity: z.number().positive().describe("Количество"),
  price: z.number().nonnegative().describe("Цена за единицу"),
  vatRate: z.enum(VAT_RATES).default("БезНДС").describe("Ставка НДС"),
  content: contentField,
};

/** Позиция документа продажи: со счетами доходов/расходов на строку. */
const saleLine = z.object({
  ...baseLineShape,
  incomeAccount: incomeAccountField,
  expenseAccount: expenseAccountField,
});

/** Позиция документа закупки: счета доходов/расходов неприменимы. */
const purchaseLine = z.object(baseLineShape);

/** Историческая строка счёта: расширенные ставки НДС, явная сумма и сумма налога. */
const historicalInvoiceLine = z
  .object({
    ...baseLineShape,
    quantity: z.number().positive().optional().describe("Количество; при emptyQuantity не задаётся"),
    emptyQuantity: z
      .boolean()
      .optional()
      .describe("Opt-in: количество в источнике пустое; поле вообще не отправляется"),
    vatRate: z.enum(BANK_VAT_RATES).default("БезНДС").describe("Ставка НДС, включая исторические значения"),
    amount: z.number().nonnegative().optional().describe("Явная сумма строки (если не quantity × price)"),
    vatAmount: z.number().nonnegative().optional().describe("Явная сумма НДС строки"),
  })
  .superRefine((line, ctx) => {
    if (line.emptyQuantity) {
      if (line.quantity !== undefined)
        ctx.addIssue({
          code: "custom",
          message: "emptyQuantity нельзя сочетать с quantity: количество должно остаться пустым.",
        });
      if (line.amount === undefined)
        ctx.addIssue({ code: "custom", message: "emptyQuantity требует явную amount." });
      if (line.vatAmount === undefined)
        ctx.addIssue({ code: "custom", message: "emptyQuantity требует явный vatAmount." });
      return;
    }
    if (line.quantity === undefined)
      ctx.addIssue({
        code: "custom",
        message: "quantity обязателен; для пустого количества используйте emptyQuantity=true.",
      });
  });

/** Убирает undefined-поля, чтобы не слать их в 1С. */
function clean(obj: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== undefined && v !== ""));
}

const resolveSet = (conn: Connection, candidates: readonly string[], human: string): Promise<string> =>
  requireEntity(conn, candidates, human);

const OPERATION_MARKER_RE = /\s*\[op:[0-9a-f-]{36}\]/gi;
const operationMarker = (operationId: string) => `[op:${operationId}]`;

/**
 * Дописывает в «Комментарий» метку операции — по ней write.operation.status находит
 * созданный объект после таймаута. Только для сущностей, у которых есть «Комментарий».
 * Метки чужих операций (напр. из копируемого документа) вырезаются.
 */
async function withOperationMarker(
  conn: Connection,
  entitySet: string,
  payload: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const operationId = currentWriteOperationId();
  if (!operationId || !conn.behavior.writeOperationMarker) return payload;
  if (!(await entityHasComment(conn, entitySet))) return payload;
  const prev = typeof payload["Комментарий"] === "string" ? payload["Комментарий"] : "";
  const base = prev.replace(OPERATION_MARKER_RE, "").trim();
  return { ...payload, Комментарий: `${base}${base ? " " : ""}${operationMarker(operationId)}` };
}

async function entityHasComment(conn: Connection, entitySet: string): Promise<boolean> {
  const meta = (await conn.getMetadata()).entities.get(entitySet);
  return !!meta?.properties.some((p) => p.name === "Комментарий");
}

/**
 * Вид договора, который 1С подставляет в документ при выборе контрагента — по направлению
 * документа. Денежные документы сюда не входят: там вид зависит от вида операции.
 */
const CONTRACT_KIND_BY_DOC: ReadonlyMap<string, "СПокупателем" | "СПоставщиком"> = new Map([
  ...[
    ...WRITE_DOCUMENTS.sales,
    ...WRITE_DOCUMENTS.customerInvoice,
    ...WRITE_DOCUMENTS.servicesAct,
    ...WRITE_DOCUMENTS.returnFromCustomer,
  ].map((set) => [set, "СПокупателем"] as const),
  ...[
    ...WRITE_DOCUMENTS.purchases,
    ...WRITE_DOCUMENTS.supplierInvoice,
    ...WRITE_DOCUMENTS.returnToSupplier,
  ].map((set) => [set, "СПоставщиком"] as const),
]);

const EMPTY_REF = "00000000-0000-0000-0000-000000000000";

/**
 * Договор не указан — поступаем как форма 1С при выборе контрагента: если у контрагента ровно один
 * договор нужного вида с этой организацией, подставляем его. Несколько или ни одного — не угадываем,
 * а предупреждаем: без договора 1С документ не проведёт (в БП 3.0 падает невнятной ошибкой модуля
 * прослеживаемости «Преобразование значения к типу Булево»).
 */
export async function withDefaultContract(
  conn: Connection,
  entitySet: string,
  payload: Record<string, unknown>,
): Promise<{ payload: Record<string, unknown>; notes: string[] }> {
  const kind = CONTRACT_KIND_BY_DOC.get(entitySet);
  const counterparty = payload["Контрагент_Key"];
  if (!kind || typeof counterparty !== "string" || !counterparty || payload["ДоговорКонтрагента_Key"]) {
    return { payload, notes: [] };
  }
  const available = await conn.available();
  const set = CATALOGS.contracts.find((c) => available.has(c));
  if (!set) return { payload, notes: [] };
  const org = typeof payload["Организация_Key"] === "string" ? payload["Организация_Key"] : undefined;
  const found = await conn.client.getCollection<ODataEntity>(
    `${set}${buildQuery({
      select: ["Ref_Key", "Description"],
      filter: and(
        cmp("Owner_Key", "eq", odataGuid(counterparty)),
        cmp("ВидДоговора", "eq", odataString(kind)),
        cmp("DeletionMark", "eq", "false"),
        org ? cmp("Организация_Key", "eq", odataGuid(org)) : undefined,
      ),
      top: 3,
    })}`,
  );
  const kindText = kind === "СПокупателем" ? "с покупателем" : "с поставщиком";
  if (found.value.length === 1) {
    const ctr = found.value[0]!;
    return {
      payload: { ...payload, ДоговорКонтрагента_Key: ctr["Ref_Key"] },
      notes: [
        `Договор не указан — подставлен единственный договор ${kindText}: «${String(ctr["Description"] ?? "")}».`,
      ],
    };
  }
  return {
    payload,
    notes: [
      found.value.length === 0
        ? `У контрагента нет договора ${kindText} — документ создастся, но 1С его не проведёт. ` +
          "Создайте договор (write.catalog.create_contract; вид уточните у пользователя) и передайте contractRef."
        : `У контрагента несколько договоров ${kindText} — укажите нужный в contractRef, ` +
          "иначе 1С документ не проведёт.",
    ],
  };
}

const SETTLE_FIELD = "СчетУчетаРасчетовСКонтрагентом_Key";
const ADVANCE_FIELD = "СчетУчетаРасчетовПоАвансам_Key";
const PAYMENT_ROWS = "РасшифровкаПлатежа";
const MONEY_DOCUMENTS: ReadonlySet<string> = new Set([
  ...WRITE_DOCUMENTS.bankIn,
  ...WRITE_DOCUMENTS.bankOut,
  ...WRITE_DOCUMENTS.cashIn,
  ...WRITE_DOCUMENTS.cashOut,
]);
const LEVEL_TEXT = {
  договор: "настройки договора",
  контрагент: "настройки контрагента",
  группа: "настройки группы",
  организация: "настройки организации",
  общая: "общей настройки",
} as const;

/** Сторона расчётов документа: по виду документа, для денежных — по виду операции. */
function documentSettlementSide(
  entitySet: string,
  payload: Record<string, unknown>,
): SettlementSide | undefined {
  const kind = CONTRACT_KIND_BY_DOC.get(entitySet);
  if (kind) return kind === "СПокупателем" ? "customer" : "supplier";
  const op = typeof payload["ВидОперации"] === "string" ? payload["ВидОперации"] : "";
  if (/Покупател/i.test(op)) return "customer";
  if (/Поставщик/i.test(op)) return "supplier";
  return undefined;
}

/** В строгом режиме defaults нужны лишь после более приоритетной настройки расчётов. */
async function strictSettlementDefaults(
  conn: Connection,
  payload: Record<string, unknown>,
  side: SettlementSide | undefined,
  hasHeader: boolean,
  choice?: { settle?: string | undefined; advance?: string | undefined },
): Promise<{ payload: Record<string, unknown>; used: boolean }> {
  if (!conn.cfg?.accounting?.strictDefaults || !side) return { payload, used: false };
  const settleGroup = side === "customer" ? "receivable" : "payable";
  const advanceGroup = side === "customer" ? "receivableAdvance" : "payableAdvance";
  const rows = Array.isArray(payload[PAYMENT_ROWS])
    ? (payload[PAYMENT_ROWS] as Record<string, unknown>[])
    : [];
  const matching = rows.filter((r) => r["ДоговорКонтрагента_Key"] === payload["ДоговорКонтрагента_Key"]);
  const missing = new Set<AccountDefault>();
  if (
    !choice?.settle &&
    ((hasHeader && !refOrUndefined(payload[SETTLE_FIELD])) ||
      matching.some((r) => !refOrUndefined(r[SETTLE_FIELD])))
  )
    missing.add(settleGroup);
  if (!choice?.advance && matching.some((r) => !refOrUndefined(r[ADVANCE_FIELD]))) missing.add(advanceGroup);
  const defaults = await resolveDefaultAccounts(conn, [...missing]);
  const fill = (target: Record<string, unknown>, advance: boolean) =>
    clean({
      ...target,
      ...(!choice?.settle && !refOrUndefined(target[SETTLE_FIELD])
        ? { [SETTLE_FIELD]: defaults.get(settleGroup) }
        : {}),
      ...(advance && !choice?.advance && !refOrUndefined(target[ADVANCE_FIELD])
        ? { [ADVANCE_FIELD]: defaults.get(advanceGroup) }
        : {}),
    });
  return {
    payload: {
      ...(hasHeader ? fill(payload, false) : payload),
      ...(rows.length ? { [PAYMENT_ROWS]: rows.map((r) => (matching.includes(r) ? fill(r, true) : r)) } : {}),
    },
    used: missing.size > 0,
  };
}

/**
 * Счета расчётов и авансов из регистра «Счета учёта расчётов с контрагентами» (#38). 1С через OData
 * настройку договора не применяет, поэтому подставляем сами: самая частная запись по договору →
 * контрагенту → организации с учётом вида расчётов. Регистр не опубликован или подходящей записи нет —
 * остаются счета по умолчанию, и предпросмотр говорит об этом прямо.
 */
export async function withSettlementAccounts(
  conn: Connection,
  entitySet: string,
  payload: Record<string, unknown>,
  side: SettlementSide | undefined = documentSettlementSide(entitySet, payload),
): Promise<{ payload: Record<string, unknown>; notes: string[] }> {
  // Для оплаты по договору счета нужны в расшифровке, а не в шапке (#39).
  // Явную шапку сохраняем; отсутствие поля в payload не заменяем наличием в metadata.
  const rowSettlement =
    MONEY_DOCUMENTS.has(entitySet) &&
    side !== undefined &&
    Array.isArray(payload[PAYMENT_ROWS]) &&
    payload[PAYMENT_ROWS].length > 0;
  const contractRef = payload["ДоговорКонтрагента_Key"];
  if (typeof contractRef !== "string" || !contractRef || contractRef === EMPTY_REF) {
    if (!conn.cfg?.accounting?.strictDefaults) return { payload, notes: [] };
    const hasHeader =
      !rowSettlement &&
      ((await conn.getMetadata()).entities.get(entitySet)?.properties.some((p) => p.name === SETTLE_FIELD) ??
        false);
    const fallback = await strictSettlementDefaults(conn, payload, side, hasHeader);
    return {
      payload: fallback.payload,
      notes: fallback.used
        ? ["Счёт расчётов — по настроенным значениям по умолчанию; договор не выбран."]
        : [],
    };
  }
  const props = new Set(
    (await conn.getMetadata()).entities.get(entitySet)?.properties.map((p) => p.name) ?? [],
  );
  let rows = Array.isArray(payload[PAYMENT_ROWS]) ? (payload[PAYMENT_ROWS] as Record<string, unknown>[]) : [];
  const hasHeader = !rowSettlement && props.has(SETTLE_FIELD);
  if (!hasHeader && rows.length === 0) return { payload, notes: [] };

  const codesOf = async (...refs: Array<unknown>) => {
    const keys = refs.filter((r): r is string => typeof r === "string" && !!r);
    const accounts = await accountsByRef(conn, keys);
    return keys.map((k) => accounts.get(k)?.code ?? "?").join(" / ");
  };

  const orgKey = typeof payload["Организация_Key"] === "string" ? payload["Организация_Key"] : undefined;
  const found = await settlementAccountsFor(conn, {
    orgKey,
    contractRef,
    side,
  });
  const choice = found.published && found.side ? found.choice : undefined;
  const fallback = await strictSettlementDefaults(conn, payload, found.side, hasHeader, choice);
  payload = fallback.payload;
  if (Array.isArray(payload[PAYMENT_ROWS])) rows = payload[PAYMENT_ROWS] as Record<string, unknown>[];
  const current = [
    payload[SETTLE_FIELD] ?? rows[0]?.[SETTLE_FIELD],
    payload[ADVANCE_FIELD] ?? rows[0]?.[ADVANCE_FIELD],
  ];
  const notes: string[] = [];
  if (choice) {
    notes.push(
      `Счета расчётов ${await codesOf(choice.settle, choice.advance)} — из ${LEVEL_TEXT[choice.level]} ` +
        "(регистр «Счета учёта расчётов с контрагентами»).",
      `Источник каждого счёта (правило MCP): ${JSON.stringify(choice.provenance)}`,
    );
  } else {
    notes.push(
      `Счета расчётов — по умолчанию (${await codesOf(...current)}): ` +
        (!found.published
          ? "регистр «Счета учёта расчётов с контрагентами» не опубликован в OData, поэтому особая настройка договора не учтена."
          : found.side
            ? "в регистре «Счета учёта расчётов с контрагентами» нет подходящей записи."
            : "вид договора не «с покупателем» и не «с поставщиком» — настройка не подбирается."),
    );
  }
  if (choice && fallback.used)
    notes.push("Недостающие в настройке счета расчётов/авансов — по настроенным значениям по умолчанию.");
  const apply = (target: Record<string, unknown>, settleKey: boolean, advanceKey: boolean) =>
    choice
      ? {
          ...target,
          ...(settleKey && choice.settle ? { [SETTLE_FIELD]: choice.settle } : {}),
          ...(advanceKey && choice.advance ? { [ADVANCE_FIELD]: choice.advance } : {}),
        }
      : target;

  // Сумма взаиморасчётов в расшифровке платежа: без неё 1С проводит оплату молча и без проводок
  // (проверено на живой базе). Для договора в рублях она равна сумме платежа при курсе 1; для валюты
  // и у.е. курс не угадываем — говорим прямо.
  const rowProps = new Set(
    (await conn.getMetadata()).entities.get(`${entitySet}_${PAYMENT_ROWS}`)?.properties.map((p) => p.name) ??
      [],
  );
  const regulated = found.kind === "РасчетыВВалютеРегламентированногоУчета";
  const withSettlementSum = (r: Record<string, unknown>) =>
    regulated && rowProps.has("СуммаВзаиморасчетов") && r["СуммаВзаиморасчетов"] === undefined
      ? {
          ...r,
          СуммаВзаиморасчетов: r["СуммаПлатежа"],
          ...(rowProps.has("КурсВзаиморасчетов") ? { КурсВзаиморасчетов: 1 } : {}),
          ...(rowProps.has("КратностьВзаиморасчетов") ? { КратностьВзаиморасчетов: 1 } : {}),
        }
      : r;

  let next: Record<string, unknown> = apply(payload, hasHeader, props.has(ADVANCE_FIELD) && !rows.length);
  if (rows.length) {
    next = {
      ...next,
      [PAYMENT_ROWS]: rows.map((r) =>
        r["ДоговорКонтрагента_Key"] === contractRef ? withSettlementSum(apply(r, true, true)) : r,
      ),
    };
    if (found.kind && !regulated) {
      notes.push(
        "Договор в иностранной валюте или у.е.: сумму взаиморасчётов и курс MCP не заполняет — укажите их " +
          "в 1С до проведения, иначе 1С проведёт оплату без проводок по расчётам.",
      );
    }
  }
  return { payload: next, notes };
}

/** После записи: счёт расчётов пуст — 1С проведёт документ без проводок по расчётам, молча. */
export function emptySettlementWarning(
  created: ODataEntity,
  payload: Record<string, unknown>,
): string | undefined {
  // Без договора расчётов с контрагентом нет (налог, зарплата, прочее списание) — пустой счёт законен.
  const contract = payload["ДоговорКонтрагента_Key"];
  if (typeof contract !== "string" || !contract || contract === EMPTY_REF) return undefined;
  const empty = (v: unknown) => v === EMPTY_REF || v === "";
  const rows = Array.isArray(created[PAYMENT_ROWS]) ? (created[PAYMENT_ROWS] as ODataEntity[]) : [];
  const headerEmpty =
    Object.hasOwn(created, SETTLE_FIELD) && empty(created[SETTLE_FIELD]) && rows.length === 0;
  const rowEmpty = rows.some((r) => Object.hasOwn(r, SETTLE_FIELD) && empty(r[SETTLE_FIELD]));
  const sumEmpty = rows.some(
    (r) =>
      r["ДоговорКонтрагента_Key"] === contract &&
      Object.hasOwn(r, "СуммаВзаиморасчетов") &&
      !r["СуммаВзаиморасчетов"],
  );
  if (!headerEmpty && !rowEmpty && !sumEmpty) return undefined;
  const what = [
    headerEmpty || rowEmpty ? "пуст счёт расчётов с контрагентом" : "",
    sumEmpty ? "не заполнена сумма взаиморасчётов в расшифровке платежа" : "",
  ]
    .filter(Boolean)
    .join(" и ");
  // Подсказка — только к инструменту, который это действительно исправляет (#61): update_entity
  // табличную часть не меняет.
  const how = [
    headerEmpty || rowEmpty ? "счёт расчётов — write.document.apply_bank_settlement_settings" : "",
    sumEmpty ? `сумму взаиморасчётов — ${ALIGN_HINT}` : "",
  ]
    .filter(Boolean)
    .join("; ");
  return (
    `ВНИМАНИЕ: в созданном документе ${what} — при проведении 1С не сформирует проводки по расчётам. ` +
    `Заполните до проведения: ${how}.`
  );
}

/**
 * Перед проведением: документ с контрагентом, но без договора 1С (БП 3.0) не проводит, причём падает
 * невнятной ошибкой модуля прослеживаемости. Говорим причину заранее, понятным текстом.
 */
async function missingContract(
  conn: Connection,
  entitySet: string,
  guid: string,
): Promise<string | undefined> {
  const em = (await conn.getMetadata()).entities.get(entitySet);
  const props = new Set(em?.properties.map((p) => p.name) ?? []);
  if (!props.has("ДоговорКонтрагента_Key") || !props.has("Контрагент_Key")) return undefined;
  const doc = await conn.client.getEntity(
    `${entitySet}(guid'${guid}')${buildQuery({ select: ["Контрагент_Key", "ДоговорКонтрагента_Key"] })}`,
  );
  const empty = (v: unknown) => !v || v === EMPTY_REF;
  if (empty(doc["Контрагент_Key"]) || !empty(doc["ДоговорКонтрагента_Key"])) return undefined;
  return (
    "В документе не заполнен договор с контрагентом — 1С его не проведёт. Укажите договор " +
    '(write.entity.update_entity с {"ДоговорКонтрагента_Key": "<Ref_Key договора>"}) и проведите снова; ' +
    "если договора нет — создайте его (write.catalog.create_contract, вид уточните у пользователя)."
  );
}

/**
 * 1С отвергла значение перечисления («Перечисление 'X' не содержит элемент 'Y'») — дописываем к
 * ошибке допустимые значения из $metadata, чтобы вызывающий мог выбрать верное, а не угадывать.
 */
export async function withEnumValues(conn: Connection, e: unknown): Promise<unknown> {
  if (!(e instanceof ODataError)) return e;
  const m = /Перечисление '([^']+)' не содержит элемент/.exec(e.message);
  if (!m) return e;
  try {
    const xml = await conn.client.getText("$metadata");
    const name = m[1]!.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const body = new RegExp(`<EnumType Name="${name}"[^>]*>([\\s\\S]*?)</EnumType>`).exec(xml)?.[1];
    const values = body ? [...body.matchAll(/Member Name="([^"]+)"/g)].map((x) => x[1]) : [];
    if (!values.length) return e;
    return new ODataError({
      kind: e.kind,
      message: `${e.message} Допустимые значения ${m[1]}: ${values.join(", ")}.`,
      cause: e,
    });
  } catch {
    return e;
  }
}

/**
 * Общий путь создания: при confirm=false возвращает предпросмотр (не пишет в 1С),
 * при confirm=true выполняет POST. Гард записи (READ_ONLY + WRITABLE) — в клиенте.
 */
async function createOrPreview(
  conn: Connection,
  entitySet: string,
  payload: Record<string, unknown>,
  confirm: boolean,
  notes?: string[],
  /** keepSettlement — счета расчётов уже заданы (копия документа): настройку договора не применять. */
  opts?: {
    keepSettlement?: boolean;
    settlementSide?: SettlementSide;
    /** Итоговый payload после договора/счетов/метки — для контрольного чтения конкретного инструмента. */
    prepared?: (payload: Record<string, unknown>) => void;
  },
): Promise<CallToolResult> {
  const contract = await withDefaultContract(conn, entitySet, payload);
  const settlement = opts?.keepSettlement
    ? { payload: contract.payload, notes: [] }
    : await withSettlementAccounts(conn, entitySet, contract.payload, opts?.settlementSide);
  payload = await withOperationMarker(conn, entitySet, settlement.payload);
  opts?.prepared?.(payload);
  const allNotes = [...(notes ?? []), ...contract.notes, ...settlement.notes];
  const extra = allNotes.length ? { notes: allNotes } : {};
  if (!confirm) {
    await conn.client.prepareCreate(entitySet, payload);
    return ok({
      dryRun: true,
      database: conn.cfg.name,
      writableBase: conn.cfg.writable,
      willCreate: entitySet,
      ...(await tryPreviewView(conn, entitySet, payload)),
      payload,
      ...extra,
      note: conn.cfg.writable
        ? "Предпросмотр. Чтобы создать запись, повторите вызов с confirm=true."
        : `Предпросмотр. ВНИМАНИЕ: запись в базу "${conn.cfg.name}" сейчас запрещена — включите ODATA_DB_${conn.cfg.name.toUpperCase()}_WRITABLE=true (и READ_ONLY=false).`,
    });
  }
  let created: ODataEntity;
  try {
    created = await conn.client.create<ODataEntity>(entitySet, payload);
  } catch (e) {
    throw await withEnumValues(conn, e);
  }
  const warning = emptySettlementWarning(created, payload);
  if (warning) allNotes.push(warning);
  const createdExtra = allNotes.length ? { notes: allNotes } : {};
  return ok({
    created: true,
    ...(created["_operation_replayed"] === true
      ? {
          replayed: true,
          note: "Эта запись уже была создана с данным operationId; повторный POST не выполнялся.",
        }
      : {}),
    database: conn.cfg.name,
    entitySet,
    ref: created["Ref_Key"],
    code: created["Code"],
    description: created["Description"],
    ...createdExtra,
  });
}

/** Первая строка «Расшифровки платежа» из прочитанного документа. */
function rowValue(doc: ODataEntity, field: string): unknown {
  const rows = doc["РасшифровкаПлатежа"];
  const row = Array.isArray(rows) ? (rows[0] as Record<string, unknown> | undefined) : undefined;
  return row?.[field];
}

/** Дополняет successful-результат инструмента, не теряя operationId и текстовый блок. */
function withStructured(result: CallToolResult, extra: Record<string, unknown>): CallToolResult {
  const structuredContent = { ...(result.structuredContent ?? {}), ...extra };
  return {
    ...result,
    content: result.content.map((item) =>
      item.type === "text" ? { ...item, text: renderJson(structuredContent) } : item,
    ),
    structuredContent,
  };
}

/** Скалярные поля шапки из fields, опубликованные у объекта, с их типами из $metadata. */
async function scalarFields(
  conn: Connection,
  entitySet: string,
  fields: Record<string, unknown>,
): Promise<Map<string, string>> {
  const types = new Map(
    ((await conn.getMetadata()).entities.get(entitySet)?.properties ?? []).map((p) => [p.name, p.type]),
  );
  return new Map(
    Object.keys(fields)
      .filter((k) => !Array.isArray(fields[k]) && types.has(k))
      .map((k) => [k, types.get(k)!]),
  );
}

/** Читает из 1С текущие значения полей (и *_Type составных ссылок). Бросает при сбое чтения. */
async function readFields(
  conn: Connection,
  entitySet: string,
  guid: string,
  keys: string[],
): Promise<Record<string, unknown>> {
  const props = new Set(
    ((await conn.getMetadata()).entities.get(entitySet)?.properties ?? []).map((p) => p.name),
  );
  const select = [...keys, ...keys.map((k) => `${k.replace(/_Key$/, "")}_Type`).filter((k) => props.has(k))];
  return conn.client.getEntity(`${entitySet}(guid'${guid}')${buildQuery({ select })}`);
}

/**
 * Текущие значения изменяемых полей шапки — для «было → станет» в предпросмотре. Не прочитались —
 * предпросмотр без изменений, а не отказ.
 */
async function currentValues(
  conn: Connection,
  entitySet: string,
  guid: string,
  fields: Record<string, unknown>,
): Promise<Record<string, unknown> | undefined> {
  const keys = [...(await scalarFields(conn, entitySet, fields)).keys()];
  if (keys.length === 0) return undefined;
  return readFields(conn, entitySet, guid, keys).catch(() => undefined);
}

/**
 * Совпадает ли сохранённое значение с отправленным — с учётом того, как 1С возвращает типы:
 * GUID в любом регистре, дата без времени — как полночь, пустая ссылка и строка — как null.
 */
export function sameStoredValue(sent: unknown, stored: unknown, type: string): boolean {
  const empty = (v: unknown) => v === null || v === undefined || v === "" || v === EMPTY_REF;
  if (empty(sent) || empty(stored)) return empty(sent) && empty(stored);
  if (type === "Edm.Guid") return String(sent).toLowerCase() === String(stored).toLowerCase();
  if (type === "Edm.Boolean") return sent === stored;
  if (/^Edm\.(Double|Decimal|Int16|Int32|Int64|Single|Byte)$/.test(type))
    return Math.abs(Number(sent) - Number(stored)) < 1e-9;
  if (type === "Edm.DateTime") {
    const d = (v: unknown) => (/^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? `${String(v)}T00:00:00` : String(v));
    return d(sent) === d(stored);
  }
  return String(sent).trimEnd() === String(stored).trimEnd();
}

/**
 * Расчётные реквизиты банковского документа: договор и счета расчётов/авансов живут и в шапке, и в
 * строках «Расшифровки платежа», а 1С при записи берёт шапку из строк. Общий PATCH шапки даёт
 * частичную правку без согласования со строками и проводками (#60), поэтому он запрещён — и для
 * проведённого, и для непроведённого документа. Поддержанный путь — apply_bank_settlement_settings.
 */
const BANK_SETTLEMENT_FIELDS = /^(ДоговорКонтрагента_Key|СчетУчетаРасчетов[\wА-Яа-яЁё]*_Key)$/;
const BANK_DOCUMENTS = new Set<string>([...WRITE_DOCUMENTS.bankIn, ...WRITE_DOCUMENTS.bankOut]);

function bankSettlementPatchRefusal(entitySet: string, fields: Record<string, unknown>): string | undefined {
  if (!BANK_DOCUMENTS.has(entitySet)) return undefined;
  const blocked = Object.keys(fields).filter((k) => BANK_SETTLEMENT_FIELDS.test(k));
  if (!blocked.length) return undefined;
  return (
    `Изменение не отправлено: ${blocked.join(", ")} банковского документа общим инструментом не меняются — ` +
    "1С берёт их из строк «Расшифровки платежа», и правка шапки была бы частичной. Порядок: если документ " +
    "проведён — отмените проведение (write.document.post_document с post=false); затем договор и счета — " +
    "write.document.apply_bank_settlement_settings (шапка и строка согласованно, с проверкой результата); " +
    "затем проведите отдельным вызовом."
  );
}

/** Общий путь изменения (PATCH): dry-run при confirm=false, иначе применяет. */
async function patchOrPreview(
  conn: Connection,
  entitySet: string,
  ref: string,
  fields: Record<string, unknown>,
  confirm: boolean,
  notes?: string[],
) {
  const guid = ref.replace(/[{}']/g, "");
  const extra = notes?.length ? { notes } : {};
  // Проверка не зависит от состояния документа — одинаково в предпросмотре и при подтверждении.
  const refusal = bankSettlementPatchRefusal(entitySet, fields);
  if (refusal) return fail(refusal);
  if (!confirm) {
    return ok({
      dryRun: true,
      database: conn.cfg.name,
      willPatch: `${entitySet}(${guid})`,
      ...(await tryPreviewView(conn, entitySet, fields, await currentValues(conn, entitySet, guid, fields))),
      fields,
      ...extra,
      note: conn.cfg.writable
        ? "Предпросмотр изменения. Чтобы применить, повторите с confirm=true."
        : `Предпросмотр. ВНИМАНИЕ: запись в базу "${conn.cfg.name}" запрещена — включите ODATA_DB_${conn.cfg.name.toUpperCase()}_WRITABLE=true (и READ_ONLY=false).`,
    });
  }
  const updated = await conn.client.patch<ODataEntity>(`${entitySet}(guid'${guid}')?$format=json`, fields);
  // Успех — только по фактически сохранённым значениям (#60): 1С может принять PATCH и при записи
  // объекта перезаполнить часть полей сама.
  const types = await scalarFields(conn, entitySet, fields);
  let stored: Record<string, unknown>;
  try {
    stored = types.size ? await readFields(conn, entitySet, guid, [...types.keys()]) : {};
  } catch (e) {
    return fail(
      `1С приняла изменение, но перечитать объект не удалось (${e instanceof Error ? e.message : String(e)}). ` +
        "Сверьте поля в 1С, прежде чем повторять.",
    );
  }
  const unconfirmed = [...types]
    .filter(([k, t]) => !sameStoredValue(fields[k], stored[k], t))
    .map(([k]) => k);
  if (unconfirmed.length) {
    return fail(
      `1С приняла изменение, но сохранила не всё: не совпали поля ${unconfirmed.join(", ")}` +
        `${unconfirmed.length < types.size ? " (остальные сохранены)" : ""}. ` +
        "Скорее всего, 1С перезаполнила их при записи объекта; повтор того же изменения даст тот же результат.",
    );
  }
  // verified — только когда сверено всё отправленное; табличные части этим путём не сверяются.
  const skipped = Object.keys(fields).filter((k) => !types.has(k));
  return ok({
    updated: true,
    verified: skipped.length === 0,
    database: conn.cfg.name,
    entitySet,
    ref: updated["Ref_Key"] ?? guid,
    description: updated["Description"],
    ...extra,
    ...(skipped.length ? { readbackNote: `Не сверялись (табличные части): ${skipped.join(", ")}.` } : {}),
  });
}

/**
 * Правка строк документа, которую нельзя слепо повторять (добавление/удаление строки:
 * инструмент заново читает документ, и повтор добавил бы ещё одну строку или удалил соседнюю).
 * Подтверждение идёт через журнал операций; для сверки после таймаута в журнал пишется число
 * строк до записи и ожидаемое после.
 */
async function patchLinesOrPreview(
  conn: Connection,
  entitySet: string,
  ref: string,
  fields: Record<string, unknown>,
  confirm: boolean,
  check: { section: string; before: number; expected: number },
) {
  const guid = ref.replace(/[{}']/g, "");
  const key = { ref: guid };
  if (!confirm) {
    await conn.client.prepareCreate(entitySet, key, { withRefKey: false });
    return patchOrPreview(conn, entitySet, ref, fields, false);
  }
  const updated = await conn.client.patchOnce<ODataEntity>(
    `${entitySet}(guid'${guid}')?$format=json`,
    fields,
    entitySet,
    key,
    { kind: "lineCount", ref: guid, ...check },
  );
  return ok({
    updated: true,
    ...(updated["_operation_replayed"] === true
      ? {
          replayed: true,
          note: "Эта правка уже была применена с данным operationId; повторный PATCH не выполнялся.",
        }
      : {}),
    database: conn.cfg.name,
    entitySet,
    ref: updated["Ref_Key"] ?? guid,
    description: updated["Description"],
  });
}

/**
 * Создаёт подчинённый объект (банковский счёт / контактное лицо) и, при makeMain,
 * проставляет его основным у владельца. dry-run при confirm=false.
 */
async function createSubordinate(
  conn: Connection,
  set: string,
  payload: Record<string, unknown>,
  confirm: boolean,
  main?: { ownerSet: string; ownerRef: string; field: string },
) {
  if (!confirm) return createOrPreview(conn, set, payload, false);
  const created = await conn.client.create<ODataEntity>(set, payload);
  const ref = String(created["Ref_Key"] ?? "");
  if (main && ref) {
    const g = main.ownerRef.replace(/[{}']/g, "");
    await conn.client.patch(`${main.ownerSet}(guid'${g}')?$format=json`, { [main.field]: ref });
  }
  return ok({
    created: true,
    ...(created["_operation_replayed"] === true
      ? {
          replayed: true,
          note: "Эта запись уже была создана с данным operationId; повторный POST не выполнялся.",
        }
      : {}),
    database: conn.cfg.name,
    entitySet: set,
    ref,
    description: created["Description"],
    ...(main ? { setAsMain: true } : {}),
  });
}

/** Раньше этого срока после попытки операцию нельзя закрыть как неприменённую даже вручную. */
const NOT_APPLIED_MIN_AGE_SEC = 120;

export function registerWriteTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "write.operation.status",
    {
      title: "Статус операции записи",
      description:
        "Состояние операции записи по operationId из предпросмотра. После потери ответа созданный объект " +
        "сверяется по назначенной ссылке (старые операции — по метке), строки документа — по числу строк, " +
        "настройка регистра — по точному составному ключу и отпечатку проверяемых полей. Совпадение фиксирует " +
        "успех без повторной отправки. Для изменения группы, Post/Unpost и неподтверждённого результата " +
        "регистра ответ unverifiable удерживает повтор и требует ручной сверки. Отсутствие объекта само по " +
        "себе не разрешает новый запрос: 1С могла ещё выполнять его, сколько бы времени ни прошло, — ответ " +
        "not_found. Закрыть операцию как not_applied можно только с markNotApplied=true после того, как " +
        "человек сам сверил базу. Новый предпросмотр допустим после not_applied либо rejected с устранённой " +
        "причиной. Ничего не записывает в 1С.",
      inputSchema: z.object({
        database: databaseField,
        operationId: z.string().uuid().describe("operationId из предпросмотра создания"),
        markNotApplied: z
          .boolean()
          .optional()
          .describe(
            "Только по явному решению человека, который сам сверил базу 1С и убедился, что записи нет. " +
              "Сверка всё равно выполняется: найденный объект фиксируется как успех. Не раньше чем через " +
              `${NOT_APPLIED_MIN_AGE_SEC} с после попытки.`,
          ),
      }),
      outputSchema: operationStatusSchema,
    },
    ({ database, operationId, markNotApplied }) =>
      guard("write.operation.status", async () => {
        const conn = ctx.db(database);
        const base = { operationId, database: conn.cfg.name };
        const entry = await conn.client.operationEntry(operationId);
        if (!entry) {
          return ok({
            ...base,
            status: "not_in_journal",
            note: "Операция не найдена в локальном журнале этой базы: подтверждение не выполнялось.",
          });
        }
        const common = {
          ...base,
          entitySet: entry.entitySet,
          // Причина неизвестного исхода или отказа — текст 1С из журнала (#55).
          ...(entry.errorMessage && entry.state !== "succeeded" ? { lastError: entry.errorMessage } : {}),
        };
        if (entry.state === "succeeded") {
          return ok({
            ...common,
            status: "succeeded",
            ref: entry.result?.["Ref_Key"],
            number: entry.result?.["Number"],
            ...(entry.result?.["RecordPath"] ? { recordPath: entry.result["RecordPath"] } : {}),
          });
        }
        if (entry.state === "prepared") {
          return ok({
            ...common,
            status: "prepared",
            note: "Предпросмотр был, подтверждение не отправлялось.",
          });
        }
        if (entry.state === "rejected") {
          return ok({
            ...common,
            status: "rejected",
            note: "1С отклонила запись. Нужен новый предпросмотр после устранения причины отказа.",
          });
        }
        if (entry.check?.kind === "manual")
          return ok({
            ...common,
            status: "unverifiable",
            note: "Исход изменения группы контрагента неизвестен. Сверьте Parent_Key и версию в 1С; повторная отправка удерживается.",
          });
        if (entry.check?.kind === "registerValues") {
          const check = entry.check;
          try {
            const record = await conn.client.getEntity(`${check.recordPath}?$format=json`);
            if (fingerprintRecordFields(record, check.fields) === check.expectedHash) {
              await conn.client.reconcileOperation(operationId, { RecordPath: check.recordPath });
              return ok({
                ...common,
                status: "found_reconciled",
                recordPath: check.recordPath,
                note: "Точная запись регистра и проверяемые ресурсы совпали с ожидаемыми; результат зафиксирован без повторной отправки.",
              });
            }
          } catch {
            // Missing records and failed reads cannot prove non-application of a delayed request.
          }
          return ok({
            ...common,
            status: "unverifiable",
            recordPath: check.recordPath,
            note: "Точная запись регистра отсутствует, отличается или не прочитана. Требуется ручная сверка; повторная отправка удерживается.",
          });
        }
        if (entry.check?.kind === "createdRef") {
          try {
            const actual = await conn.client.getEntity(
              `${entry.entitySet}(${odataGuid(entry.check.ref)})?$format=json`,
            );
            if (String(actual.Ref_Key).toLowerCase() !== entry.check.ref.toLowerCase())
              return ok({
                ...common,
                status: "unverifiable",
                note: "Прочитана другая ссылка; повтор удерживается.",
              });
            await conn.client.reconcileOperation(operationId, { Ref_Key: entry.check.ref });
            return ok({ ...common, status: "found_reconciled", ref: entry.check.ref });
          } catch (error) {
            if (!(error instanceof ODataError && error.kind === "not_found")) throw error;
            return ok({
              ...common,
              status: "unverifiable",
              note: "Группа по назначенной ссылке пока не найдена; запрос мог ещё выполняться. Повторная запись удерживается.",
            });
          }
        }
        if (entry.check?.kind === "documentState")
          return ok({
            ...common,
            status: "unverifiable",
            note: "Исход изменения документа неизвестен. Нужна ручная сверка версии, полей и состояния; повторная запись удерживается.",
          });
        if (entry.entitySet.startsWith("InformationRegister_"))
          return ok({
            ...common,
            status: "unverifiable",
            note: "Исход записи настройки неизвестен. Сверьте точный составной ключ и ресурсы вручную; повторная отправка заблокирована.",
          });
        const ageSec = Math.round((Date.now() - Date.parse(entry.updatedAt)) / 1000);
        // Отсутствие следа записи не доказывает, что запрос не применился: клиент обрывает ожидание по
        // таймауту, а 1С может закончить запись позже. Поэтому операция закрывается как not_applied только
        // по явному решению человека после ручной сверки (issue #52), и не сразу после попытки.
        const absent = async (what: string, ref?: string) => {
          if (markNotApplied && ageSec >= NOT_APPLIED_MIN_AGE_SEC) {
            await conn.client.markOperationNotApplied(operationId);
            return ok({
              ...common,
              status: "not_applied",
              ...(ref ? { ref } : {}),
              note: `${what} Операция закрыта как неприменённая по ручному подтверждению; для новой записи выполните новый предпросмотр.`,
            });
          }
          return ok({
            ...common,
            status: "not_found",
            ...(ref ? { ref } : {}),
            note:
              `${what} С последней попытки ${ageSec} с. ` +
              (markNotApplied
                ? `Закрыть как неприменённую можно не раньше чем через ${NOT_APPLIED_MIN_AGE_SEC} с — повторите позже. `
                : "Это не доказывает, что запись не дошла: 1С могла ещё выполнять запрос. Не создавайте запись заново. ") +
              "Сверьте базу 1С вручную; если записи точно нет, вызовите write.operation.status с markNotApplied=true.",
          });
        };
        // Правка строк документа — сверяем по числу строк до/после записи.
        if (entry.check?.kind === "lineCount") {
          const { ref, before, expected } = entry.check;
          const count = (await getDocInfo(conn, entry.entitySet, ref)).lines.length;
          if (count === expected && count !== before) {
            await conn.client.reconcileOperation(operationId, { Ref_Key: ref });
            return ok({
              ...common,
              status: "found_reconciled",
              ref,
              note: `В документе ${count} строк(и), как и ожидалось после правки: операция зафиксирована как успешная.`,
            });
          }
          if (count === before)
            return absent(`Правка не видна: в документе по-прежнему ${count} строк(и).`, ref);
          return ok({
            ...common,
            status: "unverifiable",
            ref,
            note: `Число строк ${count} не совпадает ни с исходным (${before}), ни с ожидаемым (${expected}): документ меняли. Проверьте его в 1С вручную.`,
          });
        }
        // Создание с назначенным Ref_Key — сверка простым GET по ссылке (справочники тоже).
        if (entry.refKey) {
          let exists: boolean;
          try {
            await conn.client.getEntity(
              `${entry.entitySet}(guid'${entry.refKey}')?$format=json&$select=Ref_Key`,
            );
            exists = true;
          } catch (e) {
            if (!(e instanceof ODataError && e.kind === "not_found")) throw e;
            exists = false;
          }
          if (exists) {
            await conn.client.reconcileOperation(operationId, { Ref_Key: entry.refKey });
            return ok({
              ...common,
              status: "found_reconciled",
              ref: entry.refKey,
              note: "Объект с назначенной ссылкой есть в 1С; операция зафиксирована как успешная. Повторное подтверждение вернёт эту ссылку.",
            });
          }
          return absent("Объекта с назначенной ссылкой в 1С нет.");
        }
        // executing | outcome_unknown, старые операции без Ref_Key — сверяем с 1С по метке
        if (!(await entityHasComment(conn, entry.entitySet))) {
          return ok({
            ...common,
            status: "unverifiable",
            note: "У этого объекта нет «Комментария», метку не поставить. Проверьте в 1С вручную (по наименованию/номеру).",
          });
        }
        // Ищем по «op:<id>» без квадратных скобок метки: в substringof это класс символов (UUID — только hex и «-»).
        // Number есть только у документов; у справочников — Code (иначе 1С отвечает 400).
        const numberField = entry.entitySet.startsWith("Document_") ? "Number" : "Code";
        const found = await conn.client.getCollection<ODataEntity>(
          `${entry.entitySet}?$format=json&$top=2&$select=Ref_Key,${numberField}&$filter=${encodeURIComponent(
            `substringof('op:${operationId}',Комментарий)`,
          )}`,
        );
        const hit = found.value[0];
        if (hit) {
          await conn.client.reconcileOperation(operationId, hit);
          return ok({
            ...common,
            status: "found_reconciled",
            ref: hit["Ref_Key"],
            number: hit[numberField],
            note: "Объект найден по метке; операция зафиксирована как успешная. Повторное подтверждение вернёт эту ссылку.",
          });
        }
        return absent("По метке объект не найден.");
      }),
  );

  server.registerTool(
    "write.counterparty.create_counterparty",
    {
      title: "Создать контрагента",
      description:
        "Создаёт нового контрагента в справочнике. ПО УМОЛЧАНИЮ это предпросмотр (dry-run) — " +
        "сначала покажите пользователю, что будет создано, и только после явного согласия " +
        "вызовите повторно с confirm=true. Запись идёт в боевую базу, поэтому без подтверждения не создавайте.",
      inputSchema: z.object({
        database: databaseField,
        name: z.string().min(1).describe("Наименование контрагента (Description)"),
        inn: z.string().optional().describe("ИНН"),
        kpp: z.string().optional().describe("КПП"),
        fullName: z.string().optional().describe("Полное наименование"),
        legalType: z
          .enum(["ЮридическоеЛицо", "ФизическоеЛицо"])
          .optional()
          .describe("Юридическое или физическое лицо"),
        phone: z.string().optional().describe("Телефон (в контактную информацию)"),
        email: z.string().optional().describe("Email (в контактную информацию)"),
        address: z.string().optional().describe("Адрес текстом (в контактную информацию, юр. адрес)"),
        ogrn: z.string().optional().describe("ОГРН (пишется, если в базе настроен доп.реквизит «ОГРН»)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, name, inn, kpp, fullName, legalType, phone, email, address, ogrn, confirm }) =>
      guard("write.counterparty.create_counterparty", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.counterparties, "Контрагенты");
        const { fields, notes } = await counterpartyExtras(conn, { phone, email, address, ogrn });
        const payload = clean({
          Description: name,
          ИНН: inn,
          КПП: kpp,
          НаименованиеПолное: fullName,
          ЮридическоеФизическоеЛицо: legalType,
          ...fields,
        });
        return createOrPreview(conn, set, payload, confirm, notes);
      }),
  );

  server.registerTool(
    "write.entity.mark_for_deletion",
    {
      title: "Пометить на удаление",
      description:
        "Ставит/снимает пометку на удаление у объекта (как в 1С) — мягкое удаление, " +
        "физически данные не стираются. По умолчанию предпросмотр (dry-run); для применения " +
        "повторите с confirm=true. Полезно, чтобы убрать ошибочно созданную запись.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя объекта, напр. Catalog_Контрагенты"),
        ref: z.string().describe("Ref_Key объекта (GUID)"),
        mark: z.boolean().default(true).describe("true — пометить на удаление, false — снять пометку"),
        confirm: confirmField,
      }),
      outputSchema: markForDeletionResultSchema,
    },
    ({ database, entitySet, ref, mark, confirm }) =>
      guard("write.entity.mark_for_deletion", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const guid = ref.replace(/[{}']/g, "");
        const path = `${entitySet}(guid'${guid}')?$format=json`;
        if (!confirm) {
          return ok({
            dryRun: true,
            database: conn.cfg.name,
            willPatch: `${entitySet}(${guid})`,
            payload: { DeletionMark: mark },
            note: "Предпросмотр. Чтобы применить пометку, повторите с confirm=true.",
          });
        }
        const updated = await conn.client.patch<ODataEntity>(path, { DeletionMark: mark });
        return ok({
          updated: true,
          database: conn.cfg.name,
          ref: updated["Ref_Key"] ?? guid,
          deletionMark: updated["DeletionMark"] ?? mark,
        });
      }),
  );

  server.registerTool(
    "write.catalog.create_nomenclature",
    {
      title: "Создать номенклатуру",
      description:
        "Создаёт новую позицию номенклатуры (товар/услуга). ПО УМОЛЧАНИЮ предпросмотр (dry-run): " +
        "покажите пользователю payload и только после согласия повторите с confirm=true. " +
        "Можно указать артикул, папку (folder — имя группы или Ref_Key) и пометить услугой (isService). " +
        "Исторические opt-in поля: собственный Code, вид номенклатуры, единица, номенклатурная группа " +
        "и ВидСтавкиНДС; при их наличии проверяются публикация, существование целей, дубль Code и readback. " +
        "Запись идёт в боевую базу.",
      inputSchema: z.object({
        database: databaseField,
        name: z.string().min(1).describe("Наименование номенклатуры (Description)"),
        fullName: z.string().optional().describe("Полное наименование"),
        article: z.string().optional().describe("Артикул"),
        code: z.string().trim().min(1).max(100).optional().describe("Исторический Code номенклатуры"),
        folder: z.string().optional().describe("Папка/группа — имя или Ref_Key (для размещения в группе)"),
        groupRef: z
          .string()
          .uuid()
          .optional()
          .describe("Ref_Key номенклатурной группы (НоменклатурнаяГруппа_Key)"),
        serviceKindRef: z
          .string()
          .uuid()
          .optional()
          .describe("Ref_Key вида номенклатуры (ВидНоменклатуры_Key)"),
        unitRef: z.string().uuid().optional().describe("Ref_Key единицы измерения (ЕдиницаИзмерения_Key)"),
        vatKind: z
          .string()
          .trim()
          .min(1)
          .max(50)
          .optional()
          .describe("ВидСтавкиНДС: БезНДС, Нулевая, Общая, ОбщаяРасчетная, Пониженная…"),
        isService: z
          .boolean()
          .default(false)
          .describe("true — услуга (Услуга=true, вид «Услуга»), false — товар"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      name,
      fullName,
      article,
      code,
      folder,
      groupRef,
      serviceKindRef,
      unitRef,
      vatKind,
      isService,
      confirm,
    }) =>
      guard("write.catalog.create_nomenclature", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.nomenclature, "Номенклатура");
        const props = new Map(
          (await conn.getMetadata()).entities.get(set)?.properties.map((p) => [p.name, p.type]) ?? [],
        );
        const historical =
          code !== undefined ||
          groupRef !== undefined ||
          serviceKindRef !== undefined ||
          unitRef !== undefined ||
          vatKind !== undefined;
        const need = (field: string, type: string): void => {
          if (props.get(field) !== type)
            throw new InputError(
              `Схема номенклатуры не публикует ${field} с типом ${type}; историческое поле перенести нельзя.`,
            );
        };
        const existingTarget = async (candidates: readonly string[], ref: string, label: string) => {
          const targetSet = resolveEntity(candidates, await conn.available());
          if (!targetSet) throw new InputError(`${label} не опубликован в OData.`);
          try {
            const target = await conn.client.getEntity(
              `${targetSet}(${odataGuid(ref)})?$format=json&$select=Ref_Key,DeletionMark`,
            );
            if (String(target.Ref_Key).toLowerCase() !== ref.toLowerCase())
              throw new InputError(`${label}: 1С вернула другой объект.`);
            if (target.DeletionMark === true) throw new InputError(`${label} помечен на удаление.`);
          } catch (error) {
            if (error instanceof ODataError && error.kind === "not_found")
              throw new InputError(`${label} с Ref_Key ${ref} не найден.`);
            throw error;
          }
        };
        if (historical) {
          if (code !== undefined) need("Code", "Edm.String");
          if (groupRef !== undefined) need("НоменклатурнаяГруппа_Key", "Edm.Guid");
          if (serviceKindRef !== undefined) need("ВидНоменклатуры_Key", "Edm.Guid");
          if (unitRef !== undefined) need("ЕдиницаИзмерения_Key", "Edm.Guid");
          if (vatKind !== undefined) {
            need("ВидСтавкиНДС", "Edm.String");
            if (!(VAT_KIND_VALUES as readonly string[]).includes(vatKind))
              throw new InputError(
                `ВидСтавкиНДС «${vatKind}» не входит в проверенное перечисление. ` +
                  `Допустимые: ${VAT_KIND_VALUES.join(", ")}.`,
              );
          }
          if (groupRef) await existingTarget(CATALOGS.nomenclatureGroups, groupRef, "Номенклатурная группа");
          if (serviceKindRef)
            await existingTarget(["Catalog_ВидыНоменклатуры"], serviceKindRef, "Вид номенклатуры");
          if (unitRef) {
            const meta = await conn.getMetadata();
            const navigation = meta.entities
              .get("Catalog_Номенклатура")
              ?.navigations.find((n) => n.name === "ЕдиницаИзмерения");
            const unitCandidates = [
              navigation?.toType,
              "Catalog_КлассификаторЕдиницИзмерения",
              "Catalog_ЕдиницыИзмерения",
            ].filter((name): name is string => Boolean(name));
            await existingTarget(unitCandidates, unitRef, "Единица измерения");
          }
          if (code !== undefined) {
            const select = ["Ref_Key", "Code", "Description"];
            let rows: ODataEntity[];
            let truncated: boolean;
            try {
              const read = await fetchAll<ODataEntity>(
                conn.client,
                set,
                { filter: cmp("Code", "eq", odataString(code)), select },
                conn.behavior.pageSize,
                200,
              );
              rows = read.rows;
              truncated = read.truncated;
            } catch (error) {
              if (!(error instanceof ODataError && error.status === 500)) throw error;
              const read = await fetchAll<ODataEntity>(
                conn.client,
                set,
                { select },
                conn.behavior.pageSize,
                5_000,
              );
              rows = read.rows;
              truncated = read.truncated;
            }
            if (truncated)
              throw new InputError(
                "Проверка кода номенклатуры: достигнут предел строк, проход неполный; " +
                  "нельзя утверждать, что код свободен.",
              );
            const conflict = rows.find((r) => String(r["Code"] ?? "").trim() === code);
            if (conflict)
              throw new InputError(
                `Код номенклатуры «${code}» уже занят: ${String(conflict["Description"] ?? "")} ` +
                  `(${String(conflict["Ref_Key"] ?? "")}).`,
              );
          }
        }
        const parentRef = await folderRefOf(conn, set, folder);
        const payload = clean({
          Description: name,
          НаименованиеПолное: fullName,
          Артикул: article,
          Code: code,
          Parent_Key: parentRef,
          НоменклатурнаяГруппа_Key: groupRef,
          ВидНоменклатуры_Key: serviceKindRef,
          ЕдиницаИзмерения_Key: unitRef,
          ВидСтавкиНДС: vatKind,
          ...(isService ? { Услуга: true } : {}),
        });
        const created = await createOrPreview(conn, set, payload, confirm);
        if (!historical || !confirm || created.isError) return created;
        const createdData = created.structuredContent as Record<string, unknown> | undefined;
        const ref = typeof createdData?.ref === "string" ? createdData.ref : "";
        if (!ref) return created;
        let actual: ODataEntity;
        try {
          actual = await conn.client.getEntity(`${set}(${odataGuid(ref)})?$format=json`);
        } catch (error) {
          const reason = error instanceof Error ? error.message : String(error);
          return {
            ...created,
            isError: true,
            _meta: { "odata-1c/errorKind": "unknown" },
            content: [
              {
                type: "text",
                text:
                  `Номенклатура создана (Ref_Key ${ref}), но перечитать её не удалось: ${reason}. ` +
                  "Повторная запись заблокирована; проверьте объект в 1С вручную.",
              },
            ],
            structuredContent: { ...createdData, verification: "failed" },
          };
        }
        const expected: Record<string, unknown> = {
          Description: name,
          ...(fullName !== undefined ? { НаименованиеПолное: fullName } : {}),
          ...(article !== undefined ? { Артикул: article } : {}),
          ...(code !== undefined ? { Code: code } : {}),
          ...(groupRef !== undefined ? { НоменклатурнаяГруппа_Key: groupRef } : {}),
          ...(serviceKindRef !== undefined ? { ВидНоменклатуры_Key: serviceKindRef } : {}),
          ...(unitRef !== undefined ? { ЕдиницаИзмерения_Key: unitRef } : {}),
          ...(vatKind !== undefined ? { ВидСтавкиНДС: vatKind } : {}),
          ...(isService ? { Услуга: true } : {}),
        };
        const mismatches: Array<{ field: string; expected: unknown; actual: unknown }> = [];
        for (const [field, value] of Object.entries(expected))
          if (!sameStoredValue(value, actual[field], props.get(field) ?? "Edm.String"))
            mismatches.push({ field, expected: value, actual: actual[field] });
        if (mismatches.length)
          return {
            ...created,
            isError: true,
            _meta: { "odata-1c/errorKind": "refused" },
            content: [
              {
                type: "text",
                text:
                  `Номенклатура создана (Ref_Key ${ref}), но контрольное чтение не подтвердило поля: ` +
                  `${mismatches.map((m) => m.field).join(", ")}. Повторная запись заблокирована.`,
              },
            ],
            structuredContent: { ...createdData, verification: "failed", mismatches },
          };
        return withStructured(created, {
          verification: "ok",
          verifiedFields: Object.keys(expected),
        });
      }),
  );

  server.registerTool(
    "write.catalog.create_contract",
    {
      title: "Создать договор",
      description:
        "Заводит договор контрагента (подчинённый справочник): владелец — контрагент, " +
        "организация, ВИД договора, номер/дата, валюта/тип цен и (опц.) руководитель/подписант. " +
        "ВАЖНО: вид договора (kind) обязателен и его НЕЛЬЗЯ выбирать самостоятельно — " +
        "уточните у пользователя, т.к. один контрагент может быть и покупателем, и поставщиком. " +
        "Номер по умолчанию «б/н», дата — сегодня. По умолчанию предпросмотр (dry-run); создание — при confirm=true.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key контрагента (владелец договора)"),
        kind: z
          .enum(CONTRACT_KINDS)
          .describe("Вид договора — ОБЯЗАТЕЛЬНО спросить у пользователя, не выбирать самому"),
        number: z.string().optional().describe("Номер договора (по умолчанию «б/н»)"),
        date: z.string().optional().describe("Дата договора YYYY-MM-DD (по умолчанию сегодня)"),
        name: z.string().optional().describe("Наименование договора (по умолчанию «Договор <номер>»)"),
        currency: z
          .string()
          .optional()
          .describe("Валюта взаиморасчётов — код (напр. 643/840) или название (RUB/USD)"),
        priceType: z
          .string()
          .optional()
          .describe("Тип цен — название или код (из справочника Типы цен номенклатуры)"),
        headName: z
          .string()
          .optional()
          .describe("ФИО руководителя/подписанта контрагента (реквизит договора)"),
        headPosition: z
          .string()
          .optional()
          .describe("Должность руководителя контрагента (напр. «Генеральный директор»)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      kind,
      number,
      date,
      name,
      currency,
      priceType,
      headName,
      headPosition,
      confirm,
    }) =>
      guard("write.catalog.create_contract", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.contracts, "Договоры контрагентов");
        const org = await resolveOrg(conn, organization);

        let cur: { ref: string; code?: string } | undefined;
        if (currency)
          cur = await resolveCatalogItem(conn, CATALOGS.currencies, "Справочник «Валюты»", currency);
        const pt = priceType
          ? await resolveCatalogItem(
              conn,
              CATALOGS.priceTypes,
              "Справочник «Типы цен номенклатуры»",
              priceType,
            )
          : undefined;
        if (cur && !cur.code?.trim())
          throw new InputError("Код выбранной валюты неизвестен: нельзя определить валютность договора.");
        const foreign = cur ? cur.code!.trim() !== (conn.behavior.regulatedCurrencyCode ?? "643") : undefined;
        const num = number ?? "б/н";

        const payload = clean({
          Description: name ?? `Договор ${num}`,
          Номер: num,
          Дата: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Owner_Key: counterpartyRef,
          ВидДоговора: kind,
          Организация_Key: org.key,
          ВалютаВзаиморасчетов_Key: cur?.ref,
          ТипЦен_Key: pt?.ref,
          РуководительКонтрагента: headName,
          ДолжностьРуководителяКонтрагента: headPosition,
          ...(foreign !== undefined ? { Валютный: foreign } : {}),
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.sales.create_invoice",
    {
      title: "Создать счёт покупателю",
      description:
        "Создаёт счёт на оплату покупателю (документ) с табличной частью «Товары». " +
        "Документ создаётся НЕПРОВЕДЁННЫМ (черновик) — провести можно вручную в 1С или инструментом post_document. " +
        "Для услуг задавайте content в строке — это «Содержание», именно оно печатается в счёте. " +
        "По умолчанию предпросмотр (dry-run); создание — при confirm=true. " +
        "Исторические opt-in поля: собственный номер и точное время, явная сумма/НДС строки и расширенные ставки НДС; " +
        "при их наличии выполняется точное контрольное чтение. Пустое количество и срок оплаты пока не поддержаны. " +
        "Контрагент и договор — по Ref_Key; позиции — по Ref_Key номенклатуры.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key контрагента-покупателя"),
        contractRef: z.string().optional().describe("Ref_Key договора (для исторического режима обязателен)"),
        date: z.string().optional().describe("Дата документа YYYY-MM-DD (по умолчанию сегодня)"),
        number: z.string().trim().min(1).max(100).optional().describe("Собственный номер документа"),
        time: z
          .string()
          .regex(/^([01]\d|2[0-3]):[0-5]\d:[0-5]\d$/)
          .optional()
          .describe("Точное время документа HH:MM:SS (по умолчанию 00:00:00)"),
        comment: z.string().trim().min(1).max(1000).optional().describe("Явный комментарий шапки"),
        sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
        orgBankAccount: orgBankAccountField,
        // Счёт на оплату проводок не делает — счета учёта в его строках не нужны.
        lines: z.array(historicalInvoiceLine).min(1).describe("Позиции счёта"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      date,
      number,
      time,
      comment,
      sumIncludesVat,
      orgBankAccount,
      lines,
      confirm,
    }) =>
      guard("write.sales.create_invoice", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.customerInvoice,
          "Документ «Счёт на оплату покупателю»",
        );
        const org = await resolveOrg(conn, organization);
        const historical =
          number !== undefined ||
          time !== undefined ||
          comment !== undefined ||
          lines.some(
            (line) =>
              line.emptyQuantity === true ||
              line.amount !== undefined ||
              line.vatAmount !== undefined ||
              !(VAT_RATES as readonly string[]).includes(line.vatRate ?? "БезНДС"),
          );
        const emptyQuantity = lines.some((line) => line.emptyQuantity === true);
        if (historical) {
          if (!contractRef)
            throw new InputError("Исторический счёт требует явный contractRef; договор не подставляется.");
          if (!orgBankAccount || !GUID_RE.test(orgBankAccount))
            throw new InputError("Исторический счёт требует явный Ref_Key банковского счёта организации.");
          await resolveCustomerContract(conn, { orgKey: org.key, counterpartyRef, contractRef });
          const bankSet = resolveEntity(CATALOGS.bankAccounts, await conn.available());
          if (!bankSet) throw new InputError("Справочник «Банковские счета» не опубликован.");
          try {
            const bankDoc = await conn.client.getEntity(
              `${bankSet}(${odataGuid(orgBankAccount)})?$format=json&$select=Ref_Key,Owner`,
            );
            if (
              String(bankDoc["Owner"] ?? "")
                .replace(/[{}]/g, "")
                .toLowerCase() !== org.key.toLowerCase()
            )
              throw new InputError("Исторический счёт: банковский счёт принадлежит не этой организации.");
          } catch (error) {
            if (error instanceof ODataError && error.kind === "not_found")
              throw new InputError(`Банковский счёт ${orgBankAccount} не найден.`);
            throw error;
          }
        }
        const bank = await resolveOrgBankAccount(conn, org.key, orgBankAccount);
        const rows = buildInvoiceRows(lines);
        const netTotal = rowsTotal(rows);
        const vatTotal =
          Math.round(rows.reduce((sum, row) => sum + Number(row["СуммаНДС"] ?? 0), 0) * 100) / 100;
        // Историческая строка с sumIncludesVat=false хранит net-сумму, а итог документа — gross.
        const documentTotal =
          historical && !sumIncludesVat ? Math.round((netTotal + vatTotal) * 100) / 100 : netTotal;
        const now = new Date();
        const day = date ?? odataDate(now).slice(0, 10);
        const documentDate = time ? new Date(`${day}T${time}`) : date ? new Date(`${date}T00:00:00`) : now;
        if (number) {
          const documentNumber = number.trim();
          const select = ["Ref_Key", "Number", "Date", "Организация_Key"];
          const sameNumber = (r: ODataEntity): boolean =>
            String(r["Организация_Key"] ?? "")
              .replace(/[{}]/g, "")
              .toLowerCase() === org.key.toLowerCase() && String(r["Number"] ?? "").trim() === documentNumber;
          let found: ODataEntity[];
          let truncated: boolean;
          try {
            const read = await fetchAll<ODataEntity>(
              conn.client,
              set,
              {
                filter: and(
                  cmp("Number", "eq", odataString(documentNumber)),
                  cmp("Организация_Key", "eq", odataGuid(org.key)),
                ),
                select,
              },
              conn.behavior.pageSize,
              200,
            );
            found = read.rows;
            truncated = read.truncated;
          } catch (error) {
            if (!(error instanceof ODataError && error.status === 500)) throw error;
            const read = await fetchAll<ODataEntity>(
              conn.client,
              set,
              { select },
              conn.behavior.pageSize,
              5_000,
            );
            found = read.rows;
            truncated = read.truncated;
          }
          if (truncated)
            throw new InputError(
              "Проверка номера счёта: достигнут предел строк, проход неполный. " +
                "Нельзя утверждать, что номер свободен; увеличьте ODATA_MAX_ROWS.",
            );
          const year = String(odataDate(documentDate)).slice(0, 4);
          const conflict = found.find((r) => sameNumber(r) && String(r["Date"] ?? "").slice(0, 4) === year);
          if (conflict)
            throw new InputError(
              `Номер «${documentNumber}» уже занят счётом ${String(conflict["Ref_Key"] ?? "")} ` +
                `за ${year} год у этой организации.`,
            );
        }
        const payload = clean({
          Date: odataDate(documentDate),
          ...(number ? { Number: number.trim() } : {}),
          Posted: false,
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          // У счёта на оплату банковский счёт организации зовётся «Структурная единица».
          СтруктурнаяЕдиница_Key: bank,
          ...(comment !== undefined ? { Комментарий: comment } : {}),
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: documentTotal,
          Товары: rows,
        });
        let effectivePayload = payload;
        const created = await createOrPreview(conn, set, payload, confirm, undefined, {
          prepared: (prepared) => {
            effectivePayload = prepared;
          },
        });
        if (!historical || !confirm || created.isError) {
          if (historical && !confirm)
            return withStructured(created, {
              ...(emptyQuantity ? { quantityVerification: "pending_live_validation" } : {}),
              ...(emptyQuantity
                ? {
                    quantityNote:
                      "Количество в источнике пустое; поле не отправлено. Точное представление пустого количества в OData ждёт live-проверки.",
                  }
                : {}),
            });
          return created;
        }
        const createdData = created.structuredContent as Record<string, unknown> | undefined;
        const ref = typeof createdData?.ref === "string" ? createdData.ref : "";
        if (!ref) return created;
        let actual: ODataEntity;
        try {
          actual = await conn.client.getEntity(`${set}(${odataGuid(ref)})?$format=json`);
        } catch (error) {
          const reason = error instanceof Error ? error.message : String(error);
          return {
            ...created,
            isError: true,
            _meta: { "odata-1c/errorKind": "unknown" },
            content: [
              {
                type: "text",
                text:
                  `Счёт создан (Ref_Key ${ref}), но перечитать его не удалось: ${reason}. ` +
                  "Повторная запись заблокирована; проверьте документ в 1С вручную.",
              },
            ],
            structuredContent: { ...createdData, verification: "failed" },
          };
        }
        const mismatches: Array<{ field: string; expected: unknown; actual: unknown }> = [];
        const expectedHeader: Record<string, unknown> = {
          Date: effectivePayload["Date"],
          ...(number ? { Number: effectivePayload["Number"] } : {}),
          ...(comment !== undefined ? { Комментарий: effectivePayload["Комментарий"] } : {}),
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: effectivePayload["ДоговорКонтрагента_Key"],
          СтруктурнаяЕдиница_Key: effectivePayload["СтруктурнаяЕдиница_Key"],
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: effectivePayload["СуммаДокумента"],
        };
        const headerTypes = new Map(
          (await conn.getMetadata()).entities.get(set)?.properties.map((p) => [p.name, p.type]) ?? [],
        );
        for (const [field, value] of Object.entries(expectedHeader))
          if (!sameStoredValue(value, actual[field], headerTypes.get(field) ?? "Edm.String"))
            mismatches.push({ field, expected: value, actual: actual[field] });
        const actualRows = actual["Товары"];
        if (!Array.isArray(actualRows) || actualRows.length !== rows.length)
          mismatches.push({
            field: "Товары",
            expected: `строк: ${rows.length}`,
            actual: Array.isArray(actualRows) ? `строк: ${actualRows.length}` : "нет таблицы",
          });
        else {
          const rowTypes = new Map(
            (await conn.getMetadata()).entities
              .get(`${set}_Товары`)
              ?.properties.map((p) => [p.name, p.type]) ?? [],
          );
          rows.forEach((expected, index) => {
            const actualRow = actualRows[index] as Record<string, unknown>;
            for (const [field, value] of Object.entries(expected)) {
              if (field === "Номенклатура_Type") continue;
              if (!sameStoredValue(value, actualRow[field], rowTypes.get(field) ?? "Edm.String"))
                mismatches.push({
                  field: `Товары[${index + 1}].${field}`,
                  expected: value,
                  actual: actualRow[field],
                });
            }
          });
        }
        if (mismatches.length)
          return {
            ...created,
            isError: true,
            _meta: { "odata-1c/errorKind": "refused" },
            content: [
              {
                type: "text",
                text:
                  `Счёт создан (Ref_Key ${ref}), но контрольное чтение не подтвердило поля: ` +
                  `${mismatches.map((m) => m.field).join(", ")}. Повторная запись заблокирована; ` +
                  "проверьте документ в 1С вручную.",
              },
            ],
            structuredContent: { ...createdData, verification: "failed", mismatches },
          };
        return withStructured(created, {
          verification: "ok",
          ...(emptyQuantity ? { quantityVerification: "pending_live_validation" } : {}),
          ...(emptyQuantity
            ? {
                quantityNote:
                  "Количество не отправлено и не сверялось: представление пустого количества в OData ещё не подтверждено live-проверкой.",
              }
            : {}),
          verifiedFields: [
            ...Object.keys(expectedHeader),
            ...rows.flatMap((row, i) => Object.keys(row).map((field) => `Товары[${i + 1}].${field}`)),
          ],
        });
      }),
  );

  server.registerTool(
    "write.document.post_document",
    {
      title: "Провести / отменить проведение документа",
      description:
        "Проводит документ (влияет на учёт) или отменяет проведение. По умолчанию предпросмотр (dry-run); " +
        "Для повторного проведения Posted=true задайте repost=true, post=true и expectedDataVersion. " +
        "Версионный Post использует PostingModeOperational=false без Unpost; проверяет состояние документа, не проводки. " +
        "применение — при confirm=true. ВНИМАНИЕ: проведение меняет учётные данные — используйте осознанно.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_СчетНаОплатуПокупателю"),
        ref: z.string().describe("Ref_Key документа (GUID)"),
        expectedDataVersion: expectedDocumentVersion
          .optional()
          .describe(
            "Точная текущая версия включает durable preview/action/readback. Post/Unpost выполняются отдельно; атомарная CAS действия не подтверждена.",
          ),
        post: z.boolean().default(true).describe("true — провести, false — отменить проведение"),
        repost: z
          .boolean()
          .optional()
          .describe(
            "Явно разрешить повторный Post уже проведённого документа; требует post=true и expectedDataVersion",
          ),
        confirm: confirmField,
      }),
      outputSchema: postDocumentResultSchema,
    },
    ({ database, entitySet, ref, post, repost, confirm, expectedDataVersion }) =>
      guard("write.document.post_document", async () => {
        if (repost && (!post || expectedDataVersion === undefined))
          throw new InputError("repost=true требует post=true и точную expectedDataVersion.");
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        // Суммы расшифровки банковского документа — и в предпросмотре: ошибку видно до подтверждения (#61).
        if (post) {
          const blocker = await bankPostingBlocker(conn, entitySet, ref.replace(/[{}']/g, ""));
          if (blocker) return fail(blocker);
        }
        if (expectedDataVersion !== undefined) {
          if (post && confirm) {
            const missing = await missingContract(conn, entitySet, ref);
            if (missing) return fail(missing);
          }
          return ok(
            await guardedDocumentAction(conn, { entitySet, ref, post, repost, expectedDataVersion }, confirm),
          );
        }
        const guid = ref.replace(/[{}']/g, "");
        const action = post ? "Post" : "Unpost";
        const path = `${entitySet}(guid'${guid}')/${action}?$format=json`;
        if (!confirm) {
          return ok({
            dryRun: true,
            database: conn.cfg.name,
            willCall: `${entitySet}(${guid})/${action}`,
            note: "Предпросмотр. Чтобы применить, повторите с confirm=true.",
          });
        }
        if (post) {
          const missing = await missingContract(conn, entitySet, guid);
          if (missing) return fail(missing);
        }
        // Успех — только по состоянию документа: 1С может ответить 200 и не провести (ошибка в
        // обработке проведения). Повторный Post проведённого документа подтверждает смена DataVersion.
        const statePath = `${entitySet}(guid'${guid}')?$format=json&$select=Posted,DataVersion`;
        const state = () => conn.client.getEntity<{ Posted?: boolean; DataVersion?: string }>(statePath);
        const before = await state();
        const response = await conn.client.action(path);
        let after: Awaited<ReturnType<typeof state>>;
        try {
          after = await state();
        } catch (e) {
          return fail(
            `1С приняла ${action}, но перечитать документ не удалось (${e instanceof Error ? e.message : String(e)}). ` +
              "Проверьте в 1С, проведён ли документ, прежде чем повторять.",
          );
        }
        const applied = post
          ? after.Posted === true && (before.Posted !== true || after.DataVersion !== before.DataVersion)
          : after.Posted === false;
        if (!applied) {
          const body = response === undefined ? "" : JSON.stringify(response).slice(0, 500);
          return fail(
            `1С ответила на ${action} без ошибки, но документ ${post ? "не проведён" : "остался проведённым"} ` +
              `(Posted=${String(after.Posted)}${post && before.Posted === true ? ", версия не изменилась" : ""}). ` +
              "Причину покажет проведение документа в самой 1С." +
              (body ? ` Ответ 1С: ${body}` : ""),
          );
        }
        // Что провёл документ — сразу в ответе: проводки по регистру Хозрасчетный.
        // Сбой чтения проводок не отменяет проведения, только поясняется.
        let postings: Record<string, unknown> = {};
        if (post) {
          try {
            const p = await getDocumentPostings(conn, entitySet, guid);
            postings = {
              postings: {
                count: p.postingsCount,
                debitTotal: p.debitTotal,
                creditTotal: p.creditTotal,
                byCorrespondence: p.byCorrespondence.slice(0, 20),
              },
              ...(p.postingsCount === 0
                ? { postingsNote: "Документ проведён, но проводок по бухгалтерскому учёту не сформировал." }
                : {}),
            };
          } catch (e) {
            postings = {
              postingsNote: `Проведено; проводки прочитать не удалось: ${e instanceof Error ? e.message : String(e)}`,
            };
          }
        }
        return ok({
          done: true,
          verified: true,
          verificationScope: "document_state",
          database: conn.cfg.name,
          ref: guid,
          action,
          ...postings,
        });
      }),
  );

  server.registerTool(
    "write.entity.update_entity",
    {
      title: "Изменить объект (общий)",
      description:
        "Изменяет произвольные поля объекта 1С через PATCH по Ref_Key. Имена полей — технические " +
        '(узнать через describe_entity), напр. {"ИНН":"7701234567"}. По умолчанию предпросмотр (dry-run); ' +
        "применение — при confirm=true. Мощный инструмент: меняйте только те поля, что указали. " +
        "После записи поля перечитываются: успех — только если 1С сохранила всё. Договор и счета расчётов " +
        "банковских документов этим инструментом не меняются — write.document.apply_bank_settlement_settings.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя объекта, напр. Catalog_Контрагенты"),
        ref: z.string().describe("Ref_Key объекта (GUID)"),
        fields: z
          .record(z.string(), z.union([z.string(), z.number(), z.boolean(), z.null()]))
          .describe("Поля для изменения: { техническоеИмя: значение }"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, entitySet, ref, fields, confirm }) =>
      guard("write.entity.update_entity", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        if (Object.keys(fields).length === 0) return fail("Не заданы поля для изменения.");
        return patchOrPreview(conn, entitySet, ref, fields, confirm);
      }),
  );

  server.registerTool(
    "write.counterparty.update_counterparty",
    {
      title: "Изменить контрагента",
      description:
        "Изменяет реквизиты контрагента по Ref_Key (наименование/ИНН/КПП/полное имя). " +
        "По умолчанию предпросмотр (dry-run); применение — при confirm=true. Указывайте только меняемые поля.",
      inputSchema: z.object({
        database: databaseField,
        ref: z.string().describe("Ref_Key контрагента"),
        name: z.string().optional().describe("Новое наименование (Description)"),
        inn: z.string().optional().describe("ИНН"),
        kpp: z.string().optional().describe("КПП"),
        fullName: z.string().optional().describe("Полное наименование"),
        phone: z.string().optional().describe("Телефон (контактная информация)"),
        email: z.string().optional().describe("Email (контактная информация)"),
        address: z.string().optional().describe("Адрес текстом (юр. адрес)"),
        ogrn: z.string().optional().describe("ОГРН (если в базе настроен доп.реквизит)"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, ref, name, inn, kpp, fullName, phone, email, address, ogrn, confirm }) =>
      guard("write.counterparty.update_counterparty", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.counterparties, "Справочник «Контрагенты»");
        const fields: Record<string, unknown> = clean({
          Description: name,
          ИНН: inn,
          КПП: kpp,
          НаименованиеПолное: fullName,
        });
        const notes: string[] = [];
        // Контакты — со слиянием: заменяем только заданные виды, прочие сохраняем.
        if (phone || email || address) {
          const kinds = await contactKindsForCounterparties(conn);
          const guid = ref.replace(/[{}']/g, "");
          const existing =
            ((await conn.client.getEntity(`${set}(guid'${guid}')${buildQuery({})}`))[
              "КонтактнаяИнформация"
            ] as Array<Record<string, unknown>>) ?? [];
          const replacedTypes = new Set(
            [phone && kinds.phone?.тип, email && kinds.email?.тип, address && kinds.address?.тип].filter(
              Boolean,
            ),
          );
          const kept = existing.filter((r) => !replacedTypes.has(String(r["Тип"])));
          const merged = [...kept, ...buildContactRows(kinds, { phone, email, address })].map((r, i) => ({
            ...r,
            LineNumber: i + 1,
          }));
          fields["КонтактнаяИнформация"] = merged;
        }
        if (ogrn) {
          const prop = await resolveAdditionalProperty(conn, "ОГРН");
          if (prop)
            fields["ДополнительныеРеквизиты"] = [{ LineNumber: 1, Свойство_Key: prop, Значение: ogrn }];
          else notes.push("ОГРН не записан: в базе нет доп.реквизита «ОГРН» для контрагентов.");
        }
        if (Object.keys(fields).length === 0) return fail("Не задано ни одного поля для изменения.");
        return patchOrPreview(conn, set, ref, fields, confirm, notes);
      }),
  );

  server.registerTool(
    "write.catalog.update_nomenclature",
    {
      title: "Изменить номенклатуру",
      description:
        "Изменяет реквизиты номенклатуры по Ref_Key (наименование/полное имя/артикул). " +
        "По умолчанию предпросмотр (dry-run); применение — при confirm=true. Указывайте только меняемые поля.",
      inputSchema: z.object({
        database: databaseField,
        ref: z.string().describe("Ref_Key номенклатуры"),
        name: z.string().optional().describe("Новое наименование (Description)"),
        fullName: z.string().optional().describe("Полное наименование"),
        article: z.string().optional().describe("Артикул"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, ref, name, fullName, article, confirm }) =>
      guard("write.catalog.update_nomenclature", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.nomenclature, "Справочник «Номенклатура»");
        const fields = clean({ Description: name, НаименованиеПолное: fullName, Артикул: article });
        if (Object.keys(fields).length === 0) return fail("Не задано ни одного поля для изменения.");
        return patchOrPreview(conn, set, ref, fields, confirm);
      }),
  );

  // Общая схема входов для товарных документов (поступление/реализация).
  const goodsDocCommon = {
    database: databaseField,
    organization: organizationField,
    counterpartyRef: z.string().describe("Ref_Key контрагента"),
    contractRef: z.string().optional().describe("Ref_Key договора (необязательно)"),
    warehouse: z.string().optional().describe("Название, код или Ref_Key склада (если в базе несколько)"),
    date: z.string().optional().describe("Дата документа YYYY-MM-DD (по умолчанию сегодня)"),
    sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
  };

  /** Схема закупки: счета доходов/расходов не применяются (Дт 41/19 Кт 60). */
  const purchaseDocInput = {
    ...goodsDocCommon,
    lines: z.array(purchaseLine).min(1).describe("Позиции документа"),
    confirm: confirmField,
  };

  /** Схема продажи: счета доходов/расходов — на документ и (приоритетнее) на строку. */
  const salesDocInput = {
    ...goodsDocCommon,
    incomeAccount: incomeAccountField.describe(
      "Счёт доходов для ВСЕХ строк документа — код как в 1С (напр. «90.01.2») или Ref_Key. " +
        "Счёт, заданный в строке, важнее. Без указания — из регистра «Счета учёта номенклатуры», иначе 90.01.1.",
    ),
    expenseAccount: expenseAccountField.describe(
      "Счёт расходов для ВСЕХ строк документа — код как в 1С (напр. «90.02.2») или Ref_Key. " +
        "Счёт, заданный в строке, важнее. Без указания — из регистра «Счета учёта номенклатуры», иначе 90.02.1.",
    ),
    lines: z.array(saleLine).min(1).describe("Позиции документа"),
    confirm: confirmField,
  };

  /** Реализация: та же схема плюс банковский счёт организации (у возврата его нет). */
  const shipmentDocInput = { ...salesDocInput, orgBankAccount: orgBankAccountField };

  const historicalPurchaseLine = z.object({
    ...baseLineShape,
    vatRate: z.enum(BANK_VAT_RATES).default("БезНДС").describe("Ставка НДС, включая исторические"),
    amount: z.number().nonnegative().optional().describe("Явная сумма строки без НДС"),
    vatAmount: z.number().nonnegative().optional().describe("Явная сумма НДС строки"),
    sourceLineId: z.string().trim().min(1).max(200).optional().describe("ИдентификаторСтроки источника"),
    inventoryAccountRef: z.string().uuid().optional().describe("Ref_Key счёта учёта товара"),
    vatAccountRef: z.string().uuid().optional().describe("Ref_Key счёта учёта входящего НДС"),
    unitRef: z.string().uuid().optional().describe("Ref_Key единицы измерения (ЕдиницаИзмерения_Key)"),
    coefficient: z.number().positive().optional().describe("Явный коэффициент (Коэффициент)"),
    countryOfOriginRef: z.string().uuid().optional().describe("Ref_Key страны происхождения товара"),
    warehouseEmpty: z.boolean().optional().describe("Opt-in: склад строки пустой (склад только в шапке)"),
  });

  const purchaseAnalytic = z.object({
    ref: z.string().uuid().describe("Ref_Key аналитики"),
    type: z.string().trim().min(1).max(200).describe("Тип ссылки, напр. StandardODATA.Catalog_Контрагенты"),
  });
  const historicalPurchaseServiceLine = z.object({
    ...baseLineShape,
    vatRate: z.enum(BANK_VAT_RATES).default("БезНДС").describe("Ставка НДС, включая исторические"),
    amount: z.number().nonnegative().optional().describe("Явная сумма строки"),
    vatAmount: z.number().nonnegative().optional().describe("Явная сумма НДС строки"),
    sourceLineId: z.string().trim().min(1).max(200).optional().describe("ИдентификаторСтроки источника"),
    costAccountRef: z.string().uuid().optional().describe("Ref_Key счёта затрат БУ"),
    costAccountNuRef: z.string().uuid().optional().describe("Ref_Key счёта затрат НУ"),
    vatAccountRef: z.string().uuid().optional().describe("Ref_Key счёта входящего НДС"),
    analytics: z.array(purchaseAnalytic).max(3).optional().describe("Субконто БУ1..3"),
    analyticsNu: z.array(purchaseAnalytic).max(3).optional().describe("СубконтоНУ1..3"),
  });

  server.registerTool(
    "write.purchase.create_purchase",
    {
      title: "Создать поступление от поставщика",
      description:
        "Создаёт документ «Поступление товаров и услуг» (закупка у поставщика) с табличной частью «Товары». " +
        "Документ НЕПРОВЕДЁННЫЙ; провести — вручную в 1С или post_document (тогда 1С сформирует проводки Дт 41/19 Кт 60; счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "Позиции могут нести content (Содержание строки) — для услуг вроде доставки. " +
        "Исторический opt-in: собственные Number/Date с секундами, входящие номер/дата, поставщик и договор, " +
        "склад в шапке при пустом складе строки, net/VAT/gross, идентификатор строки, счета расчётов/авансов, " +
        "товара и входящего НДС, флаг УПД; с контролем дубля и полным readback. " +
        "По умолчанию dry-run; создание — при confirm=true. Контрагент — поставщик, договор — вида «СПоставщиком».",
      inputSchema: z.object({
        ...goodsDocCommon,
        number: z.string().trim().min(1).max(100).optional().describe("Собственный номер документа"),
        time: z
          .string()
          .regex(/^([01]\d|2[0-3]):[0-5]\d:[0-5]\d$/)
          .optional()
          .describe("Точное время документа HH:MM:SS"),
        comment: z.string().trim().min(1).max(1000).optional().describe("Комментарий шапки"),
        incomingNumber: z.string().trim().min(1).max(100).optional().describe("Номер входящего документа"),
        incomingDate: z.string().optional().describe("Дата входящего документа YYYY-MM-DD"),
        incomingTime: z
          .string()
          .regex(/^([01]\d|2[0-3]):[0-5]\d:[0-5]\d$/)
          .optional()
          .describe("Время входящего документа HH:MM:SS"),
        documentTotal: z.number().nonnegative().optional().describe("Точный gross итог документа"),
        settlementAccountRef: z.string().uuid().optional().describe("Ref_Key счёта расчётов (60.01)"),
        advanceAccountRef: z.string().uuid().optional().describe("Ref_Key счёта авансов (60.02)"),
        universalDocument: z.boolean().optional().describe("ЭтоУниверсальныйДокумент"),
        operationKind: z
          .enum(["Товары", "Услуги", "ПокупкаКомиссия"])
          .optional()
          .describe("Исторический ВидОперации; для Товары+Услуги обязателен явный выбор"),
        allowGenericContractKind: z
          .boolean()
          .optional()
          .describe("Явное бизнес-основание принять договор вида «Прочее»"),
        lines: z.array(historicalPurchaseLine).optional().describe("Товарные позиции документа"),
        serviceLines: z
          .array(historicalPurchaseServiceLine)
          .min(1)
          .optional()
          .describe("Сервисные позиции в отдельной таблице «Услуги»"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      warehouse,
      date,
      number,
      time,
      comment,
      incomingNumber,
      incomingDate,
      incomingTime,
      documentTotal,
      settlementAccountRef,
      advanceAccountRef,
      universalDocument,
      operationKind,
      allowGenericContractKind,
      sumIncludesVat,
      lines,
      serviceLines,
      confirm,
    }) =>
      guard("write.purchase.create_purchase", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.purchases,
          "Документ «Поступление товаров и услуг»",
        );
        const org = await resolveOrg(conn, organization);
        const goodsLines = lines ?? [];
        const historical =
          number !== undefined ||
          time !== undefined ||
          comment !== undefined ||
          incomingNumber !== undefined ||
          incomingDate !== undefined ||
          incomingTime !== undefined ||
          documentTotal !== undefined ||
          settlementAccountRef !== undefined ||
          advanceAccountRef !== undefined ||
          universalDocument !== undefined ||
          operationKind !== undefined ||
          allowGenericContractKind !== undefined ||
          serviceLines !== undefined ||
          goodsLines.some(
            (line) =>
              line.amount !== undefined ||
              line.vatAmount !== undefined ||
              line.sourceLineId !== undefined ||
              line.inventoryAccountRef !== undefined ||
              line.vatAccountRef !== undefined ||
              line.unitRef !== undefined ||
              line.coefficient !== undefined ||
              line.warehouseEmpty === true ||
              !(VAT_RATES as readonly string[]).includes(line.vatRate ?? "БезНДС"),
          );
        if (!historical) {
          const warehouseKey = await resolveWarehouse(conn, warehouse);
          const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, goodsLines, "purchase");
          return createGoodsDoc(
            conn,
            set,
            {
              orgKey: org.key,
              counterpartyRef,
              contractRef,
              warehouseKey,
              date,
              sumIncludesVat,
              lines: goodsLines,
              settlement,
              lineAccountsFor,
            },
            confirm,
          );
        }
        if (!number) throw new InputError("Историческое поступление требует явный number.");
        if (!date)
          throw new InputError(
            "Историческое поступление требует явную date; time без date не подставляет текущую дату.",
          );
        if (!goodsLines.length && !serviceLines?.length)
          throw new InputError("Историческое поступление требует товарные или сервисные строки.");
        if (!contractRef) throw new InputError("Историческое поступление требует явный contractRef.");
        const hasGoods = goodsLines.length > 0;
        const hasServices = (serviceLines?.length ?? 0) > 0;
        const serviceOnly = !hasGoods && hasServices;
        if (hasGoods && hasServices && operationKind === undefined)
          throw new InputError(
            "Для исторической комбинации Товары+Услуги укажите operationKind явно: «ПокупкаКомиссия».",
          );
        const resolvedOperationKind = operationKind ?? (hasGoods ? "Товары" : "Услуги");
        if (hasGoods && hasServices && resolvedOperationKind !== "ПокупкаКомиссия")
          throw new InputError(
            "Комбинация Товары+Услуги совместима только с operationKind=«ПокупкаКомиссия».",
          );
        if (
          hasGoods &&
          !hasServices &&
          resolvedOperationKind !== "Товары" &&
          resolvedOperationKind !== "ПокупкаКомиссия"
        )
          throw new InputError("Товарная таблица совместима с operationKind «Товары»/«ПокупкаКомиссия».");
        if (!hasGoods && hasServices && resolvedOperationKind === "Товары")
          throw new InputError("Для таблицы «Услуги» operationKind «Товары» несовместим.");
        if (!serviceOnly && !warehouse)
          throw new InputError("Историческое поступление требует явный warehouse (Ref_Key).");
        const contract = await resolveCustomerContract(conn, {
          orgKey: org.key,
          counterpartyRef,
          contractRef,
        });
        const contractKind = contract.kind?.trim() ?? "";
        if (!contractKind)
          throw new InputError("Договор поставщика не имеет вида; историческое поступление невозможно.");
        if (contractKind === "СПокупателем")
          throw new InputError("Договор поставщика не может быть вида «СПокупателем»; нужен «СПоставщиком».");
        if (contractKind !== "СПоставщиком") {
          if (contractKind !== "Прочее" || allowGenericContractKind !== true)
            throw new InputError(
              `Договор поставщика вида «${contractKind}» не подходит. ` +
                "Для вида «Прочее» требуется явное allowGenericContractKind=true.",
            );
        }
        const warehouseKey = warehouse
          ? await resolveWarehouse(conn, warehouse)
          : serviceOnly
            ? EMPTY_GUID
            : undefined;
        if (!warehouseKey) throw new InputError("Склад не найден.");
        const props = new Map(
          (await conn.getMetadata()).entities.get(set)?.properties.map((p) => [p.name, p.type]) ?? [],
        );
        const rowProps = new Map(
          (await conn.getMetadata()).entities.get(`${set}_Товары`)?.properties.map((p) => [p.name, p.type]) ??
            [],
        );
        const need = (field: string, type: string): void => {
          if (props.get(field) !== type)
            throw new InputError(`Схема поступления не публикует ${field} с типом ${type}.`);
        };
        const needRow = (field: string, type: string): void => {
          if (rowProps.get(field) !== type)
            throw new InputError(`Схема строки поступления не публикует ${field} с типом ${type}.`);
        };
        for (const [field, type] of [
          ["Number", "Edm.String"],
          ["Date", "Edm.DateTime"],
          ["ВидОперации", "Edm.String"],
          ["Организация_Key", "Edm.Guid"],
          ["Склад_Key", "Edm.Guid"],
          ["Контрагент_Key", "Edm.Guid"],
          ["ДоговорКонтрагента_Key", "Edm.Guid"],
          ["НомерВходящегоДокумента", "Edm.String"],
          ["ДатаВходящегоДокумента", "Edm.DateTime"],
          ["СуммаВключаетНДС", "Edm.Boolean"],
          ["СуммаДокумента", "Edm.Double"],
          ["Posted", "Edm.Boolean"],
          ["DeletionMark", "Edm.Boolean"],
        ] as const)
          need(field, type);
        if (universalDocument !== undefined) need("ЭтоУниверсальныйДокумент", "Edm.Boolean");
        if (settlementAccountRef !== undefined) need("СчетУчетаРасчетовСКонтрагентом_Key", "Edm.Guid");
        if (advanceAccountRef !== undefined) need("СчетУчетаРасчетовПоАвансам_Key", "Edm.Guid");
        for (const [field, type] of [
          ["Номенклатура_Key", "Edm.Guid"],
          ["Количество", "Edm.Double"],
          ["Цена", "Edm.Double"],
          ["Сумма", "Edm.Double"],
          ["СтавкаНДС", "Edm.String"],
          ["СуммаНДС", "Edm.Double"],
          ["СчетУчета_Key", "Edm.Guid"],
          ["СчетУчетаНДС_Key", "Edm.Guid"],
        ] as const)
          needRow(field, type);
        if (goodsLines.some((l) => l.sourceLineId !== undefined))
          needRow("ИдентификаторСтроки", "Edm.String");
        if (goodsLines.some((l) => l.unitRef !== undefined)) needRow("ЕдиницаИзмерения_Key", "Edm.Guid");
        if (goodsLines.some((l) => l.coefficient !== undefined)) needRow("Коэффициент", "Edm.Double");
        if (goodsLines.some((l) => l.content !== undefined)) needRow("Содержание", "Edm.String");
        if (goodsLines.some((l) => l.warehouseEmpty === true)) needRow("Склад_Key", "Edm.Guid");
        if (serviceLines) {
          const serviceProps = new Map(
            (await conn.getMetadata()).entities
              .get(`${set}_Услуги`)
              ?.properties.map((p) => [p.name, p.type]) ?? [],
          );
          for (const [field, type] of [
            ["Номенклатура_Key", "Edm.Guid"],
            ["Количество", "Edm.Double"],
            ["Цена", "Edm.Double"],
            ["Сумма", "Edm.Double"],
            ["СтавкаНДС", "Edm.String"],
            ["СуммаНДС", "Edm.Double"],
            ["СчетЗатрат_Key", "Edm.Guid"],
            ["СчетЗатратНУ_Key", "Edm.Guid"],
            ["СчетУчетаНДС_Key", "Edm.Guid"],
          ] as const)
            if (serviceProps.get(field) !== type)
              throw new InputError(`Схема услуги поступления не публикует ${field} с типом ${type}.`);
          if (
            serviceLines.some((l) => l.sourceLineId !== undefined) &&
            serviceProps.get("ИдентификаторСтроки") !== "Edm.String"
          )
            throw new InputError("Схема услуги поступления не публикует ИдентификаторСтроки.");
          if (serviceLines.some((l) => l.content !== undefined)) {
            if (serviceProps.get("Содержание") !== "Edm.String")
              throw new InputError("Схема услуги поступления не публикует Содержание.");
          }
          const usedAnalytics = serviceLines.flatMap((l) => [
            ...(l.analytics ?? []),
            ...(l.analyticsNu ?? []),
          ]);
          if (usedAnalytics.length) {
            const available = await conn.available();
            for (let index = 1; index <= 3; index++) {
              if (serviceLines.some((l) => (l.analytics?.length ?? 0) >= index)) {
                if (
                  serviceProps.get(`Субконто${index}`) !== "Edm.String" ||
                  serviceProps.get(`Субконто${index}_Type`) !== "Edm.String"
                )
                  throw new InputError(`Схема услуги поступления не публикует Субконто${index} и его тип.`);
              }
              if (serviceLines.some((l) => (l.analyticsNu?.length ?? 0) >= index)) {
                if (
                  serviceProps.get(`СубконтоНУ${index}`) !== "Edm.String" ||
                  serviceProps.get(`СубконтоНУ${index}_Type`) !== "Edm.String"
                )
                  throw new InputError(`Схема услуги поступления не публикует СубконтоНУ${index} и его тип.`);
              }
            }
            for (const item of usedAnalytics) {
              const typeSet = item.type.replace(/^StandardODATA\./, "");
              if (!available.has(typeSet))
                throw new InputError(`Тип аналитики ${item.type} не опубликован в OData.`);
              const found = await conn.client.getEntity(
                `${typeSet}(${odataGuid(item.ref)})?$format=json&$select=Ref_Key,DeletionMark`,
              );
              if (
                String(found["Ref_Key"] ?? "")
                  .replace(/[{}]/g, "")
                  .toLowerCase() !== item.ref.toLowerCase()
              )
                throw new InputError(`Ссылка аналитики ${item.ref} (${item.type}) не найдена.`);
              if (found["DeletionMark"] === true)
                throw new InputError(`Ссылка аналитики ${item.ref} помечена на удаление.`);
            }
          }
          const serviceRefs = serviceLines.flatMap((l) => [
            l.costAccountRef,
            l.costAccountNuRef,
            l.vatAccountRef,
          ]);
          if (serviceRefs.some((r) => r !== undefined))
            await accountsByRef(conn, [...new Set(serviceRefs.filter((r): r is string => Boolean(r)))]);
        }
        const refs = [
          settlementAccountRef,
          advanceAccountRef,
          ...goodsLines.flatMap((l) => [l.inventoryAccountRef, l.vatAccountRef]),
        ].filter((r): r is string => Boolean(r));
        if (refs.length) await accountsByRef(conn, [...new Set(refs)]);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, goodsLines, "purchase");
        const rows = goodsLines.map((l, i) => {
          const defaults = lineAccountsFor(l);
          return clean({
            ...(l.sourceLineId !== undefined ? { ИдентификаторСтроки: l.sourceLineId } : {}),
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            ...(l.content === undefined ? {} : { Содержание: l.content }),
            Количество: l.quantity,
            Цена: l.price,
            Сумма: l.amount ?? lineSum(l),
            СтавкаНДС: l.vatRate,
            ...(l.vatAmount !== undefined ? { СуммаНДС: l.vatAmount } : {}),
            СчетУчета_Key: l.inventoryAccountRef ?? defaults["СчетУчета_Key"],
            СчетУчетаНДС_Key: l.vatAccountRef ?? defaults["СчетУчетаНДС_Key"],
            ...(l.unitRef !== undefined ? { ЕдиницаИзмерения_Key: l.unitRef } : {}),
            ...(l.coefficient !== undefined ? { Коэффициент: l.coefficient } : {}),
            ...(l.countryOfOriginRef !== undefined ? { СтранаПроисхождения_Key: l.countryOfOriginRef } : {}),
            // Пустой склад строки — явный нулевой GUID, чтобы сервер не подставил склад шапки.
            ...(l.warehouseEmpty ? { Склад_Key: EMPTY_GUID } : {}),
          });
        });
        const serviceRows = (serviceLines ?? []).map((l, i) => {
          const analytics = (
            items: typeof l.analytics | undefined,
            prefix: string,
          ): Record<string, unknown> =>
            Object.fromEntries(
              (items ?? []).flatMap((item, index) => [
                [`${prefix}${index + 1}`, item.ref],
                [`${prefix}${index + 1}_Type`, item.type],
              ]),
            );
          return clean({
            ...(l.sourceLineId !== undefined ? { ИдентификаторСтроки: l.sourceLineId } : {}),
            ...analytics(l.analytics, "Субконто"),
            ...analytics(l.analyticsNu, "СубконтоНУ"),
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            ...(l.content === undefined ? {} : { Содержание: l.content }),
            Количество: l.quantity,
            Цена: l.price,
            Сумма: l.amount ?? lineSum(l),
            СтавкаНДС: l.vatRate,
            ...(l.vatAmount !== undefined ? { СуммаНДС: l.vatAmount } : {}),
            ...(l.costAccountRef !== undefined ? { СчетЗатрат_Key: l.costAccountRef } : {}),
            ...(l.costAccountNuRef !== undefined ? { СчетЗатратНУ_Key: l.costAccountNuRef } : {}),
            ...(l.vatAccountRef !== undefined ? { СчетУчетаНДС_Key: l.vatAccountRef } : {}),
          });
        });
        const allRows = [...rows, ...serviceRows];
        const netTotal =
          Math.round(allRows.reduce((sum, row) => sum + Number(row["Сумма"] ?? 0), 0) * 100) / 100;
        const vatTotal =
          Math.round(allRows.reduce((sum, row) => sum + Number(row["СуммаНДС"] ?? 0), 0) * 100) / 100;
        const total =
          documentTotal ?? (sumIncludesVat ? netTotal : Math.round((netTotal + vatTotal) * 100) / 100);
        const now = new Date();
        const day = date ?? odataDate(now).slice(0, 10);
        const documentDate = time ? new Date(`${day}T${time}`) : date ? new Date(`${date}T00:00:00`) : now;
        const incoming = incomingDate
          ? odataDate(new Date(`${incomingDate}T${incomingTime ?? "00:00:00"}`))
          : undefined;
        // Защита от дубля: номер + организация + год.
        {
          const select = ["Ref_Key", "Number", "Date", "Организация_Key"];
          let found: ODataEntity[];
          let truncated: boolean;
          try {
            const read = await fetchAll<ODataEntity>(
              conn.client,
              set,
              {
                filter: and(
                  cmp("Number", "eq", odataString(number)),
                  cmp("Организация_Key", "eq", odataGuid(org.key)),
                ),
                select,
              },
              conn.behavior.pageSize,
              200,
            );
            found = read.rows;
            truncated = read.truncated;
          } catch (error) {
            if (!(error instanceof ODataError && error.status === 500)) throw error;
            const read = await fetchAll<ODataEntity>(
              conn.client,
              set,
              { select },
              conn.behavior.pageSize,
              5_000,
            );
            found = read.rows;
            truncated = read.truncated;
          }
          if (truncated) throw new InputError("Проверка дубля поступления: достигнут предел строк.");
          const year = String(odataDate(documentDate)).slice(0, 4);
          const conflict = found.find(
            (r) =>
              String(r["Number"] ?? "").trim() === number &&
              String(r["Организация_Key"] ?? "")
                .replace(/[{}]/g, "")
                .toLowerCase() === org.key.toLowerCase() &&
              String(r["Date"] ?? "").slice(0, 4) === year,
          );
          if (conflict)
            throw new InputError(
              `Номер «${number}» уже занят документом ${String(conflict["Ref_Key"] ?? "")} за ${year} год.`,
            );
        }
        const payload = clean({
          Date: odataDate(documentDate),
          Number: number.trim(),
          ВидОперации: resolvedOperationKind,
          Posted: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          ...(comment !== undefined ? { Комментарий: comment } : {}),
          ...(incomingNumber !== undefined ? { НомерВходящегоДокумента: incomingNumber } : {}),
          ...(incoming !== undefined ? { ДатаВходящегоДокумента: incoming } : {}),
          ...(settlementAccountRef
            ? { СчетУчетаРасчетовСКонтрагентом_Key: settlementAccountRef }
            : { СчетУчетаРасчетовСКонтрагентом_Key: settlement }),
          ...(advanceAccountRef ? { СчетУчетаРасчетовПоАвансам_Key: advanceAccountRef } : {}),
          ...(universalDocument !== undefined ? { ЭтоУниверсальныйДокумент: universalDocument } : {}),
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: total,
          ...(goodsLines.length ? { Товары: rows } : {}),
          ...(serviceLines ? { Услуги: serviceRows } : {}),
        });
        let effectivePayload = payload;
        const created = await createOrPreview(conn, set, payload, confirm, undefined, {
          prepared: (prepared) => {
            effectivePayload = prepared;
          },
        });
        if (!confirm || created.isError) return created;
        const createdData = created.structuredContent as Record<string, unknown> | undefined;
        const ref = typeof createdData?.ref === "string" ? createdData.ref : "";
        if (!ref) return created;
        let actual: ODataEntity;
        try {
          actual = await conn.client.getEntity(`${set}(${odataGuid(ref)})?$format=json`);
        } catch {
          return {
            ...created,
            isError: true,
            _meta: { "odata-1c/errorKind": "unknown" },
            content: [
              {
                type: "text",
                text: "Поступление создано, но контрольное чтение не удалось; повтор запрещён.",
              },
            ],
            structuredContent: { ...createdData, verification: "failed" },
          };
        }
        const expectedHeader: Record<string, unknown> = {
          Date: effectivePayload["Date"],
          Number: effectivePayload["Number"],
          ВидОперации: resolvedOperationKind,
          Posted: false,
          DeletionMark: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          ...(comment !== undefined ? { Комментарий: effectivePayload["Комментарий"] } : {}),
          ...(incomingNumber !== undefined ? { НомерВходящегоДокумента: incomingNumber } : {}),
          ...(incoming !== undefined ? { ДатаВходящегоДокумента: incoming } : {}),
          ...(settlementAccountRef ? { СчетУчетаРасчетовСКонтрагентом_Key: settlementAccountRef } : {}),
          ...(advanceAccountRef ? { СчетУчетаРасчетовПоАвансам_Key: advanceAccountRef } : {}),
          ...(universalDocument !== undefined ? { ЭтоУниверсальныйДокумент: universalDocument } : {}),
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: total,
        };
        const mismatches: Array<{ field: string; expected: unknown; actual: unknown }> = [];
        const types = new Map(
          (await conn.getMetadata()).entities.get(set)?.properties.map((p) => [p.name, p.type]) ?? [],
        );
        for (const [field, value] of Object.entries(expectedHeader))
          if (!sameStoredValue(value, actual[field], types.get(field) ?? "Edm.String"))
            mismatches.push({ field, expected: value, actual: actual[field] });
        const actualRows = actual["Товары"];
        if (goodsLines.length === 0) {
          // Сервисный документ: таблицы «Товары» нет.
        } else if (!Array.isArray(actualRows) || actualRows.length !== rows.length)
          mismatches.push({
            field: "Товары",
            expected: `строк: ${rows.length}`,
            actual: Array.isArray(actualRows) ? `строк: ${actualRows.length}` : "нет таблицы",
          });
        else {
          const rowTypes = new Map(
            (await conn.getMetadata()).entities
              .get(`${set}_Товары`)
              ?.properties.map((p) => [p.name, p.type]) ?? [],
          );
          rows.forEach((expected, index) => {
            const actualRow = actualRows[index] as Record<string, unknown>;
            for (const [field, value] of Object.entries(expected))
              if (!sameStoredValue(value, actualRow[field], rowTypes.get(field) ?? "Edm.String"))
                mismatches.push({
                  field: `Товары[${index + 1}].${field}`,
                  expected: value,
                  actual: actualRow[field],
                });
            if (
              actualRow["Склад_Key"] != null &&
              String(actualRow["Склад_Key"]).replace(/[{}]/g, "") !== EMPTY_GUID
            )
              mismatches.push({
                field: `Товары[${index + 1}].Склад_Key`,
                expected: null,
                actual: actualRow["Склад_Key"],
              });
          });
        }
        if (serviceLines) {
          const actualServices = actual["Услуги"];
          if (!Array.isArray(actualServices) || actualServices.length !== serviceRows.length)
            mismatches.push({
              field: "Услуги",
              expected: `строк: ${serviceRows.length}`,
              actual: Array.isArray(actualServices) ? `строк: ${actualServices.length}` : "нет таблицы",
            });
          else {
            const typesService = new Map(
              (await conn.getMetadata()).entities
                .get(`${set}_Услуги`)
                ?.properties.map((p) => [p.name, p.type]) ?? [],
            );
            serviceRows.forEach((expected, index) => {
              const actualRow = actualServices[index] as Record<string, unknown>;
              for (const [field, value] of Object.entries(expected))
                if (!sameStoredValue(value, actualRow[field], typesService.get(field) ?? "Edm.String"))
                  mismatches.push({
                    field: `Услуги[${index + 1}].${field}`,
                    expected: value,
                    actual: actualRow[field],
                  });
            });
          }
        }
        if (mismatches.length)
          return {
            ...created,
            isError: true,
            _meta: { "odata-1c/errorKind": "refused" },
            content: [
              {
                type: "text",
                text: `Поступление создано (Ref_Key ${ref}), но контрольное чтение не подтвердило: ${mismatches.map((m) => m.field).join(", ")}. Повтор запрещён.`,
              },
            ],
            structuredContent: { ...createdData, verification: "failed", mismatches },
          };
        return withStructured(created, {
          verification: "ok",
          verifiedFields: [
            ...Object.keys(expectedHeader),
            ...rows.flatMap((row, i) => Object.keys(row).map((field) => `Товары[${i + 1}].${field}`)),
            ...serviceRows.flatMap((row, i) => Object.keys(row).map((field) => `Услуги[${i + 1}].${field}`)),
          ],
        });
      }),
  );

  server.registerTool(
    "write.purchase.create_supplier_invoice",
    {
      title: "Зарегистрировать счёт от поставщика",
      description:
        "Создаёт документ «Счёт на оплату поставщика» — регистрирует входящий счёт-основание для оплаты " +
        "(подходит для предоплаты, когда товар ещё не пришёл). Проводок НЕ делает (счёт не проводится), " +
        "счета учёта не нужны. Табличная часть «Товары» вмещает и товары, и услуги (напр. доставку — задайте " +
        "content/Содержание). Реквизиты входящего счёта: incomingNumber/incomingDate. " +
        "По умолчанию dry-run; создание — при confirm=true. Контрагент — поставщик, договор — вида «СПоставщиком».",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key контрагента-поставщика"),
        contractRef: z.string().optional().describe("Ref_Key договора (вид «СПоставщиком»)"),
        supplierBankAccountRef: z
          .string()
          .optional()
          .describe("Ref_Key банковского счёта поставщика (БанковскийСчетКонтрагента)"),
        incomingNumber: z.string().optional().describe("Номер входящего счёта поставщика"),
        incomingDate: z.string().optional().describe("Дата входящего счёта YYYY-MM-DD"),
        date: z.string().optional().describe("Дата документа YYYY-MM-DD (по умолчанию сегодня)"),
        comment: z.string().optional().describe("Комментарий к документу"),
        sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
        lines: z.array(purchaseLine).min(1).describe("Позиции счёта (товары и услуги в одной таблице)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      supplierBankAccountRef,
      incomingNumber,
      incomingDate,
      date,
      comment,
      sumIncludesVat,
      lines,
      confirm,
    }) =>
      guard("write.purchase.create_supplier_invoice", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.supplierInvoice,
          "Документ «Счёт на оплату поставщика»",
        );
        const org = await resolveOrg(conn, organization);
        const rows = buildGoodsRows(lines, noAccounts);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          БанковскийСчетКонтрагента_Key: supplierBankAccountRef,
          НомерВходящегоДокумента: incomingNumber,
          ДатаВходящегоДокумента: incomingDate ? odataDate(new Date(`${incomingDate}T00:00:00`)) : undefined,
          Комментарий: comment,
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: rowsTotal(rows),
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.money.create_payout_order",
    {
      title: "Создать платёжное поручение (исходящее) контрагенту",
      description:
        "Создаёт документ «Платёжное поручение» — исходящую выплату контрагенту (напр. оплата " +
        "по агентскому договору / услугам, зеркало платёжки, уже ушедшей в банк). Документ ВСЕГДА " +
        "создаётся НЕПРОВЕДЁННЫМ (Posted=false) — эта операция намеренно не проводит документ, " +
        "проведение делайте отдельно через post_document, если вообще нужно. Номер (number) — " +
        "задаётся явно вызывающим (обычно чтобы совпасть с номером уже отправленной банковской " +
        "платёжки), не генерируется автоматически. По умолчанию dry-run (confirm=false); создание — " +
        "при confirm=true. Счёт организации и валюта — резолвятся автоматически, если не заданы.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key контрагента-получателя (из find_counterparty)"),
        counterpartyAccountRef: z
          .string()
          .optional()
          .describe("Ref_Key банковского счёта контрагента (СчетКонтрагента_Key)"),
        contractRef: z
          .string()
          .optional()
          .describe("Ref_Key договора с контрагентом (ДоговорКонтрагента_Key)"),
        amount: z.number().positive().describe("Сумма документа"),
        number: z
          .string()
          .describe(
            "Номер документа — задайте явно (напр. чтобы совпасть с номером банковской платёжки). " +
              "1С автонумерует новые документы только если Number не передан вовсе — но здесь он обязателен.",
          ),
        date: z.string().optional().describe("Дата документа YYYY-MM-DD (по умолчанию сегодня)"),
        purposeText: z.string().describe("Назначение платежа (НазначениеПлатежа)"),
        recipientText: z.string().describe("Текст получателя — как печатается в платёжке (ТекстПолучателя)"),
        recipientInn: z.string().optional().describe("ИНН получателя"),
        recipientKpp: z.string().optional().describe("КПП получателя"),
        payerText: z.string().optional().describe("Текст плательщика (по умолчанию — имя организации)"),
        payerInn: z.string().optional().describe("ИНН плательщика (по умолчанию — ИНН организации)"),
        orgBankAccount: orgBankAccountField,
        responsibleRef: z.string().optional().describe("Ref_Key ответственного (Ответственный_Key)"),
        cashflowItemRef: z
          .string()
          .optional()
          .describe("Ref_Key статьи движения денежных средств (СтатьяДвиженияДенежныхСредств_Key)"),
        currencyRef: z.string().optional().describe("Ref_Key валюты документа; без указания — рубль (643)"),
        vatRate: z.enum(VAT_RATES).default("БезНДС").describe("Ставка НДС"),
        operationKind: z
          .string()
          .default("ОплатаПоставщику")
          .describe("Вид операции (ВидОперации), напр. ОплатаПоставщику"),
        priority: z.number().int().default(5).describe("Очередность платежа (ОчередностьПлатежа)"),
        comment: z.string().optional().describe("Комментарий к документу"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      counterpartyAccountRef,
      contractRef,
      amount,
      number,
      date,
      purposeText,
      recipientText,
      recipientInn,
      recipientKpp,
      payerText,
      payerInn,
      orgBankAccount,
      responsibleRef,
      cashflowItemRef,
      currencyRef,
      vatRate,
      operationKind,
      priority,
      comment,
      confirm,
    }) =>
      guard("write.money.create_payout_order", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, WRITE_DOCUMENTS.paymentOrder, "Документ «Платёжное поручение»");
        const org = await resolveOrgOrDefault(conn, organization);
        const [resolvedOrgAccount, resolvedCurrency] = await Promise.all([
          // Через общий резолвер: принимает название/номер счёта, а не только Ref_Key.
          resolveOrgBankAccount(conn, org.ref, orgBankAccount),
          currencyRef ? Promise.resolve(currencyRef) : resolveDefaultCurrency(conn),
        ]);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Number: number,
          Организация_Key: org.ref,
          СчетОрганизации_Key: resolvedOrgAccount,
          Контрагент: counterpartyRef,
          Контрагент_Type: COUNTERPARTY_TYPE,
          СчетКонтрагента_Key: counterpartyAccountRef,
          ДоговорКонтрагента_Key: contractRef,
          СуммаДокумента: amount,
          СтавкаНДС: vatRate,
          ВидОперации: operationKind,
          ОчередностьПлатежа: priority,
          НазначениеПлатежа: purposeText,
          ТекстПолучателя: recipientText,
          ИННПолучателя: recipientInn,
          КПППолучателя: recipientKpp,
          ТекстПлательщика: payerText ?? org.name,
          ИННПлательщика: payerInn ?? org.inn,
          Ответственный_Key: responsibleRef,
          СтатьяДвиженияДенежныхСредств_Key: cashflowItemRef,
          ВалютаДокумента_Key: resolvedCurrency,
          Комментарий: comment,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.sales.create_shipment",
    {
      title: "Создать реализацию покупателю",
      description:
        "Создаёт документ «Реализация товаров и услуг» (отгрузка покупателю) с табличной частью «Товары». " +
        "Документ НЕПРОВЕДЁННЫЙ; провести — вручную в 1С или post_document (тогда 1С сформирует проводки Дт 62 Кт 90, Дт 90 Кт 41 и др.; счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "Позиции могут нести content (Содержание строки — печатается в УПД) и свои счета доходов/расходов. " +
        "По умолчанию dry-run; создание — при confirm=true. Контрагент — покупатель, договор — вида «СПокупателем».",
      inputSchema: z.object(shipmentDocInput),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      warehouse,
      date,
      sumIncludesVat,
      incomeAccount,
      expenseAccount,
      orgBankAccount,
      lines,
      confirm,
    }) =>
      guard("write.sales.create_shipment", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, WRITE_DOCUMENTS.sales, "Документ «Реализация товаров и услуг»");
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const bank = await resolveOrgBankAccount(conn, org.key, orgBankAccount);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "shipment", {
          income: incomeAccount,
          expense: expenseAccount,
        });
        return createGoodsDoc(
          conn,
          set,
          {
            orgKey: org.key,
            counterpartyRef,
            contractRef,
            warehouseKey,
            date,
            sumIncludesVat,
            lines,
            settlement,
            bankAccountKey: bank,
            lineAccountsFor,
          },
          confirm,
        );
      }),
  );

  server.registerTool(
    "write.document.update_document_lines",
    {
      title: "Изменить строки документа",
      description:
        "Заменяет табличную часть «Товары» существующего документа новым набором строк и пересчитывает сумму. " +
        "Поддержаны: счёт покупателю, поступление, реализация (и товарами, и услугами), акт об оказании услуг. " +
        "Строки пишутся в ту табличную часть, которая у документа заполнена («Товары» или «Услуги»). " +
        "Документ должен быть НЕПРОВЕДЁННЫМ " +
        "(если проведён — сначала отмените проведение через post_document). По умолчанию dry-run; применение — при confirm=true.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_РеализацияТоваровУслуг"),
        ref: z.string().describe("Ref_Key документа"),
        lines: z
          .array(saleLine)
          .min(1)
          .describe(
            "Новый полный набор строк (заменяет прежние — реквизиты старых строк не сохраняются, " +
              "задавайте content и счета заново)",
          ),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, entitySet, ref, lines, confirm }) =>
      guard("write.document.update_document_lines", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const info = await getDocInfo(conn, entitySet, ref.replace(/[{}']/g, ""));
        if (info.posted) {
          return fail(
            "Документ проведён. Сначала отмените проведение (post_document с post=false), затем меняйте строки.",
          );
        }
        const built = await buildSectionRows(conn, entitySet, info.orgKey, lines, info.section);
        if (!built) {
          return fail(
            "Поддерживаются: счёт покупателю, поступление/реализация товаров и услуг, акт об оказании " +
              "услуг, возвраты, перемещение, оприходование, списание товаров.",
          );
        }
        return patchOrPreview(
          conn,
          entitySet,
          ref,
          // Вторая ТЧ остаётся как есть, но в сумму документа входит.
          { [info.section]: built.rows, СуммаДокумента: roundMoney(built.total + info.otherTotal) },
          confirm,
        );
      }),
  );

  // === Фаза 1: товарные складские документы ===

  const whLines = z
    .array(
      z.object({
        nomenclatureRef: z.string().describe("Ref_Key номенклатуры"),
        quantity: z.number().positive().describe("Количество"),
        price: z.number().nonnegative().default(0).describe("Цена за единицу (для оценки)"),
      }),
    )
    .min(1)
    .describe("Позиции документа");

  const toGoodsLines = (
    lines: Array<{ nomenclatureRef: string; quantity: number; price: number }>,
  ): GoodsLine[] => lines.map((l) => ({ ...l, vatRate: "БезНДС" }));

  server.registerTool(
    "write.warehouse.create_return_from_customer",
    {
      title: "Создать возврат товаров от покупателя",
      description:
        "Создаёт документ «Возврат товаров от покупателя» с табличной частью «Товары». НЕПРОВЕДЁННЫЙ; " +
        "провести — post_document (проводки сторнируют реализацию: Дт 90.02 Кт 41, Дт 62 Кт 90.01 со знаком минус; счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "dry-run/confirm. Контрагент — покупатель.",
      inputSchema: z.object(salesDocInput),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      warehouse,
      date,
      incomeAccount,
      expenseAccount,
      lines,
      confirm,
    }) =>
      guard("write.warehouse.create_return_from_customer", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.returnFromCustomer,
          "Документ «Возврат товаров от покупателя»",
        );
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "shipment", {
          income: incomeAccount,
          expense: expenseAccount,
        });
        const rows = buildGoodsRows(lines, lineAccountsFor);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          СчетУчетаРасчетовСКонтрагентом_Key: settlement,
          СуммаДокумента: rowsTotal(rows),
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.warehouse.create_return_to_supplier",
    {
      title: "Создать возврат товаров поставщику",
      description:
        "Создаёт документ «Возврат товаров поставщику» с табличной частью «Товары». НЕПРОВЕДЁННЫЙ; " +
        "провести — post_document (проводки сторнируют поступление: Дт 60 Кт 41/19; счёт расчётов — по настройке договора, источник виден в предпросмотре). dry-run/confirm. " +
        "Контрагент — поставщик.",
      inputSchema: z.object(purchaseDocInput),
      outputSchema: createResultSchema,
    },
    ({ database, organization, counterpartyRef, contractRef, warehouse, date, lines, confirm }) =>
      guard("write.warehouse.create_return_to_supplier", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.returnToSupplier,
          "Документ «Возврат товаров поставщику»",
        );
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "purchase");
        const rows = buildGoodsRows(lines, lineAccountsFor);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          СчетУчетаРасчетовСКонтрагентом_Key: settlement,
          СуммаДокумента: rowsTotal(rows),
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.warehouse.create_transfer",
    {
      title: "Создать перемещение товаров",
      description:
        "Создаёт документ «Перемещение товаров» между складами. НЕПРОВЕДЁННЫЙ; провести — post_document " +
        "(проводки Дт 41(склад-получатель) Кт 41(склад-отправитель)). dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        fromWarehouse: z.string().describe("Склад-отправитель: название, код или Ref_Key"),
        toWarehouse: z.string().describe("Склад-получатель: название, код или Ref_Key"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        lines: whLines,
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, organization, fromWarehouse, toWarehouse, date, lines, confirm }) =>
      guard("write.warehouse.create_transfer", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, WRITE_DOCUMENTS.transfer, "Документ «Перемещение товаров»");
        const org = await resolveOrg(conn, organization);
        const [fromKey, toKey] = await Promise.all([
          resolveWarehouse(conn, fromWarehouse),
          resolveWarehouse(conn, toWarehouse),
        ]);
        const lineAcc = await goodsOnlyAccounts(conn, org.key, toGoodsLines(lines));
        const rows = lines.map((l, i) =>
          clean({
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            Количество: l.quantity,
            Цена: l.price,
            ...lineAcc(l.nomenclatureRef),
          }),
        );
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          СкладОтправитель_Key: fromKey,
          СкладПолучатель_Key: toKey,
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.warehouse.create_surplus",
    {
      title: "Создать оприходование товаров",
      description:
        "Создаёт документ «Оприходование товаров» (излишки). НЕПРОВЕДЁННЫЙ; провести — post_document " +
        "(проводки Дт 41 Кт 91.01). Можно указать документ-основание — инвентаризацию. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        warehouse: z.string().optional().describe("Название, код или Ref_Key склада (если несколько)"),
        inventoryRef: z.string().optional().describe("Ref_Key инвентаризации-основания (необязательно)"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        lines: whLines,
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, organization, warehouse, inventoryRef, date, lines, confirm }) =>
      guard("write.warehouse.create_surplus", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, WRITE_DOCUMENTS.surplus, "Документ «Оприходование товаров»");
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const lineAcc = await goodsOnlyAccounts(conn, org.key, toGoodsLines(lines));
        const rows = lines.map((l, i) =>
          clean({
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            Количество: l.quantity,
            Цена: l.price,
            Сумма: Math.round(l.quantity * l.price * 100) / 100,
            ...lineAcc(l.nomenclatureRef),
          }),
        );
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          ИнвентаризацияТоваровНаСкладе_Key: inventoryRef,
          СуммаДокумента: rowsTotal(rows),
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.warehouse.create_writeoff",
    {
      title: "Создать списание товаров",
      description:
        "Создаёт документ «Списание товаров» (недостача/порча). НЕПРОВЕДЁННЫЙ; провести — post_document " +
        "(проводки Дт 94 Кт 41). Можно указать документ-основание — инвентаризацию. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        warehouse: z.string().optional().describe("Название, код или Ref_Key склада (если несколько)"),
        inventoryRef: z.string().optional().describe("Ref_Key инвентаризации-основания (необязательно)"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        lines: whLines,
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, organization, warehouse, inventoryRef, date, lines, confirm }) =>
      guard("write.warehouse.create_writeoff", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, WRITE_DOCUMENTS.writeoff, "Документ «Списание товаров»");
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const lineAcc = await goodsOnlyAccounts(conn, org.key, toGoodsLines(lines));
        const rows = lines.map((l, i) =>
          clean({
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            Количество: l.quantity,
            Цена: l.price,
            Сумма: Math.round(l.quantity * l.price * 100) / 100,
            ...lineAcc(l.nomenclatureRef),
          }),
        );
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          ИнвентаризацияТоваровНаСкладе_Key: inventoryRef,
          СуммаДокумента: rowsTotal(rows),
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.warehouse.create_inventory",
    {
      title: "Создать инвентаризацию товаров",
      description:
        "Создаёт документ «Инвентаризация товаров на складе» (сверка факт/учёт). Документ НЕ проводится " +
        "(движений не делает) — служит основанием для оприходования/списания. dry-run/confirm. " +
        "Для каждой позиции: фактическое и учётное количество.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        warehouse: z.string().optional().describe("Название, код или Ref_Key склада (если несколько)"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        lines: z
          .array(
            z.object({
              nomenclatureRef: z.string().describe("Ref_Key номенклатуры"),
              factQuantity: z.number().nonnegative().describe("Фактическое количество"),
              accountingQuantity: z.number().nonnegative().describe("Количество по учёту"),
              price: z.number().nonnegative().default(0).describe("Цена (для оценки)"),
            }),
          )
          .min(1)
          .describe("Позиции инвентаризации"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, organization, warehouse, date, lines, confirm }) =>
      guard("write.warehouse.create_inventory", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.inventory,
          "Документ «Инвентаризация товаров на складе»",
        );
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const lineAcc = await goodsOnlyAccounts(
          conn,
          org.key,
          lines.map((l) => ({
            nomenclatureRef: l.nomenclatureRef,
            quantity: l.factQuantity,
            price: l.price,
            vatRate: "БезНДС",
          })),
        );
        const rows = lines.map((l, i) =>
          clean({
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            Количество: l.factQuantity,
            КоличествоУчет: l.accountingQuantity,
            Цена: l.price,
            Сумма: Math.round(l.factQuantity * l.price * 100) / 100,
            СуммаУчет: Math.round(l.accountingQuantity * l.price * 100) / 100,
            ...lineAcc(l.nomenclatureRef),
          }),
        );
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  const lineObject = saleLine;

  server.registerTool(
    "write.document.add_document_line",
    {
      title: "Добавить строку в документ",
      description:
        "Добавляет одну позицию в табличную часть «Товары» существующего НЕПРОВЕДЁННОГО документа " +
        "(счёт/поступление/реализация), сохраняя прежние строки со всеми их реквизитами " +
        "(содержание, счета учёта, номенклатурная группа). dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_РеализацияТоваровУслуг"),
        ref: z.string().describe("Ref_Key документа (GUID)"),
        line: lineObject.describe("Добавляемая позиция"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, entitySet, ref, line, confirm }) =>
      guard("write.document.add_document_line", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const info = await getDocInfo(conn, entitySet, ref.replace(/[{}']/g, ""));
        if (info.posted) return fail("Документ проведён. Сначала отмените проведение, затем меняйте строки.");
        const built = await buildSectionRows(
          conn,
          entitySet,
          info.orgKey,
          [...info.lines, line],
          info.section,
        );
        if (!built)
          return fail(
            "Поддерживаются: счёт, поступление/реализация, акт услуг, возвраты, перемещение, оприходование, списание.",
          );
        return patchLinesOrPreview(
          conn,
          entitySet,
          ref,
          // Вторая ТЧ остаётся как есть, но в сумму документа входит.
          { [info.section]: built.rows, СуммаДокумента: roundMoney(built.total + info.otherTotal) },
          confirm,
          { section: info.section, before: info.lines.length, expected: info.lines.length + 1 },
        );
      }),
  );

  server.registerTool(
    "write.document.copy_document",
    {
      title: "Скопировать документ",
      description:
        "Создаёт новый документ по образцу существующего: читает его целиком и записывает копию " +
        "со ВСЕМИ реквизитами — содержание строк, счета учёта, банковский счёт, ответственный, " +
        "адрес доставки, доп. условия. Это главный способ выставить ежемесячно повторяющийся " +
        "документ: инструменты create_* собирают документ из полей схемы и всё остальное теряют. " +
        "Копия создаётся НЕПРОВЕДЁННОЙ, с новым номером (1С нумерует сама) и датой (по умолчанию сегодня). " +
        "Правки: date — дата; fields — реквизиты шапки техническими именами (как в update_entity); " +
        "lines — правки строк по номеру (напр. количество и содержание). Сумма документа " +
        "пересчитывается по строкам, если не задана явно в fields. Счета расчётов — как в исходном документе; если в fields сменён договор — по настройке нового договора. По умолчанию dry-run; запись — при confirm=true.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_СчетНаОплатуПокупателю"),
        ref: z.string().describe("Ref_Key документа-образца (GUID)"),
        date: z.string().optional().describe("Дата копии YYYY-MM-DD (по умолчанию сегодня)"),
        fields: z
          .record(z.string(), z.union([z.string(), z.number(), z.boolean(), z.null()]))
          .optional()
          .describe(
            'Реквизиты шапки, которые надо изменить: { техническоеИмя: значение }, напр. {"Комментарий":"..."}',
          ),
        lines: z
          .array(
            z.object({
              lineNumber: z.number().int().positive().describe("Номер строки (с 1)"),
              fields: z
                .record(z.string(), z.union([z.string(), z.number(), z.boolean(), z.null()]))
                .describe(
                  'Поля строки: { техническоеИмя: значение }, напр. {"Количество":4.3,"Сумма":10750,"Содержание":"..."}',
                ),
            }),
          )
          .optional()
          .describe("Правки строк табличной части по номеру (остальные строки копируются как есть)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, entitySet, ref, date, fields, lines, confirm }) =>
      guard("write.document.copy_document", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const guid = ref.replace(/[{}']/g, "");
        const source = await conn.client.getEntity(`${entitySet}(guid'${guid}')${buildQuery({})}`);
        const payload = stripForCopy(source);
        payload["Date"] = odataDate(date ? new Date(`${date}T00:00:00`) : new Date());
        payload["Posted"] = false;

        if (lines?.length) {
          const section = pickSection(
            {
              goods: payload["Товары"] as unknown[] | undefined,
              services: payload["Услуги"] as unknown[] | undefined,
            },
            {
              servicesAct: (WRITE_DOCUMENTS.servicesAct as readonly string[]).includes(entitySet),
              operationKind: payload["ВидОперации"],
            },
          );
          const rows = sectionRows(payload, section);
          if (rows.length === 0) return fail(`В документе нет строк табличной части «${section}».`);
          const patched = applyLineOverrides(rows, lines);
          payload[section] = patched;
          const total = totalFromRows(patched);
          // Вторая ТЧ копируется как есть, но в сумму документа входит.
          if (total !== undefined)
            payload["СуммаДокумента"] = roundMoney(
              total + rowsSum(sectionRows(payload, otherSection(section))),
            );
        }
        // Реквизиты шапки применяем последними — явное указание важнее пересчёта.
        Object.assign(payload, fields ?? {});
        // Копия сохраняет счета расчётов исходного документа (бухгалтер мог поставить их вручную).
        // Настройку договора применяем, только если в fields сменили договор и счета не заданы явно.
        const has = (k: string) => !!fields && Object.hasOwn(fields, k);
        const contractChanged =
          has("ДоговорКонтрагента_Key") &&
          fields?.["ДоговорКонтрагента_Key"] !== source["ДоговорКонтрагента_Key"];
        const keepSettlement = !contractChanged || has(SETTLE_FIELD) || has(ADVANCE_FIELD);
        return createOrPreview(
          conn,
          entitySet,
          payload,
          confirm,
          keepSettlement && Object.hasOwn(payload, SETTLE_FIELD)
            ? [
                "Счета расчётов — как в исходном документе: при копировании настройка договора не применяется.",
              ]
            : undefined,
          { keepSettlement },
        );
      }),
  );

  server.registerTool(
    "write.document.remove_document_line",
    {
      title: "Удалить строку документа",
      description:
        "Удаляет строку (по номеру LineNumber, начиная с 1) из табличной части «Товары» существующего " +
        "НЕПРОВЕДЁННОГО документа. Нельзя удалить последнюю строку. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_РеализацияТоваровУслуг"),
        ref: z.string().describe("Ref_Key документа (GUID)"),
        lineNumber: z.number().int().positive().describe("Номер строки (с 1)"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, entitySet, ref, lineNumber, confirm }) =>
      guard("write.document.remove_document_line", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const info = await getDocInfo(conn, entitySet, ref.replace(/[{}']/g, ""));
        if (info.posted) return fail("Документ проведён. Сначала отмените проведение, затем меняйте строки.");
        if (lineNumber > info.lines.length) return fail(`В документе всего ${info.lines.length} строк(и).`);
        const kept = info.lines.filter((_, i) => i + 1 !== lineNumber);
        if (kept.length === 0)
          return fail("Нельзя удалить последнюю строку — в документе должна остаться хотя бы одна позиция.");
        const built = await buildSectionRows(conn, entitySet, info.orgKey, kept, info.section);
        if (!built)
          return fail(
            "Поддерживаются: счёт, поступление/реализация, акт услуг, возвраты, перемещение, оприходование, списание.",
          );
        return patchLinesOrPreview(
          conn,
          entitySet,
          ref,
          // Вторая ТЧ остаётся как есть, но в сумму документа входит.
          { [info.section]: built.rows, СуммаДокумента: roundMoney(built.total + info.otherTotal) },
          confirm,
          { section: info.section, before: info.lines.length, expected: kept.length },
        );
      }),
  );

  server.registerTool(
    "write.sales.create_act",
    {
      title: "Создать акт (реализация услуг)",
      description:
        "Создаёт «Реализация (акт, накладная)» с табличной частью УСЛУГИ (без склада/остатков). " +
        "Документ НЕПРОВЕДЁННЫЙ; проводки при проведении Дт 62 Кт 90.01 (без 41; счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "ВАЖНО для услуг: задавайте content в строке — это «Содержание», текст, который печатается " +
        "в акте и УПД (номер приложения, номер договора, период оказания). 1С его не заполнит. " +
        "Счета доходов/расходов при необходимости задаются явно (напр. 90.01.2 / 90.02.2). " +
        "dry-run/confirm. Контрагент — покупатель; позиции — услуги-номенклатура.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key покупателя"),
        contractRef: z.string().optional().describe("Ref_Key договора"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        operationKind: z
          .string()
          .trim()
          .min(1)
          .max(100)
          .optional()
          .describe(
            "Вид операции (ВидОперации). Обычно НЕ задавайте: по умолчанию «Услуги» или значение из настройки базы " +
              "(ODATA_ACCOUNTING_CONFIG, operationKinds.salesAct). Нужен, только если в конфигурации другое значение — не угадывайте.",
          ),
        sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
        incomeAccount: incomeAccountField.describe(
          "Счёт доходов для ВСЕХ строк (напр. «90.01.2»); счёт в строке важнее",
        ),
        expenseAccount: expenseAccountField.describe(
          "Счёт расходов для ВСЕХ строк (напр. «90.02.2»); счёт в строке важнее",
        ),
        orgBankAccount: orgBankAccountField,
        lines: z.array(lineObject).min(1).describe("Позиции-услуги"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      date,
      operationKind,
      sumIncludesVat,
      incomeAccount,
      expenseAccount,
      orgBankAccount,
      lines,
      confirm,
    }) =>
      guard("write.sales.create_act", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, WRITE_DOCUMENTS.sales, "Документ «Реализация товаров и услуг»");
        const org = await resolveOrg(conn, organization);
        const bank = await resolveOrgBankAccount(conn, org.key, orgBankAccount);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "service", {
          income: incomeAccount,
          expense: expenseAccount,
        });
        const rows = buildGoodsRows(lines, lineAccountsFor);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          ВидОперации: configuredOperationKind(conn.cfg, "salesAct", operationKind),
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          СчетУчетаРасчетовСКонтрагентом_Key: settlement,
          БанковскийСчетОрганизации_Key: bank,
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: rowsTotal(rows),
          Услуги: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.sales.create_services_act",
    {
      title: "Создать акт об оказании услуг",
      description:
        "Создаёт «Акт об оказании производственных услуг» (услуги с учётом доходов/расходов по " +
        "номенклатурной группе). Отличается от create_act тем, что требует номенклатурную группу " +
        "(субконто счёта 90.01) — она НЕ угадывается. Проводки Дт 62 Кт 90.01, НДС 90.03 (счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "НЕПРОВЕДЁННЫЙ; провести — post_document. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key покупателя"),
        contractRef: z.string().optional().describe("Ref_Key договора"),
        nomenclatureGroupRef: z
          .string()
          .describe(
            "Ref_Key номенклатурной группы (ОБЯЗАТЕЛЬНО, субконто счёта доходов 90.01 — не угадывается)",
          ),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
        incomeAccount: incomeAccountField.describe(
          "Счёт доходов для ВСЕХ строк (напр. «90.01.2»); счёт в строке важнее",
        ),
        expenseAccount: expenseAccountField.describe(
          "Счёт расходов для ВСЕХ строк (напр. «90.02.2»); счёт в строке важнее",
        ),
        lines: z.array(lineObject).min(1).describe("Позиции-услуги"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      nomenclatureGroupRef,
      date,
      sumIncludesVat,
      incomeAccount,
      expenseAccount,
      lines,
      confirm,
    }) =>
      guard("write.sales.create_services_act", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.servicesAct,
          "Документ «Акт об оказании производственных услуг»",
        );
        const org = await resolveOrg(conn, organization);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "service", {
          income: incomeAccount,
          expense: expenseAccount,
        });
        const grpSet = resolveEntity(CATALOGS.nomenclatureGroups, await conn.available());
        const grpSubconto = grpSet
          ? { Субконто: nomenclatureGroupRef, Субконто_Type: `StandardODATA.${grpSet}` }
          : {};
        const rows = lines.map((l, i) =>
          clean({
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            Содержание: l.content,
            Количество: l.quantity,
            Цена: l.price,
            Сумма: lineSum(l),
            СтавкаНДС: l.vatRate,
            НоменклатурнаяГруппа_Key: nomenclatureGroupRef,
            ...grpSubconto,
            ...lineAccountsFor(l),
          }),
        );
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          НоменклатурнаяГруппа_Key: nomenclatureGroupRef,
          СчетУчетаРасчетовСКонтрагентом_Key: settlement,
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: rowsTotal(rows),
          Услуги: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.money.create_payment",
    {
      title: "Создать оплату от покупателя",
      description:
        "Создаёт «Поступление на расчётный счёт» (оплата от покупателя) на сумму, по контрагенту и договору. " +
        "Документ НЕПРОВЕДЁННЫЙ; при проведении проводка Дт 51 Кт 62 (счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "Сумма взаиморасчётов для рублёвого договора заполняется сама; для договора в валюте или у.е. " +
        "курс укажите в 1С, иначе 1С проведёт оплату без проводок. " +
        "Банковский счёт организации берётся по названию или первый. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key покупателя"),
        contractRef: z.string().describe("Ref_Key договора"),
        operationKind: z
          .string()
          .trim()
          .min(1)
          .max(100)
          .optional()
          .describe(
            "Вид операции (ВидОперации). Обычно НЕ задавайте: по умолчанию «ОплатаПокупателя» или значение из настройки " +
              "базы (ODATA_ACCOUNTING_CONFIG, operationKinds.customerPayment). Нужен, только если в конфигурации другое значение — не угадывайте.",
          ),
        amount: z.number().positive().describe("Сумма оплаты"),
        orgBankAccount: orgBankAccountField,
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      operationKind,
      amount,
      orgBankAccount,
      date,
      confirm,
    }) =>
      guard("write.money.create_payment", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.bankIn,
          "Документ «Поступление на расчётный счёт»",
        );
        const org = await resolveOrg(conn, organization);
        const bank = await resolveOrgBankAccount(conn, org.key, orgBankAccount);
        const codes = conn.cfg?.accounting?.strictDefaults
          ? new Map<AccountDefault, string | undefined>()
          : await resolveDefaultAccounts(conn, ["receivable", "receivableAdvance"]);
        const settle = codes.get("receivable");
        const advance = codes.get("receivableAdvance");
        const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          ВидОперации: configuredOperationKind(conn.cfg, "customerPayment", operationKind),
          Организация_Key: org.key,
          СчетОрганизации_Key: bank,
          Контрагент: counterpartyRef,
          ...(cpSet ? { Контрагент_Type: `StandardODATA.${cpSet}` } : {}),
          ДоговорКонтрагента_Key: contractRef,
          СуммаДокумента: amount,
          РасшифровкаПлатежа: [
            clean({
              LineNumber: 1,
              ДоговорКонтрагента_Key: contractRef,
              СпособПогашенияЗадолженности: "Автоматически", // иначе платёж не распределяется и нет проводок
              СуммаПлатежа: amount,
              СтавкаНДС: "БезНДС",
              СчетУчетаРасчетовСКонтрагентом_Key: settle,
              СчетУчетаРасчетовПоАвансам_Key: advance,
            }),
          ],
        });
        return createOrPreview(conn, set, payload, confirm, undefined, { settlementSide: "customer" });
      }),
  );

  server.registerTool(
    "write.money.create_payment_from_statement",
    {
      title: "Создать оплату покупателя из строки выписки",
      description:
        "Первый участок загрузчика банковских документов БП. Создаёт «Поступление на расчётный счёт» " +
        "по строке выписки: находит организацию по точным ИНН/КПП, её банковский счёт по номеру, " +
        "контрагента-покупателя по точным ИНН/КПП и договор по номеру (или единственный подходящий). " +
        "Явные поля источника переносятся по схеме: ставка и сумма НДС (включая исторические), " +
        "назначение платежа, номер/дата входящего документа, счёт плательщика, собственный номер/время, " +
        "статья ДДС шапки и строки, счёт учёта банка, комментарий и ОтражениеВУСН. Если поле не " +
        "опубликовано или контрольное чтение после создания не совпало — отказ с диагностикой, " +
        "а не молчаливая потеря данных. Dry-run по умолчанию; для записи передайте operationId " +
        "из предпросмотра и confirm=true. Стабильный sourceKey делает повторный запуск безопасным.",
      inputSchema: z.object({
        database: databaseField,
        sourceKey: z
          .string()
          .trim()
          .min(1)
          .max(200)
          .describe("Стабильный ключ строки источника (например, идентификатор платежа в выписке)"),
        organizationInn: z.string().trim().describe("Точный ИНН организации-получателя"),
        organizationKpp: z.string().trim().optional().describe("Точный КПП организации-получателя"),
        orgBankAccountNumber: z
          .string()
          .trim()
          .describe("Точный номер банковского счёта организации-получателя (20 цифр)"),
        payerInn: z.string().trim().describe("Точный ИНН плательщика (покупателя)"),
        payerKpp: z.string().trim().optional().describe("Точный КПП плательщика"),
        payerName: z.string().trim().max(300).optional().describe("Наименование плательщика в источнике"),
        payerAccountNumber: z
          .string()
          .trim()
          .optional()
          .describe("Точный номер счёта плательщика из выписки (20 цифр), если он есть"),
        amount: z.number().positive().describe("Сумма оплаты"),
        date: z.string().describe("Дата документа YYYY-MM-DD"),
        number: z
          .string()
          .trim()
          .min(1)
          .max(100)
          .optional()
          .describe("Номер самого документа (отдельно от documentNumber входящего поручения)"),
        time: z
          .string()
          .regex(/^([01]\d|2[0-3]):[0-5]\d:[0-5]\d$/)
          .optional()
          .describe("Точное время документа HH:MM:SS (по умолчанию 00:00:00)"),
        purpose: z.string().trim().max(500).optional().describe("Назначение платежа"),
        documentNumber: z
          .string()
          .trim()
          .min(1)
          .max(100)
          .optional()
          .describe("Номер входящего банковского документа"),
        documentDate: z
          .string()
          .regex(/^\d{4}-\d{2}-\d{2}$/)
          .optional()
          .describe("Дата входящего банковского документа YYYY-MM-DD"),
        vatRate: z
          .string()
          .trim()
          .min(1)
          .max(50)
          .optional()
          .describe("Явная ставка НДС из выписки (НДС18, НДС20, БезНДС…); только вместе с vatAmount"),
        vatAmount: z
          .number()
          .min(0)
          .optional()
          .describe("Явная сумма НДС из выписки; только вместе с vatRate"),
        cashFlowItemRef: z
          .string()
          .uuid()
          .optional()
          .describe("Ref_Key статьи ДДС шапки, если она задана явно; проверяется публикация и существование"),
        rowCashFlowItemRef: z
          .string()
          .uuid()
          .optional()
          .describe("Ref_Key статьи ДДС в единственной строке расшифровки; отдельное поле схемы"),
        customerInvoiceRef: z
          .string()
          .uuid()
          .optional()
          .describe(
            "Ref_Key счёта покупателю для строки расшифровки; проверяется публикация и существование",
          ),
        bankLedgerAccountRef: z
          .string()
          .uuid()
          .optional()
          .describe("Ref_Key счёта учёта банка (СчетБанк_Key); проверяется существование в плане счетов"),
        comment: z
          .string()
          .trim()
          .min(1)
          .max(500)
          .optional()
          .describe("Явный непустой комментарий шапки документа"),
        incomeRecognition: z
          .string()
          .trim()
          .min(1)
          .max(50)
          .optional()
          .describe("ОтражениеВУСН: ВозвратРасхода, НеПринимаются, Принимаются, Распределяются"),
        contractNumber: z.string().trim().min(1).max(100).optional().describe("Точный номер договора"),
        contractRef: z.string().uuid().optional().describe("Ref_Key договора, если он известен"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      sourceKey,
      organizationInn,
      organizationKpp,
      orgBankAccountNumber,
      payerInn,
      payerKpp,
      payerName,
      payerAccountNumber,
      amount,
      date,
      number,
      time,
      purpose,
      documentNumber,
      documentDate,
      vatRate,
      vatAmount,
      cashFlowItemRef,
      rowCashFlowItemRef,
      customerInvoiceRef,
      bankLedgerAccountRef,
      comment,
      incomeRecognition,
      contractNumber,
      contractRef,
      confirm,
    }) =>
      guard("write.money.create_payment_from_statement", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.bankIn,
          "Документ «Поступление на расчётный счёт»",
        );
        const meta = await conn.getMetadata();
        const docMeta = meta.entities.get(set);
        if (!docMeta) throw new InputError(`Объект ${set} отсутствует в схеме базы.`);
        const rowSet = `${set}_РасшифровкаПлатежа`;
        const rowMeta = meta.entities.get(rowSet);
        const headerProps = new Map(docMeta.properties.map((p) => [p.name, p.type]));
        const rowProps = new Map((rowMeta?.properties ?? []).map((p) => [p.name, p.type]));
        const need = (props: Map<string, string>, field: string, type: string, where: string): void => {
          const actual = props.get(field);
          if (actual !== type)
            throw new InputError(
              `Схема ${where} не публикует поле ${field} с типом ${type}` +
                (actual ? ` (найдено ${actual})` : "") +
                ". Явное значение из выписки перенести нельзя — запись не начата.",
            );
        };
        // Поля, которые загрузчик заполняет всегда: без них договор/сумма/НДС потеряются молча.
        need(headerProps, "Организация_Key", "Edm.Guid", "документа");
        need(headerProps, "СчетОрганизации_Key", "Edm.Guid", "документа");
        need(headerProps, "Контрагент", "Edm.String", "документа");
        need(headerProps, "Контрагент_Type", "Edm.String", "документа");
        need(headerProps, "ДоговорКонтрагента_Key", "Edm.Guid", "документа");
        need(headerProps, "ВидОперации", "Edm.String", "документа");
        need(headerProps, "СуммаДокумента", "Edm.Double", "документа");
        need(rowProps, "ДоговорКонтрагента_Key", "Edm.Guid", rowSet);
        need(rowProps, "СпособПогашенияЗадолженности", "Edm.String", rowSet);
        need(rowProps, "СуммаПлатежа", "Edm.Double", rowSet);
        need(rowProps, "СтавкаНДС", "Edm.String", rowSet);
        need(rowProps, "СуммаНДС", "Edm.Double", rowSet);
        if (purpose !== undefined) need(headerProps, "НазначениеПлатежа", "Edm.String", "документа");
        if (documentNumber !== undefined)
          need(headerProps, "НомерВходящегоДокумента", "Edm.String", "документа");
        if (documentDate !== undefined)
          need(headerProps, "ДатаВходящегоДокумента", "Edm.DateTime", "документа");
        if (payerAccountNumber !== undefined)
          need(headerProps, "СчетКонтрагента_Key", "Edm.Guid", "документа");
        if (number !== undefined) need(headerProps, "Number", "Edm.String", "документа");
        if (cashFlowItemRef !== undefined)
          need(headerProps, "СтатьяДвиженияДенежныхСредств_Key", "Edm.Guid", "документа");
        if (customerInvoiceRef !== undefined) need(rowProps, "СчетНаОплату_Key", "Edm.Guid", rowSet);
        if (rowCashFlowItemRef !== undefined)
          need(rowProps, "СтатьяДвиженияДенежныхСредств_Key", "Edm.Guid", rowSet);
        if (bankLedgerAccountRef !== undefined) need(headerProps, "СчетБанк_Key", "Edm.Guid", "документа");
        if (comment !== undefined) need(headerProps, "Комментарий", "Edm.String", "документа");
        if (incomeRecognition !== undefined) {
          need(headerProps, "ОтражениеВУСН", "Edm.String", "документа");
          if (!(INCOME_RECOGNITION_VALUES as readonly string[]).includes(incomeRecognition))
            throw new InputError(
              `ОтражениеВУСН «${incomeRecognition}» не входит в проверенное перечисление. ` +
                `Допустимые: ${INCOME_RECOGNITION_VALUES.join(", ")}.`,
            );
        }

        // НДС: явные ставка и сумма — только вместе; сумму не вычисляем из суммы платежа.
        if ((vatRate === undefined) !== (vatAmount === undefined))
          throw new InputError(
            "НДС: укажите и ставку (vatRate), и сумму (vatAmount) из источника — одну величину не угадываем.",
          );
        const explicitVat = vatRate !== undefined;
        const vatRateValue = vatRate ?? "БезНДС";
        const vatAmountValue = Math.round((vatAmount ?? 0) * 100) / 100;
        if (!(BANK_VAT_RATES as readonly string[]).includes(vatRateValue))
          throw new InputError(
            `Ставка НДС «${vatRateValue}» не входит в Enum_СтавкиНДС БП 3.0. Допустимые: ${BANK_VAT_RATES.join(", ")}.`,
          );
        if (vatAmountValue > amount + 0.005)
          throw new InputError(
            `Сумма НДС ${vatAmountValue} больше суммы платежа ${amount}: явные данные выписки противоречивы.`,
          );
        if (vatRateValue === "БезНДС" && vatAmountValue !== 0)
          throw new InputError("Ставка «БезНДС» с ненулевой суммой НДС противоречива; проверьте источник.");

        /** Явная ссылка должна быть опубликована и существовать; иначе отказ до записи. */
        const existingTarget = async (
          candidates: readonly string[],
          ref: string,
          label: string,
        ): Promise<void> => {
          const targetSet = resolveEntity(candidates, await conn.available());
          if (!targetSet)
            throw new InputError(`${label} не опубликован в OData — явную ссылку перенести нельзя.`);
          try {
            const target = await conn.client.getEntity(
              `${targetSet}(${odataGuid(ref)})?$format=json&$select=Ref_Key,DeletionMark`,
            );
            if (String(target.Ref_Key).toLowerCase() !== ref.toLowerCase())
              throw new InputError(`${label}: 1С вернула другой объект.`);
            if (target.DeletionMark === true) throw new InputError(`${label} помечен на удаление.`);
          } catch (error) {
            if (error instanceof ODataError && error.kind === "not_found")
              throw new InputError(`${label} с Ref_Key ${ref} не найден в базе.`);
            throw error;
          }
        };

        // Точная резолюция — до записи: неоднозначность или неполный проход останавливают загрузку.
        const org = await resolveOrganizationByRequisites(conn, {
          inn: organizationInn,
          kpp: organizationKpp,
        });
        const ownNumber = number?.trim();
        if (ownNumber) {
          const select = ["Ref_Key", "Number", "Date", "Организация_Key"];
          const sameNumber = (r: ODataEntity): boolean => {
            const sameOrg =
              String(r["Организация_Key"] ?? "")
                .replace(/[{}]/g, "")
                .toLowerCase() === org.ref.toLowerCase();
            return String(r["Number"] ?? "").trim() === ownNumber && sameOrg;
          };
          let rows: ODataEntity[];
          let truncated: boolean;
          try {
            const read = await fetchAll<ODataEntity>(
              conn.client,
              set,
              {
                filter: and(
                  cmp("Number", "eq", odataString(ownNumber)),
                  cmp("Организация_Key", "eq", odataGuid(org.ref)),
                ),
                select,
              },
              conn.behavior.pageSize,
              200,
            );
            rows = read.rows;
            truncated = read.truncated;
          } catch (error) {
            if (!(error instanceof ODataError && error.status === 500)) throw error;
            // Публикация может запрещать фильтр по Number/Организация_Key: ограниченный проход
            // без фильтра, затем точная локальная сверка номера, года и владельца.
            const read = await fetchAll<ODataEntity>(
              conn.client,
              set,
              { select },
              conn.behavior.pageSize,
              5_000,
            );
            rows = read.rows;
            truncated = read.truncated;
          }
          if (truncated)
            throw new InputError(
              "Проверка номера документа: достигнут предел строк, проход неполный. " +
                "Нельзя утверждать, что номер свободен; увеличьте ODATA_MAX_ROWS.",
            );
          const year = date.slice(0, 4);
          const conflict = rows.find((r) => sameNumber(r) && String(r["Date"] ?? "").slice(0, 4) === year);
          if (conflict)
            throw new InputError(
              `Номер «${ownNumber}» уже занят документом ${String(conflict["Ref_Key"] ?? "")} ` +
                `за ${year} год у этой организации. Выберите другой номер или не задавайте его — 1С пронумерует сама.`,
            );
        }
        if (cashFlowItemRef) await existingTarget(CATALOGS.cashflowItems, cashFlowItemRef, "Статья ДДС");
        if (customerInvoiceRef)
          await existingTarget(WRITE_DOCUMENTS.customerInvoice, customerInvoiceRef, "Счёт покупателю");
        if (rowCashFlowItemRef)
          await existingTarget(CATALOGS.cashflowItems, rowCashFlowItemRef, "Статья ДДС строки");
        if (bankLedgerAccountRef) {
          try {
            await accountsByRef(conn, [bankLedgerAccountRef]);
          } catch {
            throw new InputError(
              `Счёт учёта банка ${bankLedgerAccountRef} не найден в плане счетов «Хозрасчётный».`,
            );
          }
        }
        const bank = await resolveOrgBankAccountByNumber(conn, org.ref, orgBankAccountNumber);
        const counterparty = await resolveCounterpartyByRequisites(conn, {
          inn: payerInn,
          kpp: payerKpp,
        });
        const payerBank =
          payerAccountNumber !== undefined
            ? await resolveCounterpartyBankAccountByNumber(conn, counterparty.ref, payerAccountNumber)
            : undefined;
        const contract = await resolveCustomerContract(conn, {
          orgKey: org.ref,
          counterpartyRef: counterparty.ref,
          ...(contractRef ? { contractRef } : {}),
          ...(contractNumber ? { contractNumber } : {}),
        });
        const codes = conn.cfg?.accounting?.strictDefaults
          ? new Map<AccountDefault, string | undefined>()
          : await resolveDefaultAccounts(conn, ["receivable", "receivableAdvance"]);
        const settle = codes.get("receivable");
        const advance = codes.get("receivableAdvance");
        if (settle) need(rowProps, "СчетУчетаРасчетовСКонтрагентом_Key", "Edm.Guid", rowSet);
        if (advance) need(rowProps, "СчетУчетаРасчетовПоАвансам_Key", "Edm.Guid", rowSet);
        const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
        const normalizedAmount = Math.round(amount * 100) / 100;
        if (normalizedAmount <= 0) throw new InputError("Сумма оплаты должна быть больше нуля.");
        const incomingDate = documentDate ? odataDate(new Date(`${documentDate}T00:00:00`)) : undefined;
        const documentDateTime = odataDate(new Date(`${date}T${time ?? "00:00:00"}`));
        const payload = clean({
          Date: documentDateTime,
          ...(ownNumber ? { Number: ownNumber } : {}),
          Posted: false,
          ВидОперации: configuredOperationKind(conn.cfg, "customerPayment"),
          Организация_Key: org.ref,
          СчетОрганизации_Key: bank.ref,
          Контрагент: counterparty.ref,
          ...(cpSet ? { Контрагент_Type: `StandardODATA.${cpSet}` } : {}),
          ДоговорКонтрагента_Key: contract.ref,
          ...(purpose !== undefined ? { НазначениеПлатежа: purpose } : {}),
          ...(documentNumber !== undefined ? { НомерВходящегоДокумента: documentNumber } : {}),
          ...(incomingDate ? { ДатаВходящегоДокумента: incomingDate } : {}),
          ...(cashFlowItemRef ? { СтатьяДвиженияДенежныхСредств_Key: cashFlowItemRef } : {}),
          ...(bankLedgerAccountRef ? { СчетБанк_Key: bankLedgerAccountRef } : {}),
          ...(comment !== undefined ? { Комментарий: comment } : {}),
          ...(incomeRecognition ? { ОтражениеВУСН: incomeRecognition } : {}),
          ...(payerBank ? { СчетКонтрагента_Key: payerBank.ref } : {}),
          СуммаДокумента: normalizedAmount,
          РасшифровкаПлатежа: [
            clean({
              LineNumber: 1,
              ДоговорКонтрагента_Key: contract.ref,
              СпособПогашенияЗадолженности: "Автоматически",
              СуммаПлатежа: normalizedAmount,
              СтавкаНДС: vatRateValue,
              СуммаНДС: vatAmountValue,
              ...(rowCashFlowItemRef ? { СтатьяДвиженияДенежныхСредств_Key: rowCashFlowItemRef } : {}),
              ...(customerInvoiceRef ? { СчетНаОплату_Key: customerInvoiceRef } : {}),
              СчетУчетаРасчетовСКонтрагентом_Key: settle,
              СчетУчетаРасчетовПоАвансам_Key: advance,
            }),
          ],
        });
        const notes: string[] = [];
        if (payerName && payerName !== counterparty.name)
          notes.push(
            `Наименование плательщика в источнике «${payerName}», в 1С «${counterparty.name}» — сверьте перед проведением.`,
          );
        notes.push(
          explicitVat
            ? `НДС переносится явно: ставка ${vatRateValue}, сумма ${vatAmountValue}.`
            : "НДС в источнике не задан: выставлены «БезНДС» и 0 — как у обычной оплаты покупателя.",
        );
        let effectivePayload = payload;
        const created = await createOrPreview(conn, set, payload, confirm, notes, {
          settlementSide: "customer",
          prepared: (prepared) => {
            effectivePayload = prepared;
          },
        });
        const createdData = created.structuredContent as Record<string, unknown> | undefined;
        if (!createdData || created.isError) return created;
        const resolved = {
          sourceKey,
          organization: { ref: org.ref, name: org.name },
          bankAccount: { ref: bank.ref, number: bank.number },
          counterparty: { ref: counterparty.ref, name: counterparty.name },
          ...(payerBank ? { payerBankAccount: { ref: payerBank.ref, number: payerBank.number } } : {}),
          ...(rowCashFlowItemRef ? { rowCashFlowItem: { ref: rowCashFlowItemRef } } : {}),
          ...(bankLedgerAccountRef ? { bankLedgerAccount: { ref: bankLedgerAccountRef } } : {}),
          contract: {
            ref: contract.ref,
            ...(contract.number ? { number: contract.number } : {}),
            ...(contract.kind ? { kind: contract.kind } : {}),
          },
          vat: { rate: vatRateValue, amount: vatAmountValue, explicit: explicitVat },
        };
        // Что именно проверяем контрольным чтением: только явно перенесённые/заполненные нами поля.
        // Счета расчётов MCP может уточнить по настройке договора — их сверяет отдельный путь.
        const expectedHeader: Record<string, unknown> = {
          Date: documentDateTime,
          Posted: false,
          ...(ownNumber ? { Number: ownNumber } : {}),
          Организация_Key: org.ref,
          СчетОрганизации_Key: bank.ref,
          Контрагент: counterparty.ref,
          ...(cpSet ? { Контрагент_Type: `StandardODATA.${cpSet}` } : {}),
          ДоговорКонтрагента_Key: contract.ref,
          ВидОперации: payload["ВидОперации"],
          СуммаДокумента: normalizedAmount,
          ...(purpose !== undefined ? { НазначениеПлатежа: purpose } : {}),
          ...(documentNumber !== undefined ? { НомерВходящегоДокумента: documentNumber } : {}),
          ...(incomingDate ? { ДатаВходящегоДокумента: incomingDate } : {}),
          ...(cashFlowItemRef ? { СтатьяДвиженияДенежныхСредств_Key: cashFlowItemRef } : {}),
          ...(bankLedgerAccountRef ? { СчетБанк_Key: bankLedgerAccountRef } : {}),
          ...(comment !== undefined ? { Комментарий: effectivePayload["Комментарий"] ?? comment } : {}),
          ...(incomeRecognition ? { ОтражениеВУСН: incomeRecognition } : {}),
          ...(payerBank ? { СчетКонтрагента_Key: payerBank.ref } : {}),
        };
        const effectiveRow =
          (effectivePayload["РасшифровкаПлатежа"] as Array<Record<string, unknown>> | undefined)?.[0] ?? {};
        const expectedRow: Record<string, unknown> = {
          ДоговорКонтрагента_Key: contract.ref,
          СпособПогашенияЗадолженности: "Автоматически",
          СуммаПлатежа: normalizedAmount,
          СтавкаНДС: vatRateValue,
          СуммаНДС: vatAmountValue,
          ...(effectiveRow["СчетУчетаРасчетовСКонтрагентом_Key"] !== undefined
            ? { СчетУчетаРасчетовСКонтрагентом_Key: effectiveRow["СчетУчетаРасчетовСКонтрагентом_Key"] }
            : {}),
          ...(effectiveRow["СчетУчетаРасчетовПоАвансам_Key"] !== undefined
            ? { СчетУчетаРасчетовПоАвансам_Key: effectiveRow["СчетУчетаРасчетовПоАвансам_Key"] }
            : {}),
          ...(rowCashFlowItemRef ? { СтатьяДвиженияДенежныхСредств_Key: rowCashFlowItemRef } : {}),
          ...(customerInvoiceRef ? { СчетНаОплату_Key: customerInvoiceRef } : {}),
        };
        const verificationPlan = [
          ...Object.keys(expectedHeader),
          ...Object.keys(expectedRow).map((f) => `РасшифровкаПлатежа.${f}`),
        ];
        if (!confirm || createdData.dryRun === true)
          return withStructured(created, { resolved, verificationPlan });
        const ref = typeof createdData.ref === "string" ? createdData.ref : "";
        if (!ref)
          return {
            ...created,
            isError: true,
            _meta: { "odata-1c/errorKind": "unknown" },
            content: [
              {
                type: "text",
                text:
                  "1С приняла запрос, но ответ не содержит Ref_Key созданного документа. " +
                  "Повторная запись заблокирована; проверьте документ в 1С вручную.",
              },
            ],
            structuredContent: { ...createdData, resolved, verification: "failed" },
          };
        let actual: ODataEntity;
        try {
          actual = await conn.client.getEntity(`${set}(${odataGuid(ref)})?$format=json`);
        } catch (error) {
          const reason = error instanceof Error ? error.message : String(error);
          return {
            ...created,
            isError: true,
            _meta: { "odata-1c/errorKind": "unknown" },
            content: [
              {
                type: "text",
                text:
                  `Документ создан (Ref_Key ${ref}), но перечитать его не удалось: ${reason}. ` +
                  "Повторная запись заблокирована; проверьте поля в 1С вручную.",
              },
            ],
            structuredContent: {
              ...createdData,
              resolved,
              verification: "failed",
              verificationNote: "Контрольное чтение не выполнено.",
            },
          };
        }
        const mismatches: Array<{ field: string; expected: unknown; actual: unknown }> = [];
        const actualRows = actual["РасшифровкаПлатежа"];
        if (!Array.isArray(actualRows) || actualRows.length !== 1)
          mismatches.push({
            field: "РасшифровкаПлатежа",
            expected: "ровно одна строка",
            actual: Array.isArray(actualRows) ? `строк: ${actualRows.length}` : "нет табличной части",
          });
        const compare = (field: string, expected: unknown, where: "header" | "row"): void => {
          const props = where === "header" ? headerProps : rowProps;
          const type = props.get(field) ?? "Edm.String";
          const actualValue = where === "header" ? actual[field] : rowValue(actual, field);
          if (!sameStoredValue(expected, actualValue, type))
            mismatches.push({
              field: where === "row" ? `РасшифровкаПлатежа.${field}` : field,
              expected,
              actual: actualValue,
            });
        };
        for (const [field, value] of Object.entries(expectedHeader)) compare(field, value, "header");
        for (const [field, value] of Object.entries(expectedRow)) compare(field, value, "row");
        if (mismatches.length)
          return {
            ...created,
            isError: true,
            _meta: { "odata-1c/errorKind": "refused" },
            content: [
              {
                type: "text",
                text:
                  `Документ создан (Ref_Key ${ref}), но контрольное чтение не подтвердило поля: ` +
                  `${mismatches.map((m) => m.field).join(", ")}. Повторная запись заблокирована; ` +
                  "исправьте документ в 1С вручную и не запускайте загрузку этого sourceKey заново как новую.",
              },
            ],
            structuredContent: {
              ...createdData,
              resolved,
              verification: "failed",
              mismatches,
              verificationNote: "Поля после создания отличаются от явных полей источника.",
            },
          };
        return withStructured(created, {
          resolved,
          verification: "ok",
          verifiedFields: verificationPlan,
        });
      }),
  );

  // === Фаза 2: денежные документы (списание с р/с, ПКО, РКО) ===

  /**
   * Строка «Расшифровка платежа» для взаиморасчётов с контрагентом. Заполняется
   * только для операций оплаты поставщику/от покупателя при заданном договоре —
   * иначе платёж без взаиморасчётов (налог/ЗП/взнос), ТЧ не нужна.
   */
  async function buildSettlementRow(
    conn: Connection,
    operationKind: string,
    contractRef: string | undefined,
    amount: number,
  ): Promise<{ settle?: string; row: Record<string, unknown> } | undefined> {
    const supplier = /Поставщик/i.test(operationKind);
    const customer = /Покупател/i.test(operationKind);
    if (!contractRef || (!supplier && !customer)) return undefined;
    const codes = conn.cfg?.accounting?.strictDefaults
      ? new Map<AccountDefault, string | undefined>()
      : await resolveDefaultAccounts(
          conn,
          supplier ? ["payable", "payableAdvance"] : ["receivable", "receivableAdvance"],
        );
    const settle = codes.get(supplier ? "payable" : "receivable");
    const advance = codes.get(supplier ? "payableAdvance" : "receivableAdvance");
    return {
      settle,
      row: clean({
        LineNumber: 1,
        ДоговорКонтрагента_Key: contractRef,
        СпособПогашенияЗадолженности: "Автоматически",
        СуммаПлатежа: amount,
        СтавкаНДС: "БезНДС",
        СчетУчетаРасчетовСКонтрагентом_Key: settle,
        СчетУчетаРасчетовПоАвансам_Key: advance,
      }),
    };
  }

  const cpTypeOf = async (conn: Connection): Promise<Record<string, string>> => {
    const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
    return cpSet ? { Контрагент_Type: `StandardODATA.${cpSet}` } : {};
  };

  server.registerTool(
    "write.money.create_bank_writeoff",
    {
      title: "Создать списание с расчётного счёта",
      description:
        "Создаёт документ «Списание с расчётного счёта» (исходящий платёж с банка). НЕПРОВЕДЁННЫЙ; " +
        "провести — post_document. Вид операции ОБЯЗАТЕЛЕН и не угадывается (напр. «ОплатаПоставщику», " +
        "«ПеречислениеНалога», «ПеречислениеЗП», «ПереводНаДругойСчет», «ПрочееСписание»). " +
        "Для оплаты поставщику с договором заполняется расшифровка: счёт расчётов — по настройке договора, источник виден в предпросмотре (обычно 60), сумма взаиморасчётов — для рублёвого договора. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        operationKind: z
          .string()
          .describe(
            "Вид операции (ОБЯЗАТЕЛЬНО, не угадывать): ОплатаПоставщику / ПеречислениеНалога / ПеречислениеЗП / ПереводНаДругойСчет / ПрочееСписание …",
          ),
        amount: z.number().positive().describe("Сумма списания"),
        counterpartyRef: z.string().optional().describe("Ref_Key контрагента (для расчётов)"),
        contractRef: z.string().optional().describe("Ref_Key договора (для расшифровки взаиморасчётов)"),
        purposeText: z.string().optional().describe("Назначение платежа (текст)"),
        cashflowItemRef: z.string().optional().describe("Ref_Key статьи ДДС"),
        orgBankAccount: orgBankAccountField,
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        comment: z.string().optional().describe("Комментарий"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      operationKind,
      amount,
      counterpartyRef,
      contractRef,
      purposeText,
      cashflowItemRef,
      orgBankAccount,
      date,
      comment,
      confirm,
    }) =>
      guard("write.money.create_bank_writeoff", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.bankOut,
          "Документ «Списание с расчётного счёта»",
        );
        const org = await resolveOrg(conn, organization);
        const bank = await resolveOrgBankAccount(conn, org.key, orgBankAccount);
        const bd = await buildSettlementRow(conn, operationKind, contractRef, amount);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          ВидОперации: operationKind,
          Организация_Key: org.key,
          СчетОрганизации_Key: bank,
          ...(counterpartyRef ? { Контрагент: counterpartyRef, ...(await cpTypeOf(conn)) } : {}),
          ДоговорКонтрагента_Key: contractRef,
          СуммаДокумента: amount,
          НазначениеПлатежа: purposeText,
          СтатьяДвиженияДенежныхСредств_Key: cashflowItemRef,
          Комментарий: comment,
          ...(bd ? { РасшифровкаПлатежа: [bd.row] } : {}),
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  const cashDocInput = {
    database: databaseField,
    organization: organizationField,
    operationKind: z.string().describe("Вид операции (ОБЯЗАТЕЛЬНО, не угадывать)"),
    amount: z.number().positive().describe("Сумма"),
    cashRef: z.string().optional().describe("Ref_Key кассы организации (СчетКасса)"),
    counterpartyRef: z.string().optional().describe("Ref_Key контрагента (для расчётов)"),
    contractRef: z.string().optional().describe("Ref_Key договора"),
    basis: z.string().optional().describe("Основание (текст)"),
    cashflowItemRef: z.string().optional().describe("Ref_Key статьи ДДС"),
    date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
    comment: z.string().optional().describe("Комментарий"),
    confirm: confirmField,
  };

  const buildCashPayload = async (
    conn: Connection,
    org: { key: string },
    a: {
      operationKind: string;
      amount: number;
      cashRef?: string;
      counterpartyRef?: string;
      contractRef?: string;
      basis?: string;
      cashflowItemRef?: string;
      date?: string;
      comment?: string;
    },
  ): Promise<Record<string, unknown>> => {
    const bd = await buildSettlementRow(conn, a.operationKind, a.contractRef, a.amount);
    return clean({
      Date: odataDate(a.date ? new Date(`${a.date}T00:00:00`) : new Date()),
      Posted: false,
      ВидОперации: a.operationKind,
      Организация_Key: org.key,
      СчетКасса_Key: a.cashRef,
      ...(a.counterpartyRef ? { Контрагент: a.counterpartyRef, ...(await cpTypeOf(conn)) } : {}),
      ДоговорКонтрагента_Key: a.contractRef,
      СуммаДокумента: a.amount,
      Основание: a.basis,
      СтатьяДвиженияДенежныхСредств_Key: a.cashflowItemRef,
      Комментарий: a.comment,
      ...(bd ? { РасшифровкаПлатежа: [bd.row] } : {}),
    });
  };

  server.registerTool(
    "write.money.create_cash_receipt",
    {
      title: "Создать приходный кассовый ордер (ПКО)",
      description:
        "Создаёт «Приходный кассовый ордер» (приём наличных в кассу). НЕПРОВЕДЁННЫЙ; провести — post_document. " +
        "Вид операции ОБЯЗАТЕЛЕН (напр. «ОплатаПокупателя», «ПрочийПриход», «ПолучениеНаличныхВБанке», " +
        "«ВозвратОтПодотчетногоЛица»). " +
        "Для оплаты по договору заполняется расшифровка: счёт расчётов — по настройке договора, источник виден в предпросмотре, сумма взаиморасчётов — для рублёвого договора. " +
        "dry-run/confirm.",
      inputSchema: z.object(cashDocInput),
      outputSchema: createResultSchema,
    },
    (a) =>
      guard("write.money.create_cash_receipt", async () => {
        const conn = ctx.db(a.database);
        const set = await requireEntity(conn, WRITE_DOCUMENTS.cashIn, "Документ «Приходный кассовый ордер»");
        const org = await resolveOrg(conn, a.organization);
        const payload = await buildCashPayload(conn, org, a);
        return createOrPreview(conn, set, payload, a.confirm);
      }),
  );

  server.registerTool(
    "write.money.create_cash_payment",
    {
      title: "Создать расходный кассовый ордер (РКО)",
      description:
        "Создаёт «Расходный кассовый ордер» (выдача наличных из кассы). НЕПРОВЕДЁННЫЙ; провести — post_document. " +
        "Вид операции ОБЯЗАТЕЛЕН (напр. «ОплатаПоставщику», «ВыдачаПодотчетномуЛицу», «ВзносНаличнымиВБанк», " +
        "«ВыплатаЗаработнойПлатыПоВедомостям»). " +
        "Для оплаты по договору заполняется расшифровка: счёт расчётов — по настройке договора, источник виден в предпросмотре, сумма взаиморасчётов — для рублёвого договора. " +
        "dry-run/confirm.",
      inputSchema: z.object(cashDocInput),
      outputSchema: createResultSchema,
    },
    (a) =>
      guard("write.money.create_cash_payment", async () => {
        const conn = ctx.db(a.database);
        const set = await requireEntity(conn, WRITE_DOCUMENTS.cashOut, "Документ «Расходный кассовый ордер»");
        const org = await resolveOrg(conn, a.organization);
        const payload = await buildCashPayload(conn, org, a);
        return createOrPreview(conn, set, payload, a.confirm);
      }),
  );

  server.registerTool(
    "write.sales.create_issued_invoice",
    {
      title: "Создать счёт-фактуру выданный",
      description:
        "Создаёт «Счёт-фактура выданный» НА ОСНОВАНИИ реализации (или иного документа отгрузки). " +
        "Реквизиты (организация, контрагент, договор, суммы, НДС) наследуются от документа-основания. " +
        "НЕПРОВЕДЁННЫЙ; провести — post_document. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        baseDocumentRef: z.string().describe("Ref_Key документа-основания (реализация товаров и услуг)"),
        baseDocumentEntity: z
          .string()
          .optional()
          .describe(
            "Имя набора документа-основания, если это не «Реализация…» (напр. Document_ОтчетКомитентуОПродажах)",
          ),
        operationCode: z.string().optional().describe("Код вида операции (по умолчанию 01)"),
        date: z.string().optional().describe("Дата и дата выставления YYYY-MM-DD (по умолчанию сегодня)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, baseDocumentRef, baseDocumentEntity, operationCode, date, confirm }) =>
      guard("write.sales.create_issued_invoice", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.issuedInvoice,
          "Документ «Счёт-фактура выданный»",
        );
        const baseSet =
          baseDocumentEntity ??
          (await requireEntity(
            conn,
            WRITE_DOCUMENTS.sales,
            "Документ-основание «Реализация товаров и услуг»",
          ));
        const b = await invoiceBasis(conn, baseSet, baseDocumentRef);
        const d = odataDate(date ? new Date(`${date}T00:00:00`) : new Date());
        const baseGuid = baseDocumentRef.replace(/[{}']/g, "");
        const baseRow = clean({
          LineNumber: 1,
          ДокументОснование: baseGuid,
          ДокументОснование_Type: `StandardODATA.${baseSet}`,
        });
        const payload = clean({
          Date: d,
          Posted: false,
          ВидСчетаФактуры: "НаРеализацию",
          Организация_Key: b.org,
          Контрагент_Key: b.counterparty,
          ДоговорКонтрагента_Key: b.contract,
          ДокументОснование: baseGuid,
          ДокументОснование_Type: `StandardODATA.${baseSet}`,
          ДатаВыставления: d,
          КодВидаОперации: operationCode ?? "01",
          СуммаДокумента: b.total,
          СуммаНДСДокумента: b.vat,
          ДокументыОснования: [baseRow],
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.purchase.create_received_invoice",
    {
      title: "Создать счёт-фактуру полученный",
      description:
        "Создаёт «Счёт-фактура полученный» НА ОСНОВАНИИ поступления товаров и услуг. " +
        "Реквизиты наследуются от основания; дату входящего счёта-фактуры продавца указывают вручную " +
        "(в поступлении её нет). НЕПРОВЕДЁННЫЙ; провести — post_document. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        baseDocumentRef: z.string().describe("Ref_Key документа-основания (поступление товаров и услуг)"),
        baseDocumentEntity: z
          .string()
          .optional()
          .describe("Имя набора документа-основания, если это не «Поступление…»"),
        incomingDate: z.string().optional().describe("Дата входящего счёта-фактуры продавца YYYY-MM-DD"),
        operationCode: z.string().optional().describe("Код вида операции (по умолчанию 01)"),
        date: z.string().optional().describe("Дата регистрации YYYY-MM-DD (по умолчанию сегодня)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, baseDocumentRef, baseDocumentEntity, incomingDate, operationCode, date, confirm }) =>
      guard("write.purchase.create_received_invoice", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          WRITE_DOCUMENTS.receivedInvoice,
          "Документ «Счёт-фактура полученный»",
        );
        const baseSet =
          baseDocumentEntity ??
          (await requireEntity(
            conn,
            WRITE_DOCUMENTS.purchases,
            "Документ-основание «Поступление товаров и услуг»",
          ));
        const b = await invoiceBasis(conn, baseSet, baseDocumentRef);
        const baseGuid = baseDocumentRef.replace(/[{}']/g, "");
        const baseRow = clean({
          LineNumber: 1,
          ДокументОснование: baseGuid,
          ДокументОснование_Type: `StandardODATA.${baseSet}`,
        });
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          ВидСчетаФактуры: "НаПоступление",
          Организация_Key: b.org,
          Контрагент_Key: b.counterparty,
          ДоговорКонтрагента_Key: b.contract,
          ДатаВходящегоДокумента: incomingDate ? odataDate(new Date(`${incomingDate}T00:00:00`)) : undefined,
          ДокументОснование: baseGuid,
          ДокументОснование_Type: `StandardODATA.${baseSet}`,
          НДСПредъявленКВычету: true,
          КодВидаОперации: operationCode ?? "01",
          СуммаДокумента: b.total,
          СуммаНДСДокумента: b.vat,
          ДокументыОснования: [baseRow],
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.counterparty.create_bank_account",
    {
      title: "Создать банковский счёт контрагента",
      description:
        "Заводит расчётный счёт контрагента (подчинённый справочник): банк по БИК + номер счёта. " +
        "Можно сделать счёт основным. По умолчанию dry-run; создание — при confirm=true. " +
        "Ref контрагента — из find_counterparty.",
      inputSchema: z.object({
        database: databaseField,
        ownerRef: z.string().describe("Ref_Key контрагента-владельца счёта"),
        accountNumber: z.string().min(1).describe("Номер расчётного счёта"),
        bik: z.string().min(1).describe("БИК банка (ищется в справочнике «Банки»)"),
        currency: z.string().optional().describe("Валюта (код/название; по умолчанию рубль)"),
        label: z.string().optional().describe("Наименование счёта (по умолчанию — номер счёта)"),
        makeMain: z.boolean().default(false).describe("Сделать основным банковским счётом контрагента"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, ownerRef, accountNumber, bik, currency, label, makeMain, confirm }) =>
      guard("write.counterparty.create_bank_account", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, CATALOGS.bankAccounts, "Справочник «Банковские счета»");
        const bank = await resolveBankByBik(conn, bik);
        const cur = await resolveCurrencyForAccount(conn, currency ?? "643");
        const payload = clean({
          Description: label ?? accountNumber,
          Owner: ownerRef,
          Owner_Type: COUNTERPARTY_TYPE,
          НомерСчета: accountNumber,
          Банк_Key: bank.ref,
          ВалютаДенежныхСредств_Key: cur.ref,
        });
        return createSubordinate(
          conn,
          set,
          payload,
          confirm,
          makeMain
            ? {
                ownerSet: await resolveSet(conn, CATALOGS.counterparties, "Контрагенты"),
                ownerRef,
                field: "ОсновнойБанковскийСчет_Key",
              }
            : undefined,
        );
      }),
  );

  server.registerTool(
    "write.organization.create_bank_account",
    {
      title: "Сопоставить или создать банковский счёт организации",
      description:
        "Нужен перед банковским поступлением, когда у организации ещё нет счёта в 1С. Сначала точным " +
        "проходом сопоставляет счёт именно организации по номеру (Owner_Type = Catalog_Организации); " +
        "если он есть — ничего не пишет и возвращает existing=true. Если счёта нет — создаёт его в том же " +
        "справочнике «Банковские счета» с владельцем-организацией (Owner = Ref_Key организации), а не " +
        "контрагентом. Dry-run по умолчанию. Счёт контрагента за счёт организации не выдаётся: владелец " +
        "проверяется до и после записи. Основной счёт организации этот путь не меняет — это делается в 1С осознанно.",
      inputSchema: z.object({
        database: databaseField,
        organizationRef: z.string().uuid().describe("Ref_Key организации-владельца счёта"),
        accountNumber: z.string().trim().describe("Точный номер расчётного счёта (20 цифр)"),
        bik: z.string().trim().describe("БИК банка (ищется в справочнике «Банки»)"),
        currency: z.string().optional().describe("Валюта (код/название; по умолчанию рубль)"),
        label: z.string().optional().describe("Наименование счёта (по умолчанию — номер счёта)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, organizationRef, accountNumber, bik, currency, label, confirm }) =>
      guard("write.organization.create_bank_account", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, CATALOGS.bankAccounts, "Справочник «Банковские счета»");
        const meta = await conn.getMetadata();
        const accMeta = meta.entities.get(set);
        const need = (field: string, type: string): void => {
          const actual = accMeta?.properties.find((p) => p.name === field)?.type;
          if (actual !== type)
            throw new InputError(
              `Схема «Банковские счета» не публикует поле ${field} с типом ${type}` +
                (actual ? ` (найдено ${actual})` : "") +
                ". Безопасное создание счёта организации невозможно — операция не начата.",
            );
        };
        need("Description", "Edm.String");
        need("Owner", "Edm.String");
        need("Owner_Type", "Edm.String");
        need("НомерСчета", "Edm.String");
        need("Банк_Key", "Edm.Guid");
        need("ВалютаДенежныхСредств_Key", "Edm.Guid");
        const orgSet = await resolveSet(conn, CATALOGS.organizations, "Организации");
        const orgRef = organizationRef.replace(/[{}']/g, "").toLowerCase();
        let orgDoc: ODataEntity;
        try {
          orgDoc = await conn.client.getEntity(
            `${orgSet}(${odataGuid(orgRef)})?$format=json&$select=Ref_Key,Description,DeletionMark`,
          );
        } catch (error) {
          if (error instanceof ODataError && error.kind === "not_found")
            throw new InputError(`Организация ${orgRef} не найдена — счёт создать не для кого.`);
          throw error;
        }
        if (orgDoc.DeletionMark === true) throw new InputError("Организация помечена на удаление.");
        if (String(orgDoc.Ref_Key).toLowerCase() !== orgRef)
          throw new InputError("1С вернула другую организацию — счёт не создан.");
        const ownerType = `StandardODATA.${orgSet}`;
        const existing = await findBankAccountsByOwnerAndNumber(conn, orgRef, accountNumber, ownerType);
        if (existing.truncated)
          return fail(
            "Банковские счета: достигнут предел строк, проход неполный. Нельзя утверждать, " +
              "что счёт с таким номером один. Увеличьте ODATA_MAX_ROWS.",
            "input",
          );
        if (existing.accounts.length > 1)
          return fail(
            `У организации несколько счетов с номером ${accountNumber} — исправьте справочник 1С.`,
            "refused",
          );
        const found = existing.accounts[0];
        if (found)
          return ok({
            existing: true,
            database: conn.cfg.name,
            entitySet: set,
            ref: found.ref,
            description: found.name,
            ownerRef: orgRef,
            ownerType,
            note: "Счёт организации с таким номером уже есть; запись не выполнялась.",
          });
        const bank = await resolveBankByBik(conn, bik);
        const cur = await resolveCurrencyForAccount(conn, currency ?? "643");
        const payload = clean({
          Description: label ?? accountNumber,
          Owner: orgRef,
          Owner_Type: ownerType,
          НомерСчета: accountNumber,
          Банк_Key: bank.ref,
          ВалютаДенежныхСредств_Key: cur.ref,
        });
        const result = await createSubordinate(conn, set, payload, confirm);
        const data = result.structuredContent as Record<string, unknown> | undefined;
        if (result.isError || !data) return result;
        if (data.dryRun === true)
          return withStructured(result, {
            ownerRef: orgRef,
            ownerType,
            diagnostic:
              "Счёт организации будет создан в «Банковских счетах» с Owner = организация и " +
              `Owner_Type = ${ownerType}; это не счёт контрагента.`,
          });
        if (data.created !== true) return result;
        const ref = typeof data.ref === "string" ? data.ref : "";
        let actual: ODataEntity;
        try {
          actual = await conn.client.getEntity(
            `${set}(${odataGuid(ref)})?$format=json&$select=Ref_Key,Owner,Owner_Type,НомерСчета,Description`,
          );
        } catch (error) {
          const reason = error instanceof Error ? error.message : String(error);
          return {
            ...result,
            isError: true,
            _meta: { "odata-1c/errorKind": "unknown" },
            content: [
              {
                type: "text",
                text:
                  `Счёт создан (Ref_Key ${ref}), но перечитать его не удалось: ${reason}. ` +
                  "Повторная запись заблокирована; проверьте владельца счёта в 1С вручную.",
              },
            ],
            structuredContent: { ...data, ownerRef: orgRef, ownerType, verification: "failed" },
          };
        }
        const sameOwner =
          String(actual.Owner ?? "").toLowerCase() === orgRef &&
          String(actual.Owner_Type ?? "") === ownerType;
        if (!sameOwner)
          return {
            ...result,
            isError: true,
            _meta: { "odata-1c/errorKind": "refused" },
            content: [
              {
                type: "text",
                text:
                  `Счёт создан (Ref_Key ${ref}), но владелец не совпал: Owner_Type=${String(actual.Owner_Type ?? "—")}, ` +
                  `Owner=${String(actual.Owner ?? "—")}; ожидался ${ownerType} / ${orgRef}. ` +
                  "Не используйте этот счёт как счёт организации; проверьте 1С вручную.",
              },
            ],
            structuredContent: {
              ...data,
              ownerRef: orgRef,
              ownerType,
              verification: "failed",
              actualOwner: String(actual.Owner ?? ""),
              actualOwnerType: String(actual.Owner_Type ?? ""),
            },
          };
        return withStructured(result, {
          ownerRef: orgRef,
          ownerType,
          verification: "ok",
          note: "Создан счёт именно организации; владелец проверен контрольным чтением.",
        });
      }),
  );

  server.registerTool(
    "write.counterparty.create_contact_person",
    {
      title: "Создать контактное лицо (директор и т.п.)",
      description:
        "Заводит контактное лицо контрагента (директор, бухгалтер, менеджер): ФИО + должность. " +
        "Можно сделать основным контактным лицом. По умолчанию dry-run; создание — при confirm=true.",
      inputSchema: z.object({
        database: databaseField,
        ownerRef: z.string().describe("Ref_Key контрагента"),
        name: z.string().min(1).describe("ФИО (наименование контактного лица)"),
        lastName: z.string().optional().describe("Фамилия"),
        firstName: z.string().optional().describe("Имя"),
        position: z.string().optional().describe("Должность (напр. «Директор»)"),
        makeMain: z.boolean().default(false).describe("Сделать основным контактным лицом контрагента"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, ownerRef, name, lastName, firstName, position, makeMain, confirm }) =>
      guard("write.counterparty.create_contact_person", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, ["Catalog_КонтактныеЛица"], "Справочник «Контактные лица»");
        const payload = clean({
          Description: name,
          Фамилия: lastName,
          Имя: firstName,
          Должность: position,
          ОбъектВладелец: ownerRef,
          ОбъектВладелец_Type: COUNTERPARTY_TYPE,
        });
        return createSubordinate(
          conn,
          set,
          payload,
          confirm,
          makeMain
            ? {
                ownerSet: await resolveSet(conn, CATALOGS.counterparties, "Контрагенты"),
                ownerRef,
                field: "ОсновноеКонтактноеЛицо_Key",
              }
            : undefined,
        );
      }),
  );

  server.registerTool(
    "write.entity.create_folder",
    {
      title: "Создать папку (группу) справочника",
      description:
        "Создаёт папку/группу в иерархическом справочнике 1С (по умолчанию — Контрагенты). " +
        "Опционально вкладывает в родительскую папку (parent — имя папки или Ref_Key). " +
        "В 1С «папки» = группы справочника; отдельных тегов в типовой Бухгалтерии нет. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z
          .string()
          .default("Catalog_Контрагенты")
          .describe("Иерархический справочник (по умолчанию Catalog_Контрагенты)"),
        name: z.string().min(1).describe("Название папки"),
        parent: z.string().optional().describe("Родительская папка — имя или Ref_Key (для вложенности)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, entitySet, name, parent, confirm }) =>
      guard("write.entity.create_folder", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const parentRef = await folderRefOf(conn, entitySet, parent);
        const payload = clean({ Description: name, IsFolder: true, Parent_Key: parentRef });
        return createOrPreview(conn, entitySet, payload, confirm);
      }),
  );

  server.registerTool(
    "write.entity.move_to_folder",
    {
      title: "Переместить объект в папку",
      description:
        "Перемещает элемент справочника (напр. контрагента) в папку/группу — задаёт родителя (Parent_Key). " +
        "folder — имя папки или Ref_Key. По умолчанию справочник Контрагенты. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z
          .string()
          .default("Catalog_Контрагенты")
          .describe("Справочник (по умолчанию Catalog_Контрагенты)"),
        ref: z.string().describe("Ref_Key перемещаемого элемента"),
        folder: z.string().describe("Папка назначения — имя или Ref_Key"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, entitySet, ref, folder, confirm }) =>
      guard("write.entity.move_to_folder", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const parentRef = await folderRefOf(conn, entitySet, folder);
        return patchOrPreview(conn, entitySet, ref, { Parent_Key: parentRef }, confirm);
      }),
  );
}
