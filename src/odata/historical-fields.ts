import { InputError } from "../errors.js";

/** Тип исторического поля в контракте 1С. */
export type HistoricalFieldType = "string" | "guid" | "datetime" | "boolean" | "number";

/** Намерение для одного поля: отсутствует, явно пусто или значение. */
export type HistoricalIntent = { kind: "omitted" } | { kind: "empty" } | { kind: "value"; value: unknown };

export interface HistoricalField {
  name: string;
  type: HistoricalFieldType;
  intent: HistoricalIntent;
  /** Обязательное поле должно присутствовать; omitted для него — ошибка до записи. */
  required?: boolean;
}

export interface HistoricalExpectation {
  name: string;
  type: HistoricalFieldType;
  expected: unknown;
  /** true — сервер должен вернуть поле; для optional explicit значений тоже true. */
  required: boolean;
  empty: boolean;
}

export interface BuiltHistoricalFields {
  payload: Record<string, unknown>;
  expectations: HistoricalExpectation[];
}

const GUID_RE = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
export const EMPTY_REF = "00000000-0000-0000-0000-000000000000";
export const EMPTY_DATE = "0001-01-01T00:00:00";

/** Кодирование явной пустоты по контракту типа; boolean не имеет пустого значения. */
function encodeEmpty(type: HistoricalFieldType): unknown {
  switch (type) {
    case "string":
      return "";
    case "guid":
      return EMPTY_REF;
    case "datetime":
      return EMPTY_DATE;
    case "number":
      return null;
    case "boolean":
      throw new InputError("Явно пустой boolean недопустим: false — это значение, а не пустота.");
  }
}

function validateValue(type: HistoricalFieldType, value: unknown, name: string): unknown {
  switch (type) {
    case "string":
      if (typeof value !== "string") throw new InputError(`${name}: ожидается строка.`);
      return value;
    case "guid":
      if (typeof value !== "string" || !GUID_RE.test(value.replace(/[{}]/g, "")))
        throw new InputError(`${name}: ожидается GUID.`);
      return value.replace(/[{}]/g, "").toLowerCase();
    case "datetime":
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(value))
        throw new InputError(`${name}: ожидается дата со временем YYYY-MM-DDTHH:MM:SS.`);
      return value;
    case "boolean":
      if (typeof value !== "boolean") throw new InputError(`${name}: ожидается boolean.`);
      return value;
    case "number":
      if (typeof value !== "number" || !Number.isFinite(value))
        throw new InputError(`${name}: ожидается число.`);
      return value;
  }
}

/**
 * Из одного описания намерений строит payload и требования readback.
 * Явная пустота кодируется по типу и никогда не становится произвольным null;
 * omitted не попадает в payload и не объявляется пустотой при чтении.
 */
export function buildHistoricalFields(fields: HistoricalField[]): BuiltHistoricalFields {
  const payload: Record<string, unknown> = {};
  const expectations: HistoricalExpectation[] = [];
  const seen = new Set<string>();
  for (const field of fields) {
    if (seen.has(field.name)) throw new InputError(`Историческое поле ${field.name} задано дважды.`);
    seen.add(field.name);
    if (field.intent.kind === "omitted") {
      if (field.required) throw new InputError(`Историческое обязательное поле ${field.name} не задано.`);
      continue;
    }
    if (field.intent.kind === "empty") {
      const encoded = encodeEmpty(field.type);
      payload[field.name] = encoded;
      expectations.push({
        name: field.name,
        type: field.type,
        expected: encoded,
        required: true,
        empty: true,
      });
      continue;
    }
    const encoded = validateValue(field.type, field.intent.value, field.name);
    payload[field.name] = encoded;
    expectations.push({
      name: field.name,
      type: field.type,
      expected: encoded,
      required: true,
      empty: false,
    });
  }
  return { payload, expectations };
}

const has = (obj: Record<string, unknown>, name: string): boolean =>
  Object.prototype.hasOwnProperty.call(obj, name) && obj[name] !== undefined;

function emptyStored(value: unknown, type: HistoricalFieldType): boolean {
  if (value === null || value === undefined || value === "") return true;
  if (type === "guid") return String(value).replace(/[{}]/g, "").toLowerCase() === EMPTY_REF;
  if (type === "datetime") return String(value).startsWith("0001-01-01");
  if (type === "number") return Number(value) === 0;
  return false;
}

function sameValue(expected: unknown, actual: unknown, type: HistoricalFieldType): boolean {
  if (type === "guid")
    return (
      String(expected ?? "")
        .replace(/[{}]/g, "")
        .toLowerCase() ===
      String(actual ?? "")
        .replace(/[{}]/g, "")
        .toLowerCase()
    );
  if (type === "boolean") return expected === actual;
  if (type === "number") return Math.abs(Number(expected) - Number(actual)) < 1e-9;
  if (type === "datetime") return String(expected) === String(actual);
  return String(expected) === String(actual);
}

/** Сверка фактического ответа 1С с требованиями; пропуск обязательного поля — mismatch. */
export function compareHistoricalFields(
  actual: Record<string, unknown>,
  expectations: HistoricalExpectation[],
): Array<{ field: string; expected: unknown; actual: unknown; reason: "missing" | "mismatch" }> {
  const mismatches: Array<{
    field: string;
    expected: unknown;
    actual: unknown;
    reason: "missing" | "mismatch";
  }> = [];
  for (const expectation of expectations) {
    if (!has(actual, expectation.name)) {
      mismatches.push({
        field: expectation.name,
        expected: expectation.expected,
        actual: undefined,
        reason: "missing",
      });
      continue;
    }
    const actualValue = actual[expectation.name];
    const ok = expectation.empty
      ? emptyStored(actualValue, expectation.type)
      : sameValue(expectation.expected, actualValue, expectation.type);
    if (!ok)
      mismatches.push({
        field: expectation.name,
        expected: expectation.expected,
        actual: actualValue,
        reason: "mismatch",
      });
  }
  return mismatches;
}
