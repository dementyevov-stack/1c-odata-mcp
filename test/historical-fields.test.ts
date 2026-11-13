import { describe, expect, it } from "vitest";
import {
  buildHistoricalFields,
  compareHistoricalFields,
  type HistoricalField,
} from "../src/odata/historical-fields.js";

const field = (
  name: string,
  type: HistoricalField["type"],
  intent: HistoricalField["intent"],
  required = true,
): HistoricalField => ({ name, type, intent, required });

describe("historical-fields: точное представление исторических значений", () => {
  it("различает omitted/empty/value и сохраняет 0, false и пустую строку", () => {
    const built = buildHistoricalFields([
      field("comment", "string", { kind: "empty" }),
      field("zero", "number", { kind: "value", value: 0 }),
      field("no", "boolean", { kind: "value", value: false }),
      field("missing", "string", { kind: "omitted" }, false),
      field("ref", "guid", { kind: "empty" }),
      field("date", "datetime", { kind: "value", value: "2026-11-01T12:34:56" }),
    ]);
    expect(built.payload).toEqual({
      comment: "",
      zero: 0,
      no: false,
      ref: "00000000-0000-0000-0000-000000000000",
      date: "2026-11-01T12:34:56",
    });
    expect(built.expectations.map((e) => e.name)).not.toContain("missing");
    expect(built.expectations.find((e) => e.name === "comment")?.empty).toBe(true);
  });

  it("отклоняет недопустимую пустоту и обязательное omitted до записи", () => {
    expect(() => buildHistoricalFields([field("flag", "boolean", { kind: "empty" })])).toThrow(/boolean/);
    expect(() => buildHistoricalFields([field("required", "string", { kind: "omitted" })])).toThrow(
      /не задано/,
    );
  });

  it("не считает пропуск поля пустотой и ловит подмену пустого значения", () => {
    const built = buildHistoricalFields([
      field("empty", "guid", { kind: "empty" }),
      field("zero", "number", { kind: "value", value: 0 }),
      field("date", "datetime", { kind: "value", value: "2026-11-01T12:34:56" }),
    ]);
    const missing = compareHistoricalFields({ zero: 0, date: "2026-11-01T12:34:56" }, built.expectations);
    expect(missing).toEqual([expect.objectContaining({ field: "empty", reason: "missing" })]);
    const substituted = compareHistoricalFields(
      {
        empty: "11111111-1111-4111-8111-111111111111",
        zero: 0,
        date: "2026-11-01T00:00:00",
      },
      built.expectations,
    );
    expect(substituted.map((m) => m.field)).toEqual(["empty", "date"]);
  });
});
