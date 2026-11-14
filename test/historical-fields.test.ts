import { describe, expect, it } from "vitest";
import { buildHistoricalFields, compareHistoricalFields } from "../src/odata/historical-fields.js";

describe("historical-fields: строгая числовая/дата пустота", () => {
  it("явный 0 не равен null/false/пустой строке", () => {
    const built = buildHistoricalFields([
      { name: "zero", type: "number", intent: { kind: "value", value: 0 }, required: true },
    ]);
    for (const actual of [null, false, ""]) {
      const mismatches = compareHistoricalFields({ zero: actual }, built.expectations);
      expect(mismatches, `actual=${String(actual)}`).toEqual([
        expect.objectContaining({ field: "zero", reason: "mismatch" }),
      ]);
    }
    expect(compareHistoricalFields({ zero: 0 }, built.expectations)).toEqual([]);
  });

  it("пустая дата — только контрактное нулевое значение", () => {
    const built = buildHistoricalFields([
      { name: "d", type: "datetime", intent: { kind: "empty" }, required: true },
    ]);
    expect(compareHistoricalFields({ d: "0001-01-01T00:00:00" }, built.expectations)).toEqual([]);
    expect(compareHistoricalFields({ d: "0001-01-02T00:00:00" }, built.expectations)).not.toEqual([]);
  });
});
