import { describe, expect, it } from "vitest";
import { fetchRegistrarPostings } from "../src/odata/accounting.js";
import { getDocumentPostings } from "../src/tools/registers.js";
import { fake1C } from "./helpers/fake-1c.js";
import { bp30Metadata, SNAPSHOT_REF } from "./helpers/metadata-snapshot.js";

/**
 * Исследование: ручная операция БП с отдельной шапкой и набором проводок.
 * Только локальная синтетика; запись в 1С и создание проводок здесь не выполняются.
 *
 * Схема снимка БП 3.0:
 *  - Document_ОперацияБух публикует шапку и таблицы регистров сведений/накопления,
 *    но НЕ табличную часть «Проводки»;
 *  - AccountingRegister_Хозрасчетный_RecordType/RecordSet не содержит ExtDimension*;
 *  - субконто читаются только из виртуальной RecordsWithExtDimensions.
 */

const DOC = "Document_ОперацияБух";
const REF = SNAPSHOT_REF;
const ACCOUNT_DR = SNAPSHOT_REF;
const ACCOUNT_CR = SNAPSHOT_REF;
const DIM_DR = "cccccccc-cccc-4ccc-8ccc-cccccccccccc";
const DIM_CR = "dddddddd-dddd-4ddd-8ddd-dddddddddddd";

const actingRecord = {
  Recorder: REF,
  Recorder_Type: `StandardODATA.${DOC}`,
  Period: "2026-11-01T12:00:00",
  LineNumber: 1,
  Active: true,
  AccountDr_Key: ACCOUNT_DR,
  AccountCr_Key: ACCOUNT_CR,
  Организация_Key: REF,
  Сумма: 100,
  СуммаНУDr: null,
  СуммаНУCr: null,
  СуммаПРDr: null,
  СуммаПРCr: null,
  СуммаВРDr: null,
  СуммаВРCr: null,
};

const dimensionRow = {
  ...actingRecord,
  ExtDimensionDr1: DIM_DR,
  ExtDimensionDr1_Type: "StandardODATA.Catalog_Контрагенты",
  ExtDimensionTypeDr1_Key: "11111111-1111-4111-8111-111111111111",
  ExtDimensionCr1: DIM_CR,
  ExtDimensionCr1_Type: "StandardODATA.Catalog_НоменклатурныеГруппы",
  ExtDimensionTypeCr1_Key: "22222222-2222-4222-8222-222222222222",
};

function syntheticManualOperation() {
  return fake1C(
    { READ_ONLY: "false", ODATA_WRITABLE: "true" },
    {
      data: (q) => {
        if (q.base === "AccountingRegister_Хозрасчетный_RecordType") return [actingRecord];
        if (q.base === "AccountingRegister_Хозрасчетный" && q.fn === "RecordsWithExtDimensions")
          return [dimensionRow];
        return undefined;
      },
    },
  );
}

describe("manual operation research: чтение активной проводки при Posted=false", () => {
  it("снимок БП 3.0 не публикует ни таблицу проводок операции, ни ExtDimension в записи регистра", () => {
    const meta = bp30Metadata();
    const operation = meta.entities.get(DOC)!;
    expect(operation.properties.some((p) => p.name === "Проводки")).toBe(false);
    expect(operation.properties.some((p) => p.name === "Posted")).toBe(true);
    expect(operation.properties.some((p) => p.name === "СуммаОперации")).toBe(true);

    const record = meta.entities.get("AccountingRegister_Хозрасчетный_RecordType")!;
    expect(record.properties.some((p) => p.name.startsWith("ExtDimension"))).toBe(false);
    const register = meta.entities.get("AccountingRegister_Хозрасчетный")!;
    expect(register.properties.find((p) => p.name === "RecordSet")?.type).toContain("Collection");
    expect(meta.entities.has("AccountingRegister_Хозрасчетный_RecordSet")).toBe(false);
  });

  it("Posted=false не скрывает активную запись; обе стороны имеют тип и ссылку субконто", async () => {
    const fake = syntheticManualOperation();
    const result = await getDocumentPostings(fake.conn, DOC, REF);
    expect(result.document.posted).toBe(false);
    expect(result.postingsCount).toBe(1);
    expect(result.postings[0]).toMatchObject({
      active: true,
      amount: 100,
      debit: {
        dimensionsStatus: "present",
        dimensions: [{ index: 1, type: "Catalog_Контрагенты", ref: DIM_DR }],
      },
      credit: {
        dimensionsStatus: "present",
        dimensions: [{ index: 1, type: "Catalog_НоменклатурныеГруппы", ref: DIM_CR }],
      },
    });
    expect(result.dimensionsNote).toBeUndefined();
    expect(fake.requests.every((r) => r.method === "GET")).toBe(true);
  });

  it("налоговые суммы пустые в записи и не превращаются в нулевые проводки", async () => {
    const fake = syntheticManualOperation();
    const raw = await fetchRegistrarPostings(fake.conn, DOC, REF);
    expect(raw.rows[0]).toMatchObject({
      СуммаНУDr: null,
      СуммаНУCr: null,
      СуммаПРDr: null,
      СуммаПРCr: null,
      СуммаВРDr: null,
      СуммаВРCr: null,
    });
    expect(raw.rows[0]).toMatchObject({ ExtDimensionDr1: DIM_DR, ExtDimensionCr1: DIM_CR });
  });

  it("в снимке нет отдельного набора RecordSet: прямая запись набора не является штатным OData-интерфейсом", () => {
    const meta = bp30Metadata();
    // RecordSet — структурная коллекция в шапке регистра, а не EntitySet для POST.
    expect(meta.entities.get("AccountingRegister_Хозрасчетный")!.properties).toEqual(
      expect.arrayContaining([expect.objectContaining({ name: "RecordSet" })]),
    );
    expect(meta.entities.has("AccountingRegister_Хозрасчетный_RecordSet")).toBe(false);
    // При этом запись и субконто в _RecordType не опубликованы; читать их можно только
    // через виртуальную RecordsWithExtDimensions (проверено выше).
    expect(
      meta.entities.get("AccountingRegister_Хозрасчетный_RecordType")!.properties.map((p) => p.name),
    ).not.toContain("ExtDimensionDr1");
  });
});
