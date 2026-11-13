import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * Синтетический контракт 1С-сторонней миграции ручной операции.
 * BSL здесь не компилируется: в проекте нет 1С/BSL-инструментов. Тест
 * фиксирует только инварианты контракта и обязательные маркеры эскиза.
 */

const contract = JSON.parse(
  readFileSync(new URL("../prototype/1c-migration-operation/contract.json", import.meta.url), "utf8"),
) as Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
const bsl = readFileSync(
  new URL("../prototype/1c-migration-operation/OperationMigration.bsl", import.meta.url),
  "utf8",
);

describe("1C manual-operation migration contract", () => {
  it("фиксирует атомарную пару шапка+движения, Posted/Active и ключ повторяемости", () => {
    expect(contract.document).toMatchObject({ entity: "Document_ОперацияБух", posted: false });
    expect(contract.transaction.atomic).toEqual(["documentHeader", "accountingRegisterMovements"]);
    expect(contract.transaction.rollbackOn).toEqual(
      expect.arrayContaining(["subconto_type_mismatch", "amount_mismatch", "readback_mismatch"]),
    );
    expect(contract.idempotency.duplicateAction).toBe("return_existing");
    expect(contract.movements).toMatchObject({
      register: "AccountingRegister_Хозрасчетный",
      oneRegistrar: true,
    });
    expect(contract.movements.rowFields).toEqual(
      expect.arrayContaining(["Active", "СубконтоДт1", "СубконтоДт1_Type", "СуммаНУДт"]),
    );
    expect(contract.movements.activePolicy).toMatch(/Copy Active/);
    expect(contract.movements.emptyAmountPolicy).toMatch(/no defaults/i);
  });

  it("содержит rollback-сценарий при несовпадении readback", () => {
    expect(contract.testScenario.name).toBe("rollback_on_subconto_type_mismatch");
    expect(contract.testScenario.expected).toEqual({
      documentCreated: false,
      movementCount: 0,
      rollback: true,
    });
    expect(contract.standardInterface.fallback).toMatch(/Not allowed/);
  });

  it("эскиз BSL требует безопасный режим, транзакцию, откат и строгую проверку субконто", () => {
    expect(bsl).toContain("БезопасныйРежим");
    expect(bsl).toContain("НачатьТранзакцию");
    expect(bsl).toContain("ОтменитьТранзакцию");
    expect(bsl).toContain("ПроверитьСубконто");
    expect(bsl).toContain("Posted");
    expect(bsl).toContain("TODO(owner)");
    // Прямая запись в таблицы БД/обход API не должна появляться в контракте.
    expect(bsl.toLowerCase()).not.toContain("insert into");
  });
});
