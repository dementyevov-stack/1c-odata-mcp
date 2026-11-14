import { describe, expect, it, vi, afterEach } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { ODataError } from "../src/odata/errors.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

const ORG = "11111111-1111-4111-8111-111111111111";
const CP = "22222222-2222-4222-8222-222222222222";
const CTR = "33333333-3333-4333-8333-333333333333";
const BASE = "44444444-4444-4444-8444-444444444444";
const NOM = "66666666-6666-4666-8666-666666666666";
const ZERO = "00000000-0000-0000-0000-000000000000";
const OTHER_CT = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";

const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(cleanup.splice(0).map((fn) => fn().catch(() => undefined)));
});

async function setup(
  options: {
    docs?: Record<string, unknown>[];
    getError?: boolean;
    dropPostResponse?: boolean;
    mutate?: (actual: Record<string, unknown>) => Record<string, unknown>;
    env?: Record<string, string>;
  } = {},
) {
  const fake = fake1C(
    { READ_ONLY: "false", ODATA_WRITABLE: "true", ...options.env },
    {
      dropPostResponse: options.dropPostResponse,
      data: (q: FakeQuery) => (q.base === "Document_СчетФактураВыданный" ? (options.docs ?? []) : undefined),
    },
  );
  const original = fake.conn.client.getEntity.bind(fake.conn.client);
  vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
    if (path.startsWith("Catalog_ДоговорыКонтрагентов(")) {
      const match = /Catalog_ДоговорыКонтрагентов\(guid'([^']+)'\)/.exec(decodeURIComponent(path));
      return {
        Ref_Key: match?.[1] ?? CTR,
        Owner_Key: CP,
        Организация_Key: ORG,
        ВидДоговора: "СПокупателем",
        DeletionMark: false,
        IsFolder: false,
      };
    }
    if (path.startsWith("Document_ПлатежноеПоручение("))
      return {
        Ref_Key: BASE,
        Организация_Key: ORG,
        Контрагент: CP,
        Контрагент_Type: "StandardODATA.Catalog_Контрагенты",
        ДоговорКонтрагента_Key: CTR,
        СуммаДокумента: 120,
      };
    if (path.startsWith("Document_ПоступлениеНаРасчетныйСчет("))
      return {
        Ref_Key: BASE,
        Организация_Key: ORG,
        Контрагент: CP,
        Контрагент_Type: "StandardODATA.Catalog_Контрагенты",
        ДоговорКонтрагента_Key: CTR,
        СуммаДокумента: 120,
      };
    if (path.startsWith("Document_РеализацияТоваровУслуг("))
      return {
        Ref_Key: BASE,
        Организация_Key: ORG,
        Контрагент_Key: CP,
        ДоговорКонтрагента_Key: CTR,
        СуммаДокумента: 120,
        Товары: [{ СуммаДокумента: 120, СуммаНДС: 20 }],
        Услуги: [],
      };
    if (path.startsWith("Document_СчетФактураВыданный(") && !path.includes("$select")) {
      if (options.getError)
        throw new ODataError({ kind: "not_found", message: "Объект не найден (404)", status: 404 });
      const actual = await original(path);
      const withMark = { DeletionMark: false, ...actual };
      return options.mutate ? options.mutate(withMark) : withMark;
    }
    return original(path);
  });
  const server = createServer(fake.ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "historical-issued-invoice-test", version: "1" });
  await client.connect(b);
  cleanup.push(async () => {
    await client.close().catch(() => undefined);
    await server.close().catch(() => undefined);
  });
  return { fake, client };
}

const call = async (client: Client, args: Record<string, unknown>) =>
  (await client.callTool({ name: "write.sales.create_issued_invoice", arguments: args })) as {
    isError?: boolean;
    structuredContent?: Record<string, unknown>;
  };
const postCount = (fake: ReturnType<typeof fake1C>): number =>
  fake.requests.filter((r) => r.method === "POST").length;

const args = {
  organization: ORG,
  baseDocumentRef: BASE,
  number: "SF-1",
  date: "2026-11-01",
  time: "12:34:56",
  issueDate: "2026-11-02",
  issueTime: "10:11:12",
  invoiceKind: "НаАванс",
  operationCode: "02",
  sumTotal: 120,
  sumVat: 20,
  basisDocuments: [
    {
      ref: BASE,
      type: "StandardODATA.Document_РеализацияТоваровУслуг",
      lineNumber: 1,
      sourceNumber: "",
      sourceDate: "2026-10-01",
      considerCorrection: false,
    },
  ],
  paymentDocuments: [],
  emptyShipmentDocuments: [],
  emptyCorrectableLineNumbers: [],
  advanceRows: [
    {
      nomenclatureRef: NOM,
      content: "Аванс по договору",
      amount: 120,
      vatRate: "НДС20",
      vatAmount: 20,
      counterpartyRef: CP,
      basisRef: BASE,
      basisType: "StandardODATA.Document_РеализацияТоваровУслуг",
      paymentDate: "2026-10-15",
      paymentNumber: "PAY-1",
      sellerInvoiceNumber: "SELL-1",
    },
  ],
};

describe("write.sales.create_issued_invoice: исторический режим", () => {
  it("обычный вызов сохраняет прежний путь", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, {
      organization: ORG,
      baseDocumentRef: BASE,
      date: "2026-11-01",
      confirm: false,
    });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({ ВидСчетаФактуры: "НаРеализацию", КодВидаОперации: "01" });
    expect(postCount(fake)).toBe(0);
  });

  it("сериализует номер, две даты/время, вид, код, основания и пустую таблицу", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      Number: "SF-1",
      Date: "2026-11-01T12:34:56",
      ДатаВыставления: "2026-11-02T10:11:12",
      ВидСчетаФактуры: "НаАванс",
      КодВидаОперации: "02",
      СуммаДокумента: 120,
      СуммаНДСДокумента: 20,
      ПлатежноРасчетныеДокументы: [],
    });
    const basis = (payload["ДокументыОснования"] as Array<Record<string, unknown>>)[0];
    expect(basis).toMatchObject({
      ДокументОснование: BASE,
      ДокументОснование_Type: "StandardODATA.Document_РеализацияТоваровУслуг",
      НомерИсходногоДокумента: "",
    });
    expect(postCount(fake)).toBe(0);
  });

  it("preflight отклоняет отсутствие вида/оснований до POST", async () => {
    const { client, fake } = await setup();
    for (const bad of [
      { ...args, invoiceKind: undefined },
      { ...args, basisDocuments: [] },
      { ...args, operationCode: undefined },
    ]) {
      const preview = await call(client, { ...bad, confirm: false });
      expect(preview.isError).toBe(true);
    }
    expect(postCount(fake)).toBe(0);
  });

  it("HTTP 500 после POST: outcome unknown, диагностика и без повтора", async () => {
    const { client, fake } = await setup();
    vi.spyOn(fake.conn.client, "create").mockRejectedValue(
      new ODataError({ kind: "server", message: "Ошибка сервера 1С (500)", status: 500 }),
    );
    const preview = await call(client, { ...args, confirm: false });
    const created = await call(client, {
      ...args,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError).toBe(true);
    expect(created.structuredContent).toMatchObject({
      stage: "POST",
      errorKind: "server",
      httpStatus: 500,
      outcome: "unknown",
    });
    expect(postCount(fake)).toBe(0);
  });

  it("потерянный ответ после сохранения даёт possibly_created, а не success", async () => {
    const { client } = await setup({ dropPostResponse: true });
    const preview = await call(client, { ...args, confirm: false });
    const created = await call(client, {
      ...args,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError).toBe(true);
    expect(created.structuredContent?.outcome).toBe("possibly_created");
  });

  it("GET 404 после успешного POST: outcome unknown", async () => {
    const { client, fake } = await setup({ getError: true });
    const preview = await call(client, { ...args, confirm: false });
    const created = await call(client, {
      ...args,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError).toBe(true);
    expect(created.structuredContent).toMatchObject({ stage: "GET", outcome: "unknown" });
    expect(postCount(fake)).toBe(1);
  });

  it("readback расходится по виду/основанию", async () => {
    const { client } = await setup({
      mutate: (actual) => {
        actual["ВидСчетаФактуры"] = "НаРеализацию";
        return actual;
      },
    });
    const preview = await call(client, { ...args, confirm: false });
    const created = await call(client, {
      ...args,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError).toBe(true);
    expect(JSON.stringify(created)).toContain("ВидСчетаФактуры");
  });

  it("транспортный HTTP 500 проходит через journal с httpStatus/causeChain", async () => {
    const { client } = await setup();
    const originalFetch = globalThis.fetch;
    vi.spyOn(globalThis, "fetch").mockImplementation(async (input, init) => {
      if (
        init?.method === "POST" &&
        decodeURIComponent(String(input)).includes("Document_СчетФактураВыданный")
      ) {
        return new Response(
          JSON.stringify({
            "odata.error": { message: { lang: "ru", value: "Не удалось записать: синтетика" } },
          }),
          { status: 500, headers: { "Content-Type": "application/json" } },
        );
      }
      return originalFetch(input, init);
    });
    const preview = await call(client, { ...args, confirm: false });
    const created = await call(client, {
      ...args,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError).toBe(true);
    expect(created.structuredContent).toMatchObject({ stage: "POST", httpStatus: 500, outcome: "unknown" });
    expect(created.structuredContent?.causeChain).toEqual(
      expect.arrayContaining([expect.objectContaining({ status: 500 })]),
    );
  });

  it("усечённый preflight не выдаётся за отсутствие", async () => {
    const { client, fake } = await setup();
    const original = fake.conn.client.getCollection.bind(fake.conn.client);
    vi.spyOn(fake.conn.client, "getCollection").mockImplementation(async (path: string) => {
      if (path.startsWith("Document_СчетФактураВыданный")) {
        return {
          value: Array.from({ length: 200 }, (_, i) => ({
            Ref_Key: `00000000-0000-4000-8000-${String(i).padStart(12, "0")}`,
          })),
        };
      }
      return original(path);
    });
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError).toBe(true);
    expect(JSON.stringify(preview)).toContain("неполная");
  });

  it("новый preview после создания не делает повторный POST", async () => {
    const { client, fake } = await setup();
    const firstPreview = await call(client, { ...args, confirm: false });
    const first = await call(client, {
      ...args,
      confirm: true,
      operationId: firstPreview.structuredContent?.operationId,
    });
    expect(first.isError, JSON.stringify(first)).toBeFalsy();
    const secondPreview = await call(client, { ...args, confirm: false });
    expect(secondPreview.isError).toBe(true);
    expect(JSON.stringify(secondPreview)).toContain("уже существует");
    expect(postCount(fake)).toBe(1);
  });

  it("два preview до первого confirm: второй confirm блокируется", async () => {
    const { client, fake } = await setup();
    const a = await call(client, { ...args, confirm: false });
    const b = await call(client, { ...args, confirm: false });
    const first = await call(client, {
      ...args,
      confirm: true,
      operationId: a.structuredContent?.operationId,
    });
    expect(first.isError, JSON.stringify(first)).toBeFalsy();
    const second = await call(client, {
      ...args,
      confirm: true,
      operationId: b.structuredContent?.operationId,
    });
    expect(second.isError).toBe(true);
    expect(postCount(fake)).toBe(1);
  });

  it("failed readback сохраняется при replay и не выдаётся за успех", async () => {
    const { client, fake } = await setup({
      mutate: (actual) => {
        actual["ВидСчетаФактуры"] = "НаРеализацию";
        return actual;
      },
    });
    const preview = await call(client, { ...args, confirm: false });
    const id = preview.structuredContent?.operationId;
    const first = await call(client, { ...args, confirm: true, operationId: id });
    expect(first.isError).toBe(true);
    const replay = await call(client, { ...args, confirm: true, operationId: id });
    expect(replay.isError).toBe(true);
    expect(replay.structuredContent?.verification).toBe("failed");
    expect(postCount(fake)).toBe(1);
  });

  it("сбой записи failed-readback оставляет pending и replay не успешен", async () => {
    const { client, fake } = await setup({
      mutate: (actual) => {
        actual["ВидСчетаФактуры"] = "НаРеализацию";
        return actual;
      },
    });
    vi.spyOn(fake.conn.client, "recordOperationVerification").mockRejectedValue(new Error("disk full"));
    const preview = await call(client, { ...args, confirm: false });
    const id = preview.structuredContent?.operationId;
    const first = await call(client, { ...args, confirm: true, operationId: id });
    expect(first.isError).toBe(true);
    const replay = await call(client, { ...args, confirm: true, operationId: id });
    expect(replay.isError).toBe(true);
    expect(replay.structuredContent?.verification).toBe("pending");
    expect(postCount(fake)).toBe(1);
  });

  it("обрыв между POST и успешным readback оставляет pending", async () => {
    const { client, fake } = await setup();
    vi.spyOn(fake.conn.client, "recordOperationVerification").mockRejectedValue(new Error("disk full"));
    const preview = await call(client, { ...args, confirm: false });
    const id = preview.structuredContent?.operationId;
    const first = await call(client, { ...args, confirm: true, operationId: id });
    expect(first.isError).toBe(true);
    expect(first.structuredContent?.verification).toBe("pending");
    const replay = await call(client, { ...args, confirm: true, operationId: id });
    expect(replay.isError).toBe(true);
    expect(replay.structuredContent?.verification).toBe("pending");
    expect(postCount(fake)).toBe(1);
  });

  it("старая квитанция без verification при restart не выдаётся за ok", async () => {
    const { client, fake } = await setup();
    vi.spyOn(fake.conn.client, "operationSettled").mockResolvedValue({
      entitySet: "Document_СчетФактураВыданный",
      result: { Ref_Key: "88888888-8888-4888-8888-888888888888" },
    });
    const replay = await call(client, {
      ...args,
      confirm: true,
      operationId: "99999999-9999-4999-8999-999999999999",
    });
    expect(replay.isError).toBe(true);
    expect(replay.structuredContent?.verification).toBe("unverified");
    expect(postCount(fake)).toBe(0);
  });

  it("новая обычная квитанция после confirm успешно replay-ится", async () => {
    const { client, fake } = await setup();
    const input = { organization: ORG, baseDocumentRef: BASE, date: "2026-11-01" };
    const preview = await call(client, { ...input, confirm: false });
    const id = preview.structuredContent?.operationId;
    const first = await call(client, { ...input, confirm: true, operationId: id });
    expect(first.isError, JSON.stringify(first)).toBeFalsy();
    const replay = await call(client, { ...input, confirm: true, operationId: id });
    expect(replay.isError, JSON.stringify(replay)).toBeFalsy();
    expect(replay.structuredContent?.replayed).toBe(true);
    expect(postCount(fake)).toBe(1);
  });

  it("advanceRows без исторических полей не уходит в обычную ветку", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, {
      organization: ORG,
      baseDocumentRef: BASE,
      date: "2026-11-01",
      advanceRows: [args.advanceRows[0]],
      confirm: false,
    });
    expect(preview.isError).toBe(true);
    expect(postCount(fake)).toBe(0);
  });

  it("явный zero GUID получателя строки аванса сохраняется и проверяется", async () => {
    const { client } = await setup({
      readback: undefined,
    });
    const zeroArgs = {
      ...args,
      advanceRows: [{ ...args.advanceRows[0], counterpartyRef: ZERO }],
    };
    const preview = await call(client, { ...zeroArgs, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const row = (preview.structuredContent?.payload as Record<string, unknown>)["Авансы"] as Array<
      Record<string, unknown>
    >;
    expect(row[0]["Контрагент_Key"]).toBe(ZERO);
    const created = await call(client, {
      ...zeroArgs,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
  });

  it("явный zero GUID договора шапки не заменяется договором основания", async () => {
    const { client } = await setup();
    const zeroArgs = { ...args, contractRef: ZERO };
    const preview = await call(client, { ...zeroArgs, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    expect((preview.structuredContent?.payload as Record<string, unknown>)["ДоговорКонтрагента_Key"]).toBe(
      ZERO,
    );
    const created = await call(client, {
      ...zeroArgs,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
  });

  it("валидный непустой договор шапки переопределяет основание", async () => {
    const { client } = await setup();
    const override = { ...args, contractRef: OTHER_CT };
    const preview = await call(client, { ...override, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    expect((preview.structuredContent?.payload as Record<string, unknown>)["ДоговорКонтрагента_Key"]).toBe(
      OTHER_CT,
    );
  });

  it("явный zero договора/строки не принимает null или непустой GUID", async () => {
    for (const mutate of [
      (actual: Record<string, unknown>) => {
        actual["ДоговорКонтрагента_Key"] = null;
        return actual;
      },
      (actual: Record<string, unknown>) => {
        actual["ДоговорКонтрагента_Key"] = CTR;
        return actual;
      },
      (actual: Record<string, unknown>) => {
        const rows = actual["Авансы"] as Array<Record<string, unknown>>;
        rows[0]["Контрагент_Key"] = null;
        return actual;
      },
    ]) {
      const { client } = await setup({ mutate });
      const zeroArgs = {
        ...args,
        contractRef: ZERO,
        advanceRows: [{ ...args.advanceRows[0], counterpartyRef: ZERO }],
      };
      const preview = await call(client, { ...zeroArgs, confirm: false });
      const created = await call(client, {
        ...zeroArgs,
        confirm: true,
        operationId: preview.structuredContent?.operationId,
      });
      expect(created.isError).toBe(true);
    }
  });

  it("allowlisted header fields сохраняют false/0/empty/value и комментарий", async () => {
    for (const marker of ["false", "true"]) {
      const { client } = await setup({ env: { ODATA_WRITE_OPERATION_MARKER: marker } });
      const extended = {
        ...args,
        contractRef: ZERO,
        currencyRef: ZERO,
        comment: "  exact comment  ",
        numberRepresentation: "NR-1",
        headerPaymentNumber: "HN-1",
        headerVatRate: "БезНДС",
        headerPaymentDate: "2026-10-01T12:13:14",
        issued: false,
        invoiceNotIssued: false,
        invoiceWithoutVat: false,
        manualCorrection: false,
        correction: false,
        sub0: true,
        initialBalanceVat: false,
        representationMethod: 0,
        sum: 1,
        vat: 0,
      };
      const preview = await call(client, { ...extended, confirm: false });
      expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
      const payload = preview.structuredContent?.payload as Record<string, unknown>;
      expect(payload).toMatchObject({
        Комментарий: "  exact comment  ",
        ПредставлениеНомера: "NR-1",
        НомерПлатежноРасчетногоДокумента: "HN-1",
        СтавкаНДС: "БезНДС",
        ДатаПлатежноРасчетногоДокумента: "2026-10-01T12:13:14",
        Выставлен: false,
        СчетФактураНеВыставляется: false,
        СчетФактураБезНДС: false,
        РучнаяКорректировка: false,
        Исправление: false,
        Под0: true,
        СформированПриВводеНачальныхОстатковНДС: false,
        КодСпособаВыставления: 0,
        Сумма: 1,
        СуммаНДС: 0,
      });
      const created = await call(client, {
        ...extended,
        confirm: true,
        operationId: preview.structuredContent?.operationId,
      });
      expect(created.isError, JSON.stringify(created)).toBeFalsy();
    }
  });

  it("basis considerCorrection=false и явно пустые таблицы сохраняются", async () => {
    const { client } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload["ДокументыОбОтгрузке"]).toEqual([]);
    expect(payload["КорректируемыеНомераСтрок"]).toEqual([]);
    const basis = (payload["ДокументыОснования"] as Array<Record<string, unknown>>)[0];
    expect(basis["УчитыватьИсправлениеИсходногоДокумента"]).toBe(false);
  });

  it("omission таблиц не равен явной пустоте", async () => {
    const { client } = await setup();
    const { emptyShipmentDocuments, emptyCorrectableLineNumbers, ...omitted } = args;
    void emptyShipmentDocuments;
    void emptyCorrectableLineNumbers;
    const preview = await call(client, { ...omitted, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).not.toHaveProperty("ДокументыОбОтгрузке");
    expect(payload).not.toHaveProperty("КорректируемыеНомераСтрок");
  });

  it("paymentDocuments принимает date и date-time, сохраняя время", async () => {
    const { client } = await setup();
    const paymentArgs = {
      ...args,
      paymentDocuments: [
        {
          number: "PAY-1",
          date: "2026-10-01T12:13:14",
          invoiceNumber: "INV-1",
          invoiceDate: "2026-10-02T01:02:03",
        },
      ],
    };
    const preview = await call(client, { ...paymentArgs, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const rows = (preview.structuredContent?.payload as Record<string, unknown>)[
      "ПлатежноРасчетныеДокументы"
    ] as Array<Record<string, unknown>>;
    expect(rows[0]).toMatchObject({
      ДатаДокумента: "2026-10-01T12:13:14",
      ДатаСчетаФактурыНаАванс: "2026-10-02T01:02:03",
    });
    const created = await call(client, {
      ...paymentArgs,
      confirm: true,
      operationId: preview.structuredContent?.operationId,
    });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
  });

  it("readback basis false/таблиц: missing, changed false или непустая таблица — failed", async () => {
    for (const mutate of [
      (actual: Record<string, unknown>) => {
        const rows = actual["ДокументыОснования"] as Array<Record<string, unknown>>;
        delete rows[0]["УчитыватьИсправлениеИсходногоДокумента"];
        return actual;
      },
      (actual: Record<string, unknown>) => {
        const rows = actual["ДокументыОснования"] as Array<Record<string, unknown>>;
        rows[0]["УчитыватьИсправлениеИсходногоДокумента"] = true;
        return actual;
      },
      (actual: Record<string, unknown>) => {
        actual["ДокументыОбОтгрузке"] = [{ НомерДокумента: "X" }];
        return actual;
      },
    ]) {
      const { client } = await setup({ mutate });
      const preview = await call(client, { ...args, confirm: false });
      const created = await call(client, {
        ...args,
        confirm: true,
        operationId: preview.structuredContent?.operationId,
      });
      expect(created.isError).toBe(true);
    }
  });

  it("номер строки после default не дублируется", async () => {
    const { client } = await setup();
    const bad = {
      ...args,
      basisDocuments: [
        { ...args.basisDocuments[0], lineNumber: 2 },
        { ...args.basisDocuments[0], lineNumber: undefined },
      ],
    };
    const preview = await call(client, { ...bad, confirm: false });
    expect(preview.isError).toBe(true);
    expect(JSON.stringify(preview)).toContain("повторяются");
  });

  it("неполная пара basisRef/basisType в строке аванса отклоняется", async () => {
    const { client } = await setup();
    for (const row of [
      { ...args.advanceRows[0], basisType: undefined },
      { ...args.advanceRows[0], basisRef: undefined },
      { ...args.advanceRows[0], basisType: "StandardODATA.Document_ПлатежноеПоручение" },
    ]) {
      const preview = await call(client, { ...args, advanceRows: [row], confirm: false });
      expect(preview.isError).toBe(true);
    }
  });

  it("аванс на банковском основании проходит с таблицей Авансы", async () => {
    const { client } = await setup();
    const bankArgs = {
      ...args,
      basisDocuments: [
        {
          ...args.basisDocuments[0],
          type: "StandardODATA.Document_ПоступлениеНаРасчетныйСчет",
          sourceDate: "2026-10-01",
        },
      ],
      advanceRows: [
        {
          ...args.advanceRows[0],
          basisType: "StandardODATA.Document_ПоступлениеНаРасчетныйСчет",
        },
      ],
    };
    const preview = await call(client, { ...bankArgs, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    expect((preview.structuredContent?.payload as Record<string, unknown>)["Авансы"]).toHaveLength(1);
  });

  it("повтор того же operationId не создаёт второй документ", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    const id = preview.structuredContent?.operationId;
    const first = await call(client, { ...args, confirm: true, operationId: id });
    const second = await call(client, { ...args, confirm: true, operationId: id });
    expect(first.isError, JSON.stringify(first)).toBeFalsy();
    expect(second.isError, JSON.stringify(second)).toBeFalsy();
    expect(postCount(fake)).toBe(1);
  });
});
