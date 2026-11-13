import { describe, expect, it, vi, afterEach } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

const ORG = "11111111-1111-4111-8111-111111111111";
const ACC = "11111111-1111-4111-8111-111111111111";
const NOM = "22222222-2222-4222-8222-222222222222";
const RECIPIENT = "33333333-3333-4333-8333-333333333333";
const EXISTING = "44444444-4444-4444-8444-444444444444";

const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(cleanup.splice(0).map((fn) => fn().catch(() => undefined)));
});

async function setup(options: { docs?: Record<string, unknown>[]; readback?: Record<string, unknown> } = {}) {
  const fake = fake1C(
    { READ_ONLY: "false", ODATA_WRITABLE: "true" },
    {
      data: (q: FakeQuery) =>
        q.base === "Document_СписаниеСРасчетногоСчета" ? (options.docs ?? []) : undefined,
    },
  );
  const original = fake.conn.client.getEntity.bind(fake.conn.client);
  vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
    if (path.startsWith("Document_СписаниеСРасчетногоСчета(") && !path.includes("$select")) {
      if (options.readback) return { DeletionMark: false, ...options.readback };
      const actual = await original(path);
      return { DeletionMark: false, ...actual };
    }
    return original(path);
  });
  const server = createServer(fake.ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "historical-bank-commission-test", version: "1" });
  await client.connect(b);
  cleanup.push(async () => {
    await client.close().catch(() => undefined);
    await server.close().catch(() => undefined);
  });
  return { fake, client };
}

const call = async (client: Client, args: Record<string, unknown>) =>
  (await client.callTool({ name: "write.money.create_bank_writeoff", arguments: args })) as {
    isError?: boolean;
    structuredContent?: Record<string, unknown>;
  };
const postCount = (fake: ReturnType<typeof fake1C>): number =>
  fake.requests.filter((r) => r.method === "POST").length;

const args = {
  organization: ORG,
  operationKind: "КомиссияБанка",
  amount: 100,
  orgBankAccount: ACC,
  number: "BC-1",
  date: "2026-11-01",
  time: "10:20:30",
  incomingNumber: "IN-BC-1",
  incomingDate: "2026-10-31",
  incomingTime: "09:08:07",
  bankAccountingAccountRef: ACC,
  debitAccountRef: ACC,
  cashflowItemRef: ACC,
  purposeText: "Комиссия банка",
  comment: "Историческая комиссия",
  recipientBankAccountRef: RECIPIENT,
  subconto: [{ ref: NOM, type: "StandardODATA.Catalog_Номенклатура" }],
};

describe("write.money.create_bank_writeoff: историческая КомиссияБанка", () => {
  it("точные реквизиты, типизированное субконто и пустая расшифровка", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      Number: "BC-1",
      Date: "2026-11-01T10:20:30",
      ВидОперации: "КомиссияБанка",
      НомерВходящегоДокумента: "IN-BC-1",
      ДатаВходящегоДокумента: "2026-10-31T09:08:07",
      СчетБанк_Key: ACC,
      СчетУчетаРасчетовСКонтрагентом_Key: ACC,
      СтатьяДвиженияДенежныхСредств_Key: ACC,
      СчетКонтрагента_Key: RECIPIENT,
      СубконтоДт1: NOM,
      СубконтоДт1_Type: "StandardODATA.Catalog_Номенклатура",
    });
    expect(payload["РасшифровкаПлатежа"]).toEqual([]);
    expect(postCount(fake)).toBe(0);
  });

  it("подтверждает, проверяет все поля и оставляет черновик", async () => {
    const { client, fake } = await setup({
      readback: {
        Number: "BC-1",
        Date: "2026-11-01T10:20:30",
        Posted: false,
        ВидОперации: "КомиссияБанка",
        Организация_Key: ORG,
        СчетОрганизации_Key: ACC,
        СчетБанк_Key: ACC,
        СчетУчетаРасчетовСКонтрагентом_Key: ACC,
        СубконтоДт1: NOM,
        СубконтоДт1_Type: "StandardODATA.Catalog_Номенклатура",
        СчетКонтрагента_Key: RECIPIENT,
        НомерВходящегоДокумента: "IN-BC-1",
        ДатаВходящегоДокумента: "2026-10-31T09:08:07",
        СуммаДокумента: 100,
        НазначениеПлатежа: "Комиссия банка",
        СтатьяДвиженияДенежныхСредств_Key: ACC,
        Комментарий: "Историческая комиссия",
        РасшифровкаПлатежа: [],
      },
    });
    const preview = await call(client, { ...args, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    const created = await call(client, { ...args, confirm: true, operationId });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({ created: true, verification: "ok" });
    expect(created.structuredContent?.verifiedFields).toEqual(
      expect.arrayContaining(["Number", "СчетБанк_Key", "СубконтоДт1_Type", "РасшифровкаПлатежа"]),
    );
    expect(postCount(fake)).toBe(1);
  });

  it("дубль по организации/году/номеру отклоняется до POST", async () => {
    const { client, fake } = await setup({
      docs: [{ Ref_Key: EXISTING, Number: "BC-1", Date: "2026-01-01T00:00:00", Организация_Key: ORG }],
    });
    const preview = await call(client, { ...args, confirm: false });
    expect(preview.isError).toBe(true);
    expect(JSON.stringify(preview)).toContain("уже занят");
    expect(postCount(fake)).toBe(0);
  });

  it("обычный вызов сохраняет прежний путь", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, {
      organization: ORG,
      operationKind: "ПрочееСписание",
      amount: 100,
      orgBankAccount: ACC,
      confirm: false,
    });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    expect((preview.structuredContent?.payload as Record<string, unknown>)["Number"]).toBeUndefined();
    expect(postCount(fake)).toBe(0);
  });
});
