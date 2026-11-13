import { describe, expect, it, vi, afterEach } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

/** Синтетический исторический договор: явные пустые номер/дата и контрольный GET. */

const ORG = "11111111-1111-4111-8111-111111111111";
const CP = "22222222-2222-4222-8222-222222222222";
const EXISTING = "33333333-3333-4333-8333-333333333333";

const cleanup: Array<() => Promise<void>> = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(cleanup.splice(0).map((fn) => fn().catch(() => undefined)));
});

interface Options {
  contracts?: Record<string, unknown>[];
  readbackDate?: string;
}

async function setup(options: Options = {}) {
  const data = (q: FakeQuery): Record<string, unknown>[] | undefined =>
    q.base === "Catalog_ДоговорыКонтрагентов" ? (options.contracts ?? []) : undefined;
  const fake = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" }, { data });
  const original = fake.conn.client.getEntity.bind(fake.conn.client);
  vi.spyOn(fake.conn.client, "getEntity").mockImplementation(async (path: string) => {
    if (
      options.readbackDate &&
      path.startsWith("Catalog_ДоговорыКонтрагентов(") &&
      !path.includes("$select")
    ) {
      const actual = await original(path);
      return { ...actual, Дата: options.readbackDate };
    }
    return original(path);
  });
  const server = createServer(fake.ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "historical-contract-test", version: "1" });
  await client.connect(b);
  cleanup.push(async () => {
    await client.close().catch(() => undefined);
    await server.close().catch(() => undefined);
  });
  return { fake, client };
}

const call = async (client: Client, args: Record<string, unknown>) =>
  (await client.callTool({ name: "write.catalog.create_contract", arguments: args })) as {
    isError?: boolean;
    structuredContent?: Record<string, unknown>;
  };
const postCount = (fake: ReturnType<typeof fake1C>): number =>
  fake.requests.filter((r) => r.method === "POST").length;

const base = {
  organization: ORG,
  counterpartyRef: CP,
  kind: "Прочее" as const,
  name: "Основной договор",
  emptyNumber: true,
  emptyDate: true,
  currency: "643",
  headName: "Иванов И.И.",
  headPosition: "Генеральный директор",
};

describe("write.catalog.create_contract: исторический opt-in", () => {
  it("сериализует явно пустые Номер/Дата, валюту и флаги", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, { ...base, confirm: false });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      Description: "Основной договор",
      Номер: "",
      Дата: "0001-01-01T00:00:00",
      Owner_Key: CP,
      Организация_Key: ORG,
      ВидДоговора: "Прочее",
      Валютный: false,
      РуководительКонтрагента: "Иванов И.И.",
      ДолжностьРуководителяКонтрагента: "Генеральный директор",
    });
    expect(preview.structuredContent?.emptyFields).toEqual(["Номер", "Дата"]);
    expect(postCount(fake)).toBe(0);
  });

  it("переносит Code, СуммаВключаетНДС, флаги и способ ставки НДС", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, {
      ...base,
      code: "HIST-001",
      sumIncludesVat: true,
      contractClosed: false,
      contractSigned: false,
      vatMethod: "Автоматически",
      confirm: false,
    });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload).toMatchObject({
      Code: "HIST-001",
      СуммаВключаетНДС: true,
      ДоговорЗакрыт: false,
      ДоговорПодписан: false,
      СпособЗаполненияСтавкиНДС: "Автоматически",
    });
    expect(postCount(fake)).toBe(0);
  });

  it("подтверждает и сверяет все отправленные поля; дата 0001-01-01 считается пустой", async () => {
    const { client, fake } = await setup({ readbackDate: "0001-01-01T00:00:00" });
    const args = {
      ...base,
      code: "HIST-001",
      sumIncludesVat: true,
      contractClosed: false,
      contractSigned: false,
      vatMethod: "Автоматически",
    };
    const preview = await call(client, { ...args, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    expect(typeof operationId).toBe("string");
    const created = await call(client, { ...args, confirm: true, operationId });
    expect(created.isError, JSON.stringify(created)).toBeFalsy();
    expect(created.structuredContent).toMatchObject({
      created: true,
      verification: "ok",
      emptyFields: ["Номер", "Дата"],
    });
    expect(created.structuredContent?.verifiedFields).toEqual(
      expect.arrayContaining([
        "Code",
        "СуммаВключаетНДС",
        "ДоговорЗакрыт",
        "ДоговорПодписан",
        "СпособЗаполненияСтавкиНДС",
        "IsFolder",
        "DeletionMark",
      ]),
    );
    expect(postCount(fake)).toBe(1);
  });

  it("непустая дата после записи при emptyDate — failed с Ref_Key, без второго POST", async () => {
    const { client, fake } = await setup({ readbackDate: "2026-11-13T00:00:00" });
    const preview = await call(client, { ...base, confirm: false });
    const operationId = preview.structuredContent?.operationId;
    expect(typeof operationId).toBe("string");
    const created = await call(client, { ...base, confirm: true, operationId });
    expect(created.isError).toBe(true);
    expect(JSON.stringify(created)).toContain("Дата");
    expect(JSON.stringify(created)).toContain("Ref_Key");
    expect(postCount(fake)).toBe(1);
  });

  it("дубль в паре владелец/организация отклоняется до POST", async () => {
    const { client, fake } = await setup({
      contracts: [
        {
          Ref_Key: EXISTING,
          Description: "Основной договор",
          ВидДоговора: "Прочее",
          Owner_Key: CP,
          Организация_Key: ORG,
        },
      ],
    });
    const preview = await call(client, { ...base, confirm: false });
    expect(preview.isError).toBe(true);
    expect(JSON.stringify(preview)).toContain("уже существует");
    expect(postCount(fake)).toBe(0);
  });

  it("обычный вызов сохраняет «б/н» и текущую дату", async () => {
    const { client, fake } = await setup();
    const preview = await call(client, {
      organization: ORG,
      counterpartyRef: CP,
      kind: "Прочее",
      name: "Обычный договор",
      confirm: false,
    });
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    const payload = preview.structuredContent?.payload as Record<string, unknown>;
    expect(payload["Номер"]).toBe("б/н");
    expect(payload["Дата"]).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/);
    expect(payload).not.toHaveProperty("emptyFields");
    expect(postCount(fake)).toBe(0);
  });

  it("emptyNumber/emptyDate нельзя сочетать с number/date", async () => {
    const { client, fake } = await setup();
    const a = await call(client, { ...base, number: "X", confirm: false });
    const b = await call(client, { ...base, date: "2026-01-01", confirm: false });
    expect(a.isError).toBe(true);
    expect(b.isError).toBe(true);
    expect(postCount(fake)).toBe(0);
  });
});
