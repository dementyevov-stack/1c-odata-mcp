import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { ODataError } from "../src/odata/errors.js";
import { WriteOperationJournal } from "../src/odata/write-journal.js";

const dirs: string[] = [];
const journal = (dir = mkdtempSync(join(tmpdir(), "1c-journal-pending-"))) => {
  dirs.push(dir);
  return new WriteOperationJournal(dir, "db", "https://fake.invalid/odata/");
};
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

const payload = { Number: "SF-1", СуммаДокумента: 1 };
const hash = "request-hash";

describe("write journal: pending verification and atomicity", () => {
  it("pending входит в fresh-process summary", async () => {
    const dir = mkdtempSync(join(tmpdir(), "1c-journal-pending-"));
    const j = journal(dir);
    await j.prepare(
      "11111111-1111-4111-8111-111111111111",
      "Document_СчетФактураВыданный",
      payload,
      hash,
      undefined,
      true,
    );
    await j.execute(
      "11111111-1111-4111-8111-111111111111",
      "Document_СчетФактураВыданный",
      payload,
      hash,
      async () => ({ Ref_Key: "22222222-2222-4222-8222-222222222222" }),
    );
    const fresh = new WriteOperationJournal(dir, "db", "https://fake.invalid/odata/");
    expect((await fresh.summary()).uncertain).toContain("11111111-1111-4111-8111-111111111111");
  });

  it("concurrent mark/execute не откатывает журнал и не шлёт второй POST", async () => {
    const j = journal();
    const op = "33333333-3333-4333-8333-333333333333";
    await j.prepare(op, "Document_СчетФактураВыданный", payload, hash);
    const replace = (j as unknown as { replace(entry: Record<string, unknown>): Promise<void> }).replace.bind(
      j,
    );
    let release!: () => void;
    let entered!: () => void;
    const held = new Promise<void>((r) => (release = r));
    const ready = new Promise<void>((r) => (entered = r));
    (j as unknown as { replace(entry: Record<string, unknown>): Promise<void> }).replace = async (entry) => {
      if (entry.state === "prepared" && entry.verification === "pending") {
        entered();
        await held;
      }
      return replace(entry);
    };
    const marking = j.markVerificationPending(op);
    await ready;
    let sends = 0;
    const send = async () => {
      sends++;
      return { Ref_Key: "44444444-4444-4444-8444-444444444444" };
    };
    await j.execute(op, "Document_СчетФактураВыданный", payload, hash, send).catch(() => undefined);
    release();
    await marking;
    await j.execute(op, "Document_СчетФактураВыданный", payload, hash, send);
    expect(sends).toBe(1);
  });

  it("concurrent confirms шлют один раз и при success, и при unknown", async () => {
    for (const mode of ["success", "unknown"] as const) {
      const j = journal();
      const op =
        mode === "success" ? "55555555-5555-4555-8555-555555555555" : "66666666-6666-4666-8666-666666666666";
      await j.prepare(op, "Document_СчетФактураВыданный", payload, hash, undefined, true);
      let sends = 0;
      const send = async () => {
        sends++;
        if (mode === "unknown") throw new ODataError({ kind: "timeout", message: "timeout" });
        return { Ref_Key: "77777777-7777-4777-8777-777777777777" };
      };
      const results = await Promise.allSettled([
        j.execute(op, "Document_СчетФактураВыданный", payload, hash, send),
        j.execute(op, "Document_СчетФактураВыданный", payload, hash, send),
      ]);
      expect(sends).toBe(1);
      expect(results.some((r) => r.status === "rejected")).toBe(true);
    }
  });
});
