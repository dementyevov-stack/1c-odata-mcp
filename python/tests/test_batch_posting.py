import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from batch_posting import (  # noqa: E402
    BatchExecutor,
    PlanError,
    Register404,
    ResponseLost,
    plan_fingerprint,
    validate_plan,
)

ORG = "ORG-SYNTHETIC"


class FakeFreshTransport:
    """Fake HTTP-ish OData transport; no network and no real accounting rows."""

    def __init__(self, docs, registers, vat_books, journals):
        self.docs = {row["Ref_Key"]: dict(row) for row in docs}
        self.registers = {tuple(key): [dict(r) for r in rows] for key, rows in registers.items()}
        self.vat_books = {key: list(value) for key, value in vat_books.items()}
        self.journals = {key: list(value) for key, value in journals.items()}
        self.posts = []
        self.lose_on = set()
        self.drop_after_lost = set()
        self.mutate_after_post = {}

    def validate_context(self, organization):
        return organization == ORG

    def find_by_source(self, source_key):
        rows = [dict(doc) for doc in self.docs.values() if doc.get("source_key") == source_key]
        return {"receipt": rows[0] if rows else None, "invoice": rows[1] if len(rows) > 1 else None}

    def get_document(self, ref):
        if ref not in self.docs:
            raise KeyError(ref)
        return dict(self.docs[ref])

    def post(self, kind, ref):
        self.posts.append((kind, ref))
        self.docs[ref]["posted"] = True
        for key, value in self.mutate_after_post.items():
            self.docs[ref][key] = value
        if ref in self.lose_on:
            if ref in self.drop_after_lost:
                del self.docs[ref]
            raise ResponseLost("synthetic response lost")

    def list_register(self, recorder_ref, recorder_type):
        key = (recorder_ref, recorder_type)
        if key not in self.registers:
            raise Register404()
        return [dict(row) for row in self.registers[key]]

    def vat_book(self, invoice_ref):
        return list(self.vat_books.get(invoice_ref, []))

    def journal(self, invoice_ref):
        return list(self.journals.get(invoice_ref, []))


def doc(ref, source_key, *, advances=(), dates=(), kind="document"):
    return {
        "Ref_Key": ref,
        "source_key": source_key,
        "kind": kind,
        "posted": False,
        "organization": ORG,
        "advances": list(advances),
        "advance_dates": list(dates),
        "stable": "same",
    }


def pair(source_key, receipt_ref, invoice_ref, advance, date, *, missing_invoice_register=False, deduction=True):
    register_rows = [{"Recorder": receipt_ref, "Recorder_Type": "AccountingRegister_Example", "Сумма": 100}]
    invoice_rows = [] if missing_invoice_register else [
        {"Recorder": invoice_ref, "Recorder_Type": "AccountingRegister_Example", "Сумма": 100}
    ]
    return {
        "source_key": source_key,
        "organization": ORG,
        "receipt_ref": receipt_ref,
        "invoice_ref": invoice_ref,
        "deduction": deduction,
        "advances": [{"ref": advance, "date": date}],
        "expected": {
            "registers": [
                {
                    "recorder_ref": receipt_ref,
                    "recorder_type": "AccountingRegister_Example",
                    "rows": register_rows,
                },
                {
                    "recorder_ref": invoice_ref,
                    "recorder_type": "AccountingRegister_Example",
                    "rows": invoice_rows,
                    "allow_404": missing_invoice_register,
                },
            ],
            "vat": 20.0,
            "vat_book": [f"VB-{source_key}"],
            "journal": [f"J-{source_key}"],
        },
    }


def plan_a():
    return {
        "schema_version": 1,
        "profile": "fresh-purchase-invoice-v1",
        "batches": [
            {
                "id": "A",
                "pairs": [pair("A1", "R-A1", "I-A1", "ADV-A1", "2026-01-01T10:00:00")],
            }
        ],
    }


def plan_b():
    return {
        "schema_version": 1,
        "profile": "fresh-purchase-invoice-v1",
        "batches": [
            {
                "id": "B",
                "pairs": [
                    pair("B1", "R-B1", "I-B1", "ADV-SHARED", "2026-01-02T11:00:00", deduction=True),
                    pair(
                        "B2",
                        "R-B2",
                        "I-B2",
                        "ADV-SHARED",
                        "2026-01-02T11:00:00",
                        missing_invoice_register=True,
                        deduction=False,
                    ),
                ],
            }
        ],
    }


def seed(plan):
    docs = []
    registers = {}
    vat_books = {}
    journals = {}
    for batch in plan["batches"]:
        for p in batch["pairs"]:
            docs.append(doc(p["receipt_ref"], f"receipt:{p['source_key']}", kind="receipt"))
            docs.append(
                {
                    **doc(
                        p["invoice_ref"],
                        f"invoice:{p['source_key']}",
                        advances=[a["ref"] for a in p["advances"]],
                        dates=[a["date"] for a in p["advances"]],
                        kind="invoice",
                    ),
                    "deduction": p.get("deduction", True),
                    "vat": p["expected"]["vat"],
                }
            )
            for reg in p["expected"]["registers"]:
                if not (reg.get("allow_404") and not reg["rows"]):
                    registers[(reg["recorder_ref"], reg["recorder_type"])] = reg["rows"]
            vat_books[p["invoice_ref"]] = p["expected"]["vat_book"]
            journals[p["invoice_ref"]] = p["expected"]["journal"]
    return FakeFreshTransport(docs, registers, vat_books, journals)


class BatchPlanTests(unittest.TestCase):
    def test_schema_and_fingerprint(self):
        plan = plan_a()
        validate_plan(plan)
        self.assertEqual(len(plan_fingerprint(plan)), 64)
        bad = json.loads(json.dumps(plan))
        bad["schema_version"] = 2
        with self.assertRaises(PlanError):
            validate_plan(bad)

    def test_receipt_hash_mismatch_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            executor = BatchExecutor(seed(plan_a()), tmp)
            executor.run(plan_a(), "A")
            changed = plan_a()
            changed["batches"][0]["pairs"][0]["receipt_ref"] = "R-OTHER"
            with self.assertRaises(PlanError):
                executor.run(changed, "A")


class BatchExecutionTests(unittest.TestCase):
    def run_plan(self, plan, batch):
        transport = seed(plan)
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        executor = BatchExecutor(transport, tmp.name)
        return executor, transport, executor.run(plan, batch), Path(tmp.name)

    def test_batches_a_and_b_use_same_executor_and_are_confirmed(self):
        with tempfile.TemporaryDirectory() as tmp:
            a_transport = seed(plan_a())
            b_transport = seed(plan_b())
            executor = BatchExecutor(a_transport, tmp)
            summary_a = executor.run(plan_a(), "A")
            executor.transport = b_transport
            summary_b = executor.run(plan_b(), "B")
            self.assertEqual(summary_a["counts"]["confirmed"], 1)
            self.assertEqual(summary_b["counts"]["confirmed"], 2)
            self.assertTrue((Path(tmp) / "A" / "summary.json").exists())
            self.assertTrue((Path(tmp) / "B" / "summary.md").exists())
            self.assertEqual(len(a_transport.posts), 2)
            self.assertEqual(len(b_transport.posts), 4)

    def test_repeat_confirmed_receipt_does_not_post(self):
        with tempfile.TemporaryDirectory() as tmp:
            transport = seed(plan_a())
            executor = BatchExecutor(transport, tmp)
            executor.run(plan_a(), "A")
            posts = len(transport.posts)
            summary = executor.run(plan_a(), "A")
            self.assertEqual(summary["counts"]["skipped"], 1)
            self.assertEqual(len(transport.posts), posts)

    def test_lost_patch_response_is_recovered_by_readback(self):
        with tempfile.TemporaryDirectory() as tmp:
            transport = seed(plan_a())
            transport.lose_on = {"R-A1"}
            executor = BatchExecutor(transport, tmp)
            summary = executor.run(plan_a(), "A")
            self.assertEqual(summary["counts"]["confirmed"], 1)
            self.assertTrue(summary["receipts"][0]["recovered"])
            self.assertEqual(len(transport.posts), 2)  # one receipt + one invoice, no duplicate

    def test_lost_and_missing_document_is_unknown_without_retry(self):
        with tempfile.TemporaryDirectory() as tmp:
            transport = seed(plan_a())
            transport.lose_on = {"R-A1"}
            transport.drop_after_lost = {"R-A1"}
            executor = BatchExecutor(transport, tmp)
            summary = executor.run(plan_a(), "A")
            self.assertEqual(summary["counts"]["unknown"], 1)
            self.assertEqual(len(transport.posts), 1)
            again = executor.run(plan_a(), "A")
            self.assertEqual(again["counts"]["unknown"], 1)
            self.assertEqual(len(transport.posts), 1)

    def test_changed_document_field_is_conflict(self):
        with tempfile.TemporaryDirectory() as tmp:
            transport = seed(plan_a())
            transport.mutate_after_post = {"unexpected": 1}
            executor = BatchExecutor(transport, tmp)
            summary = executor.run(plan_a(), "A")
            self.assertEqual(summary["counts"]["conflict"], 1)
            self.assertIn("receipt changed", " ".join(summary["receipts"][0]["reasons"]))

    def test_nonempty_expected_register_404_is_conflict(self):
        with tempfile.TemporaryDirectory() as tmp:
            transport = seed(plan_b())
            # Remove B2 invoice register although the plan expects empty rows but
            # allow_404 is true: this is allowed. Make a required register absent.
            del transport.registers[("R-B1", "AccountingRegister_Example")]
            executor = BatchExecutor(transport, tmp)
            summary = executor.run(plan_b(), "B")
            self.assertEqual(summary["counts"]["conflict"], 1)

    def test_conflicting_advance_reference_is_conflict(self):
        with tempfile.TemporaryDirectory() as tmp:
            transport = seed(plan_a())
            transport.docs["I-A1"]["advances"] = ["ADV-WRONG"]
            executor = BatchExecutor(transport, tmp)
            summary = executor.run(plan_a(), "A")
            self.assertEqual(summary["counts"]["conflict"], 1)

    def test_summaries_match_detail_receipts(self):
        with tempfile.TemporaryDirectory() as tmp:
            transport = seed(plan_b())
            executor = BatchExecutor(transport, tmp)
            summary = executor.run(plan_b(), "B")
            on_disk = json.loads((Path(tmp) / "B" / "summary.json").read_text("utf-8"))
            self.assertEqual(on_disk["counts"], summary["counts"])
            for receipt in summary["receipts"]:
                detail = json.loads((Path(tmp) / "B" / f"{receipt['source_key']}.json").read_text("utf-8"))
                self.assertEqual(detail["status"], receipt["status"])
            markdown = (Path(tmp) / "B" / "summary.md").read_text("utf-8")
            self.assertIn("B1", markdown)
            self.assertIn("B2", markdown)


if __name__ == "__main__":
    unittest.main()
