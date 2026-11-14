"""Reusable synthetic batch executor for Fresh OData purchase/invoice pairs.

This module contains no HTTP implementation and no accounting rules beyond the
explicit plan/readback checks. A real operator supplies ``transport``; tests
supply an isolated fake. Batches A and B are ordinary data for the same code.
"""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Mapping

SCHEMA_VERSION = 1
PROFILE = "fresh-purchase-invoice-v1"
STATUSES = {"confirmed", "skipped", "conflict", "unknown"}


class PlanError(ValueError):
    pass


class ResponseLost(Exception):
    """Write may have been applied; response was lost."""


class Register404(Exception):
    def __init__(self, record_type: str = "", movements: list[Any] | None = None):
        super().__init__("register set not found (404)")
        self.record_type = record_type
        self.movements = movements or []


def canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def plan_fingerprint(plan: Mapping[str, Any]) -> str:
    return hashlib.sha256(canonical_json(plan).encode("utf-8")).hexdigest()


def _require(mapping: Mapping[str, Any], key: str, kind: type) -> Any:
    value = mapping.get(key)
    if not isinstance(value, kind):
        raise PlanError(f"{key}: expected {kind.__name__}")
    return value


def validate_plan(plan: Mapping[str, Any]) -> None:
    if not isinstance(plan, Mapping):
        raise PlanError("plan must be an object")
    if plan.get("schema_version") != SCHEMA_VERSION:
        raise PlanError("unsupported schema_version")
    if plan.get("profile") != PROFILE:
        raise PlanError("unsupported profile")
    batches = _require(plan, "batches", list)
    if not batches:
        raise PlanError("batches must not be empty")
    seen_batches: set[str] = set()
    seen_sources: set[str] = set()
    for batch in batches:
        if not isinstance(batch, Mapping):
            raise PlanError("batch must be an object")
        batch_id = _require(batch, "id", str)
        if batch_id in seen_batches:
            raise PlanError(f"duplicate batch id {batch_id}")
        seen_batches.add(batch_id)
        pairs = _require(batch, "pairs", list)
        if not pairs:
            raise PlanError(f"batch {batch_id}: pairs must not be empty")
        for pair in pairs:
            if not isinstance(pair, Mapping):
                raise PlanError("pair must be an object")
            source_key = _require(pair, "source_key", str)
            if source_key in seen_sources:
                raise PlanError(f"duplicate source_key {source_key}")
            seen_sources.add(source_key)
            for key in ("organization", "receipt_ref", "invoice_ref"):
                _require(pair, key, str)
            advances = _require(pair, "advances", list)
            for advance in advances:
                if not isinstance(advance, Mapping) or not isinstance(advance.get("ref"), str):
                    raise PlanError("advance.ref must be a string")
                if not re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}", str(advance.get("date", ""))):
                    raise PlanError("advance.date must be YYYY-MM-DDTHH:MM:SS")
            expected = _require(pair, "expected", Mapping)
            if not isinstance(expected.get("registers"), list):
                raise PlanError("expected.registers must be a list")
            if "deduction" in pair and not isinstance(pair["deduction"], bool):
                raise PlanError("deduction must be boolean")
            for key in ("vat_book", "journal"):
                if not isinstance(expected.get(key), list):
                    raise PlanError(f"expected.{key} must be a list")
            if "vat" in expected and not isinstance(expected["vat"], (int, float)):
                raise PlanError("expected.vat must be numeric")


@dataclass
class Receipt:
    batch: str
    source_key: str
    status: str
    plan_hash: str
    receipt_ref: str
    invoice_ref: str
    posts: int = 0
    recovered: bool = False
    reasons: list[str] = field(default_factory=list)
    checks: dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        return {
            "schema_version": SCHEMA_VERSION,
            "profile": PROFILE,
            "batch": self.batch,
            "source_key": self.source_key,
            "status": self.status,
            "plan_hash": self.plan_hash,
            "receipt_ref": self.receipt_ref,
            "invoice_ref": self.invoice_ref,
            "posts": self.posts,
            "recovered": self.recovered,
            "reasons": self.reasons,
            "checks": self.checks,
        }


def _equal(left: Any, right: Any) -> bool:
    if type(left) is not type(right):
        return False
    if isinstance(left, Mapping):
        return set(left) == set(right) and all(_equal(left[k], right[k]) for k in left)
    if isinstance(left, list):
        return len(left) == len(right) and all(_equal(a, b) for a, b in zip(left, right, strict=True))
    return left == right


def _changed(before: Mapping[str, Any], after: Mapping[str, Any], ignore: set[str]) -> set[str]:
    return {
        key
        for key in set(before) | set(after)
        if key not in ignore and (key not in before or key not in after or not _equal(before[key], after[key]))
    }


class BatchExecutor:
    """Runs a validated plan through a transport; never retries a POST."""

    TECH_FIELDS = {"posted", "Posted", "DataVersion", "version"}

    def __init__(self, transport: Any, receipt_dir: str | Path):
        self.transport = transport
        self.receipt_dir = Path(receipt_dir)

    def _receipt_path(self, batch_id: str, source_key: str) -> Path:
        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", source_key)
        return self.receipt_dir / batch_id / f"{safe}.json"

    def _read_receipt(self, batch_id: str, source_key: str, plan_hash: str) -> Receipt | None:
        path = self._receipt_path(batch_id, source_key)
        if not path.exists():
            return None
        data = json.loads(path.read_text("utf-8"))
        if data.get("plan_hash") != plan_hash:
            raise PlanError("receipt plan hash differs from current plan")
        return Receipt(**{k: data[k] for k in (
            "batch", "source_key", "status", "plan_hash", "receipt_ref", "invoice_ref", "posts", "recovered", "reasons", "checks"
        )})

    def _write_receipt(self, receipt: Receipt) -> Path:
        path = self._receipt_path(receipt.batch, receipt.source_key)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(receipt.as_dict(), ensure_ascii=False, indent=2), "utf-8")
        return path

    def _post_and_recover(self, kind: str, ref: str) -> tuple[Mapping[str, Any] | None, str]:
        try:
            self.transport.post(kind, ref)
            recovered = False
        except ResponseLost:
            recovered = True
        try:
            doc = self.transport.get_document(ref)
        except Exception:
            return None, "unknown"
        if not doc.get("posted"):
            return doc, "unknown"
        return doc, "recovered" if recovered else "posted"

    def _run_pair(self, batch_id: str, pair: Mapping[str, Any], plan_hash: str) -> Receipt:
        source_key = str(pair["source_key"])
        receipt_ref = str(pair["receipt_ref"])
        invoice_ref = str(pair["invoice_ref"])
        existing = self._read_receipt(batch_id, source_key, plan_hash)
        if existing:
            if existing.status in {"confirmed", "skipped"}:
                existing.status = "skipped"
                existing.reasons.append("confirmed receipt already exists; no POST")
                self._write_receipt(existing)
            # conflict/unknown receipts are retained as-is: no retry POST.
            return existing

        receipt = Receipt(batch_id, source_key, "unknown", plan_hash, receipt_ref, invoice_ref)
        if not self.transport.validate_context(str(pair["organization"])):
            receipt.status = "conflict"
            receipt.reasons.append("publication/organization check failed")
            self._write_receipt(receipt)
            return receipt

        found = self.transport.find_by_source(source_key)
        if found.get("ambiguous") or found.get("incomplete"):
            receipt.status = "conflict"
            receipt.reasons.append("ambiguous or incomplete source lookup")
            self._write_receipt(receipt)
            return receipt

        try:
            before_receipt = self.transport.get_document(receipt_ref)
            before_invoice = self.transport.get_document(invoice_ref)
        except Exception:
            receipt.status = "unknown"
            receipt.reasons.append("initial card readback unavailable")
            self._write_receipt(receipt)
            return receipt

        if not before_receipt.get("posted"):
            doc, outcome = self._post_and_recover("receipt", receipt_ref)
            receipt.posts += 1
            if outcome == "unknown" or doc is None:
                receipt.status = "unknown"
                receipt.reasons.append("receipt POST outcome unknown")
                self._write_receipt(receipt)
                return receipt
            if outcome == "recovered":
                receipt.recovered = True
            after_receipt = doc
        else:
            after_receipt = before_receipt

        if not before_invoice.get("posted"):
            doc, outcome = self._post_and_recover("invoice", invoice_ref)
            receipt.posts += 1
            if outcome == "unknown" or doc is None:
                receipt.status = "unknown"
                receipt.reasons.append("invoice POST outcome unknown")
                self._write_receipt(receipt)
                return receipt
            if outcome == "recovered":
                receipt.recovered = True
            after_invoice = doc
        else:
            after_invoice = before_invoice

        reasons: list[str] = []
        changed_receipt = _changed(before_receipt, after_receipt, self.TECH_FIELDS)
        changed_invoice = _changed(before_invoice, after_invoice, self.TECH_FIELDS)
        if changed_receipt:
            reasons.append(f"receipt changed: {sorted(changed_receipt)}")
        if changed_invoice:
            reasons.append(f"invoice changed: {sorted(changed_invoice)}")
        if not after_receipt.get("posted") or not after_invoice.get("posted"):
            reasons.append("document is not posted after readback")
        if pair.get("deduction") is not None and after_invoice.get("deduction") is not pair["deduction"]:
            reasons.append("invoice deduction flag differs")
        if after_invoice.get("advances") != [a["ref"] for a in pair["advances"]]:
            reasons.append("advances do not match exact settlement documents")
        if after_invoice.get("advance_dates") != [a["date"] for a in pair["advances"]]:
            reasons.append("advance time differs")

        expected = pair["expected"]
        for reg in expected["registers"]:
            try:
                rows = self.transport.list_register(reg["recorder_ref"], reg["recorder_type"])
            except Register404 as exc:
                allowed = (
                    bool(reg.get("allow_404"))
                    and bool(after_invoice.get("posted"))
                    and not reg.get("rows", [])
                    and not exc.record_type
                    and not exc.movements
                )
                if allowed:
                    rows = []
                else:
                    reasons.append("register 404 is not the allowed empty expected set")
                    rows = None
            if rows is not None and not _equal(rows, reg["rows"]):
                reasons.append(f"register rows differ: {reg['recorder_type']}")
        if "vat" in expected and after_invoice.get("vat") != expected["vat"]:
            reasons.append("presented VAT differs")
        if self.transport.vat_book(invoice_ref) != expected["vat_book"]:
            reasons.append("VAT book differs")
        if self.transport.journal(invoice_ref) != expected["journal"]:
            reasons.append("invoice journal differs")

        receipt.checks = {
            "receipt_changed": sorted(changed_receipt),
            "invoice_changed": sorted(changed_invoice),
            "registers": len(expected["registers"]),
            "vat_book": expected["vat_book"],
            "journal": expected["journal"],
        }
        receipt.reasons = reasons
        receipt.status = "conflict" if reasons else "confirmed"
        self._write_receipt(receipt)
        return receipt

    def run(self, plan: Mapping[str, Any], batch_id: str) -> dict[str, Any]:
        validate_plan(plan)
        fingerprint = plan_fingerprint(plan)
        batch = next((b for b in plan["batches"] if b["id"] == batch_id), None)
        if batch is None:
            raise PlanError(f"batch {batch_id} not found")
        receipts: list[Receipt] = []
        for pair in batch["pairs"]:
            receipts.append(self._run_pair(batch_id, pair, fingerprint))
        summary = {
            "schema_version": SCHEMA_VERSION,
            "profile": PROFILE,
            "batch": batch_id,
            "plan_hash": fingerprint,
            "counts": {status: sum(1 for r in receipts if r.status == status) for status in sorted(STATUSES)},
            "receipts": [r.as_dict() for r in receipts],
        }
        out = self.receipt_dir / batch_id
        out.mkdir(parents=True, exist_ok=True)
        (out / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), "utf-8")
        lines = [f"# Batch {batch_id}", "", f"plan: `{fingerprint}`", ""]
        for r in receipts:
            lines.append(f"- `{r.source_key}`: **{r.status}** — receipt `{r.receipt_ref}`, invoice `{r.invoice_ref}`")
        (out / "summary.md").write_text("\n".join(lines) + "\n", "utf-8")
        return summary
