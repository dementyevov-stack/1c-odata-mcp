"""Strict, small helpers for a POST -> GET -> optional PATCH -> GET flow.

The module deliberately does not perform HTTP or choose a transport. It only
answers two questions used by a Python catalog executor:

* does a fetched object contain every expected field with the exact type/value?
* which fields changed between two fetched snapshots?

Normalization is opt-in per field. Without an explicit normalizer, ``True`` and
``1`` are different values and an absent key is different from a key whose
value is ``None``.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Mapping
from dataclasses import dataclass, field
from typing import Any

MISSING = object()

Normalizers = Mapping[str, Callable[[Any], Any]]


def _same(expected: Any, actual: Any, normalizer: Callable[[Any], Any] | None) -> bool:
    if normalizer is not None:
        expected = normalizer(expected)
        actual = normalizer(actual)
    # bool is a subclass of int in Python; check exact types so False != 0.
    if type(expected) is not type(actual):
        return False
    if isinstance(expected, Mapping):
        if set(expected) != set(actual):
            return False
        return all(_same(expected[key], actual[key], None) for key in expected)
    if isinstance(expected, (list, tuple)):
        return len(expected) == len(actual) and all(
            _same(left, right, None) for left, right in zip(expected, actual, strict=True)
        )
    return expected == actual


def has_fields(
    row: Mapping[str, Any],
    expected: Mapping[str, Any],
    normalizers: Normalizers | None = None,
) -> bool:
    """True only if every expected key is present with the exact type and value.

    ``row`` may contain additional fields; they are ignored here. A missing key
    is never treated as an explicit ``None``. Use ``normalizers`` only for a
    documented field-specific representation, e.g. ``{"DataVersion": str}``
    for an Edm.Int64 serialized as a string.
    """
    normalizers = normalizers or {}
    return all(
        key in row and _same(value, row[key], normalizers.get(key))
        for key, value in expected.items()
    )


def changed_fields(
    before: Mapping[str, Any],
    after: Mapping[str, Any],
    ignore_fields: Iterable[str] = (),
    normalizers: Normalizers | None = None,
) -> set[str]:
    """Return every field whose key presence/type/value changed.

    Missing keys are changes: ``{"Optional": None}`` -> ``{}`` yields
    ``{"Optional"}``. Pass only field names that are expected patch targets or
    explicitly documented technical fields in ``ignore_fields``; the caller
    should compare the result with the intended patch set.
    """
    ignored = set(ignore_fields)
    normalizers = normalizers or {}
    changed: set[str] = set()
    for key in before.keys() | after.keys():
        if key in ignored:
            continue
        if key not in before or key not in after:
            changed.add(key)
        elif not _same(before[key], after[key], normalizers.get(key)):
            changed.add(key)
    return changed


class ResponseLostError(Exception):
    """The HTTP response was lost; the write may or may not have happened."""


class ReadUnavailableError(Exception):
    """A control GET is unavailable; a previous write's outcome is still unknown."""


@dataclass(frozen=True)
class SearchResult:
    """Result of an exact source-identity lookup."""

    rows: list[Mapping[str, Any]]
    complete: bool = True


@dataclass(frozen=True)
class CatalogAdapter:
    """Transport boundary implemented by a real executor or a fake HTTP test."""

    find_exact: Callable[[str], SearchResult]
    post: Callable[[Mapping[str, Any]], Mapping[str, Any]]
    get: Callable[[str], Mapping[str, Any]]
    patch: Callable[[str, Mapping[str, Any]], None]


@dataclass(frozen=True)
class CatalogCyclePlan:
    """Minimal plan for one catalog create and optional point correction."""

    source_key: str
    create_payload: Mapping[str, Any]
    expected_existing: Mapping[str, Any]
    expected_created: Mapping[str, Any]
    patch_changes: Mapping[str, Any] | None = None
    expected_patched: Mapping[str, Any] | None = None
    ignore_fields: tuple[str, ...] = ()


@dataclass(frozen=True)
class CycleResult:
    status: str  # existing | created | patched | unknown | conflict
    ref: str | None = None
    row: Mapping[str, Any] | None = None
    reason: str | None = None
    recovered: bool = False
    changed: frozenset[str] = field(default_factory=frozenset)


def execute_catalog_cycle(adapter: Any, plan: CatalogCyclePlan) -> CycleResult:
    """Run the exact POST -> GET -> optional PATCH -> GET cycle.

    ``adapter`` supplies the transport (real executor or fake HTTP in tests):
    ``find_exact(source_key) -> SearchResult``, ``post(payload) -> row``,
    ``get(ref) -> row``, ``patch(ref, changes) -> None``. It may raise
    ``ResponseLostError`` or ``ReadUnavailableError``.

    No POST is attempted after an ambiguous/incomplete search. A lost PATCH or
    GET is reported as ``unknown`` until the caller re-reads the same Ref_Key;
    it never triggers a retry POST.
    """
    search = adapter.find_exact(plan.source_key)
    if not search.complete:
        return CycleResult("conflict", reason="incomplete source lookup")
    if len(search.rows) > 1:
        return CycleResult("conflict", reason="ambiguous source lookup")
    if search.rows:
        existing = search.rows[0]
        ref = existing.get("Ref_Key")
        if not isinstance(ref, str) or not ref:
            return CycleResult("conflict", reason="existing object has no Ref_Key", row=existing)
        # A search hit is not proof: re-read the exact GUID and verify the
        # complete final postcondition before declaring the work already done.
        try:
            current = adapter.get(ref)
        except ReadUnavailableError:
            return CycleResult("unknown", ref=ref, reason="control GET for existing object unavailable")
        if not has_fields(current, {"Ref_Key": ref}):
            return CycleResult("conflict", ref=ref, reason="existing readback Ref_Key differs", row=current)
        if not has_fields(current, plan.expected_existing):
            return CycleResult("conflict", ref=ref, reason="existing object differs", row=current)
        return CycleResult("existing", ref=ref, row=current)

    try:
        posted = adapter.post(dict(plan.create_payload))
    except ResponseLostError:
        return CycleResult("unknown", reason="POST response lost")
    ref = posted.get("Ref_Key")
    if not isinstance(ref, str) or not ref:
        return CycleResult("conflict", reason="POST response has no Ref_Key")
    try:
        current = adapter.get(ref)
    except ReadUnavailableError:
        return CycleResult("unknown", ref=ref, reason="GET after POST unavailable")
    if not has_fields(current, {"Ref_Key": ref}):
        return CycleResult("conflict", ref=ref, reason="readback Ref_Key differs", row=current)
    if not has_fields(current, plan.expected_created):
        return CycleResult("conflict", ref=ref, reason="created readback differs", row=current)

    if plan.patch_changes is None:
        return CycleResult("created", ref=ref, row=current)

    response_lost = False
    try:
        adapter.patch(ref, dict(plan.patch_changes))
    except ResponseLostError:
        response_lost = True
    try:
        after = adapter.get(ref)
    except ReadUnavailableError:
        return CycleResult("unknown", ref=ref, reason="control GET after PATCH unavailable")
    if not has_fields(after, {"Ref_Key": ref}):
        return CycleResult("conflict", ref=ref, reason="readback Ref_Key differs", row=after)
    expected = dict(plan.expected_created)
    expected.update(plan.expected_patched or plan.patch_changes)
    if not has_fields(after, expected):
        return CycleResult("conflict", ref=ref, reason="patched fields differ", row=after)
    changed = changed_fields(
        current,
        after,
        ignore_fields=set(plan.patch_changes) | set(plan.ignore_fields),
    )
    if changed:
        return CycleResult(
            "conflict",
            ref=ref,
            reason=f"unexpected fields changed: {sorted(changed)}",
            row=after,
            changed=frozenset(changed),
        )
    return CycleResult(
        "patched",
        ref=ref,
        row=after,
        recovered=response_lost,
        reason="PATCH response lost; recovered by exact GET" if response_lost else None,
    )
