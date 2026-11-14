"""Minimal verifier for an existing accounting-register record.

The register entity list may be empty (or stale); the authoritative lookup is a
point GET by the recorder key and recorder type. This module never posts or
creates a record: it either proves the existing row satisfies the gate or
raises a mismatch/unknown result.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import Any

from catalog_verifier import has_fields


class RegisterRecordUnknown(Exception):
    """The exact recorder lookup did not return a row; no create is allowed."""


class RegisterRecordMismatch(Exception):
    """The existing row does not satisfy the expected register contract."""


def _empty_resource(value: Any, field: str) -> None:
    # Only the explicitly named resource gets this narrow normalization.
    # bool is an int subclass and must not pass as numeric zero.
    if isinstance(value, bool) or value not in (None, 0, 0.0):
        raise RegisterRecordMismatch(f"{field}: expected only None/0/0.0 empty resource")
    # type check: 0 and 0.0 are explicitly allowed; other zero-like values are not.
    if value is not None and type(value) not in (int, float):
        raise RegisterRecordMismatch(f"{field}: unsupported empty resource type")


def verify_existing_register_record(
    transport: Any,
    *,
    recorder_key: str,
    recorder_type: str,
    record_type: str,
    expected: Mapping[str, Any],
    empty_resources: Iterable[str] = (),
) -> Mapping[str, Any]:
    """Read one record by exact recorder key/type and verify the gate.

    ``transport.get_by_recorder(recorder_key, recorder_type)`` must return a
    mapping or ``None``. ``transport.post`` is never called. ``empty_resources``
    is a deliberately small allowlist: those fields must be present and may be
    ``None``, ``int(0)`` or ``float(0.0)``; ``False`` is not accepted.
    """
    row = transport.get_by_recorder(recorder_key, recorder_type)
    if row is None:
        raise RegisterRecordUnknown("exact recorder GET returned no record")
    if not isinstance(row, Mapping):
        raise RegisterRecordMismatch("register record is not a mapping")
    if row.get("Recorder") != recorder_key:
        raise RegisterRecordMismatch("Recorder does not match the requested key")
    if row.get("Recorder_Type") != recorder_type:
        raise RegisterRecordMismatch("Recorder_Type does not match the requested type")
    if row.get("_RecordType") != record_type:
        raise RegisterRecordMismatch("_RecordType is incomplete or different")
    if not has_fields(row, expected):
        raise RegisterRecordMismatch("expected register fields differ")
    for field in empty_resources:
        if field not in row:
            raise RegisterRecordMismatch(f"{field}: empty resource field is absent")
        _empty_resource(row[field], field)
    return row
