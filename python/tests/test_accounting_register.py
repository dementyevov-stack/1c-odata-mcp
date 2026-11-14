import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from accounting_register import (  # noqa: E402
    RegisterRecordMismatch,
    RegisterRecordUnknown,
    verify_existing_register_record,
)

RECORDER = "11111111-1111-4111-8111-111111111111"
RECORDER_TYPE = "StandardODATA.Document_Example"
RECORD_TYPE = "StandardODATA.AccountingRegister_Example_RowType"


class FakeRegisterTransport:
    """Isolated fake: exact recorder GET plus counters; no network."""

    def __init__(self, row):
        self.row = row
        self.gets = []
        self.posts = 0
        self.list_calls = 0

    def get_by_recorder(self, key, type_name):
        self.gets.append((key, type_name))
        return self.row

    def list_register(self):
        self.list_calls += 1
        return []  # entity list is empty; it must not be the lookup source

    def post(self, *args, **kwargs):
        self.posts += 1
        raise AssertionError("POST is forbidden for an existing record")


def row(**overrides):
    base = {
        "Recorder": RECORDER,
        "Recorder_Type": RECORDER_TYPE,
        "_RecordType": RECORD_TYPE,
        "Amount": 100,
        "EmptyResource": None,
    }
    base.update(overrides)
    return base


def verify(transport, **overrides):
    options = dict(
        recorder_key=RECORDER,
        recorder_type=RECORDER_TYPE,
        record_type=RECORD_TYPE,
        expected={"Recorder": RECORDER, "Recorder_Type": RECORDER_TYPE, "_RecordType": RECORD_TYPE, "Amount": 100},
        empty_resources=("EmptyResource",),
    )
    options.update(overrides)
    return verify_existing_register_record(transport, **options)


class ExistingRegisterGateTests(unittest.TestCase):
    def test_exact_get_uses_existing_row_when_entity_list_is_empty(self):
        transport = FakeRegisterTransport(row())
        result = verify(transport)
        self.assertEqual(result["Recorder"], RECORDER)
        self.assertEqual(transport.gets, [(RECORDER, RECORDER_TYPE)])
        self.assertEqual(transport.list_calls, 0)
        self.assertEqual(transport.posts, 0)

    def test_none_int_zero_float_zero_are_allowed_for_named_empty_resource(self):
        for value in (None, 0, 0.0):
            result = verify(FakeRegisterTransport(row(EmptyResource=value)))
            self.assertEqual(result["EmptyResource"], value)

    def test_false_is_not_numeric_zero(self):
        with self.assertRaises(RegisterRecordMismatch):
            verify(FakeRegisterTransport(row(EmptyResource=False)))

    def test_missing_resource_property_is_mismatch(self):
        missing = row()
        del missing["EmptyResource"]
        with self.assertRaises(RegisterRecordMismatch):
            verify(FakeRegisterTransport(missing))

    def test_wrong_recorder_type_is_mismatch(self):
        with self.assertRaises(RegisterRecordMismatch):
            verify(FakeRegisterTransport(row(Recorder_Type="StandardODATA.Document_Other")))

    def test_incomplete_record_type_is_mismatch(self):
        with self.assertRaises(RegisterRecordMismatch):
            verify(FakeRegisterTransport(row(_RecordType="AccountingRegister_Example")))

    def test_absent_row_is_unknown_and_does_not_create(self):
        transport = FakeRegisterTransport(None)
        with self.assertRaises(RegisterRecordUnknown):
            verify(transport)
        self.assertEqual(transport.posts, 0)

    def test_mismatch_does_not_post(self):
        transport = FakeRegisterTransport(row(Amount=101))
        with self.assertRaises(RegisterRecordMismatch):
            verify(transport)
        self.assertEqual(transport.posts, 0)


if __name__ == "__main__":
    unittest.main()
