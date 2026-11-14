import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from catalog_verifier import (  # noqa: E402
    CatalogCyclePlan,
    ReadUnavailableError,
    ResponseLostError,
    SearchResult,
    changed_fields,
    execute_catalog_cycle,
    has_fields,
)


class HasFieldsTests(unittest.TestCase):
    def test_missing_null_field_is_not_present(self):
        self.assertFalse(has_fields({}, {"Optional": None}))
        self.assertTrue(has_fields({"Optional": None}, {"Optional": None}))

    def test_bool_is_not_int(self):
        self.assertFalse(has_fields({"Flag": 1}, {"Flag": True}))
        self.assertFalse(has_fields({"Flag": True}, {"Flag": 1}))
        self.assertFalse(has_fields({"Zero": 0}, {"Zero": False}))

    def test_ref_key_must_match(self):
        self.assertFalse(has_fields({"Ref_Key": "aaaaaaaa-..."}, {"Ref_Key": "bbbbbbbb-..."}))
        self.assertTrue(has_fields({"Ref_Key": "A"}, {"Ref_Key": "A"}))

    def test_explicit_normalizer_is_field_specific(self):
        # Edm.Int64 may be serialized as a decimal string; normalise only it.
        self.assertTrue(
            has_fields({"DataVersion": "123"}, {"DataVersion": 123}, {"DataVersion": str})
        )
        self.assertFalse(has_fields({"Other": "123"}, {"Other": 123}))


class ChangedFieldsTests(unittest.TestCase):
    def test_missing_null_field_is_a_change(self):
        self.assertEqual(changed_fields({"Optional": None}, {}), {"Optional"})
        self.assertFalse(changed_fields({"Optional": None}, {"Optional": None}))

    def test_added_or_changed_unrelated_field_is_visible(self):
        self.assertEqual(changed_fields({"Ref_Key": "A"}, {"Ref_Key": "A", "Extra": 1}), {"Extra"})
        self.assertEqual(changed_fields({"Amount": 1}, {"Amount": 2}), {"Amount"})

    def test_bool_against_number_is_a_change(self):
        self.assertEqual(changed_fields({"Flag": 1}, {"Flag": True}), {"Flag"})

    def test_only_declared_patch_and_technical_fields(self):
        before = {
            "Ref_Key": "R",
            "DisplayOrder": 1,
            "DataVersion": "1",
            "Description": "same",
            "Deleted": False,
        }
        after = {
            "Ref_Key": "R",
            "DisplayOrder": 2,
            "DataVersion": "2",
            "Description": "same",
            "Deleted": False,
        }
        self.assertEqual(
            changed_fields(before, after, ignore_fields={"DisplayOrder", "DataVersion"}),
            set(),
        )
        self.assertEqual(changed_fields(before, after, ignore_fields={"DataVersion"}), {"DisplayOrder"})


class FakeCatalogAdapter:
    """Fake HTTP-ish adapter; no network, only synthetic rows."""

    def __init__(
        self,
        *,
        search_rows=(),
        search_complete=True,
        get_fail_at=(),
        patch_lost=False,
        get_override=None,
        patch_extra=None,
        remove_after_patch=(),
    ):
        self.search_rows = [dict(row) for row in search_rows]
        self.search_complete = search_complete
        self.get_fail_at = set(get_fail_at)
        self.patch_lost = patch_lost
        self.get_override = get_override
        self.patch_extra = patch_extra or {}
        self.remove_after_patch = set(remove_after_patch)
        self.state = {
            row["Ref_Key"]: dict(row) for row in self.search_rows if isinstance(row.get("Ref_Key"), str)
        }
        self.posts = 0
        self.patches = 0
        self.gets = 0
        self.next_ref = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"

    def find_exact(self, source_key):
        return SearchResult(rows=[dict(row) for row in self.search_rows], complete=self.search_complete)

    def post(self, payload):
        self.posts += 1
        row = dict(payload)
        row["Ref_Key"] = self.next_ref
        self.state[self.next_ref] = row
        return dict(row)

    def get(self, ref):
        self.gets += 1
        if self.gets in self.get_fail_at:
            raise ReadUnavailableError("synthetic GET unavailable")
        if self.get_override is not None:
            return dict(self.get_override)
        return dict(self.state[ref])

    def patch(self, ref, changes):
        self.patches += 1
        self.state[ref].update(changes)
        self.state[ref].update(self.patch_extra)
        for key in self.remove_after_patch:
            self.state[ref].pop(key, None)
        if self.patch_lost:
            raise ResponseLostError("synthetic PATCH response lost")


def plan(**overrides):
    base = dict(
        source_key="SRC-1",
        create_payload={"Code": "SRC-1", "Description": "Synthetic", "Optional": None, "DisplayOrder": 1},
        expected_existing={"Code": "SRC-1", "Description": "Synthetic", "Optional": None, "DisplayOrder": 1},
        expected_created={"Code": "SRC-1", "Description": "Synthetic", "Optional": None, "DisplayOrder": 1},
        patch_changes={"DisplayOrder": 2},
        expected_patched={"DisplayOrder": 2},
        ignore_fields=("DataVersion",),
    )
    base.update(overrides)
    return CatalogCyclePlan(**base)


class ExecutableCatalogCycleTests(unittest.TestCase):
    """These tests call the worker cycle, not a copy of its guard logic."""

    def test_patch_applied_but_response_lost_recovers_by_exact_get(self):
        adapter = FakeCatalogAdapter(patch_lost=True)
        result = execute_catalog_cycle(adapter, plan())
        self.assertEqual(result.status, "patched")
        self.assertTrue(result.recovered)
        self.assertEqual(adapter.posts, 1)
        self.assertEqual(adapter.patches, 1)

    def test_control_get_unavailable_after_patch_is_unknown_without_retry(self):
        adapter = FakeCatalogAdapter(get_fail_at={2})
        result = execute_catalog_cycle(adapter, plan())
        self.assertEqual(result.status, "unknown")
        self.assertEqual(result.ref, "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa")
        self.assertEqual(adapter.posts, 1)
        self.assertEqual(adapter.patches, 1)

    def test_foreign_guid_in_readback_is_conflict(self):
        adapter = FakeCatalogAdapter(
            get_override={"Ref_Key": "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb", "Code": "SRC-1"},
        )
        result = execute_catalog_cycle(adapter, plan())
        self.assertEqual(result.status, "conflict")
        self.assertIn("Ref_Key", result.reason)
        self.assertEqual(adapter.posts, 1)

    def test_optional_null_disappears_strictly_after_patch_is_conflict(self):
        adapter = FakeCatalogAdapter(remove_after_patch={"Optional"})
        result = execute_catalog_cycle(adapter, plan())
        self.assertEqual(result.status, "conflict")
        self.assertEqual(adapter.posts, 1)
        self.assertEqual(adapter.patches, 1)

    def test_unrelated_field_change_after_patch_is_conflict(self):
        adapter = FakeCatalogAdapter(patch_extra={"Unrelated": 1})
        result = execute_catalog_cycle(adapter, plan())
        self.assertEqual(result.status, "conflict")
        self.assertIn("Unrelated", result.reason)

    def test_repeat_existing_exact_object_does_not_post(self):
        existing = {
            "Ref_Key": "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
            "Code": "SRC-1",
            "Description": "Synthetic",
            "Optional": None,
            "DisplayOrder": 1,
        }
        adapter = FakeCatalogAdapter(search_rows=[existing])
        result = execute_catalog_cycle(adapter, plan())
        self.assertEqual(result.status, "existing")
        self.assertEqual(result.ref, existing["Ref_Key"])
        self.assertEqual(adapter.posts, 0)

    def test_repeat_after_unknown_patch_detects_old_value_without_post(self):
        old = {
            "Ref_Key": "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
            "Code": "SRC-1",
            "Description": "Synthetic",
            "Optional": None,
            "DisplayOrder": 1,
        }
        adapter = FakeCatalogAdapter(search_rows=[old])
        result = execute_catalog_cycle(
            adapter,
            plan(
                expected_existing={
                    "Code": "SRC-1",
                    "Description": "Synthetic",
                    "Optional": None,
                    "DisplayOrder": 2,
                }
            ),
        )
        self.assertEqual(result.status, "conflict")
        self.assertEqual(adapter.posts, 0)
        self.assertEqual(adapter.patches, 0)
        self.assertIn("DisplayOrder", str(old))

    def test_known_ref_recovery_does_not_need_new_search(self):
        adapter = FakeCatalogAdapter(patch_lost=True)
        result = execute_catalog_cycle(adapter, plan())
        self.assertEqual(result.status, "patched")
        self.assertTrue(result.recovered)
        recovered = adapter.get(result.ref)
        self.assertTrue(has_fields(recovered, {"Ref_Key": result.ref, "DisplayOrder": 2}))
        self.assertEqual(adapter.posts, 1)

    def test_search_hit_without_ref_is_conflict_without_post(self):
        adapter = FakeCatalogAdapter(search_rows=[{"Code": "SRC-1", "Description": "Synthetic"}])
        result = execute_catalog_cycle(adapter, plan())
        self.assertEqual(result.status, "conflict")
        self.assertIn("Ref_Key", result.reason)
        self.assertEqual(adapter.posts, 0)
        self.assertEqual(adapter.gets, 0)

    def test_existing_ref_get_unavailable_is_unknown_without_post(self):
        existing = {
            "Ref_Key": "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee",
            "Code": "SRC-1",
            "Description": "Synthetic",
            "Optional": None,
            "DisplayOrder": 1,
        }
        adapter = FakeCatalogAdapter(search_rows=[existing], get_fail_at={1})
        result = execute_catalog_cycle(adapter, plan())
        self.assertEqual(result.status, "unknown")
        self.assertEqual(result.ref, existing["Ref_Key"])
        self.assertEqual(adapter.posts, 0)

    def test_ambiguous_or_incomplete_search_never_posts(self):
        for search_rows, complete in (
            ([{"Ref_Key": "1", "Code": "SRC-1"}, {"Ref_Key": "2", "Code": "SRC-1"}], True),
            ([{"Ref_Key": "1", "Code": "SRC-1"}], False),
        ):
            adapter = FakeCatalogAdapter(search_rows=search_rows, search_complete=complete)
            result = execute_catalog_cycle(adapter, plan())
            self.assertEqual(result.status, "conflict")
            self.assertEqual(adapter.posts, 0)
            self.assertEqual(adapter.patches, 0)

    def test_patch_changes_only_declared_field_and_technical_version(self):
        adapter = FakeCatalogAdapter(patch_extra={"DataVersion": "2"})
        result = execute_catalog_cycle(adapter, plan())
        self.assertEqual(result.status, "patched")
        self.assertEqual(result.changed, frozenset())


if __name__ == "__main__":
    unittest.main()
