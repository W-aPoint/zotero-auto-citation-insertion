"""Synthetic offline tests for Zotero collection filing."""

from contextlib import redirect_stderr, redirect_stdout
from copy import deepcopy
import importlib.util
import io
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from urllib import parse, request


SCRIPT = Path(__file__).resolve().parents[1] / "skills/zotero-word-workflow/scripts/zotero_collection.py"
SPEC = importlib.util.spec_from_file_location("cloud_filing_under_test", SCRIPT)
z = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(z)
ROOT, TARGET, OTHER, ALSO = "ROOT2222", "CHILD222", "OTHER222", "ALSO2222"
A, B, C = "AAAA2222", "BBBB2222", "CCCC2222"


def paper(i="ref", **fields):
    return {"id": i, "title": "Synthetic reference " + i, "date": "2023-06-12",
            "creators": [{"creatorType": "author", "lastName": "Sample", "firstName": "A"}], **fields}


class FakeAPI:
    """In-memory synthetic library, with one result per page and versioned writes."""

    def __init__(self, library="groups/456"):
        self.library, self.version = library, 10
        self.collections, self.items, self.calls = {}, {}, []
        self.before_write = self.after_write = None
        self.collection(ROOT, "Root", False)
        self.collection(TARGET, "Child", ROOT)

    def envelope(self, object_key, data):
        scope, number = self.library.split("/")
        return {"key": object_key, "version": self.version, "data": dict(data, key=object_key, version=self.version),
                "library": {"type": "user" if scope == "users" else "group", "id": int(number)}}

    def collection(self, object_key, name, parent):
        self.collections[object_key] = self.envelope(object_key, {"name": name, "parentCollection": parent})

    def store(self, object_key, record, membership=(), **overrides):
        self.version += 1
        data = {k: deepcopy(v) for k, v in record.items() if k not in ("id", "itemKey")}
        data.setdefault("itemType", "journalArticle")
        data.setdefault("collections", list(membership))
        data.update(overrides)
        self.items[object_key] = self.envelope(object_key, data)

    def call(self, method, path, body=None, version=None):
        route, _, query = path.partition("?")
        self.calls.append((method, path, deepcopy(body), version))
        if method == "GET":
            if route in ("items/top", "collections"):
                start = int(parse.parse_qs(query)["start"][0])
                values = list((self.items if route == "items/top" else self.collections).values())
                return deepcopy(values[start:start + 1]), {"Total-Results": str(len(values)), "Last-Modified-Version": str(self.version)}
            obj = self.items.get(route.split("/")[-1])
            if not obj:
                raise z.RemoteProblem(404, "missing")
            return deepcopy(obj), {}
        if self.before_write:
            callback, self.before_write = self.before_write, None
            callback(self, method, body)
        if method == "POST" and route == "items":
            if version != self.version:
                raise z.RemoteProblem(412, "library changed")
            object_key = next(k for k in (A, B, C) if k not in self.items)
            self.store(object_key, body[0])
            result = {"successful": {"0": deepcopy(self.items[object_key])}, "failed": {}}
        elif method == "PATCH" and route.startswith("items/"):
            object_key = route.split("/")[-1]
            obj = self.items[object_key]
            if version != obj["version"]:
                raise z.RemoteProblem(412, "item changed")
            assert set(body) == {"collections"}
            self.store(object_key, obj["data"], **body)
            result = None
        else:
            raise AssertionError((method, route))
        if self.after_write:
            callback, self.after_write = self.after_write, None
            callback(self, object_key)
        return result, {"Last-Modified-Version": str(self.version)}


def run(api, records, apply=True, path="Root/Child", selected_key=None, output=None):
    return z.execute(api, z.prepare(records), path, selected_key, apply, output or z.Output(None, None))


class CollectionTests(unittest.TestCase):
    def test_preview_paginates_identifiers_and_never_issues_markers(self):
        api = FakeAPI()
        api.store(A, paper("unrelated"))
        api.store(B, paper("legacy", extra="PMID: 4321\nDOI: 10.5555/Example"))
        out = run(api, [paper(DOI="https://doi.org/10.5555/example", PMID=4321), paper("new")], False)
        self.assertTrue(out["complete"])
        self.assertFalse(out["applied"])
        self.assertEqual(out["collection"], {"key": TARGET, "path": "Root/Child"})
        old, new = out["references"]
        self.assertEqual((old["action"], old["candidateKey"], new["action"]), ("add-to-collection", B, "create"))
        for ref in out["references"]:
            self.assertFalse(any(ref[f] for f in ("key", "uri", "marker", "verified")))
        self.assertTrue(all(call[0] == "GET" for call in api.calls))

    def test_group_existing_item_memberships_and_user_creation(self):
        api = FakeAPI()
        existing = paper(PMID="4321")
        api.store(A, paper("remote", url="https://pubmed.ncbi.nlm.nih.gov/4321/"), [OTHER])
        out = run(api, [existing], path=None, selected_key=TARGET)
        self.assertTrue(out["applied"])
        self.assertEqual(api.items[A]["data"]["collections"], [OTHER, TARGET])
        item = out["references"][0]
        self.assertEqual((item["uri"], item["marker"]), ("zg:456:" + A, f"{{ | Sample, (2023) | | |zg:456:{A}}}"))
        self.assertEqual([c[0] for c in api.calls if c[1] == f"items/{A}"], ["GET", "PATCH", "GET"])
        users = FakeAPI("users/123")
        record = paper(DOI="10.5555/SYNTH", PMID="77", publicationTitle="Synthetic journal", volume="5")
        first, second = run(users, [record]), run(users, [record])
        self.assertEqual(first["references"][0]["uri"], "zu:123:" + A)
        self.assertEqual(second["references"][0]["action"], "already-member")
        self.assertEqual(len(users.items), 1)
        self.assertEqual(users.items[A]["data"]["publicationTitle"], "Synthetic journal")
        self.assertEqual(users.items[A]["data"]["PMID"], "77")
        self.assertEqual([c[0] for c in users.calls if c[0] != "GET"], ["POST"])

    def test_concurrent_patch_preserves_new_membership(self):
        api = FakeAPI()
        p = paper(DOI="10.5555/race")
        api.store(A, p, [OTHER])
        api.before_write = lambda fake, method, body: fake.store(A, p, [OTHER, ALSO])
        out = run(api, [p])
        self.assertTrue(out["applied"])
        self.assertEqual([c[0] for c in api.calls].count("PATCH"), 2)
        self.assertEqual(api.items[A]["data"]["collections"], [OTHER, ALSO, TARGET])

    def test_concurrent_create_is_reused_on_412(self):
        api = FakeAPI()
        p = paper(PMID="9876")
        api.before_write = lambda fake, method, body: fake.store(B, p, [OTHER])
        out = run(api, [p])
        self.assertTrue(out["applied"])
        self.assertEqual(set(api.items), {B})
        self.assertEqual(out["references"][0]["key"], B)
        self.assertEqual(api.items[B]["data"]["collections"], [OTHER, TARGET])

    def test_changed_library_during_pagination_restarts_with_fresh_inventory(self):
        api = FakeAPI()
        api.store(A, paper("first"))
        api.store(B, paper("second"))
        call = api.call

        def moving(method, path, *args):
            if path == "items/top?limit=100&start=1" and C not in api.items:
                api.store(C, paper("late", PMID="99"), [TARGET])
            return call(method, path, *args)

        api.call = moving
        report = run(api, [{"id": "c", "PMID": "99"}])
        self.assertTrue(report["applied"])
        self.assertEqual(report["references"][0]["key"], C)
        self.assertEqual([r[1] for r in api.calls].count("items/top?limit=100&start=0"), 2)
        self.assertTrue(all(c[0] == "GET" for c in api.calls))

    def test_lost_post_response_partial_checkpoint_and_rerun(self):
        api = FakeAPI()
        a, b, c = [paper(x) for x in ("one", "two", "three")]
        api.store(A, a, [TARGET])

        def lost(fake, object_key):
            raise z.RemoteProblem(0, "synthetic timeout after write")

        api.after_write = lost
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "mapping.json"
            source = Path(tmp) / "papers.json"
            source.write_text(json.dumps([a, b, c]), encoding="utf-8")
            report = run(api, [a, b, c], output=z.Output(output, source))
            self.assertEqual(json.loads(output.read_text(encoding="utf-8")), report)
            self.assertFalse(report["applied"])
            self.assertEqual([r["id"] for r in report["references"]], ["one"])
            self.assertEqual(report["pending"], ["three"])
            retried = run(api, [a, b, c], output=z.Output(output, source))
        self.assertTrue(retried["applied"])
        self.assertEqual(len(api.items), 3)
        self.assertEqual(retried["references"][1]["action"], "already-member")

    def test_membership_is_read_back_before_marker(self):
        api = FakeAPI()
        api.after_write = lambda fake, k: fake.items[k]["data"].update(collections=[])
        report = run(api, [paper()])
        self.assertFalse(report["applied"])
        self.assertEqual(report["references"], [])
        self.assertIn("memberships", report["errors"][0]["message"])

    def test_ambiguous_identity_conflicts_and_explicit_resolution(self):
        api = FakeAPI()
        p = paper(DOI="10.5555/one", PMID="41")
        api.store(A, p, [TARGET])
        api.store(B, p, [TARGET])
        self.assertFalse(run(api, [p])["applied"])
        self.assertTrue(run(api, [dict(p, itemKey=B)])["applied"])
        self.assertFalse(run(api, [dict(p, itemKey=B, PMID="42")])["applied"])
        self.assertFalse(run(api, [dict(p, itemKey=C)])["applied"])
        api.items[A]["data"]["itemType"] = "attachment"
        self.assertFalse(run(api, [dict(p, itemKey=A)])["applied"])
        api.items[B]["library"]["id"] = 99
        self.assertFalse(run(api, [dict(p, itemKey=B)])["applied"])
        self.assertTrue(all(call[0] == "GET" for call in api.calls))

    def test_cautious_title_year_author_and_collection_hierarchy(self):
        api = FakeAPI()
        p = paper()
        api.store(A, p, [TARGET])
        self.assertTrue(run(api, [dict(p, title=p["title"].upper())])["applied"])
        self.assertFalse(run(api, [dict(p, date="2024")])["applied"])
        self.assertFalse(run(api, [dict(p, PMID="33")])["applied"])
        self.assertTrue(run(api, [dict(p, PMID="33", itemKey=A)])["applied"])
        api.collection(OTHER, "Child", ROOT)
        with self.assertRaisesRegex(z.Problem, "ambiguous"):
            run(api, [p])
        self.assertTrue(run(api, [p], path=None, selected_key=TARGET)["applied"])
        with self.assertRaisesRegex(z.Problem, "does not exist"):
            run(api, [p], path=None, selected_key=ALSO)

    def test_minimal_verified_identifier_or_key_reuses_library_metadata(self):
        api = FakeAPI()
        api.store(A, paper(DOI="10.5555/example", PMID="4123"), [OTHER])
        requests = [{"id": "doi", "DOI": "10.5555/example"}, {"id": "pmid", "PMID": 4123},
                    {"id": "key", "itemKey": A}]
        out = run(api, requests)
        self.assertTrue(out["applied"], out["errors"])
        self.assertEqual([r["key"] for r in out["references"]], [A] * 3)
        self.assertEqual([r["title"] for r in out["references"]], [paper()["title"]] * 3)
        self.assertFalse(run(api, [{"id": "wrong", "itemKey": A, "title": "Different paper"}])["applied"])
        for apply in (False, True):
            missing = run(api, [{"id": "missing", "DOI": "10.5555/missing"}], apply)
            self.assertFalse(missing["complete"])
            self.assertIn("to create", missing["errors"][0]["message"])
        self.assertEqual([c[0] for c in api.calls if c[0] != "GET"], ["PATCH"])

    def test_explicit_key_cannot_override_split_doi_and_pmid(self):
        api = FakeAPI()
        api.store(A, paper("one", DOI="10.5555/one"))
        api.store(B, paper("two", PMID="66"))
        report = run(api, [{"id": "conflict", "itemKey": A, "DOI": "10.5555/one", "PMID": "66"}])
        self.assertFalse(report["applied"])
        self.assertIn("different item", report["errors"][0]["message"])
        self.assertTrue(all(c[0] == "GET" for c in api.calls))

    def test_existing_verified_mapping_survives_a_preview_or_wrong_collection(self):
        api = FakeAPI()
        with tempfile.TemporaryDirectory() as tmp:
            src, dst = Path(tmp) / "papers.json", Path(tmp) / "map.json"
            src.write_text(json.dumps([paper()]), encoding="utf-8")
            with patch.object(z, "API", return_value=api), redirect_stdout(io.StringIO()):
                args = ["--library", api.library, "--papers", str(src), "--output", str(dst), "--collection", "Root/Child"]
                self.assertEqual(z.main(args), 0)
                self.assertEqual(z.main(args + ["--apply"]), 0)
                saved = dst.read_bytes()
                with redirect_stderr(io.StringIO()):
                    self.assertEqual(z.main(args), 1)
                    self.assertEqual(z.main(args[:-1] + ["Root", "--apply"]), 1)
                self.assertEqual(dst.read_bytes(), saved)

    def test_input_scope_and_output_guard(self):
        for invalid in ([paper(), paper()], [paper(DOI="made-up")], [paper(PMID="12", extra="PMID: 34")],
                        [paper(creators=[])], [paper(date="undated")], [paper(collections=[TARGET])]):
            with self.subTest(invalid=invalid), self.assertRaises(z.Problem):
                z.prepare(invalid)
        for invalid_scope in ("users/0", "https://outside.invalid", "groups/3/../users/7"):
            with self.assertRaises(z.Problem):
                z.API(invalid_scope, "synthetic")
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "papers.json"
            src.write_text("[]", encoding="utf-8")
            with self.assertRaisesRegex(z.Problem, "different from"):
                z.Output(src, src)
            output = Path(tmp) / "existing.json"
            output.write_text("{}", encoding="utf-8")
            with patch.object(Path, "samefile", return_value=True), self.assertRaisesRegex(z.Problem, "different from"):
                z.Output(output, src)
            with self.assertRaisesRegex(z.Problem, "not a Zotero mapping"):
                z.Output(output, src)

    def test_cli_without_credentials_and_redirect_do_not_echo_a_key(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(os.environ, {}, clear=True):
            src = Path(tmp) / "papers.json"
            src.write_text(json.dumps([paper()]), encoding="utf-8")
            stderr = io.StringIO()
            with redirect_stderr(stderr), redirect_stdout(io.StringIO()):
                self.assertEqual(z.main(["--library", "users/123", "--collection", "Root", "--papers", str(src)]), 1)
            self.assertIn("ZOTERO_API_KEY", stderr.getvalue())
        req = request.Request("https://api.zotero.org/", headers={"Zotero-API-Key": "synthetic-only"})
        with self.assertRaises(z.RemoteProblem) as raised:
            z.NoRedirect().redirect_request(req, None, 302, "", {}, "https://outside.invalid/")
        self.assertNotIn("synthetic-only", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
