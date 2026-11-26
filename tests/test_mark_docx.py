import copy
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest import mock
import zipfile

from lxml import etree

SCRIPTS = Path(__file__).resolve().parents[1] / "skills" / "zotero-word-workflow" / "scripts"
sys.path.insert(0, str(SCRIPTS))
import mark_docx as mark


NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + NS + "}"
CT = "application/vnd.openxmlformats-officedocument.wordprocessingml."
PKG = "http://schemas.openxmlformats.org/package/2006/"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
R1 = "{ | Lee, (2024) | | |zu:123:ABCD2345}"
R2 = "{ | Huang, (2018) | | |zu:123:WXYZ6789}"
MAPPING = {"applied": True, "collection": {"path": "Review/Studies", "key": "COLL1234"},
           "library": "users/123", "references": [
    {"id": "R1", "key": "ABCD2345", "uri": "zu:123:ABCD2345", "marker": R1},
    {"id": "R2", "key": "WXYZ6789", "uri": "zu:123:WXYZ6789", "marker": R2},
]}


def para(text):
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def request(anchor, side="after", refs=None, **kwargs):
    return dict(anchor=anchor, position=side, references=refs or ["R1"], **kwargs)


def xml(tag, body):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{tag} xmlns:w="{NS}">'
            f'{body}</w:{tag}>').encode("utf-8")


def relationships(items):
    links = [f'<Relationship Id="rId{i}" Type="{REL}{kind}" Target="{target}"/>'
             for i, (kind, target) in enumerate(items, 1)]
    return (f'<Relationships xmlns="{PKG}relationships">' + "".join(links) + '</Relationships>').encode()


CROSS = ('<w:p><w:r><w:t>Test cross</w:t></w:r><w:r><w:rPr><w:b/></w:rPr>'
         '<w:t xml:space="preserve"> run here.</w:t></w:r></w:p>')
TABLE = '<w:tbl><w:tr><w:tc>' + para("表格依据") + '</w:tc></w:tr></w:tbl>'
CITATION = ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText> ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"X"}</w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>[7]</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
BIBL = ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText> ADDIN ZOTERO_BIBL</w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>List begins</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>Existing bibliography.</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
FOOTNOTES = '<w:footnote w:id="1">' + para("附注证据") + '</w:footnote>'
ENDNOTES = '<w:endnote w:id="1">' + para("尾注证据") + '</w:endnote>'


class MarkerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.src = self.folder / "文稿 example.docx"
        self.out = self.folder / "marked.docx"

    def create(self, body=None, extras=None):
        if body is None:
            body = CROSS + TABLE + CITATION + BIBL
        kinds = {"document": "document.main", "styles": "styles", "footnotes": "footnotes", "endnotes": "endnotes"}
        types = [f'<Override PartName="/word/{name}.xml" ContentType="{CT}{kind}+xml"/>' for name, kind in kinds.items()]
        parts = {
            "[Content_Types].xml": (f'<Types xmlns="{PKG}content-types"><Default Extension="rels" '
                                    f'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                                    '<Default Extension="png" ContentType="image/png"/>' + "".join(types) + '</Types>').encode(),
            "_rels/.rels": relationships([("officeDocument", "word/document.xml")]),
            "word/_rels/document.xml.rels": relationships([(n, n + ".xml") for n in kinds if n != "document"]),
            "word/document.xml": xml("document", "<w:body>" + body + "</w:body>"),
            "word/footnotes.xml": xml("footnotes", FOOTNOTES),
            "word/endnotes.xml": xml("endnotes", ENDNOTES),
            "word/styles.xml": xml("styles", ""),
            "word/media/figure.png": bytes(range(256)),
        }
        parts.update(extras or {})
        with zipfile.ZipFile(self.src, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.comment = b"preserved package comment"
            for name, data in parts.items():
                z.writestr(name, data)
        return parts

    def insert(self, entries, mapping=MAPPING, **kwargs):
        return mark.mark_document(self.src, mapping, {"insertions": entries}, **(kwargs or {"in_place": True}))

    def test_in_place_group_cross_runs_table_notes_and_live_fields(self):
        original = self.create()
        entries = [request("cross run", refs=["R1", {"id": "R2", "locator": "p. 45"}]),
                   request("表格依据", "before"), request("附注证据", part="word/footnotes.xml"),
                   request("尾注证据", "before", part="word/endnotes.xml")]
        self.assertEqual(self.insert(entries)["inserted_groups"], 4)
        with zipfile.ZipFile(self.src) as z:
            self.assertEqual(z.namelist(), list(original))
            self.assertEqual(z.comment, b"preserved package comment")
            for name in set(original) - set(mark.PARTS):
                self.assertEqual(z.read(name), original[name])
            doc = etree.fromstring(z.read(mark.PARTS[0]))
            paragraphs = list(doc.iter(W + "p"))
            first = paragraphs[0]
            nodes = list(first.iter(W + "t"))
            self.assertEqual(nodes[0].text, "Test cross")
            self.assertEqual(nodes[1].text, " run" + R1 + R2.replace("| | |", "| p. 45 | |") + " here.")
            self.assertIsNotNone(first.find(".//" + W + "b"))
            self.assertEqual(paragraphs[1].find(".//" + W + "t").text, R1 + "表格依据")
            old = list(etree.fromstring(original[mark.PARTS[0]]).iter(W + "p"))
            self.assertEqual([etree.tostring(p, method="c14n") for p in old[2:]],
                             [etree.tostring(p, method="c14n") for p in paragraphs[2:]])
            self.assertEqual(etree.fromstring(z.read(mark.PARTS[1])).find(".//" + W + "t").text, "附注证据" + R1)
            self.assertEqual(etree.fromstring(z.read(mark.PARTS[2])).find(".//" + W + "t").text, R1 + "尾注证据")
        snapshot = self.src.read_bytes()
        self.assertEqual(self.insert(entries)["already_present"], 4)
        self.assertEqual(self.src.read_bytes(), snapshot)
        self.assertEqual(list(self.folder.iterdir()), [self.src])

    def test_cli_verified_mapping_writes_new_copy_and_refuses_overwrite(self):
        self.create(para("Result."))
        before = self.src.read_bytes()
        mapping = dict(MAPPING, verified=True)
        del mapping["applied"]
        m, p = self.folder / "mapping.json", self.folder / "plan.json"
        m.write_text(json.dumps(mapping), encoding="utf-8-sig")
        p.write_text(json.dumps({"insertions": [request("Result.")]}), encoding="utf-8")
        cmd = [sys.executable, str(SCRIPTS / "mark_docx.py"), str(self.src), "--mapping", str(m), "--plan", str(p),
               "--output", str(self.out)]
        result = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)["inserted_groups"], 1)
        self.assertEqual(self.src.read_bytes(), before)
        with zipfile.ZipFile(self.out) as z:
            doc = etree.fromstring(z.read(mark.PARTS[0]))
            self.assertEqual(doc.find(".//" + W + "t").text, "Result." + R1)
            self.assertIsNone(doc.find(".//" + W + "fldChar"))
        marked = self.out.read_bytes()
        self.assertNotEqual(subprocess.run(cmd, capture_output=True).returncode, 0)
        self.assertEqual(self.out.read_bytes(), marked)

    def test_ambiguous_missing_collision_and_conflicting_group_leave_original(self):
        self.create(para("duplicate") + para("duplicate") + para("numbered" + R1) + para("abcde"))
        before = self.src.read_bytes()
        invalid = [[request("abcde"), request("duplicate")], [request("missing")],
                   [request("numbered", refs=["R2"])], [request("abcde"), request("de")],
                   [request("abcde"), request("bc")], [request("abcde", refs=["R1", "R1"])]]
        for entries in invalid:
            with self.subTest(entries=entries), self.assertRaises(ValueError):
                self.insert(entries)
            self.assertEqual(self.src.read_bytes(), before)

    def test_selecting_protected_content_or_crossing_revision_fails(self):
        extra = ('<w:p><w:r><w:t>left</w:t></w:r><w:del w:id="1"><w:r><w:delText>gone</w:delText>'
                 '</w:r></w:del><w:r><w:t>right</w:t></w:r></w:p>'
                 '<w:p><w:ins w:id="2"><w:r><w:t>revision</w:t></w:r></w:ins></w:p>'
                 '<w:sdt><w:sdtContent>' + para("control") + '</w:sdtContent></w:sdt>'
                 '<w:p><w:bookmarkStart w:id="5" w:name="ZOTERO_BREF_a"/><w:r><w:t>saved ref</w:t></w:r>'
                 '<w:bookmarkEnd w:id="5"/></w:p><w:p><w:r><w:t>one</w:t><w:tab/><w:t>two</w:t></w:r></w:p>'
                 '<w:p><w:pPr><w:pStyle w:val="Bibliography"/></w:pPr><w:r><w:t>styled entry</w:t></w:r></w:p>'
                 '<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r></w:p>')
        self.create(CITATION + BIBL + extra)
        before = self.src.read_bytes()
        for text in ["[7]", "Existing bibliography.", "leftright", "revision", "control", "saved ref", "onetwo",
                     "styled entry", "hidden"]:
            with self.subTest(anchor=text), self.assertRaises(ValueError):
                self.insert([request(text)])
        self.assertEqual(self.src.read_bytes(), before)

    def test_mappings_signatures_protection_and_owner_locks(self):
        self.create(para("text"))
        wrong = copy.deepcopy(MAPPING)
        wrong["references"][0]["key"] = "ZZZZ9999"
        for mapping in [dict(MAPPING, applied=False), dict(MAPPING, applied=False, verified=True), wrong]:
            with self.assertRaises(ValueError):
                self.insert([request("text")], mapping=mapping)
        lock = self.folder / ("~$" + self.src.name[2:])
        lock.touch()
        with self.assertRaisesRegex(ValueError, "lock"):
            self.insert([request("text")])
        lock.unlink()
        with self.src.open("rb") as opened:
            if os.name != "nt":
                import fcntl
                fcntl.flock(opened, fcntl.LOCK_EX)
            with self.assertRaises(OSError):
                self.insert([request("text")])
        for extras in [{"_xmlsignatures/sig1.xml": b"signed"},
                       {"word/settings.xml": xml("settings", '<w:documentProtection w:enforcement="1"/>')},
                       {"word/settings.xml": xml("settings", '<w:trackRevisions/>')}]:
            self.create(para("text"), extras)
            before = self.src.read_bytes()
            with self.assertRaises(ValueError):
                self.insert([request("text")])
            self.assertEqual(self.src.read_bytes(), before)

    def test_verified_collection_and_single_library(self):
        local = copy.deepcopy(MAPPING["references"][0])
        for name in ("uri", "marker"):
            local[name] = local[name].replace("zu:123:", "zu:0:")
        manual = {"verified": True, "references": [local], "collection": {"path": "Local project"}}
        self.assertIn("R1", mark.verified_items(manual))
        self.assertIn("R1", mark.verified_items(dict(manual, library="users/0", collection={"key": "COLL1234"})))
        for invalid in [dict(MAPPING, collection={}), dict(MAPPING, library=None), dict(MAPPING, complete=False),
                        dict(manual, applied=True, library="users/0"),
                        dict(manual, collection={"key": "bad"}), dict(manual, collection={"path": "A//B"}),
                        dict(manual, library="users/123"), dict(manual, references=[local, MAPPING["references"][1]])]:
            with self.assertRaises(ValueError):
                mark.verified_items(invalid)

    def test_source_change_before_commit_discards_stage(self):
        self.create(para("text"))
        before = self.src.read_bytes()
        with mock.patch.object(mark, "fingerprint", side_effect=[b"before", b"after"]):
            with self.assertRaisesRegex(ValueError, "changed"):
                self.insert([request("text")])
        self.assertEqual(self.src.read_bytes(), before)
        self.assertEqual(list(self.folder.iterdir()), [self.src])

    @unittest.skipUnless(os.name == "nt", "Windows NTFS metadata")
    def test_windows_retains_acl_stream_and_restores_renamed_original_on_failure(self):
        self.create(para("text"))
        ads = Path(str(self.src) + ":mark_docx_test")
        ads.write_bytes(b"original named stream")
        subprocess.run(["icacls", str(self.src), "/inheritance:d"], capture_output=True, check=True)
        acl = subprocess.check_output(["icacls", str(self.src)])
        before = self.src.read_bytes()

        def interrupted(target, stage, backup):
            os.rename(target, backup)
            raise OSError("simulated interrupted ReplaceFileW")

        with mock.patch.object(mark, "windows_replace_file", side_effect=interrupted):
            with self.assertRaisesRegex(OSError, "interrupted"):
                self.insert([request("text")])
        self.assertEqual(self.src.read_bytes(), before)
        self.assertEqual(self.insert([request("text")])["inserted_groups"], 1)
        self.assertEqual(subprocess.check_output(["icacls", str(self.src)]), acl)
        self.assertEqual(ads.read_bytes(), b"original named stream")
        self.assertEqual(list(self.folder.iterdir()), [self.src])


if __name__ == "__main__":
    unittest.main()
