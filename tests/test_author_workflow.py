"""Synthetic integration. No Word process, Zotero desktop client or network access."""
from contextlib import redirect_stdout
from io import BytesIO, StringIO
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
from urllib.parse import urlsplit
from xml.etree import ElementTree as ET
from zipfile import ZipFile

SCRIPTS = Path(__file__).resolve().parents[1] / "skills/zotero-word-workflow/scripts"
sys.path.insert(0, str(SCRIPTS))
import mark_docx
import zotero_collection


class AuthorWorkflow(unittest.TestCase):
    def test_filing_then_marking_original_stops_before_word_confirmation(self):
        root_key, collection_key, item_key = "ABCDEFGH", "BCDEFGHJ", "CDEFGHJK"
        paper = {"id": "R1", "title": "Synthetic example for integration",
                 "creators": [{"creatorType": "author", "lastName": "Example"}],
                 "date": "2023", "DOI": "10.5555/synthetic-test-integration"}
        metadata = {k: v for k, v in paper.items() if k != "id"}
        metadata.update(itemType="journalArticle", collections=[root_key])
        version = 4
        writes = []

        def record(key, data):
            return {"key": key, "version": version, "library": {"type": "user", "id": 101},
                    "data": dict(data, key=key, version=version)}

        def offline_open(opener, req, timeout=None):
            nonlocal version
            url = urlsplit(req.full_url)
            self.assertEqual(url.hostname, "api.zotero.org")
            prefix = "/users/101/"
            self.assertTrue(url.path.startswith(prefix))
            path = url.path[len(prefix):]
            method = req.get_method()
            if method == "GET" and path == "collections":
                data = [record(root_key, {"name": "Project", "parentCollection": False}),
                        record(collection_key, {"name": "References", "parentCollection": root_key})]
            elif method == "GET" and path == "items/top":
                data = [record(item_key, metadata)]
            elif method == "GET" and path == "items/" + item_key:
                data = record(item_key, metadata)
            elif method == "PATCH" and path == "items/" + item_key:
                changes = json.loads(req.data)
                self.assertEqual(changes, {"collections": [root_key, collection_key]})
                self.assertEqual(dict((k.lower(), v) for k, v in req.header_items())["if-unmodified-since-version"], str(version))
                metadata.update(changes)
                writes.append(path)
                version += 1
                data = None
            else:
                self.fail(f"Unexpected request: {method} {path}")
            response = BytesIO(json.dumps(data).encode() if data is not None else b"")
            response.headers = {"Last-Modified-Version": str(version), "Total-Results": str(len(data) if isinstance(data,list) else 1)}
            return response

        with tempfile.TemporaryDirectory() as directory:
            work = Path(directory)
            papers, mapping, plan = (work / name for name in ("papers.json", "filed.json", "insertions.json"))
            papers.write_text(json.dumps([paper]), encoding="utf8")
            plan.write_text(json.dumps({"insertions":[{"anchor":"The synthetic result","position":"after","references":["R1"]}]}),encoding="utf8")
            doc = work / "original.docx"
            w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            original = f'<w:document xmlns:w="{w}"><w:body><w:p><w:r><w:t>The synthetic result.</w:t></w:r></w:p></w:body></w:document>'
            with ZipFile(doc, "w") as archive:
                archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
                archive.writestr("word/document.xml", original)
            with patch.dict(os.environ, {"ZOTERO_API_KEY":"synthetic-only"}), patch("urllib.request.OpenerDirector.open", new=offline_open), redirect_stdout(StringIO()):
                result = zotero_collection.main(["--library","users/101","--collection","Project/References",
                                                  "--papers",str(papers),"--apply","--output",str(mapping)])
                self.assertEqual(result, 0)
                mark_docx.main([str(doc),"--mapping",str(mapping),"--plan",str(plan),"--in-place"])
            filed = json.loads(mapping.read_text(encoding="utf8"))
            self.assertTrue(filed["applied"])
            self.assertEqual((filed["collection"]["key"], filed["references"][0]["key"]), (collection_key, item_key))
            self.assertEqual(writes, ["items/"+item_key])
            with ZipFile(doc) as archive:
                saved = ET.fromstring(archive.read("word/document.xml"))
            self.assertEqual("".join(n.text or "" for n in saved.iter("{"+w+"}t")), "The synthetic result"+filed["references"][0]["marker"]+".")
            for tag in ("fldChar", "instrText", "fldSimple"):
                self.assertEqual(list(saved.iter("{"+w+"}"+tag)), [])
            self.assertEqual(list(work.glob("*.docx")), [doc])


if __name__ == "__main__":
    unittest.main()
