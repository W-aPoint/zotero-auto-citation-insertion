"""Synthetic OOXML fixtures; no Word, Zotero or live library is modified."""
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from validate_scan_docx import audit, canonical_uri, digest, expand_numbers, main, W

# Deliberately synthetic identities, not real library evidence or production markers.
A, B, C = 'zu:101:ABCDEFGH', 'zu:101:BCDEFGHJ', 'zg:202:ABCDEFGH'


def marker(uri, locator=''):
    return '{ | Synthetic Author, (2025) |'+locator+' | |'+uri+'}'


def para(*runs):
    p = ET.Element(W+'p')
    for text in runs:
        r = ET.SubElement(p, W+'r')
        ET.SubElement(r, W+'t').text = text
    return p


def docx(path, paragraphs, tables=(), extra=None):
    root = ET.Element(W+'document')
    body = ET.SubElement(root, W+'body')
    body.extend(paragraphs)
    for rows in tables:
        tb = ET.SubElement(body, W+'tbl')
        for cells in rows:
            tr = ET.SubElement(tb, W+'tr')
            for p in cells:
                ET.SubElement(tr, W+'tc').append(p)
    with ZipFile(path, 'w') as z:
        z.writestr('word/document.xml', ET.tostring(root))
        for name, node in (extra or {}).items():
            z.writestr(name, ET.tostring(node))


class AuditTests(unittest.TestCase):
    def setUp(self):
        # Explicit writable parent avoids sandbox-specific default temp cleanup issues.
        self.tmp = TemporaryDirectory(prefix='scan-tests-', dir=Path.cwd())
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.input = self.root/'draft.docx'
        self.map = self.root/'map.json'
        self.map.write_text(json.dumps({'references':[
            {'number':1, 'uri':A}, {'number':2, 'uri':B}, {'number':3, 'uri':C}]}), encoding='utf-8')

    def check(self, paragraphs, tables=(), extra=None, **kw):
        docx(self.input, paragraphs, tables, extra)
        before = digest(self.input), digest(self.map)
        result = audit(self.input, self.map, **kw)
        self.assertEqual(before, (digest(self.input), digest(self.map)))
        return result

    def test_single_marker_and_locator(self):
        r = self.check([para('正文', marker(A, 'pp. 10-14'), '。')])
        self.assertTrue(r['ok'], r['errors'])
        self.assertEqual((r['citation_groups'], r['marker_occurrences'], r['unique_marker_items']), (1,1,1))
        self.assertEqual(r['markers'][0]['locator'], 'pp. 10-14')

    def test_adjacent_group_repetition_and_separate_libraries(self):
        r = self.check([para(marker(A)+marker(B)), para(marker(A)), para(marker(C))])
        self.assertTrue(r['ok'])
        self.assertEqual((r['citation_groups'], r['marker_occurrences'], r['unique_marker_items']), (3,4,3))
        self.assertEqual(r['identity_counts'][A], 2)

    def test_cross_run_marker_and_table_citation(self):
        t = marker(B)
        r = self.check([para(marker(A)[:7], marker(A)[7:])], [[[para(t[:10],t[10:])]]])
        self.assertTrue(r['ok'], r['errors'])
        self.assertEqual((r['marker_occurrences'], r['table_marker_occurrences'], r['table_citation_groups']), (2,1,1))

    def test_numeric_ranges(self):
        self.assertEqual(expand_numbers('[1,3–5，8]'), [1,3,4,5,8])
        for bad in ['[5-3]', '[1,1]', '[0]', '[1-3-5]', '[1;3]']:
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                expand_numbers(bad)

    def test_unknown_item_is_reported(self):
        r = self.check([para(marker('zu:101:ZZZZZZZZ'))])
        self.assertFalse(r['ok'])
        self.assertEqual(r['errors'][0]['paragraph'], 1)
        self.assertIn('missing', r['errors'][0]['error'])

    def test_bad_delimiters_and_fake_identity(self):
        for text in [marker(A)[:-1], marker(A).replace('|','｜'),
                     '{ | Author | |zu:101:ABCDEFGH}',
                     marker('10.1234/fake'), '{ | {Nested} | | |'+A+'}']:
            with self.subTest(text=text):
                r = self.check([para(text)])
                self.assertFalse(r['ok'])
                self.assertTrue(r['errors'])

    def test_ambiguous_number_or_conflicting_key(self):
        for data in [{'references':[{'number':1,'uri':A},{'number':1,'uri':B}]},
                     {'references':[{'uri':A,'key':'ZZZZZZZZ'}]}]:
            self.map.write_text(json.dumps(data), encoding='utf-8')
            r = self.check([para(marker(A))])
            self.assertFalse(r['ok'])

    def test_legacy_marker_mapping(self):
        self.map.write_text(json.dumps({'references':{'legacy':{'number':1,'key':'ABCDEFGH','libraryID':99,'marker':marker(A)}}}),encoding='utf-8')
        r = self.check([para(marker(A))])
        self.assertTrue(r['ok'])
        self.assertEqual(list(r['identity_counts']), [A])

    def test_active_field_and_marker_are_counted_separately(self):
        p = ET.Element(W+'p')
        code = ' ADDIN ZOTERO_ITEM CSL_CITATION '+json.dumps({'citationID':'synthetic','citationItems':[{'uris':['http://zotero.org/users/101/items/ABCDEFGH']}]})
        for typ in ['begin', 'separate', 'end']:
            r = ET.SubElement(p, W+'r')
            ET.SubElement(r, W+'fldChar').set(W+'fldCharType',typ)
            if typ == 'begin':
                for chunk in [code[:30],code[30:]]:
                    ET.SubElement(ET.SubElement(p,W+'r'),W+'instrText').text=chunk
            if typ == 'separate':
                ET.SubElement(ET.SubElement(p,W+'r'),W+'t').text='[1]'
        r = self.check([p,para(marker(B))])
        self.assertTrue(r['ok'],r['errors'])
        self.assertEqual(r['active_citation_groups'],1)
        self.assertEqual(r['marker_occurrences'],1)
        self.assertEqual(r['all_unique_item_identities'],[A,B])
        self.assertTrue(any('Mixed' in str(w) for w in r['warnings']))

    def test_footnote_and_textbox_no_double_count(self):
        notes = ET.Element(W+'footnotes');ET.SubElement(notes,W+'footnote').append(para(marker(C)))
        p = ET.Element(W+'p');box = ET.SubElement(ET.SubElement(p,W+'r'), W+'txbxContent');box.append(para(marker(A)))
        r = self.check([p],extra={'word/footnotes.xml':notes})
        self.assertTrue(r['ok'])
        self.assertEqual(r['marker_occurrences'],2)
        self.assertTrue(any(m['textbox'] for m in r['markers']))
        self.assertTrue(r['warnings'])

    def test_empty_draft_and_math_braces(self):
        r = self.check([para('Math {x+y}, interval [1,2].')])
        self.assertFalse(r['ok'])
        r = self.check([para('Math {x+y}, interval [1,2].')],allow_no_markers=True)
        self.assertTrue(r['ok'])

    def test_deleted_marker_not_counted(self):
        p = para(marker(A))
        ET.SubElement(ET.SubElement(ET.SubElement(p,W+'del'),W+'r'),W+'delText').text=marker(B)
        r = self.check([p])
        self.assertTrue(r['ok'])
        self.assertEqual(r['marker_occurrences'],1)
        self.assertTrue(any('Tracked' in str(w) for w in r['warnings']))

    def test_report_never_overwrites_input_or_existing_report(self):
        self.check([para(marker(A))])
        before=digest(self.input)
        with redirect_stderr(StringIO()), self.assertRaises(SystemExit):
            main([str(self.input),'--mapping',str(self.map),'--report',str(self.input)])
        self.assertEqual(before,digest(self.input))
        report=self.root/'qa.json'
        with redirect_stdout(StringIO()):
            self.assertEqual(main([str(self.input),'--mapping',str(self.map),'--report',str(report)]),0)
        with redirect_stderr(StringIO()):
            self.assertEqual(main([str(self.input),'--mapping',str(self.map),'--report',str(report)]),2)
        self.assertTrue(json.loads(report.read_text(encoding='utf-8'))['ok'])

    def test_uri_conversion_rejects_non_zotero_and_keeps_library(self):
        self.assertEqual(canonical_uri('https://zotero.org/groups/202/items/ABCDEFGH'),C)
        self.assertEqual(canonical_uri('http://zotero.org/users/local/HASH/items/ABCDEFGH'),'zu:0:ABCDEFGH')
        with self.assertRaises(ValueError):
            canonical_uri('https://example.com/users/101/items/ABCDEFGH')


if __name__ == '__main__':
    unittest.main(verbosity=2)
