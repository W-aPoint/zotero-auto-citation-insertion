"""Read-only Scannable Cite DOCX audit. Python 3.10+, standard library only."""
import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import sys
from urllib.parse import urlparse
import xml.etree.ElementTree as ET
from zipfile import ZipFile, BadZipFile

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
URI = re.compile(r'z[ug]:\d+:[A-Z0-9]{8}\Z')
TOKEN = re.compile(r'\{[^{}]*\}')
STORIES = re.compile(r'word/(document|footnotes|endnotes|header\d+|footer\d+)\.xml\Z')


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def canonical_uri(value):
    value = value.strip()
    if URI.fullmatch(value):
        return value
    p = urlparse(value)
    if p.scheme in ('http', 'https') and p.netloc in ('zotero.org', 'www.zotero.org'):
        m = re.fullmatch(r'/(users|groups)/(\d+)/items/([A-Z0-9]{8})/?', p.path)
        if m:
            return f"{'zu' if m[1] == 'users' else 'zg'}:{m[2]}:{m[3]}"
        m = re.fullmatch(r'/users/local/[^/]+/items/([A-Z0-9]{8})/?', p.path)
        if m:
            return 'zu:0:' + m[1]
    raise ValueError('Expected a real Scannable Cite short identity or Zotero HTTP item URI')


def parse_marker(token):
    if not token.startswith('{') or not token.endswith('}'):
        raise ValueError('Marker must have ASCII opening and closing braces')
    fields = token[1:-1].split('|')
    if len(fields) != 5 or any(c in token[1:-1] for c in '{}'):
        raise ValueError('Marker must contain exactly five fields and four ASCII pipes')
    fields = [f.strip() for f in fields]
    if not URI.fullmatch(fields[4]):
        raise ValueError('Invalid marker identity: use zu:USERID:KEY or zg:GROUPID:KEY')
    return fields


def expand_numbers(token):
    """Expand a known numeric citation, never discover citation intent automatically."""
    if not re.fullmatch(r'\[\s*\d+(?:\s*[-–—,，、]\s*\d+)*\s*\]', token):
        raise ValueError('Invalid numeric citation')
    out = []
    for chunk in re.split('[,，、]', token[1:-1]):
        pair = re.split('[-–—]', chunk.strip())
        if len(pair) == 1:
            out.append(int(pair[0]))
        elif len(pair) == 2 and int(pair[0]) <= int(pair[1]):
            if int(pair[1]) - int(pair[0]) > 10000:
                raise ValueError('Numeric range is implausibly large; inspect the source')
            out.extend(range(int(pair[0]), int(pair[1]) + 1))
        else:
            raise ValueError('Reversed or malformed numeric range')
    if any(n <= 0 for n in out) or len(set(out)) != len(out):
        raise ValueError('Citation numbers must be positive and nonduplicated')
    return out


def read_mapping(path):
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    records = data.get('references', data) if isinstance(data, dict) else data
    if isinstance(records, dict):
        records = list(records.values())
    if not isinstance(records, list) or not records:
        raise ValueError('Mapping must contain a nonempty references list or dictionary')
    identities, numbers = set(), {}
    for i, rec in enumerate(records, 1):
        if not isinstance(rec, dict):
            raise ValueError(f'Mapping record {i} is not an object')
        marker_uri = parse_marker(rec['marker'])[4] if rec.get('marker') else None
        uri = canonical_uri(rec['uri']) if rec.get('uri') else marker_uri
        if uri is None:
            raise ValueError(f'Mapping record {i} lacks marker or uri')
        if marker_uri and marker_uri != uri:
            raise ValueError(f'Mapping record {i}: marker and uri disagree')
        if rec.get('key') and rec['key'] != uri.rsplit(':', 1)[1]:
            raise ValueError(f'Mapping record {i}: item key disagrees with identity')
        identities.add(uri)
        if rec.get('number') is not None:
            num = int(rec['number'])
            if num <= 0 or (num in numbers and numbers[num] != uri):
                raise ValueError(f'Ambiguous or invalid reference number: {num}')
            numbers[num] = uri
    return identities, numbers


def read_story(xml, part):
    root = ET.fromstring(xml)
    paragraphs, fields, errors, tracked = [], [], [], 0
    stack = []

    def walk(node, current=None, table=False, excluded=False):
        nonlocal tracked
        tag = node.tag
        if tag in (W+'ins', W+'del'):
            tracked += 1
        if tag == W+'del':
            return
        if tag == W+'tbl':
            table = True
        if tag == W+'p':
            current = {'part': part, 'paragraph': len(paragraphs)+1,
                       'in_table': table, 'text': '', 'textbox': False}
            paragraphs.append(current)
        if tag == W+'txbxContent' and current is not None:
            # Nested textbox paragraphs are separate records, never duplicate their text.
            excluded = True
        if tag == W+'p' and excluded:
            current['textbox'] = True
        if tag == W+'fldSimple':
            code = node.get(W+'instr', '')
            if 'ZOTERO_ITEM' in code:
                fields.append({'part': part, 'instruction': code})
            return
        if tag == W+'fldChar':
            typ = node.get(W+'fldCharType')
            if typ == 'begin':
                stack.append('')
            elif typ == 'end':
                if not stack:
                    errors.append({'part': part, 'error': 'Field end has no matching begin'})
                else:
                    code = stack.pop()
                    if 'ZOTERO_ITEM' in code:
                        fields.append({'part': part, 'instruction': code})
        elif tag == W+'instrText' and stack:
            stack[-1] += node.text or ''
        elif current is not None and not stack:
            if tag == W+'t':
                current['text'] += node.text or ''
            elif tag in (W+'br', W+'cr'):
                current['text'] += '\n'
            elif tag == W+'tab':
                current['text'] += '\t'
        for child in node:
            walk(child, current, table, excluded)
    walk(root)
    if stack:
        errors.append({'part': part, 'error': 'Unclosed Word fields'})
    return paragraphs, fields, errors, tracked


def audit(docx, mapping, allow_no_markers=False):
    docx, mapping = Path(docx), Path(mapping)
    hashes = (digest(docx), digest(mapping))
    report = {'input': str(docx.resolve()), 'mapping': str(mapping.resolve()),
              'input_sha256': hashes[0], 'mapping_sha256': hashes[1],
              'status': 'pending_scan_conversion', 'errors': [], 'warnings': [],
              'markers': [], 'by_part': {}, 'active_citation_groups': 0}
    try:
        known, numbers = read_mapping(mapping)
    except (ValueError, TypeError, KeyError) as exc:
        report['errors'].append({'error': 'Invalid mapping: '+str(exc)})
        report['ok'] = False
        return report
    paragraphs, fields, tracked = [], [], 0
    with ZipFile(docx) as z:
        if 'word/document.xml' not in z.namelist():
            raise ValueError('Not a Word DOCX package: word/document.xml is missing')
        for part in sorted(n for n in z.namelist() if STORIES.fullmatch(n)):
            ps, fs, es, tr = read_story(z.read(part), part)
            paragraphs += ps
            fields += fs
            tracked += tr
            report['errors'] += es
    active_uris = set()
    for field in fields:
        code = field['instruction']
        try:
            obj = json.loads(code[code.index('{'):].strip())
            for item in obj['citationItems']:
                for uri in item.get('uris', []):
                    try:
                        active_uris.add(canonical_uri(uri))
                    except ValueError:
                        report['warnings'].append('An existing active item has a nonstandard URI; inspect it separately.')
        except (ValueError, KeyError, TypeError):
            report['errors'].append({'part': field['part'], 'error': 'Unparseable existing Zotero citation field'})
    report['active_citation_groups'] = len(fields)
    report['active_unique_item_identities'] = sorted(active_uris)
    groups, table_groups = 0, 0
    for p in paragraphs:
        txt = p['text']
        loc = {k:p[k] for k in ('part', 'paragraph', 'in_table', 'textbox')}
        # Valid brace tokens are processed first. Other text, including math braces, is retained.
        accounted, last_end = [], None
        part_count = report['by_part'].setdefault(p['part'], {'groups':0, 'occurrences':0, 'table_occurrences':0})
        for match in TOKEN.finditer(txt):
            token = match.group()
            if not ('zu:' in token or 'zg:' in token or token.count('|') >= 3 or re.match(r'\{\s*\|', token)):
                continue
            accounted.append((match.start(), match.end()))
            try:
                values = parse_marker(token)
            except ValueError as exc:
                report['errors'].append(loc | {'offset':match.start(), 'error':str(exc), 'token':token})
                continue
            uri = values[4]
            if uri not in known:
                report['errors'].append(loc | {'offset':match.start(), 'error':'Identity missing from supplied mapping', 'uri':uri})
            if last_end != match.start():
                groups += 1
                part_count['groups'] += 1
                table_groups += int(p['in_table'])
                if last_end is not None and txt[last_end:match.start()].isspace():
                    report['warnings'].append(loc | {'warning':'Whitespace separates markers; confirm whether these should form one citation group.'})
            last_end = match.end()
            part_count['occurrences'] += 1
            part_count['table_occurrences'] += int(p['in_table'])
            report['markers'].append(loc | {'offset':match.start(), 'uri':uri,
                                            'label':values[1], 'locator':values[2]})
        residue = list(txt)
        for a,b in accounted:
            residue[a:b] = ' '*(b-a)
        residue = ''.join(residue)
        if re.search(r'z[ug]:|\{[^{}\n]*\|', residue):
            report['errors'].append(loc | {'error':'Dangling, nested or damaged marker delimiters/identity'})
    identities = Counter(m['uri'] for m in report['markers'])
    report.update(citation_groups=groups, marker_occurrences=sum(identities.values()),
                  unique_marker_items=len(identities), table_citation_groups=table_groups,
                  table_marker_occurrences=sum(m['in_table'] for m in report['markers']),
                  identity_counts=dict(sorted(identities.items())),
                  unused_mapping_identities=sorted(known-set(identities)),
                  mapped_reference_numbers=len(numbers),
                  all_unique_item_identities=sorted(set(identities)|active_uris))
    if not identities and not allow_no_markers:
        report['errors'].append({'error':'No Scannable Cite markers found; not a prepared scan input'})
    if any(m['part'] != 'word/document.xml' or m['textbox'] for m in report['markers']):
        report['warnings'].append('Markers in non-body parts/textboxes require a plugin-version conversion test.')
    if fields and identities:
        report['warnings'].append('Mixed active citations and markers: test conversion on a copy and preserve existing item identities.')
    if tracked:
        report['warnings'].append('Tracked changes present; deleted content was excluded. Resolve the intended visible version before conversion.')
    report['input_files_unchanged'] = hashes == (digest(docx), digest(mapping))
    if not report['input_files_unchanged']:
        report['errors'].append({'error':'An input changed during audit; discard this snapshot and retry'})
    report['ok'] = not report['errors']
    report['verification_scope'] = 'Syntax and supplied-map consistency only; no live Zotero lookup, plugin scan or Word refresh.'
    return report


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('docx', type=Path)
    ap.add_argument('--mapping', required=True, type=Path)
    ap.add_argument('--report', type=Path)
    ap.add_argument('--allow-no-markers', action='store_true')
    args = ap.parse_args(argv)
    if args.report and args.report.resolve() in (args.docx.resolve(), args.mapping.resolve()):
        ap.error('Report path must not overwrite the DOCX or mapping file')
    try:
        report = audit(args.docx, args.mapping, args.allow_no_markers)
        output = json.dumps(report, ensure_ascii=False, indent=2)
        if args.report:
            with args.report.open('x', encoding='utf-8') as f:
                f.write(output+'\n')
        print(output)
        return 0 if report['ok'] else 1
    except (OSError, ValueError, BadZipFile, ET.ParseError) as exc:
        print(json.dumps({'ok':False, 'error':str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
