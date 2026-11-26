"""Insert verified plain-text Scannable Cite markers at explicitly chosen DOCX anchors."""

import argparse
import copy
import hashlib
import json
import os
import re
import shutil
import stat
import tempfile
import zipfile
from contextlib import contextmanager
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

from lxml import etree


PARTS = ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")
NAMESPACES = {"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
              "http://purl.oclc.org/ooxml/wordprocessingml/main"}
IDENTITY = re.compile(r"z([ug]):(0|[1-9][0-9]*):([A-Z0-9]{8})")
LIBRARY = re.compile(r"(?:users/(?:0|[1-9][0-9]*)|groups/[1-9][0-9]*)")
MARKER = re.compile(r"\{([^{}|\r\n]*)\|([^{}|\r\n]*)\|([^{}|\r\n]*)\|([^{}|\r\n]*)\|([^{}|\r\n]*)\}")
EXISTING_GROUP = re.compile(r"(?:\{[^{}\r\n]*\|[^{}\r\n]*\})+")
PLAIN = set("document body tbl tr tc p r t footnotes footnote endnotes endnote".split())
COMPOUND = set("fldSimple sdt hyperlink customXml smartTag del ins moveFrom moveTo txbxContent".split())
PROPERTY_CONTAINERS = {"r", "p", "tbl", "tr", "tc"}
REVISION_PROPERTIES = {"ins", "del", "cellIns", "cellDel", "cellMerge"}
PLAIN_PARAGRAPH_CHILDREN = {"pPr", "r", "proofErr", "commentRangeStart", "commentRangeEnd"}
RANGES = {
    "perm": ("permStart", "permEnd"),
    "bookmark": ("bookmarkStart", "bookmarkEnd"),
    "moveFrom": ("moveFromRangeStart", "moveFromRangeEnd"),
    "moveTo": ("moveToRangeStart", "moveToRangeEnd"),
    "customIns": ("customXmlInsRangeStart", "customXmlInsRangeEnd"),
    "customDel": ("customXmlDelRangeStart", "customXmlDelRangeEnd"),
    "customMoveFrom": ("customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd"),
    "customMoveTo": ("customXmlMoveToRangeStart", "customXmlMoveToRangeEnd"),
}
RANGE_TAGS = {name: (kind, i == 0) for kind, pair in RANGES.items() for i, name in enumerate(pair)}


def required_string(value, name):
    if not isinstance(value, str) or not value.strip() or any(ord(c) < 32 for c in value):
        raise ValueError(f"{name} must be nonempty single-line text")
    return value


def verified_items(mapping):
    if not isinstance(mapping, dict):
        raise ValueError("Expected a mapping object")
    if (mapping.get("applied", True) is not True or mapping.get("preview") is True
            or mapping.get("dry_run") is True or mapping.get("mode") in ("preview", "dry-run")
            or mapping.get("complete", True) is not True or mapping.get("errors") or mapping.get("pending")
            or not (mapping.get("applied") is True or mapping.get("verified") is True)):
        raise ValueError("Supply a complete applied or explicitly verified mapping; previews cannot mark a document")
    collection = mapping.get("collection")
    if not isinstance(collection, dict):
        raise ValueError("Identify the verified target collection with a path or key")
    key, path = collection.get("key"), collection.get("path")
    if key is not None and (not isinstance(key, str) or not re.fullmatch(r"[A-Z0-9]{8}", key)):
        raise ValueError("Invalid target collection key")
    if key is None:
        path = required_string(path, "Target collection path")
        if any(not component.strip() for component in path.split("/")):
            raise ValueError("Use a full target collection path or its verified key")
    scope = mapping.get("library")
    if scope is not None or mapping.get("applied") is True:
        if not isinstance(scope, str) or not LIBRARY.fullmatch(scope):
            raise ValueError("Identify the target Zotero library as users/N or groups/N")
        if mapping.get("applied") is True and scope == "users/0":
            raise ValueError("An applied cloud mapping requires a positive Zotero library ID")
    records = mapping.get("references")
    if not isinstance(records, list) or not records:
        raise ValueError("Mapping references must be a nonempty list")
    items = {}
    for rec in records:
        if not isinstance(rec, dict):
            raise ValueError("Expected an object for each mapping reference")
        ref_id = required_string(rec.get("id"), "Reference id")
        uri = required_string(rec.get("uri"), "Item uri")
        marker = required_string(rec.get("marker"), "Item marker")
        identity, fields = IDENTITY.fullmatch(uri), MARKER.fullmatch(marker)
        if (not identity or identity[3] != rec.get("key") or not fields
                or not fields[2].strip() or fields[5].strip() != uri):
            raise ValueError(f"{ref_id}: five-field marker, item key and uri must agree")
        item_scope = ("users/" if identity[1] == "u" else "groups/") + identity[2]
        if not LIBRARY.fullmatch(item_scope) or (scope is not None and scope != item_scope):
            raise ValueError(f"{ref_id}: item belongs to a different or invalid target library")
        scope = item_scope
        if ref_id in items:
            raise ValueError(f"Duplicate reference id: {ref_id}")
        items[ref_id] = (uri, marker)
    return items


def marker_group(references, items):
    if not isinstance(references, list) or not references:
        raise ValueError("An insertion requires a nonempty references list")
    markers, seen = [], set()
    for ref in references:
        entry = {"id": ref} if isinstance(ref, str) else ref
        if not isinstance(entry, dict) or set(entry) - {"id", "locator"}:
            raise ValueError("Use a reference id or an object with id and optional locator")
        ref_id = required_string(entry.get("id"), "Reference id")
        if ref_id not in items:
            raise ValueError(f"Unknown reference id: {ref_id}")
        uri, token = items[ref_id]
        if uri in seen:
            raise ValueError(f"Duplicate item in one citation group: {uri}")
        seen.add(uri)
        if "locator" in entry:
            locator = required_string(entry["locator"], "Locator")
            if any(ch in locator for ch in "{}|"):
                raise ValueError("A locator cannot contain marker delimiters")
            tokens = token.split("|")
            tokens[2] = f" {locator.strip()} "
            token = "|".join(tokens)
        markers.append(token)
    return "".join(markers)


def parse_xml(data):
    parser = etree.XMLParser(resolve_entities=False, no_network=True, strip_cdata=False)
    tree = etree.parse(BytesIO(data), parser)
    if tree.docinfo.doctype:
        raise ValueError("Unsupported document type declaration in Word package XML")
    return tree


@dataclass(eq=False)
class Paragraph:
    part: str
    chunks: list = field(default_factory=list)
    barriers: set = field(default_factory=set)
    size: int = 0
    text: str = ""
    groups: list = field(default_factory=list)

    def add(self, node, safe):
        end = self.size + len(node.text or "")
        if end > self.size:
            self.chunks.append((self.size, end, node, safe))
        self.size = end

    def select(self, start, end):
        chunks = [c for c in self.chunks if c[0] < end and c[1] > start]
        if any(not c[3] for c in chunks) or any(start < k < end for k in self.barriers):
            raise ValueError(f"{self.part}: anchor touches a field, revision or compound Word structure")
        return chunks


def story_paragraphs(tree, part):
    root = tree.getroot()
    ns = etree.QName(root).namespace
    expected = Path(part).stem
    if ns not in NAMESPACES or etree.QName(root).localname != expected:
        raise ValueError(f"Unsupported Word XML in {part}")
    w = "{" + ns + "}"
    revision_tags = {w + name for name in REVISION_PROPERTIES}
    paragraphs, stack, ranges, safe_cache = [], [], set(), {}
    depth = 0

    def plain(node):
        if node is None:
            return True
        if node in safe_cache:
            return safe_cache[node]
        tag = node.tag
        safe = tag.startswith(w) and tag[len(w):] in PLAIN and plain(node.getparent())
        if safe and tag in (w + "footnote", w + "endnote"):
            safe = node.get(w + "type", "normal") == "normal"
        if safe and tag[len(w):] in PROPERTY_CONTAINERS:
            props = node.find(tag + "Pr")
            if props is not None:
                safe = not any(isinstance(n.tag, str) and
                               (n.tag.endswith("Change") or n.tag in revision_tags)
                               for n in props.iter())
                style = props.find(w + "pStyle")
                if style is not None and style.get(w + "val", "").casefold() in ("bibliography", "zoterobibliography"):
                    safe = False
                for hidden in (w + "vanish", w + "webHidden"):
                    flag = props.find(hidden)
                    if flag is not None and flag.get(w + "val", "true").lower() not in ("false", "off", "0"):
                        safe = False
        safe_cache[node] = safe
        return safe

    for event, node in etree.iterwalk(root, events=("start", "end")):
        name = node.tag[len(w):] if isinstance(node.tag, str) and node.tag.startswith(w) else ""
        parent = node.getparent()
        if name == "p":
            if event == "start":
                if stack:
                    stack[-1].barriers.add(stack[-1].size)
                paragraph = Paragraph(part)
                paragraphs.append(paragraph)
                stack.append(paragraph)
            else:
                paragraph = stack.pop()
                paragraph.text = "".join(c[2].text for c in paragraph.chunks)
                paragraph.groups = list(EXISTING_GROUP.finditer(paragraph.text))
            continue
        parent_tag = parent.tag if parent is not None else None
        boundary = (name in COMPOUND or (parent_tag == w + "r" and name not in ("rPr", "t"))
                    or (parent_tag == w + "p" and name not in PLAIN_PARAGRAPH_CHILDREN and name not in RANGE_TAGS))
        if event == "start" and name == "fldChar":
            kind = node.get(w + "fldCharType")
            if kind == "begin":
                depth += 1
            elif kind == "end" and depth:
                depth -= 1
            elif kind != "separate" or not depth:
                raise ValueError(f"Unbalanced Word field in {part}")
        elif event == "start" and name in RANGE_TAGS:
            kind, opening = RANGE_TAGS[name]
            key = (kind, node.get(w + "id"))
            if opening and (kind != "bookmark" or "ZOTERO" in node.get(w + "name", "").upper()):
                ranges.add(key)
                boundary = True
            if not opening and key in ranges:
                ranges.remove(key)
                boundary = True
        if stack:
            paragraph = stack[-1]
            if boundary:
                paragraph.barriers.add(paragraph.size)
            if event == "start" and name == "t":
                paragraph.add(node, not depth and not ranges and plain(node))
        if event == "end" and name in ("footnote", "endnote") and (depth or ranges):
            raise ValueError(f"Unclosed field or protected range in a note in {part}")
    if depth or ranges:
        raise ValueError(f"Unclosed Word field or protected range in {part}")
    return paragraphs


def find_anchor(paragraphs, anchor, part):
    match = None
    for paragraph in paragraphs:
        if part and paragraph.part != part:
            continue
        start = paragraph.text.find(anchor)
        while start >= 0:
            end = start + len(anchor)
            if not any(g.start() <= start and end <= g.end() for g in paragraph.groups):
                if match is not None:
                    raise ValueError(f"Ambiguous anchor: {anchor!r}")
                match = paragraph, start, end
            start = paragraph.text.find(anchor, start + 1)
    if match is None:
        raise ValueError(f"Missing anchor in ordinary paragraph text: {anchor!r}")
    return match


def apply_plan(plan, items, paragraphs):
    entries = plan.get("insertions") if isinstance(plan, dict) else None
    if not isinstance(entries, list) or not entries:
        raise ValueError("The plan must contain a nonempty insertions list")
    edits, sites, touched = {}, {}, set()
    inserted = skipped = 0
    for entry in entries:
        if not isinstance(entry, dict) or set(entry) - {"part", "anchor", "position", "references"}:
            raise ValueError("Invalid insertion plan entry")
        anchor = required_string(entry.get("anchor"), "Anchor")
        side, part = entry.get("position"), entry.get("part")
        if side not in ("before", "after") or ("part" in entry and part not in PARTS):
            raise ValueError("Use before or after and an optional supported Word part")
        markers = marker_group(entry.get("references"), items)
        p, start, end = find_anchor(paragraphs, anchor, part)
        chunks = p.select(start, end)
        point = start if side == "before" else end
        for a, b, k in sites.setdefault(p, []):
            if k == point or a < point < b or start < k < end:
                raise ValueError(f"Plan entries collide at {anchor!r}")
        sites[p].append((start, end, point))
        adjacent = None
        for group in p.groups:
            if group.start() < end and start < group.end():
                raise ValueError(f"Anchor overlaps an existing marker: {anchor!r}")
            if (side == "before" and group.end() == start) or (side == "after" and group.start() == end):
                adjacent = group
        if adjacent is not None:
            p.select(adjacent.start(), adjacent.end())
            if adjacent.group() != markers or point in p.barriers:
                raise ValueError(f"An existing adjacent citation group conflicts at {anchor!r}")
            skipped += 1
            continue
        a, _, node, _ = chunks[0] if side == "before" else chunks[-1]
        edits.setdefault(node, []).append((point - a, markers))
        touched.add(p.part)
        inserted += 1
    for node, changes in edits.items():
        value = node.text
        for offset, markers in sorted(changes, reverse=True):
            value = value[:offset] + markers + value[offset:]
        node.text = value
    return touched, inserted, skipped


def check_package(archive):
    names = archive.namelist()
    if len(names) != len(set(names)) or "[Content_Types].xml" not in names or PARTS[0] not in names:
        raise ValueError("Expected a DOCX package with unique member names")
    types = parse_xml(archive.read("[Content_Types].xml"))
    settings = {"word/settings.xml"} if "word/settings.xml" in names else set()
    declared = [(node.get("PartName", "").lstrip("/"), node.get("ContentType", "").lower())
                for node in types.getroot() if isinstance(node.tag, str)]
    if (any(n.casefold().startswith("_xmlsignatures/") for n in names)
            or any("digital-signature" in content_type for _, content_type in declared)):
        raise ValueError("Digitally signed documents cannot be edited")
    for member, content_type in declared:
        if content_type.endswith(".settings+xml") and member:
            settings.add(member)
        if (content_type.endswith(".document.main+xml") and member != PARTS[0]
                or content_type.endswith(".footnotes+xml") and member != PARTS[1]
                or content_type.endswith(".endnotes+xml") and member != PARTS[2]):
            raise ValueError(f"Unsupported Word part location: {member}")
    for member in settings:
        if member not in names:
            raise ValueError(f"Missing declared Word settings: {member}")
        tree = parse_xml(archive.read(member))
        ns = etree.QName(tree.getroot()).namespace
        if ns not in NAMESPACES:
            raise ValueError("Unrecognized Word settings")
        w = "{" + ns + "}"
        for tag, attr in (("documentProtection", "enforcement"), ("writeProtection", None),
                          ("trackRevisions", "val")):
            for node in tree.iter(w + tag):
                if attr is None or node.get(w + attr, "true").lower() not in ("false", "0", "off"):
                    raise ValueError(f"Word settings prevent editing: {tag}")


def check_owner_files(path):
    markers = {"~$" + path.name, "~$" + path.name[2:], ".~lock." + path.name + "#"}
    if any((path.parent / name).exists() for name in markers):
        raise ValueError(f"Document has an editor lock file: {path}")


@contextmanager
def transaction_guard(path, in_place):
    if os.name != "nt" or not in_place:
        yield
        return
    # Keep cooperating writers excluded while Windows releases the source for replacement.
    # O_TEMPORARY makes Windows remove the lock even if the process exits unexpectedly.
    name = path.with_name("." + path.name + ".mark-docx.lock")
    fd = os.open(name, os.O_WRONLY | os.O_CREAT | os.O_EXCL | os.O_TEMPORARY, 0o600)
    try:
        yield
    finally:
        os.close(fd)


@contextmanager
def locked_source(path):
    if os.name == "nt":
        import ctypes
        import msvcrt
        from ctypes import wintypes

        api = ctypes.WinDLL("kernel32", use_last_error=True)
        create = api.CreateFileW
        create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                           wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        create.restype = wintypes.HANDLE
        handle = create(str(path), 0x80000000, 0, None, 3, 0x80, None)
        if handle == wintypes.HANDLE(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            fd = msvcrt.open_osfhandle(handle, os.O_RDONLY | os.O_BINARY)
        except BaseException:
            api.CloseHandle(wintypes.HANDLE(handle))
            raise
        stream = os.fdopen(fd, "rb")
    else:
        import fcntl

        stream = path.open("rb")
        try:
            fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BaseException:
            stream.close()
            raise
    with stream:
        yield stream


def fingerprint(stream):
    stream.seek(0)
    digest = hashlib.sha256()
    for block in iter(lambda: stream.read(1024 * 1024), b""):
        digest.update(block)
    return digest.digest()


def file_stamp(info):
    return info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns


def save_package(archive, stream, stage, trees, touched):
    if not touched:
        stream.seek(0)
        shutil.copyfileobj(stream, stage)
        return
    with zipfile.ZipFile(stage, "w") as result:
        result.comment = archive.comment
        for entry in archive.infolist():
            with result.open(copy.copy(entry), "w") as dst:
                if entry.filename in touched:
                    tree = trees[entry.filename]
                    info = tree.docinfo
                    dst.write(etree.tostring(tree, encoding=info.encoding or "UTF-8",
                                             xml_declaration=True, standalone=info.standalone))
                else:
                    with archive.open(entry) as src:
                        shutil.copyfileobj(src, dst)


def windows_replace_file(target, stage, backup):
    import ctypes
    from ctypes import wintypes

    replace = ctypes.WinDLL("kernel32", use_last_error=True).ReplaceFileW
    replace.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR,
                        wintypes.DWORD, wintypes.LPVOID, wintypes.LPVOID]
    replace.restype = wintypes.BOOL
    # Zero flags require Windows to merge the replaced file's DACL and named streams.
    if not replace(str(target), str(stage), str(backup), 0, None, None):
        raise ctypes.WinError(ctypes.get_last_error())


def replace_original(target, stage):
    if os.name != "nt":
        os.replace(stage, target)
        return None
    directory = Path(tempfile.mkdtemp(prefix="." + target.stem + ".recovery.", dir=target.parent))
    backup = directory / "original.docx"
    try:
        try:
            windows_replace_file(target, stage, backup)
        except OSError as error:
            # ERROR_UNABLE_TO_MOVE_REPLACEMENT_2 can move the original to the backup path.
            if backup.exists() and not os.path.lexists(target):
                try:
                    os.rename(backup, target)
                except OSError as restore_error:
                    raise OSError(f"Original preserved at {backup}: {restore_error}") from error
            if backup.exists():
                raise OSError(f"Replacement failed; original preserved at {backup}: {error}") from error
            raise
        try:
            backup.unlink(missing_ok=True)
        except OSError:
            return str(backup)
        return None
    finally:
        if not backup.exists():
            try:
                directory.rmdir()
            except OSError:
                pass  # A file monitor may briefly hold the empty recovery directory.


def mark_document(source, mapping, plan, *, in_place=False, output=None):
    if in_place == (output is not None):
        raise ValueError("Choose exactly one of --in-place and --output")
    source = Path(source).resolve(strict=True)
    target = source if in_place else Path(output).absolute()
    if source.suffix.lower() != ".docx" or target.suffix.lower() != ".docx":
        raise ValueError("Source and destination must have the .docx extension")
    if not in_place and os.path.lexists(target):
        raise FileExistsError(f"Output already exists: {target}")
    items = verified_items(mapping)
    check_owner_files(source)
    check_owner_files(target)
    with transaction_guard(source, in_place), locked_source(source) as stream:
        original = os.fstat(stream.fileno())
        if in_place and not original.st_mode & stat.S_IWUSR:
            raise PermissionError("Cannot edit a read-only original")
        initial_digest = fingerprint(stream)
        with zipfile.ZipFile(stream) as archive:
            check_package(archive)
            trees = {p: parse_xml(archive.read(p)) for p in PARTS if p in archive.namelist()}
            paragraphs = [p for part, tree in trees.items() for p in story_paragraphs(tree, part)]
            touched, inserted, skipped = apply_plan(plan, items, paragraphs)
            report = {"output": str(target), "inserted_groups": inserted, "already_present": skipped}
            if in_place and not touched:
                return report
            fd, name = tempfile.mkstemp(prefix="." + target.stem + ".", suffix=".tmp", dir=target.parent)
            stage = Path(name)
            try:
                with os.fdopen(fd, "wb") as dst:
                    save_package(archive, stream, dst, trees, touched)
                    dst.flush()
                    os.fsync(dst.fileno())
                if fingerprint(stream) != initial_digest or file_stamp(source.stat()) != file_stamp(original):
                    raise ValueError("The source changed during processing. Reread the document and retry")
                check_owner_files(source)
                check_owner_files(target)
                if in_place:
                    os.chmod(stage, stat.S_IMODE(original.st_mode))
                    if os.name == "nt":
                        stream.close()
                        if file_stamp(source.stat()) != file_stamp(original):
                            raise ValueError("The source changed before replacement")
                    retained = replace_original(target, stage)
                    if retained:
                        report["retained_original"] = retained
                elif os.name == "nt":
                    os.rename(stage, target)  # Windows fails atomically if the destination already exists.
                else:
                    os.link(stage, target)
            finally:
                stage.unlink(missing_ok=True)
            return report


def main(argv=None):
    cli = argparse.ArgumentParser(description=__doc__)
    cli.add_argument("manuscript", type=Path)
    cli.add_argument("--mapping", required=True, type=Path)
    cli.add_argument("--plan", required=True, type=Path)
    mode = cli.add_mutually_exclusive_group(required=True)
    mode.add_argument("--in-place", action="store_true")
    mode.add_argument("--output", type=Path)
    args = cli.parse_args(argv)
    try:
        mapping = json.loads(args.mapping.read_text(encoding="utf-8-sig"))
        plan = json.loads(args.plan.read_text(encoding="utf-8-sig"))
        report = mark_document(args.manuscript, mapping, plan, in_place=args.in_place, output=args.output)
    except (ValueError, OSError, zipfile.BadZipFile, etree.XMLSyntaxError, RuntimeError) as error:
        cli.exit(1, f"mark_docx: {error}\n")
    print(json.dumps(report, ensure_ascii=True))


if __name__ == "__main__":
    main()
