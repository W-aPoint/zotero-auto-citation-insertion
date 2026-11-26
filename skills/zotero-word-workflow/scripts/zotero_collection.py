"""Preview or file verified bibliographic records in an existing Zotero Web API v3 collection."""

import argparse
from collections import defaultdict
import http.client
import json
import os
from pathlib import Path
import re
import sys
import tempfile
import time
import unicodedata
from urllib import error, parse, request


LIBRARY = re.compile(r"(users|groups)/([1-9][0-9]*)\Z")
KEY = re.compile(r"[A-Z0-9]{8}\Z")
YEAR = re.compile(r"(?<!\d)([1-9][0-9]{3})(?!\d)")
NONBIB = {"note", "attachment", "annotation"}
BIB_FIELDS = set("""title date DOI PMID PMCID abstractNote publicationTitle publisher place volume issue
section partNumber partTitle pages series seriesTitle seriesText seriesNumber journalAbbreviation
url accessDate ISSN ISBN archive archiveLocation shortTitle language libraryCatalog callNumber rights extra
edition bookTitle proceedingsTitle conferenceName repository archiveID reportNumber institution university
reportType thesisType type number numberOfVolumes numPages websiteTitle websiteType""".split())


class Problem(Exception):
    pass


class RemoteProblem(Problem):
    def __init__(self, status, message):
        super().__init__(message)
        self.status = status


class NoRedirect(request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if fp is not None:
            fp.close()
        raise RemoteProblem(code, "The Zotero API redirected the request; credentials were not forwarded.")


def key(value):
    if not isinstance(value, str) or not KEY.fullmatch(value):
        raise Problem("Expected an eight-character Zotero object key.")
    return value


def integer(value, field):
    if not re.fullmatch(r"[0-9]+", str(value)):
        raise Problem(f"The Zotero response lacks a valid {field}.")
    return int(value)


class API:
    def __init__(self, library, token):
        if not LIBRARY.fullmatch(library):
            raise Problem("Choose a library such as users/123 or groups/456 using its numeric ID.")
        if not token or not re.fullmatch(r"[!-~]+", token):
            raise Problem("Set ZOTERO_API_KEY locally to a key with access to the selected library.")
        self.library = library
        self._headers = {"Zotero-API-Version": "3", "Zotero-API-Key": token, "Accept": "application/json"}
        self.opener = request.build_opener(NoRedirect())
        self.resume_at = 0.0

    def call(self, method, path, body=None, version=None):
        url = f"https://api.zotero.org/{self.library}/{path}"
        headers = self._headers.copy()
        payload = None
        if body is not None:
            payload = json.dumps(body, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"
        if version is not None:
            headers["If-Unmodified-Since-Version"] = str(version)
        wait = self.resume_at - time.monotonic()
        if wait > 0:
            time.sleep(wait)
        try:
            req = request.Request(url, method=method, headers=headers, data=payload)
            with self.opener.open(req, timeout=30) as response:
                if str(response.headers.get("Backoff", "")).isdigit():
                    self.resume_at = time.monotonic() + int(response.headers["Backoff"])
                raw = response.read()
                return (json.loads(raw) if raw else None), response.headers
        except error.HTTPError as exc:
            hint = {401: "check ZOTERO_API_KEY", 403: "check the key's access to the selected library",
                    404: "check the library or object key", 409: "library locked; rerun after syncing",
                    412: "the library or item changed"}.get(exc.code, "check the request and retry")
            if exc.code in (429, 503):
                delay = exc.headers.get("Retry-After", "")
                hint = f"retry after {delay} seconds" if str(delay).isdigit() else "retry later"
            exc.close()
            raise RemoteProblem(exc.code, f"Zotero HTTP {exc.code}: {hint}.") from None
        except (OSError, http.client.HTTPException, ValueError):
            raise RemoteProblem(0, "No usable response from Zotero. Rerun to reconcile any completed write.") from None


def unpack(obj, library):
    kind, number = library.split("/")
    if not isinstance(obj, dict) or not isinstance(obj.get("data"), dict):
        raise Problem("Invalid Zotero JSON object.")
    scope = obj.get("library") or {}
    if not isinstance(scope, dict) or scope.get("type") != {"users": "user", "groups": "group"}[kind] or str(scope.get("id")) != number:
        raise Problem("Returned object belongs to a different or unspecified Zotero library.")
    data = obj["data"]
    object_key = key(obj.get("key"))
    version = integer(obj.get("version"), "object version")
    if data.get("key", object_key) != object_key or data.get("version", version) != version:
        raise Problem("Zotero returned inconsistent object identity or version information.")
    return dict(data, key=object_key, version=version)


def snapshot(api, endpoint, accept):
    """Use numeric in-scope pagination and discard mixed-version snapshots."""
    start, version, total, seen = 0, None, None, set()
    while True:
        rows, headers = api.call("GET", f"{endpoint}?limit=100&start={start}")
        current = integer(headers.get("Last-Modified-Version"), "library version")
        count = integer(headers.get("Total-Results"), "total result count")
        if version is not None and (current != version or count != total):
            raise RemoteProblem(412, "Library changed during pagination; rerun to refresh the snapshot.")
        if not isinstance(rows, list) or start + len(rows) > count or (not rows and start < count):
            raise Problem("Zotero returned an incomplete or inconsistent page.")
        version, total = current, count
        for obj in rows:
            data = unpack(obj, api.library)
            if data["key"] in seen:
                raise Problem("A Zotero object occurred on more than one page.")
            seen.add(data["key"])
            accept(data)
        start += len(rows)
        if start == count:
            return version


def resolve_collection(api, path, selected_key):
    def add_collection(row):
        if not row.get("deleted"):
            if not isinstance(row.get("name"), str) or not row["name"]:
                raise Problem("The Zotero collection has no readable name.")
            entries[row["key"]] = row

    for attempt in range(3):
        entries = {}
        try:
            snapshot(api, "collections", add_collection)
            break
        except RemoteProblem as exc:
            if exc.status != 412 or attempt == 2:
                raise
    if selected_key:
        target = key(selected_key)
        if target not in entries:
            raise Problem("The collection key does not exist in this library.")
    else:
        if not path or any(not part.strip() for part in path.split("/")):
            raise Problem("Provide a collection path such as Root/Child, or select its key.")
        target = False
        for part in path.split("/"):
            matches = [k for k, row in entries.items() if (row.get("parentCollection") or False) == target and row.get("name") == part]
            if len(matches) != 1:
                raise Problem("Collection path is missing or ambiguous. Check its exact name or use --collection-key.")
            target = matches[0]
    lineage, visited, cursor = [], set(), target
    while cursor:
        if cursor in visited or cursor not in entries:
            raise Problem("The collection hierarchy is incomplete or contains a cycle.")
        visited.add(cursor)
        lineage.append(entries[cursor]["name"])
        cursor = entries[cursor].get("parentCollection")
    return {"key": target, "path": "/".join(reversed(lineage))}


def fold(value):
    return " ".join(unicodedata.normalize("NFKC", str(value or "")).casefold().split())


def author_year(data):
    creator = next((x for x in data.get("creators", []) if x.get("creatorType") == "author"), {})
    year = YEAR.search(str(data.get("date", "")))
    return creator.get("lastName") or creator.get("name") or "", year.group(1) if year else ""


def signature(data):
    author, year = author_year(data)
    return fold(data.get("title")), year, fold(author)


def identifiers(data):
    pairs = [(name, data.get(name, "")) for name in ("DOI", "PMID")]
    pairs.extend(re.findall(r"^\s*(DOI|PMID|PubMed ID)\s*:\s*(\S+)\s*$", str(data.get("extra", "")), re.I | re.M))
    url = str(data.get("url", "")).strip()
    if re.match(r"https?://(?:dx\.)?doi\.org/", url, re.I):
        pairs.append(("DOI", url))
    pubmed = re.fullmatch(r"https?://(?:pubmed\.ncbi\.nlm\.nih\.gov/|(?:www\.)?ncbi\.nlm\.nih\.gov/pubmed/)([1-9][0-9]*)/?", url, re.I)
    if pubmed:
        pairs.append(("PMID", pubmed.group(1)))
    found = set()
    for name, raw in pairs:
        value = str(raw or "").strip()
        if name.upper() == "DOI":
            if re.match(r"https?://(?:dx\.)?doi\.org/", value, re.I):
                value = parse.unquote(parse.urlsplit(value).path[1:])
            value = re.sub(r"^doi:\s*", "", value, flags=re.I)
            if re.fullmatch(r"10\.[0-9]{4,9}/\S+", value, re.I):
                found.add(("DOI", value.casefold()))
        elif re.fullmatch(r"[1-9][0-9]*", value):
            found.add(("PMID", value))
    return found


def prepare(raw):
    if not isinstance(raw, list) or not raw:
        raise Problem("papers.json must contain a nonempty JSON array.")
    papers, seen = [], set()
    for row in raw:
        if not isinstance(row, dict) or type(row.get("id")) not in (str, int):
            raise Problem("Each paper needs a unique, nonempty id.")
        paper_id = str(row["id"]).strip()
        if not paper_id or paper_id in seen:
            raise Problem("Paper ids must be unique and nonempty.")
        seen.add(paper_id)
        unsupported = set(row) - BIB_FIELDS - {"id", "itemKey", "itemType", "creators"}
        if unsupported:
            raise Problem(f"Unsupported paper field for {paper_id}: {', '.join(sorted(unsupported))}.")
        data = {field: row[field] for field in BIB_FIELDS.intersection(row) if row[field] is not None}
        for field in ("PMID", "date"):
            if type(data.get(field)) is int:
                data[field] = str(data[field])
        if any(not isinstance(v, str) for v in data.values()):
            raise Problem(f"Bibliographic fields must be strings for {paper_id}.")
        data["itemType"] = row.get("itemType", "journalArticle")
        if not isinstance(data["itemType"], str) or not re.fullmatch(r"[A-Za-z]+", data["itemType"]) or data["itemType"] in NONBIB:
            raise Problem(f"Choose a bibliographic item type for {paper_id}.")
        creators = row.get("creators") if row.get("creators") is not None else []
        if not isinstance(creators, list):
            raise Problem(f"Supply Zotero creators as an array for {paper_id}.")
        data["creators"] = []
        for creator in creators:
            if not isinstance(creator, dict) or set(creator) - {"creatorType", "name", "firstName", "lastName"} or not all(isinstance(v, str) for v in creator.values()):
                raise Problem(f"Invalid Zotero creator for {paper_id}.")
            if bool(fold(creator.get("name"))) == bool(fold(creator.get("lastName"))) or (creator.get("name") and "firstName" in creator):
                raise Problem(f"Use a corporate name or a personal lastName for each creator in {paper_id}.")
            data["creators"].append({"creatorType": "author", **creator})
        ids = identifiers(data)
        for field in ("DOI", "PMID"):
            values = [value for name, value in ids if name == field]
            if len(values) > 1 or (data.get(field) and not identifiers({field: data[field]})):
                raise Problem(f"Invalid or conflicting {field} for {paper_id}.")
            if data.get(field):
                data[field] = values[0]
        explicit = key(row["itemKey"]) if row.get("itemKey") not in (None, "") else None
        if not explicit and not ids and not all(signature(data)):
            raise Problem(f"Supply an itemKey, DOI, PMID, or verified title, year and lead author for {paper_id}.")
        papers.append((paper_id, explicit, data))
    return papers


def regular(data):
    return bool(data.get("itemType")) and data["itemType"] not in NONBIB and not data.get("parentItem") and not data.get("deleted")


class Inventory:
    def __init__(self, api):
        self.by_id = defaultdict(set)
        self.by_signature = defaultdict(set)
        self.titles = set()
        self.version = snapshot(api, "items/top", self.add)

    def add(self, data):
        if regular(data):
            sig = signature(data)
            self.titles.add(sig[0])
            self.by_signature[sig].add(data["key"])
            for identifier in identifiers(data):
                self.by_id[identifier].add(data["key"])

    def match(self, data):
        ids, sig = identifiers(data), signature(data)
        matches = set()
        if ids:
            for identifier in ids:
                matches.update(self.by_id.get(identifier, ()))
        else:
            matches = self.by_signature.get(sig, set())
        if len(matches) > 1:
            raise Problem("Multiple items match this record. Verify the intended item and supply its itemKey.")
        if matches:
            return next(iter(matches))
        if sig[0] and sig[0] in self.titles:
            raise Problem("An item has the same title but its identity is unresolved. Verify the existing itemKey and metadata.")
        return None


def verify_identity(source, item, explicit):
    if not regular(item):
        raise Problem("The selected object is not an active, top-level bibliographic item.")
    supplied, stored = identifiers(source), identifiers(item)
    for field, value in supplied:
        if any(name == field and other != value for name, other in stored):
            raise Problem(f"The selected item's {field} conflicts with the supplied identity.")
    if supplied & stored:
        return
    proposed, actual = signature(source), signature(item)
    if explicit and not supplied and all(not value or value == current for value, current in zip(proposed, actual)):
        return
    if (not supplied or explicit) and all(proposed) and proposed == actual:
        return
    raise Problem("The item could not be verified against the supplied identifiers or bibliographic metadata.")


def read_item(api, item_key, source, explicit):
    obj, _ = api.call("GET", f"items/{key(item_key)}")
    item = unpack(obj, api.library)
    if item["key"] != item_key:
        raise Problem("Zotero returned a different key than requested.")
    verify_identity(source, item, explicit)
    membership = item.get("collections")
    if not isinstance(membership, list) or not all(isinstance(x, str) and KEY.fullmatch(x) for x in membership):
        raise Problem("The item's collection membership could not be verified.")
    return item


def created_key(result):
    if not isinstance(result, dict):
        raise Problem("Zotero did not return the result of the item creation. Rerun to reconcile.")
    failed = result.get("failed") or {}
    if failed:
        status = integer(failed.get("0", {}).get("code"), "per-item HTTP status")
        raise RemoteProblem(status, f"Zotero rejected this item (HTTP {status}). Check its metadata and library permissions.")
    entry = (result.get("successful") or {}).get("0")
    if isinstance(entry, dict):
        return key(entry.get("key"))
    return key((result.get("success") or {}).get("0"))


class Filing:
    def __init__(self, api, collection, apply):
        self.api = api
        self.collection = collection
        self.apply = apply
        self.inventory = None

    def one(self, paper):
        for attempt in range(3):
            try:
                return self._one(*paper)
            except RemoteProblem as exc:
                if exc.status != 412 or attempt == 2:
                    raise
                self.inventory = None

    def _one(self, paper_id, explicit_key, source):
        if not explicit_key and self.inventory is None:
            self.inventory = Inventory(self.api)
        item_key = explicit_key or self.inventory.match(source)
        target = self.collection["key"]
        item = read_item(self.api, item_key, source, bool(explicit_key)) if item_key else None
        if explicit_key:
            missing_ids = identifiers(source) - identifiers(item)
            if missing_ids and self.inventory is None:
                self.inventory = Inventory(self.api)
            if any(self.inventory.by_id.get(identifier, set()) - {item_key} for identifier in missing_ids):
                raise Problem("A supplied identifier belongs to a different item from the selected itemKey.")
        current = item["collections"] if item else []
        if item is None and not all(signature(source)):
            raise Problem("No matching library item. Supply a verified title, year and lead author to create this record.")
        action = "create" if not item else ("already-member" if target in current else "add-to-collection")
        ref = {"id": paper_id, "title": (item or {}).get("title") or source.get("title", ""), "action": action, "verified": False,
               "key": None, "uri": None, "marker": None}
        if not self.apply:
            if item_key:
                ref["candidateKey"] = item_key
            return ref

        if action == "create":
            result, headers = self.api.call("POST", "items", [dict(source, collections=[target])], self.inventory.version)
            item_key = created_key(result)
            item = read_item(self.api, item_key, source, False)
            if signature(item) != signature(source) or not identifiers(source).issubset(identifiers(item)):
                raise Problem("The newly created item's identifiers or bibliographic details failed verification.")
            self.inventory.version = integer(headers.get("Last-Modified-Version"), "library version after creation")
            self.inventory.add(item)
        elif action == "add-to-collection":
            self.api.call("PATCH", f"items/{item_key}", {"collections": current + [target]}, item["version"])
            item = read_item(self.api, item_key, source, bool(explicit_key))
        if target not in item["collections"] or not set(current).issubset(item["collections"]):
            raise Problem("The item was read back without the expected collection memberships. Rerun to reconcile.")
        kind, number = self.api.library.split("/")
        uri = f"{'zu' if kind == 'users' else 'zg'}:{number}:{item_key}"
        author, year = author_year(item)
        fallback_author, fallback_year = author_year(source)
        label = author or fallback_author or ref["title"] or item_key
        label = " ".join(re.sub(r"[{}|]", " ", label).split()) or item_key
        if year or fallback_year:
            label += f", ({year or fallback_year})"
        marker = f"{{ | {label} | | |{uri}}}"
        ref.update(title=item.get("title") or ref["title"], key=item_key, uri=uri, marker=marker, verified=True)
        return ref


class Output:
    def __init__(self, filename, input_path):
        self.path, self.prior = filename, None
        if not filename:
            return
        if filename.resolve() == input_path.resolve() or (filename.exists() and input_path.samefile(filename)):
            raise Problem("Choose an output path different from the input papers file.")
        if filename.is_symlink() or filename.suffix.lower() != ".json" or not filename.parent.is_dir():
            raise Problem("Choose a regular .json output path in an existing directory.")
        if filename.exists():
            try:
                self.prior = json.loads(filename.read_text(encoding="utf-8-sig"))
            except (UnicodeError, ValueError, OSError):
                raise Problem("The output path contains another file. Choose a new mapping.json path.") from None
            if (not isinstance(self.prior, dict) or not isinstance(self.prior.get("references"), list)
                    or not isinstance(self.prior.get("collection"), dict) or type(self.prior.get("applied")) is not bool):
                raise Problem("The output file is not a Zotero mapping. Choose a new path.")

    def bind(self, library, collection, apply):
        if self.prior is not None:
            prior = self.prior
            if prior.get("library") != library or (prior.get("collection") or {}).get("key") != collection["key"]:
                raise Problem("The existing output belongs to a different library or collection. Choose a new path.")
            if prior["applied"] and not apply:
                raise Problem("Choose a separate preview path to preserve the verified mapping.")

    def write(self, report):
        if not self.path:
            return
        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=self.path.parent, delete=False) as tmp:
                temp_path = Path(tmp.name)
                json.dump(report, tmp, ensure_ascii=False, indent=2)
                tmp.write("\n")
            os.replace(temp_path, self.path)
        finally:
            if temp_path and temp_path.exists():
                temp_path.unlink()


def execute(api, papers, path, selected_key, apply, output):
    collection = resolve_collection(api, path, selected_key)
    output.bind(api.library, collection, apply)
    report = {"library": api.library, "collection": collection, "mode": "apply" if apply else "preview",
              "applied": False, "complete": False, "references": [], "errors": []}
    output.write(report)
    filing = Filing(api, collection, apply)
    for pos, paper in enumerate(papers):
        try:
            report["references"].append(filing.one(paper))
        except Problem as exc:
            report["errors"].append({"id": paper[0], "message": str(exc)})
            if isinstance(exc, RemoteProblem) and exc.status not in (400, 404, 422):
                report["pending"] = [p[0] for p in papers[pos + 1:]]
                break
        output.write(report)
    if apply and not report["errors"]:
        try:
            if resolve_collection(api, path, selected_key) != collection:
                raise Problem("The selected collection changed during filing. Verify the current collection path.")
        except Problem as exc:
            report["errors"].append({"message": str(exc)})
    report["complete"] = not report["errors"]
    report["applied"] = bool(apply and report["complete"])
    output.write(report)
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--library", required=True, help="users/123 or groups/456")
    selector = parser.add_mutually_exclusive_group(required=True)
    selector.add_argument("--collection", help="Exact slash-separated path in the selected library")
    selector.add_argument("--collection-key", help="Existing Zotero collection key")
    parser.add_argument("--papers", required=True, type=Path, help="JSON array of verified bibliographic records")
    parser.add_argument("--apply", action="store_true", help="File records and verify each item before issuing markers")
    parser.add_argument("--output", type=Path, help="Write preview or verified JSON mapping atomically")
    args = parser.parse_args(argv)
    try:
        papers = prepare(json.loads(args.papers.read_text(encoding="utf-8-sig")))
        output = Output(args.output, args.papers)
        api = API(args.library, os.environ.get("ZOTERO_API_KEY"))
        report = execute(api, papers, args.collection, args.collection_key, args.apply, output)
        # ASCII escapes preserve Unicode JSON on consoles with a legacy encoding.
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0 if report["complete"] else 1
    except (Problem, OSError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
