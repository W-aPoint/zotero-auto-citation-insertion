# 分类与稿件准备 / Prepare a collection and manuscript

## 开始前 / Before you start

记录个人库或群组库、目标分类的完整路径，以及 **原稿 / 副本 / 仅回复** 的选择。已明确的信息直接沿用。指定原稿时，工作文件和交付文件是同一份 DOCX。

Record the library, the full collection path, and the author's choice of original DOCX, separate copy, or text in chat. The assistant can prepare the metadata and insertion plan. The author confirms the final conversion in Word.

## 在云端归档 / File items without launching Zotero

Use Python 3.10+ and a Zotero Web API key with read/write access to the selected library. Configure `ZOTERO_API_KEY` in your local environment or secret manager. Keep the value out of chat, manuscripts, command arguments, and the repository. For personal or group libraries, use the corresponding numeric online ID in `users/ID` or `groups/ID`. The personal numeric ID appears on the API settings page. [API settings](https://www.zotero.org/settings/keys) · [Zotero Web API](https://www.zotero.org/support/dev/web_api/v3/write_requests)

下面的文件在本次任务的私有工作目录中准备。用论文原始记录替换示意值。检索时可使用 [Consensus](https://consensus.app/)，随后核对出版方记录中的题名、作者和年份。

`papers.json` contains verified metadata. Replace these placeholders with the actual publication record. The local `id` links the record to the insertion plan.

```json
[
  {
    "id": "R1",
    "title": "<verified article title>",
    "creators": [{"creatorType": "author", "firstName": "<given name>", "lastName": "<family name>"}],
    "date": "<publication year>",
    "DOI": "<verified DOI>"
  }
]
```

From the installed `zotero-word-workflow` skill directory, run a preview. Replace `USER_ID` with the online user ID and `Project/References` with the chosen collection path, or use `--collection-key` and the real collection key. The skill can carry out these calls on the user's behalf. A collection named by the user must match the resolved target.

```sh
python scripts/zotero_collection.py --library users/USER_ID --collection "Project/References" --papers papers.json
python scripts/zotero_collection.py --library users/USER_ID --collection "Project/References" --papers papers.json --apply --output filed.json
```

Use the verified output of the write step for marker placement. An existing item can also be identified by its verified `itemKey`, DOI or PMID. For a new record, supply the verified title, year and creators. The script handles bibliographic records and collection membership. PDF attachments remain a separate task.

For a local-only library, the author can file the items in Zotero and verify the membership manually. Such a mapping must name the target collection: set `verified` to `true`, include a `collection` object with its real `path` or `key`, and identify the library when known (for example `users/USER_ID`). All references must belong to that library. Each entry in `references` gives the local `id`, real Zotero `key`, verified five-field `marker`, and `uri`, the short identity from the marker's last field. Use `users/0` only for items whose local-library identity has actually been confirmed.

## 写入标记 / Place plain-text markers

Choose a unique, exact fragment of the target sentence as the anchor. To insert immediately before the final period, use the sentence text without the period as the anchor and choose `after`. Multiple references are joined without a separator. Install the Python package `lxml` in the environment used for DOCX placement. The cloud helper uses only Python's standard library.

```json
{
  "insertions": [
    {"anchor": "<exact text from the manuscript>", "position": "after", "references": ["R1"]}
  ]
}
```

Save the plan as `insertions.json`. Run **one** of these commands to match the author's choice. For chat-only delivery, return the marked sentences directly.

```sh
python scripts/mark_docx.py manuscript.docx --mapping filed.json --plan insertions.json --in-place
python scripts/mark_docx.py manuscript.docx --mapping filed.json --plan insertions.json --output marked-copy.docx
```

An anchor must be unique within one paragraph, but can span Word text runs. Body text, tables, footnotes and endnotes are supported. File-level placement leaves active fields and protected regions intact. If the anchor is ambiguous, give a longer exact fragment. If Word has unsaved edits, save and close that document first. A separate output path must be new. A retry skips identical, already placed groups.

After checking the marked manuscript, the author opens Zotero, completes synchronization, and explicitly confirms the conversion in Word. [Zotero sync](https://www.zotero.org/support/sync) · [Word bibliography command](https://www.zotero.org/support/word_processor_plugin_usage)
