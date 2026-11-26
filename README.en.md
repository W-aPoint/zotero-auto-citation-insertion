# Zotero Auto Citation Insertion

[简体中文](README.md) · **English**

Place Zotero's plain-text citation markers beside the sentences they support. In Word, put the cursor where you want the bibliography, click the scan button, and choose a Zotero collection. The button converts each marker in place, files the cited items in that collection, and calls Zotero's native **Add/Edit Bibliography** command.

[Download the Windows package](https://github.com/W-aPoint/zotero-auto-citation-insertion/releases/latest/download/Zotero-Auto-Citation-Insertion-Windows.zip) · [Download the Zotero companion extension](https://github.com/W-aPoint/zotero-auto-citation-insertion/releases/latest/download/WordZoteroScan.xpi)

## Install the Word button

You need desktop Microsoft Word for Windows, Zotero desktop, and Zotero's Word plugin. The workflow has been checked with Zotero 10.0.8 and ODF/DOCX Scan 3.2.2. The Word button and collection picker have Chinese labels, reproduced below.

1. Install the upstream [ODF/DOCX Scan plugin](https://zotero-odf-scan.github.io/zotero-odf-scan/). In Zotero, open **Tools → Plugins** and drag the upstream `.xpi` into the window. It provides the Scannable Cite export format and the DOCX converter.
2. Download and extract this project's Windows package. Save your documents and close Word. Open **Windows PowerShell** in the extracted directory and run `.\Install.ps1`. It registers the Word button for the current Windows user and creates the local pairing information used by Word and Zotero.
3. In Zotero's **Tools → Plugins**, install `dist/WordZoteroScan.xpi` from the extracted package. Reopen Word. The **References** tab now has a **Zotero** group with a button labeled **扫描并插入参考文献表** (Scan and insert bibliography).

If the installer cannot detect the Office architecture, check **File → Account → About Word**, then run `.\Install.ps1 -WordArchitecture x86` or `.\Install.ps1 -WordArchitecture x64`.

## From a sentence to a bibliography

1. **Copy a real marker.** In Zotero's **Settings → Export**, set the Quick Copy format to **Scannable Cite**. Select a reference, press `Ctrl+Alt+C`, and paste the marker next to the sentence it supports. For a grouped citation, place the markers directly next to one another. This example shows the shape of a personal-library marker. Use the complete marker copied from your own item.

   ```text
   { | <Author>, (<Year>) |p. 45 | |zu:<user-ID>:<item-key>}
   ```

2. **Position the bibliography.** Put the cursor on the line where the bibliography should begin, for example just below the References heading. Click **References → Zotero → 扫描并插入参考文献表**.
3. **Choose the collection.** Select a writable collection in the same Zotero library as the cited items, then click **插入或编辑参考文献表** (Insert or edit bibliography). The button creates editable citations at their original locations, preserves existing active citations, and adds the cited items to the collection. Running it again reuses existing memberships.
4. **Check and save.** Zotero inserts the new bibliography at the original cursor position. If prompted, select your citation style. If the document already has an active bibliography, Zotero opens its editor. The citations in the manuscript determine the bibliography. Your collection selection determines where the items are filed. Save the Word document when ready.

The button supports markers in the main text, including tables, and in footnotes and endnotes. It reports markers found in headers, footers, or text boxes so you can relocate them.

## Prepare markers in chat

The repository includes two optional assistant skills. Use [zotero-word-workflow](skills/zotero-word-workflow/SKILL.md) to place supplied markers in prose. Use [zotero-scannable-citations](skills/zotero-scannable-citations/SKILL.md) to replace numbered citations using a verified reference-to-item mapping.

Copy both directories under [skills](skills) into your assistant's skills directory, keeping them side by side. Codex commonly uses `~/.codex/skills/`. AntiGravity installations configured for `~/.gemini/config/skills/` can use that directory. Give the assistant the passage and the real markers copied from Zotero. For example:

> Use zotero-word-workflow to place these two citations beside the sentences they support and return the marked paragraph. I will operate Word and Zotero.

The assistant returns text for you to paste. You choose the collection and run the scan. To use the skill without the Word button, you can run the upstream **ODF Scan** yourself, open the converted DOCX, and use Word's **Zotero** tab to insert the bibliography at the cursor. [ODF/DOCX Scan instructions](https://zotero-odf-scan.github.io/zotero-odf-scan/) · [Zotero's Word guide](https://www.zotero.org/support/word_processor_plugin_usage)

## Source and license

[Word add-in](WordAddin) · [Zotero companion](zotero) · [Build script](Build.ps1). To build from source, run `.\Build.ps1` on Windows with the .NET Framework compiler and Office primary interop assemblies installed. The output appears under `dist/`. To run the included offline marker-validator tests with Python 3:

```sh
python skills/zotero-scannable-citations/scripts/test_validate_scan_docx.py
```

This project uses the [MIT License](LICENSE). The separately installed ODF/DOCX Scan plugin converts the markers, and Zotero's Word plugin formats the bibliography. The Word button communicates with Zotero over the local loopback interface, passing the markers and item identifiers needed for conversion and filing.
