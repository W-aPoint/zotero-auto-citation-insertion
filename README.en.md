# Zotero Auto Citation Insertion

[简体中文](README.md) · **English**

Tell the skill which Zotero collection should hold the papers and where you want the citation markers. It can file references while Zotero desktop is closed and place plain-text markers in your original Word document or in a copy. **You review the markers and explicitly confirm conversion to active citations and a bibliography in Word.**

[Download the Windows bundle](https://github.com/W-aPoint/zotero-auto-citation-insertion/releases/latest/download/Zotero-Auto-Citation-Insertion-Windows.zip) · [Download the Zotero companion](https://github.com/W-aPoint/zotero-auto-citation-insertion/releases/latest/download/WordZoteroScan.xpi)

## Give the skill a task

> Use My Library → My Project → References. Use Consensus to find papers supporting the second paragraph of my introduction, verify the sources, and file them in that collection. Put the five-field markers directly in my original Word document. I will confirm the conversion in Word.

Specify **the collection** and **the manuscript destination**. Identify the personal or group library and the full collection path. Use the parent path or collection key for names that occur more than once. Choices already given in the task carry forward.

| Your choice | Result |
| --- | --- |
| Edit the original | The original DOCX gains plain-text markers at the specified sentences. The file path stays the same. |
| Save a separate copy | A new DOCX at your chosen path contains the markers. |
| Return text in chat | You receive the marked sentences or paragraphs to paste yourself. |

## File, mark, and confirm

1. **Verify the papers.** We recommend using the **Consensus plugin** in your assistant (mention `@Consensus` where supported) to find candidate studies. Check the original paper or its PubMed or publisher record for the claim, authors, title, year, and DOI or PMID. You can also supply the references yourself. [Consensus](https://consensus.app/)
2. **File them in the named collection.** For a synced personal or group library, the skill can use the Zotero Web API to reuse existing records or import missing references. It reads back the real item keys and collection membership before preparing markers. This stage needs read/write access to the online library and works with Zotero desktop closed. For a local-only library or without cloud write access, file the items manually and supply the verified markers. [Commands and key setup](skills/zotero-word-workflow/references/preparation.md) · [Zotero Web API writes](https://www.zotero.org/support/dev/web_api/v3/write_requests)
3. **Place the markers.** The same five-field format works for an original document, a copy, or text returned in chat. Group citations by placing markers directly next to each other. Save and close the target Word document before file-level editing. You can keep working in it if you choose chat output. The example shows a page locator and a personal-library item. The real key comes from Zotero.

   ```text
   { | <Author>, (<Year>) |p. 45 | |zu:<user-ID>:<item-key>}
   ```

4. **Confirm in Word.** Open Zotero and let the new items sync. Open the marked manuscript, place the cursor where the bibliography should begin, and choose **References → Zotero → 扫描并插入参考文献表** (Scan and insert bibliography). Review the same target collection, then click **确认转换引文并打开参考文献表** (Confirm citation conversion and open the bibliography). Selecting a collection or cancelling leaves the citations untouched. Once you confirm, the button converts each marker in place and calls Zotero's native **Add/Edit Bibliography** at the original cursor. If a bibliography already exists, Zotero opens its editor. Review and save. [Zotero sync](https://www.zotero.org/support/sync) · [Word plugin guide](https://www.zotero.org/support/word_processor_plugin_usage)

The collection determines where the cited items are filed. The manuscript citations determine the bibliography. The button supports the main text, tables, footnotes and endnotes. It reports markers in headers, footers, and text boxes for you to relocate.

## Installation

**Skills.** Copy both directories under [skills](skills) side by side into `~/.codex/skills/` for Codex or your AntiGravity skills directory, such as `~/.gemini/config/skills/`. Use [zotero-word-workflow](skills/zotero-word-workflow/SKILL.md) for filing and placement or [zotero-scannable-citations](skills/zotero-scannable-citations/SKILL.md) for numbered citation mapping and read-only checks. The cloud helper uses Python 3.10+ and its standard library. Direct DOCX placement also requires `lxml`. Keep the API key in the local `ZOTERO_API_KEY` environment variable. [Commands and examples](skills/zotero-word-workflow/references/preparation.md)

**Word button for Windows desktop Word.**

1. Install Zotero desktop, its Word plugin, and the upstream [ODF/DOCX Scan](https://zotero-odf-scan.github.io/zotero-odf-scan/) plugin. In Zotero, open **Tools → Plugins** and install the upstream `.xpi`. It supplies the Scannable Cite export format and the DOCX converter.
2. Extract this project's Windows bundle. Save your documents and close Word. Open Windows PowerShell in the extracted directory and run `.\Install.ps1`. This registers the Word button for the current Windows user. If the installer cannot detect the Word architecture, check **File → Account → About Word** and pass `-WordArchitecture x86` or `-WordArchitecture x64`.
3. In Zotero's **Tools → Plugins**, install the bundled `dist/WordZoteroScan.xpi`, then reopen Word. The button appears under **References → Zotero**. The desktop workflow was checked with Zotero 10.0.8 and ODF/DOCX Scan 3.2.2.

You can also run the upstream **ODF Scan** yourself, open the resulting DOCX, and click **Add/Edit Bibliography** on Word's Zotero tab. [Upstream scanning guide](https://zotero-odf-scan.github.io/zotero-odf-scan/)

## Source and license

[Word add-in](WordAddin) · [Zotero companion](zotero) · [Tests](tests) · [MIT License](LICENSE). To build from source, run `.\Build.ps1` on Windows with the .NET Framework compiler and Office primary interop assemblies installed. The output is written to `dist/`.
