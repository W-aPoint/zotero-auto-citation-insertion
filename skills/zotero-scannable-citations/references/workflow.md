# 标记、映射与转换

使用 Skill 时先确定目标文库和分类，并记录作者选择的原稿、副本或仅回复。云端归档和写入原 DOCX 的方法见[分类与稿件准备](../../zotero-word-workflow/references/preparation.md)。下文的插件转换由作者在 Word 确认后执行，或由作者自行使用上游扫描器。

## 五字段 Scannable Cite

```text
{前缀|可读作者年份|定位信息|后缀|真实条目标识}
```

使用 ASCII 花括号和四个 ASCII `|` 分隔符，空字段保留；不要将管道符替换为中文全角符号。作者年份只供阅读，实际文献身份由最后一字段决定。字段中不能包含未转义的花括号或管道符；出现这些字符时从插件生成一条合法标记再处理，不自行猜测转义规则。

以下 `<USER_ID>`、`<GROUP_ID>`、`<ITEM_KEY>` 是文档说明中的占位符，不能写入实际交付稿：

```text
{ | Author, (2025) | | |zu:<USER_ID>:<ITEM_KEY>}
{See | Author, (2025) |p. 45 |for comparison |zu:<USER_ID>:<ITEM_KEY>}
{ | Author, (2025) | | |zg:<GROUP_ID>:<ITEM_KEY>}
```

优先在 Zotero 的导出/Quick Copy 中选择 **Scannable Cite**，从真实条目复制或拖放获取标记。单纯得到 RIS/BibTeX 的作者、DOI不代表已获得 Zotero 条目键。

个人库使用 `zu:用户ID:条目键`；群组库使用 `zg:群组ID:条目键`。同步用户ID/群组ID不能替换成本机 libraryID。当前插件可将实际的 `http(s)://zotero.org/users/<ID>/items/<KEY>` 和 `/groups/<ID>/items/<KEY>` 转为相应短标识。个人本地 URI `/users/local/<HASH>/items/<KEY>` 对应 `zu:0:<KEY>`，只在确有此本地身份来源时使用；不能把缺少账号信息的任意条目默认为0。实际条目键通常为8位大写字母/数字。

合引直接相邻：

```text
{ | Author A, (2024) | | |真实标识A}{ | Author B, (2025) | | |真实标识B}
```

每个标记是一次条目引用，连续相邻的一串标记按一组统计。同一文献重复出现，增加引用次数但不增加唯一文献数。不要在相邻标记间自行添加分号。脚本按严格相邻统计组数；带空格分隔会报告相邻性提示，实际合并情况仍需扫描核实。

页码写在第三字段，例如 `p. 45`、`pp. 45-48`，标签与数值间有空格；章可用 `ch. 3`，节可用 `sec. 2`。抑制作者可在可读作者字段起始加 `-`，具体效果用所选样式核实。前后缀需要斜体/粗体时按插件支持的 `*italic*` / `**bold**` 使用。

## 映射表与核验脚本

脚本接受 JSON 根数组，或根对象中的 `references` 数组/字典。每条记录至少给出 `marker` 或 `uri`；推荐同时记录原序号、题名和真实来源证据，证据可指向已有导出文件或经核实的 Zotero 条目。脚本不会连接 Zotero，也不会替你验证来源证据。

```json
{
  "references": {
    "R001": {
      "number": 1,
      "title": "从真实条目读取的题名",
      "uri": "zu:<USER_ID>:<ITEM_KEY>",
      "marker": "{ | 从真实条目读取的作者年份 | | |zu:<USER_ID>:<ITEM_KEY>}",
      "evidence_source": "真实导出文件或经核实的条目来源"
    }
  }
}
```

该示例不是可运行数据，所有占位符均需真实值。旧项目中以文献编号为记录索引、以 `marker` 保存短标识的映射表也可读取。`key` 存在时核对是否与标识末尾一致。序号相同而对应不同文献时拒绝，重复的相同身份可合并。

```text
python scripts/validate_scan_docx.py scan.docx --mapping map.json --report qa.json
```

默认要求至少一条扫描标记；仅盘点活动引文稿时加 `--allow-no-markers`。退出码0表示静态检查通过，1表示发现语法/映射/文档问题，2表示命令参数或文件读取失败。报告包含输入及映射哈希、标记分布、完整标识去重、错误和警告；不修改输入 DOCX 或映射文件。不加 `--report` 时仅向标准输出输出 JSON。

数字引文转换由编辑任务完成，脚本中的 `expand_numbers()` 仅提供严格范围展开，不会自动把所有方括号数字视作引文。转换前明确正文哪些位置是引文、哪些是公式/统计区间。脚本不负责证明正文文字或格式与另一份稿件完全一致，编辑者须另做范围外文字与结构比较。

## 转换操作和兼容性

ODF/DOCX Scan 是额外插件；“内置 RTF Scan”不是它的别名。先确认插件已启用、版本支持 DOCX，以及 Zotero Word 插件可用。DOCX 路线用 Microsoft Word；ODT 路线用 LibreOffice 和对应的 Zotero 插件。不要通过修改文件扩展名代替格式转换。

使用本项目的 Word 按钮时，在稿件内设置书目光标，核对目标分类，再明确确认引文转换。作者选择上游独立扫描器时，选择 **Markers → Zotero citations**、扫描稿和输出路径，然后在 Word 中选择样式，按需使用 **Refresh** 和 **Add/Edit Bibliography**。按钮名称随版本变化，以实际菜单为准。脚注、尾注、文本框及新旧引文混合的稿件按实际转换效果检查。

既有文献表可能含未在正文引用的背景文献，或者 Zotero Bibliography 设置了 uncited/omitted/custom 条目。不要一律断言“参考文献条数必须等于正文唯一文献数”；先解释配置差异。正常纯正文引文且无特殊配置时才检查相等。

## 与 Zotero 内置 RTF Scan 的区别

内置 RTF Scan 使用 `.rtf` 文件及 `{Smith, 2009}` 等作者年份标记，通过 Zotero 的匹配界面确认文献后排版。不得把 Scannable Cite 的五字段标记放进内置 RTF Scan，也不得把其格式化结果直接当作已验证可编辑的 Word Zotero 活动域。用户需要后续 Add/Edit Citation 的默认路线是本技能的 ODF/DOCX Scan。

## 一手来源

- [ODF/DOCX Scan 项目说明](https://zotero-odf-scan.github.io/zotero-odf-scan/)：语法、合引、定位和转换。
- [插件源代码与发行版](https://github.com/Juris-M/zotero-odf-scan-plugin)：核对版本和部件支持。
- [Zotero 内置 RTF Scan 文档](https://www.zotero.org/support/rtf_scan)：独立的 RTF 流程。

维护时以当前项目文档、已安装插件及真实样稿结果为准，不把某一次本机版本固定成通用前提。
