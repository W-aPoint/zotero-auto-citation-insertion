---
name: zotero-scannable-citations
description: 将已核实的文献或数字引文排成 Zotero Scannable Cite 五字段标记。先确定 Zotero 分类和稿件交付位置，并由作者在 Word 中确认转换。
---

# Zotero 扫描标记引文

先按 [分类归档与稿件标记流程](../zotero-word-workflow/SKILL.md) 确认个人库或群组库、分类完整路径和交付方式。作者可以选原稿、副本或仅返回文字。需要新文献时可配合 Consensus 检索，核实原论文后按指定分类导入。对已同步文库，可在桌面端关闭时通过授权的 Web API 归档。

按所支持的具体表述放置引文。已有数字引文先建立“数字序号 → 文献 → 真实 Zotero 标识”的映射。展开 `[3–5]` 等范围，区分数字引文、数据区间、公式和文后列表编号。来源不明确的数字保留原样，将缺项列在正文外。

保留五个字段、定位信息和原有顺序。连续合引标记紧贴排列。个人库和群组库分别使用其真实文库 ID 与条目键。DOI 或 PMID 可用于查证文献。具体标记语法和映射表示例见[语法与转换说明](references/workflow.md)。

选择原稿时，可按主流程直接在原 DOCX 对应句子写入纯文本标记。选择副本时使用用户指定的新路径。已有活动域、参考文献表、图表、书签和原句的其他内容保持原样。可对作者提供的映射和待扫描文档做静态检查：

```text
python scripts/validate_scan_docx.py manuscript.docx --mapping reference_map.json
```

此检查报告标记语法和映射的一致性。归档时另行回读核实真实条目及分类。交付后，作者在 Word 中审阅并明确确认转换，书目生成于作者所选的光标位置。
