# Zotero 自动引文插入

**简体中文** · [English](README.en.md)

先告诉 Skill 论文要归入哪个 Zotero 分类，以及引文标记要写到哪里。Skill 可以在 Zotero 桌面端关闭时把书目记录导入指定分类，并在你选择的 Word 原稿或副本里放好五字段标记。**真正转换成活动引文，以及在光标处生成参考文献表，由你在 Word 里亲自确认。**

[下载 Windows 完整包](https://github.com/W-aPoint/zotero-auto-citation-insertion/releases/latest/download/Zotero-Auto-Citation-Insertion-Windows.zip) · [单独下载 Zotero 配套扩展](https://github.com/W-aPoint/zotero-auto-citation-insertion/releases/latest/download/WordZoteroScan.xpi)

## 怎样给 Skill 下任务

> 目标是个人库「我的课题 → 参考文献」。请配合 Consensus 给引言第二段查找并核实文献，归入这个分类。把五字段标记直接写进我提供的 Word 原稿。最后我自己在 Word 里确认转换。

每次任务先确定两件事。**目标分类**写明个人库或群组库，以及分类完整路径。同名分类可用父级路径或分类键区分。**稿件处理方式**可以选下面任何一种。已经说明的选择会直接沿用。

| 你的选择 | 收到的结果 |
| --- | --- |
| 直接写入原稿 | 指定 DOCX 中的对应句子加入纯文本标记，交付路径就是原路径。 |
| 另存副本 | 在你指定的新 DOCX 路径写入标记，原稿保留。 |
| 只返回文字 | 在回复里给出标好的原句或段落，由你自行粘贴。 |

## 归档、标注与确认

1. **查证论文。** 推荐在助手中启用 **Consensus 插件**（如可用，可直接提及 `@Consensus`）寻找支持某句话的候选研究，再核对原论文或 PubMed、出版社记录中的题名、作者、年份、DOI 或 PMID。已有文献可以直接核对。[Consensus](https://consensus.app/)
2. **归入指定分类。** 对已同步的个人库或群组库，Skill 可通过 Zotero Web API 复用已有条目，或导入缺失的文献并加入目标分类。它回读分类成员关系和真实条目键后再生成标记。此步需要该文库的云端读写授权，可以关闭桌面 Zotero。若文库仅存于本机，或未授权云端写入，可在桌面端手动归档后提供已核验的标记。[归档命令与凭据配置](skills/zotero-word-workflow/references/preparation.md) · [Zotero 官方写入接口](https://www.zotero.org/support/dev/web_api/v3/write_requests)
3. **按选择写入标记。** 原稿和副本使用同一套五字段纯文本格式。同处多篇文献的标记紧邻排列。文件级写入前，请保存并关闭该 Word 文档。选择“只返回文字”时可以继续编辑。下面示意带页码的个人库标记，实际条目键在归档时取得。

   ```text
   { | <作者>, (<年份>) |p. 45 | |zu:<用户ID>:<条目键>}
   ```

4. **作者在 Word 确认。** 打开 Zotero，等刚归档的条目同步到本机。打开已标记的稿件，将光标放到书目应开始的位置，点 Word **引用 → Zotero → 扫描并插入参考文献表**，核对与前面相同的分类，然后亲自点 **确认转换引文并打开参考文献表**。仅选择分类或取消都不会转换引文。确认后，标记在原句处转为活动引文，原生 **Add/Edit Bibliography** 在原光标处生成书目。已有活动书目时，Zotero 打开原书目的编辑窗口。检查结果后保存。[Zotero 同步说明](https://www.zotero.org/support/sync) · [Word 插件说明](https://www.zotero.org/support/word_processor_plugin_usage)

所选分类用于收纳文中引用的文献。书目由稿件中的引文决定。按钮支持正文（含表格）、脚注和尾注，遇到页眉、页脚或文本框中的标记会提示处理。

## 安装

**Skill。** 把 [skills](skills) 下的两个文件夹作为同级目录，复制到 Codex 的 `~/.codex/skills/`，或你配置的 AntiGravity 技能目录，例如 `~/.gemini/config/skills/`。[zotero-word-workflow](skills/zotero-word-workflow/SKILL.md) 负责归档和标注，[zotero-scannable-citations](skills/zotero-scannable-citations/SKILL.md) 也提供数字引文映射和只读校验。云端助手程序使用 Python 3.10+ 标准库，直接标注 DOCX 还需安装 `lxml`。将云端密钥保存在本机环境变量 `ZOTERO_API_KEY` 中。[命令和输入示例](skills/zotero-word-workflow/references/preparation.md)

**Word 按钮（Windows 桌面版 Word）。**

1. 安装 Zotero 桌面端、Zotero 的 Word 插件和上游 [ODF/DOCX Scan](https://zotero-odf-scan.github.io/zotero-odf-scan/)。在 Zotero 的 **工具 → 插件** 窗口安装上游 `.xpi`。其转换器和 Scannable Cite 导出格式由上游提供。
2. 解压本项目的 Windows 完整包。保存并关闭 Word，在解压目录用 Windows PowerShell 运行 `.\Install.ps1`。程序为当前 Windows 用户注册 Word 按钮。如果无法自动识别 Word 位数，可按 **文件 → 账户 → 关于 Word** 显示的位数传入 `-WordArchitecture x86` 或 `-WordArchitecture x64`。
3. 在 Zotero 的 **工具 → 插件** 中安装包内的 `dist/WordZoteroScan.xpi`，重新打开 Word。按钮位于 **引用 → Zotero**。已有工作流核验所用的版本为 Zotero 10.0.8 和 ODF/DOCX Scan 3.2.2。

也可由作者使用上游 **ODF Scan** 手动扫描文件，再打开输出的 DOCX，在 Word 的 Zotero 选项卡点击 **Add/Edit Bibliography**。[上游手工扫描说明](https://zotero-odf-scan.github.io/zotero-odf-scan/)

## 源码与许可

[Word 加载项](WordAddin) · [Zotero 扩展](zotero) · [测试](tests) · [MIT 许可证](LICENSE)。从源码构建时，在有 .NET Framework 编译器和 Office 主互操作程序集的 Windows 电脑上运行 `.\Build.ps1`，输出位于 `dist/`。
