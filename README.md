# Zotero 自动引文插入

**简体中文** · [English](README.en.md)

把 Zotero 的纯文本引文标记放在它所支持的句子旁。准备好稿件后，在 Word 里把光标放到参考文献表应出现的位置，点击一次按钮，选择 Zotero 分类。按钮在原句位置转换引文，把文中引用的条目加入所选分类，再调用 Zotero 原生的 **Add/Edit Bibliography**。

[下载 Windows 安装包](https://github.com/W-aPoint/zotero-auto-citation-insertion/releases/latest/download/Zotero-Auto-Citation-Insertion-Windows.zip) · [单独下载 Zotero 配套扩展](https://github.com/W-aPoint/zotero-auto-citation-insertion/releases/latest/download/WordZoteroScan.xpi)

## 安装 Word 按钮

需要 Windows 桌面版 Microsoft Word、Zotero 桌面端及其 Word 插件。已核验这套流程使用的 Zotero 版本为 10.0.8，ODF/DOCX Scan 版本为 3.2.2。Word 按钮和分类选择窗口使用中文。

1. 安装上游 [ODF/DOCX Scan](https://zotero-odf-scan.github.io/zotero-odf-scan/)。在 Zotero 中打开 **工具 → 插件**，将上游下载的 `.xpi` 拖入窗口。它提供 Scannable Cite 导出格式及 DOCX 转换器。
2. 下载并解压本项目的 Windows 安装包。保存手头的文档并关闭 Word，在解压目录打开 **Windows PowerShell**，运行 `.\Install.ps1`。它为当前 Windows 用户注册 Word 按钮，并创建 Word 和 Zotero 在本机通信的配对信息。
3. 在 Zotero 的 **工具 → 插件** 中安装压缩包里的 `dist/WordZoteroScan.xpi`。重新打开 Word，在 **引用 → Zotero** 中可以看到 **扫描并插入参考文献表**。

如安装程序无法识别 Word 位数，在 Word 的 **文件 → 账户 → 关于 Word** 查看位数，再运行 `.\Install.ps1 -WordArchitecture x86` 或 `.\Install.ps1 -WordArchitecture x64`。

## 从一句话到参考文献表

1. **取得条目的真实标记。** 在 Zotero 的 **设置 → 导出** 中将快速复制格式设为 **Scannable Cite**。选中文献，按 `Ctrl+Alt+C` 复制，把标记贴在所支持的句子旁。同处合引时，多个标记直接相连。下面只示意个人库标记的形状，实际写作时请用自己复制的整条标记。

   ```text
   { | <作者>, (<年份>) |p. 45 | |zu:<用户ID>:<条目键>}
   ```

2. **把光标放到书目的起始位置。** 例如放在“参考文献”标题下的空行。点击 Word **引用 → Zotero → 扫描并插入参考文献表**。
3. **选择归档分类。** 选择与这些条目属于同一个 Zotero 文库的可写分类，再点 **插入或编辑参考文献表**。按钮在每条标记原处创建可编辑引文，保留已有活动引文，并将文中引用的条目加入该分类。已归档的条目可重复运行。
4. **核对并保存。** 新文档会在原光标处生成书目。如果 Zotero 提示选择引文样式，请选择所需样式。文档已有活动书目时，按钮打开原书目的编辑窗口。书目内容来自文中引文，所选分类只负责归档。核对后按平常方式保存 Word 文档。

按钮支持正文（含表格）、脚注和尾注中的标记。如果标记位于页眉、页脚或文本框中，按钮会提示用户处理。

## 只需要在回复中准备引文标记

仓库提供两个可复制到技能目录的文件夹。日常写作用 [zotero-word-workflow](skills/zotero-word-workflow/SKILL.md)。如需将既有数字引文按已核实的映射排成标记，可用 [zotero-scannable-citations](skills/zotero-scannable-citations/SKILL.md)。

将 [skills](skills) 下的两个文件夹作为同级目录复制到所用助手的技能目录。Codex 通常使用 `~/.codex/skills/`。在使用 `~/.gemini/config/skills/` 的 AntiGravity 配置里可放到该目录。向助手提供正文和从 Zotero 复制的真实标记，例如：

> 用 zotero-word-workflow 把这两篇文献标在对应句子旁，把标好的段落发给我。Word 和 Zotero 由我操作。

助手在回复中交付可粘贴正文。你本人负责粘贴、选择分类和运行扫描。只使用 Skill 时，也可自行运行上游 **ODF Scan**，打开转换后的 DOCX，再在 Word 的 **Zotero** 选项卡把书目插到光标处。[ODF/DOCX Scan 使用说明](https://zotero-odf-scan.github.io/zotero-odf-scan/) · [Zotero 官方 Word 说明](https://www.zotero.org/support/word_processor_plugin_usage)

## 源码与许可

[Word 加载项](WordAddin) · [Zotero 配套扩展](zotero) · [构建脚本](Build.ps1)。从源码构建时，在安装了 .NET Framework 编译器及 Office 主互操作程序集的 Windows 电脑上运行 `.\Build.ps1`。输出位于 `dist/`。可用 Python 3 运行随附标记校验器的离线测试：

```sh
python skills/zotero-scannable-citations/scripts/test_validate_scan_docx.py
```

本项目以 [MIT 许可证](LICENSE)发布。引文转换由另行安装的上游 ODF/DOCX Scan 提供，书目排版由 Zotero 的 Word 插件完成。Word 按钮通过本机回环地址与 Zotero 通信，仅传递转换与归档所需的标记和条目标识。
