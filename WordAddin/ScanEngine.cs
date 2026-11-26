using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordZoteroScan {
    public static class ScanEngine {
        sealed class MarkerGroup {
            internal Word.Range Range;
            internal string Text;
        }
        static readonly Regex marker = new Regex(@"^\{[^{}|\r\n]*\|[^{}|\r\n]*\|[^{}|\r\n]*\|[^{}|\r\n]*\|\s*z[ug]:\d+:[A-Z0-9]{8}\s*\}\z");
        static readonly Regex active = new Regex(@"^\s*ADDIN\s+ZOTERO_ITEM\s+(?:CSL_CITATION\s+)?(?=\{)", RegexOptions.IgnoreCase);

        static Word.Range Find(Word.Range story, int start, int end, string text) {
            if (end <= start) return null;
            var range = story.Duplicate;
            range.SetRange(start, end);
            var find = range.Find;
            find.ClearFormatting();
            bool found = find.Execute(FindText:text, MatchCase:true, MatchWholeWord:false, MatchWildcards:false,
                MatchSoundsLike:false, MatchAllWordForms:false, Forward:true, Wrap:Word.WdFindWrap.wdFindStop, Format:false);
            // In a table cell Word can return an earlier hit even with wdFindStop.
            if (!found || range.Start < start || range.End > end || range.End <= range.Start) return null;
            return range;
        }
        static void ReadMarkers(Word.Range story, List<MarkerGroup> groups, bool supported) {
            if ((story.Text ?? "").IndexOf('{') < 0) return;
            foreach (Word.Paragraph paragraph in story.Paragraphs) {
                var part = paragraph.Range;
                if ((part.Text ?? "").IndexOf('{') >= 0) ReadParagraph(part, groups, supported);
            }
        }
        static void ReadParagraph(Word.Range story, List<MarkerGroup> groups, bool supported) {
            MarkerGroup previous = null;
            int end = story.End, cursor = story.Start;
            while (cursor < end) {
                var open = Find(story, cursor, end, "{");
                if (open == null) break;
                cursor = open.End;
                var close = Find(story, cursor, end, "}");
                if (close == null) {
                    var tail = story.Duplicate;
                    tail.SetRange(open.Start, end);
                    if ((tail.Text ?? "").Contains("|zu:") || (tail.Text ?? "").Contains("|zg:"))
                        throw new InvalidOperationException("有一条引文标记缺少右花括号。");
                    continue;
                }
                var range = story.Duplicate;
                range.SetRange(open.Start,close.End);
                string text = range.Text ?? "";
                if (!marker.IsMatch(text)) {
                    if (text.IndexOf('{',1) < 0 && (text.Count(c => c == '|') >= 4 ||
                        Regex.IsMatch(text, @"\|\s*z[ug]:"))) throw new InvalidOperationException("无法识别引文标记：" + text);
                    continue;
                }
                if (!supported) throw new InvalidOperationException("在页眉、页脚或文本框等 Word Zotero 不支持的区域发现引文标记：" + text);
                if ((bool)range.Information[Word.WdInformation.wdInFieldCode] ||
                    (bool)range.Information[Word.WdInformation.wdInFieldResult])
                    throw new InvalidOperationException("引文标记位于另一个 Word 域中：" + text);
                if (previous != null && previous.Range.End == range.Start) {
                    previous.Range.End = range.End;
                    previous.Text += text;
                } else {
                    previous = new MarkerGroup { Range = range, Text = text };
                    groups.Add(previous);
                }
                cursor = close.End;
            }
        }
        static void ReadCitations(Word.Range story, List<string[]> existing) {
            foreach (Word.Field field in story.Fields) {
                string code = field.Code.Text ?? "";
                var hit = active.Match(code);
                if (!hit.Success) continue;
                var citation = Connection.Json().Deserialize<Dictionary<string,object>>(code.Substring(hit.Length).Trim());
                object raw;
                if (!citation.TryGetValue("citationItems",out raw) || !(raw is IEnumerable))
                    throw new InvalidOperationException("已有 Zotero 引文缺少条目列表。");
                foreach (var part in (IEnumerable)raw) {
                    var item = part as Dictionary<string,object>;
                    object aliases;
                    if (item == null || (!item.TryGetValue("uris", out aliases) && !item.TryGetValue("uri", out aliases)))
                        throw new InvalidOperationException("已有引文缺少条目 URI，请用 Zotero 刷新该引文后重试。");
                    if (aliases is string) existing.Add(new[]{(string)aliases});
                    else if (aliases is IEnumerable) existing.Add(((IEnumerable)aliases).Cast<object>().Select(v=>Convert.ToString(v)).ToArray());
                    else throw new InvalidOperationException("已有引文的条目 URI 无法读取。");
                }
            }
        }

        public static Word.Range CaptureCaret(Word.Selection selection) {
            var caret = selection.Range.Duplicate;
            caret.Collapse(selection.StartIsActive ? Word.WdCollapseDirection.wdCollapseStart : Word.WdCollapseDirection.wdCollapseEnd);
            return caret;
        }

        // A Word bookmark follows edits in its own story, including changes in
        // field-code length ahead of the insertion point.
        public static string InsertBibliography(Word.Window window, Word.Range caret, int collectionID) {
            var doc = window.Document;
            if (doc.ReadOnly) throw new InvalidOperationException("请在可编辑的 Word 文档中插入参考文献表。");
            string name = "WZScanPos" + Guid.NewGuid().ToString("N").Substring(0,24);
            doc.Bookmarks.Add(name, caret);
            string status;
            try {
                status = ConvertAndFile(doc, collectionID);
                if (!doc.Bookmarks.Exists(name)) {
                    throw new InvalidOperationException("原光标位于已被替换的标记中。请将光标放到参考文献表起始位置，再运行按钮。");
                }
                var destination = doc.Bookmarks[name].Range.Duplicate;
                destination.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                doc.Bookmarks[name].Delete();
                window.Activate();
                destination.Select();
            } finally {
                if (doc.Bookmarks.Exists(name)) doc.Bookmarks[name].Delete();
            }
            try {
                // Zotero locates existing BIBL fields, reads the cited items and
                // formats the bibliography. A separate asynchronous Refresh would race it.
                doc.Application.Run("ZoteroAddEditBibliography");
            } catch (COMException e) {
                throw new InvalidOperationException(status + "\n未能启动 Zotero 的 Add/Edit Bibliography。请检查 Word 的 Zotero 选项卡。", e);
            }
            doc.Application.StatusBar = status + " 已调用 Zotero 的 Add/Edit Bibliography。";
            return status;
        }

        public static string ConvertAndFile(Word.Document doc, int collectionID) {
            var groups = new List<MarkerGroup>();
            var existing = new List<string[]>();
            foreach (Word.Range first in doc.StoryRanges) {
                for (var story = first; story != null; story = story.NextStoryRange) {
                    bool supported = story.StoryType == Word.WdStoryType.wdMainTextStory ||
                        story.StoryType == Word.WdStoryType.wdFootnotesStory || story.StoryType == Word.WdStoryType.wdEndnotesStory;
                    ReadMarkers(story, groups, supported);
                    ReadCitations(story, existing);
                }
            }
            if (groups.Count == 0 && existing.Count == 0) throw new InvalidOperationException("当前文档没有 Scannable Cite 标记或 Zotero 活动引文。");
            if (groups.Count != 0 && doc.ReadOnly) throw new InvalidOperationException("当前文档是只读的，请在可编辑副本中扫描。");
            var plan = Connection.Send("prepare", new Dictionary<string,object>{
                {"target",collectionID},{"groups",groups.Select(g=>g.Text).ToArray()},{"existing",existing}});
            if (plan.fields == null || plan.fields.Length != groups.Count || plan.uris == null)
                throw new InvalidOperationException("Zotero 返回的引文组数与文档不一致。");
            var app = doc.Application;
            if (groups.Count > 0) {
                int changed = 0;
                var undo = app.UndoRecord;
                undo.StartCustomRecord("Zotero 扫描引文");
                try {
                    for (int i=groups.Count-1; i>=0; --i) {
                        var range = groups[i].Range;
                        if (range.Text != groups[i].Text) throw new InvalidOperationException("扫描期间文档中的标记发生变化。");
                        var result = plan.fields[i];
                        if (!result.instruction.StartsWith("ADDIN ZOTERO_ITEM CSL_CITATION ",StringComparison.Ordinal))
                            throw new InvalidOperationException("ODF/DOCX Scan 返回了无法识别的 Word 引文字段。");
                        // Word puts text after a resultless ADDIN field. A QUOTE field
                        // supplies the result separator before we replace its code and text.
                        var field = range.Fields.Add(range, Word.WdFieldType.wdFieldQuote, "\"Zotero\"", false);
                        changed++;
                        field.Code.Text = " " + result.instruction + " ";
                        field.Result.Text = result.display;
                        if (field.Code.Text.Trim() != result.instruction || field.Result.Text != result.display)
                            throw new InvalidOperationException("Word 未完整保存引文字段或显示文字。");
                    }
                } catch (Exception) {
                    undo.EndCustomRecord();
                    if (changed != 0) doc.Undo(1);
                    throw;
                }
                undo.EndCustomRecord();
            }
            Reply filed;
            try {
                filed = Connection.Send("archive",new Dictionary<string,object>{{"target",collectionID},{"uris",plan.uris}});
            } catch (Exception e) {
                throw new InvalidOperationException("归档未确认。再次运行可用文中现有活动引文补齐分类。\n" + e.Message,e);
            }
            string status = String.Format("已转换 {0} 组引文；所选分类收录 {1} 篇，其中本次加入 {2} 篇。",groups.Count,filed.total,filed.added);
            return status;
        }
    }
}
