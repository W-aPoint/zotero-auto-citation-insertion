using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

[assembly: ComVisible(false)]
[assembly: System.Reflection.AssemblyVersion("1.2.0.0")]
namespace WordZoteroScan {
    [ComVisible(true), Guid("7AF39E18-5D4B-4ABE-8E90-661588B28E11")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonActions { [DispId(1)] void OnScan(Office.IRibbonControl control); }

    [ComVisible(true), Guid("2077C602-306B-4F7A-9FA3-EB45C7F1AF34")]
    [ProgId("WordZoteroScan.Addin"), ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonActions))]
    public class Addin : IDTExtensibility2, Office.IRibbonExtensibility, IRibbonActions {
        Word.Application app;
        bool busy;
        public void OnConnection(object application, ext_ConnectMode mode, object addin, ref Array custom) {
            app = (Word.Application)application;
            ((Office.COMAddIn)addin).Object = this;
        }
        public void OnDisconnection(ext_DisconnectMode mode, ref Array custom) { app = null; }
        public void OnAddInsUpdate(ref Array custom) {}
        public void OnStartupComplete(ref Array custom) {}
        public void OnBeginShutdown(ref Array custom) {}

        public string GetCustomUI(string ribbonID) {
            return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs>" +
                "<tab idMso='TabReferences'><group id='WordZoteroScanGroup' label='Zotero'>" +
                "<button id='WordZoteroScanButton' label='扫描并插入参考文献表' size='large' imageMso='BibliographyInsert' " +
                "screentip='在光标处插入或编辑 Zotero 参考文献表' supertip='原位转换已有标记，归档文中文献，然后在原光标处运行 Zotero 的 Add/Edit Bibliography。' onAction='OnScan'/>" +
                "</group></tab></tabs></ribbon></customUI>";
        }
        public void OnScan(Office.IRibbonControl control) {
            if (busy || app == null) return;
            busy = true;
            try {
                if (app.Documents.Count == 0) throw new InvalidOperationException("请先在 Word 中打开要扫描的文档。");
                var window = app.ActiveWindow;
                var caret = ScanEngine.CaptureCaret(window.Selection);
                var available = Connection.Send("collections", new Dictionary<string,object>()).collections;
                var owner = new NativeWindow();
                owner.AssignHandle(new IntPtr(window.Hwnd));
                try {
                    using (var picker = new CollectionPicker(available)) {
                        if (picker.ShowDialog(owner) != DialogResult.OK) return;
                        ScanEngine.InsertBibliography(window, caret, picker.Selected.id);
                    }
                } finally { owner.ReleaseHandle(); }
            } catch (Exception e) {
                MessageBox.Show(e.Message, "Zotero 参考文献表", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } finally { busy = false; }
        }
    }

    internal sealed class CollectionPicker : Form {
        readonly TreeView tree = new TreeView { Dock = DockStyle.Fill, HideSelection = false, ShowLines = true };
        readonly Button submit = new Button { Text = "确认转换引文并打开参考文献表", AutoSize = true, Enabled = false };
        internal CollectionRow Selected { get { return tree.SelectedNode == null ? null : tree.SelectedNode.Tag as CollectionRow; } }
        internal CollectionPicker(CollectionRow[] rows) {
            Text = "核对目标 Zotero 分类并确认引文转换";
            Font = SystemFonts.MessageBoxFont;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(435, 410);
            Padding = new Padding(12);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 43, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0,7,0,0) };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true };
            buttons.Controls.Add(cancel); buttons.Controls.Add(submit);
            Controls.Add(tree); Controls.Add(buttons);
            CancelButton = cancel;
            string last = null;
            string pref = Path.Combine(Connection.LocalFolder, "last-collection.txt");
            try { if (File.Exists(pref)) last = File.ReadAllText(pref).Trim(); } catch (IOException) {}
            var libraries = new Dictionary<int,TreeNode>();
            var nodes = new Dictionary<int,TreeNode>();
            TreeNode previous = null;
            foreach (var row in rows) {
                if (!libraries.ContainsKey(row.library)) libraries[row.library] = tree.Nodes.Add(row.libraryName);
                var node = new TreeNode(row.name) { Tag = row };
                nodes.Add(row.id, node);
                if (row.library + "/" + row.key == last) previous = node;
            }
            foreach (var row in rows) {
                var parent = row.parent != 0 && nodes.ContainsKey(row.parent) ? nodes[row.parent] : libraries[row.library];
                parent.Nodes.Add(nodes[row.id]);
            }
            tree.AfterSelect += delegate { submit.Enabled = Selected != null; };
            submit.Click += delegate {
                if (Selected == null) return;
                try { File.WriteAllText(pref, Selected.library + "/" + Selected.key); } catch (IOException) {}
                DialogResult = DialogResult.OK;
            };
            foreach (TreeNode lib in tree.Nodes) lib.Expand();
            if (previous != null) { tree.SelectedNode = previous; previous.EnsureVisible(); }
        }
    }
}
