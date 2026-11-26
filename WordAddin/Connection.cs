using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;

namespace WordZoteroScan {
    public class CollectionRow {
        public int id, parent, library;
        public string name, key, libraryName;
    }
    public class FieldResult { public string instruction, display; }
    public class Reply {
        public string error;
        public CollectionRow[] collections;
        public FieldResult[] fields;
        public string[] uris;
        public int total, added;
    }
    internal static class Connection {
        internal static string LocalFolder {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordZoteroScan"); }
        }
        internal static JavaScriptSerializer Json() {
            return new JavaScriptSerializer { MaxJsonLength = Int32.MaxValue };
        }
        internal static Reply Send(string action, Dictionary<string, object> data) {
            string keyPath = Path.Combine(LocalFolder, "connection.key");
            if (!File.Exists(keyPath)) throw new InvalidOperationException("找不到本机连接配置，请运行安装程序。");
            data["action"] = action;
            data["auth"] = File.ReadAllText(keyPath).Trim();
            var req = (HttpWebRequest)WebRequest.Create("http://127.0.0.1:23119/word-zotero-scan/v1");
            req.Proxy = null;
            req.AllowAutoRedirect = false;
            req.Method = "POST";
            req.ContentType = "application/json";
            req.Timeout = 180000;
            req.ReadWriteTimeout = 180000;
            var bytes = Encoding.UTF8.GetBytes(Json().Serialize(data));
            req.ContentLength = bytes.Length;
            try {
                using (var stream = req.GetRequestStream()) stream.Write(bytes, 0, bytes.Length);
                using (var response = (HttpWebResponse)req.GetResponse())
                using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8)) {
                    return Json().Deserialize<Reply>(reader.ReadToEnd());
                }
            } catch (WebException e) {
                using (var response = e.Response as HttpWebResponse) {
                    if (response == null) throw new InvalidOperationException("请打开 Zotero 桌面端，再运行扫描。", e);
                    if (response.StatusCode == HttpStatusCode.NotFound)
                        throw new InvalidOperationException("请在 Zotero 中启用 Word Zotero Scan 扩展，随后重试。", e);
                    using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8)) {
                        string body = reader.ReadToEnd();
                        Reply detail = null;
                        try { detail = Json().Deserialize<Reply>(body); } catch (ArgumentException) {}
                        throw new InvalidOperationException(detail != null && detail.error != null ? detail.error :
                            "Zotero 返回 HTTP " + (int)response.StatusCode + "。", e);
                    }
                }
            }
        }
    }
}
