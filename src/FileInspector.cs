// UnLocked - Office 檔案結構檢查器（不需開啟 Office）
// 注意：此檔案由 Windows PowerShell 5.1 的 Add-Type 編譯，必須維持 C# 5 語法
//       （不可使用 $"" 字串插值、?. 運算子、運算式主體成員、out var 等）。
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace UnLocked
{
    public class InspectResult
    {
        public string Path = "";
        /// <summary>OOXML / CFB / UNKNOWN / EMPTY / MISSING</summary>
        public string Container = "UNKNOWN";
        /// <summary>OK / CORRUPT / PROTECTED / PASSWORD / UNKNOWN / ERROR</summary>
        public string Status = "UNKNOWN";
        /// <summary>Excel / Word / PowerPoint / 空字串</summary>
        public string DocKind = "";
        public bool IrmProtected;
        public bool PasswordEncrypted;
        public string MainPart = "";
        public string LabelId = "";
        public string LabelName = "";
        public string LabelSiteId = "";
        public string LabelMethod = "";
        public int PartCount;
        /// <summary>標籤的 contentBits（-1 = 未知）；8 = 標籤含加密</summary>
        public int LabelContentBits = -1;
        /// <summary>保護名稱（標籤／範本名稱，取自 IRM 發佈授權）</summary>
        public string ProtectionName = "";
        /// <summary>權限伺服器主機名稱</summary>
        public string RmsServer = "";
        /// <summary>Azure RMS / AD RMS</summary>
        public string RmsKind = "";
        public bool LabelEncrypts { get { return LabelContentBits >= 0 && (LabelContentBits & 8) != 0; } }
        public List<string> Issues = new List<string>();
        public List<string> Warnings = new List<string>();

        public bool IsCorrupt { get { return Status == "CORRUPT"; } }
        public string IssueText { get { return string.Join("；", Issues.ToArray()); } }
        public string WarningText { get { return string.Join("；", Warnings.ToArray()); } }
    }

    public static class FileInspector
    {
        const int MaxMessages = 30;
        const uint ENDOFCHAIN = 0xFFFFFFFE;
        const uint FREESECT = 0xFFFFFFFF;

        static readonly Regex GuidRegex = new Regex(
            "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}");
        static readonly Regex PlNameRegex = new Regex(":NAME (.+?):DESCRIPTION");
        static readonly Regex RmsUrlRegex = new Regex("https?://([^/<\"\\s]+)/_wmcs", RegexOptions.IgnoreCase);
        static readonly Regex MsipRegex = new Regex(
            "^MSIP_Label_([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})_(\\w+)$");

        public static InspectResult Inspect(string path)
        {
            InspectResult r = new InspectResult();
            r.Path = path;
            try
            {
                if (!File.Exists(path))
                {
                    r.Container = "MISSING";
                    r.Status = "ERROR";
                    r.Issues.Add("找不到檔案");
                    return r;
                }
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read,
                                                      FileShare.ReadWrite | FileShare.Delete))
                {
                    if (fs.Length == 0)
                    {
                        r.Container = "EMPTY";
                        r.Status = "CORRUPT";
                        r.Issues.Add("檔案大小為 0 位元組");
                        return r;
                    }
                    byte[] head = new byte[8];
                    int n = ReadFully(fs, head, 0, 8);
                    if (n >= 4 && head[0] == 0x50 && head[1] == 0x4B)
                    {
                        r.Container = "OOXML";
                        InspectZip(fs, r);
                    }
                    else if (n == 8 && head[0] == 0xD0 && head[1] == 0xCF && head[2] == 0x11 && head[3] == 0xE0 &&
                             head[4] == 0xA1 && head[5] == 0xB1 && head[6] == 0x1A && head[7] == 0xE1)
                    {
                        r.Container = "CFB";
                        InspectCfb(fs, r);
                    }
                    else
                    {
                        r.Container = "UNKNOWN";
                        r.Status = "UNKNOWN";
                        r.Warnings.Add("無法辨識的檔案格式（可能經第三方 DRM 加密，或並非 Office 檔），將交由 Office 判斷");
                        return r;
                    }
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                r.Status = "ERROR";
                r.Issues.Add("沒有讀取權限：" + ex.Message);
                return r;
            }
            catch (IOException ex)
            {
                r.Status = "ERROR";
                r.Issues.Add("無法讀取檔案（可能正被其他程式使用）：" + ex.Message);
                return r;
            }
            catch (Exception ex)
            {
                r.Issues.Add("檢查時發生未預期的錯誤：" + ex.Message);
            }

            if (r.Issues.Count > 0) r.Status = "CORRUPT";
            else if (r.PasswordEncrypted) r.Status = "PASSWORD";
            else if (r.IrmProtected) r.Status = "PROTECTED";
            else r.Status = "OK";
            return r;
        }

        // ------------------------------------------------------------------
        //  ZIP / OPC (xlsx, docx, pptx ...)
        // ------------------------------------------------------------------
        class ZipEntryInfo
        {
            public string Name;
            public int Flags;
            public int Method;
            public uint Crc;
            public long CSize;
            public long USize;
            public long LocalOffset;
        }

        static void InspectZip(FileStream fs, InspectResult r)
        {
            long len = fs.Length;
            int tailLen = (int)Math.Min(len, 65557L);
            byte[] tail = ReadAt(fs, len - tailLen, tailLen);
            int eocd = -1;
            for (int i = tailLen - 22; i >= 0; i--)
            {
                if (tail[i] == 0x50 && tail[i + 1] == 0x4B && tail[i + 2] == 0x05 && tail[i + 3] == 0x06)
                {
                    int commentLen = U16(tail, i + 20);
                    if (i + 22 + commentLen <= tailLen) { eocd = i; break; }
                }
            }
            if (eocd < 0)
            {
                r.Issues.Add("找不到 ZIP 中央目錄（檔案可能被截斷或寫入不完整）");
                return;
            }
            if (eocd + 22 + U16(tail, eocd + 20) < tailLen)
                r.Warnings.Add("ZIP 結尾含有多餘資料");

            long eocdPos = len - tailLen + eocd;
            long total = U16(tail, eocd + 10);
            long cdSize = U32(tail, eocd + 12);
            long cdOffset = U32(tail, eocd + 16);
            long cdEnd = eocdPos;

            if (cdOffset == 0xFFFFFFFFL || cdSize == 0xFFFFFFFFL || total == 0xFFFF)
            {
                byte[] loc = eocdPos >= 20 ? ReadAt(fs, eocdPos - 20, 20) : null;
                if (loc == null || U32(loc, 0) != 0x07064b50)
                {
                    r.Issues.Add("ZIP64 目錄定位器遺失");
                    return;
                }
                long z64Pos = (long)U64(loc, 8);
                byte[] z = (z64Pos >= 0 && z64Pos + 56 <= len) ? ReadAt(fs, z64Pos, 56) : null;
                if (z == null || U32(z, 0) != 0x06064b50)
                {
                    r.Issues.Add("ZIP64 中央目錄結尾損毀");
                    return;
                }
                total = (long)U64(z, 32);
                cdSize = (long)U64(z, 40);
                cdOffset = (long)U64(z, 48);
                cdEnd = z64Pos;
            }

            if (cdOffset < 0 || cdSize < 0 || cdOffset + cdSize > cdEnd)
            {
                r.Issues.Add("ZIP 中央目錄位置錯誤（檔案可能被截斷）");
                return;
            }
            if (cdSize > 64L * 1024 * 1024)
            {
                r.Issues.Add("ZIP 中央目錄大小異常");
                return;
            }

            byte[] cd = ReadAt(fs, cdOffset, (int)cdSize);
            List<ZipEntryInfo> entries = new List<ZipEntryInfo>();
            int p = 0;
            while (p + 46 <= cd.Length && U32(cd, p) == 0x02014b50)
            {
                ZipEntryInfo e = new ZipEntryInfo();
                e.Flags = U16(cd, p + 8);
                e.Method = U16(cd, p + 10);
                e.Crc = U32(cd, p + 16);
                e.CSize = U32(cd, p + 20);
                e.USize = U32(cd, p + 24);
                int nl = U16(cd, p + 28), xl = U16(cd, p + 30), cl = U16(cd, p + 32);
                e.LocalOffset = U32(cd, p + 42);
                if (p + 46 + nl + xl + cl > cd.Length)
                {
                    AddIssue(r, "ZIP 中央目錄項目不完整");
                    break;
                }
                e.Name = Encoding.UTF8.GetString(cd, p + 46, nl);
                if (e.USize == 0xFFFFFFFFL || e.CSize == 0xFFFFFFFFL || e.LocalOffset == 0xFFFFFFFFL)
                    ParseZip64Extra(cd, p + 46 + nl, xl, e);
                entries.Add(e);
                p += 46 + nl + xl + cl;
            }
            if (entries.Count != total)
                AddIssue(r, "ZIP 中央目錄項目數不符（宣告 " + total + "，實際讀到 " + entries.Count + "）");

            Dictionary<string, bool> names = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, byte[]> keep = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            foreach (ZipEntryInfo e in entries)
            {
                if (names.ContainsKey(e.Name))
                {
                    AddIssue(r, "封裝中有重複的部分：" + e.Name);
                    continue;
                }
                names[e.Name] = true;
                if (e.Name.EndsWith("/")) continue;
                r.PartCount++;
                CheckZipEntry(fs, e, r, keep);
            }

            CheckOpc(names, keep, r);
            ExtractLabelFromOoxml(keep, r);
        }

        static void ParseZip64Extra(byte[] b, int start, int length, ZipEntryInfo e)
        {
            int p = start, end = start + length;
            while (p + 4 <= end)
            {
                int id = U16(b, p), size = U16(b, p + 2);
                if (id == 0x0001)
                {
                    int q = p + 4;
                    if (e.USize == 0xFFFFFFFFL && q + 8 <= end) { e.USize = (long)U64(b, q); q += 8; }
                    if (e.CSize == 0xFFFFFFFFL && q + 8 <= end) { e.CSize = (long)U64(b, q); q += 8; }
                    if (e.LocalOffset == 0xFFFFFFFFL && q + 8 <= end) { e.LocalOffset = (long)U64(b, q); }
                    return;
                }
                p += 4 + size;
            }
        }

        static void CheckZipEntry(FileStream fs, ZipEntryInfo e, InspectResult r, Dictionary<string, byte[]> keep)
        {
            byte[] lh = (e.LocalOffset >= 0 && e.LocalOffset + 30 <= fs.Length) ? ReadAt(fs, e.LocalOffset, 30) : null;
            if (lh == null || U32(lh, 0) != 0x04034b50)
            {
                AddIssue(r, "部分「" + e.Name + "」的檔頭遺失或損毀");
                return;
            }
            long dataPos = e.LocalOffset + 30 + U16(lh, 26) + U16(lh, 28);
            if ((e.Flags & 1) != 0)
            {
                AddIssue(r, "部分「" + e.Name + "」使用 ZIP 加密（Office 不支援）");
                return;
            }
            if (e.Method != 0 && e.Method != 8)
            {
                AddIssue(r, "部分「" + e.Name + "」使用不支援的壓縮方式 (" + e.Method + ")");
                return;
            }
            if (dataPos + e.CSize > fs.Length)
            {
                AddIssue(r, "部分「" + e.Name + "」的資料被截斷");
                return;
            }
            if (e.CSize > int.MaxValue)
            {
                r.Warnings.Add("部分「" + e.Name + "」過大，略過內容檢查");
                return;
            }

            byte[] comp = ReadAt(fs, dataPos, (int)e.CSize);
            Stream src = new MemoryStream(comp, false);
            Stream data = e.Method == 8 ? (Stream)new DeflateStream(src, CompressionMode.Decompress) : src;
            CrcStream cs = new CrcStream(data, e.USize + 1);
            bool isXml = IsXmlPart(e.Name);
            try
            {
                try
                {
                    if (isXml && IsKeepPart(e.Name))
                    {
                        byte[] bytes = ReadAll(cs);
                        keep[e.Name] = bytes;
                        XmlDocument doc = new XmlDocument();
                        doc.XmlResolver = null;
                        using (XmlReader xr = XmlReader.Create(new MemoryStream(bytes, false), XmlSettings()))
                            doc.Load(xr);
                    }
                    else if (isXml)
                    {
                        using (XmlReader xr = XmlReader.Create(cs, XmlSettings()))
                            while (xr.Read()) { }
                    }
                }
                catch (XmlException ex)
                {
                    AddIssue(r, "部分「" + e.Name + "」XML 格式錯誤（第 " + ex.LineNumber + " 行，第 " +
                                ex.LinePosition + " 欄）：" + ex.Message);
                }
                Drain(cs);
            }
            catch (InvalidDataException)
            {
                AddIssue(r, "部分「" + e.Name + "」解壓縮失敗（資料損毀）");
                return;
            }
            catch (Exception ex)
            {
                AddIssue(r, "部分「" + e.Name + "」無法讀取：" + ex.Message);
                return;
            }
            finally
            {
                data.Dispose();
            }

            if (cs.Count != e.USize)
                AddIssue(r, "部分「" + e.Name + "」大小不符（宣告 " + e.USize + "，實際 " +
                            (cs.Count > e.USize ? "超過" : cs.Count.ToString()) + "）");
            else if (cs.Crc != e.Crc)
                AddIssue(r, "部分「" + e.Name + "」CRC 校驗失敗（內容已損毀）");
        }

        static void CheckOpc(Dictionary<string, bool> names, Dictionary<string, byte[]> keep, InspectResult r)
        {
            byte[] ct;
            if (!keep.TryGetValue("[Content_Types].xml", out ct))
            {
                if (!names.ContainsKey("[Content_Types].xml"))
                    AddIssue(r, "缺少 [Content_Types].xml（不是有效的 Office 檔案）");
            }
            else
            {
                Dictionary<string, bool> defaults = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, bool> overrides = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                XmlDocument d = TryLoad(ct);
                if (d != null)
                {
                    foreach (XmlNode node in d.GetElementsByTagName("*"))
                    {
                        XmlElement el = node as XmlElement;
                        if (el == null) continue;
                        if (el.LocalName == "Default") defaults[el.GetAttribute("Extension")] = true;
                        else if (el.LocalName == "Override") overrides[el.GetAttribute("PartName").TrimStart('/')] = true;
                    }
                    int missing = 0;
                    foreach (string name in names.Keys)
                    {
                        if (name.EndsWith("/") || name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)) continue;
                        string ext = "";
                        int dot = name.LastIndexOf('.');
                        if (dot >= 0 && dot > name.LastIndexOf('/')) ext = name.Substring(dot + 1);
                        if (!defaults.ContainsKey(ext) && !overrides.ContainsKey(name))
                        {
                            if (missing++ < 5) AddIssue(r, "部分「" + name + "」沒有宣告內容類型");
                        }
                    }
                    foreach (string o in overrides.Keys)
                        if (!names.ContainsKey(o)) r.Warnings.Add("內容類型宣告了不存在的部分：" + o);
                }
            }

            byte[] rootRels;
            if (!keep.TryGetValue("_rels/.rels", out rootRels))
            {
                AddIssue(r, "缺少 _rels/.rels（封裝關聯遺失）");
            }

            foreach (KeyValuePair<string, byte[]> kv in keep)
            {
                if (!kv.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
                XmlDocument d = TryLoad(kv.Value);
                if (d == null) continue;

                // xl/_rels/workbook.xml.rels -> 來源部分 xl/workbook.xml，基準目錄 xl/
                string relsPath = kv.Key;
                int idx = relsPath.LastIndexOf("_rels/", StringComparison.OrdinalIgnoreCase);
                if (idx < 0) continue;
                string baseDir = relsPath.Substring(0, idx);
                string sourcePart = baseDir + Path.GetFileNameWithoutExtension(relsPath.Substring(idx + 6));
                bool isRoot = relsPath.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase);
                if (!isRoot && !names.ContainsKey(sourcePart))
                    r.Warnings.Add("關聯檔「" + relsPath + "」的來源部分不存在");

                foreach (XmlNode node in d.GetElementsByTagName("*"))
                {
                    XmlElement el = node as XmlElement;
                    if (el == null || el.LocalName != "Relationship") continue;
                    if (el.GetAttribute("TargetMode").Equals("External", StringComparison.OrdinalIgnoreCase)) continue;
                    string target = el.GetAttribute("Target");
                    string type = el.GetAttribute("Type");
                    string resolved = ResolvePartName(baseDir, target);
                    bool exists = resolved != null && names.ContainsKey(resolved);
                    if (isRoot && type.EndsWith("/officeDocument", StringComparison.OrdinalIgnoreCase))
                    {
                        if (exists) r.MainPart = resolved;
                        else AddIssue(r, "找不到主要文件部分「" + target + "」");
                        continue;
                    }
                    if (!exists)
                        AddIssue(r, "「" + relsPath + "」參照了不存在的部分「" + target + "」");
                }
            }

            if (rootRels != null && r.MainPart.Length == 0 && r.Issues.Count == 0)
                AddIssue(r, "_rels/.rels 中沒有主要文件關聯");

            string probe = r.MainPart.Length > 0 ? r.MainPart : "";
            if (probe.Length == 0)
            {
                foreach (string n in names.Keys) { probe = n; if (n.StartsWith("xl/") || n.StartsWith("word/") || n.StartsWith("ppt/")) break; }
            }
            if (probe.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)) r.DocKind = "Excel";
            else if (probe.StartsWith("word/", StringComparison.OrdinalIgnoreCase)) r.DocKind = "Word";
            else if (probe.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase)) r.DocKind = "PowerPoint";
        }

        static string ResolvePartName(string baseDir, string target)
        {
            if (string.IsNullOrEmpty(target)) return null;
            int hash = target.IndexOf('#');
            if (hash >= 0) target = target.Substring(0, hash);
            try { target = Uri.UnescapeDataString(target); } catch (Exception) { }
            target = target.Replace('\\', '/');
            string combined = target.StartsWith("/") ? target.Substring(1) : baseDir + target;
            List<string> parts = new List<string>();
            foreach (string seg in combined.Split('/'))
            {
                if (seg.Length == 0 || seg == ".") continue;
                if (seg == "..")
                {
                    if (parts.Count == 0) return null;
                    parts.RemoveAt(parts.Count - 1);
                }
                else parts.Add(seg);
            }
            return string.Join("/", parts.ToArray());
        }

        static void ExtractLabelFromOoxml(Dictionary<string, byte[]> keep, InspectResult r)
        {
            // 新版格式：docMetadata/LabelInfo.xml
            foreach (KeyValuePair<string, byte[]> kv in keep)
            {
                if (!kv.Key.EndsWith("LabelInfo.xml", StringComparison.OrdinalIgnoreCase)) continue;
                if (ParseLabelList(kv.Value, r)) break;
            }

            // 舊版格式：docProps/custom.xml 的 MSIP_Label_<GUID>_<欄位>
            byte[] custom;
            if (!keep.TryGetValue("docProps/custom.xml", out custom)) return;
            XmlDocument c = TryLoad(custom);
            if (c == null) return;
            Dictionary<string, Dictionary<string, string>> labels = new Dictionary<string, Dictionary<string, string>>();
            foreach (XmlNode node in c.GetElementsByTagName("*"))
            {
                XmlElement el = node as XmlElement;
                if (el == null || el.LocalName != "property") continue;
                Match m = MsipRegex.Match(el.GetAttribute("name"));
                if (!m.Success) continue;
                string id = m.Groups[1].Value.ToLowerInvariant();
                if (!labels.ContainsKey(id)) labels[id] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                labels[id][m.Groups[2].Value] = el.InnerText.Trim();
            }
            foreach (KeyValuePair<string, Dictionary<string, string>> kv in labels)
            {
                Dictionary<string, string> f = kv.Value;
                string enabled;
                if (f.TryGetValue("Enabled", out enabled) && !enabled.Equals("true", StringComparison.OrdinalIgnoreCase)) continue;
                if (r.LabelId.Length > 0 && r.LabelId != kv.Key) continue;
                r.LabelId = kv.Key;
                string v;
                if (f.TryGetValue("Name", out v)) r.LabelName = v;
                if (r.LabelSiteId.Length == 0 && f.TryGetValue("SiteId", out v)) r.LabelSiteId = v.Trim('{', '}').ToLowerInvariant();
                if (r.LabelMethod.Length == 0 && f.TryGetValue("Method", out v)) r.LabelMethod = v;
                break;
            }
        }

        /// <summary>解析 clbl:labelList（docMetadata/LabelInfo.xml 或 IRM 容器內的 LabelInfo 資料流）。</summary>
        static bool ParseLabelList(byte[] xml, InspectResult r)
        {
            XmlDocument d = TryLoad(xml);
            if (d == null) return false;
            foreach (XmlNode node in d.GetElementsByTagName("*"))
            {
                XmlElement el = node as XmlElement;
                if (el == null || el.LocalName != "label") continue;
                if (el.GetAttribute("removed") == "1") continue;
                if (el.GetAttribute("enabled") == "0") continue;
                Match m = GuidRegex.Match(el.GetAttribute("id"));
                if (!m.Success) continue;
                r.LabelId = m.Value.ToLowerInvariant();
                Match s = GuidRegex.Match(el.GetAttribute("siteId"));
                if (s.Success) r.LabelSiteId = s.Value.ToLowerInvariant();
                r.LabelMethod = el.GetAttribute("method");
                int bits;
                if (int.TryParse(el.GetAttribute("contentBits"), out bits)) r.LabelContentBits = bits;
                return true;
            }
            return false;
        }

        /// <summary>從 IRM 發佈授權 (XrML) 取出保護名稱與權限伺服器。只讀取未加密的中繼資料。</summary>
        static void ParsePublishingLicense(byte[] b, InspectResult r)
        {
            string t16 = Encoding.Unicode.GetString(b);
            string t8 = Encoding.UTF8.GetString(b);
            string t = CountChar(t16, '<') > CountChar(t8, '<') ? t16 : t8;
            Match m = PlNameRegex.Match(t);
            if (m.Success) r.ProtectionName = m.Groups[1].Value.Trim();
            Match u = RmsUrlRegex.Match(t);
            if (u.Success)
            {
                r.RmsServer = u.Groups[1].Value.ToLowerInvariant();
                r.RmsKind = r.RmsServer.EndsWith("aadrm.com", StringComparison.OrdinalIgnoreCase) ? "Azure RMS" : "AD RMS";
            }
        }

        static int CountChar(string s, char c)
        {
            int n = 0;
            foreach (char ch in s) if (ch == c) n++;
            return n;
        }

        static bool IsXmlPart(string name)
        {
            return name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                   name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ||
                   name.EndsWith(".vml", StringComparison.OrdinalIgnoreCase);
        }

        static bool IsKeepPart(string name)
        {
            return name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ||
                   name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase) ||
                   name.Equals("docProps/custom.xml", StringComparison.OrdinalIgnoreCase) ||
                   name.EndsWith("LabelInfo.xml", StringComparison.OrdinalIgnoreCase);
        }

        static XmlReaderSettings XmlSettings()
        {
            XmlReaderSettings s = new XmlReaderSettings();
            s.DtdProcessing = DtdProcessing.Prohibit;
            s.XmlResolver = null;
            s.CheckCharacters = true;
            s.CloseInput = false;
            return s;
        }

        static XmlDocument TryLoad(byte[] bytes)
        {
            try
            {
                XmlDocument d = new XmlDocument();
                d.XmlResolver = null;
                using (XmlReader xr = XmlReader.Create(new MemoryStream(bytes, false), XmlSettings()))
                    d.Load(xr);
                return d;
            }
            catch (Exception) { return null; }
        }

        // ------------------------------------------------------------------
        //  OLE 複合檔 (xls/doc/ppt，以及受 IRM/密碼保護的 OOXML)
        // ------------------------------------------------------------------
        static void InspectCfb(FileStream fs, InspectResult r)
        {
            if (fs.Length < 512)
            {
                AddIssue(r, "OLE 檔頭不完整（檔案被截斷）");
                return;
            }
            byte[] h = ReadAt(fs, 0, 512);
            int major = U16(h, 0x1A);
            int shift = U16(h, 0x1E);
            if (U16(h, 0x1C) != 0xFFFE || (shift != 9 && shift != 12))
            {
                AddIssue(r, "OLE 檔頭欄位無效");
                return;
            }
            int ss = 1 << shift;
            long nSectors = (fs.Length - ss) / ss;
            if ((fs.Length - ss) % ss != 0)
            {
                // 最後一個磁區不完整：多數情況是被截斷，但部分程式會寫出未補齊的檔案
                r.Warnings.Add("OLE 檔案大小不是磁區的整數倍（可能被截斷）");
                nSectors++;
            }
            uint nFat = U32(h, 0x2C);
            uint dirStart = U32(h, 0x30);
            uint difat = U32(h, 0x44);
            if (nFat == 0 || nFat > nSectors)
            {
                AddIssue(r, "OLE 配置表數量異常");
                return;
            }

            List<uint> fatSecs = new List<uint>();
            for (int i = 0; i < 109 && fatSecs.Count < nFat; i++) fatSecs.Add(U32(h, 0x4C + i * 4));
            int guard = 0;
            while (fatSecs.Count < nFat && difat != ENDOFCHAIN && difat != FREESECT)
            {
                if (difat >= nSectors || guard++ > nSectors)
                {
                    AddIssue(r, "OLE DIFAT 磁區鏈損毀");
                    return;
                }
                byte[] ds = ReadSector(fs, difat, ss);
                if (ds == null) { AddIssue(r, "OLE DIFAT 磁區被截斷"); return; }
                int per = ss / 4 - 1;
                for (int i = 0; i < per && fatSecs.Count < nFat; i++) fatSecs.Add(U32(ds, i * 4));
                difat = U32(ds, per * 4);
            }
            if (fatSecs.Count < nFat)
            {
                AddIssue(r, "OLE 配置表清單不完整");
                return;
            }

            int perFat = ss / 4;
            uint[] fat = new uint[fatSecs.Count * perFat];
            for (int i = 0; i < fatSecs.Count; i++)
            {
                uint s = fatSecs[i];
                byte[] sec = s < nSectors ? ReadSector(fs, s, ss) : null;
                if (sec == null)
                {
                    AddIssue(r, "OLE 配置表磁區超出檔案範圍（檔案被截斷）");
                    return;
                }
                for (int j = 0; j < perFat; j++) fat[i * perFat + j] = U32(sec, j * 4);
            }

            string err;
            List<uint> dirChain = Chain(fat, dirStart, nSectors, out err);
            if (dirChain == null)
            {
                AddIssue(r, "OLE 目錄磁區鏈損毀：" + err);
                return;
            }

            List<string> names = new List<string>();
            Dictionary<string, long[]> streams = new Dictionary<string, long[]>();   // 名稱 → { 起始磁區, 大小 }
            long rootStart = ENDOFCHAIN, rootSize = 0;
            bool rootOk = false;
            int entriesPerSector = ss / 128;
            int index = 0;
            foreach (uint sid in dirChain)
            {
                byte[] sec = ReadSector(fs, sid, ss);
                if (sec == null) { AddIssue(r, "OLE 目錄磁區被截斷"); return; }
                for (int k = 0; k < entriesPerSector; k++, index++)
                {
                    int o = k * 128;
                    int type = sec[o + 66];
                    if (type == 0) continue;
                    int nameLen = U16(sec, o + 64);
                    if (nameLen < 2 || nameLen > 64)
                    {
                        AddIssue(r, "OLE 目錄項目名稱長度無效");
                        continue;
                    }
                    string name = Encoding.Unicode.GetString(sec, o, nameLen - 2);
                    uint start = U32(sec, o + 116);
                    long size = U32(sec, o + 120);
                    if (major == 4) size = (long)U64(sec, o + 120);
                    if (index == 0)
                    {
                        rootOk = type == 5;
                        rootStart = start;
                        rootSize = size;
                        if (size > 0 && Chain(fat, start, nSectors, out err) == null)
                            AddIssue(r, "OLE 迷你資料流磁區鏈損毀：" + err);
                        continue;
                    }
                    names.Add(name);
                    if (type == 2 && !streams.ContainsKey(name)) streams[name] = new long[] { start, size };
                    if (type == 2 && size >= 4096)
                    {
                        List<uint> chain = Chain(fat, start, nSectors, out err);
                        if (chain == null) AddIssue(r, "資料流「" + Printable(name) + "」磁區鏈損毀：" + err);
                        else if ((long)chain.Count * ss < size) AddIssue(r, "資料流「" + Printable(name) + "」資料不完整");
                    }
                }
            }
            if (!rootOk) AddIssue(r, "OLE 根目錄項目遺失");

            bool hasEncPkg = false, hasEncInfo = false, hasDrm = false;
            foreach (string n in names)
            {
                if (n == "EncryptedPackage") hasEncPkg = true;
                else if (n == "EncryptionInfo") hasEncInfo = true;
                if (n.IndexOf("DRM", StringComparison.Ordinal) >= 0) hasDrm = true;
                if (n == "Workbook" || n == "Book") r.DocKind = "Excel";
                else if (n == "WordDocument") r.DocKind = "Word";
                else if (n == "PowerPoint Document") r.DocKind = "PowerPoint";
            }
            if (hasDrm || (hasEncPkg && !hasEncInfo)) r.IrmProtected = true;
            else if (hasEncInfo) r.PasswordEncrypted = true;

            // IRM 容器內未加密的中繼資料：敏感度標籤 (LabelInfo) 與發佈授權 (\x06Primary)
            if (r.IrmProtected && r.Issues.Count == 0)
            {
                CfbReader rd = new CfbReader();
                rd.Fs = fs; rd.Fat = fat; rd.NSectors = nSectors; rd.SectorSize = ss;
                rd.MiniFatStart = U32(h, 0x3C); rd.MiniCutoff = U32(h, 0x38); rd.MiniSectorSize = 1 << U16(h, 0x20);
                rd.RootStart = (uint)rootStart; rd.RootSize = rootSize;
                long[] st;
                if (streams.TryGetValue("LabelInfo", out st))
                {
                    byte[] li = rd.Read((uint)st[0], st[1], 256 * 1024);
                    if (li != null) ParseLabelList(li, r);
                }
                if (streams.TryGetValue("\u0006Primary", out st))
                {
                    byte[] pl = rd.Read((uint)st[0], st[1], 4 * 1024 * 1024);
                    if (pl != null) ParsePublishingLicense(pl, r);
                }
                if (r.LabelId.Length > 0 && r.LabelName.Length == 0) r.LabelName = r.ProtectionName;
            }
        }

        /// <summary>讀取 OLE 資料流（含小於迷你資料流門檻、存放在迷你資料流中的資料）。失敗時回傳 null。</summary>
        class CfbReader
        {
            public FileStream Fs;
            public uint[] Fat;
            public long NSectors;
            public int SectorSize;
            public uint MiniFatStart;
            public uint MiniCutoff;
            public int MiniSectorSize;
            public uint RootStart;
            public long RootSize;

            public byte[] Read(uint start, long size, int maxBytes)
            {
                try
                {
                    if (size <= 0 || size > maxBytes) return null;
                    string err;
                    if (size >= MiniCutoff)
                    {
                        List<uint> chain = Chain(Fat, start, NSectors, out err);
                        return chain == null ? null : ReadChain(chain, size);
                    }
                    List<uint> rootChain = Chain(Fat, RootStart, NSectors, out err);
                    List<uint> miniFatChain = Chain(Fat, MiniFatStart, NSectors, out err);
                    if (rootChain == null || miniFatChain == null || RootSize > 64L * 1024 * 1024) return null;
                    byte[] mini = ReadChain(rootChain, RootSize);
                    byte[] mf = ReadChain(miniFatChain, (long)miniFatChain.Count * SectorSize);
                    if (mini == null || mf == null) return null;
                    int nMini = mf.Length / 4;
                    byte[] result = new byte[size];
                    long done = 0;
                    uint s = start;
                    int guard = 0;
                    while (done < size)
                    {
                        if (s >= nMini || guard++ > nMini) return null;
                        long off = (long)s * MiniSectorSize;
                        int take = (int)Math.Min(MiniSectorSize, size - done);
                        if (off + take > mini.Length) return null;
                        Array.Copy(mini, off, result, done, take);
                        done += take;
                        s = U32(mf, (int)s * 4);
                    }
                    return result;
                }
                catch (Exception) { return null; }
            }

            byte[] ReadChain(List<uint> chain, long size)
            {
                if (size > int.MaxValue) return null;
                byte[] result = new byte[size];
                long done = 0;
                foreach (uint sid in chain)
                {
                    if (done >= size) break;
                    byte[] sec = ReadSector(Fs, sid, SectorSize);
                    if (sec == null) return null;
                    int take = (int)Math.Min(SectorSize, size - done);
                    Array.Copy(sec, 0, result, done, take);
                    done += take;
                }
                return done == size ? result : null;
            }
        }

        static List<uint> Chain(uint[] fat, uint start, long nSectors, out string err)
        {
            List<uint> list = new List<uint>();
            uint s = start;
            err = "";
            while (s != ENDOFCHAIN)
            {
                if (s >= nSectors || s >= fat.Length)
                {
                    err = "磁區編號超出範圍";
                    return null;
                }
                list.Add(s);
                if (list.Count > nSectors)
                {
                    err = "磁區鏈出現循環";
                    return null;
                }
                s = fat[s];
            }
            return list;
        }

        static byte[] ReadSector(FileStream fs, uint sid, int ss)
        {
            long pos = ((long)sid + 1) * ss;
            if (pos >= fs.Length) return null;
            byte[] b = new byte[ss];
            fs.Position = pos;
            ReadFully(fs, b, 0, ss); // 最後一個不完整的磁區以 0 補齊
            return b;
        }

        static string Printable(string s)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char ch in s) sb.Append(ch < ' ' ? '?' : ch);
            return sb.ToString();
        }

        // ------------------------------------------------------------------
        //  共用工具
        // ------------------------------------------------------------------
        static void AddIssue(InspectResult r, string msg)
        {
            if (r.Issues.Count < MaxMessages) r.Issues.Add(msg);
            else if (r.Issues.Count == MaxMessages) r.Issues.Add("（其餘問題省略）");
        }

        static int ReadFully(Stream s, byte[] buf, int off, int count)
        {
            int total = 0;
            while (total < count)
            {
                int n = s.Read(buf, off + total, count - total);
                if (n <= 0) break;
                total += n;
            }
            return total;
        }

        static byte[] ReadAt(FileStream fs, long pos, int count)
        {
            if (pos < 0 || pos + count > fs.Length) return null;
            byte[] b = new byte[count];
            fs.Position = pos;
            if (ReadFully(fs, b, 0, count) != count) return null;
            return b;
        }

        static byte[] ReadAll(Stream s)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                byte[] buf = new byte[81920];
                int n;
                while ((n = s.Read(buf, 0, buf.Length)) > 0) ms.Write(buf, 0, n);
                return ms.ToArray();
            }
        }

        static void Drain(Stream s)
        {
            byte[] buf = new byte[81920];
            while (s.Read(buf, 0, buf.Length) > 0) { }
        }

        static int U16(byte[] b, int o) { return b[o] | (b[o + 1] << 8); }
        static uint U32(byte[] b, int o) { return (uint)(b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)); }
        static ulong U64(byte[] b, int o) { return U32(b, o) | ((ulong)U32(b, o + 4) << 32); }
    }

    /// <summary>邊讀取邊計算 CRC32 與長度的唯讀資料流。</summary>
    internal class CrcStream : Stream
    {
        static readonly uint[] Table = BuildTable();
        readonly Stream inner;
        readonly long limit;
        uint crc = 0xFFFFFFFF;
        long count;

        public CrcStream(Stream inner, long limit) { this.inner = inner; this.limit = limit; }
        public uint Crc { get { return crc ^ 0xFFFFFFFF; } }
        public long Count { get { return count; } }

        static uint[] BuildTable()
        {
            uint[] t = new uint[256];
            for (uint i = 0; i < 256; i++)
            {
                uint c = i;
                for (int k = 0; k < 8; k++) c = (c & 1) != 0 ? 0xEDB88320 ^ (c >> 1) : c >> 1;
                t[i] = c;
            }
            return t;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (this.count >= limit) return 0;
            if (count > limit - this.count) count = (int)(limit - this.count);
            int n = inner.Read(buffer, offset, count);
            for (int i = 0; i < n; i++) crc = Table[(crc ^ buffer[offset + i]) & 0xFF] ^ (crc >> 8);
            this.count += n;
            return n;
        }

        public override bool CanRead { get { return true; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return false; } }
        public override long Length { get { throw new NotSupportedException(); } }
        public override long Position { get { return count; } set { throw new NotSupportedException(); } }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
    }
}
