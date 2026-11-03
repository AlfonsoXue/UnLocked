// UnLocked.exe 啟動程式
// - 內嵌 src\*.ps1 / *.cs，第一次執行時解出到 %LOCALAPPDATA%\UnLocked\app\<版本雜湊>\src
// - 在本程序內以 Windows PowerShell 5.1 引擎 (System.Management.Automation) 執行 UnLocked.ps1，
//   不另外開 powershell.exe，也不會出現主控台視窗
// - 加上 -NoGui 時為命令列模式，輸出會寫到呼叫端的主控台
// 注意：由 .NET Framework 內建的 csc.exe 編譯，必須維持 C# 5 語法。
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace UnLocked.Launcher
{
    internal static class Program
    {
        const string ResourcePrefix = "UnLocked.src.";
        static readonly Regex ParamName = new Regex("^-[A-Za-z][A-Za-z0-9]*:?$");

        [DllImport("kernel32.dll")]
        static extern bool AttachConsole(int processId);

        [STAThread]
        static int Main(string[] args)
        {
            bool cli = false;
            foreach (string a in args)
                if (string.Equals(a, "-NoGui", StringComparison.OrdinalIgnoreCase)) cli = true;
            if (cli) AttachConsole(-1);   // 附加到呼叫端的主控台（從 cmd / PowerShell 執行時）

            try
            {
                string srcDir = ExtractScripts();
                string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                Environment.SetEnvironmentVariable("UNLOCKED_HOME", exeDir);
                return RunScript(Path.Combine(srcDir, "UnLocked.ps1"), args, cli);
            }
            catch (Exception ex)
            {
                string msg = "UnLocked 無法啟動：" + Environment.NewLine + Environment.NewLine + Describe(ex);
                if (cli) Console.Error.WriteLine(msg);
                else MessageBox.Show(msg, "UnLocked", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 2;
            }
        }

        static string Describe(Exception ex)
        {
            RuntimeException rex = ex as RuntimeException;
            if (rex != null && rex.ErrorRecord != null)
            {
                string pos = rex.ErrorRecord.InvocationInfo != null ? rex.ErrorRecord.InvocationInfo.PositionMessage : "";
                return rex.Message + Environment.NewLine + pos;
            }
            return ex.Message;
        }

        // ------------------------------------------------------------------
        //  解出內嵌的指令碼
        // ------------------------------------------------------------------
        static string ExtractScripts()
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            List<string> names = new List<string>();
            foreach (string n in asm.GetManifestResourceNames())
                if (n.StartsWith(ResourcePrefix, StringComparison.Ordinal)) names.Add(n);
            if (names.Count == 0) throw new InvalidOperationException("執行檔中找不到內嵌的指令碼，請重新建置。");
            names.Sort(StringComparer.Ordinal);

            Dictionary<string, byte[]> files = new Dictionary<string, byte[]>();
            string hash;
            using (SHA256 sha = SHA256.Create())
            {
                foreach (string n in names)
                {
                    byte[] data = ReadResource(asm, n);
                    files[n.Substring(ResourcePrefix.Length)] = data;
                    byte[] nameBytes = Encoding.UTF8.GetBytes(n);
                    sha.TransformBlock(nameBytes, 0, nameBytes.Length, null, 0);
                    sha.TransformBlock(data, 0, data.Length, null, 0);
                }
                sha.TransformFinalBlock(new byte[0], 0, 0);
                hash = BitConverter.ToString(sha.Hash, 0, 8).Replace("-", "").ToLowerInvariant();
            }

            string baseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "UnLocked", "app");
            string srcDir = Path.Combine(Path.Combine(baseDir, hash), "src");
            Directory.CreateDirectory(srcDir);
            foreach (KeyValuePair<string, byte[]> kv in files)
            {
                string target = Path.Combine(srcDir, kv.Key);
                // 每次啟動都比對內容，避免使用被修改過的檔案
                if (File.Exists(target) && SameBytes(File.ReadAllBytes(target), kv.Value)) continue;
                string tmp = target + "." + Guid.NewGuid().ToString("N") + ".tmp";
                File.WriteAllBytes(tmp, kv.Value);
                try
                {
                    if (File.Exists(target)) File.Delete(target);
                    File.Move(tmp, target);
                }
                catch (IOException)
                {
                    // 可能同時開了兩個 UnLocked：若另一個已寫好相同內容就沿用
                    if (File.Exists(tmp)) File.Delete(tmp);
                    if (!File.Exists(target) || !SameBytes(File.ReadAllBytes(target), kv.Value)) throw;
                }
            }
            CleanupOldVersions(baseDir, hash);
            return srcDir;
        }

        static byte[] ReadResource(Assembly asm, string name)
        {
            using (Stream s = asm.GetManifestResourceStream(name))
            using (MemoryStream ms = new MemoryStream())
            {
                s.CopyTo(ms);
                return ms.ToArray();
            }
        }

        static bool SameBytes(byte[] a, byte[] b)
        {
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++) if (a[i] != b[i]) return false;
            return true;
        }

        static void CleanupOldVersions(string baseDir, string current)
        {
            try
            {
                foreach (string dir in Directory.GetDirectories(baseDir))
                {
                    if (string.Equals(Path.GetFileName(dir), current, StringComparison.OrdinalIgnoreCase)) continue;
                    // 只清除 7 天前的舊版本，避免影響仍在執行中的舊版 UnLocked
                    if (Directory.GetLastWriteTimeUtc(dir) > DateTime.UtcNow.AddDays(-7)) continue;
                    try { Directory.Delete(dir, true); } catch (Exception) { }
                }
            }
            catch (Exception) { }
        }

        // ------------------------------------------------------------------
        //  在本程序內執行 PowerShell 指令碼
        // ------------------------------------------------------------------
        static int RunScript(string scriptPath, string[] args, bool cli)
        {
            InitialSessionState iss = InitialSessionState.CreateDefault();
            // 只影響本程序（與 powershell.exe -ExecutionPolicy Bypass 相同）；群組原則強制的設定仍然優先
            iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;
            using (Runspace rs = RunspaceFactory.CreateRunspace(iss))
            {
                rs.ApartmentState = System.Threading.ApartmentState.STA;
                rs.ThreadOptions = PSThreadOptions.UseCurrentThread;   // 在主要 STA 執行緒上執行（Windows Forms 需要）
                rs.Open();
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = rs;
                    ps.AddScript(BuildCommand(scriptPath, args));
                    if (cli) HookConsoleStreams(ps);
                    Collection<PSObject> output = ps.Invoke();
                    if (cli)
                        foreach (ErrorRecord er in ps.Streams.Error) Console.Error.WriteLine(er.ToString());
                    for (int i = output.Count - 1; i >= 0; i--)
                    {
                        if (output[i] == null) continue;
                        object v = output[i].BaseObject;
                        if (v is int) return (int)v;
                        break;
                    }
                    return 0;
                }
            }
        }

        static string BuildCommand(string scriptPath, string[] args)
        {
            StringBuilder sb = new StringBuilder("& ");
            sb.Append(Quote(scriptPath));
            foreach (string a in args)
            {
                sb.Append(' ');
                if (ParamName.IsMatch(a)) sb.Append(a);   // 參數名稱，例如 -NoGui
                else sb.Append(Quote(a));                 // 其他一律當成字串，避免被當成程式碼執行
            }
            sb.Append("; if ($LASTEXITCODE -is [int]) { $LASTEXITCODE } else { 0 }");
            return sb.ToString();
        }

        static string Quote(string s)
        {
            StringBuilder sb = new StringBuilder("'");
            foreach (char c in s)
            {
                sb.Append(c);
                // PowerShell 單引號字串中，' 與全形單引號都要重複一次才代表字元本身
                if (c == '\'' || c == '‘' || c == '’' || c == '‚' || c == '‛') sb.Append(c);
            }
            return sb.Append('\'').ToString();
        }

        static void HookConsoleStreams(PowerShell ps)
        {
            ps.Streams.Information.DataAdded += delegate(object sender, DataAddedEventArgs e)
            {
                InformationRecord rec = ps.Streams.Information[e.Index];
                HostInformationMessage msg = rec.MessageData as HostInformationMessage;
                if (msg == null) { Console.WriteLine(rec.MessageData); return; }
                bool colored = false;
                ConsoleColor old = ConsoleColor.Gray;
                try
                {
                    if (msg.ForegroundColor.HasValue)
                    {
                        old = Console.ForegroundColor;
                        Console.ForegroundColor = msg.ForegroundColor.Value;
                        colored = true;
                    }
                }
                catch (Exception) { }   // 輸出被導向檔案時無法設定顏色
                if (msg.NoNewLine.GetValueOrDefault()) Console.Write(msg.Message); else Console.WriteLine(msg.Message);
                if (colored) { try { Console.ForegroundColor = old; } catch (Exception) { } }
            };
            ps.Streams.Warning.DataAdded += delegate(object sender, DataAddedEventArgs e)
            {
                Console.Error.WriteLine("警告：" + ps.Streams.Warning[e.Index].Message);
            };
        }
    }
}
