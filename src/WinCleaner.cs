// WinCleaner — очистка Windows от временных файлов и кэшей.
// Собирается компилятором, встроенным в Windows (.NET Framework 4.x), см. build.cmd,
// поэтому код на C# 5: без $"", ?. и прочего синтаксиса новее 2012 года.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("WinCleaner")]
[assembly: AssemblyDescription("Очистка Windows от временных файлов и кэшей")]
[assembly: AssemblyProduct("WinCleaner")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace WinCleaner
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.ThreadException += (s, e) =>
                MessageBox.Show(e.Exception.Message, "WinCleaner", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Application.Run(new MainForm());
        }
    }

    // ============================== Модель ==============================

    sealed class Ctx
    {
        public bool DryRun;                 // true — только посчитать размер, ничего не удалять
        public volatile bool Cancel;
        public readonly List<string> Notes = new List<string>();
        public Action<string> Log = delegate { };
        public Action<string> Status = delegate { };
        public void Note(string s) { lock (Notes) Notes.Add(s); }
    }

    sealed class AppDef
    {
        public string Name;
        public string[] Procs;
        public string PathLike;             // чтобы не спутать, например, browser.exe Яндекса с чужим
        public string[] Local = new string[0];
        public string[] Roaming = new string[0];
        public bool Firefox;
    }

    sealed class CleanItem
    {
        public const long Unknown = -2;
        public string Title, Description, Warning;
        public bool Deep, Net, Checked, SkipIfEmpty;
        public bool Action;                 // не освобождает место, а выполняет действие (сетевые пункты)
        public bool NeedsReboot;
        public AppDef[] Apps;
        public Func<Ctx, long> Run;
        public long Size = -1, Freed = -1;
        public List<string> Notes = new List<string>();
        public ListViewItem Row;
    }

    // ============================== Файлы ==============================

    static class Fs
    {
        // Удаляет содержимое папки (саму папку оставляет). Ссылки/junction не открывает и не трогает,
        // занятые файлы молча пропускает. Возвращает объём удалённого (или найденного при DryRun).
        public static long ClearDir(Ctx c, string path, double minAgeHours = 0, string[] keep = null)
        {
            if (string.IsNullOrEmpty(path) || !Directory.Exists(path)) return 0;
            DateTime cutoff = minAgeHours > 0 ? DateTime.Now.AddHours(-minAgeHours) : DateTime.MaxValue;
            DirectoryInfo root;
            try
            {
                root = new DirectoryInfo(path);
                if ((root.Attributes & FileAttributes.ReparsePoint) != 0) return 0;
            }
            catch { return 0; }

            long freed = 0;
            var dirs = new List<DirectoryInfo>();
            var stack = new Stack<DirectoryInfo>();
            stack.Push(root);
            while (stack.Count > 0 && !c.Cancel)
            {
                var d = stack.Pop();
                FileSystemInfo[] children;
                try { children = d.GetFileSystemInfos(); } catch { continue; }
                foreach (var fsi in children)
                {
                    FileAttributes a;
                    try { a = fsi.Attributes; } catch { continue; }
                    if ((a & FileAttributes.ReparsePoint) != 0) continue;
                    var sub = fsi as DirectoryInfo;
                    if (sub != null) { dirs.Add(sub); stack.Push(sub); continue; }
                    var f = (FileInfo)fsi;
                    if (keep != null && keep.Contains(f.Name, StringComparer.OrdinalIgnoreCase)) continue;
                    freed += DeleteFile(f, cutoff, c.DryRun);
                }
            }
            if (!c.DryRun)
            {
                // Дочерние папки всегда найдены позже родительских — удаляем с конца, пустые уйдут.
                for (int i = dirs.Count - 1; i >= 0; i--)
                {
                    try { dirs[i].Delete(false); }
                    catch
                    {
                        try { dirs[i].Attributes = FileAttributes.Directory; dirs[i].Delete(false); } catch { }
                    }
                }
            }
            return freed;
        }

        static long DeleteFile(FileInfo f, DateTime cutoff, bool dry)
        {
            try
            {
                if (f.LastWriteTime > cutoff || f.CreationTime > cutoff) return 0;
                long len = f.Length;
                if (!dry)
                {
                    if ((f.Attributes & FileAttributes.ReadOnly) != 0) f.Attributes = FileAttributes.Normal;
                    f.Delete();
                }
                return len;
            }
            catch { return 0; }
        }

        public static long RemoveFiles(Ctx c, string dir, string pattern)
        {
            if (!Directory.Exists(dir)) return 0;
            long s = 0;
            try
            {
                foreach (var f in new DirectoryInfo(dir).GetFiles(pattern))
                    if ((f.Attributes & FileAttributes.ReparsePoint) == 0) s += DeleteFile(f, DateTime.MaxValue, c.DryRun);
            }
            catch { }
            return s;
        }

        public static long RemoveFile(Ctx c, string path)
        {
            try
            {
                var f = new FileInfo(path);
                return f.Exists ? DeleteFile(f, DateTime.MaxValue, c.DryRun) : 0;
            }
            catch { return 0; }
        }

        public static long Size(string path) { return ClearDir(new Ctx { DryRun = true }, path); }

        public static List<string> SubDirs(string path)
        {
            var res = new List<string>();
            if (!Directory.Exists(path)) return res;
            try
            {
                foreach (var d in new DirectoryInfo(path).GetDirectories())
                    if ((d.Attributes & FileAttributes.ReparsePoint) == 0) res.Add(d.FullName);
            }
            catch { }
            return res;
        }

        public static long FreeSpace(string drive)
        {
            try { return new DriveInfo(drive).AvailableFreeSpace; } catch { return 0; }
        }

        public static string Fmt(long b)
        {
            if (b >= 1L << 30) return (b / (double)(1L << 30)).ToString("0.00") + " ГБ";
            if (b >= 1L << 20) return (b / (double)(1L << 20)).ToString("0.0") + " МБ";
            if (b >= 1L << 10) return (b / 1024.0).ToString("0") + " КБ";
            return b + " Б";
        }
    }

    // ============================== Система ==============================

    static class Sys
    {
        public static readonly string Win = Environment.GetFolderPath(Environment.SpecialFolder.Windows);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHEmptyRecycleBin(IntPtr hwnd, string root, uint flags);

        // Папки профилей всех пользователей компьютера (локальные, Microsoft- и Azure AD-аккаунты).
        public static List<string> UserProfiles()
        {
            var res = new List<string>();
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"))
                {
                    if (key != null)
                    {
                        foreach (var sid in key.GetSubKeyNames())
                        {
                            if (!sid.StartsWith("S-1-5-21-") && !sid.StartsWith("S-1-12-1-")) continue;
                            using (var k = key.OpenSubKey(sid))
                            {
                                var p = k == null ? null : k.GetValue("ProfileImagePath") as string;
                                if (string.IsNullOrEmpty(p)) continue;
                                p = Environment.ExpandEnvironmentVariables(p);
                                if (Directory.Exists(p) && !res.Contains(p, StringComparer.OrdinalIgnoreCase)) res.Add(p);
                            }
                        }
                    }
                }
            }
            catch { }
            string me = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (!string.IsNullOrEmpty(me) && !res.Contains(me, StringComparer.OrdinalIgnoreCase)) res.Add(me);
            return res;
        }

        public static int Run(string exe, string args, Action<string> onLine, int timeoutMs)
        {
            var psi = new ProcessStartInfo(exe, args) { UseShellExecute = false, CreateNoWindow = true };
            if (onLine != null)
            {
                var enc = Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage);
                psi.RedirectStandardOutput = true;
                psi.RedirectStandardError = true;
                psi.StandardOutputEncoding = enc;
                psi.StandardErrorEncoding = enc;
            }
            try
            {
                using (var p = Process.Start(psi))
                {
                    if (onLine != null)
                    {
                        DataReceivedEventHandler h = (s, e) => { if (!string.IsNullOrEmpty(e.Data)) onLine(e.Data); };
                        p.OutputDataReceived += h;
                        p.ErrorDataReceived += h;
                        p.BeginOutputReadLine();
                        p.BeginErrorReadLine();
                    }
                    if (!p.WaitForExit(timeoutMs)) { try { p.Kill(); } catch { } return -1; }
                    p.WaitForExit();
                    return p.ExitCode;
                }
            }
            catch { return -1; }
        }

        public static List<string> StopServices(params string[] names)
        {
            var stopped = new List<string>();
            foreach (var n in names)
            {
                try
                {
                    using (var sc = new ServiceController(n))
                    {
                        if (sc.Status != ServiceControllerStatus.Running) continue;
                        sc.Stop();
                        stopped.Add(n);
                        sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30));
                    }
                }
                catch { }
            }
            return stopped;
        }

        public static void StartServices(IEnumerable<string> names)
        {
            foreach (var n in names.Reverse())
            {
                try { using (var sc = new ServiceController(n)) sc.Start(); } catch { }
            }
        }

        // Запуск встроенной «Очистки диска» Windows (cleanmgr /sagerun) только с нужными пунктами.
        public static void CleanMgr(Ctx c, params string[] handlers)
        {
            string exe = Path.Combine(Win, @"System32\cleanmgr.exe");
            if (!File.Exists(exe)) { c.Note("встроенная «Очистка диска» Windows не найдена"); return; }
            const string root = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches";
            const string flag = "StateFlags0042";
            var set = new List<string>();
            try
            {
                using (var vc = Registry.LocalMachine.OpenSubKey(root, true))
                {
                    if (vc == null) return;
                    foreach (var h in handlers)
                        using (var k = vc.OpenSubKey(h, true))
                            if (k != null) { k.SetValue(flag, 2, RegistryValueKind.DWord); set.Add(h); }
                }
                if (set.Count > 0) Run(exe, "/sagerun:42", null, 60 * 60 * 1000);
            }
            finally
            {
                try
                {
                    using (var vc = Registry.LocalMachine.OpenSubKey(root, true))
                        if (vc != null)
                            foreach (var h in set)
                                using (var k = vc.OpenSubKey(h, true))
                                    if (k != null) k.DeleteValue(flag, false);
                }
                catch { }
            }
        }

        public static void EmptyRecycleBin()
        {
            try { SHEmptyRecycleBin(IntPtr.Zero, null, 1 | 2 | 4); } catch { } // без вопросов, без окна, без звука
        }

        public static void FlushDns() { Run(Path.Combine(Win, @"System32\ipconfig.exe"), "/flushdns", null, 15000); }

        // Системная команда с результатом одной строкой в журнале (и последней строкой вывода при ошибке).
        public static bool Step(Ctx c, string exe, string args, int timeoutMs = 60000)
        {
            var output = new List<string>();
            int code = Run(Path.Combine(Win, "System32", exe), args, line =>
            {
                lock (output) { if (line.Trim().Length > 0) output.Add(line.Trim()); }
            }, timeoutMs);
            string cmd = Path.GetFileNameWithoutExtension(exe) + " " + args;
            if (code == 0) c.Log("    " + cmd + " — OK");
            else
            {
                string last;
                lock (output) last = output.Count > 0 ? output[output.Count - 1] : "";
                c.Log("    " + cmd + " — код " + code + (last.Length > 0 ? ": " + last : ""));
            }
            return code == 0;
        }

        [DllImport("wininet.dll", SetLastError = true)]
        static extern bool InternetSetOption(IntPtr hInternet, int option, IntPtr buffer, int length);

        // Прокси текущего пользователя (Параметры → Сеть → Прокси): выключаем ручной и скрипт автонастройки.
        public static void ResetUserProxy(Ctx c)
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings", true))
                {
                    if (k == null) return;
                    k.SetValue("ProxyEnable", 0, RegistryValueKind.DWord);
                    k.DeleteValue("AutoConfigURL", false);
                }
                InternetSetOption(IntPtr.Zero, 39, IntPtr.Zero, 0); // INTERNET_OPTION_SETTINGS_CHANGED
                InternetSetOption(IntPtr.Zero, 37, IntPtr.Zero, 0); // INTERNET_OPTION_REFRESH
                c.Log("    прокси пользователя — выключен");
            }
            catch (Exception ex) { c.Note("прокси пользователя: " + ex.Message); }
        }

        public static long HiberSize(string root)
        {
            try
            {
                var f = new DirectoryInfo(root).GetFiles("hiberfil.sys");
                return f.Length > 0 ? f[0].Length : 0;
            }
            catch { return 0; }
        }
    }

    static class AppControl
    {
        public static List<Process> Find(AppDef a)
        {
            var res = new List<Process>();
            foreach (var name in a.Procs)
            {
                foreach (var p in Process.GetProcessesByName(name))
                {
                    if (a.PathLike != null)
                    {
                        string path = null;
                        try { path = p.MainModule.FileName; } catch { }
                        if (path != null && path.IndexOf(a.PathLike, StringComparison.OrdinalIgnoreCase) < 0) { p.Dispose(); continue; }
                    }
                    res.Add(p);
                }
            }
            return res;
        }

        public static bool IsRunning(AppDef a)
        {
            var l = Find(a);
            foreach (var p in l) p.Dispose();
            return l.Count > 0;
        }

        // Сначала вежливо (как крестик окна), через 8 секунд — принудительно.
        public static void Close(AppDef a)
        {
            foreach (var p in Find(a)) using (p) { try { p.CloseMainWindow(); } catch { } }
            for (int i = 0; i < 16 && IsRunning(a); i++) Thread.Sleep(500);
            foreach (var p in Find(a)) using (p) { try { p.Kill(); } catch { } }
            Thread.Sleep(700);
        }
    }

    // ============================== Что чистим ==============================

    static class Catalog
    {
        public static List<string> Users = new List<string>();

        static readonly string[] ChromiumCache = {
            "Cache", "Code Cache", "GPUCache", "DawnCache", "DawnGraphiteCache", "DawnWebGPUCache",
            "GrShaderCache", "GraphiteDawnCache", "ShaderCache",
            @"Service Worker\CacheStorage", @"Service Worker\ScriptCache"
        };

        static readonly AppDef[] Browsers = {
            new AppDef { Name = "Google Chrome", Procs = new[] { "chrome" }, PathLike = @"\Google\", Local = new[] { @"Google\Chrome\User Data" } },
            new AppDef { Name = "Microsoft Edge", Procs = new[] { "msedge" }, Local = new[] { @"Microsoft\Edge\User Data" } },
            new AppDef { Name = "Яндекс Браузер", Procs = new[] { "browser" }, PathLike = @"\Yandex\", Local = new[] { @"Yandex\YandexBrowser\User Data" } },
            new AppDef { Name = "Opera", Procs = new[] { "opera" },
                Local = new[] { @"Opera Software\Opera Stable", @"Opera Software\Opera GX Stable" },
                Roaming = new[] { @"Opera Software\Opera Stable", @"Opera Software\Opera GX Stable" } },
            new AppDef { Name = "Brave", Procs = new[] { "brave" }, Local = new[] { @"BraveSoftware\Brave-Browser\User Data" } },
            new AppDef { Name = "Vivaldi", Procs = new[] { "vivaldi" }, Local = new[] { @"Vivaldi\User Data" } },
            new AppDef { Name = "Chromium", Procs = new[] { "chrome" }, PathLike = @"\Chromium\", Local = new[] { @"Chromium\User Data" } },
            new AppDef { Name = "Firefox", Procs = new[] { "firefox" }, Firefox = true },
        };

        static readonly AppDef[] Programs = {
            new AppDef { Name = "Discord", Procs = new[] { "Discord" }, Roaming = new[] { "discord" } },
            new AppDef { Name = "Slack", Procs = new[] { "slack" }, Roaming = new[] { "Slack" } },
            new AppDef { Name = "Microsoft Teams", Procs = new[] { "Teams" }, Roaming = new[] { @"Microsoft\Teams" } },
            new AppDef { Name = "Steam", Procs = new[] { "steam", "steamwebhelper" }, Local = new[] { @"Steam\htmlcache" } },
        };

        static IEnumerable<string> Roots(AppDef a, string user)
        {
            foreach (var x in a.Local) yield return Path.Combine(user, @"AppData\Local", x);
            foreach (var x in a.Roaming) yield return Path.Combine(user, @"AppData\Roaming", x);
        }

        static bool IsPresent(AppDef a)
        {
            foreach (var u in Users)
            {
                if (a.Firefox) { if (Directory.Exists(Path.Combine(u, @"AppData\Local\Mozilla\Firefox\Profiles"))) return true; }
                else foreach (var r in Roots(a, u)) if (Directory.Exists(r)) return true;
            }
            return false;
        }

        // В папке браузера удаляем только папки кэша — в корне и в каждом профиле (Default, Profile 1…).
        // Пароли, cookies, история, закладки лежат в других файлах и не затрагиваются.
        static long ClearChromium(Ctx c, string root)
        {
            if (!Directory.Exists(root)) return 0;
            var bases = new List<string> { root };
            bases.AddRange(Fs.SubDirs(root));
            long s = 0;
            foreach (var b in bases)
                foreach (var n in ChromiumCache)
                    s += Fs.ClearDir(c, Path.Combine(b, n));
            return s;
        }

        static long ClearApps(Ctx c, AppDef[] apps)
        {
            long total = 0;
            foreach (var a in apps)
            {
                if (c.Cancel) break;
                if (!IsPresent(a)) continue;
                if (!c.DryRun && AppControl.IsRunning(a)) { c.Note(a.Name + ": пропущено — программа открыта"); continue; }
                long s = 0;
                foreach (var u in Users)
                {
                    if (a.Firefox)
                    {
                        foreach (var p in Fs.SubDirs(Path.Combine(u, @"AppData\Local\Mozilla\Firefox\Profiles")))
                            foreach (var n in new[] { "cache2", "startupCache", "thumbnails", "jumpListCache" })
                                s += Fs.ClearDir(c, Path.Combine(p, n));
                    }
                    else
                    {
                        foreach (var r in Roots(a, u)) s += ClearChromium(c, r);
                    }
                }
                if (s > 0) c.Note(a.Name + ": " + Fs.Fmt(s));
                total += s;
            }
            return total;
        }

        static long ForUsers(Func<string, long> f)
        {
            long s = 0;
            foreach (var u in Users) s += f(u);
            return s;
        }

        public static List<CleanItem> Build()
        {
            string win = Sys.Win;
            string sys = Path.GetPathRoot(win);
            string pd = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            Users = Sys.UserProfiles();
            var L = new List<CleanItem>();

            // ---------- Рекомендуется ----------

            L.Add(new CleanItem
            {
                Title = "Временные файлы программ",
                Description = "Папки Temp всех пользователей. Файлы моложе суток не трогаются — их может использовать программа или установщик, который работает прямо сейчас.",
                Run = c => ForUsers(u => Fs.ClearDir(c, Path.Combine(u, @"AppData\Local\Temp"), 24))
            });

            L.Add(new CleanItem
            {
                Title = "Временные файлы Windows",
                Description = "Папка Windows\\Temp. Файлы моложе суток не трогаются.",
                Run = c => Fs.ClearDir(c, Path.Combine(win, "Temp"), 24)
            });

            L.Add(new CleanItem
            {
                Title = "Кэш браузеров",
                Description = "Chrome, Edge, Яндекс, Opera, Brave, Vivaldi, Firefox. Удаляется только кэш: пароли, история, закладки, cookies и вкладки остаются. Если браузер открыт, программа предложит его закрыть.",
                Apps = Browsers,
                Run = c => ClearApps(c, Browsers)
            });

            L.Add(new CleanItem
            {
                Title = "Кэш Discord, Steam, Teams, Slack",
                Description = "Кэш встроенного браузера этих программ. Переписка, настройки и игры не затрагиваются.",
                Apps = Programs,
                Run = c => ClearApps(c, Programs)
            });

            L.Add(new CleanItem
            {
                Title = "Кэш WebView и приложений Microsoft Store",
                Description = "Кэш Internet Explorer / WebView и временные файлы приложений из Microsoft Store.",
                Run = c => ForUsers(u =>
                {
                    long s = Fs.ClearDir(c, Path.Combine(u, @"AppData\Local\Microsoft\Windows\INetCache"));
                    foreach (var p in Fs.SubDirs(Path.Combine(u, @"AppData\Local\Packages")))
                        s += Fs.ClearDir(c, Path.Combine(p, @"AC\INetCache")) + Fs.ClearDir(c, Path.Combine(p, @"AC\Temp"));
                    return s;
                })
            });

            L.Add(new CleanItem
            {
                Title = "Отчёты об ошибках и дампы памяти",
                Description = "Отчёты о сбоях программ и дампы памяти после «синих экранов». Нужны только для диагностики сбоев.",
                Run = c =>
                {
                    long s = 0;
                    foreach (var n in new[] { "ReportArchive", "ReportQueue", "Temp" })
                        s += Fs.ClearDir(c, Path.Combine(pd, @"Microsoft\Windows\WER", n));
                    s += Fs.RemoveFiles(c, Path.Combine(win, "Minidump"), "*.dmp");
                    s += Fs.RemoveFile(c, Path.Combine(win, "MEMORY.DMP"));
                    s += Fs.ClearDir(c, Path.Combine(win, "LiveKernelReports"));
                    s += ForUsers(u => Fs.ClearDir(c, Path.Combine(u, @"AppData\Local\CrashDumps"))
                                     + Fs.ClearDir(c, Path.Combine(u, @"AppData\Local\Microsoft\Windows\WER")));
                    return s;
                }
            });

            L.Add(new CleanItem
            {
                Title = "Кэш оптимизации доставки обновлений",
                Description = "Части обновлений Windows, которыми компьютер делится с другими в сети. Если понадобятся, Windows скачает их снова.",
                Run = c =>
                {
                    string d = Path.Combine(win, @"ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache");
                    if (c.DryRun) return Fs.Size(d);
                    long before = Fs.Size(d);
                    Sys.Run(Path.Combine(win, @"System32\WindowsPowerShell\v1.0\powershell.exe"),
                        "-NoProfile -NonInteractive -Command \"Delete-DeliveryOptimizationCache -Force\"", null, 120000);
                    Fs.ClearDir(c, d); // если командлета нет (старая Windows) — удаляем сами
                    return Math.Max(0, before - Fs.Size(d));
                }
            });

            L.Add(new CleanItem
            {
                Title = "Старые журналы Windows",
                Description = "Логи CBS, DISM и Windows Update старше 7 дней.",
                Run = c =>
                {
                    long s = 0;
                    foreach (var n in new[] { "CBS", "DISM", "WindowsUpdate" })
                        s += Fs.ClearDir(c, Path.Combine(win, "Logs", n), 24 * 7);
                    return s;
                }
            });

            L.Add(new CleanItem
            {
                Title = "Корзина",
                Description = "Файлы в корзине на всех дисках компьютера.",
                Warning = "удалённые файлы нельзя будет восстановить",
                Run = c =>
                {
                    long s = 0;
                    foreach (var drive in DriveInfo.GetDrives())
                    {
                        if (drive.DriveType != DriveType.Fixed) continue;
                        foreach (var sid in Fs.SubDirs(Path.Combine(drive.RootDirectory.FullName, "$Recycle.Bin")))
                            s += Fs.ClearDir(c, sid, 0, new[] { "desktop.ini" });
                    }
                    if (!c.DryRun) Sys.EmptyRecycleBin();
                    return s;
                }
            });

            // ---------- Глубокая очистка ----------

            L.Add(new CleanItem
            {
                Deep = true,
                Title = "Кэш загрузок Windows Update",
                Description = "Установочные файлы уже установленных обновлений (SoftwareDistribution\\Download). Службы обновления на время очистки останавливаются и потом запускаются снова.",
                Run = c =>
                {
                    string d = Path.Combine(win, @"SoftwareDistribution\Download");
                    if (c.DryRun) return Fs.Size(d);
                    var stopped = Sys.StopServices("wuauserv", "bits");
                    try { return Fs.ClearDir(c, d); }
                    finally { Sys.StartServices(stopped); }
                }
            });

            L.Add(new CleanItem
            {
                Deep = true,
                Title = "Кэш шейдеров видеокарты",
                Description = "DirectX, NVIDIA, AMD, Intel. Пересоздаётся сам; при первом запуске игр возможны короткие подтормаживания.",
                Run = c =>
                {
                    long s = ForUsers(u =>
                    {
                        long x = 0;
                        foreach (var n in new[] { "D3DSCache", @"NVIDIA\DXCache", @"NVIDIA\GLCache", @"AMD\DxCache", @"AMD\DxcCache", @"AMD\GLCache", @"AMD\VkCache" })
                            x += Fs.ClearDir(c, Path.Combine(u, @"AppData\Local", n));
                        foreach (var n in new[] { @"NVIDIA\PerDriverVersion\DXCache", @"NVIDIA\PerDriverVersion\GLCache", @"Intel\ShaderCache" })
                            x += Fs.ClearDir(c, Path.Combine(u, @"AppData\LocalLow", n));
                        return x;
                    });
                    return s + Fs.ClearDir(c, Path.Combine(pd, @"NVIDIA Corporation\NV_Cache"));
                }
            });

            L.Add(new CleanItem
            {
                Deep = true,
                Title = "Остатки установщиков драйверов",
                Description = "Распакованные установщики драйверов AMD и NVIDIA (C:\\AMD, C:\\NVIDIA) и скачанные драйверы NVIDIA. Уже установленным драйверам они не нужны.",
                Run = c => Fs.ClearDir(c, Path.Combine(sys, "AMD"))
                         + Fs.ClearDir(c, Path.Combine(sys, "NVIDIA"))
                         + Fs.ClearDir(c, Path.Combine(pd, @"NVIDIA Corporation\Downloader"))
            });

            L.Add(new CleanItem
            {
                Deep = true,
                Title = "Кэши разработчика (npm, pip, yarn, NuGet, Go)",
                Description = "Скачанные пакеты для программирования. Скачаются заново при следующей установке или сборке.",
                Run = c => ForUsers(u =>
                {
                    long x = 0;
                    foreach (var n in new[] { "npm-cache", @"Yarn\Cache", @"pip\cache", @"NuGet\v3-cache", "go-build" })
                        x += Fs.ClearDir(c, Path.Combine(u, @"AppData\Local", n));
                    return x;
                })
            });

            L.Add(new CleanItem
            {
                Deep = true,
                Title = "Файлы установки и обновления Windows",
                Description = "Остатки после крупных обновлений Windows, журналы установки, старые файлы chkdsk. Выполняется встроенной «Очисткой диска» Windows — может занять несколько минут.",
                Run = c =>
                {
                    if (c.DryRun) return Fs.Size(Path.Combine(sys, "$WINDOWS.~BT")) + Fs.Size(Path.Combine(sys, "$WINDOWS.~WS"));
                    long f0 = Fs.FreeSpace(sys);
                    Sys.CleanMgr(c, "Temporary Setup Files", "Windows Upgrade Log Files", "Upgrade Discarded Files", "Setup Log Files", "Old ChkDsk Files");
                    return Math.Max(0, Fs.FreeSpace(sys) - f0);
                }
            });

            L.Add(new CleanItem
            {
                Deep = true,
                SkipIfEmpty = true,
                Title = "Предыдущая версия Windows (Windows.old)",
                Description = "Копия прошлой версии Windows, которая остаётся после крупного обновления. После удаления откатиться на неё будет нельзя.",
                Warning = "вернуться к прошлой версии Windows будет нельзя",
                Run = c =>
                {
                    string d = Path.Combine(sys, "Windows.old");
                    if (c.DryRun) return Fs.Size(d);
                    if (!Directory.Exists(d)) return 0;
                    long f0 = Fs.FreeSpace(sys);
                    Sys.CleanMgr(c, "Previous Installations");
                    return Math.Max(0, Fs.FreeSpace(sys) - f0);
                }
            });

            L.Add(new CleanItem
            {
                Deep = true,
                Title = "Хранилище компонентов WinSxS",
                Description = "Удаляет заменённые старые версии системных компонентов (DISM /StartComponentCleanup). Безопасно, но может идти 5–30 минут. Сколько освободится — заранее неизвестно.",
                Run = c =>
                {
                    if (c.DryRun) return CleanItem.Unknown;
                    long f0 = Fs.FreeSpace(sys);
                    var rx = new Regex(@"(\d{1,3}(?:[.,]\d)?)\s*%");
                    c.Status("Хранилище компонентов WinSxS: запуск DISM (может занять до 30 минут)…");
                    int code = Sys.Run(Path.Combine(win, @"System32\dism.exe"), "/Online /Cleanup-Image /StartComponentCleanup", line =>
                    {
                        var m = rx.Match(line);
                        if (m.Success) c.Status("Хранилище компонентов WinSxS: " + m.Groups[1].Value + "% (может занять до 30 минут)");
                        else if (line.Trim().Length > 0) c.Log("    dism: " + line.Trim());
                    }, 3 * 60 * 60 * 1000);
                    if (code != 0) c.Note("DISM завершился с кодом " + code + " — обычно это значит, что Windows ждёт перезагрузки после обновлений. Повторите после перезагрузки.");
                    return Math.Max(0, Fs.FreeSpace(sys) - f0);
                }
            });

            L.Add(new CleanItem
            {
                Deep = true,
                SkipIfEmpty = true,
                Title = "Отключить гибернацию (hiberfil.sys)",
                Description = "Удаляет файл гибернации. Пропадут режим «Гибернация» и «Быстрый запуск» Windows. На ноутбуках лучше не трогать.",
                Warning = "гибернация и «быстрый запуск» будут отключены",
                Run = c =>
                {
                    long size = Sys.HiberSize(sys);
                    if (c.DryRun || size == 0) return size;
                    Sys.Run(Path.Combine(win, @"System32\powercfg.exe"), "/hibernate off", null, 60000);
                    Thread.Sleep(1000);
                    return Math.Max(0, size - Sys.HiberSize(sys));
                }
            });

            // ---------- Сеть ----------

            L.Add(new CleanItem
            {
                Net = true, Action = true,
                Title = "Сбросить кэш DNS",
                Description = "Помогает, когда отдельные сайты не открываются или открывается старая версия сайта после смены адреса. Безопасно.",
                Run = c =>
                {
                    if (!c.DryRun) Sys.Step(c, "ipconfig.exe", "/flushdns");
                    return 0;
                }
            });

            L.Add(new CleanItem
            {
                Net = true, Action = true,
                Title = "Очистить кэш ARP и NetBIOS",
                Description = "Забытые адреса устройств локальной сети. Помогает, когда компьютер не видит роутер, принтер или другие компьютеры после их замены или перезагрузки. Безопасно.",
                Run = c =>
                {
                    if (c.DryRun) return 0;
                    Sys.Step(c, "netsh.exe", "interface ip delete arpcache");
                    Sys.Step(c, "nbtstat.exe", "-R");
                    Sys.Step(c, "nbtstat.exe", "-RR");
                    return 0;
                }
            });

            L.Add(new CleanItem
            {
                Net = true, Action = true,
                Title = "Получить IP-адрес заново",
                Description = "Заново запрашивает адрес у роутера (ipconfig /release и /renew). Помогает при «Без доступа к интернету» и конфликте IP-адресов. Интернет пропадёт на несколько секунд.",
                Warning = "интернет пропадёт на несколько секунд",
                Run = c =>
                {
                    if (c.DryRun) return 0;
                    Sys.Step(c, "ipconfig.exe", "/release", 60000);
                    if (!Sys.Step(c, "ipconfig.exe", "/renew", 120000))
                        c.Note("не все адаптеры получили адрес — если интернет не появился, перезагрузите роутер");
                    return 0;
                }
            });

            L.Add(new CleanItem
            {
                Net = true, Action = true,
                Title = "Сбросить прокси",
                Description = "Выключает прокси, прописанный для Windows и в настройках Интернета. Помогает, если после VPN, «ускорителей интернета» или вирусов перестали открываться сайты или Microsoft Store. Если прокси вам настроили специально (на работе) — не отмечайте.",
                Warning = "если прокси нужен специально (например, на работе), его придётся настроить заново",
                Run = c =>
                {
                    if (c.DryRun) return 0;
                    Sys.Step(c, "netsh.exe", "winhttp reset proxy");
                    Sys.ResetUserProxy(c);
                    return 0;
                }
            });

            L.Add(new CleanItem
            {
                Net = true, Action = true, NeedsReboot = true,
                Title = "Сбросить Winsock и TCP/IP",
                Description = "Возвращает сетевые настройки Windows к исходным (netsh winsock reset, netsh int ip reset). Самое сильное средство, когда интернет не работает, а остальное не помогло. Нужна перезагрузка. Статический IP и некоторые VPN, возможно, придётся настроить заново.",
                Warning = "понадобится перезагрузка; статический IP и некоторые VPN, возможно, придётся настроить заново",
                Run = c =>
                {
                    if (c.DryRun) return 0;
                    Sys.Step(c, "netsh.exe", "winsock reset");
                    Sys.Step(c, "netsh.exe", "int ip reset");
                    Sys.Step(c, "netsh.exe", "int ipv6 reset");
                    return 0;
                }
            });

            foreach (var it in L) it.Checked = !it.Deep && !it.Net;
            return L;
        }
    }

    // ============================== Интерфейс ==============================

    sealed class BufferedListView : ListView
    {
        public BufferedListView() { DoubleBuffered = true; }
    }

    sealed class DiskBar : Control
    {
        string drive = "C:";
        long total, free;

        public DiskBar()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        }

        public void Set(string d, long t, long f)
        {
            drive = d.TrimEnd('\\'); total = t; free = f;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(BackColor);
            int th = TextRenderer.MeasureText("Ag", Font).Height;
            using (var bold = new Font(Font, FontStyle.Bold))
                TextRenderer.DrawText(g, "Диск " + drive, bold, new Point(0, 0), ForeColor, TextFormatFlags.NoPadding);
            if (total > 0)
            {
                string right = "свободно " + Fs.Fmt(free) + " из " + Fs.Fmt(total);
                TextRenderer.DrawText(g, right, Font, new Rectangle(0, 0, Width, th), Color.FromArgb(107, 114, 128),
                    TextFormatFlags.Right | TextFormatFlags.NoPadding);
            }
            var r = new Rectangle(0, th + 4, Width - 1, Math.Max(6, Height - th - 6));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var bg = new SolidBrush(Color.FromArgb(229, 231, 235)))
            using (var p = Round(r, r.Height / 2))
                g.FillPath(bg, p);
            if (total > 0)
            {
                double used = 1.0 - (double)free / total;
                var ur = new Rectangle(r.X, r.Y, Math.Max(r.Height, (int)(r.Width * used)), r.Height);
                var col = free < total / 10 ? Color.FromArgb(220, 38, 38) : Color.FromArgb(37, 99, 235);
                using (var b = new SolidBrush(col))
                using (var p = Round(ur, r.Height / 2))
                    g.FillPath(b, p);
            }
        }

        static GraphicsPath Round(Rectangle r, int rad)
        {
            int d = Math.Max(1, Math.Min(rad * 2, Math.Min(r.Width, r.Height)));
            var p = new GraphicsPath();
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }
    }

    sealed class MainForm : Form
    {
        static readonly Color Accent = Color.FromArgb(37, 99, 235);
        static readonly Color Danger = Color.FromArgb(220, 38, 38);
        static readonly Color Ink = Color.FromArgb(17, 24, 39);
        static readonly Color Muted = Color.FromArgb(107, 114, 128);
        static readonly Color Good = Color.FromArgb(22, 163, 74);

        readonly float k;
        readonly List<CleanItem> items;
        readonly string sysDrive;
        readonly ListView list;
        readonly Label desc, status, total;
        readonly DiskBar disk;
        readonly Button btnScan, btnClean;
        readonly ProgressBar bar;
        readonly TextBox log;
        bool busy, cleaning, ready;
        Ctx current;

        int S(int px) { return (int)Math.Round(px * k); }

        public MainForm()
        {
            using (var g = Graphics.FromHwnd(IntPtr.Zero)) k = g.DpiX / 96f;
            sysDrive = Path.GetPathRoot(Sys.Win);
            items = Catalog.Build();

            Text = "WinCleaner — очистка Windows";
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            ForeColor = Ink;
            AutoScaleMode = AutoScaleMode.None;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(S(780), Math.Min(S(800), Screen.PrimaryScreen.WorkingArea.Height - S(60)));
            MinimumSize = new Size(S(660), S(620));
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            // --- шапка ---
            // Ширину задаём сразу: якоря Right считаются от исходного размера панели (по умолчанию 200).
            var header = new Panel { Dock = DockStyle.Top, Size = new Size(S(780), S(142)) };
            header.Controls.Add(new Label
            {
                Text = "Очистка Windows",
                Font = new Font("Segoe UI Semibold", 16f),
                AutoSize = true,
                Location = new Point(S(17), S(12))
            });
            header.Controls.Add(new Label
            {
                Text = "Отметьте, что удалить, и нажмите «Очистить». Документы, фото, загрузки, пароли и история браузеров не затрагиваются.",
                ForeColor = Muted,
                Location = new Point(S(20), S(50)),
                Size = new Size(S(740), S(40)),
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right
            });
            disk = new DiskBar
            {
                Location = new Point(S(20), S(94)),
                Size = new Size(S(740), S(42)),
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right
            };
            header.Controls.Add(disk);

            // --- список ---
            var listHost = new Panel { Dock = DockStyle.Fill, Padding = new Padding(S(20), S(6), S(20), S(4)) };
            list = new BufferedListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                CheckBoxes = true,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false,
                ShowItemToolTips = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                BorderStyle = BorderStyle.FixedSingle,
                SmallImageList = new ImageList { ImageSize = new Size(1, S(28)) }
            };
            list.Columns.Add("Что очистить", S(560));
            list.Columns.Add("Размер", S(140), HorizontalAlignment.Right);
            var gSafe = new ListViewGroup("safe", "Рекомендуется — безопасно");
            var gDeep = new ListViewGroup("deep", "Глубокая очистка — по желанию");
            var gNet = new ListViewGroup("net", "Сеть — если интернет работает плохо (отмечайте вручную)");
            list.Groups.Add(gSafe);
            list.Groups.Add(gDeep);
            list.Groups.Add(gNet);
            foreach (var it in items)
            {
                var row = new ListViewItem(it.Title)
                {
                    Checked = it.Checked,
                    Group = it.Net ? gNet : it.Deep ? gDeep : gSafe,
                    Tag = it,
                    ToolTipText = it.Description,
                    UseItemStyleForSubItems = false
                };
                row.SubItems.Add("…");
                it.Row = row;
                list.Items.Add(row);
            }
            list.ItemCheck += (s, e) => { if (busy) e.NewValue = e.CurrentValue; };
            list.ItemChecked += (s, e) =>
            {
                ((CleanItem)e.Item.Tag).Checked = e.Item.Checked;
                UpdateTotal();
            };
            list.SelectedIndexChanged += (s, e) => ShowDesc();
            list.Resize += (s, e) => FitColumns();
            listHost.Controls.Add(list);

            // --- описание выбранного пункта ---
            var descHost = new Panel { Dock = DockStyle.Bottom, Height = S(62), Padding = new Padding(S(20), S(4), S(20), 0) };
            desc = new Label { Dock = DockStyle.Fill, ForeColor = Muted, Text = "Выберите пункт, чтобы увидеть подробности." };
            descHost.Controls.Add(desc);

            // --- итог, кнопки, прогресс ---
            var actions = new Panel { Dock = DockStyle.Bottom, Size = new Size(S(780), S(92)) };
            total = new Label
            {
                AutoSize = true,
                Location = new Point(S(20), S(6)),
                Font = new Font("Segoe UI Semibold", 10.5f),
                Text = "Считаю…"
            };
            var links = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Location = new Point(S(360), S(6)),
                Size = new Size(S(400), S(26)),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                WrapContents = false
            };
            links.Controls.Add(MakeLink("Ничего", () => SetAll(x => false)));
            links.Controls.Add(MakeLink("Всю очистку", () => SetAll(x => !x.Net)));
            links.Controls.Add(MakeLink("Рекомендуемые", () => SetAll(x => !x.Deep && !x.Net)));
            links.Controls.Add(new Label { Text = "Отметить:", AutoSize = true, ForeColor = Muted, Margin = new Padding(0, S(3), 0, 0) });
            status = new Label
            {
                Location = new Point(S(20), S(42)),
                Size = new Size(S(440), S(22)),
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right,
                AutoEllipsis = true,
                ForeColor = Muted
            };
            bar = new ProgressBar
            {
                Location = new Point(S(20), S(68)),
                Size = new Size(S(440), S(8)),
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right
            };
            btnScan = new Button
            {
                Text = "Пересчитать",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Ink,
                Location = new Point(S(470), S(38)),
                Size = new Size(S(130), S(40)),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Cursor = Cursors.Hand
            };
            btnScan.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
            btnScan.Click += (s, e) => StartScan();
            btnClean = new Button
            {
                Text = "Очистить",
                FlatStyle = FlatStyle.Flat,
                BackColor = Accent,
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 10.5f),
                Location = new Point(S(610), S(38)),
                Size = new Size(S(150), S(40)),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Cursor = Cursors.Hand
            };
            btnClean.FlatAppearance.BorderSize = 0;
            btnClean.Click += (s, e) => StartClean();
            actions.Controls.AddRange(new Control[] { total, links, status, bar, btnScan, btnClean });

            // --- журнал ---
            var logHost = new Panel { Dock = DockStyle.Bottom, Height = S(118), Padding = new Padding(S(20), S(6), S(20), S(18)) };
            log = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(249, 250, 251),
                Font = new Font("Consolas", 9f)
            };
            logHost.Controls.Add(log);

            // Порядок важен: Fill добавляется первым, последний добавленный Bottom-блок оказывается в самом низу.
            Controls.Add(listHost);
            Controls.Add(header);
            Controls.Add(descHost);
            Controls.Add(actions);
            Controls.Add(logHost);

            ready = true;
            RefreshDisk();
            FitColumns();
            AppendLog("Профилей пользователей на компьютере: " + Catalog.Users.Count);
            Shown += (s, e) => StartScan();
            FormClosing += OnClosingForm;
        }

        LinkLabel MakeLink(string text, Action onClick)
        {
            var l = new LinkLabel
            {
                Text = text,
                AutoSize = true,
                LinkColor = Accent,
                ActiveLinkColor = Accent,
                LinkBehavior = LinkBehavior.HoverUnderline,
                Margin = new Padding(S(10), S(3), 0, 0)
            };
            l.LinkClicked += (s, e) => { if (!busy) onClick(); };
            return l;
        }

        void SetAll(Func<CleanItem, bool> pick)
        {
            foreach (var it in items) it.Row.Checked = pick(it);
        }

        void FitColumns()
        {
            if (list.Columns.Count < 2) return;
            int w = list.ClientSize.Width - list.Columns[1].Width - S(4);
            if (w > S(200)) list.Columns[0].Width = w;
        }

        void RefreshDisk()
        {
            try
            {
                var d = new DriveInfo(sysDrive);
                disk.Set(sysDrive, d.TotalSize, d.AvailableFreeSpace);
            }
            catch { }
        }

        void AppendLog(string s)
        {
            log.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + s + Environment.NewLine);
        }

        void ShowDesc()
        {
            if (list.SelectedItems.Count == 0) return;
            var it = (CleanItem)list.SelectedItems[0].Tag;
            string text = it.Description;
            var notes = it.Notes;
            if (notes.Count > 0) text += Environment.NewLine + string.Join(" · ", notes);
            desc.Text = text;
        }

        void UpdateRow(CleanItem it)
        {
            var cell = it.Row.SubItems[1];
            if (it.Freed >= 0) { cell.Text = it.Action ? "✓ выполнено" : "✓ " + Fs.Fmt(it.Freed); cell.ForeColor = Good; }
            else if (it.Action) { cell.Text = "действие"; cell.ForeColor = Muted; }
            else if (it.Size == -1) { cell.Text = "…"; cell.ForeColor = Muted; }
            else if (it.Size == CleanItem.Unknown) { cell.Text = "неизвестно"; cell.ForeColor = Muted; }
            else if (it.Size == 0) { cell.Text = "пусто"; cell.ForeColor = Muted; }
            else { cell.Text = Fs.Fmt(it.Size); cell.ForeColor = Ink; }
            if (it.Row.Selected) ShowDesc();
        }

        void UpdateTotal()
        {
            if (!ready) return;
            long sum = 0;
            bool unknown = false, pending = false;
            foreach (var it in items)
            {
                if (!it.Checked || it.Freed >= 0) continue;
                if (it.Size > 0) sum += it.Size;
                else if (it.Size == CleanItem.Unknown) unknown = true;
                else if (it.Size == -1) pending = true;
            }
            total.Text = "Можно освободить: " + Fs.Fmt(sum) + (unknown ? " + WinSxS" : "") + (pending ? "  (считаю…)" : "");
        }

        void SetBusy(bool b, bool isClean)
        {
            busy = b;
            cleaning = b && isClean;
            btnScan.Enabled = !b;
            btnClean.Enabled = true;
            btnClean.Text = b ? "Остановить" : "Очистить";
            btnClean.BackColor = b ? Danger : Accent;
            if (!b) current = null;
        }

        Ctx NewCtx(bool dry)
        {
            var c = new Ctx { DryRun = dry };
            c.Log = s => Ui(() => AppendLog(s));
            c.Status = s => Ui(() => status.Text = s);
            current = c;
            return c;
        }

        static long Exec(CleanItem it, Ctx ctx)
        {
            lock (ctx.Notes) ctx.Notes.Clear();
            long r;
            try { r = it.Run(ctx); }
            catch (Exception ex) { ctx.Note("ошибка: " + ex.Message); r = 0; }
            lock (ctx.Notes) it.Notes = new List<string>(ctx.Notes);
            return r;
        }

        void Ui(Action a)
        {
            if (IsDisposed || !IsHandleCreated) return;
            try { BeginInvoke(a); } catch (InvalidOperationException) { }
        }

        void RunBackground(Action work, Action done)
        {
            var t = new Thread(() =>
            {
                try { work(); }
                catch (Exception ex) { Ui(() => AppendLog("Ошибка: " + ex.Message)); }
                Ui(done);
            });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        void StartScan()
        {
            if (busy) return;
            foreach (var it in items) { it.Size = -1; it.Freed = -1; it.Notes = new List<string>(); UpdateRow(it); }
            var ctx = NewCtx(true);
            SetBusy(true, false);
            status.Text = "Считаю, сколько места можно освободить…";
            bar.Maximum = items.Count;
            bar.Value = 0;
            UpdateTotal();
            RunBackground(() =>
            {
                int i = 0;
                foreach (var it in items)
                {
                    if (ctx.Cancel) break;
                    var item = it;
                    Ui(() => { item.Row.SubItems[1].Text = "считаю…"; });
                    item.Size = Exec(item, ctx);
                    int done = ++i;
                    Ui(() => { UpdateRow(item); bar.Value = done; UpdateTotal(); });
                }
            }, () =>
            {
                SetBusy(false, false);
                RefreshDisk();
                UpdateTotal();
                status.Text = ctx.Cancel ? "Подсчёт остановлен." : "Готово. Проверьте галочки и нажмите «Очистить».";
            });
        }

        void StartClean()
        {
            if (busy)
            {
                if (current != null)
                {
                    current.Cancel = true;
                    btnClean.Enabled = false;
                    status.Text = "Останавливаю после текущего пункта…";
                }
                return;
            }

            var sel = items.Where(x => x.Checked && !(x.SkipIfEmpty && x.Size == 0)).ToList();
            if (sel.Count == 0)
            {
                MessageBox.Show(this, "Ничего не выбрано (или в отмеченных пунктах пусто).", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var warn = sel.Where(x => x.Warning != null).ToList();
            if (warn.Count > 0)
            {
                string msg = "Обратите внимание:" + Environment.NewLine + Environment.NewLine +
                    string.Join(Environment.NewLine, warn.Select(x =>
                        "• " + x.Title + (x.Size > 0 ? " (" + Fs.Fmt(x.Size) + ")" : "") + " — " + x.Warning)) +
                    Environment.NewLine + Environment.NewLine + "Продолжить?";
                if (MessageBox.Show(this, msg, "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                        MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            }

            var running = sel.Where(x => x.Apps != null).SelectMany(x => x.Apps)
                             .Where(a => AppControl.IsRunning(a)).ToList();
            bool closeApps = false;
            if (running.Count > 0)
            {
                var r = MessageBox.Show(this,
                    "Сейчас открыты: " + string.Join(", ", running.Select(a => a.Name)) + "." + Environment.NewLine + Environment.NewLine +
                    "Пока программа открыта, её кэш почистить нельзя." + Environment.NewLine + Environment.NewLine +
                    "Да — закрыть их сейчас (браузеры восстановят вкладки при следующем запуске)" + Environment.NewLine +
                    "Нет — не закрывать и пропустить их кэш",
                    "Открытые программы", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (r == DialogResult.Cancel) return;
                closeApps = r == DialogResult.Yes;
            }

            var ctx = NewCtx(false);
            SetBusy(true, true);
            long free0 = Fs.FreeSpace(sysDrive);
            bar.Maximum = sel.Count;
            bar.Value = 0;
            AppendLog("Начинаю очистку…");
            RunBackground(() =>
            {
                if (closeApps)
                {
                    Ui(() => status.Text = "Закрываю программы…");
                    foreach (var a in running) AppControl.Close(a);
                }
                int i = 0;
                foreach (var it in sel)
                {
                    if (ctx.Cancel) break;
                    var item = it;
                    Ui(() =>
                    {
                        status.Text = (item.Action ? "Выполняю: " : "Очищаю: ") + item.Title + "…";
                        item.Row.SubItems[1].Text = item.Action ? "выполняю…" : "очищаю…";
                        item.Row.SubItems[1].ForeColor = Accent;
                        item.Row.EnsureVisible();
                    });
                    if (item.Action) ctx.Log(item.Title + ":");
                    item.Freed = Math.Max(0, Exec(item, ctx));
                    int done = ++i;
                    Ui(() =>
                    {
                        UpdateRow(item);
                        bar.Value = done;
                        if (!item.Action) AppendLog(item.Title + " — освобождено " + Fs.Fmt(item.Freed));
                        foreach (var n in item.Notes) AppendLog("    " + n);
                    });
                }
                Sys.FlushDns();
            }, () =>
            {
                long free1 = Fs.FreeSpace(sysDrive);
                long sum = sel.Where(x => x.Freed > 0).Sum(x => x.Freed);
                SetBusy(false, false);
                RefreshDisk();
                int actions = sel.Count(x => x.Action && x.Freed >= 0);
                bool reboot = sel.Any(x => x.NeedsReboot && x.Freed >= 0);
                string res = (ctx.Cancel ? "Очистка остановлена. " : "Готово! ") + "Освобождено: " + Fs.Fmt(sum) + ".";
                if (actions > 0) res += " Сетевых действий выполнено: " + actions + ".";
                string diskLine = "Свободно на диске " + sysDrive.TrimEnd('\\') + ": было " + Fs.Fmt(free0) + ", стало " + Fs.Fmt(free1) + ".";
                status.Text = res;
                total.Text = "Освобождено: " + Fs.Fmt(sum);
                AppendLog(res + " " + diskLine);
                string msg = res + Environment.NewLine + Environment.NewLine + diskLine;
                if (!reboot)
                {
                    MessageBox.Show(this, msg, "WinCleaner", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                msg += Environment.NewLine + Environment.NewLine +
                       "Сброс Winsock и TCP/IP вступит в силу после перезагрузки." + Environment.NewLine +
                       "Перезагрузить компьютер сейчас? Сначала сохраните открытые документы.";
                if (MessageBox.Show(this, msg, "Нужна перезагрузка", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    Sys.Run(Path.Combine(Sys.Win, @"System32\shutdown.exe"), "/r /t 5", null, 10000);
            });
        }

        void OnClosingForm(object sender, FormClosingEventArgs e)
        {
            if (!busy) return;
            if (cleaning && MessageBox.Show(this, "Очистка ещё идёт. Прервать и закрыть программу?", "WinCleaner",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                e.Cancel = true;
                return;
            }
            if (current != null) current.Cancel = true;
        }
    }
}
