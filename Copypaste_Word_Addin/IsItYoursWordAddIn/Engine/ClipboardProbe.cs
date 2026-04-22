// ClipboardProbe.cs
// C# 7.3 ¡ª self-contained clipboard listener for the add-in.
// Usage:
//   private IClipboardProbe _clip;
//   _clip = new ClipboardProbe();
//   _clip.CandidateAvailable += c => { Provenance.SetCandidate(_engine.State, c); };
//   _clip.Start();  // and later: _clip.Stop(); _clip.Dispose();

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace IsItYoursWordAddIn
{
    public sealed class ClipboardCandidate
    {
        public DateTime Utc;
        public string Process;      // e.g., "chrome.exe", "WINWORD.EXE", "firefox.exe"
        public string Text;         // FULL clipboard text (UTF-16)
        public string SourceUrl;    // CF_HTML SourceURL (if present)
        public string ChromiumUrl;  // optional
        public string FirefoxTitle; // optional
    }

    public interface IClipboardProbe : IDisposable
    {
        event Action<ClipboardCandidate> CandidateAvailable;
        void Start();
        void Stop();
    }

    internal sealed class ClipboardProbe : IClipboardProbe
    {
        public event Action<ClipboardCandidate> CandidateAvailable;

        // ---- Config / state ----
        private const int WM_CLIPBOARDUPDATE = 0x031D;
        private const int WM_DESTROY = 0x0002;
        private const uint CF_UNICODETEXT = 13;

        private const int MIN_COPY_LEN = 20;
        private readonly TimeSpan _dedupWindow = TimeSpan.FromMilliseconds(800);

        private Thread _thread;
        private IntPtr _hwnd = IntPtr.Zero;
        private volatile bool _running;
        private volatile bool _disposed;

        private string _lastEventKey = "";
        private DateTime _lastEventUtc = DateTime.MinValue;

        private static readonly WndProcDelegate WndProcThunk = WndProc;
        private delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, UIntPtr wParam, IntPtr lParam);

        private enum BrowserKind { Chromium, Firefox, Other }

        public void Start()
        {
            if (_disposed || _running) return;
            _running = true;
            _thread = new Thread(MessageThread) { IsBackground = true };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        public void Stop()
        {
            _running = false;
            var h = _hwnd;
            if (h != IntPtr.Zero)
            {
                try { PostMessageW(h, WM_DESTROY, UIntPtr.Zero, IntPtr.Zero); } catch { }
            }
            if (_thread != null && _thread.IsAlive)
            {
                try { _thread.Join(300); } catch { }
            }
            _hwnd = IntPtr.Zero;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Stop();
        }

        // ---- Message thread & window ----

        private void MessageThread()
        {
            var wc = new WNDCLASSEX
            {
                cbSize = (uint)Marshal.SizeOf(typeof(WNDCLASSEX)),
                style = 0x0020, // CS_OWNDC
                lpfnWndProc = Marshal.GetFunctionPointerForDelegate(WndProcThunk),
                cbClsExtra = 0,
                cbWndExtra = 0,
                hInstance = GetModuleHandleW(null),
                hIcon = IntPtr.Zero,
                hCursor = IntPtr.Zero,
                hbrBackground = IntPtr.Zero,
                lpszMenuName = null,
                lpszClassName = "ClipLiteWnd",
                hIconSm = IntPtr.Zero
            };

            if (RegisterClassExW(in wc) == 0) { _running = false; return; }

            _hwnd = CreateWindowExW(0, wc.lpszClassName, "ClipLiteHidden", 0,
                                    0, 0, 0, 0, IntPtr.Zero, IntPtr.Zero, wc.hInstance, IntPtr.Zero);
            if (_hwnd == IntPtr.Zero) { _running = false; return; }

            if (!AddClipboardFormatListener(_hwnd))
            {
                DestroyWindow(_hwnd);
                _hwnd = IntPtr.Zero;
                _running = false;
                return;
            }

            // Message loop
            MSG msg;
            while (_running && GetMessageW(out msg, IntPtr.Zero, 0, 0) > 0)
            {
                TranslateMessage(in msg);
                DispatchMessageW(in msg);
            }

            // Cleanup if WM_DESTROY didn¡¯t run
            if (_hwnd != IntPtr.Zero)
            {
                try { RemoveClipboardFormatListener(_hwnd); } catch { }
                try { DestroyWindow(_hwnd); } catch { }
                _hwnd = IntPtr.Zero;
            }
        }

        // Static WndProc forwards to the instance via global lookup on window; keep simple: single instance
        private static ClipboardProbe _singleton; // safe: single add-in instance
        private static IntPtr WndProc(IntPtr hWnd, uint msg, UIntPtr wParam, IntPtr lParam)
        {
            var self = _singleton;
            if (self == null) return DefWindowProcW(hWnd, msg, wParam, lParam);

            if (msg == WM_CLIPBOARDUPDATE)
            {
                try { self.HandleClipboardUpdate(); } catch { }
                return IntPtr.Zero;
            }
            else if (msg == WM_DESTROY)
            {
                try { RemoveClipboardFormatListener(hWnd); } catch { }
                try { PostQuitMessage(0); } catch { }
                return IntPtr.Zero;
            }
            return DefWindowProcW(hWnd, msg, wParam, lParam);
        }

        // ctor sets singleton pointer
        public ClipboardProbe() { _singleton = this; }

        // ---- Core handler (ported from Program.cs, without Console I/O) ----

        private void HandleClipboardUpdate()
        {
            string utcIso = DateTime.UtcNow.ToString("o");

            BrowserKind kind = BrowserKind.Other;
            string procName = "";
            try
            {
                IntPtr owner = GetClipboardOwner();
                if (owner != IntPtr.Zero)
                {
                    GetWindowThreadProcessId(owner, out uint pid);
                    if (pid != 0)
                    {
                        try { procName = Process.GetProcessById((int)pid).ProcessName.ToLowerInvariant(); }
                        catch { procName = ""; }
                    }
                }
                if (procName.Contains("chrome") || procName.Contains("msedge") ||
                    procName.Contains("opera") || procName.Contains("brave") ||
                    procName.Contains("chromium"))
                    kind = BrowserKind.Chromium;
                else if (procName.Contains("firefox"))
                    kind = BrowserKind.Firefox;
            }
            catch { }

            if (!TryOpenClipboard()) return;

            string text = ReadUnicodeText(); // always
            string cfHtml = ReadCFHtmlRaw(); // may be empty
            string sourceUrlFromHtml = ExtractSourceUrlFromCFHtml(cfHtml);
            string chromiumInternalUrl = (kind == BrowserKind.Chromium) ? ReadChromiumInternalUrl() : "";

            CloseClipboard();

            string firefoxTitle = "";
            if (kind == BrowserKind.Firefox)
            {
                try
                {
                    IntPtr fg = GetForegroundWindow();
                    firefoxTitle = GetWindowTitle(fg);
                }
                catch { }
            }

            // Threshold
            if (string.IsNullOrEmpty(text) || text.Length < MIN_COPY_LEN) return;

            // De-dup across apps
            string dedupKey = string.Join("|",
                procName ?? "",
                text ?? "",
                sourceUrlFromHtml ?? "",
                (kind == BrowserKind.Chromium ? chromiumInternalUrl ?? "" : ""),
                (kind == BrowserKind.Firefox ? firefoxTitle ?? "" : "")
            );

            DateTime nowUtc = DateTime.UtcNow;
            if (dedupKey == _lastEventKey && (nowUtc - _lastEventUtc) <= _dedupWindow) return;
            _lastEventKey = dedupKey;
            _lastEventUtc = nowUtc;

            // Emit candidate
            var c = new ClipboardCandidate
            {
                Utc = DateTime.UtcNow,
                Process = procName ?? "",
                Text = text ?? "",
                SourceUrl = sourceUrlFromHtml ?? "",
                ChromiumUrl = chromiumInternalUrl ?? "",
                FirefoxTitle = firefoxTitle ?? ""
            };

            try { CandidateAvailable?.Invoke(c); } catch { }
        }

        // ---- Helpers (ported) ----

        private static bool TryOpenClipboard()
        {
            if (OpenClipboard(IntPtr.Zero)) return true;
            Thread.Sleep(30);
            return OpenClipboard(IntPtr.Zero);
        }

        private static string ReadUnicodeText()
        {
            try
            {
                IntPtr h = GetClipboardData(CF_UNICODETEXT);
                if (h == IntPtr.Zero) return "";
                IntPtr p = IntPtr.Zero;
                try
                {
                    p = GlobalLock(h);
                    if (p == IntPtr.Zero) return "";
                    string s = Marshal.PtrToStringUni(p);
                    return s ?? "";
                }
                finally
                {
                    if (p != IntPtr.Zero) GlobalUnlock(h);
                }
            }
            catch { return ""; }
        }

        private static string ReadCFHtmlRaw()
        {
            var bytes = ReadClipboardBytesByFormat("text/html", 256 * 1024);
            if (bytes.Length == 0)
                bytes = ReadClipboardBytesByFormat("HTML Format", 256 * 1024);
            return DecodePreferUtf8(bytes);
        }

        private static string ReadChromiumInternalUrl()
        {
            var bytes = ReadClipboardBytesByFormat("Chromium internal source URL", 8 * 1024);
            return DecodePreferUtf8(bytes);
        }

        private static byte[] ReadClipboardBytesByFormat(string formatName, int maxBytes)
        {
            try
            {
                ushort fmt = RegisterClipboardFormatW(formatName);
                if (fmt == 0) return new byte[0];
                IntPtr h = GetClipboardData(fmt);
                if (h == IntPtr.Zero) return new byte[0];

                IntPtr p = IntPtr.Zero;
                try
                {
                    p = GlobalLock(h);
                    if (p == IntPtr.Zero) return new byte[0];

                    var buf = new byte[maxBytes];
                    int i = 0;
                    for (; i < maxBytes; i++)
                    {
                        byte b = Marshal.ReadByte(p, i);
                        if (b == 0) break;
                        buf[i] = b;
                    }
                    if (i == 0) return new byte[0];
                    Array.Resize(ref buf, i);
                    return buf;
                }
                finally
                {
                    if (p != IntPtr.Zero) GlobalUnlock(h);
                }
            }
            catch { return new byte[0]; }
        }

        private static string DecodePreferUtf8(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) return "";
            try
            {
                var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
                string s = utf8.GetString(bytes);
                return s;
            }
            catch
            {
                try { return Encoding.Default.GetString(bytes); } catch { return ""; }
            }
        }

        private static string ExtractSourceUrlFromCFHtml(string cfHtml)
        {
            if (string.IsNullOrEmpty(cfHtml)) return "";
            const string marker = "SourceURL:";
            int idx = cfHtml.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return "";
            int start = idx + marker.Length;
            int end = start;
            while (end < cfHtml.Length && cfHtml[end] != '\r' && cfHtml[end] != '\n') end++;
            string url = cfHtml.Substring(start, end - start).Trim();
            return string.IsNullOrEmpty(url) ? "" : url;
        }

        private static string GetWindowTitle(IntPtr hwnd)
        {
            try
            {
                var sb = new StringBuilder(1024);
                int r = GetWindowTextW(hwnd, sb, sb.Capacity);
                return r > 0 ? sb.ToString() : "";
            }
            catch { return ""; }
        }

        // ---- P/Invoke ----

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern ushort RegisterClassExW(in WNDCLASSEX lpwcx);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateWindowExW(
            uint dwExStyle, string lpClassName, string lpWindowName, uint dwStyle,
            int X, int Y, int nWidth, int nHeight,
            IntPtr hWndParent, IntPtr hMenu, IntPtr hInstance, IntPtr lpParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr DefWindowProcW(IntPtr hWnd, uint msg, UIntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern sbyte GetMessageW(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport("user32.dll")]
        private static extern bool TranslateMessage(in MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern IntPtr DispatchMessageW(in MSG lpMsg);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool CloseClipboard();

        [DllImport("user32.dll")]
        private static extern IntPtr GetClipboardOwner();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern ushort RegisterClipboardFormatW(string lpszFormat);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandleW(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern void PostQuitMessage(int nExitCode);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool PostMessageW(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WNDCLASSEX
        {
            public uint cbSize;
            public uint style;
            public IntPtr lpfnWndProc;
            public int cbClsExtra;
            public int cbWndExtra;
            public IntPtr hInstance;
            public IntPtr hIcon;
            public IntPtr hCursor;
            public IntPtr hbrBackground;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpszMenuName;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpszClassName;
            public IntPtr hIconSm;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public UIntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public int pt_x;
            public int pt_y;
            public uint lPrivate;
        }
    }
}
