// Program.cs
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

internal static class Program
{
    private const int WM_CLIPBOARDUPDATE = 0x031D;
    private const int WM_DESTROY = 0x0002;
    private const uint CF_UNICODETEXT = 13;

    // Win32 P/Invoke
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

    // Win32 structs
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

    // Keep wndproc delegate alive
    private static readonly WndProcDelegate WndProcThunk = WndProc;
    private delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, UIntPtr wParam, IntPtr lParam);

    private enum BrowserKind { Chromium, Firefox, Other }
    // ---- De-dup + threshold (global) ----
    private const int MIN_COPY_LEN = 20;
    private static string _lastEventKey = "";
    private static DateTime _lastEventUtc = DateTime.MinValue;
    private static readonly TimeSpan _dedupWindow = TimeSpan.FromMilliseconds(800);

    [STAThread]
    private static void Main()
    {
        Console.OutputEncoding = Encoding.UTF8;
        Console.WriteLine("Minimal clipboard listener (UTC timestamp, text, URLs). Ctrl+C to exit.");

        var wc = new WNDCLASSEX
        {
            cbSize = (uint)Marshal.SizeOf<WNDCLASSEX>(),
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

        ushort atom = RegisterClassExW(in wc);
        if (atom == 0)
        {
            Console.WriteLine("RegisterClassExW failed: " + Marshal.GetLastWin32Error());
            return;
        }

        IntPtr hwnd = CreateWindowExW(0, wc.lpszClassName, "ClipLiteHidden", 0, 0, 0, 0, 0,
                                      IntPtr.Zero, IntPtr.Zero, wc.hInstance, IntPtr.Zero);
        if (hwnd == IntPtr.Zero)
        {
            Console.WriteLine("CreateWindowExW failed: " + Marshal.GetLastWin32Error());
            return;
        }

        if (!AddClipboardFormatListener(hwnd))
        {
            Console.WriteLine("AddClipboardFormatListener failed: " + Marshal.GetLastWin32Error());
            DestroyWindow(hwnd);
            return;
        }

        Console.CancelKeyPress += (s, e) =>
        {
            RemoveClipboardFormatListener(hwnd);
            DestroyWindow(hwnd);
        };

        // Message loop
        MSG msg;
        while (GetMessageW(out msg, IntPtr.Zero, 0, 0) > 0)
        {
            TranslateMessage(in msg);
            DispatchMessageW(in msg);
        }
    }

    private static IntPtr WndProc(IntPtr hWnd, uint msg, UIntPtr wParam, IntPtr lParam)
    {
        if (msg == WM_CLIPBOARDUPDATE)
        {
            try { HandleClipboardUpdate(); }
            catch (Exception ex) { Console.WriteLine("Error: " + ex.Message); }
            return IntPtr.Zero;
        }
        else if (msg == WM_DESTROY)
        {
            RemoveClipboardFormatListener(hWnd);
            PostQuitMessage(0);
            return IntPtr.Zero;
        }
        return DefWindowProcW(hWnd, msg, wParam, lParam);
    }

    private static void HandleClipboardUpdate()
    {
        string utc = DateTime.UtcNow.ToString("o");

        // Determine browser kind from clipboard owner process
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
        catch { /* ignore */ }

        // Open clipboard
        if (!TryOpenClipboard())
        {
            // If busy, skip this event
            return;
        }

        string text = ReadUnicodeText();                           // always
        string cfHtml = ReadCFHtmlRaw();                           // may be empty
        string sourceUrlFromHtml = ExtractSourceUrlFromCFHtml(cfHtml);
        string chromiumInternalUrl = (kind == BrowserKind.Chromium) ? ReadChromiumInternalUrl() : "";

        CloseClipboard();

        // For Firefox, grab foreground window title
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
        // ---- Threshold filter (<20 chars) ----
        if (string.IsNullOrEmpty(text) || text.Length < MIN_COPY_LEN)
        {
            CloseClipboard();
            return; // ignore tiny copies by design
        }

        // ---- Global de-dup (all apps, incl. Word close) ----
        // Key includes process, text, both URL sources, and window title if Firefox
        string dedupKey = string.Join("|",
            procName ?? "",
            text,
            sourceUrlFromHtml ?? "",
            (kind == BrowserKind.Chromium ? chromiumInternalUrl ?? "" : ""),
            (kind == BrowserKind.Firefox ? firefoxTitle ?? "" : "")
        );

        DateTime nowUtc = DateTime.UtcNow;
        if (dedupKey == _lastEventKey && (nowUtc - _lastEventUtc) <= _dedupWindow)
        {
            CloseClipboard();
            return; // drop duplicate burst
        }
        _lastEventKey = dedupKey;
        _lastEventUtc = nowUtc;


        // ---- Output (only what’s required) ----
        Console.WriteLine("TimeUTC: " + utc);
        Console.WriteLine("Text: " + Quote(text));

        if (kind == BrowserKind.Chromium)
        {
            Console.WriteLine("CF_HTML SourceURL: " + (string.IsNullOrEmpty(sourceUrlFromHtml) ? "(none)" : sourceUrlFromHtml));
            Console.WriteLine("ChromiumURL: " + (string.IsNullOrEmpty(chromiumInternalUrl) ? "(none)" : chromiumInternalUrl));
        }
        else if (kind == BrowserKind.Firefox)
        {
            Console.WriteLine("CF_HTML SourceURL: " + (string.IsNullOrEmpty(sourceUrlFromHtml) ? "(none)" : sourceUrlFromHtml));
            Console.WriteLine("WindowTitle: " + (string.IsNullOrEmpty(firefoxTitle) ? "(none)" : firefoxTitle));
        }
        else
        {
            // Other apps: still print CF_HTML SourceURL if present
            if (!string.IsNullOrEmpty(sourceUrlFromHtml))
                Console.WriteLine("CF_HTML SourceURL: " + sourceUrlFromHtml);
        }

        Console.WriteLine(); // blank line between events
    }

    // ---- Helpers ----
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
        // Try modern "text/html" first (often UTF-8), then legacy "HTML Format"
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
            if (fmt == 0) return Array.Empty<byte>();
            IntPtr h = GetClipboardData(fmt);
            if (h == IntPtr.Zero) return Array.Empty<byte>();

            IntPtr p = IntPtr.Zero;
            try
            {
                p = GlobalLock(h);
                if (p == IntPtr.Zero) return Array.Empty<byte>();

                // read up to maxBytes or until NUL byte
                var buf = new byte[maxBytes];
                int i = 0;
                for (; i < maxBytes; i++)
                {
                    byte b = Marshal.ReadByte(p, i);
                    if (b == 0) break;
                    buf[i] = b;
                }
                if (i == 0) return Array.Empty<byte>();
                Array.Resize(ref buf, i);
                return buf;
            }
            finally
            {
                if (p != IntPtr.Zero) GlobalUnlock(h);
            }
        }
        catch { return Array.Empty<byte>(); }
    }

    private static string DecodePreferUtf8(byte[] bytes)
    {
        if (bytes == null || bytes.Length == 0) return "";
        try
        {
            // Try UTF-8 (strict). If it round-trips, keep it; else fall back.
            var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
            string s = utf8.GetString(bytes);
            return s;
        }
        catch
        {
            // Fallback to system ANSI (matches CF_HTML legacy behavior on Windows)
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

    private static string Quote(string s)
    {
        if (s == null) return "\"\"";
        // show as one line; trim long tail for console
        var one = s.Replace("\r", "\\r").Replace("\n", "\\n");
        if (one.Length > 500) one = one.Substring(0, 500) + "...";
        return "\"" + one + "\"";
    }
}
