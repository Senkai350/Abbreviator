using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Abbreviator.Interop
{
    /// <summary>Win32-вызовы для режима живой подсветки (оверлея).</summary>
    internal static class NativeMethods
    {
        // ------------------------------------------------------------------
        // Структуры
        // ------------------------------------------------------------------

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            public int Width { get { return Right - Left; } }
            public int Height { get { return Bottom - Top; } }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y) { X = x; Y = y; }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct SIZE
        {
            public int Cx;
            public int Cy;

            public SIZE(int cx, int cy) { Cx = cx; Cy = cy; }
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct BLENDFUNCTION
        {
            public byte BlendOp;
            public byte BlendFlags;
            public byte SourceConstantAlpha;
            public byte AlphaFormat;
        }

        // ------------------------------------------------------------------
        // Константы
        // ------------------------------------------------------------------

        public const int GWL_HWNDPARENT = -8;
        public const int GA_ROOT = 2;

        public const int SW_SHOWNOACTIVATE = 4;
        public const int SW_HIDE = 0;

        public const int ULW_ALPHA = 0x00000002;
        public const byte AC_SRC_OVER = 0x00;
        public const byte AC_SRC_ALPHA = 0x01;

        // ------------------------------------------------------------------
        // user32
        // ------------------------------------------------------------------

        public delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int processId);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);

        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr hwnd, EnumChildProc callback, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hwnd, StringBuilder buffer, int maxCount);

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hwnd, int flags);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hwnd, int cmdShow);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UpdateLayeredWindow(
            IntPtr hwnd, IntPtr hdcDst, ref POINT pptDst, ref SIZE psize,
            IntPtr hdcSrc, ref POINT pptSrc, int crKey,
            ref BLENDFUNCTION pblend, int dwFlags);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongW")]
        private static extern int SetWindowLong32(IntPtr hwnd, int index, int value);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hwnd, int index, IntPtr value);

        /// <summary>SetWindowLongPtr есть только в 64-разрядной user32.</summary>
        public static IntPtr SetWindowLongPtr(IntPtr hwnd, int index, IntPtr value)
        {
            if (IntPtr.Size == 8) return SetWindowLongPtr64(hwnd, index, value);
            return (IntPtr)SetWindowLong32(hwnd, index, value.ToInt32());
        }

        // ------------------------------------------------------------------
        // gdi32
        // ------------------------------------------------------------------

        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        public static extern bool DeleteDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        public static extern IntPtr SelectObject(IntPtr hdc, IntPtr handle);

        [DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr handle);

        // ------------------------------------------------------------------
        // Помощники
        // ------------------------------------------------------------------

        /// <summary>
        /// Окно на переднем плане принадлежит нашему процессу (то есть Word,
        /// его диалогам или нашему же оверлею).
        ///
        /// Сравнивать с конкретным HWND нельзя: оверлей — самостоятельное
        /// окно, и такая проверка давала бы «Word не активен» через тик,
        /// отчего подчёркивания мигали.
        /// </summary>
        public static bool ForegroundBelongsToThisProcess()
        {
            IntPtr foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero) return false;

            int pid;
            GetWindowThreadProcessId(foreground, out pid);
            return pid == System.Diagnostics.Process.GetCurrentProcess().Id;
        }

        /// <summary>
        /// Область документа в окне Word — дочернее окно класса "_WwG".
        /// При разделённом окне таких окон два; берём самое большое видимое.
        /// </summary>
        public static bool TryGetDocumentPane(IntPtr wordWindow, out RECT pane)
        {
            RECT best = new RECT();
            long bestArea = 0;
            var buffer = new StringBuilder(64);

            EnumChildProc callback = delegate(IntPtr child, IntPtr lp)
            {
                buffer.Length = 0;
                if (GetClassName(child, buffer, buffer.Capacity) > 0 &&
                    buffer.ToString() == "_WwG" &&
                    IsWindowVisible(child))
                {
                    RECT r;
                    if (GetWindowRect(child, out r))
                    {
                        long area = (long)r.Width * r.Height;
                        if (area > bestArea)
                        {
                            bestArea = area;
                            best = r;
                        }
                    }
                }
                return true;
            };

            EnumChildWindows(wordWindow, callback, IntPtr.Zero);
            GC.KeepAlive(callback);

            pane = best;
            return bestArea > 0;
        }
    }
}
