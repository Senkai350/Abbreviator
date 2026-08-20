using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace Abbreviator.Interop
{
    /// <summary>Win32-вызовы для режима живой подсветки (оверлея).</summary>
    internal static class NativeMethods
    {
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

        public delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

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

        public const int GA_ROOT = 2;

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
