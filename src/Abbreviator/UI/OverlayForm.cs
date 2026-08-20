using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Abbreviator.Interop;

namespace Abbreviator.UI
{
    /// <summary>
    /// Прозрачное окно поверх окна Word: волнистые подчёркивания рисуются на
    /// нём, а документ не изменяется вовсе — так же, как Word рисует
    /// орфографию.
    ///
    /// Окно слоёное (WS_EX_LAYERED) и обновляется через UpdateLayeredWindow:
    /// кадр отдаётся системе целиком, без WM_PAINT и без затирания фона,
    /// поэтому мерцания нет в принципе. Прозрачность — попиксельная (альфа),
    /// а не по цветовому ключу, так что линии выходят сглаженными.
    ///
    /// WS_EX_TRANSPARENT | WS_EX_NOACTIVATE: окно не принимает ни фокус, ни
    /// клики — мышь и клавиатура проходят сквозь него в Word, поэтому правый
    /// клик по сокращению работает как обычно.
    ///
    /// Владельцем окна назначается окно Word (GWL_HWNDPARENT). Благодаря
    /// этому оверлей всегда над Word, сворачивается вместе с ним и не
    /// перекрывает другие приложения — TopMost для этого не нужен.
    /// </summary>
    public sealed class OverlayForm : Form
    {
        public struct Mark
        {
            public Rectangle Rect;
            public bool Known;
        }

        private const int WS_EX_TRANSPARENT = 0x00000020;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_LAYERED = 0x00080000;
        private const int WS_EX_NOACTIVATE = 0x08000000;

        private readonly Pen _knownPen;
        private readonly Pen _unknownPen;

        private Bitmap _frame;
        private IntPtr _owner = IntPtr.Zero;
        private bool _shown;

        public OverlayForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            Visible = false;

            // Цвета непрозрачные, сглаживание выключено: UpdateLayeredWindow
            // ждёт альфу, умноженную на цвет, а GetHbitmap её не умножает.
            // При полностью прозрачных и полностью непрозрачных пикселях
            // умножение — тождество, поэтому ореола вокруг линий не будет.
            _knownPen = new Pen(Color.FromArgb(255, 40, 150, 60), 1f);
            _unknownPen = new Pen(Color.FromArgb(255, 214, 48, 34), 1f);
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;
                return cp;
            }
        }

        /// <summary>Привязать оверлей к окну Word как к окну-владельцу.</summary>
        public void AttachTo(IntPtr wordWindow)
        {
            if (wordWindow == _owner) return;
            _owner = wordWindow;
            NativeMethods.SetWindowLongPtr(Handle, NativeMethods.GWL_HWNDPARENT, wordWindow);
        }

        /// <summary>
        /// Отрисовать кадр: положение окна — панель документа Word,
        /// подчёркивания заданы в экранных координатах.
        /// </summary>
        public void Render(Rectangle pane, List<Mark> marks)
        {
            if (pane.Width <= 0 || pane.Height <= 0) { HideOverlay(); return; }

            if (_frame == null || _frame.Width != pane.Width || _frame.Height != pane.Height)
            {
                if (_frame != null) _frame.Dispose();
                _frame = new Bitmap(pane.Width, pane.Height, PixelFormat.Format32bppArgb);
            }

            using (var g = Graphics.FromImage(_frame))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.None;

                foreach (var m in marks)
                {
                    var r = m.Rect;
                    r.Offset(-pane.X, -pane.Y);           // экран -> холст кадра
                    DrawWave(g, r, m.Known ? _knownPen : _unknownPen);
                }
            }

            Push(pane);
        }

        public void HideOverlay()
        {
            if (!_shown) return;
            _shown = false;
            if (IsHandleCreated) NativeMethods.ShowWindow(Handle, NativeMethods.SW_HIDE);
        }

        // ------------------------------------------------------------------

        /// <summary>Отдать готовый кадр системе одним вызовом.</summary>
        private void Push(Rectangle pane)
        {
            IntPtr screenDc = NativeMethods.GetDC(IntPtr.Zero);
            IntPtr memDc = NativeMethods.CreateCompatibleDC(screenDc);
            IntPtr bitmap = IntPtr.Zero;
            IntPtr previous = IntPtr.Zero;

            try
            {
                bitmap = _frame.GetHbitmap(Color.FromArgb(0));
                previous = NativeMethods.SelectObject(memDc, bitmap);

                var size = new NativeMethods.SIZE(pane.Width, pane.Height);
                var destination = new NativeMethods.POINT(pane.X, pane.Y);
                var source = new NativeMethods.POINT(0, 0);
                var blend = new NativeMethods.BLENDFUNCTION
                {
                    BlendOp = NativeMethods.AC_SRC_OVER,
                    BlendFlags = 0,
                    SourceConstantAlpha = 255,
                    AlphaFormat = NativeMethods.AC_SRC_ALPHA
                };

                NativeMethods.UpdateLayeredWindow(Handle, screenDc, ref destination, ref size,
                                                  memDc, ref source, 0, ref blend,
                                                  NativeMethods.ULW_ALPHA);

                if (!_shown)
                {
                    _shown = true;
                    NativeMethods.ShowWindow(Handle, NativeMethods.SW_SHOWNOACTIVATE);
                }
            }
            finally
            {
                NativeMethods.ReleaseDC(IntPtr.Zero, screenDc);
                if (bitmap != IntPtr.Zero)
                {
                    NativeMethods.SelectObject(memDc, previous);
                    NativeMethods.DeleteObject(bitmap);
                }
                NativeMethods.DeleteDC(memDc);
            }
        }

        /// <summary>Волнистая линия под нижней кромкой текста.</summary>
        private static void DrawWave(Graphics g, Rectangle rect, Pen pen)
        {
            if (rect.Width < 4) return;

            int y = rect.Bottom - 1;
            int x = rect.Left;
            int end = rect.Right;

            var points = new List<Point>();
            bool up = true;
            while (x <= end)
            {
                points.Add(new Point(x, up ? y - 2 : y));
                up = !up;
                x += 3;
            }

            if (points.Count >= 2) g.DrawLines(pen, points.ToArray());
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _knownPen.Dispose();
                _unknownPen.Dispose();
                if (_frame != null) _frame.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
