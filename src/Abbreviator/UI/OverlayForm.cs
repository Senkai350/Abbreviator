using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Abbreviator.UI
{
    /// <summary>
    /// Прозрачное окно поверх окна Word, на котором рисуются волнистые
    /// подчёркивания — так же, как Word рисует орфографию: документ при этом
    /// не меняется вовсе.
    ///
    /// Окно не принимает ни фокус, ни клики (WS_EX_TRANSPARENT |
    /// WS_EX_NOACTIVATE): мышь и клавиатура проходят сквозь него в Word,
    /// поэтому правый клик по сокращению работает как обычно.
    /// </summary>
    public sealed class OverlayForm : Form
    {
        public struct Mark
        {
            public Rectangle Rect;
            public bool Known;
        }

        private List<Mark> _marks = new List<Mark>();
        private Rectangle _clip;

        private readonly Pen _knownPen;
        private readonly Pen _unknownPen;

        private const int WS_EX_TRANSPARENT = 0x00000020;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_NOACTIVATE = 0x08000000;

        public OverlayForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;

            // Цветовой ключ: фон этого цвета полностью прозрачен,
            // непрозрачны только сами подчёркивания.
            BackColor = Color.Magenta;
            TransparencyKey = Color.Magenta;

            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);

            _knownPen = new Pen(Color.FromArgb(46, 155, 66), 1.6f);
            _unknownPen = new Pen(Color.FromArgb(214, 56, 42), 1.6f);
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
                cp.ExStyle |= WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;
                return cp;
            }
        }

        /// <summary>
        /// Обновляет оверлей: положение поверх окна Word, область отсечения
        /// (панель документа) и список подчёркиваний в экранных координатах.
        /// </summary>
        public void UpdateMarks(Rectangle wordWindow, Rectangle documentPane, List<Mark> marks)
        {
            if (Bounds != wordWindow) Bounds = wordWindow;

            // Экранные координаты -> координаты клиентской области оверлея.
            var clip = documentPane;
            clip.Offset(-wordWindow.X, -wordWindow.Y);

            var local = new List<Mark>(marks.Count);
            foreach (var m in marks)
            {
                var r = m.Rect;
                r.Offset(-wordWindow.X, -wordWindow.Y);
                if (r.IntersectsWith(clip))
                    local.Add(new Mark { Rect = r, Known = m.Known });
            }

            bool changed = clip != _clip || !SameMarks(local, _marks);
            _clip = clip;
            _marks = local;

            if (!Visible) Show();
            if (changed) Invalidate();
        }

        public void HideOverlay()
        {
            if (Visible) Hide();
        }

        private static bool SameMarks(List<Mark> a, List<Mark> b)
        {
            if (a.Count != b.Count) return false;
            for (int i = 0; i < a.Count; i++)
                if (a[i].Rect != b[i].Rect || a[i].Known != b[i].Known) return false;
            return true;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (_marks.Count == 0) return;

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.SetClip(_clip);

            foreach (var m in _marks)
                DrawWave(e.Graphics, m.Rect, m.Known ? _knownPen : _unknownPen);
        }

        /// <summary>Волнистая линия под нижней кромкой прямоугольника текста.</summary>
        private static void DrawWave(Graphics g, Rectangle rect, Pen pen)
        {
            if (rect.Width < 4) return;

            int y = rect.Bottom - 1;
            int x = rect.Left;
            int end = rect.Right;

            // Зигзаг с шагом 3 и амплитудой 2 — визуально как у проверки орфографии.
            var points = new List<Point>();
            bool up = true;
            while (x <= end)
            {
                points.Add(new Point(x, up ? y - 2 : y));
                up = !up;
                x += 3;
            }

            if (points.Count >= 2)
                g.DrawLines(pen, points.ToArray());
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _knownPen.Dispose();
                _unknownPen.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
