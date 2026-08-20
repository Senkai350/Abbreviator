using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using Abbreviator.Interop;
using Abbreviator.UI;

namespace Abbreviator.Core
{
    /// <summary>
    /// Живая подсветка: волнистые подчёркивания поверх окна Word, как у
    /// проверки орфографии. Документ не изменяется вообще.
    ///
    /// Таймер берёт видимый на экране диапазон текста
    /// (Window.RangeFromPoint по сетке точек панели документа), находит в нём
    /// сокращения и для каждого запрашивает экранный прямоугольник
    /// (Window.GetPoint). Событий прокрутки объектная модель Word не даёт,
    /// поэтому опрос. Кадр перерисовывается только когда изменилась «подпись»
    /// состояния — видимый диапазон, геометрия панели и позиция первой видимой
    /// строки; при неподвижном экране работы не делается вовсе.
    ///
    /// Тот же таймер отслеживает смену активного документа и запускает
    /// автоопределение листов перечня, поэтому он работает и при выключенной
    /// подсветке.
    /// </summary>
    public sealed class LiveCheck : IDisposable
    {
        private const int MaxVisibleChars = 30000;
        private const int MaxMarks = 250;
        private const int ExpandChars = 80;

        private readonly AddInController _controller;
        private Timer _timer;
        private OverlayForm _overlay;

        private string _signature;
        private int _errorCount;

        private AbbreviationRecognizer _recognizer;

        public LiveCheck(AddInController controller)
        {
            _controller = controller;
        }

        public void Start()
        {
            if (_timer == null)
            {
                _timer = new Timer();
                _timer.Tick += OnTick;
            }
            _timer.Interval = Math.Max(60, _controller.Settings.LiveIntervalMs);
            _timer.Start();
        }

        /// <summary>Спрятать подчёркивания, не останавливая слежение за документом.</summary>
        public void HideMarks()
        {
            _signature = null;
            if (_overlay != null && !_overlay.IsDisposed) _overlay.HideOverlay();
        }

        /// <summary>Сбросить кэши: перечень, настройки или содержимое изменились.</summary>
        public void Invalidate()
        {
            _signature = null;
            _recognizer = null;
            if (_timer != null) _timer.Interval = Math.Max(60, _controller.Settings.LiveIntervalMs);
        }

        public void Dispose()
        {
            try
            {
                if (_timer != null)
                {
                    _timer.Stop();
                    _timer.Dispose();
                    _timer = null;
                }
                if (_overlay != null)
                {
                    _overlay.Close();
                    _overlay.Dispose();
                    _overlay = null;
                }
            }
            catch { }
        }

        // ==================================================================

        private void OnTick(object sender, EventArgs e)
        {
            try
            {
                _controller.PollActiveDocument();

                if (_controller.Settings.LiveHighlight) Tick();
                else HideMarks();

                _errorCount = 0;
            }
            catch (Exception ex)
            {
                // Первые сбои пишем в журнал, дальше молчим, чтобы не раздуть его.
                if (++_errorCount <= 3) Diag.Error("живая подсветка", ex);
                HideMarks();
            }
        }

        private void Tick()
        {
            dynamic app = _controller.App;
            dynamic doc = _controller.ActiveDocument;
            if (app == null || doc == null) { HideMarks(); return; }

            dynamic window = app.ActiveWindow;
            if (window == null) { HideMarks(); return; }

            // Рисуем, только пока активно окно нашего процесса. Сравнение с
            // конкретным HWND здесь не годится: оверлей — отдельное окно, и
            // такая проверка через тик давала бы «Word не активен», отчего
            // подчёркивания мигали.
            if (!NativeMethods.ForegroundBelongsToThisProcess()) { HideMarks(); return; }

            IntPtr wordHwnd = GetWindowHandle(window);
            if (wordHwnd == IntPtr.Zero || NativeMethods.IsIconic(wordHwnd)) { HideMarks(); return; }

            NativeMethods.RECT paneRect;
            if (!NativeMethods.TryGetDocumentPane(wordHwnd, out paneRect)) { HideMarks(); return; }
            Rectangle pane = ToRectangle(paneRect);
            if (pane.Width <= 0 || pane.Height <= 0) { HideMarks(); return; }

            int contentEnd = WordUtil.ContentEnd(doc);

            int visStart, visEnd;
            if (!TryGetVisibleRange(window, paneRect, contentEnd, out visStart, out visEnd))
            {
                HideMarks();
                return;
            }

            // Подпись состояния. Прямоугольник первой видимой строки ловит
            // прокрутку даже на пару пикселей, когда первый видимый символ
            // остался прежним.
            Rectangle anchor;
            TryGetScreenRect(window, doc, visStart, Math.Min(visStart + 1, contentEnd), out anchor);

            string signature = string.Join("|", new[]
            {
                visStart.ToString(), visEnd.ToString(), contentEnd.ToString(),
                pane.ToString(), anchor.ToString()
            });

            if (signature == _signature) return;   // экран неподвижен — работы нет
            _signature = signature;

            string text = WordUtil.RangeText(doc, visStart, visEnd);
            if (text.Length == 0 || text.Length > MaxVisibleChars) { HideMarks(); return; }

            var settings = _controller.Settings;
            DocumentState state = _controller.GetState(doc);
            AddInController.DocIndex index = _controller.GetIndex(doc, state);
            AbbreviationRecognizer recognizer = GetRecognizer(app);

            var marks = new List<OverlayForm.Mark>();

            foreach (var token in recognizer.Find(text))
            {
                if (marks.Count >= MaxMarks) break;

                int pos = visStart + token.Index;

                if (state.IsIgnored(token.Abbr)) continue;
                if (settings.SkipTableOfContents && index.Toc.IsInToc(pos)) continue;
                if (!settings.HighlightDictionarySection && InRanges(index.DictionaryRanges, pos)) continue;

                bool known = index.Known.ContainsKey(token.Abbr);
                if (known && !settings.LiveShowKnown) continue;

                Rectangle rect;
                if (!TryGetScreenRect(window, doc, pos, pos + token.Length, out rect)) continue;
                if (!rect.IntersectsWith(pane)) continue;

                marks.Add(new OverlayForm.Mark { Rect = rect, Known = known });
            }

            EnsureOverlay(wordHwnd).Render(pane, marks);
        }

        private OverlayForm EnsureOverlay(IntPtr wordHwnd)
        {
            if (_overlay == null || _overlay.IsDisposed) _overlay = new OverlayForm();
            _overlay.AttachTo(wordHwnd);
            return _overlay;
        }

        private static IntPtr GetWindowHandle(dynamic window)
        {
            try { return new IntPtr(Convert.ToInt64(window.Hwnd)); }
            catch { return NativeMethods.GetForegroundWindow(); }
        }

        // ==================================================================
        // Видимый диапазон
        // ==================================================================

        /// <summary>
        /// Диапазон текста, попадающего в панель документа: пробуем сетку точек
        /// и берём минимальный Start и максимальный End из удачных попаданий.
        /// RangeFromPoint бросает исключение на точках вне текста — это норма.
        /// </summary>
        private static bool TryGetVisibleRange(dynamic window, NativeMethods.RECT pane,
                                               int contentEnd, out int start, out int end)
        {
            start = int.MaxValue;
            end = -1;

            int[] xs = { pane.Left + pane.Width / 10, pane.Left + pane.Width / 2, pane.Right - pane.Width / 10 };
            int[] ys = { pane.Top + 8, pane.Top + pane.Height / 4, pane.Top + pane.Height / 2,
                         pane.Top + 3 * pane.Height / 4, pane.Bottom - 8 };

            foreach (int y in ys)
            {
                foreach (int x in xs)
                {
                    try
                    {
                        dynamic r = window.RangeFromPoint(x, y);
                        if (r == null) continue;
                        int s = (int)r.Start;
                        int e = (int)r.End;
                        if (s < start) start = s;
                        if (e > end) end = e;
                    }
                    catch { }
                }
            }

            if (end < 0 || start == int.MaxValue) return false;

            // Расширяем диапазон, чтобы не потерять токен, разрезанный краем экрана.
            start = Math.Max(0, start - ExpandChars);
            end = Math.Min(contentEnd, end + ExpandChars);
            return end > start;
        }

        // ==================================================================
        // Экранные координаты диапазона
        // ==================================================================

        /// <summary>
        /// Window.GetPoint возвращает прямоугольник диапазона в пикселях экрана.
        /// У метода четыре out-параметра, поэтому при позднем связывании он
        /// вызывается через InvokeMember с ParameterModifier.
        /// </summary>
        private static bool TryGetScreenRect(dynamic window, dynamic doc,
                                             int start, int end, out Rectangle rect)
        {
            rect = Rectangle.Empty;
            if (end <= start) return false;

            try
            {
                object range = doc.Range(start, end);
                object[] args = { 0, 0, 0, 0, range };
                var byRef = new ParameterModifier(5);
                byRef[0] = byRef[1] = byRef[2] = byRef[3] = true;

                object win = window;
                win.GetType().InvokeMember(
                    "GetPoint",
                    BindingFlags.InvokeMethod,
                    null, win, args, new[] { byRef }, null, null);

                int left = Convert.ToInt32(args[0]);
                int top = Convert.ToInt32(args[1]);
                int width = Convert.ToInt32(args[2]);
                int height = Convert.ToInt32(args[3]);

                if (width <= 0 || height <= 0 || height > 200) return false;

                rect = new Rectangle(left, top, width, height);
                return true;
            }
            catch
            {
                return false;
            }
        }

        // ==================================================================

        private static bool InRanges(List<int[]> ranges, int position)
        {
            foreach (var r in ranges)
                if (position >= r[0] && position < r[1]) return true;
            return false;
        }

        private AbbreviationRecognizer GetRecognizer(dynamic app)
        {
            if (_recognizer == null)
                _recognizer = new DocumentScanner(_controller.Settings).CreateRecognizer(app);
            return _recognizer;
        }

        private static Rectangle ToRectangle(NativeMethods.RECT r)
        {
            return new Rectangle(r.Left, r.Top, r.Width, r.Height);
        }
    }
}
