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
    /// Раз в 250 мс сервис берёт видимый на экране диапазон текста
    /// (Window.RangeFromPoint по нескольким точкам панели документа), находит
    /// в нём сокращения и для каждого запрашивает экранный прямоугольник
    /// (Window.GetPoint). Прямоугольники отдаются оверлею. При прокрутке,
    /// смене масштаба или правке текста картинка обновляется на следующем
    /// тике сама — Word не шлёт событий прокрутки, поэтому именно опрос.
    ///
    /// Все вызовы COM здесь внутрипроцессные и идут в потоке Word (таймер
    /// WinForms работает на его цикле сообщений), так что 250 мс хватает
    /// с большим запасом.
    /// </summary>
    public sealed class LiveCheck : IDisposable
    {
        private const int IntervalMs = 250;
        private const int MaxVisibleChars = 30000;
        private const int MaxMarks = 300;
        private const int ExpandChars = 80;

        private readonly AddInController _controller;
        private Timer _timer;
        private OverlayForm _overlay;

        // Кэш перечня: перечитывать его на каждом тике нельзя, а во время
        // набора текста длина документа меняется каждый тик. Поэтому индекс
        // перестраивается только когда длина документа стабильна два тика
        // подряд (пользователь сделал паузу).
        private Dictionary<string, DictEntry> _known =
            new Dictionary<string, DictEntry>(StringComparer.Ordinal);
        private List<int[]> _dictRanges = new List<int[]>();
        private int _knownContentEnd = -1;
        private int _lastTickContentEnd = -1;
        private string _knownDocKey;

        private AbbreviationRecognizer _recognizer;
        private int _errorCount;

        public LiveCheck(AddInController controller)
        {
            _controller = controller;
        }

        public bool Enabled
        {
            get { return _timer != null && _timer.Enabled; }
        }

        public void Start()
        {
            if (_timer == null)
            {
                _timer = new Timer { Interval = IntervalMs };
                _timer.Tick += OnTick;
            }
            _timer.Start();
        }

        public void Stop()
        {
            if (_timer != null) _timer.Stop();
            if (_overlay != null) _overlay.HideOverlay();
        }

        /// <summary>Сбросить кэши: перечень или настройки изменились.</summary>
        public void Invalidate()
        {
            _knownContentEnd = -1;
            _lastTickContentEnd = -1;
            _recognizer = null;
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
                Tick();
                _errorCount = 0;
            }
            catch (Exception ex)
            {
                // Первые сбои пишем в журнал, дальше молчим, чтобы не раздуть его.
                if (++_errorCount <= 3) Diag.Error("живая подсветка", ex);
                HideOverlay();
            }
        }

        private void Tick()
        {
            dynamic app = _controller.App;
            dynamic doc = _controller.ActiveDocument;
            if (app == null || doc == null) { HideOverlay(); return; }

            dynamic window = app.ActiveWindow;
            if (window == null) { HideOverlay(); return; }

            // Рисуем только когда окно Word на переднем плане: иначе оверлей
            // висел бы поверх других приложений и наших же диалогов.
            IntPtr wordHwnd = (IntPtr)(int)window.Hwnd;
            IntPtr foreground = NativeMethods.GetForegroundWindow();
            if (foreground != wordHwnd &&
                NativeMethods.GetAncestor(foreground, NativeMethods.GA_ROOT) != wordHwnd)
            {
                HideOverlay();
                return;
            }
            if (NativeMethods.IsIconic(wordHwnd)) { HideOverlay(); return; }

            NativeMethods.RECT windowRect;
            if (!NativeMethods.GetWindowRect(wordHwnd, out windowRect)) { HideOverlay(); return; }

            NativeMethods.RECT paneRect;
            if (!NativeMethods.TryGetDocumentPane(wordHwnd, out paneRect))
                paneRect = windowRect;

            // --- видимый диапазон текста -----------------------------------
            int contentEnd = WordUtil.ContentEnd(doc);
            int visStart, visEnd;
            if (!TryGetVisibleRange(window, paneRect, contentEnd, out visStart, out visEnd))
            {
                HideOverlay();
                return;
            }

            RefreshKnownIndex(doc, contentEnd);

            string text = WordUtil.RangeText(doc, visStart, visEnd);
            if (text.Length == 0 || text.Length > MaxVisibleChars)
            {
                HideOverlay();
                return;
            }

            // --- сокращения и их экранные прямоугольники -------------------
            var settings = _controller.Settings;
            var state = _controller.GetState(doc);
            var recognizer = GetRecognizer(app);

            var marks = new List<OverlayForm.Mark>();
            var pane = ToRectangle(paneRect);

            foreach (var token in recognizer.Find(text))
            {
                if (marks.Count >= MaxMarks) break;

                int pos = visStart + token.Index;

                if (state.IsIgnored(token.Abbr)) continue;

                bool known = _known.ContainsKey(token.Abbr);
                if (known && !settings.LiveShowKnown) continue;
                if (!settings.HighlightDictionarySection && InDictionary(pos)) continue;

                Rectangle rect;
                if (!TryGetScreenRect(window, doc, pos, pos + token.Length, out rect)) continue;
                if (!rect.IntersectsWith(pane)) continue;

                marks.Add(new OverlayForm.Mark { Rect = rect, Known = known });
            }

            if (_overlay == null || _overlay.IsDisposed) _overlay = new OverlayForm();
            _overlay.UpdateMarks(ToRectangle(windowRect), pane, marks);
        }

        private void HideOverlay()
        {
            if (_overlay != null && !_overlay.IsDisposed) _overlay.HideOverlay();
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
            int[] ys = { pane.Top + 8, pane.Top + pane.Height / 2, pane.Bottom - 8 };

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
        /// У метода четыре out-параметра, поэтому вызов идёт через
        /// InvokeMember с ParameterModifier — надёжный способ передать byref
        /// при позднем связывании.
        /// </summary>
        private static bool TryGetScreenRect(dynamic window, dynamic doc,
                                             int start, int end, out Rectangle rect)
        {
            rect = Rectangle.Empty;
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
                    null,
                    win,
                    args,
                    new[] { byRef },
                    null,
                    null);

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
        // Кэш перечня
        // ==================================================================

        private void RefreshKnownIndex(dynamic doc, int contentEnd)
        {
            string key;
            try { key = (string)doc.FullName; } catch { key = null; }

            bool stale = contentEnd != _knownContentEnd || key != _knownDocKey;
            bool stable = contentEnd == _lastTickContentEnd;
            _lastTickContentEnd = contentEnd;

            if (!stale) return;

            // Во время набора длина меняется каждый тик; ждём паузу.
            if (!stable && _knownContentEnd >= 0 && key == _knownDocKey) return;

            var state = _controller.GetState(doc);
            _known = _controller.GetKnownIndex(doc, state);

            _dictRanges = new List<int[]>();
            if (state.DictionaryPages.Count > 0)
            {
                List<int[]> bounds = WordUtil.PageBoundsFor(doc, state.DictionaryPages);
                foreach (int page in state.DictionaryPages)
                {
                    int[] b = WordUtil.BoundsOfPage(bounds, page);
                    if (b != null && b[1] > b[0]) _dictRanges.Add(b);
                }
            }

            _knownContentEnd = contentEnd;
            _knownDocKey = key;
        }

        private bool InDictionary(int position)
        {
            foreach (var r in _dictRanges)
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
