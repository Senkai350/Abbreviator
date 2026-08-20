using System;
using System.Collections.Generic;
using Abbreviator.Interop;

namespace Abbreviator.Core
{
    /// <summary>
    /// Подсветка найденных аббревиатур фоном:
    /// зелёный — сокращение есть в перечне, красный — нет.
    /// </summary>
    public sealed class HighlightService
    {
        private readonly AppSettings _settings;

        public HighlightService(AppSettings settings)
        {
            _settings = settings ?? new AppSettings();
        }

        /// <summary>Сколько вхождений не удалось сопоставить с текстом документа.</summary>
        public int Misaligned { get; private set; }

        /// <summary>
        /// Красит все вхождения из результата разбора.
        /// Возвращает число окрашенных диапазонов.
        /// </summary>
        public int Apply(dynamic app, dynamic doc, ScanResult scan)
        {
            if (doc == null || scan == null) return 0;

            Misaligned = 0;
            int painted = 0;
            int drift = 0;

            // doc приходит как dynamic; ResolveRange принимает ref-параметр,
            // а по dynamic такой вызов не диспетчеризуется. Приводим к object,
            // чтобы вызов связался статически.
            object docRef = doc;

            bool undo = WordUtil.TryStartUndoRecord(app, "Abbreviator: подсветка");
            WordUtil.SetScreenUpdating(app, false);
            try
            {
                foreach (var occ in scan.Occurrences)
                {
                    if (occ.InDictionary && !_settings.HighlightDictionarySection) continue;

                    int color;
                    switch (occ.Status)
                    {
                        case AbbrStatus.Known:
                            color = _settings.KnownColor;
                            break;
                        case AbbrStatus.Unknown:
                            color = _settings.UnknownColor;
                            break;
                        default:
                            if (!_settings.HighlightIgnored) continue;
                            color = Wd.Gray25;
                            break;
                    }

                    dynamic range = ResolveRange(docRef, occ, ref drift);
                    if (range == null) { Misaligned++; continue; }

                    try
                    {
                        range.HighlightColorIndex = color;
                        painted++;
                    }
                    catch { Misaligned++; }
                }
            }
            finally
            {
                WordUtil.SetScreenUpdating(app, true);
                if (undo) WordUtil.EndUndoRecord(app);
            }

            return painted;
        }

        /// <summary>Снимает подсветку только с найденных аббревиатур.</summary>
        public int Clear(dynamic app, dynamic doc, ScanResult scan)
        {
            if (doc == null || scan == null) return 0;

            int cleared = 0;
            int drift = 0;
            object docRef = doc;

            bool undo = WordUtil.TryStartUndoRecord(app, "Abbreviator: снять подсветку");
            WordUtil.SetScreenUpdating(app, false);
            try
            {
                foreach (var occ in scan.Occurrences)
                {
                    dynamic range = ResolveRange(docRef, occ, ref drift);
                    if (range == null) continue;
                    try
                    {
                        range.HighlightColorIndex = Wd.NoHighlight;
                        cleared++;
                    }
                    catch { }
                }
            }
            finally
            {
                WordUtil.SetScreenUpdating(app, true);
                if (undo) WordUtil.EndUndoRecord(app);
            }

            return cleared;
        }

        /// <summary>Снимает подсветку во всём документе, включая чужую.</summary>
        public void ClearAll(dynamic app, dynamic doc)
        {
            if (doc == null) return;
            bool undo = WordUtil.TryStartUndoRecord(app, "Abbreviator: снять всю подсветку");
            try
            {
                doc.Content.HighlightColorIndex = Wd.NoHighlight;
            }
            catch { }
            finally
            {
                if (undo) WordUtil.EndUndoRecord(app);
            }
        }

        /// <summary>Красит одно вхождение (после добавления статьи в перечень).</summary>
        public void PaintOne(dynamic doc, int start, int end, AbbrStatus status)
        {
            try
            {
                int color = status == AbbrStatus.Known ? _settings.KnownColor
                          : status == AbbrStatus.Unknown ? _settings.UnknownColor
                          : Wd.NoHighlight;
                doc.Range(start, end).HighlightColorIndex = color;
            }
            catch { }
        }

        // ------------------------------------------------------------------

        /// <summary>
        /// Возвращает диапазон документа, соответствующий вхождению.
        ///
        /// Индексы берутся из текстовой модели, а она может разъехаться с
        /// координатами Range (поля, внедрённые объекты). Поэтому текст
        /// диапазона проверяется, и при расхождении вхождение ищется рядом;
        /// найденный сдвиг запоминается и применяется к следующим вхождениям.
        /// </summary>
        private dynamic ResolveRange(object document, Occurrence occ, ref int drift)
        {
            dynamic doc = document;
            int len = occ.End - occ.Start;
            if (len <= 0) return null;

            dynamic candidate = TryRange(doc, occ.Start + drift, occ.End + drift, occ.Abbr);
            if (candidate != null) return candidate;

            if (drift != 0)
            {
                candidate = TryRange(doc, occ.Start, occ.End, occ.Abbr);
                if (candidate != null) { drift = 0; return candidate; }
            }

            // Ищем сокращение в окне вокруг ожидаемой позиции.
            const int Window = 200;
            int from = Math.Max(0, occ.Start + drift - Window);
            int to = occ.End + drift + Window;

            string around;
            try { around = doc.Range(from, to).Text as string ?? string.Empty; }
            catch { return null; }

            int best = -1;
            int center = occ.Start + drift - from;
            for (int i = around.IndexOf(occ.Abbr, StringComparison.Ordinal);
                 i >= 0;
                 i = around.IndexOf(occ.Abbr, i + 1, StringComparison.Ordinal))
            {
                if (best < 0 || Math.Abs(i - center) < Math.Abs(best - center)) best = i;
            }

            if (best < 0) return null;

            int newStart = from + best;
            drift = newStart - occ.Start;
            return TryRange(doc, newStart, newStart + len, occ.Abbr);
        }

        private static dynamic TryRange(dynamic doc, int start, int end, string expected)
        {
            if (start < 0 || end <= start) return null;
            try
            {
                dynamic r = doc.Range(start, end);
                string text = r.Text as string;
                if (text != null && string.Equals(text, expected, StringComparison.Ordinal))
                    return r;
            }
            catch { }
            return null;
        }
    }
}
