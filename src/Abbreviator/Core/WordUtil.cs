using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Abbreviator.Interop;

namespace Abbreviator.Core
{
    /// <summary>
    /// Тонкая обёртка над объектной моделью Word при позднем связывании.
    /// Все методы принимают dynamic и стараются не бросать исключений наружу.
    /// </summary>
    public static class WordUtil
    {
        /// <summary>Неразрывный пробел.</summary>
        public const char Nbsp = (char)0x00A0;

        /// <summary>Текст всего основного потока документа одним вызовом COM.</summary>
        public static string GetContentText(dynamic doc)
        {
            try
            {
                string t = doc.Content.Text as string;
                return t ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public static int ContentStart(dynamic doc)
        {
            try { return (int)doc.Content.Start; } catch { return 0; }
        }

        public static int ContentEnd(dynamic doc)
        {
            try { return (int)doc.Content.End; } catch { return 0; }
        }

        public static dynamic Range(dynamic doc, int start, int end)
        {
            return doc.Range(start, end);
        }

        public static string RangeText(dynamic doc, int start, int end)
        {
            try
            {
                string t = doc.Range(start, end).Text as string;
                return t ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public static int PageCount(dynamic doc)
        {
            try { return (int)doc.ComputeStatistics(Wd.StatisticPages); }
            catch { return 0; }
        }

        public static int PageOf(dynamic doc, int position)
        {
            try
            {
                dynamic r = doc.Range(position, position);
                return (int)r.Information[Wd.ActiveEndPageNumber];
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>Позиция начала страницы pageNumber (1-based) или -1.</summary>
        public static int PageStart(dynamic doc, int pageNumber)
        {
            try
            {
                dynamic r = doc.GoTo(Wd.GoToPage, Wd.GoToAbsolute, pageNumber, Type.Missing);
                return (int)r.Start;
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>Границы всех страниц документа: [start, end) в координатах Range.</summary>
        public static List<int[]> PageBounds(dynamic doc)
        {
            var list = new List<int[]>();
            int pages = PageCount(doc);
            if (pages <= 0) return list;

            int contentEnd = ContentEnd(doc);
            var starts = new int[pages + 2];
            for (int p = 1; p <= pages; p++)
            {
                int s = PageStart(doc, p);
                starts[p] = s < 0 ? (p == 1 ? 0 : starts[p - 1]) : s;
            }

            for (int p = 1; p <= pages; p++)
            {
                int s = starts[p];
                int e = (p < pages) ? starts[p + 1] : contentEnd;
                if (e < s) e = s;
                list.Add(new[] { s, e });
            }
            return list;
        }

        /// <summary>
        /// Границы только нужных страниц. Для контекстного меню это заметно
        /// дешевле PageBounds: перечень занимает одну-две страницы, а не весь
        /// документ. Незапрошенные страницы остаются нулевыми.
        /// </summary>
        public static List<int[]> PageBoundsFor(dynamic doc, IEnumerable<int> pages)
        {
            var wanted = (pages ?? Enumerable.Empty<int>()).Where(p => p > 0)
                                                           .Distinct()
                                                           .OrderBy(p => p)
                                                           .ToList();
            var list = new List<int[]>();
            if (wanted.Count == 0) return list;

            int total = PageCount(doc);
            int contentEnd = ContentEnd(doc);

            for (int i = 0; i < wanted[wanted.Count - 1]; i++) list.Add(new[] { 0, 0 });

            foreach (int p in wanted)
            {
                if (p > total) continue;

                int s = PageStart(doc, p);
                if (s < 0) continue;

                int e = p < total ? PageStart(doc, p + 1) : contentEnd;
                if (e <= s) e = contentEnd;

                list[p - 1] = new[] { s, e };
            }

            return list;
        }

        /// <summary>Границы конкретной страницы или null.</summary>
        public static int[] BoundsOfPage(List<int[]> bounds, int pageNumber)
        {
            if (bounds == null) return null;
            if (pageNumber < 1 || pageNumber > bounds.Count) return null;
            return bounds[pageNumber - 1];
        }

        /// <summary>Номер страницы по позиции — двоичным поиском по границам.</summary>
        public static int PageByPosition(List<int[]> bounds, int position)
        {
            if (bounds == null || bounds.Count == 0) return 0;
            int lo = 0, hi = bounds.Count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                if (position < bounds[mid][0]) hi = mid - 1;
                else if (position >= bounds[mid][1]) lo = mid + 1;
                else return mid + 1;
            }
            return Math.Min(Math.Max(lo, 1), bounds.Count);
        }

        public static void Repaginate(dynamic doc)
        {
            try { doc.Repaginate(); } catch { }
        }

        public static void ScrollTo(dynamic doc, int start, int end)
        {
            try
            {
                dynamic r = doc.Range(start, end);
                r.Select();
                doc.ActiveWindow.ScrollIntoView(r, true);
            }
            catch { }
        }

        public static void GoToPageInView(dynamic doc, int pageNumber)
        {
            try
            {
                dynamic r = doc.GoTo(Wd.GoToPage, Wd.GoToAbsolute, pageNumber, Type.Missing);
                r.Select();
                doc.ActiveWindow.ScrollIntoView(r, true);
            }
            catch { }
        }

        public static bool TryStartUndoRecord(dynamic app, string name)
        {
            try
            {
                app.UndoRecord.StartCustomRecord(name);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void EndUndoRecord(dynamic app)
        {
            try { app.UndoRecord.EndCustomRecord(); } catch { }
        }

        public static void SetScreenUpdating(dynamic app, bool value)
        {
            try { app.ScreenUpdating = value; } catch { }
        }

        public static dynamic ActiveDocument(dynamic app)
        {
            try
            {
                if (app == null) return null;
                if ((int)app.Documents.Count == 0) return null;
                return app.ActiveDocument;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Служебный символ Word, не несущий текста: маркеры абзаца и ячейки,
        /// разрывы строки и страницы, коды полей, ссылки на сноски и примечания.
        /// Все они лежат в диапазоне управляющих символов.
        /// </summary>
        public static bool IsServiceChar(char c)
        {
            return (c < ' ' && c != '\t') || c == (char)0x7F;
        }

        /// <summary>Текст абзаца без служебных символов Word.</summary>
        public static string CleanParagraph(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var sb = new StringBuilder(text.Length);
            foreach (char c in text)
            {
                if (IsServiceChar(c)) continue;
                sb.Append(c == Nbsp ? ' ' : c);
            }
            return sb.ToString().Trim();
        }

        /// <summary>Свернуть пробелы, убрать пунктуацию, привести к нижнему регистру.</summary>
        public static string NormalizeForCompare(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var sb = new StringBuilder(text.Length);
            bool space = true;
            foreach (char c in text)
            {
                char ch = c == Nbsp ? ' ' : c;
                if (IsServiceChar(ch) || char.IsWhiteSpace(ch))
                {
                    if (!space) { sb.Append(' '); space = true; }
                    continue;
                }
                if (char.IsPunctuation(ch) && ch != '-') continue;
                sb.Append(char.ToLowerInvariant(ch));
                space = false;
            }
            return sb.ToString().Trim();
        }
    }
}
