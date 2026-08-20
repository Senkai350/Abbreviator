using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Abbreviator.Core
{
    /// <summary>
    /// Поиск страниц содержания (оглавления).
    ///
    /// Такие страницы исключаются и из поиска аббревиатур, и из определения
    /// перечня: в содержании есть строка «Перечень принятых сокращений», и без
    /// этой проверки надстройка принимала бы за перечень страницу содержания.
    ///
    /// Признаки, в порядке надёжности:
    ///   1. поле оглавления Word (Document.TablesOfContents) — точный признак;
    ///   2. заголовок «Содержание» / «Оглавление» в начале страницы;
    ///   3. строки вида «Название .......... 12» или «Название → 12» —
    ///      заголовок, отбивка и номер страницы в конце.
    /// </summary>
    public sealed class TocDetector
    {
        /// <summary>Строка содержания: текст, отбивка (точки, подчёркивания, табуляция) и номер.</summary>
        private static readonly Regex TocLineRx = new Regex(
            @"^\s*\S.*?(?:\t+|[\.…_\-\s]{3,})\s*\d{1,4}\s*$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>Заголовки самой страницы содержания.</summary>
        private static readonly string[] TocHeadings =
        {
            "содержание", "оглавление", "ведомость документа", "состав документа"
        };

        /// <summary>Сколько строк-ссылок подряд считаем содержанием.</summary>
        private const int MinTocLines = 3;

        private readonly HashSet<int> _pages = new HashSet<int>();
        private readonly List<int[]> _ranges = new List<int[]>();

        public HashSet<int> Pages { get { return _pages; } }
        public List<int[]> Ranges { get { return _ranges; } }

        public bool IsEmpty
        {
            get { return _pages.Count == 0 && _ranges.Count == 0; }
        }

        public bool IsTocPage(int page)
        {
            return _pages.Contains(page);
        }

        public bool IsInToc(int position)
        {
            foreach (var r in _ranges)
                if (position >= r[0] && position < r[1]) return true;
            return false;
        }

        /// <summary>Абзац похож на строку содержания.</summary>
        public static bool LooksLikeTocLine(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            if (text.Length > 200) return false;
            return TocLineRx.IsMatch(text);
        }

        public static bool LooksLikeTocHeading(string text)
        {
            string norm = WordUtil.NormalizeForCompare(text);
            if (norm.Length == 0 || norm.Length > 40) return false;

            foreach (var heading in TocHeadings)
                if (norm == heading || norm.StartsWith(heading + " ", StringComparison.Ordinal))
                    return true;
            return false;
        }

        // ==================================================================

        public static TocDetector Build(dynamic doc, TextModel model, List<int[]> pageBounds)
        {
            var toc = new TocDetector();
            if (model == null) return toc;

            toc.CollectFields(doc);
            toc.CollectByText(model, pageBounds);
            toc.MarkPagesOfRanges(pageBounds);
            toc.AddPagesAsRanges(pageBounds);
            return toc;
        }

        /// <summary>Поля оглавления и списков иллюстраций — самый надёжный признак.</summary>
        private void CollectFields(dynamic doc)
        {
            AddRangesOf(doc, "TablesOfContents");
            AddRangesOf(doc, "TablesOfFigures");
        }

        private void AddRangesOf(dynamic doc, string collectionName)
        {
            try
            {
                dynamic collection = doc.GetType().InvokeMember(
                    collectionName,
                    System.Reflection.BindingFlags.GetProperty,
                    null, doc, null);

                int count = (int)collection.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic range = collection[i].Range;
                    int start = (int)range.Start;
                    int end = (int)range.End;
                    if (end > start) _ranges.Add(new[] { start, end });
                }
            }
            catch
            {
                // Коллекции может не быть — не страшно, останутся текстовые признаки.
            }
        }

        /// <summary>Содержание, набранное вручную: заголовок и строки с номерами.</summary>
        private void CollectByText(TextModel model, List<int[]> pageBounds)
        {
            var linesPerPage = new Dictionary<int, int>();
            var headingPages = new HashSet<int>();

            foreach (var p in model.Paragraphs)
            {
                if (p.IsEmpty) continue;

                int page = WordUtil.PageByPosition(pageBounds, p.Start);
                if (page <= 0) continue;

                if (LooksLikeTocHeading(p.Text))
                {
                    headingPages.Add(page);
                    continue;
                }

                if (LooksLikeTocLine(p.Text))
                {
                    int n;
                    linesPerPage.TryGetValue(page, out n);
                    linesPerPage[page] = n + 1;
                }
            }

            foreach (var pair in linesPerPage)
            {
                // Заголовок «Содержание» снижает порог: страница с ним и парой
                // ссылок — уже содержание.
                int threshold = headingPages.Contains(pair.Key) ? 2 : MinTocLines;
                if (pair.Value >= threshold) _pages.Add(pair.Key);
            }

            // Заголовок без строк-ссылок (содержание из одного поля) тоже считаем.
            foreach (int page in headingPages)
            {
                int n;
                if (linesPerPage.TryGetValue(page, out n) && n > 0) _pages.Add(page);
            }
        }

        /// <summary>
        /// Границы страниц содержания добавляем в список диапазонов: тогда
        /// проверка «позиция в содержании» — обычное сравнение чисел, без
        /// обращения к Word на каждое найденное слово.
        /// </summary>
        private void AddPagesAsRanges(List<int[]> pageBounds)
        {
            foreach (int page in _pages)
            {
                int[] b = WordUtil.BoundsOfPage(pageBounds, page);
                if (b != null && b[1] > b[0]) _ranges.Add(new[] { b[0], b[1] });
            }
        }

        private void MarkPagesOfRanges(List<int[]> pageBounds)
        {
            if (pageBounds == null) return;

            foreach (var r in _ranges)
            {
                int first = WordUtil.PageByPosition(pageBounds, r[0]);
                int last = WordUtil.PageByPosition(pageBounds, Math.Max(r[0], r[1] - 1));
                if (first <= 0 || last <= 0) continue;
                for (int p = first; p <= last; p++) _pages.Add(p);
            }
        }
    }
}
