using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Abbreviator.Core
{
    /// <summary>
    /// Разбор «Перечня принятых сокращений»: поиск заголовка, определение
    /// страниц перечня и извлечение статей вида «АСУ - расшифровка».
    /// </summary>
    public sealed class DictionaryParser
    {
        /// <summary>
        /// Статья перечня. Левая часть — до 60 символов, дальше разделитель
        /// (дефис, тире, двоеточие или табуляция) и расшифровка.
        /// </summary>
        private static readonly Regex EntryRx = new Regex(
            @"^\s*(?<term>[^\t]{1,60}?)\s*(?:\t+|\s[-–—]\s|[-–—]\s|\s[-–—]|:\s)\s*(?<def>\S.*)$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>Токен, который выглядит как сокращение, даже если распознаватель его отфильтровал.</summary>
        private static readonly Regex LooseTermRx = new Regex(
            @"^[A-ZА-ЯЁ0-9][A-ZА-ЯЁ0-9\-\./ ]{0,30}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private readonly AppSettings _settings;
        private readonly AbbreviationRecognizer _recognizer;

        /// <summary>
        /// Страницы содержания. В содержании есть строка «Перечень принятых
        /// сокращений», и без этой проверки она принималась бы за заголовок
        /// самого перечня.
        /// </summary>
        public TocDetector Toc;

        public DictionaryParser(AppSettings settings, AbbreviationRecognizer recognizer)
        {
            _settings = settings ?? new AppSettings();
            _recognizer = recognizer;
        }

        /// <summary>Абзац лежит в содержании и должен пропускаться.</summary>
        private bool InToc(ParaSpan p, List<int[]> pageBounds)
        {
            if (!_settings.SkipTableOfContents || Toc == null) return false;
            if (Toc.IsInToc(p.Start)) return true;
            if (TocDetector.LooksLikeTocLine(p.Text)) return true;

            int page = WordUtil.PageByPosition(pageBounds, p.Start);
            return page > 0 && Toc.IsTocPage(page);
        }

        // ------------------------------------------------------------------
        // Поиск заголовка
        // ------------------------------------------------------------------

        /// <summary>
        /// Индекс абзаца с заголовком перечня или -1.
        /// Заголовком считается короткий абзац, целиком состоящий из одного
        /// из вариантов названия (номер раздела и знаки препинания игнорируются).
        /// </summary>
        public int FindHeaderParagraph(TextModel model)
        {
            return FindHeaderParagraph(model, null);
        }

        /// <summary>
        /// Индекс абзаца с заголовком перечня или -1.
        /// Абзацы на страницах содержания пропускаются.
        /// </summary>
        public int FindHeaderParagraph(TextModel model, List<int[]> pageBounds)
        {
            var variants = _settings.HeaderVariants
                .Select(WordUtil.NormalizeForCompare)
                .Where(v => v.Length > 0)
                .OrderByDescending(v => v.Length)
                .ToList();

            if (variants.Count == 0) return -1;

            for (int i = 0; i < model.Paragraphs.Count; i++)
            {
                var para = model.Paragraphs[i];
                if (InToc(para, pageBounds)) continue;

                var norm = WordUtil.NormalizeForCompare(para.Text);
                if (norm.Length == 0) continue;

                foreach (var v in variants)
                {
                    int pos = norm.IndexOf(v, StringComparison.Ordinal);
                    if (pos < 0) continue;

                    // Заголовок, а не упоминание в тексте: строка почти целиком
                    // состоит из названия (допускаем нумерацию вида "3.1").
                    if (pos <= 12 && norm.Length <= v.Length + 12) return i;
                }
            }
            return -1;
        }

        // ------------------------------------------------------------------
        // Автоопределение страниц перечня
        // ------------------------------------------------------------------

        /// <summary>
        /// Страницы, на которых расположен перечень. Раздел прослеживается от
        /// заголовка вперёд, пока встречаются статьи; благодаря этому в список
        /// попадают и страницы продолжения.
        /// </summary>
        public List<int> DetectPages(TextModel model, List<int[]> pageBounds, out int headerStart, out int headerPage)
        {
            headerStart = -1;
            headerPage = 0;

            var pages = new List<int>();
            int headerIndex = FindHeaderParagraph(model, pageBounds);
            if (headerIndex < 0) return pages;

            var header = model.Paragraphs[headerIndex];
            headerStart = header.Start;
            headerPage = WordUtil.PageByPosition(pageBounds, header.Start);

            var otherHeaders = _settings.HeaderVariants
                .Select(WordUtil.NormalizeForCompare)
                .ToList();

            int lastEntryStart = header.Start;
            int misses = 0;

            for (int i = headerIndex + 1; i < model.Paragraphs.Count; i++)
            {
                var p = model.Paragraphs[i];

                if (p.IsEmpty) continue;

                if (LooksLikeEntry(p))
                {
                    lastEntryStart = p.Start;
                    misses = 0;
                    continue;
                }

                // Новый заголовок такого же уровня — перечень закончился.
                var norm = WordUtil.NormalizeForCompare(p.Text);
                if (otherHeaders.Any(v => v.Length > 0 && norm.StartsWith(v, StringComparison.Ordinal)))
                    break;

                // Заголовок, набранный капслоком, — тоже граница раздела.
                if (_recognizer != null && _recognizer.IsUpperCaseHeading(p.Text) && norm.Length < 80)
                    break;

                if (++misses >= 3) break;
            }

            int lastPage = WordUtil.PageByPosition(pageBounds, lastEntryStart);
            if (headerPage <= 0) headerPage = 1;
            if (lastPage < headerPage) lastPage = headerPage;

            for (int p = headerPage; p <= lastPage; p++)
            {
                if (_settings.SkipTableOfContents && Toc != null && Toc.IsTocPage(p)) continue;
                pages.Add(p);
            }
            return pages;
        }

        private bool LooksLikeEntry(ParaSpan p)
        {
            if (p.IsEmpty) return false;

            if (p.InTable)
                return ExtractTerms(p.Text).Count > 0 || p.Text.Length < 200;

            var m = EntryRx.Match(p.Text);
            if (!m.Success) return ExtractTerms(p.Text).Count > 0 && p.Text.Length <= 40;

            return ExtractTerms(m.Groups["term"].Value).Count > 0;
        }

        // ------------------------------------------------------------------
        // Разбор статей
        // ------------------------------------------------------------------

        /// <summary>Статьи перечня на указанных страницах.</summary>
        public List<DictEntry> ParseEntries(TextModel model, List<int[]> pageBounds, IEnumerable<int> pages)
        {
            var result = new List<DictEntry>();
            if (pages == null) return result;

            foreach (int page in pages.Distinct().OrderBy(x => x))
            {
                var bounds = WordUtil.BoundsOfPage(pageBounds, page);
                if (bounds == null) continue;
                ParseRange(model, page, bounds[0], bounds[1], result);
            }

            return result;
        }

        private void ParseRange(TextModel model, int page, int start, int end, List<DictEntry> sink)
        {
            var cells = new List<ParaSpan>();

            foreach (var p in model.ParagraphsIn(start, end))
            {
                if (p.InTable)
                {
                    if (p.IsEmpty)
                    {
                        FlushRow(cells, page, sink);   // маркер конца строки таблицы
                        cells.Clear();
                    }
                    else
                    {
                        cells.Add(p);
                    }
                    continue;
                }

                if (cells.Count > 0)
                {
                    FlushRow(cells, page, sink);
                    cells.Clear();
                }

                if (p.IsEmpty) continue;
                if (_settings.SkipTableOfContents && TocDetector.LooksLikeTocLine(p.Text)) continue;

                var entry = ParseParagraph(p, page);
                if (entry != null) sink.Add(entry);
            }

            if (cells.Count > 0) FlushRow(cells, page, sink);
        }

        private void FlushRow(List<ParaSpan> cells, int page, List<DictEntry> sink)
        {
            if (cells.Count == 0) return;

            var first = cells[0];
            var terms = ExtractTerms(first.Text);
            if (terms.Count == 0) return;

            string definition = cells.Count > 1
                ? string.Join(" ", cells.Skip(1).Select(c => c.Text)).Trim()
                : string.Empty;

            sink.Add(new DictEntry
            {
                Term = first.Text.Trim(),
                Definition = definition,
                Page = page,
                Start = first.Start,
                End = cells[cells.Count - 1].End,
                IsTableRow = true,
                Terms = terms
            });
        }

        private DictEntry ParseParagraph(ParaSpan p, int page)
        {
            string term, definition;

            var m = EntryRx.Match(p.Text);
            if (m.Success)
            {
                term = m.Groups["term"].Value.Trim();
                definition = m.Groups["def"].Value.Trim();
            }
            else
            {
                // Строка без разделителя: считаем статьёй, только если она короткая
                // и целиком состоит из сокращения.
                term = p.Text.Trim();
                definition = string.Empty;
                if (term.Length > 40) return null;
            }

            var terms = ExtractTerms(term);
            if (terms.Count == 0) return null;

            return new DictEntry
            {
                Term = term,
                Definition = definition,
                Page = page,
                Start = p.Start,
                End = p.End,
                IsTableRow = false,
                Terms = terms
            };
        }

        /// <summary>Сокращения в левой части статьи.</summary>
        public List<string> ExtractTerms(string termText)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(termText)) return list;

            if (_recognizer != null)
            {
                foreach (var t in _recognizer.Find(termText))
                    if (!list.Contains(t.Abbr)) list.Add(t.Abbr);
            }

            if (list.Count > 0) return list;

            // Распознаватель мог отсеять токен (стоп-слово, орфография и т.п.),
            // но в перечне он объявлен явно — принимаем по мягкому правилу.
            string trimmed = termText.Trim();
            if (LooseTermRx.IsMatch(trimmed))
            {
                foreach (var part in trimmed.Split(',', '/', ';'))
                {
                    var t = part.Trim();
                    if (t.Length >= 2 && t.Any(char.IsLetter) && !list.Contains(t))
                        list.Add(t);
                }
            }

            return list;
        }
    }
}
