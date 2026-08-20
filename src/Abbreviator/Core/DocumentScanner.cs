using System;
using System.Collections.Generic;
using System.Linq;

namespace Abbreviator.Core
{
    /// <summary>
    /// Полный разбор документа: поиск аббревиатур, разбор перечня принятых
    /// сокращений и сопоставление одного с другим.
    /// </summary>
    public sealed class DocumentScanner
    {
        private readonly AppSettings _settings;

        public DocumentScanner(AppSettings settings)
        {
            _settings = settings ?? new AppSettings();
        }

        /// <summary>
        /// Разбирает документ. Список страниц перечня берётся из состояния
        /// документа; если он пуст, страницы определяются автоматически и
        /// записываются в состояние.
        /// </summary>
        public ScanResult Scan(dynamic app, dynamic doc, DocumentState state)
        {
            var result = new ScanResult();
            if (doc == null) return result;

            // Типы указаны явно: аргумент doc/app имеет тип dynamic, поэтому
            // при var результат тоже стал бы dynamic, а по dynamic не работают
            // ни LINQ, ни out-параметры.
            AbbreviationRecognizer recognizer = CreateRecognizer(app);
            var parser = new DictionaryParser(_settings, recognizer);

            WordUtil.Repaginate(doc);
            List<int[]> pageBounds = WordUtil.PageBounds(doc);
            result.TotalPages = pageBounds.Count;

            TextModel model = TextModel.Build(doc);

            // --- 0. Страницы содержания ------------------------------------
            // В содержании встречается строка «Перечень принятых сокращений»,
            // поэтому такие страницы исключаются и из поиска заголовка, и из
            // поиска аббревиатур.
            TocDetector toc = _settings.SkipTableOfContents
                ? TocDetector.Build(doc, model, pageBounds)
                : new TocDetector();
            parser.Toc = toc;
            result.TocPages = new List<int>(toc.Pages);

            // --- 1. Страницы перечня ---------------------------------------
            int headerStart, headerPage;
            var detected = parser.DetectPages(model, pageBounds, out headerStart, out headerPage);
            result.HeaderStart = headerStart;
            result.HeaderPage = headerPage;

            // Список страниц ведёт пользователь; автоопределение срабатывает,
            // только пока список пуст.
            if (state.DictionaryPages.Count == 0)
                state.SetPages(detected);

            result.DictionaryPages = state.DictionaryPages.ToList();

            // --- 2. Статьи перечня -----------------------------------------
            result.Entries = parser.ParseEntries(model, pageBounds, state.DictionaryPages);

            var known = new Dictionary<string, DictEntry>(StringComparer.Ordinal);
            foreach (var e in result.Entries)
                foreach (var term in e.Terms)
                    if (!known.ContainsKey(term)) known[term] = e;

            // --- 3. Диапазоны страниц перечня ------------------------------
            var dictRanges = new List<int[]>();
            foreach (int page in state.DictionaryPages)
            {
                int[] b = WordUtil.BoundsOfPage(pageBounds, page);
                if (b != null) dictRanges.Add(b);
            }

            // --- 4. Вхождения в тексте -------------------------------------
            foreach (var p in model.Paragraphs)
            {
                if (p.IsEmpty) continue;
                if (recognizer.IsUpperCaseHeading(p.Text)) continue;

                if (_settings.SkipTableOfContents)
                {
                    int paraPage = WordUtil.PageByPosition(pageBounds, p.Start);
                    if (toc.IsTocPage(paraPage) || toc.IsInToc(p.Start)) continue;
                    if (TocDetector.LooksLikeTocLine(p.Text)) continue;
                }

                foreach (var token in recognizer.Find(p.Raw))
                {
                    int start = p.Start + token.Index;
                    int end = start + token.Length;

                    var occ = new Occurrence
                    {
                        Abbr = token.Abbr,
                        Raw = token.Raw,
                        Start = start,
                        End = end,
                        InDictionary = InAnyRange(dictRanges, start),
                        Page = WordUtil.PageByPosition(pageBounds, start)
                    };

                    if (state.IsIgnored(token.Abbr)) occ.Status = AbbrStatus.Ignored;
                    else if (known.ContainsKey(token.Abbr)) occ.Status = AbbrStatus.Known;
                    else occ.Status = AbbrStatus.Unknown;

                    result.Occurrences.Add(occ);
                }
            }

            // --- 5. Сводка --------------------------------------------------
            foreach (var occ in result.Occurrences)
            {
                AbbrInfo info;
                if (!result.ByAbbr.TryGetValue(occ.Abbr, out info))
                {
                    info = new AbbrInfo { Abbr = occ.Abbr, Status = occ.Status };
                    result.ByAbbr[occ.Abbr] = info;
                }

                if (occ.InDictionary)
                {
                    info.DictionaryCount++;
                }
                else
                {
                    info.UsageCount++;
                    if (info.FirstUse < 0) info.FirstUse = occ.Start;
                }
            }

            // Сокращения, объявленные в перечне, но ни разу не использованные,
            // тоже должны попасть в сводку.
            foreach (var pair in known)
            {
                AbbrInfo info;
                if (!result.ByAbbr.TryGetValue(pair.Key, out info))
                {
                    info = new AbbrInfo { Abbr = pair.Key, Status = AbbrStatus.Known };
                    result.ByAbbr[pair.Key] = info;
                }
                info.Status = state.IsIgnored(pair.Key) ? AbbrStatus.Ignored : AbbrStatus.Known;
                info.Definition = pair.Value.Definition;
                info.EntryStart = pair.Value.Start;
            }

            foreach (var info in result.ByAbbr.Values)
            {
                switch (info.Status)
                {
                    case AbbrStatus.Known: result.KnownCount++; break;
                    case AbbrStatus.Unknown: result.UnknownCount++; break;
                    case AbbrStatus.Ignored: result.IgnoredCount++; break;
                }
            }

            return result;
        }

        /// <summary>
        /// Распознаватель с подключённой проверкой орфографии Word: она отсеивает
        /// обычные слова, набранные капслоком («ВВЕДЕНИЕ», «ЗАКЛЮЧЕНИЕ»).
        /// </summary>
        public AbbreviationRecognizer CreateRecognizer(dynamic app)
        {
            var recognizer = new AbbreviationRecognizer(_settings);

            if (_settings.UseSpellCheckFilter && app != null)
            {
                recognizer.IsCommonWord = word =>
                {
                    try { return (bool)app.CheckSpelling(word); }
                    catch { return false; }
                };
            }

            return recognizer;
        }

        private static bool InAnyRange(List<int[]> ranges, int position)
        {
            foreach (var r in ranges)
                if (position >= r[0] && position < r[1]) return true;
            return false;
        }
    }
}
