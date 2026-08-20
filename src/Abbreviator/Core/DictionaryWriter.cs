using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Abbreviator.Interop;

namespace Abbreviator.Core
{
    /// <summary>Результат добавления статьи в перечень.</summary>
    public sealed class WriteResult
    {
        public bool Success;
        public string Error;

        /// <summary>Диапазон новой статьи в документе.</summary>
        public int Start = -1;
        public int End = -1;

        /// <summary>Страница, на которую попала статья.</summary>
        public int Page;
    }

    /// <summary>
    /// Дописывает статью в «Перечень принятых сокращений».
    /// Формат строки: «СОКРАЩЕНИЕ - расшифровка».
    /// </summary>
    public sealed class DictionaryWriter
    {
        private readonly AppSettings _settings;

        public DictionaryWriter(AppSettings settings)
        {
            _settings = settings ?? new AppSettings();
        }

        public WriteResult AddEntry(dynamic app, dynamic doc, DocumentState state, ScanResult scan,
                                    string abbr, string definition)
        {
            var result = new WriteResult();

            if (doc == null)
            {
                result.Error = "Нет активного документа.";
                return result;
            }
            if (string.IsNullOrWhiteSpace(abbr))
            {
                result.Error = "Не задано сокращение.";
                return result;
            }
            if (state.DictionaryPages.Count == 0)
            {
                result.Error = "Не задан ни один лист перечня принятых сокращений. " +
                               "Откройте «Листы перечня» и добавьте страницы вручную " +
                               "или запустите автоопределение.";
                return result;
            }

            abbr = abbr.Trim();
            definition = Flatten(definition);

            var anchor = ChooseAnchor(scan, state, abbr);
            bool insertBefore = anchor != null && _settings.KeepAlphabeticalOrder &&
                                Compare(anchor.Terms.FirstOrDefault() ?? anchor.Term, abbr) > 0;

            bool undo = WordUtil.TryStartUndoRecord(app, "Abbreviator: добавить сокращение");
            try
            {
                if (anchor == null)
                    WriteAfterHeader(doc, scan, state, abbr, definition, result);
                else if (anchor.IsTableRow)
                    WriteTableRow(doc, anchor, abbr, definition, insertBefore, result);
                else
                    WriteParagraph(doc, anchor, abbr, definition, insertBefore, result);
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = "Не удалось вставить статью: " + ex.Message;
            }
            finally
            {
                if (undo) WordUtil.EndUndoRecord(app);
            }

            if (result.Success)
            {
                WordUtil.Repaginate(doc);
                result.Page = WordUtil.PageOf(doc, result.Start);
                if (result.Page > 0) state.AddPage(result.Page);
                state.Save(doc);
            }

            return result;
        }

        // ------------------------------------------------------------------
        // Выбор точки вставки
        // ------------------------------------------------------------------

        /// <summary>
        /// Статья, относительно которой вставляем новую.
        /// В алфавитном режиме — первая статья, которая должна идти после новой;
        /// иначе — последняя статья на последнем листе перечня.
        /// </summary>
        private DictEntry ChooseAnchor(ScanResult scan, DocumentState state, string abbr)
        {
            if (scan == null || scan.Entries.Count == 0) return null;

            var pages = new HashSet<int>(state.DictionaryPages);
            var entries = scan.Entries.Where(e => pages.Contains(e.Page))
                                      .OrderBy(e => e.Start)
                                      .ToList();
            if (entries.Count == 0) entries = scan.Entries.OrderBy(e => e.Start).ToList();
            if (entries.Count == 0) return null;

            if (_settings.KeepAlphabeticalOrder)
            {
                foreach (var e in entries)
                {
                    string term = e.Terms.FirstOrDefault() ?? e.Term;
                    if (Compare(term, abbr) > 0) return e;
                }
            }

            int lastPage = state.LastDictionaryPage;
            var onLastPage = entries.Where(e => e.Page == lastPage).ToList();
            return onLastPage.Count > 0 ? onLastPage[onLastPage.Count - 1] : entries[entries.Count - 1];
        }

        private static int Compare(string a, string b)
        {
            return string.Compare(a ?? string.Empty, b ?? string.Empty,
                                  CultureInfo.GetCultureInfo("ru-RU"), CompareOptions.IgnoreCase);
        }

        // ------------------------------------------------------------------
        // Вставка
        // ------------------------------------------------------------------

        private void WriteParagraph(dynamic doc, DictEntry anchor, string abbr, string definition,
                                    bool insertBefore, WriteResult result)
        {
            string line = BuildLine(abbr, definition);
            dynamic source = doc.Range(anchor.Start, anchor.End);

            int contentEnd = (int)doc.Content.End;
            dynamic target;

            if (insertBefore)
            {
                target = doc.Range(anchor.Start, anchor.Start);
                target.InsertAfter(line + "\r");
            }
            else if (anchor.End < contentEnd)
            {
                target = doc.Range(anchor.End, anchor.End);
                target.InsertAfter(line + "\r");
            }
            else
            {
                // Статья — последний абзац документа.
                target = doc.Range(contentEnd - 1, contentEnd - 1);
                target.InsertAfter("\r" + line);
            }

            CopyFormatting(source, target);

            result.Success = true;
            result.Start = (int)target.Start;
            result.End = (int)target.End;
        }

        private void WriteTableRow(dynamic doc, DictEntry anchor, string abbr, string definition,
                                   bool insertBefore, WriteResult result)
        {
            dynamic anchorRange = doc.Range(anchor.Start, anchor.End);
            dynamic table = anchorRange.Tables[1];
            int rowIndex = (int)anchorRange.Cells[1].RowIndex;
            int rowCount = (int)table.Rows.Count;

            dynamic newRow;
            if (insertBefore)
            {
                newRow = table.Rows.Add(table.Rows[rowIndex]);
            }
            else if (rowIndex >= rowCount)
            {
                newRow = table.Rows.Add(Type.Missing);
            }
            else
            {
                newRow = table.Rows.Add(table.Rows[rowIndex + 1]);
            }

            int cells = (int)newRow.Cells.Count;
            newRow.Cells[1].Range.Text = abbr;
            if (cells >= 2) newRow.Cells[2].Range.Text = definition;
            else if (definition.Length > 0) newRow.Cells[1].Range.Text = BuildLine(abbr, definition);

            result.Success = true;
            result.Start = (int)newRow.Range.Start;
            result.End = (int)newRow.Range.End;
        }

        /// <summary>Перечень найден, но статей в нём ещё нет — пишем сразу после заголовка.</summary>
        private void WriteAfterHeader(dynamic doc, ScanResult scan, DocumentState state,
                                      string abbr, string definition, WriteResult result)
        {
            int anchorPos = scan != null && scan.HeaderStart >= 0 ? scan.HeaderStart : -1;

            if (anchorPos < 0)
            {
                // Заголовок не найден — становимся в конец последнего листа перечня.
                List<int[]> bounds = WordUtil.PageBounds(doc);
                int[] b = WordUtil.BoundsOfPage(bounds, state.LastDictionaryPage);
                if (b == null)
                {
                    result.Error = "Не удалось определить место вставки: перечень не найден.";
                    return;
                }
                anchorPos = Math.Max(b[0], b[1] - 1);
            }
            else
            {
                dynamic headerPara = doc.Range(anchorPos, anchorPos).Paragraphs[1];
                anchorPos = (int)headerPara.Range.End;
            }

            int contentEnd = (int)doc.Content.End;
            string line = BuildLine(abbr, definition);
            dynamic target;

            if (anchorPos < contentEnd)
            {
                target = doc.Range(anchorPos, anchorPos);
                target.InsertAfter(line + "\r");
            }
            else
            {
                target = doc.Range(contentEnd - 1, contentEnd - 1);
                target.InsertAfter("\r" + line);
            }

            result.Success = true;
            result.Start = (int)target.Start;
            result.End = (int)target.End;
        }

        private static void CopyFormatting(dynamic source, dynamic target)
        {
            try { target.ParagraphFormat = source.ParagraphFormat; } catch { }
            try { target.Style = source.Style; } catch { }
            try { target.Font = source.Font; } catch { }
            try { target.HighlightColorIndex = Wd.NoHighlight; } catch { }
        }

        private string BuildLine(string abbr, string definition)
        {
            string sep = string.IsNullOrEmpty(_settings.EntrySeparator) ? " - " : _settings.EntrySeparator;
            return string.IsNullOrEmpty(definition) ? abbr : abbr + sep + definition;
        }

        private static string Flatten(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var chars = text.Select(c => WordUtil.IsServiceChar(c) ? ' ' : c).ToArray();
            return new string(chars).Trim();
        }
    }
}
