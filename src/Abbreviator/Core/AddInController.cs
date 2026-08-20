using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Abbreviator.Interop;
using Abbreviator.UI;

namespace Abbreviator.Core
{
    /// <summary>
    /// Точка сборки всей логики надстройки: разбор документа, подсветка,
    /// работа со списком листов перечня и с игнорируемыми сокращениями.
    /// </summary>
    public sealed class AddInController
    {
        private readonly Dictionary<string, DocumentState> _states =
            new Dictionary<string, DocumentState>(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<string, KnownIndex> _indexes =
            new Dictionary<string, KnownIndex>(StringComparer.OrdinalIgnoreCase);

        private Target _cachedTarget;
        private DateTime _cachedTargetAt = DateTime.MinValue;
        private int _cachedTargetPos = -1;

        private MainForm _mainForm;

        public dynamic App { get; private set; }

        public AppSettings Settings
        {
            get { return AppSettings.Current; }
        }

        public AddInController(dynamic app)
        {
            App = app;
        }

        public void Shutdown()
        {
            try
            {
                if (_mainForm != null && !_mainForm.IsDisposed) _mainForm.Close();
            }
            catch { }
            _mainForm = null;
            App = null;
        }

        // ==================================================================
        // Документ и его состояние
        // ==================================================================

        public dynamic ActiveDocument
        {
            get { return WordUtil.ActiveDocument(App); }
        }

        private static string KeyOf(dynamic doc)
        {
            try
            {
                string full = doc.FullName as string;
                if (!string.IsNullOrEmpty(full)) return full;
            }
            catch { }
            try { return (doc.Name as string) ?? "?"; }
            catch { return "?"; }
        }

        public DocumentState GetState(dynamic doc)
        {
            if (doc == null) return new DocumentState();

            string key = KeyOf(doc);
            DocumentState state;
            if (!_states.TryGetValue(key, out state))
            {
                state = DocumentState.Load(doc);
                _states[key] = state;
            }
            return state;
        }

        public void SaveState(dynamic doc, DocumentState state)
        {
            if (doc == null || state == null) return;
            state.Save(doc);
            string cacheKey = KeyOf(doc);
            _indexes.Remove(cacheKey);
            InvalidateTarget();
        }

        public void InvalidateTarget()
        {
            _cachedTarget = null;
            _cachedTargetAt = DateTime.MinValue;
            _cachedTargetPos = -1;
        }

        // ==================================================================
        // Полная проверка документа
        // ==================================================================

        /// <summary>Разобрать документ, подсветить аббревиатуры, обновить окно.</summary>
        public void CheckDocument(bool silent = false)
        {
            dynamic doc = ActiveDocument;
            if (doc == null)
            {
                if (!silent) Warn("Откройте документ Word.");
                return;
            }

            DocumentState state = GetState(doc);
            var scanner = new DocumentScanner(Settings);
            var highlighter = new HighlightService(Settings);

            Cursor.Current = Cursors.WaitCursor;
            ScanResult scan;
            int painted;
            try
            {
                scan = scanner.Scan(App, doc, state);
                state.LastScan = scan;
                state.Save(doc);
                string cacheKey = KeyOf(doc);
                _indexes.Remove(cacheKey);

                highlighter.Clear(App, doc, scan);
                painted = highlighter.Apply(App, doc, scan);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

            InvalidateTarget();
            RefreshMainForm();

            if (silent) return;

            var sb = new StringBuilder();
            sb.AppendLine("Проверка завершена.");
            sb.AppendLine();
            sb.AppendLine("Уникальных сокращений: " + scan.ByAbbr.Count);
            sb.AppendLine("  в перечне (зелёные): " + scan.KnownCount);
            sb.AppendLine("  нет в перечне (красные): " + scan.UnknownCount);
            sb.AppendLine("  игнорируемые: " + scan.IgnoredCount);
            sb.AppendLine("Всего вхождений подсвечено: " + painted);
            sb.AppendLine();
            sb.AppendLine(scan.DictionaryPages.Count > 0
                ? "Листы перечня: " + string.Join(", ", scan.DictionaryPages)
                : "Перечень принятых сокращений не найден. Задайте его листы вручную.");

            if (highlighter.Misaligned > 0)
                sb.AppendLine().AppendLine("Не удалось подсветить вхождений: " + highlighter.Misaligned +
                                           " (обычно это текст внутри полей).");

            Info(sb.ToString());
        }

        public void ClearHighlights()
        {
            dynamic doc = ActiveDocument;
            if (doc == null) { Warn("Откройте документ Word."); return; }

            DocumentState state = GetState(doc);
            var highlighter = new HighlightService(Settings);

            if (state.LastScan == null)
            {
                var scanner = new DocumentScanner(Settings);
                state.LastScan = scanner.Scan(App, doc, state);
            }

            int cleared = highlighter.Clear(App, doc, state.LastScan);
            Info("Подсветка снята с " + cleared + " вхождений.");
        }

        public void ClearAllHighlights()
        {
            dynamic doc = ActiveDocument;
            if (doc == null) { Warn("Откройте документ Word."); return; }

            if (MessageBox.Show(Owner,
                    "Снять подсветку во всём документе, включая ту, что расставлена вручную?",
                    "Abbreviator", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            new HighlightService(Settings).ClearAll(App, doc);
        }

        // ==================================================================
        // Листы перечня
        // ==================================================================

        /// <summary>Результат автоопределения перечня.</summary>
        public sealed class PageDetection
        {
            public List<int> Pages = new List<int>();
            public int HeaderPage;

            public bool Found
            {
                get { return Pages.Count > 0; }
            }
        }

        /// <summary>
        /// Определить листы перечня автоматически.
        /// Результат возвращается объектом, а не out-параметром: doc имеет тип
        /// dynamic, а вызовы с out по dynamic не связываются.
        /// </summary>
        public PageDetection AutoDetectPages(dynamic doc, DocumentState state)
        {
            var detection = new PageDetection();
            if (doc == null) return detection;

            AbbreviationRecognizer recognizer = new DocumentScanner(Settings).CreateRecognizer(App);
            var parser = new DictionaryParser(Settings, recognizer);

            WordUtil.Repaginate(doc);
            List<int[]> bounds = WordUtil.PageBounds(doc);
            TextModel model = TextModel.Build(doc);

            int headerStart, headerPage;
            detection.Pages = parser.DetectPages(model, bounds, out headerStart, out headerPage);
            detection.HeaderPage = headerPage;
            return detection;
        }

        public void GoToDictionary()
        {
            dynamic doc = ActiveDocument;
            if (doc == null) { Warn("Откройте документ Word."); return; }

            DocumentState state = GetState(doc);
            int page = state.FirstDictionaryPage;
            if (page <= 0)
            {
                Warn("Листы перечня не заданы. Откройте «Листы перечня» и добавьте их.");
                return;
            }
            WordUtil.GoToPageInView(doc, page);
        }

        public void GoToPage(int page)
        {
            dynamic doc = ActiveDocument;
            if (doc == null) return;
            WordUtil.GoToPageInView(doc, page);
        }

        public int CurrentPage()
        {
            try
            {
                dynamic sel = App.Selection;
                return (int)sel.Information[Wd.ActiveEndPageNumber];
            }
            catch
            {
                return 0;
            }
        }

        // ==================================================================
        // Индекс принятых сокращений (лёгкий разбор, только перечень)
        // ==================================================================

        private sealed class KnownIndex
        {
            public Dictionary<string, DictEntry> Entries;
            public int ContentEnd;
            public int PagesHash;
        }

        /// <summary>
        /// Быстрый доступ к перечню без полного разбора документа.
        /// Кэш сбрасывается, если изменилась длина документа или список листов.
        /// </summary>
        public Dictionary<string, DictEntry> GetKnownIndex(dynamic doc, DocumentState state)
        {
            var empty = new Dictionary<string, DictEntry>(StringComparer.Ordinal);
            if (doc == null) return empty;

            string key = KeyOf(doc);
            int contentEnd = WordUtil.ContentEnd(doc);
            int pagesHash = string.Join(",", state.DictionaryPages).GetHashCode();

            KnownIndex cached;
            if (_indexes.TryGetValue(key, out cached) &&
                cached.ContentEnd == contentEnd && cached.PagesHash == pagesHash)
                return cached.Entries;

            var map = new Dictionary<string, DictEntry>(StringComparer.Ordinal);

            if (state.DictionaryPages.Count > 0)
            {
                AbbreviationRecognizer recognizer = new DocumentScanner(Settings).CreateRecognizer(App);
                var parser = new DictionaryParser(Settings, recognizer);
                List<int[]> bounds = WordUtil.PageBoundsFor(doc, state.DictionaryPages);
                TextModel model = TextModel.Build(doc);

                foreach (var e in parser.ParseEntries(model, bounds, state.DictionaryPages))
                    foreach (var term in e.Terms)
                        if (!map.ContainsKey(term)) map[term] = e;
            }

            _indexes[key] = new KnownIndex
            {
                Entries = map,
                ContentEnd = contentEnd,
                PagesHash = pagesHash
            };
            return map;
        }

        // ==================================================================
        // Аббревиатура под курсором
        // ==================================================================

        /// <summary>
        /// Сокращение, на котором стоит курсор. Результат кэшируется на секунду,
        /// потому что контекстное меню опрашивает надстройку несколько раз подряд.
        /// </summary>
        public Target ResolveTarget()
        {
            int position = SelectionStart();
            if (_cachedTarget != null && position == _cachedTargetPos &&
                (DateTime.Now - _cachedTargetAt).TotalMilliseconds < 1000)
                return _cachedTarget;

            var target = ResolveTargetCore();
            _cachedTarget = target;
            _cachedTargetAt = DateTime.Now;
            _cachedTargetPos = position;
            return target;
        }

        private int SelectionStart()
        {
            try { return (int)App.Selection.Start; }
            catch { return -1; }
        }

        private Target ResolveTargetCore()
        {
            var target = new Target();

            dynamic doc = ActiveDocument;
            if (doc == null) return target;

            try
            {
                dynamic sel = App.Selection;
                int selStart = (int)sel.Start;
                int selEnd = (int)sel.End;

                dynamic paraRange = sel.Range.Paragraphs[1].Range;
                int paraStart = (int)paraRange.Start;
                string raw = paraRange.Text as string ?? string.Empty;
                if (raw.Length == 0) return target;

                DocumentState state = GetState(doc);
                AbbreviationRecognizer recognizer = new DocumentScanner(Settings).CreateRecognizer(App);

                int caret = selStart - paraStart;
                TokenMatch? best = null;
                int bestDistance = int.MaxValue;

                foreach (var token in recognizer.Find(raw))
                {
                    int from = token.Index;
                    int to = token.Index + token.Length;

                    // Пересечение с выделением или курсор внутри токена.
                    bool hit = selEnd > selStart
                        ? (from < selEnd - paraStart && to > selStart - paraStart)
                        : (caret >= from && caret <= to);

                    if (!hit) continue;

                    int distance = Math.Abs(from - caret);
                    if (distance < bestDistance)
                    {
                        bestDistance = distance;
                        best = token;
                    }
                }

                if (best == null) return target;

                var match = best.Value;
                target.Abbr = match.Abbr;
                target.Start = paraStart + match.Index;
                target.End = target.Start + match.Length;

                Dictionary<string, DictEntry> known = GetKnownIndex(doc, state);
                DictEntry entry;

                if (state.IsIgnored(match.Abbr))
                {
                    target.Status = AbbrStatus.Ignored;
                }
                else if (known.TryGetValue(match.Abbr, out entry))
                {
                    target.Status = AbbrStatus.Known;
                    target.Definition = entry.Definition;
                    target.EntryStart = entry.Start;
                }
                else
                {
                    target.Status = AbbrStatus.Unknown;
                }

                if (state.LastScan != null)
                {
                    AbbrInfo info;
                    if (state.LastScan.ByAbbr.TryGetValue(match.Abbr, out info))
                        target.UsageCount = info.UsageCount;
                }
            }
            catch
            {
                return new Target();
            }

            return target;
        }

        // ==================================================================
        // Действия контекстного меню
        // ==================================================================

        /// <summary>Спросить расшифровку и дописать статью в перечень.</summary>
        public void AddTargetToDictionary()
        {
            dynamic doc = ActiveDocument;
            if (doc == null) { Warn("Откройте документ Word."); return; }

            var target = ResolveTarget();
            if (!target.HasValue)
            {
                Warn("Поставьте курсор на сокращение.");
                return;
            }

            AddAbbreviation(target.Abbr, null);
        }

        /// <summary>
        /// Добавляет сокращение в перечень: спрашивает расшифровку, вставляет
        /// статью на последний лист перечня и переходит к ней.
        /// </summary>
        public void AddAbbreviation(string abbr, string presetDefinition)
        {
            dynamic doc = ActiveDocument;
            if (doc == null) { Warn("Откройте документ Word."); return; }

            DocumentState state = GetState(doc);

            if (state.DictionaryPages.Count == 0)
            {
                PageDetection detected = AutoDetectPages(doc, state);
                if (detected.Found)
                {
                    state.SetPages(detected.Pages);
                    state.Save(doc);
                }
                else
                {
                    Warn("Перечень принятых сокращений не найден.\r\n\r\n" +
                         "Откройте «Листы перечня», добавьте страницы перечня вручную " +
                         "и повторите попытку.");
                    ShowMainForm(1);
                    return;
                }
            }

            string definition = presetDefinition;
            if (definition == null)
            {
                using (var dialog = new AddEntryForm(abbr, Settings.EntrySeparator))
                {
                    if (dialog.ShowDialog(Owner) != DialogResult.OK) return;
                    abbr = dialog.Abbreviation;
                    definition = dialog.Definition;
                }
            }

            // Для вставки нужен свежий разбор перечня.
            var scanner = new DocumentScanner(Settings);
            var scan = state.LastScan;
            if (scan == null || scan.Entries.Count == 0)
            {
                scan = scanner.Scan(App, doc, state);
                state.LastScan = scan;
            }

            var writer = new DictionaryWriter(Settings);
            WriteResult result = writer.AddEntry(App, doc, state, scan, abbr, definition);

            if (!result.Success)
            {
                Warn(result.Error ?? "Не удалось добавить статью.");
                return;
            }

            string cacheKey = KeyOf(doc);
            _indexes.Remove(cacheKey);
            InvalidateTarget();

            // Перечень изменился — пересчитываем и перекрашиваем документ.
            state.LastScan = scanner.Scan(App, doc, state);
            var highlighter = new HighlightService(Settings);
            highlighter.Clear(App, doc, state.LastScan);
            highlighter.Apply(App, doc, state.LastScan);
            state.Save(doc);

            RefreshMainForm();

            // Переносим пользователя на страницу перечня, к новой статье.
            WordUtil.ScrollTo(doc, result.Start, result.End);
        }

        /// <summary>Пометить сокращение как игнорируемое.</summary>
        public void IgnoreTarget()
        {
            var target = ResolveTarget();
            if (!target.HasValue)
            {
                Warn("Поставьте курсор на сокращение.");
                return;
            }
            IgnoreAbbreviation(target.Abbr);
        }

        public void IgnoreAbbreviation(string abbr)
        {
            dynamic doc = ActiveDocument;
            if (doc == null || string.IsNullOrWhiteSpace(abbr)) return;

            DocumentState state = GetState(doc);
            state.Ignored.Add(abbr.Trim());
            state.Save(doc);
            InvalidateTarget();

            // Снимаем подсветку со всех вхождений этого сокращения.
            if (state.LastScan != null)
            {
                foreach (var occ in state.LastScan.Occurrences.Where(o => o.Abbr == abbr))
                {
                    occ.Status = AbbrStatus.Ignored;
                    try { doc.Range(occ.Start, occ.End).HighlightColorIndex = Wd.NoHighlight; }
                    catch { }
                }

                AbbrInfo info;
                if (state.LastScan.ByAbbr.TryGetValue(abbr, out info)) info.Status = AbbrStatus.Ignored;
            }

            RefreshMainForm();
        }

        public void UnignoreAbbreviation(string abbr)
        {
            dynamic doc = ActiveDocument;
            if (doc == null || string.IsNullOrWhiteSpace(abbr)) return;

            DocumentState state = GetState(doc);
            state.Ignored.Remove(abbr.Trim());
            Settings.GlobalIgnored.Remove(abbr.Trim());
            state.Save(doc);
            InvalidateTarget();
            RefreshMainForm();
        }

        /// <summary>Перейти к статье перечня для сокращения под курсором.</summary>
        public void GoToTargetEntry()
        {
            dynamic doc = ActiveDocument;
            if (doc == null) return;

            var target = ResolveTarget();
            if (!target.HasValue || target.EntryStart < 0)
            {
                Warn("Статья для этого сокращения в перечне не найдена.");
                return;
            }
            WordUtil.ScrollTo(doc, target.EntryStart, target.EntryStart);
        }

        public void GoToOccurrence(int start, int end)
        {
            dynamic doc = ActiveDocument;
            if (doc == null) return;
            WordUtil.ScrollTo(doc, start, end);
        }

        // ==================================================================
        // Окна
        // ==================================================================

        public IWin32Window Owner
        {
            get { return WordOwner.Current; }
        }

        public void ShowMainForm(int tabIndex = 0)
        {
            if (_mainForm == null || _mainForm.IsDisposed)
                _mainForm = new MainForm(this);

            _mainForm.SelectTab(tabIndex);
            if (!_mainForm.Visible) _mainForm.Show(Owner);
            _mainForm.ReloadFromDocument();
            _mainForm.BringToFront();
        }

        public void RefreshMainForm()
        {
            if (_mainForm != null && !_mainForm.IsDisposed && _mainForm.Visible)
                _mainForm.ReloadFromDocument();
        }

        public void ShowIgnoreForm()
        {
            dynamic doc = ActiveDocument;
            if (doc == null) { Warn("Откройте документ Word."); return; }

            DocumentState state = GetState(doc);
            using (var dialog = new IgnoreForm(state, Settings))
            {
                if (dialog.ShowDialog(Owner) == DialogResult.OK)
                {
                    state.Save(doc);
                    Settings.Save();
                    InvalidateTarget();
                    CheckDocument(true);
                }
            }
        }

        public void ShowSettingsForm()
        {
            using (var dialog = new SettingsForm(Settings))
            {
                if (dialog.ShowDialog(Owner) == DialogResult.OK)
                {
                    Settings.Save();
                    _indexes.Clear();
                    InvalidateTarget();
                }
            }
        }

        public void ShowAbout()
        {
            Info("Abbreviator — надстройка Word для проверки аббревиатур.\r\n\r\n" +
                 "Зелёный фон — сокращение есть в перечне принятых сокращений.\r\n" +
                 "Красный фон — сокращения в перечне нет.\r\n\r\n" +
                 "Правая кнопка мыши на сокращении — добавить его в перечень " +
                 "или в игнорируемые.\r\n\r\n" +
                 "Файл настроек: " + AppSettings.FilePath);
        }

        public void OpenSettingsFile()
        {
            try
            {
                Settings.Save();
                Process.Start(AppSettings.FilePath);
            }
            catch (Exception ex)
            {
                Warn("Не удалось открыть файл настроек: " + ex.Message);
            }
        }

        // ==================================================================

        public void Info(string text)
        {
            MessageBox.Show(Owner, text, "Abbreviator", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Warn(string text)
        {
            MessageBox.Show(Owner, text, "Abbreviator", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
