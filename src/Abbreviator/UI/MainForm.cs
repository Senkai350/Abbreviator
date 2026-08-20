using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Abbreviator.Core;

namespace Abbreviator.UI
{
    /// <summary>
    /// Главное окно надстройки.
    ///
    /// Вкладка «Листы перечня» — список страниц, с которых берутся принятые
    /// сокращения. Перечень может начинаться на одной странице и продолжаться
    /// на следующих, поэтому список ведётся явно и правится вручную.
    ///
    /// Вкладка «Сокращения» — что нашлось в документе и в каком оно статусе.
    /// </summary>
    public sealed class MainForm : Form
    {
        private readonly AddInController _controller;

        private readonly TabControl _tabs;
        private readonly ListBox _pagesList;
        private readonly Label _pagesInfo;
        private readonly ListView _resultsList;
        private readonly CheckBox _onlyUnknown;
        private readonly Label _summary;

        private bool _loading;

        public MainForm(AddInController controller)
        {
            _controller = controller;

            Text = "Abbreviator — аббревиатуры и перечень сокращений";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(720, 470);
            MinimumSize = new Size(640, 420);
            ShowInTaskbar = false;
            Font = SystemFonts.MessageBoxFont;

            _tabs = new TabControl { Dock = DockStyle.Fill };

            // ---------------- Вкладка 1: листы перечня ----------------
            var pagesTab = new TabPage("Листы перечня");

            _pagesList = new ListBox
            {
                Location = new Point(12, 12),
                Size = new Size(220, 320),
                SelectionMode = SelectionMode.MultiExtended,
                IntegralHeight = false
            };
            _pagesList.DoubleClick += (s, e) => GoToSelectedPage();

            var addCurrent = MakeButton("Добавить текущую страницу", 248, 12, 250);
            addCurrent.Click += OnAddCurrentPage;

            var addManual = MakeButton("Добавить страницу или диапазон…", 248, 44, 250);
            addManual.Click += OnAddManualPages;

            var remove = MakeButton("Удалить выбранные", 248, 76, 250);
            remove.Click += OnRemovePages;

            var clear = MakeButton("Очистить список", 248, 108, 250);
            clear.Click += OnClearPages;

            var detect = MakeButton("Определить автоматически", 248, 150, 250);
            detect.Click += OnAutoDetect;

            var goTo = MakeButton("Перейти к странице", 248, 182, 250);
            goTo.Click += (s, e) => GoToSelectedPage();

            _pagesInfo = new Label
            {
                Location = new Point(248, 220),
                Size = new Size(440, 112),
                ForeColor = SystemColors.GrayText
            };

            pagesTab.Controls.AddRange(new Control[]
            {
                _pagesList, addCurrent, addManual, remove, clear, detect, goTo, _pagesInfo
            });

            // ---------------- Вкладка 2: найденные сокращения ----------------
            var resultsTab = new TabPage("Сокращения");

            _resultsList = new ListView
            {
                Location = new Point(12, 12),
                Size = new Size(480, 320),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = false,
                HideSelection = false
            };
            _resultsList.Columns.Add("Сокращение", 120);
            _resultsList.Columns.Add("Исп.", 45, HorizontalAlignment.Right);
            _resultsList.Columns.Add("Статус", 120);
            _resultsList.Columns.Add("Расшифровка", 185);
            _resultsList.DoubleClick += (s, e) => GoToSelectedAbbr();

            _onlyUnknown = new CheckBox
            {
                Text = "Только те, которых нет в перечне",
                Location = new Point(12, 338),
                AutoSize = true
            };
            _onlyUnknown.CheckedChanged += (s, e) => FillResults();

            var goToUse = MakeButton("Перейти к вхождению", 508, 12, 180);
            goToUse.Click += (s, e) => GoToSelectedAbbr();

            var addToDict = MakeButton("Добавить в перечень…", 508, 44, 180);
            addToDict.Click += OnAddSelectedToDictionary;

            var ignore = MakeButton("Игнорировать", 508, 76, 180);
            ignore.Click += OnIgnoreSelected;

            var unignore = MakeButton("Снять игнорирование", 508, 108, 180);
            unignore.Click += OnUnignoreSelected;

            var goToEntry = MakeButton("Показать в перечне", 508, 150, 180);
            goToEntry.Click += OnGoToEntry;

            resultsTab.Controls.AddRange(new Control[]
            {
                _resultsList, _onlyUnknown, goToUse, addToDict, ignore, unignore, goToEntry
            });

            _tabs.TabPages.Add(pagesTab);
            _tabs.TabPages.Add(resultsTab);

            // ---------------- Нижняя панель ----------------
            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 46 };

            var check = MakeButton("Проверить документ", 12, 10, 160);
            check.Click += (s, e) => { _controller.CheckDocument(true); ReloadFromDocument(); };

            var clearHl = MakeButton("Снять подсветку", 180, 10, 140);
            clearHl.Click += (s, e) => _controller.ClearHighlights();

            _summary = new Label
            {
                Location = new Point(332, 15),
                Size = new Size(270, 20),
                ForeColor = SystemColors.GrayText
            };

            var close = MakeButton("Закрыть", 610, 10, 90);
            close.Click += (s, e) => Hide();
            close.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            bottom.Controls.AddRange(new Control[] { check, clearHl, _summary, close });

            Controls.Add(_tabs);
            Controls.Add(bottom);

            FormClosing += (s, e) =>
            {
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                    Hide();
                }
            };
        }

        private static Button MakeButton(string text, int x, int y, int width)
        {
            return new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, 26),
                TextAlign = ContentAlignment.MiddleCenter
            };
        }

        public void SelectTab(int index)
        {
            if (index >= 0 && index < _tabs.TabPages.Count) _tabs.SelectedIndex = index;
        }

        // ==================================================================

        public void ReloadFromDocument()
        {
            _loading = true;
            try
            {
                FillPages();
                FillResults();
                FillSummary();
            }
            finally
            {
                _loading = false;
            }
        }

        private DocumentState CurrentState
        {
            get
            {
                dynamic doc = _controller.ActiveDocument;
                return doc == null ? null : _controller.GetState(doc);
            }
        }

        private void FillPages()
        {
            var state = CurrentState;
            _pagesList.BeginUpdate();
            _pagesList.Items.Clear();

            if (state != null)
                foreach (var page in state.DictionaryPages)
                    _pagesList.Items.Add("Страница " + page);

            _pagesList.EndUpdate();

            if (state == null)
            {
                _pagesInfo.Text = "Нет открытого документа.";
                return;
            }

            _pagesInfo.Text =
                "Сокращения с этих страниц считаются принятыми и подсвечиваются зелёным.\r\n\r\n" +
                "Перечень может начинаться на одной странице и продолжаться на следующих — " +
                "добавьте сюда все страницы продолжения, иначе сокращения с них будут " +
                "помечены красным.\r\n\r\n" +
                (state.DictionaryPages.Count == 0
                    ? "Список пуст. Нажмите «Определить автоматически»."
                    : "Листов в перечне: " + state.DictionaryPages.Count);
        }

        private void FillResults()
        {
            var state = CurrentState;
            _resultsList.BeginUpdate();
            _resultsList.Items.Clear();

            if (state != null && state.LastScan != null)
            {
                var rows = state.LastScan.ByAbbr.Values
                    .Where(i => !_onlyUnknown.Checked || i.Status == AbbrStatus.Unknown)
                    .OrderBy(i => i.Status == AbbrStatus.Unknown ? 0 : 1)
                    .ThenBy(i => i.Abbr, StringComparer.CurrentCulture)
                    .ToList();

                foreach (var info in rows)
                {
                    var item = new ListViewItem(info.Abbr);
                    item.SubItems.Add(info.UsageCount.ToString());
                    item.SubItems.Add(StatusText(info));
                    item.SubItems.Add(info.Definition ?? string.Empty);
                    item.Tag = info;

                    switch (info.Status)
                    {
                        case AbbrStatus.Known:
                            item.BackColor = Color.FromArgb(226, 246, 226);
                            break;
                        case AbbrStatus.Unknown:
                            item.BackColor = Color.FromArgb(255, 228, 228);
                            break;
                        default:
                            item.ForeColor = SystemColors.GrayText;
                            break;
                    }

                    _resultsList.Items.Add(item);
                }
            }

            _resultsList.EndUpdate();
        }

        private static string StatusText(AbbrInfo info)
        {
            switch (info.Status)
            {
                case AbbrStatus.Known:
                    return info.IsUnused ? "в перечне, не исп." : "в перечне";
                case AbbrStatus.Unknown:
                    return "нет в перечне";
                default:
                    return "игнорируется";
            }
        }

        private void FillSummary()
        {
            var state = CurrentState;
            if (state == null || state.LastScan == null)
            {
                _summary.Text = "Документ ещё не проверялся.";
                return;
            }

            var scan = state.LastScan;
            _summary.Text = string.Format("Всего {0} · в перечне {1} · нет {2} · игнор. {3}",
                scan.ByAbbr.Count, scan.KnownCount, scan.UnknownCount, scan.IgnoredCount);
        }

        // ==================================================================
        // Листы перечня
        // ==================================================================

        private void OnAddCurrentPage(object sender, EventArgs e)
        {
            dynamic doc = _controller.ActiveDocument;
            if (doc == null) return;

            int page = _controller.CurrentPage();
            if (page <= 0)
            {
                MessageBox.Show(this, "Не удалось определить текущую страницу.", "Abbreviator",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DocumentState state = _controller.GetState(doc);
            state.AddPage(page);
            _controller.SaveState(doc, state);
            ReloadFromDocument();
        }

        private void OnAddManualPages(object sender, EventArgs e)
        {
            dynamic doc = _controller.ActiveDocument;
            if (doc == null) return;

            string input = Prompt.Show(this,
                "Номера страниц перечня через запятую.\r\nМожно указывать диапазоны: 12, 13-15",
                "Добавить страницы", string.Empty);
            if (string.IsNullOrWhiteSpace(input)) return;

            DocumentState state = _controller.GetState(doc);
            foreach (var page in ParsePages(input)) state.AddPage(page);
            _controller.SaveState(doc, state);
            ReloadFromDocument();
        }

        private static IEnumerable<int> ParsePages(string input)
        {
            foreach (var chunk in input.Split(',', ';', ' '))
            {
                var t = chunk.Trim();
                if (t.Length == 0) continue;

                int dash = t.IndexOf('-');
                if (dash > 0 && dash < t.Length - 1)
                {
                    int a, b;
                    if (int.TryParse(t.Substring(0, dash), out a) &&
                        int.TryParse(t.Substring(dash + 1), out b) && a > 0 && b >= a)
                    {
                        for (int p = a; p <= b; p++) yield return p;
                        continue;
                    }
                }

                int one;
                if (int.TryParse(t, out one) && one > 0) yield return one;
            }
        }

        private void OnRemovePages(object sender, EventArgs e)
        {
            dynamic doc = _controller.ActiveDocument;
            if (doc == null) return;

            DocumentState state = _controller.GetState(doc);
            foreach (int index in _pagesList.SelectedIndices.Cast<int>().OrderByDescending(i => i))
            {
                if (index >= 0 && index < state.DictionaryPages.Count)
                    state.RemovePage(state.DictionaryPages[index]);
            }
            _controller.SaveState(doc, state);
            ReloadFromDocument();
        }

        private void OnClearPages(object sender, EventArgs e)
        {
            dynamic doc = _controller.ActiveDocument;
            if (doc == null) return;

            DocumentState state = _controller.GetState(doc);
            state.DictionaryPages.Clear();
            _controller.SaveState(doc, state);
            ReloadFromDocument();
        }

        private void OnAutoDetect(object sender, EventArgs e)
        {
            dynamic doc = _controller.ActiveDocument;
            if (doc == null) return;

            DocumentState state = _controller.GetState(doc);
            Cursor.Current = Cursors.WaitCursor;
            AddInController.PageDetection detection;
            try
            {
                detection = _controller.AutoDetectPages(doc, state);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

            if (!detection.Found)
            {
                MessageBox.Show(this,
                    "Заголовок перечня не найден.\r\n\r\n" +
                    "Проверьте, что в документе есть строка «Перечень принятых сокращений», " +
                    "либо добавьте страницы вручную. Список допустимых заголовков " +
                    "настраивается в файле настроек.",
                    "Abbreviator", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var page in detection.Pages) state.AddPage(page);
            _controller.SaveState(doc, state);
            ReloadFromDocument();

            MessageBox.Show(this,
                "Перечень найден на странице " + detection.HeaderPage + ".\r\n" +
                "Добавлены листы: " + string.Join(", ", detection.Pages) + ".\r\n\r\n" +
                "Если перечень продолжается дальше, добавьте оставшиеся страницы вручную.",
                "Abbreviator", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GoToSelectedPage()
        {
            var state = CurrentState;
            if (state == null) return;

            int index = _pagesList.SelectedIndex;
            if (index < 0 || index >= state.DictionaryPages.Count) return;

            _controller.GoToPage(state.DictionaryPages[index]);
        }

        // ==================================================================
        // Найденные сокращения
        // ==================================================================

        private AbbrInfo SelectedInfo
        {
            get
            {
                if (_resultsList.SelectedItems.Count == 0) return null;
                return _resultsList.SelectedItems[0].Tag as AbbrInfo;
            }
        }

        private void GoToSelectedAbbr()
        {
            var info = SelectedInfo;
            if (info == null) return;

            if (info.FirstUse < 0)
            {
                MessageBox.Show(this, "Это сокращение объявлено в перечне, но в тексте не встречается.",
                                "Abbreviator", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _controller.GoToOccurrence(info.FirstUse, info.FirstUse + info.Abbr.Length);
        }

        private void OnAddSelectedToDictionary(object sender, EventArgs e)
        {
            var info = SelectedInfo;
            if (info == null) return;
            _controller.AddAbbreviation(info.Abbr, null);
            ReloadFromDocument();
        }

        private void OnIgnoreSelected(object sender, EventArgs e)
        {
            var info = SelectedInfo;
            if (info == null) return;
            _controller.IgnoreAbbreviation(info.Abbr);
            ReloadFromDocument();
        }

        private void OnUnignoreSelected(object sender, EventArgs e)
        {
            var info = SelectedInfo;
            if (info == null) return;
            _controller.UnignoreAbbreviation(info.Abbr);
            _controller.CheckDocument(true);
            ReloadFromDocument();
        }

        private void OnGoToEntry(object sender, EventArgs e)
        {
            var info = SelectedInfo;
            if (info == null) return;

            if (info.EntryStart < 0)
            {
                MessageBox.Show(this, "Статьи для этого сокращения в перечне нет.", "Abbreviator",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _controller.GoToOccurrence(info.EntryStart, info.EntryStart);
        }
    }
}
