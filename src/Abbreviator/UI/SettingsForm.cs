using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Abbreviator.Core;
using Abbreviator.Interop;

namespace Abbreviator.UI
{
    /// <summary>Настройки распознавания, подсветки и перечня.</summary>
    public sealed class SettingsForm : Form
    {
        private readonly AppSettings _settings;

        private readonly NumericUpDown _minLetters;
        private readonly NumericUpDown _maxLength;
        private readonly CheckBox _skipRoman;
        private readonly CheckBox _skipCaps;
        private readonly CheckBox _useSpell;
        private readonly NumericUpDown _spellMin;

        private readonly ComboBox _knownColor;
        private readonly ComboBox _unknownColor;
        private readonly CheckBox _highlightDict;
        private readonly CheckBox _highlightIgnored;

        private readonly TextBox _separator;
        private readonly CheckBox _alphabetical;
        private readonly CheckBox _autoDetect;
        private readonly TextBox _headers;

        private static readonly int[] Colors =
        {
            Wd.BrightGreen, Wd.Green, Wd.Turquoise, Wd.Teal, Wd.Yellow, Wd.DarkYellow,
            Wd.Red, Wd.DarkRed, Wd.Pink, Wd.Violet, Wd.Blue, Wd.DarkBlue,
            Wd.Gray25, Wd.Gray50, Wd.NoHighlight
        };

        public SettingsForm(AppSettings settings)
        {
            _settings = settings;

            Text = "Настройки Abbreviator";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(560, 600);
            Font = SystemFonts.MessageBoxFont;

            int y = 12;

            Controls.Add(Group("Распознавание", 12, y, 536, 168));
            y += 26;

            Controls.Add(Lab("Минимум букв в сокращении:", 28, y + 3));
            _minLetters = Num(320, y, 1, 6, settings.MinLetters);
            Controls.Add(_minLetters);
            y += 30;

            Controls.Add(Lab("Максимальная длина:", 28, y + 3));
            _maxLength = Num(320, y, 2, 40, settings.MaxLength);
            Controls.Add(_maxLength);
            y += 30;

            _skipRoman = Check("Пропускать римские цифры (IV, XII)", 28, y, settings.SkipRomanNumerals);
            Controls.Add(_skipRoman);
            y += 26;

            _skipCaps = Check("Пропускать заголовки, набранные ПРОПИСНЫМИ", 28, y, settings.SkipUpperCaseParagraphs);
            Controls.Add(_skipCaps);
            y += 26;

            _useSpell = Check("Отсеивать обычные слова проверкой орфографии Word", 28, y, settings.UseSpellCheckFilter);
            Controls.Add(_useSpell);
            y += 26;

            Controls.Add(Lab("Проверять орфографией слова длиннее:", 28, y + 3));
            _spellMin = Num(320, y, 2, 20, settings.SpellCheckMinLength);
            Controls.Add(_spellMin);
            y += 44;

            Controls.Add(Group("Подсветка", 12, y, 536, 116));
            y += 26;

            Controls.Add(Lab("Есть в перечне:", 28, y + 3));
            _knownColor = Combo(320, y, settings.KnownColor);
            Controls.Add(_knownColor);
            y += 30;

            Controls.Add(Lab("Нет в перечне:", 28, y + 3));
            _unknownColor = Combo(320, y, settings.UnknownColor);
            Controls.Add(_unknownColor);
            y += 30;

            _highlightDict = Check("Подсвечивать сокращения на страницах самого перечня", 28, y, settings.HighlightDictionarySection);
            Controls.Add(_highlightDict);
            y += 24;

            _highlightIgnored = Check("Помечать игнорируемые серым", 28, y, settings.HighlightIgnored);
            Controls.Add(_highlightIgnored);
            y += 42;

            Controls.Add(Group("Перечень", 12, y, 536, 150));
            y += 26;

            Controls.Add(Lab("Разделитель статьи:", 28, y + 3));
            _separator = new TextBox
            {
                Location = new Point(320, y),
                Width = 100,
                Text = settings.EntrySeparator
            };
            Controls.Add(_separator);
            Controls.Add(new Label
            {
                Text = "напр. \" - \"",
                Location = new Point(428, y + 3),
                AutoSize = true,
                ForeColor = SystemColors.GrayText
            });
            y += 30;

            _alphabetical = Check("Вставлять статьи в алфавитном порядке", 28, y, settings.KeepAlphabeticalOrder);
            Controls.Add(_alphabetical);
            y += 24;

            _autoDetect = Check("Определять листы перечня автоматически при проверке", 28, y, settings.AutoDetectPagesOnScan);
            Controls.Add(_autoDetect);
            y += 28;

            Controls.Add(Lab("Варианты заголовка (по одному в строке):", 28, y));
            y += 20;

            _headers = new TextBox
            {
                Location = new Point(28, y),
                Size = new Size(504, 46),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Text = string.Join("\r\n", settings.HeaderVariants)
            };
            Controls.Add(_headers);
            y += 60;

            var reset = new Button
            {
                Text = "Сбросить",
                Location = new Point(12, y),
                Width = 100
            };
            reset.Click += OnReset;

            var ok = new Button
            {
                Text = "ОК",
                DialogResult = DialogResult.OK,
                Location = new Point(372, y),
                Width = 84
            };
            ok.Click += OnOk;

            var cancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new Point(464, y),
                Width = 84
            };

            Controls.AddRange(new Control[] { reset, ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;
        }

        // ------------------------------------------------------------------

        /// <summary>
        /// Заголовок раздела. Намеренно не GroupBox: элементы раскладываются по
        /// абсолютным координатам формы, и рамка перекрывала бы их по z-порядку.
        /// </summary>
        private static Label Group(string text, int x, int y, int w, int h)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, 20),
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                ForeColor = SystemColors.HotTrack
            };
        }

        private static Label Lab(string text, int x, int y)
        {
            return new Label { Text = text, Location = new Point(x, y), AutoSize = true };
        }

        private static NumericUpDown Num(int x, int y, int min, int max, int value)
        {
            return new NumericUpDown
            {
                Location = new Point(x, y),
                Width = 70,
                Minimum = min,
                Maximum = max,
                Value = Math.Min(Math.Max(value, min), max)
            };
        }

        private static CheckBox Check(string text, int x, int y, bool value)
        {
            return new CheckBox
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                Checked = value
            };
        }

        private static ComboBox Combo(int x, int y, int selected)
        {
            var box = new ComboBox
            {
                Location = new Point(x, y),
                Width = 160,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            foreach (int c in Colors) box.Items.Add(Wd.ColorName(c));
            int index = Array.IndexOf(Colors, selected);
            box.SelectedIndex = index >= 0 ? index : 0;
            return box;
        }

        private static int ColorFrom(ComboBox box)
        {
            int index = box.SelectedIndex;
            return index >= 0 && index < Colors.Length ? Colors[index] : Wd.NoHighlight;
        }

        // ------------------------------------------------------------------

        private void OnOk(object sender, EventArgs e)
        {
            _settings.MinLetters = (int)_minLetters.Value;
            _settings.MaxLength = (int)_maxLength.Value;
            _settings.SkipRomanNumerals = _skipRoman.Checked;
            _settings.SkipUpperCaseParagraphs = _skipCaps.Checked;
            _settings.UseSpellCheckFilter = _useSpell.Checked;
            _settings.SpellCheckMinLength = (int)_spellMin.Value;

            _settings.KnownColor = ColorFrom(_knownColor);
            _settings.UnknownColor = ColorFrom(_unknownColor);
            _settings.HighlightDictionarySection = _highlightDict.Checked;
            _settings.HighlightIgnored = _highlightIgnored.Checked;

            _settings.EntrySeparator = string.IsNullOrEmpty(_separator.Text) ? " - " : _separator.Text;
            _settings.KeepAlphabeticalOrder = _alphabetical.Checked;
            _settings.AutoDetectPagesOnScan = _autoDetect.Checked;

            var headers = _headers.Lines
                .Select(l => l.Trim())
                .Where(l => l.Length > 0)
                .ToList();
            if (headers.Count > 0) _settings.HeaderVariants = headers;
        }

        private void OnReset(object sender, EventArgs e)
        {
            if (MessageBox.Show(this, "Вернуть настройки по умолчанию?", "Abbreviator",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            var defaults = new AppSettings();

            _minLetters.Value = defaults.MinLetters;
            _maxLength.Value = defaults.MaxLength;
            _skipRoman.Checked = defaults.SkipRomanNumerals;
            _skipCaps.Checked = defaults.SkipUpperCaseParagraphs;
            _useSpell.Checked = defaults.UseSpellCheckFilter;
            _spellMin.Value = defaults.SpellCheckMinLength;

            _knownColor.SelectedIndex = Math.Max(0, Array.IndexOf(Colors, defaults.KnownColor));
            _unknownColor.SelectedIndex = Math.Max(0, Array.IndexOf(Colors, defaults.UnknownColor));
            _highlightDict.Checked = defaults.HighlightDictionarySection;
            _highlightIgnored.Checked = defaults.HighlightIgnored;

            _separator.Text = defaults.EntrySeparator;
            _alphabetical.Checked = defaults.KeepAlphabeticalOrder;
            _autoDetect.Checked = defaults.AutoDetectPagesOnScan;
            _headers.Text = string.Join("\r\n", defaults.HeaderVariants);
        }
    }
}
