using System;
using System.Drawing;
using System.Windows.Forms;

namespace Abbreviator.UI
{
    /// <summary>
    /// Ввод расшифровки: «[Сокращение] - [пользователь вводит сам]».
    /// </summary>
    public sealed class AddEntryForm : Form
    {
        private readonly TextBox _abbrBox;
        private readonly TextBox _defBox;
        private readonly Label _preview;
        private readonly string _separator;

        public string Abbreviation
        {
            get { return _abbrBox.Text.Trim(); }
        }

        public string Definition
        {
            get { return _defBox.Text.Trim(); }
        }

        public AddEntryForm(string abbr, string separator)
        {
            _separator = string.IsNullOrEmpty(separator) ? " - " : separator;

            Text = "Добавить сокращение в перечень";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(460, 210);
            Font = SystemFonts.MessageBoxFont;

            var abbrLabel = new Label
            {
                Text = "Сокращение:",
                Location = new Point(12, 15),
                AutoSize = true
            };

            _abbrBox = new TextBox
            {
                Text = abbr ?? string.Empty,
                Location = new Point(12, 34),
                Width = 160,
                CharacterCasing = CharacterCasing.Normal
            };

            var defLabel = new Label
            {
                Text = "Расшифровка:",
                Location = new Point(12, 68),
                AutoSize = true
            };

            _defBox = new TextBox
            {
                Location = new Point(12, 87),
                Width = 436,
                Height = 46,
                Multiline = true
            };

            _preview = new Label
            {
                Location = new Point(12, 140),
                Width = 436,
                Height = 18,
                ForeColor = SystemColors.GrayText,
                AutoEllipsis = true
            };

            var ok = new Button
            {
                Text = "Добавить",
                DialogResult = DialogResult.OK,
                Location = new Point(272, 170),
                Width = 84
            };

            var cancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new Point(364, 170),
                Width = 84
            };

            _abbrBox.TextChanged += (s, e) => UpdatePreview();
            _defBox.TextChanged += (s, e) => UpdatePreview();
            ok.Click += OnOk;

            Controls.AddRange(new Control[] { abbrLabel, _abbrBox, defLabel, _defBox, _preview, ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;

            UpdatePreview();
            ActiveControl = _defBox;
        }

        private void UpdatePreview()
        {
            _preview.Text = "В перечень будет добавлено:  " +
                            _abbrBox.Text.Trim() + _separator + _defBox.Text.Trim();
        }

        private void OnOk(object sender, EventArgs e)
        {
            if (Abbreviation.Length == 0)
            {
                MessageBox.Show(this, "Укажите сокращение.", "Abbreviator",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            if (Definition.Length == 0)
            {
                var answer = MessageBox.Show(this,
                    "Расшифровка не заполнена. Добавить статью без неё?",
                    "Abbreviator", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (answer != DialogResult.Yes) DialogResult = DialogResult.None;
            }
        }
    }
}
