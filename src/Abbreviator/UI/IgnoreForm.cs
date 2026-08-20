using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Abbreviator.Core;

namespace Abbreviator.UI
{
    /// <summary>
    /// Игнорируемые сокращения: отдельно для документа и общие для всех документов.
    /// </summary>
    public sealed class IgnoreForm : Form
    {
        private readonly DocumentState _state;
        private readonly AppSettings _settings;

        private readonly ListBox _docList;
        private readonly ListBox _globalList;

        public IgnoreForm(DocumentState state, AppSettings settings)
        {
            _state = state;
            _settings = settings;

            Text = "Игнорируемые сокращения";
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(620, 400);
            MinimumSize = new Size(560, 360);
            Font = SystemFonts.MessageBoxFont;

            var docLabel = new Label
            {
                Text = "В этом документе:",
                Location = new Point(12, 12),
                AutoSize = true
            };

            _docList = new ListBox
            {
                Location = new Point(12, 34),
                Size = new Size(260, 280),
                SelectionMode = SelectionMode.MultiExtended,
                IntegralHeight = false
            };

            var globalLabel = new Label
            {
                Text = "Во всех документах:",
                Location = new Point(292, 12),
                AutoSize = true
            };

            _globalList = new ListBox
            {
                Location = new Point(292, 34),
                Size = new Size(260, 280),
                SelectionMode = SelectionMode.MultiExtended,
                IntegralHeight = false
            };

            var addDoc = Small("+", 12, 320, 40);
            addDoc.Click += (s, e) => AddTo(_docList, _state.Ignored);

            var delDoc = Small("−", 58, 320, 40);
            delDoc.Click += (s, e) => RemoveFrom(_docList, _state.Ignored);

            var toGlobal = Small("→ во все", 104, 320, 90);
            toGlobal.Click += OnMoveToGlobal;

            var addGlobal = Small("+", 292, 320, 40);
            addGlobal.Click += (s, e) => AddTo(_globalList, _settings.GlobalIgnored);

            var delGlobal = Small("−", 338, 320, 40);
            delGlobal.Click += (s, e) => RemoveFrom(_globalList, _settings.GlobalIgnored);

            var ok = new Button
            {
                Text = "ОК",
                DialogResult = DialogResult.OK,
                Location = new Point(430, 358),
                Width = 84
            };

            var cancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new Point(522, 358),
                Width = 84
            };

            var hint = new Label
            {
                Text = "Игнорируемые сокращения не подсвечиваются и не требуют статьи в перечне.",
                Location = new Point(12, 358),
                Size = new Size(410, 34),
                ForeColor = SystemColors.GrayText
            };

            Controls.AddRange(new Control[]
            {
                docLabel, _docList, globalLabel, _globalList,
                addDoc, delDoc, toGlobal, addGlobal, delGlobal,
                hint, ok, cancel
            });

            AcceptButton = ok;
            CancelButton = cancel;

            Reload();
        }

        private static Button Small(string text, int x, int y, int width)
        {
            return new Button { Text = text, Location = new Point(x, y), Size = new Size(width, 26) };
        }

        private void Reload()
        {
            _docList.BeginUpdate();
            _docList.Items.Clear();
            foreach (var item in _state.Ignored.OrderBy(x => x, StringComparer.CurrentCulture))
                _docList.Items.Add(item);
            _docList.EndUpdate();

            _globalList.BeginUpdate();
            _globalList.Items.Clear();
            foreach (var item in _settings.GlobalIgnored.OrderBy(x => x, StringComparer.CurrentCulture))
                _globalList.Items.Add(item);
            _globalList.EndUpdate();
        }

        private void AddTo(ListBox list, System.Collections.Generic.HashSet<string> target)
        {
            string value = Prompt.Show(this, "Сокращение, которое нужно игнорировать:",
                                       "Добавить", string.Empty);
            if (string.IsNullOrWhiteSpace(value)) return;

            foreach (var part in value.Split(',', ';', ' '))
            {
                var t = part.Trim();
                if (t.Length > 0) target.Add(t);
            }
            Reload();
        }

        private void RemoveFrom(ListBox list, System.Collections.Generic.HashSet<string> target)
        {
            foreach (var item in list.SelectedItems.Cast<object>().ToList())
                target.Remove(item as string);
            Reload();
        }

        private void OnMoveToGlobal(object sender, EventArgs e)
        {
            foreach (var item in _docList.SelectedItems.Cast<object>().ToList())
            {
                var t = item as string;
                if (string.IsNullOrEmpty(t)) continue;
                _settings.GlobalIgnored.Add(t);
                _state.Ignored.Remove(t);
            }
            Reload();
        }
    }
}
