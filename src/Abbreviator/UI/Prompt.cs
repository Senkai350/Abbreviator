using System;
using System.Drawing;
using System.Windows.Forms;

namespace Abbreviator.UI
{
    /// <summary>Простой диалог ввода строки — замена InputBox.</summary>
    public static class Prompt
    {
        public static string Show(IWin32Window owner, string message, string caption, string initial)
        {
            using (var form = new Form())
            using (var label = new Label())
            using (var input = new TextBox())
            using (var ok = new Button())
            using (var cancel = new Button())
            {
                form.Text = caption;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.StartPosition = FormStartPosition.CenterParent;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.ShowInTaskbar = false;
                form.ClientSize = new Size(420, 150);
                form.Font = SystemFonts.MessageBoxFont;

                label.Text = message;
                label.Location = new Point(12, 12);
                label.Size = new Size(396, 60);

                input.Text = initial ?? string.Empty;
                input.Location = new Point(12, 76);
                input.Width = 396;

                ok.Text = "ОК";
                ok.DialogResult = DialogResult.OK;
                ok.Location = new Point(232, 110);
                ok.Width = 84;

                cancel.Text = "Отмена";
                cancel.DialogResult = DialogResult.Cancel;
                cancel.Location = new Point(324, 110);
                cancel.Width = 84;

                form.Controls.AddRange(new Control[] { label, input, ok, cancel });
                form.AcceptButton = ok;
                form.CancelButton = cancel;

                return form.ShowDialog(owner) == DialogResult.OK ? input.Text : null;
            }
        }
    }
}
