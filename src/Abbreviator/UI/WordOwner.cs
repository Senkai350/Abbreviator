using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace Abbreviator.UI
{
    /// <summary>
    /// Владелец модальных окон — главное окно Word.
    /// Без него диалоги умеют уходить за окно Word.
    /// </summary>
    public sealed class WordOwner : IWin32Window
    {
        private readonly IntPtr _handle;

        private WordOwner(IntPtr handle)
        {
            _handle = handle;
        }

        public IntPtr Handle
        {
            get { return _handle; }
        }

        public static IWin32Window Current
        {
            get
            {
                try
                {
                    IntPtr h = Process.GetCurrentProcess().MainWindowHandle;
                    if (h != IntPtr.Zero) return new WordOwner(h);
                }
                catch { }
                return null;
            }
        }
    }
}
