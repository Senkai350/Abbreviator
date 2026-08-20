using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Abbreviator.Core;
using Abbreviator.Interop;

namespace Abbreviator
{
    /// <summary>
    /// Точка входа COM-надстройки Word.
    ///
    /// Класс помечен AutoDual: обратные вызовы ленты и контекстного меню Office
    /// ищет по имени через IDispatch, поэтому все они должны быть публичными
    /// методами этого класса.
    /// </summary>
    [ComVisible(true)]
    [Guid("FE63C2BE-14F3-45D9-BE38-98A925BDF989")]
    [ProgId("Abbreviator.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private AddInController _controller;
        private dynamic _ribbon;

        // ==================================================================
        // IDTExtensibility2
        // ==================================================================

        public void OnConnection(object application, ext_ConnectMode connectMode,
                                 object addInInst, ref Array custom)
        {
            try
            {
                _controller = new AddInController(application);
            }
            catch (Exception ex)
            {
                Report("Не удалось загрузить надстройку", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                if (_controller != null) _controller.Shutdown();
            }
            catch { }
            finally
            {
                _controller = null;
                _ribbon = null;
            }
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom) { }

        // ==================================================================
        // IRibbonExtensibility
        // ==================================================================

        public string GetCustomUI(string ribbonID)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                string name = assembly.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith("Ribbon.xml", StringComparison.OrdinalIgnoreCase));
                if (name == null) return string.Empty;

                using (var stream = assembly.GetManifestResourceStream(name))
                {
                    if (stream == null) return string.Empty;
                    using (var reader = new StreamReader(stream))
                        return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                Report("Не удалось загрузить разметку ленты", ex);
                return string.Empty;
            }
        }

        public void OnRibbonLoad(object ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        /// <summary>Заставить Office перечитать getLabel/getVisible.</summary>
        public void InvalidateRibbon()
        {
            try { if (_ribbon != null) _ribbon.Invalidate(); }
            catch { }
        }

        // ==================================================================
        // Кнопки ленты
        // ==================================================================

        public void OnCheck(object control) { Run(() => _controller.CheckDocument()); }

        public void OnClear(object control) { Run(() => _controller.ClearHighlights()); }

        public void OnClearAll(object control) { Run(() => _controller.ClearAllHighlights()); }

        public void OnPages(object control) { Run(() => _controller.ShowMainForm(0)); }

        public void OnList(object control) { Run(() => _controller.ShowMainForm(1)); }

        public void OnGotoDict(object control) { Run(() => _controller.GoToDictionary()); }

        public void OnIgnoreList(object control) { Run(() => _controller.ShowIgnoreForm()); }

        public void OnSettings(object control) { Run(() => _controller.ShowSettingsForm()); }

        public void OnSettingsFile(object control) { Run(() => _controller.OpenSettingsFile()); }

        public void OnAbout(object control) { Run(() => _controller.ShowAbout()); }

        public void OnDetect(object control)
        {
            Run(() =>
            {
                dynamic doc = _controller.ActiveDocument;
                if (doc == null)
                {
                    _controller.Warn("Откройте документ Word.");
                    return;
                }

                DocumentState state = _controller.GetState(doc);
                AddInController.PageDetection detection = _controller.AutoDetectPages(doc, state);

                if (!detection.Found)
                {
                    _controller.Warn("Заголовок перечня принятых сокращений не найден.\r\n\r\n" +
                                     "Задайте листы вручную в окне «Листы перечня».");
                    _controller.ShowMainForm(0);
                    return;
                }

                foreach (var page in detection.Pages) state.AddPage(page);
                _controller.SaveState(doc, state);
                _controller.ShowMainForm(0);
                _controller.Info("Перечень найден на странице " + detection.HeaderPage + ".\r\n" +
                                 "Листы перечня: " + string.Join(", ", detection.Pages) + ".");
            });
        }

        // ==================================================================
        // Контекстное меню
        // ==================================================================

        private Target CurrentTarget
        {
            get
            {
                try { return _controller == null ? new Target() : _controller.ResolveTarget(); }
                catch { return new Target(); }
            }
        }

        public bool GetAddVisible(object control)
        {
            var t = CurrentTarget;
            return t.HasValue && t.Status != AbbrStatus.Known;
        }

        public string GetAddLabel(object control)
        {
            var t = CurrentTarget;
            return "Abbreviator: добавить «" + (t.Abbr ?? string.Empty) + "» в перечень…";
        }

        public bool GetIgnoreVisible(object control)
        {
            var t = CurrentTarget;
            return t.HasValue && t.Status != AbbrStatus.Ignored;
        }

        public string GetIgnoreLabel(object control)
        {
            var t = CurrentTarget;
            string suffix = t.UsageCount == 1 ? " (используется 1 раз)"
                          : t.UsageCount > 1 ? " (используется " + t.UsageCount + " раз)"
                          : string.Empty;
            return "Abbreviator: игнорировать «" + (t.Abbr ?? string.Empty) + "»" + suffix;
        }

        public bool GetUnignoreVisible(object control)
        {
            var t = CurrentTarget;
            return t.HasValue && t.Status == AbbrStatus.Ignored;
        }

        public string GetUnignoreLabel(object control)
        {
            var t = CurrentTarget;
            return "Abbreviator: не игнорировать «" + (t.Abbr ?? string.Empty) + "»";
        }

        public bool GetGotoVisible(object control)
        {
            var t = CurrentTarget;
            return t.HasValue && t.Status == AbbrStatus.Known && t.EntryStart >= 0;
        }

        public string GetGotoLabel(object control)
        {
            var t = CurrentTarget;
            return "Abbreviator: показать «" + (t.Abbr ?? string.Empty) + "» в перечне";
        }

        public void OnCtxAdd(object control) { Run(() => _controller.AddTargetToDictionary()); }

        public void OnCtxIgnore(object control) { Run(() => _controller.IgnoreTarget()); }

        public void OnCtxUnignore(object control)
        {
            Run(() =>
            {
                var t = _controller.ResolveTarget();
                if (t.HasValue) _controller.UnignoreAbbreviation(t.Abbr);
                _controller.CheckDocument(true);
            });
        }

        public void OnCtxGoto(object control) { Run(() => _controller.GoToTargetEntry()); }

        // ==================================================================

        private void Run(Action action)
        {
            if (_controller == null) return;
            try
            {
                action();
            }
            catch (Exception ex)
            {
                Report("Ошибка Abbreviator", ex);
            }
        }

        private static void Report(string title, Exception ex)
        {
            try
            {
                MessageBox.Show(title + ":\r\n\r\n" + ex.Message, "Abbreviator",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch { }
        }
    }
}
