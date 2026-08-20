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
    /// Класс-интерфейс — AutoDispatch, а не AutoDual. Обратные вызовы ленты и
    /// контекстного меню Office ищет по имени через IDispatch::GetIDsOfNames,
    /// а его AutoDispatch обслуживает рефлексией по публичным методам класса —
    /// именно так устроены классы лент в VSTO.
    ///
    /// AutoDual вдобавок требует сгенерировать для класса ITypeInfo, то есть
    /// экспортировать все публичные члены в библиотеку типов. Если экспорт не
    /// удаётся, CCW не создаётся: объект уже сконструирован, но QueryInterface
    /// не отвечает, и Word сообщает «произошла ошибка с надстройкой», не вызвав
    /// ни одного метода. AutoDispatch этот шаг не выполняет вовсе.
    ///
    /// Все обратные вызовы обязаны быть публичными методами этого класса.
    /// </summary>
    [ComVisible(true)]
    [Guid("FE63C2BE-14F3-45D9-BE38-98A925BDF989")]
    [ProgId("Abbreviator.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private AddInController _controller;
        private dynamic _ribbon;

        public Connect()
        {
            // Первая точка, куда попадает управление: если в журнале нет даже
            // этой строки, COM-объект вообще не был создан — дело в регистрации.
            Diag.Write("Connect: объект создан");
        }

        // ==================================================================
        // IDTExtensibility2
        // ==================================================================

        public void OnConnection(object application, ext_ConnectMode connectMode,
                                 object addInInst, ref Array custom)
        {
            try
            {
                Diag.Write("OnConnection: режим " + connectMode);
                Diag.WriteHeader(application);

                _controller = new AddInController(application);

                Diag.Write("OnConnection: надстройка загружена");
            }
            catch (Exception ex)
            {
                Diag.Error("OnConnection", ex);
                Report("Не удалось загрузить надстройку", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Diag.Write("OnDisconnection: режим " + removeMode);
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
                Diag.Write("GetCustomUI: запрошена разметка для " + ribbonID);

                var assembly = Assembly.GetExecutingAssembly();
                string name = assembly.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith("Ribbon.xml", StringComparison.OrdinalIgnoreCase));

                if (name == null)
                {
                    Diag.Error("GetCustomUI: ресурс Ribbon.xml не найден в сборке", null);
                    return string.Empty;
                }

                using (var stream = assembly.GetManifestResourceStream(name))
                {
                    if (stream == null)
                    {
                        Diag.Error("GetCustomUI: поток ресурса " + name + " пуст", null);
                        return string.Empty;
                    }

                    using (var reader = new StreamReader(stream))
                    {
                        string xml = reader.ReadToEnd();
                        Diag.Write("GetCustomUI: отдано " + xml.Length + " символов");
                        return xml;
                    }
                }
            }
            catch (Exception ex)
            {
                Diag.Error("GetCustomUI", ex);
                Report("Не удалось загрузить разметку ленты", ex);
                return string.Empty;
            }
        }

        public void OnRibbonLoad(object ribbonUI)
        {
            Diag.Write("OnRibbonLoad: лента построена");
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

        public bool GetLivePressed(object control)
        {
            return Run(() => _controller.LiveEnabled, false);
        }

        public void OnLiveToggle(object control, bool pressed)
        {
            Run(() => _controller.SetLiveEnabled(pressed));
        }

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
            return Run(() =>
            {
                var t = CurrentTarget;
                return t.HasValue && t.Status != AbbrStatus.Known;
            }, false);
        }

        public string GetAddLabel(object control)
        {
            return Run(() => "Abbreviator: добавить «" + (CurrentTarget.Abbr ?? string.Empty) +
                             "» в перечень…", "Abbreviator: добавить в перечень…");
        }

        public bool GetIgnoreVisible(object control)
        {
            return Run(() =>
            {
                var t = CurrentTarget;
                return t.HasValue && t.Status != AbbrStatus.Ignored;
            }, false);
        }

        public string GetIgnoreLabel(object control)
        {
            return Run(() =>
            {
                var t = CurrentTarget;
                string suffix = t.UsageCount == 1 ? " (используется 1 раз)"
                              : t.UsageCount > 1 ? " (используется " + t.UsageCount + " раз)"
                              : string.Empty;
                return "Abbreviator: игнорировать «" + (t.Abbr ?? string.Empty) + "»" + suffix;
            }, "Abbreviator: игнорировать");
        }

        public bool GetUnignoreVisible(object control)
        {
            return Run(() =>
            {
                var t = CurrentTarget;
                return t.HasValue && t.Status == AbbrStatus.Ignored;
            }, false);
        }

        public string GetUnignoreLabel(object control)
        {
            return Run(() => "Abbreviator: не игнорировать «" + (CurrentTarget.Abbr ?? string.Empty) + "»",
                       "Abbreviator: не игнорировать");
        }

        public bool GetGotoVisible(object control)
        {
            return Run(() =>
            {
                var t = CurrentTarget;
                return t.HasValue && t.Status == AbbrStatus.Known && t.EntryStart >= 0;
            }, false);
        }

        public string GetGotoLabel(object control)
        {
            return Run(() => "Abbreviator: показать «" + (CurrentTarget.Abbr ?? string.Empty) + "» в перечне",
                       "Abbreviator: показать в перечне");
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
            if (_controller == null)
            {
                Diag.Error("команда вызвана, но надстройка не инициализирована", null);
                return;
            }

            try
            {
                action();
            }
            catch (Exception ex)
            {
                Diag.Error("выполнение команды", ex);
                Report("Ошибка Abbreviator", ex);
            }
        }

        /// <summary>
        /// Обёртка для обратных вызовов, возвращающих значение: исключение,
        /// улетевшее в Word, приводит к отключению надстройки.
        /// </summary>
        private T Run<T>(Func<T> action, T fallback)
        {
            if (_controller == null) return fallback;

            try
            {
                return action();
            }
            catch (Exception ex)
            {
                Diag.Error("обратный вызов контекстного меню", ex);
                return fallback;
            }
        }

        private static void Report(string title, Exception ex)
        {
            try
            {
                MessageBox.Show(title + ":\r\n\r\n" + ex.Message +
                                "\r\n\r\nПодробности: " + Diag.LogPath,
                                "Abbreviator", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch { }
        }
    }
}
