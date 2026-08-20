using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;

namespace Abbreviator.Core
{
    /// <summary>
    /// Журнал загрузки и работы надстройки.
    ///
    /// Word сообщает об ошибках надстройки одной общей фразой и не показывает
    /// исключение, поэтому каждая точка входа из COM пишет сюда. По журналу
    /// видно, докуда дошла загрузка: создан ли объект, вызван ли OnConnection,
    /// запросил ли Word разметку ленты.
    ///
    /// Ни один метод не бросает исключений: сбой журнала не должен ронять
    /// надстройку.
    /// </summary>
    public static class Diag
    {
        private const long MaxBytes = 1024 * 1024;

        private static readonly object Sync = new object();

        public static string Folder
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Abbreviator");
            }
        }

        public static string LogPath
        {
            get { return Path.Combine(Folder, "abbreviator.log"); }
        }

        public static void Write(string message)
        {
            try
            {
                lock (Sync)
                {
                    Directory.CreateDirectory(Folder);
                    Truncate();

                    string line = string.Format(CultureInfo.InvariantCulture,
                        "{0:yyyy-MM-dd HH:mm:ss.fff}  {1}{2}",
                        DateTime.Now, message, Environment.NewLine);

                    File.AppendAllText(LogPath, line, new UTF8Encoding(true));
                }
            }
            catch
            {
                // Журнал не критичен.
            }
        }

        public static void Error(string context, Exception ex)
        {
            if (ex == null)
            {
                Write("ОШИБКА: " + context);
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("ОШИБКА: " + context);

            for (Exception e = ex; e != null; e = e.InnerException)
            {
                sb.AppendLine("    " + e.GetType().FullName + ": " + e.Message);
                if (!string.IsNullOrEmpty(e.StackTrace)) sb.AppendLine(e.StackTrace);
            }

            Write(sb.ToString().TrimEnd());
        }

        /// <summary>Шапка журнала: версии, пути, разрядность процесса.</summary>
        public static void WriteHeader(dynamic application)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();

                var sb = new StringBuilder();
                sb.AppendLine("=== Abbreviator ===");
                sb.AppendLine("    сборка:      " + assembly.GetName().Version);
                sb.AppendLine("    путь:        " + assembly.Location);
                sb.AppendLine("    CLR:         " + Environment.Version);
                sb.AppendLine("    процесс:     " + (IntPtr.Size == 8 ? "64-разрядный" : "32-разрядный"));

                string version = "?", build = "?";
                try { version = Convert.ToString(application.Version); } catch { }
                try { build = Convert.ToString(application.Build); } catch { }
                sb.Append("    Word:        " + version + " (" + build + ")");

                Write(sb.ToString());
            }
            catch (Exception ex)
            {
                Error("не удалось собрать сведения об окружении", ex);
            }
        }
    }
}
