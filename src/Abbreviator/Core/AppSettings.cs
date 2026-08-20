using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Abbreviator.Interop;

namespace Abbreviator.Core
{
    /// <summary>
    /// Глобальные настройки надстройки.
    /// Хранятся в %APPDATA%\Abbreviator\settings.ini — обычный текстовый файл,
    /// его можно править руками.
    /// </summary>
    public sealed class AppSettings
    {
        // --- распознавание ---
        public int MinLetters = 2;
        public int MaxLength = 15;
        public bool SkipRomanNumerals = true;
        public bool SkipUpperCaseParagraphs = true;
        public bool UseSpellCheckFilter = true;
        public int SpellCheckMinLength = 5;

        // --- подсветка ---
        public int KnownColor = Wd.BrightGreen;
        public int UnknownColor = Wd.Red;
        public bool HighlightDictionarySection = false;
        public bool HighlightIgnored = false;

        // --- перечень ---
        public string EntrySeparator = " - ";
        public bool KeepAlphabeticalOrder = false;
        public bool AutoDetectPagesOnScan = true;

        /// <summary>Варианты заголовка перечня. Сравнение — по нормализованной строке.</summary>
        public List<string> HeaderVariants = new List<string>
        {
            "Перечень принятых сокращений",
            "Перечень сокращений",
            "Список принятых сокращений",
            "Список сокращений",
            "Перечень условных обозначений и сокращений",
            "Перечень условных обозначений, символов и сокращений",
            "Принятые сокращения",
            "Условные обозначения и сокращения",
            "Список используемых сокращений"
        };

        /// <summary>Слова, которые никогда не считаются аббревиатурами.</summary>
        public HashSet<string> StopWords = new HashSet<string>(StringComparer.Ordinal)
        {
            "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ", "ПРИЛОЖЕНИЕ", "ПРИЛОЖЕНИЯ",
            "РЕФЕРАТ", "АННОТАЦИЯ", "ГЛАВА", "РАЗДЕЛ", "ЧАСТЬ", "ПЕРЕЧЕНЬ", "СПИСОК",
            "ЛИТЕРАТУРА", "ИСТОЧНИКИ", "ПРИНЯТЫХ", "СОКРАЩЕНИЙ", "СОКРАЩЕНИЯ",
            "ОБОЗНАЧЕНИЙ", "ОБОЗНАЧЕНИЯ", "УСЛОВНЫХ", "ТЕРМИНОВ", "ТЕРМИНЫ",
            "ТАБЛИЦА", "РИСУНОК", "ФОРМУЛА", "ПРИМЕЧАНИЕ", "ПРИМЕР",
            "ЗАДАНИЕ", "ЦЕЛЬ", "ЗАДАЧИ", "ВЫВОДЫ", "ИТОГО", "ВСЕГО",
            "УТВЕРЖДАЮ", "СОГЛАСОВАНО", "ДА", "НЕТ"
        };

        /// <summary>Глобальный (общий для всех документов) список игнорируемых сокращений.</summary>
        public HashSet<string> GlobalIgnored = new HashSet<string>(StringComparer.Ordinal);

        // ------------------------------------------------------------------
        // Загрузка / сохранение
        // ------------------------------------------------------------------

        public static string Folder
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Abbreviator");
            }
        }

        public static string FilePath
        {
            get { return Path.Combine(Folder, "settings.ini"); }
        }

        private static AppSettings _current;

        public static AppSettings Current
        {
            get
            {
                if (_current == null) _current = Load();
                return _current;
            }
        }

        public static AppSettings Load()
        {
            var s = new AppSettings();
            try
            {
                if (!File.Exists(FilePath)) return s;
                var map = ReadIni(FilePath);

                s.MinLetters = GetInt(map, "MinLetters", s.MinLetters);
                s.MaxLength = GetInt(map, "MaxLength", s.MaxLength);
                s.SkipRomanNumerals = GetBool(map, "SkipRomanNumerals", s.SkipRomanNumerals);
                s.SkipUpperCaseParagraphs = GetBool(map, "SkipUpperCaseParagraphs", s.SkipUpperCaseParagraphs);
                s.UseSpellCheckFilter = GetBool(map, "UseSpellCheckFilter", s.UseSpellCheckFilter);
                s.SpellCheckMinLength = GetInt(map, "SpellCheckMinLength", s.SpellCheckMinLength);

                s.KnownColor = GetInt(map, "KnownColor", s.KnownColor);
                s.UnknownColor = GetInt(map, "UnknownColor", s.UnknownColor);
                s.HighlightDictionarySection = GetBool(map, "HighlightDictionarySection", s.HighlightDictionarySection);
                s.HighlightIgnored = GetBool(map, "HighlightIgnored", s.HighlightIgnored);

                s.EntrySeparator = GetStr(map, "EntrySeparator", s.EntrySeparator);
                s.KeepAlphabeticalOrder = GetBool(map, "KeepAlphabeticalOrder", s.KeepAlphabeticalOrder);
                s.AutoDetectPagesOnScan = GetBool(map, "AutoDetectPagesOnScan", s.AutoDetectPagesOnScan);

                var headers = GetStr(map, "HeaderVariants", null);
                if (!string.IsNullOrWhiteSpace(headers))
                    s.HeaderVariants = SplitList(headers);

                var stop = GetStr(map, "StopWords", null);
                if (!string.IsNullOrWhiteSpace(stop))
                    s.StopWords = new HashSet<string>(SplitList(stop), StringComparer.Ordinal);

                var ignored = GetStr(map, "GlobalIgnored", null);
                if (!string.IsNullOrWhiteSpace(ignored))
                    s.GlobalIgnored = new HashSet<string>(SplitList(ignored), StringComparer.Ordinal);
            }
            catch
            {
                // Битый файл настроек не должен ломать надстройку.
            }
            return s;
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Folder);
                var sb = new StringBuilder();
                sb.AppendLine("; Настройки надстройки Abbreviator");
                sb.AppendLine("; Списки разделяются символом |");
                sb.AppendLine();
                sb.AppendLine("[Recognition]");
                sb.AppendLine("MinLetters=" + MinLetters);
                sb.AppendLine("MaxLength=" + MaxLength);
                sb.AppendLine("SkipRomanNumerals=" + Fmt(SkipRomanNumerals));
                sb.AppendLine("SkipUpperCaseParagraphs=" + Fmt(SkipUpperCaseParagraphs));
                sb.AppendLine("UseSpellCheckFilter=" + Fmt(UseSpellCheckFilter));
                sb.AppendLine("SpellCheckMinLength=" + SpellCheckMinLength);
                sb.AppendLine();
                sb.AppendLine("[Highlight]");
                sb.AppendLine("KnownColor=" + KnownColor);
                sb.AppendLine("UnknownColor=" + UnknownColor);
                sb.AppendLine("HighlightDictionarySection=" + Fmt(HighlightDictionarySection));
                sb.AppendLine("HighlightIgnored=" + Fmt(HighlightIgnored));
                sb.AppendLine();
                sb.AppendLine("[Dictionary]");
                sb.AppendLine("EntrySeparator=" + EntrySeparator);
                sb.AppendLine("KeepAlphabeticalOrder=" + Fmt(KeepAlphabeticalOrder));
                sb.AppendLine("AutoDetectPagesOnScan=" + Fmt(AutoDetectPagesOnScan));
                sb.AppendLine("HeaderVariants=" + string.Join("|", HeaderVariants));
                sb.AppendLine();
                sb.AppendLine("[Lists]");
                sb.AppendLine("StopWords=" + string.Join("|", StopWords.OrderBy(x => x, StringComparer.Ordinal)));
                sb.AppendLine("GlobalIgnored=" + string.Join("|", GlobalIgnored.OrderBy(x => x, StringComparer.Ordinal)));

                File.WriteAllText(FilePath, sb.ToString(), new UTF8Encoding(true));
            }
            catch
            {
                // Нет прав на запись — работаем с настройками только в памяти.
            }
        }

        public static void Reload()
        {
            _current = Load();
        }

        // ------------------------------------------------------------------

        private static List<string> SplitList(string value)
        {
            return value.Split('|')
                        .Select(x => x.Trim())
                        .Where(x => x.Length > 0)
                        .ToList();
        }

        private static string Fmt(bool value)
        {
            return value ? "1" : "0";
        }

        private static Dictionary<string, string> ReadIni(string path)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
            {
                var line = raw.Trim();
                if (line.Length == 0) continue;
                if (line[0] == ';' || line[0] == '#' || line[0] == '[') continue;
                int eq = line.IndexOf('=');
                if (eq <= 0) continue;
                map[line.Substring(0, eq).Trim()] = line.Substring(eq + 1);
            }
            return map;
        }

        private static string GetStr(Dictionary<string, string> map, string key, string fallback)
        {
            string v;
            return map.TryGetValue(key, out v) ? v : fallback;
        }

        private static int GetInt(Dictionary<string, string> map, string key, int fallback)
        {
            string v;
            int r;
            if (map.TryGetValue(key, out v) &&
                int.TryParse(v.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out r))
                return r;
            return fallback;
        }

        private static bool GetBool(Dictionary<string, string> map, string key, bool fallback)
        {
            string v;
            if (!map.TryGetValue(key, out v)) return fallback;
            v = v.Trim();
            return v == "1" || v.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                   v.Equals("да", StringComparison.OrdinalIgnoreCase);
        }
    }
}
