using System;
using System.Collections.Generic;
using System.Linq;

namespace Abbreviator.Core
{
    /// <summary>
    /// Состояние надстройки, привязанное к конкретному документу:
    /// список страниц перечня и список игнорируемых сокращений.
    ///
    /// Хранится в Document.Variables — переменные документа переживают
    /// сохранение файла и не видны в тексте.
    /// </summary>
    public sealed class DocumentState
    {
        public const string VarPages = "AbbreviatorPages";
        public const string VarIgnored = "AbbreviatorIgnored";
        public const string VarVersion = "AbbreviatorVersion";

        /// <summary>Номера страниц, с которых берётся перечень принятых сокращений.</summary>
        public List<int> DictionaryPages = new List<int>();

        /// <summary>Сокращения, помеченные пользователем как игнорируемые.</summary>
        public HashSet<string> Ignored = new HashSet<string>(StringComparer.Ordinal);

        /// <summary>Последний разбор документа (в памяти, не сохраняется).</summary>
        public ScanResult LastScan;

        // ------------------------------------------------------------------

        public static DocumentState Load(dynamic doc)
        {
            var st = new DocumentState();
            if (doc == null) return st;

            st.DictionaryPages = ParsePages(GetVar(doc, VarPages));
            var ignored = GetVar(doc, VarIgnored);
            if (!string.IsNullOrWhiteSpace(ignored))
            {
                foreach (var item in ignored.Split('|', ';', ','))
                {
                    var t = item.Trim();
                    if (t.Length > 0) st.Ignored.Add(t);
                }
            }
            return st;
        }

        public void Save(dynamic doc)
        {
            if (doc == null) return;
            SetVar(doc, VarPages, string.Join(",", DictionaryPages.OrderBy(p => p)));
            SetVar(doc, VarIgnored, string.Join("|", Ignored.OrderBy(x => x, StringComparer.Ordinal)));
            SetVar(doc, VarVersion, "1");
        }

        public void SetPages(IEnumerable<int> pages)
        {
            DictionaryPages = (pages ?? Enumerable.Empty<int>())
                .Where(p => p > 0)
                .Distinct()
                .OrderBy(p => p)
                .ToList();
        }

        public void AddPage(int page)
        {
            if (page <= 0) return;
            if (!DictionaryPages.Contains(page))
            {
                DictionaryPages.Add(page);
                DictionaryPages.Sort();
            }
        }

        public void RemovePage(int page)
        {
            DictionaryPages.Remove(page);
        }

        public bool IsIgnored(string abbr)
        {
            return Ignored.Contains(abbr) || AppSettings.Current.GlobalIgnored.Contains(abbr);
        }

        public int LastDictionaryPage
        {
            get { return DictionaryPages.Count == 0 ? 0 : DictionaryPages.Max(); }
        }

        public int FirstDictionaryPage
        {
            get { return DictionaryPages.Count == 0 ? 0 : DictionaryPages.Min(); }
        }

        // ------------------------------------------------------------------

        private static List<int> ParsePages(string raw)
        {
            var list = new List<int>();
            if (string.IsNullOrWhiteSpace(raw)) return list;

            foreach (var chunk in raw.Split(',', ';', ' '))
            {
                var t = chunk.Trim();
                if (t.Length == 0) continue;

                // Поддерживаем диапазоны вида "12-15".
                int dash = t.IndexOf('-');
                if (dash > 0 && dash < t.Length - 1)
                {
                    int a, b;
                    if (int.TryParse(t.Substring(0, dash), out a) &&
                        int.TryParse(t.Substring(dash + 1), out b) && a > 0 && b >= a)
                    {
                        for (int p = a; p <= b; p++) list.Add(p);
                        continue;
                    }
                }

                int one;
                if (int.TryParse(t, out one) && one > 0) list.Add(one);
            }

            return list.Distinct().OrderBy(p => p).ToList();
        }

        private static string GetVar(dynamic doc, string name)
        {
            try
            {
                dynamic vars = doc.Variables;
                int count = (int)vars.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic v = vars[i];
                    if (string.Equals((string)v.Name, name, StringComparison.OrdinalIgnoreCase))
                        return v.Value as string;
                }
            }
            catch { }
            return null;
        }

        private static void SetVar(dynamic doc, string name, string value)
        {
            try
            {
                doc.Variables[name].Value = value ?? string.Empty;
            }
            catch
            {
                try { doc.Variables.Add(name, value ?? string.Empty); } catch { }
            }
        }
    }
}
