using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Abbreviator.Core
{
    /// <summary>Кандидат в аббревиатуры, найденный в куске текста.</summary>
    public struct TokenMatch
    {
        /// <summary>Нормализованная аббревиатура (без падежного окончания).</summary>
        public string Abbr;

        /// <summary>Полный текст вхождения вместе с окончанием ("ГОСТы").</summary>
        public string Raw;

        /// <summary>Смещение Abbr внутри переданной строки.</summary>
        public int Index;

        /// <summary>Длина Abbr.</summary>
        public int Length;
    }

    /// <summary>
    /// Поиск аббревиатур в тексте.
    ///
    /// Аббревиатурой считается последовательность из двух и более прописных букв
    /// (кириллица или латиница), возможно с цифрами и дефисами: АСУ, ГОСТ, ЭВМ,
    /// САПР, API, SQL-92, ТУ-154. Допускается короткое строчное окончание —
    /// "ГОСТы", "АСУшник" — оно отбрасывается при нормализации.
    /// </summary>
    public sealed class AbbreviationRecognizer
    {
        /// <summary>
        /// abbr — само сокращение, tail — падежное окончание строчными буквами.
        /// Границы слева и справа: не буква, не цифра, не подчёркивание.
        /// </summary>
        private const string Pattern =
            @"(?<![\p{L}\p{Nd}_])" +
            @"(?<abbr>[A-ZА-ЯЁ][A-ZА-ЯЁ0-9]*(?:-[A-ZА-ЯЁ0-9]+)*)" +
            @"(?<tail>[а-яёa-z]{0,3})" +
            @"(?![\p{L}\p{Nd}_])";

        private static readonly Regex Rx =
            new Regex(Pattern, RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex RomanRx =
            new Regex(@"^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$",
                      RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private readonly AppSettings _settings;
        private readonly List<Regex> _ignorePatterns = new List<Regex>();

        /// <summary>
        /// Внешняя проверка «это обычное слово, а не аббревиатура».
        /// Подставляется орфографией Word; может быть null.
        /// </summary>
        public Func<string, bool> IsCommonWord;

        private readonly Dictionary<string, bool> _commonWordCache =
            new Dictionary<string, bool>(StringComparer.Ordinal);

        public AbbreviationRecognizer(AppSettings settings)
        {
            _settings = settings ?? new AppSettings();

            foreach (var pattern in _settings.IgnorePatterns)
            {
                if (string.IsNullOrWhiteSpace(pattern)) continue;
                try
                {
                    _ignorePatterns.Add(new Regex(pattern, RegexOptions.CultureInvariant));
                }
                catch (ArgumentException)
                {
                    // Ошибка в пользовательском выражении не должна ломать разбор.
                }
            }
        }

        /// <summary>Все аббревиатуры в строке.</summary>
        public IEnumerable<TokenMatch> Find(string text)
        {
            if (string.IsNullOrEmpty(text)) yield break;

            var excluded = BuildExcludedSpans(text);

            foreach (Match m in Rx.Matches(text))
            {
                var g = m.Groups["abbr"];
                string abbr = g.Value;

                if (InExcludedSpan(excluded, g.Index)) continue;
                if (!IsAbbreviation(abbr)) continue;

                yield return new TokenMatch
                {
                    Abbr = abbr,
                    Raw = m.Value,
                    Index = g.Index,
                    Length = g.Length
                };
            }
        }

        /// <summary>
        /// Участки текста, внутри которых аббревиатуры не ищутся: обозначения
        /// конструкторской документации вида ИЯУК.123456.789-01 и подобные.
        /// Отсеиваются именно диапазоном, а не по одному токену, поэтому
        /// производные формы обозначения тоже игнорируются целиком.
        /// </summary>
        private List<int[]> BuildExcludedSpans(string text)
        {
            var spans = new List<int[]>();
            foreach (var rx in _ignorePatterns)
            {
                foreach (Match m in rx.Matches(text))
                    if (m.Length > 0) spans.Add(new[] { m.Index, m.Index + m.Length });
            }
            return spans;
        }

        private static bool InExcludedSpan(List<int[]> spans, int index)
        {
            for (int i = 0; i < spans.Count; i++)
                if (index >= spans[i][0] && index < spans[i][1]) return true;
            return false;
        }

        /// <summary>Проверка одиночного токена (для контекстного меню).</summary>
        public bool IsAbbreviation(string token)
        {
            if (string.IsNullOrEmpty(token)) return false;
            if (token.Length > _settings.MaxLength) return false;

            int letters = 0;
            foreach (char c in token)
            {
                if (char.IsLetter(c))
                {
                    if (!char.IsUpper(c)) return false;
                    letters++;
                }
                else if (!char.IsDigit(c) && c != '-')
                {
                    return false;
                }
            }
            if (letters < _settings.MinLetters) return false;

            if (_settings.StopWords.Contains(token)) return false;
            if (_settings.SkipRomanNumerals && IsRoman(token)) return false;
            if (LooksLikeCommonWord(token)) return false;

            return true;
        }

        /// <summary>Нормализация: обрезаем падежное окончание, если оно есть.</summary>
        public string Normalize(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return raw;
            int end = raw.Length;
            while (end > 0 && char.IsLower(raw[end - 1])) end--;
            return end == 0 ? raw : raw.Substring(0, end);
        }

        private static bool IsRoman(string token)
        {
            foreach (char c in token)
                if ("IVXLCDM".IndexOf(c) < 0) return false;
            return RomanRx.IsMatch(token);
        }

        private bool LooksLikeCommonWord(string token)
        {
            if (!_settings.UseSpellCheckFilter) return false;
            if (IsCommonWord == null) return false;
            if (token.Length < _settings.SpellCheckMinLength) return false;
            if (token.IndexOf('-') >= 0) return false;

            foreach (char c in token)
                if (!char.IsLetter(c)) return false;

            bool cached;
            if (_commonWordCache.TryGetValue(token, out cached)) return cached;

            bool result;
            try
            {
                // Приводим к виду "Введение": слово из словаря останется корректным,
                // а настоящая аббревиатура ("Асу", "Гост") — нет.
                string probe = token.Substring(0, 1) + token.Substring(1).ToLowerInvariant();
                result = IsCommonWord(probe);
            }
            catch
            {
                result = false;
            }

            _commonWordCache[token] = result;
            return result;
        }

        /// <summary>
        /// Абзац похож на заголовок, набранный капслоком: два и более слова,
        /// все буквы прописные. Такие абзацы пропускаем целиком.
        /// </summary>
        public bool IsUpperCaseHeading(string paragraphText)
        {
            if (!_settings.SkipUpperCaseParagraphs) return false;
            if (string.IsNullOrWhiteSpace(paragraphText)) return false;

            string t = paragraphText.Trim();
            if (t.Length < 4) return false;

            int upper = 0, lower = 0, words = 1;
            bool prevSpace = false;
            foreach (char c in t)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (!prevSpace) words++;
                    prevSpace = true;
                    continue;
                }
                prevSpace = false;
                if (!char.IsLetter(c)) continue;
                if (char.IsUpper(c)) upper++; else lower++;
            }

            if (words < 2) return false;
            if (upper + lower == 0) return false;
            return lower * 10 <= upper; // не более 10% строчных
        }

        public void ClearCache()
        {
            _commonWordCache.Clear();
        }
    }
}
