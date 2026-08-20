using System;
using System.Collections.Generic;

namespace Abbreviator.Core
{
    /// <summary>Статус аббревиатуры относительно перечня принятых сокращений.</summary>
    public enum AbbrStatus
    {
        /// <summary>Есть в перечне — подсвечивается зелёным.</summary>
        Known = 0,

        /// <summary>Нет в перечне — подсвечивается красным.</summary>
        Unknown = 1,

        /// <summary>Помечена пользователем как игнорируемая — не подсвечивается.</summary>
        Ignored = 2
    }

    /// <summary>Одно вхождение аббревиатуры в текст документа.</summary>
    public sealed class Occurrence
    {
        /// <summary>Нормализованная аббревиатура (без падежного окончания), напр. "АСУ".</summary>
        public string Abbr;

        /// <summary>Как встретилось в тексте, напр. "АСУшник" → "АСУ" + окончание.</summary>
        public string Raw;

        /// <summary>Позиция начала в координатах Document.Range.</summary>
        public int Start;

        /// <summary>Позиция конца в координатах Document.Range.</summary>
        public int End;

        /// <summary>Номер страницы (0 — не вычислялся).</summary>
        public int Page;

        /// <summary>Вхождение находится на странице перечня сокращений.</summary>
        public bool InDictionary;

        public AbbrStatus Status;

        public override string ToString()
        {
            return Abbr + " @" + Start + ".." + End;
        }
    }

    /// <summary>Агрегированная информация по одной аббревиатуре.</summary>
    public sealed class AbbrInfo
    {
        public string Abbr;
        public AbbrStatus Status;

        /// <summary>Число вхождений в основном тексте (без учёта самого перечня).</summary>
        public int UsageCount;

        /// <summary>Число вхождений на страницах перечня.</summary>
        public int DictionaryCount;

        /// <summary>Расшифровка из перечня, если она найдена.</summary>
        public string Definition;

        /// <summary>Позиция первого вхождения в основном тексте.</summary>
        public int FirstUse = -1;

        /// <summary>Позиция статьи в перечне.</summary>
        public int EntryStart = -1;

        public bool IsUnused
        {
            get { return UsageCount == 0; }
        }
    }

    /// <summary>Статья перечня принятых сокращений: "АСУ - автоматизированная система управления".</summary>
    public sealed class DictEntry
    {
        public string Term;
        public string Definition;
        public int Page;
        public int Start;
        public int End;

        /// <summary>Статья оформлена строкой таблицы.</summary>
        public bool IsTableRow;

        /// <summary>Индекс строки в таблице (1-based), если IsTableRow.</summary>
        public int RowIndex;

        /// <summary>Все аббревиатуры, найденные в левой части статьи.</summary>
        public List<string> Terms = new List<string>();

        public override string ToString()
        {
            return Term + " - " + Definition;
        }
    }

    /// <summary>Результат полного разбора документа.</summary>
    public sealed class ScanResult
    {
        public List<Occurrence> Occurrences = new List<Occurrence>();
        public List<DictEntry> Entries = new List<DictEntry>();
        public List<int> DictionaryPages = new List<int>();

        /// <summary>Страницы содержания — пропускаются при разборе.</summary>
        public List<int> TocPages = new List<int>();
        public Dictionary<string, AbbrInfo> ByAbbr =
            new Dictionary<string, AbbrInfo>(StringComparer.Ordinal);

        public int TotalPages;
        public DateTime ScannedAt = DateTime.Now;

        /// <summary>Заголовок перечня найден и его позиция в документе.</summary>
        public int HeaderStart = -1;
        public int HeaderPage;

        public int KnownCount;
        public int UnknownCount;
        public int IgnoredCount;
    }

    /// <summary>Аббревиатура под курсором — то, с чем работает контекстное меню.</summary>
    public sealed class Target
    {
        public string Abbr;
        public int Start;
        public int End;
        public AbbrStatus Status;
        public string Definition;
        public int UsageCount;
        public int EntryStart = -1;

        public bool HasValue
        {
            get { return !string.IsNullOrEmpty(Abbr); }
        }
    }
}
