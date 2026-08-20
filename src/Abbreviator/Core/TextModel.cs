using System;
using System.Collections.Generic;

namespace Abbreviator.Core
{
    /// <summary>Абзац документа: границы в координатах Range и очищенный текст.</summary>
    public sealed class ParaSpan
    {
        /// <summary>Позиция первого символа абзаца в координатах Document.Range.</summary>
        public int Start;

        /// <summary>Позиция за последним символом абзаца (включая знак абзаца).</summary>
        public int End;

        /// <summary>Сырой текст абзаца вместе со служебными символами.</summary>
        public string Raw;

        /// <summary>Текст без служебных символов, обрезанный по краям.</summary>
        public string Text;

        /// <summary>Абзац является ячейкой таблицы (заканчивается маркером ячейки).</summary>
        public bool InTable;

        public bool IsEmpty
        {
            get { return string.IsNullOrWhiteSpace(Text); }
        }

        public override string ToString()
        {
            return Start + ".." + End + " " + Text;
        }
    }

    /// <summary>
    /// Текстовая модель документа: весь текст основного потока читается одним
    /// вызовом COM, дальше вся работа идёт над строкой в памяти.
    ///
    /// Индекс i в строке соответствует позиции Base + i в координатах Range.
    /// Для документов без полей и внедрённых объектов это соответствие точное;
    /// на всякий случай ResolveRange умеет корректировать сдвиг.
    /// </summary>
    public sealed class TextModel
    {
        public string Text { get; private set; }
        public int Base { get; private set; }
        public List<ParaSpan> Paragraphs { get; private set; }

        private TextModel() { }

        public static TextModel Build(dynamic doc)
        {
            var model = new TextModel
            {
                Text = WordUtil.GetContentText(doc),
                Base = WordUtil.ContentStart(doc),
                Paragraphs = new List<ParaSpan>()
            };
            model.SplitParagraphs();
            return model;
        }

        private void SplitParagraphs()
        {
            int start = 0;
            for (int i = 0; i < Text.Length; i++)
            {
                char c = Text[i];
                if (c != '\r') continue;

                // Ячейка таблицы: за знаком абзаца идёт маркер \a.
                int end = i + 1;
                bool inTable = end < Text.Length && Text[end] == '\a';
                if (inTable) end++;

                AddSpan(start, end, inTable);
                start = end;
                i = end - 1;
            }

            if (start < Text.Length) AddSpan(start, Text.Length, false);
        }

        private void AddSpan(int start, int end, bool inTable)
        {
            string raw = Text.Substring(start, end - start);
            Paragraphs.Add(new ParaSpan
            {
                Start = Base + start,
                End = Base + end,
                Raw = raw,
                Text = WordUtil.CleanParagraph(raw),
                InTable = inTable
            });
        }

        /// <summary>Смещение в строке Text по позиции документа.</summary>
        public int ToIndex(int position)
        {
            return position - Base;
        }

        /// <summary>Позиция документа по смещению в строке Text.</summary>
        public int ToPosition(int index)
        {
            return index + Base;
        }

        /// <summary>Абзац, содержащий указанную позицию документа, или null.</summary>
        public ParaSpan ParagraphAt(int position)
        {
            if (Paragraphs.Count == 0) return null;
            int lo = 0, hi = Paragraphs.Count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                var p = Paragraphs[mid];
                if (position < p.Start) hi = mid - 1;
                else if (position >= p.End) lo = mid + 1;
                else return p;
            }
            return null;
        }

        /// <summary>Индекс абзаца, содержащего позицию, или -1.</summary>
        public int IndexOfParagraphAt(int position)
        {
            int lo = 0, hi = Paragraphs.Count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                var p = Paragraphs[mid];
                if (position < p.Start) hi = mid - 1;
                else if (position >= p.End) lo = mid + 1;
                else return mid;
            }
            return -1;
        }

        /// <summary>Абзацы, целиком или частично попадающие в диапазон [start, end).</summary>
        public IEnumerable<ParaSpan> ParagraphsIn(int start, int end)
        {
            foreach (var p in Paragraphs)
            {
                if (p.End <= start) continue;
                if (p.Start >= end) yield break;
                yield return p;
            }
        }
    }
}
