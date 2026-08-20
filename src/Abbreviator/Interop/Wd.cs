namespace Abbreviator.Interop
{
    /// <summary>
    /// Числовые значения перечислений Word, нужные надстройке.
    /// При позднем связывании типизированные enum-ы Word недоступны, поэтому
    /// используем константы с теми же значениями, что и в объектной модели.
    /// </summary>
    public static class Wd
    {
        // WdColorIndex
        public const int NoHighlight = 0;
        public const int Black = 1;
        public const int Blue = 2;
        public const int Turquoise = 3;
        public const int BrightGreen = 4;
        public const int Pink = 5;
        public const int Red = 6;
        public const int Yellow = 7;
        public const int White = 8;
        public const int DarkBlue = 9;
        public const int Teal = 10;
        public const int Green = 11;
        public const int Violet = 12;
        public const int DarkRed = 13;
        public const int DarkYellow = 14;
        public const int Gray50 = 15;
        public const int Gray25 = 16;

        // WdStatistic
        public const int StatisticPages = 2;

        // WdGoToItem
        public const int GoToPage = 1;

        // WdGoToDirection
        public const int GoToAbsolute = 1;

        // WdInformation
        public const int ActiveEndPageNumber = 3;
        public const int WithInTable = 12;

        // WdUnits
        public const int UnitCharacter = 1;
        public const int UnitWord = 2;
        public const int UnitParagraph = 4;
        public const int UnitStory = 6;

        // WdCollapseDirection
        public const int CollapseEnd = 0;
        public const int CollapseStart = 1;

        // WdStoryType
        public const int MainTextStory = 1;

        /// <summary>Человекочитаемое имя цвета подсветки (для окна настроек).</summary>
        public static string ColorName(int index)
        {
            switch (index)
            {
                case NoHighlight: return "нет";
                case Turquoise: return "бирюзовый";
                case BrightGreen: return "ярко-зелёный";
                case Pink: return "розовый";
                case Red: return "красный";
                case Yellow: return "жёлтый";
                case Teal: return "тёмно-бирюзовый";
                case Green: return "зелёный";
                case Violet: return "фиолетовый";
                case DarkRed: return "тёмно-красный";
                case DarkYellow: return "тёмно-жёлтый";
                case Gray50: return "серый 50%";
                case Gray25: return "серый 25%";
                case Blue: return "синий";
                case DarkBlue: return "тёмно-синий";
                default: return index.ToString();
            }
        }
    }
}
