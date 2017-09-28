using System;

namespace AllowanceDocumentCreator
{
    public sealed class UsingConsoleColor : IDisposable
    {
        private readonly ConsoleColor _originalBackground;
        private readonly ConsoleColor _originalForeground;

        public UsingConsoleColor(ConsoleColor foreground)
            : this(ConsoleColor.Black, foreground)
        {
        }

        public UsingConsoleColor(ConsoleColor background, ConsoleColor foreground)
        {
            _originalBackground = Console.BackgroundColor;
            Console.BackgroundColor = background;

            _originalForeground = Console.ForegroundColor;
            Console.ForegroundColor = foreground;
        }

        public void Dispose()
        {
            Console.BackgroundColor = _originalBackground;
            Console.ForegroundColor = _originalForeground;
        }
    }
}
