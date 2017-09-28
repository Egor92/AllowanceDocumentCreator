using System;
using System.Linq;
using MoreLinq;

namespace AllowanceDocumentCreator
{
    public class CellIndex
    {
        #region Ctor

        public CellIndex(string column, int row)
        {
            if (string.IsNullOrEmpty(column))
                throw new ArgumentException("Value cannot be null or empty.", nameof(column));
            if (row <= 0)
                throw new ArgumentOutOfRangeException(nameof(row));

            const int letterCount = 'Z' - 'A' + 1;
            Column = column.ToUpper()
                           .Select(x => x - 'A' + 1)
                           .Reverse()
                           .Index()
                           .Select(x => x.Value * (int) Math.Pow(letterCount, x.Key))
                           .Sum();
            Row = row;
        }

        private CellIndex(int column, int row)
        {
            if (column <= 0)
                throw new ArgumentOutOfRangeException(nameof(column));
            if (row <= 0)
                throw new ArgumentOutOfRangeException(nameof(row));

            Column = column;
            Row = row;
        }

        #endregion

        #region Row

        public int Row { get; }

        #endregion

        #region Column

        public int Column { get; }

        #endregion
    }
}
