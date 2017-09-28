using System;
using System.Linq;

namespace AllowanceDocumentCreator
{
    public class OutputData
    {
        #region Ctor

        public OutputData(OutputDataItem[] items)
        {
            Items = items;

            B = Items.Sum(x => x.B);
            C = Items.Sum(x => x.C);
            D = Items.Sum(x => x.D1 + x.D2);
            E = Items.Sum(x => x.E);

            var total = B + C + D + E;
            TotalRubles = (int) total;
            TotalKopecks = (int)Math.Round(total *100 % 100);
        }

        #endregion

        #region Items

        public OutputDataItem[] Items { get; }

        #endregion

        #region B

        public double B { get; }

        #endregion

        #region C

        public double C { get; }

        #endregion

        #region D

        public double D { get; }

        #endregion

        #region E

        public double E { get; }

        #endregion

        #region TotalRubles

        public int TotalRubles { get; }

        #endregion

        #region TotalKopecks

        public int TotalKopecks { get; }

        #endregion
    }
}
