using System;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace AllowanceDocumentCreator
{
    public sealed class DocumentWriter : IDisposable
    {
        private enum Alignment
        {
            Left,
            Right,
        }

        #region Fields

        private readonly string _filePath;
        private Application _excelApp;
        private Workbook _workBook;
        private Worksheet _workSheet;

        #endregion

        #region Ctor

        public DocumentWriter(string filePath)
        {
            _filePath = filePath;
            OpenDocument();
        }

        #endregion

        #region Implementation of IDisposable

        public void Dispose()
        {
            Marshal.ReleaseComObject(_workSheet);

            _workBook.Close();
            Marshal.ReleaseComObject(_workBook);

            _excelApp.Quit();
            Marshal.ReleaseComObject(_excelApp);
        }

        #endregion

        private void OpenDocument()
        {
            _excelApp = new Application();
            _workBook = _excelApp.Workbooks.Open(_filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true,
                                                1, 0);
            _workSheet = (Worksheet) _workBook.Worksheets.Item[1];
        }

        public void Write(OutputData outputData)
        {
            for (int i = 0; i < outputData.Items.Length; i++)
            {
                var outputDataItem = outputData.Items[i];
                WriteText(outputDataItem.LastName, new CellIndex("D", GetOffset(i, 0)), 19, Alignment.Left);
                WriteText(outputDataItem.FirstName, new CellIndex("D", GetOffset(i, 1)), 19, Alignment.Left);
                WriteText(outputDataItem.FatherName, new CellIndex("D", GetOffset(i, 2)), 19, Alignment.Left);

                WriteText(outputDataItem.DaysCount, new CellIndex("Y", GetOffset(i, 0)), 1, Alignment.Left);

                WriteText(outputDataItem.A, new CellIndex("AB", GetOffset(i, 0)), 7, Alignment.Right);

                WriteText(outputDataItem.B, new CellIndex("AK", GetOffset(i, 0)), 8, Alignment.Right);

                WriteText(outputDataItem.C, new CellIndex("AU", GetOffset(i, 0)), 8, Alignment.Right);
                WriteToCell(outputDataItem.CPercent, new CellIndex("BD", GetOffset(i, 0)));

                WriteText(outputDataItem.D1, new CellIndex("BG", GetOffset(i, 0)), 8, Alignment.Right);
                WriteToCell(outputDataItem.D1Percent, new CellIndex("BP", GetOffset(i, 0)));
                WriteText(outputDataItem.D2, new CellIndex("BG", GetOffset(i, 2)), 8, Alignment.Right);
                WriteToCell(outputDataItem.D2Percent, new CellIndex("BP", GetOffset(i, 2)));

                WriteText(outputDataItem.E, new CellIndex("BS", GetOffset(i, 0)), 8, Alignment.Right);
                WriteToCell(outputDataItem.EPercent, new CellIndex("CB", GetOffset(i, 0)));
            }

            WriteText(outputData.B, new CellIndex("AG", 71), 10, Alignment.Right);
            WriteText(outputData.C, new CellIndex("AU", 71), 10, Alignment.Right);
            WriteText(outputData.D, new CellIndex("BG", 71), 10, Alignment.Right);
            WriteText(outputData.E, new CellIndex("BS", 71), 10, Alignment.Right);

            WriteText(outputData.TotalRubles, new CellIndex("F", 74), 8, Alignment.Right);
            WriteText(outputData.TotalKopecks, new CellIndex("R", 74), 2, Alignment.Right);
        }

        private void WriteText(int value, CellIndex startCellIndex, int cellCount, Alignment alignment)
        {
            WriteText(value.ToString(), startCellIndex, cellCount, alignment);
        }

        private void WriteText(double value, CellIndex startCellIndex, int cellCount, Alignment alignment)
        {
            var cultureInfo = CultureInfo.GetCultureInfo("en-US");
            var text = value.ToString("F2", cultureInfo);
            WriteText(text, startCellIndex, cellCount, alignment);
        }

        private void WriteText(string text, CellIndex startCellIndex, int cellCount, Alignment alignment)
        {
            var spaceCount = Math.Max(cellCount - text.Length, 0);
            var spaces = Enumerable.Repeat(' ', spaceCount);
            var textChars = text.ToCharArray();

            var chars = alignment == Alignment.Left
                ? Enumerable.Concat(textChars, spaces)
                            .ToList()
                : Enumerable.Concat(spaces, textChars)
                            .ToList();

            for (int i = 0; i < chars.Count; i++)
            {
                var symbol = chars[i];
                int rowIndex = startCellIndex.Row;
                int columnIndex = startCellIndex.Column + i;
                _workSheet.Cells[rowIndex, columnIndex] = symbol.ToString();
            }
        }

        private void WriteToCell(double value, CellIndex cellIndex)
        {
            var cultureInfo = CultureInfo.GetCultureInfo("en-US");
            var text = value.ToString(cultureInfo);
            WriteToCell(text, cellIndex);
        }

        private void WriteToCell(string text, CellIndex cellIndex)
        {
            _workSheet.Cells[cellIndex.Row, cellIndex.Column] = text;
        }

        private static int GetOffset(int tableRowNumber, int lineNumber)
        {
            const int tableIndex = 38;
            const int rowHeight = 8;
            return tableIndex + tableRowNumber * rowHeight + lineNumber * 2 + 1;
        }
    }
}
