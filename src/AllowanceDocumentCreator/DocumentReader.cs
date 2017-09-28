using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace AllowanceDocumentCreator
{
    public class DocumentReader : IDisposable
    {
        #region Fields

        private readonly string _filePath;
        private Application _xlApp;
        private Workbook _xlWorkBook;
        private Worksheet _xlWorkSheet;

        #endregion

        #region Ctor

        public DocumentReader(string filePath)
        {
            _filePath = filePath;
            OpenDocument();
        }

        #endregion

        #region Implementation of IDisposable

        public void Dispose()
        {
            Marshal.ReleaseComObject(_xlWorkSheet);

            _xlWorkBook.Close();
            Marshal.ReleaseComObject(_xlWorkBook);

            _xlApp.Quit();
            Marshal.ReleaseComObject(_xlApp);
        }

        #endregion

        public List<InputDataItem> Read()
        {
            List<InputDataItem> dataItems = new List<InputDataItem>();
            while (true)
            {
                int rowIndex = dataItems.Count * 2 + 1;
                var hasData = !string.IsNullOrWhiteSpace(_xlWorkSheet.Cells[rowIndex, 1].Text);
                if (!hasData)
                    break;

                var rowData = new InputDataItem
                {
                    A = Convert.ToDouble(_xlWorkSheet.Cells[rowIndex, 1].Text),
                    B = Convert.ToDouble(_xlWorkSheet.Cells[rowIndex, 2].Text),
                    C = Convert.ToDouble(_xlWorkSheet.Cells[rowIndex, 3].Text),
                    D1 = Convert.ToDouble(_xlWorkSheet.Cells[rowIndex, 4].Text),
                    D2 = Convert.ToDouble(_xlWorkSheet.Cells[rowIndex + 1, 4].Text),
                    E = Convert.ToDouble(_xlWorkSheet.Cells[rowIndex, 5].Text),
                };
                dataItems.Add(rowData);
            }

            return dataItems;
        }

        private void OpenDocument()
        {
            _xlApp = new Application();
            _xlWorkBook = _xlApp.Workbooks.Open(_filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true,
                                                1, 0);
            _xlWorkSheet = (Worksheet) _xlWorkBook.Worksheets.Item[1];
        }
    }
}
