using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace AllowanceDocumentCreator
{
    public class RowData
    {
        public double A { get; set; }

        public double B { get; set; }

        public double C { get; set; }

        public double D1 { get; set; }

        public double D2 { get; set; }

        public double E { get; set; }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            string dataFilePath = string.Empty;

            if (args.Length == 0)
            {
                Console.WriteLine("Введите путь к файлу");
                dataFilePath = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(dataFilePath))
                {
                    dataFilePath = Path.Combine(Directory.GetCurrentDirectory(), @"docs\sample_data.xlsx");
                    Console.WriteLine(dataFilePath);
                }
            }
            else if (args.Length > 1)
            {
                dataFilePath = args[0];
            }

            Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            try
            {
                xlApp = new Application();
                xlWorkBook = xlApp.Workbooks.Open(dataFilePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0,
                                                  true, 1, 0);
                xlWorkSheet = (Worksheet) xlWorkBook.Worksheets.Item[1];
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Произошла ошибка при открытии документа");
                Console.WriteLine(e);
                Console.ForegroundColor = ConsoleColor.Gray;
                return;
            }

            List<RowData> data = new List<RowData>();
            try
            {
                while (true)
                {
                    int rowIndex = data.Count * 2 + 1;
                    var hasData = !string.IsNullOrWhiteSpace(xlWorkSheet.Cells[rowIndex, 1].Text);
                    if (!hasData)
                        break;

                    var rowData = new RowData
                    {
                        A = Convert.ToDouble(xlWorkSheet.Cells[rowIndex, 1].Text),
                        B = Convert.ToDouble(xlWorkSheet.Cells[rowIndex, 2].Text),
                        C = Convert.ToDouble(xlWorkSheet.Cells[rowIndex, 3].Text),
                        D1 = Convert.ToDouble(xlWorkSheet.Cells[rowIndex, 4].Text),
                        D2 = Convert.ToDouble(xlWorkSheet.Cells[rowIndex + 1, 4].Text),
                        E = Convert.ToDouble(xlWorkSheet.Cells[rowIndex, 5].Text),
                    };
                    data.Add(rowData);
                }
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Произошла ошибка при чтении документа");
                Console.WriteLine(e);
                Console.ForegroundColor = ConsoleColor.Gray;
                return;
            }

            Marshal.ReleaseComObject(xlWorkSheet);

            //close and release
            xlWorkBook.Close();
            Marshal.ReleaseComObject(xlWorkBook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
