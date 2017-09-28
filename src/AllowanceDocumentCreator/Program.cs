using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AllowanceDocumentCreator
{
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                var dataFilePath = GetDataFilePath(args);
                WriteToConsole($"Выбранный файл: {dataFilePath}", ConsoleColor.Yellow);

                var inputDataItems = GetInputDataItems(dataFilePath);
                var outputData = GetOutputData(inputDataItems);

                var copyDocumentTemplateResult = CopyDocumentTemplate(dataFilePath);
                if (!copyDocumentTemplateResult.IsSuccess)
                {
                    WriteErrorMessage(copyDocumentTemplateResult.Message);
                    return;
                }

                var outputDocumentPath = copyDocumentTemplateResult.Data;
                WriteDataToDocument(outputData, outputDocumentPath);

                DeleteDocumentTemplate(outputDocumentPath);
            }
            catch (Exception e)
            {
                WriteErrorMessage(e);
            }

            Console.WriteLine("Нажмите любую клавишу, чтобы завершить работу программы");
            Console.Read();
        }

        private static string GetDataFilePath(string[] args)
        {
            string dataFilePath;

            if (args.Length == 0)
            {
                Console.WriteLine("Введите путь к файлу");
                dataFilePath = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(dataFilePath))
                {
                    dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"sample_data.xlsx");
                    Console.WriteLine(dataFilePath);
                }
            }
            else
            {
                dataFilePath = args[0];
            }
            return dataFilePath;
        }

        private static InputDataItem[] GetInputDataItems(string dataFilePath)
        {
            List<InputDataItem> inputDataItems;
            WriteToConsole($"Получение данных из файла {Path.GetFileName(dataFilePath)}", ConsoleColor.Yellow);
            Console.WriteLine("Открытие файла...");
            using (DocumentReader documentReader = new DocumentReader(dataFilePath))
            {
                WriteToConsole("Файл открыт", ConsoleColor.Green);
                Console.WriteLine("Считывание данных...");
                inputDataItems = documentReader.Read();
                WriteToConsole("Данные считаны", ConsoleColor.Green);
            }
            return inputDataItems.ToArray();
        }

        private static OutputData GetOutputData(InputDataItem[] inputDataItems)
        {
            var outputDataItems = inputDataItems.Select(x => new OutputDataItem(x))
                                                .ToArray();
            return new OutputData(outputDataItems);
        }

        private static Result<string> CopyDocumentTemplate(string dataFilePath)
        {
            var documentTemplateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"document_template.xlsx");
            if (!File.Exists(documentTemplateFilePath))
            {
                return Result.Fault<string>($"Не найден шаблон документа по пути {documentTemplateFilePath}");
            }

            var dataFileName = Path.GetTempFileName();
            var outputDirectoryPath = Path.GetDirectoryName(dataFilePath);

            var outputDocumentPath = Path.Combine(outputDirectoryPath, $"{dataFileName}.tmp");

            Console.WriteLine("Копирование файла шаблона...");
            File.Copy(documentTemplateFilePath, outputDocumentPath);
            WriteToConsole("Копирование файла шаблона завершено", ConsoleColor.Green);

            return Result.Success(outputDocumentPath);
        }

        private static void DeleteDocumentTemplate(string filePath)
        {
            File.Delete(filePath);
        }

        private static void WriteDataToDocument(OutputData outputData, string outputDocumentPath)
        {
            Console.WriteLine("Открытие файла на запись...");
            using (DocumentWriter dataReader = new DocumentWriter(outputDocumentPath))
            {
                WriteToConsole("Файл открыт", ConsoleColor.Green);
                Console.WriteLine("Запись данных...");
                dataReader.Write(outputData);
                WriteToConsole("Данные записаны", ConsoleColor.Green);
            }
        }

        private static void WriteErrorMessage(object obj)
        {
            using (new UsingConsoleColor(ConsoleColor.Red))
            {
                Console.WriteLine("Произошла ошибка");
                Console.WriteLine(obj);
            }
        }

        private static void WriteToConsole(string str, ConsoleColor foreground)
        {
            using (new UsingConsoleColor(foreground))
            {
                Console.WriteLine(str);
            }
        }
    }
}
