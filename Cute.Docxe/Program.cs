using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Console = Colorful.Console;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;

namespace DocConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Автор: Ducheved", Color.Yellow);
            Console.WriteLine("Версия: 1.0", Color.Yellow);
            Console.WriteLine("Лицензия: Apache", Color.Yellow);
            Console.WriteLine("Программа предоставляет возможность конвертации doc старых файлов из папки одновременно в xml и docx.", Color.Yellow);
            Console.WriteLine();
            Console.WriteLine("Введите путь к папке с .doc файлами:", Color.Cyan);
            string? inputFolder = Console.ReadLine();
            if (string.IsNullOrEmpty(inputFolder))
            {
                Console.WriteLine("Не указана папка с .doc файлами.");
                return;
            }

            Console.WriteLine("Введите путь к папке для сохранения .xml и .docx файлов:", Color.Cyan);
            string? outputFolder = Console.ReadLine();

            if (string.IsNullOrEmpty(inputFolder) || !Directory.Exists(inputFolder))
            {
                Console.WriteLine("Указанная папка с .doc файлами не существует.");
                return;
            }

            if (string.IsNullOrEmpty(outputFolder))
            {
                Console.WriteLine("Не указана папка для сохранения файлов.");
                return;
            }

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            string xmlFolder = Path.Combine(outputFolder, "XML");
            string docxFolder = Path.Combine(outputFolder, "DOCX");

            if (!Directory.Exists(xmlFolder))
            {
                Directory.CreateDirectory(xmlFolder);
            }

            if (!Directory.Exists(docxFolder))
            {
                Directory.CreateDirectory(docxFolder);
            }

            ConvertFiles(inputFolder, xmlFolder, docxFolder, outputFolder).Wait();
        }

        static async System.Threading.Tasks.Task ConvertFiles(string inputFolder, string xmlFolder, string docxFolder, string outputFolder)
        {
            string[] docFiles = Directory.GetFiles(inputFolder, "*.doc");
            int maxWorkers = Environment.ProcessorCount;

            ConcurrentQueue<string> fileQueue = new ConcurrentQueue<string>(docFiles);
            ConcurrentBag<string> failedFiles = new ConcurrentBag<string>();

            var partitioner = Partitioner.Create(docFiles, true);
            Parallel.ForEach(partitioner, new ParallelOptions { MaxDegreeOfParallelism = maxWorkers }, filePath =>
            {
                Application word = new Application();

                if (!ConvertFile(word, filePath, xmlFolder, docxFolder))
                {
                    failedFiles.Add(filePath);
                }

                word.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(word);
            });

            string reportPath = Path.Combine(outputFolder, "conversion_report.txt");
            using (StreamWriter reportFile = new StreamWriter(new FileStream(reportPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, true)))
            {
                reportFile.WriteLine("Обработка файлов завершена. Не удалось обработать следующие файлы из-за ошибок:");
                foreach (string failedFile in failedFiles)
                {
                    await reportFile.WriteLineAsync(failedFile);
                }
            }

            Console.WriteLine($"\nОтчет о конвертации сохранен в файл: {reportPath}");
        }

        static bool ConvertFile(Application word, string inputPath, string xmlFolder, string docxFolder)
        {
            string? fileName = Path.GetFileNameWithoutExtension(inputPath);
            if (fileName == null)
            {
                Console.WriteLine($"Не удалось получить имя файла из пути: {inputPath}", Color.Red);
                return false;
            }

            Console.WriteLine($"Поток {Thread.CurrentThread.ManagedThreadId} начал обработку файла {fileName}", Color.Blue);

            string outputDocxPath = Path.Combine(docxFolder, $"{fileName}.docx");
            string outputXmlPath = Path.Combine(xmlFolder, $"{fileName}.xml");

            Document doc = null;

            try
            {
                doc = word.Documents.Open(inputPath);

                doc.SaveAs2(outputDocxPath, WdSaveFormat.wdFormatDocumentDefault);
                if (File.Exists(outputDocxPath) && new FileInfo(outputDocxPath).Length > 0)
                {
                    Console.WriteLine($"Успех: Файл {fileName}.doc успешно конвертирован в DOCX: {outputDocxPath} (Обработан потоком {Thread.CurrentThread.ManagedThreadId})", Color.Green);
                }
                else
                {
                    Console.WriteLine($"Ошибка: Не удалось конвертировать файл {fileName}.doc в DOCX (Обработан потоком {Thread.CurrentThread.ManagedThreadId})", Color.Red);
                    return false;
                }

                doc.SaveAs2(outputXmlPath, WdSaveFormat.wdFormatXML);
                if (File.Exists(outputXmlPath) && new FileInfo(outputXmlPath).Length > 0)
                {
                    Console.WriteLine($"Успех: Файл {fileName}.doc успешно конвертирован в XML: {outputXmlPath} (Обработан потоком {Thread.CurrentThread.ManagedThreadId})", Color.Green);
                }
                else
                {
                    Console.WriteLine($"Ошибка: Не удалось конвертировать файл {fileName}.doc в XML (Обработан потоком {Thread.CurrentThread.ManagedThreadId})", Color.Red);
                    return false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Ошибка: Не удалось конвертировать файл {fileName}.doc: {e.Message} (Обработан потоком {Thread.CurrentThread.ManagedThreadId})", Color.Red);
                return false;
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }
                Console.ResetColor();
                Process currentProcess = Process.GetCurrentProcess();
                Console.WriteLine($"Поток {Thread.CurrentThread.ManagedThreadId} завершил обработку файла {fileName}. Использовано памяти: {currentProcess.WorkingSet64 / 1024 / 1024} МБ. Загрузка процессора: {currentProcess.TotalProcessorTime.TotalSeconds} секунд.", Color.Cyan);
            }

            return true;
        }
    }
}