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
            Console.WriteLine("Author: Ducheved", Color.Yellow);
            Console.WriteLine("Version: 1.0", Color.Yellow);
            Console.WriteLine("License: Apache", Color.Yellow);
            Console.WriteLine("This program allows you to convert old doc files from a folder to xml and docx simultaneously.", Color.Yellow);
            Console.WriteLine();
            Console.WriteLine("Enter the path to the folder with .doc files:", Color.Cyan);
            string? inputFolder = Console.ReadLine();
            if (string.IsNullOrEmpty(inputFolder))
            {
                Console.WriteLine("Folder with .doc files is not specified.");
                return;
            }

            Console.WriteLine("Enter the path to the folder to save .xml and .docx files:", Color.Cyan);
            string? outputFolder = Console.ReadLine();

            if (string.IsNullOrEmpty(inputFolder) || !Directory.Exists(inputFolder))
            {
                Console.WriteLine("The specified folder with .doc files does not exist.");
                return;
            }

            if (string.IsNullOrEmpty(outputFolder))
            {
                Console.WriteLine("Folder for saving files is not specified.");
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

        static async Task ConvertFiles(string inputFolder, string xmlFolder, string docxFolder, string outputFolder)
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
                reportFile.WriteLine("File processing is complete. The following files could not be processed due to errors:");
                foreach (string failedFile in failedFiles)
                {
                    await reportFile.WriteLineAsync(failedFile);
                }
            }

            Console.WriteLine($"\nConversion report saved to file: {reportPath}");
        }

        static bool ConvertFile(Application word, string inputPath, string xmlFolder, string docxFolder)
        {
            string? fileName = Path.GetFileNameWithoutExtension(inputPath);
            if (fileName == null)
            {
                Console.WriteLine($"Failed to get file name from path: {inputPath}", Color.Red);
                return false;
            }

            Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId} started processing file {fileName}", Color.Blue);

            string outputDocxPath = Path.Combine(docxFolder, $"{fileName}.docx");
            string outputXmlPath = Path.Combine(xmlFolder, $"{fileName}.xml");

            Document doc = null;

            try
            {
                doc = word.Documents.Open(inputPath);

                doc.SaveAs2(outputDocxPath, WdSaveFormat.wdFormatDocumentDefault);
                if (File.Exists(outputDocxPath) && new FileInfo(outputDocxPath).Length > 0)
                {
                    Console.WriteLine($"Success: File {fileName}.doc successfully converted to DOCX: {outputDocxPath} (Processed by thread {Thread.CurrentThread.ManagedThreadId})", Color.Green);
                }
                else
                {
                    Console.WriteLine($"Error: Failed to convert file {fileName}.doc to DOCX (Processed by thread {Thread.CurrentThread.ManagedThreadId})", Color.Red);
                    return false;
                }

                doc.SaveAs2(outputXmlPath, WdSaveFormat.wdFormatXML);
                if (File.Exists(outputXmlPath) && new FileInfo(outputXmlPath).Length > 0)
                {
                    Console.WriteLine($"Success: File {fileName}.doc successfully converted to XML: {outputXmlPath} (Processed by thread {Thread.CurrentThread.ManagedThreadId})", Color.Green);
                }
                else
                {
                    Console.WriteLine($"Error: Failed to convert file {fileName}.doc to XML (Processed by thread {Thread.CurrentThread.ManagedThreadId})", Color.Red);
                    return false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: Failed to convert file {fileName}.doc: {e.Message} (Processed by thread {Thread.CurrentThread.ManagedThreadId})", Color.Red);
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
                Console.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId} finished processing file {fileName}. Memory used: {currentProcess.WorkingSet64 / 1024 / 1024} MB. CPU time: {currentProcess.TotalProcessorTime.TotalSeconds} seconds.", Color.Cyan);
            }

            return true;
        }
    }
}