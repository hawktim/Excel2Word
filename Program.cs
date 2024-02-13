using Excel2Word.Activity;
using System;
using System.IO;

namespace Excel2Word
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var appPath = AppDomain.CurrentDomain.BaseDirectory;
            var inFile = Path.Combine(appPath, "Data.xlsb");
            var outFile = Path.Combine(appPath, "отчет1.docx");

            if (!File.Exists(inFile))
            {
                Console.WriteLine("Приложение будет закрыто. Необходим файл с данными: " + inFile);
                Console.ReadKey();
                return;
            }
            try
            {
                if (File.Exists(outFile))
                    File.Delete(outFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Возникли проблемы с удалением файла");
                Console.WriteLine(ex.Message);
                Console.ReadKey();
                return;
            }

            var readXls = new ReadXls();
            Console.WriteLine("Чтение файла данных");
            var data = readXls.ReadFile(inFile);
            Console.WriteLine("Запись в файл");
            var writeWord = new BuildWord();
            writeWord.WriteFile(data, outFile);
            Console.WriteLine("Данные успешно сохраненны в файл:" + outFile);
            System.Diagnostics.Process.Start(outFile);
            Console.ReadKey();
        }
    }
}