using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Text;
using System.Threading.Tasks;

namespace GeneralskiiTask
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            Console.Write("Введите кол-во строк массива: ");
            int n;
            while (!int.TryParse(Console.ReadLine(), out n) || n <= 0)
            {
                Console.Write("Введите положительное число и число отличающееся от 0: ");
            }

            string[] generatedArray = GenerateStringArray(n);

            var sortedAscending = generatedArray.OrderBy(s => s).ToArray();
            var sortedDescending = generatedArray.OrderByDescending(s => s).ToArray();

            Console.WriteLine("Начальные сгенерированные данные:");
            foreach (var item in generatedArray)
            {
                Console.WriteLine(item);
            }

            Console.WriteLine("\nОтсортированные по возрастанию:");
            foreach (var item in sortedAscending)
            {
                Console.WriteLine(item);
            }

            Console.WriteLine("\nОтсортированные по убыванию:");
            foreach (var item in sortedDescending)
            {
                Console.WriteLine(item);
            }

            Console.Write("Введите путь для сохранения файла (например, C:\\Users\\User\\Desktop): ");
            string filePath = Console.ReadLine();

            if (!Directory.Exists(filePath))
            {
                Console.WriteLine("Указанный путь не существует.");
                Console.WriteLine("Проверьте введенный путь и попробуйте снова.");
                return;
            }

            string fileName = $"SortedData-{DateTime.Now:yyyy-MMMM-dd HH-mm-ss}.xlsx";
            string fullPath = Path.Combine(filePath, fileName);

            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Data");

                    worksheet.Cells[1, 1].Value = "Начальные данные";
                    worksheet.Cells[1, 2].Value = "Отсортированные по возрастанию";
                    worksheet.Cells[1, 3].Value = "Отсортированные по убыванию";

                    for (int i = 0; i < n; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = generatedArray[i];
                        worksheet.Cells[i + 2, 2].Value = sortedAscending[i];
                        worksheet.Cells[i + 2, 3].Value = sortedDescending[i];
                    }

                    package.SaveAs(new FileInfo(fullPath));
                }

                Console.WriteLine($"Данные успешно сохранены в {fullPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при сохранении файла: {ex.Message}");
            }

            Console.WriteLine("Нажмите любую клавишу для завершения...");
            Console.ReadKey();
        }

        static string[] GenerateStringArray(int n)
        {
            Random random = new Random();
            string[] array = new string[n];
            for (int i = 0; i < n; i++)
            {
                array[i] = GenerateRandomString(random);
            }
            return array;
        }

        static string GenerateRandomString(Random random)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            int length = random.Next(10, 21);
            var stringChars = new char[length];
            bool hasLetter = false;
            bool hasDigit = false;

            for (int i = 0; i < length; i++)
            {
                char c = chars[random.Next(chars.Length)];
                if (char.IsLetter(c)) hasLetter = true;
                if (char.IsDigit(c)) hasDigit = true;
                stringChars[i] = c;
            }

            if (!hasLetter)
            {
                stringChars[random.Next(length)] = 'A';
            }
            if (!hasDigit)
            {
                stringChars[random.Next(length)] = '0';
            }

            return new string(stringChars);
        }
    }
}
