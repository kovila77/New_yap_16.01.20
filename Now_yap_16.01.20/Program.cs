using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Now_yap_16._01._20
{
    class Program
    {
        static void Main(string[] args)
        {
            string settingsFileName = @"C:\Users\35498\source\repos\New_yap_16.01.20\wishes.xlsx";
            
            List<List<string>> wishes;
            List<string> names;

            Console.WriteLine("Стандартный пут к файлу Excel? д/н");
            if (Console.ReadLine() == "н")
            {
                Console.WriteLine("Введите путь: ");
                settingsFileName = Console.ReadLine();
            }

            Console.WriteLine($"Установленный путь: {settingsFileName}");
            ExcelSettingReader excelSettingReader = new ExcelSettingReader(settingsFileName);

            Console.WriteLine($"Чтение имени файла шаблона из листа {excelSettingReader.settingSheetName} в ячейке {excelSettingReader.xTemplateName} {excelSettingReader.yTemplateName}...");
            WordWishesWriter wordWishesWriter = new WordWishesWriter(excelSettingReader.TemplateName);

            wordWishesWriter.fontName = excelSettingReader.FontName;

            Console.WriteLine($"Чтение имён с листа {excelSettingReader.namesSheetName}...");
            names = excelSettingReader.GetNames();
            if (names.Count < 1)
            {
                Console.WriteLine($"Нет имён :\\");
                Console.ReadKey();
                excelSettingReader.closeApp();
                return;
            }

            Console.WriteLine($"Чтение пожеланий с листа {excelSettingReader.wishesSheetName}...");
            wishes = excelSettingReader.GetWishes();

            excelSettingReader.closeApp();

            if (wishes.Count < 3)
            {
                Console.WriteLine($"Слишком мало групп пожеланий!!!");
                Console.ReadKey();
                return;
                //throw new ArgumentException("Слишком мало тем пожеланий");
            }

            WishesGenerator wishesGenerator = new WishesGenerator(wishes);

            if (!wishesGenerator.isThereEnoughtCombination(names.Count()))
            {
                Console.WriteLine($"Слишком мало пожеланий для такого количества имён!!!");
                Console.ReadKey();
                return;
            }

            Console.WriteLine($"Генерация пожеланий...");
            wishesGenerator.generateWishes();

            //wishesGenerator.writeAllCombinations();

            Console.WriteLine($"Сброс комбинаций в файл Docx");
            wordWishesWriter.createCongratulationsDoc(wishesGenerator, names);

            wordWishesWriter.showWord();

            Console.WriteLine($"Для выхода нажмите любую кнопку...");
            Console.ReadKey();
        }
    }
}
