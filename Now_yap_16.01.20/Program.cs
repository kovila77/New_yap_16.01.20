using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Word = Microsoft.Office.Interop.Word;
//using Excel = Microsoft.Office.Interop.Excel;

namespace Now_yap_16._01._20
{
    class Program
    {
        static void Main(string[] args)
        {
            const string settingsFileName = @"C:\Users\35498\source\repos\New_yap_16.01.20\wishes.xlsx";


            List<List<string>> wishes = new List<List<string>>();
            List<string> names = new List<string>();

            ExcelSettingReader excelSettingReader = new ExcelSettingReader(settingsFileName);

            WordWishesWriter wordWishesWriter = new WordWishesWriter(excelSettingReader.TemplateName);

            names = excelSettingReader.GetNames();

            wishes = excelSettingReader.GetWishes();

            excelSettingReader.closeApp();

            WishesGenerator wishesGenerator = new WishesGenerator(wishes);

            //wishesGenerator.writeAllCombinations();

            wordWishesWriter.createCongratulationsDoc(wishesGenerator, names);

            wordWishesWriter.show();
            Console.ReadLine();
        }
    }
}
