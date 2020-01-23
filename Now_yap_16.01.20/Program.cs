using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
//using Excel = Microsoft.Office.Interop.Excel;

namespace Now_yap_16._01._20
{
    class Program
    {
        static void Main(string[] args)
        {
            //Excel.Application excelApp = new Excel.Application();
            const string settingsFileName = @"C:\Users\35498\source\repos\New_yap_16.01.20\wishes.xlsx";
            string templateName;

            List<List<string>> wishes = new List<List<string>>();
            List<string> names = new List<string>();

            ExcelSettingReader excelSettingReader = new ExcelSettingReader(settingsFileName);

            templateName = excelSettingReader.TemplateName;
            templateName = excelSettingReader.TemplateName;

            names = excelSettingReader.GetNames();

            wishes = excelSettingReader.GetWishes();


            Word.Application wordApp = new Word.Application();
            Word.Document wDoc = null;

            //try create doc
            try
            {
                wDoc = wordApp.Documents.Add(templateName);
            }
            catch (Exception e)
            {
                Console.WriteLine("Unable to open template");
                Console.WriteLine(e.Message);
                wordApp.Quit();
                Console.ReadLine();
                return;
            }

            WishesGenerator wishesGenerator = new WishesGenerator(wishes);

            //wishesGenerator.writeAllCombinations();

            for (int j = 0; j < names.Count; j++)
            {
                wishesGenerator.newTrio();
                wordApp.ActiveDocument.Bookmarks["Name"].Range.Text = names[j];
                for (int i = 0; i < 3; i++)
                {
                    wordApp.ActiveDocument.Bookmarks[$"Wish{i + 1}"].Range.Text = (wishesGenerator.WishesTrio)[i];
                }
                if (j + 1 < names.Count)
                {
                    wordApp.Selection.EndKey(Word.WdUnits.wdStory);
                    wordApp.Selection.InsertNewPage();
                    wordApp.Selection.InsertFile(templateName);
                }
            }
            wordApp.Visible = true;
            Console.ReadLine();
        }
    }
}
