using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excelApp = new Excel.Application();
            const string settingsFileName = @"C:\Users\35498\OneDrive\PSU\SYAP\New_yap_16.01.20\wishes.xlsx";
            string templateName;
            const string settingSheetName = "Settings";
            const string namesSheetName = "Names";
            const string wishesSheetName = "Wishes";
            int countWishGenerationAtOnce = 3;
            Excel.Worksheet worksheet;
            Excel.Workbook workbook;
            List<List<string>> wishes = new List<List<string>>();
            List<string> names = new List<string>();

            // try open excel file
            try
            {
                workbook = excelApp.Workbooks.Open(settingsFileName);
            }
            catch (Exception e)
            {
                Console.WriteLine("Unable to open settings file");
                Console.WriteLine(e.Message);
                excelApp.Quit();
                Console.ReadLine();
                return;
            }

            //try open sheet Settings
            try
            {
                worksheet = (Excel.Worksheet)workbook.Worksheets[settingSheetName];
                templateName = worksheet.Cells[1, 2].Text;
                countWishGenerationAtOnce = Convert.ToInt32(worksheet.Cells[2, 2].Value);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Unable to open \"{settingSheetName}\" sheet");
                Console.WriteLine(e.Message);
                excelApp.Quit();
                Console.ReadLine();
                return;
            }

            //try load names sheet
            try
            {
                worksheet = (Excel.Worksheet)workbook.Worksheets[namesSheetName];
                for (int i = 1; worksheet.Cells[i, 1].Value != null; i++)
                    names.Add(worksheet.Cells[i, 1].Text);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Unable to open \"{namesSheetName}\" sheet");
                Console.WriteLine(e.Message);
                excelApp.Quit();
                Console.ReadLine();
                return;
            }

            //try load wishes sheet
            try
            {
                worksheet = (Excel.Worksheet)workbook.Worksheets[wishesSheetName];
                for (int i = 1; worksheet.Cells[1, i].Value != null; i++)
                {
                    wishes.Add(new List<string>());
                    for (int j = 2; worksheet.Cells[j, i].Value != null; j++)
                    {
                        wishes[i - 1].Add(worksheet.Cells[j, i].Text);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Unable to open \"{wishesSheetName}\" sheet");
                Console.WriteLine(e.Message);
                excelApp.Quit();
                Console.ReadLine();
                return;
            }

            //for (int i = 0; i < wishes.Count; i++)
            //    for (int j = 0; j < wishes[i].Count; j++)
            //        Console.WriteLine(wishes[i][j]);
            //foreach (var t in names)
            //{
            //    Console.WriteLine(t);
            //}
            //Console.WriteLine(templateName);

            excelApp.Quit();

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

            WishesGenerator wishesGenerator = new WishesGenerator(wishes, countWishGenerationAtOnce);

            for (int j = 0; j < names.Count; j++)
            {
                wishesGenerator.generateNewWish();
                wordApp.ActiveDocument.Bookmarks["Name"].Range.Text = names[j];
                for (int i = 0; i < countWishGenerationAtOnce; i++)
                {
                    wordApp.ActiveDocument.Bookmarks[$"Wish{i + 1}"].Range.Text = wishesGenerator.Wishes[i];
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

        class WishesGenerator
        {
            private List<List<string>> allWishes;
            private List<string> wishes;
            private int countExistW = 0;
            private int countWishGenerationAtOnce;

            public List<string> Wishes { get { return wishes; } }

            public WishesGenerator(List<List<string>> allWishes, int countWishGenerationAtOnce)
            {
                this.allWishes = allWishes;
                if (countWishGenerationAtOnce > 0)
                    this.countWishGenerationAtOnce = countWishGenerationAtOnce;
                else
                    this.countWishGenerationAtOnce = 1;
                wishes = new List<string>();
                wishes.Add("счастья");
                wishes.Add("здоровья");
                wishes.Add("денег");
            }

            public void generateNewWish()
            {
            }
        }
    }
}
