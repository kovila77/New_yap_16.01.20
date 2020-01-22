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
            const string settingsFileName = @"C:\Users\35498\source\repos\New_yap_16.01.20\wishes.xlsx";
            string templateName;
            const string settingSheetName = "Settings";
            const string namesSheetName = "Names";
            const string wishesSheetName = "Wishes";
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

            WishesGenerator wishesGenerator = new WishesGenerator(wishes);
            wishesGenerator.generateWishes();

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

        class WishesGenerator
        {
            private List<List<string>> allWishes;
            private List<List<string>> combinationsWishes;
            private List<string> wishesTrio;
            private int countExistW = 0;
            private int gg;

            public List<string> WishesTrio { get { return wishesTrio; } }

            public WishesGenerator(List<List<string>> allWishes)
            {
                this.allWishes = allWishes;
                wishesTrio = new List<string>();
                //wishesTrio.Add(IEnumerable)"счастья");
                // wishesTrio.Add("здоровья");
                //wishesTrio.Add("денег");
            }

            public void generateWishes()
            {
                combinationsWishes = new List<List<string>>();
                for (int i = 0; i < allWishes.Count; i++)
                {
                    for (int j = i + 1; j < allWishes.Count; j++)
                    {
                        for (int k = j + 1; k < allWishes.Count; k++)
                        {
                            var tmp = new List<List<string>>();
                            tmp.Add(allWishes[i]);
                            tmp.Add(allWishes[j]);
                            tmp.Add(allWishes[k]);
                            combinationsWishes.AddRange(CartesianProduct(tmp));

                            gg = 0;
                        }
                    }
                }
            }

            public void newTrio()
            {
                wishesTrio = combinationsWishes[gg++];
            }

            //private IEnumerable<IEnumerable<string>> CartesianProduct(IEnumerable<IEnumerable<string>> sequences)
            //{
            //    IEnumerable<IEnumerable<string>> result = new[] { Enumerable.Empty<string>() };
            //    foreach (var sequence in sequences)
            //    {
            //        //var s = sequence;
            //        //result =
            //        //from seq in result
            //        //from item in s
            //        //select seq.Concat(new[] { item });
            //        foreach (var seq in result)
            //        {
            //            foreach (var item in sequence)
            //            {

            //            }
            //        }
            //    }
            //    return result;
            //}

            private List<List<string>> CartesianProduct(List<List<string>> sequences)
            {
                List<List<string>> result = new List<List<string>>();
                result.Add(new List<string>());
                foreach (var sequence in sequences)
                {
                    //var s = sequence;
                    //result =
                    //from seq in result
                    //from item in s
                    //select seq.Concat(new[] { item });
                    foreach (var seq in result)
                    {
                        foreach (var item in sequence)
                        {
                            seq.Add(item);
                        }
                    }
                }
                return result;
            }
        }
    }
}
