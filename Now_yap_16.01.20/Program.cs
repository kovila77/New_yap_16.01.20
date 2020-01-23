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
            wishesGenerator.generateWishes(12313);

            int ff = 0;
            for (int j = 0; j < names.Count; j++)
            {
                wishesGenerator.newTrio();
                for (int i = 0; i < 3; i++)
                {
                    Console.WriteLine((wishesGenerator.WishesTrio)[i]);
                    ff++;
                }
            }
            Console.WriteLine(ff);

            if (false)
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
            List<List<bool>> allWithesUsed;
            private List<List<string>> combinationsWishesUnique3;
            private int curCombW3;
            private List<List<string>> combinationsWishesUnique2;
            private int curCombW2;
            private List<List<string>> combinationsWishesUnique1;
            private int curCombW1;
            private List<string> wishesTrio;
            private int countExistW = 0;


            public List<string> WishesTrio { get { return wishesTrio; } }

            public WishesGenerator(List<List<string>> allWishes)
            {
                this.allWishes = allWishes;
                allWithesUsed = new List<List<bool>>();
                for (int i = 0; i < allWishes.Count; i++)
                {
                    allWithesUsed.Add(new List<bool>());
                    for (int j = 0; j < allWishes[i].Count; j++)
                    {
                        allWithesUsed[i].Add(false);
                    }
                }
                wishesTrio = new List<string>();
                //wishesTrio.Add(IEnumerable)"счастья");
                // wishesTrio.Add("здоровья");
                //wishesTrio.Add("денег");
            }

            public void generateWishes(int coutnWishToCreate)
            {
                combinationsWishesUnique1 = new List<List<string>>();
                combinationsWishesUnique2 = new List<List<string>>();
                combinationsWishesUnique3 = new List<List<string>>();
                bool[] topicsUse = new bool[allWishes.Count];

                for (int topic0 = 0; topic0 < allWishes.Count; topic0++)
                {
                    for (int topic1 = topic0 + 1; topic1 < allWishes.Count; topic1++)
                    {
                        for (int topic2 = topic1 + 1; topic2 < allWishes.Count; topic2++)
                        {
                            for (int i = 0; i < allWishes[topic0].Count; i++)
                            {
                                for (int j = 0; j < allWishes[topic1].Count; j++)
                                {
                                    for (int k = 0; k < allWishes[topic2].Count; k++)
                                    {
                                        if (notUse3Wish(topic0, topic1, topic2, i, j, k))
                                        {
                                            combinationsWishesUnique3.Add(
                                                new List<string> {
                                                (string)allWishes[topic0][i].Clone(),
                                                (string)allWishes[topic1][j].Clone(),
                                                (string)allWishes[topic2][k].Clone()
                                            });
                                        }
                                        else if (notUse2Wish(topic0, topic1, topic2, i, j, k))
                                        {
                                            combinationsWishesUnique2.Add(
                                                new List<string> {
                                                (string)allWishes[topic0][i].Clone(),
                                                (string)allWishes[topic1][j].Clone(),
                                                (string)allWishes[topic2][k].Clone()
                                            });
                                        }
                                        else
                                        {
                                            combinationsWishesUnique1.Add(
                                                new List<string> {
                                                (string)allWishes[topic0][i].Clone(),
                                                (string)allWishes[topic1][j].Clone(),
                                                (string)allWishes[topic2][k].Clone()
                                            });
                                        }
                                        allWithesUsed[topic0][i] = true;
                                        allWithesUsed[topic1][j] = true;
                                        allWithesUsed[topic2][k] = true;
                                    }
                                }
                            }
                        }
                    }
                }

                curCombW3 = curCombW2 = curCombW1 = 0;
            }

            public void newTrio()
            {
                if (!(curCombW3 == combinationsWishesUnique3.Count))
                {
                    wishesTrio = combinationsWishesUnique3[curCombW3++];
                }
                else if (!(curCombW2 == combinationsWishesUnique2.Count))
                {
                    wishesTrio = combinationsWishesUnique3[curCombW2++];
                }
                else
                {
                    wishesTrio = combinationsWishesUnique3[curCombW1++];
                }
            }

            private bool notUse3Wish(int topic0, int topic1, int topic2, int i, int j, int k)
            {
                return !allWithesUsed[topic0][i] && !allWithesUsed[topic1][j] && !allWithesUsed[topic2][k];
            }

            private bool notUse2Wish(int topic0, int topic1, int topic2, int i, int j, int k)
            {
                return !allWithesUsed[topic0][i] && !allWithesUsed[topic1][j] && !allWithesUsed[topic2][k]
                 && !allWithesUsed[topic0][i] && !allWithesUsed[topic1][j] && allWithesUsed[topic2][k]
                 && !allWithesUsed[topic0][i] && allWithesUsed[topic1][j] && !allWithesUsed[topic2][k]
                 && allWithesUsed[topic0][i] && !allWithesUsed[topic1][j] && !allWithesUsed[topic2][k];
            }

            private bool notUse1Wish(int topic0, int topic1, int topic2, int i, int j, int k)
            {
                return !allWithesUsed[topic0][i] && allWithesUsed[topic1][j] && allWithesUsed[topic2][k]
                 && allWithesUsed[topic0][i] && !allWithesUsed[topic1][j] && allWithesUsed[topic2][k]
                 && allWithesUsed[topic0][i] && allWithesUsed[topic1][j] && !allWithesUsed[topic2][k];
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

            private List<List<string>> CreatWishesCombination(List<List<string>> sourse)
            {
                if (sourse.Count != 3)
                    throw new ArgumentException("There was not 3 topic of wish");

                List<List<string>> result = new List<List<string>>();
                for (int i = 0; i < sourse[0].Count; i++)
                    for (int j = 0; j < sourse[1].Count; j++)
                        for (int k = 0; k < sourse[2].Count; k++)
                            result.Add(new List<string> {
                                (string)sourse[0][i].Clone(),
                                (string)sourse[1][j].Clone(),
                                (string)sourse[2][k].Clone()
                            });
                return result;
            }
        }
    }
}
