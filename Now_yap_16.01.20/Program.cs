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
                templateName = worksheet.Cells[2, 2].Text;
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
                        wishes[i-1].Add(worksheet.Cells[j, i].Text);
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
            Console.ReadLine();
            //excelApp.Quit();
            //Document doc = null;

            //try
            //{
            //    doc = app.Documents.Add(templateName);
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine("Unable to open template");
            //    Console.WriteLine(e.Message);
            //    app.Quit();
            //    Console.ReadLine();
            //    return;
            //}
            //app.ActiveDocument.Bookmarks["Name"].Range.Text = "n1";
            //app.ActiveDocument.Bookmarks["Wish1"].Range.Text = "w11";
            //app.ActiveDocument.Bookmarks["Wish2"].Range.Text = "w12";
            //app.ActiveDocument.Bookmarks["Wish3"].Range.Text = "w13";

            //app.Selection.EndKey(WdUnits.wdStory);
            //app.Selection.InsertNewPage();
            //app.Selection.InsertFile(templateName, "", false, false, false);

            //app.ActiveDocument.Bookmarks["Name"].Range.Text = "n2";
            //app.ActiveDocument.Bookmarks["Wish1"].Range.Text = "w21";
            //app.ActiveDocument.Bookmarks["Wish2"].Range.Text = "w22";
            //app.ActiveDocument.Bookmarks["Wish3"].Range.Text = "w23";

            //Console.WriteLine("Generation complited");
            //Console.ReadLine();
            //app.Visible = true;


        }
    }
}
