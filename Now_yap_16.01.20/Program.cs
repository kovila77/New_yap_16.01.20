using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new Application();
            var templateName = @"C:\Users\35498\Desktop\PSU\nyap\template.dotx";
            Document doc = null;

            try
            {
                doc = app.Documents.Add(templateName);
            }
            catch (Exception e)
            {
                Console.WriteLine("Unable to open template");
                Console.WriteLine(e.Message);
                app.Quit();
                Console.ReadLine();
                return;
            }
            app.ActiveDocument.Bookmarks["Name"].Range.Text = "n1";
            app.ActiveDocument.Bookmarks["Wish1"].Range.Text = "w11";
            app.ActiveDocument.Bookmarks["Wish2"].Range.Text = "w12";
            app.ActiveDocument.Bookmarks["Wish3"].Range.Text = "w13";

            app.Selection.EndKey(WdUnits.wdStory);
            app.Selection.InsertNewPage();
            app.Selection.InsertFile(templateName, "", false, false, false);

            app.ActiveDocument.Bookmarks["Name"].Range.Text = "n2";
            app.ActiveDocument.Bookmarks["Wish1"].Range.Text = "w21";
            app.ActiveDocument.Bookmarks["Wish2"].Range.Text = "w22";
            app.ActiveDocument.Bookmarks["Wish3"].Range.Text = "w23";

            Console.WriteLine("Generation complited");
            Console.ReadLine();
            app.Visible = true;


        }
    }
}
