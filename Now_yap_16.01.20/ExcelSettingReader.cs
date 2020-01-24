using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Now_yap_16._01._20
{
    class ExcelSettingReader
    {
        public string settingsFileName;
        public string settingSheetName = "Settings";
        public string namesSheetName = "Names";
        public string wishesSheetName = "Wishes";
        public int xTemplateName = 1;
        public int yTemplateName = 2;
        public int xfontName = 2;
        public int yfontName = 2;
        private string templateName = null;
        private string fontName = null;
        private Excel.Application excelApp = null;
        private Excel.Worksheet worksheetSettings = null;
        private Excel.Worksheet worksheetNames = null;
        private Excel.Worksheet worksheetWishes = null;
        private Excel.Workbook workbook = null;

        private Excel.Application ExcelApp { get { if (excelApp == null) excelApp = new Excel.Application(); return excelApp; } }
        private Excel.Worksheet WorksheetSettings { get { if (worksheetSettings == null) worksheetSettings = (Excel.Worksheet)Workbook.Worksheets[settingSheetName]; return worksheetSettings; } }
        private Excel.Worksheet WorksheetNames { get { if (worksheetNames == null) worksheetNames = (Excel.Worksheet)Workbook.Worksheets[namesSheetName]; return worksheetNames; } }
        private Excel.Worksheet WorksheetWishes { get { if (worksheetWishes == null) worksheetWishes = (Excel.Worksheet)Workbook.Worksheets[wishesSheetName]; return worksheetWishes; } }
        private Excel.Workbook Workbook { get { if (workbook == null) workbook = ExcelApp.Workbooks.Open(settingsFileName); return workbook; } }

        public string TemplateName { get { if (templateName == null) templateName = WorksheetSettings.Cells[xTemplateName, yTemplateName].Text; return templateName; } }
        public string FontName { get { if (fontName == null) fontName = WorksheetSettings.Cells[xfontName, yfontName].Text; return fontName; } }

        public ExcelSettingReader(string settingsFileName)
        {
            this.settingsFileName = settingsFileName;
        }

        public List<string> GetNames()
        {
            List<string> names = new List<string>();
            for (int i = 1; WorksheetNames.Cells[i, 1].Value != null; i++)
                names.Add(WorksheetNames.Cells[i, 1].Text);
            return names;
        }

        public List<List<string>> GetWishes()
        {
            List<List<string>> wishes = new List<List<string>>();
            for (int i = 1; WorksheetWishes.Cells[1, i].Value != null; i++)
            {
                wishes.Add(new List<string>());
                for (int j = 2; WorksheetWishes.Cells[j, i].Value != null; j++)
                {
                    wishes[i - 1].Add(WorksheetWishes.Cells[j, i].Text);
                }
            }
            return wishes;
        }

        public void closeApp()
        {
            Workbook.Close();
            ExcelApp.Quit();
            templateName = null;
            excelApp = null;
            worksheetSettings = null;
            worksheetNames = null;
            worksheetWishes = null;
            workbook = null;
            GC.Collect();
        }
    }
}
