using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace Now_yap_16._01._20
{
    class WordWishesWriter
    {
        private Word.Application wordApp = new Word.Application();
        private Word.Document wDoc = null;
        public string templateName = null;
        public string fontName = null;

        private Word.Application WordApp { get { if (wordApp == null) { wordApp = new Word.Application(); loadDoc(); } return wordApp; } }
        private Word.Document WDoc { get { if (wDoc == null) loadDoc(); return wDoc; } }


        public WordWishesWriter(string templateName)
        {
            this.templateName = templateName;
            loadDoc();
        }

        public void createCongratulationsDoc(WishesGenerator wishesGenerator, List<string> names)
        {
            for (int j = 0; j < names.Count; j++)
            {
                wishesGenerator.newTrio();
                WordApp.ActiveDocument.Bookmarks["Name"].Range.Text = names[j];
                for (int i = 0; i < 3; i++)
                {
                    WordApp.ActiveDocument.Bookmarks[$"Wish{i + 1}"].Range.Text = (wishesGenerator.WishesTrio)[i];
                }
                if (j + 1 < names.Count)
                {
                    WordApp.Selection.EndKey(Word.WdUnits.wdStory);
                    WordApp.Selection.InsertNewPage();
                    WordApp.Selection.InsertFile(templateName);
                }
            }

            if (fontName != null)
            {
                WordApp.Selection.WholeStory();
                WordApp.Selection.Font.Name = fontName;
            }
        }

        public void showWord()
        {
            WordApp.Visible = true;
        }

        private void loadDoc()
        {
            wDoc = WordApp.Documents.Add(templateName);
        }
    }
}
