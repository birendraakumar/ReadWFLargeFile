using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadMSWF
{
    class Program
    {
        static void Main(string[] args)
        {
            //print();

        }

        static private void FindAndReplace(Microsoft.Office.Interop.Word.Application WordApp, object findText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object nmatchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
            object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            WordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
            ref nmatchAllWordForms, ref forward,
            ref wrap, ref format, ref replaceWithText,
            ref replaceAll, ref matchKashida,
            ref matchDiacritics, ref matchAlefHamza,
            ref matchControl);
        }
        static public void print()
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            //string clearancename = Form1.Texts;
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open("C:\\Users\\HP\\Desktop\\document.docx");
            Microsoft.Office.Interop.Word.Words wds = doc.Sections[1].Range.Words;
            doc.Activate();

            //Iterate the word need to change font
            foreach (Microsoft.Office.Interop.Word.Range wd in wds)
            {
                if (wd.Text.Equals("<") || wd.Text.Equals(">") || wd.Text.Equals("name"))
                    wd.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;
            }

            FindAndReplace(app, "Sad", "[Birendra Kumar]");

            //doc.PrintPreview();
            object FileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;
            doc.SaveAs2("C:\\Users\\HP\\Desktop\\Find-And-Replace-Text.pdf", ref FileFormat);
        }
    }
}
