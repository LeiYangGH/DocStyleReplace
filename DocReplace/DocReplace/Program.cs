﻿using System;
using System.IO;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
namespace DocReplace
{
    class Program
    {
        static void SaveAsText(string fullName)
        {
            object oMissing = System.Reflection.Missing.Value;
            Word._Application app;
            Word._Document doc;
            app = new Word.Application();
            doc = app.Documents.Open(fullName);
            string txtFileName = fullName.Replace(@".doc", @".txt");
            //string text = doc.Content.Text;
            //File.WriteAllText(txtFileName, text);
            object formatAsObject = Word.WdSaveFormat.wdFormatText;
            doc.SaveAs2(txtFileName, formatAsObject);
            app.Quit();
            Console.WriteLine(txtFileName);
        }

        static Regex regWJBM = new Regex(@"([^A-Z]+)([A-Z].+)", RegexOptions.Compiled);

        //https://stackoverflow.com/questions/19252252/c-sharp-word-interop-find-and-replace-everything
        static void FindAndReplace(Word.Application app, Word.Document doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            foreach (Word.Range rng in doc.StoryRanges)
            {
                app.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            }
        }

        static void SearchReplace(string fullName)
        {
            object oMissing = System.Reflection.Missing.Value;
            Word.Application app;
            Word.Document doc;
            app = new Word.Application();
            doc = app.Documents.Open(fullName);
            string MixedName = Path.GetFileNameWithoutExtension(fullName).Trim().Replace("（包材）", "").Replace("（原料）", "").Replace("（辅料）", "");
            string wjmc = regWJBM.Match(MixedName).Groups[1].Value;
            string wjbm = regWJBM.Match(MixedName).Groups[2].Value;

            //FindAndReplace(app, doc, "[WJMC]", wjmc);
            //FindAndReplace(app, doc, "[WJBM]", wjbm);

            foreach (Word.Section section in doc.Sections)
            {
                object missing = null;
                object replaceAll = Word.WdReplace.wdReplaceAll;

                //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Replace("[WJMC]", wjmc);
                //section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Replace("[WJBM]", wjbm);
                //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Replace("[WJMC]", wjmc);
                //section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text.Replace("[WJBM]", wjbm);
                Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Find.Text = "[WJMC]";
                footerRange.Find.Replacement.Text = wjmc;
                footerRange.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceAll, ref missing, ref missing, ref missing, ref missing);

            }

            //Word.Find findObject = app.Selection.Find;
            //findObject.ClearFormatting();
            //findObject.Text = "[WJMC]";
            //findObject.Replacement.ClearFormatting();
            //findObject.Replacement.Text = wjmc;

            //object replaceAll = Word.WdReplace.wdReplaceAll;
            //findObject.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            //    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            //    ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            //Word.Find findObject1 = app.Selection.Find;
            //findObject1.ClearFormatting();
            //findObject1.Text = "[WJBM]";
            //findObject1.Replacement.ClearFormatting();
            //findObject1.Replacement.Text = wjbm;

            //object replaceAll1 = Word.WdReplace.wdReplaceAll;
            //findObject1.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            //    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            //    ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            doc.Save();
            app.Quit();
        }

        static void Main(string[] args)
        {
            //foreach (string docFullName in Directory.GetFiles(@"C:\DocReplace\Old", "*.doc"))
            foreach (string docFullName in Directory.GetFiles(@"C:\DocReplace\TestSrc", "*.doc"))
            {
                SaveAsText(docFullName);
            }

            //foreach (string docFullName in Directory.GetFiles(@"C:\DocReplace\TestDes", "*.doc"))
            //{
            //    SearchReplace(docFullName);
            //}
            Console.WriteLine("ok");
            Console.ReadLine();
        }
    }
}
