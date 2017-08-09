using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
            //foreach (Word.Range rng in doc.StoryRanges)
            //{
            app.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
            ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
            ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            //}
        }

        static void SearchReplace(string fullName)
        {
            object oMissing = System.Reflection.Missing.Value;
            Word.Application app;
            Word.Document doc;
            app = new Word.Application();
            doc = app.Documents.Open(fullName);
            //string MixedName = Path.GetFileNameWithoutExtension(fullName).Trim().Replace("（包材）", "").Replace("（原料）", "").Replace("（辅料）", "");
            //string wjmc = regWJBM.Match(MixedName).Groups[1].Value;
            //string wjbm = regWJBM.Match(MixedName).Groups[2].Value;

            FindAndReplace(app, doc, "号", "**号**");
            //FindAndReplace(app, doc, "[WJBM]", wjbm);

            //foreach (Word.Section section in doc.Sections)
            //{
            //    object missing = null;
            //    object replaceAll = Word.WdReplace.wdReplaceAll;


            //    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            //    footerRange.Find.Text = "[WJMC]";
            //    footerRange.Find.Replacement.Text = wjmc;
            //    footerRange.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceAll, ref missing, ref missing, ref missing, ref missing);

            //}

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

        static void ExtractNoDic2File()
        {
            string rawFileName = @"C:\DocReplace\Template\nodicraw.txt";
            //Regex reg = new Regex(@"\s(.{5,50})\s(.{5,15})\s(.{5,15})");
            Regex reg = new Regex(@"\s(.{5,50})\s([A-Z]{2}.{5,15}\d\d)\s([A-Z]{2}.{5,15}\d\d)");
            //foreach (string line in File.ReadAllLines(rawFileName, Encoding.Unicode))
            using (StreamWriter sw = new StreamWriter(@"C:\DocReplace\Template\nodic.txt", false, Encoding.Unicode))
                foreach (string line in File.ReadAllLines(rawFileName))
                {
                    if (reg.IsMatch(line))
                    {
                        Match m = reg.Match(line);
                        MedStandard meds = new MedStandard(m.Groups[1].Value, m.Groups[2].Value, m.Groups[3].Value);
                        Console.WriteLine(meds);
                        sw.WriteLine(meds);
                    }
                }
        }


        static List<MedStandard> lstMedStandards = new List<MedStandard>();

        static void ReadLstMedStandards()
        {
            lstMedStandards = File.ReadAllLines(@"C:\DocReplace\Template\nodic.txt", Encoding.Unicode)
                .Select(x => x.Split(new char[] { '\t' }, StringSplitOptions.RemoveEmptyEntries))
                .Select(x => new MedStandard(x[0], x[1], x[2])).ToList();
        }

        static void Main(string[] args)
        {
            //ExtractNoDic2File();
            ReadLstMedStandards();
            //foreach (MedStandard meds in lstMedStandards)
            //    Console.WriteLine(meds);

            //foreach (string docFullName in Directory.GetFiles(@"C:\DocReplace\Old", "*.doc"))
            //foreach (string docFullName in Directory.GetFiles(@"C:\DocReplace\ReplaceHeaderFooter", "*.doc",SearchOption.AllDirectories))
            //{
            //    SaveAsText(docFullName);
            //}
            Regex regTxtNo = new Regex(@"[A-Z]{2}/.{3,15}\d\d", RegexOptions.Compiled);
            foreach (string docFullName in Directory.GetFiles(@"C:\DocReplace\ReplaceHeaderFooter", "*.doc", SearchOption.AllDirectories))
            {
                Console.WriteLine(docFullName);
                string MixedName = Path.GetFileNameWithoutExtension(docFullName).Trim().Replace("（包材）", "").Replace("（原料）", "").Replace("（辅料）", "");
                //string wjmc = regWJBM.Match(MixedName).Groups[1].Value;
                string wjbm = regWJBM.Match(MixedName).Groups[2].Value;
                Console.WriteLine(wjbm);
                var meds = lstMedStandards.FirstOrDefault(x =>
                  MedStandard.EqualAfterReplaceSpecialChars(x.OldNO, wjbm));
                if (meds != null)
                {
                    Console.WriteLine($"************{meds.NewNO}");
                    string copyTo = docFullName.Replace("ReplaceHeaderFooter", "ReplaceDicNo")
                        .Replace(".doc", $"新编号{meds.NewNO.Replace("/", "／")}.doc");
                    File.Copy(docFullName, copyTo);
                    //.Replace("/", "斜杠")
                }

                //string content = File.ReadAllText(txtFullName);
                //if (regTxtNo.IsMatch(content))
                //{
                //    foreach (Match m in regTxtNo.Matches(content))
                //    {
                //        string oldNo = m.Value;
                //        Console.WriteLine(oldNo);
                //        var meds = lstMedStandards.FirstOrDefault(x =>
                //          MedStandard.EqualAfterReplaceSpecialChars(x.OldNO, oldNo));
                //        if (meds != null)
                //            Console.WriteLine($"************{meds.NewNO}");
                //    }
                //}
            }
            Console.WriteLine("ok");
            Console.ReadLine();



            //foreach (string docFullName in Directory.GetFiles(@"C:\DocReplace\TestDes", "*.doc"))
            //{
            //    SearchReplace(docFullName);
            //}
        }
    }
}
