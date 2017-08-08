using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            string text = doc.Content.Text;
            string txtFileName = fullName.Replace(@".doc", @".txt");
            File.WriteAllText(txtFileName, text);
            app.Quit();
            Console.WriteLine(txtFileName);
        }

        static void Main(string[] args)
        {
            foreach (string docFullName in Directory.GetFiles(@"C:\DocReplace\Old", "*.doc"))
            {
                SaveAsText(docFullName);
            }
            Console.WriteLine("ok");
            Console.ReadLine();
        }
    }
}
