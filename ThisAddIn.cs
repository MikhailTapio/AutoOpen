using System;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace AutoOpen_Word
{
    public partial class ThisAddIn
    {
        private const string tempFileName = "OpenedDocs.txt";
        private readonly List<string> closedDocs = new List<string>();

        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            string fileName = Path.Combine(Path.GetTempPath(), tempFileName);
            if (File.Exists(fileName))
            {
                string[] lines = File.ReadAllLines(fileName);
                foreach (string line in lines)
                {
                    Application.Documents.Open(line);
                }
                File.Delete(fileName);
            }
            Application.DocumentOpen += Application_DocumentOpen;
            Application.DocumentBeforeClose += Application_DocumentBeforeClose;
        }

        private void Application_DocumentOpen(Word.Document Doc)
        {
            closedDocs.Add(Doc.FullName);
        }

        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            string fileName = Path.Combine(Path.GetTempPath(), tempFileName);
            File.WriteAllLines(fileName, closedDocs);
        }
    }
}
