using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace AutoOpen_Word
{
    public partial class ThisAddIn
    {
        private List<string> closedDocs = new List<string>();

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string fileName = Path.Combine(Path.GetTempPath(), "OpenedDocuments.txt");
            if (File.Exists(fileName))
            {
                string[] lines = File.ReadAllLines(fileName);
                foreach (string line in lines)
                {
                    Application.Documents.Open(line);
                }
                File.Delete(fileName);
            }
            this.Application.DocumentOpen += Application_DocumentOpen;
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
        }

        private void Application_DocumentOpen(Word.Document Doc)
        {
            closedDocs.Add(Doc.FullName);
        }

        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            string fileName = Path.Combine(Path.GetTempPath(), "OpenedDocuments.txt");
            File.WriteAllLines(fileName, closedDocs);
        }
    }
}
