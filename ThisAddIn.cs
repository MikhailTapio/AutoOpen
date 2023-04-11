using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.IO;

namespace AutoOpen_PPT
{
    public partial class ThisAddIn
    {
        private const string tempFileName = "OpenedPres.txt";
        private List<string> closedPres = new List<string>();
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string fileName = Path.Combine(Path.GetTempPath(), tempFileName);
            if (File.Exists(fileName))
            {
                string[] lines = File.ReadAllLines(fileName);
                foreach (string line in lines)
                {
                    Application.Presentations.Open(line);
                }
                File.Delete(fileName);
            }
            this.Application.PresentationOpen += Application_PresentationOpen;
            this.Application.PresentationClose += Application_PresentationClose;
        }

        private void Application_PresentationClose(PowerPoint.Presentation Pres)
        {
            string fileName = Path.Combine(Path.GetTempPath(), tempFileName);
            File.WriteAllLines(fileName, closedPres);
        }

        private void Application_PresentationOpen(PowerPoint.Presentation Pres)
        {
            closedPres.Add(Pres.FullName);
        }
    }
}
