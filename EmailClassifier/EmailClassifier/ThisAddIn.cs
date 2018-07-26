using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.IO;
using CsvHelper;

namespace EmailClassifier
{
    public partial class ThisAddIn
    {
        Outlook.Explorer thisExplorer;
   
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            thisExplorer = this.Application.ActiveExplorer();
            thisExplorer.SelectionChange += new Microsoft.Office.Interop.Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Access_All_Form_Regions);
        }

        // Writes email information to a csv file
        public void writeToFile(string label)
        { 
            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selObject = this.Application.ActiveExplorer().Selection[1];
                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                    String from = mailItem.SenderName;
                    String to = mailItem.To;
                    String cc = mailItem.CC;
                    String subject = mailItem.Subject;
                    String body = mailItem.Body;
                       
                    using (TextWriter writer = new StreamWriter(@"C:\Users\Ruedi\OneDrive\MS\OutlookPlugin\EmailClassifier\data2.csv", append: true))
                    {
                        var csv = new CsvWriter(writer);
                        if( Console.WriteLine(@"C:\Users\Ruedi\OneDrive\MS\OutlookPlugin\EmailClassifier\data2.csv"))
                        {
                            var list = new List<string[]>
                            {
                                new[] { "label", "from", "to", "cc", "subject", "body" },
                                new[] { label, from, to, cc, subject, body },
                            };
                        }
                        else
                        {
                            var list = new List<string[]>
                            {
                                new[] { label, from, to, cc, subject, body },
                            };
                        }
                        foreach (var item in list)
                        {
                            foreach (var field in item)
                            {
                                csv.WriteField(field);
                            }
                            csv.NextRecord();
                        }
                        writer.Flush();
                    }
                }
            }
        }

        // pulls email information and thne calls R script which returns the predicition
        public string classifyEmail(string rCodeFilePath, string rScriptExecutablePath)
        {
            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                var info = new ProcessStartInfo();
                Object selObject = this.Application.ActiveExplorer().Selection[1];
                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                    String from = mailItem.SenderName;
                    String to = mailItem.To;
                    String cc = mailItem.CC;
                    String subject = mailItem.Subject;
                    String body = mailItem.Body;
                    info.Arguments = rCodeFilePath + " " + "\"" + body + "\"";
                }

                string file = rCodeFilePath;
                string result;

                info.FileName = rScriptExecutablePath;
                info.WorkingDirectory = Path.GetDirectoryName(rScriptExecutablePath);
                info.RedirectStandardInput = false;
                info.RedirectStandardOutput = true;
                info.UseShellExecute = false;
                info.CreateNoWindow = true;
                using (var proc = new Process())
                {
                    proc.StartInfo = info;
                    proc.Start();
                    result = proc.StandardOutput.ReadToEnd();
                }
                return result;
            }
            return "error";
        }

        // do I still need this????
        private void Access_All_Form_Regions()
        {
            foreach (Microsoft.Office.Tools.Outlook.IFormRegion formRegion in Globals.FormRegions)
            {
                if (formRegion is ML_Form)
                {
                    ML_Form formRegion1 = (ML_Form)formRegion;
                }
            }
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}


