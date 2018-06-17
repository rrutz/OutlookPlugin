using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace EmailClassifier
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        public void writeToFile(string text)
        {
            String path = @"C:\Users\Ruedi\OneDrive\OutlookPlugin\OutlookPlugin\EmailClassifier\test.txt";

            
            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selObject = this.Application.ActiveExplorer().Selection[1];
                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                    String itemMessage = mailItem.Subject;
                    mailItem.Display(true);
                    if (!System.IO.File.Exists(path))
                    {
                        using (System.IO.StreamWriter sw = System.IO.File.AppendText(path))
                        { 
                            sw.WriteLine("Sender, Header, Receiver");
                            sw.WriteLine(itemMessage);
                        }
                    }
                    else
                    {
                        using (System.IO.StreamWriter sw = System.IO.File.AppendText(path))
                        {
                            sw.WriteLine(itemMessage);
                        }
                    }
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
