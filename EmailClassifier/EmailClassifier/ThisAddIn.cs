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

        Outlook.Explorer thisExplorer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            thisExplorer = this.Application.ActiveExplorer();
            thisExplorer.SelectionChange += new Microsoft.Office.Interop.Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Access_All_Form_Regions);
 
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        public void writeToFile(string text, bool writeToFile)
        {
            String path = @"C:\Users\Ruedi\OneDrive\OutlookPlugin\OutlookPlugin\EmailClassifier\test.txt";
            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selObject = this.Application.ActiveExplorer().Selection[1];
                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                    String itemMessage = mailItem.Subject;

                    if ( writeToFile )
                    {
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
                    // call R function
                }
            }
        }



        private void Access_All_Form_Regions()
        {
            foreach (Microsoft.Office.Tools.Outlook.IFormRegion formRegion in Globals.FormRegions)
            {
                if (formRegion is ML_Form)
                {
                    ML_Form formRegion1 = (ML_Form)formRegion;
                    formRegion1.button_read.Text = "wwwwwwww";
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
