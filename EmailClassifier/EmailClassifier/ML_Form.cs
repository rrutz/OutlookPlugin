using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.Threading;

namespace EmailClassifier
{
    public partial class ML_Form
    {

        #region Form Region Factory 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("EmailClassifier.FormRegion1")]
        public partial class FormRegion1Factory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void FormRegion1Factory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
            }
        }

        #endregion

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void FormRegion1_FormRegionShowing(object sender, System.EventArgs e)
        {
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void FormRegion1_FormRegionClosed(object sender, System.EventArgs e)
        {
        }
        private void threadFunction()
        { 
            Globals.ThisAddIn.classifyEmail(@"C:\Users\Ruedi\OneDrive\MS\OutlookPlugin\EmailClassifier\EmailClassifier\FitModel.r.r", @"C:\Program Files\Microsoft\R Client\R_SERVER\bin\Rscript.exe");
        }

        private void Button_trainModel_Click( object sender, System.EventArgs e)
        {
            Thread th = new Thread(threadFunction);
            th.Start();
        }

        private void button_read_Click(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.writeToFile(this.button_read.Text);
        }

        private void button_delete_Click(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.writeToFile(this.button_delete.Text);
        }

        private void button_ignore_Click(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.writeToFile(this.button_ignore.Text);
        }

        private void button_followUp_Click(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.writeToFile(this.button_followUp.Text);
        }

        public string ButtonText { get { return this.button_read.Text; } set { this.button_read.Text = value; } }
    }
}
