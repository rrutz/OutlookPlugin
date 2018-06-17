using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmailClassifier
{
    partial class ML_Form
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

        private void Button_trainModel_Click( object sender, System.EventArgs e)
        {
            // this.button_trainModel.Text = Convert.ToString(isClicked);
        }

        bool isClicked = false;
        private void CB_addToTrainingData_Click( object sender, System.EventArgs e)
        {
            if (isClicked)
            {
                isClicked = false;
            }
            else
            {
                isClicked = true;
            }
        }

        private void button_read_Click(object sender, System.EventArgs e)
        {
            if (isClicked)
            { 
                Globals.ThisAddIn.writeToFile(this.button_read.Text);
            }
        }


    }
}
