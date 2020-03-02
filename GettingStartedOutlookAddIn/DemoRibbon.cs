using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace GettingStartedOutlookAddIn
{
    public partial class DemoRibbon
    {
        private void DemoRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonDemo_Click(object sender, RibbonControlEventArgs e)
        {
            // Get Application application
            Outlook.Application application = Globals.ThisAddIn.Application;

            // Get the current item for this Inspecto object and check if is type
            // of MailItem
            Outlook.Inspector inspector = application.ActiveInspector();
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                MessageBox.Show("Subject: " + mailItem.Subject);
            }
        }

        private void button2Demo_Click(object sender, RibbonControlEventArgs e)
        {
            // Get Application application
            Outlook.Application application = Globals.ThisAddIn.Application;

            // Get the current item for this Inspecto object and check if is type
            // of MailItem
            Outlook.Inspector inspector = application.ActiveInspector();
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                Word.Document document = (Word.Document) inspector.WordEditor;
                string selectedText = document.Application.Selection.Text;
                MessageBox.Show(selectedText);
            }
        }
    }
}
