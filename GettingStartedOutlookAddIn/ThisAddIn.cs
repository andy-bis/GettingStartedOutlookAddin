using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace GettingStartedOutlookAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Get the Application object
            Outlook.Application application = this.Application;

            // Get the Inspectors objects
            Outlook.Inspectors inspectors = application.Inspectors;

            // Get the active Inspector
            Outlook.Inspector activeInspector = application.ActiveInspector();
            if (activeInspector != null)
            {
                // Get the active item's title when Outlook start
                MessageBox.Show("Active Inspector: " + activeInspector.Caption);
            }

            // Get the Explorers objects
            Outlook.Explorers explorers = application.Explorers;

            // Get the active Explorer object
            Outlook.Explorer activeExplorer = application.ActiveExplorer();
            if (activeExplorer != null)
            {
                // Get the active folder's title when Outlook start
                MessageBox.Show("Active Explorer: " + activeExplorer.Caption);
            }

            // Add a new Inspector to the application
            inspectors.NewInspector += 
                new Outlook.InspectorsEvents_NewInspectorEventHandler(
                    Inspectors_AddTextToNewMail);

            // Subscribe to the ItemSend event, that it's triggered when an email is sent
            application.ItemSend += 
                new Outlook.ApplicationEvents_11_ItemSendEventHandler(
                    ItemSend_BeforeSend);

            // Add a new Inspector to the application
            inspectors.NewInspector += 
                new Outlook.InspectorsEvents_NewInspectorEventHandler(
                    Inspectors_RegisterEventWordDocument);
        }
        
        void ItemSend_BeforeSend(object item, ref bool cancel)
        {
            Outlook.MailItem mailItem = (Outlook.MailItem) item;
            if (mailItem != null)
            {
                mailItem.Body += "Modified by GettingStartedOutlookAddIn";
            }
            cancel = false;
        }

        void Inspectors_AddTextToNewMail(Outlook.Inspector inspector)
        {
            // Get the current item for this Inspecto object and check if is type
            // of MailItem
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;            
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "My subject text";
                    mailItem.Body = "My body text";
                }
            }
        }

        void Inspectors_RegisterEventWordDocument(Outlook.Inspector inspector)
        {
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                // Check that the email editor is Word editor
                // Although "always" is a Word editor in Outlook 2013, it's best done perform this check
                if (inspector.EditorType == Outlook.OlEditorType.olEditorWord && inspector.IsWordMail())
                {
                    // Get the Word document
                    Word.Document document = inspector.WordEditor;
                    if (document != null)
                    {
                        // Subscribe to the BeforeDoubleClick event of the Word document
                        document.Application.WindowBeforeDoubleClick += 
                            new Word.ApplicationEvents4_WindowBeforeDoubleClickEventHandler(
                                ApplicationOnWindowBeforeDoubleClick);
                    }
                }
            }
        }

        private void ApplicationOnWindowBeforeDoubleClick(Word.Selection selection, ref bool cancel)
        {
            // Get the selected word
            Word.Words words = selection.Words;
            MessageBox.Show("Selection: " + words.First.Text);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
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
