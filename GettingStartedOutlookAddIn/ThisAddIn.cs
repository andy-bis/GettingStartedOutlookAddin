using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Diagnostics;
using System;
using System.Collections.Generic;

namespace GettingStartedOutlookAddIn
{
    public partial class ThisAddIn
    {
		private List<Outlook.Items> mSubscribedItems = new List<Outlook.Items>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
			/*while (!Debugger.IsAttached)
			{
				System.Threading.Thread.Sleep(1000);
			}

			Debugger.Break();*/

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
                // MessageBox.Show("Active Explorer: " + activeExplorer.Caption);
            }

			AddHandlerToStores();

            // Add a new Inspector to the application
            inspectors.NewInspector += 
                new Outlook.InspectorsEvents_NewInspectorEventHandler(
                    Inspectors_AddTextToNewMail);

            // Subscribe to the ItemSend event, that it's triggered when an email is sent
            application.ItemSend += 
                new Outlook.ApplicationEvents_11_ItemSendEventHandler(
                    ItemSend_BeforeSend);

			application.ItemLoad += this.ItemLoadHandler;

			application.NewMail += this.NewMailHandler;
			application.NewMailEx += this.NewMailExHandler;

            // Add a new Inspector to the application
            inspectors.NewInspector += 
                new Outlook.InspectorsEvents_NewInspectorEventHandler(
                    Inspectors_RegisterEventWordDocument);
        }

		private void AddHandlerToStores()
		{
			List<string> names = new List<string>();

			foreach (Outlook.Store store in this.Application.Session.Stores)
			{
				AddHandlerToFolder(store.GetRootFolder(), names);

				//Outlook.MAPIFolder inbox = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
				//inbox.Items.ItemAdd += this.InboxItemAdd;

				// MessageBox.Show(string.Format("Store {0}", store.DisplayName));
				/*Outlook.MAPIFolder fldr = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

				foreach (Outlook.MailItem item in fldr.Items)
				{
					Outlook.UserProperty ss = item.UserProperties.Find("SecretSubj");
					//if (ss == null)
					//{
					//	Outlook.UserProperty nup = item.UserProperties.Add("SecretSubj", Outlook.OlUserPropertyType.olText, false);
					//	nup.Value = item.Subject;
					//	item.Subject = "public subj";
					//}

					Outlook.UserProperty sb = item.UserProperties.Find("SecretBody");
					//if (sb == null)
					//{
					//	Outlook.UserProperty nup = item.UserProperties.Add("SecretBody", Outlook.OlUserPropertyType.olText, false);
					//	nup.Value = item.Body;
					//	item.Body = "public text";
					//}

					if (ss != null && sb != null)
					{
						MessageBox.Show(string.Format("secrets: {0}\n{1}", ss.Value, sb.Value));
					}
				}*/
			}
		}

		private void AddHandlerToFolder(Outlook.MAPIFolder folder, List<string> names)
		{
			Outlook.Items items = folder.Items;
			
			items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(
				GetFolderItemAddHandler(folder.Name));

			mSubscribedItems.Add(items);
			names.Add(folder.Name);

			foreach (Outlook.MAPIFolder subfolder in folder.Folders)
			{
				AddHandlerToFolder(subfolder, names);
			}
		}

		private void InboxItemAdd(object item)
		{
			Outlook.MailItem mailItem = (Outlook.MailItem)item;
			if (mailItem != null)
			{
				string message = string.Format("Added email: {0}", mailItem.Subject);
				MessageBox.Show(message);
			}
			else
			{
				string message = string.Format("Added: {0}", item.GetType().Name);
				MessageBox.Show(message);
			}
		}

		private Action<object> GetFolderItemAddHandler(string folderName)
		{
			return (item) => FolderItemAddHandler(folderName, item);
		}

		private void FolderItemAddHandler(string folderName, object item)
		{
			if (folderName == "Outbox" || folderName == "Sent Items")
			{
				return;
			}

			Outlook.MailItem mailItem = (Outlook.MailItem)item;
			if (mailItem != null)
			{
				Outlook.UserProperty hd = mailItem.UserProperties.Find("secret");
				if (hd != null)
				{
					string message = string.Format("Added email with {0} to {1}", hd.Value, folderName);
					MessageBox.Show(message);
				}
				else
				{
					string message = string.Format("Added email: {0} to {1}", mailItem.Subject, folderName);
					MessageBox.Show(message);
				}
			}
			else
			{
				string message = string.Format("Added {0} to {1}", item.GetType().Name, folderName);
				MessageBox.Show(message);
			}
		}

		private void NewMailHandler()
		{
			System.Threading.Thread.Sleep(1000);
			//MessageBox.Show("NewMail");
		}

		private void NewMailExHandler(string EntryIDCollection)
		{
			System.Threading.Thread.Sleep(1000);
			//string message = string.Format("NewMailEx: {0}", EntryIDCollection);
			//MessageBox.Show(message);
		}

		private void ItemLoadHandler(object item)
		{
			/*Outlook.MailItem mailItem = (Outlook.MailItem)item;
			if (mailItem != null)
			{
				string message = string.Format("Loaded email: {0}", mailItem.Subject);
				MessageBox.Show(message);
			}
			else
			{
				string message = string.Format("Loaded: {0}", item.GetType().Name);
				MessageBox.Show(message);
			}*/
		}

		void ItemSend_BeforeSend(object item, ref bool cancel)
        {
			try
			{
				Outlook.MailItem mailItem = (Outlook.MailItem)item;
				if (mailItem != null)
				{
					mailItem.Body += "Modified by GettingStartedOutlookAddIn";
					// mailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/8582000B", true);
					mailItem.UserProperties.Add("secret", Outlook.OlUserPropertyType.olText).Value = "encrypted";
				}
				cancel = false;
			}
			catch (Exception ex)
			{
			}
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
