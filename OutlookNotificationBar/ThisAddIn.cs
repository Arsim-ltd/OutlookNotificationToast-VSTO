using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Tulpep.NotificationWindow;
namespace OutlookNotificationBar
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            items = inbox.Items;
            items.ItemAdd +=new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        void items_ItemAdd(object Item)
        {
            try
            {
                if (this.Application.ActiveExplorer().WindowState == Outlook.OlWindowState.olMinimized)
                {
                    Outlook.MailItem mail = (Outlook.MailItem)Item;
                    if (Item != null)
                    {
                        PopupNotifier notify = new PopupNotifier();
                        notify.TitleColor = System.Drawing.Color.White;
                        notify.TitleText = " מאת: " + mail.SenderName;
                        notify.ContentColor = System.Drawing.Color.White;
                        notify.ContentText = "כותרת: " + mail.Subject;
                        notify.IsRightToLeft = true;
                        notify.Click += (sender, e) => c_click(sender, e, mail, notify);
                        notify.Popup();
                    }
                }
            }
            catch (Exception rx) { Console.WriteLine("error during popup attempt: " + rx.Message); }

        }

        void c_click(object sender, EventArgs e, Outlook.MailItem mail,PopupNotifier notify)
        {
            mail.Display();
            notify.Hide();
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
