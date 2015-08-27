using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using RettieAddIn.Helper;
using System.Runtime.InteropServices;

namespace RettieAddIn
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors inspectors;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            ContactsHelper.Application = this.Application;
            inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(inspectors_NewInspector);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            object item = ((Outlook.Inspector)Inspector).CurrentItem;
            ContactItem OpenedContactItem;

            if (item == null)
            {
                return;
            }
            else
            {
                if (!(item is Outlook.ContactItem))
                {
                    Marshal.ReleaseComObject(item);
                    return;
                }
            }

            // It's a Contact 
            try
            {


                OpenedContactItem = item as Outlook.ContactItem;
                WrappedContactItem objWrappedContactItem = new WrappedContactItem(OpenedContactItem);
                //((ItemEvents_10_Event)OpenedContactItem).Close += new ItemEvents_10_CloseEventHandler(Contact_Close);

                //zviContactList.Add(new ZviContact() { EContactAddress = "abc@yahoo.com", SuperID = 1, UserBucketID = 22 });
                //zviContactList.Add(new ZviContact() { EContactAddress = "abc@yahoo.com", SuperID = 1, UserBucketID = 22 });
                //zviContactList.Add(new ZviContact() { EContactAddress = "abc@yahoo.com", SuperID = 1, UserBucketID = 22 });
            }
            catch (InvalidCastException)
            {
                Marshal.ReleaseComObject(item);
                return;
            }

            //int test = myMail.HTMLBody.IndexOf(Globals.ThisAddIn.dataHandler.ACC_META_NAME);

            //Marshal.ReleaseComObject(myMail);

            //if(Inspector.CurrentItem
            //(Inspector.CurrentItem as MailItem).BeforeCheckNames += new ItemEvents_10_BeforeCheckNamesEventHandler(ThisAddIn_BeforeCheckNames);
            //MessageBox.Show("");
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
