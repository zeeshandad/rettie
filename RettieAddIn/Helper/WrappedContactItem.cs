using Microsoft.Office.Interop.Outlook;
using RettieAddIn.Model;
using RettieAddIn.Regions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RettieAddIn.Helper
{
    class WrappedContactItem
    {
        ContactItem item;
        private const string responsibilityProperty = "Responsibility";
        private const string LastCheckedDateProperty = "LastCheckedDate";
        bool dirty = false;

        public WrappedContactItem(ContactItem c)
        {
            item = c;

            ((ItemEvents_10_Event)item).Open += WrappedContactItem_Open;
            ((ItemEvents_10_Event)item).Close += WrappedContactItem_Close;
            ((ItemEvents_10_Event)item).Write += WrappedContactItem_Write; //this is the save event
        }

        void WrappedContactItem_Close(ref bool Cancel)
        {
            if (dirty && item.Saved)
                if (MessageBox.Show("Do you want to save your changes?", "Changes", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    ((ItemEvents_10_Event)item).Write -= WrappedContactItem_Write; //this is the save event
                    SaveCustomInformation();
                }

            Marshal.ReleaseComObject(item);
        }

        void WrappedContactItem_Open(ref bool Cancel)
        {
            var inspector = item.GetInspector;
            RettieContactRegion objContactRegion = Globals.FormRegions[inspector].First() as RettieContactRegion;
            if (objContactRegion.rettieDetailsUserControl1.DataContext == null)
            {
                object resp = item.User1; //GetProperty(responsibilityProperty);
                if (resp != null)
                    ContactViewModel.Responsibility = FindContact(resp.ToString());
                object dt = item.Anniversary;//GetProperty(LastCheckedDateProperty);
                if (dt != null && DateTime.Parse(dt.ToString()) != DateTime.MinValue && DateTime.Parse(dt.ToString()).Year <= DateTime.Now.Year)
                    ContactViewModel.LastCheckedDate = DateTime.Parse(dt.ToString());
                //else
                //    ContactViewModel.LastCheckedDate = DateTime.Now;

                ContactViewModel.DirtyFlagChanged += ContactViewModel_DirtyFlagChanged;
                objContactRegion.rettieDetailsUserControl1.DataContext = ContactViewModel;
            }
        }

        private RettieContact FindContact(String email)
        {
            RettieContact contact = ContactsHelper.GetListOfContactsAsync().Result.Find(r => r.Email.ToLower() == email.ToLower());
            return contact;
        }

        void ContactViewModel_DirtyFlagChanged(object sender, EventArgs e)
        {
            //Our custom values are changed, we need to prompt user if he tries to close window without saving them
            dirty = true;
        }

        void WrappedContactItem_Write(ref bool Cancel)
        {
            SaveCustomInformation();
        }

        private void SaveCustomInformation()
        {
            dirty = false;
            if (ContactViewModel.Responsibility != null)
            {
                //SetProperty(responsibilityProperty, ContactViewModel.Responsibility, OlUserPropertyType.olText);
                //SetProperty(LastCheckedDateProperty, ContactViewModel.LastCheckedDate, OlUserPropertyType.olDateTime);
                if (!String.IsNullOrEmpty(ContactViewModel.Responsibility.Email))
                    item.User1 = ContactViewModel.Responsibility.Email;
                if (ContactViewModel.LastCheckedDate.HasValue)
                    item.Anniversary = ContactViewModel.LastCheckedDate.Value;
                item.Save();
            }
            //item.UserProperties[responsibilityProperty].Value = ContactViewModel.Responsibility;
            //item.UserProperties[LastCheckedDateProperty].Value = ContactViewModel.LastCheckedDate;
            //OlUserPropertyType.olText
            //MessageBox.Show(ContactViewModel.Responsibility);
        }

        ContactModel _ContactViewModel;
        public ContactModel ContactViewModel
        {
            get
            {
                if (_ContactViewModel == null)
                {
                    _ContactViewModel = new ContactModel();
                }
                return _ContactViewModel;
            }
        }

        private void SetProperty(string propertyName, Object value, OlUserPropertyType type)
        {
            if (item.UserProperties.Find(propertyName) == null)
            {
                item.UserProperties.Add(propertyName, type, false);
            }

            item.UserProperties[propertyName].Value = value;
        }

        private object GetProperty(string propertyName)
        {
            if (item.UserProperties.Find(propertyName) != null)
            {
                return item.UserProperties[propertyName].Value;
            }

            return null;
        }
    }
}
