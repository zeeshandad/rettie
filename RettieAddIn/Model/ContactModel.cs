using RettieAddIn.Helper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RettieAddIn.Model
{
    public class ContactModel : INotifyPropertyChanged
    {
        public delegate void DirtyFlagChangedEventHandler(object sender, EventArgs e);
        public event DirtyFlagChangedEventHandler DirtyFlagChanged;

        // Invoke the Changed event; called whenever list changes
        protected virtual void OnDirtyFlagChanged(EventArgs e)
        {
            if (DirtyFlagChanged != null)
                DirtyFlagChanged(this, e);
        }

        List<RettieContact> _ResponsibilityList;
        public List<RettieContact> ResponsibilityList
        {

            get
            {
                if (_ResponsibilityList == null)
                    _ResponsibilityList = ContactsHelper.GetListOfContacts();

                return _ResponsibilityList;
            }
            set
            {
                if (value != _ResponsibilityList)
                {
                    _ResponsibilityList = value;
                    NotifyPropertyChanged("ResponsibilityList");
                    OnDirtyFlagChanged(new EventArgs());
                }
            }
        }

        RettieContact _Responsibility;
        public RettieContact Responsibility
        {

            get { return _Responsibility; }
            set
            {
                if (value != _Responsibility)
                {
                    _Responsibility = value;
                    NotifyPropertyChanged("Responsibility");
                    OnDirtyFlagChanged(new EventArgs());
                }

            }
        }

        DateTime? _LastCheckedDate;
        public DateTime? LastCheckedDate
        {
            get { return _LastCheckedDate; }
            set
            {
                if (_LastCheckedDate != value)
                {
                    _LastCheckedDate = value;
                    NotifyPropertyChanged("LastCheckedDate");
                    //OnDirtyFlagChanged(new EventArgs());
                }
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged(String info)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(info));
            }
        }
    }
}
