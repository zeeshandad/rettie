using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace RettieAddIn.Helper
{
    public class ContactsHelper
    {
        // List<RettieContact> rettieContactList = null;

        internal static Outlook._Application Application
        {
            get;
            set;
        }

        public static List<RettieContact> RettieContactList { get; set; }

        public static List<RettieContact> GetListOfContacts()
        {
            if (RettieContactList != null)
                return RettieContactList;

            RettieContactList = new List<RettieContact>();

            Task.Factory.StartNew(() =>
           {

               AddressList addressList = null;
               try
               {

                   addressList = Globals.ThisAddIn.Application.Session.GetGlobalAddressList();//.OfType<OlAddressListType.olExchangeContainer.GetType()>;

                   //foreach (AddressList lst in lists)
                   //{

                   //    Debug.WriteLine(lst.AddressListType.ToString());
                   //    Debug.WriteLine(lst.AddressEntries.Count.ToString());
                   //    if (lst.AddressListType == OlAddressListType.olExchangeContainer)
                   //    {
                   //        //addressList = lst;
                   //        //break;

                   //        foreach (AddressEntry a in lst.AddressEntries)
                   //        {
                   //            Debug.WriteLine(a.Name);
                   //            Debug.WriteLine(a.Address);
                   //            Debug.WriteLine(a.AddressEntryUserType);
                   //        }
                   //    }

                   //}

                   if (addressList != null)
                       for (int iI = 1; iI < addressList.AddressEntries.Count; iI++)
                       {
                           AddressEntry addressEntry = addressList.AddressEntries[iI];

                           if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry
                               || addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                           {
                               Outlook.ExchangeUser exchUser = addressEntry.GetExchangeUser();
                               //Debug.WriteLine(exchUser.Name + " " + exchUser.PrimarySmtpAddress);
                               RettieContactList.Add(new RettieContact() { Name = exchUser.Name, Email = exchUser.PrimarySmtpAddress });
                               //if (addressEntry.Name == DisplayName)
                               //rettieContactList.Add(new RettieContact() { Name = addressEntry.Name, Email = addressEntry.Address });
                           }
                       }
               }
               catch
               {
                   //            GetGlobalAddressList supports only Exchange servers. It returns an error if the Global Address List is not available or cannot be found.
                   //It also returns an error if no connection is available or the user is set to work offline.
               }
               finally
               {
                   if (addressList != null)
                       Marshal.ReleaseComObject(addressList);
               }
           });
            return RettieContactList;


            //List<RettieContact> rettieContactList = null;
            ////List<Outlook.ContactItem> contactItemsList = null;
            //Outlook.Items folderItems = null;
            //Outlook.MAPIFolder folderSuggestedContacts = null;
            //Outlook.NameSpace ns = null;
            //Outlook.MAPIFolder folderContacts = null;
            //object itemObj = null;
            //try
            //{
            //    //contactItemsList = new List<Outlook.ContactItem>();
            //    rettieContactList = new List<RettieContact>();

            //    ns = Application.GetNamespace("MAPI");
            //    // getting items from the Contacts folder in Outlook
            //    folderContacts = ns.GetDefaultFolder(Outlook.ol);
            //    folderItems = folderContacts.Items;
            //    for (int i = 1; folderItems.Count >= i; i++)
            //    {
            //        itemObj = folderItems[i];
            //        if (itemObj is Outlook.ContactItem)
            //            rettieContactList.Add(new RettieContact()
            //            {
            //                Name = (itemObj as Outlook.ContactItem).FullName,
            //                Email = (itemObj as Outlook.ContactItem).Email1Address
            //            });
            //        else
            //            Marshal.ReleaseComObject(itemObj);
            //    }
            //    Marshal.ReleaseComObject(folderItems);
            //    folderItems = null;
            //    //// getting items from the Suggested Contacts folder in Outlook
            //    //folderSuggestedContacts = ns.GetDefaultFolder(
            //    //                          Outlook.OlDefaultFolders.olFolderSuggestedContacts);
            //    //folderItems = folderSuggestedContacts.Items;
            //    //for (int i = 1; folderItems.Count >= i; i++)
            //    //{
            //    //    itemObj = folderItems[i];
            //    //    if (itemObj is Outlook.ContactItem)
            //    //        contactItemsList.Add(itemObj as Outlook.ContactItem);
            //    //    else
            //    //        Marshal.ReleaseComObject(itemObj);
            //    //}
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show(ex.Message);
            //}
            //finally
            //{
            //    if (folderItems != null)
            //        Marshal.ReleaseComObject(folderItems);
            //    if (folderContacts != null)
            //        Marshal.ReleaseComObject(folderContacts);
            //    if (folderSuggestedContacts != null)
            //        Marshal.ReleaseComObject(folderSuggestedContacts);
            //    if (ns != null)
            //        Marshal.ReleaseComObject(ns);
            //}
            //return rettieContactList;
        }

        public static async Task<List<RettieContact>> GetListOfContactsAsync()
        {
            if (RettieContactList != null)
                return RettieContactList;

            RettieContactList = new List<RettieContact>();

            await Task.Factory.StartNew(() =>
                {

                    AddressList addressList = null;
                    try
                    {

                        addressList = Globals.ThisAddIn.Application.Session.GetGlobalAddressList();//.OfType<OlAddressListType.olExchangeContainer.GetType()>;

                        //foreach (AddressList lst in lists)
                        //{

                        //    Debug.WriteLine(lst.AddressListType.ToString());
                        //    Debug.WriteLine(lst.AddressEntries.Count.ToString());
                        //    if (lst.AddressListType == OlAddressListType.olExchangeContainer)
                        //    {
                        //        //addressList = lst;
                        //        //break;

                        //        foreach (AddressEntry a in lst.AddressEntries)
                        //        {
                        //            Debug.WriteLine(a.Name);
                        //            Debug.WriteLine(a.Address);
                        //            Debug.WriteLine(a.AddressEntryUserType);
                        //        }
                        //    }

                        //}

                        if (addressList != null)
                            for (int iI = 1; iI < addressList.AddressEntries.Count; iI++)
                            {
                                AddressEntry addressEntry = addressList.AddressEntries[iI];

                                if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry
                                    || addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                                {
                                    Outlook.ExchangeUser exchUser = addressEntry.GetExchangeUser();
                                    //Debug.WriteLine(exchUser.Name + " " + exchUser.PrimarySmtpAddress);
                                    RettieContactList.Add(new RettieContact() { Name = exchUser.Name, Email = exchUser.PrimarySmtpAddress });
                                    //if (addressEntry.Name == DisplayName)
                                    //rettieContactList.Add(new RettieContact() { Name = addressEntry.Name, Email = addressEntry.Address });
                                }
                            }
                    }
                    catch
                    {
                        //            GetGlobalAddressList supports only Exchange servers. It returns an error if the Global Address List is not available or cannot be found.
                        //It also returns an error if no connection is available or the user is set to work offline.
                    }
                    finally
                    {
                        if (addressList != null)
                            Marshal.ReleaseComObject(addressList);
                    }
                });
            return RettieContactList;


            //List<RettieContact> rettieContactList = null;
            ////List<Outlook.ContactItem> contactItemsList = null;
            //Outlook.Items folderItems = null;
            //Outlook.MAPIFolder folderSuggestedContacts = null;
            //Outlook.NameSpace ns = null;
            //Outlook.MAPIFolder folderContacts = null;
            //object itemObj = null;
            //try
            //{
            //    //contactItemsList = new List<Outlook.ContactItem>();
            //    rettieContactList = new List<RettieContact>();

            //    ns = Application.GetNamespace("MAPI");
            //    // getting items from the Contacts folder in Outlook
            //    folderContacts = ns.GetDefaultFolder(Outlook.ol);
            //    folderItems = folderContacts.Items;
            //    for (int i = 1; folderItems.Count >= i; i++)
            //    {
            //        itemObj = folderItems[i];
            //        if (itemObj is Outlook.ContactItem)
            //            rettieContactList.Add(new RettieContact()
            //            {
            //                Name = (itemObj as Outlook.ContactItem).FullName,
            //                Email = (itemObj as Outlook.ContactItem).Email1Address
            //            });
            //        else
            //            Marshal.ReleaseComObject(itemObj);
            //    }
            //    Marshal.ReleaseComObject(folderItems);
            //    folderItems = null;
            //    //// getting items from the Suggested Contacts folder in Outlook
            //    //folderSuggestedContacts = ns.GetDefaultFolder(
            //    //                          Outlook.OlDefaultFolders.olFolderSuggestedContacts);
            //    //folderItems = folderSuggestedContacts.Items;
            //    //for (int i = 1; folderItems.Count >= i; i++)
            //    //{
            //    //    itemObj = folderItems[i];
            //    //    if (itemObj is Outlook.ContactItem)
            //    //        contactItemsList.Add(itemObj as Outlook.ContactItem);
            //    //    else
            //    //        Marshal.ReleaseComObject(itemObj);
            //    //}
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show(ex.Message);
            //}
            //finally
            //{
            //    if (folderItems != null)
            //        Marshal.ReleaseComObject(folderItems);
            //    if (folderContacts != null)
            //        Marshal.ReleaseComObject(folderContacts);
            //    if (folderSuggestedContacts != null)
            //        Marshal.ReleaseComObject(folderSuggestedContacts);
            //    if (ns != null)
            //        Marshal.ReleaseComObject(ns);
            //}
            //return rettieContactList;
        }
    }
}
