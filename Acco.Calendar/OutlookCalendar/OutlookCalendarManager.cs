//
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Threading.Tasks;
//
using Microsoft.Office.Interop.Outlook;
//
using Acco.Calendar.Person;
using Acco.Calendar.Event;
using Acco.Calendar.Location;
using Acco.Calendar.Database;
using MongoDB.Driver.Builders;
using System.Text.RegularExpressions;
using Acco.Calendar.Utilities;
//

namespace Acco.Calendar.Manager
{
    public sealed class OutlookCalendarManager : ICalendarManager
    {
        private Application OutlookApplication { get; set; }
        private NameSpace MapiNameSpace { get; set; }
        private MAPIFolder CalendarFolder { get; set; }
        private static readonly OutlookCalendarManager instance = new OutlookCalendarManager();
        // hidden constructor
        private OutlookCalendarManager() { Initialize(); }
        public static OutlookCalendarManager Instance { get { return instance; } }

        private void Initialize()
        {
            OutlookApplication = new Application();
            MapiNameSpace = OutlookApplication.GetNamespace("MAPI");
            CalendarFolder = MapiNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
        }

        public bool Push(ICalendar calendar)
        {
            bool result = true;
            // TODO: set various infos here
            //
            foreach(var evt in calendar.Events)
            {

            }
            //
            return result;
        }

        public ICalendar Pull()
        {
            return Pull(from:   DateTime.Now.Add(new TimeSpan(-30 /*days*/, 0 /* hours */, 0 /*minutes*/, 0 /* seconds*/)), 
                        to:     DateTime.Now.Add(new TimeSpan(30 /*days*/, 0 /* hours */, 0 /*minutes*/, 0 /* seconds*/)));
        }

        private void Events_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            // todo: work in progress
            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    // note: to know which item was added, use NewItems.
                    foreach (GenericEvent item in e.NewItems) //todo: check if its possible to add the list of added events
                    {
                        var r = Storage.Instance.Appointments.Save(item);
                        if (!r.Ok)
                        {
                            Console.BackgroundColor = ConsoleColor.Red; // add these in utils.
                            Console.ForegroundColor = ConsoleColor.White;  // add these in utils. (Utilities.Warning(...) - Utilities.Error(...) - Utilities.Info(...)
                            Console.WriteLine("Event [{0}] was not added", item.Id);
                            Console.ResetColor();
                        }
                        else
                        {
                            Console.BackgroundColor = ConsoleColor.Green; // add these in utils.
                            Console.ForegroundColor = ConsoleColor.Black;  // add these in utils. (Utilities.Warning(...) - Utilities.Error(...) - Utilities.Info(...)
                            Console.WriteLine("Event [{0}] added", item.Id);
                            Console.ResetColor();
                        }
                    }
                    break;
                case NotifyCollectionChangedAction.Remove:
                    foreach (GenericEvent item in e.OldItems) //todo: check if its possible to delete the list of removed events
                    {
                        Console.WriteLine("Event [{0}] removed", item.Id);
                        var query = Query<GenericEvent>.EQ(evt => evt.Id, item.Id);
                        Storage.Instance.Appointments.Remove(query);
                    }
                    break;
                default:
                    throw new System.Exception("Unmanaged Action => " + e.Action);
            }
        }

        private List<string> RecipientsEmailAddresses(AppointmentItem item)
        {
            string PR_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            var emails = new List<string>();
            ///
            if (item == null)
            {
                throw new ArgumentNullException();
            }
            //
            foreach (Microsoft.Office.Interop.Outlook.Recipient recipient in item.Recipients)
            {
                var recipientAddressEntry = recipient.AddressEntry;
                if (recipientAddressEntry.Type == "EX")
                {
                    if (recipientAddressEntry != null)
                    {
                        //Now we have an AddressEntry representing the Sender
                        if (recipientAddressEntry.AddressEntryUserType == Microsoft.Office.Interop.Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                            recipientAddressEntry.AddressEntryUserType == Microsoft.Office.Interop.Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                        {
                            //Use the ExchangeUser object PrimarySMTPAddress
                            var exchUser = recipientAddressEntry.GetExchangeUser();
                            if (exchUser != null)
                            {
                                emails.Add(exchUser.PrimarySmtpAddress);
                            }
                            else
                            {
                                throw new System.Exception("No email found for " + recipientAddressEntry.Address);
                            }
                        }
                        else
                        {
                            emails.Add(recipientAddressEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string);
                        }
                    }
                    else
                    {
                        throw new System.Exception("No email found for " + recipientAddressEntry.Address);
                    }
                }
                else
                {
                    // try to match the address against a regex
                    Regex email = new Regex(Defines.EmailRegularExpression);
                    if (email.IsMatch(recipient.Address) && 
                        !emails.Contains(recipient.Address)) // avoid unnecessary duplicates
                    {
                        emails.Add(recipient.Address);
                    }
                }
            }
         
            return emails;
        }

        public async Task<bool> PushAsync(ICalendar calendar)
        {
            var push = Task.Factory.StartNew(() => Push(calendar));
            return await push;
        }

        public async Task<ICalendar> PullAsync()
        {
            var pull = Task.Factory.StartNew(() => Pull());
            return await pull;
        }

        public ICalendar Pull(DateTime from, DateTime to)
        {
            var myCalendar = new GenericCalendar();
            try
            {
                myCalendar.Id = CalendarFolder.EntryID;
                myCalendar.Name = CalendarFolder.Name;
#if !OLD_OFFICE_ASSEMBLY
                myCalendar.Creator = new GenericPerson();
                myCalendar.Creator.Email = CalendarFolder.Store.DisplayName; // warning: display name isn't always an email.
#endif
                //
                myCalendar.Events = new ObservableCollection<GenericEvent>();
                myCalendar.Events.CollectionChanged += Events_CollectionChanged;
                //
                Items evts = CalendarFolder.Items;
                evts.Sort("[Start]");
                var filter = "[Start] >= '"
                            + from.ToString("g")
                            + "' AND [End] <= '"
                            + to.ToString("g") + "'";
                evts = evts.Restrict(filter);
                //
                foreach (AppointmentItem evt in evts)
                {
                    //
                    var myEvt = new GenericEvent(   Id: evt.EntryID,
                                                    Summary: evt.Subject,
                                                    Description: evt.Body,
                                                    Location: new GenericLocation { Name = evt.Location });
                    // Start
                    myEvt.Start = evt.Start;
                    // End
                    if (!evt.AllDayEvent)
                    {
                        myEvt.End = evt.End;
                    }
                    //
#if !OLD_OFFICE_ASSEMBLY // this only works with office 2010
                    // Creator and organizer are the same person.
                    // Creator
                    myEvt.Creator = new GenericPerson();
                    myEvt.Creator.Email = evt.GetOrganizer().Address;
                    myEvt.Creator.Name = evt.GetOrganizer().Name;
                    // Organizer
                    myEvt.Organizer = new GenericPerson();
                    myEvt.Organizer.Email = evt.GetOrganizer().Address;
                    myEvt.Organizer.Name = evt.GetOrganizer().Name;
#endif
                    // Attendees
                    myEvt.Attendees = new List<GenericPerson>();
                    foreach(var attendeeEmail in RecipientsEmailAddresses(evt))
                    {
                        myEvt.Attendees.Add(new GenericPerson
                        {
                            Email = attendeeEmail
                        });
                    }
                    // Recurrence
                    if (evt.IsRecurring)
                    {
                        myEvt.Recurrence = new OutlookRecurrence();
                        ((OutlookRecurrence)myEvt.Recurrence).Parse(evt.GetRecurrencePattern());
                    }
                    // add it to calendar events.
                    myCalendar.Events.Add(myEvt);
                }
                //
                myCalendar.Events.CollectionChanged -= Events_CollectionChanged;
            }
            catch(System.Exception ex)
            {
                Console.WriteLine("Exception: [{0}]", ex.Message);
            }
            //
            return myCalendar;
        }

        public async Task<ICalendar> PullAsync(DateTime from, DateTime to)
        {
            var pull = Task.Factory.StartNew(() => Pull(from, to));
            return await pull;
        }
    }
}
