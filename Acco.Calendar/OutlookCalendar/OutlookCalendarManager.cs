using System.Linq;
using Acco.Calendar.Event;
using Acco.Calendar.Location;

//
using Acco.Calendar.Person;
using Acco.Calendar.Utilities;

//
using Microsoft.Office.Interop.Outlook;

//
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

//

namespace Acco.Calendar.Manager
{
    public sealed class OutlookCalendarManager : GenericCalendarManager
    {
        private Application OutlookApplication { get; set; }

        private NameSpace MapiNameSpace { get; set; }

        private MAPIFolder CalendarFolder { get; set; }

        private static readonly OutlookCalendarManager instance = new OutlookCalendarManager();

        // hidden constructor
        private OutlookCalendarManager()
        {
            Initialize();
        }

        public static OutlookCalendarManager Instance { get { return instance; } }

        private void Initialize()
        {
            OutlookApplication = new Application();
            MapiNameSpace = OutlookApplication.GetNamespace("MAPI");
            CalendarFolder = MapiNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
        }

        public override bool Push(ICalendar calendar)
        {
            var result = true;
            // TODO: set various infos here
            //
            foreach (var evt in calendar.Events)
            {
            }
            //
            return result;
        }

        public override ICalendar Pull()
        {
            return Pull(from: DateTime.Now.Add(new TimeSpan(-30 /*days*/, 0 /* hours */, 0 /*minutes*/, 0 /* seconds*/)),
                        to: DateTime.Now.Add(new TimeSpan(30 /*days*/, 0 /* hours */, 0 /*minutes*/, 0 /* seconds*/)));
        }

        private static List<GenericPerson> ExtractRecipientInfos(AppointmentItem item)
        {
            const string prSmtpAddress = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            var people = new List<GenericPerson>();
            if (item == null)
            {
                throw new ArgumentNullException();
            }
            //
            foreach (Recipient recipient in item.Recipients)
            {
                var person = new GenericPerson
                {
                    Name = recipient.Name
                };
                //
                var recipientAddressEntry = recipient.AddressEntry;
                if (recipientAddressEntry.Type == "EX")
                {
                    if (recipientAddressEntry != null)
                    {
                        //Now we have an AddressEntry representing the Sender
                        if (recipientAddressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry ||
                            recipientAddressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                        {
                            //Use the ExchangeUser object PrimarySMTPAddress
                            var exchUser = recipientAddressEntry.GetExchangeUser();
                            if (exchUser != null)
                            {
                                person.Email = exchUser.PrimarySmtpAddress;
                            }
                            else
                            {
                                throw new System.Exception("No email found for " + recipientAddressEntry.Address);
                            }
                        }
                        else
                        {
                            person.Email = recipientAddressEntry.PropertyAccessor.GetProperty(prSmtpAddress) as string;
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
                    var email = new Regex(Defines.EmailRegularExpression);
                    if (email.IsMatch(recipient.Address)) 
                    {
                        person.Email = recipient.Address;
                    }
                }
                //
                if (people.All(p => p.Name != person.Name) && people.All(e => e.Email != person.Email))
                {
                    people.Add(person);
                }
            }
            //
            return people;
        }

        public override async Task<bool> PushAsync(ICalendar calendar)
        {
            var push = Task.Factory.StartNew(() => Push(calendar));
            return await push;
        }

        public override async Task<ICalendar> PullAsync()
        {
            var pull = Task.Factory.StartNew(() => Pull());
            return await pull;
        }

        public override ICalendar Pull(DateTime from, DateTime to)
        {
            var myCalendar = new GenericCalendar
            {
                Events = new DBCollection<GenericEvent>()
            };
            try
            {
                myCalendar.Id = CalendarFolder.EntryID;
                myCalendar.Name = CalendarFolder.Name;
#if !OLD_OFFICE_ASSEMBLY
                myCalendar.Creator = new GenericPerson
                {
                    Email = CalendarFolder.Store.DisplayName
                };
#endif
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
                    var myEvt = new GenericEvent(   id: evt.EntryID,
                                                    summary: evt.Subject,
                                                    description: evt.Body,
                                                    location: new GenericLocation { Name = evt.Location });
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
                    myEvt.Creator = new GenericPerson
                    {
                        Email = evt.GetOrganizer().Address,
                        Name = evt.GetOrganizer().Name
                    };
                    // Organizer
                    myEvt.Organizer = new GenericPerson
                    {
                        Email = evt.GetOrganizer().Address,
                        Name = evt.GetOrganizer().Name
                    };
#endif
                    // Attendees
                    myEvt.Attendees = ExtractRecipientInfos(evt);
                    // Recurrence
                    if (evt.IsRecurring)
                    {
                        myEvt.Recurrence = new OutlookRecurrence();
                        ((OutlookRecurrence)myEvt.Recurrence).Parse(evt.GetRecurrencePattern());
                    }
                    // add it to calendar events.
                    myCalendar.Events.Add(myEvt);
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Exception: [{0}]", ex.Message);
            }
            //
            return myCalendar;
        }

        public override async Task<ICalendar> PullAsync(DateTime from, DateTime to)
        {
            var pull = Task.Factory.StartNew(() => Pull(from, to));
            return await pull;
        }
    }
}