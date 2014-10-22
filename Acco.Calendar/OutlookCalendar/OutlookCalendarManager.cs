
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
using System.Linq;

//

namespace Acco.Calendar.Manager
{
    public sealed class OutlookCalendarManager : GenericCalendarManager
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

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
            Log.Info("Initializing...");
            OutlookApplication = new Application();
            MapiNameSpace = OutlookApplication.GetNamespace("MAPI");
            CalendarFolder = MapiNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
        }

        public override IEnumerable<UpdateOutcome> Push(ICalendar calendar)
        {
            Log.Info(String.Format("Pushing calendar [{0}] to outlook", calendar.Id));
            var result = new List<UpdateOutcome>();
            foreach (var evt in calendar.Events)
            {
            }
            return result;
        }

        public override ICalendar Pull()
        {
            Log.Info("Pulling calendar from outlook");
            return Pull(from: DateTime.Now.Add(new TimeSpan(-30 /*days*/, 0 /* hours */, 0 /*minutes*/, 0 /* seconds*/)),
                        to: DateTime.Now.Add(new TimeSpan(30 /*days*/, 0 /* hours */, 0 /*minutes*/, 0 /* seconds*/)));
        }

        private static List<GenericAttendee> ExtractRecipientInfos(AppointmentItem item)
        {
            Log.Info("Extracting recipients infos");
            const string prSmtpAddress = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            var people = new List<GenericAttendee>();
            if (item == null)
            {
                throw new ArgumentNullException();
            }
            //
            foreach (Recipient recipient in item.Recipients)
            {
                var person = new GenericAttendee
                {
                    Name = recipient.Name,
                    Response = ResponseStatus.None
                };
                //
                var recipientAddressEntry = recipient.AddressEntry;
                // find attedee's email
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
                // find attendee's response to the meeting 
                // todo: check why OlResponseStatus is always olResponseNone
                Log.Debug(String.Format("[{0}] response status is [{1}]", person.Email, recipient.MeetingResponseStatus));
                switch (recipient.MeetingResponseStatus)
                {
                    case OlResponseStatus.olResponseAccepted:
                        person.Response = ResponseStatus.Accepted;
                        break;
                    case OlResponseStatus.olResponseTentative:
                        person.Response = ResponseStatus.Tentative;
                        break;
                    case OlResponseStatus.olResponseDeclined:
                        person.Response = ResponseStatus.Declined;
                        break;
                    case OlResponseStatus.olResponseOrganized:
                        person.Response = ResponseStatus.Organized;
                        break;
                    case OlResponseStatus.olResponseNone:
                        person.Response = ResponseStatus.None;
                        break;
                    case OlResponseStatus.olResponseNotResponded:
                        person.Response = ResponseStatus.NotResponded;
                        break;
                }
                // free busy
                // todo: finish parsing freebusy info
                //var fb = recipient.FreeBusy(DateTime.Now /* what date do we pass here? */, 30 /* minutes */); // every "1" means that 30 minutes are busy
                //
                if (people.All(p => p.Name != person.Name) && people.All(e => e.Email != person.Email))
                {
                    Log.Info(String.Format("Adding [{0}] to people", person.Email));
                    people.Add(person);
                }
            }
            //
            return people;
        }

        public override async Task<IEnumerable<UpdateOutcome>> PushAsync(ICalendar calendar)
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
            Log.Info(String.Format("Pulling calendar from outlook, from[{0}] to [{1}]", from, to));
            var myCalendar = new GenericCalendar
            {
                Events = new DbCollection<GenericEvent>()
            };

            myCalendar.Id = CalendarFolder.EntryID;
            myCalendar.Name = CalendarFolder.Name;
#if !OLD_OFFICE_ASSEMBLY
            myCalendar.Creator = new GenericPerson
            {
                Email = CalendarFolder.Store.DisplayName
            };
#endif
            //
            var items = CalendarFolder.Items;
            items.Sort("[Start]");
            var filter = "[Start] >= '"
                        + from.ToString("g")
                        + "' AND [End] <= '"
                        + to.ToString("g") + "'";
            Log.Debug(String.Format("Filter string [{0}]", filter));
            items = items.Restrict(filter);
            //
            foreach (AppointmentItem evt in items)
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
            LastCalendar = myCalendar;
            return myCalendar;
        }

        public override async Task<ICalendar> PullAsync(DateTime from, DateTime to)
        {
            var pull = Task.Factory.StartNew(() => Pull(from, to));
            return await pull;
        }

        private async void RemoveEvents(IEnumerable<GenericEvent> eventsToRemove)
        {
            //todo: do something here..
        }
    }
}