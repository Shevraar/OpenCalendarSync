//
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
//
using Microsoft.Office.Interop.Outlook;
//
using Acco.Calendar.Person;
using Acco.Calendar.Event;
using Acco.Calendar.Location;

namespace Acco.Calendar.Manager
{

    public sealed class OutlookCalendarManager : ICalendarManager
    {
        #region Variables
        private Application OutlookApplication { get; set; }
        private NameSpace MapiNameSpace { get; set; }
        private MAPIFolder CalendarFolder { get; set; }
        #endregion

        #region Singleton directives
        private static readonly OutlookCalendarManager instance = new OutlookCalendarManager();
        // hidden constructor
        private OutlookCalendarManager() { Initialize(); }



        public static OutlookCalendarManager Instance { get { return instance; } }
        #endregion

        private void Initialize()
        {
            OutlookApplication = new Application();
            MapiNameSpace = OutlookApplication.GetNamespace("MAPI");
            CalendarFolder = MapiNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
        }

        public bool Push(GenericCalendar calendar)
        {
            bool result = true;
            // TODO: set various infos here
            //
            foreach(GenericEvent evt in calendar.Events)
            {

            }
            //
            return result;
        }

        public GenericCalendar Pull()
        {
            GenericCalendar myCalendar = new GenericCalendar();
            //
            myCalendar.Id = CalendarFolder.EntryID;
            myCalendar.Name = CalendarFolder.Name;
#if !OLD_OFFICE_ASSEMBLY
            myCalendar.Creator = new GenericPerson();
            myCalendar.Creator.Email = CalendarFolder.Store.DisplayName;
#endif
            //
            myCalendar.Events = new List<GenericEvent>();
            //
            Items evts = CalendarFolder.Items;
            //evts.IncludeRecurrences = true;
            evts.Sort("[Start]");
            string filter = "[Start] >= '"
                + DateTime.Now.Add(new TimeSpan(-30 /*days*/, 0 /* hours */, 0 /*minutes*/, 0 /* seconds*/)).ToString("g")
                + "' AND [End] <= '"
                + DateTime.Now.Add(new TimeSpan(30 /*days*/, 0 /* hours */, 0 /*minutes*/, 0 /* seconds*/)).ToString("g") + "'";
            evts = evts.Restrict(filter);
            foreach (AppointmentItem evt in evts)
            {
                //
                GenericEvent myEvt = new GenericEvent();
                // Start
                myEvt.Start = evt.Start;
                // End
                if(!evt.AllDayEvent)
                { 
                    myEvt.End = evt.End;
                }
                // Description
                myEvt.Description = evt.Subject;
                // Summary
                myEvt.Summary = evt.Body;
                // Location
                myEvt.Location = new GenericLocation();
                myEvt.Location.Name = evt.Location;
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
                // todo: get the address from exchange... another fucking thing from microsoft!
                //foreach (Microsoft.Office.Interop.Outlook.Recipient rcpt in evt.Recipients)
                //{
                //    Microsoft.Office.Interop.Outlook.AddressEntry rcptAE = rcpt.AddressEntry;
                //    Console.WriteLine("rcpt {0}", rcptAE.GetExchangeUser().PrimarySmtpAddress);
                //}
                myEvt.Attendees = new List<GenericPerson>();
                string[] requiredAttendees = null;
                if(evt.RequiredAttendees != null)
                { 
                    requiredAttendees = evt.RequiredAttendees.Split(';');
                }
                string[] optionalAttendees = null;
                if(evt.OptionalAttendees != null)
                { 
                    optionalAttendees = evt.OptionalAttendees.Split(';');
                }
                //
                if(requiredAttendees != null)
                { 
                    foreach(string attendee in requiredAttendees)
                    {
                        myEvt.Attendees.Add(new GenericPerson 
                        {
                            Email = attendee.Trim() // todo: add some validation to test if attendee contains an email or not (regex)
                        });
                    }
                }
                //
                if(optionalAttendees != null)
                { 
                    foreach(string optionalAttendee in optionalAttendees)
                    {
                        myEvt.Attendees.Add(new GenericPerson
                        {
                            Email = optionalAttendee.Trim() // todo: add some validation to test if attendee contains an email or not (regex)
                        });
                    }
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
            return myCalendar;
        }

        //private string GetSenderSMTPAddress(AppointmentItem item)
        //{
        //    string PR_SMTP_ADDRESS =
        //        @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        //    if (item == null)
        //    {
        //        throw new ArgumentNullException();
        //    }
        //    if (item.Recipients == "EX")
        //    {
        //        Outlook.AddressEntry sender =
        //            mail.Sender;
        //        if (sender != null)
        //        {
        //            //Now we have an AddressEntry representing the Sender
        //            if (sender.AddressEntryUserType ==
        //                Outlook.OlAddressEntryUserType.
        //                olExchangeUserAddressEntry
        //                || sender.AddressEntryUserType ==
        //                Outlook.OlAddressEntryUserType.
        //                olExchangeRemoteUserAddressEntry)
        //            {
        //                //Use the ExchangeUser object PrimarySMTPAddress
        //                Outlook.ExchangeUser exchUser =
        //                    sender.GetExchangeUser();
        //                if (exchUser != null)
        //                {
        //                    return exchUser.PrimarySmtpAddress;
        //                }
        //                else
        //                {
        //                    return null;
        //                }
        //            }
        //            else
        //            {
        //                return sender.PropertyAccessor.GetProperty(
        //                    PR_SMTP_ADDRESS) as string;
        //            }
        //        }
        //        else
        //        {
        //            return null;
        //        }
        //    }
        //    else
        //    {
        //        return mail.SenderEmailAddress;
        //    }
        //}

        public async Task<bool> PushAsync(GenericCalendar calendar)
        {
            Task<bool> push = Task.Factory.StartNew(() => Push(calendar));
            return await push;
        }

        public async Task<GenericCalendar> PullAsync()
        {
            Task<GenericCalendar> pull = Task.Factory.StartNew(() => Pull());
            return await pull;
        }
    }
}
