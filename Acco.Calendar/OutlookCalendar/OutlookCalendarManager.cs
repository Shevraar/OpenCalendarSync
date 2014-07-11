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

        private void Initialize()
        {
            OutlookApplication = new Application();
            MapiNameSpace = OutlookApplication.GetNamespace("MAPI");
            CalendarFolder = MapiNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
        }

        public static OutlookCalendarManager Instance { get { return instance; } }
        #endregion

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
            myCalendar.Id = CalendarFolder.Name;
            myCalendar.Creator = new GenericPerson();
            myCalendar.Creator.Name = OutlookApplication.DefaultProfileName;
            //
            myCalendar.Events = new List<GenericEvent>();
            //
            Items appointments = CalendarFolder.Items;
            foreach(AppointmentItem appointment in appointments)
            {
                GenericEvent myEvt = new GenericEvent();
                // Start
                myEvt.Start = appointment.Start;
                // End
                if(!appointment.AllDayEvent)
                { 
                    myEvt.End = appointment.End;
                }
                // Description
                myEvt.Description = appointment.Subject;
                // Location
                myEvt.Location = new GenericLocation();
                myEvt.Location.Name = appointment.Location;
                // Creator
                //todo: understand how to retrieve creator info... otherwise it'll be empty
                myEvt.Creator = new GenericPerson();
                // Attendees
                myEvt.Attendees = new List<GenericPerson>();
                string[] requiredAttendees = appointment.RequiredAttendees.Split(';');
                string[] optionalAttendees = appointment.OptionalAttendees.Split(';');
                //
                foreach(string requiredAttendee in requiredAttendees)
                {
                    myEvt.Attendees.Add(new GenericPerson 
                    {
                        Email = requiredAttendee
                    });
                }
                foreach(string optionalAttendee in optionalAttendees)
                {
                    myEvt.Attendees.Add(new GenericPerson
                    {
                        Email = optionalAttendee
                    });
                }
                // Recurrency
                if (appointment.IsRecurring)
                {
                    //RecurrencePattern rp = appointment.GetRecurrencePattern();
                    //DateTime first = new DateTime(1999, 1, 1, appointment.Start.Hour, appointment.Start.Minute, 0);
                    //DateTime last = DateTime.Now;
                    ///AppointmentItem recur = null;
                    myEvt.Recurrency = new GenericRecurrency();
                }
            }
            //
            return myCalendar;
        }

        public async Task<bool> PushAsync(GenericCalendar calendar)
        {
            return Push(calendar);
        }

        public async Task<GenericCalendar> PullAsync()
        {
            return Pull();
        }
    }
}
