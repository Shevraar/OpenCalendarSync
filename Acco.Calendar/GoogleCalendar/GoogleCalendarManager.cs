using Acco.Calendar.Event;
using Acco.Calendar.Location;
using Acco.Calendar.Person;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Acco.Calendar.Manager
{
    public sealed class GoogleCalendarManager : ICalendarManager
    {
        #region Variables + Constants
        CalendarService Service { get; set; }
        UserCredential Credential { get; set; }
        FileDataStore DataStore { get; set; }
        string DataStorePath { get { return "Acco.Calendar.GoogleCalendarManager"; } }
        string AppName { get { return "Outlook 2007 Calendar Importer"; } }
        string MyCalendarId { get; set; }
        string MyCalendarName { get; set; }
        #endregion

        #region Singleton directives
        private static readonly GoogleCalendarManager instance = new GoogleCalendarManager();
        // hidden constructor
        private GoogleCalendarManager() { }

        public static GoogleCalendarManager Instance { get { return instance; } }
        #endregion

        public bool Push(GenericCalendar calendar)
        {
            // to be implemented
            return false;
        }

        public GenericCalendar Pull()
        {
            // to be implemented
            return null;
        }

        public async Task<bool> PushAsync(GenericCalendar calendar)
        {
            bool res = false;
            res = await PushEvents(calendar.Events);
            return res;
        }

        public async Task<GenericCalendar> PullAsync()
        {
            GenericCalendar calendar = new GenericCalendar();
            calendar.Events = await PullEvents();
            calendar.Id = MyCalendarId;
            calendar.Name = MyCalendarName;
            return calendar;
        }

        #region Initialization
        public async Task<bool> Initialize(string _ClientId, string _ClientSecret, string _CalendarName)
        {
            bool res = await Authenticate(_ClientId, _ClientSecret);
            //
            MyCalendarId = await DataStore.GetAsync<string>( _CalendarName + "_Id");
            MyCalendarName = _CalendarName;
            // 
            if (MyCalendarId != null)
            {
                MyCalendarId = (await GetCalendar(MyCalendarId)).Id;
            }
            //
            if (res == false || MyCalendarId == null)
            {
                MyCalendarId = (await CreateCalendar()).Id;
                await DataStore.StoreAsync<string>(_CalendarName +  "_Id", MyCalendarId);
            }
            //
            await DataStore.StoreAsync<string>( _CalendarName + "_Name", _CalendarName);
            //
            return res;
        }
        private async Task<bool> Authenticate(string _ClientId, string _ClientSecret)
        {
            bool res = true;
            try
            {
                DataStore = new FileDataStore(DataStorePath);
                Credential = await 
                    GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets
                    {
                        ClientId = _ClientId,
                        ClientSecret = _ClientSecret
                    },
                    new[] { CalendarService.Scope.Calendar },
                    "user",
                    CancellationToken.None,
                    DataStore);
                //
                Service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = Credential,
                    ApplicationName = AppName
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: {0}", ex.Message);
                res = false;
            }
            return res;
        }
        #endregion

        #region Low Level Google Calendar API v3 Operations
        private async Task<Google.Apis.Calendar.v3.Data.Calendar> GetCalendar(string id)
        {
            return await Service.Calendars.Get(id).ExecuteAsync();
        }
        private async Task<Google.Apis.Calendar.v3.Data.Calendar> CreateCalendar()
        {
            return await Service.Calendars.Insert(new Google.Apis.Calendar.v3.Data.Calendar()
            {
                Summary = MyCalendarName,
                TimeZone = "Europe/Rome",
                Description = "Automatically created: " + DateTime.Now.ToString("g")
            }).ExecuteAsync();
        }

        private async Task<bool> RemoveCalendar(string calendarId)
        {
            bool res = false;
            var deleteResult = await Service.Calendars.Delete(calendarId).ExecuteAsync();
            if (deleteResult == "")
            {
                res = true;
            }
            else
            {
                Console.WriteLine("Error while removing calendar [{0}] => [{1}]", calendarId, deleteResult);
            }
            return res;
        }

        private async Task<bool> PushEvent(GenericEvent evt)
        {
            bool res = false;
            //
            try
            {
                Google.Apis.Calendar.v3.Data.Event myEvt = new Google.Apis.Calendar.v3.Data.Event();
                // Organizer
                if (evt.Organizer != null)
                {
                    myEvt.Organizer = new Google.Apis.Calendar.v3.Data.Event.OrganizerData();
                    myEvt.Organizer.DisplayName = evt.Organizer.Name;
                    myEvt.Organizer.Email = evt.Organizer.Email;
                }
                // Creator
                if(evt.Creator != null)
                {
                    myEvt.Creator = new Google.Apis.Calendar.v3.Data.Event.CreatorData();
                    myEvt.Creator.DisplayName = evt.Creator.Name;
                    myEvt.Creator.Email = evt.Creator.Email;
                }
                // Summary
                if (evt.Summary != "")
                {
                    myEvt.Summary = evt.Summary;
                }
                // Location
                if (evt.Location != null)
                {
                    myEvt.Location = evt.Location.Name;
                }
                // Attendees
                if (evt.Attendees != null)
                {
                    myEvt.Attendees = new List<Google.Apis.Calendar.v3.Data.EventAttendee>();
                    foreach (GenericPerson person in evt.Attendees)
                    {
                        myEvt.Attendees.Add(new Google.Apis.Calendar.v3.Data.EventAttendee
                        {
                            Email = person.Email,
                            DisplayName = person.Name,
                        });
                    }
                }
                // Start 
                if (evt.Start.HasValue)
                {
                    myEvt.Start = new Google.Apis.Calendar.v3.Data.EventDateTime();
                    myEvt.Start.DateTime = evt.Start;
                    myEvt.Start.TimeZone = "Europe/Rome";
                }
                // End
                if (evt.End.HasValue)
                {
                    myEvt.End = new Google.Apis.Calendar.v3.Data.EventDateTime();
                    myEvt.End.DateTime = evt.End;
                    myEvt.End.TimeZone = "Europe/Rome";
                }
                else
                {
                    myEvt.EndTimeUnspecified = true;
                }
                // Recurrency
                if (evt.Recurrency != null)
                {
                    myEvt.Recurrence = new List<string>();
                    GoogleRecurrency temporaryRecurrency = new GoogleRecurrency();
                    // this is bad, dunno how to do otherwise..
                    temporaryRecurrency.Days = evt.Recurrency.Days;
                    temporaryRecurrency.Expiry = evt.Recurrency.Expiry;
                    temporaryRecurrency.Type = evt.Recurrency.Type;
                    myEvt.Recurrence.Add(temporaryRecurrency.ToString());
                }
                // Creation date
                if (evt.Created.HasValue)
                {
                    myEvt.Created = evt.Created;
                }

                myEvt.Reminders = new Google.Apis.Calendar.v3.Data.Event.RemindersData();
                myEvt.Reminders.UseDefault = true;

                Google.Apis.Calendar.v3.Data.Event newlyCreatedEvent = await Service.Events.Insert(myEvt, MyCalendarId).ExecuteAsync();
                if(newlyCreatedEvent != null)
                {
                    res = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: [{0}]", ex.Message);
                res = false;
            }
            //
            return res;
        }

        private async Task<bool> PushEvents(List<GenericEvent> evts)
        {
            bool res = false;
            //
            foreach (GenericEvent evt in evts)
            {
                res = await PushEvent(evt);
                if(res == false)
                {
                    break;
                }
            }
            //
            return res;
        }

        private async Task<List<GenericEvent>> PullEvents()
        {
            List<GenericEvent> myEvts = new List<GenericEvent>();
            try
            {
                Google.Apis.Calendar.v3.Data.Events evts = await Service.Events.List(MyCalendarId).ExecuteAsync();
                foreach(Google.Apis.Calendar.v3.Data.Event evt in evts.Items)
                {
                    GenericEvent myEvt = new GenericEvent();
                    // Summary
                    if (evt.Summary != "")
                    {
                        myEvt.Summary = evt.Summary;
                    }
                    // Description
                    if (evt.Description != "")
                    {
                        myEvt.Description = evt.Description;
                    }
                    // Organizer
                    if (evt.Organizer != null)
                    {
                        myEvt.Organizer = new GenericPerson();
                        myEvt.Organizer.Email = evt.Organizer.Email;
                        myEvt.Organizer.Name = evt.Organizer.DisplayName;
                    }
                    // Creator
                    if (evt.Creator != null)
                    {
                        myEvt.Creator = new GenericPerson();
                        myEvt.Creator.Email = evt.Creator.Email;
                        myEvt.Creator.Name = evt.Creator.DisplayName;
                    }
                    // Location
                    if (evt.Location != "")
                    {
                        myEvt.Location = new GenericLocation();
                        myEvt.Location.Name = evt.Location;
                    }
                    // Start
                    if (evt.Start != null)
                    {
                        myEvt.Start = evt.Start.DateTime;
                    }
                    // End
                    if (evt.End != null)
                    {
                        myEvt.End = evt.End.DateTime;
                    }
                    // Creation date
                    if (evt.Created.HasValue)
                    {
                        myEvt.Created = evt.Created;
                    }
                    // Recurrency
                    if (evt.Recurrence != null)
                    {
                        // TODO: recurrency is a list of stuff coming from google... how this thing is
                        // formatted is unknown.
                        myEvt.Recurrency = new GenericRecurrency();
                        ((GoogleRecurrency)myEvt.Recurrency).FromString(evt.Recurrence[0]);
                    }
                    // Attendees
                    if (evt.Attendees != null)
                    {
                        foreach (Google.Apis.Calendar.v3.Data.EventAttendee attendee in evt.Attendees)
                        {
                            myEvt.Attendees = new List<GenericPerson>();
                            GenericPerson myAttendee = new GenericPerson();
                            //
                            myAttendee.Email = attendee.Email;
                            myAttendee.Name = attendee.DisplayName;
                            //
                            myEvt.Attendees.Add(myAttendee);
                        }
                    }
                    //
                    myEvts.Add(myEvt);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: [{0}]", ex.Message);
            }
            return myEvts;
        }

        #endregion
    }
}
