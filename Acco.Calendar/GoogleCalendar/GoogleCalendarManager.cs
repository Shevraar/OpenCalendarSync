using Acco.Calendar.Event;
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
        string GoogleApplicationName { get { return "Outlook 2007 Calendar Importer"; } }
        string CalendarId { get; set; }
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
            return null;
        }

        #region Initialization
        public async Task<bool> Initialize(string _ClientId, string _ClientSecret)
        {
            bool res = await Authenticate(_ClientId, _ClientSecret);
            //
            CalendarId = await DataStore.GetAsync<string>("CalendarId");
            //
            if (CalendarId != null)
            {
                CalendarId = (await GetCalendar(CalendarId)).Id;
            }
            //
            if (res == false || CalendarId == null)
            {
                CalendarId = (await CreateCalendar()).Id;
                await DataStore.StoreAsync<string>("CalendarId", CalendarId);
            }
            //
            return res;
        }
        private async Task<bool> Authenticate(string _ClientId, string _ClientSecret)
        {
            bool res = true;
            try
            {
                DataStore = new FileDataStore(DataStorePath);
                Credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets
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
                    ApplicationName = GoogleApplicationName
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
                Summary = GoogleApplicationName,
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
                // Creator/Organizer
                if (evt.Creator != null)
                {
                    myEvt.Organizer = new Google.Apis.Calendar.v3.Data.Event.OrganizerData();
                    myEvt.Organizer.DisplayName = evt.Creator.Name;
                    myEvt.Organizer.Email = evt.Creator.Email;
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
                }
                // End
                if (evt.End.HasValue)
                {
                    myEvt.End = new Google.Apis.Calendar.v3.Data.EventDateTime();
                    myEvt.End.DateTime = evt.End;
                }
                else
                {
                    myEvt.EndTimeUnspecified = true;
                }
                // Recurrency
                if (evt.Recurrency != null)
                {
                    myEvt.Recurrence = new List<string>();
                    myEvt.Recurrence.Add(((GoogleRecurrency)evt.Recurrency).ToString());
                }
                if (evt.Created.HasValue)
                {
                    myEvt.Created = evt.Created;
                }

                myEvt.Reminders = new Google.Apis.Calendar.v3.Data.Event.RemindersData();
                myEvt.Reminders.UseDefault = true;

                await Service.Events.Insert(myEvt, CalendarId).ExecuteAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR {0}", ex.Message);
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
                await PushEvent(evt);
            }
            //
            return res;
        }

        private async Task<List<GenericEvent>> PullEvents()
        {
            return null;
        }

        #endregion
    }
}
