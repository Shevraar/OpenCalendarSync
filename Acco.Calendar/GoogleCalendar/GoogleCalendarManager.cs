using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.ObjectModel;
//
using Acco.Calendar.Event;
using Acco.Calendar.Location;
using Acco.Calendar.Person;
using Acco.Calendar.Database;
//
using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Collections.Specialized;
using MongoDB.Driver.Builders;
//

namespace Acco.Calendar.Manager
{
    /// <summary>
    /// Implementation of a Google Calendar Manager based on the latest API version:
    /// https://developers.google.com/google-apps/calendar/v3/reference/
    /// </summary>
    public sealed class GoogleCalendarManager : ICalendarManager
    {
        CalendarService Service { get; set; }
        UserCredential Credential { get; set; }
        FileDataStore DataStore { get; set; }
        string DataStorePath { get { return "Acco.Calendar.GoogleCalendarManager"; } } // todo: replace this with mongodb config
        string AppName { get { return "Outlook 2007 Calendar Importer"; } } // todo: replace this with mongodb config
        string MyCalendarId { get; set; }  // todo: replace this with mongodb config
        string MyCalendarName { get; set; } // todo: replace this with mongodb config
        private static readonly GoogleCalendarManager instance = new GoogleCalendarManager();
        // hidden constructor
        private GoogleCalendarManager() { }
        public static GoogleCalendarManager Instance { get { return instance; } }

        public bool Push(ICalendar calendar)
        {
            var pushTask = PushAsync(calendar);
            pushTask.RunSynchronously();
            return pushTask.Result;
        }

        public ICalendar Pull()
        {
            var pullTask = PullAsync();
            pullTask.RunSynchronously();
            return pullTask.Result;
        }

        public async Task<bool> PushAsync(ICalendar calendar)
        {
            var res = false;
            res = await PushEvents(calendar.Events);
            return res;
        }

        public async Task<ICalendar> PullAsync()
        {
            var calendar = new GenericCalendar();
            calendar.Events = await PullEvents() as ObservableCollection<GenericEvent>;
            calendar.Id = MyCalendarId;
            calendar.Name = MyCalendarName;
            return calendar;
        }

        private void Events_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            // todo: work in progress
            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    // note: to know which item was added, use NewItems.
                    Console.WriteLine("Event added");
                    foreach (GenericEvent item in e.NewItems) //todo: check if its possible to add the list of added events
                    {
                        Storage.Instance.Appointments.Save(item);
                    }
                    break;
                case NotifyCollectionChangedAction.Remove:
                    Console.WriteLine("Event removed");
                    foreach (GenericEvent item in e.OldItems) //todo: check if its possible to delete the list of removed events
                    {
                        var query = Query<GenericEvent>.EQ(evt => evt.Id, item.Id);
                        Storage.Instance.Appointments.Remove(query);
                    }
                    break;
                default:
                    throw new System.Exception("Unmanaged Action => " + e.Action);
            }
        }

        public async Task<bool> Initialize(string _ClientId, string _ClientSecret, string _CalendarName)
        {
            // todo: stop using google's datastore to store this kind of stuff and start using MongoDB
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
                Console.WriteLine("Exception: [{0}]", ex.Message);
                res = false;
            }
            return res;
        }

        private async Task<Google.Apis.Calendar.v3.Data.Calendar> GetCalendar(string id)
        {
            return await Service.Calendars.Get(id).ExecuteAsync();
        }
        private async Task<Google.Apis.Calendar.v3.Data.Calendar> CreateCalendar()
        {
            return await Service.Calendars.Insert(new Google.Apis.Calendar.v3.Data.Calendar()
            {
                Summary = MyCalendarName,
                TimeZone = "Europe/Rome", //todo: configurable
                Description = "Automatically created: " + DateTime.Now.ToString("g")
            }).ExecuteAsync();
        }

        private async Task<bool> RemoveCalendar(string calendarId)
        {
            var res = false;
            var deleteResult = await Service.Calendars.Delete(calendarId).ExecuteAsync();
            if (deleteResult == "") { res = true; }
            else { Console.WriteLine("Error while removing calendar [{0}] => [{1}]", calendarId, deleteResult); }
            return res;
        }

        private async Task<bool> PushEvent(GenericEvent evt)
        {
            var res = false;
            //
            try
            {
                var myEvt = new Google.Apis.Calendar.v3.Data.Event();
                // Id
                myEvt.Id = evt.Id;
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
                    foreach (var person in evt.Attendees)
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
                    myEvt.Start.TimeZone = "Europe/Rome"; // documentation says its optional, but google still needs it...
                }
                // End
                if (evt.End.HasValue)
                {
                    myEvt.End = new Google.Apis.Calendar.v3.Data.EventDateTime();
                    myEvt.End.DateTime = evt.End;
                    myEvt.End.TimeZone = "Europe/Rome"; // documentation says its optional, but google still needs it...
                }
                else
                {
                    myEvt.EndTimeUnspecified = true;
                }
                // Recurrency
                if (evt.Recurrence != null)
                {
                    myEvt.Recurrence = new List<string>();
                    myEvt.Recurrence.Add(evt.Recurrence.Get());
                }
                // Creation date
                if (evt.Created.HasValue)
                {
                    myEvt.Created = evt.Created;
                }
                //
                myEvt.Reminders = new Google.Apis.Calendar.v3.Data.Event.RemindersData();
                myEvt.Reminders.UseDefault = true;
                //
                var createdEvent = await Service.Events.Insert(myEvt, MyCalendarId).ExecuteAsync();
                //
                if(createdEvent != null) { res = true; }
            }
            catch(GoogleApiException ex)
            {
                Console.WriteLine("GoogleApiException: [{0}]", ex.Message); //todo: add improved logging...
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: [{0}]", ex.Message); //todo: add improved logging...
                res = false;
            }
            //
            return res;
        }

        private async Task<bool> PushEvents(IList<GenericEvent> evts)
        {
            bool res = false;
            //
            foreach (var evt in evts)
            {
                res = await PushEvent(evt);
                if(res == false)
                {
                    //Console.WriteLine("Event [{0}] was not pushed to google calendar. Aborting", evt.Id);
                    throw new PushException("PushEvent failed", evt);
                }
            }
            //
            return res;
        }

        private async Task<IList<GenericEvent>> PullEvents()
        {
            var myEvts = new ObservableCollection<GenericEvent>();
            myEvts.CollectionChanged += Events_CollectionChanged;
            try
            {
                var evts = await Service.Events.List(MyCalendarId).ExecuteAsync();
                foreach(var evt in evts.Items)
                {
                    var myEvt = new GenericEvent(   Id: evt.Id, 
                                                    Summary: evt.Summary, 
                                                    Description: evt.Description, 
                                                    Location: new GenericLocation{ Name = evt.Location }    );
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
                    // Recurrence
                    if (evt.Recurrence != null)
                    {
                        myEvt.Recurrence = new GoogleRecurrence();
                        ((GoogleRecurrence)myEvt.Recurrence).Parse<String>(evt.Recurrence[0]); //warning: this only parses one line inside Recurrence...
                    }
                    // Attendees
                    if (evt.Attendees != null)
                    {
                        foreach (var attendee in evt.Attendees)
                        {
                            myEvt.Attendees = new List<GenericPerson>();
                            var myAttendee = new GenericPerson();
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
            myEvts.CollectionChanged -= Events_CollectionChanged;
            return myEvts;
        }

        private async Task<IList<GenericEvent>> PullEvents(DateTime from, DateTime to)
        {
            var myEvts = new ObservableCollection<GenericEvent>();
            var evts = (await PullEvents()) as ObservableCollection<GenericEvent>;
            // note: google doesn't provide a direct way to filter events when listing them
            //       so we have to filter them manually
            evts.CollectionChanged += Events_CollectionChanged;
            var excludedEvts = evts.Where(x => (x.Start < from && x.End > to)).ToList(); // todo: have to test this
            foreach(var excludedEvt in excludedEvts)
            {
                evts.Remove(excludedEvt);
            }
            evts.CollectionChanged -= Events_CollectionChanged;
            return myEvts;
        }

        public ICalendar Pull(DateTime from, DateTime to)
        {
            var pullTask = PullAsync();
            pullTask.RunSynchronously();
            return pullTask.Result;
        }

        public async Task<ICalendar> PullAsync(DateTime from, DateTime to)
        {
            var calendar = new GenericCalendar();
            calendar.Events.CollectionChanged += Events_CollectionChanged;
            calendar.Events = await PullEvents(from, to) as ObservableCollection<GenericEvent>;
            calendar.Id = MyCalendarId;
            calendar.Name = MyCalendarName;
            calendar.Events.CollectionChanged -= Events_CollectionChanged;
            return calendar;
        }
    }
}
