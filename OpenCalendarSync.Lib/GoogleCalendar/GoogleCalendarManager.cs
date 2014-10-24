using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using MongoDB.Driver.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OpenCalendarSync.Lib.Event;
using OpenCalendarSync.Lib.Location;
using OpenCalendarSync.Lib.Person;
using OpenCalendarSync.Lib.Utilities;

namespace OpenCalendarSync.Lib.Manager
{
    /// <summary>
    /// Implementation of a Google Calendar Manager based on the latest API version:
    /// https://developers.google.com/google-apps/calendar/v3/reference/
    /// </summary>
    public sealed class GoogleCalendarManager : GenericCalendarManager
    {
        private static readonly log4net.ILog Log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private CalendarService Service { get; set; }

        private UserCredential Credential { get; set; }

        private FileDataStore DataStore { get; set; }

        private static string DataStorePath
        {
            get { return "outlook2googlecalendar"; }
        }

        //private static string SettingsPath
        //{
        //    get { return "googlecalendar.settings"; }
        //}

        private GoogleCalendarParameters _googleCalendarParameters;
        private static readonly GoogleCalendarManager instance = new GoogleCalendarManager();

        // hidden constructor
        private GoogleCalendarManager()
        {
        }

        public static GoogleCalendarManager Instance
        {
            get { return instance; }
        }

        public override IEnumerable<UpdateOutcome> Push(ICalendar calendar)
        {
            var pushTask = PushAsync(calendar);
            pushTask.Wait();
            return pushTask.Result;
        }

        public override ICalendar Pull()
        {
            Log.Info("Pulling calendar from google");
            var pullTask = PullAsync();
            pullTask.RunSynchronously();
            return pullTask.Result;
        }

        public override async Task<IEnumerable<UpdateOutcome>> PushAsync(ICalendar calendar)
        {
            Log.Info(String.Format("Pushing calendar to google [{0}]", calendar.Id));
            if (LastCalendar != null)
            {
                var eventsToRemove = LastCalendar.Events.Where(e => !calendar.Events.Any(elc => elc.Id == e.Id));
                RemoveEvents(eventsToRemove);
            }
            LastCalendar = calendar;
            var res = await PushEvents(calendar.Events);
            return res;
        }

        public override async Task<ICalendar> PullAsync()
        {
            Log.Info("Pulling calendar from google");
            var calendar = new GenericCalendar
            {
                Events = await PullEvents() as DbCollection<GenericEvent>,
                Id = _googleCalendarParameters.Id,
                Name = _googleCalendarParameters.Name
            };
            LastCalendar = calendar;
            return calendar;
        }

        private static DbCollection<GenericEvent> RetrieveEvents()
        {
            Log.Info("Retrieving events from database");
            var query =
            from e in Database.Storage.Instance.Appointments.AsQueryable()
            select e;
            var ret = new DbCollection<GenericEvent>();
            foreach (var evt in query)
            {
                ret.Add(evt);
            }
            return ret;
        }

        /// <summary>
        /// Authenticate to google services with the provided client id and client secret
        /// </summary>
        /// <param name="clientId">The client id provided by google's developer console</param>
        /// <param name="clientSecret">The client secret provided by google's developer console</param>
        /// <returns>True or False if authentication has been successful or not</returns>
        public async Task<bool> Login(string clientId, string clientSecret)
        {
            if(LastCalendar == null)
            {
                LastCalendar = new GenericCalendar
                {
                    Events = RetrieveEvents()
                };
            }
            if(!LoggedIn)
            {
                LoggedIn = await Authenticate(clientId, clientSecret);
            }
            return LoggedIn;
        }

        /// <summary>
        /// Initialize calendar operations by checking if the provided calendar id 
        /// exists on google calendar, if it does not, it creates the calendar and returns the calendar id.
        /// </summary>
        /// <param name="calendarId">Google calendar's id</param>
        /// <param name="calendarName">The summary of google's calendar</param>
        /// <returns>Google's calendar id</returns>
        public async Task<string> Initialize(string calendarId, string calendarName)
        {
            Log.Debug("Initializing...");
            if(LoggedIn)
            {
                // if passed caledar id is not valid
                if(string.IsNullOrEmpty(calendarId))
                { 
                    // create the calendar
                    var onlineCalendar = await CreateCalendar(calendarName);
                    _googleCalendarParameters.Id = onlineCalendar.Id;
                    _googleCalendarParameters.Name = onlineCalendar.Summary;
                }
                else
                {
                    // try to get the calendar with the specified calendarDd
                    var onlineCalendar = await GetCalendar(calendarId);

                    // if the result is KO (no calendar found)
                    if (onlineCalendar == null)
                    {
                        // create a new calendar, with the specified calendarName
                        _googleCalendarParameters.Id = (await CreateCalendar(calendarName)).Id;
                        _googleCalendarParameters.Name = calendarName;
                    }
                    else
                    {
                        // if the result is OK, retrieve the existing calendar
                        _googleCalendarParameters.Id = onlineCalendar.Id;
                        _googleCalendarParameters.Name = onlineCalendar.Description;
                        if (onlineCalendar.Id == calendarId)
                        {
                            Log.Debug("Calendar IDs match!");
                        }
                        else
                        {
                            throw new Exception(String.Format("Passed calendar id [{0}] doesn't match the one on google [{1}]", calendarId, onlineCalendar.Id));
                        }
                        // check if passed calendar name is the same as the one online.
                        if (calendarName != onlineCalendar.Summary)
                        {
                            Log.Warn(String.Format("Online calendar name[{0}] is different from the one stored[{1}]", onlineCalendar.Description, calendarName));
                        }
                    }
                }
            }
            else
            {
                Log.Error("Not logged in, try to log in first");
            }
            Log.Debug("Finished initialization...");
            return _googleCalendarParameters.Id;
        }

        private async Task<bool> Authenticate(string clientId, string clientSecret, string applicationName = "OpenCalendarSync")
        {
            Log.Info("Authenticating to google");
            var res = true;
            try
            {
                DataStore = new FileDataStore(DataStorePath);
                Credential = await
                    GoogleWebAuthorizationBroker.AuthorizeAsync(
                        new ClientSecrets
                        {
                            ClientId = clientId,
                            ClientSecret = clientSecret
                        },
                        new[] {CalendarService.Scope.Calendar},
                        "user",
                        CancellationToken.None,
                        DataStore);
                //
                Service = new CalendarService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = Credential,
                    ApplicationName = applicationName
                });
            }
            catch (Exception ex)
            {
                Log.Error("Exception", ex);
                res = false;
            }
            return res;
        }

        private async Task<Google.Apis.Calendar.v3.Data.Calendar> GetCalendar(string id)
        {
            return await Service.Calendars.Get(id).ExecuteAsync();
        }

        private async Task<Google.Apis.Calendar.v3.Data.Calendar> CreateCalendar(string calendarName)
        {
            Log.Info("Creating calendar");
            try
            {
                return await Service.Calendars.Insert(new Google.Apis.Calendar.v3.Data.Calendar
                {
                    Summary = calendarName,
                    TimeZone = _googleCalendarParameters.TimeZone,
                    Description = "Automatically created: " + DateTime.Now.ToString("g")
                }).ExecuteAsync();
            }
            catch (GoogleApiException ex)
            {
                Log.Error("GoogleApiException", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Exception", ex);
            }
            return null;
        }

        private async Task<UpdateOutcome> PushEvent(IEvent evt)
        {
            var googleEventId = StringHelper.GoogleBase32.ToBaseString(StringHelper.GetBytes(evt.Id)).ToLower();
            Log.Debug(String.Format("Pushing event with googleEventId[{0}]", googleEventId));
            Log.Debug(String.Format("and iCalUID [{0}]", evt.Id));
            var res = new UpdateOutcome {Event = evt as GenericEvent};
            //
            try
            {
                /*
                    Identifier of the event. When creating new single or recurring events, you can specify their IDs. Provided IDs must follow these rules:
                    characters allowed in the ID are those used in base32hex encoding, i.e. lowercase letters a-v and digits 0-9, see section 3.1.2 in RFC2938
                    the length of the ID must be between 5 and 1024 characters
                    the ID must be unique per calendar
                    Due to the globally distributed nature of the system, we cannot guarantee that ID collisions will be detected at event creation time. To minimize the risk of collisions we recommend using an established UUID algorithm such as one described in RFC4122.
                 */
                var myEvt = new Google.Apis.Calendar.v3.Data.Event
                {
                    Id = googleEventId
                };
                // Id
                // Organizer
                if (evt.Organizer != null)
                {
                    myEvt.Organizer = new Google.Apis.Calendar.v3.Data.Event.OrganizerData
                    {
                        DisplayName = evt.Organizer.Name,
                        Email = evt.Organizer.Email
                    };
                }
                // Creator
                if (evt.Creator != null)
                {
                    myEvt.Creator = new Google.Apis.Calendar.v3.Data.Event.CreatorData
                    {
                        DisplayName = evt.Creator.Name,
                        Email = evt.Creator.Email
                    };
                }
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
                        var r = person.Response.GetAttributeOfType<GoogleResponseStatus>();
                        myEvt.Attendees.Add(new Google.Apis.Calendar.v3.Data.EventAttendee
                        {
                            Email = person.Email,
                            DisplayName = person.Name,
                            ResponseStatus = r.Text
                        });
                    }
                }
                // Start
                if (evt.Start.HasValue)
                {
                    myEvt.Start = new Google.Apis.Calendar.v3.Data.EventDateTime
                    {
                        DateTime = evt.Start,
                        TimeZone = "Europe/Rome"
                    };
                }
                // End
                if (evt.End.HasValue)
                {
                    myEvt.End = new Google.Apis.Calendar.v3.Data.EventDateTime
                    {
                        DateTime = evt.End,
                        TimeZone = "Europe/Rome"
                    };
                }
                else
                {
                    myEvt.EndTimeUnspecified = true;
                }
                // Recurrency
                if (evt.Recurrence != null)
                {
                    myEvt.Recurrence = new List<string> {evt.Recurrence.Get()};
                }
                // Creation date
                if (evt.Created.HasValue)
                {
                    myEvt.Created = evt.Created;
                }
                //
                myEvt.Reminders = new Google.Apis.Calendar.v3.Data.Event.RemindersData {UseDefault = true};
                //
                var createdEvent = await Service.Events.Insert(myEvt, _googleCalendarParameters.Id).ExecuteAsync();
                //
                if (createdEvent != null)
                {
                    res.Successful = true;
                }
            }
            catch(GoogleApiException ex)
            {
                Log.Error("GoogleApiException", ex);
                res.Successful = false;
            }
            catch (AggregateException ex)
            {
                foreach (var e in ex.InnerExceptions)
                {
                    Log.Error(e.GetType().ToString(), e);
                }
                res.Successful = false;
            }
            //
            return res;
        }

        private async Task<IEnumerable<UpdateOutcome>> PushEvents(IEnumerable<IEvent> evts)
        {
            var res = new List<UpdateOutcome>();
            // handle exceptions in a bulk
            var pushExceptions = new List<Exception>();
            // new events management
            var events = evts as IList<IEvent> ?? evts.ToList();
            foreach (var newEvent in events.Where(evt => evt.Action == EventAction.Add))
            {
                var currentEvent = await PushEvent(newEvent);
                res.Add(currentEvent); // add it anyway
                if (currentEvent.Successful == false) { pushExceptions.Add(new PushException("PushEvent failed", newEvent as GenericEvent)); }
            }
            // updated events management
            foreach(var updatedEvent in events.Where(evt => evt.Action == EventAction.Update))
            {
                try 
                { 
                    var currentEvent = await UpdateEvent(updatedEvent);
                    res.Add(currentEvent);
                    if (currentEvent.Successful == false) { pushExceptions.Add(new PushException("UpdateEvent failed", updatedEvent as GenericEvent)); }
                }
                catch(Exception ex)
                {
                    Log.Error("Exception", ex);
                    pushExceptions.Add(new PushException("UpdateEvent failed", updatedEvent as GenericEvent));
                }
            }
            // throw the exceptions, if any
            if (pushExceptions.Count > 0)
            {
                throw new AggregateException(pushExceptions);
            }
            return res;
        }

        private async Task<IList<GenericEvent>> PullEvents()
        {
            Log.Info("Pulling events");
            var myEvts = new DbCollection<GenericEvent>();
            try
            {
                var evts = await Service.Events.List(_googleCalendarParameters.Id).ExecuteAsync();
                foreach (var evt in evts.Items)
                {
                    var iCalUid = StringHelper.GetString(StringHelper.GoogleBase32.FromBaseString(evt.Id));
                    var myEvt = new GenericEvent(   id: iCalUid,
                                                    summary: evt.Summary,
                                                    description: evt.Description,
                                                    location: new GenericLocation {Name = evt.Location});
                    // Organizer
                    if (evt.Organizer != null)
                    {
                        myEvt.Organizer = new GenericPerson
                        {
                            Email = evt.Organizer.Email,
                            Name = evt.Organizer.DisplayName
                        };
                    }
                    // Creator
                    if (evt.Creator != null)
                    {
                        myEvt.Creator = new GenericPerson
                        {
                            Email = evt.Creator.Email,
                            Name = evt.Creator.DisplayName
                        };
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
                        ((GoogleRecurrence) myEvt.Recurrence).Parse(evt.Recurrence[0]);
                            //warning: this only parses one line inside Recurrence...
                    }
                    // Attendees
                    if (evt.Attendees != null)
                    {
                        myEvt.Attendees = new List<GenericAttendee>();
                        foreach (var attendee in evt.Attendees)
                        {
                            ResponseStatus r;
                            switch (attendee.ResponseStatus)
                            {
                                case "accepted":
                                    r = ResponseStatus.Accepted;
                                    break;
                                case "tentative":
                                    r = ResponseStatus.Tentative;
                                    break;
                                case "needsAction":
                                    r = ResponseStatus.NotResponded;
                                    break;
                                case "declined":
                                    r = ResponseStatus.Declined;
                                    break;
                                default:
                                    r = ResponseStatus.None;
                                    break;
                            }
                            myEvt.Attendees.Add(
                                new GenericAttendee
                                {
                                    Email = attendee.Email,
                                    Name = attendee.DisplayName,
                                    Response = r
                                }
                                );
                        }
                    }
                    myEvts.Add(myEvt);
                }
            }
            catch (Exception ex)
            {
                Log.Error("Exception", ex);
            }
            return myEvts;
        }

        private async Task<IList<GenericEvent>> PullEvents(DateTime from, DateTime to)
        {
            Log.Info(String.Format("Pulling events from [{0}] to [{1}]", from, to));
            var myEvts = new DbCollection<GenericEvent>();
            var evts = (await PullEvents()) as DbCollection<GenericEvent>;
            // note: google doesn't provide a direct way to filter events when listing them
            //       so we have to filter them manually
            if (evts != null)
            {
                var excludedEvts = evts.Where(x => (x.Start < @from && x.End > to)).ToList(); // todo: have to test this
                foreach (var excludedEvt in excludedEvts)
                {
                    evts.Remove(excludedEvt);
                }
            }
            else
            {
                throw new Exception();
            }
            return myEvts;
        }

        public override ICalendar Pull(DateTime from, DateTime to)
        {
            var pullTask = PullAsync(from, to);
            pullTask.RunSynchronously();
            return pullTask.Result;
        }

        public override async Task<ICalendar> PullAsync(DateTime from, DateTime to)
        {
            var calendar = new GenericCalendar
            {
                Events = new DbCollection<GenericEvent>()
            };
            calendar.Events = await PullEvents(from, to) as DbCollection<GenericEvent>;
            calendar.Id = _googleCalendarParameters.Id;
            calendar.Name = _googleCalendarParameters.Name;
            LastCalendar = calendar;
            return calendar;
        }

        private async void RemoveEvents(IEnumerable<GenericEvent> eventsToRemove)
        {
            foreach (var evt in eventsToRemove)
            {
                try
                {
                    Log.Debug(String.Format("Remove event with google id [{0}]", StringHelper.GoogleBase32.ToBaseString(StringHelper.GetBytes(evt.Id))));
                    Log.Debug(String.Format("and iCalUID [{0}]", evt.Id));
                    var res = await Service.Events.Delete(_googleCalendarParameters.Id, StringHelper.GoogleBase32.ToBaseString(StringHelper.GetBytes(evt.Id))).ExecuteAsync();
                    if (!string.IsNullOrEmpty(res))
                    {
                        Log.Debug(res);
                    }
                }
                catch (GoogleApiException ex)
                {
                    Log.Error("GoogleApiException", ex);
                }
                catch (Exception ex)
                {
                    Log.Error("Exception", ex);
                }
            }
        }

        private async Task<UpdateOutcome> UpdateEvent(IEvent updatedEvent)
        {
            Log.Debug(String.Format("Updating existing event [{0}]...", updatedEvent.Id));
            var res = new UpdateOutcome { Event = updatedEvent as GenericEvent, Successful = false };
            //
            var googleEventId = StringHelper.GoogleBase32.ToBaseString(StringHelper.GetBytes(updatedEvent.Id)).ToLower();
            /*
                Identifier of the event. When creating new single or recurring events, you can specify their IDs. Provided IDs must follow these rules:
                characters allowed in the ID are those used in base32hex encoding, i.e. lowercase letters a-v and digits 0-9, see section 3.1.2 in RFC2938
                the length of the ID must be between 5 and 1024 characters
                the ID must be unique per calendar
                Due to the globally distributed nature of the system, we cannot guarantee that ID collisions will be detected at event creation time. To minimize the risk of collisions we recommend using an established UUID algorithm such as one described in RFC4122.
             */
            var myEvt = new Google.Apis.Calendar.v3.Data.Event
            {
                Id = googleEventId
            };
            //
            // Id
            // Organizer
            if (updatedEvent.Organizer != null)
            {
                myEvt.Organizer = new Google.Apis.Calendar.v3.Data.Event.OrganizerData
                {
                    DisplayName = updatedEvent.Organizer.Name,
                    Email = updatedEvent.Organizer.Email
                };
            }
            // Creator
            if (updatedEvent.Creator != null)
            {
                myEvt.Creator = new Google.Apis.Calendar.v3.Data.Event.CreatorData
                {
                    DisplayName = updatedEvent.Creator.Name,
                    Email = updatedEvent.Creator.Email
                };
            }
            // Summary
            if (updatedEvent.Summary != "")
            {
                myEvt.Summary = updatedEvent.Summary;
            }
            // Description
            if (updatedEvent.Description != "")
            {
                myEvt.Description = updatedEvent.Description;
            }
            // Location
            if (updatedEvent.Location != null)
            {
                myEvt.Location = updatedEvent.Location.Name;
            }
            // Attendees
            if (updatedEvent.Attendees != null)
            {
                myEvt.Attendees = new List<Google.Apis.Calendar.v3.Data.EventAttendee>();
                foreach (var person in updatedEvent.Attendees)
                {
                    var r = person.Response.GetAttributeOfType<GoogleResponseStatus>();
                    myEvt.Attendees.Add(new Google.Apis.Calendar.v3.Data.EventAttendee
                    {
                        Email = person.Email,
                        DisplayName = person.Name,
                        ResponseStatus = r.Text
                    });
                }
            }
            // Start
            if (updatedEvent.Start.HasValue)
            {
                myEvt.Start = new Google.Apis.Calendar.v3.Data.EventDateTime
                {
                    DateTime = updatedEvent.Start,
                    TimeZone = _googleCalendarParameters.TimeZone
                };
            }
            // End
            if (updatedEvent.End.HasValue)
            {
                myEvt.End = new Google.Apis.Calendar.v3.Data.EventDateTime
                {
                    DateTime = updatedEvent.End,
                    TimeZone = _googleCalendarParameters.TimeZone
                };
            }
            else
            {
                myEvt.EndTimeUnspecified = true;
            }
            // Recurrency
            if (updatedEvent.Recurrence != null)
            {
                myEvt.Recurrence = new List<string> { updatedEvent.Recurrence.Get() };
            }
            // Creation date
            if (updatedEvent.Created.HasValue)
            {
                myEvt.Created = updatedEvent.Created;
            }
            //
            myEvt.Reminders = new Google.Apis.Calendar.v3.Data.Event.RemindersData { UseDefault = true };
            var existingEvent = await Service.Events.Get(_googleCalendarParameters.Id, myEvt.Id).ExecuteAsync();
            if (existingEvent.Sequence.HasValue)
            {
                myEvt.Sequence = existingEvent.Sequence.Value + 1;
            }
            else
            { 
                throw new Exception(String.Format("Failed to get sequence number for existing event[{0}], better luck next time", existingEvent.Id));
            }
            //
            var update = await Service.Events.Update(myEvt, _googleCalendarParameters.Id, myEvt.Id).ExecuteAsync();
            if (update != null)
            {
                res.Successful = true; 
            }
            return res;
        }

        public async Task<bool> SetCalendarColor(string foregroundColor, string backgroundColor)
        {
            if (!LoggedIn)
            { 
                return false;
            }
            var request = Service.CalendarList.Update(new Google.Apis.Calendar.v3.Data.CalendarListEntry { BackgroundColor = backgroundColor, ForegroundColor = foregroundColor }, _googleCalendarParameters.Id);
            request.ColorRgbFormat = true; // if we don't do this, google wants a colorId, which we don't have.
            var ret = await request.ExecuteAsync();
            if (ret != null)
                return true;
            return false;
        }

        public async Task<string> DropCurrentCalendar()
        {
            if(!LoggedIn)
            {
                return "not logged in";
            }
            return await Service.Calendars.Delete(_googleCalendarParameters.Id).ExecuteAsync();
        }

        public bool LoggedIn { get; private set; }

        private struct GoogleCalendarParameters
        {
            public string Id { get; set; }

            public string Name { get; set; }

            public string TimeZone { get; set; }
        }
    }
}