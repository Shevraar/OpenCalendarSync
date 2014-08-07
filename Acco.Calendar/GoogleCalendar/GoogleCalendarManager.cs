using Acco.Calendar.Event;
using Acco.Calendar.Location;
using Acco.Calendar.Person;
using Acco.Calendar.Utilities;
using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using MongoDB.Bson;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WallF.BaseNEncodings;

namespace Acco.Calendar.Manager
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

        private static string SettingsPath
        {
            get { return "googlecalendar.settings"; }
        }

        private GoogleCalendarSettings _settings = new GoogleCalendarSettings();
        private static readonly GoogleCalendarManager instance = new GoogleCalendarManager();

        // hidden constructor
        private GoogleCalendarManager()
        {
        }

        public static GoogleCalendarManager Instance
        {
            get { return instance; }
        }

        public override bool Push(ICalendar calendar)
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

        public override async Task<bool> PushAsync(ICalendar calendar)
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
                Id = _settings.CalendarId,
                Name = _settings.CalendarName
            };
            LastCalendar = calendar;
            return calendar;
        }

        public async Task<bool> Initialize(string clientId, string clientSecret, string calendarName)
        {
            Log.Info(String.Format("Initializing google calendar [{0}]", calendarName));
            var authenticated = await Authenticate(clientId, clientSecret);
            if (authenticated)
            {
                _settings = await CreateSettings(calendarName); 
                var theirCalendarId = (await GetCalendar(_settings.CalendarId)).Id;
                if (_settings.CalendarId == theirCalendarId)
                {
                    Log.Debug(String.Format("Our calendar id matches the one on google: id[{0}]", _settings.CalendarId));
                }
                else
                {
                    throw new Exception(String.Format("Stored calendar id [{0}] doesn't match the one on google [{1}]",
                        _settings.CalendarId, theirCalendarId));
                }
            }
            return authenticated;
        }

        private async Task<bool> Authenticate(string clientId, string clientSecret)
        {
            Log.Info("Authenticating to google");
            var res = true;
            try
            {
                DataStore = new FileDataStore(DataStorePath);
                Credential = await
                    GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets
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
                    ApplicationName = _settings.ApplicationName
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
                    TimeZone = "Europe/Rome", //todo: configurable
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

        private async Task<bool> RemoveCalendar(string calendarId)
        {
            Log.Info(String.Format("Removing calendar [{0}]", calendarId));
            var res = false;
            var deleteResult = await Service.Calendars.Delete(calendarId).ExecuteAsync();
            if (deleteResult == "")
            {
                res = true;
            }
            else
            {
                Log.Error(String.Format("Error removing calendar [{0}], deleteResult [{1}]", calendarId, deleteResult));
            }
            return res;
        }

        private async Task<bool> PushEvent(IEvent evt)
        {
            var googleEventId = StringHelper.GoogleBase32.ToBaseString(StringHelper.GetBytes(evt.Id)).ToLower();
            Log.Debug(String.Format("Pushing event with googleEventId[{0}]", googleEventId));
            Log.Debug(String.Format("and iCalUID [{0}]", evt.Id));
            var res = false;
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
                var createdEvent = await Service.Events.Insert(myEvt, _settings.CalendarId).ExecuteAsync();
                //
                if (createdEvent != null)
                {
                    res = true;
                }
            }
            catch(GoogleApiException ex)
            {
                Log.Error("GoogleApiException", ex);
                res = false;
            }
            catch (AggregateException ex)
            {
                foreach (var e in ex.InnerExceptions)
                {
                    Log.Error(e.GetType().ToString(), e);
                }
                res = false;
            }
            //
            return res;
        }

        private async Task<bool> PushEvents(IEnumerable<IEvent> evts)
        {
            var res = true;
            //
            foreach (var evt in evts.Where(evt => evt.EventAction == EventAction.Add))
            {
                res = await PushEvent(evt);
                if (res == false)
                {
                    throw new PushException("PushEvent failed", evt as GenericEvent);
                }
            }
            //
            return res;
        }

        private async Task<IList<GenericEvent>> PullEvents()
        {
            Log.Info("Pulling events");
            var myEvts = new DbCollection<GenericEvent>();
            try
            {
                var evts = await Service.Events.List(_settings.CalendarId).ExecuteAsync();
                foreach (var evt in evts.Items)
                {
                    var iCalUID = StringHelper.GetString(StringHelper.GoogleBase32.FromBaseString(evt.Id));
                    var myEvt = new GenericEvent(   id: iCalUID,
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
            calendar.Id = _settings.CalendarId;
            calendar.Name = _settings.CalendarName;
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
                    var res = await Service.Events.Delete(_settings.CalendarId, StringHelper.GoogleBase32.ToBaseString(StringHelper.GetBytes(evt.Id))).ExecuteAsync();
                    Log.Debug(res);
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

        private async Task<GoogleCalendarSettings> CreateSettings(string calendarName)
        {
            GoogleCalendarSettings temporarySettings = null;
            if (File.Exists(SettingsPath))
            {
                using (var r = new StreamReader(SettingsPath))
                {
                    var json = r.ReadToEnd();
                    temporarySettings = JsonConvert.DeserializeObject<GoogleCalendarSettings>(json);
                    if (temporarySettings.CalendarName != calendarName)
                    {
                        Log.Warn(String.Format("Calendar name mismatch stored:[{0}], provided:[{1}]",
                            temporarySettings.CalendarName, calendarName));
                        Log.Warn("Deleting old calendar and making a new one");
                        var isCalendarDeleted = await RemoveCalendar(temporarySettings.CalendarId);
                        if (isCalendarDeleted)
                        {
                            Log.Info("Calendar successfully deleted");
                        }
                        else
                        {
                            throw new Exception(
                                String.Format("Failed to delete calendar id[{0}] and name [{1}]",
                                    temporarySettings.CalendarId, temporarySettings.CalendarId));
                        }
                    }
                }
            }
            else
            {
                try
                {
                    var calendarId = (await CreateCalendar(calendarName)).Id;
                    temporarySettings = new GoogleCalendarSettings
                    {
                        CalendarName = calendarName,
                        CalendarId = calendarId
                    };
                }
                catch (Exception ex)
                {
                    Log.Error("Exception", ex);
                }
                var jsonSettings = JsonConvert.SerializeObject(temporarySettings);
                using (var sw = new StreamWriter(SettingsPath))
                {
                    await sw.WriteAsync(jsonSettings);
                }
            }
            Log.Info(temporarySettings.ToJson());
            return temporarySettings;
        }

        [Serializable]
        internal class GoogleCalendarSettings
        {
            public string CalendarId { get; set; }

            public string CalendarName { get; set; }

            public string ApplicationName
            {
                get { return "Outlook2GoogleCalendar"; }
            }
        }
    }
}