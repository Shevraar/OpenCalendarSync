//
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
//
using Acco.Calendar.Event;
using Acco.Calendar.Location;
using Acco.Calendar.Person;
//
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Http;
using Google.Apis.Services;
using Google.Apis.Util.Store;
//
using Newtonsoft.Json;

namespace Acco.Calendar
{
    public interface ICalendarManager
    {
        bool Push(GenericCalendar calendar);
        Task<bool> PushAsync(GenericCalendar calendar);
        GenericCalendar Pull();
        Task<GenericCalendar> PullAsync();

    }

    public class OutlookCalendarManager : ICalendarManager
    {
        public bool Push(GenericCalendar calendar)
        {
            return false;
        }

        public GenericCalendar Pull()
        {
            return null;
        }

        public async Task<bool> PushAsync(GenericCalendar calendar)
        {
            return false;
        }

        public async Task<GenericCalendar> PullAsync()
        {
            return null;
        }
    }

    public sealed class GoogleCalendarManager : ICalendarManager
    {
        #region Variables + Constants
        CalendarService Service { get; set; }
        UserCredential Credential { get; set; }
        FileDataStore DataStore { get; set; }
        string DataStorePath { get { return "Acco.Calendar.GoogleCalendarManager"; } }
        string GoogleApplicationName { get { return "Outlook 2007 Calendar Importer"; } }
        string MyCalendarId { get; set; }
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

            MyCalendarId = await DataStore.GetAsync<string>("MyCalendarId");

            if (MyCalendarId != null)
            {
                res = await CheckCalendar(MyCalendarId);
            }
            
            if(res == false || MyCalendarId == null)
            {
                MyCalendarId = await CreateCalendar();
                await DataStore.StoreAsync<string>("MyCalendarId", MyCalendarId);
            }

            return res;
        }
        private async Task<bool> Authenticate(string _ClientId, string _ClientSecret)
        {
            bool res = true;
            try
            {
                DataStore = new FileDataStore(DataStorePath);
                Credential = await GoogleWebAuthorizationBroker.AuthorizeAsync( new ClientSecrets 
                                                                                    {
                                                                                        ClientId = _ClientId,
                                                                                        ClientSecret = _ClientSecret
                                                                                    },
                                                                                    new[] { CalendarService.Scope.Calendar },
                                                                                    "user",
                                                                                    CancellationToken.None,
                                                                                    DataStore);

                Service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = Credential,
                    ApplicationName = GoogleApplicationName
                });
            }
            catch(Exception ex)
            {
                Console.WriteLine("ERROR: {0}", ex.Message);
                res = false;
            }

            return res;
        }


        #endregion

        #region Push + Pull events to Google Calendar
        private async Task<bool> PushEvent(GenericEvent evt)
        {
            bool res = false;

            return res;
        }

        private async Task<bool> PushEvents(List<GenericEvent> evts)
        {
            bool res = false;

            return res;
        }

        private async Task<List<GenericEvent>> PullEvents()
        {
            return null;
        }
        #endregion

        #region First time access - create calendar and store its id
        private async Task<bool> CheckCalendar(string id)
        {
            bool res = false;

            var calendar = await Service.Calendars.Get(id).ExecuteAsync();
            if (calendar != null)
            { 
                res = true;
            }

            return res;
        }
        private async Task<string> CreateCalendar()
        {
            string res = "";

            var calendar = await Service.Calendars.Insert(new Google.Apis.Calendar.v3.Data.Calendar() { Summary = "Automatically created by " + GoogleApplicationName + " @ " + DateTime.Now.ToString("g")}).ExecuteAsync();
            res = calendar.Id;

            return res;
        }
        #endregion
    }
}
