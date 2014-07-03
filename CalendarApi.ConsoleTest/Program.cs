using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace CalendarApiConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Calendar API Sample: List MyLibrary");
            Console.WriteLine("================================");
            try
            {
                new Program().Run().Wait();
            }
            catch (AggregateException ex)
            {
                foreach (var e in ex.InnerExceptions)
                {
                    Console.WriteLine("ERROR: " + e.Message);
                }
            }
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        private async Task Run()
        {
            UserCredential credential;
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new[] { CalendarService.Scope.Calendar },
                    "user", CancellationToken.None, new FileDataStore("Books.ListMyLibrary"));
            }

            // Create the service.
            var service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Books API Sample",
                });
            
            // list calendars
            foreach (CalendarListEntry entry in service.CalendarList.List().Execute().Items)
            {
                // generic calendar info
                Console.WriteLine("=Calendar======================================");
                Console.WriteLine("\t id[{0}]", entry.Id);
                Console.WriteLine("\t descr[{0}]", entry.Description);
                Console.WriteLine("\t kind[{0}]", entry.Kind);
                Console.WriteLine("\t isSelected[{0}]", entry.Selected.HasValue ? ((bool)entry.Selected ?  "true" : "false") : "no-val");
                Console.WriteLine("\t isHidden[{0}]", entry.Hidden.HasValue ? ((bool)entry.Hidden ? "true" : "false") : "no-val");

                IList<Event> events = service.Events.List(entry.Id).Execute().Items;
                JsonSerializer serializer = new JsonSerializer();
                serializer.Converters.Add(new JavaScriptDateTimeConverter());
                serializer.NullValueHandling = NullValueHandling.Include;
                serializer.Formatting = Formatting.Indented;
                using (StreamWriter sw = new StreamWriter(@"./calendar+" + entry.Id + ".json"))
                using (JsonWriter writer = new JsonTextWriter(sw))
                {
                    foreach(Event evt in events)
                    {
                        Console.WriteLine("====Event=====================================");
                        Console.WriteLine("\t id[{0}]", evt.Id);
                        Console.WriteLine("\t descr[{0}]", evt.Description);
                        Console.WriteLine("\t start[{0}]", evt.Start.DateTime.HasValue ? evt.Start.DateTime.ToString() : "no-val");
                        Console.WriteLine("\t end[{0}]", evt.End.DateTime.HasValue ? evt.End.DateTime.ToString() : "no-val");
                        Console.WriteLine("\t location[{0}]", evt.Location);
                        serializer.Serialize(writer, evt);
                    }
                }
            }
        }
    }
}
