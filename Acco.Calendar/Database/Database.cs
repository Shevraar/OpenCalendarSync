using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//
using Acco.Calendar.Event;
//
using MongoDB.Driver;

namespace Acco.Calendar.Database
{
    public sealed class Storage
    {
        string ConnectionString { get { return "mongodb://localhost"; } }
        MongoClient Client { get; set; }
        MongoServer Server { get; set; }
        public MongoDatabase Database { get; private set; }
        public MongoCollection<GenericEvent> Appointments { get; private set; }
        private static volatile Storage instance;
        private static object syncRoot = new Object();

        private Storage()
        {
            Client = new MongoClient(ConnectionString);
            Server = Client.GetServer();
            Database = Server.GetDatabase("AccoCalendar");
            Appointments = Database.GetCollection<GenericEvent>("appointments");
        }

        public static Storage Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                            instance = new Storage();
                    }
                }
                return instance;
            }
        } 
    }
}
