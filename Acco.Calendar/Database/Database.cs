using System;
//
using Acco.Calendar.Event;
//
using MongoDB.Driver;

namespace Acco.Calendar.Database
{
    public sealed class Storage
    {
        static string ConnectionString { get { return "mongodb://localhost"; } }
        MongoClient Client { get; set; }
        MongoServer Server { get; set; }
        public MongoDatabase Database { get; private set; }
        public MongoCollection<GenericEvent> Appointments { get; private set; }
        private static volatile Storage instance;
        private static readonly object SyncRoot = new Object();

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
                    lock (SyncRoot)
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
