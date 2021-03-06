﻿using MongoDB.Driver;
using System;
using OpenCalendarSync.Lib.Event;

namespace OpenCalendarSync.Lib.Database
{
    public sealed class Storage
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static volatile Storage _instance;
        private static readonly object SyncRoot = new Object();

        private Storage()
        {
            Log.Info("Initializing storage...");
            Client = new MongoClient(ConnectionString);
            Server = Client.GetServer();
            Database = Server.GetDatabase("OpenCalendarSync");
            Appointments = Database.GetCollection<GenericEvent>("appointments");
        }

        private static string ConnectionString
        {
            get { return "mongodb://localhost"; }
        }

        private MongoClient Client { get; set; }

        private MongoServer Server { get; set; }

        internal MongoDatabase Database { get; private set; }

        public MongoCollection<GenericEvent> Appointments { get; private set; }

        public static Storage Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (SyncRoot)
                    {
                        if (_instance == null)
                            _instance = new Storage();
                    }
                }
                return _instance;
            }
        }
    }
}