﻿using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Threading;
using OpenCalendarSync.Lib.Event;

namespace OpenCalendarSync.Lib.Manager
{
    public struct UpdateOutcome
    {
        public IEvent Event { get; set; }
        public bool Successful { get; set; }
    }

    [Serializable]
    public class PushException : Exception
    {
        /// <summary>
        /// Exception fired when there's an error whilst pushing an event to the desired calendar
        /// </summary>
        /// <param name="message">Description of the error</param>
        /// <param name="failedEvent">Which event has failed</param>
        public PushException(string message, GenericEvent failedEvent)
            : base(message)
        {
            FailedEvent = failedEvent;
        }

        public GenericEvent FailedEvent { get; private set; }
    }

    public interface ICalendarManager
    {
        IEnumerable<UpdateOutcome> Push(ICalendar calendar);

        Task<IEnumerable<UpdateOutcome>> PushAsync(ICalendar calendar);

        ICalendar Pull(); // this gets all the events synchronously

        Task<ICalendar> PullAsync();  // this gets all the events asynchronously

        ICalendar Pull(DateTime from, DateTime to);

        Task<ICalendar> PullAsync(DateTime from, DateTime to);

        IEnumerable<ICalendarManager> Subscribers { get; set; }

        void StartLookingForChanges(TimeSpan updateInterval);
    }

    public abstract class GenericCalendarManager : ICalendarManager
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public abstract IEnumerable<UpdateOutcome> Push(ICalendar calendar);

        public abstract Task<IEnumerable<UpdateOutcome>> PushAsync(ICalendar calendar);

        public abstract ICalendar Pull(); // this gets all the events synchronously

        public abstract Task<ICalendar> PullAsync(); // this gets all the events asynchronously

        public abstract ICalendar Pull(DateTime from, DateTime to);

        public abstract Task<ICalendar> PullAsync(DateTime from, DateTime to);

        public IEnumerable<ICalendarManager> Subscribers { get; set; }

        public void StartLookingForChanges(TimeSpan updateInterval)
        {
            UpdateInterval = updateInterval;
            var timer = new Timer(LookForCalendarChanges);
            timer.Change(updateInterval, TimeSpan.FromMilliseconds(-1));
        }

        private async void LookForCalendarChanges(object state)
        {
            Log.Info("Updating calendar...");
            var t = (Timer) state;
            t.Dispose();
            await UpdateAsync();
            var timer = new Timer(LookForCalendarChanges);
            timer.Change(UpdateInterval, TimeSpan.FromMilliseconds(-1));
        }

        private Task UpdateAsync()
        {
            var t = Task.Factory.StartNew(async () =>
            {
                try
                {
                    var newCalendar = await PullAsync();
                    if (newCalendar != null && LastCalendar != null)
                    {
                        foreach (var subscriber in Subscribers)
                        {
                            var res = await subscriber.PushAsync(newCalendar);
                            Log.Debug(res);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("Exception", ex);
                }
            });
            return t;
        }

        protected void Update()
        {
            UpdateAsync().RunSynchronously();
        }

        protected ICalendar LastCalendar { get; set; }

        private TimeSpan UpdateInterval { get; set; }
    }
}