using Acco.Calendar.Event;
using System;

//
using System.Threading.Tasks;

//
using MongoDB.Driver.Builders;
using System.Collections.Generic;
using System.Threading;

namespace Acco.Calendar
{
    [Serializable]
    public class PushException : Exception
    {
        public PushException(string message, GenericEvent failedEvent)
            : base(message)
        {
            FailedEvent = failedEvent;
        }

        public GenericEvent FailedEvent { get; private set; }
    }

    public interface ICalendarManager
    {
        bool Push(ICalendar calendar);

        Task<bool> PushAsync(ICalendar calendar);

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

        public abstract bool Push(ICalendar calendar);

        public abstract Task<bool> PushAsync(ICalendar calendar);

        public abstract ICalendar Pull(); // this gets all the events synchronously

        public abstract Task<ICalendar> PullAsync(); // this gets all the events asynchronously

        public abstract ICalendar Pull(DateTime from, DateTime to);

        public abstract Task<ICalendar> PullAsync(DateTime from, DateTime to);

        public IEnumerable<ICalendarManager> Subscribers { get; set; }

        public void StartLookingForChanges(TimeSpan updateInterval)
        {
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
            timer.Change(TimeSpan.FromSeconds(30), TimeSpan.FromMilliseconds(-1));
        }

        protected internal Task UpdateAsync()
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

        protected internal ICalendar LastCalendar { get; set; }
    }
}