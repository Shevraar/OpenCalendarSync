using Acco.Calendar.Database;
using Acco.Calendar.Event;
using MongoDB.Bson;
using MongoDB.Driver.Builders;
using System;
using System.Collections.Specialized;

//
using System.Threading.Tasks;

//

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
    }

    public abstract class GenericCalendarManager : ICalendarManager
    {
        public abstract bool Push(ICalendar calendar);

        public abstract Task<bool> PushAsync(ICalendar calendar);

        public abstract ICalendar Pull(); // this gets all the events synchronously

        public abstract Task<ICalendar> PullAsync();  // this gets all the events asynchronously

        public abstract ICalendar Pull(DateTime from, DateTime to);

        public abstract Task<ICalendar> PullAsync(DateTime from, DateTime to);
    }
}