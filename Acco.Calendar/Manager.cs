using Acco.Calendar.Event;
using System;
//
using System.Threading.Tasks;
//
//
//

namespace Acco.Calendar
{
    public class PushException : Exception
    {
        public PushException(string message, GenericEvent _FailedEvent) 
            : base(message)
        {
            FailedEvent = _FailedEvent;
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
}
