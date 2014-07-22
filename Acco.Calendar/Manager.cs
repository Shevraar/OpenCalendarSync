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

        protected bool AlreadySynced(IEvent evt)
        {
            // first: check if the item has already been added to the shared database
            var query = Query<GenericEvent>.EQ(x => x.Id, evt.Id);
            if (Storage.Instance.Appointments.FindOneAs<GenericEvent>(query) == null)
            {
                return false;
            }
            // 
            Console.BackgroundColor = ConsoleColor.Yellow;
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("Event [{0}] is already present on database", evt.Id);
            Console.ResetColor();
            return true;
        }

        protected void Events_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Add)
            {
                // note: to know which item was added, use NewItems.
                foreach (GenericEvent item in e.NewItems) //todo: check if its possible to add the list of added events
                {
                    var r = Storage.Instance.Appointments.Save(item);
                    if (!r.Ok)
                    {
                        Console.BackgroundColor = ConsoleColor.Red; // add these in utils.
                        Console.ForegroundColor = ConsoleColor.White;
                        // add these in utils. (Utilities.Warning(...) - Utilities.Error(...) - Utilities.Info(...)
                        Console.WriteLine("Event [{0}] was not added", item.Id);
                        Console.ResetColor();
                    }
                    else
                    {
                        Console.BackgroundColor = ConsoleColor.Green; // add these in utils.
                        Console.ForegroundColor = ConsoleColor.Black;
                        // add these in utils. (Utilities.Warning(...) - Utilities.Error(...) - Utilities.Info(...)
                        Console.WriteLine("Event [{0}] added", item.Id);
                        Console.ResetColor();
                    }
                }
            }
            else if (e.Action == NotifyCollectionChangedAction.Remove)
            {
                foreach (GenericEvent item in e.OldItems)
                //todo: check if its possible to delete the list of removed events
                {
                    Console.WriteLine("Event [{0}] removed", item.Id);
                    var query = Query<GenericEvent>.EQ(evt => evt.Id, item.Id);
                    Storage.Instance.Appointments.Remove(query);
                }
            }
            else
            {
                throw new Exception("Unmanaged Action => " + e.Action);
            }
        }
    }
}