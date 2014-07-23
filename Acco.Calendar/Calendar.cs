using Acco.Calendar.Database;
using Acco.Calendar.Event;
using Acco.Calendar.Person;
using MongoDB.Driver.Builders;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations;

namespace Acco.Calendar
{
    public interface IDBActions<T> where T: IEvent
    {
        void Save(T item);
        void Delete(T item);
        bool IsAlreadySynced(T item);
    }

    public class DBCollection<T> : Collection<T>, IDBActions<T> where T: IEvent
    {
        public void Save(T item)
        {
            var r = Storage.Instance.Appointments.Save(item);
            if (!r.Ok)
            {
                Console.BackgroundColor = ConsoleColor.Red; // add these in utils.
                Console.ForegroundColor = ConsoleColor.White; // add these in utils. (Utilities.Warning(...) - Utilities.Error(...) - Utilities.Info(...)
                Console.WriteLine("Event [{0}] was not added", item.Id);
                Console.ResetColor();
            }
            else
            {
                Console.BackgroundColor = ConsoleColor.Green; // add these in utils.
                Console.ForegroundColor = ConsoleColor.Black; // add these in utils. (Utilities.Warning(...) - Utilities.Error(...) - Utilities.Info(...)
                Console.WriteLine("Event [{0}] added", item.Id);
                Console.ResetColor();
            }
        }

        public void Delete(T item)
        {
            var query = Query<T>.EQ(e => e.Id, item.Id);
            var r = Storage.Instance.Appointments.Remove(query);
            if (!r.Ok)
            {
                Console.BackgroundColor = ConsoleColor.Yellow; // add these in utils.
                Console.ForegroundColor = ConsoleColor.Black; // add these in utils. (Utilities.Warning(...) - Utilities.Error(...) - Utilities.Info(...)
                Console.WriteLine("Event [{0}] was not removed", item.Id);
                Console.ResetColor();
            }
            else
            {
                Console.BackgroundColor = ConsoleColor.Red; // add these in utils.
                Console.ForegroundColor = ConsoleColor.White; // add these in utils. (Utilities.Warning(...) - Utilities.Error(...) - Utilities.Info(...)
                Console.WriteLine("Event [{0}] removed", item.Id);
                Console.ResetColor();
            }
        }

        public bool IsAlreadySynced(T item)
        {
            var isPresent = false;
            // first: check if the item has already been added to the shared database
            var query = Query<T>.EQ(x => x.Id, item.Id);
            if (Storage.Instance.Appointments.FindOneAs<T>(query) == null)
            {
                isPresent = false;
            }
            else 
            { 
                Console.BackgroundColor = ConsoleColor.Yellow;
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine("Event [{0}] is already present on database", item.Id);
                Console.ResetColor();
                isPresent = true;
            }
            return isPresent;
        }

        protected override void InsertItem(int index, T item)
        {
            if(IsAlreadySynced(item) == false)
            {
                Save(item); 
                base.InsertItem(index, item);
            }
        }

        protected override void RemoveItem(int index)
        {
            Delete(Items[index]);
            base.RemoveItem(index);
        }
    }

    public interface ICalendar
    {
        string Id { get; set; }

        string Name { get; set; }

        GenericPerson Creator { get; set; }

        DBCollection<GenericEvent> Events { get; set; }
    }

    public class GenericCalendar : ICalendar
    {
        [Required(ErrorMessage = "This field is required")]
        public string Id { get; set; }

        public string Name { get; set; }

        public GenericPerson Creator { get; set; }

        public DBCollection<GenericEvent> Events { get; set; }
    }
}