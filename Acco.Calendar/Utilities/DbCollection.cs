using System;
using System.Collections.ObjectModel;
using Acco.Calendar.Database;
using Acco.Calendar.Event;
using MongoDB.Driver.Builders;

namespace Acco.Calendar
{
    public interface IDbActions<in T> where T : IEvent
    {
        void Save(T item);
        void Delete(T item);
        bool IsAlreadySynced(T item);
    }

    public class DbCollection<T> : Collection<T>, IDbActions<T> where T: IEvent
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public void Save(T item)
        {
            var r = Storage.Instance.Appointments.Save(item);
            if (!r.Ok)
            {
                Log.Error(String.Format("[{0}] was not added", item.Id));
            }
            else
            {
                Log.Info(String.Format("[{0}] was added", item.Id));
            }
        }

        public void Delete(T item)
        {
            var query = Query<T>.EQ(e => e.Id, item.Id);
            var r = Storage.Instance.Appointments.Remove(query);
            if (!r.Ok)
            {
                Log.Warn(String.Format("[{0}] was not removed", item.Id));
            }
            else
            {
                Log.Info(String.Format("[{0}] was removed", item.Id));
            }
        }

        public bool IsAlreadySynced(T item)
        {
            bool isPresent;
            // first: check if the item has already been added to the shared database
            var query = Query<T>.EQ(x => x.Id, item.Id);
            var appointment = Storage.Instance.Appointments.FindOneAs<T>(query);
            if (appointment == null)
            {
                Log.Debug(String.Format("[{0}] was not found", item.Id));
                isPresent = false;
            }
            else 
            { 
                // todo: add item comparison here
                if(appointment as GenericEvent == item as GenericEvent)
                {
                    Log.Info(String.Format("[{0}] is a duplicate", item.Id));
                    isPresent = true;
                    item.Action = EventAction.Duplicate;
                }
                else
                {
                    isPresent = true;
                    Log.Info(String.Format("[{0}] has to be updated", item.Id));
                    item.Action = EventAction.Update;
                    //todo: more specifically -> update only needed fields.. and use the update function..
                    //var update = Update<T>.Set(x => x, item); 
                    var saveResult = Storage.Instance.Appointments.Save(item); // todo: parse the result and gg - also change 
                }
            }
            return isPresent;
        }

        protected override void InsertItem(int index, T item)
        {
            if(IsAlreadySynced(item) == false)
            {
                Save(item); 
                item.Action = EventAction.Add;
            }
            base.InsertItem(index, item);
        }

        protected override void RemoveItem(int index)
        {
            Items[index].Action = EventAction.Remove;
            Delete(Items[index]);
            base.RemoveItem(index);
        }
    }
}