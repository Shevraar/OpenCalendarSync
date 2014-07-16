using System;
using System.Collections.Generic;
//
using Acco.Calendar.Location;
using Acco.Calendar.Person;

namespace Acco.Calendar.Event
{
    public enum RecurrencyType : ushort
    {
        DAILY = 0,
        WEEKLY,
        MONTHLY,
        YEARLY
    }

    [FlagsAttribute]
    public enum DayOfWeek : short
    {
        None = -64,
        Monday = 1,
        Tuesday = 2,
        Wednesday = 4,
        Thursday = 8,
        Friday = 16,
        Saturday = 32,
        Sunday = 64
    }

    public interface IRecurrence
    {
        RecurrencyType Type { get; set; }
        DateTime? Expiry { get; set; }
        DayOfWeek Days { get; set; }
        void Parse(string format);
        string Get();
    }

    public class GenericRecurrence : IRecurrence
    {
        public RecurrencyType Type { get; set; }
        public DateTime? Expiry { get; set; }
        public DayOfWeek Days { get; set; }
        public virtual void Parse(string format) { return; }
        public virtual string Get() { return "";  }
    }

    public interface IEvent
    {
        GenericPerson Organizer { get; set; }
        GenericPerson Creator { get; set; }
        DateTime? Created { get; set; }
        DateTime? LastModified { get; set; }
        DateTime? Start { get; set; }
        DateTime? End { get; set; }
        string Summary { get; set; }
        string Description { get; set; }
        GenericLocation Location { get; set; }
        GenericRecurrence Recurrence { get; set; }
        List<GenericPerson> Attendees { get; set; }
    }

    public class GenericEvent : IEvent
    {
        public GenericPerson Organizer { get; set; }
        public GenericPerson Creator { get; set; }
        public DateTime? Created { get; set; }
        public DateTime? LastModified { get; set; }
        public DateTime? Start { get; set; }
        public DateTime? End { get; set; }
        public string Summary { get; set; }
        public string Description { get; set; }
        public GenericLocation Location { get; set; }
        public GenericRecurrence Recurrence { get; set; }
        public List<GenericPerson> Attendees { get; set; }
    }
}
