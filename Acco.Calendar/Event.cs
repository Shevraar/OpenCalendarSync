using System;
using System.Collections.Generic;

using Acco.Calendar.Location;
using Acco.Calendar.Person;

namespace Acco.Calendar.Event
{
    public enum RecurrencyType
    {
        DAILY = 0,
        WEEKLY,
        MONTHLY,
        YEARLY
    }

    public interface IRecurrency
    {
        RecurrencyType Type { get; set; }
        DateTime? Expiry { get; set; }
    }

    public class GenericRecurrency : IRecurrency
    {
        public RecurrencyType Type { get; set; }
        public DateTime? Expiry { get; set; }
    }

    public class GoogleRecurrency : GenericRecurrency
    {
        public override string ToString()
        {
            string ret = "";
            string offset = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Today).ToString("-hh:mm");
            // yyyymmddThhmmss-OFFSET
            // 20110701T100000-07:00
            ret = @"RRULE:FREQ=" + Type.ToString() + ";UNTIL=" + Expiry.Value.ToString("yyyymmddThhmmss") + offset;
            return ret;
        }
    }

    public interface IEvent
    {
        GenericPerson Creator { get; set; }
        DateTime? Created { get; set; }
        DateTime? LastModified { get; set; }
        DateTime? Start { get; set; }
        DateTime? End { get; set; }
        string Summary { get; set; }
        string Description { get; set; }
        GenericLocation Location { get; set; }
        GenericRecurrency Recurrency { get; set; }
        List<GenericPerson> Attendees { get; set; }
    }

    public class GenericEvent : IEvent
    {
        public GenericPerson Creator { get; set; }
        public DateTime? Created { get; set; }
        public DateTime? LastModified { get; set; }
        public DateTime? Start { get; set; }
        public DateTime? End { get; set; }
        public string Summary { get; set; }
        public string Description { get; set; }
        public GenericLocation Location { get; set; }
        public GenericRecurrency Recurrency { get; set; }
        public List<GenericPerson> Attendees { get; set; }
    }
}
