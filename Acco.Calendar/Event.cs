using System;
using System.Collections.Generic;

using Acco.Calendar.Location;
using Acco.Calendar.Person;

namespace Acco.Calendar.Event
{
    public interface IEvent
    {
        GenericPerson Creator { get; set; }
        DateTime? Created { get; set; }
        DateTime? LastModified { get; set; }
        DateTime? Start { get; set; }
        DateTime? End { get; set; }
        string Description { get; set; }
        GenericLocation Location { get; set; }
        bool? Recurrent { get; set; }
        List<GenericPerson> Attendees { get; set; }
    }

    public class GenericEvent : IEvent
    {
        public GenericPerson Creator { get; set; }
        public DateTime? Created { get; set; }
        public DateTime? LastModified { get; set; }
        public DateTime? Start { get; set; }
        public DateTime? End { get; set; }
        public string Description { get; set; }
        public GenericLocation Location { get; set; }
        public bool? Recurrent { get; set; }
        public List<GenericPerson> Attendees { get; set; }
    }
}
