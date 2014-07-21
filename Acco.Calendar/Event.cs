//
//
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using Acco.Calendar.Location;
using Acco.Calendar.Person;
using DDay.iCal;

//

namespace Acco.Calendar.Event
{
    public interface IRecurrence
    {
        void Parse<T>(T rules);
        string Get();
    }

    public class GenericRecurrence : IRecurrence
    {
        protected RecurrencePattern Pattern { get; set; }
        public virtual void Parse<T>(T rules) { throw new NotImplementedException(); }
        public virtual string Get() { throw new NotImplementedException(); }
        public override string ToString()
        {
            string s = "";
            if (Pattern != null)
            {
                s = Pattern.ToString();
            }
            return s;
        }
    }

    [Serializable]
    public class RecurrenceParseException : Exception
    {
        public RecurrenceParseException(string message, Type _TypeOfRule) :
            base(message)
        {
            TypeOfRule = _TypeOfRule;
        }

        public Type TypeOfRule { get; private set; }
    }

    public interface IEvent
    {
        [Required(ErrorMessage = "This field is required")]
        // todo: add other DataAnnotations validation stuff (such as min lenght and max lenght, etc)
        string Id { get; set; }

        GenericPerson Organizer { get; set; }
        GenericPerson Creator { get; set; }
        DateTime? Created { get; set; }
        DateTime? LastModified { get; set; }
        DateTime? Start { get; set; }
        DateTime? End { get; set; }

        [Required(ErrorMessage = "This field is required")]
        string Summary { get; set; }

        [Required(ErrorMessage = "This field is required")]
        string Description { get; set; }

        [Required(ErrorMessage = "This field is required")]
        GenericLocation Location { get; set; }

        GenericRecurrence Recurrence { get; set; }
        List<GenericPerson> Attendees { get; set; }
    }

    public class GenericEvent : IEvent
    {
        public GenericEvent(string Id)
        {
            this.Id = Id;
            Summary = "No Summary";
            Description = "No Description";
            Location = new GenericLocation {Name = "No Location"};
        }

        public GenericEvent(string Id, string Summary, string Description)
        {
            this.Id = Id;
            this.Summary = Summary;
            this.Description = Description;
            Location = new GenericLocation {Name = "No Location"};
        }

        public GenericEvent(string Id, string Summary, string Description, ILocation Location)
        {
            this.Id = Id;
            this.Summary = Summary;
            this.Description = Description;
            this.Location = Location as GenericLocation;
        }

        public string Id { get; set; }
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

        public override string ToString()
        {
            string eventString = "[";
            eventString += "Id: " + Id;
            eventString += "\n";
            eventString += "Summary: " + Summary;
            eventString += "\n";
            eventString += "Location: " + Location.Name;
            eventString += "]";
            return eventString;
        }
    }
}