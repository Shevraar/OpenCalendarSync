using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations;
//
using Acco.Calendar.Location;
using Acco.Calendar.Person;
//
using DDay;
using DDay.iCal;
using DDay.Collections;
using DDay.iCal.Serialization;
using Acco.Calendar.Utilities;
//

namespace Acco.Calendar.Event
{
    public interface IRecurrence
    {
        RecurrencePattern Pattern { get; set; }
        void Parse<T>(T rules);
        string Get();
    }

    public abstract class GenericRecurrence : IRecurrence
    {
        public RecurrencePattern Pattern { get; set; }
        public abstract void Parse<T>(T rules);
        public abstract string Get();
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
        [Required(ErrorMessage = "This field is required")] // todo: add other DataAnnotations validation stuff (such as min lenght and max lenght, etc)
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
        private GenericEvent()
        {
            // have comment this exception, until I understand how MongoDB de-serialization works...
            //throw new Exception("You shouldn't invoke this constructor");
        }
        public GenericEvent(string Id) 
        { 
            this.Id = Id;
            Summary = "No Summary";
            Description = "No Description";
            Location = new GenericLocation { Name = "No Location" };
        }
        public GenericEvent(string Id, string Summary, string Description)
        {
            this.Id = Id;
            this.Summary = Summary;
            this.Description = Description;
            Location = new GenericLocation { Name = "No Location" };
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
            var eventString = "[";
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
