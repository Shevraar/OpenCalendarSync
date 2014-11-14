using System.ComponentModel.DataAnnotations;
using OpenCalendarSync.Lib.Event;
using OpenCalendarSync.Lib.Person;

namespace OpenCalendarSync.Lib
{
    public interface ICalendar
    {
        string Id { get; set; }

        string Name { get; set; }

        GenericPerson Creator { get; set; }

        DbCollection<GenericEvent> Events { get; set; }
    }

    public class GenericCalendar : ICalendar
    {
        [Required(ErrorMessage = "This field is required")]
        public string Id { get; set; }

        public string Name { get; set; }

        public GenericPerson Creator { get; set; }

        public DbCollection<GenericEvent> Events { get; set; }
    }
}