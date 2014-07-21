using Acco.Calendar.Event;
using Acco.Calendar.Person;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations;

namespace Acco.Calendar
{
    public interface ICalendar
    {
        string Id { get; set; }

        string Name { get; set; }

        GenericPerson Creator { get; set; }

        ObservableCollection<GenericEvent> Events { get; set; }
    }

    public class GenericCalendar : ICalendar
    {
        [Required(ErrorMessage = "This field is required")]
        public string Id { get; set; }

        public string Name { get; set; }

        public GenericPerson Creator { get; set; }

        public ObservableCollection<GenericEvent> Events { get; set; }
    }
}