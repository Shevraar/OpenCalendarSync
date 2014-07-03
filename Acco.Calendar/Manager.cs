//
using System.Threading.Tasks;
//
//
//

namespace Acco.Calendar
{
    public interface ICalendarManager
    {
        bool Push(GenericCalendar calendar);
        Task<bool> PushAsync(GenericCalendar calendar);
        GenericCalendar Pull();
        Task<GenericCalendar> PullAsync();

    }


    
}
