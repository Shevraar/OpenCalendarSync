using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acco.Calendar.Manager
{

    public class OutlookCalendarManager : ICalendarManager
    {
        public bool Push(GenericCalendar calendar)
        {
            return false;
        }

        public GenericCalendar Pull()
        {
            return null;
        }

        public async Task<bool> PushAsync(GenericCalendar calendar)
        {
            return false;
        }

        public async Task<GenericCalendar> PullAsync()
        {
            return null;
        }
    }
}
