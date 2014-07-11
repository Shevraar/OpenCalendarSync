using System.Threading.Tasks;
using NetOffice.Outlook;

namespace Acco.Calendar.Manager
{

    public sealed class OutlookCalendarManager : ICalendarManager
    {
        #region Singleton directives
        private static readonly OutlookCalendarManager instance = new OutlookCalendarManager();
        // hidden constructor
        private OutlookCalendarManager() { }

        public static OutlookCalendarManager Instance { get { return instance; } }
        #endregion

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
            return Push(calendar);
        }

        public async Task<GenericCalendar> PullAsync()
        {
            return Pull();
        }
    }
}
