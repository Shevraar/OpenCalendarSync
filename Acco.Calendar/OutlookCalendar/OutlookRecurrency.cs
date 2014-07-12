using System;
//
using Microsoft.Office.Interop.Outlook;

namespace Acco.Calendar.Event
{
    class OutlookRecurrency : GenericRecurrency
    {
        /*
            The following table shows the properties that are valid for the different recurrence types. 
            An error occurs if the item is saved and the property is null or contains an invalid value. 
            Monthly and yearly patterns are only valid for a single day. 
            Weekly patterns are only valid as the Or of the DayOfWeekMask...
            -----------------------------------------------------------------------------------------------------------------------------------------------------
            OlRecurrenceType value   |  Valid RecurrencePattern properties
            -----------------------------------------------------------------------------------------------------------------------------------------------------
            olRecursDaily            |  Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
            olRecursWeekly           |  DayOfWeekMask, Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
            olRecursMonthly          |  DayOfMonth, Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
            olRecursMonthNth         |  DayOfWeekMask, Duration, EndTime, Interval, Instance, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
            olRecursYearly           |  DayOfMonth, Duration, EndTime, Interval, MonthOfYear, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
            olRecursYearNth          |  DayOfWeekMask, Duration, EndTime, Interval, Instance, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
            -----------------------------------------------------------------------------------------------------------------------------------------------------
        */
        /*
            The Interval property is used when the AppointmentItem is to recur less often than every recurrence unit, 
            such as once every three days, once every two weeks, or once every six months. 
            Interval contains a value representing the frequency of recurrence in terms of recurrence units.
            Interval is always valid on a newly created RecurrencePattern object and defaults to 1, which is its minimum value. 
            Its maximum value depends on the setting of the RecurrenceType property as follows:
            ----------------------------------------------------------------------
            RecurrenceType setting                         |Maximum Interval value
            ----------------------------------------------------------------------
            CdoRecurTypeDaily                              |999
            CdoRecurTypeMonthly CdoRecurTypeMonthlyNth     |99
            CdoRecurTypeYearly CdoRecurTypeYearlyNth       |1
            CdoRecurTypeWeekly                             |99
            ----------------------------------------------------------------------
            Setting the Interval property causes CDO to force certain other recurrence pattern properties into conformance. 
            PatternEndDate is recalculated from PatternStartDate, Occurrences, and Interval. 
            If the resulting PatternEndDate is January 1, 4000 or later, NoEndDate is automatically reset to True, Occurrences is reset to 1,490,000, 
            PatternEndDate is reset to the month and day of PatternStartDate in the year 4001, and the recurrence pattern is considered to extend infinitely far into the future.
            Changes you make to properties on a RecurrencePattern object take effect when you call the underlying appointment's Send or Update method.
         */

        public void Parse(RecurrencePattern recurrencePattern)
        {
            switch(recurrencePattern.RecurrenceType)
            {
                case OlRecurrenceType.olRecursDaily:
                    ParseDailyRecurrency(recurrencePattern);
                    break;
                case OlRecurrenceType.olRecursWeekly:
                    ParseWeeklyRecurrency(recurrencePattern);
                    break;
                case OlRecurrenceType.olRecursMonthly:
                case OlRecurrenceType.olRecursMonthNth:
                    ParseMonthlyRecurrency(recurrencePattern);
                    break;
                case OlRecurrenceType.olRecursYearly:
                case OlRecurrenceType.olRecursYearNth:
                    ParseYearlyRecurrency(recurrencePattern);
                    break;
            }
        }

        public void ParseDailyRecurrency(RecurrencePattern recurrencePattern)
        {
            //olRecursDaily            |  Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
            Type = RecurrencyType.DAILY;
        }

        public void ParseWeeklyRecurrency(RecurrencePattern recurrencePattern)
        {
            //olRecursWeekly           |  DayOfWeekMask, Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
            // the weekly pattern is treated as a daily pattern with daysofweek set... fucking microsoft.
            Type = RecurrencyType.DAILY;
            Parse(recurrencePattern.DayOfWeekMask);
        }

        public void ParseMonthlyRecurrency(RecurrencePattern recurrencePattern)
        {
            // not yet implemented
        }

        public void ParseYearlyRecurrency(RecurrencePattern recurrencePattern)
        {
            // not yet implemented
        }

        public void Parse(OlDaysOfWeek daysOfWeek)
        {
            if (daysOfWeek.HasFlag(OlDaysOfWeek.olMonday))
                Days |= DayOfWeek.Monday;
            if (daysOfWeek.HasFlag(OlDaysOfWeek.olTuesday))
                Days |= DayOfWeek.Tuesday;
            if (daysOfWeek.HasFlag(OlDaysOfWeek.olWednesday))
                Days |= DayOfWeek.Wednesday;
            if (daysOfWeek.HasFlag(OlDaysOfWeek.olThursday))
                Days |= DayOfWeek.Thursday;
            if (daysOfWeek.HasFlag(OlDaysOfWeek.olFriday))
                Days |= DayOfWeek.Friday;
            if (daysOfWeek.HasFlag(OlDaysOfWeek.olSaturday))
                Days |= DayOfWeek.Saturday;
            if (daysOfWeek.HasFlag(OlDaysOfWeek.olSunday))
                Days |= DayOfWeek.Sunday;
        }
    }
}
