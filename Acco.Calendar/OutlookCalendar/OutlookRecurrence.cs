using System;
//
using Microsoft.Office.Interop.Outlook;
using DDay.iCal;
using System.Collections.Generic;

namespace Acco.Calendar.Event
{
    class OutlookRecurrence : GenericRecurrence
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

        public void Parse(Microsoft.Office.Interop.Outlook.RecurrencePattern outlookRecurrencePattern)
        {
            switch(outlookRecurrencePattern.RecurrenceType)
            {
                case OlRecurrenceType.olRecursDaily:
                    {
                        //olRecursDaily            |  Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
                        Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Daily, outlookRecurrencePattern.Interval);
                        Pattern.FirstDayOfWeek = System.DayOfWeek.Monday; // always monday, c.b.a. to change it for other counties.
                        Pattern.RestrictionType = RecurrenceRestrictionType.Default;
                        Pattern.Count = outlookRecurrencePattern.Occurrences; // check if count is intended as duration...
                        Pattern.Until = outlookRecurrencePattern.PatternEndDate;   
                    }
                    break;
                case OlRecurrenceType.olRecursWeekly:
                    {
                        //olRecursWeekly           |  DayOfWeekMask, Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
                        Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Weekly, outlookRecurrencePattern.Interval);
                        Pattern.FirstDayOfWeek = System.DayOfWeek.Monday; // always monday, c.b.a. to change it for other counties.
                        Pattern.RestrictionType = RecurrenceRestrictionType.Default;
                        Pattern.Count = outlookRecurrencePattern.Occurrences; // check if count is intended as duration...
                        Pattern.Until = outlookRecurrencePattern.PatternEndDate;
                        Pattern.ByDay = new List<IWeekDay>();
                        // 
                        if (outlookRecurrencePattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olMonday))
                        { 
                            Pattern.ByDay.Add(new WeekDay() 
                            { 
                                DayOfWeek = System.DayOfWeek.Monday
                            });
                        }
                        if (outlookRecurrencePattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olTuesday))
                        {
                            Pattern.ByDay.Add(new WeekDay()
                            {
                                DayOfWeek = System.DayOfWeek.Tuesday
                            });
                        }
                        if (outlookRecurrencePattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olWednesday))
                        {
                            Pattern.ByDay.Add(new WeekDay()
                            {
                                DayOfWeek = System.DayOfWeek.Wednesday
                            });
                        }
                        if (outlookRecurrencePattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olThursday))
                        {
                            Pattern.ByDay.Add(new WeekDay()
                            {
                                DayOfWeek = System.DayOfWeek.Thursday
                            });
                        }
                        if (outlookRecurrencePattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olFriday))
                        {
                            Pattern.ByDay.Add(new WeekDay()
                            {
                                DayOfWeek = System.DayOfWeek.Friday
                            });
                        }
                        if (outlookRecurrencePattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olSaturday))
                        {
                            Pattern.ByDay.Add(new WeekDay()
                            {
                                DayOfWeek = System.DayOfWeek.Saturday
                            });
                        }
                        if (outlookRecurrencePattern.DayOfWeekMask.HasFlag(OlDaysOfWeek.olSunday))
                        {
                            Pattern.ByDay.Add(new WeekDay()
                            {
                                DayOfWeek = System.DayOfWeek.Sunday
                            });
                        }
                        //
                    }
                    break;
                case OlRecurrenceType.olRecursMonthly:
                case OlRecurrenceType.olRecursMonthNth:
                    {
                        //
                    }
                    break;
                case OlRecurrenceType.olRecursYearly:
                case OlRecurrenceType.olRecursYearNth:
                    {
                        //
                    }
                    break;
            }
        }

        public override string Get()
        {
            return Pattern.ToString();
        }

        public override void Parse(string format)
        {
            base.Parse(format);
        }
    }
}
