using System;
using System.Collections.Generic;
//
using Microsoft.Office.Interop.Outlook;
//
using DDay.iCal;
using DDay.iCal.Serialization;
using DDay.iCal.Serialization.iCalendar;

namespace Acco.Calendar.Event
{
    public class OutlookRecurrence : GenericRecurrence
    {
        /// <summary>
        /// Outlook Recurrence - Parse Microsoft.Office.Interop.Outlook.RecurrencePattern to DDay.ICal.RecurrencePattern
        /// </summary>
        /// <remarks>
        /// Directly from MSDN:
        /// 
        /// The following table shows the properties that are valid for the different recurrence types. 
        /// An error occurs if the item is saved and the property is null or contains an invalid value. 
        /// Monthly and yearly patterns are only valid for a single day. 
        /// Weekly patterns are only valid as the Or of the DayOfWeekMask...
        /// -----------------------------------------------------------------------------------------------------------------------------------------------------
        /// OlRecurrenceType value   |  Valid RecurrencePattern properties
        /// -----------------------------------------------------------------------------------------------------------------------------------------------------
        /// olRecursDaily            |  Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
        /// olRecursWeekly           |  DayOfWeekMask, Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
        /// olRecursMonthly          |  DayOfMonth, Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
        /// olRecursMonthNth         |  DayOfWeekMask, Duration, EndTime, Interval, Instance, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
        /// olRecursYearly           |  DayOfMonth, Duration, EndTime, Interval, MonthOfYear, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
        /// olRecursYearNth          |  DayOfWeekMask, Duration, EndTime, Interval, Instance, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
        /// -----------------------------------------------------------------------------------------------------------------------------------------------------
        ///  The Interval property is used when the AppointmentItem is to recur less often than every recurrence unit, 
        /// such as once every three days, once every two weeks, or once every six months. 
        /// Interval contains a value representing the frequency of recurrence in terms of recurrence units.
        /// Interval is always valid on a newly created RecurrencePattern object and defaults to 1, which is its minimum value. 
        /// Its maximum value depends on the setting of the RecurrenceType property as follows:
        /// ----------------------------------------------------------------------
        /// RecurrenceType setting                         |Maximum Interval value
        /// ----------------------------------------------------------------------
        /// CdoRecurTypeDaily                              |999
        /// CdoRecurTypeMonthly CdoRecurTypeMonthlyNth     |99
        /// CdoRecurTypeYearly CdoRecurTypeYearlyNth       |1
        /// CdoRecurTypeWeekly                             |99
        /// ----------------------------------------------------------------------
        /// Setting the Interval property causes CDO to force certain other recurrence pattern properties into conformance. 
        /// PatternEndDate is recalculated from PatternStartDate, Occurrences, and Interval. 
        /// If the resulting PatternEndDate is January 1, 4000 or later, NoEndDate is automatically reset to True, Occurrences is reset to 1,490,000, 
        /// PatternEndDate is reset to the month and day of PatternStartDate in the year 4001, and the recurrence pattern is considered to extend infinitely far into the future.
        /// Changes you make to properties on a RecurrencePattern object take effect when you call the underlying appointment's Send or Update method.
        /// </remarks>

        public override void Parse<T>(T rules)
        {
            if (rules is Microsoft.Office.Interop.Outlook.RecurrencePattern)
            {
                var outlookRP = rules as Microsoft.Office.Interop.Outlook.RecurrencePattern;
                switch (outlookRP.RecurrenceType)
                {
                    case OlRecurrenceType.olRecursDaily:
                    {
                        Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Daily, outlookRP.Interval);
                        Pattern.FirstDayOfWeek = System.DayOfWeek.Monday; // always monday, c.b.a. to change it for other counties.
                        Pattern.RestrictionType = RecurrenceRestrictionType.Default;
                        if (outlookRP.Occurrences > 0) { Pattern.Count = outlookRP.Occurrences; }
                        else if (outlookRP.PatternEndDate > DateTime.Now) { Pattern.Until = outlookRP.PatternEndDate; }
                    }
                    break;
                    case OlRecurrenceType.olRecursWeekly:
                    {
                        Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Weekly, outlookRP.Interval);
                        Pattern.FirstDayOfWeek = System.DayOfWeek.Monday; // always monday, c.b.a. to change it for other countries.
                        Pattern.RestrictionType = RecurrenceRestrictionType.NoRestriction;
                        if (outlookRP.Occurrences > 0) { Pattern.Count = outlookRP.Occurrences; }
                        else if (outlookRP.PatternEndDate > DateTime.Now) { Pattern.Until = outlookRP.PatternEndDate; }
                        Pattern.ByDay = extractDaysOfWeek(outlookRP.DayOfWeekMask);
                    }
                    break;
                    case OlRecurrenceType.olRecursMonthly:
                    {
                        // warning: untested
                        // see also: http://msdn.microsoft.com/en-us/library/office/aa211012(v=office.11).aspx 
                        Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Monthly, outlookRP.Interval);
                        Pattern.FirstDayOfWeek = System.DayOfWeek.Monday; // always monday, c.b.a. to change it for other countries.
                        Pattern.RestrictionType = RecurrenceRestrictionType.NoRestriction;
                        // how many times this is going to occur
                        if (outlookRP.Occurrences > 0) { Pattern.Count = outlookRP.Occurrences; }
                        else if (outlookRP.DayOfMonth > 0) { Pattern.ByMonthDay = new List<int> { outlookRP.DayOfMonth }; }
                    }
                    break;
                    case OlRecurrenceType.olRecursMonthNth:
                    {
                        // see also: http://msdn.microsoft.com/en-us/library/office/aa211012(v=office.11).aspx 
                        Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Monthly, outlookRP.Interval);
                        Pattern.FirstDayOfWeek = System.DayOfWeek.Monday; // always monday, c.b.a. to change it for other countries.
                        Pattern.RestrictionType = RecurrenceRestrictionType.NoRestriction;
                        // how many times this is going to occur
                        if (outlookRP.Occurrences > 0) { Pattern.Count = outlookRP.Occurrences; }
                        // instance states e.g.: "The Nth Tuesday"
                        if (outlookRP.Instance > 0) { Pattern.BySetPosition = new List<int> { outlookRP.Instance }; }
                        if (outlookRP.DayOfWeekMask > 0) { Pattern.ByDay = extractDaysOfWeek(outlookRP.DayOfWeekMask); }
                    }
                    break;
                    case OlRecurrenceType.olRecursYearly:
                    {
                        // warning: untested
                        Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Yearly, outlookRP.Interval);
                        Pattern.FirstDayOfWeek = System.DayOfWeek.Monday; // always monday, c.b.a. to change it for other countries.
                        Pattern.RestrictionType = RecurrenceRestrictionType.NoRestriction;
                        if (outlookRP.Occurrences > 0) { Pattern.Count = outlookRP.Occurrences; }
                        else if (outlookRP.PatternEndDate > DateTime.Now) { Pattern.Until = outlookRP.PatternEndDate; }
                        if (outlookRP.DayOfWeekMask > 0) { Pattern.ByDay = extractDaysOfWeek(outlookRP.DayOfWeekMask); }
                        else if (outlookRP.DayOfMonth > 0) { Pattern.ByMonthDay = new List<int> { outlookRP.DayOfMonth }; }
                    }
                    break;
                    case OlRecurrenceType.olRecursYearNth:
                    {
                        // warning: untested
                        Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Yearly, outlookRP.Interval);
                        Pattern.FirstDayOfWeek = System.DayOfWeek.Monday; // always monday, c.b.a. to change it for other countries.
                        Pattern.RestrictionType = RecurrenceRestrictionType.NoRestriction;
                        if (outlookRP.Occurrences > 0) { Pattern.Count = outlookRP.Occurrences; }
                        else if (outlookRP.PatternEndDate > DateTime.Now) { Pattern.Until = outlookRP.PatternEndDate; }
                        if (outlookRP.DayOfWeekMask > 0) { Pattern.ByDay = extractDaysOfWeek(outlookRP.DayOfWeekMask); }
                        else if (outlookRP.DayOfMonth > 0) { Pattern.ByMonthDay = new List<int> { outlookRP.DayOfMonth }; }
                    }
                    break;
                }
            }
            else
            {
                throw new RecurrenceParseException("OutlookRecurrence: Unsupported type.", typeof(T));
            }
        }

        public override string Get()
        {
            var _modifiedPattern = Pattern.ToString();
            _modifiedPattern = "RRULE:" + _modifiedPattern;
            // todo: find a way to set the time correctly
            // note: error from google: "message": "Invalid value for: Invalid format: \"20140731T0000000000\" is malformed at \"T00000000000\""
            return _modifiedPattern;
        }

        private List<IWeekDay> extractDaysOfWeek(OlDaysOfWeek outlookDaysOfWeekMask)
        {
            var myDaysOfWeekMask =  new List<IWeekDay>();
            // 
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olMonday))
            {
                myDaysOfWeekMask.Add(new WeekDay()
                {
                    DayOfWeek = System.DayOfWeek.Monday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olTuesday))
            {
                myDaysOfWeekMask.Add(new WeekDay()
                {
                    DayOfWeek = System.DayOfWeek.Tuesday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olWednesday))
            {
                myDaysOfWeekMask.Add(new WeekDay()
                {
                    DayOfWeek = System.DayOfWeek.Wednesday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olThursday))
            {
                myDaysOfWeekMask.Add(new WeekDay()
                {
                    DayOfWeek = System.DayOfWeek.Thursday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olFriday))
            {
                myDaysOfWeekMask.Add(new WeekDay()
                {
                    DayOfWeek = System.DayOfWeek.Friday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olSaturday))
            {
                myDaysOfWeekMask.Add(new WeekDay()
                {
                    DayOfWeek = System.DayOfWeek.Saturday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olSunday))
            {
                myDaysOfWeekMask.Add(new WeekDay()
                {
                    DayOfWeek = System.DayOfWeek.Sunday
                });
            }
            return myDaysOfWeekMask;
        }
    }
}
