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
        /// </summary>

        public override void Parse<T>(T rules)
        {
            if (rules is Microsoft.Office.Interop.Outlook.RecurrencePattern)
            {
                Microsoft.Office.Interop.Outlook.RecurrencePattern outlookRecurrencePattern = rules as Microsoft.Office.Interop.Outlook.RecurrencePattern;
                switch (outlookRecurrencePattern.RecurrenceType)
                {
                    case OlRecurrenceType.olRecursDaily:
                        {
                            Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Daily, outlookRecurrencePattern.Interval);
                            Pattern.FirstDayOfWeek = System.DayOfWeek.Monday; // always monday, c.b.a. to change it for other counties.
                            Pattern.RestrictionType = RecurrenceRestrictionType.Default;
                            if (outlookRecurrencePattern.Occurrences > 0)
                                Pattern.Count = outlookRecurrencePattern.Occurrences;
                            else if (outlookRecurrencePattern.PatternEndDate > DateTime.Now)
                                Pattern.Until = outlookRecurrencePattern.PatternEndDate;
                        }
                        break;
                    case OlRecurrenceType.olRecursWeekly:
                        {
                            //iCalendar _icalendar = new iCalendar();
                            //_icalendar.AddLocalTimeZone();
                            //DDay.iCal.Event _event = new DDay.iCal.Event();
                            Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Weekly, outlookRecurrencePattern.Interval);
                            Pattern.FirstDayOfWeek = System.DayOfWeek.Monday; // always monday, c.b.a. to change it for other countries.
                            Pattern.RestrictionType = RecurrenceRestrictionType.NoRestriction;
                            if (outlookRecurrencePattern.Occurrences > 0)
                                Pattern.Count = outlookRecurrencePattern.Occurrences;
                            else if(outlookRecurrencePattern.PatternEndDate > DateTime.Now)
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
                            //_event.RecurrenceRules = new List<IRecurrencePattern>();
                            //_event.RecurrenceRules.Add(Pattern);
                            //_icalendar.Events.Add(_event);
                            //iCalendarSerializer serializer = new iCalendarSerializer(_icalendar);
                            //string _serializedicalendar = serializer.SerializeToString();
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
            else
            {
                throw new RecurrenceParseException("OutlookRecurrence: Unsupported type.", typeof(T));
            }
        }

        public override string Get()
        {
            string _modifiedPattern = Pattern.ToString();
            // todo: eventually modify this pattern to add timezones manually...
            _modifiedPattern = "RRULE:" + _modifiedPattern;
            // todo: find a way to set the time correctly
            //  "message": "Invalid value for: Invalid format: \"20140731T0000000000\" is malformed at \"T00000000000\""
            //RFC3339DateTime _rfc3339dt = new RFC3339DateTime(_modifiedPattern.Substring((x) => (x = 1), (y) => ( y = 2)));
            return _modifiedPattern;
        }
    }
}
