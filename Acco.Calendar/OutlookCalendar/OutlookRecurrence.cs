//
using DDay.iCal;

//
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using OutlookRecurrencePattern = Microsoft.Office.Interop.Outlook.RecurrencePattern;

namespace Acco.Calendar.Event
{
    public class OutlookRecurrence : GenericRecurrence
    {
        public override void Parse<T>(T rules)
        {
            if (rules is OutlookRecurrencePattern)
            {
                var outlookRp = rules as OutlookRecurrencePattern;
                switch (outlookRp.RecurrenceType)
                {
                    case OlRecurrenceType.olRecursDaily:
                        {
                            Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Daily, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = RecurrenceRestrictionType.Default
                            };
                            if (outlookRp.Occurrences > 0) { Pattern.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { Pattern.Until = outlookRp.PatternEndDate; }
                        }
                        break;

                    case OlRecurrenceType.olRecursWeekly:
                        {
                            Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Weekly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = RecurrenceRestrictionType.NoRestriction
                            };
                            if (outlookRp.Occurrences > 0) { Pattern.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { Pattern.Until = outlookRp.PatternEndDate; }
                            Pattern.ByDay = extractDaysOfWeek(outlookRp.DayOfWeekMask);
                        }
                        break;

                    case OlRecurrenceType.olRecursMonthly:
                        {
                            // warning: untested
                            // see also: http://msdn.microsoft.com/en-us/library/office/aa211012(v=office.11).aspx
                            Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Monthly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = RecurrenceRestrictionType.NoRestriction
                            };
                            // how many times this is going to occur
                            if (outlookRp.Occurrences > 0) { Pattern.Count = outlookRp.Occurrences; }
                            else if (outlookRp.DayOfMonth > 0) { Pattern.ByMonthDay = new List<int> { outlookRp.DayOfMonth }; }
                        }
                        break;

                    case OlRecurrenceType.olRecursMonthNth:
                        {
                            // see also: http://msdn.microsoft.com/en-us/library/office/aa211012(v=office.11).aspx
                            Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Monthly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = RecurrenceRestrictionType.NoRestriction
                            };
                            // how many times this is going to occur
                            if (outlookRp.Occurrences > 0) { Pattern.Count = outlookRp.Occurrences; }
                            // instance states e.g.: "The Nth Tuesday"
                            if (outlookRp.Instance > 0) { Pattern.BySetPosition = new List<int> { outlookRp.Instance }; }
                            if (outlookRp.DayOfWeekMask > 0) { Pattern.ByDay = extractDaysOfWeek(outlookRp.DayOfWeekMask); }
                        }
                        break;

                    case OlRecurrenceType.olRecursYearly:
                        {
                            // warning: untested
                            Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Yearly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = RecurrenceRestrictionType.NoRestriction
                            };
                            if (outlookRp.Occurrences > 0) { Pattern.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { Pattern.Until = outlookRp.PatternEndDate; }
                            if (outlookRp.DayOfWeekMask > 0) { Pattern.ByDay = extractDaysOfWeek(outlookRp.DayOfWeekMask); }
                            else if (outlookRp.DayOfMonth > 0) { Pattern.ByMonthDay = new List<int> { outlookRp.DayOfMonth }; }
                        }
                        break;

                    case OlRecurrenceType.olRecursYearNth:
                        {
                            // warning: untested
                            Pattern = new DDay.iCal.RecurrencePattern(FrequencyType.Yearly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = RecurrenceRestrictionType.NoRestriction
                            };
                            if (outlookRp.Occurrences > 0) { Pattern.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { Pattern.Until = outlookRp.PatternEndDate; }
                            if (outlookRp.DayOfWeekMask > 0) { Pattern.ByDay = extractDaysOfWeek(outlookRp.DayOfWeekMask); }
                            else if (outlookRp.DayOfMonth > 0) { Pattern.ByMonthDay = new List<int> { outlookRp.DayOfMonth }; }
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
            var modifiedPattern = Pattern.ToString();
            modifiedPattern = "RRULE:" + modifiedPattern;
            // todo: find a way to set the time correctly
            // note: error from google: "message": "Invalid value for: Invalid format: \"20140731T0000000000\" is malformed at \"T00000000000\""
            return modifiedPattern;
        }

        private List<IWeekDay> extractDaysOfWeek(OlDaysOfWeek outlookDaysOfWeekMask)
        {
            var myDaysOfWeekMask = new List<IWeekDay>();
            //
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olMonday))
            {
                myDaysOfWeekMask.Add(new WeekDay
                {
                    DayOfWeek = DayOfWeek.Monday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olTuesday))
            {
                myDaysOfWeekMask.Add(new WeekDay
                {
                    DayOfWeek = DayOfWeek.Tuesday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olWednesday))
            {
                myDaysOfWeekMask.Add(new WeekDay
                {
                    DayOfWeek = DayOfWeek.Wednesday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olThursday))
            {
                myDaysOfWeekMask.Add(new WeekDay
                {
                    DayOfWeek = DayOfWeek.Thursday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olFriday))
            {
                myDaysOfWeekMask.Add(new WeekDay
                {
                    DayOfWeek = DayOfWeek.Friday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olSaturday))
            {
                myDaysOfWeekMask.Add(new WeekDay
                {
                    DayOfWeek = DayOfWeek.Saturday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olSunday))
            {
                myDaysOfWeekMask.Add(new WeekDay
                {
                    DayOfWeek = DayOfWeek.Sunday
                });
            }
            return myDaysOfWeekMask;
        }
    }
}