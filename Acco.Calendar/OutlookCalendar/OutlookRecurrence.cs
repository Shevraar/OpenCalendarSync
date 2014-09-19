//
using RecPatt = DDay.iCal.RecurrencePattern;
using iCal = DDay.iCal;

//
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using OutlookRecurrencePattern = Microsoft.Office.Interop.Outlook.RecurrencePattern;

namespace Acco.Calendar.Event
{
    public class OutlookRecurrence : GenericRecurrence
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public override void Parse<T>(T rules)
        {
            if (rules is OutlookRecurrencePattern)
            {
                Log.Info("Parsing OutlookRecurrencePattern...");
                var outlookRp = rules as OutlookRecurrencePattern;
                Log.Info(String.Format("Recurrence type is [{0}]", outlookRp.RecurrenceType));
                switch (outlookRp.RecurrenceType)
                {
                    case OlRecurrenceType.olRecursDaily:
                        {
                            _RecPatt = new RecPatt(iCal.FrequencyType.Daily, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.Default
                            };
                            if (outlookRp.Occurrences > 0) { _RecPatt.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { _RecPatt.Until = outlookRp.PatternEndDate; }
                        }
                        break;

                    case OlRecurrenceType.olRecursWeekly:
                        {
                            _RecPatt = new RecPatt(iCal.FrequencyType.Weekly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.NoRestriction
                            };
                            if (outlookRp.Occurrences > 0) { _RecPatt.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { _RecPatt.Until = outlookRp.PatternEndDate; }
                            _RecPatt.ByDay = ExtractDaysOfWeek(outlookRp.DayOfWeekMask);
                        }
                        break;

                    case OlRecurrenceType.olRecursMonthly:
                        {
                            // warning: untested
                            // see also: http://msdn.microsoft.com/en-us/library/office/aa211012(v=office.11).aspx
                            _RecPatt = new RecPatt(iCal.FrequencyType.Monthly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.NoRestriction
                            };
                            // how many times this is going to occur
                            if (outlookRp.Occurrences > 0) { _RecPatt.Count = outlookRp.Occurrences; }
                            else if (outlookRp.DayOfMonth > 0) { _RecPatt.ByMonthDay = new List<int> { outlookRp.DayOfMonth }; }
                        }
                        break;

                    case OlRecurrenceType.olRecursMonthNth:
                        {
                            // see also: http://msdn.microsoft.com/en-us/library/office/aa211012(v=office.11).aspx
                            _RecPatt = new DDay.iCal.RecurrencePattern(iCal.FrequencyType.Monthly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.NoRestriction
                            };
                            // how many times this is going to occur
                            if (outlookRp.Occurrences > 0) { _RecPatt.Count = outlookRp.Occurrences; }
                            // instance states e.g.: "The Nth Tuesday"
                            if (outlookRp.Instance > 0) { _RecPatt.BySetPosition = new List<int> { outlookRp.Instance }; }
                            if (outlookRp.DayOfWeekMask > 0) { _RecPatt.ByDay = ExtractDaysOfWeek(outlookRp.DayOfWeekMask); }
                        }
                        break;

                    case OlRecurrenceType.olRecursYearly:
                        {
                            // warning: untested
                            _RecPatt = new RecPatt(iCal.FrequencyType.Yearly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.NoRestriction
                            };
                            if (outlookRp.Occurrences > 0) { _RecPatt.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { _RecPatt.Until = outlookRp.PatternEndDate; }
                            if (outlookRp.DayOfWeekMask > 0) { _RecPatt.ByDay = ExtractDaysOfWeek(outlookRp.DayOfWeekMask); }
                            else if (outlookRp.DayOfMonth > 0) { _RecPatt.ByMonthDay = new List<int> { outlookRp.DayOfMonth }; }
                        }
                        break;

                    case OlRecurrenceType.olRecursYearNth:
                        {
                            // warning: untested
                            _RecPatt = new RecPatt(iCal.FrequencyType.Yearly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.NoRestriction
                            };
                            if (outlookRp.Occurrences > 0) { _RecPatt.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { _RecPatt.Until = outlookRp.PatternEndDate; }
                            if (outlookRp.DayOfWeekMask > 0) { _RecPatt.ByDay = ExtractDaysOfWeek(outlookRp.DayOfWeekMask); }
                            else if (outlookRp.DayOfMonth > 0) { _RecPatt.ByMonthDay = new List<int> { outlookRp.DayOfMonth }; }
                        }
                        break;
                }
                Log.Debug(String.Format("iCalendar recurrence pattern is [{0}]", _RecPatt));
            }
            else
            {
                throw new RecurrenceParseException("OutlookRecurrence: Unsupported type.", typeof(T));
            }
        }

        public override string Get()
        {
            var modifiedPattern = _RecPatt.ToString();
            modifiedPattern = "RRULE:" + modifiedPattern;
            // todo: find a way to set the time correctly
            // note: error from google: "message": "Invalid value for: Invalid format: \"20140731T0000000000\" is malformed at \"T00000000000\""
            return modifiedPattern;
        }

        private static List<iCal.IWeekDay> ExtractDaysOfWeek(OlDaysOfWeek outlookDaysOfWeekMask)
        {
            Log.Info(String.Format("Extracting days of week from outlookDaysOfWeekMask [{0}]", outlookDaysOfWeekMask));
            var myDaysOfWeekMask = new List<iCal.IWeekDay>();
            //
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olMonday))
            {
                Log.Debug("Monday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Monday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olTuesday))
            {
                Log.Debug("Tuesday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Tuesday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olWednesday))
            {
                Log.Debug("Wednesday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Wednesday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olThursday))
            {
                Log.Debug("Thursday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Thursday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olFriday))
            {
                Log.Debug("Friday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Friday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olSaturday))
            {
                Log.Debug("Saturday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Saturday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olSunday))
            {
                Log.Debug("Sunday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Sunday
                });
            }
            return myDaysOfWeekMask;
        }
    }
}