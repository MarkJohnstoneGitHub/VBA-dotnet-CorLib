Attribute VB_Name = "CalendarWeekRuleEnum"
Attribute VB_Description = "Defines different rules for determining the first week of the year."
'Rubberduck annotations
'@Folder("VBADotNetCorLib.System.Globalization")
'@ModuleDescription "Defines different rules for determining the first week of the year."

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 February 10, 2023
'@LastModified February 10, 2023

'@DotNetReferences
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendarweekrule?view=net-7.0
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/Globalization/CalendarWeekRule.cs

'@Remarks
'   A member of the CalendarWeekRule enumeration is returned by the DateTimeFormatInfo.CalendarWeekRule
'   property and is used by the culture's current calendar to determine the calendar week rule.
'   The enumeration value is also used as a parameter to the Calendar.GetWeekOfYear method.
'
'   Calendar week rules depend on the System.DayOfWeek value that indicates the first day of the
'   week in addition to depending on a CalendarWeekRule value. The DateTimeFormatInfo.FirstDayOfWeek
'   property provides the default value for a culture, but any DayOfWeek value can be specified as
'   the first day of the week in the Calendar.GetWeekOfYear method.
'
'   The first week based on the FirstDay value can have one to seven days. The first week based on
'   the FirstFullWeek value always has seven days. The first week based on the FirstFourDayWeek
'   value can have four to seven days.

Option Explicit

Public Enum CalendarWeekRule
    FirstDay = 0            ' Week 1 begins on the first day of the year

    FirstFullWeek = 1       ' Week 1 begins on first FirstDayOfWeek not before the first day of the year

    FirstFourDayWeek = 2    ' Week 1 begins on first FirstDayOfWeek such that FirstDayOfWeek+3 is not before the first day of the year
End Enum
