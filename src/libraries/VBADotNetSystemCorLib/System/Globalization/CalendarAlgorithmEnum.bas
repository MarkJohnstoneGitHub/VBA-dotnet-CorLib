Attribute VB_Name = "CalendarAlgorithmEnum"
Attribute VB_Description = "Specifies whether a calendar is solar-based, lunar-based, or lunisolar-based."
'Rubberduck annotations
'@Folder("VBADotNetCorLib.System.Globalization")
'@ModuleDescription "Specifies whether a calendar is solar-based, lunar-based, or lunisolar-based."

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 February 10, 2023
'@LastModified February 10, 2023

'@DotNetReferences
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendaralgorithmtype?view=net-7.0
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/Globalization/CalendarAlgorithmType.cs

'@Remarks
'   A date calculation for a particular calendar depends on whether the calendar is solar-based,
'   lunar-based, or lunisolar-based. For example, the GregorianCalendar, JapaneseCalendar, and
'   JulianCalendar classes are solar-based, the HijriCalendar and UmAlQuraCalendar classes are
'   lunar-based,.and the HebrewCalendar and JapaneseLunisolarCalendar classes are lunisolar-based,
'   thus using solar calculations for the year and lunar calculations for the month and day.

'   A CalendarAlgorithmType value, which is returned by a calendar member such as the
'   Calendar.AlgorithmType property, specifies the foundation for a particular calendar.

Option Explicit

Public Enum CalendarAlgorithmType

    Unknown = 0             ' This is the default value to return in the Calendar base class.
    SolarCalendar = 1       ' Solar-base calendar, such as GregorianCalendar, jaoaneseCalendar, JulianCalendar, etc.
                            ' Solar calendars are based on the solar year and seasons.
    LunarCalendar = 2       ' Lunar-based calendar, such as Hijri and UmAlQuraCalendar.
                            ' Lunar calendars are based on the path of the moon.  The seasons are not accurately represented.
    LunisolarCalendar = 3   ' Lunisolar-based calendar which use leap month rule, such as HebrewCalendar and Asian Lunisolar calendars.
                            ' Lunisolar calendars are based on the cycle of the moon, but consider the seasons as a secondary consideration,
                            ' so they align with the seasons as well as lunar events.
End Enum
