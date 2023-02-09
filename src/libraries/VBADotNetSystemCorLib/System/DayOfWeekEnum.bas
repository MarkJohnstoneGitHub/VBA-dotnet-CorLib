Attribute VB_Name = "DayOfWeekEnum"
Attribute VB_Description = "Specifies the day of the week."
'Rubberduck annotations
'@Folder "VBADotNetCorLib.System"
'@ModuleDescription "Specifies the day of the week."

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 February 10, 2023
'@LastModified February 10, 2023

'@DotNetReferences
' https://learn.microsoft.com/en-us/dotnet/api/system.dayofweek?view=net-7.0
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/DayOfWeek.cs

Option Explicit

''
'@Remarks
'   The DayOfWeek enumeration represents the day of the week in calendars that have seven days per week.
'   The value of the constants in this enumeration ranges from Sunday to Saturday. If cast to an integer,
'   its value ranges from zero (which indicates Sunday) to six (which indicates Saturday).
'
'   This enumeration is useful when it is desirable to have a strongly typed specification of the day of
'   the week. For example, this enumeration is the type of the property value for the DateTime.DayOfWeek
'   and DateTimeOffset.DayOfWeek properties.
'
'   The members of the DayOfWeek enumeration are not localized. To return the localized name of the
'   day of the week, call the DateTime.ToString(String) or the DateTime.ToString(String, IFormatProvider)
'   method with either the "ddd" or "dddd" format strings. The former format string produces the
'   abbreviated weekday name; the latter produces the full weekday name.
'
'@Note The DayOfWeek is 0 based, where as VB's vbDayOfWeek is 1 based.

Public Enum DayOfWeek
    Sunday              ' 0 Indicates Sunday.
    Monday              ' 1 Indicates Monday.
    Tuesday             ' 2 Indicates Tuesday.
    Wednesday           ' 3 Indicates Wednesday.
    Thursday            ' 4 Indicates Thursday.
    Friday              ' 5 Indicates Friday.
    Saturday            ' 6 Indicates Saturday.
End Enum
