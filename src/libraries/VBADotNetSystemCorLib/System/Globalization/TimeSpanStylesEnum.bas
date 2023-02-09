Attribute VB_Name = "TimeSpanStylesEnum"
Attribute VB_Description = "Defines the formatting options that customize string parsing for the ParseExact and TryParseExact methods. This enumeration supports a bitwise combination of its member values."
'Rubberduck annotations
'@Folder "VBADotNetCorLib.System.Globalization"
'@ModuleDescription "Defines the formatting options that customize string parsing for the ParseExact and TryParseExact methods. This enumeration supports a bitwise combination of its member values."

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 February 10, 2023
'@LastModified February 10, 2023

'@DotNetReferences
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.timespanstyles?view=net-7.0
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/Globalization/TimeSpanStyles.cs

Option Explicit

Public Enum TimeSpanStyles
    None = "&H000000000"            'Indicates that input is interpreted as a negative time interval only if a negative sign is present.
    AssumeNegative = "&H000000001"  'Indicates that input is always interpreted as a negative time interval.
End Enum
