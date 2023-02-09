Attribute VB_Name = "DateTimeKindEnum"
Attribute VB_Description = "Specifies whether a DateTime object represents a local time, a Coordinated Universal Time (UTC), or is not specified as either local time or UTC."
'Rubberduck annotations
'@Folder "VBADotNetCorLib.System"
'@ModuleDescription "Specifies whether a DateTime object represents a local time, a Coordinated Universal Time (UTC), or is not specified as either local time or UTC."

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 February 10, 2023
'@LastModified February 10, 2023

'@DotNetReferences
' https://learn.microsoft.com/en-us/dotnet/api/system.datetimekind?view=net-7.0
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/DateTimeKind.cs

Option Explicit

''
'@Remarks
'   A member of the DateTimeKind enumeration is returned by the DateTime.Kind property.
'
'   The members of the DateTimeKind enumeration are used in conversion operations between local time
'   and Coordinated Universal Time (UTC), but not in comparison or arithmetic operations.
'   For more information about time conversions, see Converting Times Between Time Zones.
''
Public Enum DateTimeKind
    Unspecified = 0         'The time represented is not specified as either local time or Coordinated Universal Time (UTC).
    Utc = 1                 'The time represented is UTC.
    Locale = 2              'The time represented is local time. Note: Renamed from local due to VBA reserved word
    LocalAmbiguousDst = 3   '@TODO Added to handle AmbiguousDst in DateTime ??
End Enum
