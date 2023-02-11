Attribute VB_Name = "ErrorNumbers"
Attribute VB_Description = "Temporary error numbers  and messages untill exception handling implemented."
'Rubberduck annotations
'@Folder "VBADotNetCorLib.ErrorHandling"
'@ModuleDescription "Temporary error numbers  and messages untill exception handling implemented."

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 February 11, 2023
'@LastModified February 11, 2023

Option Explicit

Public Const ArgumentOutOfRangeException                As Long = 1000
Public Const ArgumentException                          As Long = 1001
Public Const ThrowArgumentNullException                 As Long = 1002
Public Const ThrowArgumentOutOfRange_TimeSpanTooLong    As Long = 1101
Public Const ThrowArgumentOutOfRange_BadHourMinuteSecond As Long = 1102
Public Const ArgumentOutOfRange_Year                    As Long = 1103
Public Const ThrowArgumentOutOfRange_BadYearMonthDay    As Long = 1104
Public Const ThrowArgumentOutOfRange_Month              As Long = 1105

Public Const OverflowException                          As Long = 2000


'DateTime Errors
'ArgumentOutOfRange_DateTimeBadTicks
'Argument_InvalidDateTimeKind
'Arg_DateTimeRange


'Not Implemented
Public Const NotImplementedException        As Long = 9999
Public Const NotImplementedError            As String = "Not implemented"
