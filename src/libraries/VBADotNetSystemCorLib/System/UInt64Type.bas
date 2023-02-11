Attribute VB_Name = "UInt64Type"
Attribute VB_Description = "Represents a 64-bit unsigned integer."
'Rubberduck annotations
'@Folder("VBADotNetCorLib.System")
'@ModuleDescription "Represents a 64-bit unsigned integer."

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 February 11, 2023
'@LastModified February 11, 2023

'@DotNetReferences
' https://learn.microsoft.com/en-us/dotnet/api/system.uint64?view=net-7.0
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/UInt64.cs
' https://github.com/dotnet/runtime/blob/ba1d6253f058c0ddc3399119dd1ec52017b704d7/src/libraries/System.Private.CoreLib/src/System/UInt64.cs

'@Remarks
'   The UInt64 value type represents unsigned integers with values ranging from 0 to 18,446,744,073,709,551,615.

Option Explicit

#If VBA7 Then
    Public Type UInt64
        value As LongLong
    End Type
#Else
    Public Type UInt64
        LowPart As Long
        HighPart As Long
    End Type
#End If
