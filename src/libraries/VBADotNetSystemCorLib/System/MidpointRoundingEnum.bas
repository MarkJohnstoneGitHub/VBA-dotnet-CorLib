Attribute VB_Name = "MidpointRoundingEnum"
Attribute VB_Description = "Specifies the strategy that mathematical rounding methods should use to round a number."
'Rubberduck annotations
'@ModuleDescription "Specifies the strategy that mathematical rounding methods should use to round a number."
'@Folder "VBADotNetCorLib.System"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 February 10, 2023
'@LastModified February 10, 2023

'@DotNetReferences
' https://learn.microsoft.com/en-us/dotnet/api/system.midpointrounding?view=net-7.0
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/MidpointRounding.cs

''
'@Remarks
'   Use MidpointRounding with appropriate overloads of Math.Round, MathF.Round, and Decimal.Round
'   to provide more control of the rounding process.
'
'   There are two overall rounding strategies, round to nearest and directing rounding, and each
'   enumeration field participates in exactly one of these strategies.
'
'   Round to nearest
'       Fields:
'           AwayFromZero
'           ToEven
'   Directed rounding
'       Fields:
'           ToNegativeInfinity
'           ToPositiveInfinity
'           ToZero
'
'Fields
'   AwayFromZero 1
'       The strategy of rounding to the nearest number, and when a number is halfway between two others,
'       it's rounded toward the nearest number that's away from zero.
'
'   ToEven 0
'       The strategy of rounding to the nearest number, and when a number is halfway between two others,
'       it's rounded toward the nearest even number.
'
'   ToNegativeInfinity 3
'       The strategy of downwards-directed rounding, with the result closest to and no greater than the
'       infinitely precise result.
'
'   ToPositiveInfinity 4
'       The strategy of upwards-directed rounding, with the result closest to and no less than the
'       infinitely precise result.
'
'   ToZero 2
'       The strategy of directed rounding toward zero, with the result closest to and no greater in
'       magnitude than the infinitely precise result.
''
Option Explicit

Public Enum MidpointRounding
    ToEven = 0
    AwayFromZero = 1
    ToZero = 2
    ToNegativeInfinity = 3
    ToPositiveInfinity = 4
End Enum
