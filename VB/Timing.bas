Attribute VB_Name = "Timing"
Option Explicit

Public Type LARGE_INTEGER
 LowPart As Long
 HighPart As Long
End Type

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function IntToCurrency(iInput As LARGE_INTEGER) As Currency
 CopyMemory IntToCurrency, iInput, LenB(iInput)
 IntToCurrency = IntToCurrency * 10000
End Function
