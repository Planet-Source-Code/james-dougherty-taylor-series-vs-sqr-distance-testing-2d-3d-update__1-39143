Attribute VB_Name = "Taylor"
Option Explicit

'C++
'_declspec(dllexport) int _stdcall TaylorDistance2DC(int X, int Y)
Public Declare Function TaylorDistance2DC Lib "Taylor.dll" (ByVal X As Integer, ByVal Y As Integer) As Integer

'_declspec(dllexport) float _stdcall TaylorDistance3DC(float X, float Y, float Z)
Public Declare Function TaylorDistance3DC Lib "Taylor.dll" (ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Single

Private Function ShiftLeft(ByVal iNumber As Long, ByVal iBitsLeft As Long) As Long
 Dim iPower As Long
 iPower = RaisePower(iBitsLeft)
 ShiftLeft = CLng(iPower * iNumber)
End Function

Private Function ShiftRight(ByVal iNumber As Long, ByVal iBitsRight As Long) As Long
 Dim iPower As Long
 iPower = RaisePower(iBitsRight)
 ShiftRight = CLng((1 / iPower) * iNumber)
End Function

Private Function RaisePower(ByVal iPower As Long) As Long
 Dim iCount As Long
 Dim iRaised As Long
 iRaised = 1
 For iCount = 1 To iPower: iRaised = iRaised * 2: Next
 RaisePower = iRaised
End Function

Public Function TaylorDistance2D(X As Long, Y As Long) As Long
 Dim fMin As Single
 X = Abs(X)
 Y = Abs(Y)
 If X < Y Then fMin = X
 If X > Y Then fMin = Y
 TaylorDistance2D = (X + Y - (ShiftRight(fMin, 1)) - ShiftRight(fMin, 2) + ShiftRight(fMin, 4))
End Function

Public Function TaylorDistance3D(X As Single, Y As Single, Z As Single) As Single
 Dim iDistance As Long
 Dim iTmp As Long
 Dim iX As Long
 Dim iY As Long
 Dim iZ As Long
 
 iX = CLng(Abs(X) * 1024)
 iY = CLng(Abs(Y) * 1024)
 iZ = CLng(Abs(Z) * 1024)
  
 If iY < iX Then iTmp = iX: iX = iY: iY = iTmp: iTmp = 0
 If iZ < iY Then iTmp = iY: iY = iZ: iZ = iTmp: iTmp = 0
 If iY < iX Then iTmp = iX: iX = iY: iY = iTmp: iTmp = 0
 
 iDistance = (iZ + 11 * ShiftRight(iY, 5) + ShiftRight(iX, 2))
 TaylorDistance3D = (CSng(ShiftRight(iDistance, 10)))
 
End Function
