Attribute VB_Name = "SqrMethod"
Option Explicit

'_declspec(dllexport) float _stdcall SqrDistance2DC(float X1, float Y1, float X2, float Y2)
Public Declare Function SqrDistance2DC Lib "Taylor.dll" (ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single

'_declspec(dllexport) float _stdcall SqrDistance3DC(float X1, float Y1, float Z1, float X2, float Y2, float Z2)
Public Declare Function SqrDistance3DC Lib "Taylor.dll" (ByVal X1 As Single, ByVal Y1 As Single, ByVal Z1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Z2 As Single) As Single

Public Function SqrDistance2D(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
 SqrDistance2D = Sqr((X2 - X1) * (X2 - X1) + (Y2 - Y1) * (Y2 - Y1))
End Function

Public Function SqrDistance3D(X1 As Single, Y1 As Single, Z1 As Single, X2 As Single, Y2 As Single, Z2 As Single) As Single
 SqrDistance3D = Sqr((X2 - X1) * (X2 - X1) + (Y2 - Y1) * (Y2 - Y1) + (Z2 - Z1) * (Z2 - Z1))
End Function
