VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Taylor Series Distance"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk2D 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.OptionButton optType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C++"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   14
      Top             =   1320
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Visual Basic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtDistance 
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtDistance 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdSQR 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Test Sqr() Func."
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdTaylor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Test Taylor Series"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblCoord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z1 = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   2760
      TabIndex        =   16
      Top             =   120
      Width           =   405
   End
   Begin VB.Label lblCoord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z2 = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   2760
      TabIndex        =   15
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click 2x to get real time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   945
      TabIndex        =   12
      Top             =   3000
      Width           =   1995
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time - 0.000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2400
      TabIndex        =   11
      Top             =   2640
      Width           =   1035
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time - 0.000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   2640
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2040
      TabIndex        =   9
      Top             =   1800
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   750
   End
   Begin VB.Label lblCoord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y2 = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   405
   End
   Begin VB.Label lblCoord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X2 = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lblCoord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y1 = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   405
   End
   Begin VB.Label lblCoord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X1 = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  
'COMPILE THE TAYLOR DLL INTO THIS DIR OR IT WILL NOT WORK
  
  
Private fX1 As Single
Private fX2 As Single
Private fY1 As Single
Private fY2 As Single
Private fZ1 As Single
Private fZ2 As Single

Private Sub cmdTaylor_Click()
 Dim iStart As LARGE_INTEGER
 Dim iStop As LARGE_INTEGER
 Dim cStart As Currency
 Dim cStop As Currency
 
 'Timing
 QueryPerformanceCounter iStart
 
 If chk2D.Value Then '2D
  If optType(0).Value Then 'VB method
   txtDistance(0).Text = TaylorDistance2D(CLng(fX1) - CLng(fX2), CLng(fY1) - CLng(fY2))
  Else 'C++ method
   txtDistance(0).Text = TaylorDistance2DC(CLng(fX1) - CLng(fX2), CLng(fY1) - CLng(fY2))
  End If
 Else '3D
  If optType(0).Value Then 'VB method
   txtDistance(0).Text = TaylorDistance3D(fX1 - fX2, fY1 - fY1, fZ1 - fZ2)
  Else 'C++ method
   txtDistance(0).Text = TaylorDistance3DC(CLng(fX1) - CLng(fX2), CLng(fY1) - CLng(fY1), CLng(fZ1) - CLng(fZ2))
  End If
 End If
 
 'Timing
 QueryPerformanceCounter iStop
 cStart = IntToCurrency(iStart)
 cStop = IntToCurrency(iStop)
 
 'Dislay how long it took
 lblTime(0) = "Time - " & Format$((cStop - cStart), "000")
 
End Sub

Private Sub cmdSQR_Click()
 Dim iStart As LARGE_INTEGER
 Dim iStop As LARGE_INTEGER
 Dim cStart As Currency
 Dim cStop As Currency
 
 'Timing
 QueryPerformanceCounter iStart
 
 If chk2D.Value Then '2D
  If optType(0).Value Then 'VB method
   txtDistance(1).Text = SqrDistance2D(fX1, fY1, fX2, fY2)
  Else 'C++ method
   txtDistance(1).Text = SqrDistance2DC(fX1, fY1, fX2, fY2)
  End If
 Else
  If optType(0).Value Then 'VB method
   txtDistance(1).Text = SqrDistance3D(fX1, fY1, fZ1, fX2, fY2, fZ2)
  Else 'C++ method
   txtDistance(1).Text = SqrDistance3DC(fX1, fY1, fZ1, fX2, fY2, fZ2)
  End If
 End If
 
 'Timing
 QueryPerformanceCounter iStop
 cStart = IntToCurrency(iStart)
 cStop = IntToCurrency(iStop)
 
 'Dislay how long it took
 lblTime(1) = "Time - " & Format$((cStop - cStart), "000")
 
End Sub

Private Sub Form_Load()
 
 'Fill in the values to be checked
 fX1 = 55: fY1 = 550: fZ1 = 250
 fX2 = 10: fY2 = 500: fZ2 = -200
 
 'then show the values
 lblCoord(0) = lblCoord(0) & fX1
 lblCoord(1) = lblCoord(1) & fY1
 lblCoord(5) = lblCoord(5) & fZ1
 lblCoord(2) = lblCoord(2) & fX2
 lblCoord(3) = lblCoord(3) & fY2
 lblCoord(4) = lblCoord(4) & fZ2
 
End Sub
