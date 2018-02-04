VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   0  'None
   ClientHeight    =   1425
   ClientLeft      =   7110
   ClientTop       =   6000
   ClientWidth     =   4665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Welcomefrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   150
      Left            =   4680
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   4680
      Top             =   600
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   495
      Left            =   960
      Top             =   360
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   150
      Picture         =   "Welcomefrm.frx":000C
      Top             =   150
      Width           =   720
   End
   Begin VB.Label lblwelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'global variable declarations
Dim Appear_Counter As Integer
'global constants declarations
Const LWA_COLORKEY = &H3
Const LWA_ALPHA = &H3
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
'API functions declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long



Private Sub Form_Load()
Dim Ret As Long

lblwelcome.Caption = frmlogin.UserNameCmb
Appear_Counter = 0
Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
Ret = Ret Or WS_EX_LAYERED
SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
Timer1.Enabled = True
Timer2.Enabled = False
End Sub
Private Sub Timer1_Timer()

SetLayeredWindowAttributes Me.hwnd, 0, Appear_Counter, LWA_ALPHA
Appear_Counter = Appear_Counter + 10
If Appear_Counter = 260 Then
    Appear_Counter = 260
    Timer1.Enabled = False
On Error GoTo err
Unload Me
err:

End If
End Sub

Private Sub Timer2_Timer()
On Error GoTo err
SetLayeredWindowAttributes Me.hwnd, 0, Appear_Counter, LWA_ALPHA
Appear_Counter = Appear_Counter - 5
Label2.Caption = Appear_Counter
If Appear_Counter = 0 Then
 
    End
End If
err:
    err.Clear
    Exit Sub
    Unload Me
    End
End Sub


