VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5280
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   8745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSplash.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   5280
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   150
      Left            =   1680
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   600
      Top             =   3240
   End
   Begin VB.Label Label3 
      Caption         =   "Portfolio Management"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "© 2015 by BPC. All rights reserved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5280
      TabIndex        =   0
      Top             =   4560
      Width           =   975
   End
End
Attribute VB_Name = "frmSplash"
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
If Appear_Counter = 160 Then
    Appear_Counter = 160
    Timer1.Enabled = False
On Error GoTo err
 
 frmlogin.Show
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
