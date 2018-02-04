VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmwallpaper 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WELCOME TO SIS"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   Icon            =   "frmwallpaper.frx":0000
   MaxButton       =   0   'False
   MouseIcon       =   "frmwallpaper.frx":08CA
   MousePointer    =   99  'Custom
   Picture         =   "frmwallpaper.frx":1194
   ScaleHeight     =   10080
   ScaleWidth      =   19080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   4440
      Top             =   6360
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   7080
      TabIndex        =   1
      Top             =   6600
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   99
      MouseIcon       =   "frmwallpaper.frx":37A96
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   13680
      TabIndex        =   0
      Top             =   6600
      Width           =   2535
   End
End
Attribute VB_Name = "frmwallpaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim a As Integer
Private Sub Form_Load()
a = 1
End Sub

Private Sub Timer1_Timer()
a = a + 1

Label1.Caption = CStr(a) & "% " & "System Shutdown"
ProgressBar1.Value = a
If a = 100 Then End
End Sub
