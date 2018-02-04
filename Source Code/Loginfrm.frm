VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1845
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4590
   Icon            =   "Loginfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Loginfrm.frx":000C
   ScaleHeight     =   1090.087
   ScaleMode       =   0  'User
   ScaleWidth      =   4309.762
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox UserNameCmb 
      BackColor       =   &H00FF8080&
      Height          =   315
      ItemData        =   "Loginfrm.frx":BFAA
      Left            =   1290
      List            =   "Loginfrm.frx":BFB7
      TabIndex        =   6
      Text            =   "ADMIN"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1575
      Picture         =   "Loginfrm.frx":BFD1
      TabIndex        =   3
      Top             =   1140
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3180
      Picture         =   "Loginfrm.frx":17F6F
      TabIndex        =   4
      Top             =   1140
      Width           =   1140
   End
   Begin VB.TextBox Passwordtxt 
      BackColor       =   &H00FF8080&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblforgot 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
reccheck
rec.Open ("select * from LOGINTABLE where username = '" & Trim(UserNameCmb.Text) & "' and pwd ='" & Trim(Passwordtxt.Text) & "'"), con, adOpenDynamic, adLockOptimistic
If ((UserNameCmb.Text = "") Or (Passwordtxt.Text = "")) Then
  If (Passwordtxt.Text = "") Then
     MsgBox "Please enter the password", , "Login"
  End If
  If (UserNameCmb.Text = "") Then
    MsgBox "Username not entered", , "Login"
  End If
Else
 If rec.EOF = False Then
  If (UserNameCmb.Text = rec.Fields(0)) And (Passwordtxt.Text = rec.Fields(1)) Then
     'Unload Me
     frmMain.Show
     LoginSucceeded = True
     Me.Hide
  Else
     MsgBox "Invalid Password, try again!", , "Login"
     Passwordtxt.SetFocus
     SendKeys "{Home}+{End}"
  End If
 End If
End If
End Sub

Private Sub Form_Load()
  connection
End Sub
