VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Welcome to SIS"
   ClientHeight    =   11520
   ClientLeft      =   6915
   ClientTop       =   7155
   ClientWidth     =   18270
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmLogin.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   11520
   ScaleWidth      =   18270
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbAcType 
      Height          =   315
      ItemData        =   "frmLogin.frx":1194
      Left            =   9480
      List            =   "frmLogin.frx":119E
      TabIndex        =   9
      Text            =   "ADMIN"
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "E:\Portfolio Management\PotfolioDb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Admin"
      Top             =   7320
      Width           =   1425
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Portfolio Management\PotfolioDb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Client"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox UserNameCmb 
      DataField       =   "Username"
      DataSource      =   "Data1"
      Height          =   315
      ItemData        =   "frmLogin.frx":11B1
      Left            =   9480
      List            =   "frmLogin.frx":11BB
      TabIndex        =   6
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox passwordtxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   9480
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User Type :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Index           =   2
      Left            =   12840
      Top             =   7680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Index           =   1
      Left            =   11040
      Top             =   8400
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Index           =   0
      Left            =   9960
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label lblHint 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   13080
      TabIndex        =   7
      Top             =   8160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image imgHint 
      Height          =   375
      Left            =   12840
      Picture         =   "frmLogin.frx":11CE
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image errorimg 
      Height          =   480
      Left            =   8040
      Picture         =   "frmLogin.frx":1751
      Top             =   9000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image fieldblank 
      Height          =   480
      Left            =   8040
      Picture         =   "frmLogin.frx":1B93
      Top             =   9000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label fieldlbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Fields can not be blank."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   9240
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label errorlbl 
      BackStyle       =   0  'Transparent
      Caption         =   "The username or password is incorrect. Please try again."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   9240
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image imgCancel 
      Height          =   375
      Left            =   11040
      MouseIcon       =   "frmLogin.frx":1FD5
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":289F
      ToolTipText     =   "Close"
      Top             =   8400
      Width           =   375
   End
   Begin VB.Image imgOK 
      Height          =   375
      Left            =   9960
      MouseIcon       =   "frmLogin.frx":2D7F
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":3649
      ToolTipText     =   "Login"
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "For use  by Authorized Personnel Only....                  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   1
      Left            =   8160
      TabIndex        =   3
      Top             =   5520
      Width           =   3675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   7700
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   0
      Left            =   8040
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean



Private Sub cmbAcType_Change()
   If cmbAcType.Text = "CLIENT" Then
      UserNameCmb.Visible = True
      Label1(0).Visible = True
   End If
End Sub

Private Sub imgHint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   reccheck
   rec.Open ("select HINT from LOGINTABLE where username = '" & Trim(UserNameCmb.Text) & "'"), con, adOpenDynamic, adLockOptimistic
   If rec.EOF = False Then
      lblHint.Caption = rec.Fields(0)
      lblHint.Visible = True
   End If
End Sub

Private Sub imgHint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblHint.Visible = False
End Sub

Private Sub passwordtxt_Change()
  If Not passwordtxt.Text = "" Then
     fieldlbl.Visible = False
     fieldblank.Visible = False
  End If
  If passwordtxt.Text = "" Then
     fieldlbl.Visible = False
     fieldblank.Visible = False
     errorlbl.Visible = False
     errorimg.Visible = False
  End If
End Sub

Private Sub usernamecmb_Change()
  If UserNameCmb.Text = "" Then
     passwordtxt.Text = ""
     errorlbl.Visible = False
     errorimg.Visible = False
     fieldlbl.Visible = False
     fieldblank.Visible = False
  End If
  If Not UserNameCmb.Text = "" And passwordtxt.Text = "" Then
     fieldlbl.Visible = False
     fieldblank.Visible = False
  End If
End Sub

Private Sub imgCancel_Click()
  LoginSucceeded = False
  Unload Me
End Sub

Private Sub imgOK_Click()
  If UserNameCmb.Text = "" Then
     UserNameCmb.SetFocus
     fieldlbl.Visible = True
     fieldblank.Visible = True
     Exit Sub
  End If

  If passwordtxt.Text = "" Then
     UserNameCmb.SetFocus
     fieldlbl.Visible = True
     fieldblank.Visible = True
     Exit Sub
  End If
  
  
  If (UserNameCmb.Text = Data1.Recordset.Fields(4)) And (passwordtxt.Text = Data1.Recordset.Fields(0)) Then
      frmAdmin.Show
      LoginSucceeded = True
      Me.Hide
  Else
      errorlbl.Visible = True
      errorimg.Visible = True
      fieldlbl.Visible = False
      fieldblank.Visible = False
      errorlbl.Refresh
      errorimg.Refresh
      passwordtxt.SetFocus
  End If
End Sub
