VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAdminOptions 
   BackColor       =   &H80000005&
   Caption         =   "Administrator Options"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16410
   Icon            =   "AdminOptionsfrm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "AdminOptionsfrm.frx":000C
   ScaleHeight     =   9750
   ScaleWidth      =   16410
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   2280
      TabIndex        =   0
      Top             =   1200
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   8421504
      TabCaption(0)   =   "Change Password"
      TabPicture(0)   =   "AdminOptionsfrm.frx":1C0C3
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmbUserName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtOPassword"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtHPassword"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtCPassword"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtNPassword"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Shape1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgCancel(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label9"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label7"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label5"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Shape2(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Shape3(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Shape4(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Input Configuration"
      TabPicture(1)   =   "AdminOptionsfrm.frx":1C0DF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "imgCancel(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label14"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Shape2(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Shape3(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Shape4(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Theme"
      TabPicture(2)   =   "AdminOptionsfrm.frx":1C0FB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmbTheme"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "imgCancel(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Shape2(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Shape3(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Shape4(0)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lblApply"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label4"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label3"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Search Configuration"
      TabPicture(3)   =   "AdminOptionsfrm.frx":1C117
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "imgCancel(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Shape2(3)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Shape3(3)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Shape4(3)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Shape5"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label15"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label16"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "lblOK"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtQuestion"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txtQuery"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).ControlCount=   10
      Begin VB.TextBox txtQuery 
         Height          =   375
         Left            =   2520
         TabIndex        =   47
         Top             =   1560
         Width           =   6135
      End
      Begin VB.TextBox txtQuestion 
         Height          =   375
         Left            =   2520
         TabIndex        =   46
         Top             =   840
         Width           =   6135
      End
      Begin VB.Frame Frame2 
         Caption         =   " Remove Anything"
         Height          =   2535
         Index           =   1
         Left            =   -70200
         TabIndex        =   35
         Top             =   2280
         Width           =   4335
         Begin VB.TextBox txtSchool 
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   39
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox txtCast 
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   38
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox txtReligion 
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   37
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox txtProgramme 
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   36
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "School"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   43
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cast"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   42
            Top             =   1440
            Width           =   315
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Religion"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   41
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Programme"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   40
            Top             =   480
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Add More"
         Height          =   2535
         Index           =   0
         Left            =   -74640
         TabIndex        =   21
         Top             =   2280
         Width           =   4335
         Begin VB.TextBox txtSchool 
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   34
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox txtCast 
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   33
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox txtReligion 
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   32
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox txtProgramme 
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   31
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "School"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   27
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cast"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Width           =   315
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Religion"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   25
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Programme"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   480
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Settings"
         Height          =   1335
         Left            =   -74640
         TabIndex        =   20
         Top             =   840
         Width           =   8775
         Begin VB.TextBox txtAdditMark 
            Height          =   375
            Left            =   2160
            TabIndex        =   30
            Top             =   800
            Width           =   615
         End
         Begin VB.TextBox txtMaxMark 
            Height          =   375
            Left            =   2160
            TabIndex        =   29
            Top             =   320
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Additional Mark"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "Maximum Internal Mark"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.ComboBox cmbUserName 
         Height          =   315
         ItemData        =   "AdminOptionsfrm.frx":1C133
         Left            =   -72240
         List            =   "AdminOptionsfrm.frx":1C140
         TabIndex        =   15
         Text            =   "ADMIN"
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtOPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   -72240
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   2040
         Width           =   3375
      End
      Begin VB.PictureBox Picture1 
         Height          =   3800
         Left            =   -74280
         Picture         =   "AdminOptionsfrm.frx":1C15A
         ScaleHeight     =   3735
         ScaleWidth      =   3735
         TabIndex        =   7
         Top             =   1680
         Width           =   3800
      End
      Begin VB.ComboBox cmbTheme 
         Height          =   315
         ItemData        =   "AdminOptionsfrm.frx":1E066
         Left            =   -74280
         List            =   "AdminOptionsfrm.frx":1E088
         TabIndex        =   4
         Text            =   "Default"
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtHPassword 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72240
         TabIndex        =   3
         Top             =   3840
         Width           =   3375
      End
      Begin VB.TextBox txtCPassword 
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   -72240
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox txtNPassword 
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   -72240
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label lblOK 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         Height          =   495
         Left            =   7560
         TabIndex        =   48
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   45
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Question"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H8000000D&
         Height          =   4695
         Left            =   360
         Top             =   600
         Width           =   8895
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         Height          =   4095
         Left            =   -74640
         Top             =   960
         Width           =   8895
      End
      Begin VB.Shape Shape4 
         Height          =   495
         Index           =   3
         Left            =   7560
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         Height          =   405
         Index           =   3
         Left            =   600
         Top             =   6000
         Width           =   405
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000D&
         Height          =   975
         Index           =   3
         Left            =   360
         Top             =   5760
         Width           =   8895
      End
      Begin VB.Image imgCancel 
         Height          =   375
         Index           =   3
         Left            =   615
         Picture         =   "AdminOptionsfrm.frx":1E0FA
         Top             =   6015
         Width           =   375
      End
      Begin VB.Image imgCancel 
         Height          =   375
         Index           =   0
         Left            =   -74390
         Picture         =   "AdminOptionsfrm.frx":1E5DA
         Top             =   6010
         Width           =   375
      End
      Begin VB.Image imgCancel 
         Height          =   375
         Index           =   2
         Left            =   -74390
         Picture         =   "AdminOptionsfrm.frx":1EABA
         Top             =   6010
         Width           =   375
      End
      Begin VB.Image imgCancel 
         Height          =   375
         Index           =   1
         Left            =   -74390
         Picture         =   "AdminOptionsfrm.frx":1EF9A
         Top             =   6010
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         Height          =   495
         Left            =   -67440
         TabIndex        =   28
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Password lefts Null"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   -68520
         TabIndex        =   19
         Top             =   3240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Password lefts Null"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   -68520
         TabIndex        =   18
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
         Height          =   375
         Left            =   -67440
         TabIndex        =   17
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Wrong Password entered"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -68520
         TabIndex        =   16
         Top             =   2055
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Password Hint"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74160
         TabIndex        =   13
         Top             =   3960
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Conform Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74160
         TabIndex        =   12
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74160
         TabIndex        =   11
         Top             =   2760
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74160
         TabIndex        =   10
         Top             =   2160
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74160
         TabIndex        =   9
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000D&
         Height          =   975
         Index           =   2
         Left            =   -74640
         Top             =   5760
         Width           =   8895
      End
      Begin VB.Shape Shape3 
         Height          =   400
         Index           =   2
         Left            =   -74400
         Top             =   6000
         Width           =   400
      End
      Begin VB.Shape Shape4 
         Height          =   495
         Index           =   2
         Left            =   -67440
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000D&
         Height          =   975
         Index           =   1
         Left            =   -74640
         Top             =   5760
         Width           =   8895
      End
      Begin VB.Shape Shape3 
         Height          =   400
         Index           =   1
         Left            =   -74400
         Top             =   6000
         Width           =   400
      End
      Begin VB.Shape Shape4 
         Height          =   495
         Index           =   1
         Left            =   -67440
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000D&
         Height          =   975
         Index           =   0
         Left            =   -74640
         Top             =   5760
         Width           =   8895
      End
      Begin VB.Shape Shape3 
         Height          =   400
         Index           =   0
         Left            =   -74400
         Top             =   6000
         Width           =   400
      End
      Begin VB.Shape Shape4 
         Height          =   495
         Index           =   0
         Left            =   -67440
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Label lblApply 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Apply"
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   -67440
         TabIndex        =   8
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Sample:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Theme:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAdminOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fso As New FileSystemObject

Private Sub Form_Load()
   Theme frmAdminOptions
   connection
End Sub

Private Sub Image3_Click()

End Sub

Private Sub imgCancel_Click(Index As Integer)
  Unload Me
End Sub

Private Sub Label14_Click()
   On Error GoTo error_para
   
   Dim validate As Integer
   
   validate = 0
   If txtMaxMark.Text <> "" Then
      reccheck
      rec.Open "update table SISSETTINGS set MAXMARK = '" & Trim(txtMaxMark.Text) & "'", con, adOpenDynamic, adLockPessimistic
      reccheck
      rec.Open "update table SISSETTINGS set MAXMARK = '" & Trim(txtMaxMark.Text) & "'", con, adOpenDynamic, adLockPessimistic
   Else
      validate = validate + 1
   End If
   
   If txtMaxMark.Text <> "" Then
      reccheck
      rec.Open "update table SISSETTINGS set ADDITMARK = '" & Trim(txtAdditMark.Text) & "'", con, adOpenDynamic, adLockPessimistic
      reccheck
      rec.Open "update table SISSETTINGS set ADDITMARK = '" & Trim(txtAdditMark.Text) & "'", con, adOpenDynamic, adLockPessimistic
   Else
      validate = validate + 1
   End If
   
   If AddInfo(txtProgramme(0), "PROGRAMME") = False Then validate = validate + 1
   
   If AddInfo(txtReligion(0), "RELIGION") = False Then validate = validate + 1
    
   If AddInfo(txtCast(0), "CASTE") = False Then validate = validate + 1
   
   If AddInfo(txtSchool(0), "SCHOOL") = False Then validate = validate + 1
   
   If RemoveInfo(txtProgramme(1), "PROGRAMME") = False Then validate = validate + 1
   
   If RemoveInfo(txtReligion(1), "RELIGION") = False Then validate = validate + 1
   
   If RemoveInfo(txtCast(1), "CASTE") = False Then validate = validate + 1
   
   If RemoveInfo(txtSchool(1), "SCHOOL") = False Then validate = validate + 1
   
   If validate = 10 Then
      MsgBox "You must provide atleast one information for made changes", vbCritical, "Add More Info"
   Else
      MsgBox "Changes made Successfully", vbInformation, "Add More Info"
   End If
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub Label9_Click()
   On Error GoTo error_para
   
   Dim validate As Boolean
   validate = True
   If txtNPassword.Enabled = True Then
      If txtNPassword.Text <> "" Then
         Label10(0).Visible = False
      Else
         validate = False
         Label10(0).Visible = True
      End If
      If txtCPassword.Text <> "" Then
         Label10(1).Visible = False
      Else
         validate = False
         Label10(1).Visible = True
      End If
      If validate = True Then
         If Trim(txtNPassword.Text) = Trim(txtCPassword.Text) Then
            reccheck
            rec.Open "update LOGINTABLE set PWD = '" & Trim(txtNPassword.Text) & "' where USERNAME = '" & Trim(cmbUserName.Text) & "'", con, adOpenDynamic, adLockPessimistic
            rec.Open "update LOGINTABLE set HINT = '" & Trim(txtHPassword.Text) & "' where USERNAME = '" & Trim(cmbUserName.Text) & "'", con, adOpenDynamic, adLockPessimistic
            MsgBox "Password changed Successfully", vbInformation, "Change Password"
            txtNPassword.Text = ""
            txtCPassword.Text = ""
            txtHPassword.Text = ""
         Else
            MsgBox "New Password and Conform Password are not Match", vbCritical, "Change Password"
         End If
      End If
   Else
      MsgBox "You must conform the Old Password to Change Password", vbCritical, "Change Password"
   End If
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbTheme_Change()
  On Error GoTo error_para
  
  If fso.FolderExists(App.Path & "\Theme") = False Then fso.CreateFolder App.Path & "\Theme"
  Select Case cmbTheme.Text
     Case "Default":
                      If fso.FileExists(App.Path & "\Theme\Default.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\Default.jpg")
                      End If
     Case "Cloud Burst":
     
                      If fso.FileExists(App.Path & "\Theme\CloudBurst.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\CloudBurst.jpg")
                      End If
     Case "Game Fight":
                      If fso.FileExists(App.Path & "\Theme\GameFight.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\GameFight.jpg")
                      End If
     
     Case "Colors":
                      If fso.FileExists(App.Path & "\Theme\Colors.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\Colors.jpg")
                      End If
     
     Case "Green":
                      If fso.FileExists(App.Path & "\Theme\Green.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\Green.jpg")
                      End If
     
     Case "Blue":
                      If fso.FileExists(App.Path & "\Theme\Blue.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\Blue.jpg")
                      End If
     
     Case "Blue Shades":
                      If fso.FileExists(App.Path & "\Theme\BlueShades.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\BlueShades.jpg")
                      End If
     Case "Orange Shades":
                      If fso.FileExists(App.Path & "\Theme\OrangeShades.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\OrangeShades.jpg")
                      End If
     Case "Gray Shades":
                      If fso.FileExists(App.Path & "\Theme\GrayShades.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\GrayShades.jpg")
                      End If
     Case "Green Shades":
                      If fso.FileExists(App.Path & "\Theme\GreenShades.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\GreenShades.jpg")
                      End If
   
  End Select
  Picture1.Refresh
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbTheme_Click()
  On Error GoTo error_para
  
  If fso.FolderExists(App.Path & "\Theme") = False Then fso.CreateFolder App.Path & "\Theme"
  Select Case cmbTheme.Text
     Case "Default":
                      If fso.FileExists(App.Path & "\Theme\Default.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\Default.jpg")
                      End If
     Case "Cloud Burst":
     
                      If fso.FileExists(App.Path & "\Theme\CloudBurst.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\CloudBurst.jpg")
                      End If
     Case "Game Fight":
                      If fso.FileExists(App.Path & "\Theme\GameFight.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\GameFight.jpg")
                      End If
     
     Case "Colors":
                      If fso.FileExists(App.Path & "\Theme\Colors.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\Colors.jpg")
                      End If
     
     Case "Green":
                      If fso.FileExists(App.Path & "\Theme\Green.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\Green.jpg")
                      End If
     
     Case "Blue":
                      If fso.FileExists(App.Path & "\Theme\Blue.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\Blue.jpg")
                      End If
    
     Case "Blue Shades":
                      If fso.FileExists(App.Path & "\Theme\BlueShades.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\BlueShades.jpg")
                      End If
     Case "Orange Shades":
                      If fso.FileExists(App.Path & "\Theme\OrangeShades.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\OrangeShades.jpg")
                      End If
     Case "Gray Shades":
                      If fso.FileExists(App.Path & "\Theme\GrayShades.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\GrayShades.jpg")
                      End If
     Case "Green Shades":
                      If fso.FileExists(App.Path & "\Theme\GreenShades.jpg") = True Then
                         Picture1 = LoadPicture(App.Path & "\Theme\GreenShades.jpg")
                      End If
  End Select
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub lblApply_Click()
   Dim Theme As Integer
   Select Case cmbTheme.Text
      Case "Default": Theme = 0
      Case "Cloud Burst": Theme = 1
      Case "Game Fight": Theme = 2
      Case "Colors": Theme = 3
      Case "Green": Theme = 4
      Case "Blue": Theme = 5
      Case "Blue Shades": Theme = 6
      Case "Orange Shades": Theme = 7
      Case "Gray Shades": Theme = 8
      Case "Green Shades": Theme = 9
   End Select
   
   If fso.FolderExists(App.Path & "\Theme") = False Then fso.CreateFolder App.Path & "\Theme"
   Open App.Path & "\Theme\Theme.txt" For Output As #1
   Print #1, Theme
   Close #1
   MsgBox "Chages Apply only when restaring the System", vbInformation, ""
   Unload Me
End Sub

Private Sub lblOK_Click()
  On Error GoTo error_para

  Dim temp As Integer
  If txtQuestion.Text = "" Or txtQuery.Text = "" Then
     MsgBox "Question and Query left blank"
  Else
     reccheck
     rec.Open "select MAX(QNO) from SEARCHENGINE", con, adOpenDynamic, adLockPessimistic
     If rec.EOF = False Then
        temp = rec.Fields(0)
     Else
        temp = 1
     End If
     reccheck
     rec.Open "insert into SEARCHENGINE values('" & temp & "','" & txtQuery.Text & "','" & txtQuestion.Text & "')", con, adOpenDynamic, adLockPessimistic
     MsgBox "Question added sucessfully"
  End If
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub txtOPassword_Change()
  On Error GoTo error_para
  
  reccheck
  rec.Open "select * from LOGINTABLE where username = '" & Trim(cmbUserName.Text) & "'", con, adOpenDynamic, adLockOptimistic

  If rec.EOF = False Then
    If Not (txtOPassword.Text = rec.Fields(1)) Then
       Label8.Visible = True
       txtNPassword.Enabled = False
       txtCPassword.Enabled = False
       txtHPassword.Enabled = False
       txtOPassword.SetFocus
    Else
       Label8.Visible = False
       txtNPassword.Enabled = True
       txtCPassword.Enabled = True
       txtHPassword.Enabled = True
       txtNPassword.SetFocus
    End If
  End If
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub txtQuestion_KeyPress(KeyAscii As Integer)
   ValName KeyAscii
End Sub


