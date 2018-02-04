VERSION 5.00
Begin VB.Form frmPhysicalInfo 
   Caption         =   "Extra Curricular Activities"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16410
   Icon            =   "frmPhysical.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmPhysical.frx":000C
   ScaleHeight     =   11010
   ScaleWidth      =   16410
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRegNo 
      Height          =   375
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   3480
      MaxLength       =   35
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ComboBox cmbCourse 
      Height          =   315
      ItemData        =   "frmPhysical.frx":1C0C3
      Left            =   6360
      List            =   "frmPhysical.frx":1C0DF
      TabIndex        =   1
      Text            =   "--Select--"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      ItemData        =   "frmPhysical.frx":1C11B
      Left            =   9360
      List            =   "frmPhysical.frx":1C1A9
      TabIndex        =   0
      Text            =   "Year"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image imgAttach 
      Height          =   375
      Left            =   11760
      Picture         =   "frmPhysical.frx":1C2C1
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblPosition 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label lblSubCategory 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   3960
      Width           =   5175
   End
   Begin VB.Label lblCategory 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3360
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Top             =   4605
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   3495
      Left            =   1920
      Top             =   2880
      Width           =   10695
   End
   Begin VB.Image imgCancel 
      Height          =   375
      Left            =   2280
      Picture         =   "frmPhysical.frx":1E23A
      Top             =   7140
      Width           =   375
   End
   Begin VB.Image imgPrint 
      Height          =   450
      Left            =   11640
      Picture         =   "frmPhysical.frx":1E71A
      Top             =   7140
      Width           =   375
   End
   Begin VB.Image imgReport 
      Height          =   450
      Left            =   10200
      Picture         =   "frmPhysical.frx":1EB28
      Top             =   7140
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   855
      Left            =   1920
      Top             =   6960
      Width           =   10695
   End
   Begin VB.Shape Shape4 
      Height          =   405
      Index           =   0
      Left            =   2280
      Top             =   7140
      Width           =   375
   End
   Begin VB.Shape Shape4 
      Height          =   465
      Index           =   1
      Left            =   10200
      Top             =   7140
      Width           =   375
   End
   Begin VB.Shape Shape4 
      Height          =   465
      Index           =   2
      Left            =   11640
      Top             =   7140
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000D&
      Height          =   1455
      Left            =   1920
      Top             =   1080
      Width           =   10695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reg No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   12
      Left            =   2520
      TabIndex        =   7
      Top             =   1395
      Width           =   855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   6
      Top             =   1845
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   8520
      TabIndex        =   5
      Top             =   1845
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   1845
      Width           =   855
   End
End
Attribute VB_Name = "frmPhysicalInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Theme frmPhysicalInfo
  connection
End Sub

Private Sub imgAttach_Click()
  frmAttachment.imgOpen.Enabled = False
  frmAttachment.imgSave.Enabled = False
  frmAttachment.txtRegNo.Text = frmPhysicalInfo.txtRegNo.Text
  frmAttachment.Show
End Sub

Private Sub imgReport_Click()
   reccheck
   rec.Open "select * from PHYSICALTABLE where REGNO = " & Trim(txtRegNo.Text), con, adOpenDynamic, adLockPessimistic
   If rec.EOF = False Then
      lblCategory.Caption = rec.Fields(0)
      lblSubCategory.Caption = rec.Fields(1)
      lblPosition.Caption = rec.Fields(2)
   Else
      MsgBox "No such record found"
   End If
End Sub

Private Sub txtRegNo_Change()
   lblCategory.Caption = ""
   lblSubCategory.Caption = ""
   lblPosition.Caption = ""
End Sub

Private Sub txtRegNo_KeyPress(KeyAscii As Integer)
  ValRegNo KeyAscii
End Sub
