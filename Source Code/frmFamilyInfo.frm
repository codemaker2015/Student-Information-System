VERSION 5.00
Begin VB.Form frmFamilyInfo 
   Caption         =   "Family Information"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14475
   Icon            =   "frmFamilyInfo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmFamilyInfo.frx":000C
   ScaleHeight     =   9675
   ScaleWidth      =   14475
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbYear 
      Height          =   315
      ItemData        =   "frmFamilyInfo.frx":1C0C3
      Left            =   8400
      List            =   "frmFamilyInfo.frx":1C151
      TabIndex        =   3
      Text            =   "Year"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cmbCourse 
      Height          =   315
      ItemData        =   "frmFamilyInfo.frx":1C269
      Left            =   5400
      List            =   "frmFamilyInfo.frx":1C285
      TabIndex        =   2
      Text            =   "--Select--"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2520
      MaxLength       =   35
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtRegNo 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      Height          =   470
      Index           =   2
      Left            =   10680
      Top             =   8100
      Width           =   375
   End
   Begin VB.Shape Shape4 
      Height          =   470
      Index           =   1
      Left            =   9240
      Top             =   8100
      Width           =   375
   End
   Begin VB.Shape Shape4 
      Height          =   400
      Index           =   0
      Left            =   1320
      Top             =   8100
      Width           =   375
   End
   Begin VB.Label Label1 
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
      Height          =   1575
      Index           =   9
      Left            =   6360
      TabIndex        =   27
      Top             =   6000
      Width           =   4815
   End
   Begin VB.Label lblPhone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      Height          =   240
      Index           =   1
      Left            =   1560
      TabIndex        =   26
      Top             =   6240
      Width           =   585
   End
   Begin VB.Label lblMOccupation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
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
      Height          =   240
      Left            =   1560
      TabIndex        =   25
      Top             =   5640
      Width           =   1020
   End
   Begin VB.Label lblMotherName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mother's Name"
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
      Height          =   240
      Left            =   1560
      TabIndex        =   24
      Top             =   5160
      Width           =   1365
   End
   Begin VB.Label AddressLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Height          =   240
      Left            =   1560
      TabIndex        =   23
      Top             =   4080
      Width           =   765
   End
   Begin VB.Label lblOccupation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
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
      Height          =   240
      Left            =   1560
      TabIndex        =   22
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label lblFatherName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name"
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
      Height          =   240
      Left            =   1560
      TabIndex        =   21
      Top             =   3120
      Width           =   1320
   End
   Begin VB.Label lblGPhone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      Height          =   240
      Left            =   6360
      TabIndex        =   20
      Top             =   4800
      Width           =   585
   End
   Begin VB.Label lblGAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Height          =   240
      Left            =   6360
      TabIndex        =   19
      Top             =   3600
      Width           =   765
   End
   Begin VB.Label lblGName 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   6360
      TabIndex        =   18
      Top             =   3120
      Width           =   555
   End
   Begin VB.Label Label1 
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
      Index           =   0
      Left            =   3240
      TabIndex        =   17
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   3240
      TabIndex        =   16
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   975
      Index           =   2
      Left            =   3240
      TabIndex        =   15
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Index           =   3
      Left            =   3240
      TabIndex        =   14
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Index           =   4
      Left            =   3240
      TabIndex        =   13
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Index           =   5
      Left            =   3240
      TabIndex        =   12
      Top             =   6140
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Index           =   6
      Left            =   7800
      TabIndex        =   11
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   1095
      Index           =   7
      Left            =   7800
      TabIndex        =   10
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   5175
      Left            =   960
      Top             =   2520
      Width           =   10695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "About Sisters and Brothers"
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
      Index           =   0
      Left            =   6360
      TabIndex        =   9
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label Label1 
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
      Index           =   8
      Left            =   7800
      TabIndex        =   8
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   855
      Left            =   960
      Top             =   7920
      Width           =   10695
   End
   Begin VB.Image imgReport 
      Height          =   450
      Left            =   9240
      Picture         =   "frmFamilyInfo.frx":1C2C1
      Top             =   8100
      Width           =   375
   End
   Begin VB.Image imgPrint 
      Height          =   450
      Left            =   10680
      Picture         =   "frmFamilyInfo.frx":1C878
      Top             =   8100
      Width           =   375
   End
   Begin VB.Image imgCancel 
      Height          =   375
      Left            =   1320
      Picture         =   "frmFamilyInfo.frx":1CC86
      Top             =   8100
      Width           =   375
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
      Left            =   1560
      TabIndex        =   7
      Top             =   1725
      Width           =   855
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
      Left            =   7560
      TabIndex        =   6
      Top             =   1725
      Width           =   615
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
      Left            =   4680
      TabIndex        =   5
      Top             =   1725
      Width           =   615
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
      Left            =   1560
      TabIndex        =   4
      Top             =   1275
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000D&
      Height          =   1455
      Left            =   960
      Top             =   960
      Width           =   10695
   End
End
Attribute VB_Name = "frmFamilyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()
  txtRegNo.SetFocus
  If txtRegNo.Text <> "" Then imgReport_Click
End Sub

Private Sub Form_Load()
   Theme frmFamilyInfo
End Sub

Private Sub imgCancel_Click()
  Unload Me
End Sub

Private Sub imgPrint_Click()
   On Error Resume Next
    
    imgPrint.Visible = False
    imgCancel.Visible = False
    imgReport.Visible = False
    
    PrintForm
    
    imgPrint.Visible = True
    imgCancel.Visible = True
    imgReport.Visible = True
End Sub

Private Sub txtName_Change()
  If txtName.Text = "" Then
     Label1(12).Enabled = True
     txtRegNo.Enabled = True
  Else
     Label1(12).Enabled = False
     txtRegNo.Enabled = False
  End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
  ValName KeyAscii
End Sub

Private Sub txtRegNo_Change()
  Dim i As Integer
  If txtRegNo.Text = "" Then
     Label2(1).Enabled = True
     Label3.Enabled = True
     Label4.Enabled = True
     txtName.Enabled = True
     cmbCourse.Enabled = True
     cmbYear.Enabled = True
  Else
     Label2(1).Enabled = False
     Label3.Enabled = False
     Label4.Enabled = False
     txtName.Enabled = False
     cmbCourse.Enabled = False
     cmbYear.Enabled = False
     For i = 0 To 9
       Label1(i).Caption = ""
     Next i
  End If
End Sub

Private Sub cmbCourse_Change()
  If cmbCourse.Text = "" Then
     Label1(12).Enabled = True
     txtRegNo.Enabled = True
  Else
     Label1(12).Enabled = False
     txtRegNo.Enabled = False
  End If
End Sub

Private Sub cmbYear_Change()
  If cmbYear.Text = "" Then
     Label1(12).Enabled = True
     txtRegNo.Enabled = True
  Else
     Label1(12).Enabled = False
     txtRegNo.Enabled = False
  End If
End Sub

Private Sub imgReport_Click()
  On Error Resume Next
  
  txtRegNo.Enabled = True
  If txtRegNo.Text <> "" Then
     If CheckRegNo(txtRegNo, 8) = True Then
        'do nothing
     End If
  Else
     connection
     reccheck
     Dim i As Integer
     i = 0
     rec.Open "select * from MAINTABLE where STUDENTNAME ='" & Trim(txtName.Text) & "' and COURSE ='" & Trim(cmbCourse.Text) & "' and YEAROFSTUDY ='" & Trim(cmbYear.Text) & "'", con, adOpenDynamic, adLockPessimistic
     If rec.EOF = False Then
       txtRegNo.Text = rec.Fields(0)
     Else
       MsgBox "No such result found", , "Personal Information"
       Exit Sub
     End If
  End If
  
  '------------------------Search for Records--------------------------------
  connection
  reccheck
  
  rec.Open "select * from FAMILYTABLE where regno=" & Trim(txtRegNo.Text), con, adOpenDynamic, adLockPessimistic
  If rec.EOF = False Then
     For i = 0 To 9
        Label1(i).Caption = rec.Fields(i + 1)
     Next i
  Else
    MsgBox "No such result found", , "Personal Information"
    Exit Sub
  End If
End Sub

Private Sub txtRegNo_KeyPress(KeyAscii As Integer)
   ValRegNo KeyAscii
End Sub
