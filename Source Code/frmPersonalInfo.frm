VERSION 5.00
Begin VB.Form frmPersonalInfo 
   Caption         =   "Personal Information"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14940
   Icon            =   "frmPersonalInfo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmPersonalInfo.frx":000C
   ScaleHeight     =   9495
   ScaleWidth      =   14940
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRegNo 
      Height          =   375
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   24
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2640
      MaxLength       =   35
      TabIndex        =   25
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ComboBox cmbCourse 
      Height          =   315
      ItemData        =   "frmPersonalInfo.frx":1C0C3
      Left            =   5520
      List            =   "frmPersonalInfo.frx":1C0DF
      TabIndex        =   26
      Text            =   "--Select--"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      ItemData        =   "frmPersonalInfo.frx":1C11B
      Left            =   8520
      List            =   "frmPersonalInfo.frx":1C1A9
      TabIndex        =   27
      Text            =   "Year"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Image imgAttach 
      Height          =   375
      Left            =   11640
      Picture         =   "frmPersonalInfo.frx":1C2C1
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape6 
      Height          =   470
      Index           =   1
      Left            =   11760
      Top             =   8040
      Width           =   375
   End
   Begin VB.Shape Shape6 
      Height          =   470
      Index           =   0
      Left            =   10320
      Top             =   8040
      Width           =   375
   End
   Begin VB.Shape Shape5 
      Height          =   375
      Left            =   1320
      Top             =   8040
      Width           =   375
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000D&
      Height          =   1935
      Left            =   10680
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Image imgPhoto 
      Height          =   1650
      Left            =   10800
      Picture         =   "frmPersonalInfo.frx":1E23A
      Top             =   2880
      Width           =   1500
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000D&
      Height          =   1695
      Left            =   1080
      Top             =   600
      Width           =   11535
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
      Left            =   1680
      TabIndex        =   31
      Top             =   1035
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
      Left            =   4800
      TabIndex        =   30
      Top             =   1485
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
      Left            =   7680
      TabIndex        =   29
      Top             =   1485
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
      Left            =   1680
      TabIndex        =   28
      Top             =   1485
      Width           =   855
   End
   Begin VB.Image imgCancel 
      Height          =   375
      Left            =   1320
      Picture         =   "frmPersonalInfo.frx":1EF44
      Top             =   8040
      Width           =   375
   End
   Begin VB.Image imgPrint 
      Height          =   450
      Left            =   11760
      Picture         =   "frmPersonalInfo.frx":1F424
      Top             =   8040
      Width           =   375
   End
   Begin VB.Image imgReport 
      Height          =   450
      Left            =   10320
      Picture         =   "frmPersonalInfo.frx":1F832
      Top             =   8040
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   855
      Left            =   1080
      Top             =   7860
      Width           =   11535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   4815
      Left            =   1080
      Top             =   2640
      Width           =   11535
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
      Index           =   11
      Left            =   8280
      TabIndex        =   23
      Top             =   6000
      Width           =   3135
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
      Height          =   735
      Index           =   10
      Left            =   8280
      TabIndex        =   22
      Top             =   5040
      Width           =   3255
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
      Index           =   9
      Left            =   8280
      TabIndex        =   21
      Top             =   4440
      Width           =   1935
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
      Left            =   8280
      TabIndex        =   20
      Top             =   3840
      Width           =   2655
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
      Index           =   7
      Left            =   8280
      TabIndex        =   19
      Top             =   3240
      Width           =   1935
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
      Left            =   3240
      TabIndex        =   18
      Top             =   6840
      Width           =   3135
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
      TabIndex        =   17
      Top             =   6240
      Width           =   3135
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
      TabIndex        =   16
      Top             =   5640
      Width           =   3135
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
      TabIndex        =   15
      Top             =   5040
      Width           =   3135
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
      Index           =   2
      Left            =   3240
      TabIndex        =   14
      Top             =   4440
      Width           =   3135
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
      TabIndex        =   13
      Top             =   3840
      Width           =   3135
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
      TabIndex        =   12
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label lblRegNo 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   1680
      TabIndex        =   11
      Top             =   3240
      Width           =   705
   End
   Begin VB.Label lblName 
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
      Left            =   1680
      TabIndex        =   10
      Top             =   3840
      Width           =   555
   End
   Begin VB.Label YearOfStudyLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year of study"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   4440
      Width           =   1170
   End
   Begin VB.Label lblAddress 
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
      Left            =   6720
      TabIndex        =   8
      Top             =   4920
      Width           =   765
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
      Index           =   0
      Left            =   6720
      TabIndex        =   7
      Top             =   6000
      Width           =   585
   End
   Begin VB.Label lblProgramme 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Programe"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   5040
      Width           =   900
   End
   Begin VB.Label lblGender 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   5640
      Width           =   675
   End
   Begin VB.Label lblReligion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   6240
      Width           =   750
   End
   Begin VB.Label lblCast 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cast"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   6840
      Width           =   405
   End
   Begin VB.Label lblBloodGroup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Group"
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
      Left            =   6720
      TabIndex        =   2
      Top             =   3240
      Width           =   1140
   End
   Begin VB.Label lblIncome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anual Income"
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
      Left            =   6720
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblDBirth 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
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
      Left            =   6720
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Photo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   10920
      TabIndex        =   32
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "frmPersonalInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fso As New FileSystemObject

Private Sub Form_GotFocus()
   If txtRegNo.Text <> "" Then Call imgReport_Click
End Sub

Private Sub Form_Load()
   Theme frmPersonalInfo
End Sub

Private Sub imgAttach_Click()
  frmAttachment.txtRegNo.Text = txtRegNo.Text
  frmAttachment.imgSave.Enabled = False
  frmAttachment.Show
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
     Label2.Enabled = True
     Label3.Enabled = True
     Label4.Enabled = True
     txtName.Enabled = True
     cmbCourse.Enabled = True
     cmbYear.Enabled = True
  Else
     Label2.Enabled = False
     Label3.Enabled = False
     Label4.Enabled = False
     txtName.Enabled = False
     cmbCourse.Enabled = False
     cmbYear.Enabled = False
     For i = 0 To 11
       Label1(i).Caption = ""
     Next i
     imgPhoto.Picture = LoadPicture()
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
  
  Dim i As Integer
  rec.Open "select * from PERSONALTABLE where regno=" & Trim(txtRegNo.Text), con, adOpenDynamic, adLockPessimistic
  If rec.EOF = False Then
     For i = 0 To 11
        If i <> 8 Then Label1(i).Caption = rec.Fields(i)
        If i = 9 Then Label1(i) = Format(rec.Fields(i), "dd-mm-yyyy")
        If i = 4 Then
           If rec.Fields(i) = "M" Then
              Label1(i).Caption = "Male"
           End If
           If rec.Fields(i) = "F" Then
              Label1(i).Caption = "Male"
           End If
        End If
     Next i
     Select Case rec.Fields(8)
        Case 25000: Label1(8) = "Below 25000"
        Case 50000: Label1(8) = "Between 25000 and 50000"
        Case 100000: Label1(8) = "Between 50000 and 2 Lakh"
        Case 200000: Label1(8) = "Above 2 Lakh"
     End Select
     If fso.FileExists(rec.Fields(12)) = True Then imgPhoto.Picture = LoadPicture(rec.Fields(12))
  Else
     MsgBox "No such record found", , "Personal Table"
  End If
  
End Sub

Private Sub txtRegNo_KeyPress(KeyAscii As Integer)
   ValRegNo KeyAscii
End Sub
