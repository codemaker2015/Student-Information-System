VERSION 5.00
Begin VB.Form frmEducationalInfo 
   Caption         =   "Educational Information"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   Icon            =   "frmEducationalInfo.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "frmEducationalInfo.frx":000C
   ScaleHeight     =   11010
   ScaleWidth      =   14865
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbYear 
      Height          =   315
      ItemData        =   "frmEducationalInfo.frx":26884E
      Left            =   8640
      List            =   "frmEducationalInfo.frx":2688DC
      TabIndex        =   3
      Text            =   "Year"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox cmbCourse 
      Height          =   315
      ItemData        =   "frmEducationalInfo.frx":2689F4
      Left            =   5640
      List            =   "frmEducationalInfo.frx":268A10
      TabIndex        =   2
      Text            =   "--Select--"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2760
      MaxLength       =   35
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtRegNo 
      Height          =   375
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      Height          =   400
      Index           =   2
      Left            =   1560
      Top             =   8460
      Width           =   375
   End
   Begin VB.Shape Shape4 
      Height          =   465
      Index           =   1
      Left            =   12240
      Top             =   8460
      Width           =   375
   End
   Begin VB.Shape Shape4 
      Height          =   470
      Index           =   0
      Left            =   10800
      Top             =   8460
      Width           =   375
   End
   Begin VB.Label lblGrade 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   10800
      TabIndex        =   36
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label lblGrade 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   10800
      TabIndex        =   35
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblGrade 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   10800
      TabIndex        =   34
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Grade"
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
      Index           =   1
      Left            =   10920
      TabIndex        =   33
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   9480
      TabIndex        =   32
      Top             =   3840
      Width           =   975
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      X1              =   9360
      X2              =   12000
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   5
      X1              =   12000
      X2              =   12000
      Y1              =   3240
      Y2              =   6000
   End
   Begin VB.Label lblChance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   12240
      TabIndex        =   31
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblChance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   12240
      TabIndex        =   30
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblChance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   12240
      TabIndex        =   29
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblMarks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   9480
      TabIndex        =   28
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblMarks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   9480
      TabIndex        =   27
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblMarks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   9480
      TabIndex        =   26
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblSchool 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   25
      Top             =   5520
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lblSchool 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   24
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Label lblSchool 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   23
      Top             =   4320
      Width           =   3735
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   22
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   21
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   20
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblRegNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblRegNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   18
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblRegNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   17
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblCourse 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   16
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblCourse 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   15
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblCourse 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   14
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      Index           =   2
      X1              =   1200
      X2              =   13200
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   1200
      X2              =   13200
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   4
      X1              =   10680
      X2              =   10680
      Y1              =   3720
      Y2              =   6000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   3
      X1              =   9360
      X2              =   9360
      Y1              =   3240
      Y2              =   6000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   2
      X1              =   5400
      X2              =   5400
      Y1              =   3240
      Y2              =   6000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   4200
      X2              =   4200
      Y1              =   3240
      Y2              =   6000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   3000
      X2              =   3000
      Y1              =   3240
      Y2              =   6000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   1200
      X2              =   13200
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   2775
      Left            =   1200
      Top             =   3240
      Width           =   12015
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Course"
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
      Left            =   1320
      TabIndex        =   13
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reg. No."
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
      Left            =   3240
      TabIndex        =   12
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Year of Passing"
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
      Height          =   555
      Left            =   4440
      TabIndex        =   11
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of the Institution"
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
      Left            =   6240
      TabIndex        =   10
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marks"
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
      Left            =   10440
      TabIndex        =   9
      Top             =   3360
      Width           =   555
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of    Chances"
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
      Height          =   435
      Left            =   12240
      TabIndex        =   8
      Top             =   3480
      Width           =   840
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
      Left            =   1800
      TabIndex        =   7
      Top             =   1845
      Width           =   855
   End
   Begin VB.Label Label5 
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
      Left            =   7800
      TabIndex        =   6
      Top             =   1845
      Width           =   615
   End
   Begin VB.Label Label4 
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
      Left            =   4920
      TabIndex        =   5
      Top             =   1845
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1395
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000D&
      Height          =   1455
      Left            =   1200
      Top             =   1080
      Width           =   12015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   855
      Left            =   1200
      Top             =   8280
      Width           =   12015
   End
   Begin VB.Image imgReport 
      Height          =   450
      Left            =   10800
      Picture         =   "frmEducationalInfo.frx":268A4C
      Top             =   8460
      Width           =   375
   End
   Begin VB.Image imgPrint 
      Height          =   450
      Left            =   12240
      Picture         =   "frmEducationalInfo.frx":269003
      Top             =   8460
      Width           =   375
   End
   Begin VB.Image imgCancel 
      Height          =   375
      Left            =   1560
      Picture         =   "frmEducationalInfo.frx":269411
      Top             =   8460
      Width           =   375
   End
End
Attribute VB_Name = "frmEducationalInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_GotFocus()
  txtRegNo.SetFocus
  If txtRegNo.Text <> "" Then imgReport_Click
End Sub

Private Sub Form_Load()
  Theme frmEducationalInfo
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
     Label4.Enabled = True
     Label5.Enabled = True
     txtName.Enabled = True
     cmbCourse.Enabled = True
     cmbYear.Enabled = True
  Else
     Label2(1).Enabled = False
     Label4.Enabled = False
     Label5.Enabled = False
     txtName.Enabled = False
     cmbCourse.Enabled = False
     cmbYear.Enabled = False
     For i = 0 To 2
        lblCourse(i).Caption = ""
        lblRegNo(i).Caption = ""
        lblYear(i).Caption = ""
        lblSchool(i).Caption = ""
        lblMarks(i).Caption = ""
        lblGrade(i).Caption = ""
        lblChance(i).Caption = ""
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
'  On Error Resume Next
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
  Dim i As Integer, J As Integer
  i = 0
  J = 0
  rec.Open "select * from EXAMTABLE where regno=" & Trim(txtRegNo.Text), con, adOpenDynamic, adLockPessimistic
  If rec.EOF = False Then
     Do While Not rec.EOF
       If J = 3 Then Exit Do
       lblCourse(i).Caption = rec.Fields(2)
       lblRegNo(i).Caption = rec.Fields(1)
       lblYear(i).Caption = rec.Fields(3)
       lblSchool(i).Caption = rec.Fields(4)
       lblGrade(i).Caption = rec.Fields(5)
       lblMarks(i).Caption = rec.Fields(6)
       lblChance(i).Caption = rec.Fields(7)
       rec.MoveNext
       i = i + 1
       J = J + 1
     Loop
  Else
     MsgBox "No such record found", , "Educational Information"
     Exit Sub
  End If
End Sub

Private Sub txtRegNo_KeyPress(KeyAscii As Integer)
  ValRegNo KeyAscii
End Sub
