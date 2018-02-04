VERSION 5.00
Begin VB.Form frmMarkPrint 
   BackColor       =   &H80000005&
   Caption         =   "Internal Mark"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14595
   FillColor       =   &H80000005&
   Icon            =   "frmMarkPrint.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmMarkPrint.frx":000C
   ScaleHeight     =   9315
   ScaleWidth      =   14595
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstSem 
      Height          =   255
      ItemData        =   "frmMarkPrint.frx":1C0C3
      Left            =   6000
      List            =   "frmMarkPrint.frx":1C0D9
      TabIndex        =   17
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtRegNo 
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      Height          =   400
      Index           =   2
      Left            =   10200
      Top             =   7320
      Width           =   375
   End
   Begin VB.Shape Shape6 
      Height          =   470
      Index           =   1
      Left            =   6960
      Top             =   7320
      Width           =   375
   End
   Begin VB.Shape Shape6 
      Height          =   470
      Index           =   0
      Left            =   8760
      Top             =   7320
      Width           =   375
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   3000
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   1440
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000D&
      Height          =   735
      Left            =   960
      Top             =   960
      Width           =   10215
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   9720
      TabIndex        =   49
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   9720
      TabIndex        =   48
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   9720
      TabIndex        =   47
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   9720
      TabIndex        =   46
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   9720
      TabIndex        =   45
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   9720
      TabIndex        =   44
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblSeminar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   8040
      TabIndex        =   43
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label lblSeminar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   8040
      TabIndex        =   42
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblSeminar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   8040
      TabIndex        =   41
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblSeminar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   8040
      TabIndex        =   40
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblSeminar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   39
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblSeminar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   38
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblAttendance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   6840
      TabIndex        =   37
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblAttendance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   36
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblAttendance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   35
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblAttendance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   6840
      TabIndex        =   34
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblAttendance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   33
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblAttendance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   32
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblTest2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   31
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label lblTest2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   30
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblTest2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   29
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lblTest2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   28
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblTest2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   27
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblTest2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   5760
      TabIndex        =   26
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblTest1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   25
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label lblTest1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   24
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblTest1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   23
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lblTest1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   22
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblTest1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   21
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblTest1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   20
      Top             =   3000
      Width           =   855
   End
   Begin VB.Image imgCancel 
      Height          =   375
      Left            =   10200
      Picture         =   "frmMarkPrint.frx":1C0EF
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image imgPrint 
      Height          =   450
      Left            =   8760
      Picture         =   "frmMarkPrint.frx":1C5CF
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image imgReport 
      Height          =   450
      Left            =   6960
      Picture         =   "frmMarkPrint.frx":1C9DD
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label1 
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
      Left            =   1440
      TabIndex        =   19
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Semester"
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
      Left            =   4800
      TabIndex        =   18
      Top             =   1200
      Width           =   870
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   975
      Left            =   960
      Top             =   7080
      Width           =   10200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Send via SMS"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   15
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Send via E-Mail"
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   14
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label lblPerfomance 
      BackColor       =   &H80000005&
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
      Left            =   2880
      TabIndex        =   13
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label lblPerfo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Perfomance"
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
      Left            =   960
      TabIndex        =   12
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   3375
      Left            =   960
      Top             =   2400
      Width           =   10200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   960
      X2              =   11160
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   4560
      X2              =   4560
      Y1              =   2400
      Y2              =   5760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   5640
      X2              =   5640
      Y1              =   2400
      Y2              =   5760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   2
      X1              =   6720
      X2              =   6720
      Y1              =   2400
      Y2              =   5760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   3
      X1              =   7920
      X2              =   7920
      Y1              =   2400
      Y2              =   5760
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subjects"
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
      Left            =   2280
      TabIndex        =   11
      Top             =   2520
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test 1"
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
      Left            =   4800
      TabIndex        =   10
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test 2"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance"
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
      Left            =   6840
      TabIndex        =   8
      Top             =   2520
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seminar / Viva"
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
      Left            =   8040
      TabIndex        =   7
      Top             =   2520
      Width           =   1305
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   4
      X1              =   9600
      X2              =   9600
      Y1              =   2400
      Y2              =   5760
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Mark"
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
      Left            =   9840
      TabIndex        =   6
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   960
      X2              =   11160
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   960
      X2              =   11160
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      Index           =   2
      X1              =   960
      X2              =   11160
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      Index           =   3
      X1              =   960
      X2              =   11160
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      Index           =   4
      X1              =   960
      X2              =   11160
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   3000
      Width           =   45
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   3480
      Width           =   45
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   3
      Left            =   1080
      TabIndex        =   2
      Top             =   4440
      Width           =   45
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   4
      Left            =   1080
      TabIndex        =   1
      Top             =   4920
      Width           =   45
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   5
      Left            =   1080
      TabIndex        =   0
      Top             =   5400
      Width           =   45
   End
End
Attribute VB_Name = "frmMarkPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_GotFocus()
  If txtRegNo.Text <> "" Then
     'report click
  End If
End Sub

Private Sub Form_Load()
  Dim i As Integer
  On Error Resume Next
  
  Theme frmMarkPrint
  
  lstSem.Text = lstSem.List(0)
  i = 0
  If txtRegNo.Text <> "" Then imgReport_Click

  If txtRegNo = "" Then
     'do nothing
  Else
  
    connection
    reccheck
    rec.Open ("select * from INTERNALMARKTABLE where regno = '" & Trim(txtRegNo.Text) & "' and sem = '" & Trim(lstSem.Text) & "'"), con, adOpenDynamic, adLockOptimistic
    Do While Not rec.EOF
       If i = 6 Then Exit Do
       lblSubject(i).Caption = rec.Fields(4)
       lblTest1(i).Caption = rec.Fields(5)
       lblTest2(i).Caption = rec.Fields(6)
       lblAttendance(i).Caption = rec.Fields(7)
       lblSeminar(i).Caption = rec.Fields(8)
       lblInternal(i).Caption = rec.Fields(9)
       i = i + 1
       rec.MoveNext
    Loop
  End If
End Sub

Private Sub imgCancel_Click()
  Unload Me
End Sub

Private Sub imgPrint_Click()
  On Error Resume Next
    
    imgPrint.Visible = False
    imgCancel.Visible = False
    imgReport.Visible = False
    Shape4.Visible = False
    Shape5.Visible = False
    Label2(0).Visible = False
    Label2(1).Visible = False
    
    PrintForm
    
    imgPrint.Visible = True
    imgCancel.Visible = True
    imgReport.Visible = True
    Shape4.Visible = True
    Shape5.Visible = True
    Label2(0).Visible = True
    Label2(1).Visible = True
    
End Sub
Private Sub GradeCalc(mark As Integer, ByRef lbl As Label)
    Select Case mark
        Case 5: lbl.Caption = "A"
        Case 4: lbl.Caption = "B"
        Case 3: lbl.Caption = "C"
        Case 2: lbl.Caption = "D"
        Case 1: lbl.Caption = "E"
    End Select
End Sub

Private Sub imgReport_Click()
  On Error Resume Next
  If txtRegNo.Text <> "" Then
     If CheckRegNo(txtRegNo, 8) = True Then
        'do nothing
     End If
  End If
  
  '------------------------Search for Records--------------------------------
  Dim i As Integer, perfomance As Integer
  perfomance = 0
  connection
  reccheck
  i = 0
  rec.Open ("select * from INTERNALMARKTABLE where REGNO = '" & Trim(txtRegNo.Text) & "' and SEM = '" & Trim(lstSem.Text) & "'"), con, adOpenDynamic, adLockOptimistic
  If rec.EOF = False Then
     Do While Not rec.EOF
        If i = 6 Then Exit Do
        
        lblSubject(i).Caption = rec.Fields(4)
        
        If rec.Fields(5) < 6 Then
        Select Case rec.Fields(5)
           Case 5: lblTest1(i).Caption = "A"
           Case 4: lblTest1(i).Caption = "B"
           Case 3: lblTest1(i).Caption = "C"
           Case 2: lblTest1(i).Caption = "D"
           Case 1: lblTest1(i).Caption = "E"
        End Select
        End If
        If rec.Fields(6) < 6 Then GradeCalc rec.Fields(6), lblTest2(i)
        lblAttendance(i).Caption = rec.Fields(7)
        If rec.Fields(8) < 6 Then GradeCalc rec.Fields(8), lblSeminar(i)
        If rec.Fields(9) < 6 Then
           GradeCalc rec.Fields(9), lblInternal(i)
           perfomance = perfomance + rec.Fields(9) * 4
        End If
        i = i + 1
        perfomance = perfomance + rec.Fields(9)
        rec.MoveNext
     Loop
     If perfomance > 90 Then lblPerfomance.Caption = "Excellent"
     If perfomance > 70 And perfomance < 91 Then lblPerfomance.Caption = "Good"
     If perfomance > 30 And perfomance < 71 Then lblPerfomance.Caption = "Satisfied"
     If perfomance < 31 Then lblPerfomance.Caption = "Poor"
     
  Else
     MsgBox "No such record found", , "Internal Mark"
  End If
End Sub

Private Sub Label2_Click(Index As Integer)
   MsgBox "No internet connection", , "Sending Error"
End Sub

Private Sub txtRegNo_Change()
   Dim i As Integer
   For i = 0 To 5
     lblSubject(i).Caption = ""
     lblTest1(i).Caption = ""
     lblTest2(i).Caption = ""
     lblAttendance(i).Caption = ""
     lblSeminar(i).Caption = ""
     lblInternal(i).Caption = ""
   Next i
End Sub

Private Sub txtRegNo_KeyPress(KeyAscii As Integer)
   ValRegNo KeyAscii
End Sub

Private Sub txtRegNo_LostFocus()
   CheckRegNo txtRegNo, 8
End Sub
