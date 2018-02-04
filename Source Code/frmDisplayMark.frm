VERSION 5.00
Begin VB.Form frmDisplayMark 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Internal Marks"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14970
   Icon            =   "frmDisplayMark.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmDisplayMark.frx":000C
   ScaleHeight     =   8850
   ScaleWidth      =   14970
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstSem 
      Height          =   255
      ItemData        =   "frmDisplayMark.frx":1C0C3
      Left            =   7440
      List            =   "frmDisplayMark.frx":1C0D9
      TabIndex        =   14
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtRegNo 
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ListBox lstCourse 
      Height          =   255
      ItemData        =   "frmDisplayMark.frx":1C0EF
      Left            =   4560
      List            =   "frmDisplayMark.frx":1C108
      TabIndex        =   12
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7560
      Picture         =   "frmDisplayMark.frx":1C13D
      Top             =   7560
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   5880
      Picture         =   "frmDisplayMark.frx":1CE07
      Top             =   7440
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send"
      Height          =   495
      Left            =   9360
      TabIndex        =   48
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label12 
      Height          =   255
      Index           =   5
      Left            =   9240
      TabIndex        =   47
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label12 
      Height          =   255
      Index           =   4
      Left            =   9240
      TabIndex        =   46
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label12 
      Height          =   255
      Index           =   3
      Left            =   9240
      TabIndex        =   45
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label12 
      Height          =   255
      Index           =   2
      Left            =   9240
      TabIndex        =   44
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label12 
      Height          =   255
      Index           =   1
      Left            =   9240
      TabIndex        =   43
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label11 
      Height          =   255
      Index           =   5
      Left            =   7680
      TabIndex        =   42
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label11 
      Height          =   255
      Index           =   4
      Left            =   7680
      TabIndex        =   41
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label11 
      Height          =   255
      Index           =   3
      Left            =   7680
      TabIndex        =   40
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label11 
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   39
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label11 
      Height          =   255
      Index           =   1
      Left            =   7680
      TabIndex        =   38
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label10 
      Height          =   255
      Index           =   5
      Left            =   6480
      TabIndex        =   37
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label10 
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   36
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label10 
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   35
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label10 
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   34
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label10 
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   33
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lbltest2 
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   32
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label lbltest2 
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   31
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label lbltest2 
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   30
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label lbltest2 
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   29
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lbltest2 
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   28
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label12 
      Height          =   255
      Index           =   0
      Left            =   9240
      TabIndex        =   27
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label11 
      Height          =   255
      Index           =   0
      Left            =   7680
      TabIndex        =   26
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label10 
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   25
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lbltest2 
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   24
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lbltest1 
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   23
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label lbltest1 
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   22
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label lbltest1 
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   21
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label lbltest1 
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   20
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lbltest1 
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   19
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lbltest1 
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   18
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblRegno 
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
      Index           =   1
      Left            =   1440
      TabIndex        =   17
      Top             =   2520
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
      Left            =   6480
      TabIndex        =   16
      Top             =   2520
      Width           =   870
   End
   Begin VB.Label lblCourse 
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
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   3375
      Left            =   1440
      Top             =   3360
      Width           =   9255
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   10680
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   4200
      X2              =   4200
      Y1              =   3360
      Y2              =   6720
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   5280
      X2              =   5280
      Y1              =   3360
      Y2              =   6720
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   6360
      X2              =   6360
      Y1              =   3360
      Y2              =   6720
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   7560
      X2              =   7560
      Y1              =   3360
      Y2              =   6720
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
      Top             =   3480
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
      Left            =   4440
      TabIndex        =   10
      Top             =   3480
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
      Left            =   5520
      TabIndex        =   9
      Top             =   3480
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
      Left            =   6480
      TabIndex        =   8
      Top             =   3480
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
      Left            =   7680
      TabIndex        =   7
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   9120
      X2              =   9120
      Y1              =   3360
      Y2              =   6720
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
      Left            =   9360
      TabIndex        =   6
      Top             =   3480
      Width           =   1140
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   1440
      X2              =   10680
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   1440
      X2              =   10680
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   1440
      X2              =   10680
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line3 
      Index           =   3
      X1              =   1440
      X2              =   10680
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line3 
      Index           =   4
      X1              =   1440
      X2              =   10680
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 0"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   3960
      Width           =   675
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Top             =   4440
      Width           =   675
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 2"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   4920
      Width           =   675
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 3"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   3
      Left            =   1680
      TabIndex        =   2
      Top             =   5400
      Width           =   675
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 4"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   4
      Left            =   1680
      TabIndex        =   1
      Top             =   5880
      Width           =   675
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 5"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   5
      Left            =   1680
      TabIndex        =   0
      Top             =   6360
      Width           =   675
   End
End
Attribute VB_Name = "frmDisplayMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Theme frmDisplayMark
End Sub

Private Sub Image2_Click()
  Unload Me
End Sub

Private Sub Label1_Click()
 ' If my.computer.network.isavailable = True Then
  '   MsgBox "Connection Ok"
 ' Else
 '   MsgBox "Connection not Ok"
'  End If
End Sub

