VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmPerfomance 
   Caption         =   "Internal Mark"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16395
   Icon            =   "MarkGraphfrm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "MarkGraphfrm.frx":000C
   ScaleHeight     =   10215
   ScaleWidth      =   16395
   WindowState     =   2  'Maximized
   Begin MSChart20Lib.MSChart MSChartMark 
      Height          =   4695
      Left            =   1080
      OleObjectBlob   =   "MarkGraphfrm.frx":1C0C3
      TabIndex        =   0
      Top             =   2400
      Width           =   8055
   End
   Begin VB.ListBox lstSem 
      Height          =   255
      ItemData        =   "MarkGraphfrm.frx":1D936
      Left            =   5640
      List            =   "MarkGraphfrm.frx":1D94C
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtRegNo 
      Height          =   375
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   5
      Left            =   10080
      TabIndex        =   10
      Top             =   6040
      Width           =   4335
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   4
      Left            =   10080
      TabIndex        =   9
      Top             =   5440
      Width           =   4335
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   3
      Left            =   10080
      TabIndex        =   8
      Top             =   4840
      Width           =   4335
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   2
      Left            =   10080
      TabIndex        =   7
      Top             =   4240
      Width           =   4335
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   10080
      TabIndex        =   6
      Top             =   3640
      Width           =   4335
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   10080
      TabIndex        =   5
      Top             =   3040
      Width           =   4335
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   9600
      Top             =   6000
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   9600
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   9600
      Top             =   4800
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   9600
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   9600
      Top             =   3600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   9600
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape4 
      Height          =   405
      Index           =   2
      Left            =   960
      Top             =   8160
      Width           =   375
   End
   Begin VB.Shape Shape4 
      Height          =   480
      Index           =   1
      Left            =   12120
      Top             =   8160
      Width           =   375
   End
   Begin VB.Shape Shape4 
      Height          =   480
      Index           =   0
      Left            =   13440
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image imgCancel 
      Height          =   375
      Left            =   960
      Picture         =   "MarkGraphfrm.frx":1D962
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image imgPrint 
      Height          =   450
      Left            =   13440
      Picture         =   "MarkGraphfrm.frx":1DE42
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image imgReport 
      Height          =   450
      Left            =   12120
      Picture         =   "MarkGraphfrm.frx":1E250
      Top             =   8160
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000D&
      Height          =   855
      Left            =   600
      Top             =   7920
      Width           =   14175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   5535
      Left            =   600
      Top             =   2040
      Width           =   14175
   End
   Begin VB.Label lblSem 
      BackStyle       =   0  'Transparent
      Caption         =   "Semester:"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblRegNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg No:"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   975
      Left            =   600
      Top             =   600
      Width           =   14175
   End
End
Attribute VB_Name = "frmPerfomance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_GotFocus()
  txtRegNo.SetFocus
End Sub

Private Sub Form_Load()
  Theme frmPerfomance

  lstSem.Text = lstSem.List(0)
  connection
  
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

Private Sub imgReport_Click()
  Dim arrVal(1 To 8)
  Dim i As Integer
  Dim course As String
  course = " "
  i = 1
  reccheck
  rec.Open "select SUBJECTNAME,INTERNALMARK from INTERNALMARKTABLE where regno = '" & Trim(txtRegNo.Text) & "' and sem = '" & Trim(lstSem.Text) & "'", con, adOpenDynamic, adLockOptimistic
  If rec.EOF = False Then
    
    While Not rec.EOF
       arrVal(i) = rec.Fields(1)
       lblSubject(i - 1).Caption = rec.Fields(0)
       i = i + 1
       rec.MoveNext
     Wend
     MSChartMark.ChartData = arrVal
     MSChartMark.Refresh
  Else
     MsgBox "No such record found"
  End If
End Sub

Private Sub txtRegNo_KeyPress(KeyAscii As Integer)
  ValRegNo KeyAscii
End Sub

Private Sub txtRegNo_LostFocus()
  CheckRegNo txtRegNo, 8
End Sub
