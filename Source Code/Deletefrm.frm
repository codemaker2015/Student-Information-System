VERSION 5.00
Begin VB.Form frmDelete 
   Caption         =   "Delete Student Information"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
   Icon            =   "Deletefrm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Deletefrm.frx":000C
   ScaleHeight     =   8820
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      ItemData        =   "Deletefrm.frx":1C0C3
      Left            =   3960
      List            =   "Deletefrm.frx":1C0D9
      TabIndex        =   9
      Text            =   "All"
      Top             =   4200
      Width           =   1935
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      ItemData        =   "Deletefrm.frx":1C123
      Left            =   10080
      List            =   "Deletefrm.frx":1C1B1
      TabIndex        =   6
      Text            =   "Year"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox cmbCourse 
      Height          =   315
      ItemData        =   "Deletefrm.frx":1C2C9
      Left            =   7080
      List            =   "Deletefrm.frx":1C2E5
      TabIndex        =   5
      Text            =   "--Select--"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   3960
      MaxLength       =   35
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtDelete 
      Height          =   375
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   2040
      Top             =   5760
      Width           =   10215
   End
   Begin VB.Shape Shape1 
      Height          =   2775
      Left            =   2040
      Top             =   2640
      Width           =   10215
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   4245
      Width           =   975
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
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3525
      Width           =   855
   End
   Begin VB.Image cmdCancel 
      Height          =   375
      Left            =   2400
      Picture         =   "Deletefrm.frx":1C321
      Top             =   6000
      Width           =   375
   End
   Begin VB.Image cmdOK 
      Height          =   375
      Left            =   11520
      Picture         =   "Deletefrm.frx":1C801
      Top             =   6000
      Width           =   375
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
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   3525
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
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   3525
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   3075
      Width           =   855
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error GoTo error_para
  
  If txtDelete.Text <> "" Then
    CheckRegNo txtDelete, 8
  Else
    MsgBox "Register Number lefts Empty", , "Delete"
    Exit Sub
  End If
  '---------------------------------delete records------------------------------------
  
  connection
  reccheck
  Select Case cmbCategory.Text
         Case "Personal":
                             rec.Open "select * from PERSONALTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                             If rec.EOF = False Then
                               connection
                               reccheck
                               rec.Open "delete from PERSONALTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                               MsgBox "Record Deleted Successfully", vbInformation, "Delete Information"
                               txtDelete.Text = ""
                             Else
                               MsgBox "No such record found", vbInformation, "Delete Information"
                               Exit Sub
                             End If
         Case "Educational":
                             rec.Open "select * from EXAMTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                             If rec.EOF = False Then
                               connection
                               reccheck
                               rec.Open "delete from EXAMTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                               MsgBox "Record Deleted Successfully", vbInformation, "Delete Information"
                               txtDelete.Text = ""
                             Else
                               MsgBox "No such record found", vbInformation, "Delete Information"
                               Exit Sub
                             End If
         Case "Family":      rec.Open "select * from FAMILYTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                             If rec.EOF = False Then
                               connection
                               reccheck
                               rec.Open "delete from FAMILYTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                               MsgBox "Record Deleted Successfully", vbInformation, "Delete Information"
                               txtDelete.Text = ""
                             Else
                               MsgBox "No such record found", vbInformation, "Delete Information"
                               Exit Sub
                             End If
         Case "Extra Curiccular":  rec.Open "select * from PHYSICALTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                             If rec.EOF = False Then
                               connection
                               reccheck
                               rec.Open "delete from PHYSICALTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                               MsgBox "Record Deleted Successfully", vbInformation, "Delete Information"
                               txtDelete.Text = ""
                             Else
                               MsgBox "No such record found", vbInformation, "Delete Information"
                               Exit Sub
                             End If
         Case "Internal Exams":  rec.Open "select * from INTERNALMARKTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                             If rec.EOF = False Then
                               connection
                               reccheck
                               rec.Open "delete from INTERNALMARKTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                               MsgBox "Record Deleted Successfully", vbInformation, "Delete Information"
                               txtDelete.Text = ""
                             Else
                               MsgBox "No such record found", vbInformation, "Delete Information"
                               Exit Sub
                             End If
         Case "All":   rec.Open "delete from PERSONALTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                       connection
                       reccheck
                       rec.Open "delete from EXAMTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                       connection
                       reccheck
                       rec.Open "delete from FAMILYTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                       connection
                       reccheck
                       rec.Open "delete from PHYSICALTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                       connection
                       reccheck
                       rec.Open "delete from INTERNALMARKTABLE where REGNO =" & Trim(txtDelete.Text), con, adOpenDynamic, adLockPessimistic
                       MsgBox "Record Deleted Successfully", vbInformation, "Delete Information"
                       txtDelete.Text = ""
  End Select
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbYear_LostFocus()
  On Error GoTo error_para
  
  If txtDelete.Text <> "" Then
     If CheckRegNo(txtDelete, 8) = True Then
        'do nothing
     End If
  Else
     connection
     reccheck
     
     rec.Open "select REGNO from MAINTABLE where COURSE ='" & Trim(cmbCourse.Text) & "' and YEAROFSTUDY ='" & Trim(cmbYear.Text) & "' and " & "STUDENTNAME like '" & Trim(txtName.Text) & "%'", con, adOpenDynamic, adLockPessimistic
     If rec.EOF = False Then
       txtDelete.Text = rec.Fields(0)
     Else
       MsgBox "No such result found", vbInformation, "Delete Information"
       Exit Sub
     End If
  End If
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub Form_GotFocus()
  Unload frmLoading
End Sub

Private Sub Form_Load()
   Theme frmDelete
End Sub

Private Sub txtDelete_Change()
  If txtDelete.Text = "" Then
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
  End If
End Sub

Private Sub txtName_Change()
  If txtName.Text = "" Then
     Label1.Enabled = True
     txtDelete.Enabled = True
  Else
     Label1.Enabled = False
     txtDelete.Enabled = False
  End If
End Sub

Private Sub cmbCourse_Change()
  If cmbCourse.Text = "" Then
     Label1.Enabled = True
     txtDelete.Enabled = True
  Else
     Label1.Enabled = False
     txtDelete.Enabled = False
  End If
End Sub

Private Sub cmbYear_Change()
  If cmbYear.Text = "" Then
     Label1.Enabled = True
     txtDelete.Enabled = True
  Else
     Label1.Enabled = False
     txtDelete.Enabled = False
  End If
End Sub
