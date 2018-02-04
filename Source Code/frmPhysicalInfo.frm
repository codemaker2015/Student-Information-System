VERSION 5.00
Begin VB.Form frmPhysical 
   Caption         =   "Extra Curricular Activities"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16350
   Icon            =   "frmPhysicalInfo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmPhysicalInfo.frx":000C
   ScaleHeight     =   9705
   ScaleWidth      =   16350
   Begin VB.TextBox txtRegNo 
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox cmbPosition 
      Height          =   315
      ItemData        =   "frmPhysicalInfo.frx":26884E
      Left            =   4440
      List            =   "frmPhysicalInfo.frx":268861
      TabIndex        =   8
      Text            =   "--Select--"
      Top             =   4560
      Width           =   3375
   End
   Begin VB.ComboBox cmbSubCategory 
      Height          =   315
      Left            =   4440
      TabIndex        =   7
      Text            =   "--Select--"
      Top             =   3920
      Width           =   3375
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      ItemData        =   "frmPhysicalInfo.frx":2688A8
      Left            =   4440
      List            =   "frmPhysicalInfo.frx":2688BB
      TabIndex        =   6
      Text            =   "--Select--"
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1920
      Width           =   855
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
      Left            =   2520
      TabIndex        =   5
      Top             =   4600
      Width           =   1335
   End
   Begin VB.Label lblAdd 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
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
      Height          =   315
      Left            =   9720
      TabIndex        =   4
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
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
      Height          =   315
      Left            =   11520
      TabIndex        =   3
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
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
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   1215
      Left            =   2280
      Top             =   6240
      Width           =   10815
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Index           =   1
      Left            =   2640
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      Height          =   495
      Left            =   9720
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      Height          =   495
      Left            =   11520
      Top             =   6600
      Width           =   1095
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
      Left            =   2520
      TabIndex        =   1
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
      Left            =   2520
      TabIndex        =   0
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Image imgAttach 
      Height          =   375
      Left            =   9240
      Picture         =   "frmPhysicalInfo.frx":2688DF
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000D&
      Height          =   1095
      Index           =   0
      Left            =   2280
      Top             =   1560
      Width           =   10815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   2895
      Left            =   2280
      Top             =   3000
      Width           =   10815
   End
End
Attribute VB_Name = "frmPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim attach As Boolean
Private Sub Form_Load()
  Theme frmPhysical
  attach = False
  connection
End Sub

Private Sub imgAttach_Click()
  frmAttachment.txtRegNo.Text = txtRegNo.Text
  frmAttachment.imgSave.Enabled = False
  frmAttachment.Show
  attach = True
End Sub

Private Sub lblAdd_Click()
   On Error GoTo error_para
   
   If attach = False Then
     MsgBox "You should attach the documents before submitting details", vbCritical, ""
   Else
     'code for adding marks
     
     If CheckRegNo(txtRegNo, 8) = True Then
        reccheck
        rec.Open "select * from PHYSICALTABLE where REGNO = " & txtRegNo.Text
        If rec.EOF = False Then
           If rec.Fields(1) = cmbSubCategory.Text Then
              If MsgBox("Database already contains this information. Do you wish to overwrite it", vbYesNo) = vbYes Then
                 reccheck
                 rec.Open "delete PHYSICALTABLE where SUBCATEGORY = '" & cmbSubCategory.Text & "'", con, adOpenDynamic, adLockPessimistic
                 reccheck
                 rec.Open "insert into PHYSICALTABLE values('" & Val(Trim(txtRegNo.Text)) & "','" & Trim(cmbCategory.Text) & "','" & Trim(cmbSubCategory.Text) & "','" & Trim(cmbPosition.Text) & "')", con, adOpenDynamic, adLockPessimistic
                 MsgBox "Record added suceesfully"
              End If
           End If
        End If
     Else
        MsgBox "Reg No entered is incorrect", vbCritical
     End If
   End If
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub lblExit_Click()
  Unload Me
End Sub

Private Sub txtRegNo_KeyPress(KeyAscii As Integer)
   ValRegNo KeyAscii
End Sub

