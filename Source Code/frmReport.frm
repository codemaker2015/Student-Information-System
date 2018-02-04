VERSION 5.00
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16380
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmReport.frx":000C
   ScaleHeight     =   11010
   ScaleWidth      =   16380
   WindowState     =   2  'Maximized
   Begin VB.Shape Shape8 
      BorderColor     =   &H8000000D&
      Height          =   5775
      Left            =   2400
      Top             =   1200
      Width           =   9975
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H8000000D&
      Height          =   1215
      Left            =   10560
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H8000000D&
      Height          =   1215
      Left            =   8280
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H8000000D&
      Height          =   1215
      Left            =   5760
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000D&
      Height          =   1215
      Left            =   3480
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000D&
      Height          =   1215
      Left            =   8280
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   1215
      Left            =   5760
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   1215
      Left            =   3480
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblPerfomance 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Image imgPerfomance 
      Height          =   1200
      Left            =   8280
      Picture         =   "frmReport.frx":1C0C3
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label lblAttach 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Attachment"
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
      Left            =   10560
      TabIndex        =   5
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Image imgAttach 
      Height          =   1200
      Left            =   10560
      Picture         =   "frmReport.frx":1CE8B
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label lblPersonal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personal"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblFamily 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Family"
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
      Left            =   5760
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblEducational 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Educational"
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
      Left            =   8280
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblMark 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Marks"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Image imgPersonal 
      Height          =   1200
      Left            =   3480
      Picture         =   "frmReport.frx":1DA7F
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Image imgFamily 
      Height          =   1200
      Left            =   5760
      Picture         =   "frmReport.frx":1E273
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Image imgEducational 
      Height          =   1200
      Left            =   8280
      Picture         =   "frmReport.frx":1FB98
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Image imgMark 
      Height          =   1200
      Left            =   3480
      Picture         =   "frmReport.frx":211D5
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Image imgPhysical 
      Height          =   1200
      Left            =   5760
      Picture         =   "frmReport.frx":21EC3
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Physical"
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
      Left            =   5760
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_GotFocus()
  Unload frmLoading
End Sub

Private Sub Form_Load()
  Theme frmReport
End Sub

Private Sub imgPersonal_Click()
   frmPersonalInfo.Show
End Sub

Private Sub lblPersonal_Click()
   frmPersonalInfo.Show
End Sub

Private Sub imgEducational_Click()
   frmEducationalInfo.Show
End Sub

Private Sub lblEducational_Click()
  frmEducationalInfo.Show
End Sub

Private Sub imgFamily_Click()
  frmFamilyInfo.Show
End Sub

Private Sub lblFamily_Click()
  frmFamilyInfo.Show
End Sub

Private Sub imgMark_Click()
  frmMarkPrint.Show
End Sub

Private Sub lblMark_Click()
  frmMarkPrint.Show
End Sub

Private Sub imgPhysical_Click()
  frmPhysicalInfo.Show
End Sub

Private Sub Label1_Click()
  frmPhysicalInfo.Show
End Sub

Private Sub imgPerfomance_Click()
   frmPerfomance.Show
End Sub

Private Sub lblPerfomance_Click()
   frmPerfomance.Show
End Sub

Private Sub imgAttach_Click()
  frmAttachment.imgOpen.Enabled = False
  frmAttachment.imgSave.Enabled = False
  frmAttachment.Show
End Sub

Private Sub lblAttach_Click()
  frmAttachment.imgOpen.Enabled = False
  frmAttachment.imgSave.Enabled = False
  frmAttachment.Show
End Sub

