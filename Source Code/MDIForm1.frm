VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000F&
   Caption         =   "Student Information System 2015"
   ClientHeight    =   6300
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00404000&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   6300
      Left            =   0
      Picture         =   "MDIForm1.frx":EBC6E
      ScaleHeight     =   6240
      ScaleWidth      =   2580
      TabIndex        =   0
      Top             =   0
      Width           =   2640
      Begin VB.Image Image3 
         Height          =   1215
         Left            =   480
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   855
         Left            =   600
         TabIndex        =   2
         Top             =   3840
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   975
         Left            =   360
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   615
         Left            =   480
         TabIndex        =   1
         Top             =   1440
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   360
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Menu FileMnu 
      Caption         =   "File"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Label1_Click()
  AddStudentInfofrm.Show
End Sub

Private Sub Label2_Click()
  AddMarkfrm.Show
End Sub

Private Sub Label3_Click()
  Searchfrm.Show
End Sub
