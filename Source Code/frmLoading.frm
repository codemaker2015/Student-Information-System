VERSION 5.00
Begin VB.Form frmLoading 
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19260
   Icon            =   "frmLoading.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmLoading.frx":000C
   ScaleHeight     =   10215
   ScaleWidth      =   19260
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3960
      Top             =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   6120
      TabIndex        =   0
      Top             =   4560
      Width           =   2655
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private i As Integer

Private Sub Form_Load()
  Theme frmLoading
  
  i = 0
  
End Sub

Private Sub Timer1_Timer()
  i = i + 1
  Dim k As Integer
  k = i Mod 3
  Select Case k
    Case 1: Label1.Caption = "Loading."
    Case 2: Label1.Caption = "Loading.."
    Case 0: Label1.Caption = "Loading..."
  End Select
  Label1.Refresh
  If i = 3 Then
    Select Case choice
       Case 1: frmAddStudentInfo.Show
       Case 2: frmAddMark.Show
       Case 3: frmSearch.Show
       Case 4: frmReport.Show
       Case 5: frmDelete.Show
       Case 6: frmAdminOptions.SSTab1.Tab = 2
               frmAdminOptions.Show
       Case 7: frmAdminOptions.SSTab1.Tab = 0
               frmAdminOptions.Show
    End Select
    Unload Me
  End If
End Sub
