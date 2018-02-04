VERSION 5.00
Begin VB.Form frmExit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
   Icon            =   "frmExit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmExit.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   1110
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "           No"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "          Yes"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure you want to quit ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label2_Click()
  Dim i As Integer
  For i = Forms.Count - 1 To 0 Step -1
     Unload Forms(i)
  Next i
End Sub

Private Sub Label3_Click()
  Unload Me
End Sub
