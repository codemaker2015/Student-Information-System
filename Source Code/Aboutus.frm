VERSION 5.00
Begin VB.Form aboutfrm 
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "ABOUT"
   ClientHeight    =   4725
   ClientLeft      =   4035
   ClientTop       =   3480
   ClientWidth     =   9495
   FillColor       =   &H80000002&
   ForeColor       =   &H80000006&
   Icon            =   "Aboutus.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Aboutus.frx":08CA
   ScaleHeight     =   4725
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      MouseIcon       =   "Aboutus.frx":170DA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To contact me,please Email nass_ibra@hotmail.com or make a call +919052006833"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   4320
      TabIndex        =   4
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Aboutus.frx":1722C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1215
      Left            =   4080
      TabIndex        =   3
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HYDERABAD"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TANZANIA STUDENT ASSOCIATION"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   2985
      Left            =   240
      Picture         =   "Aboutus.frx":17302
      Top             =   480
      Width           =   3000
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   3600
      X2              =   3600
      Y1              =   360
      Y2              =   4440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABOUT PROGRAMMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   120
      Top             =   360
      Width           =   9135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   360
      Top             =   480
      Width           =   9015
   End
End
Attribute VB_Name = "aboutfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
