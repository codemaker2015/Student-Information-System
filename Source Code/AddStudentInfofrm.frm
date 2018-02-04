VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form AddStudentInfofrm 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Educational Information"
      TabPicture(0)   =   "AddStudentInfofrm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Family Information"
      TabPicture(1)   =   "AddStudentInfofrm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "More Information"
      TabPicture(2)   =   "AddStudentInfofrm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Summary"
      TabPicture(3)   =   "AddStudentInfofrm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.Frame Frame1 
         Caption         =   "Student Info"
         Height          =   4815
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   8415
         Begin VB.PictureBox Picture1 
            Height          =   1815
            Left            =   5400
            ScaleHeight     =   1755
            ScaleWidth      =   2355
            TabIndex        =   22
            Top             =   2160
            Width           =   2415
            Begin VB.Image Image1 
               Height          =   1455
               Left            =   120
               Top             =   120
               Width           =   2055
            End
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   5400
            TabIndex        =   21
            Text            =   "Text5"
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox Text4 
            Height          =   975
            Left            =   5400
            TabIndex        =   20
            Text            =   "Text4"
            Top             =   600
            Width           =   2415
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   1320
            TabIndex        =   19
            Text            =   "Combo5"
            Top             =   3120
            Width           =   1815
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   1320
            TabIndex        =   18
            Text            =   "Combo4"
            Top             =   2760
            Width           =   1815
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   1320
            TabIndex        =   17
            Text            =   "Combo3"
            Top             =   2400
            Width           =   1815
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1320
            TabIndex        =   16
            Text            =   "Combo2"
            Top             =   2040
            Width           =   1815
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1320
            TabIndex        =   15
            Text            =   "Combo1"
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1320
            TabIndex        =   14
            Text            =   "Text3"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1320
            TabIndex        =   13
            Text            =   "Text2"
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1320
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Browse picture"
            Height          =   375
            Left            =   5400
            TabIndex        =   23
            Top             =   3960
            Width           =   2415
         End
         Begin VB.Label Label10 
            Caption         =   "Address"
            Height          =   255
            Left            =   4320
            TabIndex        =   11
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Phone"
            Height          =   255
            Left            =   4320
            TabIndex        =   10
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Anual income"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Blood group"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Caste"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Religion"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Gender"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Year of study"
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Name"
            Height          =   255
            Left            =   360
            TabIndex        =   3
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Re 
            Caption         =   "Reg No"
            Height          =   255
            Left            =   360
            TabIndex        =   2
            Top             =   600
            Width           =   855
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   4335
            Left            =   240
            Top             =   360
            Width           =   7935
         End
      End
   End
End
Attribute VB_Name = "AddStudentInfofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label2_Click()

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  Select Case SSTab1.Caption
    Case "Family Information": MsgBox "Family Information"
    Case "More Information": MsgBox "More Information"
  End Select
    
End Sub
