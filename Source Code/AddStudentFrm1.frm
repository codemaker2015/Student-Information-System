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
      TabPicture(0)   =   "AddStudentFrm1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Family Information"
      TabPicture(1)   =   "AddStudentFrm1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "More Information"
      TabPicture(2)   =   "AddStudentFrm1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Summary"
      TabPicture(3)   =   "AddStudentFrm1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.Frame Frame4 
         Caption         =   "Details of Brother(s) and Sister(s)"
         Height          =   2295
         Left            =   -70440
         TabIndex        =   26
         Top             =   2760
         Width           =   4095
         Begin VB.Shape Shape4 
            BorderColor     =   &H80000001&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   1815
            Left            =   120
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Guardian Info"
         Height          =   2175
         Left            =   -70440
         TabIndex        =   25
         Top             =   480
         Width           =   4095
         Begin VB.Shape Shape3 
            BorderColor     =   &H80000001&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   1815
            Left            =   120
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Parents Info"
         Height          =   4575
         Left            =   -74760
         TabIndex        =   24
         Top             =   480
         Width           =   4095
         Begin VB.Label Label20 
            Caption         =   "Label20"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   4200
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "Label19"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   3840
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Label18"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   3480
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Label16"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Label15"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "Label14"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "Label13"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Label12"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H80000001&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   4215
            Left            =   120
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Student Info"
         Height          =   4815
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   8415
         Begin VB.PictureBox Picture1 
            Height          =   1455
            Left            =   5520
            ScaleHeight     =   1395
            ScaleWidth      =   2355
            TabIndex        =   22
            Top             =   2160
            Width           =   2415
            Begin VB.Image Image1 
               Height          =   1215
               Left            =   120
               Top             =   120
               Width           =   2175
            End
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   5520
            TabIndex        =   21
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox Text3 
            Height          =   975
            Left            =   5520
            TabIndex        =   20
            Top             =   600
            Width           =   2415
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   1680
            TabIndex        =   19
            Text            =   "--Select--"
            Top             =   3120
            Width           =   2415
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   1680
            TabIndex        =   18
            Text            =   "--Select--"
            Top             =   2760
            Width           =   2415
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   1680
            TabIndex        =   17
            Text            =   "--Select--"
            Top             =   2400
            Width           =   2415
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   1680
            TabIndex        =   16
            Text            =   "--Select--"
            Top             =   2040
            Width           =   2415
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1680
            TabIndex        =   15
            Text            =   "--Select--"
            Top             =   1680
            Width           =   2415
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1680
            TabIndex        =   14
            Text            =   "--Select--"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1680
            TabIndex        =   13
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1680
            TabIndex        =   12
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Browse Picture"
            Height          =   255
            Left            =   5520
            TabIndex        =   23
            Top             =   3600
            Width           =   2415
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            Height          =   195
            Left            =   4560
            TabIndex        =   11
            Top             =   1680
            Width           =   465
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   4560
            TabIndex        =   10
            Top             =   600
            Width           =   570
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Anual income"
            Height          =   195
            Left            =   360
            TabIndex        =   9
            Top             =   3120
            Width           =   960
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Blood group"
            Height          =   195
            Left            =   360
            TabIndex        =   8
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caste"
            Height          =   195
            Left            =   360
            TabIndex        =   7
            Top             =   2400
            Width           =   405
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Religion"
            Height          =   195
            Left            =   360
            TabIndex        =   6
            Top             =   2040
            Width           =   570
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            Height          =   195
            Left            =   360
            TabIndex        =   5
            Top             =   1680
            Width           =   525
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year of study"
            Height          =   195
            Left            =   360
            TabIndex        =   4
            Top             =   1320
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   360
            TabIndex        =   3
            Top             =   960
            Width           =   420
         End
         Begin VB.Label Re 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg No"
            Height          =   195
            Left            =   360
            TabIndex        =   2
            Top             =   600
            Width           =   555
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
