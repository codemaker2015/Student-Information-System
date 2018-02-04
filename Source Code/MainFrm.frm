VERSION 5.00
Begin VB.Form mainfrm 
   AutoRedraw      =   -1  'True
   Caption         =   "Student Information System"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   8640
      TabIndex        =   25
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Next"
      Height          =   495
      Left            =   7080
      TabIndex        =   24
      Top             =   9240
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2640
      TabIndex        =   18
      Text            =   "Text5"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.ComboBox Combo6 
      BackColor       =   &H8000000F&
      Height          =   315
      ItemData        =   "MainFrm.frx":0000
      Left            =   2640
      List            =   "MainFrm.frx":0016
      TabIndex        =   17
      Text            =   "--Select--"
      Top             =   7200
      Width           =   975
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H8000000F&
      Height          =   315
      ItemData        =   "MainFrm.frx":002F
      Left            =   7800
      List            =   "MainFrm.frx":0039
      TabIndex        =   16
      Text            =   "--Select--"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H8000000F&
      Height          =   315
      ItemData        =   "MainFrm.frx":004C
      Left            =   2640
      List            =   "MainFrm.frx":0053
      TabIndex        =   15
      Text            =   "--Select--"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2640
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "MainFrm.frx":005E
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7800
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   315
      ItemData        =   "MainFrm.frx":0064
      Left            =   2640
      List            =   "MainFrm.frx":007A
      TabIndex        =   10
      Text            =   "--Select--"
      Top             =   3000
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   315
      ItemData        =   "MainFrm.frx":009C
      Left            =   11160
      List            =   "MainFrm.frx":0100
      TabIndex        =   9
      Text            =   "--Select--"
      Top             =   2160
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   315
      ItemData        =   "MainFrm.frx":01C4
      Left            =   7800
      List            =   "MainFrm.frx":01CE
      TabIndex        =   8
      Text            =   "--Select--"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      Caption         =   "Anual Family Income"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      TabIndex        =   23
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      Caption         =   "Blood Group"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      TabIndex        =   22
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      Caption         =   "Caste *"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6120
      TabIndex        =   21
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      Caption         =   "Religion *"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      TabIndex        =   20
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      Caption         =   "Phone **"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      TabIndex        =   19
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      Caption         =   "Permanent Address  *"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      Caption         =   "Reg No **"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Programme *"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      Caption         =   "Year of Study *"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9600
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Sex"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Name *"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Baselios Mount, Mulakulam North P.O, Piravom - 686664"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   960
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "BASELIOS POULOSE II CATHOLICOS COLLEGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub
