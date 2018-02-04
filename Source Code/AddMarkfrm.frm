VERSION 5.00
Begin VB.Form frmAddMark 
   BackColor       =   &H80000005&
   Caption         =   "Add Mark Details"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15765
   FillColor       =   &H00FFFFFF&
   Icon            =   "AddMarkfrm.frx":0000
   LinkTopic       =   "MDIForm1"
   MDIChild        =   -1  'True
   PaletteMode     =   2  'Custom
   Picture         =   "AddMarkfrm.frx":000C
   ScaleHeight     =   11010
   ScaleWidth      =   15765
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      Caption         =   "Mark System"
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
      Height          =   255
      Left            =   2040
      TabIndex        =   59
      Top             =   1200
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00404040&
      Caption         =   "Grade System"
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
      Height          =   255
      Left            =   4320
      TabIndex        =   58
      Top             =   1200
      Width           =   1695
   End
   Begin VB.ListBox lstCourse 
      Height          =   255
      ItemData        =   "AddMarkfrm.frx":1C0C3
      Left            =   6000
      List            =   "AddMarkfrm.frx":1C0DC
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "AddMarkfrm.frx":1C111
      Left            =   3000
      List            =   "AddMarkfrm.frx":1C121
      TabIndex        =   35
      Top             =   7920
      Width           =   1695
   End
   Begin VB.TextBox txtInternalMark 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   34
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtInternalMark 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   33
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txtInternalMark 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   32
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtInternalMark 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   31
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtInternalMark 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   30
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtInternalMark 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   29
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ComboBox cmbSeminar 
      Height          =   315
      Index           =   5
      ItemData        =   "AddMarkfrm.frx":1C147
      Left            =   9600
      List            =   "AddMarkfrm.frx":1C15A
      TabIndex        =   28
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbAttendance 
      Height          =   315
      Index           =   5
      ItemData        =   "AddMarkfrm.frx":1C16D
      Left            =   8280
      List            =   "AddMarkfrm.frx":1C1C2
      TabIndex        =   27
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txttest2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   26
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txttest1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   6120
      MaxLength       =   3
      TabIndex        =   25
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbSeminar 
      Height          =   315
      Index           =   4
      ItemData        =   "AddMarkfrm.frx":1C239
      Left            =   9600
      List            =   "AddMarkfrm.frx":1C24C
      TabIndex        =   24
      Top             =   6840
      Width           =   1095
   End
   Begin VB.ComboBox cmbAttendance 
      Height          =   315
      Index           =   4
      ItemData        =   "AddMarkfrm.frx":1C25F
      Left            =   8280
      List            =   "AddMarkfrm.frx":1C2B4
      TabIndex        =   3
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txttest2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   23
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox txttest1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   6120
      MaxLength       =   3
      TabIndex        =   22
      Top             =   6840
      Width           =   855
   End
   Begin VB.ComboBox cmbSeminar 
      Height          =   315
      Index           =   3
      ItemData        =   "AddMarkfrm.frx":1C32B
      Left            =   9600
      List            =   "AddMarkfrm.frx":1C33E
      TabIndex        =   21
      Top             =   6360
      Width           =   1095
   End
   Begin VB.ComboBox cmbAttendance 
      Height          =   315
      Index           =   3
      ItemData        =   "AddMarkfrm.frx":1C351
      Left            =   8280
      List            =   "AddMarkfrm.frx":1C3A6
      TabIndex        =   20
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox txttest2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   19
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox txttest1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   6120
      MaxLength       =   3
      TabIndex        =   18
      Top             =   6360
      Width           =   855
   End
   Begin VB.ComboBox cmbSeminar 
      Height          =   315
      Index           =   2
      ItemData        =   "AddMarkfrm.frx":1C41D
      Left            =   9600
      List            =   "AddMarkfrm.frx":1C430
      TabIndex        =   17
      Top             =   5880
      Width           =   1095
   End
   Begin VB.ComboBox cmbAttendance 
      Height          =   315
      Index           =   2
      ItemData        =   "AddMarkfrm.frx":1C443
      Left            =   8280
      List            =   "AddMarkfrm.frx":1C498
      TabIndex        =   16
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txttest2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   15
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox txttest1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   6120
      MaxLength       =   3
      TabIndex        =   14
      Top             =   5880
      Width           =   855
   End
   Begin VB.ComboBox cmbSeminar 
      Height          =   315
      Index           =   1
      ItemData        =   "AddMarkfrm.frx":1C50F
      Left            =   9600
      List            =   "AddMarkfrm.frx":1C522
      TabIndex        =   13
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ComboBox cmbAttendance 
      Height          =   315
      Index           =   1
      ItemData        =   "AddMarkfrm.frx":1C535
      Left            =   8280
      List            =   "AddMarkfrm.frx":1C58A
      TabIndex        =   12
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txttest2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   11
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txttest1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   6120
      MaxLength       =   3
      TabIndex        =   10
      Top             =   5400
      Width           =   855
   End
   Begin VB.ComboBox cmbSeminar 
      Height          =   315
      Index           =   0
      ItemData        =   "AddMarkfrm.frx":1C601
      Left            =   9600
      List            =   "AddMarkfrm.frx":1C614
      TabIndex        =   9
      Top             =   4920
      Width           =   1095
   End
   Begin VB.ComboBox cmbAttendance 
      Height          =   315
      Index           =   0
      ItemData        =   "AddMarkfrm.frx":1C627
      Left            =   8280
      List            =   "AddMarkfrm.frx":1C67C
      TabIndex        =   8
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txttest2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   7
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txttest1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   6120
      MaxLength       =   3
      TabIndex        =   6
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtRegNo 
      Height          =   375
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ListBox lstSem 
      Height          =   255
      ItemData        =   "AddMarkfrm.frx":1C6F3
      Left            =   9480
      List            =   "AddMarkfrm.frx":1C709
      TabIndex        =   5
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "5"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3360
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "25"
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblPhysical 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Extra Curricular activities"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   61
      Top             =   8400
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Marking System"
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
      Height          =   255
      Left            =   1800
      TabIndex        =   60
      Top             =   600
      Width           =   1575
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H8000000D&
      Height          =   735
      Left            =   1560
      Top             =   960
      Width           =   4695
   End
   Begin VB.Shape Shape6 
      Height          =   495
      Left            =   10800
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      Height          =   495
      Left            =   9000
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      Height          =   495
      Left            =   3960
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   1920
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Line Line3 
      Index           =   5
      X1              =   1560
      X2              =   12360
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label lblCourse 
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
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5040
      TabIndex        =   57
      Top             =   3480
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   1215
      Left            =   1560
      Top             =   8880
      Width           =   10815
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
      Left            =   1920
      TabIndex        =   44
      Top             =   9360
      Width           =   1005
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
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
      Left            =   3960
      TabIndex        =   45
      Top             =   9360
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
      Left            =   10800
      TabIndex        =   43
      Top             =   9360
      Width           =   1005
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
      Left            =   9000
      TabIndex        =   42
      Top             =   9360
      Width           =   1005
   End
   Begin VB.Label lblPerfo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Height          =   240
      Left            =   1560
      TabIndex        =   41
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Mark"
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
      Height          =   240
      Left            =   1560
      TabIndex        =   56
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Mark"
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
      Height          =   240
      Left            =   1560
      TabIndex        =   55
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   240
      Index           =   5
      Left            =   1800
      TabIndex        =   54
      Top             =   7320
      Width           =   45
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 4"
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
      Height          =   240
      Index           =   4
      Left            =   1800
      TabIndex        =   40
      Top             =   6840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 3"
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
      Height          =   240
      Index           =   3
      Left            =   1800
      TabIndex        =   39
      Top             =   6360
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 2"
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
      Height          =   240
      Index           =   2
      Left            =   1800
      TabIndex        =   38
      Top             =   5880
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 1"
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
      Height          =   240
      Index           =   1
      Left            =   1800
      TabIndex        =   37
      Top             =   5400
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 0"
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
      Height          =   240
      Index           =   0
      Left            =   1800
      TabIndex        =   36
      Top             =   4920
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      Index           =   4
      X1              =   1560
      X2              =   12360
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      Index           =   3
      X1              =   1560
      X2              =   12360
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      Index           =   2
      X1              =   1560
      X2              =   12360
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   1560
      X2              =   12360
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   1560
      X2              =   12360
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Mark"
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
      Height          =   240
      Left            =   11160
      TabIndex        =   53
      Top             =   4440
      Width           =   1140
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   4
      X1              =   10920
      X2              =   10920
      Y1              =   4320
      Y2              =   7680
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seminar / Viva"
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
      Height          =   240
      Left            =   9480
      TabIndex        =   52
      Top             =   4440
      Width           =   1305
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance"
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
      Height          =   240
      Left            =   8280
      TabIndex        =   51
      Top             =   4440
      Width           =   1020
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test 2"
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
      Height          =   240
      Left            =   7320
      TabIndex        =   50
      Top             =   4440
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test 1"
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
      Height          =   240
      Left            =   6240
      TabIndex        =   49
      Top             =   4440
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subjects"
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
      Height          =   240
      Left            =   2400
      TabIndex        =   48
      Top             =   4440
      Width           =   780
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   3
      X1              =   9360
      X2              =   9360
      Y1              =   4320
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   2
      X1              =   8160
      X2              =   8160
      Y1              =   4320
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   7080
      X2              =   7080
      Y1              =   4320
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   6000
      X2              =   6000
      Y1              =   4320
      Y2              =   7680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   1560
      X2              =   12360
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   3375
      Left            =   1560
      Top             =   4320
      Width           =   10815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Semester"
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
      Height          =   240
      Left            =   8400
      TabIndex        =   47
      Top             =   3480
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reg. No."
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
      Height          =   240
      Left            =   1560
      TabIndex        =   46
      Top             =   3480
      Width           =   795
   End
End
Attribute VB_Name = "frmAddMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Label13_Click()
  Unload Me
End Sub

Private Sub Form_GotFocus()
  Unload frmLoading
End Sub

Private Sub Form_Load()
  Theme frmAddMark
  test1 = test2 = attendance = seminar = internal = total = 0
  Option1_Click
  lstCourse.Text = lstCourse.List(0)
  lstSem.Text = lstSem.List(0)
  connection
  reccheck
  rec.Open ("select SUBJECTNAME from SUBJECTTABLE where course = '" & Trim(lstCourse.Text) & "' and sem = '" & Trim(lstSem.Text) & "'"), con, adOpenDynamic, adLockOptimistic

  Dim i As Integer
  'i = 0
  While Not rec.EOF
    lblSubject(i).Visible = True
    txttest1(i).Visible = True
    txttest2(i).Visible = True
    cmbAttendance(i).Visible = True
    cmbSeminar(i).Visible = True
    txtInternalMark(i).Visible = True
    
    lblSubject(i).Caption = rec.Fields(0)
    rec.MoveNext
    i = i + 1
  Wend
End Sub


Private Sub lblAdd_Click()
  'On Error GoTo error_para
  
  Dim i As Integer, J As Integer
  J = 0
  On Error Resume Next
  If txttest1(0).Text = "" Or txttest1(1).Text = "" Or txttest1(2).Text = "" Or txttest1(3).Text = "" Or txttest1(4).Text = "" Then
     Dialog.Label1(J).Visible = True
     Dialog.Label2.Caption = 0
     Dialog.Label1(J).Caption = "Test 1 Mark left blank"
     J = J + 1
  End If
  If txttest2(0).Text = "" Or txttest2(1).Text = "" Or txttest2(2).Text = "" Or txttest2(3).Text = "" Or txttest2(4).Text = "" Then
     Dialog.Label1(J).Visible = True
     Dialog.Label2.Caption = 0
     Dialog.Label1(J).Caption = "Test 2 Mark left blank"
     J = J + 1
  End If
  If cmbAttendance(0).Text = "" Or cmbAttendance(1).Text = "" Or cmbAttendance(2).Text = "" Or cmbAttendance(3).Text = "" Or cmbAttendance(4).Text = "" Then
     Dialog.Label1(J).Visible = True
     Dialog.Label2.Caption = 0
     Dialog.Label1(J).Caption = "Attendance Mark left blank"
     J = J + 1
  End If
  If cmbSeminar(0).Text = "" Or cmbSeminar(1).Text = "" Or cmbSeminar(2).Text = "" Or cmbSeminar(3).Text = "" Or cmbSeminar(4).Text = "" Then
     Dialog.Label1(J).Visible = True
     Dialog.Label2.Caption = 0
     Dialog.Label1(J).Caption = "Seminar Mark left blank"
     J = J + 1
  End If
  If J <> 0 Then
     Dialog.Show
  Else
      connection
      reccheck
           
      If MsgBox("Are You sure, The entered details are correct?", vbQuestion + vbYesNo, "Conformation") = vbYes Then
        rec.Open "select REGNO from INTERNALMARKTABLE where REGNO ='" & Trim(txtRegNo.Text) & "'", con, adOpenDynamic, adLockPessimistic
        If rec.EOF = False Then
            If MsgBox("Student information is already in database, Do you wish to Replace it with this one?", vbYesNo, "Data entry error") = vbNo Then
                Exit Sub
            Else
              reccheck
              rec.Open "delete from INTERNALMARKTABLE where REGNO = " & Trim(txtRegNo.Text), con, adOpenDynamic, adLockPessimistic
            End If
       End If
    End If
    
    For i = 0 To 5
      If lblSubject(i) = "Subject " + CStr(i) Then Exit For
    
      connection
      reccheck
      
      
      Dim Temp1 As Integer, Temp2 As Integer, Temp3 As Integer
      If marksystem = "Grade" Then
      Select Case Trim(txttest1(i).Text)
         Case "A": Temp1 = 5
         Case "B": Temp1 = 4
         Case "C": Temp1 = 3
         Case "D": Temp1 = 2
         Case "E": Temp1 = 1
      End Select
      End If
      
      If marksystem = "Grade" Then
      Select Case Trim(txttest2(i).Text)
         Case "A": Temp2 = 5
         Case "B": Temp2 = 4
         Case "C": Temp2 = 3
         Case "D": Temp2 = 2
         Case "E": Temp2 = 1
      End Select
      End If
      
      If marksystem = "Grade" Then
      Select Case Trim(txtInternalMark(i).Text)
         Case "A": Temp3 = 5
         Case "B": Temp3 = 4
         Case "C": Temp3 = 3
         Case "D": Temp3 = 2
         Case "E": Temp3 = 1
      End Select
      End If
        
      connection
      reccheck
       
       rec.Open ("insert into INTERNALMARKTABLE values('" & Val(Trim(txtRegNo.Text)) - 1000000 & "','" & Val(Trim(txtRegNo.Text)) & "','" & Val(Trim(lstSem.Text)) & "','" & Trim(lstCourse.Text) & "','" & _
       Trim(lblSubject(i).Caption) & "','" & Temp1 & "','" & Temp2 & "','" & Val(Trim(cmbAttendance(i).Text)) & "','" & Val(Trim(cmbSeminar(i).Text)) & "','" & Temp3 & "')"), con, adOpenDynamic, adLockOptimistic
    Next i
    MsgBox "Mark addes successfully", vbInformation, "Information"
  End If
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub lblDelete_Click()
  'On Error GoTo error_para

  If CheckRegNo(txtRegNo, 8) = True Then
    connection
    reccheck
    rec.Open ("select STUDENTNAME from MAINTABLE where REGNO = '" & Val(Trim(txtRegNo.Text)) & "'"), con, adOpenDynamic, adLockOptimistic
    If MsgBox("Do you wish to delete Mark details of " + rec.Fields(0), vbYesNo, "Conformation") = vbYes Then
       frmDelete.Show
       frmDelete.txtDelete.Text = txtRegNo.Text
       Unload Me
    End If
  End If
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub lblEdit_Click()
  frmAddStudentInfo.Show
End Sub

Private Sub lblExit_Click()
  Unload Me
End Sub

Private Sub lblPhysical_Click()
   frmPhysical.Show
End Sub

Private Sub lstCourse_LostFocus()
  'On Error GoTo error_para
  
  connection
  reccheck
  rec.Open ("select SUBJECTNAME from SUBJECTTABLE where course = '" & Trim(lstCourse.Text) & "' and sem = '" & Trim(lstSem.Text) & "'"), con, adOpenDynamic, adLockOptimistic

  Dim i As Integer
  While Not rec.EOF
    lblSubject(i).Visible = True
    txttest1(i).Visible = True
    txttest2(i).Visible = True
    cmbAttendance(i).Visible = True
    cmbSeminar(i).Visible = True
    txtInternalMark(i).Visible = True
  
    lblSubject(i).Caption = rec.Fields(0)
    rec.MoveNext
    i = i + 1
  Wend
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub lstSem_GotFocus()
  'On Error GoTo error_para

  connection
  reccheck
  rec.Open ("select SUBJECTNAME from SUBJECTTABLE where course = '" & Trim(lstCourse.Text) & "' and sem = '" & Trim(lstSem.Text) & "'"), con, adOpenDynamic, adLockOptimistic

  Dim i As Integer
  While Not rec.EOF
    lblSubject(i).Visible = True
    txttest1(i).Visible = True
    txttest2(i).Visible = True
    cmbAttendance(i).Visible = True
    cmbSeminar(i).Visible = True
    txtInternalMark(i).Visible = True
  
    lblSubject(i).Caption = rec.Fields(0)
    rec.MoveNext
    i = i + 1
  Wend
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub lstSem_LostFocus()
  'On Error GoTo error_para

  connection
  reccheck
  rec.Open ("select SUBJECTNAME from SUBJECTTABLE where course = '" & Trim(lstCourse.Text) & "' and sem = '" & Trim(lstSem.Text) & "'"), con, adOpenDynamic, adLockOptimistic

  Dim i As Integer
  While Not rec.EOF
    lblSubject(i).Visible = True
    txttest1(i).Visible = True
    txttest2(i).Visible = True
    cmbAttendance(i).Visible = True
    cmbSeminar(i).Visible = True
    txtInternalMark(i).Visible = True
  
    lblSubject(i).Caption = rec.Fields(0)
    rec.MoveNext
    i = i + 1
  Wend
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub
Private Sub Option1_Click()
  Label10.Caption = "Maximum Mark"
  Label11.Caption = "Additional Mark"
  Text1.Text = "20"
  Text2.Text = "5"
  marksystem = "Mark"
End Sub

Private Sub Option2_Click()
  Label10.Caption = "Maximum Grade"
  Label11.Caption = "Additional Grade"
  Text1.Text = "A"
  Text2.Text = "C"
  marksystem = "Grade"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If marksystem = "Grade" Then
     If KeyAscii < 65 Or KeyAscii > 69 Then
      If KeyAscii <> 8 Then
         KeyAscii = 0
      End If
     End If
   End If
   If marksystem = "Mark" Then ValRegNo KeyAscii
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If marksystem = "Grade" Then
     If KeyAscii < 65 Or KeyAscii > 69 Then
      If KeyAscii <> 8 Then
         KeyAscii = 0
      End If
     End If
   End If
   If marksystem = "Mark" Then ValRegNo KeyAscii
End Sub

Private Sub txtRegNo_KeyPress(KeyAscii As Integer)
   ValRegNo KeyAscii
End Sub

Private Sub txtRegNo_LostFocus()
   CheckRegNo txtRegNo, 8
End Sub

Private Sub txttest1_Change(Index As Integer)
   Dim percent As Double
   If marksystem = "Grade" Then
      Select Case Trim(txttest1(Index).Text)
         Case "A": test1 = 5
         Case "B": test1 = 4
         Case "C": test1 = 3
         Case "D": test1 = 2
         Case "E": test1 = 1
      End Select
   End If
   If marksystem = "Mark" Then
      percent = Val(Trim(txttest1(Index).Text)) * 100 / Val(Trim(Text1.Text))
      If percent > 90 Then test1 = 5
      If percent > 85 And percent < 89 Then test1 = 4
      If percent > 80 And percent < 84 Then test1 = 3
      If percent > 76 And percent < 79 Then test1 = 2
      If percent < 75 Then test1 = 1
   End If
End Sub

Private Sub txttest1_KeyPress(Index As Integer, KeyAscii As Integer)
   If marksystem = "Grade" Then
     If KeyAscii < 65 Or KeyAscii > 69 Then
      If KeyAscii <> 8 Then
         KeyAscii = 0
      End If
     End If
   End If
   If marksystem = "Mark" Then ValRegNo KeyAscii
End Sub

Private Sub txttest2_Change(Index As Integer)
  Dim percent As Double
  If marksystem = "Grade" Then
      Select Case Trim(txttest2(Index).Text)
         Case "A": test2 = 5
         Case "B": test2 = 4
         Case "C": test2 = 3
         Case "D": test2 = 2
         Case "E": test2 = 1
      End Select
   End If
   If marksystem = "Mark" Then
      percent = Val(Trim(txttest2(Index).Text)) * 100 / Val(Trim(Text1.Text))
      If percent > 90 Then test2 = 5
      If percent > 85 And percent < 89 Then test2 = 4
      If percent > 80 And percent < 84 Then test2 = 3
      If percent > 76 And percent < 79 Then test2 = 2
      If percent < 75 Then test2 = 1
   End If
End Sub

Private Sub cmbAttendance_Change(Index As Integer)
    'If marksystem = "Grade" Then
    If Val(Trim(cmbAttendance(Index).Text)) > 90 Then attendance = 5
    If Val(Trim(cmbAttendance(Index).Text)) > 85 And Val(Trim(cmbAttendance(Index).Text)) < 89 Then attendance = 4
    If Val(Trim(cmbAttendance(Index).Text)) > 80 And Val(Trim(cmbAttendance(Index).Text)) < 84 Then attendance = 3
    If Val(Trim(cmbAttendance(Index).Text)) > 76 And Val(Trim(cmbAttendance(Index).Text)) < 79 Then attendance = 2
    If Val(Trim(cmbAttendance(Index).Text)) < 75 Then attendance = 1
          
     'End If
End Sub

Private Sub cmbSeminar_Change(Index As Integer)
   seminar = Val(Trim(cmbSeminar(Index).Text))
End Sub

Private Sub cmbSeminar_LostFocus(Index As Integer)
   
   If marksystem = "Grade" Then
      Select Case Trim(Text2.Text)
         Case "A": internal = (test1 + test2) / 2 + attendance + seminar + 5
         Case "B": internal = (test1 + test2) / 2 + attendance + seminar + 4
         Case "C": internal = (test1 + test2) / 2 + attendance + seminar + 3
         Case "D": internal = (test1 + test2) / 2 + attendance + seminar + 2
      End Select
      
      If internal > 15 Then txtInternalMark(Index).Text = "A"
      If internal > 11 And internal < 16 Then txtInternalMark(Index).Text = "B"
      If internal > 7 And internal < 12 Then txtInternalMark(Index).Text = "C"
      If internal > 3 And internal < 8 Then txtInternalMark(Index).Text = "D"
      If internal < 4 Then txtInternalMark(Index).Text = "E"
   End If
   If marksystem = "Mark" Then
      txtInternalMark(Index).Text = internal
      internal = (test1 + test2) / 2 + attendance + seminar + Val(Trim(Text2.Text))
   End If
   txtInternalMark(Index).Refresh
   
   total = total + internal
   If total > 70 Then List1.Text = List1.List(0)
   If total > 60 And total < 69 Then List1.Text = List1.List(1)
   If total > 40 And total < 59 Then List1.Text = List1.List(2)
   If total < 40 Then List1.Text = List1.List(3)
End Sub

Private Sub txttest2_KeyPress(Index As Integer, KeyAscii As Integer)
  If marksystem = "Grade" Then
     If KeyAscii < 65 Or KeyAscii > 69 Then
      If KeyAscii <> 8 Then
         KeyAscii = 0
      End If
     End If
   End If
   If marksystem = "Mark" Then ValRegNo KeyAscii
End Sub
