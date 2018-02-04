VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAddStudentInfo 
   BackColor       =   &H80000005&
   Caption         =   "Add Student Information"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   Icon            =   "AddStudentFrm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "AddStudentFrm.frx":000C
   ScaleHeight     =   9600
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabAddinfo 
      Height          =   6615
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Personal Information"
      TabPicture(0)   =   "AddStudentFrm.frx":1C0C3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame7(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Family Information"
      TabPicture(1)   =   "AddStudentFrm.frx":1C0DF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7(2)"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Frame2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Examination Information"
      TabPicture(2)   =   "AddStudentFrm.frx":1C0FB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7(1)"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Summary"
      TabPicture(3)   =   "AddStudentFrm.frx":1C117
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "Frame7(0)"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame7 
         Height          =   680
         Index           =   0
         Left            =   -74880
         TabIndex        =   70
         Top             =   5280
         Width           =   9015
         Begin VB.Image imgSave 
            Height          =   420
            Left            =   8040
            Picture         =   "AddStudentFrm.frx":1C133
            Top             =   180
            Width           =   420
         End
         Begin VB.Image imgBack 
            Height          =   480
            Index           =   2
            Left            =   6720
            Picture         =   "AddStudentFrm.frx":1CB1D
            Top             =   120
            Width           =   480
         End
         Begin VB.Image imgCancel 
            Height          =   480
            Index           =   3
            Left            =   240
            Picture         =   "AddStudentFrm.frx":1CF5F
            Top             =   140
            Width           =   480
         End
      End
      Begin VB.Frame Frame7 
         Height          =   680
         Index           =   1
         Left            =   -74880
         TabIndex        =   69
         Top             =   5280
         Width           =   9015
         Begin VB.Image imgBack 
            Height          =   480
            Index           =   1
            Left            =   6720
            Picture         =   "AddStudentFrm.frx":1DC29
            Top             =   120
            Width           =   480
         End
         Begin VB.Image imgNext 
            Height          =   480
            Index           =   2
            Left            =   8040
            Picture         =   "AddStudentFrm.frx":1E06B
            Top             =   120
            Width           =   480
         End
         Begin VB.Image imgCancel 
            Height          =   480
            Index           =   2
            Left            =   240
            Picture         =   "AddStudentFrm.frx":1E4AD
            Top             =   140
            Width           =   480
         End
      End
      Begin VB.Frame Frame7 
         Height          =   680
         Index           =   2
         Left            =   -74880
         TabIndex        =   68
         Top             =   5280
         Width           =   9015
         Begin VB.Image imgCancel 
            Height          =   480
            Index           =   1
            Left            =   240
            Picture         =   "AddStudentFrm.frx":1F177
            Top             =   140
            Width           =   480
         End
         Begin VB.Image imgNext 
            Height          =   480
            Index           =   1
            Left            =   8040
            Picture         =   "AddStudentFrm.frx":1FE41
            Top             =   120
            Width           =   480
         End
         Begin VB.Image imgBack 
            Height          =   480
            Index           =   0
            Left            =   6720
            Picture         =   "AddStudentFrm.frx":20283
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.Frame Frame7 
         Height          =   680
         Index           =   3
         Left            =   120
         TabIndex        =   67
         Top             =   5280
         Width           =   9015
         Begin VB.Image imgCancel 
            Height          =   480
            Index           =   0
            Left            =   240
            Picture         =   "AddStudentFrm.frx":206C5
            ToolTipText     =   "Cancel"
            Top             =   140
            Width           =   480
         End
         Begin VB.Image imgNext 
            Height          =   480
            Index           =   0
            Left            =   8040
            Picture         =   "AddStudentFrm.frx":2138F
            ToolTipText     =   "Next"
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Summary"
         Height          =   4695
         Left            =   -74880
         TabIndex        =   59
         Top             =   480
         Width           =   9015
         Begin RichTextLib.RichTextBox rtbinfo 
            Height          =   4335
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   7646
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"AddStudentFrm.frx":217D1
         End
         Begin RichTextLib.RichTextBox rtbinfo 
            Height          =   4335
            Index           =   1
            Left            =   4440
            TabIndex        =   75
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   7646
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"AddStudentFrm.frx":21853
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Examination Info"
         Height          =   4695
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   9015
         Begin VB.ComboBox cmbGrade 
            Height          =   315
            Index           =   2
            ItemData        =   "AddStudentFrm.frx":218D5
            Left            =   6080
            List            =   "AddStudentFrm.frx":218E8
            TabIndex        =   60
            Top             =   2760
            Width           =   615
         End
         Begin VB.ComboBox cmbGrade 
            Height          =   315
            Index           =   1
            ItemData        =   "AddStudentFrm.frx":218FB
            Left            =   6080
            List            =   "AddStudentFrm.frx":2190E
            TabIndex        =   73
            Top             =   2280
            Width           =   615
         End
         Begin VB.ComboBox cmbGrade 
            Height          =   315
            Index           =   0
            ItemData        =   "AddStudentFrm.frx":21921
            Left            =   6080
            List            =   "AddStudentFrm.frx":21934
            TabIndex        =   72
            Top             =   1800
            Width           =   615
         End
         Begin VB.ListBox lstNoOfChances 
            Height          =   255
            Index           =   2
            ItemData        =   "AddStudentFrm.frx":21947
            Left            =   7680
            List            =   "AddStudentFrm.frx":21954
            TabIndex        =   58
            Top             =   2760
            Width           =   975
         End
         Begin VB.ComboBox cmbPercent 
            Height          =   315
            Index           =   2
            ItemData        =   "AddStudentFrm.frx":21961
            Left            =   6790
            List            =   "AddStudentFrm.frx":21A0A
            TabIndex        =   57
            Top             =   2760
            Width           =   615
         End
         Begin VB.ComboBox cmbSchoolName 
            Height          =   315
            Index           =   2
            ItemData        =   "AddStudentFrm.frx":21AEB
            Left            =   3720
            List            =   "AddStudentFrm.frx":21AFB
            TabIndex        =   56
            Text            =   "--Select--"
            Top             =   2760
            Width           =   2175
         End
         Begin VB.ComboBox cmbYearOfPassing 
            Height          =   315
            Index           =   2
            ItemData        =   "AddStudentFrm.frx":21B6E
            Left            =   2760
            List            =   "AddStudentFrm.frx":21BFC
            TabIndex        =   55
            Text            =   "Year"
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox txtRegNo 
            Height          =   285
            Index           =   3
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   54
            Top             =   2760
            Width           =   855
         End
         Begin VB.ListBox lstCourse 
            Height          =   255
            Index           =   2
            ItemData        =   "AddStudentFrm.frx":21D14
            Left            =   360
            List            =   "AddStudentFrm.frx":21D21
            TabIndex        =   53
            Top             =   2760
            Width           =   1095
         End
         Begin VB.ListBox lstNoOfChances 
            Height          =   255
            Index           =   1
            ItemData        =   "AddStudentFrm.frx":21D38
            Left            =   7680
            List            =   "AddStudentFrm.frx":21D45
            TabIndex        =   52
            Top             =   2280
            Width           =   975
         End
         Begin VB.ComboBox cmbPercent 
            Height          =   315
            Index           =   1
            ItemData        =   "AddStudentFrm.frx":21D52
            Left            =   6790
            List            =   "AddStudentFrm.frx":21DFB
            TabIndex        =   51
            Top             =   2280
            Width           =   615
         End
         Begin VB.ComboBox cmbSchoolName 
            Height          =   315
            Index           =   1
            ItemData        =   "AddStudentFrm.frx":21EDC
            Left            =   3720
            List            =   "AddStudentFrm.frx":21EEC
            TabIndex        =   50
            Text            =   "--Select--"
            Top             =   2280
            Width           =   2175
         End
         Begin VB.ComboBox cmbYearOfPassing 
            Height          =   315
            Index           =   1
            ItemData        =   "AddStudentFrm.frx":21F5F
            Left            =   2760
            List            =   "AddStudentFrm.frx":21FED
            TabIndex        =   49
            Text            =   "Year"
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtRegNo 
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   9
            TabIndex        =   48
            Top             =   2280
            Width           =   855
         End
         Begin VB.ListBox lstCourse 
            Height          =   255
            Index           =   1
            ItemData        =   "AddStudentFrm.frx":22105
            Left            =   360
            List            =   "AddStudentFrm.frx":22112
            TabIndex        =   47
            Top             =   2280
            Width           =   1095
         End
         Begin VB.ListBox lstNoOfChances 
            Height          =   255
            Index           =   0
            ItemData        =   "AddStudentFrm.frx":22129
            Left            =   7680
            List            =   "AddStudentFrm.frx":22136
            TabIndex        =   45
            Top             =   1800
            Width           =   975
         End
         Begin VB.ComboBox cmbPercent 
            Height          =   315
            Index           =   0
            ItemData        =   "AddStudentFrm.frx":22143
            Left            =   6800
            List            =   "AddStudentFrm.frx":221EC
            TabIndex        =   44
            Top             =   1800
            Width           =   615
         End
         Begin VB.ComboBox cmbSchoolName 
            Height          =   315
            Index           =   0
            ItemData        =   "AddStudentFrm.frx":222CD
            Left            =   3720
            List            =   "AddStudentFrm.frx":222E0
            TabIndex        =   43
            Text            =   "--Select--"
            Top             =   1800
            Width           =   2175
         End
         Begin VB.ComboBox cmbYearOfPassing 
            Height          =   315
            Index           =   0
            ItemData        =   "AddStudentFrm.frx":2235B
            Left            =   2760
            List            =   "AddStudentFrm.frx":223E9
            TabIndex        =   42
            Text            =   "Year"
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox txtRegNo 
            Height          =   285
            Index           =   0
            Left            =   1680
            MaxLength       =   7
            TabIndex        =   41
            Top             =   1800
            Width           =   855
         End
         Begin VB.ListBox lstCourse 
            Height          =   255
            Index           =   0
            ItemData        =   "AddStudentFrm.frx":22501
            Left            =   360
            List            =   "AddStudentFrm.frx":2250E
            TabIndex        =   40
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CommandButton cmdPercent 
            Caption         =   "%"
            Height          =   195
            Left            =   6720
            TabIndex        =   39
            Top             =   1440
            Width           =   615
         End
         Begin VB.CommandButton cmdGrade 
            Caption         =   "Grade"
            Height          =   195
            Left            =   6120
            TabIndex        =   38
            Top             =   1440
            Width           =   615
         End
         Begin VB.Image imgAttach 
            Height          =   375
            Left            =   8280
            Picture         =   "AddStudentFrm.frx":22525
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "** Atleast one should be filled"
            Height          =   195
            Left            =   240
            TabIndex        =   46
            Top             =   3240
            Width           =   2055
         End
         Begin VB.Line Line5 
            X1              =   240
            X2              =   8760
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line4 
            Index           =   0
            X1              =   240
            X2              =   8760
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line3 
            X1              =   6000
            X2              =   7440
            Y1              =   1425
            Y2              =   1425
         End
         Begin VB.Line Line2 
            Index           =   5
            X1              =   6720
            X2              =   6720
            Y1              =   1440
            Y2              =   3120
         End
         Begin VB.Line Line2 
            Index           =   4
            X1              =   1560
            X2              =   1560
            Y1              =   1080
            Y2              =   3120
         End
         Begin VB.Line Line2 
            Index           =   3
            X1              =   2640
            X2              =   2640
            Y1              =   1080
            Y2              =   3120
         End
         Begin VB.Line Line2 
            Index           =   2
            X1              =   3600
            X2              =   3600
            Y1              =   1080
            Y2              =   3120
         End
         Begin VB.Line Line2 
            Index           =   1
            X1              =   6000
            X2              =   6000
            Y1              =   1080
            Y2              =   3120
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   7440
            X2              =   7440
            Y1              =   1080
            Y2              =   3120
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   8760
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Shape Shape6 
            Height          =   2055
            Left            =   240
            Top             =   1080
            Width           =   8535
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "No. of    Chances"
            Height          =   435
            Left            =   7800
            TabIndex        =   37
            Top             =   1200
            Width           =   840
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marks"
            Height          =   195
            Left            =   6480
            TabIndex        =   36
            Top             =   1200
            Width           =   435
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name of the Institution"
            Height          =   195
            Left            =   3960
            TabIndex        =   35
            Top             =   1200
            Width           =   1590
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Year of Passing"
            Height          =   435
            Left            =   2880
            TabIndex        =   34
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg. No."
            Height          =   195
            Left            =   1800
            TabIndex        =   33
            Top             =   1200
            Width           =   645
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name of Course"
            Height          =   195
            Left            =   360
            TabIndex        =   32
            Top             =   1200
            Width           =   1140
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H80000001&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   3495
            Left            =   120
            Top             =   480
            Width           =   8775
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Details of Brother(s) and Sister(s)"
         Height          =   2295
         Left            =   -70200
         TabIndex        =   10
         Top             =   2760
         Width           =   4335
         Begin VB.TextBox txtBroSis 
            Height          =   1575
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   480
            Width           =   3855
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H80000001&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   1815
            Left            =   120
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Guardian Info"
         Height          =   2175
         Left            =   -70200
         TabIndex        =   9
         Top             =   480
         Width           =   4335
         Begin VB.TextBox txtPhone 
            Height          =   285
            Index           =   2
            Left            =   1560
            MaxLength       =   13
            TabIndex        =   29
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox txtGAddress 
            Height          =   735
            Left            =   1560
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtGName 
            Height          =   285
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   27
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label lblGPhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   1560
            Width           =   465
         End
         Begin VB.Label lblGAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   720
            Width           =   570
         End
         Begin VB.Label lblGName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   420
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H80000001&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   1815
            Left            =   120
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Parents Info"
         Height          =   4575
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   4455
         Begin VB.TextBox txtPhone 
            Height          =   285
            Index           =   1
            Left            =   1560
            MaxLength       =   13
            TabIndex        =   23
            Top             =   3240
            Width           =   2535
         End
         Begin VB.ComboBox cmbMOccupation 
            Height          =   315
            ItemData        =   "AddStudentFrm.frx":2449E
            Left            =   1560
            List            =   "AddStudentFrm.frx":244C3
            TabIndex        =   22
            Text            =   "--Select--"
            Top             =   2640
            Width           =   2535
         End
         Begin VB.TextBox txtMotherName 
            Height          =   285
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   21
            Top             =   2280
            Width           =   2535
         End
         Begin VB.TextBox txtFAddress 
            Height          =   735
            Left            =   1560
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000005&
            Caption         =   "Same as student"
            Height          =   255
            Left            =   1560
            TabIndex        =   19
            Top             =   1080
            Width           =   2535
         End
         Begin VB.ComboBox cmbFOccupation 
            Height          =   315
            ItemData        =   "AddStudentFrm.frx":24526
            Left            =   1560
            List            =   "AddStudentFrm.frx":2454B
            TabIndex        =   18
            Text            =   "--Select--"
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtFatherName 
            Height          =   285
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   17
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label lblPhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   3240
            Width           =   465
         End
         Begin VB.Label lblMOccupation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   2760
            Width           =   825
         End
         Begin VB.Label lblMotherName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mother's Name"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   2280
            Width           =   1065
         End
         Begin VB.Label AddressLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   1440
            Width           =   570
         End
         Begin VB.Label lblOccupation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   825
         End
         Begin VB.Label lblFatherName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Father's Name"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   1020
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H80000001&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   4215
            Left            =   120
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Student Info"
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   9015
         Begin VB.TextBox txtPhone 
            Height          =   285
            Index           =   0
            Left            =   5880
            MaxLength       =   13
            TabIndex        =   87
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox txtAddress 
            Height          =   975
            Left            =   5880
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   86
            Top             =   600
            Width           =   2415
         End
         Begin VB.ComboBox cmbIncome 
            Height          =   315
            ItemData        =   "AddStudentFrm.frx":245AE
            Left            =   1440
            List            =   "AddStudentFrm.frx":245BE
            TabIndex        =   84
            Text            =   "--Select--"
            Top             =   3480
            Width           =   2415
         End
         Begin VB.ComboBox cmbBloodGroup 
            Height          =   315
            ItemData        =   "AddStudentFrm.frx":24610
            Left            =   1440
            List            =   "AddStudentFrm.frx":24626
            TabIndex        =   83
            Text            =   "--Select--"
            Top             =   3120
            Width           =   2415
         End
         Begin VB.ComboBox cmbCast 
            Height          =   315
            ItemData        =   "AddStudentFrm.frx":24641
            Left            =   1440
            List            =   "AddStudentFrm.frx":24690
            TabIndex        =   82
            Text            =   "--Select--"
            Top             =   2760
            Width           =   2415
         End
         Begin VB.ComboBox cmbReligion 
            Height          =   315
            ItemData        =   "AddStudentFrm.frx":247EF
            Left            =   1440
            List            =   "AddStudentFrm.frx":247FF
            TabIndex        =   81
            Text            =   "--Select--"
            Top             =   2400
            Width           =   2415
         End
         Begin VB.ComboBox cmbGender 
            Height          =   315
            ItemData        =   "AddStudentFrm.frx":24824
            Left            =   1440
            List            =   "AddStudentFrm.frx":2482E
            TabIndex        =   80
            Text            =   "--Select--"
            Top             =   2040
            Width           =   2415
         End
         Begin VB.ComboBox cmbProgramme 
            Height          =   315
            ItemData        =   "AddStudentFrm.frx":24840
            Left            =   1440
            List            =   "AddStudentFrm.frx":24842
            TabIndex        =   79
            Text            =   "--Select--"
            Top             =   1680
            Width           =   2415
         End
         Begin VB.ComboBox cmbYearOfStudy 
            Height          =   315
            ItemData        =   "AddStudentFrm.frx":24844
            Left            =   1440
            List            =   "AddStudentFrm.frx":248D2
            TabIndex        =   78
            Text            =   "--Select--"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1440
            MaxLength       =   40
            TabIndex        =   77
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtRegNo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0000000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   76
            Top             =   600
            Width           =   2415
         End
         Begin VB.PictureBox Picture1 
            Height          =   1935
            Left            =   6240
            ScaleHeight     =   1875
            ScaleWidth      =   1755
            TabIndex        =   7
            Top             =   2160
            Width           =   1815
            Begin VB.Image imgPhoto 
               Height          =   1650
               Left            =   120
               Top             =   120
               Width           =   1500
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Caption         =   "Photo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   89
               Top             =   840
               Width           =   1215
            End
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1440
            TabIndex        =   85
            Top             =   3840
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Format          =   58720257
            CurrentDate     =   34335
            MaxDate         =   402133
            MinDate         =   2
         End
         Begin VB.Label lblDBirth 
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth"
            Height          =   255
            Left            =   360
            TabIndex        =   71
            Top             =   3840
            Width           =   1215
         End
         Begin VB.Label lblIncome 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Anual Income"
            Height          =   195
            Left            =   360
            TabIndex        =   66
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label lblBloodGroup 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Blood Group"
            Height          =   195
            Left            =   360
            TabIndex        =   65
            Top             =   3120
            Width           =   885
         End
         Begin VB.Label lblCast 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cast"
            Height          =   195
            Left            =   360
            TabIndex        =   64
            Top             =   2760
            Width           =   315
         End
         Begin VB.Label lblReligion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Religion"
            Height          =   195
            Left            =   360
            TabIndex        =   63
            Top             =   2400
            Width           =   570
         End
         Begin VB.Label lblGender 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            Height          =   195
            Left            =   360
            TabIndex        =   62
            Top             =   2040
            Width           =   525
         End
         Begin VB.Label lblProgramme 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Programe"
            Height          =   195
            Left            =   360
            TabIndex        =   61
            Top             =   1680
            Width           =   675
         End
         Begin VB.Label lblPicture 
            Alignment       =   2  'Center
            Caption         =   "Browse Picture"
            Height          =   255
            Left            =   6240
            TabIndex        =   88
            Top             =   4080
            Width           =   1815
         End
         Begin VB.Label lblPhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            Height          =   195
            Index           =   0
            Left            =   5040
            TabIndex        =   6
            Top             =   1680
            Width           =   465
         End
         Begin VB.Label lblAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   5040
            TabIndex        =   5
            Top             =   600
            Width           =   570
         End
         Begin VB.Label YearOfStudyLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year of study"
            Height          =   195
            Left            =   360
            TabIndex        =   4
            Top             =   1320
            Width           =   930
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   360
            TabIndex        =   3
            Top             =   960
            Width           =   420
         End
         Begin VB.Label lblRegNo 
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
            BorderColor     =   &H80000001&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   4095
            Left            =   240
            Top             =   360
            Width           =   8655
         End
      End
   End
End
Attribute VB_Name = "frmAddStudentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fso As New FileSystemObject

Private Sub preview()
   Dim i As Integer
    Dim selLength As Integer
    Dim selStart As Integer
    Dim smFound As Boolean
    Dim fn As Boolean

    rtbinfo(0).Text = rtbinfo(0).Text & _
    vbNewLine
    
    rtbinfo(0).Text = ""
    rtbinfo(0).Text = rtbinfo(0).Text & "Personal Information" & vbNewLine
    
    rtbinfo(0).Text = rtbinfo(0).Text & _
    vbNewLine
    
    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Student Register Number: " & Me.txtRegNo(1).Text & vbNewLine

    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Name: " & Me.txtName.Text & vbNewLine
    
    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Year of Study: " & Me.cmbYearOfStudy.Text & vbNewLine

    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Programme: " & Me.cmbProgramme.Text & vbNewLine

    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Gender: " & Me.cmbGender.Text & vbNewLine

    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Religion: " & Me.cmbReligion.Text & vbNewLine

    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Cast: " & Me.cmbCast.Text & vbNewLine
    
    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Blood Group: " & Me.cmbBloodGroup.Text & vbNewLine

    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Anual Income: " & Me.cmbIncome.Text & vbNewLine
    
    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Date of Birth: " & Me.DTPicker1.Value & vbNewLine
    
    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Address: " & Me.txtAddress.Text & vbNewLine
    
    rtbinfo(0).Text = rtbinfo(0).Text & _
    "Phone: " & Me.txtPhone(0).Text & vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & "Family Information" & vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Father's Name: " & Me.txtFatherName.Text & vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Occupation: " & Me.cmbFOccupation.Text & vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Mother's Name: " & Me.txtMotherName.Text & vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Occupation: " & Me.cmbMOccupation.Text & vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Parent Phone No: " & Me.txtPhone(1).Text & vbNewLine
   
     rtbinfo(1).Text = rtbinfo(1).Text & _
    "Guardian Name: " & Me.txtGName.Text & vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Guardian Phone No: " & Me.txtPhone(2).Text & vbNewLine
        
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Details of Brother(s) and Sister(s) " & vbNewLine
     
    rtbinfo(1).Text = rtbinfo(1).Text & _
    Me.txtBroSis.Text & vbNewLine
   
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Previous Examination Information" & vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Name of Course: " & Me.lstCourse(0).Text & vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Register No: " & Me.txtRegNo(0).Text & vbNewLine
    
    If cmdGrade.Visible = True Then
       rtbinfo(1).Text = rtbinfo(1).Text & _
       "Mark Obtained: " & Me.cmbGrade(0).Text & vbNewLine
    Else
        rtbinfo(1).Text = rtbinfo(1).Text & _
       "Mark Obtained: " & Me.cmbPercent(0).Text & vbNewLine
    End If
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Name of Course: " & Me.lstCourse(1).Text & vbNewLine
    
    rtbinfo(1).Text = rtbinfo(1).Text & _
    "Register No: " & Me.txtRegNo(2).Text & vbNewLine
    
    If cmdGrade.Visible = True Then
       rtbinfo(1).Text = rtbinfo(1).Text & _
       "Mark Obtained: " & Me.cmbGrade(1).Text & vbNewLine
    Else
        rtbinfo(1).Text = rtbinfo(1).Text & _
       "Mark Obtained: " & Me.cmbPercent(1).Text & vbNewLine
    End If
    
    'set color
    
    rtbinfo(0).selStart = 0
    rtbinfo(0).selLength = Len(rtbinfo(0).Text)
    rtbinfo(0).SelColor = &H584620
    rtbinfo(0).SelBold = False
    
    For i = 1 To Len(rtbinfo(0).Text) + 1
    
        If Mid(rtbinfo(0).Text, i, 1) = ":" Then
            smFound = True
            selStart = i
            selLength = 0
        End If
        
        If smFound = True Then
            selLength = selLength + 1

            If Mid(rtbinfo(0).Text, i, 2) = vbNewLine Then
                
                rtbinfo(0).selStart = selStart
                rtbinfo(0).selLength = selLength
                rtbinfo(0).SelFontSize = 10
                rtbinfo(0).SelColor = &H0&
                rtbinfo(0).SelBold = True
                
                If fn = False Then
                    rtbinfo(0).SelFontSize = 10
                    fn = True
                End If
                
                rtbinfo(0).selLength = 0
                
                smFound = False
            End If
            
        End If
    Next
    
    rtbinfo(1).selStart = 0
    rtbinfo(1).selLength = Len(rtbinfo(1).Text)
    rtbinfo(1).SelColor = &H584620
    rtbinfo(1).SelBold = False
    
    For i = 1 To Len(rtbinfo(1).Text) + 1
    
        If Mid(rtbinfo(1).Text, i, 1) = ":" Then
            smFound = True
            selStart = i
            selLength = 0
        End If
        
        If smFound = True Then
            selLength = selLength + 1

            If Mid(rtbinfo(1).Text, i, 2) = vbNewLine Then
                
                rtbinfo(1).selStart = selStart
                rtbinfo(1).selLength = selLength
                rtbinfo(1).SelFontSize = 10
                rtbinfo(1).SelColor = &H0&
                rtbinfo(1).SelBold = True
                
                If fn = False Then
                    rtbinfo(1).SelFontSize = 10
                    fn = True
                End If
                
                rtbinfo(1).selLength = 0
                
                smFound = False
            End If
        End If
    Next
End Sub

Private Sub Check1_Click()
  If Check1.Value = Checked Then
    txtFAddress.Text = Trim(txtAddress.Text)
    'txtMAddress.Text = Trim(txtAddress.Text)
    txtGAddress.Text = Trim(txtAddress.Text)
    'txtFAddress.Refresh
    'txtMAddress.Refresh
    'txtGAddress.Refresh
  Else
    If Check1.Value = Unchecked Then
       txtFAddress.Text = " "
      ' txtMAddress.Text = " "
       txtGAddress.Text = " "
    End If
  End If
End Sub

Private Sub cmdGrade_Click()
  cmbPercent(0).Visible = False
  cmbPercent(1).Visible = False
  cmbPercent(2).Visible = False
  cmbGrade(0).Visible = True
  cmbGrade(1).Visible = True
  cmbGrade(2).Visible = True
End Sub

Private Sub cmdPercent_Click()
  cmbGrade(0).Visible = False
  cmbGrade(1).Visible = False
  cmbGrade(2).Visible = False
  cmbPercent(0).Visible = True
  cmbPercent(1).Visible = True
  cmbPercent(2).Visible = True
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub Form_GotFocus()
  Unload frmLoading
  
End Sub

Private Sub Form_Load()
  Theme frmAddStudentInfo
  tabAddinfo.TabIndex = 0
  
  connection
  
End Sub

Private Sub imgAttach_Click()
  frmAttachment.imgReport.Enabled = False
  frmAttachment.txtRegNo.Text = frmAddStudentInfo.txtRegNo(1).Text
  frmAttachment.Show
End Sub

Private Sub imgCancel_Click(Index As Integer)
  Unload Me
End Sub

Private Sub imgNext_Click(Index As Integer)
  Dim i As Integer
  i = 0
'  tabAddinfo.Tab = Index + 1
  If Index = 2 Then preview
  Select Case Index
    Case 0:
             If txtRegNo(1).Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 0
                Dialog.Label1(i).Caption = "Register Number left blank"
                i = i + 1
             End If
             
             If txtName.Text = "" Then
                Dialog.Label1(i).Visible = True
                
                Dialog.Label2.Caption = 200
                Dialog.Label1(i).Caption = "Student Name left blank"
                i = i + 1
             End If
             If cmbYearOfStudy.Text = "" Or cmbYearOfStudy.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 400
                Dialog.Label1(i).Caption = "Year of study left blank "
                i = i + 1
             End If
             If cmbProgramme.Text = "" Or cmbProgramme.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 600
                Dialog.Label1(i).Caption = "Programme left blank"
                i = i + 1
             End If
             If cmbGender.Text = "" Or cmbGender.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 800
                Dialog.Label1(i).Caption = "Gender left blank"
                i = i + 1
             End If
             If cmbReligion.Text = "" Or cmbReligion.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1000
                Dialog.Label1(i).Caption = "Religion left blank"
                i = i + 1
             End If
             If cmbCast.Text = "" Or cmbCast.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1200
                Dialog.Label1(i).Caption = "Cast left blank"
                i = i + 1
             End If
             If cmbBloodGroup.Text = "" Or cmbBloodGroup.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1400
                Dialog.Label1(i).Caption = "Blood group left blank"
                i = i + 1
             End If
             If cmbIncome.Text = "" Or cmbIncome.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1600
                Dialog.Label1(i).Caption = "Anual income left blank"
                i = i + 1
             End If
             If txtAddress.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1800
                Dialog.Label1(i).Caption = "Address left blank"
                i = i + 1
             End If
             If txtPhone(0).Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 2000
                Dialog.Label1(i).Caption = "Phone left blank"
                i = i + 1
             End If
             If imgPhoto.Picture = Null Then Dialog.Label1(i).Caption = "Choose student photo"
             
             If i <> 0 Then
                tabAddinfo.Tab = 0
                Dialog.Show
             Else
                tabAddinfo.Tab = tabAddinfo.Tab + 1
             End If
    Case 1: i = 0
            If txtFatherName.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 0
                Dialog.Label1(i).Caption = "Father Name left blank"
                i = i + 1
             End If
             If cmbFOccupation.Text = "" Or cmbIncome.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 200
                Dialog.Label1(i).Caption = "Father's Occupation left blank"
                i = i + 1
             End If
             
             If txtFAddress.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 400
                Dialog.Label1(i).Caption = "Father's Address left blank"
                i = i + 1
             End If
             If txtMotherName.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 600
                Dialog.Label1(i).Caption = "Mother Name left blank"
                i = i + 1
             End If
             If cmbMOccupation.Text = "" Or cmbIncome.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 800
                Dialog.Label1(i).Caption = "Mother's Occupation left blank"
                i = i + 1
             End If
             'If txtMAddress.Text = "" Then
              '  Dialog.Label1(i).Visible = True
               ' Dialog.Label2.Caption = 1000
                'Dialog.Label1(i).Caption = "Mother's Address left blank"
                'i = i + 1
             'End If
             If txtGName.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1200
                Dialog.Label1(i).Caption = "Guardian Name left blank"
                i = i + 1
             End If
             If txtGAddress.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1400
                Dialog.Label1(i).Caption = "Guardian's Address left blank"
                i = i + 1
             End If
             
             If i <> 0 Then
                tabAddinfo.Tab = 1
                Dialog.Show
             Else
                tabAddinfo.Tab = tabAddinfo.Tab + 1
             End If
 
    Case 2:  lstCourse(0).Text = lstCourse(0).List(0)
             lstCourse(1).Text = lstCourse(1).List(0)
             lstCourse(2).Text = lstCourse(2).List(0)
             lstNoOfChances(0).Text = lstNoOfChances(0).List(0)
             lstNoOfChances(1).Text = lstNoOfChances(1).List(0)
             lstNoOfChances(2).Text = lstNoOfChances(2).List(0)
             
             i = 0
             'If lstCourse(0).Text = "" Or lstCourse(1).Text = "" Then
              '  Dialog.Label1(i).Visible = True
               ' Dialog.Label2.Caption = 0
                'Dialog.Label1(i).Caption = "Course left blank"
                'i = i + 1
             'End If
             If txtRegNo(0).Text = "" Or txtRegNo(2).Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 200
                Dialog.Label1(i).Caption = "Register Number left blank"
                i = i + 1
             End If
             If cmbYearOfPassing(0).Text = "" Or cmbYearOfPassing(0).Text = "Year" Or cmbYearOfPassing(1).Text = "" Or cmbYearOfPassing(1).Text = "Year" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 400
                Dialog.Label1(i).Caption = "Year of Passing left blank"
                i = i + 1
             End If
             If cmbSchoolName(0).Text = "" Or cmbSchoolName(0).Text = "--Select--" Or cmbSchoolName(1).Text = "" Or cmbSchoolName(1).Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 600
                Dialog.Label1(i).Caption = "School Name left blank"
                i = i + 1
             End If
             If cmbGrade(0).Visible = True Then
                If cmbGrade(0).Text = "" Or cmbGrade(1).Text = "" Then
                   Dialog.Label1(i).Visible = True
                   Dialog.Label2.Caption = 800
                   Dialog.Label1(i).Caption = "Grade left blank"
                   i = i + 1
                End If
             End If
             If cmbPercent(0).Visible = True Then
                If cmbPercent(0).Text = "" Or cmbPercent(1).Text = "" Then
                   Dialog.Label1(i).Visible = True
                   Dialog.Label2.Caption = 1000
                   Dialog.Label1(i).Caption = "Percentage left blank"
                   i = i + 1
                End If
             End If
             'If lstNoOfChances(0).Text = "" Or lstNoOfChances(1).Text = "" Then
              '  Dialog.Label1(i).Visible = True
              '  Dialog.Label2.Caption = 1200
              '  Dialog.Label1(i).Caption = "No of Chances left blank"
               ' i = i + 1
             'End If
             
             If i <> 0 Then
                tabAddinfo.Tab = 2
                Dialog.Show
             Else
                tabAddinfo.Tab = tabAddinfo.Tab + 1
             End If
 
    Case 3: 'save settings
  End Select
End Sub

Private Sub InfoValidate()
  Dim i As Integer
  i = 0
  Select Case tabAddinfo.Caption
    Case "Family Information":
            If txtRegNo(1).Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 0
                Dialog.Label1(i).Caption = "Register Number left blank"
                i = i + 1
             End If
             
             If txtName.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 200
                Dialog.Label1(i).Caption = "Student Name left blank"
                i = i + 1
             End If
             If cmbYearOfStudy.Text = "" Or cmbYearOfStudy.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 400
                Dialog.Label1(i).Caption = "Year of study left blank "
                i = i + 1
             End If
             If cmbProgramme.Text = "" Or cmbProgramme.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 600
                Dialog.Label1(i).Caption = "Programme left blank"
                i = i + 1
             End If
             If cmbGender.Text = "" Or cmbGender.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 800
                Dialog.Label1(i).Caption = "Gender left blank"
                i = i + 1
             End If
             If cmbReligion.Text = "" Or cmbReligion.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1000
                Dialog.Label1(i).Caption = "Religion left blank"
                i = i + 1
             End If
             If cmbCast.Text = "" Or cmbCast.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1200
                Dialog.Label1(i).Caption = "Cast left blank"
                i = i + 1
             End If
             If cmbBloodGroup.Text = "" Or cmbBloodGroup.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1400
                Dialog.Label1(i).Caption = "Blood group left blank"
                i = i + 1
             End If
             If cmbIncome.Text = "" Or cmbIncome.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1600
                Dialog.Label1(i).Caption = "Anual income left blank"
                i = i + 1
             End If
             If txtAddress.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1800
                Dialog.Label1(i).Caption = "Address left blank"
                i = i + 1
             End If
             If txtPhone(0).Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 2000
                Dialog.Label1(i).Caption = "Phone left blank"
                i = i + 1
             End If
             If imgPhoto.Picture = Null Then Dialog.Label1(i).Caption = "Choose student photo"
             If i <> 0 Then
                'Dialog.Show
                tabAddinfo.Tab = 0
             Else
                tabAddinfo.Tab = tabAddinfo.Tab + 1
             End If
    Case "Examination Information": i = 0
            If txtFatherName.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 0
                Dialog.Label1(i).Caption = "Father Name left blank"
                i = i + 1
             End If
             If cmbFOccupation.Text = "" Or cmbIncome.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1600
                Dialog.Label1(i).Caption = "Father's Occupation left blank"
                i = i + 1
             End If
             
             If txtFAddress.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 200
                Dialog.Label1(i).Caption = "Father's Address left blank"
                i = i + 1
             End If
             If txtMotherName.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 0
                Dialog.Label1(i).Caption = "Mother Name left blank"
                i = i + 1
             End If
             If cmbMOccupation.Text = "" Or cmbIncome.Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1600
                Dialog.Label1(i).Caption = "Mother's Occupation left blank"
                i = i + 1
             End If
             'If txtMAddress.Text = "" Then
              '  Dialog.Label1(i).Visible = True
               ' Dialog.Label2.Caption = 200
                'Dialog.Label1(i).Caption = "Mother's Address left blank"
                'i = i + 1
             'End If
             If txtGName.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 0
                Dialog.Label1(i).Caption = "Guardian Name left blank"
                i = i + 1
             End If
             If txtGAddress.Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 200
                Dialog.Label1(i).Caption = "Guardian's Address left blank"
                i = i + 1
             End If
             If i <> 0 Then
                tabAddinfo.Tab = 1
                'Dialog.Show
             Else
                tabAddinfo.Tab = tabAddinfo.Tab + 1
             End If
 
    Case "Summary":  i = 0
             If lstCourse(0).Text = "" Or lstCourse(0).Text = "--Select--" Or lstCourse(1).Text = "" Or lstCourse(1).Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 0
                Dialog.Label1(i).Caption = "Course left blank"
                i = i + 1
             End If
             If txtRegNo(0).Text = "" Or txtRegNo(2).Text = "" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 200
                Dialog.Label1(i).Caption = "Register Number left blank"
                i = i + 1
             End If
             If cmbYearOfPassing(0).Text = "" Or cmbYearOfPassing(0).Text = "--Select--" Or cmbYearOfPassing(1).Text = "" Or cmbYearOfPassing(1).Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 400
                Dialog.Label1(i).Caption = "Year of Passing left blank"
                i = i + 1
             End If
             If cmbSchoolName(0).Text = "" Or cmbSchoolName(0).Text = "--Select--" Or cmbSchoolName(1).Text = "" Or cmbSchoolName(1).Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 600
                Dialog.Label1(i).Caption = "School Name left blank"
                i = i + 1
             End If
             
             If cmbGrade(0).Text = "" Or cmbGrade(0).Text = "--Select--" Or cmbGrade(1).Text = "" Or cmbGrade(1).Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 800
                Dialog.Label1(i).Caption = "Grade left blank"
                i = i + 1
             End If
             If cmbPercent(0).Text = "" Or cmbPercent(0).Text = "--Select--" Or cmbPercent(1).Text = "" Or cmbPercent(1).Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1000
                Dialog.Label1(i).Caption = "Percentage left blank"
                i = i + 1
             End If
             If lstNoOfChances(0).Text = "" Or lstNoOfChances(0).Text = "--Select--" Or lstNoOfChances(1).Text = "" Or lstNoOfChances(1).Text = "--Select--" Then
                Dialog.Label1(i).Visible = True
                Dialog.Label2.Caption = 1200
                Dialog.Label1(i).Caption = "No of Chances left blank"
                i = i + 1
             End If
             
             If i <> 0 Then
                tabAddinfo.Tab = 2
                'Dialog.Show
             Else
                tabAddinfo.Tab = tabAddinfo.Tab + 1
             End If
 
    'save details
  End Select
    
End Sub

Private Sub imgBack_Click(Index As Integer)
    If Index > -1 Then
       tabAddinfo.Tab = Index
    End If
End Sub

Private Sub lblPicture_Click()
  MsgBox "Size of the picture must be 100 x 100", vbExclamation
  
  With CommonDialog1
    .FileName = ""
    .Filter = "All Picture Files | *.bmp;*.jpg;*.gif | JPEG (*.jpg) | *.jpg | Bitmap (*.bmp) | *.bmp "
    .DialogTitle = "Open Image..."
    .CancelError = True
    .ShowOpen
    If .FileName <> "" Then
      imgPhoto.Picture = LoadPicture(CommonDialog1.FileName)
      If imgPhoto.height > 1650 Or imgPhoto.width > 1500 Then MsgBox "Picture Dimension is not suitable", vbCritical
    End If
  End With
End Sub

Private Sub imgSave_Click()
 ' On Error GoTo error_para
  
  Dim income As Double
  Dim dob As String
  Dim picpath As String
  dob = Format(DTPicker1.Value, "mm-dd-yyyy")
  Select Case cmbIncome
     Case "Below 25000": income = 25000
     Case "Between 25000 and 50000": income = 50000
     Case "Between 50000 and 2 Lakh": income = 100000
     Case "Above 2 Lakh": income = 200000
  End Select
  
    picpath = App.Path & "\sis\photo\" & Val(txtRegNo(1).Text) + 1729 & ".jpg"
  
    If MsgBox("Are You sure, The entered details are correct?", vbQuestion + vbYesNo, "Conformation") = vbYes Then
    
     
      If fso.FolderExists(App.Path & "\sis") = False Then fso.CreateFolder App.Path & "\sis"
      If fso.FolderExists(App.Path & "\sis\photo") = False Then fso.CreateFolder App.Path & "\sis\photo"
      fso.CopyFile CommonDialog1.FileName, picpath
    
      reccheck
    
      rec.Open "select REGNO from MAINTABLE where REGNO ='" & Trim(txtRegNo(1).Text) & "'", con, adOpenDynamic, adLockPessimistic
      If rec.EOF = False Then
        If MsgBox("Student information is already in database, Do you wish to Replace it with this one?", vbYesNo, "Data entry error") = vbNo Then
           Exit Sub
        Else
          reccheck
          rec.Open "delete from MAINTABLE where REGNO = " & Trim(txtRegNo(1).Text), con, adOpenDynamic, adLockPessimistic
         
          reccheck
          rec.Open "delete from PERSONALTABLE where REGNO = " & Trim(txtRegNo(1).Text), con, adOpenDynamic, adLockPessimistic
         
          reccheck
          rec.Open "delete from FAMILYTABLE where REGNO = " & Trim(txtRegNo(1).Text), con, adOpenDynamic, adLockPessimistic
         
          reccheck
          rec.Open "delete from EXAMTABLE where EXAMNO = " & Trim(txtRegNo(0).Text), con, adOpenDynamic, adLockPessimistic
      
          reccheck
          rec.Open "delete from EXAMTABLE where EXAMNO = " & Trim(txtRegNo(2).Text), con, adOpenDynamic, adLockPessimistic
      
          If Trim(txtRegNo(3).Text) <> "" Then
             reccheck
             rec.Open "delete from EXAMTABLE where EXAMNO = " & Trim(txtRegNo(3).Text), con, adOpenDynamic, adLockPessimistic
          End If
        End If
      End If
    
      reccheck
      rec.Open "insert into MAINTABLE values('" & Trim(txtRegNo(1).Text) & "','" & Trim(txtName.Text) & "','" & Trim(cmbProgramme.Text) & "','" & Trim(cmbYearOfStudy.Text) & "')", con, adOpenDynamic, adLockPessimistic
    
      reccheck
      rec.Open "insert into PERSONALTABLE values('" & Trim(txtRegNo(1).Text) & "','" & Trim(txtName.Text) & "','" & Val(cmbYearOfStudy.Text) & "','" & _
      Trim(cmbProgramme.Text) & "','" & Trim(cmbGender.Text) & "','" & Trim(cmbReligion.Text) & "','" & Trim(cmbCast.Text) & "','" & Trim(cmbBloodGroup.Text) & "','" & income & "','" & dob & "','" & Trim(txtAddress.Text) & "','" & Trim(txtPhone(0).Text) & "','" & picpath & "')", con, adOpenDynamic, adLockOptimistic
    
      reccheck
      rec.Open "insert into FAMILYTABLE values('" & Trim(txtRegNo(1).Text) & "','" & Trim(txtFatherName.Text) & "','" & Trim(cmbFOccupation.Text) & "','" & _
      Trim(txtFAddress.Text) & "','" & Trim(txtMotherName.Text) & "','" & Trim(cmbMOccupation.Text) & "','" & Trim(txtPhone(1).Text) & "','" & Trim(txtGName.Text) & "','" & Trim(txtGAddress.Text) & "','" & Trim(txtPhone(2).Text) & "','" & Trim(txtBroSis.Text) & "')", con, adOpenDynamic, adLockOptimistic

      reccheck
    
      rec.Open "insert into EXAMTABLE values('" & Trim(txtRegNo(1).Text) & "','" & Trim(txtRegNo(0).Text) & "','" & Trim(lstCourse(0).Text) & "','" & _
      Val(Trim(cmbYearOfPassing(0).Text)) & "','" & Trim(cmbSchoolName(0).Text) & "','" & Trim(cmbGrade(0).Text) & "','" & Trim(cmbPercent(0).Text) & "','" & Trim(lstNoOfChances(0).Text) & "')", con, adOpenDynamic, adLockPessimistic
    
      reccheck
      rec.Open "insert into EXAMTABLE values('" & Trim(txtRegNo(1).Text) & "','" & Trim(txtRegNo(2).Text) & "','" & Trim(lstCourse(1).Text) & "','" & _
      Val(Trim(cmbYearOfPassing(1).Text)) & "','" & Trim(cmbSchoolName(1).Text) & "','" & Trim(cmbGrade(1).Text) & "','" & Trim(cmbPercent(1).Text) & "','" & Trim(lstNoOfChances(1).Text) & "')", con, adOpenDynamic, adLockPessimistic
   
      If txtRegNo(3).Text <> "" Then
         connection
         reccheck
         rec.Open "insert into EXAMTABLE values('" & Trim(txtRegNo(1).Text) & "','" & Trim(txtRegNo(3).Text) & "','" & Trim(lstCourse(2).Text) & "','" & _
         Val(Trim(cmbYearOfPassing(2).Text)) & "','" & Trim(cmbSchoolName(2).Text) & "','" & Trim(cmbGrade(2).Text) & "','" & Trim(cmbPercent(2).Text) & "','" & Trim(lstNoOfChances(2).Text) & "')", con, adOpenDynamic, adLockPessimistic
      End If
    
      MsgBox "Data Added Successfully"
      frmPrintDetails.Show
    
  End If
'error_para:
  ' MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub tabAddinfo_GotFocus()
  tabAddinfo.Tab = 0
  txtRegNo(1).SetFocus
  cmdPercent_Click
End Sub

Private Sub cmbProgramme_LostFocus()
 '  On Error GoTo error_para
   
   Dim Programme As String
   If cmbProgramme.Text = "Others" Then
      Programme = InputBox("Enter the Programme Name : ", "Programme Name")
      cmbProgramme.Text = Programme
      If CheckCombo(cmbProgramme, " Programme Name") = True Then
         reccheck
         rec.Open "delete from PROGRAMME where NAME = 'Others'", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into PROGRAMME values('" & Programme & "')", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into PROGRAMME values('Others')", con, adOpenDynamic, adLockPessimistic
      End If
   End If
'error_para:
  ' MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbProgramme_Click()
 '  On Error GoTo error_para
   
   Dim Programme As String
   If cmbProgramme.Text = "Others" Then
      Programme = InputBox("Enter the Programme Name : ", "Programme Name")
      cmbProgramme.Text = Programme
      cmbProgramme.Refresh
      If CheckCombo(cmbProgramme, " Programme Name") = True Then
         reccheck
         rec.Open "delete from PROGRAMME where NAME = 'Others'", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into PROGRAMME values('" & Programme & "')", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into PROGRAMME values('Others')", con, adOpenDynamic, adLockPessimistic
      End If
   End If
'error_para:
  ' MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbCast_Click()
 '  On Error GoTo error_para

   Dim Cast As String
   If cmbCast.Text = "Others" Then
      Cast = InputBox("Enter the Case Name : ", "Case Name")
      cmbCast.Text = Cast
      cmbCast.Refresh
      If CheckCombo(cmbCast, " Cast Name") = True Then
         reccheck
         rec.Open "delete from CAST where NAME = 'Others'", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into CAST values('" & Cast & "')", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into CAST values('Others')", con, adOpenDynamic, adLockPessimistic
      End If
   End If
'error_para:
   'MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbCast_LostFocus()
  ' On Error GoTo error_para
   
   Dim Cast As String
   If cmbCast.Text = "Others" Then
      Cast = InputBox("Enter the Case Name : ", "Case Name")
      cmbCast.Text = Cast
      cmbCast.Refresh
      If CheckCombo(cmbCast, " Cast Name") = True Then
         reccheck
         rec.Open "delete from CAST where NAME = 'Others'", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into CAST values('" & Cast & "')", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into CAST values('Others')", con, adOpenDynamic, adLockPessimistic
      End If
   End If
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbReligion_Click()
  ' On Error GoTo error_para
   
   Dim Religion As String
   If cmbReligion.Text = "Others" Then
      Religion = InputBox("Enter the Religion Name : ", "Religion Name")
      cmbReligion.Text = Religion
      cmbReligion.Refresh
      If CheckCombo(cmbReligion, " Religion Name") = True Then
         reccheck
         rec.Open "delete from RELIGION where NAME = 'Others'", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into RELIGION values('" & Religion & "')", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into RELIGION values('Others')", con, adOpenDynamic, adLockPessimistic
      End If
   End If
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbReligion_LostFocus()
 '  On Error GoTo error_para
   
   Dim Religion As String
   If cmbReligion.Text = "Others" Then
      Religion = InputBox("Enter the Religion Name : ", "Religion Name")
      cmbReligion.Text = Religion
      cmbReligion.Refresh
      If CheckCombo(cmbReligion, " Religion Name") = True Then
         reccheck
         rec.Open "delete from RELIGION where NAME = 'Others'", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into RELIGION values('" & Religion & "')", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into RELIGION values('Others')", con, adOpenDynamic, adLockPessimistic
      End If
   End If
'error_para:
  ' MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbSchoolName_Click(Index As Integer)
  ' On Error GoTo error_para
   
   Dim school As String
   If cmbSchoolName(Index).Text = "Others" Then
      school = InputBox("Enter the School Name : ", "School Name")
      cmbSchoolName(Index).Text = school
      cmbSchoolName(Index).Refresh
      If CheckCombo(cmbSchoolName(Index), " School Name") = True Then
         reccheck
         rec.Open "delete from SCHOOL where NAME = 'Others'", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into SCHOOL values('" & school & "')", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into SCHOOL values('Others')", con, adOpenDynamic, adLockPessimistic
      End If
   End If
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbSchoolName_LostFocus(Index As Integer)
   'On Error GoTo error_para
   
   Dim school As String
   If cmbSchoolName(Index).Text = "Others" Then
      school = InputBox("Enter the School Name : ", "School Name")
      cmbSchoolName(Index).Text = school
      cmbSchoolName(Index).Refresh
      If CheckCombo(cmbSchoolName, " School Name") = True Then
         reccheck
         rec.Open "delete from SCHOOL where NAME = 'Others'", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into SCHOOL values('" & school & "')", con, adOpenDynamic, adLockPessimistic
         reccheck
         rec.Open "insert into SCHOOL values('Others')", con, adOpenDynamic, adLockPessimistic
      End If
   End If
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

'Validation of fields in the first tab

Private Sub txtRegNo_KeyPress(Index As Integer, KeyAscii As Integer)
   ValRegNo KeyAscii
End Sub

Private Sub txtName_GotFocus()
  If CheckRegNo(txtRegNo(1), 8) = True Then
     txtName.SetFocus
  End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
   ValName KeyAscii
End Sub

Private Sub cmbYearOfStudy_GotFocus()
   If CheckName(txtName, 30) = True Then
      cmbYearOfStudy.SetFocus
   End If
End Sub

Private Sub cmbProgramme_GotFocus()
  'On Error GoTo error_para
  
  If CheckCombo(cmbYearOfStudy, " YearOfStudy") = True Then
      cmbProgramme.SetFocus
      cmbProgramme.Clear
      
      reccheck
  
      rec.Open "select * from PROGRAMME", con, adOpenDynamic, adLockPessimistic
      If rec.EOF = False Then
        cmbProgramme.Text = rec.Fields(0)
        While Not rec.EOF
          cmbProgramme.AddItem (rec.Fields(0))
          rec.MoveNext
        Wend
      End If
  End If
'error_para:
  ' MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbGender_GotFocus()
  If CheckCombo(cmbYearOfStudy, " Programme") = True Then
     cmbGender.SetFocus
  End If
End Sub

Private Sub cmbReligion_GotFocus()
 ' On Error GoTo error_para
  
  If CheckCombo(cmbGender, " Gender") = True Then
     cmbReligion.SetFocus
     cmbReligion.Clear
      
     reccheck
  
     rec.Open "select * from RELIGION", con, adOpenDynamic, adLockPessimistic
     If rec.EOF = False Then
        cmbReligion.Text = rec.Fields(0)
        While Not rec.EOF
          cmbReligion.AddItem (rec.Fields(0))
          rec.MoveNext
        Wend
     End If
  End If
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbCast_GotFocus()
  'On Error GoTo error_para
  
  If CheckCombo(cmbReligion, " Religion") = True Then
     cmbCast.SetFocus
     cmbCast.Clear
      
     reccheck
  
     rec.Open "select * from CASTE", con, adOpenDynamic, adLockPessimistic
     If rec.EOF = False Then
        cmbCast.Text = rec.Fields(0)
        While Not rec.EOF
          cmbCast.AddItem (rec.Fields(0))
          rec.MoveNext
        Wend
     End If
  End If
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub cmbBloodGroup_GotFocus()
  If CheckCombo(cmbCast, " Cast") = True Then
     cmbBloodGroup.SetFocus
  End If
End Sub

Private Sub cmbIncome_GotFocus()
  If CheckCombo(cmbBloodGroup, " BloodGroup") = True Then
     cmbIncome.SetFocus
  End If
End Sub

Private Sub DTPicker1_GotFocus()
  If CheckCombo(cmbIncome, " Income") = True Then
     DTPicker1.SetFocus
  End If
End Sub

Private Sub txtAddress_GotFocus()
'  If CheckCombo(DTPicker1, " Date of birth") = True Then
 '    DTPicker1.SetFocus
 ' End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
   ValAddress KeyAscii
End Sub

Private Sub txtPhone_GotFocus(Index As Integer)
  If CheckName(txtAddress, 70) = True Then
     'Do nothing
  End If
End Sub

Private Sub cmbSchoolName_GotFocus(Index As Integer)
  '  On Error GoTo error_para
    
    cmbSchoolName(Index).Clear
      
    reccheck
  
     rec.Open "select * from SCHOOL", con, adOpenDynamic, adLockPessimistic
     If rec.EOF = False Then
        cmbSchoolName(Index).Text = rec.Fields(0)
        While Not rec.EOF
          cmbSchoolName(Index).AddItem (rec.Fields(0))
          rec.MoveNext
        Wend
     End If
'error_para:
 '  MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub txtPhone_KeyPress(Index As Integer, KeyAscii As Integer)
   ValPhone KeyAscii
End Sub

Private Sub txtFatherName_KeyPress(KeyAscii As Integer)
   ValName KeyAscii
End Sub

Private Sub txtFAddress_KeyPress(KeyAscii As Integer)
   ValAddress KeyAscii
End Sub

Private Sub txtMotherName_KeyPress(KeyAscii As Integer)
   ValName KeyAscii
End Sub

Private Sub txtMAddress_KeyPress(KeyAscii As Integer)
   ValAddress KeyAscii
End Sub

Private Sub txtGName_KeyPress(KeyAscii As Integer)
   ValName KeyAscii
End Sub


Private Sub txtGAddress_KeyPress(KeyAscii As Integer)
  ValAddress KeyAscii
End Sub
