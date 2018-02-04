VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form StudentFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Entry Form"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9000
      Top             =   720
   End
   Begin TabDlg.SSTab SSTab_AddStudent 
      Height          =   5940
      Left            =   120
      TabIndex        =   64
      Top             =   600
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   10478
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16054778
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Student Information"
      TabPicture(0)   =   "AddStudentFrm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label37"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label38"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label29"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label30"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label31"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label32"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label33"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label34"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label35"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label36"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Religion"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label20"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDOB"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Container4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdNext1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "jcbutton1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtmContact"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtPOB"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtID"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cboReligion"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtFname"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtMname"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtLname"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtAge"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtCitizenship"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cboSex"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cboStatus"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cboBloodType"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtAdd"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cd"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Command1"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).ControlCount=   37
      TabCaption(1)   =   "Parents Information"
      TabPicture(1)   =   "AddStudentFrm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Container1"
      Tab(1).Control(1)=   "jcbutton2"
      Tab(1).Control(2)=   "cmdBack1"
      Tab(1).Control(3)=   "jcbutton3"
      Tab(1).Control(4)=   "Shape2"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "School Information"
      TabPicture(2)   =   "AddStudentFrm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Container3"
      Tab(2).Control(1)=   "jcbutton4"
      Tab(2).Control(2)=   "jcbutton6"
      Tab(2).Control(3)=   "jcbutton5"
      Tab(2).Control(4)=   "Shape1"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Summary"
      TabPicture(3)   =   "AddStudentFrm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "rtbInfo"
      Tab(3).Control(1)=   "jcbutton7"
      Tab(3).Control(2)=   "jcbutton9"
      Tab(3).Control(3)=   "jcbutton8"
      Tab(3).Control(4)=   "Shape4"
      Tab(3).ControlCount=   5
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1350
         Width           =   2775
      End
      Begin VB.PictureBox Command1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5400
         ScaleHeight     =   285
         ScaleWidth      =   2745
         TabIndex        =   15
         Top             =   3960
         Width           =   2805
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   9000
         Top             =   4440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtAdd 
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox cboBloodType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "AddStudentFrm.frx":0070
         Left            =   1320
         List            =   "AddStudentFrm.frx":0080
         TabIndex        =   10
         Top             =   4200
         Width           =   2745
      End
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "AddStudentFrm.frx":0091
         Left            =   1320
         List            =   "AddStudentFrm.frx":009E
         TabIndex        =   9
         Top             =   3840
         Width           =   2775
      End
      Begin VB.ComboBox cboSex 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "AddStudentFrm.frx":00BA
         Left            =   1350
         List            =   "AddStudentFrm.frx":00C4
         TabIndex        =   4
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtCitizenship 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "Filipino"
         Top             =   3480
         Width           =   2745
      End
      Begin VB.TextBox txtAge 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   3120
         Width           =   2745
      End
      Begin VB.TextBox txtLname 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   1
         Top             =   960
         Width           =   2745
      End
      Begin VB.TextBox txtMname 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1680
         Width           =   2745
      End
      Begin VB.TextBox txtFname 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1320
         Width           =   2745
      End
      Begin VB.TextBox cboReligion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2400
         Width           =   2745
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   2745
      End
      Begin VB.TextBox txtPOB 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   11
         Top             =   4560
         Width           =   2745
      End
      Begin VB.TextBox txtmContact 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   12
         Top             =   600
         Width           =   2745
      End
      Begin RichTextLib.RichTextBox rtbInfo 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   42
         Top             =   360
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   8493
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"AddStudentFrm.frx":00D6
      End
      Begin VB.PictureBox Container1 
         Height          =   4815
         Left            =   -75000
         ScaleHeight     =   4755
         ScaleWidth      =   9675
         TabIndex        =   47
         Top             =   360
         Width           =   9735
         Begin VB.Frame Frame3 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Name of brother(s) and sister(s)"
            ForeColor       =   &H00C25418&
            Height          =   2295
            Left            =   4560
            TabIndex        =   85
            Top             =   2040
            Width           =   4305
            Begin VB.TextBox txtBroSis 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1875
               Left            =   240
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Top             =   240
               Width           =   3825
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Guardian"
            ForeColor       =   &H00C25418&
            Height          =   1875
            Left            =   4560
            TabIndex        =   55
            Top             =   120
            Width           =   4305
            Begin VB.TextBox txtGAddress 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   705
               Left            =   1410
               MaxLength       =   70
               TabIndex        =   30
               Top             =   990
               Width           =   2745
            End
            Begin VB.TextBox txtGContact 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   20
               TabIndex        =   29
               Top             =   630
               Width           =   2745
            End
            Begin VB.TextBox txtGuardianName 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   50
               TabIndex        =   28
               Top             =   270
               Width           =   2745
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               Height          =   195
               Left            =   210
               TabIndex        =   58
               Top             =   1020
               Width           =   585
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Number"
               Height          =   195
               Left            =   210
               TabIndex        =   57
               Top             =   660
               Width           =   1170
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               Height          =   195
               Left            =   210
               TabIndex        =   56
               Top             =   330
               Width           =   405
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Parents"
            ForeColor       =   &H00C25418&
            Height          =   4215
            Left            =   120
            TabIndex        =   48
            Top             =   120
            Width           =   4305
            Begin VB.TextBox txtOrg 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   20
               TabIndex        =   27
               Top             =   3750
               Width           =   2745
            End
            Begin VB.TextBox txtPNo 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   20
               TabIndex        =   26
               Top             =   3390
               Width           =   2745
            End
            Begin VB.TextBox txtFOAge 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   50
               TabIndex        =   23
               Top             =   2070
               Width           =   2745
            End
            Begin VB.TextBox txtMOAge 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   50
               TabIndex        =   20
               Top             =   990
               Width           =   2745
            End
            Begin VB.TextBox txtPContact 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   20
               TabIndex        =   24
               Top             =   2430
               Width           =   2745
            End
            Begin VB.TextBox txtPAddress 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   1410
               MaxLength       =   70
               TabIndex        =   25
               Top             =   2790
               Width           =   2745
            End
            Begin VB.TextBox txtFOccupation 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   20
               TabIndex        =   22
               Top             =   1710
               Width           =   2745
            End
            Begin VB.TextBox txtFatherName 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   50
               TabIndex        =   21
               Top             =   1350
               Width           =   2745
            End
            Begin VB.TextBox txtMOccupation 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   50
               TabIndex        =   19
               Top             =   630
               Width           =   2745
            End
            Begin VB.TextBox txtMotherName 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   50
               TabIndex        =   18
               Top             =   270
               Width           =   2745
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Organization"
               Height          =   195
               Left            =   180
               TabIndex        =   87
               Top             =   3780
               Width           =   885
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Precint No."
               Height          =   195
               Left            =   180
               TabIndex        =   86
               Top             =   3420
               Width           =   795
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Age"
               Height          =   195
               Left            =   180
               TabIndex        =   81
               Top             =   2100
               Width           =   285
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Age"
               Height          =   195
               Left            =   180
               TabIndex        =   80
               Top             =   1020
               Width           =   285
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Number"
               Height          =   195
               Left            =   180
               TabIndex        =   54
               Top             =   2490
               Width           =   1170
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               Height          =   195
               Left            =   180
               TabIndex        =   53
               Top             =   2760
               Width           =   585
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Occupation"
               Height          =   195
               Left            =   180
               TabIndex        =   52
               Top             =   1770
               Width           =   915
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Father's Name"
               Height          =   195
               Left            =   180
               TabIndex        =   51
               Top             =   1410
               Width           =   1035
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Occupation"
               Height          =   195
               Left            =   180
               TabIndex        =   50
               Top             =   690
               Width           =   825
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mother's Name"
               Height          =   195
               Left            =   180
               TabIndex        =   49
               Top             =   330
               Width           =   1065
            End
         End
      End
      Begin VB.PictureBox Container3 
         Height          =   4815
         Left            =   -75000
         ScaleHeight     =   4755
         ScaleWidth      =   9675
         TabIndex        =   59
         Top             =   360
         Width           =   9735
         Begin VB.Frame Frame4 
            BackColor       =   &H00F6F8F8&
            Caption         =   "School Info"
            ForeColor       =   &H00C25418&
            Height          =   2595
            Left            =   120
            TabIndex        =   60
            Top             =   120
            Width           =   6105
            Begin VB.TextBox txtGenAve 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   20
               TabIndex        =   39
               Top             =   2100
               Width           =   2745
            End
            Begin VB.TextBox txtSY 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   20
               TabIndex        =   38
               Top             =   1740
               Width           =   2745
            End
            Begin VB.TextBox txtSchoolName 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   50
               TabIndex        =   35
               Top             =   270
               Width           =   4545
            End
            Begin VB.TextBox txtSchoolContact 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1410
               MaxLength       =   20
               TabIndex        =   36
               Top             =   630
               Width           =   2745
            End
            Begin VB.TextBox txtSchoolAdd 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   705
               Left            =   1410
               MaxLength       =   70
               TabIndex        =   37
               Top             =   990
               Width           =   2745
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "General Ave."
               Height          =   195
               Left            =   180
               TabIndex        =   83
               Top             =   2130
               Width           =   930
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "School Year"
               Height          =   195
               Left            =   180
               TabIndex        =   82
               Top             =   1770
               Width           =   870
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "School Name"
               Height          =   195
               Left            =   210
               TabIndex        =   63
               Top             =   330
               Width           =   960
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Number"
               Height          =   195
               Left            =   210
               TabIndex        =   62
               Top             =   660
               Width           =   1170
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               Height          =   195
               Left            =   210
               TabIndex        =   61
               Top             =   1020
               Width           =   585
            End
         End
      End
      Begin VB.PictureBox jcbutton1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   16
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox cmdNext1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   17
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox jcbutton2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66480
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   34
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox cmdBack1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67680
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   33
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox jcbutton3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   32
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox jcbutton4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66480
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   43
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox jcbutton6 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67680
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   41
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox jcbutton5 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   40
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox jcbutton7 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66480
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   46
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox jcbutton9 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67680
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   45
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox jcbutton8 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   44
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox Container4 
         Height          =   2205
         Left            =   5400
         ScaleHeight     =   2145
         ScaleWidth      =   2745
         TabIndex        =   65
         Top             =   1800
         Width           =   2805
         Begin VB.Image Image2 
            DataField       =   "PHOTO"
            DataSource      =   "Data1"
            Height          =   2000
            Left            =   120
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2550
         End
      End
      Begin MSComCtl2.DTPicker txtDOB 
         Height          =   330
         Left            =   1320
         TabIndex        =   6
         Top             =   2760
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16515073
         CurrentDate     =   36892
         MinDate         =   2
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Picture path"
         Height          =   195
         Left            =   4320
         TabIndex        =   84
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Type"
         Height          =   195
         Left            =   240
         TabIndex        =   79
         Top             =   4200
         Width           =   795
      End
      Begin VB.Label Religion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         Height          =   195
         Left            =   240
         TabIndex        =   78
         Top             =   2400
         Width           =   555
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         Height          =   195
         Left            =   240
         TabIndex        =   77
         Top             =   2760
         Width           =   720
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship"
         Height          =   195
         Left            =   240
         TabIndex        =   76
         Top             =   3480
         Width           =   765
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Left            =   240
         TabIndex        =   75
         Top             =   3840
         Width           =   465
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         Height          =   195
         Left            =   240
         TabIndex        =   74
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   195
         Left            =   240
         TabIndex        =   73
         Top             =   3120
         Width           =   285
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   195
         Left            =   240
         TabIndex        =   72
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         Height          =   195
         Left            =   240
         TabIndex        =   71
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   195
         Left            =   240
         TabIndex        =   70
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   195
         Left            =   4320
         TabIndex        =   69
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Place Of Birth"
         Height          =   195
         Left            =   240
         TabIndex        =   68
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg No"
         Height          =   195
         Left            =   240
         TabIndex        =   67
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Left            =   4320
         TabIndex        =   66
         Top             =   960
         Width           =   570
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   1200
         Top             =   -330
         Width           =   4695
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00F4F9FA&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   5550
         Left            =   -74955
         Top             =   345
         Width           =   9720
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00F4F9FA&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   5550
         Left            =   -74955
         Top             =   345
         Width           =   9720
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00F4F9FA&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   5550
         Left            =   -74955
         Top             =   345
         Width           =   9720
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00F4F9FA&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000013&
         BorderWidth     =   2
         Height          =   4575
         Left            =   120
         Top             =   480
         Width           =   9600
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00F4F9FA&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   5550
         Left            =   45
         Top             =   345
         Width           =   9720
      End
   End
   Begin VB.PictureBox jcFrames2 
      FillColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   10035
      TabIndex        =   88
      Top             =   0
      Width           =   10095
   End
   Begin VB.Image Image1 
      DataField       =   "PHOTO"
      DataSource      =   "Data1"
      Height          =   1530
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1890
   End
End
Attribute VB_Name = "StudentFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim clsData As New clsStudents
Private Sub chkTransferee_Click()
If chkTransferee.Value = 0 Then
cmbTransfereeYL.Enabled = False
Else
cmbTransfereeYL.Enabled = True
End If
End Sub

Private Sub cmdBack1_Click()
SSTab_AddStudent.Tab = 0
End Sub

Private Sub cmdNext1_Click()
SSTab_AddStudent.Tab = 1
txtMotherName.SetFocus
End Sub
Private Sub Command1_Click()
On Error Resume Next
Dim extractedpath, thepic As String
Dim lentxt As Integer
Dim imge(1) As IPictureDisp
'On Error Resume Next
With cd
    .DialogTitle = "Browse NONESCOST MPC Member's Pictures"
    .Filter = "JPEG Files(*.jpg)|*.jpg|BMP Files(*.bmp)|*.bmp"
    .InitDir = "c:\"
    .ShowOpen
    Set imge(1) = LoadPicture(cd.FileName)
    Image2.Picture = imge(1)
    If .FileName = "" Then
        Image2.Picture = Image2.Picture
        'Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.Width, Picture1.Height
    Else
    Text1.Text = (cd.FileName)
    thepic = (cd.FileName)
    lentxt = Len(Text1.Text)
    'lentxt = Len(Text1.Text)
    'extractedpath = Mid(Text1.Text, 12, lentxt) '- 12 + 1)
    extractedpath = Text1.Text
    Text1.Text = extractedpath
    Image2.Picture = LoadPicture(thepic)
     
   ' Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.Width, Picture1.Height
    End If
End With
SSTab_AddStudent.Tab = 1
SSTab_AddStudent.Tab = 0
End Sub
Private Sub Form_Load()
'txtID = clsData.GetID
txtyear = Format(Now, "yyyy")
txtCombine = txtyear & "-" & txtID
SSTab_AddStudent.Tab = 0
End Sub



Private Sub jcbutton1_Click()
Unload Me
End Sub

Private Sub jcbutton10_Click()
PreviewSuumary
End Sub

Private Sub jcbutton2_Click()
If ValidateStudentInfo = False Then
SSTab_AddStudent.Tab = 2
txtSchoolName.SetFocus
End If
'SSTab_AddStudent.Tab = 1
End Sub
Private Sub jcbutton3_Click()
Unload Me
End Sub
Private Sub jcbutton4_Click()
If ValidateParentsInfo = False Then
SSTab_AddStudent.Tab = 3
Timer1.Enabled = True
Timer1.Enabled = False
Else
SSTab_AddStudent.Tab = 1
End If
'SSTab_AddStudent.Tab = 2
End Sub

Private Sub jcbutton5_Click()
Unload Me
End Sub

Private Sub jcbutton6_Click()
SSTab_AddStudent.Tab = 1
End Sub

Private Sub jcbutton7_Click()
'On Error Resume Next
'Dim FS As New FileSystemObject
If jcbutton7.Caption = "&Save" And ValidateStudentInfo = False And ValidateParentsInfo = False Then
clsData.AddStudents txtID, txtLname, txtFname, txtMname, cboSex, txtDOB, txtAge, cboReligion, txtCitizenship, cboStatus, cboBloodType, txtPOB, txtmContact, txtMotherName, txtMOccupation, txtFatherName, txtFOccupation, txtPContact, txtPAddress, txtGuardianName, txtGContact, txtGAddress, txtSchoolName, txtSchoolAdd, txtSchoolContact, Text1.Text, txtMOAge, txtFOAge, txtSY, txtGenAve, txtAdd, txtBroSis, txtPNo, txtOrg
'FS.CopyFile Text1.Text, "D:\Documents and Settings\admin\My Documents", False
'Set FS = Nothing
Unload Me
StudentFrm.Show
ElseIf jcbutton7.Caption = "&Update" And ValidateStudentInfo = False And ValidateParentsInfo = False Then
clsData.UpdateStudents txtID, txtLname, txtFname, txtMname, cboSex, txtDOB, txtAge, cboReligion, txtCitizenship, cboStatus, cboBloodType, txtPOB, txtmContact, txtMotherName, txtMOccupation, txtFatherName, txtFOccupation, txtPContact, txtPAddress, txtGuardianName, txtGContact, txtGAddress, txtSchoolName, txtSchoolAdd, txtSchoolContact, Text1.Text, txtID.Text, txtMOAge, txtFOAge, txtSY, txtGenAve, txtAdd, txtBroSis, txtPNo, txtOrg
Unload Me
StudentListFrm.Show
End If
End Sub

Private Sub jcbutton8_Click()
Unload Me
End Sub

Private Sub jcbutton9_Click()
SSTab_AddStudent.Tab = 2
txtMotherName.SetFocus
End Sub

Function ValidateStudentInfo() As Boolean
ValidateStudentInfo = True
    If Trim(txtLname.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtLname.SetFocus
    Exit Function
    ElseIf Trim(txtFname.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtFname.SetFocus
    Exit Function
    ElseIf Trim(txtMname.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtMname.SetFocus
    Exit Function
    ElseIf Trim(cboSex.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    cboSex.SetFocus
    Exit Function
    ElseIf Trim(cboReligion.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    cboReligion.SetFocus
    Exit Function
    ElseIf Trim(txtCitizenship.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtCitizenship.SetFocus
    Exit Function
    ElseIf Trim(cboStatus.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    cboStatus.SetFocus
    Exit Function
    ElseIf Trim(cboBloodType.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    cboBloodType.SetFocus
    Exit Function
    ElseIf Trim(txtPOB.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtPOB.SetFocus
    Exit Function
    ElseIf Trim(txtmContact.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtmContact.SetFocus
    Exit Function
    ElseIf Trim(txtAdd.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtAdd.SetFocus
    Exit Function
    End If
ValidateStudentInfo = False
End Function
Function ValidateParentsInfo() As Boolean
ValidateParentsInfo = True
    If Trim(txtMotherName.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtMotherName.SetFocus
    Exit Function
    ElseIf Trim(txtMOccupation.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtMOccupation.SetFocus
    Exit Function
    ElseIf Trim(txtFatherName.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtFatherName.SetFocus
    Exit Function
    ElseIf Trim(txtFOccupation.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtFOccupation.SetFocus
    Exit Function
    ElseIf Trim(txtPContact.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtPContact.SetFocus
    Exit Function
    ElseIf Trim(txtPAddress.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtPAddress.SetFocus
    Exit Function
    ElseIf Trim(txtGuardianName.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtGuardianName.SetFocus
    Exit Function
    ElseIf Trim(txtGContact.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtGContact.SetFocus
    Exit Function
    ElseIf Trim(txtGAddress.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtGAddress.SetFocus
    Exit Function
    ElseIf Trim(txtSchoolName.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtSchoolName.SetFocus
    Exit Function
    ElseIf Trim(txtSchoolAdd.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtSchoolAdd.SetFocus
    Exit Function
    ElseIf Trim(txtSchoolContact.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtSchoolContact.SetFocus
    Exit Function
    ElseIf Trim(txtMOAge.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtMOAge.SetFocus
    Exit Function
    ElseIf Trim(txtFOAge.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtFOAge.SetFocus
    Exit Function
    ElseIf Trim(txtPNo.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtPNo.SetFocus
    Exit Function
    ElseIf Trim(txtOrg.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtOrg.SetFocus
    Exit Function
    ElseIf Trim(txtBroSis.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtBroSis.SetFocus
    Exit Function
    ElseIf Trim(txtSY.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtSY.SetFocus
    Exit Function
    ElseIf Trim(txtGenAve.Text) = "" Then
    MsgBox "Don't leave the field empty.", vbCritical, ""
    txtGenAve.SetFocus
    Exit Function
    End If
ValidateParentsInfo = False
End Function
Private Sub SSTab_AddStudent_Click(PreviousTab As Integer)
Select Case SSTab_AddStudent.Caption
    Case "Parents Information"
    cmdNext1_Click
    'SSTab_AddStudent.Tab = 1
    Case "Summary"
    jcbutton4_Click
    Case "School Information"
    jcbutton2_Click
    'SSTab_AddStudent.Tab = 2
    End Select
End Sub

Private Sub Timer1_Timer()
PreviewSuumary
End Sub

Private Sub txtDOB_Change()
On Error Resume Next
    txtAge.Text = (Now - txtDOB.Value) \ 365
End Sub



Private Sub txtFOAge_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Private Sub txtGContact_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub
Private Sub txtGenAve_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub
Private Sub txtmContact_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub





Private Sub txtMOAge_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Private Sub txtPContact_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub



Private Sub txtSContact_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub

Public Sub PreviewSuumary()
Dim i As Integer
    Dim selLength As Integer
    Dim selStart As Integer
    Dim smFound As Boolean
    Dim fn As Boolean

rtbInfo.Text = ""
    
    rtbInfo.Text = rtbInfo.Text & _
    "Student ID Number: " & Me.txtID.Text & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Name: " & Me.txtLname.Text & ", " & Me.txtFname.Text & " " & Me.txtMname.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Gender: " & Me.cboSex.Text & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Religion: " & Me.cboReligion.Text & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Date of Birth: " & Me.txtDOB.Value & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Age: " & Me.txtAge.Text & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Citizenship: " & Me.txtCitizenship.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Status: " & Me.cboStatus.Text & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Blood Type: " & Me.cboBloodType.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Place of Birth: " & Me.txtPOB.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Contact: " & Me.txtmContact.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Address: " & Me.txtAdd.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    vbNewLine
        
    rtbInfo.Text = rtbInfo.Text & _
    vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Parents Information" & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Mother's Name: " & Me.txtMotherName.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Mother's Occupation: " & Me.txtMOccupation.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Mother's Age: " & Me.txtMOAge.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Father's Name: " & Me.txtFatherName.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Father's Occupation: " & Me.txtFOccupation.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Father's Age: " & Me.txtFOAge.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Contact: " & Me.txtPContact.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Address: " & Me.txtPAddress.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Precint No: " & Me.txtPNo.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Organization: " & Me.txtOrg.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Guardian's Name: " & Me.txtGuardianName.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Contact: " & Me.txtGContact.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Address: " & Me.txtGAddress.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "School Information" & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Last School Attended: " & Me.txtSchoolName.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Contact: " & Me.txtSchoolContact.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Address: " & Me.txtSchoolAdd.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "School Year: " & Me.txtSY.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "General Ave.: " & Me.txtGenAve.Text & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    vbNewLine
    
     'set color
    rtbInfo.selStart = 0
    rtbInfo.selLength = Len(rtbInfo.Text)
    rtbInfo.SelColor = &H584620
    rtbInfo.SelBold = False
    
    For i = 1 To Len(rtbInfo.Text) + 1
    
        If Mid(rtbInfo.Text, i, 1) = ":" Then
            smFound = True
            selStart = i
            selLength = 0
        End If
        
        If smFound = True Then
            selLength = selLength + 1

            If Mid(rtbInfo.Text, i, 2) = vbNewLine Then
                
                rtbInfo.selStart = selStart
                rtbInfo.selLength = selLength
                rtbInfo.SelFontSize = 10
                rtbInfo.SelColor = &H0&
                rtbInfo.SelBold = True
                
                If fn = False Then
                    rtbInfo.SelFontSize = 10
                    fn = True
                End If
                
                rtbInfo.selLength = 0
                
                smFound = False
            End If
            
        End If
        

    Next
End Sub
