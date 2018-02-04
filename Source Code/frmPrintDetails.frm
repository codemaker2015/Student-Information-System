VERSION 5.00
Begin VB.Form frmPrintDetails 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Student Information System"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11775
   Icon            =   "frmPrintDetails.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.Label lblChance 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   9960
      TabIndex        =   82
      Top             =   8340
      Width           =   615
   End
   Begin VB.Label lblChance 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Chance"
      Height          =   255
      Index           =   1
      Left            =   9960
      TabIndex        =   81
      Top             =   7965
      Width           =   615
   End
   Begin VB.Label lblChance 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Chance"
      Height          =   255
      Index           =   0
      Left            =   9960
      TabIndex        =   80
      Top             =   7605
      Width           =   615
   End
   Begin VB.Label lblPercent 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   9120
      TabIndex        =   79
      Top             =   8340
      Width           =   615
   End
   Begin VB.Label lblPercent 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   1
      Left            =   9120
      TabIndex        =   78
      Top             =   7965
      Width           =   615
   End
   Begin VB.Label lblPercent 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   0
      Left            =   9120
      TabIndex        =   77
      Top             =   7605
      Width           =   615
   End
   Begin VB.Label lblGrade 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   8160
      TabIndex        =   76
      Top             =   8340
      Width           =   615
   End
   Begin VB.Label lblGrade 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Grade"
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   75
      Top             =   7965
      Width           =   615
   End
   Begin VB.Label lblGrade 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Grade"
      Height          =   255
      Index           =   0
      Left            =   8160
      TabIndex        =   74
      Top             =   7605
      Width           =   615
   End
   Begin VB.Label lblInstitution 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   73
      Top             =   8340
      Width           =   4695
   End
   Begin VB.Label lblInstitution 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Institution name"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   72
      Top             =   7965
      Width           =   4695
   End
   Begin VB.Label lblInstitution 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Institution name"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   71
      Top             =   7605
      Width           =   4695
   End
   Begin VB.Label lblYear 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   70
      Top             =   8340
      Width           =   855
   End
   Begin VB.Label lblYear 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   69
      Top             =   7965
      Width           =   855
   End
   Begin VB.Label lblYear 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   68
      Top             =   7605
      Width           =   855
   End
   Begin VB.Label lblCourse 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   67
      Top             =   8340
      Width           =   1215
   End
   Begin VB.Label lblCourse 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   66
      Top             =   7965
      Width           =   1215
   End
   Begin VB.Label lblCourse 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   65
      Top             =   7605
      Width           =   1215
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   8040
      X2              =   9840
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   9000
      X2              =   9000
      Y1              =   7320
      Y2              =   8640
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   9840
      X2              =   9840
      Y1              =   7080
      Y2              =   8640
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   6
      Left            =   9300
      TabIndex        =   64
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Grade"
      Height          =   255
      Index           =   5
      Left            =   8280
      TabIndex        =   63
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Chances"
      Height          =   495
      Index           =   4
      Left            =   9960
      TabIndex        =   62
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Marks"
      Height          =   255
      Index           =   3
      Left            =   8760
      TabIndex        =   61
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of the Institution"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   60
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Year of Passing"
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   59
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Course"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   58
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name and Signature of the Guardian"
      Height          =   195
      Left            =   7560
      TabIndex        =   57
      Top             =   10080
      Width           =   2595
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Signature of the Student"
      Height          =   195
      Left            =   960
      TabIndex        =   56
      Top             =   10080
      Width           =   1725
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3720
      TabIndex        =   55
      Top             =   9480
      Width           =   4575
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15. Suggestion by the Parent/Guardian"
      Height          =   195
      Left            =   600
      TabIndex        =   54
      Top             =   9480
      Width           =   2775
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3720
      TabIndex        =   53
      Top             =   9120
      Width           =   4575
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "14. Details of Scholarship/Stipend"
      Height          =   195
      Left            =   600
      TabIndex        =   52
      Top             =   9120
      Width           =   2415
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3720
      TabIndex        =   51
      Top             =   8760
      Width           =   4575
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "13. Co-curricular Activities"
      Height          =   195
      Left            =   600
      TabIndex        =   50
      Top             =   8760
      Width           =   1830
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   600
      X2              =   10680
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   600
      X2              =   10680
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Index           =   1
      Left            =   600
      Top             =   7080
      Width           =   10095
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   600
      X2              =   10680
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   2040
      X2              =   2040
      Y1              =   7080
      Y2              =   8640
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   2880
      X2              =   2880
      Y1              =   7080
      Y2              =   8640
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   8040
      X2              =   8040
      Y1              =   7080
      Y2              =   8640
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12. Details of Examination passed"
      Height          =   195
      Left            =   600
      TabIndex        =   49
      Top             =   6840
      Width           =   2385
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "About Sisters and brothers"
      Height          =   1455
      Left            =   8520
      TabIndex        =   48
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   615
      Index           =   2
      Left            =   5880
      TabIndex        =   47
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   615
      Index           =   1
      Left            =   3240
      TabIndex        =   46
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   615
      Index           =   0
      Left            =   720
      TabIndex        =   45
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   6930
      TabIndex        =   44
      Top             =   5310
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
      Height          =   255
      Index           =   1
      Left            =   4290
      TabIndex        =   43
      Top             =   5310
      Width           =   1365
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
      Height          =   255
      Index           =   0
      Left            =   1695
      TabIndex        =   42
      Top             =   5310
      Width           =   1365
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   41
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   40
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   1245
      TabIndex        =   39
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address with Phone No"
      Height          =   195
      Index           =   2
      Left            =   5880
      TabIndex        =   38
      Top             =   5580
      Width           =   1665
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address with Phone No"
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   37
      Top             =   5580
      Width           =   1665
   End
   Begin VB.Label lblOccupation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
      Height          =   195
      Index           =   2
      Left            =   5880
      TabIndex        =   36
      Top             =   5310
      Width           =   825
   End
   Begin VB.Label lblOccupation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   35
      Top             =   5310
      Width           =   825
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Index           =   2
      Left            =   5880
      TabIndex        =   34
      Top             =   5040
      Width           =   420
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   33
      Top             =   5040
      Width           =   420
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address with Phone No"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   32
      Top             =   5580
      Width           =   1665
   End
   Begin VB.Label lblOccupation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   31
      Top             =   5310
      Width           =   825
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   30
      Top             =   5040
      Width           =   420
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sister(s) and Brother(s)"
      Height          =   195
      Index           =   3
      Left            =   8760
      TabIndex        =   29
      Top             =   4725
      Width           =   1590
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian"
      Height          =   195
      Index           =   2
      Left            =   6720
      TabIndex        =   28
      Top             =   4725
      Width           =   645
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mother"
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   27
      Top             =   4725
      Width           =   495
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   26
      Top             =   4725
      Width           =   450
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   8400
      X2              =   8400
      Y1              =   4680
      Y2              =   6600
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   5760
      X2              =   5760
      Y1              =   4680
      Y2              =   6600
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   3120
      X2              =   3120
      Y1              =   4680
      Y2              =   6600
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   600
      X2              =   10680
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   0
      Left            =   600
      Top             =   4680
      Width           =   10095
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "11. Information about Family"
      Height          =   195
      Left            =   600
      TabIndex        =   25
      Top             =   4440
      Width           =   1980
   End
   Begin VB.Label lblPhoto 
      BackStyle       =   0  'Transparent
      Caption         =   "Photo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   24
      Top             =   960
      Width           =   855
   End
   Begin VB.Image imgStudentPicture 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   9480
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image imgExit 
      Height          =   480
      Left            =   6000
      Picture         =   "frmPrintDetails.frx":000C
      Top             =   9720
      Width           =   480
   End
   Begin VB.Image imgPrint 
      Height          =   480
      Left            =   4800
      Picture         =   "frmPrintDetails.frx":0CD6
      Top             =   9720
      Width           =   480
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Year of study"
      Height          =   195
      Index           =   10
      Left            =   7440
      TabIndex        =   23
      Top             =   3720
      Width           =   930
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cast"
      Height          =   195
      Index           =   9
      Left            =   7440
      TabIndex        =   22
      Top             =   3360
      Width           =   315
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Reg No"
      Height          =   195
      Index           =   8
      Left            =   7440
      TabIndex        =   21
      Top             =   1920
      Width           =   555
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      Height          =   195
      Index           =   7
      Left            =   7440
      TabIndex        =   20
      Top             =   1560
      Width           =   270
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Income"
      Height          =   195
      Index           =   6
      Left            =   2760
      TabIndex        =   19
      Top             =   4080
      Width           =   525
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Blood"
      Height          =   195
      Index           =   5
      Left            =   2760
      TabIndex        =   18
      Top             =   3720
      Width           =   405
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
      Height          =   195
      Index           =   4
      Left            =   2760
      TabIndex        =   17
      Top             =   3360
      Width           =   570
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      Height          =   195
      Index           =   3
      Left            =   7440
      TabIndex        =   16
      Top             =   2280
      Width           =   465
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Permanent Address"
      Height          =   975
      Index           =   2
      Left            =   2760
      TabIndex        =   15
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Programme"
      Height          =   195
      Index           =   1
      Left            =   2760
      TabIndex        =   14
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   13
      Top             =   1560
      Width           =   420
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "10. Anual Family Income"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   12
      Top             =   4080
      Width           =   1725
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "8. Blood Group"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   11
      Top             =   3720
      Width           =   1065
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "     Caste *"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6000
      TabIndex        =   10
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "7. Religion *"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "6. Phone **"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6000
      TabIndex        =   8
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5. Permanent Address  *"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   2280
      Width           =   1710
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "4. Reg No **"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6000
      TabIndex        =   6
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "3. Programme *"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "9. Year of Study *"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6000
      TabIndex        =   4
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "2. Sex"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6000
      TabIndex        =   3
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1. Name *"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   8175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAddMark 
         Caption         =   "Add Mark"
      End
      Begin VB.Menu mnuExtra 
         Caption         =   "Add Extra-Curricular Activities"
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
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmPrintDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  
  imgStudentPicture.Picture = LoadPicture(frmAddStudentInfo.CommonDialog1.FileName)
  If imgStudentPicture.Picture <> (None) Then
    lblPhoto.Visible = False
  End If
  Label14(0).Caption = frmAddStudentInfo.txtName.Text
  Label14(1).Caption = frmAddStudentInfo.cmbProgramme.Text
  Label14(2).Caption = frmAddStudentInfo.txtAddress.Text
  Label14(3).Caption = frmAddStudentInfo.txtPhone(0).Text
  Label14(4).Caption = frmAddStudentInfo.cmbReligion.Text
  Label14(5).Caption = frmAddStudentInfo.cmbBloodGroup.Text
  Label14(6).Caption = frmAddStudentInfo.cmbIncome.Text
  Label14(7).Caption = frmAddStudentInfo.cmbGender.Text
  Label14(8).Caption = frmAddStudentInfo.txtRegNo(0).Text
  Label14(9).Caption = frmAddStudentInfo.cmbCast.Text
  Label14(10).Caption = frmAddStudentInfo.cmbYearOfStudy.Text
  
  Label17(0).Caption = frmAddStudentInfo.txtFatherName.Text
  Label18(0).Caption = frmAddStudentInfo.cmbFOccupation.Text
  Label19(0).Caption = frmAddStudentInfo.txtFAddress.Text + "     " + "Ph:" + frmAddStudentInfo.txtPhone(1).Text
  Label17(1).Caption = frmAddStudentInfo.txtMotherName.Text
  Label18(1).Caption = frmAddStudentInfo.cmbMOccupation.Text
  Label19(1).Caption = frmAddStudentInfo.txtFAddress.Text + "     " + "Ph:"
  Label17(2).Caption = frmAddStudentInfo.txtGName.Text
  Label19(2).Caption = frmAddStudentInfo.txtGAddress.Text + "     " + "Ph:" + frmAddStudentInfo.txtPhone(2).Text
  Label20.Caption = frmAddStudentInfo.txtBroSis.Text
  
  lblCourse(0).Caption = frmAddStudentInfo.lstCourse(0).Text
  lblCourse(1).Caption = frmAddStudentInfo.lstCourse(1).Text
  If lblYear(2).Caption <> "" Or lblYear(2).Caption <> "Year" Then
     If lblCourse(2).Caption <> "" Then lblCourse(2).Caption = frmAddStudentInfo.lstCourse(2).Text
  End If
  lblYear(0).Caption = frmAddStudentInfo.cmbYearOfPassing(0).Text
  lblYear(1).Caption = frmAddStudentInfo.cmbYearOfPassing(1).Text
  If frmAddStudentInfo.cmbYearOfPassing(2).Text <> "" Or frmAddStudentInfo.cmbYearOfPassing(2).Text <> "Year" Then lblYear(2).Caption = frmAddStudentInfo.cmbYearOfPassing(2).Text
  lblInstitution(0).Caption = frmAddStudentInfo.cmbSchoolName(0).Text
  lblInstitution(1).Caption = frmAddStudentInfo.cmbSchoolName(1).Text
  If frmAddStudentInfo.cmbSchoolName(2).Text <> "" Or frmAddStudentInfo.cmbSchoolName(2).Text <> "--Select--" Then lblInstitution(2).Caption = frmAddStudentInfo.cmbSchoolName(2).Text
  lblGrade(0).Caption = frmAddStudentInfo.cmbGrade(0).Text
  lblGrade(1).Caption = frmAddStudentInfo.cmbGrade(1).Text
  If frmAddStudentInfo.cmbGrade(2).Text <> "" Then lblGrade(2).Caption = frmAddStudentInfo.cmbGrade(2).Text
  lblPercent(0).Caption = frmAddStudentInfo.cmbPercent(0).Text
  lblPercent(1).Caption = frmAddStudentInfo.cmbPercent(1).Text
  If frmAddStudentInfo.cmbPercent(2).Text <> "" Then lblPercent(2).Caption = frmAddStudentInfo.cmbPercent(2).Text
  lblChance(0).Caption = frmAddStudentInfo.lstNoOfChances(0).Text
  lblChance(1).Caption = frmAddStudentInfo.lstNoOfChances(1).Text
  If frmAddStudentInfo.lstNoOfChances(2).Text <> "" Then lblChance(2).Caption = frmAddStudentInfo.lstNoOfChances(2).Text
  
End Sub

Private Sub mnuAddMark_Click()
   frmAddMark.txtRegNo.Text = frmAddStudentInfo.txtRegNo(1).Text
   frmAddMark.Show
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

Private Sub mnuExtra_Click()
  frmPhysical.Show
End Sub

Private Sub mnuPrint_Click()
  imgPrint_Click
End Sub

Private Sub imgPrint_Click()
    On Error Resume Next
    
    imgPrint.Visible = False
    imgExit.Visible = False
    
    PrintForm
    
    imgPrint.Visible = True
    imgExit.Visible = True
    
End Sub

Private Sub imgExit_Click()
    Unload Me
End Sub
