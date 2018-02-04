VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSearch 
   BackColor       =   &H80000005&
   Caption         =   "Search Engine"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15090
   Icon            =   "Searchfrm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Searchfrm.frx":000C
   ScaleHeight     =   9270
   ScaleWidth      =   15090
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   1800
      TabIndex        =   0
      Top             =   2520
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   -2147483643
      TabCaption(0)   =   "Search Engine 1"
      TabPicture(0)   =   "Searchfrm.frx":1C0C3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search Engine 2"
      TabPicture(1)   =   "Searchfrm.frx":1C0DF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Search Engine 3"
      TabPicture(2)   =   "Searchfrm.frx":1C0FB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H80000005&
         Caption         =   "Search Student Info"
         Height          =   4815
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   9135
         Begin VB.ComboBox cmbInfoType 
            Height          =   315
            ItemData        =   "Searchfrm.frx":1C117
            Left            =   600
            List            =   "Searchfrm.frx":1C127
            TabIndex        =   30
            Text            =   "--Select Information Type--"
            Top             =   3600
            Width           =   2775
         End
         Begin VB.ComboBox cmbStudent 
            Height          =   315
            ItemData        =   "Searchfrm.frx":1C157
            Left            =   600
            List            =   "Searchfrm.frx":1C159
            TabIndex        =   29
            Text            =   "--Select Student--"
            Top             =   3120
            Width           =   2775
         End
         Begin VB.TextBox txtName 
            Height          =   375
            Left            =   1440
            TabIndex        =   27
            Top             =   1080
            Width           =   1935
         End
         Begin VB.ComboBox cmbYear 
            Height          =   315
            ItemData        =   "Searchfrm.frx":1C15B
            Left            =   1440
            List            =   "Searchfrm.frx":1C1E9
            TabIndex        =   26
            Text            =   "--Select--"
            Top             =   1920
            Width           =   1935
         End
         Begin VB.ComboBox cmbCourse 
            Height          =   315
            ItemData        =   "Searchfrm.frx":1C301
            Left            =   1440
            List            =   "Searchfrm.frx":1C31A
            TabIndex        =   25
            Text            =   "--Select--"
            Top             =   1560
            Width           =   1935
         End
         Begin VB.TextBox txtRegNo 
            Height          =   375
            Left            =   1440
            TabIndex        =   20
            Top             =   600
            Width           =   1935
         End
         Begin VB.Image imgCancel 
            Height          =   375
            Left            =   3120
            Picture         =   "Searchfrm.frx":1C34F
            Top             =   4080
            Width           =   375
         End
         Begin VB.Image imgReport 
            Height          =   450
            Left            =   2280
            Picture         =   "Searchfrm.frx":1C82F
            Top             =   4080
            Width           =   375
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "No such  results found"
            Height          =   255
            Left            =   600
            TabIndex        =   28
            Top             =   2520
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label lblName 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   375
            Left            =   600
            TabIndex        =   24
            Top             =   1155
            Width           =   735
         End
         Begin VB.Label lblYear 
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            Height          =   375
            Left            =   600
            TabIndex        =   23
            Top             =   1965
            Width           =   735
         End
         Begin VB.Label lblCourse 
            BackStyle       =   0  'Transparent
            Caption         =   "Program"
            Height          =   375
            Left            =   600
            TabIndex        =   22
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblRegNo 
            BackColor       =   &H80000005&
            Caption         =   "Reg No"
            Height          =   375
            Left            =   600
            TabIndex        =   21
            Top             =   675
            Width           =   615
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000001&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   4455
            Index           =   2
            Left            =   120
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Caption         =   "Serach Anything"
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   9135
         Begin VB.TextBox txtSearch 
            Height          =   405
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   480
            Width           =   3495
         End
         Begin VB.Shape Shape2 
            Height          =   2415
            Left            =   4800
            Top             =   1320
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Physical Information"
            Height          =   375
            Index           =   4
            Left            =   5040
            TabIndex        =   43
            Top             =   3360
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Internal Mark Information"
            Height          =   375
            Index           =   3
            Left            =   5040
            TabIndex        =   42
            Top             =   2880
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Educational Information"
            Height          =   375
            Index           =   2
            Left            =   5040
            TabIndex        =   41
            Top             =   2400
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Family Information"
            Height          =   375
            Index           =   1
            Left            =   5040
            TabIndex        =   40
            Top             =   1920
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Personal Information"
            Height          =   375
            Index           =   0
            Left            =   5040
            TabIndex        =   39
            Top             =   1440
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Label four 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   34
            Top             =   4120
            Width           =   255
         End
         Begin VB.Label three 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   255
            Index           =   0
            Left            =   2120
            TabIndex        =   33
            Top             =   4120
            Width           =   255
         End
         Begin VB.Label two 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   32
            Top             =   4120
            Width           =   255
         End
         Begin VB.Label one 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   31
            Top             =   4120
            Width           =   255
         End
         Begin VB.Image imgSearch1 
            Height          =   240
            Left            =   3960
            Picture         =   "Searchfrm.frx":1CDE6
            Top             =   480
            Width           =   240
         End
         Begin VB.Image imgNext 
            Height          =   375
            Index           =   0
            Left            =   3000
            Picture         =   "Searchfrm.frx":1D0EE
            Top             =   3960
            Width           =   375
         End
         Begin VB.Image imgPrevious 
            Height          =   375
            Index           =   0
            Left            =   600
            Picture         =   "Searchfrm.frx":1D49F
            Top             =   3960
            Width           =   375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   9
            Top             =   3360
            Width           =   3500
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   8
            Top             =   3000
            Width           =   3500
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   7
            Top             =   2640
            Width           =   3500
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   6
            Top             =   2280
            Width           =   3500
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   5
            Top             =   1920
            Width           =   3500
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   4
            Top             =   1560
            Width           =   3500
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No such  results found"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   1080
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000001&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   4455
            Index           =   0
            Left            =   120
            Top             =   240
            Width           =   8895
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000005&
         Caption         =   "Search Questions"
         Height          =   4815
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   9135
         Begin VB.ComboBox cmbSearch 
            Height          =   315
            ItemData        =   "Searchfrm.frx":1D9B8
            Left            =   360
            List            =   "Searchfrm.frx":1D9BF
            TabIndex        =   11
            Text            =   "--Select Question--"
            Top             =   480
            Width           =   3500
         End
         Begin VB.Label four 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   38
            Top             =   4125
            Width           =   255
         End
         Begin VB.Label three 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   255
            Index           =   1
            Left            =   2115
            TabIndex        =   37
            Top             =   4125
            Width           =   255
         End
         Begin VB.Label two 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   36
            Top             =   4125
            Width           =   255
         End
         Begin VB.Label one 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   35
            Top             =   4125
            Width           =   255
         End
         Begin VB.Image imgPrevious 
            Height          =   375
            Index           =   1
            Left            =   600
            Picture         =   "Searchfrm.frx":1D9DB
            Top             =   3960
            Width           =   375
         End
         Begin VB.Image imgNext 
            Height          =   375
            Index           =   1
            Left            =   3000
            Picture         =   "Searchfrm.frx":1DEF4
            Top             =   3960
            Width           =   375
         End
         Begin VB.Image imgSearch2 
            Height          =   240
            Left            =   3960
            Picture         =   "Searchfrm.frx":1E2A5
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No such  results found"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   19
            Top             =   1080
            Visible         =   0   'False
            Width           =   3500
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   11
            Left            =   360
            TabIndex        =   18
            Top             =   1560
            Visible         =   0   'False
            Width           =   3500
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   10
            Left            =   360
            TabIndex        =   17
            Top             =   1920
            Visible         =   0   'False
            Width           =   3500
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   16
            Top             =   2280
            Visible         =   0   'False
            Width           =   3500
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   15
            Top             =   2640
            Visible         =   0   'False
            Width           =   3500
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   14
            Top             =   3000
            Visible         =   0   'False
            Width           =   3500
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   13
            Top             =   3360
            Visible         =   0   'False
            Width           =   3500
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000001&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   4455
            Index           =   1
            Left            =   120
            Top             =   240
            Width           =   4335
         End
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private regno As String

Private Sub Form_Load()
   Theme frmSearch
   
   regno = 0
   connection
   reccheck
   
   rec.Open "select QUESTION from SEARCHENGINE", con, adOpenDynamic, adLockPessimistic
   cmbSearch.Clear
   While Not rec.EOF
     cmbSearch.AddItem (rec.Fields(0))
     rec.MoveNext
   Wend
End Sub

Private Sub Form_GotFocus()
  Unload frmLoading
End Sub

Private Sub imgCancel_Click()
  Unload Me
End Sub

Private Sub imgNext_Click(Index As Integer)
   If three(Index).FontBold = True Then
      three(Index).FontBold = False
      four(Index).FontBold = True
   End If
   If two(Index).FontBold = True Then
      two(Index).FontBold = False
      three(Index).FontBold = True
   End If
   If one(Index).FontBold = True Then
      one(Index).FontBold = False
      two(Index).FontBold = True
   End If
End Sub

Private Sub imgPrevious_Click(Index As Integer)
   If two(Index).FontBold = True Then
      two(Index).FontBold = False
      one(Index).FontBold = True
   End If
   If three(Index).FontBold = True Then
      three(Index).FontBold = False
      two(Index).FontBold = True
   End If
   If four(Index).FontBold = True Then
      four(Index).FontBold = False
      three(Index).FontBold = True
   End If
End Sub


Private Sub imgSearch1_Click()
   Dim length As Integer, i As Integer
   Dim search As String
   Dim temp(5) As String
   search = Trim(txtSearch.Text)
   length = Len(search)
   
   connection
   reccheck
   
   Shape2.Visible = False
   
   For i = 0 To 4
      Label3(i).Visible = False
   Next i
   
   If IsNumeric(search) = True Then
     'search for numeric results
     If length = 8 Or length = 9 Then
        rec.Open "select STUDENTNAME from MAINTABLE where REGNO = " & search, con, adOpenDynamic, adLockPessimistic
        If rec.EOF = False Then
           Label1(0).Visible = False
           Label2(0).Caption = rec.Fields(0)
           Label2(0).Visible = True
           Exit Sub
        Else
           Label1(0).Visible = True
           For i = 0 To 5
             Label2(i).Visible = False
           Next i
           Exit Sub
        End If
     End If
     
     reccheck
     
     If length > 5 And length < 9 Then
        rec.Open "select REGNO from EXAMTABLE where EXAMNO = " & search, con, adOpenDynamic, adLockPessimistic
        If rec.EOF = False Then
           temp(0) = rec.Fields(0)
           reccheck
           rec.Open "select STUDENTNAME from MAINTABLE where REGNO = " & temp(0), con, adOpenDynamic, adLockPessimistic
           If rec.EOF = False Then
             Label1(0).Visible = False
             Label2(0).Caption = rec.Fields(0)
             Label2(0).Visible = True
             Exit Sub
           Else
             Label1(0).Visible = True
             For i = 0 To 5
                Label2(i).Visible = False
             Next i
             Exit Sub
           End If
        Else
           Label1(0).Visible = True
           For i = 0 To 5
             Label2(i).Visible = False
           Next i
           Exit Sub
        End If
     End If
     
     reccheck
     
     If length > 9 Or length < 14 Then
        rec.Open "select STUDENTNAME from PERSONALTABLE where PH = " & search, con, adOpenDynamic, adLockPessimistic
        If rec.EOF = False Then
           Label1(0).Visible = False
           Label2(0).Caption = rec.Fields(0)
           Label2(0).Visible = True
           Exit Sub
        Else
           Label1(0).Visible = True
           For i = 0 To 5
             Label2(i).Visible = False
           Next i
           Exit Sub
        End If
     End If
     
   Else
      'search for alphabetic results
       Dim section As Integer
       Dim J As Integer
       section = 0
       For i = 1 To Len(search)
           If Mid(search, i, 1) Like "[a-z]" Or Mid(search, i, 1) = " " Or Mid(search, i, 1) = ":" Then
              'do nothing
           Else
              Label1(0).Visible = True
              For J = 0 To 5
                Label2(J).Caption = ""
              Next J
              Exit For
              Exit Sub
           End If
       Next i
       
       If length < 25 Then
         reccheck
         rec.Open "select REGNO from FAMILYTABLE  where FATHEROCCUPATION like '%" & search & "'", con, adOpenDynamic, adLockPessimistic
         If rec.EOF = False Then
            section = 1
            i = 0
            Do While Not rec.EOF
               temp(i) = rec.Fields(0)
               If i = 5 Then Exit Do
               i = i + 1
               rec.MoveNext
            Loop
               For J = 0 To i
                   reccheck
                   rec.Open "select STUDENTNAME from MAINTABLE where REGNO = '" & temp(J) & "'", con, adOpenDynamic, adLockPessimistic
                   If rec.EOF = False Then Label2(J).Caption = rec.Fields(0)
               Next J
         End If
         
         If section = 0 Then
            reccheck
            rec.Open "select REGNO from FAMILYTABLE  where MOTHEROCCUPATION like '%" & search & "'", con, adOpenDynamic, adLockPessimistic
            If rec.EOF = False Then
               i = 0
               section = 2
               Do While Not rec.EOF
                  temp(i) = rec.Fields(0)
                  If i = 5 Then Exit Do
                  i = i + 1
                  rec.MoveNext
               Loop
               For J = 0 To i
                   reccheck
                   rec.Open "select STUDENTNAME from MAINTABLE where REGNO = '" & temp(J) & "'", con, adOpenDynamic, adLockPessimistic
                   If rec.EOF = False Then Label2(J).Caption = rec.Fields(0)
               Next J
            End If
         End If
        
       End If
       
       If length < 40 Then
         If section = 0 Then
            reccheck
            rec.Open "select distinct STUDENTNAME from PERSONALTABLE  where STUDENTNAME like '%" & search & "'", con, adOpenDynamic, adLockPessimistic
            If rec.EOF = False Then
               section = 4
               i = 0
               Do While Not rec.EOF
                  Label2(i).Caption = rec.Fields(0)
                  If i = 5 Then Exit Do
                  i = i + 1
                  rec.MoveNext
               Loop
            End If
         End If
         
         If section = 0 Then
            reccheck
            rec.Open "select REGNO from FAMILYTABLE  where FATHERNAME like '%" & search & "'", con, adOpenDynamic, adLockPessimistic
            If rec.EOF = False Then
               i = 0
               section = 5
               Do While Not rec.EOF
                  temp(i) = rec.Fields(0)
                  If i = 5 Then Exit Do
                     i = i + 1
                  rec.MoveNext
               Loop
               For J = 0 To i
                   reccheck
                   rec.Open "select STUDENTNAME from MAINTABLE where REGNO = '" & temp(J) & "'", con, adOpenDynamic, adLockPessimistic
                   If rec.EOF = False Then Label2(J).Caption = rec.Fields(0)
               Next J
            End If
         End If
         
         If section = 0 Then
            reccheck
            rec.Open "select REGNO from FAMILYTABLE  where MOTHERNAME like '%" & search & "'", con, adOpenDynamic, adLockPessimistic
            If rec.EOF = False Then
               i = 0
               section = 6
               Do While Not rec.EOF
                  temp(i) = rec.Fields(0)
                  If i = 5 Then Exit Do
                  i = i + 1
                  rec.MoveNext
               Loop
               For J = 0 To i
                   reccheck
                   rec.Open "select STUDENTNAME from MAINTABLE where REGNO = '" & temp(J) & "'", con, adOpenDynamic, adLockPessimistic
                   If rec.EOF = False Then Label2(J).Caption = rec.Fields(0)
               Next J
            End If
         End If
         
         If section = 0 Then
            reccheck
            rec.Open "select REGNO from FAMILYTABLE  where GUARDIANNAME like '%" & search & "'", con, adOpenDynamic, adLockPessimistic
            If rec.EOF = False Then
               i = 0
               section = 7
               Do While Not rec.EOF
                  temp(i) = rec.Fields(0)
                  If i = 5 Then Exit Do
                  i = i + 1
                  rec.MoveNext
               Loop
               For J = 0 To i
                   reccheck
                   rec.Open "select STUDENTNAME from MAINTABLE where REGNO = '" & temp(J) & "'", con, adOpenDynamic, adLockPessimistic
                   If rec.EOF = False Then Label2(J).Caption = rec.Fields(0)
               Next J
            End If
         End If
          
       End If
        
       If length < 75 Then
          If section = 0 Then
             reccheck
             rec.Open "select STUDENTNAME from PERSONALTABLE where STUDENTADDRESS like '%" & search & "'", con, adOpenDynamic, adLockPessimistic
             If rec.EOF = False Then
                section = 8
                i = 0
                Do While Not rec.EOF
                  Label2(i).Caption = rec.Fields(0)
                  i = i + 1
                  rec.MoveNext
                Loop
             End If
          End If
          
          If section = 0 Then
             reccheck
             rec.Open "select REGNO from FAMILYTABLE  where FATHERADDRESS like '%" & search & "'", con, adOpenDynamic, adLockPessimistic
             If rec.EOF = False Then
                i = 0
                section = 9
                Do While Not rec.EOF
                   temp(i) = rec.Fields(0)
                   If i = 5 Then Exit Do
                   i = i + 1
                   rec.MoveNext
                Loop
                For J = 0 To i
                    reccheck
                    rec.Open "select STUDENTNAME from MAINTABLE where REGNO = '" & temp(J) & "'", con, adOpenDynamic, adLockPessimistic
                    If rec.EOF = False Then Label2(J).Caption = rec.Fields(0)
                Next J
             End If
          End If
             
          If section = 0 Then
             reccheck
             rec.Open "select REGNO from FAMILYTABLE  where GUARDIANADDRESS like '%" & search & "'", con, adOpenDynamic, adLockPessimistic
             If rec.EOF = False Then
                i = 0
                section = 11
                Do While Not rec.EOF
                   temp(i) = rec.Fields(0)
                   If i = 5 Then Exit Do
                   i = i + 1
                   rec.MoveNext
                Loop
                For J = 0 To i
                    reccheck
                    rec.Open "select STUDENTNAME from MAINTABLE where REGNO = '" & temp(J) & "'", con, adOpenDynamic, adLockPessimistic
                    If rec.EOF = False Then Label2(J).Caption = rec.Fields(0)
                Next J
             End If
          End If
             
       End If
          
       If length < 100 Then
             If section = 0 Then
                reccheck
                rec.Open "select REGNO from FAMILYTABLE  where BROSIS like '%" & search & "'", con, adOpenDynamic, adLockPessimistic
                If rec.EOF = False Then
                   i = 0
                   section = 12
                   Do While Not rec.EOF
                      temp(i) = rec.Fields(0)
                      If i = 5 Then Exit Do
                      i = i + 1
                      rec.MoveNext
                   Loop
                   For J = 0 To i
                       reccheck
                       rec.Open "select STUDENTNAME from MAINTABLE where REGNO = '" & temp(J) & "'", con, adOpenDynamic, adLockPessimistic
                       If rec.EOF = False Then Label2(J).Caption = rec.Fields(0)
                    Next J
                End If
             End If
       End If
       
       If section = 0 Then
          Label1(0).Visible = True
       Else
          Label1(0).Visible = False
       End If
              
   End If
End Sub

Private Sub imgSearch2_Click()
   connection
   reccheck
   
   rec.Open "select * from SEARCHENGINE", con, adOpenDynamic, adLockPessimistic
   
   Dim sprocedure As String
   Dim i As Integer
   i = 11
   sprocedure = ""
   
   If rec.EOF = False Then
     Do While Not rec.EOF
       If rec.Fields(2) = cmbSearch.Text Then
          sprocedure = rec.Fields(1)
          Exit Do
       End If
       rec.MoveNext
     Loop
   End If
   If sprocedure = "" Then
      Label1(1).Visible = True
      For i = 6 To 11
        Label2(i).Visible = False
      Next i
   Else
      Label1(1).Visible = False
      
      connection
      reccheck
      
      rec.Open sprocedure, con, adOpenDynamic, adLockPessimistic
      If rec.EOF = False Then
         While Not rec.EOF And i > 6
            Label2(i).Caption = rec.Fields(0)
            Label2(i).Visible = True
            rec.MoveNext
            i = i + 1
         Wend
      Else
        For i = 6 To 11
          Label2(i).Visible = True
        Next i
      End If
   End If
End Sub

Private Sub Label2_Click(Index As Integer)
  Dim i As Integer
  Shape2.Visible = True
  For i = 0 To 4
    Label3(i).Visible = True
  Next i
  
  reccheck
  rec.Open "select REGNO from MAINTABLE where STUDENTNAME = '" & Label2(Index).Caption & "'", con, adOpenDynamic, adLockPessimistic
  regno = rec.Fields(0)
End Sub

Private Sub Label3_Click(Index As Integer)
   Select Case Index
      Case 0: frmPersonalInfo.txtRegNo.Text = regno
              frmPersonalInfo.Show
      Case 1: frmFamilyInfo.txtRegNo.Text = regno
              frmFamilyInfo.Show
      Case 2: frmEducationalInfo.txtRegNo.Text = regno
              frmEducationalInfo.Show
      Case 3: frmMarkPrint.txtRegNo.Text = regno
              frmMarkPrint.Show
      Case 4: frmPhysicalInfo.txtRegNo.Text = regno
              frmPhysicalInfo.Show
   End Select
End Sub

Private Sub txtRegNo_Change()
  If txtRegNo.Text = "" Then
     lblName.Enabled = True
     lblCourse.Enabled = True
     lblYear.Enabled = True
     txtName.Enabled = True
     cmbCourse.Enabled = True
     cmbYear.Enabled = True
  Else
     lblName.Enabled = False
     lblCourse.Enabled = False
     lblYear.Enabled = False
     txtName.Enabled = False
     cmbCourse.Enabled = False
     cmbYear.Enabled = False
  End If
End Sub

Private Sub cmbYear_Change()
  If cmbYear.Text = "" Then
     lblRegNo.Enabled = True
     txtRegNo.Enabled = True
  Else
     lblRegNo.Enabled = False
     txtRegNo.Enabled = False
  End If
End Sub

Private Sub cmbCourse_Change()
  If cmbCourse.Text = "" Then
     lblRegNo.Enabled = True
     txtRegNo.Enabled = True
  Else
     lblRegNo.Enabled = False
     txtRegNo.Enabled = False
  End If
End Sub

Private Sub txtName_Change()
  If txtName.Text = "" Then
     lblRegNo.Enabled = True
     txtRegNo.Enabled = True
  Else
     lblRegNo.Enabled = False
     txtRegNo.Enabled = False
  End If
End Sub

Private Sub imgReport_Click()
  If txtRegNo.Text <> "" Then
     If CheckRegNo(txtRegNo, 8) = True Then
        'do nothing
     End If
  Else
     connection
     reccheck
     
     rec.Open "select STUDENTNAME from MAINTABLE where STUDENTNAME ='" & Trim(txtName.Text) & "' and COURSE ='" & Trim(cmbCourse.Text) & "' and YEAROFSTUDY ='" & Trim(cmbYear.Text) & "'", con, adOpenDynamic, adLockPessimistic
     If rec.EOF = False Then
        While Not rec.EOF
          cmbStudent.AddItem (rec.Fields(0))
        Wend
        txtRegNo.Text = rec.Fields(0)
     Else
       Label8.Visible = True
       txtRegNo.Text = ""
       Exit Sub
     End If
  End If
  
  '------------------------Search for Records--------------------------------
  
  If CheckCombo(cmbInfoType, "Information Type") = True Then
     Select Case cmbInfoType.Text
         Case "Personal": frmPersonalInfo.txtRegNo.Text = txtRegNo.Text
                         ' frmPersonalInfo.imgReport_Click()
                          frmPersonalInfo.Show
                          'frmPersonalInfo.imgReport_Click
         Case "Educational":  frmEducationalInfo.txtRegNo.Text = txtRegNo.Text
                             ' frmEducationalInfo.imgReport_Click
                              frmEducationalInfo.Show
                              
         Case "Family":  frmFamilyInfo.txtRegNo.Text = txtRegNo.Text
                         frmFamilyInfo.Show
                         'frmFamilyInfo.imgReport_Click
         Case "Internal Exams": frmMarkPrint.txtRegNo.Text = txtRegNo.Text
                                frmMarkPrint.Show
                          '      frmMarkPrint.imgReport_Click
     End Select
 End If
End Sub

Private Sub txtRegNo_LostFocus()
   If txtRegNo.Text <> "" Then
     If CheckRegNo(txtRegNo, 8) = True Then
        'do nothing
     End If
   End If
End Sub

Private Sub cmbStudent_Change()
   connection
   reccheck
   
   rec.Open "select REGNO from MAINTABLE where STUDENTNAME = '" & cmbStudent.Text, con, adOpenDynamic, adLockPessimistic
   txtRegNo.Text = rec.Fields(0)
End Sub

Private Sub cmbYear_LostFocus()
     connection
     reccheck
     
     rec.Open "select STUDENTNAME from MAINTABLE where COURSE ='" & Trim(cmbCourse.Text) & "' and YEAROFSTUDY ='" & Trim(cmbYear.Text) & "' and " & "STUDENTNAME like '" & Trim(txtName.Text) & "%'", con, adOpenDynamic, adLockPessimistic
     If rec.EOF = False Then
       'txtRegNo.Text = rec.Fields(0)
       cmbStudent.Clear
       While Not rec.EOF
          cmbStudent.AddItem (rec.Fields(0))
          rec.MoveNext
       Wend
     Else
       Label8.Visible = True
       txtRegNo.Text = ""
       Exit Sub
     End If
End Sub

Private Sub txtSearch_Change()
  Dim i As Integer
  For i = 0 To 5
    Label2(i).Caption = ""
  Next i
  Label1(0).Visible = False
End Sub
