VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Student Information System 2015"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20370
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "frmMain.frx":05B2
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   10710
      Left            =   0
      Picture         =   "frmMain.frx":1C669
      ScaleHeight     =   10710
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2310
         Left            =   840
         TabIndex        =   11
         Top             =   7800
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483645
         Appearance      =   0
         StartOfWeek     =   16515073
         TitleBackColor  =   -2147483637
         CurrentDate     =   41916
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H8000000D&
         Height          =   2775
         Left            =   240
         Top             =   7560
         Width           =   3735
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H8000000D&
         Height          =   615
         Left            =   240
         Top             =   240
         Width           =   3735
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H8000000D&
         Height          =   6495
         Left            =   240
         Top             =   960
         Width           =   3735
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H8000000D&
         Height          =   1215
         Left            =   2520
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H8000000D&
         Height          =   1215
         Left            =   480
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H8000000D&
         Height          =   1215
         Left            =   2520
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000D&
         Height          =   1215
         Left            =   480
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000D&
         Height          =   1215
         Left            =   2520
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         Height          =   1215
         Left            =   480
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblReport 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   2520
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   4680
         Width           =   975
      End
      Begin VB.Image imgReport 
         Height          =   1200
         Left            =   2520
         MouseIcon       =   "frmMain.frx":32729
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":32FF3
         ToolTipText     =   "Report"
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lblDelete 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   6960
         Width           =   975
      End
      Begin VB.Image imgDelete 
         Height          =   1200
         Left            =   480
         MouseIcon       =   "frmMain.frx":340BE
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":34988
         ToolTipText     =   "Delete"
         Top             =   5640
         Width           =   1200
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   280
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   280
         Width           =   855
      End
      Begin VB.Label lblExit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   2520
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   6960
         Width           =   1215
      End
      Begin VB.Image imgExit 
         Height          =   1200
         Left            =   2520
         MouseIcon       =   "frmMain.frx":357C0
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":3608A
         ToolTipText     =   "Exit"
         Top             =   5640
         Width           =   1200
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   600
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   4680
         Width           =   645
      End
      Begin VB.Image imgSearch 
         Height          =   1200
         Left            =   480
         MouseIcon       =   "frmMain.frx":36AD2
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":3739C
         ToolTipText     =   "Search"
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edition"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   2660
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   2520
         Width           =   645
      End
      Begin VB.Image imgEdit 
         Height          =   1200
         Left            =   2520
         MouseIcon       =   "frmMain.frx":383DD
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":38CA7
         ToolTipText     =   "Edit"
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblRegistration 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   360
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   2520
         Width           =   1170
      End
      Begin VB.Image imgRegistration 
         Height          =   1200
         Left            =   480
         MouseIcon       =   "frmMain.frx":3A539
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":3AE03
         ToolTipText     =   "Registration"
         Top             =   1200
         Width           =   1200
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   20370
      TabIndex        =   12
      Top             =   0
      Width           =   20370
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuRegistration 
         Caption         =   "Registration"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdition 
         Caption         =   "Edition"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuViewReport 
         Caption         =   "View Report"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "Utility"
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuTheme 
         Caption         =   "Theme"
      End
      Begin VB.Menu mnuSISSettings 
         Caption         =   "System Settings"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuAboutCollege 
         Caption         =   "About College"
      End
   End
   Begin VB.Menu mnuLogout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAlarmCollection As Collection
Private mbHidden As Boolean
Private mbIsBusy As Boolean

Private Sub ShowIconTray()
    Dim lRet As Long
    Dim nd As NOTIFYICONDATA
    With nd
      .cbSize = Len(nd)
      .hwnd = Picture1.hwnd
      .uID = 1&
      .szTip = "SIS" & Chr(0)
      .uCallbackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    End With
    lRet = Shell_NotifyIconA(NIM_ADD, nd)
End Sub

Private Sub MDIForm_Load()
  'center the form:
  Dim fso As New FileSystemObject
  Dim th As Integer
  Me.Top = (Screen.height - Me.height) / 2
  Me.Left = (Screen.width - Me.width) / 2
  DateTime lblDate, lblTime
  
  th = Theme(frmMain)
  If fso.FileExists(App.Path & "\Theme\Side" & th & ".jpg") = True Then Picture1.Picture = LoadPicture(App.Path & "\Theme\Side" & th & ".jpg")
  If fso.FileExists(App.Path & "\Theme\Registration" & th & ".jpg") = True Then imgRegistration.Picture = LoadPicture(App.Path & "\Theme\Registration" & th & ".jpg")
  If fso.FileExists(App.Path & "\Theme\Edit" & th & ".jpg") = True Then imgEdit.Picture = LoadPicture(App.Path & "\Theme\Edit" & th & ".jpg")
  If fso.FileExists(App.Path & "\Theme\Search" & th & ".jpg") = True Then imgSearch.Picture = LoadPicture(App.Path & "\Theme\Search" & th & ".jpg")
  If fso.FileExists(App.Path & "\Theme\Report" & th & ".jpg") = True Then imgReport.Picture = LoadPicture(App.Path & "\Theme\Report" & th & ".jpg")
  If fso.FileExists(App.Path & "\Theme\Delete" & th & ".jpg") = True Then imgDelete.Picture = LoadPicture(App.Path & "\Theme\Delete" & th & ".jpg")
  If fso.FileExists(App.Path & "\Theme\Exit" & th & ".jpg") = True Then imgExit.Picture = LoadPicture(App.Path & "\Theme\Exit" & th & ".jpg")
     
  ShowIconTray
    
   'lblTimeDisplay.Caption = Format(Now, "short time")
   ' Me.cmdTurnOff.Enabled = False
    'LoadAlarms mAlarmCollection
    'ListAlarms
    mbHidden = CBool(GetSetting(App.EXEName, "Settings", "Hidden", "False"))
    If mbHidden Then
        Me.Hide
    End If
    
  frmWelcome.Show

End Sub

Private Sub imgRegistration_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Then
      choice = 1
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub lblRegistration_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Then
      choice = 1
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub imgEdit_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Then
      choice = 2
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub lblEdit_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Then
      choice = 2
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub imgSearch_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Or frmlogin.UserNameCmb.Text = "STUDENT" Then
      choice = 3
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub lblSearch_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Or frmlogin.UserNameCmb.Text = "STUDENT" Then
      choice = 3
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub imgReport_Click()
  choice = 4
  frmLoading.Show
End Sub

Private Sub lblReport_Click()
  choice = 4
  frmLoading.Show
End Sub

Private Sub imgDelete_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Then
      choice = 5
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub lblDelete_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Then
      choice = 5
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub imgExit_Click()
  Call MDIForm_QueryUnload(1, 1)
End Sub

Private Sub lblExit_Click()
  Call MDIForm_QueryUnload(1, 1)
End Sub

Private Sub DeleteIconTray()
    Dim nd As NOTIFYICONDATA
    Dim iRet As Long
    With nd
      .cbSize = Len(nd)
      .hwnd = Picture1.hwnd
      .uID = 1&
    End With
    iRet = Shell_NotifyIconA(NIM_DELETE, nd)
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MsgBox("Are you sure you want to quit ?", vbYesNo, "Exit") = vbYes Then
     SaveSetting App.EXEName, "Settings", "Hidden", mbHidden
     DeleteIconTray
     frmshutdown.Label2.Caption = "Shutdown"
     frmshutdown.Show
     'End
  End If
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show
End Sub

Private Sub mnuLogout_Click()
   frmMain.Hide
   frmlogin.passwordtxt = ""
   frmlogin.Show
End Sub

Private Sub mnuNotepad_Click()
  On Error GoTo err
    Shell "Notepad.exe", vbNormalFocus
    Exit Sub
err:
    MsgBox "You don't have Notepad installed in your computer.", vbExclamation, "Notepad Missing"
End Sub

Private Sub mnuRegistration_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Then
      choice = 1
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub mnuEdition_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Then
      choice = 2
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub mnuSearch_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Or frmlogin.UserNameCmb.Text = "STUDENT" Then
      choice = 3
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub mnuViewReport_Click()
   choice = 4
   frmLoading.Show
End Sub

Private Sub mnuDelete_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Then
      choice = 5
      frmLoading.Show
   Else
      MsgBox "You should have Administrator previlege to use this option", , "Login Error"
   End If
End Sub

Private Sub mnuExit_Click()
   Call MDIForm_QueryUnload(1, 1)
End Sub

Private Sub mnuSISSettings_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Then
      choice = 7
      frmLoading.Show
   Else
      MsgBox "login as ADMIN for change System Settings", , "Login Error"
   End If
End Sub

Private Sub mnuTheme_Click()
   If frmlogin.UserNameCmb.Text = "ADMIN" Then
      choice = 6
      frmLoading.Show
   Else
      MsgBox "login as ADMIN for change Theme", , "Login Error"
   End If
End Sub
