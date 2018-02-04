VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAttachment 
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18180
   Icon            =   "frmAtttachment.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAtttachment.frx":000C
   ScaleHeight     =   11010
   ScaleWidth      =   18180
   Begin VB.TextBox txtRegNo 
      Height          =   375
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15840
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgPrev 
      Height          =   375
      Left            =   9720
      Picture         =   "frmAtttachment.frx":1C0C3
      Top             =   9840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgNext 
      Height          =   375
      Left            =   10680
      Picture         =   "frmAtttachment.frx":1C5DC
      Top             =   9840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgReport 
      Height          =   450
      Left            =   12360
      Picture         =   "frmAtttachment.frx":1C98D
      Top             =   9840
      Width           =   375
   End
   Begin VB.Image imgCancel 
      Height          =   375
      Left            =   840
      Picture         =   "frmAtttachment.frx":1CF44
      Top             =   9840
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   855
      Index           =   1
      Left            =   480
      Top             =   9600
      Width           =   15255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   855
      Index           =   0
      Left            =   480
      Top             =   240
      Width           =   15255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg No:"
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
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   8295
      Left            =   480
      Top             =   1200
      Width           =   15255
   End
   Begin VB.Image imgSave 
      Height          =   450
      Left            =   13320
      Picture         =   "frmAtttachment.frx":1D424
      Top             =   9840
      Width           =   375
   End
   Begin VB.Image imgOpen 
      Height          =   450
      Left            =   14280
      Picture         =   "frmAtttachment.frx":1D81B
      Top             =   9840
      Width           =   375
   End
   Begin VB.Image imgFile 
      Height          =   7995
      Left            =   600
      Top             =   1320
      Width           =   15000
   End
End
Attribute VB_Name = "frmAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private picpath As String
Private fso As New FileSystemObject
Private Sub Form_Load()
      On Error GoTo error_para
      
      Theme frmAttachment
      connection
       
      If fso.FolderExists(App.Path & "\sis") = False Then fso.CreateFolder (App.Path & "\sis")
      If fso.FolderExists(App.Path & "\sis\certificate") = False Then fso.CreateFolder (App.Path & "\sis\certificate")
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub imgCancel_Click()
  Unload Me
End Sub

Private Sub imgOpen_Click()
  On Error GoTo error_para

  MsgBox "Size of the picture must be 1000 x 1000", vbExclamation
  
  With CommonDialog1
    .FileName = ""
    .Filter = "All Picture Files | *.bmp;*.jpg;*.gif | JPEG (*.jpg) | *.jpg | Bitmap (*.bmp) | *.bmp "
    .DialogTitle = "Open Image..."
    .CancelError = True
    .ShowOpen
    If .FileName <> "" Then
      imgFile.Picture = LoadPicture(CommonDialog1.FileName)
      If imgFile.height > 10000 Or imgFile.width > 16000 Then MsgBox "Picture Dimension is not suitable", vbCritical
    End If
  End With
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub imgReport_Click()
   On Error GoTo error_para
   
   Dim i As Integer
   i = 0
   reccheck
   rec.Open "select PICPATH from ATTACHMENTTABLE where REGNO = '" & Val(Trim(txtRegNo.Text)) & "'", con, adOpenDynamic, adLockPessimistic
   If rec.EOF = False Then
     While Not rec.EOF
       i = i + 1
       rec.MoveNext
     Wend
     
     reccheck
     rec.Open "select PICPATH from ATTACHMENTTABLE where REGNO = '" & Val(Trim(txtRegNo.Text)) & "'", con, adOpenDynamic, adLockPessimistic
     
     If i > 1 Then
        imgNext.Visible = True
        imgFile.Picture = LoadPicture(rec.Fields(0))
     Else
        imgFile.Picture = LoadPicture(rec.Fields(0))
     End If
   Else
      MsgBox "No such Attachment found", vbExclamation, ""
   End If
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub imgSave_Click()
  On Error GoTo error_para
  
  Dim temp As String
  temp = "iiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii"
  Dim i As Integer
  i = 1
  If imgFile.height < 10000 And imgFile.width < 16000 Then
     reccheck
     CheckRegNo txtRegNo, 8
     picpath = App.Path & "\sis\certificate\" & Val(Trim(txtRegNo.Text)) + 1729 & ".jpg"
     Do While True
       If fso.FileExists(picpath) Then
          picpath = App.Path & "\sis\certificate\" & Val(Trim(txtRegNo.Text)) + 1729 & Mid(temp, 1, i) & ".jpg"
          i = i + 1
       Else
          Exit Do
       End If
     Loop
     fso.CopyFile CommonDialog1.FileName, picpath
     
     rec.Open "insert into ATTACHMENTTABLE values('" & Trim(txtRegNo.Text) & "','" & picpath & "')", con, adOpenDynamic, adLockPessimistic
     MsgBox "Attachment added Successfully", vbInformation, ""
  End If
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub imgNext_Click()
   On Error GoTo error_para
   
   imgPrev.Visible = True
   
   If rec.EOF = False Then rec.MoveNext
   If rec.EOF = False Then
      imgFile.Picture = LoadPicture(rec.Fields(0))
      imgFile.Refresh
   End If
error_para:
   MsgBox "Oops error Occured", vbOKOnly, ""
End Sub

Private Sub imgPrev_Click()
   If rec.BOF = False Then rec.MovePrevious
   If rec.BOF = False Then imgFile.Picture = LoadPicture(rec.Fields(0))
End Sub

Private Sub txtRegNo_KeyPress(KeyAscii As Integer)
   ValRegNo KeyAscii
End Sub

Private Sub txtRegNo_LinkClose()
  CheckRegNo txtregbo, 8
End Sub
