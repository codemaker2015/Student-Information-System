Attribute VB_Name = "Functions"
Option Explicit

Public rs1 As Recordset
Public db As Database
Public td As TableDef
Public fl As Field
Public igRow As Integer, igColumn As Integer
Public iFields As Integer, iRecords As Integer
Public vargBookmark As Variant

Public test1 As Integer, test2 As Integer, attendance As Integer, seminar As Integer, internal As Integer, total As Integer
Public marksystem As String
Public choice As Integer

Function CreateDB(ByRef dbname As String, tbname As String)
  
         Set db = CreateDatabase(dbname, dbLangGeneral)
         Set td = db.CreateTableDef(tbname)

         ' After you create the database, you need to add fields to it.
         For iFields = 1 To 2 ' The last number can be changed to the
                              ' number of fields you want in the database.
            Set fl = td.CreateField("Field " & CStr(iFields), dbText)
            td.Fields.Append fl
         Next iFields

         db.TableDefs.Append td
    
         db.Close
End Function

Function DateTime(ByVal lblDate As Label, ByVal lblTime As Label)
   'display date & time
    lblDate.Caption = Format$(Date, "d/m/yyyy")
    lblTime.Caption = Format$(Time, "h:nn AM/PM")
End Function

Function CheckName(ByVal txt As TextBox, ByVal length As Integer) As Boolean
  Dim str As String
  Dim i As Integer
  str = LCase(Trim(txt.Text))
  If Len(str) > length Then
    MsgBox "Length of the name can't be greather than " + CStr(length), vbCritical, "Data Entry Error"
    txt.SetFocus
    CheckName = False
    Exit Function
  End If
  For i = 1 To Len(str)
  If Mid(str, i, 1) Like "[a-z]" Or Mid(str, i, 1) = " " Then
     'do nothing
  Else
     MsgBox "Name contains an illegal character"
     Exit For
     txt.SetFocus
     CheckName = False
     Exit Function
  End If
  Next i
  CheckName = True
End Function

Function CheckCombo(ByVal cmb As ComboBox, ByVal fillitem As String) As Boolean
  Select Case cmb.Text
    Case "--Select--": MsgBox "You should select " + fillitem, vbCritical, "Data Entry Error"
         cmb.SetFocus
         CheckCombo = False
         Exit Function
    Case Empty: MsgBox fillitem + " can't become null", vbCritical, "Data Entry Error"
         cmb.SetFocus
         CheckCombo = False
         Exit Function
  End Select
  CheckCombo = True
End Function

Function CheckRegNo(ByVal txt As TextBox, ByVal length As Integer) As Boolean
   Dim reg As String
   Dim i As Integer
   reg = Trim(txt.Text)
    If Len(reg) = 0 Then
      MsgBox "Register Number can't become null", vbCritical, "Data Entry Error"
      txt.SetFocus
      CheckRegNo = False
      Exit Function
   End If
   If Len(reg) < length Then
      MsgBox "Register Number is too short", vbCritical, "Data Entry Error"
      txt.SetFocus
      CheckRegNo = False
      Exit Function
   End If
   If Len(reg) > length + 1 Then
      MsgBox "Register Number is too long", vbCritical, "Data Entry Error"
      txt.SetFocus
      CheckRegNo = False
      Exit Function
   End If
   For i = 1 To Len(reg) - 1
      If Mid(reg, i, 1) Like "[0-9]" Then
        If i = 1 Then
           If Mid(reg, i, 1) Like "[1-9]" Then
              'Do nothing
           Else
              MsgBox "First digit of Register Number can't become zero", vbCritical, "Data Entry Error"
              txt.SetFocus
              Exit For
              CheckRegNo = False
              Exit Function
           End If
        End If
      Else
        MsgBox "Register Number entered is invalid", vbCritical, "Data Entry Error"
        txt.SetFocus
        CheckRegNo = False
        Exit Function
      End If
   Next i
   CheckRegNo = True
End Function


Function CheckPhone(ByVal txt As TextBox) As Boolean
   Dim reg As String
   Dim i As Integer
   reg = Trim(txt.Text)
   If Len(reg) < 10 Then
      MsgBox "Phone Number is too short", vbCritical, "Data Entry Error"
      txt.SetFocus
      CheckPhone = False
      Exit Function
   End If
   If Len(reg) > 12 Then
      MsgBox "Phone Number is too long", vbCritical, "Data Entry Error"
      txt.SetFocus
      CheckPhone = False
      Exit Function
   End If
   For i = 1 To Len(reg) - 1
      If Mid(reg, i, 1) Like "[0-9]" Then
        'Do nothing
      Else
        MsgBox "Phone Number entered is invalid", vbCritical, "Data Entry Error"
        txt.SetFocus
        Exit Function
      End If
   Next i
   CheckPhone = True
End Function

Function ValRegNo(ByRef KeyAscii As Integer)
   If KeyAscii < 48 Or KeyAscii > 57 Then
      If KeyAscii <> 8 Then
         MsgBox "Only numbers are allowed", , "Error"
         KeyAscii = 0
      End If
   End If
End Function

Function ValAddress(ByRef KeyAscii As Integer)
  If KeyAscii < 96 Then
     If KeyAscii < 64 Or KeyAscii > 91 Then
        If KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
           MsgBox "Only Alphabets are allowed", , "Error"
           KeyAscii = 0
        End If
     End If
  Else
     If KeyAscii > 123 Then
        KeyAscii = 0
     End If
  End If
End Function

Function ValName(ByRef KeyAscii As Integer)
  If KeyAscii < 96 Then
     If KeyAscii < 64 Or KeyAscii > 91 Then
        'If KeyAscii = 32 Then
        If KeyAscii <> 8 And KeyAscii <> 32 Then
           MsgBox "Only Alphabets are allowed", , "Error"
           KeyAscii = 0
        End If
     End If
  Else
     If KeyAscii > 123 Then
        KeyAscii = 0
     End If
  End If
End Function

Function ValPhone(ByRef KeyAscii As Integer)
   If KeyAscii < 48 Or KeyAscii > 57 Then
      If KeyAscii <> 8 And KeyAscii <> 43 Then
         MsgBox "Only numbers are allowed", , "Error"
         KeyAscii = 0
      End If
   End If
End Function

Function Theme(ByRef frm As Form) As Integer
  Dim t As Integer
  Dim fso As New FileSystemObject
  If fso.FileExists(App.Path & "\Theme\Theme.txt") = True Then
     Open App.Path & "\Theme\Theme.txt" For Input As #1
     t = Input(LOF(1), #1)
     Close #1
     If t > -1 And t < 10 Then
        If fso.FileExists(App.Path & "\Theme\" & t & ".jpg") = True Then frm.Picture = LoadPicture(App.Path & "\Theme\" & t & ".jpg")
     End If
  End If
  Theme = t
End Function

Function AddInfo(ByVal txt As TextBox, table As String) As Boolean
  If Trim(txt.Text) <> "" Then
      reccheck
      rec.Open "delete from " & table & " where NAME = 'Others'", con, adOpenDynamic, adLockPessimistic
      reccheck
      rec.Open "insert into " & table & " values('" & Trim(txt.Text) & "')", con, adOpenDynamic, adLockPessimistic
      reccheck
      rec.Open "insert into " & table & " values('Others')", con, adOpenDynamic, adLockPessimistic
      AddInfo = True
   Else
      AddInfo = False
   End If
End Function

Function RemoveInfo(ByVal txt As TextBox, table As String) As Boolean
  If Trim(txt.Text) <> "" Then
      If Trim(txt.Text) = "Others" Or Trim(txt.Text) = "--Select--" Then
         MsgBox "Could not perform the Operation", vbCritical, "Remove Info"
      Else
         reccheck
         rec.Open "delete from " & table & " where NAME = '" & Trim(txt.Text) & "'", con, adOpenDynamic, adLockPessimistic
         RemoveInfo = True
      End If
   Else
      RemoveInfo = False
   End If
End Function
