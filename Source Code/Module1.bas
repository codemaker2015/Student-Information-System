Attribute VB_Name = "Module1"
Function CheckCombo(ByVal cmb As ComboBox, ByVal fillitem As String)
  Select Case cmb.Text
    Case "--Select--": MsgBox "You should select " + fillitem, vbCritical, "Data Entry Error"
    Case Empty: MsgBox fillitem + " can't become null", vbCritical, "Data Entry Error"
  End Select
End Function

Function CheckRegNo(ByVal txt As TextBox, ByVal length As Integer)
   Dim reg As String
   reg = Trim(txt.Text)
   If Len(reg) < length Then
      MsgBox "Register Number is too short", vbCritical, "Data Entry Error"
      Exit Function
   End If
   If Len(reg) > length + 1 Then
      MsgBox "Register Number is too long", vbCritical, "Data Entry Error"
      Exit Function
   End If
   For i = 1 To Len(reg) - 1
      If Mid(reg, i, 1) Like "[0-9]" Then
        If i = 1 Then
           If Mid(reg, i, 1) Like "[1-9]" Then
              'Do nothing
           Else
              MsgBox "First digit of Register Number can't become zero", vbCritical, "Data Entry Error"
              Exit For
           End If
        End If
      Else
        MsgBox "Register Number entered is invalid", vbCritical, "Data Entry Error"
        Exit Function
      End If
   Next i
End Function


Function CheckPhone(ByVal txt As TextBox)
   Dim reg As String
   reg = Trim(txt.Text)
   If Len(reg) < 10 Then
      MsgBox "Phone Number is too short", vbCritical, "Data Entry Error"
      Exit Function
   End If
   If Len(reg) > 12 Then
      MsgBox "Phone Number is too long", vbCritical, "Data Entry Error"
      Exit Function
   End If
   For i = 1 To Len(reg) - 1
      If Mid(reg, i, 1) Like "[0-9]" Then
        'Do nothing
      Else
        MsgBox "Phone Number entered is invalid", vbCritical, "Data Entry Error"
        Exit Function
      End If
   Next i
End Function
