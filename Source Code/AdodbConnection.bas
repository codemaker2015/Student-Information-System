Attribute VB_Name = "AdodbConnection"
'DECLARATION AS USED IN THE PROJECT FOR CONNECTION

Public Con As New ADODB.Connection
Public rs_find As New ADODB.Recordset
Public rs_student As New ADODB.Recordset
Public rs_stugrid As New ADODB.Recordset
Public rs_att As New ADODB.Recordset
Public rs_class As New ADODB.Recordset
Public rs_temp As New ADODB.Recordset
Dim str As String

Public Sub connect()
'SUB FOR CREATING CONNECTION
Set Con = New ADODB.Connection
Set rs_student = New ADODB.Recordset
Set rs_find = New ADODB.Recordset
Set rs_stugrid = New ADODB.Recordset
Set rs_att = New ADODB.Recordset
Set rs_class = New ADODB.Recordset
Con.CursorLocation = adUseClient
Con.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=STUDENTINFORMATIONSYSTEM;Data Source=."

'CONNECTION IS OPEN
Con.Open
rs_student.Open "SELECT *  FROM student_details ", Con, adOpenStatic, adLockPessimistic

End Sub

Public Sub student_id()
On Error Resume Next
Call connect
Con.Refresh
With rs_find
.Open "select * from student_details", Con, adOpenDynamic, adLockPessimistic
.MoveLast
If IsNull(.Fields("ID").Value) Then
Texstu_id.Text = 1
Else
no = .Fields("ID") + 1
Texstu_id.Text = no
End If
.Close
End With
End Sub

Public Sub pic()
If frmsturegister1.cdb.FileName <> "" Then
       frmsturegister1.pcbox.Picture = LoadPicture(frmsturegister1.cdb.FileName)
        pic_name = frmsturegister1.cdb.FileName
        pic_ext = Right(frmsturegister1.cdb.FileTitle, 4)
        pic_changed = True
    End If
Call connect
End Sub
