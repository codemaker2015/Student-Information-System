Attribute VB_Name = "Module1"
Option Explicit
Public rs As ADODB.Recordset
Public con As ADODB.connection
Public cmd As ADODB.Command
Public constring As String
Public Sub connection()
  constring = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=STUDENTINFORMATIONSYSTEM;Data Source=."
  Set rs = New ADODB.Recordset
  Set con = New ADODB.connection
  Set cmd = New ADODB.Command
  
  con.ConnectionString = constring
  con.Open
  cmd.ActiveConnection = con
'  cmd.CommandText = "SELECT SUBJECTNAME  FROM SUBJECTTABLE where  course = " + course + " and sem = " + lstSem.Text
 ' Set rs = cmd.Execute
End Sub
