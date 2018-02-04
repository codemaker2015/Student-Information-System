Attribute VB_Name = "DBConnect"
Public rs As New ADODB.Recordset
Public conn As New ADODB.Connection
Public sql As String
Public Constr As String
Public CurrentUser As String
Public UserTitle As String
Public UserLog As Integer
Public rAdd As Boolean, rDelete As Boolean, rUpdate As Boolean, rPrint As Boolean
Public SuccessLogin As Boolean

Sub Main()
Constr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=STUDENTINFORMATIONSYSTEM;Data Source=."
conn.Open Con
With rs
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
End With
'LoginFrm.Show
'MainFrm.Show
'Form1.Show
SplashFrm.Show
End Sub
