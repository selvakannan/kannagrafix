Attribute VB_Name = "Module1"
' put the following in a module

Option Explicit
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim username As String
Dim passwd As String
Dim serverIP As String
Dim db As String
Public Function connectMysql(username As String, passwd As String, serverIP As String, db As String, conn As ADODB.Connection, rs As ADODB.Recordset)
   Set conn = New ADODB.Connection
   Set rs = New ADODB.Recordset
   conn.CursorLocation = adUseClient
   conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & serverIP & ";UID=" & username & ";PWD=" & passwd & ";DATABASE=" & db & ";" _
   & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 163841
   conn.Open
End Function

'****************************

