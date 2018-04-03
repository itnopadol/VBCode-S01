Attribute VB_Name = "Module1"
Option Explicit

Global ConnectionString As String
Global vConnection As New ADODB.Connection
Global vMemPrinter As String

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Sub InitializeDataBase()
If vConnection.State <> 0 Then
vConnection.Close
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=VBUser;Password='132';Data Source = S02DB;Initial Catalog = BCNP"
vConnection.Open (ConnectionString)
Else
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=VBUser;Password='132';Data Source = S02DB;Initial Catalog = BCNP"
vConnection.Open (ConnectionString)
End If
End Sub


Public Function OpenDataBase(gConnection As ADODB.Connection, vRecordset As ADODB.Recordset, vQuery As String) As String
gConnection.CursorLocation = adUseClient
vRecordset.Open vQuery, gConnection, adOpenDynamic, adLockOptimistic
OpenDataBase = vRecordset.RecordCount
End Function
