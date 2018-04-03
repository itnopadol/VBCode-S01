Attribute VB_Name = "Module1"
Option Explicit
Global gConnection As New ADODB.Connection
Global gRecordset As New ADODB.Recordset
Global vConnection As New ADODB.Connection
Global ConnectionString As String
Global vCompany1 As String, vUserID1 As String, vPassword1 As String
Global vCompany As String, vUserID As String, vPassword As String
Global conn As New ADODB.Connection
Global rs As New ADODB.Recordset
Global Sqlstr As String
Global vMemServer As String
Global vMemDatabase As String
Global vMemUserID As String
Global vMemPassword As String


Public Sub InitializeDataBase1()
If gConnection.State = 0 Then
    ConnectionString = "Provider = SQLOLEDB.1;Persist Security Info= false;User ID = vbuser;Password = 132;Data Source = bi;Initial Catalog = bcvat"
    gConnection.Open (ConnectionString)
ElseIf gConnection.State = 1 Then
    gConnection.Close
    ConnectionString = "Provider = SQLOLEDB.1;Persist Security Info= false;User ID = vbuser;Password = 132;Data Source = bi;Initial Catalog = bcvat "
    gConnection.Open (ConnectionString)
End If
End Sub

Public Sub InitializeGetDataBase()
If gConnection.State = 0 Then
    ConnectionString = "Provider = SQLOLEDB.1;Persist Security Info= false;User ID = " & vMemUserID & ";Password = " & vMemPassword & ";Data Source = " & vMemServer & ";Initial Catalog = " & vMemDatabase & " "
    gConnection.Open (ConnectionString)
ElseIf gConnection.State = 1 Then
    gConnection.Close
    ConnectionString = "Provider = SQLOLEDB.1;Persist Security Info= false;User ID = " & vMemUserID & ";Password = " & vMemPassword & ";Data Source = " & vMemServer & ";Initial Catalog = " & vMemDatabase & " "
    gConnection.Open (ConnectionString)
End If
End Sub

Public Function OpenDataBase(gConnection As ADODB.Connection, gRecordset As ADODB.Recordset, vQuery As String) As String
gConnection.CursorLocation = adUseClient
gRecordset.Open vQuery, gConnection, adOpenDynamic, adLockOptimistic
OpenDataBase = gRecordset.RecordCount
End Function

Public Sub InitializeDataBase()
If vConnection.State <> 0 Then
vConnection.Close
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=vbuser;Password=132;Data Source = bi;Initial Catalog = bcvat  "
vConnection.Open (ConnectionString)
Else
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=vbuser;Password=132;Data Source = bi;Initial Catalog = bcvat "
vConnection.Open (ConnectionString)
End If
End Sub


Public Sub GetConnect()
If conn.State = 0 Then
    conn.Provider = "SQLOLEDB.1"
    conn.Properties("Persist Security Info") = False
    conn.Properties("User ID") = ""
    conn.Properties("Password") = ""
    conn.Properties("Data Source") = "Nebula"
    conn.Properties("Initial Catalog") = "BCNP"
    conn.CursorLocation = adUseClient
    conn.Open
Else
    conn.Close
    conn.Provider = "SQLOLEDB.1"
    conn.Properties("Persist Security Info") = False
    conn.Properties("User ID") = ""
    conn.Properties("Password") = ""
    conn.Properties("Data Source") = "Nebula"
    conn.Properties("Initial Catalog") = "BCNP"
    conn.CursorLocation = adUseClient
    conn.Open

End If
End Sub


Public Sub ReportSetLocation(pCrystalReport As CrystalReport)
    Dim vStop As Integer
    Dim i As Integer
    Dim gDefaultDatabaseReport As String
    Dim gCurrentDatabaseName As String
    Dim vName As String
     
    vStop = pCrystalReport.RetrieveDataFiles - 1
    gDefaultDatabaseReport = "BCVAT47"
    gCurrentDatabaseName = vCompany1
    vName = pCrystalReport.DataFiles(0)
    For i = 0 To vStop
        pCrystalReport.DataFiles(i) = ReplaceString(pCrystalReport.DataFiles(i), _
        gDefaultDatabaseReport, gCurrentDatabaseName)
    Next
End Sub

Public Function ReplaceString(pSourceString As String, pReplaceString As String, pWithString As String) As String
    Dim vPosition As Integer
    Dim vReturnValue As String
    Dim vTemp As String
    Dim vLengthReplaceString As Integer
    
    vLengthReplaceString = Len(pReplaceString)
    vTemp = pSourceString
    vReturnValue = ""
    Do Until vTemp = ""
        vPosition = InStr(1, vTemp, pReplaceString, vbTextCompare)
        If vPosition = 0 Then
            vReturnValue = vReturnValue & vTemp
            vTemp = ""
        Else
            vReturnValue = vReturnValue & Left(vTemp, vPosition - 1) & pWithString
            vTemp = Right(vTemp, Len(vTemp) - (vPosition + vLengthReplaceString - 1))
        End If
    Loop
    ReplaceString = vReturnValue
End Function

