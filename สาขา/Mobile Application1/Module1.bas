Attribute VB_Name = "Module1"
Option Explicit
Global ConnectionString As String
Global gConnection As New ADODB.Connection
Global vConnection As New ADODB.Connection
Global vCompany As String
Global vUserID As String
Global vPassword As String
Global vConnect As Integer
Global vDateCountStock As Date
Global vItemClick As Integer
Global vCheckValueNumber As Boolean

Global vRemStockOpen As Integer  ' เก็บเลขที่นับสต๊อกเพื่อตรวจสอบการบันทึกข้อมูล
Global vRemStockDocNo As String
Global vRemDocDate As String
Global vRemCountID As String
Global vRemWHCode As String
Global vRemStoreCode As String
Global vRemRowID As String

Public Sub InitializeDatabase()
Dim vQuery As String

If gConnection.State <> 0 Then
gConnection.Close
ConnectionString = "Provider = SQLOLEDB.1;Persist Security Info =False;User ID = vbuser;Password=132;Data Source  =S02DB;Initial Catalog = BCNP "
gConnection.Open (ConnectionString)
Else
ConnectionString = "Provider = SQLOLEDB.1;Persist Security Info =False;User ID = vbuser;Password=132;Data Source  =S02DB;Initial Catalog = BCNP "
gConnection.Open (ConnectionString)
End If
End Sub

Public Function OpenDataBase(gConnection As ADODB.Connection, vRecordset As ADODB.Recordset, vQuery As String) As String
gConnection.CursorLocation = adUseClient
vRecordset.Open vQuery, gConnection, adOpenDynamic, adLockOptimistic
OpenDataBase = vRecordset.RecordCount
End Function

Public Sub InitializeDataBase1()
If gConnection.State <> 0 Then
gConnection.Close
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=" & vUserID & ";Password=" & vPassword & ";Data Source = S02DB;Initial Catalog = BCNP"
gConnection.Open (ConnectionString)
Else
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=" & vUserID & ";Password=" & vPassword & ";Data Source = S02DB;Initial Catalog = BCNP"
gConnection.Open (ConnectionString)
End If
End Sub

Public Sub InitializeDataBase2()
If vConnection.State <> 0 Then
vConnection.Close
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=" & vUserID & ";Password=" & vPassword & ";Data Source = S02DB;Initial Catalog = BCNP"
vConnection.Open (ConnectionString)
Else
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=" & vUserID & ";Password=" & vPassword & ";Data Source = S02DB;Initial Catalog = BCNP"
vConnection.Open (ConnectionString)
End If
End Sub

Public Function CheckDegit(vCheckString As String) As Currency
    Dim vLenString As Integer
    Dim i As Integer
    Dim vString As String
    Dim vStrTemp As String
    Dim vStr1 As String
    
    vLenString = Len(vCheckString)
    For i = 1 To vLenString
        vString = Mid(vCheckString, i, 1)
        If Asc(vString) = 46 Or Asc(vString) = 47 Or Asc(vString) = 48 Or Asc(vString) = 49 Or Asc(vString) = 50 Or Asc(vString) = 51 Or Asc(vString) = 52 Or Asc(vString) = 53 Or Asc(vString) = 54 Or Asc(vString) = 55 Or Asc(vString) = 56 Or Asc(vString) = 57 Then
            vStr1 = vString
            vStrTemp = vStrTemp & vStr1
        Else
            vStr1 = ""
            vStrTemp = vStrTemp & vStr1
        End If
    Next i
    CheckDegit = vStrTemp
End Function

Public Sub CheckNumber(vData As String)
Dim vDocNo As String
Dim vText As String
Dim i As Integer

On Error GoTo ErrDescription

vDocNo = Trim(vData)

For i = 1 To Len(vData)
    If Mid(vDocNo, i, 1) = 0 Or Mid(vDocNo, i, 1) = 1 Or Mid(vDocNo, i, 1) = 2 Or Mid(vDocNo, i, 1) = 3 Or Mid(vDocNo, i, 1) = 4 Or Mid(vDocNo, i, 1) = 5 Or Mid(vDocNo, i, 1) = 6 Or Mid(vDocNo, i, 1) = 7 Or Mid(vDocNo, i, 1) = 8 Or Mid(vDocNo, i, 1) = 9 Or Mid(vDocNo, i, 1) = "." Then
        vCheckValueNumber = True
    Else
        vCheckValueNumber = False
        Exit Sub
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Function CheckDot(vData As String) As Integer
Dim vDocNo As String
Dim vText As String
Dim i As Integer
Dim vDot As Integer

On Error GoTo ErrDescription

vDocNo = Trim(vData)

For i = 1 To Len(vData)
    If Mid(vDocNo, i, 1) = "." Then
        vDot = vDot + 1
    End If
Next i
CheckDot = vDot

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Function
End If
End Function
