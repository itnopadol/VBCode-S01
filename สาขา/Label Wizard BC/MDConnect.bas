Attribute VB_Name = "MDConnect"
Option Explicit

Global ConnAccess As New ADODB.Connection        ' Connection Access
Global ConnSQL As New ADODB.Connection        ' Connection SQL Server
Global ConnDEV As New ADODB.Connection
Global strUsername, strPassword As String
Global SPrice() As String         ' Delaration Array SPrice For Many Module
Global tmpItemNumber, tmpBarcod, tmpItemDesc, tmpUOFM, tmpPrice, tmpSPrice, tmpWHCode As String, tmpShelfCode As String
Global IntStep As Integer, IntStep1 As Integer           ' Count Step
Global Rs1 As New ADODB.Recordset
Global Rs2 As New ADODB.Recordset
Global Rs3 As New ADODB.Recordset
Global Rs4 As New ADODB.Recordset
Global Rs5 As New ADODB.Recordset
Public Const dbTipPath = "\\S02DB\BCS\Doc\dbTip.mdb"
Global vConnectionString As String
Global vConnection As New ADODB.Connection
Global vCheckRecProduct As Integer


'Public Const dbTipPath = "C:\dbTip.mdb"
Public Sub ConnTipDB()
   On Error GoTo error1:
    With ConnAccess
        If .State = adStateOpen Then .Close
                .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = '" & dbTipPath & "'"
                .Open
    End With
    Exit Sub
' Error Handings
error1:
    MsgBox Err.Number & " - " & Err.Description
    End
End Sub
Public Sub ConnectSQL()
    With ConnSQL
        If .State = adStateOpen Then .Close
                .Provider = "SQLOLEDB"
                .Properties("Persist Security Info").Value = False
                .Properties("User ID").Value = strUsername
                .Properties("Password").Value = strPassword
                .Properties("Data Source").Value = "S02DB" '"GALAXY"
                .Properties("Initial Catalog").Value = "BCNP" 'เปลี่ยนฐานข้อมูลเป็น BCNP
                .CursorLocation = adUseClient
                .Open
    End With
End Sub
Public Sub ConnectNPDEV()
    With ConnDEV
        If .State = adStateOpen Then .Close
                .Provider = "SQLOLEDB"
                .Properties("Persist Security Info").Value = False
                .Properties("User ID").Value = strUsername
                .Properties("Password").Value = strPassword
                .Properties("Data Source").Value = "S02DB" '"GALAXY"
                .Properties("Initial Catalog").Value = "NPDEV"
                .CursorLocation = adUseClient
                .Open
    End With
End Sub
Public Sub Main()
    frmWizard.Show
    frmWizard.txtUsername.SetFocus
End Sub

Public Sub InitializeDatabase()
    If vConnection.State = 0 Then
        vConnectionString = "Provider = SQLOLEDB.1;Persist Security Info = false;User ID = " & strUsername & ";Password = " & strPassword & ";Data Source = S02DB;Initial Catalog = BCNP"
        vConnection.Open (vConnectionString)
    Else
        vConnection.Close
        vConnectionString = "Provider = SQLOLEDB.1;Persist Security Info = false;User ID = " & strUsername & ";Password = " & strPassword & ";Data Source = S02DB;Initial Catalog = BCNP"
        vConnection.Open (vConnectionString)
    End If
End Sub

Public Function OpenDatabase(gConnection As ADODB.Connection, gRecordset As ADODB.Recordset, vQuery As String) As String
gConnection.CursorLocation = adUseClient
gRecordset.Open vQuery, gConnection, adOpenDynamic, adLockOptimistic
OpenDatabase = gRecordset.RecordCount
End Function


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
