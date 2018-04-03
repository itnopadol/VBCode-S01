Attribute VB_Name = "Module1"
Option Explicit
Global ConnectionString As String
Global qConnection As New ADODB.Connection
Global vConnection As New ADODB.Connection
Global sConnection As New ADODB.Connection
Global gConnection1 As New ADODB.Connection
Global gConnection2 As New ADODB.Connection
Global gConnection3 As New ADODB.Connection
Global vCompany As String
Global vUserID As String
Global vPassword As String
Global vPrintDocno As String
Global vIndexFinish As Integer
Global vIndexComplete As Integer
Global vIndexBegin As Integer
Global vTimeID As Integer
Global vCheckClickListview As Integer
Global vSelectZoneID As Integer
Global vRefNoReceive As String
Global vWHCodeReceive As String
Global vInvoiceNoReceive As String

Global vDepartment As String
Global vUserAuthority As Integer

Public Sub InitializeDataBase()
If vConnection.State <> 0 Then
  vConnection.Close
  ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=" & vUserID & ";Password=" & vPassword & ";Data Source = Nebula;Initial Catalog = BCNP"
  vConnection.Open (ConnectionString)
Else
  ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=" & vUserID & ";Password=" & vPassword & ";Data Source = Nebula;Initial Catalog = BCNP"
  vConnection.Open (ConnectionString)
End If
End Sub

Public Sub InitializeDataBase1()
If sConnection.State <> 0 Then
  sConnection.Close
  ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=vbuser;Password=132;Data Source = Nebula;Initial Catalog = BCNP"
  sConnection.Open (ConnectionString)
Else
  ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=vbuser;Password=132;Data Source = Nebula;Initial Catalog = BCNP"
  sConnection.Open (ConnectionString)
End If
End Sub

Public Sub InitializeDataBase2()
If qConnection.State <> 0 Then
  qConnection.Close
  ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=vbuser;Password=132;Data Source = Nebula;Initial Catalog = BCNP"
  qConnection.Open (ConnectionString)
Else
  ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=vbuser;Password=132;Data Source = Nebula;Initial Catalog = BCNP"
  qConnection.Open (ConnectionString)
End If
End Sub

Public Sub InitializeConnectDataBase1()
If gConnection1.State <> 0 Then
  gConnection1.Close
  ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=vbuser;Password=132;Data Source = Nebula;Initial Catalog = BCNP"
  gConnection1.Open (ConnectionString)
Else
  ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=vbuser;Password=132;Data Source = Nebula;Initial Catalog = BCNP"
  gConnection1.Open (ConnectionString)
End If
End Sub

Public Function OpenDataBase(gConnection As ADODB.Connection, gRecordset As ADODB.Recordset, vQuery As String) As String
  gConnection.CursorLocation = adUseClient
  gRecordset.Open vQuery, vConnection, adOpenDynamic, adLockOptimistic
  OpenDataBase = gRecordset.RecordCount
End Function

Public Function OpenDataBase1(gConnection As ADODB.Connection, gRecordset As ADODB.Recordset, vQuery As String) As String
  gConnection.CursorLocation = adUseClient
  gRecordset.Open vQuery, sConnection, adOpenDynamic, adLockOptimistic
  OpenDataBase1 = gRecordset.RecordCount
End Function
Public Function OpenDataBase2(gConnection As ADODB.Connection, gRecordset As ADODB.Recordset, vQuery As String) As String
  gConnection.CursorLocation = adUseClient
  gRecordset.Open vQuery, qConnection, adOpenDynamic, adLockOptimistic
  OpenDataBase2 = gRecordset.RecordCount
End Function

Public Function OpenGetDataBase1(gConnection As ADODB.Connection, gRecordset As ADODB.Recordset, vQuery As String) As String
  gConnection.CursorLocation = adUseClient
  gRecordset.Open vQuery, gConnection1, adOpenDynamic, adLockOptimistic
  OpenGetDataBase1 = gRecordset.RecordCount
End Function

Public Sub RefreshQueueFinish()
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vCount As Integer
Dim vPrinted As Integer
Dim vQuery As String


FrmQueue.ListView104.ListItems.Clear
vQuery = "exec dbo.USP_QM_SearchQueueFinishZone5 " & vSelectZoneID & " "
If OpenDataBase2(vConnection, vRecordset, vQuery) <> 0 Then
    While Not vRecordset.EOF
        Set vListItem = FrmQueue.ListView104.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        vListItem.SubItems(1) = Trim(vRecordset.Fields("arname").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("saleorderno").Value) 'timeid
        vListItem.SubItems(3) = Trim(vRecordset.Fields("picker").Value)
        vListItem.SubItems(4) = Trim(vRecordset.Fields("timeid").Value)
        vListItem.SubItems(5) = Trim(vRecordset.Fields("salename").Value)
        vListItem.SubItems(6) = Trim(vRecordset.Fields("doctype").Value)
        vListItem.SubItems(7) = Trim(vRecordset.Fields("customerzone").Value)
        vListItem.SubItems(8) = Trim(vRecordset.Fields("pickingstatus").Value)
        vListItem.SubItems(9) = Trim(vRecordset.Fields("docdate").Value)
        vListItem.SubItems(10) = Trim(vRecordset.Fields("timepick").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub

Public Sub RefreshQueueBegin()
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim vCount As Integer
Dim vPrinted As Integer
Dim vQuery As String
Dim vARName As String
Dim vRefNo As String
Dim vSaleCode As String
Dim vCustomerZone As String
Dim vQueueDate As String

FrmQueue.ListView101.ListItems.Clear
vQuery = "exec dbo.USP_QM_SearchQueueZone5   " & vSelectZoneID & " "
If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
  If vCheckClickListview <> 2 Then
    vARName = Trim(vRecordset.Fields("arname").Value)
    vRefNo = Trim(vRecordset.Fields("saleorderno").Value)
    vSaleCode = Trim(vRecordset.Fields("salename").Value)
    vQueueDate = Trim(vRecordset.Fields("docdate").Value)
     Select Case Trim(vRecordset.Fields("customerzone").Value)
     Case 0
       vCustomerZone = "ลูกค้ารอรับของตามจุดออกใบหยิบ"
     Case 1
       vCustomerZone = "ลูกค้ารอรับของฝั่ง : สำนักงานใหญ่ "
     Case 2
       vCustomerZone = "ลูกค้ารอรับของฝั่ง : OutLet"
     End Select
 
    FrmQueue.LBLQueueDate.Caption = vQueueDate
    FrmQueue.LBLARName.Caption = vARName
    FrmQueue.LBLRefNo.Caption = vRefNo
    FrmQueue.LBLUserPick.Caption = Trim("-")
    FrmQueue.LBLSale.Caption = vSaleCode
    FrmQueue.LBLCustomerZone.Caption = vCustomerZone
  End If
    While Not vRecordset.EOF
        Set vListItem = FrmQueue.ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        vListItem.SubItems(1) = Trim(vRecordset.Fields("salename").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("requesttime").Value)
        vListItem.SubItems(3) = Trim(vRecordset.Fields("saleorderno").Value)
        vListItem.SubItems(4) = Trim(vRecordset.Fields("doctype").Value)
        vListItem.SubItems(5) = Trim(vRecordset.Fields("arname").Value)
        vListItem.SubItems(6) = Trim(vRecordset.Fields("timeid").Value)
        vListItem.SubItems(7) = Trim(vRecordset.Fields("customerzone").Value)
        vListItem.SubItems(8) = Trim(vRecordset.Fields("docdate").Value)
        vListItem.SubItems(9) = Trim(vRecordset.Fields("timepick").Value)
        vListItem.SubItems(10) = Trim(vRecordset.Fields("zoneid").Value)
        vListItem.SubItems(11) = Trim(vRecordset.Fields("familycode").Value)
        vListItem.SubItems(12) = Trim(vRecordset.Fields("whcode").Value)
        vListItem.SubItems(13) = Trim(vRecordset.Fields("shelfgroup").Value)
        vListItem.SubItems(14) = Trim(vRecordset.Fields("pickzone").Value)
    vRecordset.MoveNext
    Wend
Else
    FrmQueue.LBLQueueDate.Caption = ""
    FrmQueue.LBLARName.Caption = ""
    FrmQueue.LBLRefNo.Caption = ""
    FrmQueue.LBLUserPick.Caption = ""
    FrmQueue.LBLSale.Caption = ""
    FrmQueue.LBLCustomerZone.Caption = ""
End If
vRecordset.Close
End Sub

Public Sub RefreshQueuePicking()
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim vCount As Integer
Dim vPrinted As Integer
Dim vQuery As String
Dim vStartTime As String
Dim vARName As String
Dim vRefNo As String
Dim vPicker As String
Dim vSaleCode As String
Dim vQueueDate As String

FrmQueue.ListView102.ListItems.Clear
vQuery = "exec dbo.USP_QM_SearchQueuePickingZone5   " & vSelectZoneID & " "
If OpenDataBase(qConnection, vRecordset, vQuery) <> 0 Then
    vCount = 1
  If vCheckClickListview = 2 Then
    vQueueDate = Trim(vRecordset.Fields("docdate").Value)
    vARName = Trim(vRecordset.Fields("arname").Value)
    vRefNo = Trim(vRecordset.Fields("saleorderno").Value)
    vPicker = Trim(vRecordset.Fields("picker").Value)
    vSaleCode = Trim(vRecordset.Fields("salename").Value)
    
    FrmQueue.LBLQueueDate.Caption = vQueueDate
    FrmQueue.LBLARName.Caption = vARName
    FrmQueue.LBLRefNo.Caption = vRefNo
    FrmQueue.LBLUserPick.Caption = vPicker
    FrmQueue.LBLSale.Caption = vSaleCode
  End If
    While Not vRecordset.EOF
        Set vListItem = FrmQueue.ListView102.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        vListItem.SubItems(1) = Trim(vRecordset.Fields("starttime").Value)
        vListItem.SubItems(9) = Trim(vRecordset.Fields("startdatetime").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("pickingtime").Value)
        'vListItem.SubItems(2) = Format(-1 * DateDiff("s", Now, vListItem.SubItems(9)) / 60, "##,##0.00")
        vListItem.SubItems(3) = Trim(vRecordset.Fields("saleorderno").Value)
        vListItem.SubItems(4) = Trim(vRecordset.Fields("doctype").Value)
        vListItem.SubItems(5) = Trim(vRecordset.Fields("arname").Value)
        vListItem.SubItems(6) = Trim(vRecordset.Fields("timeid").Value)
        vListItem.SubItems(7) = Trim(vRecordset.Fields("picker").Value)
        vListItem.SubItems(8) = Trim(vRecordset.Fields("salename").Value)
        vListItem.SubItems(10) = Trim(vRecordset.Fields("customerzone").Value)
        vListItem.SubItems(11) = Trim(vRecordset.Fields("docdate").Value)
        vListItem.SubItems(12) = Trim(vRecordset.Fields("timepick").Value)
        
        vStartTime = FrmQueue.ListView102.ListItems.Item(vCount).SubItems(9)
        'FrmQueue.ListView102.ListItems.Item(vCount).SubItems(2) = Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")
  
        If CCur(Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")) >= 15 Then
          FrmQueue.ListView102.ListItems(vCount).ForeColor = "&H000000FF"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(1).ForeColor = "&H000000FF" 'red
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(2).ForeColor = "&H000000FF"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(3).ForeColor = "&H000000FF"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(4).ForeColor = "&H000000FF"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(5).ForeColor = "&H000000FF"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(6).ForeColor = "&H000000FF"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(7).ForeColor = "&H000000FF"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(8).ForeColor = "&H000000FF"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(9).ForeColor = "&H000000FF"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(10).ForeColor = "&H000000FF"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(11).ForeColor = "&H000000FF"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(12).ForeColor = "&H000000FF"
        End If

        If CCur(Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")) >= 10 And CCur(Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")) < 15 Then 'dark blue
          FrmQueue.ListView102.ListItems(vCount).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(1).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(2).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(3).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(4).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(5).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(6).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(7).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(8).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(9).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(10).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(11).ForeColor = "&H00FF0000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(12).ForeColor = "&H00FF0000"
        End If

        If CCur(Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")) >= 5 And CCur(Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")) < 10 Then 'green
          FrmQueue.ListView102.ListItems(vCount).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(1).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(2).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(3).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(4).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(5).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(6).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(7).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(8).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(9).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(10).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(11).ForeColor = "&H00008000"
          FrmQueue.ListView102.ListItems.Item(vCount).ListSubItems(12).ForeColor = "&H00008000"
        End If
    
    vCount = vCount + 1
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub

Public Sub ChekAuthorityAccess()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

vQuery = "exec USP_NP_AccessProgram '" & vUserID & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) > 0 Then
    vDepartment = Trim(vRecordset.Fields("departmentcode").Value)
    vUserAuthority = Trim(vRecordset.Fields("pagestatus").Value)
Else
    vUserAuthority = 0
End If
vRecordset.Close

End Sub
