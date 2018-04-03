Attribute VB_Name = "Module1"
Option Explicit
Global gConnectionString As String
Global vUserID As String
Global vPassword As String
Global gConnection As New ADODB.Connection
Global vConnection As New ADODB.Connection
Global bConnection As New ADODB.Connection
Global posConnection As New ADODB.Connection
Global vCheckJob As Integer
Global vCheckJob1 As Integer
Global vMemberDiscount As Integer
Global vMemCommand As Integer
Global vCheckUsedID As Integer
Global vCheckSetFocus As Integer
Global vCountListItem As Integer
Global vGetConnect As New ADODB.Connection
Global vGetconnectString As String
Global vCheckSelect1 As Integer
Global vCheckSelect2 As Integer
Global vCheckSelect3 As Integer
Global vFormName As String
Global vCheckStatusPrint  As Integer
Global vAccess As Integer
Global vFormActivate As String
Global vIsOpen1 As Integer

Global vMemIsExpire As Integer

Global rs As New ADODB.Recordset


Public Sub InitializeDatabase()
If gConnection.State = 0 Then
    gConnectionString = "Provider = SQLOLEDB.1;Persist Security Info = False;User ID = " & vUserID & ";password = " & vPassword & ";Data Source = NEBULA;Initial Catalog = BCNP "
    gConnection.Open (gConnectionString)
ElseIf gConnection.State = 1 Then
    gConnection.Close
    gConnectionString = "Provider = SQLOLEDB.1;Persist Security Info = False;User ID = " & vUserID & ";password = " & vPassword & ";Data Source = NEBULA;Initial Catalog = BCNP "
    gConnection.Open (gConnectionString)
End If
End Sub


Public Sub InitializeDatabasePOS()
If posConnection.State = 0 Then
    gConnectionString = "Provider = SQLOLEDB.1;Persist Security Info = False;User ID = sa;password = [ibdkifu;Data Source = s01;Initial Catalog = POS"
    posConnection.Open (gConnectionString)
ElseIf posConnection.State = 1 Then
    posConnection.Close
    gConnectionString = "Provider = SQLOLEDB.1;Persist Security Info = False;User ID = sa;password = [ibdkifu;Data Source = s01;Initial Catalog = POS"
    posConnection.Open (gConnectionString)
End If
End Sub

Public Sub InitializeDatabaseBranch()
If bConnection.State = 0 Then
    gConnectionString = "Provider = SQLOLEDB.1;Persist Security Info = False;User ID = " & vUserID & ";password = " & vPassword & ";Data Source = S02DB;Initial Catalog = BCNP "
    bConnection.Open (gConnectionString)
ElseIf bConnection.State = 1 Then
    bConnection.Close
    gConnectionString = "Provider = SQLOLEDB.1;Persist Security Info = False;User ID = " & vUserID & ";password = " & vPassword & ";Data Source = S02DB;Initial Catalog = BCNP "
    bConnection.Open (gConnectionString)
End If
End Sub

'Public Function OpenDatabasePOS(posConnection As ADODB.Connection, posRecordset As ADODB.Recordset, vQuery As String) As String
'posConnection.CursorLocation = adUseClient
'posRecordset.Open vQuery, posConnection, adOpenDynamic, adLockOptimistic
'OpenDatabase = posRecordset.RecordCount
'End Function


Public Function OpenDatabase(gConnection As ADODB.Connection, gRecordset As ADODB.Recordset, vQuery As String) As String
gConnection.CursorLocation = adUseClient
gRecordset.Open vQuery, gConnection, adOpenDynamic, adLockOptimistic
OpenDatabase = gRecordset.RecordCount
End Function

Public Function OpenDatabaseBranch(bConnection As ADODB.Connection, bRecordset As ADODB.Recordset, vQuery As String) As String
bConnection.CursorLocation = adUseClient
bRecordset.Open vQuery, bConnection, adOpenDynamic, adLockOptimistic
OpenDatabaseBranch = bRecordset.RecordCount
End Function

Public Function OpenDataBase1(vConnection As ADODB.Connection, gRecordset1 As ADODB.Recordset, vQuery As String) As String
vConnection.CursorLocation = adUseClient
gRecordset1.Open vQuery, vConnection, adOpenDynamic, adLockOptimistic
OpenDataBase1 = gRecordset1.RecordCount
End Function

Public Function OpenDatabaseBPlus(vConnection As ADODB.Connection, gRecordset As ADODB.Recordset, vQuery As String) As String
vGetConnect.CursorLocation = adUseClient
gRecordset.Open vQuery, vConnection, adOpenDynamic, adLockOptimistic
OpenDatabaseBPlus = gRecordset.RecordCount
End Function


Public Sub InitializeSendEmail()
If vGetConnect.State = 0 Then
    vGetconnectString = "Provider = SQLOLEDB.1;Persist Security Info = False;User ID = vbuser;Password = 132;Data Source = Nebula;Initial Catalog = BCNP"
    vGetConnect.Open (vGetconnectString)
Else
    vGetConnect.Close
    vGetconnectString = "Provider = SQLOLEDB.1;Persist Security Info = False;User ID = vbuser;Password = 132;Data Source = Nebula;Initial Catalog = BCNP"
    vGetConnect.Open (vGetconnectString)
End If
End Sub

Public Sub ChekAuthorityAccess()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDepartment As String
Dim vLevelID As Integer
Dim vFormName As String
Dim vUserAuthority As Integer

vQuery = "select department ,levelid from bcnp.dbo.vw_np_UserAutorityProgram where code = '" & vUserID & "' "
If OpenDatabaseBPlus(vGetConnect, vRecordset, vQuery) <> 0 Then
    vDepartment = Trim(vRecordset.Fields("department").Value)
    vLevelID = Trim(vRecordset.Fields("levelid").Value)
    vUserAuthority = 1
Else
    vUserAuthority = 0
End If
vRecordset.Close

If vUserAuthority = 1 Then
vFormName = vFormActivate
vQuery = "select pagestatus from vw_np_AccessProgram where departmentcode = '" & vDepartment & "' and levelid = " & vLevelID & " and pagename = '" & vFormName & "' "
If OpenDatabaseBPlus(vGetConnect, vRecordset, vQuery) <> 0 Then
    vAccess = Trim(vRecordset.Fields("pagestatus").Value)
Else
    vAccess = 0
End If
vRecordset.Close
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
    If vStrTemp <> "" Then
    CheckDegit = vStrTemp
    Else
    CheckDegit = 0
    End If
End Function

Public Sub SelectItemPromo()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItem As ListItem
Dim vPromotion As String
Dim vSecMan As String
Dim i As Integer
Dim vPromoDocNo As String

    If Form103.Text101.Text <> "" Then
        Form103.ListView101.ListItems.Clear
        i = 0
        Form103.Check101.Value = 0
        vPromotion = Left(Trim(Form103.Text101.Text), InStr(Trim(Form103.Text101.Text), "/") - 1)
        vSecMan = Trim(Form103.Text102.Text)
        vPromoDocNo = Trim(Form103.Text104)
        Form103.ListView101.ListItems.Clear
        vQuery = "execute USP_PM_SelectItemPrintLabel '" & vPromotion & "','" & vSecMan & "','" & vPromoDocNo & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            Form103.Label103.Caption = vRecordset.RecordCount
            vCountListItem = vRecordset.RecordCount
            Form103.Text101.Text = Trim(vRecordset.Fields("pmname").Value)
            Form103.Text102.Text = Trim(vRecordset.Fields("secman").Value)
            Form103.Label101.Caption = Trim(vRecordset.Fields("datestart").Value)
            Form103.Label102.Caption = Trim(vRecordset.Fields("dateend").Value)
            vMemIsExpire = Trim(vRecordset.Fields("expire").Value)
            
            While Not vRecordset.EOF
                i = i + 1
                Set vListItem = Form103.ListView101.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
                If Not IsNull(Trim(vRecordset.Fields("barcode3").Value)) Then
                vListItem.SubItems(1) = Trim(vRecordset.Fields("barcode3").Value)
                Else
                vListItem.SubItems(1) = "รหัสสินค้ามีปัญหา ติดต่อผู้ดูแลสินค้า"
                End If
                vListItem.SubItems(2) = Trim(vRecordset.Fields("priceerect").Value)
                vListItem.SubItems(3) = Trim(vRecordset.Fields("price").Value)
                vListItem.SubItems(4) = Trim(vRecordset.Fields("promoprice").Value)
                vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
                vListItem.SubItems(6) = Trim(vRecordset.Fields("itemname").Value)
                vListItem.SubItems(7) = Trim(vRecordset.Fields("discount").Value)
                vListItem.SubItems(8) = Trim(vRecordset.Fields("ismember").Value)
                vListItem.SubItems(9) = Trim(vRecordset.Fields("docno").Value)
                If IsNull(Trim(vRecordset.Fields("barcode3").Value)) Then
                    Form103.ListView101.ListItems(i).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H000000FF"
                End If
            vRecordset.MoveNext
            Wend

        End If
        vRecordset.Close
        On Error Resume Next
        Form103.ListView101.SetFocus
Else
  MsgBox "กรุณากรอก ทะเบียนโปรโมชั่นด้วย", vbCritical, "Send Error Message"
End If
End Sub

Public Sub SelectItemPromoPrintLabel(vPMCode As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItem As ListItem
Dim vPromotion As String
Dim vSecMan As String
Dim i As Integer
Dim vPromoDocNo As String

    If Form103.Text104.Text <> "" Then
        Form103.ListView101.ListItems.Clear
        i = 0
        Form103.Check101.Value = 0
        vPromotion = vPMCode 'Left(Trim(Form103.Text101.Text), InStr(Trim(Form103.Text101.Text), "/") - 1)
        vSecMan = Trim(Form103.Text102.Text)
        vPromoDocNo = Trim(Form103.Text104)
        Form103.ListView101.ListItems.Clear
        vQuery = "execute USP_PM_SelectItemPrintLabel '" & vPromotion & "','" & vSecMan & "','" & vPromoDocNo & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            Form103.Label103.Caption = vRecordset.RecordCount
            vCountListItem = vRecordset.RecordCount
            Form103.Text101.Text = Trim(vRecordset.Fields("pmname").Value)
            Form103.Text102.Text = Trim(vRecordset.Fields("secman").Value)
            Form103.Label101.Caption = Trim(vRecordset.Fields("datestart").Value)
            Form103.Label102.Caption = Trim(vRecordset.Fields("dateend").Value)
            vMemIsExpire = Trim(vRecordset.Fields("expire").Value)
            
            While Not vRecordset.EOF
                i = i + 1
                Set vListItem = Form103.ListView101.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
                If Not IsNull(Trim(vRecordset.Fields("barcode3").Value)) Then
                vListItem.SubItems(1) = Trim(vRecordset.Fields("barcode3").Value)
                Else
                vListItem.SubItems(1) = "รหัสสินค้ามีปัญหา ติดต่อผู้ดูแลสินค้า"
                End If
                vListItem.SubItems(2) = Trim(vRecordset.Fields("priceerect").Value)
                vListItem.SubItems(3) = Trim(vRecordset.Fields("price").Value)
                vListItem.SubItems(4) = Trim(vRecordset.Fields("promoprice").Value)
                vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
                vListItem.SubItems(6) = Trim(vRecordset.Fields("itemname").Value)
                vListItem.SubItems(7) = Trim(vRecordset.Fields("discount").Value)
                vListItem.SubItems(8) = Trim(vRecordset.Fields("ismember").Value)
                vListItem.SubItems(9) = Trim(vRecordset.Fields("docno").Value)
                If IsNull(Trim(vRecordset.Fields("barcode3").Value)) Then
                    Form103.ListView101.ListItems(i).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
                    Form103.ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H000000FF"
                End If
            vRecordset.MoveNext
            Wend

        End If
        vRecordset.Close
        On Error Resume Next
        Form103.ListView101.SetFocus
Else
  MsgBox "กรุณากรอก เลขที่เอกสารที่ต้องการพิมพ์ป้ายโปรโมชั่นด้วย", vbCritical, "Send Error Message"
End If
End Sub

