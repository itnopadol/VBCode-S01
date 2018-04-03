Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Net.NetworkInformation


Imports System
Imports System.Net.DNS
Imports System.Management
Imports System.Security
Imports System.Security.Principal.WindowsIdentity
Imports System.Net
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.ComponentModel
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.Text

Public Class frmNopadolTransDataToBranch
    Dim vQuery As String
    Dim vStrExecute As String
    Dim hostname As String
    Dim ipaddress As String
    Dim h As System.Net.IPHostEntry = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName)

    Dim mc As System.Management.ManagementClass
    Dim mo As ManagementObject

    Declare Function SendARP Lib "iphlpapi.dll" Alias "SendARP" (ByVal DestIP As Int32, ByVal SrcIP As Int32, ByVal pMacAddr() As Byte, ByRef PhyAddrLen As Int32) As Int32
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


    Dim vIsConnect As Integer

    Dim thread1 As System.Threading.Thread
    Dim thread2 As System.Threading.Thread
    Dim thread3 As System.Threading.Thread
    Dim thread4 As System.Threading.Thread
    Dim thread5 As System.Threading.Thread


    Dim daNEBULA1 As SqlDataAdapter
    Dim dsNEBULA1 As DataSet
    Dim dtNEBULA1 As DataTable
    Dim dvNEBULA1 As DataView
    Dim cmdNEBULA1 As SqlCommand

    Dim daNEBULA2 As SqlDataAdapter
    Dim dsNEBULA2 As DataSet
    Dim dtNEBULA2 As DataTable
    Dim dvNEBULA2 As DataView
    Dim cmdNEBULA2 As SqlCommand

    Dim daNEBULA3 As SqlDataAdapter
    Dim dsNEBULA3 As DataSet
    Dim dtNEBULA3 As DataTable
    Dim dvNEBULA3 As DataView
    Dim cmdNEBULA3 As SqlCommand


    Dim vConnectionStringNEBULA As String
    Dim vConnectionNEBULA As SqlConnection
    Dim daNEBULA As SqlDataAdapter
    Dim dsNEBULA As DataSet
    Dim dtNEBULA As DataTable
    Dim dvNEBULA As DataView
    Dim cmdNEBULA As SqlCommand

    Dim vConnectionStringS02DB As String
    Dim vConnectionS02DB As SqlConnection
    Dim daS02DB As SqlDataAdapter
    Dim dsS02DB As DataSet
    Dim dtS02DB As DataTable
    Dim dvS02DB As DataView
    Dim cmdS02DB As SqlCommand

    Dim daS02DB1 As SqlDataAdapter
    Dim dsS02DB1 As DataSet
    Dim dtS02DB1 As DataTable
    Dim dvS02DB1 As DataView
    Dim cmdS02DB1 As SqlCommand



    Dim vNEBULAServer As String
    Dim vNEBULADatabase As String
    Dim vNEBULAUserID As String
    Dim vNEBULAPassword As String

    Dim vS02DBServer As String
    Dim vS02DBDatabase As String
    Dim vS02DBUserID As String
    Dim vS02DBPassword As String

    Dim vTableName As String
    Dim vModuleID As String
    Dim vDocNo As String
    Dim vMemCancelStatus As Integer

    Dim vCheckExist As Integer
    Dim vDocSearchType As Integer
    Dim vLink As Boolean

    Dim vMemHeadOfficeIsCancel As Integer
    Dim vMemHeadOfficeIsConfirm As Integer
    Dim vMemHeadOfficeBillStatus As Integer
    Dim vMemHeadOfficeLastEditorCode As String
    Dim vMemHeadOfficeLastEditDateT As Date

    Dim vMemBranchIsCancel As Integer
    Dim vMemBranchIsConfirm As Integer
    Dim vMemBranchBillStatus As Integer
    Dim vMemBranchLastEditorCode As String
    Dim vMemBranchLastEditDateT As Date

    Dim vHeadOfficeExist As Integer
    Dim vBranchExist As Integer
    Dim vCheckError As Integer
    Dim vCountDepositUse As Integer
    Dim vSendTrnStatus As Integer

    Dim vTrnState As Integer
    Dim vIsTransfer As Integer
    Dim vMemBranchDepExist As Integer

    Dim vMemOfficeIsConnect As Integer
    Dim vMemBranchIsConnect As Integer

    Dim vMemBranchDocCancel As Integer
    Dim vMemBranchDeleteDepInvoice As Integer
    Dim vMemBranchDepInvoiceHaveUse As Integer


    Private Sub frmNopadolTransDataToBranch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        Dim c As Process = Process.GetCurrentProcess()
        Dim p As Process

        For Each p In Process.GetProcessesByName(c.ProcessName)
            If p.Id <> c.Id Then
                If p.MainModule.FileName = c.MainModule.FileName Then
                    Application.Exit()
                End If
            End If

        Next p

        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("th-TH")

        Dim keyName As String = Registry.CurrentUser.ToString() & "\Control Panel\International"
        Dim valueName As String = "sShortDate"
        Dim s As String = Registry.GetValue(keyName, valueName, String.Empty).ToString()
        Registry.SetValue(keyName, valueName, "dd/MM/yyyy")

        Me.WindowState = 1
        If (Me.WindowState = FormWindowState.Minimized) Then
            Me.Visible = False
            NotifyIcon1.Visible = True
        End If


        Me.CheckForIllegalCrossThreadCalls = False

        Me.PBActive.Visible = False
        Me.PBNotConnect.Visible = True
        Me.TBLink.Text = "ไม่สามารถติดต่อสาขาได้"
    End Sub

    Public Sub ConnectOffice()
        On Error GoTo ErrMyDescription

        vConnectionStringNEBULA = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =NEBULA;Initial Catalog =BCNP"
        vConnectionNEBULA = New SqlConnection(vConnectionStringNEBULA)
        vConnectionNEBULA.Open()

        vMemOfficeIsConnect = 1


ErrMyDescription:
        If Err.Description = "" Then
            vMemOfficeIsConnect = 1
        Else
            vMemOfficeIsConnect = 0
        End If
    End Sub

    Public Sub ConnectBranch()
        On Error GoTo ErrMyDescription

        vConnectionStringS02DB = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =192.168.2.2;Initial Catalog =BCNP; Connection Timeout = 180;"
        vConnectionS02DB = New SqlConnection(vConnectionStringS02DB)
        vConnectionS02DB.Open()

        vMemBranchIsConnect = 1

ErrMyDescription:
        If Err.Description = "" Then
            vMemBranchIsConnect = 1
        Else
            vMemBranchIsConnect = 0
        End If

    End Sub

    Public Sub CheckDataHeadOffice(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        On Error Resume Next

        vQuery = "exec dbo.USP_NP_CheckStatusDocHeadOffice " & vType & ",'" & vTableName & "','" & vValue & "'"
        daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA = New DataSet
        daNEBULA.Fill(dsNEBULA, "DataHeadOffice")
        dtNEBULA = dsNEBULA.Tables("DataHeadOffice")
        If dsNEBULA.Tables("DataHeadOffice").Rows.Count > 0 Then
            vHeadOfficeExist = 1
            vMemHeadOfficeIsCancel = dsNEBULA.Tables("DataHeadOffice").Rows(0).Item("iscancel")
            vMemHeadOfficeIsConfirm = dsNEBULA.Tables("DataHeadOffice").Rows(0).Item("isconfirm")
            vMemHeadOfficeBillStatus = dsNEBULA.Tables("DataHeadOffice").Rows(0).Item("billstatus")
            vMemHeadOfficeLastEditorCode = dsNEBULA.Tables("DataHeadOffice").Rows(0).Item("lasteditorcode")
            vMemHeadOfficeLastEditDateT = dsNEBULA.Tables("DataHeadOffice").Rows(0).Item("lasteditdatet")
        Else
            vHeadOfficeExist = 0
            vMemHeadOfficeIsCancel = 0
            vMemHeadOfficeIsConfirm = 0
            vMemHeadOfficeBillStatus = 0
            vMemHeadOfficeLastEditorCode = ""
            vMemHeadOfficeLastEditDateT = ""
        End If
    End Sub

    Public Sub CheckDataBranch(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        On Error Resume Next

        If vConnectionS02DB.State = ConnectionState.Closed Then
            Call ConnectBranch()
        End If

        vQuery = "exec dbo.USP_NP_CheckStatusDocBranch " & vType & ",'" & vTableName & "','" & vValue & "'"
        daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
        dsS02DB = New DataSet
        daS02DB.Fill(dsS02DB, "DataBranch")
        dtS02DB = dsS02DB.Tables("DataBranch")
        If dsS02DB.Tables("DataBranch").Rows.Count > 0 Then
            vBranchExist = 1
            vMemBranchIsCancel = dsS02DB.Tables("DataBranch").Rows(0).Item("iscancel")
            vMemBranchIsConfirm = dsS02DB.Tables("DataBranch").Rows(0).Item("isconfirm")
            vMemBranchBillStatus = dsS02DB.Tables("DataBranch").Rows(0).Item("billstatus")
            vMemBranchLastEditorCode = dsS02DB.Tables("DataBranch").Rows(0).Item("lasteditorcode")
            vMemBranchLastEditDateT = dsS02DB.Tables("DataBranch").Rows(0).Item("lasteditdatet")
        Else
            vBranchExist = 0
            vMemBranchIsCancel = 0
            vMemBranchIsConfirm = 0
            vMemBranchBillStatus = 0
            vMemBranchLastEditorCode = ""
            vMemBranchLastEditDateT = ""
        End If
    End Sub


    Public Sub CompareBranchDepUse(ByVal vInvNo As String, ByVal vDepNo As String, ByVal vInvAmount As Double)

        On Error Resume Next

        If vConnectionS02DB.State = ConnectionState.Closed Then
            Call ConnectBranch()
        End If

        vQuery = "exec dbo.USP_NP_CheckDataDepUse '" & vInvNo & "','" & vDepNo & "'," & vInvAmount & ""
        daS02DB1 = New SqlDataAdapter(vQuery, vConnectionS02DB)
        dsS02DB1 = New DataSet
        daS02DB1.Fill(dsS02DB1, "DataDep")
        dtS02DB1 = dsS02DB1.Tables("DataDep")

        If dsS02DB1.Tables("DataDep").Rows.Count > 0 Then
            vMemBranchDepExist = dsS02DB1.Tables("DataDep").Rows(0).Item("countdoc")
        End If
    End Sub


    Public Sub CheckDeleteDepositInvoice(ByVal vInvNo As String, ByVal vDepositNo As String)

        On Error Resume Next

        If vConnectionNEBULA.State = ConnectionState.Closed Then
            Call ConnectOffice()
        End If

        vQuery = "exec dbo.USP_PTF_CheckDataDepositINV '" & vInvNo & "','" & vDepositNo & "'"
        daNEBULA3 = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA3 = New DataSet
        daNEBULA3.Fill(dsNEBULA3, "DataDepAmount")
        dtNEBULA3 = dsNEBULA3.Tables("DataDepAmount")

        If dsNEBULA3.Tables("DataDepAmount").Rows.Count > 0 Then
            vMemBranchDeleteDepInvoice = dsNEBULA3.Tables("DataDepAmount").Rows(0).Item("sumofdeposit1")
            vMemBranchDepInvoiceHaveUse = dsNEBULA3.Tables("DataDepAmount").Rows(0).Item("vCount")
        End If
    End Sub


    Public Sub CheckCancelDataBranch(ByVal vInvNo As String)

        On Error Resume Next

        If vConnectionNEBULA.State = ConnectionState.Closed Then
            Call ConnectOffice()
        End If

        vQuery = "exec dbo.USP_PTF_CheckCancelDataTable '" & vInvNo & "'"
        daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA = New DataSet
        daNEBULA.Fill(dsNEBULA, "DataCancel")
        dtNEBULA = dsNEBULA.Tables("DataCancel")

        If dsNEBULA.Tables("DataCancel").Rows.Count > 0 Then
            vMemBranchDocCancel = dsNEBULA.Tables("DataCancel").Rows(0).Item("iscancel")
        End If
    End Sub



    Public Sub CheckData(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        On Error Resume Next

        vQuery = "exec dbo.USP_PTF_SearchDataTransfer " & vType & ",'" & vTableName & "','" & vValue & "'"
        daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA = New DataSet
        daNEBULA.Fill(dsNEBULA, "Data")
        dtNEBULA = dsNEBULA.Tables("Data")
        If dsNEBULA.Tables("Data").Rows.Count <= 0 Then
            vCheckExist = 0
        Else
            vCheckExist = 1
        End If
    End Sub

    Public Sub PrepareData(ByVal vType As Integer, ByVal vTableName As String, ByVal vValues As String)
        On Error Resume Next

        vQuery = "exec dbo.USP_PTF_PrepareDataTable " & vType & ",'" & vTableName & "','" & vValues & "'"
        cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
        cmdNEBULA.ExecuteNonQuery()

    End Sub

    Public Sub DeleteTableBranch(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        'On Error GoTo ErrDescription

        On Error Resume Next

        'vQuery = "begin tran"
        'cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        'cmdS02DB.ExecuteNonQuery()

        If vConnectionS02DB.State = ConnectionState.Closed Then
            Call ConnectBranch()
        End If

        vQuery = "exec dbo.USP_PTF_DeleteDataTable " & vType & ",'" & vTableName & "','" & vValue & "' "
        cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        cmdS02DB.ExecuteNonQuery()

        'vQuery = "commit tran"
        'cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        'cmdS02DB.ExecuteNonQuery()

        'ErrDescription:

    End Sub


    Public Sub DeleteDepUseBranch(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String, ByVal vInvAmount As Double)
        'On Error GoTo ErrDescription

        On Error Resume Next

        'vQuery = "begin tran"
        'cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        'cmdS02DB.ExecuteNonQuery()

        If vConnectionS02DB.State = ConnectionState.Closed Then
            Call ConnectBranch()
        End If

        If vConnectionNEBULA.State = ConnectionState.Closed Then
            Call ConnectOffice()
        End If

        vType = 6

        vQuery = "exec dbo.USP_PTF_DeleteDataTable " & vType & ",'" & vTableName & "','" & vValue & "' "
        cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        cmdS02DB.ExecuteNonQuery()

        vQuery = "exec dbo.USP_DP_InsertDepositUseLogs 'ลบข้อมูลสาขา','" & vTableName & "','" & vValue & "'," & vInvAmount & ""
        cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
        cmdNEBULA.ExecuteNonQuery()

        'vQuery = "commit tran"
        'cmdS02DB = New SqlCommand(vQuery, vConnectionnebula)
        'cmdS02DB.ExecuteNonQuery()

        'ErrDescription:

    End Sub

    Public Sub DropTable(ByVal vTableName As String)
        On Error Resume Next

        vQuery = "exec dbo.USP_PTF_DropDataTable '" & vTableName & "'"
        cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
        cmdNEBULA.ExecuteNonQuery()
    End Sub

    Public Sub InsertLogs(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        Dim vDatasource As String
        Dim vDestination As String

        On Error Resume Next


        If vType = 0 Then
            vDatasource = "NEBULA"
            vDestination = "S02DB"
        End If

        If vType = 1 Then
            vDatasource = "S02DB"
            vDestination = "NEBULA"
        End If

        vQuery = "exec dbo.USP_PTF_InsertTransferDataLogs '" & vDatasource & "','" & vDestination & "','" & vTableName & "','" & vValue & "','PrgTransfer',1"
        cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
        cmdNEBULA.ExecuteNonQuery()
    End Sub

    Public Sub InsertDepositDelDepUse(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        Dim vDatasource As String
        Dim vDestination As String

        On Error Resume Next

        If vType = 0 Then
            vDatasource = "NEBULA"
            vDestination = "S02DB"
        End If

        If vType = 1 Then
            vDatasource = "S02DB"
            vDestination = "NEBULA"
        End If

        vQuery = "exec dbo.USP_PTF_InsertTransferDataLogs '" & vDatasource & "','" & vDestination & "','" & vTableName & "','" & vValue & "','PrgTransfer',0"
        cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
        cmdNEBULA.ExecuteNonQuery()
    End Sub

    Public Sub TransferDoc(ByVal vGetDocNo As String)
        Dim vCountTransfer As Integer
        Dim i As Integer
        Dim vInvNo As String
        Dim vInvAmount As Double

        On Error Resume Next

        vIsTransfer = 1

        '===============================================================================================================
        '===============================================================================================================

        If vConnectionS02DB.State = ConnectionState.Closed Then
            Call ConnectBranch()
        End If

        '===============================================================================================================
        '===============================================================================================================
        vTableName = "BCSALEORDER"
        vDocSearchType = 0

        Call CheckDataHeadOffice(0, vTableName, vGetDocNo)
        Call CheckDataBranch(0, vTableName, vGetDocNo)

        Call CheckData(vDocSearchType, vTableName, vGetDocNo)
        If vCheckExist = 0 Then
            vSendTrnStatus = 2
            Exit Sub
        End If

        Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        vCountTransfer = 0
        vQuery = "select * from tempdb.dbo.BCSALEORDER where docno ='" & vGetDocNo & "'"
        daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA = New DataSet
        daNEBULA.Fill(dsNEBULA, "Data")
        dtNEBULA = dsNEBULA.Tables("Data")
        If dsNEBULA.Tables("Data").Rows.Count > 0 Then
            For i = 0 To dsNEBULA.Tables("Data").Rows.Count - 1
                vQuery = "exec dbo.USP_PTF_InsertSaleOrder '" & dsNEBULA.Tables("Data").Rows(i).Item("DocNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DocDate") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxType") & "," & dsNEBULA.Tables("Data").Rows(i).Item("BillType") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("ArCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("CreditDay") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DueDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DeliveryAddr") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxRate") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("RefDocNo") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("BillStatus") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SOStatus") & "," & dsNEBULA.Tables("Data").Rows(i).Item("HoldingStatus") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("CrAuthorizeMan") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ContactCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("NetAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCancel") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCompleteSave") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("RecurName") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("IsConditionSend") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("OwnReceive") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CarLicense") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ApproveCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ApproveDateTime") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("IsUseRobotSale") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TotalTransport") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumOtherValue") & "," & dsNEBULA.Tables("Data").Rows(i).Item("READYFORPAY") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("TimeTransport") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CarType") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CondPayCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("PrintCount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("PORefNo") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("IsImport") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("USERGROUP") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("METHODEPAYBILL") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("METHODEPAYBILL2") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("DELIVERYDAY") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DELIVERYDATE") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("DELIVERYSTATUS") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("PROMOTIONCODE") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
                cmdS02DB.ExecuteNonQuery()
                vTrnState = 1
            Next
        Else
            vTrnState = 0
        End If

        Call DropTable(vTableName)
        Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)

        vSendTrnStatus = 1

        '===============================================================================================================

        vTableName = "BCSALEORDERSUB"

        Call PrepareData(vDocSearchType, vTableName, vGetDocNo)

        vQuery = "select * from tempdb.dbo.BCSALEORDERSUB where docno ='" & vGetDocNo & "'"
        daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA = New DataSet
        daNEBULA.Fill(dsNEBULA, "DataSub")
        dtNEBULA = dsNEBULA.Tables("DataSub")
        If dsNEBULA.Tables("DataSub").Rows.Count > 0 Then
            For i = 0 To dsNEBULA.Tables("DataSub").Rows.Count - 1
                Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)
                vQuery = "exec dbo.USP_PTF_InsertSaleOrderSub '" & dsNEBULA.Tables("datasub").Rows(i).Item("DocNo") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("DepositNo") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("MyDescription") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("Balance") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("Amount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("DeposTaxType") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("CurrencyCode") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("DPExchangeRate") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("NewExchangeRate") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("HomeAmount1") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("HomeAmount2") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ExchangeProfit") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
                cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
                cmdS02DB.ExecuteNonQuery()
            Next
        End If

        Call DropTable(vTableName)

        vTrnState = 1
        vSendTrnStatus = 1

        Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        vSendTrnStatus = 1

        If vCheckError = 0 And vTrnState = 1 Then
            vQuery = "exec dbo.USP_NP_UpdateDepTransStatus '" & vGetDocNo & "' "
            cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
            cmdNEBULA.ExecuteNonQuery()
        End If


        '===============================================================================================================
        '===============================================================================================================

        vTableName = "BCARDEPOSIT"
        vDocSearchType = 0

        Call CheckDataHeadOffice(3, vTableName, vGetDocNo)
        Call CheckDataBranch(3, vTableName, vGetDocNo)

        Call CheckData(vDocSearchType, vTableName, vGetDocNo)
        If vCheckExist = 0 Then
            vSendTrnStatus = 2
            Exit Sub
        End If

        Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        vCountTransfer = 0
        vQuery = "select * from tempdb.dbo.BCARDEPOSIT where docno ='" & vGetDocNo & "'"
        daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA = New DataSet
        daNEBULA.Fill(dsNEBULA, "Data")
        dtNEBULA = dsNEBULA.Tables("Data")
        If dsNEBULA.Tables("Data").Rows.Count > 0 Then
            For i = 0 To dsNEBULA.Tables("Data").Rows.Count - 1
                vQuery = "exec dbo.USP_PTF_InsertARDeposit '" & dsNEBULA.Tables("Data").Rows(i).Item("DocNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DocDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("TaxDate") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("TaxNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ArCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DepartCode") & "',0,'" & dsNEBULA.Tables("Data").Rows(i).Item("DueDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("SaleCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxRate") & ",1,'" & dsNEBULA.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumOfWTax") & "," & dsNEBULA.Tables("Data").Rows(i).Item("NetAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("BillBalance") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ExcessAmount1") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ExcessAmount2") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ChargeAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("RefNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumCashAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumChqAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumCreditAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumBankAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("GLFormat") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("GLStartPosting") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsPostGL") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCancel") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsReturnMoney") & "," & dsNEBULA.Tables("Data").Rows(i).Item("GLTransData") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("RecurName") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("JobNo") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsNEBULA.Tables("Data").Rows(i).Item("OTHERINCOME") & "," & dsNEBULA.Tables("Data").Rows(i).Item("OTHEREXPENSE") & "," & dsNEBULA.Tables("Data").Rows(i).Item("CHANGEAMOUNT") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DepositNo") & "'"
                cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
                cmdS02DB.ExecuteNonQuery()
                vTrnState = 1
            Next
        Else
            vTrnState = 0
        End If

        Call DropTable(vTableName)
        Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)

        vSendTrnStatus = 1

        '=======================================================================================================================================================================================================

        vTableName = "BCARDEPOSITUSE"

        Call PrepareData(3, vTableName, vGetDocNo)



        vQuery = "select * from dbo.BCARDepositUse where depositno ='" & vGetDocNo & "'"
        daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
        dsS02DB = New DataSet
        daS02DB.Fill(dsS02DB, "DataSub")
        dtS02DB = dsS02DB.Tables("DataSub")
        If dsS02DB.Tables("DataSub").Rows.Count > 0 Then
            For i = 0 To dsS02DB.Tables("DataSub").Rows.Count - 1

                vInvNo = dsS02DB.Tables("DataSub").Rows(i).Item("DocNo")
                vInvAmount = dsS02DB.Tables("DataSub").Rows(i).Item("Amount")

                If vb6.Left(vInvNo, 3) = "S01" Or vb6.Left(vInvNo, 3) = "S01" Then

                    Call CheckCancelDataBranch(vInvNo)

                    If vMemBranchDocCancel = 1 Then
                        Call DeleteDepUseBranch(6, vGetDocNo, vInvNo, vInvAmount)
                    End If

                    Call CheckDeleteDepositInvoice(vInvNo, vGetDocNo)

                    If vMemBranchDeleteDepInvoice = 0 Then
                        Call DeleteDepUseBranch(6, vGetDocNo, vInvNo, vInvAmount)
                    ElseIf vMemBranchDeleteDepInvoice > 0 And vMemBranchDepInvoiceHaveUse = 0 Then
                        Call DeleteDepUseBranch(6, vGetDocNo, vInvNo, vInvAmount)
                    End If

                End If

            Next
        End If



        vCountTransfer = 0
        vQuery = "select * from tempdb.dbo.BCARDepositUse where depositno ='" & vGetDocNo & "'"
        daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA = New DataSet
        daNEBULA.Fill(dsNEBULA, "DataSub")
        dtNEBULA = dsNEBULA.Tables("DataSub")
        If dsNEBULA.Tables("DataSub").Rows.Count > 0 Then
            For i = 0 To dsNEBULA.Tables("DataSub").Rows.Count - 1

                vInvNo = dsNEBULA.Tables("datasub").Rows(i).Item("DocNo")
                vInvAmount = dsNEBULA.Tables("datasub").Rows(i).Item("Amount")

                Call CompareBranchDepUse(vInvNo, vGetDocNo, vInvAmount)

                If vMemBranchDepExist = 0 Then
                    Call DeleteDepUseBranch(6, vGetDocNo, vInvNo, vInvAmount)
                    vQuery = "exec dbo.USP_PTF_InsertARDepositUse '" & dsNEBULA.Tables("datasub").Rows(i).Item("DocNo") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("DepositNo") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("MyDescription") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("Balance") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("Amount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("DeposTaxType") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("CurrencyCode") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("DPExchangeRate") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("NewExchangeRate") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("HomeAmount1") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("HomeAmount2") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ExchangeProfit") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
                    cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
                    cmdS02DB.ExecuteNonQuery()

                    vQuery = "exec dbo.USP_DP_InsertDepositUseLogs 'เพิ่มข้อมูลสาขา','" & dsNEBULA.Tables("datasub").Rows(i).Item("DepositNo") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("DocNo") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("Amount") & ""
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()

                End If

            Next
        End If

        Call DropTable(vTableName)

        vTrnState = 1
        vSendTrnStatus = 1

        Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        vSendTrnStatus = 1

        If vCheckError = 0 And vTrnState = 1 Then
            vQuery = "exec dbo.USP_NP_UpdateDepTransStatus '" & vGetDocNo & "' "
            cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
            cmdNEBULA.ExecuteNonQuery()
        End If


        '===============================================================================================================
        '===============================================================================================================
        vIsTransfer = 0

LineEnd:

        'ErrDescription:
        '        If Err.Description <> "" Then
        '            MsgBox(Err.Description)
        '            vCheckError = 1
        '            Me.TMTransfer.Enabled = False
        '            Exit Sub
        '        Else
        '            vCheckError = 0
        '        End If
    End Sub

    Public Sub TransferData(ByVal vDepositNo As String)
        Dim vGetDocNo As String
        Dim vCountTransfer As Integer
        Dim i As Integer

        On Error Resume Next

        vGetDocNo = vDepositNo
        vTableName = "BCARDEPOSIT"
        vDocSearchType = 0

        Call CheckDataHeadOffice(3, vTableName, vGetDocNo)
        Call CheckDataBranch(3, vTableName, vGetDocNo)

        If vBranchExist > 0 Then
            If vMemBranchBillStatus <> 0 Then
                vSendTrnStatus = 2
            End If

            If vMemHeadOfficeIsCancel = 1 Then
                If vMemHeadOfficeIsCancel = vMemBranchIsCancel Then
                    vSendTrnStatus = 2
                End If
            End If
        End If

        Call CheckData(vDocSearchType, vTableName, vGetDocNo)
        If vCheckExist = 0 Then
            vSendTrnStatus = 2
            Exit Sub
        End If

        Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        vCountTransfer = 0
        vQuery = "select * from tempdb.dbo.BCARDEPOSIT where docno ='" & vGetDocNo & "'"
        daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA = New DataSet
        daNEBULA.Fill(dsNEBULA, "Data")
        dtNEBULA = dsNEBULA.Tables("Data")
        If dsNEBULA.Tables("Data").Rows.Count > 0 Then
            For i = 0 To dsNEBULA.Tables("Data").Rows.Count - 1
                vQuery = "exec dbo.USP_PTF_InsertARDeposit '" & dsNEBULA.Tables("Data").Rows(i).Item("DocNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DocDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("TaxDate") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("TaxNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ArCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DepartCode") & "',0,'" & dsNEBULA.Tables("Data").Rows(i).Item("DueDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("SaleCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxRate") & ",1,'" & dsNEBULA.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumOfWTax") & "," & dsNEBULA.Tables("Data").Rows(i).Item("NetAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("BillBalance") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ExcessAmount1") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ExcessAmount2") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ChargeAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("RefNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumCashAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumChqAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumCreditAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumBankAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("GLFormat") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("GLStartPosting") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsPostGL") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCancel") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsReturnMoney") & "," & dsNEBULA.Tables("Data").Rows(i).Item("GLTransData") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("RecurName") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("JobNo") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsNEBULA.Tables("Data").Rows(i).Item("OTHERINCOME") & "," & dsNEBULA.Tables("Data").Rows(i).Item("OTHEREXPENSE") & "," & dsNEBULA.Tables("Data").Rows(i).Item("CHANGEAMOUNT") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DepositNo") & "'"
                cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
                cmdS02DB.ExecuteNonQuery()
            Next
        End If
        Call DropTable(vTableName)
        Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        vSendTrnStatus = 1
    End Sub

    Public Sub CheckCountDepositUse(ByVal vDepositNo As String)
        On Error Resume Next

        vQuery = "exec dbo.USP_NP_CheckCountInvoiceUseDep '" & vDepositNo & "'"
        daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA = New DataSet
        daNEBULA.Fill(dsNEBULA, "Data")
        dtNEBULA = dsNEBULA.Tables("Data")
        If dsNEBULA.Tables("Data").Rows.Count > 0 Then
            vCountDepositUse = dsNEBULA.Tables("Data").Rows(0).Item("vcount")
        Else
            vCountDepositUse = 0
        End If

    End Sub

    Private Sub TMTransfer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TMTransfer.Tick
        On Error Resume Next

        If vMemOfficeIsConnect = 1 And vIsTransfer = 0 Then
            thread3 = New System.Threading.Thread(AddressOf SearchListDocTransfer)
            thread3.Start()
        End If

        'Dim i As Integer
        'Dim vListItem As ListViewItem
        'Dim n As Integer
        'Dim a As Integer

        ' '''On Error Resume Next

        'Call vConnctDataBase()

        ''vQuery = "exec dbo.USP_NP_SearchDocNotTransfer 0"
        'vQuery = "exec dbo.USP_DP_SearchDepNotTransfer 0"
        'daNEBULA1 = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        'dsNEBULA1 = New DataSet
        'daNEBULA1.Fill(dsNEBULA1, "Docno")
        'dtNEBULA1 = dsNEBULA1.Tables("Docno")
        'If dtNEBULA1.Rows.Count > 0 Then
        '    For i = 0 To dtNEBULA1.Rows.Count - 1
        '        vDocNo = dtNEBULA1.Rows(i).Item("depositno")

        '        Call TransferDoc(vDocNo)

        '        If vCheckError = 0 And vTrnState = 1 Then
        '            'vQuery = "exec dbo.USP_NP_UpdateDocTransfered '" & vDocNo & "'"
        '            vQuery = "exec dbo.USP_NP_UpdateDepTransStatus '" & vDocNo & "' "
        '            cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
        '            cmdNEBULA.ExecuteNonQuery()
        '        End If

        '    Next

        '    Me.ListViewListTrn.Items.Clear()

        '    vQuery = "exec dbo.USP_DP_SearchDepNotTransfer 1"
        '    daNEBULA2 = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA2 = New DataSet
        '    daNEBULA2.Fill(dsNEBULA2, "Docno")
        '    dtNEBULA2 = dsNEBULA2.Tables("Docno")
        '    If dtNEBULA2.Rows.Count > 0 Then
        '        For i = 0 To dtNEBULA2.Rows.Count - 1
        '            n = n + 1
        '            vListItem = Me.ListViewListTrn.Items.Add(n)
        '            vListItem.SubItems.Add(0).Text = dtNEBULA2.Rows(i).Item("depositno")
        '            vListItem.SubItems.Add(1).Text = dtNEBULA2.Rows(i).Item("trfdate")
        '        Next
        '    End If

        '    If Me.ListViewListTrn.Items.Count > 0 Then
        '        For a = 0 To Me.ListViewListTrn.Items.Count - 1
        '            If a Mod 2 = 0 Then
        '                Me.ListViewListTrn.Items(a).BackColor = Color.AliceBlue
        '            Else
        '                Me.ListViewListTrn.Items(a).BackColor = Color.White
        '            End If
        '        Next
        '    End If
        'End If

        ''ErrDescription:
        ''        If Err.Description <> "" Then
        ''            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        ''            Exit Sub
        ''        End If
    End Sub

    Public Sub SearchDocNotTransfer()
        Dim i As Integer

        On Error Resume Next

        If vConnectionNEBULA.State = ConnectionState.Closed Then
            Call ConnectOffice()
        End If

        If vMemOfficeIsConnect = 1 And vIsTransfer = 0 Then

            vQuery = "exec dbo.USP_DP_SearchDepNotTransfer 0"
            daNEBULA1 = New SqlDataAdapter(vQuery, vConnectionNEBULA)
            dsNEBULA1 = New DataSet
            daNEBULA1.Fill(dsNEBULA1, "Docno")
            dtNEBULA1 = dsNEBULA1.Tables("Docno")
            If dtNEBULA1.Rows.Count > 0 Then
                For i = 0 To dtNEBULA1.Rows.Count - 1
                    vDocNo = dtNEBULA1.Rows(i).Item("depositno")

                    Call TransferDoc(vDocNo)

                    'If vCheckError = 0 And vTrnState = 1 Then
                    '    vQuery = "exec dbo.USP_NP_UpdateDepTransStatus '" & vDocNo & "' "
                    '    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    '    cmdNEBULA.ExecuteNonQuery()
                    'End If

                Next
            End If

            If vIsTransfer = 0 Then
                vConnectionNEBULA.Close()
            End If
        End If
    End Sub


    Public Sub SearchListDocTransfer()
        Dim i As Integer
        Dim vListItem As ListViewItem
        Dim n As Integer
        Dim a As Integer

        On Error Resume Next

        If vConnectionNEBULA.State = ConnectionState.Closed Then
            Call ConnectOffice()
        End If

        If vMemOfficeIsConnect = 1 And vIsTransfer = 0 Then
            Me.ListViewListTrn.Items.Clear()

            vQuery = "exec dbo.USP_DP_SearchDepNotTransfer 1"
            daNEBULA3 = New SqlDataAdapter(vQuery, vConnectionNEBULA)
            dsNEBULA3 = New DataSet
            daNEBULA3.Fill(dsNEBULA3, "Docno")
            dtNEBULA3 = dsNEBULA3.Tables("Docno")
            If dtNEBULA3.Rows.Count > 0 Then
                For i = 0 To dtNEBULA3.Rows.Count - 1
                    n = n + 1
                    vListItem = Me.ListViewListTrn.Items.Add(n)
                    vListItem.SubItems.Add(0).Text = dtNEBULA3.Rows(i).Item("depositno")
                    vListItem.SubItems.Add(1).Text = dtNEBULA3.Rows(i).Item("trfdate")
                Next
            End If

            If vIsTransfer = 0 Then
                vConnectionNEBULA.Close()
            End If

            If Me.ListViewListTrn.Items.Count > 0 Then
                For a = 0 To Me.ListViewListTrn.Items.Count - 1
                    If a Mod 2 = 0 Then
                        Me.ListViewListTrn.Items(a).BackColor = Color.AliceBlue
                    Else
                        Me.ListViewListTrn.Items(a).BackColor = Color.White
                    End If
                Next
            End If
        End If
    End Sub

    Public Function PingAddress(ByVal ServerIP As String) As Boolean
        Dim ReplySuccess As Boolean
        Dim IpAddress As New System.Net.IPAddress(New Byte() {192, 168, 2, 2})

        On Error Resume Next

        If System.Net.IPAddress.TryParse(ServerIP, IpAddress) Then
            Dim png As New System.Net.NetworkInformation.Ping
            Dim reply As System.Net.NetworkInformation.PingReply
            Dim timeout As Integer
            Dim data() As Byte

            data = New Byte() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 28, 30, 31, 32}
            timeout = 10000
            reply = png.Send(IpAddress, timeout, data)
            ReplySuccess = (reply.Status = Net.NetworkInformation.IPStatus.Success)
        End If

        Return ReplySuccess

    End Function

    Private Sub Time_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Time.Tick
        Me.LBLTime.Text = Now.Hour & ":" & Now.Minute

        On Error Resume Next

        If Now.Hour > 7 And Now.Hour < 20 Then
            Me.TMTransfer.Enabled = True
            'Me.TBLink.Text = "โปรแกรมกำลังทำงาน"
            Me.TMActive.Enabled = True
        ElseIf Now.Hour > 7 And Now.Hour < 20 Then
            Me.TMTransfer.Enabled = True
            Me.TMActive.Enabled = True
        Else
            Me.TMTransfer.Enabled = False
            'Me.TBLink.Text = "โปรแกรมหยุดทำงาน"
            Me.TMActive.Enabled = False
            Me.PBActive.Visible = False
        End If

        If vMemBranchIsConnect = 1 Then
            Me.PBActive.Visible = True
            Me.PBNotConnect.Visible = False
            Me.TBLink.Text = "ติดต่อสาขาได้"
        End If

        If vMemBranchIsConnect = 0 Then
            Me.PBActive.Visible = False
            Me.PBNotConnect.Visible = True
            Me.TBLink.Text = "ไม่สามารถติดต่อสาขาได้"
        End If

    End Sub

    Private Sub TMActive_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TMActive.Tick
        On Error Resume Next

        If vMemBranchIsConnect = 1 Then
            Me.PBActive.Visible = True
            Me.PBNotConnect.Visible = False
            Me.TBLink.Text = "ติดต่อสาขาได้"
        Else
            Me.PBActive.Visible = False
            Me.PBNotConnect.Visible = True
            Me.TBLink.Text = "ไม่สามารถติดต่อสาขาได้"
        End If

    End Sub

    Public Sub vConnctDataBase()
        On Error Resume Next

        'vConnectionStringNEBULA = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Max Pool Size = 10000;Min Pool Size = 5;Data Source =NEBULA;Initial Catalog =BCNP"
        vConnectionStringNEBULA = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =NEBULA;Initial Catalog =BCNP"
        vConnectionNEBULA = New SqlConnection(vConnectionStringNEBULA)
        vConnectionNEBULA.Open()

        vConnectionStringS02DB = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Max Pool Size = 10000;Min Pool Size = 5;Data Source =192.168.2.2;Initial Catalog =BCNP"
        'vConnectionStringS02DB = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =192.168.0.15;Initial Catalog =BCNP"
        vConnectionS02DB = New SqlConnection(vConnectionStringS02DB)
        vConnectionS02DB.Open()
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        On Error Resume Next

        NotifyIcon1.Visible = False
        Me.Visible = True
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub BTNCloseProgram_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseProgram.Click
        Dim vAnswer As Integer

        On Error Resume Next

        vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่ การออกโปรแกรมต้องไม่อยู่ระหว่างการโอนมิฉะนั้นออกไม่ได้", MsgBoxStyle.YesNo, "Send Question Message")

        If vAnswer = 6 Then
            If vIsprocess = 0 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub BTNMinimize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNMinimize.Click
        On Error Resume Next

        Me.WindowState = 1
        If (Me.WindowState = FormWindowState.Minimized) Then
            Me.Visible = False
            NotifyIcon1.Visible = True
        End If
    End Sub

    Private Sub TConnect_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TConnect.Tick
        On Error Resume Next

        thread1 = New System.Threading.Thread(AddressOf ConnectOffice)
        thread1.Start()

        thread2 = New System.Threading.Thread(AddressOf ConnectBranch)
        thread2.Start()

        Me.TConnect.Enabled = False
    End Sub

    Private Sub TMNotTransfer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TMNotTransfer.Tick
        On Error Resume Next

        If vMemOfficeIsConnect = 1 And vMemBranchIsConnect = 1 And vIsTransfer = 0 Then
            thread3 = New System.Threading.Thread(AddressOf SearchDocNotTransfer)
            thread3.Start()
        End If
    End Sub

    Private Sub TCheckConnect_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TCheckConnect.Tick

        If vMemBranchIsConnect = 1 Then
            Me.PBActive.Visible = True
            Me.PBNotConnect.Visible = False
            Me.TBLink.Text = "ติดต่อสาขาได้"
        End If

        If vMemBranchIsConnect = 0 Then
            Me.PBActive.Visible = False
            Me.PBNotConnect.Visible = True
            Me.TBLink.Text = "ไม่สามารถติดต่อสาขาได้"
        End If

        thread2 = New System.Threading.Thread(AddressOf ConnectBranch)
        thread2.Start()
    End Sub
End Class
