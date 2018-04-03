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


Public Class FormBranchTransferToNopadol

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


    Dim daS02DB1 As SqlDataAdapter
    Dim dsS02DB1 As DataSet
    Dim dtS02DB1 As DataTable
    Dim dvS02DB1 As DataView
    Dim cmdS02DB1 As SqlCommand

    Dim daS02DB2 As SqlDataAdapter
    Dim dsS02DB2 As DataSet
    Dim dtS02DB2 As DataTable
    Dim dvS02DB2 As DataView
    Dim cmdS02DB2 As SqlCommand

    Dim daS02DB3 As SqlDataAdapter
    Dim dsS02DB3 As DataSet
    Dim dtS02DB3 As DataTable
    Dim dvS02DB3 As DataView
    Dim cmdS02DB3 As SqlCommand

    Dim vConnectionStringNEBULA As String
    Dim vConnectionNEBULA As SqlConnection
    Dim daNEBULA As SqlDataAdapter
    Dim dsNEBULA As DataSet
    Dim dtNEBULA As DataTable
    Dim dvNEBULA As DataView
    Dim cmdNEBULA As SqlCommand

    Dim daNEBULA1 As SqlDataAdapter
    Dim dsNEBULA1 As DataSet
    Dim dtNEBULA1 As DataTable
    Dim dvNEBULA1 As DataView
    Dim cmdNEBULA1 As SqlCommand


    Dim vConnectionStringS02DB As String
    Dim vConnectionS02DB As SqlConnection
    Dim daS02DB As SqlDataAdapter
    Dim dsS02DB As DataSet
    Dim dtS02DB As DataTable
    Dim dvS02DB As DataView
    Dim cmdS02DB As SqlCommand

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
    Dim vMemOfficeDocCancel As Integer
    Dim vMemOfficeDeleteDepInvoice As Integer
    Dim vMemOfficeDepInvoiceHaveUse As Integer


    Private Sub FormBranchTransferToNopadol_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'On Error Resume Next

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
        Me.TBLink.Text = "ไม่สามารถติดต่อสำนักงานใหญ่ได้"
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


        vConnectionStringS02DB = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =192.168.2.2;Initial Catalog =BCNP"
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
        'On Error Resume Next

        If vConnectionNEBULA.State = ConnectionState.Closed Then
            Call ConnectOffice()
        End If

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
        'On Error Resume Next

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

    Public Sub CheckData(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        'On Error Resume Next

        vQuery = "exec dbo.USP_PTF_SearchDataTransfer " & vType & ",'" & vTableName & "','" & vValue & "'"
        daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
        dsS02DB = New DataSet
        daS02DB.Fill(dsS02DB, "Data")
        dtS02DB = dsS02DB.Tables("Data")
        If dsS02DB.Tables("Data").Rows.Count <= 0 Then
            vCheckExist = 0
        Else
            vCheckExist = 1
        End If
    End Sub

    Public Sub PrepareData(ByVal vType As Integer, ByVal vTableName As String, ByVal vValues As String)
        'On Error Resume Next

        vQuery = "exec dbo.USP_PTF_PrepareDataTable " & vType & ",'" & vTableName & "','" & vValues & "'"
        cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        cmdS02DB.ExecuteNonQuery()

    End Sub

    Public Sub DeleteTableHeadOffice(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        'On Error Resume Next

        If vConnectionNEBULA.State = ConnectionState.Closed Then
            Call ConnectOffice()
        End If
        vQuery = "exec dbo.USP_PTF_DeleteDataTable " & vType & ",'" & vTableName & "','" & vValue & "' "
        cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
        cmdNEBULA.ExecuteNonQuery()
    End Sub

    Public Sub DropTableBranch(ByVal vTableName As String)
        'On Error Resume Next

        vQuery = "exec dbo.USP_PTF_DropDataTable '" & vTableName & "'"
        cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        cmdS02DB.ExecuteNonQuery()
    End Sub


    Public Sub InsertLogs(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        Dim vDatasource As String
        Dim vDestination As String

        'On Error Resume Next

        If vType = 0 Then
            vDatasource = "NEBULA"
            vDestination = "S02DB"
        End If

        If vType = 1 Then
            vDatasource = "S02DB"
            vDestination = "NEBULA"
        End If

        vQuery = "exec dbo.USP_PTF_InsertTransferDataLogs '" & vDatasource & "','" & vDestination & "','" & vTableName & "','" & vValue & "','PrgTransfer',1"
        cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        cmdS02DB.ExecuteNonQuery()
    End Sub


    Public Sub CheckCancelDataOffice(ByVal vInvNo As String)

        'On Error Resume Next

        If vConnectionS02DB.State = ConnectionState.Closed Then
            Call ConnectBranch()
        End If


        vQuery = "exec dbo.USP_PTF_CheckCancelDataTable '" & vInvNo & "'"
        daS02DB1 = New SqlDataAdapter(vQuery, vConnectionS02DB)
        dsS02DB1 = New DataSet
        daS02DB1.Fill(dsS02DB1, "DataCancel")
        dtS02DB1 = dsS02DB1.Tables("DataCancel")

        If dsS02DB1.Tables("DataCancel").Rows.Count > 0 Then
            vMemOfficeDocCancel = dsS02DB1.Tables("DataCancel").Rows(0).Item("iscancel")
        End If
    End Sub


    Public Sub TransferDoc(ByVal vGetDocNo As String)
        Dim vCountTransfer As Integer
        Dim i As Integer
        Dim vInvNo As String
        Dim vInvAmount As Double

        'On Error Resume Next

        vIsTransfer = 1

        If vConnectionNEBULA.State = ConnectionState.Closed Then
            Call ConnectOffice()
        End If

        '===============================================================================================================
        '===============================================================================================================

        vTableName = "BCARDEPOSIT"
        vDocSearchType = 0

        Call CheckDataBranch(3, vTableName, vGetDocNo)
        Call CheckDataHeadOffice(3, vTableName, vGetDocNo)


        Call CheckData(vDocSearchType, vTableName, vGetDocNo)
        If vCheckExist = 0 Then
            vSendTrnStatus = 2
            Exit Sub
        End If

        Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        Call DeleteTableHeadOffice(vDocSearchType, vTableName, vGetDocNo)

        vCountTransfer = 0
        vQuery = "select * from tempdb.dbo.BCARDeposit where docno ='" & vGetDocNo & "'"
        daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
        dsS02DB = New DataSet
        daS02DB.Fill(dsS02DB, "Data")
        dtS02DB = dsS02DB.Tables("Data")
        If dsS02DB.Tables("Data").Rows.Count > 0 Then
            For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                vQuery = "exec dbo.USP_PTF_InsertARDeposit '" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("TaxDate") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("TaxNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ArCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "',0,'" & dsS02DB.Tables("Data").Rows(i).Item("DueDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("SaleCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxRate") & ",1,'" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsS02DB.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumOfWTax") & "," & dsS02DB.Tables("Data").Rows(i).Item("NetAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("BillBalance") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExcessAmount1") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExcessAmount2") & "," & dsS02DB.Tables("Data").Rows(i).Item("ChargeAmount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("RefNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumCashAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumChqAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumCreditAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumBankAmount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("GLFormat") & "'," & dsS02DB.Tables("Data").Rows(i).Item("GLStartPosting") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsPostGL") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsReturnMoney") & "," & dsS02DB.Tables("Data").Rows(i).Item("GLTransData") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsS02DB.Tables("Data").Rows(i).Item("RecurName") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("JobNo") & "'," & dsS02DB.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("OTHERINCOME") & "," & dsS02DB.Tables("Data").Rows(i).Item("OTHEREXPENSE") & "," & dsS02DB.Tables("Data").Rows(i).Item("CHANGEAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DepositNo") & "'"
                cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                cmdNEBULA.ExecuteNonQuery()
                vTrnState = 1
            Next
        Else
            vTrnState = 0
        End If

        Call DropTableBranch(vTableName)
        Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        vSendTrnStatus = 1

        '=======================================================================================================================================================================================================

        vTableName = "BCARDEPOSITUSE"

        Call PrepareData(3, vTableName, vGetDocNo)



        vQuery = "select * from dbo.BCARDepositUse where depositno ='" & vGetDocNo & "'"
        daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA = New DataSet
        daNEBULA.Fill(dsNEBULA, "DataSub")
        dtNEBULA = dsNEBULA.Tables("DataSub")
        If dsNEBULA.Tables("DataSub").Rows.Count > 0 Then
            For i = 0 To dsNEBULA.Tables("DataSub").Rows.Count - 1

                vInvNo = dsNEBULA.Tables("DataSub").Rows(i).Item("DocNo")
                vInvAmount = dsNEBULA.Tables("DataSub").Rows(i).Item("Amount")

                If vb6.Left(vInvNo, 3) = "S02" Or vb6.Left(vInvNo, 3) = "W02" Then

                    Call CheckCancelDataOffice(vInvNo)

                    If vMemOfficeDocCancel = 1 Then
                        Call DeleteDepUseOffice(6, vGetDocNo, vInvNo, vInvAmount)
                    End If


                    Call CheckDeleteDepositInvoice(vInvNo, vGetDocNo)

                    If vMemOfficeDeleteDepInvoice = 0 Then
                        Call DeleteDepUseOffice(6, vGetDocNo, vInvNo, vInvAmount)
                    ElseIf vMemOfficeDeleteDepInvoice > 0 And vMemOfficeDepInvoiceHaveUse = 0 Then
                        Call DeleteDepUseOffice(6, vGetDocNo, vInvNo, vInvAmount)
                    End If
                End If

            Next
        End If


        vCountTransfer = 0
        vQuery = "select * from tempdb.dbo.BCARDepositUse where depositno ='" & vGetDocNo & "'"
        daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
        dsS02DB = New DataSet
        daS02DB.Fill(dsS02DB, "DataSub")
        dtS02DB = dsS02DB.Tables("DataSub")
        If dsS02DB.Tables("DataSub").Rows.Count > 0 Then
            For i = 0 To dsS02DB.Tables("DataSub").Rows.Count - 1

                vInvNo = dsS02DB.Tables("DataSub").Rows(i).Item("DocNo")
                vInvAmount = dsS02DB.Tables("DataSub").Rows(i).Item("Amount")

                Call CompareOfficeDepUse(vInvNo, vGetDocNo, vInvAmount)

                'Call CheckCancelDataOffice(vInvNo)

                'If vMemOfficeDocCancel = 1 Then
                '    Call DeleteDepUseOffice(6, vGetDocNo, vInvNo)
                'End If


                If vMemBranchDepExist = 0 Then
                    Call DeleteDepUseOffice(6, vGetDocNo, vInvNo, vInvAmount)
                    vQuery = "exec dbo.USP_PTF_InsertARDepositUse '" & dsS02DB.Tables("datasub").Rows(i).Item("DocNo") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("DepositNo") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("MyDescription") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("Balance") & "," & dsS02DB.Tables("datasub").Rows(i).Item("Amount") & "," & dsS02DB.Tables("datasub").Rows(i).Item("DeposTaxType") & "," & dsS02DB.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsS02DB.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("CurrencyCode") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("DPExchangeRate") & "," & dsS02DB.Tables("datasub").Rows(i).Item("NewExchangeRate") & "," & dsS02DB.Tables("datasub").Rows(i).Item("HomeAmount1") & "," & dsS02DB.Tables("datasub").Rows(i).Item("HomeAmount2") & "," & dsS02DB.Tables("datasub").Rows(i).Item("ExchangeProfit") & "," & dsS02DB.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()

                    vQuery = "exec dbo.USP_DP_InsertDepositUseLogs 'เพิ่มข้อมูลสนญ','" & dsS02DB.Tables("datasub").Rows(i).Item("DepositNo") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("DocNo") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("Amount") & ""
                    cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
                    cmdS02DB.ExecuteNonQuery()
                End If

            Next
        End If

        Call DropTableBranch(vTableName)

        vTrnState = 1
        vSendTrnStatus = 1

        Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        vSendTrnStatus = 1

        If vCheckError = 0 And vTrnState = 1 Then
            vQuery = "exec dbo.USP_NP_UpdateDepTransStatus '" & vDocNo & "' "
            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
            cmdS02DB.ExecuteNonQuery()
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

    Public Sub TransferDeposit(ByVal vDepositNo As String)
        Dim vGetDocNo As String
        Dim vCountTransfer As Integer
        Dim i As Integer

        'On Error Resume Next

        vGetDocNo = vDepositNo
        vTableName = "BCARDEPOSIT"
        vDocSearchType = 0

        Call CheckDataBranch(3, vTableName, vGetDocNo)
        Call CheckDataHeadOffice(3, vTableName, vGetDocNo)

        If vBranchExist > 0 Then
            If vMemBranchIsCancel = 1 Then
                If vMemBranchIsCancel = vMemHeadOfficeIsCancel Then
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
        Call DeleteTableHeadOffice(vDocSearchType, vTableName, vGetDocNo)

        vCountTransfer = 0
        vQuery = "select * from tempdb.dbo.BCARDeposit where docno ='" & vGetDocNo & "'"
        daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
        dsS02DB = New DataSet
        daS02DB.Fill(dsS02DB, "Data")
        dtS02DB = dsS02DB.Tables("Data")
        If dsS02DB.Tables("Data").Rows.Count > 0 Then
            For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                vQuery = "exec dbo.USP_PTF_InsertARDeposit '" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("TaxDate") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("TaxNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ArCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "',0,'" & dsS02DB.Tables("Data").Rows(i).Item("DueDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("SaleCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxRate") & ",1,'" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsS02DB.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumOfWTax") & "," & dsS02DB.Tables("Data").Rows(i).Item("NetAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("BillBalance") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExcessAmount1") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExcessAmount2") & "," & dsS02DB.Tables("Data").Rows(i).Item("ChargeAmount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("RefNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumCashAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumChqAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumCreditAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumBankAmount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("GLFormat") & "'," & dsS02DB.Tables("Data").Rows(i).Item("GLStartPosting") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsPostGL") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsReturnMoney") & "," & dsS02DB.Tables("Data").Rows(i).Item("GLTransData") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsS02DB.Tables("Data").Rows(i).Item("RecurName") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("JobNo") & "'," & dsS02DB.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("OTHERINCOME") & "," & dsS02DB.Tables("Data").Rows(i).Item("OTHEREXPENSE") & "," & dsS02DB.Tables("Data").Rows(i).Item("CHANGEAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DepositNo") & "'"
                cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                cmdNEBULA.ExecuteNonQuery()
            Next
        End If
        Call DropTableBranch(vTableName)
        Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        vSendTrnStatus = 1
        vQuery = "exec dbo.USP_NP_UpdateDocTransfered '" & vGetDocNo & "'"
        cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        cmdS02DB.ExecuteNonQuery()

    End Sub

    Public Sub CheckCountDepositUse(ByVal vDepositNo As String)
        'On Error Resume Next

        vQuery = "exec dbo.USP_NP_CheckCountInvoiceUseDep '" & vDepositNo & "'"
        daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
        dsS02DB = New DataSet
        daS02DB.Fill(dsS02DB, "Data")
        dtS02DB = dsS02DB.Tables("Data")
        If dsS02DB.Tables("Data").Rows.Count <= 0 Then
            vCountDepositUse = dsS02DB.Tables("Data").Rows(0).Item("vCount")
        Else
            vCountDepositUse = 0
        End If

    End Sub


    Public Sub CheckDeleteDepositInvoice(ByVal vInvNo As String, ByVal vDepositNo As String)

        'On Error Resume Next

        If vConnectionS02DB.State = ConnectionState.Closed Then
            Call ConnectBranch()
        End If


        vQuery = "exec dbo.USP_PTF_CheckDataDepositINV '" & vInvNo & "','" & vDepositNo & "'"
        daS02DB3 = New SqlDataAdapter(vQuery, vConnectionS02DB)
        dsS02DB3 = New DataSet
        daS02DB3.Fill(dsS02DB3, "DataDep")
        dtS02DB3 = dsS02DB3.Tables("DataDep")

        If dsS02DB3.Tables("DataDep").Rows.Count > 0 Then
            vMemOfficeDeleteDepInvoice = dsS02DB3.Tables("DataDep").Rows(0).Item("sumofdeposit1")
            vMemOfficeDepInvoiceHaveUse = dsS02DB3.Tables("DataDep").Rows(0).Item("vCount")
        End If

    End Sub


    Public Sub vConnctDataBase()
        'On Error Resume Next

        vConnectionStringNEBULA = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =NEBULA;Initial Catalog =BCNP"
        vConnectionNEBULA = New SqlConnection(vConnectionStringNEBULA)
        vConnectionNEBULA.Open()

        vConnectionStringS02DB = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =192.168.2.2;Initial Catalog =BCNP"
        vConnectionS02DB = New SqlConnection(vConnectionStringS02DB)
        vConnectionS02DB.Open()
    End Sub

    Private Sub TMTransfer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TMTransfer.Tick
        'On Error Resume Next

        If vMemOfficeIsConnect = 1 Then
            thread3 = New System.Threading.Thread(AddressOf SearchListDocTransfer)
            thread3.Start()
        End If

        'Dim i As Integer
        'Dim vListItem As ListViewItem
        'Dim n As Integer
        'Dim a As Integer

        '''On Error Resume Next

        'Call vConnctDataBase()

        'vQuery = "exec dbo.USP_NP_SearchDocNotTransfer 0"
        'daS02DB1 = New SqlDataAdapter(vQuery, vConnectionS02DB)
        'dsS02DB1 = New DataSet
        'daS02DB1.Fill(dsS02DB1, "Docno")
        'dtS02DB1 = dsS02DB1.Tables("Docno")
        'If dtS02DB1.Rows.Count > 0 Then
        '    For i = 0 To dtS02DB1.Rows.Count - 1
        '        vTableName = dtS02DB1.Rows(i).Item("tablename")
        '        vModuleID = dtS02DB1.Rows(i).Item("moduleid")
        '        vDocNo = dtS02DB1.Rows(i).Item("Docno")
        '        vMemCancelStatus = dtS02DB1.Rows(i).Item("cancelstatus")

        '        Call TransferDoc(vTableName, vDocNo, vMemCancelStatus)

        '        If vCheckError = 0 Then 'And vTrnState = 1 Then
        '            vQuery = "exec dbo.USP_NP_UpdateDocTransfered '" & vDocNo & "'"
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        End If

        '    Next

        '    Me.ListViewListTrn.Items.Clear()

        '    vQuery = "exec dbo.USP_NP_SearchDocNotTransfer 1"
        '    daS02DB2 = New SqlDataAdapter(vQuery, vConnectionS02DB)
        '    dsS02DB2 = New DataSet
        '    daS02DB2.Fill(dsS02DB2, "Docno")
        '    dtS02DB2 = dsS02DB2.Tables("Docno")
        '    If dtS02DB2.Rows.Count > 0 Then
        '        For i = 0 To dtS02DB2.Rows.Count - 1
        '            n = n + 1
        '            vListItem = Me.ListViewListTrn.Items.Add(n)
        '            vListItem.SubItems.Add(0).Text = dtS02DB2.Rows(i).Item("tablename")
        '            vListItem.SubItems.Add(1).Text = dtS02DB2.Rows(i).Item("Docno")
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
    End Sub


    Private Sub Time_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Time.Tick
        'On Error Resume Next

        Me.LBLTime.Text = Now.Hour & ":" & Now.Minute

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

        If vMemOfficeIsConnect = 1 Then
            Me.PBActive.Visible = True
            Me.PBNotConnect.Visible = False
            Me.TBLink.Text = "ติดต่อสำนักงานใหญ่ได้"
        End If

        If vMemOfficeIsConnect = 0 Then
            Me.PBActive.Visible = False
            Me.PBNotConnect.Visible = True
            Me.TBLink.Text = "ไม่สามารถติดต่อสำนักงานใหญ่ได้"
        End If
    End Sub

    Private Sub TMActive_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TMActive.Tick
        'On Error Resume Next

        If vMemBranchIsConnect = 1 Then
            Me.PBActive.Visible = True
            Me.PBNotConnect.Visible = False
            Me.TBLink.Text = "ติดต่อสำนักงานใหญ่ได้"
        Else
            Me.PBActive.Visible = False
            Me.PBNotConnect.Visible = True
            Me.TBLink.Text = "ไม่สามารถติดต่อสำนักงานใหญ่ได้"
        End If
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        'On Error Resume Next

        NotifyIcon1.Visible = False
        Me.Visible = True
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub BTNMinimize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNMinimize.Click
        'On Error Resume Next

        Me.WindowState = 1
        If (Me.WindowState = FormWindowState.Minimized) Then
            Me.Visible = False
            NotifyIcon1.Visible = True
        End If
    End Sub

    Private Sub BTNCloseProgram_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseProgram.Click
        Dim vAnswer As Integer

        'On Error Resume Next

        vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่ การออกโปรแกรมต้องไม่อยู่ระหว่างการโอนมิฉะนั้นออกไม่ได้", MsgBoxStyle.YesNo, "Send Question Message")

        If vAnswer = 6 Then
            If vIsprocess = 0 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub TMNotTransfer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TMNotTransfer.Tick
        'On Error Resume Next

        If vMemOfficeIsConnect = 1 And vMemBranchIsConnect = 1 And vIsTransfer = 0 Then
            thread3 = New System.Threading.Thread(AddressOf SearchDocNotTransfer)
            thread3.Start()
        End If

        'Me.TMNotTransfer.Enabled = False
    End Sub

    Private Sub TConnect_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TConnect.Tick
        'On Error Resume Next

        thread1 = New System.Threading.Thread(AddressOf ConnectOffice)
        thread1.Start()

        thread2 = New System.Threading.Thread(AddressOf ConnectBranch)
        thread2.Start()

        Me.TConnect.Enabled = False
    End Sub

    Public Sub SearchListDocTransfer()
        Dim i As Integer
        Dim vListItem As ListViewItem
        Dim n As Integer
        Dim a As Integer

        'On Error Resume Next

        If vConnectionS02DB.State = ConnectionState.Closed Then
            Call ConnectBranch()
        End If

        If vMemBranchIsConnect = 1 And vIsTransfer = 0 Then
            Me.ListViewListTrn.Items.Clear()

            vQuery = "exec dbo.USP_DP_SearchDepNotTransfer 1"
            daS02DB3 = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB3 = New DataSet
            daS02DB3.Fill(dsS02DB3, "Docno")
            dtS02DB3 = dsS02DB3.Tables("Docno")
            If dtS02DB3.Rows.Count > 0 Then
                For i = 0 To dtS02DB3.Rows.Count - 1
                    n = n + 1
                    vListItem = Me.ListViewListTrn.Items.Add(n)
                    vListItem.SubItems.Add(0).Text = dtS02DB3.Rows(i).Item("depositno")
                    vListItem.SubItems.Add(1).Text = dtS02DB3.Rows(i).Item("trfdate")
                Next
            End If

            If vIsTransfer = 0 Then
                vConnectionS02DB.Close()
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

    Public Sub SearchDocNotTransfer()
        Dim i As Integer

        'On Error Resume Next

        If vConnectionS02DB.State = ConnectionState.Closed Then
            Call ConnectBranch()
        End If

        If vMemBranchIsConnect = 1 And vIsTransfer = 0 Then

            vQuery = "exec dbo.USP_DP_SearchDepNotTransfer 0"
            daS02DB1 = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB1 = New DataSet
            daS02DB1.Fill(dsS02DB1, "Docno")
            dtS02DB1 = dsS02DB1.Tables("Docno")
            If dtS02DB1.Rows.Count > 0 Then
                For i = 0 To dtS02DB1.Rows.Count - 1
                    vDocNo = dtS02DB1.Rows(i).Item("depositno")

                    Call TransferDoc(vDocNo)

                Next
            End If

            If vIsTransfer = 0 Then
                vConnectionS02DB.Close()
            End If
        End If
    End Sub

    Public Sub CompareOfficeDepUse(ByVal vInvNo As String, ByVal vDepNo As String, ByVal vInvAmount As Double)

        'On Error Resume Next

        If vConnectionNEBULA.State = ConnectionState.Closed Then
            Call ConnectOffice()
        End If

        vQuery = "exec dbo.USP_NP_CheckDataDepUse '" & vInvNo & "','" & vDepNo & "'," & vInvAmount & ""
        daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA = New DataSet
        daNEBULA.Fill(dsNEBULA, "DataDep")
        dtNEBULA = dsNEBULA.Tables("DataDep")

        If dsNEBULA.Tables("DataDep").Rows.Count > 0 Then
            vMemBranchDepExist = dsNEBULA.Tables("DataDep").Rows(0).Item("countdoc")
        End If
    End Sub

    Public Sub DeleteDepUseOffice(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String, ByVal vInvAmount As Double)
        'On Error GoTo ErrDescription

        'On Error Resume Next

        'vQuery = "begin tran"
        'cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        'cmdS02DB.ExecuteNonQuery()

        If vConnectionNEBULA.State = ConnectionState.Closed Then
            Call ConnectOffice()
        End If

        If vConnectionS02DB.State = ConnectionState.Closed Then
            Call ConnectBranch()
        End If

        vType = 6

        vQuery = "exec dbo.USP_PTF_DeleteDataTable " & vType & ",'" & vTableName & "','" & vValue & "' "
        cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
        cmdNEBULA.ExecuteNonQuery()


        vQuery = "exec dbo.USP_DP_InsertDepositUseLogs 'ลบข้อมูลสนญ','" & vTableName & "','" & vValue & "'," & vInvAmount & ""
        cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        cmdS02DB.ExecuteNonQuery()


        'vQuery = "commit tran"
        'cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        'cmdS02DB.ExecuteNonQuery()

        'ErrDescription:

    End Sub

    Private Sub TCheckConnect_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TCheckConnect.Tick
        If vMemOfficeIsConnect = 1 Then
            Me.PBActive.Visible = True
            Me.PBNotConnect.Visible = False
            Me.TBLink.Text = "ติดต่อสำนักงานใหญ่ได้"
        End If

        If vMemOfficeIsConnect = 0 Then
            Me.PBActive.Visible = False
            Me.PBNotConnect.Visible = True
            Me.TBLink.Text = "ไม่สามารถติดต่อสำนักงานใหญ่ได้"
        End If

        thread1 = New System.Threading.Thread(AddressOf ConnectOffice)
        thread1.Start()
    End Sub
End Class