Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Net.NetworkInformation

Public Class FormTrnDataNopadolToBranch
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

    Dim vConnectionStringNEBULA As String
    Dim vConnectionNEBULA As SqlConnection
    Dim daNEBULA As SqlDataAdapter
    Dim dsNEBULA As DataSet
    Dim dtNEBULA As DataTable
    Dim dvNEBULA As DataView
    Dim vQuery As String
    Dim cmdNEBULA As SqlCommand

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

    Private Sub FormTransferApp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'vConnectionStringNEBULA = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Max Pool Size = 10000;Min Pool Size = 5;Data Source =NEBULA;Initial Catalog =BCNP"
        vConnectionStringNEBULA = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =NEBULA;Initial Catalog =BCNP"
        vConnectionNEBULA = New SqlConnection(vConnectionStringNEBULA)
        vConnectionNEBULA.Open()

        'vConnectionStringS02DB = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Max Pool Size = 10000;Min Pool Size = 5;Data Source =192.168.2.2;Initial Catalog =BCNP"
        vConnectionStringS02DB = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =192.168.2.2;Initial Catalog =BCNP"
        vConnectionS02DB = New SqlConnection(vConnectionStringS02DB)
        vConnectionS02DB.Open()

    End Sub

    Public Sub CheckDataHeadOffice(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
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
        vQuery = "exec dbo.USP_PTF_PrepareDataTable " & vType & ",'" & vTableName & "','" & vValues & "'"
        'cmdNEBULA = New SqlCommand(vQuery, vConnectionS02DB)
        cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
        cmdNEBULA.ExecuteNonQuery()

    End Sub

    Public Sub DeleteTableBranch(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        'On Error GoTo ErrDescription

        'vQuery = "begin tran"
        'cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        'cmdS02DB.ExecuteNonQuery()

        vQuery = "exec dbo.USP_PTF_DeleteDataTable " & vType & ",'" & vTableName & "','" & vValue & "' "
        cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        cmdS02DB.ExecuteNonQuery()

        'vQuery = "commit tran"
        'cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        'cmdS02DB.ExecuteNonQuery()

        'ErrDescription:

    End Sub

    Public Sub DropTable(ByVal vTableName As String)
        vQuery = "exec dbo.USP_PTF_DropDataTable '" & vTableName & "'"
        'cmdNEBULA = New SqlCommand(vQuery, vConnectionS02DB)
        cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
        cmdNEBULA.ExecuteNonQuery()
    End Sub

    Public Sub InsertLogs(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        Dim vDatasource As String
        Dim vDestination As String

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

    Public Sub TransferDoc(ByVal vGetTableName As String, ByVal vGetDocNo As String, ByVal vIsCancel As Integer)
        Dim vCountTransfer As Integer
        Dim i As Integer

        On Error Resume Next

        '===============================================================================================================
        '===============================================================================================================

        'If vGetTableName = "BCSTKTRANSFER" Then
        '    vTableName = "BCSTKTRANSFER"
        '    vDocSearchType = 0

        '    Call CheckDataHeadOffice(3, vTableName, vGetDocNo)
        '    Call CheckDataBranch(3, vTableName, vGetDocNo)

        '    If vBranchExist > 0 Then

        '        If vMemHeadOfficeIsCancel = 1 Then
        '            If vMemHeadOfficeIsCancel = vMemBranchIsCancel Then
        '                vSendTrnStatus = 2
        '                GoTo LineEnd
        '            End If
        '        End If

        '        If vMemHeadOfficeLastEditorCode <> "" Then
        '            If vMemHeadOfficeLastEditDateT <= vMemBranchLastEditDateT Then
        '                vSendTrnStatus = 2
        '                GoTo LineEnd
        '            End If
        '        End If
        '    End If

        '    Call CheckData(vDocSearchType, vTableName, vGetDocNo)
        '    If vCheckExist = 0 Then
        '        vSendTrnStatus = 2
        '        Exit Sub
        '    End If
        '    Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        '    Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        '    vCountTransfer = 0
        '    vQuery = "select * from tempdb.dbo.BCStktransfer where docno ='" & vGetDocNo & "'"
        '    daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA = New DataSet
        '    daNEBULA.Fill(dsNEBULA, "Data")
        '    dtNEBULA = dsNEBULA.Tables("Data")
        '    If dsNEBULA.Tables("Data").Rows.Count > 0 Then
        '        For i = 0 To dsNEBULA.Tables("Data").Rows.Count - 1
        '            'vQuery = "exec dbo.USP_PTF_InsertSTKTransfer '" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "',1,'" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsS02DB.Tables("Data").Rows(i).Item("SumOfQty") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCompleteSave") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("RecurName") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsS02DB.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISPOS") & "," & dsS02DB.Tables("Data").Rows(i).Item("SUMOFAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DepositNo") & "'"
        '            vQuery = "exec dbo.USP_PTF_InsertSTKTransfer '" & dsNEBULA.Tables("Data").Rows(i).Item("DocNo") & "',1,'" & dsNEBULA.Tables("Data").Rows(i).Item("DocDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("SumOfQty") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCancel") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCompleteSave") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("RecurName") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISPOS") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SUMOFAMOUNT") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DepositNo") & "'"
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        Next
        '    End If
        '    Call DropTable(vTableName)
        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1

        '    '=======================================================================================================================================================================================================

        '    vTableName = "BCSTKTRANSFSUB"
        '    Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        '    Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        '    vCountTransfer = 0
        '    vQuery = "select * from tempdb.dbo.BCStktransfSub where docno ='" & vGetDocNo & "'"
        '    daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA = New DataSet
        '    daNEBULA.Fill(dsNEBULA, "Data")
        '    dtNEBULA = dsNEBULA.Tables("Data")
        '    If dsNEBULA.Tables("Data").Rows.Count > 0 Then
        '        For i = 0 To dsNEBULA.Tables("Data").Rows.Count - 1
        '            vQuery = "exec dbo.USP_PTF_InsertSTKTransferSub '" & dsNEBULA.Tables("Data").Rows(i).Item("DocNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DocDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("FromWH") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("FromShelf") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ToWH") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ToShelf") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("Qty") & "," & dsNEBULA.Tables("Data").Rows(i).Item("Price") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumOfCost") & "," & dsNEBULA.Tables("Data").Rows(i).Item("Amount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("RefNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("OppositeQty") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("CategoryCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("GroupCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("BrandCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("TypeCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("FormatCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("BarCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TransState") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCancel") & "," & dsNEBULA.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("LotOutNumber") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LotInNumber") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LotExpireDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LotMyDesc") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("StyleCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("BEHINDINDEX") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISPOS") & "," & dsNEBULA.Tables("Data").Rows(i).Item("REFLINENUMBER") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        Next
        '    End If
        '    Call DropTable(vTableName)
        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1

        '    '=======================================================================================================================================================================================================

        '    vTableName = "BCSTKTRANSFSUB3"
        '    Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        '    Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        '    vCountTransfer = 0
        '    vQuery = "select * from tempdb.dbo.BCStktransfSub3 where docno ='" & vGetDocNo & "'"
        '    daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA = New DataSet
        '    daNEBULA.Fill(dsNEBULA, "Data")
        '    dtNEBULA = dsNEBULA.Tables("Data")
        '    If dsNEBULA.Tables("Data").Rows.Count > 0 Then
        '        For i = 0 To dsNEBULA.Tables("Data").Rows.Count - 1
        '            vQuery = "exec dbo.USP_PTF_InsertSTKTransferSub3 " & dsNEBULA.Tables("Data").Rows(i).Item("BehindIndex") & "," & dsNEBULA.Tables("Data").Rows(i).Item("MyType") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DocNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DocDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("WHCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ShelfCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("Qty") & "," & dsNEBULA.Tables("Data").Rows(i).Item("Price") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumOfCost") & "," & dsNEBULA.Tables("Data").Rows(i).Item("Amount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("RefNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("OppositeQty") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("fromwhcode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("fromshelfcode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("reftrlinenum") & "," & dsNEBULA.Tables("Data").Rows(i).Item("FixUnitCost") & "," & dsNEBULA.Tables("Data").Rows(i).Item("FixUnitQty") & "," & dsNEBULA.Tables("Data").Rows(i).Item("AVERAGECOST") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("LotNumber") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LotExpireDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LotMyDesc") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("StyleCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TAXRATE") & "," & dsNEBULA.Tables("Data").Rows(i).Item("PACKINGRATE1") & "," & dsNEBULA.Tables("Data").Rows(i).Item("PACKINGRATE2") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISCANCEL") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISPOS") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISPROCESS") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISLOCKCOST") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        Next
        '    End If
        '    Call DropTable(vTableName)
        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1
        'End If

        ''===============================================================================================================
        ''===============================================================================================================

        'If vGetTableName = "BCSALEORDER" Then
        '    vTableName = "BCSALEORDER"
        '    vDocSearchType = 0

        '    Call CheckDataHeadOffice(vDocSearchType, vTableName, vGetDocNo)
        '    Call CheckDataBranch(vDocSearchType, vTableName, vGetDocNo)

        '    If vBranchExist > 0 Then
        '        If vMemBranchBillStatus <> 0 Then
        '            vSendTrnStatus = 2
        '            GoTo LineEnd
        '        End If

        '        If vMemHeadOfficeIsCancel = 1 Then
        '            If vMemHeadOfficeIsCancel = vMemBranchIsCancel Then
        '                vSendTrnStatus = 2
        '                GoTo LineEnd
        '            End If
        '        End If

        '        If vMemHeadOfficeLastEditorCode <> "" Then
        '            If vMemHeadOfficeLastEditDateT <= vMemBranchLastEditDateT Then
        '                vSendTrnStatus = 2
        '                GoTo LineEnd
        '            End If
        '        End If
        '    End If

        '    Call CheckData(vDocSearchType, vTableName, vGetDocNo)
        '    If vCheckExist = 0 Then
        '        vSendTrnStatus = 2
        '        Exit Sub
        '    End If

        '    Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        '    Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        '    vCountTransfer = 0
        '    vQuery = "select * from tempdb.dbo.BCSaleOrder where docno ='" & vGetDocNo & "'"
        '    daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA = New DataSet
        '    daNEBULA.Fill(dsNEBULA, "Data")
        '    dtNEBULA = dsNEBULA.Tables("Data")
        '    If dsNEBULA.Tables("Data").Rows.Count > 0 Then
        '        For i = 0 To dsNEBULA.Tables("Data").Rows.Count - 1
        '            vQuery = "exec dbo.USP_PTF_InsertSaleOrder '" & dsNEBULA.Tables("Data").Rows(i).Item("DocNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DocDate") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxType") & "," & dsNEBULA.Tables("Data").Rows(i).Item("BillType") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("ArCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("CreditDay") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DueDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DeliveryAddr") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxRate") & ",1,'" & dsNEBULA.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("RefDocNo") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("BillStatus") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SOStatus") & "," & dsNEBULA.Tables("Data").Rows(i).Item("HoldingStatus") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("CrAuthorizeMan") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ContactCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("NetAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCancel") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCompleteSave") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("RecurName") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("IsConditionSend") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("OwnReceive") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CarLicense") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ApproveCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ApproveDateTime") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("IsUseRobotSale") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TotalTransport") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumOtherValue") & "," & dsNEBULA.Tables("Data").Rows(i).Item("READYFORPAY") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("TimeTransport") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CarType") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CondPayCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("PrintCount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("PORefNo") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("IsImport") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("USERGROUP") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("METHODEPAYBILL") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("METHODEPAYBILL2") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("DELIVERYDAY") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DELIVERYDATE") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("DELIVERYSTATUS") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("PROMOTIONCODE") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("ISEXPORT") & ""
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        Next
        '    End If
        '    Call DropTable(vTableName)
        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1

        '    '=======================================================================================================================================================================================================

        '    vTableName = "BCSALEORDERSUB"
        '    Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        '    Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        '    vCountTransfer = 0
        '    vQuery = "select * from tempdb.dbo.BCSaleOrderSub where docno ='" & vGetDocNo & "'"
        '    daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA = New DataSet
        '    daNEBULA.Fill(dsNEBULA, "DataSub")
        '    dtNEBULA = dsNEBULA.Tables("DataSub")
        '    If dsNEBULA.Tables("DataSub").Rows.Count > 0 Then
        '        For i = 0 To dsNEBULA.Tables("DataSub").Rows.Count - 1
        '            vQuery = "exec dbo.USP_PTF_InsertSaleOrderSub '" & dsNEBULA.Tables("datasub").Rows(i).Item("DocNo") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("TaxType") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("ItemCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("ArCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("DepartCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("SaleCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("MyDescription") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("ItemName") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("WHCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("ShelfCode") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("Qty") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("RemainQty") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("Price") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("DiscountWord") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("DiscountAmount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("Amount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("HomeAmount") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("UnitCode") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("OppositePrice2") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("OppositeUnit") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("OppositeQty") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("StkReserveNo") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("SOStatus") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("HoldingStatus") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("RefType") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ItemType") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ExceptTax") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("AllocateCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("ProjectCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("CurrencyCode") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("ExchangeRate") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("TransState") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("IsCancel") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("CategoryCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("GroupCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("BrandCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("TypeCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("FormatCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("BarCode") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("IsPromotion") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("PremiumStatus") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("IsUseRobotSale") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("IsConditionSend") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("TransportAmount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("OtherValue") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("REMAINPAYQTY") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ISPAYMENT") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("Colorcode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("SizeCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("StyleCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("itemsetcode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("PORefNo") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("BEHINDINDEX") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("TAXRATE") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("PACKINGRATE1") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("PACKINGRATE2") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("USERGROUP") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("ISPROCESS") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ISLOCKCOST") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("DELIVERYQTY") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("PROMOTIONCODE") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("MASTERITEMCODE") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        Next
        '    End If
        '    Call DropTable(vTableName)
        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1

        'End If

        ''===============================================================================================================
        ''===============================================================================================================

        'If vGetTableName = "BCARDEPOSIT" Then
        '    vTableName = "BCARDEPOSIT"
        '    vDocSearchType = 0

        '    Call CheckDataHeadOffice(3, vTableName, vGetDocNo)
        '    Call CheckDataBranch(3, vTableName, vGetDocNo)

        '    If vBranchExist > 0 Then
        '        If vMemBranchBillStatus <> 0 Then
        '            vSendTrnStatus = 2
        '            GoTo LineEnd
        '        End If

        '        If vMemHeadOfficeIsCancel = 1 Then
        '            If vMemHeadOfficeIsCancel = vMemBranchIsCancel Then
        '                vSendTrnStatus = 2
        '                GoTo LineEnd
        '            End If
        '        End If

        '    End If

        '    Call CheckData(vDocSearchType, vTableName, vGetDocNo)
        '    If vCheckExist = 0 Then
        '        vSendTrnStatus = 2
        '        Exit Sub
        '    End If

        '    Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        '    Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        '    vCountTransfer = 0
        '    vQuery = "select * from tempdb.dbo.BCARDEPOSIT where docno ='" & vGetDocNo & "'"
        '    daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA = New DataSet
        '    daNEBULA.Fill(dsNEBULA, "Data")
        '    dtNEBULA = dsNEBULA.Tables("Data")
        '    If dsNEBULA.Tables("Data").Rows.Count > 0 Then
        '        For i = 0 To dsNEBULA.Tables("Data").Rows.Count - 1
        '            vQuery = "exec dbo.USP_PTF_InsertARDeposit '" & dsNEBULA.Tables("Data").Rows(i).Item("DocNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DocDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("TaxDate") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("TaxNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ArCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DepartCode") & "',0,'" & dsNEBULA.Tables("Data").Rows(i).Item("DueDate") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("SaleCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxRate") & ",1,'" & dsNEBULA.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumOfWTax") & "," & dsNEBULA.Tables("Data").Rows(i).Item("NetAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("BillBalance") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ExcessAmount1") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ExcessAmount2") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ChargeAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("RefNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumCashAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumChqAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumCreditAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("SumBankAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("GLFormat") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("GLStartPosting") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsPostGL") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCancel") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsReturnMoney") & "," & dsNEBULA.Tables("Data").Rows(i).Item("GLTransData") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("RecurName") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("JobNo") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsNEBULA.Tables("Data").Rows(i).Item("OTHERINCOME") & "," & dsNEBULA.Tables("Data").Rows(i).Item("OTHEREXPENSE") & "," & dsNEBULA.Tables("Data").Rows(i).Item("CHANGEAMOUNT") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DepositNo") & "'"
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        Next
        '    End If
        '    Call DropTable(vTableName)
        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1

        '    '=======================================================================================================================================================================================================

        '    vTableName = "BCARDEPOSITUSE"

        '    Call PrepareData(3, vTableName, vGetDocNo)
        '    Call DeleteTableBranch(3, vTableName, vGetDocNo)

        '    vCountTransfer = 0
        '    vQuery = "select * from tempdb.dbo.BCARDepositUse where depositno ='" & vGetDocNo & "'"
        '    daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA = New DataSet
        '    daNEBULA.Fill(dsNEBULA, "DataSub")
        '    dtNEBULA = dsNEBULA.Tables("DataSub")
        '    If dsNEBULA.Tables("DataSub").Rows.Count > 0 Then
        '        For i = 0 To dsNEBULA.Tables("DataSub").Rows.Count - 1
        '            vQuery = "exec dbo.USP_PTF_InsertARDepositUse '" & dsNEBULA.Tables("datasub").Rows(i).Item("DocNo") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("DepositNo") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("MyDescription") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("Balance") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("Amount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("DeposTaxType") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("CurrencyCode") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("DPExchangeRate") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("NewExchangeRate") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("HomeAmount1") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("HomeAmount2") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ExchangeProfit") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        Next
        '    End If
        '    Call DropTable(vTableName)
        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1

        'End If

        ''===============================================================================================================
        ''===============================================================================================================
        'If vGetTableName = "BCPURCHASEORDER" Then
        '    vTableName = "BCPURCHASEORDER"
        '    vDocSearchType = 0

        '    Call CheckDataHeadOffice(vDocSearchType, vTableName, vGetDocNo)
        '    Call CheckDataBranch(vDocSearchType, vTableName, vGetDocNo)

        '    If vBranchExist > 0 Then
        '        If vMemBranchBillStatus <> 0 Then
        '            vSendTrnStatus = 2
        '            GoTo LineEnd
        '        End If

        '        If vMemHeadOfficeIsCancel = 1 Then
        '            If vMemHeadOfficeIsCancel = vMemBranchIsCancel Then
        '                vSendTrnStatus = 2
        '                GoTo LineEnd
        '            End If
        '        End If

        '        If vMemHeadOfficeLastEditorCode <> "" Then
        '            If vMemHeadOfficeLastEditDateT <= vMemBranchLastEditDateT Then
        '                vSendTrnStatus = 2
        '                GoTo LineEnd
        '            End If
        '        End If
        '    End If

        '    Call CheckData(vDocSearchType, vTableName, vGetDocNo)
        '    If vCheckExist = 0 Then
        '        vSendTrnStatus = 2
        '        Exit Sub
        '    End If
        '    Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        '    Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        '    vCountTransfer = 0
        '    vQuery = "select * from tempdb.dbo.bcpurchaseorder where docno ='" & vGetDocNo & "'"
        '    daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA = New DataSet
        '    daNEBULA.Fill(dsNEBULA, "Data")
        '    dtNEBULA = dsNEBULA.Tables("Data")
        '    If dsNEBULA.Tables("Data").Rows.Count > 0 Then
        '        For i = 0 To dsNEBULA.Tables("Data").Rows.Count - 1
        '            vQuery = "exec dbo.USP_PTF_InsertPurchaseOrder '" & dsNEBULA.Tables("Data").Rows(i).Item("DocNo") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DocDate") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("ApCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("CreditDay") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DueDate") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("LeadTime") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("LeadDate") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("ExpireCredit") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("ExpireDate") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("TaxRate") & ",1,'" & dsNEBULA.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("PoStatus") & "," & dsNEBULA.Tables("Data").Rows(i).Item("BillStatus") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCancel") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsCompleteSave") & "," & dsNEBULA.Tables("Data").Rows(i).Item("IsProcessOK") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ContactCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsNEBULA.Tables("Data").Rows(i).Item("NetAmount") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("RecurName") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("IsConditionSend") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("OrderToArCode") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("IsImport") & "," & dsNEBULA.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("Buyer") & "'"
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        Next
        '    End If
        '    Call DropTable(vTableName)
        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1

        '    '=======================================================================================================================================================================================================

        '    vTableName = "BCPURCHASEORDERSUB"
        '    Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        '    Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        '    vCountTransfer = 0
        '    vQuery = "select * from tempdb.dbo.bcpurchaseordersub where docno ='" & vGetDocNo & "'"
        '    daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA = New DataSet
        '    daNEBULA.Fill(dsNEBULA, "DataSub")
        '    dtNEBULA = dsNEBULA.Tables("DataSub")
        '    If dsNEBULA.Tables("DataSub").Rows.Count > 0 Then
        '        For i = 0 To dsNEBULA.Tables("DataSub").Rows.Count - 1
        '            vQuery = "exec dbo.USP_PTF_InsertPurchaseOrderSub '" & dsNEBULA.Tables("datasub").Rows(i).Item("DocNo") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("TaxType") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("ItemCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("ApCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("DepartCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("MyDescription") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("ItemName") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("WHCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("ShelfCode") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("Qty") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("RemainQty") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("Price") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("DiscountWord") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("DiscountAmount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("Amount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("HomeAmount") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("UnitCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("OppositeUnit") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("OppositeQty") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("StkReqNo") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("ConfirmNo") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("ItemType") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ExceptTax") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("AllocateCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("ProjectCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("CurrencyCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("ExchangeRate") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("TransState") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("IsCancel") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("CategoryCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("GroupCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("BrandCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("TypeCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("FormatCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("BarCode") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("IsPromotion") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("RefLineNumber") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("Colorcode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("SizeCode") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("StyleCode") & "'," & dsNEBULA.Tables("datasub").Rows(i).Item("BEHINDINDEX") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("TAXRATE") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("PACKINGRATE1") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("PACKINGRATE2") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ISPROCESS") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ISLOCKCOST") & "," & dsNEBULA.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsNEBULA.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        Next
        '    End If
        '    Call DropTable(vTableName)
        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1

        'End If

        ''===============================================================================================================
        ''===============================================================================================================

        'If vGetTableName = "BCCOUPON" Then
        '    vTableName = "BCCOUPON"
        '    vDocSearchType = 1

        '    Call CheckDataHeadOffice(vDocSearchType, vTableName, vGetDocNo)
        '    Call CheckDataBranch(vDocSearchType, vTableName, vGetDocNo)

        '    If vBranchExist > 0 Then
        '        If vMemBranchBillStatus <> 0 Then
        '            vSendTrnStatus = 2
        '            GoTo LineEnd
        '        End If

        '        If vMemHeadOfficeIsCancel = 1 Then
        '            If vMemHeadOfficeIsCancel = vMemBranchIsCancel Then
        '                vSendTrnStatus = 2
        '                GoTo LineEnd
        '            End If
        '        End If

        '        If vMemHeadOfficeLastEditorCode <> "" Then
        '            If vMemHeadOfficeLastEditDateT <= vMemBranchLastEditDateT Then
        '                vSendTrnStatus = 2
        '                GoTo LineEnd
        '            End If
        '        End If
        '    End If

        '    Call CheckData(vDocSearchType, vTableName, vGetDocNo)
        '    If vCheckExist = 0 Then
        '        vSendTrnStatus = 2
        '        Exit Sub
        '    End If
        '    Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        '    Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        '    vCountTransfer = 0
        '    vQuery = "select * from tempdb.dbo.bccoupon where code ='" & vGetDocNo & "'"
        '    daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA = New DataSet
        '    daNEBULA.Fill(dsNEBULA, "Data")
        '    dtNEBULA = dsNEBULA.Tables("Data")
        '    If dsNEBULA.Tables("Data").Rows.Count > 0 Then
        '        For i = 0 To dsNEBULA.Tables("Data").Rows.Count - 1
        '            vQuery = "exec dbo.USP_PTF_InsertCoupon '" & dsNEBULA.Tables("Data").Rows(i).Item("CODE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("NAME") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("COUPONTYPE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("STARTBOOK") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("STOPBOOK") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("STARTNO") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("STOPNO") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("BEGINDATE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("EXPIREDATE") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("COUPONVAL") & "," & dsNEBULA.Tables("Data").Rows(i).Item("COUPONSTATUS") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        Next
        '    End If
        '    Call DropTable(vTableName)
        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1

        'End If

        ''===============================================================================================================
        ''===============================================================================================================
        'If vGetTableName = "BCCOUPONRECEIVE" Then
        '    vTableName = "BCCOUPONRECEIVE"
        '    vDocSearchType = 4

        '    Call CheckDataHeadOffice(vDocSearchType, vTableName, vGetDocNo)
        '    Call CheckDataBranch(vDocSearchType, vTableName, vGetDocNo)

        '    If vBranchExist > 0 Then
        '        If vMemBranchBillStatus <> 0 Then
        '            vSendTrnStatus = 2
        '            GoTo LineEnd
        '        End If

        '        If vMemHeadOfficeIsCancel = 1 Then
        '            If vMemHeadOfficeIsCancel = vMemBranchIsCancel Then
        '                vSendTrnStatus = 2
        '                GoTo LineEnd
        '            End If
        '        End If

        '        If vMemHeadOfficeLastEditorCode <> "" Then
        '            If vMemHeadOfficeLastEditDateT <= vMemBranchLastEditDateT Then
        '                vSendTrnStatus = 2
        '                GoTo LineEnd
        '            End If
        '        End If
        '    End If

        '    Call CheckData(vDocSearchType, vTableName, vGetDocNo)
        '    If vCheckExist = 0 Then
        '        vSendTrnStatus = 2
        '        Exit Sub
        '    End If
        '    Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
        '    Call DeleteTableBranch(vDocSearchType, vTableName, vGetDocNo)

        '    vCountTransfer = 0
        '    vQuery = "select * from tempdb.dbo.BCCouponReceive where CouponNo ='" & vGetDocNo & "'"
        '    daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        '    dsNEBULA = New DataSet
        '    daNEBULA.Fill(dsNEBULA, "Data")
        '    dtNEBULA = dsNEBULA.Tables("Data")
        '    If dsNEBULA.Tables("Data").Rows.Count > 0 Then
        '        For i = 0 To dsNEBULA.Tables("Data").Rows.Count - 1
        '            vQuery = "exec dbo.USP_PTF_InsertCouponReceive '" & dsNEBULA.Tables("Data").Rows(i).Item("COUPONCODE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("COUPONTYPE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("COUPONNO") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("TOCOUPONNO") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("COUPONCOUNT") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("DOCNO") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("BOOK") & "'," & dsNEBULA.Tables("Data").Rows(i).Item("COUPONVAL") & "," & dsNEBULA.Tables("Data").Rows(i).Item("LINENUMBER") & "," & dsNEBULA.Tables("Data").Rows(i).Item("PERCENTDISC") & "," & dsNEBULA.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsNEBULA.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsNEBULA.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
        '            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '            cmdS02DB.ExecuteNonQuery()
        '        Next
        '    End If
        '    Call DropTable(vTableName)
        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1

        'End If

        ''===============================================================================================================
        ''===============================================================================================================

        'If vGetTableName = "TB_NP_TransferDocLogs" And vIsCancel = 1 Then
        '    vTableName = "TB_NP_TransferDocLogs"
        '    vDocSearchType = 5

        '    vQuery = "delete dbo.bcardeposituse where docno = '" & vGetDocNo & "'"
        '    cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        '    cmdS02DB.ExecuteNonQuery()

        '    Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
        '    vSendTrnStatus = 1

        'End If

        '===============================================================================================================
        '===============================================================================================================

        If vGetTableName = "TB_NP_CheckCancelDep" Then
            vTableName = "TB_NP_CheckCancelDep"
            vDocSearchType = 0

            Dim n As Integer
            Dim vDocNoOffice(0) As String
            Dim vDepositNoOffice(0) As String
            Dim vCountOffice As Integer
            Dim vDocNoBranch(0) As String
            Dim vDepositNoBranch(0) As String
            Dim vCountBranch As Integer


            vQuery = "select docno,depositno dbo.bcardeposituse where docno = '" & vGetDocNo & "'"
            daNEBULA = New SqlDataAdapter(vQuery, vConnectionNEBULA)
            dsNEBULA = New DataSet
            daNEBULA.Fill(dsNEBULA, "DepBranch")
            dtNEBULA = dsNEBULA.Tables("DepBranch")
            If dsNEBULA.Tables("DepBranch").Rows.Count > 0 Then
                vCountOffice = dsNEBULA.Tables("DepBranch").Rows.Count

                ReDim vDocNoBranch(vCountOffice)
                ReDim vDepositNoBranch(vCountOffice)

                For n = 0 To vCountOffice - 1
                    vDocNoOffice(n) = dsNEBULA.Tables("DepBranch").Rows(n).Item("docno")
                    vDepositNoOffice(n) = dsNEBULA.Tables("DepBranch").Rows(n).Item("depositno")
                Next
            End If


            vQuery = "select docno,depositno dbo.bcardeposituse where docno = '" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "DepBranch")
            dtS02DB = dsS02DB.Tables("DepBranch")
            If dsS02DB.Tables("DepBranch").Rows.Count > 0 Then
                vCountBranch = dsS02DB.Tables("DepBranch").Rows.Count

                ReDim vDocNoBranch(vCountBranch)
                ReDim vDepositNoBranch(vCountBranch)

                For n = 0 To vCountBranch - 1
                    vDocNoBranch(n) = dsS02DB.Tables("DepBranch").Rows(n).Item("docno")
                    vDepositNoBranch(n) = dsS02DB.Tables("DepBranch").Rows(n).Item("depositno")
                Next
            End If

            Dim a As Integer
            Dim b As Integer
            Dim vCheckInvoiceNoOffice As String
            Dim vCheckDepositNoOffice As String
            Dim vCheckInvoiceNoBranch As String
            Dim vCheckDepositNoBranch As String
            Dim vExist As Integer
            Dim vCountExist As Integer


            If vCountOffice = 0 And vCountBranch > 0 Then
                MsgBox("", MsgBoxStyle.Critical, "error")
                vQuery = "delete dbo.bcardeposituse where docno = '" & vGetDocNo & "'"
                cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
                cmdS02DB.ExecuteNonQuery()

                Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
                vSendTrnStatus = 1
            End If

            If vCountBranch > vCountOffice Then
                MsgBox("", MsgBoxStyle.Critical, "error")
                For a = 0 To vCountBranch - 1
                    vCheckInvoiceNoBranch = vDocNoBranch(a)
                    vCheckDepositNoBranch = vDepositNoBranch(a)

                    For b = 0 To vCountOffice - 1
                        vCheckInvoiceNoOffice = vDocNoOffice(b)
                        vCheckDepositNoOffice = vDepositNoOffice(b)

                        If vCheckInvoiceNoBranch = vCheckInvoiceNoOffice And vCheckDepositNoBranch = vCheckDepositNoOffice Then
                            vExist = 1
                        Else
                            vExist = 0
                        End If
                        vCountExist = vCountExist + vExist

                        If vCountExist = 0 Then
                            vQuery = "delete dbo.bcardeposituse where docno = '" & vGetDocNo & "'"
                            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
                            cmdS02DB.ExecuteNonQuery()

                            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
                            vSendTrnStatus = 1
                        End If
                    Next
                Next
            End If

            If vCountOffice > vCountBranch Then
                MsgBox("", MsgBoxStyle.Critical, "error")
                For a = 0 To vCountOffice - 1
                    vCheckInvoiceNoOffice = vDocNoOffice(a)
                    vCheckDepositNoOffice = vDepositNoOffice(a)

                    For b = 0 To vCountBranch - 1
                        vCheckInvoiceNoBranch = vDocNoBranch(b)
                        vCheckDepositNoBranch = vDepositNoBranch(b)

                        If vCheckInvoiceNoBranch = vCheckInvoiceNoOffice And vCheckDepositNoBranch = vCheckDepositNoOffice Then
                            vExist = 1
                        Else
                            vExist = 0
                        End If
                        vCountExist = vCountExist + vExist

                        If vCountExist = 0 Then
                            vQuery = "delete dbo.bcardeposituse where docno = '" & vGetDocNo & "'"
                            cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
                            cmdS02DB.ExecuteNonQuery()

                            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
                            vSendTrnStatus = 1
                        End If
                    Next
                Next
            End If

        End If


        '===============================================================================================================
        '===============================================================================================================


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

    Public Sub CheckCountDepositUse(ByVal vDepositNo As String)
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
        Dim i As Integer
        Dim vListItem As ListViewItem
        Dim n As Integer
        Dim a As Integer

        On Error Resume Next

        Call vConnctDataBase()

        vQuery = "exec dbo.USP_NP_SearchDocNotTransfer 0"
        daNEBULA1 = New SqlDataAdapter(vQuery, vConnectionNEBULA)
        dsNEBULA1 = New DataSet
        daNEBULA1.Fill(dsNEBULA1, "Docno")
        dtNEBULA1 = dsNEBULA1.Tables("Docno")
        If dtNEBULA1.Rows.Count > 0 Then
            For i = 0 To dtNEBULA1.Rows.Count - 1
                vTableName = dtNEBULA1.Rows(i).Item("tablename")
                vModuleID = dtNEBULA1.Rows(i).Item("moduleid")
                vDocNo = dtNEBULA1.Rows(i).Item("Docno")
                vMemCancelStatus = dtNEBULA1.Rows(i).Item("cancelstatus")

                Call TransferDoc(vTableName, vDocNo, vMemCancelStatus)

                If vCheckError = 0 Then
                    vQuery = "exec dbo.USP_NP_UpdateDocTransfered '" & vDocNo & "'"
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                End If

            Next

            Me.ListViewListTrn.Items.Clear()

            vQuery = "exec dbo.USP_NP_SearchDocNotTransfer 1"
            daNEBULA2 = New SqlDataAdapter(vQuery, vConnectionNEBULA)
            dsNEBULA2 = New DataSet
            daNEBULA2.Fill(dsNEBULA2, "Docno")
            dtNEBULA2 = dsNEBULA2.Tables("Docno")
            If dtNEBULA2.Rows.Count > 0 Then
                For i = 0 To dtNEBULA2.Rows.Count - 1
                    n = n + 1
                    vListItem = Me.ListViewListTrn.Items.Add(n)
                    vListItem.SubItems.Add(0).Text = dtNEBULA2.Rows(i).Item("tablename")
                    vListItem.SubItems.Add(1).Text = dtNEBULA2.Rows(i).Item("Docno")
                Next
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

        'ErrDescription:
        '        If Err.Description <> "" Then
        '            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        '            Exit Sub
        '        End If
    End Sub

    Public Function PingAddress(ByVal ServerIP As String) As Boolean
        Dim ReplySuccess As Boolean
        Dim IpAddress As New System.Net.IPAddress(New Byte() {192, 168, 2, 2})

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

        If Now.Hour > 7 And Now.Hour < 19 Then
            Me.TMTransfer.Enabled = True
            Me.TBLink.Text = "โปรแกรมกำลังทำงาน"
            Me.TMActive.Enabled = True
        ElseIf Now.Hour > 7 And Now.Hour < 19 Then
            Me.TMTransfer.Enabled = True
            Me.TMActive.Enabled = True
        Else
            Me.TMTransfer.Enabled = False
            Me.TBLink.Text = "โปรแกรมหยุดทำงาน"
            Me.TMActive.Enabled = False
            Me.PBActive.Visible = False
        End If

    End Sub

    Private Sub TMActive_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TMActive.Tick
        If Me.PBActive.Visible = True Then
            Me.PBActive.Visible = False
        Else
            Me.PBActive.Visible = True
        End If
    End Sub

    Public Sub vConnctDataBase()
        'vConnectionStringNEBULA = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Max Pool Size = 10000;Min Pool Size = 5;Data Source =NEBULA;Initial Catalog =BCNP"
        vConnectionStringNEBULA = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =NEBULA;Initial Catalog =BCNP"
        vConnectionNEBULA = New SqlConnection(vConnectionStringNEBULA)
        vConnectionNEBULA.Open()

        'vConnectionStringS02DB = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Max Pool Size = 10000;Min Pool Size = 5;Data Source =192.168.2.2;Initial Catalog =BCNP"
        vConnectionStringS02DB = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =192.168.2.2;Initial Catalog =BCNP"
        vConnectionS02DB = New SqlConnection(vConnectionStringS02DB)
        vConnectionS02DB.Open()
    End Sub

End Class

