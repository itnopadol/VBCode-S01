Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Net.NetworkInformation

Public Class FormTrnDataBranchToNopadol
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

    Private Sub FormTrnDataBranchToNopadol_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        vConnectionStringNEBULA = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =NEBULA;Initial Catalog =BCNP"
        vConnectionNEBULA = New SqlConnection(vConnectionStringNEBULA)
        vConnectionNEBULA.Open()

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
        vQuery = "exec dbo.USP_PTF_PrepareDataTable " & vType & ",'" & vTableName & "','" & vValues & "'"
        cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        cmdS02DB.ExecuteNonQuery()

    End Sub

    Public Sub DeleteTableHeadOffice(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        vQuery = "exec dbo.USP_PTF_DeleteDataTable " & vType & ",'" & vTableName & "','" & vValue & "' "
        cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
        cmdNEBULA.ExecuteNonQuery()
    End Sub

    Public Sub DropTableBranch(ByVal vTableName As String)
        vQuery = "exec dbo.USP_PTF_DropDataTable '" & vTableName & "'"
        cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        cmdS02DB.ExecuteNonQuery()
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
        cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
        cmdS02DB.ExecuteNonQuery()
    End Sub

    Public Sub TransferDoc(ByVal vGetTableName As String, ByVal vGetDocNo As String, ByVal vIsCancel As Integer)
        Dim vCountTransfer As Integer
        Dim i As Integer

        On Error Resume Next

        '===============================================================================================================
        '===============================================================================================================

        If vGetTableName = "BCQUOTATION" Then
            vTableName = "BCQUOTATION"
            vDocSearchType = 0

            Call CheckDataBranch(3, vTableName, vGetDocNo)
            Call CheckDataHeadOffice(3, vTableName, vGetDocNo)
            If vBranchExist > 0 Then

                If vMemBranchIsCancel = 1 Then
                    If vMemBranchIsCancel = vMemHeadOfficeIsCancel Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
                    End If
                End If

                If vMemBranchLastEditorCode <> "" Then
                    If vMemBranchLastEditDateT <= vMemHeadOfficeLastEditDateT Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
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
            vQuery = "select * from tempdb.dbo.BCQuotation where docno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "Data")
            dtS02DB = dsS02DB.Tables("Data")
            If dsS02DB.Tables("Data").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertBackOrder '" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("ArCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("CreditDay") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DueDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("Subject") & "','" & dsS02DB.Tables("Data").Rows(i).Item("Enclosures") & "'," & dsS02DB.Tables("Data").Rows(i).Item("Validity") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("SaleCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxRate") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription1") & "','" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription2") & "'," & dsS02DB.Tables("Data").Rows(i).Item("BillStatus") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExpireCredit") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CallCheckDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ExpireDate") & "'," & dsS02DB.Tables("Data").Rows(i).Item("CustomerAssert") & "," & dsS02DB.Tables("Data").Rows(i).Item("AssertStatus") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ContactCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsS02DB.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsS02DB.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("NetAmount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("RecurName") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TotalTransport") & "," & dsS02DB.Tables("Data").Rows(i).Item("BillType") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsImport") & "," & dsS02DB.Tables("Data").Rows(i).Item("PRINTCOUNT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("USERGROUP") & "'," & dsS02DB.Tables("Data").Rows(i).Item("DELIVERYDAY") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DELIVERYDATE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("METHODEPAYBILL") & "','" & dsS02DB.Tables("Data").Rows(i).Item("METHODEPAYBILL2") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

            '=======================================================================================================================================================================================================

            vTableName = "BCQUOTATIONSUB"
            Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
            Call DeleteTableHeadOffice(vDocSearchType, vTableName, vGetDocNo)

            vCountTransfer = 0
            vQuery = "select * from tempdb.dbo.BCQuotationSub where docno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "Data")
            dtS02DB = dsS02DB.Tables("Data")
            If dsS02DB.Tables("Data").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertBackOrderSub '" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ArCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ItemName") & "','" & dsS02DB.Tables("Data").Rows(i).Item("WHCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ShelfCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("Qty") & "," & dsS02DB.Tables("Data").Rows(i).Item("RemainQty") & "," & dsS02DB.Tables("Data").Rows(i).Item("Price") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsS02DB.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("Amount") & "," & dsS02DB.Tables("Data").Rows(i).Item("NetAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("HomeAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("TransState") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("UnitCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("OppositePrice2") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsS02DB.Tables("Data").Rows(i).Item("OppositeQty") & "," & dsS02DB.Tables("Data").Rows(i).Item("ItemType") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExceptTax") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsS02DB.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CategoryCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("GroupCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("BrandCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("TypeCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("FormatCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("BarCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("IsPromotion") & "," & dsS02DB.Tables("Data").Rows(i).Item("PremiumStatus") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("QuoStatusCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("IsConditionSend") & "," & dsS02DB.Tables("Data").Rows(i).Item("TransportAmount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("StyleCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TAXRATE") & "," & dsS02DB.Tables("Data").Rows(i).Item("PACKINGRATE1") & "," & dsS02DB.Tables("Data").Rows(i).Item("PACKINGRATE2") & "," & dsS02DB.Tables("Data").Rows(i).Item("BEHINDINDEX") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("USERGROUP") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ITEMSETCODE") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITDATET") & "' "
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

        End If

        '===============================================================================================================
        '===============================================================================================================

        If vGetTableName = "BCSTKREQUEST" Then
            vTableName = "BCSTKREQUEST"
            vDocSearchType = 0

            Call CheckDataBranch(3, vTableName, vGetDocNo)
            Call CheckDataHeadOffice(3, vTableName, vGetDocNo)
            If vBranchExist > 0 Then

                If vMemBranchIsCancel = 1 Then
                    If vMemBranchIsCancel = vMemHeadOfficeIsCancel Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
                    End If
                End If

                If vMemBranchLastEditorCode <> "" Then
                    If vMemBranchLastEditDateT <= vMemHeadOfficeLastEditDateT Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
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
            vQuery = "select * from tempdb.dbo.BCStkRequest where docno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "Data")
            dtS02DB = dsS02DB.Tables("Data")
            If dsS02DB.Tables("Data").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertStkRequest '" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxRate") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsS02DB.Tables("Data").Rows(i).Item("WorkMan") & "'," & dsS02DB.Tables("Data").Rows(i).Item("BillStatus") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsS02DB.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("RecurName") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsS02DB.Tables("Data").Rows(i).Item("IsImport") & "," & dsS02DB.Tables("Data").Rows(i).Item("PRINTCOUNT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("APCODE") & "'," & dsS02DB.Tables("Data").Rows(i).Item("CREDITDAY") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DUEDATE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CONTACTCODE") & "'," & dsS02DB.Tables("Data").Rows(i).Item("SUMOFITEMAMOUNT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DISCOUNTWORD") & "'," & dsS02DB.Tables("Data").Rows(i).Item("DISCOUNTAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("AFTERDISCOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("BEFORETAXAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("TAXAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("TOTALAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("DISCOUNTCASE") & "," & dsS02DB.Tables("Data").Rows(i).Item("EXCEPTTAXAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("ZEROTAXAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("NETAMOUNT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CURRENCYCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("EXCHANGERATE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("SALEAPPROVED") & "','" & dsS02DB.Tables("Data").Rows(i).Item("WORKSITE") & "'," & dsS02DB.Tables("Data").Rows(i).Item("AMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("Buyer") & "'"
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

            '=======================================================================================================================================================================================================

            vTableName = "BCSTKREQUESTSUB"
            Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
            Call DeleteTableHeadOffice(vDocSearchType, vTableName, vGetDocNo)

            Dim vRefLineNumber As Integer

            vCountTransfer = 0
            vQuery = "select * from tempdb.dbo.BCStkRequestSub where docno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "Data")
            dtS02DB = dsS02DB.Tables("Data")
            If dsS02DB.Tables("Data").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                    If Not IsDBNull(dsS02DB.Tables("Data").Rows(i).Item("RefLineNumber")) Then
                        vRefLineNumber = dsS02DB.Tables("Data").Rows(i).Item("RefLineNumber")
                    Else
                        vRefLineNumber = 0
                    End If
                    vQuery = "exec dbo.USP_PTF_InsertStkRequestSub '" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsS02DB.Tables("Data").Rows(i).Item("Qty") & "," & dsS02DB.Tables("Data").Rows(i).Item("RemainQty") & "," & dsS02DB.Tables("Data").Rows(i).Item("TransState") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & "," & dsS02DB.Tables("Data").Rows(i).Item("Priority") & "," & dsS02DB.Tables("Data").Rows(i).Item("WantDay") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("WantDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsS02DB.Tables("Data").Rows(i).Item("OppositeQty") & "," & dsS02DB.Tables("Data").Rows(i).Item("ItemType") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExceptTax") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CategoryCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("GroupCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("BrandCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("TypeCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("FormatCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsS02DB.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("BarCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ItemName") & "','" & dsS02DB.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("StyleCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TAXRATE") & "," & dsS02DB.Tables("Data").Rows(i).Item("PACKINGRATE1") & "," & dsS02DB.Tables("Data").Rows(i).Item("PACKINGRATE2") & "," & dsS02DB.Tables("Data").Rows(i).Item("BEHINDINDEX") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("APCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("WHCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("SHELFCODE") & "'," & dsS02DB.Tables("Data").Rows(i).Item("PRICE") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DISCOUNTWORD") & "'," & dsS02DB.Tables("Data").Rows(i).Item("DISCOUNTAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("AMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("NETAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("HOMEAMOUNT") & "," & vRefLineNumber & "," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITDATET") & "' "
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

        End If

        '===============================================================================================================
        '===============================================================================================================

        If vGetTableName = "BCSTKTRANSFER" Then
            vTableName = "BCSTKTRANSFER"
            vDocSearchType = 0

            Call CheckDataBranch(3, vTableName, vGetDocNo)
            Call CheckDataHeadOffice(3, vTableName, vGetDocNo)
            If vBranchExist > 0 Then

                If vMemBranchIsCancel = 1 Then
                    If vMemBranchIsCancel = vMemHeadOfficeIsCancel Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
                    End If
                End If

                If vMemBranchLastEditorCode <> "" Then
                    If vMemBranchLastEditDateT <= vMemHeadOfficeLastEditDateT Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
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
            vQuery = "select * from tempdb.dbo.BCStktransfer where docno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "Data")
            dtS02DB = dsS02DB.Tables("Data")
            If dsS02DB.Tables("Data").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertSTKTransfer '" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "',1,'" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsS02DB.Tables("Data").Rows(i).Item("SumOfQty") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCompleteSave") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("RecurName") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsS02DB.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISPOS") & "," & dsS02DB.Tables("Data").Rows(i).Item("SUMOFAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DepositNo") & "'"
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

            '=======================================================================================================================================================================================================

            vTableName = "BCSTKTRANSFSUB"
            Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
            Call DeleteTableHeadOffice(vDocSearchType, vTableName, vGetDocNo)

            vCountTransfer = 0
            vQuery = "select * from tempdb.dbo.BCStktransfSub where docno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "Data")
            dtS02DB = dsS02DB.Tables("Data")
            If dsS02DB.Tables("Data").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertSTKTransferSub '" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsS02DB.Tables("Data").Rows(i).Item("FromWH") & "','" & dsS02DB.Tables("Data").Rows(i).Item("FromShelf") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ToWH") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ToShelf") & "'," & dsS02DB.Tables("Data").Rows(i).Item("Qty") & "," & dsS02DB.Tables("Data").Rows(i).Item("Price") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumOfCost") & "," & dsS02DB.Tables("Data").Rows(i).Item("Amount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("RefNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsS02DB.Tables("Data").Rows(i).Item("OppositeQty") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CategoryCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("GroupCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("BrandCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("TypeCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("FormatCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("BarCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TransState") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & "," & dsS02DB.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("LotOutNumber") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LotInNumber") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LotExpireDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LotMyDesc") & "','" & dsS02DB.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("StyleCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("BEHINDINDEX") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISPOS") & "," & dsS02DB.Tables("Data").Rows(i).Item("REFLINENUMBER") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

            '=======================================================================================================================================================================================================

            vTableName = "BCSTKTRANSFSUB3"
            Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
            Call DeleteTableHeadOffice(vDocSearchType, vTableName, vGetDocNo)

            vCountTransfer = 0
            vQuery = "select * from tempdb.dbo.BCStktransfSub3 where docno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "Data")
            dtS02DB = dsS02DB.Tables("Data")
            If dsS02DB.Tables("Data").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertSTKTransferSub3 " & dsS02DB.Tables("Data").Rows(i).Item("BehindIndex") & "," & dsS02DB.Tables("Data").Rows(i).Item("MyType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsS02DB.Tables("Data").Rows(i).Item("WHCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ShelfCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("Qty") & "," & dsS02DB.Tables("Data").Rows(i).Item("Price") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumOfCost") & "," & dsS02DB.Tables("Data").Rows(i).Item("Amount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("RefNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsS02DB.Tables("Data").Rows(i).Item("OppositeQty") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("fromwhcode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("fromshelfcode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("reftrlinenum") & "," & dsS02DB.Tables("Data").Rows(i).Item("FixUnitCost") & "," & dsS02DB.Tables("Data").Rows(i).Item("FixUnitQty") & "," & dsS02DB.Tables("Data").Rows(i).Item("AVERAGECOST") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("LotNumber") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LotExpireDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LotMyDesc") & "','" & dsS02DB.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("StyleCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TAXRATE") & "," & dsS02DB.Tables("Data").Rows(i).Item("PACKINGRATE1") & "," & dsS02DB.Tables("Data").Rows(i).Item("PACKINGRATE2") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISCANCEL") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISPOS") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISPROCESS") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISLOCKCOST") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

        End If


        '===============================================================================================================
        '===============================================================================================================

        If vGetTableName = "BCSALEORDER" Then
            vTableName = "BCSALEORDER"
            vDocSearchType = 0

            Call CheckDataBranch(3, vTableName, vGetDocNo)
            Call CheckDataHeadOffice(3, vTableName, vGetDocNo)
            If vBranchExist > 0 Then

                If vMemBranchIsCancel = 1 Then
                    If vMemBranchIsCancel = vMemHeadOfficeIsCancel Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
                    End If
                End If

                If vMemBranchLastEditorCode <> "" Then
                    If vMemBranchLastEditDateT <= vMemHeadOfficeLastEditDateT Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
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
            vQuery = "select * from tempdb.dbo.BCSaleOrder where docno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "Data")
            dtS02DB = dsS02DB.Tables("Data")
            If dsS02DB.Tables("Data").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertSaleOrder '" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxType") & "," & dsS02DB.Tables("Data").Rows(i).Item("BillType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("ArCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("CreditDay") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DueDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DeliveryAddr") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxRate") & ",1,'" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsS02DB.Tables("Data").Rows(i).Item("RefDocNo") & "'," & dsS02DB.Tables("Data").Rows(i).Item("BillStatus") & "," & dsS02DB.Tables("Data").Rows(i).Item("SOStatus") & "," & dsS02DB.Tables("Data").Rows(i).Item("HoldingStatus") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CrAuthorizeMan") & "','" & dsS02DB.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ContactCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsS02DB.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsS02DB.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("NetAmount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCompleteSave") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("RecurName") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsS02DB.Tables("Data").Rows(i).Item("IsConditionSend") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("OwnReceive") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CarLicense") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ApproveCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ApproveDateTime") & "'," & dsS02DB.Tables("Data").Rows(i).Item("IsUseRobotSale") & "," & dsS02DB.Tables("Data").Rows(i).Item("TotalTransport") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumOtherValue") & "," & dsS02DB.Tables("Data").Rows(i).Item("READYFORPAY") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("TimeTransport") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CarType") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CondPayCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("PrintCount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("PORefNo") & "'," & dsS02DB.Tables("Data").Rows(i).Item("IsImport") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("USERGROUP") & "','" & dsS02DB.Tables("Data").Rows(i).Item("METHODEPAYBILL") & "','" & dsS02DB.Tables("Data").Rows(i).Item("METHODEPAYBILL2") & "'," & dsS02DB.Tables("Data").Rows(i).Item("DELIVERYDAY") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DELIVERYDATE") & "'," & dsS02DB.Tables("Data").Rows(i).Item("DELIVERYSTATUS") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("PROMOTIONCODE") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

            '=======================================================================================================================================================================================================

            vTableName = "BCSALEORDERSUB"
            Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
            Call DeleteTableHeadOffice(vDocSearchType, vTableName, vGetDocNo)

            vCountTransfer = 0
            vQuery = "select * from tempdb.dbo.BCSaleOrderSub where docno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "DataSub")
            dtS02DB = dsS02DB.Tables("DataSub")
            If dsS02DB.Tables("DataSub").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("DataSub").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertSaleOrderSub '" & dsS02DB.Tables("datasub").Rows(i).Item("DocNo") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("TaxType") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("ItemCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("ArCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("DepartCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("SaleCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("MyDescription") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("ItemName") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("WHCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("ShelfCode") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("Qty") & "," & dsS02DB.Tables("datasub").Rows(i).Item("RemainQty") & "," & dsS02DB.Tables("datasub").Rows(i).Item("Price") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("DiscountWord") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("DiscountAmount") & "," & dsS02DB.Tables("datasub").Rows(i).Item("Amount") & "," & dsS02DB.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsS02DB.Tables("datasub").Rows(i).Item("HomeAmount") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("UnitCode") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("OppositePrice2") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("OppositeUnit") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("OppositeQty") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("StkReserveNo") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("SOStatus") & "," & dsS02DB.Tables("datasub").Rows(i).Item("HoldingStatus") & "," & dsS02DB.Tables("datasub").Rows(i).Item("RefType") & "," & dsS02DB.Tables("datasub").Rows(i).Item("ItemType") & "," & dsS02DB.Tables("datasub").Rows(i).Item("ExceptTax") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("AllocateCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("ProjectCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("CurrencyCode") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("ExchangeRate") & "," & dsS02DB.Tables("datasub").Rows(i).Item("TransState") & "," & dsS02DB.Tables("datasub").Rows(i).Item("IsCancel") & "," & dsS02DB.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("CategoryCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("GroupCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("BrandCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("TypeCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("FormatCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("BarCode") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("IsPromotion") & "," & dsS02DB.Tables("datasub").Rows(i).Item("PremiumStatus") & "," & dsS02DB.Tables("datasub").Rows(i).Item("IsUseRobotSale") & "," & dsS02DB.Tables("datasub").Rows(i).Item("IsConditionSend") & "," & dsS02DB.Tables("datasub").Rows(i).Item("TransportAmount") & "," & dsS02DB.Tables("datasub").Rows(i).Item("OtherValue") & "," & dsS02DB.Tables("datasub").Rows(i).Item("REMAINPAYQTY") & "," & dsS02DB.Tables("datasub").Rows(i).Item("ISPAYMENT") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("Colorcode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("SizeCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("StyleCode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("itemsetcode") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("PORefNo") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("BEHINDINDEX") & "," & dsS02DB.Tables("datasub").Rows(i).Item("TAXRATE") & "," & dsS02DB.Tables("datasub").Rows(i).Item("PACKINGRATE1") & "," & dsS02DB.Tables("datasub").Rows(i).Item("PACKINGRATE2") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("USERGROUP") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("ISPROCESS") & "," & dsS02DB.Tables("datasub").Rows(i).Item("ISLOCKCOST") & "," & dsS02DB.Tables("datasub").Rows(i).Item("DELIVERYQTY") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("PROMOTIONCODE") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("MASTERITEMCODE") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

        End If

        '===============================================================================================================
        '===============================================================================================================

        If vGetTableName = "BCARDEPOSIT" Then
            vTableName = "BCARDEPOSIT"
            vDocSearchType = 0

            Call CheckDataBranch(3, vTableName, vGetDocNo)
            Call CheckDataHeadOffice(3, vTableName, vGetDocNo)

            If vBranchExist > 0 Then

                'Call CheckCountDepositUse(vGetDocNo)
                'If vCountDepositUse = 0 Then
                '    vSendTrnStatus = 2
                '    GoTo LineEnd
                'End If


                If vMemBranchIsCancel = 1 Then
                    If vMemBranchIsCancel = vMemHeadOfficeIsCancel Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
                    End If
                End If


                'If vMemBranchLastEditorCode <> "" Then
                '    If vMemBranchLastEditDateT <= vMemHeadOfficeLastEditDateT Then
                '        GoTo LineEnd
                '    End If
                'End If
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

            '=======================================================================================================================================================================================================

            vTableName = "BCARDEPOSITUSE"
            Call PrepareData(3, vTableName, vGetDocNo)
            Call DeleteTableHeadOffice(3, vTableName, vGetDocNo)

            vCountTransfer = 0
            vQuery = "select * from tempdb.dbo.BCARDepositUse where depositno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "DataSub")
            dtS02DB = dsS02DB.Tables("DataSub")
            If dsS02DB.Tables("DataSub").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("DataSub").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertARDepositUse '" & dsS02DB.Tables("datasub").Rows(i).Item("DocNo") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("DepositNo") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("MyDescription") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("Balance") & "," & dsS02DB.Tables("datasub").Rows(i).Item("Amount") & "," & dsS02DB.Tables("datasub").Rows(i).Item("DeposTaxType") & "," & dsS02DB.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsS02DB.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("CurrencyCode") & "'," & dsS02DB.Tables("datasub").Rows(i).Item("DPExchangeRate") & "," & dsS02DB.Tables("datasub").Rows(i).Item("NewExchangeRate") & "," & dsS02DB.Tables("datasub").Rows(i).Item("HomeAmount1") & "," & dsS02DB.Tables("datasub").Rows(i).Item("HomeAmount2") & "," & dsS02DB.Tables("datasub").Rows(i).Item("ExchangeProfit") & "," & dsS02DB.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsS02DB.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

        End If

        '===============================================================================================================
        '===============================================================================================================

        If vGetTableName = "BCAPINVOICE" Then
            vTableName = "BCAPINVOICE"
            vDocSearchType = 0

            Call CheckDataBranch(3, vTableName, vGetDocNo)
            Call CheckDataHeadOffice(3, vTableName, vGetDocNo)
            If vBranchExist > 0 Then

                If vMemBranchIsCancel = 1 Then
                    If vMemBranchIsCancel = vMemHeadOfficeIsCancel Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
                    End If
                End If

                If vMemBranchLastEditorCode <> "" Then
                    If vMemBranchLastEditDateT <= vMemHeadOfficeLastEditDateT Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
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
            vQuery = "select * from tempdb.dbo.BCAPInvoice where docno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "Data")
            dtS02DB = dsS02DB.Tables("Data")
            If dsS02DB.Tables("Data").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertAPInvoice '" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("TaxNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("ApCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("CreditDay") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DueDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("StatementDate") & "'," & dsS02DB.Tables("Data").Rows(i).Item("StatementState") & "," & dsS02DB.Tables("Data").Rows(i).Item("TaxRate") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsS02DB.Tables("Data").Rows(i).Item("BillType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ContactCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsS02DB.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsS02DB.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("PettyCashAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumCashAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumChqAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumBankAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("DepositIncTax") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumOfDeposit1") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumOfDeposit2") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumOfWTax") & "," & dsS02DB.Tables("Data").Rows(i).Item("NetDebtAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("HomeAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("OtherIncome") & "," & dsS02DB.Tables("Data").Rows(i).Item("OtherExpense") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExcessAmount1") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExcessAmount2") & "," & dsS02DB.Tables("Data").Rows(i).Item("BillBalance") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("GLFormat") & "'," & dsS02DB.Tables("Data").Rows(i).Item("GLStartPosting") & "," & dsS02DB.Tables("Data").Rows(i).Item("GRBillStatus") & "," & dsS02DB.Tables("Data").Rows(i).Item("GRIRBillStatus") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsPostGL") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCreditNote") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsDebitNote") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCompleteSave") & "," & dsS02DB.Tables("Data").Rows(i).Item("GLTransData") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("RecurName") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ExchangeProfit") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CancelDateTime") & "','" & dsS02DB.Tables("Data").Rows(i).Item("Remark1") & "','" & dsS02DB.Tables("Data").Rows(i).Item("Remark2") & "','" & dsS02DB.Tables("Data").Rows(i).Item("Remark3") & "','" & dsS02DB.Tables("Data").Rows(i).Item("Remark4") & "','" & dsS02DB.Tables("Data").Rows(i).Item("Remark5") & "','" & dsS02DB.Tables("Data").Rows(i).Item("OwnerTransport") & "'," & dsS02DB.Tables("Data").Rows(i).Item("PayBillAmount") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("RefDocNo") & "'," & dsS02DB.Tables("Data").Rows(i).Item("IsImport") & "," & dsS02DB.Tables("Data").Rows(i).Item("PRINTCOUNT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("JOBNO") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ISGRBILL") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISIRBILL") & "," & dsS02DB.Tables("Data").Rows(i).Item("SUMOFWTAXCASH") & "," & dsS02DB.Tables("Data").Rows(i).Item("SUMBASEWTAXCASH") & "," & dsS02DB.Tables("Data").Rows(i).Item("SUMCOMWTAXCASH") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & " "
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

            '=======================================================================================================================================================================================================

            vTableName = "BCAPINVOICESUB"
            Call PrepareData(vDocSearchType, vTableName, vGetDocNo)
            Call DeleteTableHeadOffice(vDocSearchType, vTableName, vGetDocNo)

            vCountTransfer = 0
            vQuery = "select * from tempdb.dbo.BCAPInvoiceSub where docno ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "Data")
            dtS02DB = dsS02DB.Tables("Data")
            If dsS02DB.Tables("Data").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertAPInvoiceSub " & dsS02DB.Tables("Data").Rows(i).Item("BehindIndex") & "," & dsS02DB.Tables("Data").Rows(i).Item("MyType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DocNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("TaxNo") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DocDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ApCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ItemName") & "','" & dsS02DB.Tables("Data").Rows(i).Item("WHCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ShelfCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("CNQty") & "," & dsS02DB.Tables("Data").Rows(i).Item("GRRemainQty") & "," & dsS02DB.Tables("Data").Rows(i).Item("Qty") & "," & dsS02DB.Tables("Data").Rows(i).Item("Price") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsS02DB.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("Amount") & "," & dsS02DB.Tables("Data").Rows(i).Item("NetAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("HomeAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("BalanceAmount") & "," & dsS02DB.Tables("Data").Rows(i).Item("SumOfExpCost") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsS02DB.Tables("Data").Rows(i).Item("OppositeQty") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("PORefNo") & "','" & dsS02DB.Tables("Data").Rows(i).Item("IRRefNo") & "'," & dsS02DB.Tables("Data").Rows(i).Item("StockType") & "," & dsS02DB.Tables("Data").Rows(i).Item("ExceptTax") & "," & dsS02DB.Tables("Data").Rows(i).Item("TransState") & "," & dsS02DB.Tables("Data").Rows(i).Item("IsCancel") & "," & dsS02DB.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("BarCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CategoryCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("GroupCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("BrandCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("TypeCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("FormatCode") & "'," & dsS02DB.Tables("Data").Rows(i).Item("IsPromotion") & "," & dsS02DB.Tables("Data").Rows(i).Item("PORefLinenum") & "," & dsS02DB.Tables("Data").Rows(i).Item("FixUnitCost") & "," & dsS02DB.Tables("Data").Rows(i).Item("FixUnitQty") & "," & dsS02DB.Tables("Data").Rows(i).Item("AVERAGECOST") & "," & dsS02DB.Tables("Data").Rows(i).Item("StatusReceive") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("LotNumber") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LotExpireDate") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LotMyDesc") & "'," & dsS02DB.Tables("Data").Rows(i).Item("SumOfCost") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("StyleCode") & "','" & dsS02DB.Tables("Data").Rows(i).Item("JobNo") & "'," & dsS02DB.Tables("Data").Rows(i).Item("TAXRATE") & "," & dsS02DB.Tables("Data").Rows(i).Item("PACKINGRATE1") & "," & dsS02DB.Tables("Data").Rows(i).Item("PACKINGRATE2") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISPROCESS") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISLOCKCOST") & "," & dsS02DB.Tables("Data").Rows(i).Item("DISCCASHCARD") & "," & dsS02DB.Tables("Data").Rows(i).Item("WTAXAMOUNT") & "," & dsS02DB.Tables("Data").Rows(i).Item("BASEWTAX") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITDATET") & "' "
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

        End If

        '===============================================================================================================
        '===============================================================================================================

        If vGetTableName = "BCCOUPONRECEIVE" Then
            vTableName = "BCCOUPONRECEIVE"
            vDocSearchType = 4

            Call CheckDataBranch(vDocSearchType, vTableName, vGetDocNo)
            Call CheckDataHeadOffice(vDocSearchType, vTableName, vGetDocNo)
            If vBranchExist > 0 Then

                If vMemBranchIsCancel = 1 Then
                    If vMemBranchIsCancel = vMemHeadOfficeIsCancel Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
                    End If
                End If

                If vMemBranchLastEditorCode <> "" Then
                    If vMemBranchLastEditDateT <= vMemHeadOfficeLastEditDateT Then
                        vSendTrnStatus = 2
                        GoTo LineEnd
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
            vQuery = "select * from tempdb.dbo.BCCouponReceive where CouponNo ='" & vGetDocNo & "'"
            daS02DB = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB = New DataSet
            daS02DB.Fill(dsS02DB, "Data")
            dtS02DB = dsS02DB.Tables("Data")
            If dsS02DB.Tables("Data").Rows.Count > 0 Then
                For i = 0 To dsS02DB.Tables("Data").Rows.Count - 1
                    vQuery = "exec dbo.USP_PTF_InsertCouponReceive '" & dsS02DB.Tables("Data").Rows(i).Item("COUPONCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("COUPONTYPE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("COUPONNO") & "','" & dsS02DB.Tables("Data").Rows(i).Item("TOCOUPONNO") & "'," & dsS02DB.Tables("Data").Rows(i).Item("COUPONCOUNT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("DOCNO") & "','" & dsS02DB.Tables("Data").Rows(i).Item("BOOK") & "'," & dsS02DB.Tables("Data").Rows(i).Item("COUPONVAL") & "," & dsS02DB.Tables("Data").Rows(i).Item("LINENUMBER") & "," & dsS02DB.Tables("Data").Rows(i).Item("PERCENTDISC") & "," & dsS02DB.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsS02DB.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsS02DB.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                    cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
                    cmdNEBULA.ExecuteNonQuery()
                Next
            End If
            Call DropTableBranch(vTableName)
            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

        End If

        '===============================================================================================================
        '===============================================================================================================

        If vGetTableName = "TB_NP_TransferDocLogs" And vIsCancel = 1 Then
            vTableName = "TB_NP_TransferDocLogs"
            vDocSearchType = 5

            vQuery = "delete dbo.bcardeposituse where docno = '" & vGetDocNo & "'"
            cmdNEBULA = New SqlCommand(vQuery, vConnectionNEBULA)
            cmdNEBULA.ExecuteNonQuery()

            Call InsertLogs(vDocSearchType, vTableName, vGetDocNo)
            vSendTrnStatus = 1

        End If

        '===============================================================================================================
        '===============================================================================================================

LineEnd:
    End Sub

    Public Sub CheckCountDepositUse(ByVal vDepositNo As String)
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

    Public Sub vConnctDataBase()

        vConnectionStringNEBULA = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =NEBULA;Initial Catalog =BCNP"
        vConnectionNEBULA = New SqlConnection(vConnectionStringNEBULA)
        vConnectionNEBULA.Open()

        vConnectionStringS02DB = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Data Source =192.168.2.2;Initial Catalog =BCNP"
        vConnectionS02DB = New SqlConnection(vConnectionStringS02DB)
        vConnectionS02DB.Open()
    End Sub

    Private Sub TMTransfer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TMTransfer.Tick
        Dim i As Integer
        Dim vListItem As ListViewItem
        Dim n As Integer
        Dim a As Integer

        On Error Resume Next

        Call vConnctDataBase()

        vQuery = "exec dbo.USP_NP_SearchDocNotTransfer 0"
        daS02DB1 = New SqlDataAdapter(vQuery, vConnectionS02DB)
        dsS02DB1 = New DataSet
        daS02DB1.Fill(dsS02DB1, "Docno")
        dtS02DB1 = dsS02DB1.Tables("Docno")
        If dtS02DB1.Rows.Count > 0 Then
            For i = 0 To dtS02DB1.Rows.Count - 1
                vTableName = dtS02DB1.Rows(i).Item("tablename")
                vModuleID = dtS02DB1.Rows(i).Item("moduleid")
                vDocNo = dtS02DB1.Rows(i).Item("Docno")
                vMemCancelStatus = dtS02DB1.Rows(i).Item("cancelstatus")

                Call TransferDoc(vTableName, vDocNo, vMemCancelStatus)

                If vCheckError = 0 Then
                    vQuery = "exec dbo.USP_NP_UpdateDocTransfered '" & vDocNo & "'"
                    cmdS02DB = New SqlCommand(vQuery, vConnectionS02DB)
                    cmdS02DB.ExecuteNonQuery()
                End If

            Next

            Me.ListViewListTrn.Items.Clear()

            vQuery = "exec dbo.USP_NP_SearchDocNotTransfer 1"
            daS02DB2 = New SqlDataAdapter(vQuery, vConnectionS02DB)
            dsS02DB2 = New DataSet
            daS02DB2.Fill(dsS02DB2, "Docno")
            dtS02DB2 = dsS02DB2.Tables("Docno")
            If dtS02DB2.Rows.Count > 0 Then
                For i = 0 To dtS02DB2.Rows.Count - 1
                    n = n + 1
                    vListItem = Me.ListViewListTrn.Items.Add(n)
                    vListItem.SubItems.Add(0).Text = dtS02DB2.Rows(i).Item("tablename")
                    vListItem.SubItems.Add(1).Text = dtS02DB2.Rows(i).Item("Docno")
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
    End Sub


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
End Class