Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.SqlServer
Imports Microsoft.VisualBasic
Imports ASP
Partial Class TransferData
    Inherits System.Web.UI.Page

    Dim vConnectionStringFrom As String
    Dim vConnectionFrom As SqlConnection
    Dim daFrom As SqlDataAdapter
    Dim dsFrom As DataSet
    Dim dtFrom As DataTable
    Dim dvFrom As DataView
    Dim vQuery As String
    Dim vCommandFrom As SqlCommand

    Dim vConnectionStringTo As String
    Dim vConnectionTo As SqlConnection
    Dim daTo As SqlDataAdapter
    Dim dsTo As DataSet
    Dim dtTo As DataTable
    Dim dvTo As DataView
    Dim vCommandTo As SqlCommand

    Dim vConnectionString As String
    Dim vConnection As SqlConnection
    Dim da As SqlDataAdapter
    Dim ds As DataSet
    Dim vdt As DataTable
    Dim dt As New DataTable
    Dim dv As DataView

    Dim vFromServer As String
    Dim vFromDatabase As String
    Dim vFromUserID As String
    Dim vFromPassword As String

    Dim vToServer As String
    Dim vToDatabase As String
    Dim vToUserID As String
    Dim vToPassword As String

    Dim vTrnUserID As String

    Dim vLocation As String
    Dim vListIndex As Integer
    Dim vNewDocList As Integer
    Dim vDocList As Integer
    Dim vNewTransType As Integer
    Dim vTransType As Integer

    Dim vCheckExist As Integer
    Dim vAccess As Integer
    Dim vDepartment As String
    Dim vPrgID As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Request.QueryString.Count > 0 Then
            vFromServer = Request.QueryString("Server").ToString()
            vFromDatabase = Request.QueryString("DataBase").ToString()
            vFromUserID = Request.QueryString("User").ToString()
            vFromPassword = Request.QueryString("Password").ToString()
            vTrnUserID = Request.QueryString("TrnUserID").ToString()
        End If

        Session("_List") = Me.DropDownList1.SelectedIndex
        vDocList = Session("_List")
        vNewDocList = Session("_NewList")

        Session("_TransTypeList") = Me.DropDownList2.SelectedIndex
        vTransType = Session("_TransTypeList")
        vNewTransType = Session("_NewTransTypeList")

        If vDocList <> vNewDocList Or vTransType <> vNewTransType Then
            Me.Label15.Visible = False
            Me.Label16.Visible = False
            Me.TextBox9.Text = ""
            Me.TextBox10.Text = ""
            Me.TextBox11.Text = ""
            Me.TextBox12.Text = ""
            Me.TextBox9.Visible = False
            Me.TextBox10.Visible = False
            Me.TextBox11.Visible = False
            Me.TextBox12.Visible = False
            Me.Button1.Visible = False
            Me.Button3.Visible = False
            Me.Button4.Visible = False

            Call CreateDataSource()
            Call DataBind()
        End If

        If Not IsPostBack Then
            Me.TextBox1.Text = vFromServer
            Me.TextBox2.Text = vFromDatabase
            Me.TextBox3.Text = vFromUserID
            Me.TextBox4.Text = vFromPassword
            Me.TextBox13.Text = vTrnUserID
        End If

        vFromServer = Me.TextBox1.Text
        vFromDatabase = Me.TextBox2.Text
        vFromUserID = Me.TextBox3.Text
        vFromPassword = Me.TextBox4.Text

        vConnectionStringFrom = "Persist Security Info =False;User ID =" & vFromUserID & ";Password =" & vFromPassword & ";Data Source =" & vFromServer & ";Initial Catalog =BCNP"
        vConnectionFrom = New SqlConnection(vConnectionStringFrom)
        vConnectionFrom.Open()

        vToServer = Me.TextBox5.Text
        vToDatabase = Me.TextBox6.Text
        vToUserID = Me.TextBox7.Text
        vToPassword = Me.TextBox8.Text

        vConnectionStringTo = "Persist Security Info =False;User ID =" & vToUserID & ";Password =" & vToPassword & ";Data Source =192.168.2.2;Initial Catalog =BCNP"
        vConnectionTo = New SqlConnection(vConnectionStringTo)
        vConnectionTo.Open()


        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand

        vQuery = "exec dbo.USP_NP_AccessWebProgram '" & vTrnUserID & "' "
        daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
        dsFrom = New DataSet
        daFrom.Fill(dsFrom, "UserAccess")
        dtFrom = dsFrom.Tables("UserAccess")
        If dsFrom.Tables("UserAccess").Rows.Count > 0 Then
            vAccess = dsFrom.Tables("UserAccess").Rows(0).Item("pagestatus")
            vDepartment = dsFrom.Tables("UserAccess").Rows(0).Item("pageid")
            vPrgID = dsFrom.Tables("UserAccess").Rows(0).Item("prgid")
        End If

        If vAccess = 0 Or vPrgID <> "04" Then
            MsgBox("คุณไม่มีสิทธิ์ เข้าใช้งานโปรแกรม กรุณาแจ้งแผนกคอมฯ กรณีต้องการใช้งาน", MsgBoxStyle.Critical, "Send Error Message")
            Me.Visible = False
        End If
    End Sub


    Public Sub PrepareData(ByVal vType As Integer, ByVal vTableName As String, ByVal vValues As String)
        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand


        vQuery = "exec dbo.USP_PTF_PrepareDataTable " & vType & ",'" & vTableName & "','" & vValues & "'"
        With cmd
            .CommandType = CommandType.Text
            .CommandText = vQuery
            .Connection = vConnectionFrom
            .ExecuteNonQuery()
        End With

    End Sub

    Public Sub PrepareDataBranch(ByVal vType As Integer, ByVal vTableName As String, ByVal vValues As String)
        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand


        vQuery = "exec dbo.USP_PTF_PrepareDataTable " & vType & ",'" & vTableName & "','" & vValues & "'"
        With cmd
            .CommandType = CommandType.Text
            .CommandText = vQuery
            .Connection = vConnectionTo
            .ExecuteNonQuery()
        End With

    End Sub

    Public Sub DropTable(ByVal vTableName As String)
        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand

        vQuery = "exec dbo.USP_PTF_DropDataTable '" & vTableName & "'"
        With cmd
            .CommandType = CommandType.Text
            .CommandText = vQuery
            .Connection = vConnectionFrom
            .ExecuteNonQuery()
        End With
    End Sub
    Public Sub DeleteTableBranch(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand
        vQuery = "exec dbo.USP_PTF_DeleteDataTable " & vType & ",'" & vTableName & "','" & vValue & "' "
        With cmd
            .CommandType = CommandType.Text
            .CommandText = vQuery
            .Connection = vConnectionTo
            .ExecuteNonQuery()
        End With

    End Sub

    Public Sub DeleteTable(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand
        vQuery = "exec dbo.USP_PTF_DeleteDataTable " & vType & ",'" & vTableName & "','" & vValue & "' "
        With cmd
            .CommandType = CommandType.Text
            .CommandText = vQuery
            .Connection = vConnectionFrom
            .ExecuteNonQuery()
        End With

    End Sub

    Public Sub DropTableBranch(ByVal vTableName As String)
        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand

        vQuery = "exec dbo.USP_PTF_DropDataTable '" & vTableName & "'"
        With cmd
            .CommandType = CommandType.Text
            .CommandText = vQuery
            .Connection = vConnectionTo
            .ExecuteNonQuery()
        End With
    End Sub


    Public Sub InsertLogs(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand

        Dim vDatasource As String
        Dim vDestination As String
        Dim vUserID As String

        If vType = 0 Then
            vDatasource = "NEBULA"
            vDestination = "S02DB"
        End If

        If vType = 1 Then
            vDatasource = "S02DB"
            vDestination = "NEBULA"
        End If

        vUserID = vTrnUserID

        vQuery = "exec dbo.USP_PTF_InsertTransferDataLogs '" & vDatasource & "','" & vDestination & "','" & vTableName & "','" & vValue & "','" & vUserID & "',1"
        With cmd
            .CommandType = CommandType.Text
            .CommandText = vQuery
            .Connection = vConnectionFrom
            .ExecuteNonQuery()
        End With
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand
        Dim i As Integer
        Dim vDocNo As String
        Dim vItemCode As String
        Dim vType As Integer
        Dim vTableName As String
        Dim vCountTransfer As Integer
        Dim vTransType As Integer

        vTransType = Me.DropDownList2.SelectedIndex
        vType = Me.DropDownList1.SelectedIndex

        '=======================================================================================================================================================================================================
        '=======================================================================================================================================================================================================
        '=======================================================================================================================================================================================================

        If vTransType = 0 Then 'โอนจากสำนักงานใหญ่ไปสาขา
            If vType = 0 Then

                vDocNo = UCase(Me.TextBox9.Text)

                '=======================================================================================================================================================================================================

                vTableName = "BCSaleOrder"
                Call CheckData(0, vTableName, vDocNo)
                If vCheckExist = 0 Then
                    Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                    Call CreateDataSource()
                    Call BindGrid()
                    Me.Button1.Visible = False
                    Me.DataGrid1.Visible = False
                    MsgBox("ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.bcsaleorder where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertSaleOrder '" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxType") & "," & dsFrom.Tables("Data").Rows(i).Item("BillType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("ArCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("CreditDay") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DueDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DeliveryAddr") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxRate") & "," & dsFrom.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("Data").Rows(i).Item("RefDocNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("BillStatus") & "," & dsFrom.Tables("Data").Rows(i).Item("SOStatus") & "," & dsFrom.Tables("Data").Rows(i).Item("HoldingStatus") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CrAuthorizeMan") & "','" & dsFrom.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsFrom.Tables("Data").Rows(i).Item("ContactCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsFrom.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsFrom.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsFrom.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("NetAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCompleteSave") & ",'" & dsFrom.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("RecurName") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsConditionSend") & ",'" & dsFrom.Tables("Data").Rows(i).Item("OwnReceive") & "','" & dsFrom.Tables("Data").Rows(i).Item("CarLicense") & "','" & dsFrom.Tables("Data").Rows(i).Item("ApproveCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ApproveDateTime") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsUseRobotSale") & "," & dsFrom.Tables("Data").Rows(i).Item("TotalTransport") & "," & dsFrom.Tables("Data").Rows(i).Item("SumOtherValue") & "," & dsFrom.Tables("Data").Rows(i).Item("READYFORPAY") & ",'" & dsFrom.Tables("Data").Rows(i).Item("TimeTransport") & "','" & dsFrom.Tables("Data").Rows(i).Item("CarType") & "','" & dsFrom.Tables("Data").Rows(i).Item("CondPayCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("PrintCount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("PORefNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsImport") & ",'" & dsFrom.Tables("Data").Rows(i).Item("USERGROUP") & "','" & dsFrom.Tables("Data").Rows(i).Item("METHODEPAYBILL") & "','" & dsFrom.Tables("Data").Rows(i).Item("METHODEPAYBILL2") & "'," & dsFrom.Tables("Data").Rows(i).Item("DELIVERYDAY") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DELIVERYDATE") & "'," & dsFrom.Tables("Data").Rows(i).Item("DELIVERYSTATUS") & ",'" & dsFrom.Tables("Data").Rows(i).Item("PROMOTIONCODE") & "'," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If

                Me.DataGrid1.Items(0).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(0).Cells(2).Text = "โอนเรียบร้อย"

                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCSaleOrderSub"
                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.bcsaleordersub where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "DataSub")
                dtFrom = dsFrom.Tables("DataSub")
                If dsFrom.Tables("DataSub").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("DataSub").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertSaleOrderSub '" & dsFrom.Tables("datasub").Rows(i).Item("DocNo") & "'," & dsFrom.Tables("datasub").Rows(i).Item("TaxType") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ArCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("SaleCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ItemName") & "','" & dsFrom.Tables("datasub").Rows(i).Item("WHCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ShelfCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("Qty") & "," & dsFrom.Tables("datasub").Rows(i).Item("RemainQty") & "," & dsFrom.Tables("datasub").Rows(i).Item("Price") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("DiscountWord") & "'," & dsFrom.Tables("datasub").Rows(i).Item("DiscountAmount") & "," & dsFrom.Tables("datasub").Rows(i).Item("Amount") & "," & dsFrom.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsFrom.Tables("datasub").Rows(i).Item("HomeAmount") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("UnitCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("OppositePrice2") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("OppositeUnit") & "'," & dsFrom.Tables("datasub").Rows(i).Item("OppositeQty") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("StkReserveNo") & "'," & dsFrom.Tables("datasub").Rows(i).Item("SOStatus") & "," & dsFrom.Tables("datasub").Rows(i).Item("HoldingStatus") & "," & dsFrom.Tables("datasub").Rows(i).Item("RefType") & "," & dsFrom.Tables("datasub").Rows(i).Item("ItemType") & "," & dsFrom.Tables("datasub").Rows(i).Item("ExceptTax") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("AllocateCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("CurrencyCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("ExchangeRate") & "," & dsFrom.Tables("datasub").Rows(i).Item("TransState") & "," & dsFrom.Tables("datasub").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("CategoryCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("GroupCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("BrandCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("TypeCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("FormatCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("BarCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("IsPromotion") & "," & dsFrom.Tables("datasub").Rows(i).Item("PremiumStatus") & "," & dsFrom.Tables("datasub").Rows(i).Item("IsUseRobotSale") & "," & dsFrom.Tables("datasub").Rows(i).Item("IsConditionSend") & "," & dsFrom.Tables("datasub").Rows(i).Item("TransportAmount") & "," & dsFrom.Tables("datasub").Rows(i).Item("OtherValue") & "," & dsFrom.Tables("datasub").Rows(i).Item("REMAINPAYQTY") & "," & dsFrom.Tables("datasub").Rows(i).Item("ISPAYMENT") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("Colorcode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("SizeCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("StyleCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("itemsetcode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("PORefNo") & "'," & dsFrom.Tables("datasub").Rows(i).Item("BEHINDINDEX") & "," & dsFrom.Tables("datasub").Rows(i).Item("TAXRATE") & "," & dsFrom.Tables("datasub").Rows(i).Item("PACKINGRATE1") & "," & dsFrom.Tables("datasub").Rows(i).Item("PACKINGRATE2") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("USERGROUP") & "'," & dsFrom.Tables("datasub").Rows(i).Item("ISPROCESS") & "," & dsFrom.Tables("datasub").Rows(i).Item("ISLOCKCOST") & "," & dsFrom.Tables("datasub").Rows(i).Item("DELIVERYQTY") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("PROMOTIONCODE") & "','" & dsFrom.Tables("datasub").Rows(i).Item("MASTERITEMCODE") & "'," & dsFrom.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If

                Me.DataGrid1.Items(1).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(1).Cells(2).Text = "โอนเรียบร้อย"

                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

            End If

            '=======================================================================================================================================================================================================
            '=======================================================================================================================================================================================================


            If vType = 1 Then
                vDocNo = UCase(Me.TextBox9.Text)

                '=======================================================================================================================================================================================================

                vTableName = "BCPurchaseOrder"
                Call CheckData(0, vTableName, vDocNo)
                If vCheckExist = 0 Then
                    Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                    Call CreateDataSource()
                    Call BindGrid()
                    Me.Button1.Visible = False
                    Me.DataGrid1.Visible = False
                    MsgBox("ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.bcpurchaseorder where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertPurchaseOrder '" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("ApCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("CreditDay") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DueDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("LeadTime") & ",'" & dsFrom.Tables("Data").Rows(i).Item("LeadDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExpireCredit") & ",'" & dsFrom.Tables("Data").Rows(i).Item("ExpireDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxRate") & "," & dsFrom.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsFrom.Tables("Data").Rows(i).Item("PoStatus") & "," & dsFrom.Tables("Data").Rows(i).Item("BillStatus") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCompleteSave") & "," & dsFrom.Tables("Data").Rows(i).Item("IsProcessOK") & ",'" & dsFrom.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsFrom.Tables("Data").Rows(i).Item("ContactCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsFrom.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsFrom.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsFrom.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("NetAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsFrom.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("RecurName") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsConditionSend") & ",'" & dsFrom.Tables("Data").Rows(i).Item("OrderToArCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsImport") & "," & dsFrom.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("Buyer") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If

                Me.DataGrid1.Items(0).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(0).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCPurchaseOrderSub"
                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.bcpurchaseordersub where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "DataSub")
                dtFrom = dsFrom.Tables("DataSub")
                If dsFrom.Tables("DataSub").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("DataSub").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertPurchaseOrderSub '" & dsFrom.Tables("datasub").Rows(i).Item("DocNo") & "'," & dsFrom.Tables("datasub").Rows(i).Item("TaxType") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ApCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ItemName") & "','" & dsFrom.Tables("datasub").Rows(i).Item("WHCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ShelfCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("Qty") & "," & dsFrom.Tables("datasub").Rows(i).Item("RemainQty") & "," & dsFrom.Tables("datasub").Rows(i).Item("Price") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("DiscountWord") & "'," & dsFrom.Tables("datasub").Rows(i).Item("DiscountAmount") & "," & dsFrom.Tables("datasub").Rows(i).Item("Amount") & "," & dsFrom.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsFrom.Tables("datasub").Rows(i).Item("HomeAmount") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("OppositeUnit") & "'," & dsFrom.Tables("datasub").Rows(i).Item("OppositeQty") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("StkReqNo") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ConfirmNo") & "'," & dsFrom.Tables("datasub").Rows(i).Item("ItemType") & "," & dsFrom.Tables("datasub").Rows(i).Item("ExceptTax") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("AllocateCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("CurrencyCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ExchangeRate") & "'," & dsFrom.Tables("datasub").Rows(i).Item("TransState") & "," & dsFrom.Tables("datasub").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("CategoryCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("GroupCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("BrandCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("TypeCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("FormatCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("BarCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("IsPromotion") & "," & dsFrom.Tables("datasub").Rows(i).Item("RefLineNumber") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("Colorcode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("SizeCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("StyleCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("BEHINDINDEX") & "," & dsFrom.Tables("datasub").Rows(i).Item("TAXRATE") & "," & dsFrom.Tables("datasub").Rows(i).Item("PACKINGRATE1") & "," & dsFrom.Tables("datasub").Rows(i).Item("PACKINGRATE2") & "," & dsFrom.Tables("datasub").Rows(i).Item("ISPROCESS") & "," & dsFrom.Tables("datasub").Rows(i).Item("ISLOCKCOST") & "," & dsFrom.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(1).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(1).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

            End If

            '=======================================================================================================================================================================================================
            '=======================================================================================================================================================================================================

            If vType = 2 Then

                vDocNo = UCase(Me.TextBox9.Text)

                '=======================================================================================================================================================================================================

                vTableName = "BCStktransfer"
                Call CheckData(0, vTableName, vDocNo)
                If vCheckExist = 0 Then
                    Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                    Call CreateDataSource()
                    Call BindGrid()
                    Me.Button1.Visible = False
                    Me.DataGrid1.Visible = False
                    MsgBox("ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCStktransfer where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertSTKTransfer '" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsFrom.Tables("Data").Rows(i).Item("SumOfQty") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCompleteSave") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("RecurName") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsFrom.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("ISPOS") & "," & dsFrom.Tables("Data").Rows(i).Item("SUMOFAMOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DepositNo") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(0).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(0).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCStktransfSub"
                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCStktransfSub where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertSTKTransferSub '" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("Data").Rows(i).Item("FromWH") & "','" & dsFrom.Tables("Data").Rows(i).Item("FromShelf") & "','" & dsFrom.Tables("Data").Rows(i).Item("ToWH") & "','" & dsFrom.Tables("Data").Rows(i).Item("ToShelf") & "'," & dsFrom.Tables("Data").Rows(i).Item("Qty") & "," & dsFrom.Tables("Data").Rows(i).Item("Price") & "," & dsFrom.Tables("Data").Rows(i).Item("SumOfCost") & "," & dsFrom.Tables("Data").Rows(i).Item("Amount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("RefNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsFrom.Tables("Data").Rows(i).Item("OppositeQty") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CategoryCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("GroupCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BrandCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("TypeCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("FormatCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BarCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("TransState") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsFrom.Tables("Data").Rows(i).Item("LotOutNumber") & "','" & dsFrom.Tables("Data").Rows(i).Item("LotInNumber") & "','" & dsFrom.Tables("Data").Rows(i).Item("LotExpireDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("LotMyDesc") & "','" & dsFrom.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsFrom.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("StyleCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("BEHINDINDEX") & "," & dsFrom.Tables("Data").Rows(i).Item("ISPOS") & "," & dsFrom.Tables("Data").Rows(i).Item("REFLINENUMBER") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(1).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(1).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCStktransfSub3"
                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCStktransfSub3 where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertSTKTransferSub3 " & dsFrom.Tables("Data").Rows(i).Item("BehindIndex") & "," & dsFrom.Tables("Data").Rows(i).Item("MyType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("Data").Rows(i).Item("WHCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ShelfCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("Qty") & "," & dsFrom.Tables("Data").Rows(i).Item("Price") & "," & dsFrom.Tables("Data").Rows(i).Item("SumOfCost") & "," & dsFrom.Tables("Data").Rows(i).Item("Amount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("RefNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsFrom.Tables("Data").Rows(i).Item("OppositeQty") & ",'" & dsFrom.Tables("Data").Rows(i).Item("fromwhcode") & "','" & dsFrom.Tables("Data").Rows(i).Item("fromshelfcode") & "'," & dsFrom.Tables("Data").Rows(i).Item("reftrlinenum") & "," & dsFrom.Tables("Data").Rows(i).Item("FixUnitCost") & "," & dsFrom.Tables("Data").Rows(i).Item("FixUnitQty") & "," & dsFrom.Tables("Data").Rows(i).Item("AVERAGECOST") & ",'" & dsFrom.Tables("Data").Rows(i).Item("LotNumber") & "','" & dsFrom.Tables("Data").Rows(i).Item("LotExpireDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("LotMyDesc") & "','" & dsFrom.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsFrom.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("StyleCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("TAXRATE") & "," & dsFrom.Tables("Data").Rows(i).Item("PACKINGRATE1") & "," & dsFrom.Tables("Data").Rows(i).Item("PACKINGRATE2") & "," & dsFrom.Tables("Data").Rows(i).Item("ISCANCEL") & "," & dsFrom.Tables("Data").Rows(i).Item("ISPOS") & "," & dsFrom.Tables("Data").Rows(i).Item("ISPROCESS") & "," & dsFrom.Tables("Data").Rows(i).Item("ISLOCKCOST") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(2).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(2).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

            End If

            '=======================================================================================================================================================================================================
            '=======================================================================================================================================================================================================

            If vType = 3 Then
                vItemCode = UCase(Me.TextBox10.Text)

                '=======================================================================================================================================================================================================

                vTableName = "BCItem"
                Call CheckData(1, vTableName, vItemCode)
                If vCheckExist = 0 Then
                    Me.TextBox12.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                    Call CreateDataSource()
                    Call BindGrid()
                    Me.Button1.Visible = False
                    Me.DataGrid1.Visible = False
                    MsgBox("ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
                Call PrepareData(1, vTableName, vItemCode)
                Call DeleteTableBranch(1, vTableName, vItemCode)


                Dim vSalePrice1 As Double
                Dim vSalePrice2 As Double
                Dim vSalePrice3 As Double
                Dim vSalePrice4 As Double
                Dim vSalePrice5 As Double
                Dim vSalePrice6 As Double
                Dim vSalePrice7 As Double
                Dim vSalePrice8 As Double
                Dim vSalePrice9 As Double
                Dim vISExport As Integer
                Dim vPrintYear As Integer

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.bcitem where code ='" & vItemCode & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "BCItem")
                dtFrom = dsFrom.Tables("BCItem")
                If dsFrom.Tables("BCItem").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("BCItem").Rows.Count - 1

                        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice1")) Then
                            vSalePrice1 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice1")
                        End If

                        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice2")) Then
                            vSalePrice2 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice2")
                        End If

                        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice3")) Then
                            vSalePrice3 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice3")
                        End If

                        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice4")) Then
                            vSalePrice4 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice4")
                        End If

                        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice5")) Then
                            vSalePrice5 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice5")
                        End If

                        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice6")) Then
                            vSalePrice6 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice6")
                        End If

                        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice7")) Then
                            vSalePrice7 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice7")
                        End If

                        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice8")) Then
                            vSalePrice8 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice8")
                        End If

                        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice9")) Then
                            vSalePrice9 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice9")
                        End If

                        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("isexport")) Then
                            vISExport = dsFrom.Tables("BCItem").Rows(i).Item("isexport")
                        End If

                        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("printyear")) Then
                            vPrintYear = dsFrom.Tables("BCItem").Rows(i).Item("printyear")
                        End If

                        vQuery = "exec dbo.USP_PTF_InsertItemMaster '" & dsFrom.Tables("BCItem").Rows(i).Item("Code") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Name1") & "', '" & dsFrom.Tables("BCItem").Rows(i).Item("Name2") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ShortName") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CategoryCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("GroupCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("BrandCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("TypeCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("FormatCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ColorCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("MyGrade") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("myclass") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("GenericCode") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("UnitType") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("DefStkUnitCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefSaleUnitCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefBuyUnitCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("OppositeUnit") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("OppositeUnit2") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("Opposite1Price") & "," & dsFrom.Tables("BCItem").Rows(i).Item("OppositeRate") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("MySize") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("Weight") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("LastUpdate") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("StockType") & "," & dsFrom.Tables("BCItem").Rows(i).Item("StockProcess") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("VenderCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DepositCondit") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DeliveryCharge") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("InstallCharge") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ReturnRemark") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("ReturnStatus") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("ReturnCharge") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ItemStatus") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("NewItemStatus") & "," & dsFrom.Tables("BCItem").Rows(i).Item("RenovateStatus") & "," & dsFrom.Tables("BCItem").Rows(i).Item("ExceptTax") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PromotionType") & "," & dsFrom.Tables("BCItem").Rows(i).Item("WTaxRate") & " ,'" & dsFrom.Tables("BCItem").Rows(i).Item("WaveFile") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("CostType") & "," & dsFrom.Tables("BCItem").Rows(i).Item("OrderPoint") & "," & dsFrom.Tables("BCItem").Rows(i).Item("StockMin") & "," & dsFrom.Tables("BCItem").Rows(i).Item("StockMax") & "," & dsFrom.Tables("BCItem").Rows(i).Item("StockQty") & "," & dsFrom.Tables("BCItem").Rows(i).Item("AverageCost") & "," & dsFrom.Tables("BCItem").Rows(i).Item("Amount") & "," & dsFrom.Tables("BCItem").Rows(i).Item("CostOfArea") & "," & dsFrom.Tables("BCItem").Rows(i).Item("LastBuyPrice") & "," & dsFrom.Tables("BCItem").Rows(i).Item("StandardCost") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("PicFileName1") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("PicFileName2") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("PicFileName3") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("PicFileName4") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("SpecFileName") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("AviFileName") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("AccGroupCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefBuyWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefBuyShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefSaleWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefSaleShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefBadWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefBadShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefFGWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefFGShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefRepairWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefRepairShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefRawWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefRawShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Formula1") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Formula2") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Formula3") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Discount") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("ReserveQty") & "," & dsFrom.Tables("BCItem").Rows(i).Item("RemainOutQty") & "," & dsFrom.Tables("BCItem").Rows(i).Item("RemainInQty") & "," & dsFrom.Tables("BCItem").Rows(i).Item("BasePrice") & "," & dsFrom.Tables("BCItem").Rows(i).Item("ActiveStatus") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CancelDateTime") & "'," & vSalePrice1 & "," & vSalePrice2 & "," & vSalePrice3 & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("DefFixUnitCode") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("DefFixUnitCost") & "," & dsFrom.Tables("BCItem").Rows(i).Item("FixQty") & "," & dsFrom.Tables("BCItem").Rows(i).Item("FixCost") & "," & dsFrom.Tables("BCItem").Rows(i).Item("SalePriceNV1") & "," & dsFrom.Tables("BCItem").Rows(i).Item("SalePriceNV2") & "," & dsFrom.Tables("BCItem").Rows(i).Item("SalePriceNV3") & "," & dsFrom.Tables("BCItem").Rows(i).Item("ProcessStatus") & "," & dsFrom.Tables("BCItem").Rows(i).Item("IsGift") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PromoMember") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PromoType") & "," & dsFrom.Tables("BCItem").Rows(i).Item("AverageCostReal") & "," & dsFrom.Tables("BCItem").Rows(i).Item("OverReceive") & "," & dsFrom.Tables("BCItem").Rows(i).Item("ContainerCapacity") & "," & dsFrom.Tables("BCItem").Rows(i).Item("ContainerWeight") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("CapacityUnit") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("LastBuyPriceHome") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("LastBuyCurrencyCode") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("LastAvgCost") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("UserGroup") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("stockqty1") & "," & dsFrom.Tables("BCItem").Rows(i).Item("stockqty2") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("StockQtyWord") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CustGroup") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("PublisherCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("WriterCode") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("PrintOrdinal") & "," & vPrintYear & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("LastSaleDate") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("LastBuyDate") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("LastSalePrice") & "," & dsFrom.Tables("BCItem").Rows(i).Item("LastSalePriceHome") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("LastSaleCurrencyCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Specification") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ManufactoryCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("AFT_remark") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("UnitWeight") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("WeightUnitCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("LeadTime") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("OfferedBy") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("OfferedDate") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("HS_SMX_remark") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Description") & "'," & vSalePrice4 & "," & vSalePrice5 & "," & vSalePrice6 & "," & vSalePrice7 & "," & vSalePrice8 & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT1") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT2") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT3") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT4") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT5") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT6") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT7") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT8") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("ITEMBARCODE") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("PRICECODE") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Merchandiser") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CLASSFICATION") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("CANMAKEPO") & "," & dsFrom.Tables("BCItem").Rows(i).Item("DISCCASHCARD") & "," & vISExport & "," & dsFrom.Tables("BCItem").Rows(i).Item("ISVOLUME") & ""
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(0).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(0).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vItemCode)

                vQuery = "update dbo.bcitem set defbuywhcode = 'S02',defbuyshelf = 'AVL',defsalewhcode = 'S02',defsaleshelf = 'AVL',defbadwhcode = 'S02',Defbadshelf='DMG',deffgwhcode = 'S02',deffgshelf = 'AVL',defrepairwhcode = 'S02',defrepairshelf = 'RTV',defrawwhcode = 'S02',defrawshelf = 'AVL' where code = '" & vItemCode & "'"
                With cmd
                    .CommandType = CommandType.Text
                    .CommandText = vQuery
                    .Connection = vConnectionTo
                    .ExecuteNonQuery()
                End With

                '=======================================================================================================================================================================================================

                vTableName = "BCBarcodeMaster"
                Call PrepareData(2, vTableName, vItemCode)
                Call DeleteTableBranch(2, vTableName, vItemCode)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.bcbarcodemaster where itemcode ='" & vItemCode & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        If Not IsDBNull(dsFrom.Tables("Data").Rows(i).Item("isexport")) Then
                            vISExport = dsFrom.Tables("Data").Rows(i).Item("isexport")
                        End If

                        vQuery = "exec dbo.USP_PTF_InsertBarCodeMaster '" & dsFrom.Tables("Data").Rows(i).Item("Barcode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BarcodeName") & "','" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("Remark") & "'," & dsFrom.Tables("Data").Rows(i).Item("ActiveStatus") & ",'" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("WHCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("SHELFCODE") & "'," & vISExport & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If

                vQuery = "update dbo.bcbarcodemaster set whcode = 'S02',shelfcode = 'AVL' where itemcode = '" & vItemCode & "'"
                With cmd
                    .CommandType = CommandType.Text
                    .CommandText = vQuery
                    .Connection = vConnectionTo
                    .ExecuteNonQuery()
                End With

                Me.DataGrid1.Items(1).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(1).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vItemCode)

                '=======================================================================================================================================================================================================

                vTableName = "BCPriceList"
                Call PrepareData(2, vTableName, vItemCode)
                Call DeleteTableBranch(2, vTableName, vItemCode)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.bcpricelist where itemcode ='" & vItemCode & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        If Not IsDBNull(dsFrom.Tables("Data").Rows(i).Item("isexport")) Then
                            vISExport = dsFrom.Tables("Data").Rows(i).Item("isexport")
                        End If
                        vQuery = "exec dbo.USP_PTF_InsertPriceList '" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("Remark") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxType") & "," & dsFrom.Tables("Data").Rows(i).Item("FromQty") & "," & dsFrom.Tables("Data").Rows(i).Item("ToQty") & ",'" & dsFrom.Tables("Data").Rows(i).Item("StartDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("StopDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("SaleType") & "," & dsFrom.Tables("Data").Rows(i).Item("TransportType") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice1") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice2") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice3") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice4") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice5") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice6") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice7") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice8") & "," & dsFrom.Tables("Data").Rows(i).Item("LowPrice") & "," & dsFrom.Tables("Data").Rows(i).Item("LineNumber") & "," & vISExport & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(2).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(2).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vItemCode)

                '=======================================================================================================================================================================================================

                vTableName = "BPSPriceList"
                Call PrepareData(2, vTableName, vItemCode)
                Call DeleteTableBranch(2, vTableName, vItemCode)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.bpspricelist where itemcode ='" & vItemCode & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertBPSPriceList '" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("Remark") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("BarCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("FromQty") & "," & dsFrom.Tables("Data").Rows(i).Item("ToQty") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice1") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice2") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice3") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice4") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice5") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice6") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice7") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice8") & "," & dsFrom.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(3).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(3).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vItemCode)

                '=======================================================================================================================================================================================================

                vTableName = "BCItemWareHouse"
                Call PrepareData(2, vTableName, vItemCode)
                Call DeleteTableBranch(2, vTableName, vItemCode)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.bcitemwarehouse where itemcode ='" & vItemCode & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertItemWareHouse '" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("WHCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ShelfCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(4).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(4).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vItemCode)

                '=======================================================================================================================================================================================================

                vTableName = "BCPriceErect"
                Call PrepareData(2, vTableName, vItemCode)
                Call DeleteTableBranch(2, vTableName, vItemCode)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.bcpriceerect where itemcode ='" & vItemCode & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertPriceErect '" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("PriceErect") & "','" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(5).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(5).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vItemCode)
            End If

            '=======================================================================================================================================================================================================
            '=======================================================================================================================================================================================================

            If vType = 4 Then

                vDocNo = UCase(Me.TextBox9.Text)

                vTableName = "BCARInvoice"
                Call CheckData(0, vTableName, vDocNo)
                If vCheckExist = 0 Then
                    Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                    Call CreateDataSource()
                    Call BindGrid()
                    Me.Button1.Visible = False
                    Me.DataGrid1.Visible = False
                    MsgBox("ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                Dim vTimeTransport As String
                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCARInvoice where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1

                        If Not IsDBNull(dsFrom.Tables("Data").Rows(i).Item("timetransport")) Then
                            vTimeTransport = dsFrom.Tables("Data").Rows(i).Item("timetransport")
                        End If

                        vQuery = "exec dbo.USP_PTF_InsertARInvoice '" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("TaxNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("ArCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("PassBillTo") & "','" & dsFrom.Tables("Data").Rows(i).Item("ArName") & "','" & dsFrom.Tables("Data").Rows(i).Item("ArAddress") & "','" & dsFrom.Tables("Data").Rows(i).Item("CashierCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ShiftNo") & ",'" & dsFrom.Tables("Data").Rows(i).Item("MachineNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("MachineCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("PosStatus") & ",'" & dsFrom.Tables("Data").Rows(i).Item("BillTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreditType") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreditBranch") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreditDueDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreditNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("CofirmNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("ChargeWord") & "'," & dsFrom.Tables("Data").Rows(i).Item("CreditBaseAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("ChargeAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("GrandTotal") & "," & dsFrom.Tables("Data").Rows(i).Item("CoupongAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CoupongDesc") & "'," & dsFrom.Tables("Data").Rows(i).Item("ChangeAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("SumMarkCount1") & "," & dsFrom.Tables("Data").Rows(i).Item("SumMarkCount2") & "," & dsFrom.Tables("Data").Rows(i).Item("SumMarkValue1") & "," & dsFrom.Tables("Data").Rows(i).Item("SumMarkValue2") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("CreditDay") & "," & dsFrom.Tables("Data").Rows(i).Item("DeliveryDay") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DeliveryDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("DueDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("PayBillDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("SaleAreaCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxRate") & "," & dsFrom.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsFrom.Tables("Data").Rows(i).Item("BillType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsFrom.Tables("Data").Rows(i).Item("RefDocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("CrAuthorizeMan") & "','" & dsFrom.Tables("Data").Rows(i).Item("DeliveryAddr") & "','" & dsFrom.Tables("Data").Rows(i).Item("ContactCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("TransportCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsFrom.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsFrom.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsFrom.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("SumCashAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("SumChqAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("SumCreditAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("SumBankAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("DepositIncTax") & "," & dsFrom.Tables("Data").Rows(i).Item("SumOfDeposit1") & "," & dsFrom.Tables("Data").Rows(i).Item("SumOfDeposit2") & "," & dsFrom.Tables("Data").Rows(i).Item("SumOfWTax") & "," & dsFrom.Tables("Data").Rows(i).Item("NetDebtAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("HomeAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("OtherIncome") & "," & dsFrom.Tables("Data").Rows(i).Item("OtherExpense") & "," & dsFrom.Tables("Data").Rows(i).Item("ExcessAmount1") & "," & dsFrom.Tables("Data").Rows(i).Item("ExcessAmount2") & "," & dsFrom.Tables("Data").Rows(i).Item("BillBalance") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsFrom.Tables("Data").Rows(i).Item("GLFormat") & "'," & dsFrom.Tables("Data").Rows(i).Item("GLStartPosting") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCreditNote") & "," & dsFrom.Tables("Data").Rows(i).Item("IsDebitNote") & "," & dsFrom.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCompleteSave") & "," & dsFrom.Tables("Data").Rows(i).Item("IsPostGL") & "," & dsFrom.Tables("Data").Rows(i).Item("GLTransData") & "," & dsFrom.Tables("Data").Rows(i).Item("PayBillStatus") & ",'" & dsFrom.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("RecurName") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExchangeProfit") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CustTypeCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CustGroupCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("Remark1") & "','" & dsFrom.Tables("Data").Rows(i).Item("Remark2") & "','" & dsFrom.Tables("Data").Rows(i).Item("Remark3") & "','" & dsFrom.Tables("Data").Rows(i).Item("Remark4") & "','" & dsFrom.Tables("Data").Rows(i).Item("Remark5") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsReceiveMoney") & "," & dsFrom.Tables("Data").Rows(i).Item("IsConditionSend") & "," & dsFrom.Tables("Data").Rows(i).Item("TotalTransport") & "," & dsFrom.Tables("Data").Rows(i).Item("PayBillAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("GrossWeight") & "," & dsFrom.Tables("Data").Rows(i).Item("PrintCount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("SORefNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("HoldingStatus") & ",'" & dsFrom.Tables("Data").Rows(i).Item("TimeTransport") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsImport") & ",'" & dsFrom.Tables("Data").Rows(i).Item("JOBNO") & "'," & dsFrom.Tables("Data").Rows(i).Item("REFTYPE") & "," & dsFrom.Tables("Data").Rows(i).Item("BILLTEMPORARY") & "," & dsFrom.Tables("Data").Rows(i).Item("ISMULTITAXABB") & ",'" & dsFrom.Tables("Data").Rows(i).Item("TAXABBNO_1") & "','" & dsFrom.Tables("Data").Rows(i).Item("TAXABBNO_2") & "','" & dsFrom.Tables("Data").Rows(i).Item("TAXABBNO_3") & "','" & dsFrom.Tables("Data").Rows(i).Item("TAXABBNO_4") & "','" & dsFrom.Tables("Data").Rows(i).Item("TAXABBNO_5") & "'," & dsFrom.Tables("Data").Rows(i).Item("TAXABBAMOUNT1") & "," & dsFrom.Tables("Data").Rows(i).Item("TAXABBAMOUNT2") & "," & dsFrom.Tables("Data").Rows(i).Item("TAXABBAMOUNT3") & "," & dsFrom.Tables("Data").Rows(i).Item("TAXABBAMOUNT4") & "," & dsFrom.Tables("Data").Rows(i).Item("TAXABBAMOUNT5") & ",'" & dsFrom.Tables("Data").Rows(i).Item("APPROVEDCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("APPROVEDDATE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CANCELDESC") & "','" & dsFrom.Tables("Data").Rows(i).Item("SHIFTCODE") & "'," & dsFrom.Tables("Data").Rows(i).Item("CREDITVAT") & "," & dsFrom.Tables("Data").Rows(i).Item("CREDITSUMVAT") & "," & dsFrom.Tables("Data").Rows(i).Item("OTHERAMOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("OTHERFEE") & "," & dsFrom.Tables("Data").Rows(i).Item("DIFFAMOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("ISREWARD") & "," & dsFrom.Tables("Data").Rows(i).Item("POSCREDIT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("USERGROUP") & "'," & dsFrom.Tables("Data").Rows(i).Item("NETWEIGHT") & "," & dsFrom.Tables("Data").Rows(i).Item("NUMOFPALLET") & ",'" & dsFrom.Tables("Data").Rows(i).Item("INVOICETYPE") & "'," & dsFrom.Tables("Data").Rows(i).Item("QTYAMOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("QTYDEFAULT") & "," & dsFrom.Tables("Data").Rows(i).Item("QTYCOPY") & "," & dsFrom.Tables("Data").Rows(i).Item("MERGEITEM") & "," & dsFrom.Tables("Data").Rows(i).Item("NEWLINE") & "," & dsFrom.Tables("Data").Rows(i).Item("CALCTAXFLAG") & "," & dsFrom.Tables("Data").Rows(i).Item("PRICECOPY") & "," & dsFrom.Tables("Data").Rows(i).Item("WHCOPY") & ",'" & dsFrom.Tables("Data").Rows(i).Item("METHODEPAYBILL") & "','" & dsFrom.Tables("Data").Rows(i).Item("METHODEPAYBILL2") & "','" & dsFrom.Tables("Data").Rows(i).Item("DOREFNO") & "','" & dsFrom.Tables("Data").Rows(i).Item("PROMOTIONCODE") & "'," & dsFrom.Tables("Data").Rows(i).Item("SUMOFWTAXCASH") & "," & dsFrom.Tables("Data").Rows(i).Item("SUMBASEWTAXCASH") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ""

                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(0).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(0).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCARInvoiceSub"

                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                Dim vDORemainQTY As Double

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCARInvoiceSub where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1

                        If Not IsDBNull(dsFrom.Tables("Data").Rows(i).Item("DORemainQTY")) Then
                            vDORemainQTY = dsFrom.Tables("Data").Rows(i).Item("DORemainQTY")
                        End If

                        vQuery = "exec dbo.USP_PTF_InsertARInvoiceSub " & dsFrom.Tables("Data").Rows(i).Item("BehindIndex") & "," & dsFrom.Tables("Data").Rows(i).Item("MyType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("TaxNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("ArCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("Data").Rows(i).Item("ItemName") & "','" & dsFrom.Tables("Data").Rows(i).Item("WHCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ShelfCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("CNQty") & "," & dsFrom.Tables("Data").Rows(i).Item("Qty") & "," & dsFrom.Tables("Data").Rows(i).Item("Price") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsFrom.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("Amount") & "," & dsFrom.Tables("Data").Rows(i).Item("NetAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("HomeAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("SumOfCost") & "," & dsFrom.Tables("Data").Rows(i).Item("BalanceAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsFrom.Tables("Data").Rows(i).Item("OppositeQty") & "," & dsFrom.Tables("Data").Rows(i).Item("OppositePrice2") & ",'" & dsFrom.Tables("Data").Rows(i).Item("SORefNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("PORefNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("StockType") & "," & dsFrom.Tables("Data").Rows(i).Item("ExceptTax") & "," & dsFrom.Tables("Data").Rows(i).Item("LineNumber") & "," & dsFrom.Tables("Data").Rows(i).Item("RefLineNumber") & "," & dsFrom.Tables("Data").Rows(i).Item("TransState") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & ",'" & dsFrom.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsFrom.Tables("Data").Rows(i).Item("BarCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CustTypeCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CustGroupCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("SaleAreaCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CategoryCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("GroupCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BrandCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("TypeCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("FormatCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("MachineNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("MachineCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BillTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CashierCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ShiftNo") & "," & dsFrom.Tables("Data").Rows(i).Item("PosStatus") & "," & dsFrom.Tables("Data").Rows(i).Item("PriceStatus") & "," & dsFrom.Tables("Data").Rows(i).Item("IsPromotion") & "," & dsFrom.Tables("Data").Rows(i).Item("PremiumStatus") & "," & dsFrom.Tables("Data").Rows(i).Item("FixUnitCost") & "," & dsFrom.Tables("Data").Rows(i).Item("FixUnitQty") & "," & dsFrom.Tables("Data").Rows(i).Item("IsConditionSend") & "," & dsFrom.Tables("Data").Rows(i).Item("TransportAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("AVERAGECOST") & ",'" & dsFrom.Tables("Data").Rows(i).Item("LotNumber") & "','" & dsFrom.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsFrom.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("StyleCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("itemsetcode") & "','" & dsFrom.Tables("Data").Rows(i).Item("JobNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("TAXRATE") & "," & dsFrom.Tables("Data").Rows(i).Item("PACKINGRATE1") & "," & dsFrom.Tables("Data").Rows(i).Item("PACKINGRATE2") & "," & dsFrom.Tables("Data").Rows(i).Item("REFTYPE") & ",'" & dsFrom.Tables("Data").Rows(i).Item("SHIFTCODE") & "'," & dsFrom.Tables("Data").Rows(i).Item("PROMOSTATUS") & "," & dsFrom.Tables("Data").Rows(i).Item("OLDPRICE") & ",'" & dsFrom.Tables("Data").Rows(i).Item("USERCODE") & "'," & dsFrom.Tables("Data").Rows(i).Item("USERMODIFY") & "," & dsFrom.Tables("Data").Rows(i).Item("POSCREDIT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("USERGROUP") & "','" & dsFrom.Tables("Data").Rows(i).Item("PRICECODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("INVOICETYPE") & "'," & dsFrom.Tables("Data").Rows(i).Item("ISPROCESS") & "," & dsFrom.Tables("Data").Rows(i).Item("ISLOCKCOST") & ",'" & dsFrom.Tables("Data").Rows(i).Item("ITEMNAMEDESC") & "','" & dsFrom.Tables("Data").Rows(i).Item("DOREFNO") & "'," & dsFrom.Tables("Data").Rows(i).Item("DELIVERYSTATUS") & ",'" & dsFrom.Tables("Data").Rows(i).Item("PROMOTIONCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("MASTERITEMCODE") & "'," & dsFrom.Tables("Data").Rows(i).Item("BTDISC") & "," & dsFrom.Tables("Data").Rows(i).Item("DISCCASHCARD") & "," & dsFrom.Tables("Data").Rows(i).Item("WTAXAMOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("BASEWTAX") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'," & vDORemainQTY & ""
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(1).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(1).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCRecMoney"

                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCRecMoney where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertRecMoney '" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("ArCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsFrom.Tables("Data").Rows(i).Item("PayAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("HomeAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("ChqTotalAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("PaymentType") & "," & dsFrom.Tables("Data").Rows(i).Item("SaveFrom") & "," & dsFrom.Tables("Data").Rows(i).Item("PayChqState") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CreditType") & "'," & dsFrom.Tables("Data").Rows(i).Item("ChargeAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("ChargeWord") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsFrom.Tables("Data").Rows(i).Item("RefNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("BankCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BankBranchCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("TransBankDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("Data").Rows(i).Item("RefDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("CHANGEAMOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(2).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(2).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCOutPutTax"

                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                Dim vExchangeRate As Double

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCOutPutTax where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1

                        If Not IsDBNull(dsFrom.Tables("Data").Rows(i).Item("ExchangeRate")) Then
                            vExchangeRate = dsFrom.Tables("Data").Rows(i).Item("ExchangeRate")
                        End If

                        vQuery = "exec dbo.USP_PTF_InsertOutPutTax " & dsFrom.Tables("Data").Rows(i).Item("SaveFrom") & " ,'" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("BookCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("Source") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("TaxDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("TaxNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("ArCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ShortTaxDesc") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxRate") & "," & dsFrom.Tables("Data").Rows(i).Item("Process") & "," & dsFrom.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("LineNumber") & "," & dsFrom.Tables("Data").Rows(i).Item("IsMultiCurrency") & "," & dsFrom.Tables("Data").Rows(i).Item("FAmount") & "," & vExchangeRate & ",'" & dsFrom.Tables("Data").Rows(i).Item("TaxGroup") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("CancelOutPeriod") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CancelDocNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsPos") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CANCELDOCDATE") & "'," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(3).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(3).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)


                '=======================================================================================================================================================================================================

                vTableName = "BCCreditCard"

                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCCreditCard where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertCreditCard '" & dsFrom.Tables("Data").Rows(i).Item("BankCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreditCardNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("ArCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ReceiveDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("DueDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("BookNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("Status") & "," & dsFrom.Tables("Data").Rows(i).Item("SaveFrom") & ",'" & dsFrom.Tables("Data").Rows(i).Item("StatusDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("StatusDocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BankBranchCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("Amount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CreditType") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("ChargeAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsFrom.Tables("Data").Rows(i).Item("CreditVatRate") & "," & dsFrom.Tables("Data").Rows(i).Item("CreditVat") & "," & dsFrom.Tables("Data").Rows(i).Item("CreditSumVat") & "," & dsFrom.Tables("Data").Rows(i).Item("PRINTCOUNT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("POSDOCNO") & "'," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(4).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(4).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCCHQIN"

                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                Dim vReciveConfirm As Integer
                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCCHQIN where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1

                        If Not IsDBNull(dsFrom.Tables("Data").Rows(i).Item("RECIVECONFIRM")) Then
                            vReciveConfirm = dsFrom.Tables("Data").Rows(i).Item("RECIVECONFIRM")
                        End If

                        vQuery = "exec dbo.USP_PTF_InsertCHQIN '" & dsFrom.Tables("Data").Rows(i).Item("BankCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ChqNumber") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("ArCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ExtendStatus") & "','" & dsFrom.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ReceiveDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("DueDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("BookNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("Status") & "," & dsFrom.Tables("Data").Rows(i).Item("SaveFrom") & ",'" & dsFrom.Tables("Data").Rows(i).Item("StatusDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("StatusDocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BankBranchCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("Amount") & "," & dsFrom.Tables("Data").Rows(i).Item("Balance") & ",'" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsFrom.Tables("Data").Rows(i).Item("ChqUseStatus") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("RECIVECHQBY") & "'," & vReciveConfirm & "," & dsFrom.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(5).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(5).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)


                '=======================================================================================================================================================================================================

                vTableName = "BCTrans"

                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCTrans where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertBCTrans '" & dsFrom.Tables("Data").Rows(i).Item("BatchNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("BookCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("RefDocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("RefDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("Amount") & "," & dsFrom.Tables("Data").Rows(i).Item("FAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("FCurrency") & "','" & dsFrom.Tables("Data").Rows(i).Item("FExchangeRate") & "'," & dsFrom.Tables("Data").Rows(i).Item("Source") & "," & dsFrom.Tables("Data").Rows(i).Item("TransType") & "," & dsFrom.Tables("Data").Rows(i).Item("IsManualKey") & ",'" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("Data").Rows(i).Item("RecurName") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsConfirm") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("IsPassError") & "," & dsFrom.Tables("Data").Rows(i).Item("TaxCount") & "," & dsFrom.Tables("Data").Rows(i).Item("CheqCount") & "," & dsFrom.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(6).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(6).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCTransSub"

                Call PrepareData(0, vTableName, vDocNo)
                Call DeleteTableBranch(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCTransSub where docno ='" & vDocNo & "'"
                daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                dsFrom = New DataSet
                daFrom.Fill(dsFrom, "Data")
                dtFrom = dsFrom.Tables("Data")
                If dsFrom.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertBCTransSub '" & dsFrom.Tables("Data").Rows(i).Item("BatchNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsFrom.Tables("Data").Rows(i).Item("BookCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("AccountCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ExtDescription") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("PartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("SideCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("JobCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BranchCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("Debit") & "," & dsFrom.Tables("Data").Rows(i).Item("Credit") & "," & dsFrom.Tables("Data").Rows(i).Item("FDebit") & "," & dsFrom.Tables("Data").Rows(i).Item("FCredit") & "," & dsFrom.Tables("Data").Rows(i).Item("Source") & "," & dsFrom.Tables("Data").Rows(i).Item("IsManualKey") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("IsConfirm") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionTo
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(7).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(7).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTable(vTableName)
                Call InsertLogs(0, vTableName, vDocNo)

            End If

        End If


        '=======================================================================================================================================================================================================
        '=======================================================================================================================================================================================================
        '=======================================================================================================================================================================================================

        'โอนจากสาขามาสำนักงานใหญ่
        If vTransType = 1 Then
            If vType = 0 Then
                MsgBox("ยังไม่เปิด การใช้งานในส่วนการโอนข้อมูลประเภทนี้", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
                'vDocNo = UCase(Me.TextBox9.Text)

                ''=======================================================================================================================================================================================================

                'vTableName = "BCSaleOrder"
                'Call CheckDataBranch(0, vTableName, vDocNo)
                'If vCheckExist = 0 Then
                '    Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                '    Call CreateDataSource()
                '    Call BindGrid()
                '    Me.Button1.Visible = False
                '    Me.DataGrid1.Visible = False
                '    MsgBox("ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                '    Exit Sub
                'End If
                'Call PrepareDatabranch(0, vTableName, vDocNo)
                'Call DeleteTable(0, vTableName, vDocNo)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.bcsaleorder where docno ='" & vDocNo & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "Data")
                'dtFrom = dsFrom.Tables("Data")
                'If dsFrom.Tables("Data").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                '        vQuery = "exec dbo.USP_PTF_InsertSaleOrder '" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxType") & "," & dsFrom.Tables("Data").Rows(i).Item("BillType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("ArCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("CreditDay") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DueDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DeliveryAddr") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxRate") & "," & dsFrom.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("Data").Rows(i).Item("RefDocNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("BillStatus") & "," & dsFrom.Tables("Data").Rows(i).Item("SOStatus") & "," & dsFrom.Tables("Data").Rows(i).Item("HoldingStatus") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CrAuthorizeMan") & "','" & dsFrom.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsFrom.Tables("Data").Rows(i).Item("ContactCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsFrom.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsFrom.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsFrom.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("NetAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCompleteSave") & ",'" & dsFrom.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("RecurName") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsConditionSend") & ",'" & dsFrom.Tables("Data").Rows(i).Item("OwnReceive") & "','" & dsFrom.Tables("Data").Rows(i).Item("CarLicense") & "','" & dsFrom.Tables("Data").Rows(i).Item("ApproveCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ApproveDateTime") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsUseRobotSale") & "," & dsFrom.Tables("Data").Rows(i).Item("TotalTransport") & "," & dsFrom.Tables("Data").Rows(i).Item("SumOtherValue") & "," & dsFrom.Tables("Data").Rows(i).Item("READYFORPAY") & ",'" & dsFrom.Tables("Data").Rows(i).Item("TimeTransport") & "','" & dsFrom.Tables("Data").Rows(i).Item("CarType") & "','" & dsFrom.Tables("Data").Rows(i).Item("CondPayCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("PrintCount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("PORefNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsImport") & ",'" & dsFrom.Tables("Data").Rows(i).Item("USERGROUP") & "','" & dsFrom.Tables("Data").Rows(i).Item("METHODEPAYBILL") & "','" & dsFrom.Tables("Data").Rows(i).Item("METHODEPAYBILL2") & "'," & dsFrom.Tables("Data").Rows(i).Item("DELIVERYDAY") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DELIVERYDATE") & "'," & dsFrom.Tables("Data").Rows(i).Item("DELIVERYSTATUS") & ",'" & dsFrom.Tables("Data").Rows(i).Item("PROMOTIONCODE") & "'," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If

                'Me.DataGrid1.Items(0).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(0).Cells(2).Text = "โอนเรียบร้อย"

                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vDocNo)

                ''=======================================================================================================================================================================================================

                'vTableName = "BCSaleOrderSub"
                'Call PrepareDatabranch(0, vTableName, vDocNo)
                'Call DeleteTable(0, vTableName, vDocNo)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.bcsaleordersub where docno ='" & vDocNo & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "DataSub")
                'dtFrom = dsFrom.Tables("DataSub")
                'If dsFrom.Tables("DataSub").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("DataSub").Rows.Count - 1
                '        vQuery = "exec dbo.USP_PTF_InsertSaleOrderSub '" & dsFrom.Tables("datasub").Rows(i).Item("DocNo") & "'," & dsFrom.Tables("datasub").Rows(i).Item("TaxType") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ArCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("SaleCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ItemName") & "','" & dsFrom.Tables("datasub").Rows(i).Item("WHCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ShelfCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("Qty") & "," & dsFrom.Tables("datasub").Rows(i).Item("RemainQty") & "," & dsFrom.Tables("datasub").Rows(i).Item("Price") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("DiscountWord") & "'," & dsFrom.Tables("datasub").Rows(i).Item("DiscountAmount") & "," & dsFrom.Tables("datasub").Rows(i).Item("Amount") & "," & dsFrom.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsFrom.Tables("datasub").Rows(i).Item("HomeAmount") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("UnitCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("OppositePrice2") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("OppositeUnit") & "'," & dsFrom.Tables("datasub").Rows(i).Item("OppositeQty") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("StkReserveNo") & "'," & dsFrom.Tables("datasub").Rows(i).Item("SOStatus") & "," & dsFrom.Tables("datasub").Rows(i).Item("HoldingStatus") & "," & dsFrom.Tables("datasub").Rows(i).Item("RefType") & "," & dsFrom.Tables("datasub").Rows(i).Item("ItemType") & "," & dsFrom.Tables("datasub").Rows(i).Item("ExceptTax") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("AllocateCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("CurrencyCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("ExchangeRate") & "," & dsFrom.Tables("datasub").Rows(i).Item("TransState") & "," & dsFrom.Tables("datasub").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("CategoryCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("GroupCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("BrandCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("TypeCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("FormatCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("BarCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("IsPromotion") & "," & dsFrom.Tables("datasub").Rows(i).Item("PremiumStatus") & "," & dsFrom.Tables("datasub").Rows(i).Item("IsUseRobotSale") & "," & dsFrom.Tables("datasub").Rows(i).Item("IsConditionSend") & "," & dsFrom.Tables("datasub").Rows(i).Item("TransportAmount") & "," & dsFrom.Tables("datasub").Rows(i).Item("OtherValue") & "," & dsFrom.Tables("datasub").Rows(i).Item("REMAINPAYQTY") & "," & dsFrom.Tables("datasub").Rows(i).Item("ISPAYMENT") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("Colorcode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("SizeCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("StyleCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("itemsetcode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("PORefNo") & "'," & dsFrom.Tables("datasub").Rows(i).Item("BEHINDINDEX") & "," & dsFrom.Tables("datasub").Rows(i).Item("TAXRATE") & "," & dsFrom.Tables("datasub").Rows(i).Item("PACKINGRATE1") & "," & dsFrom.Tables("datasub").Rows(i).Item("PACKINGRATE2") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("USERGROUP") & "'," & dsFrom.Tables("datasub").Rows(i).Item("ISPROCESS") & "," & dsFrom.Tables("datasub").Rows(i).Item("ISLOCKCOST") & "," & dsFrom.Tables("datasub").Rows(i).Item("DELIVERYQTY") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("PROMOTIONCODE") & "','" & dsFrom.Tables("datasub").Rows(i).Item("MASTERITEMCODE") & "'," & dsFrom.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If

                'Me.DataGrid1.Items(1).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(1).Cells(2).Text = "โอนเรียบร้อย"

                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vDocNo)

            End If

            '=======================================================================================================================================================================================================
            '=======================================================================================================================================================================================================


            If vType = 1 Then
                MsgBox("ยังไม่เปิด การใช้งานในส่วนการโอนข้อมูลประเภทนี้", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
                'vDocNo = UCase(Me.TextBox9.Text)

                ''=======================================================================================================================================================================================================

                'vTableName = "BCPurchaseOrder"
                'Call CheckDataBranch(0, vTableName, vDocNo)
                'If vCheckExist = 0 Then
                '    Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                '    Call CreateDataSource()
                '    Call BindGrid()
                '    Me.Button1.Visible = False
                '    Me.DataGrid1.Visible = False
                '    MsgBox("ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                '    Exit Sub
                'End If
                'Call PrepareDatabranch(0, vTableName, vDocNo)
                'Call DeleteTable(0, vTableName, vDocNo)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.bcpurchaseorder where docno ='" & vDocNo & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "Data")
                'dtFrom = dsFrom.Tables("Data")
                'If dsFrom.Tables("Data").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                '        vQuery = "exec dbo.USP_PTF_InsertPurchaseOrder '" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("ApCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("CreditDay") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DueDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("LeadTime") & ",'" & dsFrom.Tables("Data").Rows(i).Item("LeadDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExpireCredit") & ",'" & dsFrom.Tables("Data").Rows(i).Item("ExpireDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxRate") & "," & dsFrom.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsFrom.Tables("Data").Rows(i).Item("PoStatus") & "," & dsFrom.Tables("Data").Rows(i).Item("BillStatus") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCompleteSave") & "," & dsFrom.Tables("Data").Rows(i).Item("IsProcessOK") & ",'" & dsFrom.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsFrom.Tables("Data").Rows(i).Item("ContactCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsFrom.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsFrom.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsFrom.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsFrom.Tables("Data").Rows(i).Item("NetAmount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsFrom.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("RecurName") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsConditionSend") & ",'" & dsFrom.Tables("Data").Rows(i).Item("OrderToArCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsImport") & "," & dsFrom.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("Buyer") & "'"
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If

                'Me.DataGrid1.Items(0).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(0).Cells(2).Text = "โอนเรียบร้อย"
                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vDocNo)

                ''=======================================================================================================================================================================================================

                'vTableName = "BCPurchaseOrderSub"
                'Call PrepareDatabranch(0, vTableName, vDocNo)
                'Call DeleteTable(0, vTableName, vDocNo)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.bcpurchaseordersub where docno ='" & vDocNo & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "DataSub")
                'dtFrom = dsFrom.Tables("DataSub")
                'If dsFrom.Tables("DataSub").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("DataSub").Rows.Count - 1
                '        vQuery = "exec dbo.USP_PTF_InsertPurchaseOrderSub '" & dsFrom.Tables("datasub").Rows(i).Item("DocNo") & "'," & dsFrom.Tables("datasub").Rows(i).Item("TaxType") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ApCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ItemName") & "','" & dsFrom.Tables("datasub").Rows(i).Item("WHCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ShelfCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("Qty") & "," & dsFrom.Tables("datasub").Rows(i).Item("RemainQty") & "," & dsFrom.Tables("datasub").Rows(i).Item("Price") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("DiscountWord") & "'," & dsFrom.Tables("datasub").Rows(i).Item("DiscountAmount") & "," & dsFrom.Tables("datasub").Rows(i).Item("Amount") & "," & dsFrom.Tables("datasub").Rows(i).Item("NetAmount") & "," & dsFrom.Tables("datasub").Rows(i).Item("HomeAmount") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("OppositeUnit") & "'," & dsFrom.Tables("datasub").Rows(i).Item("OppositeQty") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("StkReqNo") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ConfirmNo") & "'," & dsFrom.Tables("datasub").Rows(i).Item("ItemType") & "," & dsFrom.Tables("datasub").Rows(i).Item("ExceptTax") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("AllocateCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ProjectCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("CurrencyCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("ExchangeRate") & "'," & dsFrom.Tables("datasub").Rows(i).Item("TransState") & "," & dsFrom.Tables("datasub").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("datasub").Rows(i).Item("LineNumber") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("CategoryCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("GroupCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("BrandCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("TypeCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("FormatCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("BarCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("IsPromotion") & "," & dsFrom.Tables("datasub").Rows(i).Item("RefLineNumber") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("Colorcode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("SizeCode") & "','" & dsFrom.Tables("datasub").Rows(i).Item("StyleCode") & "'," & dsFrom.Tables("datasub").Rows(i).Item("BEHINDINDEX") & "," & dsFrom.Tables("datasub").Rows(i).Item("TAXRATE") & "," & dsFrom.Tables("datasub").Rows(i).Item("PACKINGRATE1") & "," & dsFrom.Tables("datasub").Rows(i).Item("PACKINGRATE2") & "," & dsFrom.Tables("datasub").Rows(i).Item("ISPROCESS") & "," & dsFrom.Tables("datasub").Rows(i).Item("ISLOCKCOST") & "," & dsFrom.Tables("datasub").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("datasub").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("datasub").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("datasub").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("datasub").Rows(i).Item("LASTEDITDATET") & "'"
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If
                'Me.DataGrid1.Items(1).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(1).Cells(2).Text = "โอนเรียบร้อย"
                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vDocNo)

            End If

            '=======================================================================================================================================================================================================
            '=======================================================================================================================================================================================================

            If vType = 2 Then
                MsgBox("ยังไม่เปิด การใช้งานในส่วนการโอนข้อมูลประเภทนี้", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
                'vDocNo = UCase(Me.TextBox9.Text)

                ''=======================================================================================================================================================================================================

                'vTableName = "BCStktransfer"
                'Call CheckDataBranch(0, vTableName, vDocNo)
                'If vCheckExist = 0 Then
                '    Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                '    Call CreateDataSource()
                '    Call BindGrid()
                '    Me.Button1.Visible = False
                '    Me.DataGrid1.Visible = False
                '    MsgBox("ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                '    Exit Sub
                'End If
                'Call PrepareDatabranch(0, vTableName, vDocNo)
                'Call DeleteTable(0, vTableName, vDocNo)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.BCStktransfer where docno ='" & vDocNo & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "Data")
                'dtFrom = dsFrom.Tables("Data")
                'If dsFrom.Tables("Data").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                '        vQuery = "exec dbo.USP_PTF_InsertSTKTransfer '" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "'," & dsFrom.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsFrom.Tables("Data").Rows(i).Item("SumOfQty") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCompleteSave") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("RecurName") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsFrom.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("ISPOS") & "," & dsFrom.Tables("Data").Rows(i).Item("SUMOFAMOUNT") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DepositNo") & "'"
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If
                'Me.DataGrid1.Items(0).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(0).Cells(2).Text = "โอนเรียบร้อย"
                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vDocNo)

                ''=======================================================================================================================================================================================================

                'vTableName = "BCStktransfSub"
                'Call PrepareDatabranch(0, vTableName, vDocNo)
                'Call DeleteTable(0, vTableName, vDocNo)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.BCStktransfSub where docno ='" & vDocNo & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "Data")
                'dtFrom = dsFrom.Tables("Data")
                'If dsFrom.Tables("Data").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                '        vQuery = "exec dbo.USP_PTF_InsertSTKTransferSub '" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("Data").Rows(i).Item("FromWH") & "','" & dsFrom.Tables("Data").Rows(i).Item("FromShelf") & "','" & dsFrom.Tables("Data").Rows(i).Item("ToWH") & "','" & dsFrom.Tables("Data").Rows(i).Item("ToShelf") & "'," & dsFrom.Tables("Data").Rows(i).Item("Qty") & "," & dsFrom.Tables("Data").Rows(i).Item("Price") & "," & dsFrom.Tables("Data").Rows(i).Item("SumOfCost") & "," & dsFrom.Tables("Data").Rows(i).Item("Amount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("RefNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsFrom.Tables("Data").Rows(i).Item("OppositeQty") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CategoryCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("GroupCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BrandCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("TypeCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("FormatCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BarCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("TransState") & "," & dsFrom.Tables("Data").Rows(i).Item("IsCancel") & "," & dsFrom.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsFrom.Tables("Data").Rows(i).Item("LotOutNumber") & "','" & dsFrom.Tables("Data").Rows(i).Item("LotInNumber") & "','" & dsFrom.Tables("Data").Rows(i).Item("LotExpireDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("LotMyDesc") & "','" & dsFrom.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsFrom.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("StyleCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("BEHINDINDEX") & "," & dsFrom.Tables("Data").Rows(i).Item("ISPOS") & "," & dsFrom.Tables("Data").Rows(i).Item("REFLINENUMBER") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If
                'Me.DataGrid1.Items(1).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(1).Cells(2).Text = "โอนเรียบร้อย"
                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vDocNo)

                ''=======================================================================================================================================================================================================

                'vTableName = "BCStktransfSub3"
                'Call PrepareDatabranch(0, vTableName, vDocNo)
                'Call DeleteTable(0, vTableName, vDocNo)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.BCStktransfSub3 where docno ='" & vDocNo & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "Data")
                'dtFrom = dsFrom.Tables("Data")
                'If dsFrom.Tables("Data").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                '        vQuery = "exec dbo.USP_PTF_InsertSTKTransferSub3 " & dsFrom.Tables("Data").Rows(i).Item("BehindIndex") & "," & dsFrom.Tables("Data").Rows(i).Item("MyType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("DocNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("DocDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("Data").Rows(i).Item("WHCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ShelfCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("Qty") & "," & dsFrom.Tables("Data").Rows(i).Item("Price") & "," & dsFrom.Tables("Data").Rows(i).Item("SumOfCost") & "," & dsFrom.Tables("Data").Rows(i).Item("Amount") & ",'" & dsFrom.Tables("Data").Rows(i).Item("RefNo") & "','" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsFrom.Tables("Data").Rows(i).Item("OppositeQty") & ",'" & dsFrom.Tables("Data").Rows(i).Item("fromwhcode") & "','" & dsFrom.Tables("Data").Rows(i).Item("fromshelfcode") & "'," & dsFrom.Tables("Data").Rows(i).Item("reftrlinenum") & "," & dsFrom.Tables("Data").Rows(i).Item("FixUnitCost") & "," & dsFrom.Tables("Data").Rows(i).Item("FixUnitQty") & "," & dsFrom.Tables("Data").Rows(i).Item("AVERAGECOST") & ",'" & dsFrom.Tables("Data").Rows(i).Item("LotNumber") & "','" & dsFrom.Tables("Data").Rows(i).Item("LotExpireDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("LotMyDesc") & "','" & dsFrom.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsFrom.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("StyleCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("TAXRATE") & "," & dsFrom.Tables("Data").Rows(i).Item("PACKINGRATE1") & "," & dsFrom.Tables("Data").Rows(i).Item("PACKINGRATE2") & "," & dsFrom.Tables("Data").Rows(i).Item("ISCANCEL") & "," & dsFrom.Tables("Data").Rows(i).Item("ISPOS") & "," & dsFrom.Tables("Data").Rows(i).Item("ISPROCESS") & "," & dsFrom.Tables("Data").Rows(i).Item("ISLOCKCOST") & "," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If
                'Me.DataGrid1.Items(2).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(2).Cells(2).Text = "โอนเรียบร้อย"
                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vDocNo)

            End If

            '=======================================================================================================================================================================================================
            '=======================================================================================================================================================================================================

            If vType = 3 Then

                MsgBox("ยังไม่เปิด การใช้งานในส่วนการโอนข้อมูลประเภทนี้", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
                'vItemCode = UCase(Me.TextBox10.Text)

                ''=======================================================================================================================================================================================================

                'vTableName = "BCItem"
                'Call CheckDataBranch(1, vTableName, vItemCode)
                'If vCheckExist = 0 Then
                '    Me.TextBox12.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                '    Call CreateDataSource()
                '    Call BindGrid()
                '    Me.Button1.Visible = False
                '    Me.DataGrid1.Visible = False
                '    MsgBox("ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                '    Exit Sub
                'End If
                'Call PrepareDatabranch(1, vTableName, vItemCode)
                'Call DeleteTable(1, vTableName, vItemCode)


                'Dim vSalePrice1 As Double
                'Dim vSalePrice2 As Double
                'Dim vSalePrice3 As Double
                'Dim vSalePrice4 As Double
                'Dim vSalePrice5 As Double
                'Dim vSalePrice6 As Double
                'Dim vSalePrice7 As Double
                'Dim vSalePrice8 As Double
                'Dim vSalePrice9 As Double
                'Dim vISExport As Integer
                'Dim vPrintYear As Integer

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.bcitem where code ='" & vItemCode & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "BCItem")
                'dtFrom = dsFrom.Tables("BCItem")
                'If dsFrom.Tables("BCItem").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("BCItem").Rows.Count - 1

                '        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice1")) Then
                '            vSalePrice1 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice1")
                '        End If

                '        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice2")) Then
                '            vSalePrice2 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice2")
                '        End If

                '        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice3")) Then
                '            vSalePrice3 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice3")
                '        End If

                '        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice4")) Then
                '            vSalePrice4 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice4")
                '        End If

                '        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice5")) Then
                '            vSalePrice5 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice5")
                '        End If

                '        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice6")) Then
                '            vSalePrice6 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice6")
                '        End If

                '        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice7")) Then
                '            vSalePrice7 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice7")
                '        End If

                '        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice8")) Then
                '            vSalePrice8 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice8")
                '        End If

                '        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("saleprice9")) Then
                '            vSalePrice9 = dsFrom.Tables("BCItem").Rows(i).Item("saleprice9")
                '        End If

                '        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("isexport")) Then
                '            vISExport = dsFrom.Tables("BCItem").Rows(i).Item("isexport")
                '        End If

                '        If Not IsDBNull(dsFrom.Tables("BCItem").Rows(i).Item("printyear")) Then
                '            vPrintYear = dsFrom.Tables("BCItem").Rows(i).Item("printyear")
                '        End If

                '        vQuery = "exec dbo.USP_PTF_InsertItemMaster '" & dsFrom.Tables("BCItem").Rows(i).Item("Code") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Name1") & "', '" & dsFrom.Tables("BCItem").Rows(i).Item("Name2") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ShortName") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CategoryCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("GroupCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("BrandCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("TypeCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("FormatCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ColorCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("MyGrade") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("myclass") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("GenericCode") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("UnitType") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("DefStkUnitCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefSaleUnitCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefBuyUnitCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("OppositeUnit") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("OppositeUnit2") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("Opposite1Price") & "," & dsFrom.Tables("BCItem").Rows(i).Item("OppositeRate") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("MySize") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("Weight") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("LastUpdate") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("StockType") & "," & dsFrom.Tables("BCItem").Rows(i).Item("StockProcess") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("VenderCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DepositCondit") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DeliveryCharge") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("InstallCharge") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ReturnRemark") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("ReturnStatus") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("ReturnCharge") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ItemStatus") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("NewItemStatus") & "," & dsFrom.Tables("BCItem").Rows(i).Item("RenovateStatus") & "," & dsFrom.Tables("BCItem").Rows(i).Item("ExceptTax") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PromotionType") & "," & dsFrom.Tables("BCItem").Rows(i).Item("WTaxRate") & " ,'" & dsFrom.Tables("BCItem").Rows(i).Item("WaveFile") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("CostType") & "," & dsFrom.Tables("BCItem").Rows(i).Item("OrderPoint") & "," & dsFrom.Tables("BCItem").Rows(i).Item("StockMin") & "," & dsFrom.Tables("BCItem").Rows(i).Item("StockMax") & "," & dsFrom.Tables("BCItem").Rows(i).Item("StockQty") & "," & dsFrom.Tables("BCItem").Rows(i).Item("AverageCost") & "," & dsFrom.Tables("BCItem").Rows(i).Item("Amount") & "," & dsFrom.Tables("BCItem").Rows(i).Item("CostOfArea") & "," & dsFrom.Tables("BCItem").Rows(i).Item("LastBuyPrice") & "," & dsFrom.Tables("BCItem").Rows(i).Item("StandardCost") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("PicFileName1") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("PicFileName2") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("PicFileName3") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("PicFileName4") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("SpecFileName") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("AviFileName") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("AccGroupCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("MyDescription") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefBuyWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefBuyShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefSaleWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefSaleShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefBadWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefBadShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefFGWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefFGShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefRepairWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefRepairShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefRawWHCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("DefRawShelf") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Formula1") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Formula2") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Formula3") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Discount") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("ReserveQty") & "," & dsFrom.Tables("BCItem").Rows(i).Item("RemainOutQty") & "," & dsFrom.Tables("BCItem").Rows(i).Item("RemainInQty") & "," & dsFrom.Tables("BCItem").Rows(i).Item("BasePrice") & "," & dsFrom.Tables("BCItem").Rows(i).Item("ActiveStatus") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("CreatorCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CreateDateTime") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("LastEditorCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("LastEditDateT") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ConfirmCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ConfirmDateTime") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CancelCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CancelDateTime") & "'," & vSalePrice1 & "," & vSalePrice2 & "," & vSalePrice3 & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("DefFixUnitCode") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("DefFixUnitCost") & "," & dsFrom.Tables("BCItem").Rows(i).Item("FixQty") & "," & dsFrom.Tables("BCItem").Rows(i).Item("FixCost") & "," & dsFrom.Tables("BCItem").Rows(i).Item("SalePriceNV1") & "," & dsFrom.Tables("BCItem").Rows(i).Item("SalePriceNV2") & "," & dsFrom.Tables("BCItem").Rows(i).Item("SalePriceNV3") & "," & dsFrom.Tables("BCItem").Rows(i).Item("ProcessStatus") & "," & dsFrom.Tables("BCItem").Rows(i).Item("IsGift") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PromoMember") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PromoType") & "," & dsFrom.Tables("BCItem").Rows(i).Item("AverageCostReal") & "," & dsFrom.Tables("BCItem").Rows(i).Item("OverReceive") & "," & dsFrom.Tables("BCItem").Rows(i).Item("ContainerCapacity") & "," & dsFrom.Tables("BCItem").Rows(i).Item("ContainerWeight") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("CapacityUnit") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("LastBuyPriceHome") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("LastBuyCurrencyCode") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("LastAvgCost") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("UserGroup") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("stockqty1") & "," & dsFrom.Tables("BCItem").Rows(i).Item("stockqty2") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("StockQtyWord") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CustGroup") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("PublisherCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("WriterCode") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("PrintOrdinal") & "," & vPrintYear & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("LastSaleDate") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("LastBuyDate") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("LastSalePrice") & "," & dsFrom.Tables("BCItem").Rows(i).Item("LastSalePriceHome") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("LastSaleCurrencyCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Specification") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("ManufactoryCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("AFT_remark") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("UnitWeight") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("WeightUnitCode") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("LeadTime") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("OfferedBy") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("OfferedDate") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("HS_SMX_remark") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Description") & "'," & vSalePrice4 & "," & vSalePrice5 & "," & vSalePrice6 & "," & vSalePrice7 & "," & vSalePrice8 & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT1") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT2") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT3") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT4") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT5") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT6") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT7") & "," & dsFrom.Tables("BCItem").Rows(i).Item("PRICEVATOUT8") & ",'" & dsFrom.Tables("BCItem").Rows(i).Item("ITEMBARCODE") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("PRICECODE") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("Merchandiser") & "','" & dsFrom.Tables("BCItem").Rows(i).Item("CLASSFICATION") & "'," & dsFrom.Tables("BCItem").Rows(i).Item("CANMAKEPO") & "," & dsFrom.Tables("BCItem").Rows(i).Item("DISCCASHCARD") & "," & vISExport & "," & dsFrom.Tables("BCItem").Rows(i).Item("ISVOLUME") & ""
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If
                'Me.DataGrid1.Items(0).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(0).Cells(2).Text = "โอนเรียบร้อย"
                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vitemcode)

                'vQuery = "update dbo.bcitem set defbuywhcode = 'S01',defbuyshelf = 'AVL',defsalewhcode = 'S01',defsaleshelf = 'AVL',defbadwhcode = 'S01',Defbadshelf='DMG',deffgwhcode = 'S01',deffgshelf = 'AVL',defrepairwhcode = 'S01',defrepairshelf = 'RTV',defrawwhcode = 'S01',defrawshelf = 'AVL' where code = '" & vItemCode & "'"
                'With cmd
                '    .CommandType = CommandType.Text
                '    .CommandText = vQuery
                '    .Connection = vConnectionTo
                '    .ExecuteNonQuery()
                'End With

                ''=======================================================================================================================================================================================================

                'vTableName = "BCBarcodeMaster"
                'Call PrepareDatabranch(2, vTableName, vItemCode)
                'Call DeleteTable(2, vTableName, vItemCode)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.bcbarcodemaster where itemcode ='" & vItemCode & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "Data")
                'dtFrom = dsFrom.Tables("Data")
                'If dsFrom.Tables("Data").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                '        If Not IsDBNull(dsFrom.Tables("Data").Rows(i).Item("isexport")) Then
                '            vISExport = dsFrom.Tables("Data").Rows(i).Item("isexport")
                '        End If

                '        vQuery = "exec dbo.USP_PTF_InsertBarCodeMaster '" & dsFrom.Tables("Data").Rows(i).Item("Barcode") & "','" & dsFrom.Tables("Data").Rows(i).Item("BarcodeName") & "','" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("Remark") & "'," & dsFrom.Tables("Data").Rows(i).Item("ActiveStatus") & ",'" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("WHCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("SHELFCODE") & "'," & vISExport & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If

                'vQuery = "update dbo.bcbarcodemaster set whcode = 'S01',shelfcode = 'AVL' where itemcode = '" & vItemCode & "'"
                'With cmd
                '    .CommandType = CommandType.Text
                '    .CommandText = vQuery
                '    .Connection = vConnectionTo
                '    .ExecuteNonQuery()
                'End With

                'Me.DataGrid1.Items(1).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(1).Cells(2).Text = "โอนเรียบร้อย"
                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vitemcode)

                ''=======================================================================================================================================================================================================

                'vTableName = "BCPriceList"
                'Call PrepareDatabranch(2, vTableName, vItemCode)
                'Call DeleteTable(2, vTableName, vItemCode)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.bcpricelist where itemcode ='" & vItemCode & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "Data")
                'dtFrom = dsFrom.Tables("Data")
                'If dsFrom.Tables("Data").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                '        If Not IsDBNull(dsFrom.Tables("Data").Rows(i).Item("isexport")) Then
                '            vISExport = dsFrom.Tables("Data").Rows(i).Item("isexport")
                '        End If
                '        vQuery = "exec dbo.USP_PTF_InsertPriceList '" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("Remark") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxType") & "," & dsFrom.Tables("Data").Rows(i).Item("FromQty") & "," & dsFrom.Tables("Data").Rows(i).Item("ToQty") & ",'" & dsFrom.Tables("Data").Rows(i).Item("StartDate") & "','" & dsFrom.Tables("Data").Rows(i).Item("StopDate") & "'," & dsFrom.Tables("Data").Rows(i).Item("SaleType") & "," & dsFrom.Tables("Data").Rows(i).Item("TransportType") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice1") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice2") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice3") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice4") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice5") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice6") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice7") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice8") & "," & dsFrom.Tables("Data").Rows(i).Item("LowPrice") & "," & dsFrom.Tables("Data").Rows(i).Item("LineNumber") & "," & vISExport & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If
                'Me.DataGrid1.Items(2).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(2).Cells(2).Text = "โอนเรียบร้อย"
                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vitemcode)

                ''=======================================================================================================================================================================================================

                'vTableName = "BPSPriceList"
                'Call PrepareDatabranch(2, vTableName, vItemCode)
                'Call DeleteTable(2, vTableName, vItemCode)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.bpspricelist where itemcode ='" & vItemCode & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "Data")
                'dtFrom = dsFrom.Tables("Data")
                'If dsFrom.Tables("Data").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                '        vQuery = "exec dbo.USP_PTF_InsertBPSPriceList '" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("Remark") & "'," & dsFrom.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsFrom.Tables("Data").Rows(i).Item("BarCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("FromQty") & "," & dsFrom.Tables("Data").Rows(i).Item("ToQty") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice1") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice2") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice3") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice4") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice5") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice6") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice7") & "," & dsFrom.Tables("Data").Rows(i).Item("SalePrice8") & "," & dsFrom.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If
                'Me.DataGrid1.Items(3).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(3).Cells(2).Text = "โอนเรียบร้อย"
                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vitemcode)

                ''=======================================================================================================================================================================================================

                'vTableName = "BCItemWareHouse"
                'Call PrepareDatabranch(2, vTableName, vItemCode)
                'Call DeleteTable(2, vTableName, vItemCode)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.bcitemwarehouse where itemcode ='" & vItemCode & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "Data")
                'dtFrom = dsFrom.Tables("Data")
                'If dsFrom.Tables("Data").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                '        vQuery = "exec dbo.USP_PTF_InsertItemWareHouse '" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("WHCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("ShelfCode") & "'," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If
                'Me.DataGrid1.Items(4).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(4).Cells(2).Text = "โอนเรียบร้อย"
                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vitemcode)

                ''=======================================================================================================================================================================================================

                'vTableName = "BCPriceErect"
                'Call PrepareDatabranch(2, vTableName, vItemCode)
                'Call DeleteTable(2, vTableName, vItemCode)

                'vCountTransfer = 0
                'vQuery = "select * from tempdb.dbo.bcpriceerect where itemcode ='" & vItemCode & "'"
                'daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                'dsFrom = New DataSet
                'daFrom.Fill(dsFrom, "Data")
                'dtFrom = dsFrom.Tables("Data")
                'If dsFrom.Tables("Data").Rows.Count > 0 Then
                '    For i = 0 To dsFrom.Tables("Data").Rows.Count - 1
                '        vQuery = "exec dbo.USP_PTF_InsertPriceErect '" & dsFrom.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsFrom.Tables("Data").Rows(i).Item("PriceErect") & "','" & dsFrom.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsFrom.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsFrom.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsFrom.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                '        With cmd
                '            .CommandType = CommandType.Text
                '            .CommandText = vQuery
                '            .Connection = vConnectionTo
                '            .ExecuteNonQuery()
                '        End With
                '        vCountTransfer = vCountTransfer + 1
                '    Next
                'End If
                'Me.DataGrid1.Items(5).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                'Me.DataGrid1.Items(5).Cells(2).Text = "โอนเรียบร้อย"
                'Call DropTableBranch(vTableName)
                'Call InsertLogs(1, vTableName, vitemcode)

            End If

            '=======================================================================================================================================================================================================
            '=======================================================================================================================================================================================================

            If vType = 4 Then

                vDocNo = UCase(Me.TextBox9.Text)

                vTableName = "BCARInvoice"
                Call CheckDataBranch(0, vTableName, vDocNo)
                If vCheckExist = 0 Then
                    Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                    Call CreateDataSource()
                    Call BindGrid()
                    Me.Button1.Visible = False
                    Me.DataGrid1.Visible = False
                    MsgBox("ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
                Call PrepareDataBranch(0, vTableName, vDocNo)
                Call DeleteTable(0, vTableName, vDocNo)

                Dim vTimeTransport As String
                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCARInvoice where docno ='" & vDocNo & "'"
                daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                dsTo = New DataSet
                daTo.Fill(dsTo, "Data")
                dtTo = dsTo.Tables("Data")
                If dsTo.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsTo.Tables("Data").Rows.Count - 1

                        If Not IsDBNull(dsTo.Tables("Data").Rows(i).Item("timetransport")) Then
                            vTimeTransport = dsTo.Tables("Data").Rows(i).Item("timetransport")
                        End If

                        vQuery = "exec dbo.USP_PTF_InsertARInvoice '" & dsTo.Tables("Data").Rows(i).Item("DocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("DocDate") & "','" & dsTo.Tables("Data").Rows(i).Item("TaxNo") & "'," & dsTo.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsTo.Tables("Data").Rows(i).Item("ArCode") & "','" & dsTo.Tables("Data").Rows(i).Item("PassBillTo") & "','" & dsTo.Tables("Data").Rows(i).Item("ArName") & "','" & dsTo.Tables("Data").Rows(i).Item("ArAddress") & "','" & dsTo.Tables("Data").Rows(i).Item("CashierCode") & "'," & dsTo.Tables("Data").Rows(i).Item("ShiftNo") & ",'" & dsTo.Tables("Data").Rows(i).Item("MachineNo") & "','" & dsTo.Tables("Data").Rows(i).Item("MachineCode") & "'," & dsTo.Tables("Data").Rows(i).Item("PosStatus") & ",'" & dsTo.Tables("Data").Rows(i).Item("BillTime") & "','" & dsTo.Tables("Data").Rows(i).Item("CreditType") & "','" & dsTo.Tables("Data").Rows(i).Item("CreditBranch") & "','" & dsTo.Tables("Data").Rows(i).Item("CreditDueDate") & "','" & dsTo.Tables("Data").Rows(i).Item("CreditNo") & "','" & dsTo.Tables("Data").Rows(i).Item("CofirmNo") & "','" & dsTo.Tables("Data").Rows(i).Item("ChargeWord") & "'," & dsTo.Tables("Data").Rows(i).Item("CreditBaseAmount") & "," & dsTo.Tables("Data").Rows(i).Item("ChargeAmount") & "," & dsTo.Tables("Data").Rows(i).Item("GrandTotal") & "," & dsTo.Tables("Data").Rows(i).Item("CoupongAmount") & ",'" & dsTo.Tables("Data").Rows(i).Item("CoupongDesc") & "'," & dsTo.Tables("Data").Rows(i).Item("ChangeAmount") & "," & dsTo.Tables("Data").Rows(i).Item("SumMarkCount1") & "," & dsTo.Tables("Data").Rows(i).Item("SumMarkCount2") & "," & dsTo.Tables("Data").Rows(i).Item("SumMarkValue1") & "," & dsTo.Tables("Data").Rows(i).Item("SumMarkValue2") & ",'" & dsTo.Tables("Data").Rows(i).Item("DepartCode") & "'," & dsTo.Tables("Data").Rows(i).Item("CreditDay") & "," & dsTo.Tables("Data").Rows(i).Item("DeliveryDay") & ",'" & dsTo.Tables("Data").Rows(i).Item("DeliveryDate") & "','" & dsTo.Tables("Data").Rows(i).Item("DueDate") & "','" & dsTo.Tables("Data").Rows(i).Item("PayBillDate") & "','" & dsTo.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsTo.Tables("Data").Rows(i).Item("SaleAreaCode") & "'," & dsTo.Tables("Data").Rows(i).Item("TaxRate") & "," & dsTo.Tables("Data").Rows(i).Item("IsConfirm") & ",'" & dsTo.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsTo.Tables("Data").Rows(i).Item("BillType") & ",'" & dsTo.Tables("Data").Rows(i).Item("BillGroup") & "','" & dsTo.Tables("Data").Rows(i).Item("RefDocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("CrAuthorizeMan") & "','" & dsTo.Tables("Data").Rows(i).Item("DeliveryAddr") & "','" & dsTo.Tables("Data").Rows(i).Item("ContactCode") & "','" & dsTo.Tables("Data").Rows(i).Item("TransportCode") & "'," & dsTo.Tables("Data").Rows(i).Item("SumOfItemAmount") & ",'" & dsTo.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsTo.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsTo.Tables("Data").Rows(i).Item("AfterDiscount") & "," & dsTo.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsTo.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsTo.Tables("Data").Rows(i).Item("TotalAmount") & "," & dsTo.Tables("Data").Rows(i).Item("DiscountCase") & "," & dsTo.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsTo.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsTo.Tables("Data").Rows(i).Item("SumCashAmount") & "," & dsTo.Tables("Data").Rows(i).Item("SumChqAmount") & "," & dsTo.Tables("Data").Rows(i).Item("SumCreditAmount") & "," & dsTo.Tables("Data").Rows(i).Item("SumBankAmount") & "," & dsTo.Tables("Data").Rows(i).Item("DepositIncTax") & "," & dsTo.Tables("Data").Rows(i).Item("SumOfDeposit1") & "," & dsTo.Tables("Data").Rows(i).Item("SumOfDeposit2") & "," & dsTo.Tables("Data").Rows(i).Item("SumOfWTax") & "," & dsTo.Tables("Data").Rows(i).Item("NetDebtAmount") & "," & dsTo.Tables("Data").Rows(i).Item("HomeAmount") & "," & dsTo.Tables("Data").Rows(i).Item("OtherIncome") & "," & dsTo.Tables("Data").Rows(i).Item("OtherExpense") & "," & dsTo.Tables("Data").Rows(i).Item("ExcessAmount1") & "," & dsTo.Tables("Data").Rows(i).Item("ExcessAmount2") & "," & dsTo.Tables("Data").Rows(i).Item("BillBalance") & ",'" & dsTo.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsTo.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsTo.Tables("Data").Rows(i).Item("GLFormat") & "'," & dsTo.Tables("Data").Rows(i).Item("GLStartPosting") & "," & dsTo.Tables("Data").Rows(i).Item("IsCancel") & "," & dsTo.Tables("Data").Rows(i).Item("IsCreditNote") & "," & dsTo.Tables("Data").Rows(i).Item("IsDebitNote") & "," & dsTo.Tables("Data").Rows(i).Item("IsProcessOK") & "," & dsTo.Tables("Data").Rows(i).Item("IsCompleteSave") & "," & dsTo.Tables("Data").Rows(i).Item("IsPostGL") & "," & dsTo.Tables("Data").Rows(i).Item("GLTransData") & "," & dsTo.Tables("Data").Rows(i).Item("PayBillStatus") & ",'" & dsTo.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsTo.Tables("Data").Rows(i).Item("RecurName") & "'," & dsTo.Tables("Data").Rows(i).Item("ExchangeProfit") & ",'" & dsTo.Tables("Data").Rows(i).Item("CustTypeCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CustGroupCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsTo.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsTo.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsTo.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsTo.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CancelDateTime") & "','" & dsTo.Tables("Data").Rows(i).Item("Remark1") & "','" & dsTo.Tables("Data").Rows(i).Item("Remark2") & "','" & dsTo.Tables("Data").Rows(i).Item("Remark3") & "','" & dsTo.Tables("Data").Rows(i).Item("Remark4") & "','" & dsTo.Tables("Data").Rows(i).Item("Remark5") & "'," & dsTo.Tables("Data").Rows(i).Item("IsReceiveMoney") & "," & dsTo.Tables("Data").Rows(i).Item("IsConditionSend") & "," & dsTo.Tables("Data").Rows(i).Item("TotalTransport") & "," & dsTo.Tables("Data").Rows(i).Item("PayBillAmount") & "," & dsTo.Tables("Data").Rows(i).Item("GrossWeight") & "," & dsTo.Tables("Data").Rows(i).Item("PrintCount") & ",'" & dsTo.Tables("Data").Rows(i).Item("SORefNo") & "'," & dsTo.Tables("Data").Rows(i).Item("HoldingStatus") & ",'" & dsTo.Tables("Data").Rows(i).Item("TimeTransport") & "'," & dsTo.Tables("Data").Rows(i).Item("IsImport") & ",'" & dsTo.Tables("Data").Rows(i).Item("JOBNO") & "'," & dsTo.Tables("Data").Rows(i).Item("REFTYPE") & "," & dsTo.Tables("Data").Rows(i).Item("BILLTEMPORARY") & "," & dsTo.Tables("Data").Rows(i).Item("ISMULTITAXABB") & ",'" & dsTo.Tables("Data").Rows(i).Item("TAXABBNO_1") & "','" & dsTo.Tables("Data").Rows(i).Item("TAXABBNO_2") & "','" & dsTo.Tables("Data").Rows(i).Item("TAXABBNO_3") & "','" & dsTo.Tables("Data").Rows(i).Item("TAXABBNO_4") & "','" & dsTo.Tables("Data").Rows(i).Item("TAXABBNO_5") & "'," & dsTo.Tables("Data").Rows(i).Item("TAXABBAMOUNT1") & "," & dsTo.Tables("Data").Rows(i).Item("TAXABBAMOUNT2") & "," & dsTo.Tables("Data").Rows(i).Item("TAXABBAMOUNT3") & "," & dsTo.Tables("Data").Rows(i).Item("TAXABBAMOUNT4") & "," & dsTo.Tables("Data").Rows(i).Item("TAXABBAMOUNT5") & ",'" & dsTo.Tables("Data").Rows(i).Item("APPROVEDCODE") & "','" & dsTo.Tables("Data").Rows(i).Item("APPROVEDDATE") & "','" & dsTo.Tables("Data").Rows(i).Item("CANCELDESC") & "','" & dsTo.Tables("Data").Rows(i).Item("SHIFTCODE") & "'," & dsTo.Tables("Data").Rows(i).Item("CREDITVAT") & "," & dsTo.Tables("Data").Rows(i).Item("CREDITSUMVAT") & "," & dsTo.Tables("Data").Rows(i).Item("OTHERAMOUNT") & "," & dsTo.Tables("Data").Rows(i).Item("OTHERFEE") & "," & dsTo.Tables("Data").Rows(i).Item("DIFFAMOUNT") & "," & dsTo.Tables("Data").Rows(i).Item("ISREWARD") & "," & dsTo.Tables("Data").Rows(i).Item("POSCREDIT") & ",'" & dsTo.Tables("Data").Rows(i).Item("USERGROUP") & "'," & dsTo.Tables("Data").Rows(i).Item("NETWEIGHT") & "," & dsTo.Tables("Data").Rows(i).Item("NUMOFPALLET") & ",'" & dsTo.Tables("Data").Rows(i).Item("INVOICETYPE") & "'," & dsTo.Tables("Data").Rows(i).Item("QTYAMOUNT") & "," & dsTo.Tables("Data").Rows(i).Item("QTYDEFAULT") & "," & dsTo.Tables("Data").Rows(i).Item("QTYCOPY") & "," & dsTo.Tables("Data").Rows(i).Item("MERGEITEM") & "," & dsTo.Tables("Data").Rows(i).Item("NEWLINE") & "," & dsTo.Tables("Data").Rows(i).Item("CALCTAXFLAG") & "," & dsTo.Tables("Data").Rows(i).Item("PRICECOPY") & "," & dsTo.Tables("Data").Rows(i).Item("WHCOPY") & ",'" & dsTo.Tables("Data").Rows(i).Item("METHODEPAYBILL") & "','" & dsTo.Tables("Data").Rows(i).Item("METHODEPAYBILL2") & "','" & dsTo.Tables("Data").Rows(i).Item("DOREFNO") & "','" & dsTo.Tables("Data").Rows(i).Item("PROMOTIONCODE") & "'," & dsTo.Tables("Data").Rows(i).Item("SUMOFWTAXCASH") & "," & dsTo.Tables("Data").Rows(i).Item("SUMBASEWTAXCASH") & "," & dsTo.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionFrom
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(0).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(0).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTableBranch(vTableName)
                Call InsertLogs(1, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCARInvoiceSub"

                Call PrepareDataBranch(0, vTableName, vDocNo)
                Call DeleteTable(0, vTableName, vDocNo)

                Dim vDORemainQTY As Double

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCARInvoiceSub where docno ='" & vDocNo & "'"
                daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                dsTo = New DataSet
                daTo.Fill(dsTo, "Data")
                dtTo = dsTo.Tables("Data")
                If dsTo.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsTo.Tables("Data").Rows.Count - 1

                        If Not IsDBNull(dsTo.Tables("Data").Rows(i).Item("DORemainQTY")) Then
                            vDORemainQTY = dsTo.Tables("Data").Rows(i).Item("DORemainQTY")
                        End If

                        vQuery = "exec dbo.USP_PTF_InsertARInvoiceSub " & dsTo.Tables("Data").Rows(i).Item("BehindIndex") & "," & dsTo.Tables("Data").Rows(i).Item("MyType") & ",'" & dsTo.Tables("Data").Rows(i).Item("DocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("TaxNo") & "'," & dsTo.Tables("Data").Rows(i).Item("TaxType") & ",'" & dsTo.Tables("Data").Rows(i).Item("ItemCode") & "','" & dsTo.Tables("Data").Rows(i).Item("DocDate") & "','" & dsTo.Tables("Data").Rows(i).Item("ArCode") & "','" & dsTo.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsTo.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsTo.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsTo.Tables("Data").Rows(i).Item("ItemName") & "','" & dsTo.Tables("Data").Rows(i).Item("WHCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ShelfCode") & "'," & dsTo.Tables("Data").Rows(i).Item("CNQty") & "," & dsTo.Tables("Data").Rows(i).Item("Qty") & "," & dsTo.Tables("Data").Rows(i).Item("Price") & ",'" & dsTo.Tables("Data").Rows(i).Item("DiscountWord") & "'," & dsTo.Tables("Data").Rows(i).Item("DiscountAmount") & "," & dsTo.Tables("Data").Rows(i).Item("Amount") & "," & dsTo.Tables("Data").Rows(i).Item("NetAmount") & "," & dsTo.Tables("Data").Rows(i).Item("HomeAmount") & "," & dsTo.Tables("Data").Rows(i).Item("SumOfCost") & "," & dsTo.Tables("Data").Rows(i).Item("BalanceAmount") & ",'" & dsTo.Tables("Data").Rows(i).Item("UnitCode") & "','" & dsTo.Tables("Data").Rows(i).Item("OppositeUnit") & "'," & dsTo.Tables("Data").Rows(i).Item("OppositeQty") & "," & dsTo.Tables("Data").Rows(i).Item("OppositePrice2") & ",'" & dsTo.Tables("Data").Rows(i).Item("SORefNo") & "','" & dsTo.Tables("Data").Rows(i).Item("PORefNo") & "'," & dsTo.Tables("Data").Rows(i).Item("StockType") & "," & dsTo.Tables("Data").Rows(i).Item("ExceptTax") & "," & dsTo.Tables("Data").Rows(i).Item("LineNumber") & "," & dsTo.Tables("Data").Rows(i).Item("RefLineNumber") & "," & dsTo.Tables("Data").Rows(i).Item("TransState") & "," & dsTo.Tables("Data").Rows(i).Item("IsCancel") & ",'" & dsTo.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsTo.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsTo.Tables("Data").Rows(i).Item("BarCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CustTypeCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CustGroupCode") & "','" & dsTo.Tables("Data").Rows(i).Item("SaleAreaCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CategoryCode") & "','" & dsTo.Tables("Data").Rows(i).Item("GroupCode") & "','" & dsTo.Tables("Data").Rows(i).Item("BrandCode") & "','" & dsTo.Tables("Data").Rows(i).Item("TypeCode") & "','" & dsTo.Tables("Data").Rows(i).Item("FormatCode") & "','" & dsTo.Tables("Data").Rows(i).Item("MachineNo") & "','" & dsTo.Tables("Data").Rows(i).Item("MachineCode") & "','" & dsTo.Tables("Data").Rows(i).Item("BillTime") & "','" & dsTo.Tables("Data").Rows(i).Item("CashierCode") & "'," & dsTo.Tables("Data").Rows(i).Item("ShiftNo") & "," & dsTo.Tables("Data").Rows(i).Item("PosStatus") & "," & dsTo.Tables("Data").Rows(i).Item("PriceStatus") & "," & dsTo.Tables("Data").Rows(i).Item("IsPromotion") & "," & dsTo.Tables("Data").Rows(i).Item("PremiumStatus") & "," & dsTo.Tables("Data").Rows(i).Item("FixUnitCost") & "," & dsTo.Tables("Data").Rows(i).Item("FixUnitQty") & "," & dsTo.Tables("Data").Rows(i).Item("IsConditionSend") & "," & dsTo.Tables("Data").Rows(i).Item("TransportAmount") & "," & dsTo.Tables("Data").Rows(i).Item("AVERAGECOST") & ",'" & dsTo.Tables("Data").Rows(i).Item("LotNumber") & "','" & dsTo.Tables("Data").Rows(i).Item("Colorcode") & "','" & dsTo.Tables("Data").Rows(i).Item("SizeCode") & "','" & dsTo.Tables("Data").Rows(i).Item("StyleCode") & "','" & dsTo.Tables("Data").Rows(i).Item("itemsetcode") & "','" & dsTo.Tables("Data").Rows(i).Item("JobNo") & "'," & dsTo.Tables("Data").Rows(i).Item("TAXRATE") & "," & dsTo.Tables("Data").Rows(i).Item("PACKINGRATE1") & "," & dsTo.Tables("Data").Rows(i).Item("PACKINGRATE2") & "," & dsTo.Tables("Data").Rows(i).Item("REFTYPE") & ",'" & dsTo.Tables("Data").Rows(i).Item("SHIFTCODE") & "'," & dsTo.Tables("Data").Rows(i).Item("PROMOSTATUS") & "," & dsTo.Tables("Data").Rows(i).Item("OLDPRICE") & ",'" & dsTo.Tables("Data").Rows(i).Item("USERCODE") & "'," & dsTo.Tables("Data").Rows(i).Item("USERMODIFY") & "," & dsTo.Tables("Data").Rows(i).Item("POSCREDIT") & ",'" & dsTo.Tables("Data").Rows(i).Item("USERGROUP") & "','" & dsTo.Tables("Data").Rows(i).Item("PRICECODE") & "','" & dsTo.Tables("Data").Rows(i).Item("INVOICETYPE") & "'," & dsTo.Tables("Data").Rows(i).Item("ISPROCESS") & "," & dsTo.Tables("Data").Rows(i).Item("ISLOCKCOST") & ",'" & dsTo.Tables("Data").Rows(i).Item("ITEMNAMEDESC") & "','" & dsTo.Tables("Data").Rows(i).Item("DOREFNO") & "'," & dsTo.Tables("Data").Rows(i).Item("DELIVERYSTATUS") & ",'" & dsTo.Tables("Data").Rows(i).Item("PROMOTIONCODE") & "','" & dsTo.Tables("Data").Rows(i).Item("MASTERITEMCODE") & "'," & dsTo.Tables("Data").Rows(i).Item("BTDISC") & "," & dsTo.Tables("Data").Rows(i).Item("DISCCASHCARD") & "," & dsTo.Tables("Data").Rows(i).Item("WTAXAMOUNT") & "," & dsTo.Tables("Data").Rows(i).Item("BASEWTAX") & "," & dsTo.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsTo.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsTo.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsTo.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsTo.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'," & vDORemainQTY & ""
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionFrom
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(1).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(1).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTableBranch(vTableName)
                Call InsertLogs(1, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCRecMoney"

                Call PrepareDataBranch(0, vTableName, vDocNo)
                Call DeleteTable(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCRecMoney where docno ='" & vDocNo & "'"
                daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                dsTo = New DataSet
                daTo.Fill(dsTo, "Data")
                dtTo = dsTo.Tables("Data")
                If dsTo.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsTo.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertRecMoney '" & dsTo.Tables("Data").Rows(i).Item("DocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("DocDate") & "','" & dsTo.Tables("Data").Rows(i).Item("ArCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsTo.Tables("Data").Rows(i).Item("ExchangeRate") & "," & dsTo.Tables("Data").Rows(i).Item("PayAmount") & "," & dsTo.Tables("Data").Rows(i).Item("HomeAmount") & "," & dsTo.Tables("Data").Rows(i).Item("ChqTotalAmount") & "," & dsTo.Tables("Data").Rows(i).Item("PaymentType") & "," & dsTo.Tables("Data").Rows(i).Item("SaveFrom") & "," & dsTo.Tables("Data").Rows(i).Item("PayChqState") & ",'" & dsTo.Tables("Data").Rows(i).Item("CreditType") & "'," & dsTo.Tables("Data").Rows(i).Item("ChargeAmount") & ",'" & dsTo.Tables("Data").Rows(i).Item("ChargeWord") & "','" & dsTo.Tables("Data").Rows(i).Item("ConfirmNo") & "'," & dsTo.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsTo.Tables("Data").Rows(i).Item("RefNo") & "','" & dsTo.Tables("Data").Rows(i).Item("BankCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsTo.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsTo.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsTo.Tables("Data").Rows(i).Item("BankBranchCode") & "','" & dsTo.Tables("Data").Rows(i).Item("TransBankDate") & "','" & dsTo.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsTo.Tables("Data").Rows(i).Item("RefDate") & "'," & dsTo.Tables("Data").Rows(i).Item("CHANGEAMOUNT") & "," & dsTo.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsTo.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsTo.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsTo.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsTo.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionFrom
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(2).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(2).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTableBranch(vTableName)
                Call InsertLogs(1, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCOutPutTax"

                Call PrepareDataBranch(0, vTableName, vDocNo)
                Call DeleteTable(0, vTableName, vDocNo)

                Dim vExchangeRate As Double

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCOutPutTax where docno ='" & vDocNo & "'"
                daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                dsTo = New DataSet
                daTo.Fill(dsTo, "Data")
                dtTo = dsTo.Tables("Data")
                If dsTo.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsTo.Tables("Data").Rows.Count - 1

                        If Not IsDBNull(dsTo.Tables("Data").Rows(i).Item("ExchangeRate")) Then
                            vExchangeRate = dsTo.Tables("Data").Rows(i).Item("ExchangeRate")
                        End If

                        vQuery = "exec dbo.USP_PTF_InsertOutPutTax " & dsTo.Tables("Data").Rows(i).Item("SaveFrom") & " ,'" & dsTo.Tables("Data").Rows(i).Item("DocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("BookCode") & "'," & dsTo.Tables("Data").Rows(i).Item("Source") & ",'" & dsTo.Tables("Data").Rows(i).Item("DocDate") & "','" & dsTo.Tables("Data").Rows(i).Item("TaxDate") & "','" & dsTo.Tables("Data").Rows(i).Item("TaxNo") & "','" & dsTo.Tables("Data").Rows(i).Item("ArCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ShortTaxDesc") & "'," & dsTo.Tables("Data").Rows(i).Item("TaxRate") & "," & dsTo.Tables("Data").Rows(i).Item("Process") & "," & dsTo.Tables("Data").Rows(i).Item("BeforeTaxAmount") & "," & dsTo.Tables("Data").Rows(i).Item("TaxAmount") & "," & dsTo.Tables("Data").Rows(i).Item("ExceptTaxAmount") & "," & dsTo.Tables("Data").Rows(i).Item("ZeroTaxAmount") & "," & dsTo.Tables("Data").Rows(i).Item("LineNumber") & "," & dsTo.Tables("Data").Rows(i).Item("IsMultiCurrency") & "," & dsTo.Tables("Data").Rows(i).Item("FAmount") & "," & vExchangeRate & ",'" & dsTo.Tables("Data").Rows(i).Item("TaxGroup") & "','" & dsTo.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsTo.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsTo.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsTo.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsTo.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsTo.Tables("Data").Rows(i).Item("IsCancel") & "," & dsTo.Tables("Data").Rows(i).Item("CancelOutPeriod") & ",'" & dsTo.Tables("Data").Rows(i).Item("CancelDocNo") & "'," & dsTo.Tables("Data").Rows(i).Item("IsPos") & ",'" & dsTo.Tables("Data").Rows(i).Item("CANCELDOCDATE") & "'," & dsTo.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionFrom
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(3).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(3).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTableBranch(vTableName)
                Call InsertLogs(1, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCCreditCard"

                Call PrepareDataBranch(0, vTableName, vDocNo)
                Call DeleteTable(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCCreditCard where docno ='" & vDocNo & "'"
                daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                dsTo = New DataSet
                daTo.Fill(dsTo, "Data")
                dtTo = dsTo.Tables("Data")
                If dsTo.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsTo.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertCreditCard '" & dsTo.Tables("Data").Rows(i).Item("BankCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CreditCardNo") & "','" & dsTo.Tables("Data").Rows(i).Item("DocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("ArCode") & "','" & dsTo.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ReceiveDate") & "','" & dsTo.Tables("Data").Rows(i).Item("DueDate") & "','" & dsTo.Tables("Data").Rows(i).Item("BookNo") & "'," & dsTo.Tables("Data").Rows(i).Item("Status") & "," & dsTo.Tables("Data").Rows(i).Item("SaveFrom") & ",'" & dsTo.Tables("Data").Rows(i).Item("StatusDate") & "','" & dsTo.Tables("Data").Rows(i).Item("StatusDocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsTo.Tables("Data").Rows(i).Item("BankBranchCode") & "'," & dsTo.Tables("Data").Rows(i).Item("Amount") & ",'" & dsTo.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsTo.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsTo.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsTo.Tables("Data").Rows(i).Item("CreditType") & "','" & dsTo.Tables("Data").Rows(i).Item("ConfirmNo") & "'," & dsTo.Tables("Data").Rows(i).Item("ChargeAmount") & ",'" & dsTo.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsTo.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsTo.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsTo.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsTo.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CancelDateTime") & "'," & dsTo.Tables("Data").Rows(i).Item("CreditVatRate") & "," & dsTo.Tables("Data").Rows(i).Item("CreditVat") & "," & dsTo.Tables("Data").Rows(i).Item("CreditSumVat") & "," & dsTo.Tables("Data").Rows(i).Item("PRINTCOUNT") & ",'" & dsTo.Tables("Data").Rows(i).Item("POSDOCNO") & "'," & dsTo.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionFrom
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(4).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(4).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTableBranch(vTableName)
                Call InsertLogs(1, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCCHQIN"

                Call PrepareDataBranch(0, vTableName, vDocNo)
                Call DeleteTable(0, vTableName, vDocNo)

                Dim vReciveConfirm As Integer
                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCCHQIN where docno ='" & vDocNo & "'"
                daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                dsTo = New DataSet
                daTo.Fill(dsTo, "Data")
                dtTo = dsTo.Tables("Data")
                If dsTo.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsTo.Tables("Data").Rows.Count - 1

                        If Not IsDBNull(dsTo.Tables("Data").Rows(i).Item("RECIVECONFIRM")) Then
                            vReciveConfirm = dsTo.Tables("Data").Rows(i).Item("RECIVECONFIRM")
                        End If

                        vQuery = "exec dbo.USP_PTF_InsertCHQIN '" & dsTo.Tables("Data").Rows(i).Item("BankCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ChqNumber") & "','" & dsTo.Tables("Data").Rows(i).Item("DocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("ArCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ExtendStatus") & "','" & dsTo.Tables("Data").Rows(i).Item("SaleCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ReceiveDate") & "','" & dsTo.Tables("Data").Rows(i).Item("DueDate") & "','" & dsTo.Tables("Data").Rows(i).Item("BookNo") & "'," & dsTo.Tables("Data").Rows(i).Item("Status") & "," & dsTo.Tables("Data").Rows(i).Item("SaveFrom") & ",'" & dsTo.Tables("Data").Rows(i).Item("StatusDate") & "','" & dsTo.Tables("Data").Rows(i).Item("StatusDocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsTo.Tables("Data").Rows(i).Item("BankBranchCode") & "'," & dsTo.Tables("Data").Rows(i).Item("Amount") & "," & dsTo.Tables("Data").Rows(i).Item("Balance") & ",'" & dsTo.Tables("Data").Rows(i).Item("MyDescription") & "'," & dsTo.Tables("Data").Rows(i).Item("ChqUseStatus") & ",'" & dsTo.Tables("Data").Rows(i).Item("CurrencyCode") & "'," & dsTo.Tables("Data").Rows(i).Item("ExchangeRate") & ",'" & dsTo.Tables("Data").Rows(i).Item("CreatorCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CreateDateTime") & "','" & dsTo.Tables("Data").Rows(i).Item("LastEditorCode") & "','" & dsTo.Tables("Data").Rows(i).Item("LastEditDateT") & "','" & dsTo.Tables("Data").Rows(i).Item("ConfirmCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ConfirmDateTime") & "','" & dsTo.Tables("Data").Rows(i).Item("CancelCode") & "','" & dsTo.Tables("Data").Rows(i).Item("CancelDateTime") & "','" & dsTo.Tables("Data").Rows(i).Item("RECIVECHQBY") & "'," & vReciveConfirm & "," & dsTo.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsTo.Tables("Data").Rows(i).Item("ISEXPORT") & ""
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionFrom
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(5).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(5).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTableBranch(vTableName)
                Call InsertLogs(1, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCTrans"

                Call PrepareDataBranch(0, vTableName, vDocNo)
                Call DeleteTable(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCTrans where docno ='" & vDocNo & "'"
                daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                dsTo = New DataSet
                daTo.Fill(dsTo, "Data")
                dtTo = dsTo.Tables("Data")
                If dsTo.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsTo.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertBCTrans '" & dsTo.Tables("Data").Rows(i).Item("BatchNo") & "','" & dsTo.Tables("Data").Rows(i).Item("BookCode") & "','" & dsTo.Tables("Data").Rows(i).Item("DocDate") & "','" & dsTo.Tables("Data").Rows(i).Item("DocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("RefDocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("RefDate") & "'," & dsTo.Tables("Data").Rows(i).Item("Amount") & "," & dsTo.Tables("Data").Rows(i).Item("FAmount") & ",'" & dsTo.Tables("Data").Rows(i).Item("FCurrency") & "','" & dsTo.Tables("Data").Rows(i).Item("FExchangeRate") & "'," & dsTo.Tables("Data").Rows(i).Item("Source") & "," & dsTo.Tables("Data").Rows(i).Item("TransType") & "," & dsTo.Tables("Data").Rows(i).Item("IsManualKey") & ",'" & dsTo.Tables("Data").Rows(i).Item("MyDescription") & "','" & dsTo.Tables("Data").Rows(i).Item("RecurName") & "'," & dsTo.Tables("Data").Rows(i).Item("IsConfirm") & "," & dsTo.Tables("Data").Rows(i).Item("IsCancel") & "," & dsTo.Tables("Data").Rows(i).Item("IsPassError") & "," & dsTo.Tables("Data").Rows(i).Item("TaxCount") & "," & dsTo.Tables("Data").Rows(i).Item("CheqCount") & "," & dsTo.Tables("Data").Rows(i).Item("PRINTCOUNT") & "," & dsTo.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsTo.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsTo.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsTo.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsTo.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionFrom
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(6).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(6).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTableBranch(vTableName)
                Call InsertLogs(1, vTableName, vDocNo)

                '=======================================================================================================================================================================================================

                vTableName = "BCTransSub"

                Call PrepareDataBranch(0, vTableName, vDocNo)
                Call DeleteTable(0, vTableName, vDocNo)

                vCountTransfer = 0
                vQuery = "select * from tempdb.dbo.BCTransSub where docno ='" & vDocNo & "'"
                daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                dsTo = New DataSet
                daTo.Fill(dsTo, "Data")
                dtTo = dsTo.Tables("Data")
                If dsTo.Tables("Data").Rows.Count > 0 Then
                    For i = 0 To dsTo.Tables("Data").Rows.Count - 1
                        vQuery = "exec dbo.USP_PTF_InsertBCTransSub '" & dsTo.Tables("Data").Rows(i).Item("BatchNo") & "'," & dsTo.Tables("Data").Rows(i).Item("LineNumber") & ",'" & dsTo.Tables("Data").Rows(i).Item("BookCode") & "','" & dsTo.Tables("Data").Rows(i).Item("DocDate") & "','" & dsTo.Tables("Data").Rows(i).Item("DocNo") & "','" & dsTo.Tables("Data").Rows(i).Item("AccountCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ExtDescription") & "','" & dsTo.Tables("Data").Rows(i).Item("DepartCode") & "','" & dsTo.Tables("Data").Rows(i).Item("ProjectCode") & "','" & dsTo.Tables("Data").Rows(i).Item("AllocateCode") & "','" & dsTo.Tables("Data").Rows(i).Item("PartCode") & "','" & dsTo.Tables("Data").Rows(i).Item("SideCode") & "','" & dsTo.Tables("Data").Rows(i).Item("JobCode") & "','" & dsTo.Tables("Data").Rows(i).Item("BranchCode") & "'," & dsTo.Tables("Data").Rows(i).Item("Debit") & "," & dsTo.Tables("Data").Rows(i).Item("Credit") & "," & dsTo.Tables("Data").Rows(i).Item("FDebit") & "," & dsTo.Tables("Data").Rows(i).Item("FCredit") & "," & dsTo.Tables("Data").Rows(i).Item("Source") & "," & dsTo.Tables("Data").Rows(i).Item("IsManualKey") & "," & dsTo.Tables("Data").Rows(i).Item("IsCancel") & "," & dsTo.Tables("Data").Rows(i).Item("IsConfirm") & "," & dsTo.Tables("Data").Rows(i).Item("ISEXPORT") & ",'" & dsTo.Tables("Data").Rows(i).Item("CREATORCODE") & "','" & dsTo.Tables("Data").Rows(i).Item("CREATEDATETIME") & "','" & dsTo.Tables("Data").Rows(i).Item("LASTEDITORCODE") & "','" & dsTo.Tables("Data").Rows(i).Item("LASTEDITDATET") & "'"
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = vQuery
                            .Connection = vConnectionFrom
                            .ExecuteNonQuery()
                        End With
                        vCountTransfer = vCountTransfer + 1
                    Next
                End If
                Me.DataGrid1.Items(7).Cells(1).Text = Format(vCountTransfer, "##,##0.00")
                Me.DataGrid1.Items(7).Cells(2).Text = "โอนเรียบร้อย"
                Call DropTableBranch(vTableName)
                Call InsertLogs(1, vTableName, vDocNo)

            End If

            '=======================================================================================================================================================================================================
            '=======================================================================================================================================================================================================
        End If

        Me.Label15.Visible = False
        Me.Label16.Visible = False
        Me.TextBox9.Visible = False
        Me.TextBox10.Visible = False
        Me.Button1.Visible = False
        Me.Button3.Visible = False
        Me.Button4.Visible = False
        Me.TextBox11.Visible = False
        Me.TextBox12.Visible = False
        Me.TextBox9.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox12.Text = ""

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim i As Integer

        If Me.DropDownList1.Text <> "" Then
            i = Me.DropDownList1.SelectedIndex

            If i = 0 And (vDepartment <> "AC" And vDepartment <> "IT") Then
                MsgBox("คุณไม่มีสิทธิ์ เข้าใช้งานในส่วนของการโอนข้อมูลใบสั่งขาย/จอง กรุณาแจ้งแผนกคอมฯ กรณีต้องการใช้งาน", MsgBoxStyle.Critical, "Send Error Message")
                Me.Visible = False
                Exit Sub
            End If

            If i = 1 And (vDepartment <> "MC" And vDepartment <> "IT") Then
                MsgBox("คุณไม่มีสิทธิ์ เข้าใช้งานในส่วนของการโอนข้อมูลใบสั่งซื้อ กรุณาแจ้งแผนกคอมฯ กรณีต้องการใช้งาน", MsgBoxStyle.Critical, "Send Error Message")
                Me.Visible = False
                Exit Sub
            End If

            If i = 2 And (vDepartment <> "AC" And vDepartment <> "IT") Then
                MsgBox("คุณไม่มีสิทธิ์ เข้าใช้งานในส่วนของการโอนข้อมูลใบโอนสินค้า กรุณาแจ้งแผนกคอมฯ กรณีต้องการใช้งาน", MsgBoxStyle.Critical, "Send Error Message")
                Me.Visible = False
                Exit Sub
            End If

            If i = 3 And (vDepartment <> "MC" And vDepartment <> "IT") Then
                MsgBox("คุณไม่มีสิทธิ์ เข้าใช้งานในส่วนของการโอนข้อมูลทะเบียนสินค้า กรุณาแจ้งแผนกคอมฯ กรณีต้องการใช้งาน", MsgBoxStyle.Critical, "Send Error Message")
                Me.Visible = False
                Exit Sub
            End If

            If i = 4 And (vDepartment <> "AC" And vDepartment <> "IT") Then
                MsgBox("คุณไม่มีสิทธิ์ เข้าใช้งานในส่วนของการโอนข้อมูลบิลขาย กรุณาแจ้งแผนกคอมฯ กรณีต้องการใช้งาน", MsgBoxStyle.Critical, "Send Error Message")
                Me.Visible = False
                Exit Sub
            End If

            Me.Button1.Visible = False
            Me.Button3.Visible = True
            Me.Button4.Visible = True

            Call CreateDataSource()
            Call BindGrid()

            Me.DataGrid1.Visible = False

            If i = 3 Then
                Me.Label16.Visible = True
                Me.TextBox10.Visible = True
                Me.TextBox12.Visible = True
                Me.TextBox10.Enabled = True
                Me.TextBox9.Enabled = False

                Me.Label15.Visible = False
                Me.TextBox9.Visible = False
                Me.TextBox11.Visible = False

                Dim vScript As String = "<SCRIPT language='javascript'>form1.TextBox10.focus();</SCRIPT>"
                Page.RegisterStartupScript("focus", vScript)
            Else
                Me.Label15.Visible = True
                Me.TextBox9.Visible = True
                Me.TextBox11.Visible = True
                Me.TextBox10.Enabled = False

                Me.TextBox9.Enabled = True

                Me.Label16.Visible = False
                Me.TextBox10.Visible = False
                Me.TextBox12.Visible = False

                Dim vScript As String = "<SCRIPT language='javascript'>form1.TextBox9.focus();</SCRIPT>"
                Page.RegisterStartupScript("focus", vScript)
            End If
        End If
    End Sub

    Private Sub CreateDataSource()
        On Error Resume Next
        Dim i As Integer
        Dim dr As DataRow

        dt.Columns.Clear()
        dt.Columns.Add(New DataColumn("ชื่อตาราง", GetType(String)))
        dt.Columns.Add(New DataColumn("จำนวนรายการ", GetType(String)))
        dt.Columns.Add(New DataColumn("สถานะการโอน", GetType(String)))

        'For i = 0 To dtFrom.Rows.Count - 1
        '    dr = dt.NewRow()
        '    dr(0) = Trim(dtFrom.Rows(i).Item("tablename"))
        '    dr(1) = String.Format("{0:N0}", (dtFrom.Rows(i).Item("recordcount")))
        '    dr(2) = String.Format("รอโอนข้อมูล")
        '    dt.Rows.Add(dr)
        'Next i
        dv = New DataView(dt)
        Return
    End Sub

    Public Sub SelectTable(ByVal vType As Integer)
        On Error Resume Next
        Dim i As Integer
        Dim dr As DataRow
        Dim vTableName As String

        vTableName = Me.DropDownList1.Items(Me.DropDownList1.SelectedIndex).Text
        dt.Columns.Clear()
        dt.Columns.Add(New DataColumn("ชื่อตาราง", GetType(String)))
        dt.Columns.Add(New DataColumn("จำนวนรายการ", GetType(String)))
        dt.Columns.Add(New DataColumn("สถานะการโอน", GetType(String)))

        If vType = 0 Then
            For i = 1 To 2
                If i = 1 Then
                    vTableName = "BCSaleOrder"
                ElseIf i = 2 Then
                    vTableName = "BCSaleOrderSub"
                End If
                dr = dt.NewRow()
                dr(0) = Trim(vTableName)
                dr(1) = String.Format(0)
                dr(2) = String.Format("{0:N0}", "รอโอนข้อมูล")
                dt.Rows.Add(dr)
            Next i
            dv = New DataView(dt)
        End If

        If vType = 1 Then
            For i = 1 To 2
                If i = 1 Then
                    vTableName = "BCPurchaseOrder"
                ElseIf i = 2 Then
                    vTableName = "BCPurchaseOrderSub"
                End If
                dr = dt.NewRow()
                dr(0) = Trim(vTableName)
                dr(1) = String.Format(0)
                dr(2) = String.Format("{0:N0}", "รอโอนข้อมูล")
                dt.Rows.Add(dr)
            Next i
            dv = New DataView(dt)
        End If

        If vType = 2 Then
            For i = 1 To 3
                If i = 1 Then
                    vTableName = "BCSTKTransfer"
                ElseIf i = 2 Then
                    vTableName = "BCSTKTransfSub"
                ElseIf i = 3 Then
                    vTableName = "BCSTKTransfSub3"
                End If
                dr = dt.NewRow()
                dr(0) = Trim(vTableName)
                dr(1) = String.Format(0)
                dr(2) = String.Format("{0:N0}", "รอโอนข้อมูล")
                dt.Rows.Add(dr)
            Next i
            dv = New DataView(dt)
        End If

        If vType = 3 Then
            For i = 1 To 6
                If i = 1 Then
                    vTableName = "BCItem"
                ElseIf i = 2 Then
                    vTableName = "BCBarcodeMaster"
                ElseIf i = 3 Then
                    vTableName = "BCPriceList"
                ElseIf i = 4 Then
                    vTableName = "BPSPriceList"
                ElseIf i = 5 Then
                    vTableName = "BCItemWareHouse"
                ElseIf i = 6 Then
                    vTableName = "BCPriceErect"
                End If
                dr = dt.NewRow()
                dr(0) = Trim(vTableName)
                dr(1) = String.Format(0)
                dr(2) = String.Format("{0:N0}", "รอโอนข้อมูล")
                dt.Rows.Add(dr)
            Next i
            dv = New DataView(dt)
        End If

        If vType = 4 Then
            For i = 1 To 8
                If i = 1 Then
                    vTableName = "BCARInvoice"
                ElseIf i = 2 Then
                    vTableName = "BCARInvoiceSub"
                ElseIf i = 3 Then
                    vTableName = "BCRecMoney"
                ElseIf i = 4 Then
                    vTableName = "BCOutPutTax"
                ElseIf i = 5 Then
                    vTableName = "BCCreditCard"
                ElseIf i = 6 Then
                    vTableName = "BCChqIn"
                ElseIf i = 7 Then
                    vTableName = "BCTrans"
                ElseIf i = 8 Then
                    vTableName = "BCTransSub"
                End If
                dr = dt.NewRow()
                dr(0) = Trim(vTableName)
                dr(1) = String.Format(0)
                dr(2) = String.Format("{0:N0}", "รอโอนข้อมูล")
                dt.Rows.Add(dr)
            Next i
            dv = New DataView(dt)
        End If
        Return
    End Sub
    Private Sub BindGrid()
        On Error Resume Next

        DataGrid1.DataSource = dv
        DataGrid1.DataBind()
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand

        Dim vType As Integer
        Dim vTableName As String
        Dim vDocNo As String
        Dim vItemCode As String
        Dim i As Integer
        Dim vTransType As Integer

        'เอาไว้ตรวจสอบข้อมูลว่า มีข้อมูลพร้อมที่จะโอนหรือยัง

        If Me.TextBox9.Text <> "" Or Me.TextBox10.Text <> "" Then
            vTransType = Me.DropDownList2.SelectedIndex
            vType = Me.DropDownList1.SelectedIndex
            Call SelectTable(vType)
            Call BindGrid()

            If vTransType = 0 Then 'โอนจากสำนักงานใหญ่ไปสาขา
                If vType = 0 Then 'ใบสั่งขาย/สั่งจอง
                    vDocNo = Me.TextBox9.Text

                    vTableName = "BCSaleOrder"

                    vQuery = "exec dbo.USP_PTF_SearchDataTransfer 0,'" & vTableName & "','" & vDocNo & "'"
                    daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                    dsFrom = New DataSet
                    daFrom.Fill(dsFrom, "Data")
                    dtFrom = dsFrom.Tables("Data")
                    If dsFrom.Tables("Data").Rows.Count > 0 Then
                        Me.TextBox11.Text = "ผู้สร้างเอกสาร : " & dsFrom.Tables("Data").Rows(i).Item("creatorcode") & "  " & "มูลค่าเอกสาร = " & Format(dsFrom.Tables("Data").Rows(i).Item("netamount"), "##,000.00")
                        Me.Button1.Visible = True
                        Me.DataGrid1.Visible = True
                    Else
                        Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                        Call CreateDataSource()
                        Call BindGrid()
                        Me.Button1.Visible = False
                        Me.DataGrid1.Visible = False
                    End If

                End If

                If vType = 1 Then 'ใบสั่งซื้อ
                    vDocNo = Me.TextBox9.Text

                    vTableName = "BCPurchaseOrder"

                    vQuery = "exec dbo.USP_PTF_SearchDataTransfer 0,'" & vTableName & "','" & vDocNo & "'"
                    daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                    dsFrom = New DataSet
                    daFrom.Fill(dsFrom, "Data")
                    dtFrom = dsFrom.Tables("Data")
                    If dsFrom.Tables("Data").Rows.Count > 0 Then
                        Me.TextBox11.Text = "ผู้สร้างเอกสาร : " & dsFrom.Tables("Data").Rows(i).Item("creatorcode") & "  " & "มูลค่าเอกสาร = " & Format(dsFrom.Tables("Data").Rows(i).Item("netamount"), "##,000.00")
                        Me.Button1.Visible = True
                        Me.DataGrid1.Visible = True
                    Else
                        Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                        Call CreateDataSource()
                        Call BindGrid()
                        Me.Button1.Visible = False
                        Me.DataGrid1.Visible = False
                    End If
                End If

                If vType = 2 Then 'ใบโอนสินค้า

                    vDocNo = Me.TextBox9.Text

                    vTableName = "BCStktransfer"

                    vQuery = "exec dbo.USP_PTF_SearchDataTransfer 0,'" & vTableName & "','" & vDocNo & "'"
                    daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                    dsFrom = New DataSet
                    daFrom.Fill(dsFrom, "Data")
                    dtFrom = dsFrom.Tables("Data")
                    If dsFrom.Tables("Data").Rows.Count > 0 Then
                        Me.TextBox11.Text = "ผู้สร้างเอกสาร : " & dsFrom.Tables("Data").Rows(i).Item("creatorcode") & "" '& Format(dsFrom.Tables("Data").Rows(i).Item("netamount"), "##,000.00")
                        Me.Button1.Visible = True
                        Me.DataGrid1.Visible = True
                    Else
                        Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                        Call CreateDataSource()
                        Call BindGrid()
                        Me.Button1.Visible = False
                        Me.DataGrid1.Visible = False
                    End If
                End If

                If vType = 3 Then 'ทะเบียนสินค้า
                    vItemCode = Me.TextBox10.Text

                    vTableName = "BCItem"

                    vQuery = "exec dbo.USP_PTF_SearchDataTransfer 1,'" & vTableName & "','" & vItemCode & "'"
                    daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                    dsFrom = New DataSet
                    daFrom.Fill(dsFrom, "Data")
                    dtFrom = dsFrom.Tables("Data")
                    If dsFrom.Tables("Data").Rows.Count > 0 Then
                        Me.TextBox12.Text = "ชื่อสินค้า : " & dsFrom.Tables("Data").Rows(i).Item("name1")
                        Me.Button1.Visible = True
                        Me.DataGrid1.Visible = True
                    Else
                        Me.TextBox12.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                        Call CreateDataSource()
                        Call BindGrid()
                        Me.Button1.Visible = False
                        Me.DataGrid1.Visible = False
                    End If

                End If

                If vType = 4 Then 'บิลขาย

                    vDocNo = Me.TextBox9.Text

                    vTableName = "BCARInvoice"

                    vQuery = "exec dbo.USP_PTF_SearchDataTransfer 0,'" & vTableName & "','" & vDocNo & "'"
                    daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
                    dsFrom = New DataSet
                    daFrom.Fill(dsFrom, "Data")
                    dtFrom = dsFrom.Tables("Data")
                    If dsFrom.Tables("Data").Rows.Count > 0 Then
                        Me.TextBox11.Text = "ผู้สร้างเอกสาร : " & dsFrom.Tables("Data").Rows(i).Item("creatorcode") & "  " & "มูลค่าเอกสาร = " & Format(dsFrom.Tables("Data").Rows(i).Item("netdebtamount"), "##,000.00")
                        Me.Button1.Visible = True
                        Me.DataGrid1.Visible = True
                    Else
                        Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                        Call CreateDataSource()
                        Call BindGrid()
                        Me.Button1.Visible = False
                        Me.DataGrid1.Visible = False
                    End If
                End If
            End If

            If vTransType = 1 Then 'โอนข้อมูลจากสาขามาสำนักงานใหญ่
                If vType = 0 Then 'ใบสั่งขาย/สั่งจอง
                    vDocNo = Me.TextBox9.Text

                    vTableName = "BCSaleOrder"

                    vQuery = "exec dbo.USP_PTF_SearchDataTransfer 0,'" & vTableName & "','" & vDocNo & "'"
                    daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                    dsTo = New DataSet
                    daTo.Fill(dsTo, "Data")
                    dtTo = dsTo.Tables("Data")
                    If dsTo.Tables("Data").Rows.Count > 0 Then
                        Me.TextBox11.Text = "ผู้สร้างเอกสาร : " & dsTo.Tables("Data").Rows(i).Item("creatorcode") & "  " & "มูลค่าเอกสาร = " & Format(dsTo.Tables("Data").Rows(i).Item("netamount"), "##,000.00")
                        Me.Button1.Visible = True
                        Me.DataGrid1.Visible = True
                    Else
                        Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                        Call CreateDataSource()
                        Call BindGrid()
                        Me.Button1.Visible = False
                        Me.DataGrid1.Visible = False
                    End If

                End If

                If vType = 1 Then 'ใบสั่งซื้อ
                    vDocNo = Me.TextBox9.Text

                    vTableName = "BCPurchaseOrder"

                    vQuery = "exec dbo.USP_PTF_SearchDataTransfer 0,'" & vTableName & "','" & vDocNo & "'"
                    daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                    dsTo = New DataSet
                    daTo.Fill(dsTo, "Data")
                    dtTo = dsTo.Tables("Data")
                    If dsTo.Tables("Data").Rows.Count > 0 Then
                        Me.TextBox11.Text = "ผู้สร้างเอกสาร : " & dsTo.Tables("Data").Rows(i).Item("creatorcode") & "  " & "มูลค่าเอกสาร = " & Format(dsTo.Tables("Data").Rows(i).Item("netamount"), "##,000.00")
                        Me.Button1.Visible = True
                        Me.DataGrid1.Visible = True
                    Else
                        Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                        Call CreateDataSource()
                        Call BindGrid()
                        Me.Button1.Visible = False
                        Me.DataGrid1.Visible = False
                    End If
                End If

                If vType = 2 Then 'ใบโอนสินค้า

                    vDocNo = Me.TextBox9.Text

                    vTableName = "BCStktransfer"

                    vQuery = "exec dbo.USP_PTF_SearchDataTransfer 0,'" & vTableName & "','" & vDocNo & "'"
                    daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                    dsTo = New DataSet
                    daTo.Fill(dsTo, "Data")
                    dtTo = dsTo.Tables("Data")
                    If dsTo.Tables("Data").Rows.Count > 0 Then
                        Me.TextBox11.Text = "ผู้สร้างเอกสาร : " & dsTo.Tables("Data").Rows(i).Item("creatorcode") & "  " & "มูลค่าเอกสาร = " & Format(dsTo.Tables("Data").Rows(i).Item("netamount"), "##,000.00")
                        Me.Button1.Visible = True
                        Me.DataGrid1.Visible = True
                    Else
                        Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                        Call CreateDataSource()
                        Call BindGrid()
                        Me.Button1.Visible = False
                        Me.DataGrid1.Visible = False
                    End If
                End If

                If vType = 3 Then ' ทะเบียนสินค้า
                    vItemCode = Me.TextBox10.Text

                    vTableName = "BCItem"

                    vQuery = "exec dbo.USP_PTF_SearchDataTransfer 1,'" & vTableName & "','" & vItemCode & "'"
                    daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                    dsTo = New DataSet
                    daTo.Fill(dsTo, "Data")
                    dtTo = dsTo.Tables("Data")
                    If dsTo.Tables("Data").Rows.Count > 0 Then
                        Me.TextBox12.Text = "ชื่อสินค้า : " & dsTo.Tables("Data").Rows(i).Item("name1")
                        Me.Button1.Visible = True
                        Me.DataGrid1.Visible = True
                    Else
                        Me.TextBox12.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                        Call CreateDataSource()
                        Call BindGrid()
                        Me.Button1.Visible = False
                        Me.DataGrid1.Visible = False
                    End If

                End If

                If vType = 4 Then 'บิลขาย

                    vDocNo = Me.TextBox9.Text

                    vTableName = "BCARInvoice"

                    vQuery = "exec dbo.USP_PTF_SearchDataTransfer 0,'" & vTableName & "','" & vDocNo & "'"
                    daTo = New SqlDataAdapter(vQuery, vConnectionTo)
                    dsTo = New DataSet
                    daTo.Fill(dsTo, "Data")
                    dtTo = dsTo.Tables("Data")
                    If dsTo.Tables("Data").Rows.Count > 0 Then
                        Me.TextBox11.Text = "ผู้สร้างเอกสาร : " & dsTo.Tables("Data").Rows(i).Item("creatorcode") & "  " & "มูลค่าเอกสาร = " & Format(dsTo.Tables("Data").Rows(i).Item("netdebtamount"), "##,000.00")
                        Me.Button1.Visible = True
                        Me.DataGrid1.Visible = True
                    Else
                        Me.TextBox11.Text = "ไม่พบข้อมูล ที่จะทำการโอนข้อมูล กรุณาตรวจสอบ"
                        Call CreateDataSource()
                        Call BindGrid()
                        Me.Button1.Visible = False
                        Me.DataGrid1.Visible = False
                    End If
                End If

            End If
        End If
    End Sub

    Public Sub CheckData(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand

        If Me.TextBox9.Text <> "" Or Me.TextBox10.Text <> "" Then

            vQuery = "exec dbo.USP_PTF_SearchDataTransfer " & vType & ",'" & vTableName & "','" & vValue & "'"
            daFrom = New SqlDataAdapter(vQuery, vConnectionFrom)
            dsFrom = New DataSet
            daFrom.Fill(dsFrom, "Data")
            dtFrom = dsFrom.Tables("Data")
            If dsFrom.Tables("Data").Rows.Count <= 0 Then
                vCheckExist = 0
            Else
                vCheckExist = 1
            End If
        End If
    End Sub

    Public Sub CheckDataBranch(ByVal vType As Integer, ByVal vTableName As String, ByVal vValue As String)
        Dim vQuery As String
        Dim cmd As SqlCommand = New SqlCommand

        If Me.TextBox9.Text <> "" Or Me.TextBox10.Text <> "" Then

            vQuery = "exec dbo.USP_PTF_SearchDataTransfer " & vType & ",'" & vTableName & "','" & vValue & "'"
            daTo = New SqlDataAdapter(vQuery, vConnectionTo)
            dsTo = New DataSet
            daTo.Fill(dsTo, "Data")
            dtTo = dsTo.Tables("Data")
            If dsTo.Tables("Data").Rows.Count <= 0 Then
                vCheckExist = 0
            Else
                vCheckExist = 1
            End If
        End If
    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        Session("_NewList") = Me.DropDownList1.SelectedIndex
    End Sub

    Protected Sub DropDownList2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList2.SelectedIndexChanged
        Session("_NewTransTypeList") = Me.DropDownList2.SelectedIndex
    End Sub

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.TextBox9.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox12.Text = ""
        Me.Button1.Visible = False

        Call CreateDataSource()
        Call BindGrid()

        Me.DataGrid1.Visible = False

    End Sub

End Class
