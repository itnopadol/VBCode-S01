Imports System.data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms
Imports System.Web.SessionState
Imports System.Net
Imports System.Web


'Imports System.Web.Configuration
'Imports System.IO.StringWriter '.stringWrite = new System.IO.StringWriter();
'Imports System.Web.UI.HtmlTextWriter ' htmlWrite = new HtmlTextWriter(stringWrite);
'Imports CrystalDecisions.CrystalReports.Engine
'Imports CrystalDecisions.ReportSource
'Imports CrystalDecisions.Shared
'Imports CrystalDecisions.Web
'Imports System.Data.OracleClient


Partial Class PGPickup
    Inherits System.Web.UI.Page

    Dim vQuery As String
    Dim Conn As New SqlConnection
    Dim vConnectionString As String = "Persist Security Info = False;User ID=vbuser;Password=132;Data Source = Nebula;Initial Catalog = BCNP"
    Dim vConnection As SqlConnection
    Dim vCommand As SqlCommand
    Dim da As SqlDataAdapter
    Dim ds As DataSet
    Dim dt As New DataTable
    Dim dv As DataView

    Dim dt1 As DataTable
    Dim dt2 As DataTable

    Dim vBarCode As String
    Dim vLicense As String
    Dim vSaleCode As String
    Dim vNetAmount As Double
    Dim vPointID As Integer

    Dim vRecQTY As Double
    Dim vRecItemCode As String
    Dim vRecItemName As String
    Dim vRecUnitCode As String
    Dim vRecOnHand As Double
    Dim vRecPrice As Double
    Dim vRecNetAmount As Double
    Dim vRecLicense As String
    Dim vRecSaleCode As String
    Dim vRecPointID As String

    Dim vPage As Integer
    Dim vIsOpen As Integer
    Dim vConnectZone As Integer
    Dim vCountItemZoneOld As Integer
    Dim vCountItemOld As Integer


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim vNewItemCode As String
        Dim vOldItemCode As String
        Dim vNewUnitCode As String
        Dim vOldUnitCode As String
        Dim vNewQty As Double
        Dim vOldQty As Double
        Dim i As Integer

        If Not IsPostBack Then
            Dim vScript As String = "<SCRIPT language='javascript'>form1.TBBarCode.focus();</SCRIPT>"
            Page.RegisterStartupScript("focus", vScript)

            If Request.QueryString.Count > 0 Then
                vRecLicense = Request.QueryString("License").ToString()
                vRecSaleCode = Request.QueryString("SaleCode").ToString()
                vRecPointID = Request.QueryString("PointID").ToString()
                vRecItemCode = Request.QueryString("BarCodeID").ToString()
                vRecQTY = Request.QueryString("QTY").ToString()
                vRecOnHand = Request.QueryString("OnHand").ToString()
                vRecUnitCode = Request.QueryString("UnitCode").ToString()
                vRecItemName = Request.QueryString("ItemName").ToString()
                vRecPrice = Request.QueryString("Price").ToString()
                vRecNetAmount = Request.QueryString("NetAmount").ToString()

                Me.LBLLicense.Text = vRecLicense
                Me.LBLSaleCode.Text = vRecSaleCode
                Me.LBLPointID.Text = vRecPointID
                Me.LBLNetAmount.Text = Format(vRecNetAmount, "##,##0.00")

                If vRecItemCode <> "" Then
                    Dim dat As DataTable

                    If Not IsNothing(Session("MyDataSource1")) Then
                        dat = Session("MyDataSource1")

                        If dat.Rows.Count > 0 Then
                            vNewItemCode = dat.Rows(0).Item(0)
                            vNewItemCode = dat.Rows(0).Item(0)
                            vNewItemCode = dat.Rows(0).Item(0)

                            vNewItemCode = vRecItemCode
                            vNewUnitCode = vRecUnitCode
                            vNewQty = vRecQTY

                            For i = 0 To dat.Rows.Count - 1
                                vOldItemCode = dat.Rows(i).Item(0)
                                vOldUnitCode = dat.Rows(i).Item(3)
                                vOldQty = dat.Rows(i).Item(2)

                                If vNewItemCode = vOldItemCode And vNewUnitCode = vOldUnitCode Then
                                    Dim dtt1 As DataTable = Me.GetDataSourceVisible()
                                    Me.BindGrid1(dtt1)
                                    Dim dtt2 As DataTable = Me.GetDataSourceHide()
                                    Me.BindGrid2(dtt2)
                                    Me.LBLMessage.Text = "สินค้ารหัส" & vNewItemCode & "มีอยู่แล้วในรายการที่" & i + 1
                                    Exit Sub
                                End If
                                i = i + 1
                            Next

                        End If

                    End If


                    Dim dt1 As DataTable = Me.GetDataSourceVisible()
                    Dim dr1 As DataRow = dt1.NewRow

                    dr1("รหัส") = vRecItemCode
                    dr1("ชื่อ") = vRecItemName
                    dr1("จำนวน") = vRecQTY
                    dr1("หน่วย") = vRecUnitCode
                    dr1("ราคา") = vRecPrice
                    dt1.Rows.Add(dr1)
                    Me.BindGrid1(dt1)

                    Dim dt2 As DataTable = Me.GetDataSourceHide()
                    Dim dr2 As DataRow = dt2.NewRow

                    dr2("รหัส") = vRecItemCode
                    dr2("ชื่อ") = vRecItemName
                    dr2("จำนวน") = vRecQTY
                    dr2("หน่วย") = vRecUnitCode
                    dr2("ราคา") = vRecPrice
                    dr2("รวม") = vRecPrice * vRecQTY
                    dr2("คลัง") = "S01"
                    dr2("ชั้นเก็บ") = "AVL"
                    dr2("บาร์") = vRecItemCode
                    dr2("ที่เก็บ") = ""
                    dr2("โซน") = Me.LBLPointID.Text
                    dr2("จุดจ่าย") = ""

                    dt2.Rows.Add(dr2)
                    Me.BindGrid2(dt2)
                End If
            End If
        End If
    End Sub

    Private Function GetDataSourceVisible() As DataTable
        Dim dt1 As DataTable = Session("MyDataSource1")
        If dt1 Is Nothing Then
            dt1 = New DataTable()
            dt1.Columns.Add("รหัส")
            dt1.Columns.Add("ชื่อ")
            dt1.Columns.Add("จำนวน")
            dt1.Columns.Add("หน่วย")
            dt1.Columns.Add("ราคา")
            Session("MyDataSource1") = dt1
        End If
        Return dt1
    End Function

    Private Function ClearSession()
        Me.GridView1.DataBind()
        Me.GridView2.DataBind()
        Me.LBLNetAmount.Text = ""
        Session("MyDataSource1") = Nothing
        Session("MyDataSource2") = Nothing
        Session.RemoveAll()
        Session.Remove("MyDataSource1")
        Session.Remove("MyDataSource2")
    End Function

    Private Function GetDataSourceHide() As DataTable
        Dim dt2 As DataTable = Session("MyDataSource2")
        If dt2 Is Nothing Then
            dt2 = New DataTable()
            dt2.Columns.Add("รหัส")
            dt2.Columns.Add("ชื่อ")
            dt2.Columns.Add("จำนวน")
            dt2.Columns.Add("หน่วย")
            dt2.Columns.Add("ราคา")
            dt2.Columns.Add("รวม")
            dt2.Columns.Add("คลัง")
            dt2.Columns.Add("ชั้นเก็บ")
            dt2.Columns.Add("บาร์")
            dt2.Columns.Add("ที่เก็บ")
            dt2.Columns.Add("โซน")
            dt2.Columns.Add("จุดจ่าย")

            Session("MyDataSource2") = dt2
        End If
        Return dt2
    End Function

    Private Sub BindGrid1(ByVal dt1 As DataTable)
        GridView2.DataSource = dt1
        GridView2.DataBind()
    End Sub

    Private Sub BindGrid2(ByVal dt2 As DataTable)
        GridView1.DataSource = dt2
        GridView1.DataBind()
    End Sub

    'Protected Sub LinkButton4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkButton4.Click

    '    Dim dt1 As DataTable = Me.GetDataSourceVisible()
    '    Dim dr1 As DataRow = dt1.NewRow

    '    dr1("รหัสสินค้า") = vRecItemCode
    '    dr1("ชื่อสินค้า") = vRecItemName
    '    dr1("จำนวน") = vRecQTY
    '    dr1("หน่วย") = vRecUnitCode
    '    dt1.Rows.Add(dr1)
    '    Me.BindGrid1(dt1)

    '    Dim dt2 As DataTable = Me.GetDataSourceHide()
    '    Dim dr2 As DataRow = dt2.NewRow

    '    dr2("รหัสสินค้า") = vRecItemCode
    '    dr2("ชื่อสินค้า") = vRecItemName
    '    dr2("จำนวน") = vRecQTY
    '    dr2("หน่วย") = vRecUnitCode
    '    dr2("คลัง") = ""
    '    dr2("ชั้นเก็บ") = ""
    '    dr2("โซน") = ""
    '    dt2.Rows.Add(dr2)
    '    Me.BindGrid2(dt2)
    'End Sub

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        ''If e.Row.RowType = DataControlRowType.DataRow Then
        ''    Session("val") = e.Row.Cells(0).Text
        ''    e.Row.Attributes.Add("onmouseover", "javascript:this.style.cursor = 'pointer';")
        ''    e.Row.Attributes.Add("onclick", "win = window.location.href='otherpage.aspx';")
        ''End If

        'If e.Row.RowType = DataControlRowType.DataRow Then

        '    e.Row.Attributes("onmouseover") = "this.style.background = '#CCCCCC';"
        '    e.Row.Attributes("onmouseout") = "this.style.background = '#FFFFFF';"

        '    e.Row.Attributes("onclick") = ClientScript.GetPostBackClientHyperlink(Me.GridView2, e.Row.RowIndex)
        'End If
    End Sub

    Protected Sub GridView2_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles GridView2.RowDeleted

    End Sub

    Protected Sub GridView2_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView2.RowDeleting
        Dim vQty As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vNetAmount As Double

        Dim dt1 As DataTable = Me.GetDataSourceVisible()
        dt1.Rows(e.RowIndex).Delete()
        Me.BindGrid1(dt1)

        GridView2.DataSource = dt1
        GridView2.DataBind()

        Dim dt2 As DataTable = Me.GetDataSourceHide()
        dt2.Rows(e.RowIndex).Delete()
        Me.BindGrid2(dt2)

        GridView1.DataSource = dt2
        GridView1.DataBind()

        For Each gr As GridViewRow In Me.GridView2.Rows

            vQty = gr.Cells(3).Text
            vPrice = gr.Cells(5).Text
            vAmount = vQty * vPrice

            vNetAmount = vNetAmount + vAmount
        Next

        Me.LBLNetAmount.Text = Format(vNetAmount, "##,##0.00")
    End Sub

    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged

    End Sub

    Public Sub SaveData()
        Dim vCountItem As Integer
        Dim vHeader As String
        Dim vNumber As Integer
        Dim vDocNumber As String

        Dim vDocNo As String
        Dim vDocDate As String
        Dim vARCode As String
        Dim vSaleCode As String
        Dim vMemberID As String
        Dim vRefNo As String
        Dim vTotalNetAmount As Double
        Dim vBeforeTaxAmount As Double
        Dim vTaxAmount As Double

        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vDiscountWord As String
        Dim vDiscountAmount As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vUserID As String
        Dim i As Integer
        Dim vBarCode As String
        Dim vShelfID As String
        Dim vZoneID As String
        Dim vLinePickZone As String
        Dim vLineNumber As Integer

        Dim a As Integer
        Dim b As Integer
        Dim vCheckItemCode As String
        Dim vCheckUnitCode As String
        Dim vCheckBarCode As String
        Dim vCheckWHCode As String
        Dim vCheckShelfCode As String
        Dim vCheckZoneID As String
        Dim vCheckPickZone As String

        Dim vOldItem As String
        Dim vOldUnit As String
        Dim vOldBar As String
        Dim vOldWH As String
        Dim vOldShelf As String
        Dim vOldZone As String
        Dim vOldPick As String
        Dim vOld As Integer

        Dim vCountItemPickZone As Integer
        Dim vItemPickZone As String
        Dim vCount As Integer
        Dim vQueZone As String

        Dim vCheckIsConfirm As Integer
        Dim vCheckHoldBillNo As String

        Dim vInstrAr As Integer
        Dim vLenAr As Integer
        Dim vInstrSale As Integer
        Dim vLenSale As Integer


        On Error GoTo ErrDescription

        If Me.GridView2.Rows.Count > 0 And Me.LBLNetAmount.Text <> "" Then
            vUserID = Me.LBLUserID.Text

            If Me.LBLDocNo.Text = "" Then
                vQuery = "exec dbo.usp_np_searchnewdocno 29"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count > 0 Then
                    vHeader = Trim(ds.Tables(0).Rows(0)("header").ToString)
                    vNumber = ds.Tables(0).Rows(0)("autonumber").ToString
                    vDocNumber = Trim(ds.Tables(0).Rows(0)("docnumber").ToString)
                End If

                vDocNo = Trim(vDocNumber & vHeader & "-" & Format(vNumber, "0000"))
            Else
                vDocNo = Me.LBLDocNo.Text
            End If

            If vDocNo <> "" Then
                vQuery = "exec dbo.USP_NP_CheckDocDate"
                Dim vService7 As New WebReference.WebServiceCalc
                Dim ds7 As DataSet = vService7.vGetQueryAnlyzer(vQuery)
                If ds7.Tables(0).Rows.Count > 0 Then
                    vDocDate = ds7.Tables(0).Rows(0)("vdocdate").ToString
                End If

                vRefNo = Me.LBLLicense.Text

                vConnectZone = "01"
                vQueZone = "A"

                For Each gr As GridViewRow In Me.GridView1.Rows
                    'MsgBox(gr.Cells(1).Text & "," & gr.Cells(2).Text & "," & gr.Cells(3).Text & "," & gr.Cells(4).Text & "," & gr.Cells(5).Text & "," & gr.Cells(6).Text & "," & gr.Cells(7).Text & "," & gr.Cells(8).Text & "," & gr.Cells(9).Text & "," & gr.Cells(10).Text)
                    vItemPickZone = gr.Cells(10).Text
                    If vConnectZone = vItemPickZone Then
                        vCountItemPickZone = vCountItemPickZone + 1
                    End If
                Next

                If vCountItemPickZone = 0 Then
                    If vCountItemZoneOld = 0 Then
                        'Call ClearSaveData()
                        Exit Sub
                    End If
                End If

                vInstrAr = InStr(Me.LBLArCode.Text, "/")
                vLenAr = Len(Me.LBLArCode.Text)
                vARCode = vb6.Left(Me.LBLArCode.Text, vInstrAr - 1)

                If Me.LBLArCode.Text = "1/เงินสด" Then
                    vARCode = "99999"
                End If

                vInstrSale = InStr(Me.LBLSaleCode.Text, "/")
                vLenSale = Len(Me.LBLSaleCode.Text)
                vSaleCode = vb6.Left(Me.LBLSaleCode.Text, vInstrSale - 1)

                vMemberID = ""
                vTotalNetAmount = Me.LBLNetAmount.Text
                vBeforeTaxAmount = (vTotalNetAmount * 100) / 107
                vTaxAmount = vTotalNetAmount - vBeforeTaxAmount

                If vIsOpen = 0 Then
                    vQuery = "exec dbo.usp_np_insertdriveinslipno '" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vMemberID & "','" & vSaleCode & "','" & vRefNo & "'," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalNetAmount & ",'" & vUserID & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

                    For Each gr As GridViewRow In Me.GridView1.Rows
                        vItemCode = gr.Cells(0).Text 'Me.ListViewItem.Items(i).SubItems(4).Text
                        vItemName = gr.Cells(1).Text 'Me.ListViewItem.Items(i).SubItems(1).Text
                        vWHCode = gr.Cells(6).Text 'Me.ListViewItem.Items(i).SubItems(7).Text
                        vShelfCode = gr.Cells(7).Text 'Me.ListViewItem.Items(i).SubItems(8).Text
                        vQTY = gr.Cells(2).Text 'Me.ListViewItem.Items(i).SubItems(2).Text
                        vPrice = gr.Cells(4).Text 'Me.ListViewItem.Items(i).SubItems(5).Text
                        vAmount = gr.Cells(5).Text 'Me.ListViewItem.Items(i).SubItems(6).Text
                        vUnitCode = gr.Cells(3).Text 'Me.ListViewItem.Items(i).SubItems(3).Text
                        vBarCode = gr.Cells(8).Text 'Me.ListViewItem.Items(i).SubItems(9).Text
                        vShelfID = gr.Cells(9).Text 'Me.ListViewItem.Items(i).SubItems(10).Text
                        vZoneID = gr.Cells(10).Text 'Me.ListViewItem.Items(i).SubItems(11).Text
                        vLinePickZone = gr.Cells(11).Text 'Me.ListViewItem.Items(i).SubItems(12).Text
                        vDiscountWord = ""
                        vDiscountAmount = 0
                        vLineNumber = i

                        vQuery = "exec dbo.USP_NP_InsertDriveInSlipNoSub '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "','" & vLinePickZone & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vAmount & ",'" & vBarCode & "','" & vSaleCode & "'," & vLineNumber & " "
                        Dim vService2 As New WebReference.WebServiceCalc
                        Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                    Next

                    vQuery = "exec dbo.usp_np_updatenewdocno 29"
                    Dim vService3 As New WebReference.WebServiceCalc
                    Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

                    'MsgBox("บันทึกข้อมูลได้เลขที่เอกสาร " & vDocNo & " ", MsgBoxStyle.Information, "Send Information Message")


                    Dim vAnswer As Integer

                    'vAnswer = MsgBox("คุณต้องการส่ง เอกสารไปทำการ Check Out หรือไม่", MsgBoxStyle.YesNo, "Send Information Message")
                    If vAnswer = 6 Then
                        'Call SendCheckQue(vDocNo, vDocDate, vConnectZone)
                        'Call ClearSaveData()

                    Else
                        'Call ClearSaveData()
                    End If


                End If


                If vIsOpen = 1 Then
                    'Call BeforeSaveData()
                    vQuery = "exec dbo.USP_NP_CheckQueHoldBill1 '" & vDocNo & "','" & vARCode & "'"
                    Dim vService As New WebReference.WebServiceCalc
                    Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                    If ds.Tables(0).Rows.Count > 0 Then
                        vCheckIsConfirm = ds.Tables(0).Rows(0)("isconfirm").ToString()
                        vCheckHoldBillNo = ds.Tables(0).Rows(0)("holdbillno").ToString()
                    End If

                    If vCheckIsConfirm = 1 And vCheckHoldBillNo <> "" Then
                        'MsgBox("เลขที่เอกสารได้ทำการ ตรวจสอบสินค้าและทำเอกสารพักบิลเรียบร้อยแล้ว ไม่สามารถส่งคิวแก้ไขเอกสารได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                        'Call ClearSaveData()
                        'Call AfterSaveData()
                        'Me.TBRefNo.Focus()
                        '.TBRefNo.SelectAll()
                        Exit Sub
                    End If

                    vQuery = "exec dbo.USP_NP_InsertDriveInSlipNo '" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vMemberID & "','" & vSaleCode & "','" & vRefNo & "'," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalNetAmount & ",'" & vUserID & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

                    vCountItem = Me.GridView1.Rows.Count

                    For a = 0 To vCountItemOld - 1
                        vOldItem = "" 'vMemItemCodeOld(a)
                        vOldUnit = "" 'vMemUnitCodeOld(a)
                        vOldBar = "" 'vMemBarCodeOld(a)
                        vOldWH = "" 'vMemWHCodeOld(a)
                        vOldShelf = "" 'vMemShelfCodeOld(a)
                        vOldZone = "" 'vMemZoneIDOld(a)
                        vOldPick = "" 'vMemPickZoneOld(a)

                        For Each gr As GridViewRow In Me.GridView1.Rows
                            vCheckItemCode = gr.Cells(1).Text 'Me.ListViewItem.Items(b).SubItems(4).Text
                            vCheckUnitCode = gr.Cells(1).Text 'Me.ListViewItem.Items(b).SubItems(3).Text
                            vCheckBarCode = gr.Cells(1).Text 'Me.ListViewItem.Items(b).SubItems(9).Text
                            vCheckWHCode = gr.Cells(1).Text 'Me.ListViewItem.Items(b).SubItems(7).Text
                            vCheckShelfCode = gr.Cells(1).Text 'Me.ListViewItem.Items(b).SubItems(8).Text
                            vCheckZoneID = gr.Cells(1).Text 'Me.ListViewItem.Items(b).SubItems(11).Text
                            vCheckPickZone = gr.Cells(1).Text 'Me.ListViewItem.Items(b).SubItems(12).Text

                            If vCheckItemCode = vOldItem And vCheckUnitCode = vOldUnit And vCheckBarCode = vOldBar And vCheckWHCode = vOldWH And vCheckShelfCode = vOldShelf And vCheckZoneID = vOldZone And vCheckPickZone = vOldPick Then
                                vOld = 1
                                GoTo Line1
                            Else
                                vOld = 0
                            End If
                        Next
Line1:

                        If vOld = 0 Then
                            vItemCode = vOldItem
                            vWHCode = vOldWH
                            vShelfCode = vOldShelf
                            vUnitCode = vOldUnit
                            vBarCode = vOldBar
                            vZoneID = vOldZone
                            vLinePickZone = vOldPick

                            vQuery = "exec dbo.USP_NP_InsertDriveInSlipSubNo '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vZoneID & "','" & vLinePickZone & "','" & vUnitCode & "','" & vBarCode & "'," & vTotalNetAmount & " "
                            Dim vService2 As New WebReference.WebServiceCalc
                            Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                        End If
                    Next

                    For Each gr As GridViewRow In Me.GridView1.Rows
                        vItemCode = gr.Cells(0).Text 'Me.ListViewItem.Items(i).SubItems(4).Text
                        vItemName = gr.Cells(1).Text 'Me.ListViewItem.Items(i).SubItems(1).Text
                        vWHCode = gr.Cells(6).Text 'Me.ListViewItem.Items(i).SubItems(7).Text
                        vShelfCode = gr.Cells(7).Text 'Me.ListViewItem.Items(i).SubItems(8).Text
                        vQTY = gr.Cells(2).Text 'Me.ListViewItem.Items(i).SubItems(2).Text
                        vPrice = gr.Cells(4).Text 'Me.ListViewItem.Items(i).SubItems(5).Text
                        vAmount = gr.Cells(5).Text 'Me.ListViewItem.Items(i).SubItems(6).Text
                        vUnitCode = gr.Cells(3).Text 'Me.ListViewItem.Items(i).SubItems(3).Text
                        vBarCode = gr.Cells(8).Text 'Me.ListViewItem.Items(i).SubItems(9).Text
                        vShelfID = gr.Cells(9).Text 'Me.ListViewItem.Items(i).SubItems(10).Text
                        vZoneID = gr.Cells(10).Text 'Me.ListViewItem.Items(i).SubItems(11).Text
                        vLinePickZone = gr.Cells(11).Text 'Me.ListViewItem.Items(i).SubItems(12).Text
                        vDiscountWord = ""
                        vDiscountAmount = 0
                        vLineNumber = i

                        If vConnectZone = vLinePickZone Then
                            vQuery = "exec dbo.USP_NP_InsertDriveInSlipNo '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "','" & vLinePickZone & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vAmount & ",'" & vBarCode & "','" & vSaleCode & "'," & vLineNumber & " "
                            Dim vService3 As New WebReference.WebServiceCalc
                            Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)
                        End If
                    Next
                    'MsgBox("แก้ไขเลขที่เอกสาร " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")

                    Dim vAnswer As Integer

                    'vAnswer = MsgBox("คุณต้องการส่ง เอกสารไปทำการ Check Out หรือไม่", MsgBoxStyle.YesNo, "Send Information Message")
                    If vAnswer = 6 Then

                        Dim m As Integer
                        Dim vQueItemCode As String
                        Dim vQueItemName As String
                        Dim vQueUnit As String
                        Dim vQueQty As Double
                        Dim vQueID As Integer
                        Dim vQueArName As String
                        Dim vQueSaleName As String
                        Dim vQueZoneID As String
                        Dim vQueRefNo As String
                        Dim vIndex As Integer
                        Dim vQueDocNo As String
                        Dim vQueWHCode As String
                        Dim vQueShelfCode As String
                        Dim vQueShelfID As String
                        Dim vQueBarCode As String
                        Dim vQuePickZone As String


                        vQuery = "exec dbo.USP_NP_CheckQueDriveIn1 '" & vDocNo & "','" & vDocDate & "','" & vQueZone & "' "
                        Dim vService4 As New WebReference.WebServiceCalc
                        Dim ds4 As DataSet = vService4.vGetQueryAnlyzer(vQuery)

                        If ds4.Tables(0).Rows.Count > 0 Then

                            Me.LBLLicense.Text = Trim(ds4.Tables(0).Rows(0)("refno").ToString)
                            Me.LBLArCode.Text = Trim(ds4.Tables(0).Rows(0)("arcode").ToString) & "/" & Trim(ds4.Tables(0).Rows(0)("arname").ToString)

                            For m = 0 To ds4.Tables(0).Rows.Count - 1
                                vIndex = vIndex + 1
                                vQueItemCode = Trim(ds4.Tables(0).Rows(m)("itemcode").ToString)
                                vQueItemName = Trim(ds4.Tables(0).Rows(m)("itemname").ToString)
                                vQueUnit = Trim(ds4.Tables(0).Rows(m)("unitcode").ToString)
                                vQueQty = Trim(ds4.Tables(0).Rows(m)("qty").ToString)
                                vQueID = Trim(ds4.Tables(0).Rows(m)("queid").ToString)
                                vQueArName = Trim(ds4.Tables(0).Rows(m)("arname").ToString)
                                vQueSaleName = Trim(ds4.Tables(0).Rows(m)("salename").ToString)
                                vQueZoneID = Trim(ds4.Tables(0).Rows(m)("quezone").ToString)
                                vQueRefNo = Trim(ds4.Tables(0).Rows(m)("refno").ToString)
                                vQueDocNo = Trim(ds4.Tables(0).Rows(m)("docno").ToString)
                                vQueWHCode = Trim(ds4.Tables(0).Rows(m)("whcode").ToString)
                                vQueShelfCode = Trim(ds4.Tables(0).Rows(m)("shelfcode").ToString)
                                vQueShelfID = Trim(ds4.Tables(0).Rows(m)("shelfid").ToString)
                                vQueBarCode = Trim(ds4.Tables(0).Rows(m)("barcode").ToString)
                                vQuePickZone = Trim(ds4.Tables(0).Rows(m)("pickzone").ToString)

                                Dim listItem As New ListViewItem(vIndex)
                                listItem.SubItems.Add(vQueItemName)
                                listItem.SubItems.Add(Format(vQueQty, "##,##0.00"))
                                listItem.SubItems.Add(vQueUnit)
                                listItem.SubItems.Add(vQueID)
                                listItem.SubItems.Add(vQueZoneID)
                                listItem.SubItems.Add(vQueDocNo)
                                listItem.SubItems.Add(vQueItemCode)
                                listItem.SubItems.Add(vQueWHCode)
                                listItem.SubItems.Add(vQueShelfCode)
                                listItem.SubItems.Add(vQueBarCode)
                                listItem.SubItems.Add(vQuePickZone)
                                listItem.SubItems.Add(vQueShelfID)
                                'Me.ListViewItemLastSend.Items.Add(listItem)
                            Next

                        Else
                            'Call SendCheckQue(vDocNo, vDocDate, vConnectZone)
                            'Call ClearSaveData()
                        End If
                    Else
                        'Call ClearSaveData()
                    End If

                End If

            End If

        End If


ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Protected Sub TBBarCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged
        vBarCode = Me.TBBarCode.Text
        vLicense = Me.LBLLicense.Text
        vSaleCode = Me.LBLSaleCode.Text
        vPointID = Me.LBLPointID.Text

        If Me.LBLNetAmount.Text <> "" Then
            vNetAmount = Me.LBLNetAmount.Text
        End If

        vQuery = "exec bcnp.dbo.usp_MB_SearchBarcode '" & vBarCode & "' "
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
        If ds.Tables(0).Rows.Count > 0 Then
            Response.Redirect("PGPickupSearchItem.aspx?BarCodeID=" & vBarCode & "&License=" & vLicense & "&SaleCode=" & vSaleCode & "&NetAmount=" & vNetAmount & "&PointID=" & vPointID)
        Else
            Me.TBBarCode.Text = ""
            Dim vScript As String = "<SCRIPT language='javascript'>form1.TBBarCode.focus();</SCRIPT>"
            Page.RegisterStartupScript("focus", vScript)
        End If
    End Sub

    Private Function IsNumber(ByVal value As String) As Boolean
        Try
            Integer.Parse(value)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    'Private Sub displayreportPDF(ByVal crRpt As ReportDocument)
    '    Dim _export As New MemoryStream
    '    _export = crRpt.ExportToStream(ExportFormatType.PortableDocFormat)
    '    Response.Clear()
    '    Response.Buffer = True
    '    Response.ClearContent()
    '    Response.ClearHeaders()
    '    Response.ContentType = "application/pdf"
    '    Try
    '        Response.BinaryWrite(_export.ToArray())
    '        Response.End()
    '        Response.Flush()
    '        Response.Close()
    '    Catch ex As Exception
    '        Response.Write("Network Busy")
    '        Throw
    '    Finally
    '        crRpt.Close()
    '        crRpt.Dispose()
    '    End Try
    'End Sub

    'Protected Sub LinkButton2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkButton2.Click
    '    Dim crRpt As New ReportDocument
    '    crRpt.Load("ที่อยู่ Report")
    '    'crRpt.SetDatabaseLogon(UserId, Pwd, ServerName, DbName)
    '    cryRpt.DisplayGroupTree = False
    '    cryRpt.ReportSource = crRpt
    '    cryRpt.Visible = True
    '    displayreportPDF(crRpt)
    'End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vLicense As String
        Dim vUserID As String
        Dim vMemSaleName As String

        Call SaveData()
        Call ClearSession()

        vLicense = Me.LBLLicense.Text
        vMemSaleName = Me.LBLSaleCode.Text
        vPointID = Me.LBLPointID.Text

        Response.Redirect("PickupLicense.aspx?UserID=" & vUserID & "&SaleCode=" & vMemSaleName & "&PointID=" & vPointID)
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNPrevoius.Click
        Dim vMemSaleName As String
        Dim vUserID As String

        vUserID = Me.LBLUserID.Text
        vPointID = Me.LBLPointID.Text
        vMemSaleName = Me.LBLSaleCode.Text

        Response.Redirect("PickupLicense.aspx?UserID=" & vUserID & "&SaleCode=" & vMemSaleName & "&PointID=" & vPointID)

    End Sub
End Class
