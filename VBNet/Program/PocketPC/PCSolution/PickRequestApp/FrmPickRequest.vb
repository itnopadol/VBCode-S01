Imports System.data
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms

Public Class FrmPickRequest
    Dim ds As DataSet
    Dim vQuery As String

    Dim vUserCode As String
    Dim vPassWord As String

    Dim vMemSaleName As String

    Private Sub FrmPickRequest_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub TBBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
        Dim vBarCode As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vPrice As Double
        Dim vRate As Integer
        Dim vUnitCode As String
        Dim i As Integer
        Dim vWHCode As String
        Dim vStore As String
        Dim vStkUnit As String
        Dim vStkQTY As Double
        Dim vReserveQTY As Double
        Dim vDefWHCode As String
        Dim vDefShelfCode As String
        Dim vShelfID As String


        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBRefNo.Focus()
        End If

        If e.KeyCode = Keys.Down Then
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(0).Selected = True
                Me.ListViewItem.Items(0).Focused = True
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If


        'If e.KeyCode = Keys.Enter Then
        If e.KeyCode = 0 Then
            If Me.TBBarCode.Text <> "" Then
                vBarCode = Me.TBBarCode.Text
            Else
                Me.TBBarCode.Focus()
            End If

            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetDataBarCode(vBarCode)
            Me.ListViewStock.Items.Clear()
            Me.ListViewWareHouse.Items.Clear()

            If ds.Tables(0).Rows.Count > 0 Then
                vItemCode = ds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = ds.Tables(0).Rows(0)("itemname").ToString
                vPrice = ds.Tables(0).Rows(0)("price").ToString
                vRate = ds.Tables(0).Rows(0)("rate").ToString
                vUnitCode = ds.Tables(0).Rows(0)("unitcode").ToString
                vDefWHCode = ds.Tables(0).Rows(0)("defsalewhcode").ToString
                vDefShelfCode = ds.Tables(0).Rows(0)("defsaleshelf").ToString
                vShelfID = ds.Tables(0).Rows(0)("shelfid").ToString

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    vWHCode = ds.Tables(0).Rows(i)("whcode").ToString
                    vStore = ds.Tables(0).Rows(i)("shelfcode").ToString
                    vStkUnit = ds.Tables(0).Rows(i)("stkunitcode").ToString
                    vStkQTY = ds.Tables(0).Rows(i)("stock").ToString

                    Dim listItem As New ListViewItem(vWHCode)
                    listItem.SubItems.Add(vStore)
                    listItem.SubItems.Add(Format(vStkQTY, "##,##0.00"))
                    listItem.SubItems.Add(vStkUnit)
                    Me.ListViewStock.Items.Add(listItem)
                Next

                Dim n As Integer
                Dim vCheckItemCode As String
                Dim vCheckQTY As Double

                If Me.ListViewItem.Items.Count > 0 Then
                    For n = 0 To Me.ListViewItem.Items.Count - 1
                        vCheckItemCode = Me.ListViewItem.Items(n).SubItems(4).Text
                        vCheckQTY = Me.ListViewItem.Items(n).SubItems(2).Text
                        If vItemCode = vCheckItemCode Then
                            Me.TBQTY.Text = Format(vCheckQTY, "##,##0.00")
                            GoTo Line1
                        End If
                    Next
                End If

                Dim vRemainInQty As Double
                Dim vRemainOutQty As Double
                Dim vGetWHCode As String
                Dim m As Integer

                vQuery = "exec dbo.USP_MB_SearchItemWareHouse '" & vBarCode & "'"
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)

                Me.ListViewWareHouse.Items.Clear()
                If ds1.Tables(0).Rows.Count > 0 Then
                    For m = 0 To ds1.Tables(0).Rows.Count - 1
                        vGetWHCode = ds1.Tables(0).Rows(m)("whcode").ToString
                        vReserveQTY = ds1.Tables(0).Rows(m)("reserveqty").ToString
                        vRemainInQty = ds1.Tables(0).Rows(m)("remaininqty").ToString
                        vRemainOutQty = ds1.Tables(0).Rows(m)("remainoutqty").ToString

                        Dim listItem As New ListViewItem(vGetWHCode)
                        listItem.SubItems.Add(Format(vReserveQTY, "##,##0.00"))
                        listItem.SubItems.Add(Format(vRemainInQty, "##,##0.00"))
                        listItem.SubItems.Add(Format(vRemainOutQty, "##,##0.00"))
                        Me.ListViewWareHouse.Items.Add(listItem)

                    Next
                End If

Line1:
                Me.TBQTY.Focus()
                Me.TBQTY.SelectAll()
            Else
                Me.TBBarCode.Focus()
                Me.TBQTY.SelectAll()
            End If

            Me.TBItem.Text = vItemCode
            Me.TBItemName.Text = vItemName
            Me.TBPrice.Text = Format(vPrice, "##,##0.00")
            Me.TBRate.Text = Format(vRate, "##,##0.00")
            Me.TBUnit.Text = vUnitCode
            Me.TBWHCode.Text = vDefWHCode
            Me.TBShelfCode.Text = vDefShelfCode
            Me.TBMemBarCode.Text = vBarCode
            Me.TBShelfID.Text = vShelfID

        End If

        If e.KeyCode = Keys.Back Then
            Me.TBItem.Text = ""
            Me.TBItemName.Text = ""
            Me.TBPrice.Text = ""
            Me.TBUnit.Text = ""
            Me.TBWHCode.Text = ""
            Me.TBShelfCode.Text = ""
            Me.TBShelfID.Text = ""
            Me.TBQTY.Text = ""
            Me.TBRate.Text = ""
            Me.TBMemBarCode.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.ListViewWareHouse.Items.Clear()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If
    End Sub

    Private Sub TBBarCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged
        If Me.TBBarCode.Text <> "" Then
            Me.PNItemDetails.Visible = True
            Me.PNItemDetails.BringToFront()
            Me.BTNSave.Visible = False
        Else
            Me.PNItemDetails.Visible = False
            Me.PNDriveIn.Visible = True
            Me.PNDriveIn.BringToFront()
            Me.BTNSave.Visible = True
        End If
    End Sub

    Private Sub BTNAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim vItemCode As String
        'Dim vItemName As String
        'Dim vWHCode As String
        'Dim vShelfCode As String
        'Dim vQTY As Double
        'Dim vPrice As Double
        'Dim vAmount As Double
        'Dim vUnitCode As String
        'Dim vIndex As Integer

        'If Me.ListViewStock.Items.Count > 0 And Me.TBItem.Text <> "" And Me.TBQTY.Text <> "" Then
        '    vItemCode = Me.TBItem.Text
        '    vItemName = Me.TBItemName.Text
        '    vWHCode = Me.TBWHCode.Text
        '    vShelfCode = Me.TBShelfCode.Text
        '    vQTY = Me.TBQTY.Text
        '    vPrice = Me.TBPrice.Text
        '    vAmount = vQTY * vPrice
        '    vUnitCode = Me.TBUnit.Text
        '    vIndex = Me.ListViewItem.Items.Count + 1

        '    Dim listItem As New ListViewItem(vIndex)
        '    listItem.SubItems.Add(vItemName)
        '    listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
        '    listItem.SubItems.Add(vUnitCode)
        '    listItem.SubItems.Add(vItemCode)
        '    listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
        '    listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
        '    listItem.SubItems.Add(vWHCode)
        '    listItem.SubItems.Add(vShelfCode)
        '    Me.ListViewItem.Items.Add(listItem)

        '    Call CalcItemAmount()

        '    Me.TBItem.Text = ""
        '    Me.TBItemName.Text = ""
        '    Me.TBPrice.Text = ""
        '    Me.TBReserve.Text = ""
        '    Me.TBUnit.Text = ""
        '    Me.TBWHCode.Text = ""
        '    Me.TBShelfCode.Text = ""
        '    Me.TBQTY.Text = ""
        '    Me.ListViewStock.Items.Clear()
        '    Me.PNItemDetails.Visible = False
        '    Me.TBBarCode.Focus()
        'End If
    End Sub

    Private Sub CalcItemAmount()
        Dim i As Integer
        Dim vAmount As Double
        Dim vSumAmount As Double

        If Me.ListViewItem.Items.Count > 0 Then
            For i = 0 To Me.ListViewItem.Items.Count - 1
                vAmount = Me.ListViewItem.Items(i).SubItems(6).Text
                vSumAmount = vSumAmount + vAmount
            Next
            Me.TBItemAmount.Text = Format(vSumAmount, "##,##0.00")
        Else
            Me.TBItemAmount.Text = Format(0, "##,##0.00")
        End If
    End Sub


    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vHeader As String
        Dim vNumber As Integer
        Dim vDocNumber As String

        Dim vDocNo As String
        Dim vDocDate As String
        Dim vARCode As String
        Dim vMemberID As String
        Dim vSaleCode As String
        Dim vBeforeTaxAmount As Double
        Dim vTaxAmount As Double
        Dim vRefNo As String
        Dim vPickZone As String
        Dim vTotalNetAmount As Double
        Dim vMyDescription As String
        Dim vIsConditionSend As Integer
        Dim vReqTime As String

        Dim vItemCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vDiscountWord As String
        Dim vDiscountAmount As Double
        Dim vNetAmount As Double
        Dim vUnitCode As String
        Dim vUserID As String
        Dim i As Integer
        Dim vBarCode As String
        Dim vLineNumber As Integer

        Dim vInstrSale As Integer
        Dim vLenSale As Integer

        Dim vAnswer As Integer


        If Me.ListViewItem.Items.Count > 0 And Me.TBItemAmount.Text <> "" Then
            vUserID = Me.TBUserID.Text

            If Me.TBRefNo.Text = "" Then
                Me.TBRefNo.Text = "N/A"
            End If

            If Me.TBDocNo.Text = "" Then
                vQuery = "exec dbo.usp_np_searchnewdocno 26"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count > 0 Then
                    vHeader = Trim(ds.Tables(0).Rows(0)("header").ToString)
                    vNumber = ds.Tables(0).Rows(0)("autonumber").ToString
                    vDocNumber = Trim(ds.Tables(0).Rows(0)("docnumber").ToString)
                End If

                vDocNo = Trim(vDocNumber & vHeader & "-" & Format(vNumber, "0000"))
            Else
                vDocNo = Me.TBDocNo.Text
            End If

            If vDocNo <> "" Then

                Call BeforeSave()

                Me.LBLSaveMessage.Text = "กำลังบันทึกแก้ไขข้อมูล กรุณารอสักครู่"

                vDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year

                vQuery = "exec dbo.USP_NP_CheckDocDate"
                Dim vService7 As New WebReference.WebServiceCalc
                Dim ds7 As DataSet = vService7.vGetQueryAnlyzer(vQuery)
                If ds7.Tables(0).Rows.Count > 0 Then
                    vDocDate = ds7.Tables(0).Rows(0)("vdocdate").ToString
                End If

                vRefNo = Me.TBRefNo.Text

                If Me.RDZone1.Checked = True Then
                    vPickZone = "01"
                ElseIf Me.RDZone2.Checked = True Then
                    vPickZone = "02"
                ElseIf Me.RDZone3.Checked = True Then
                    vPickZone = "03"
                End If

                vConnectZone = vPickZone

                If Me.TBARCode.Text = "1" Then
                    vARCode = "99999"
                Else
                    vARCode = Me.TBARCode.Text
                End If

                vInstrSale = InStr(Me.TBSaleCode.Text, "/")
                If vInstrSale = 0 Then
                    MsgBox("กรุณากรอกรหัสพนักงานให้ถูกต้องตามโปรแกรมที่ระบุไว้ คีย์รหัสพนักงานแล้วกด Enter อีกครั้ง", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBSaleCode.Focus()
                    Me.TBSaleCode.SelectAll()
                    Exit Sub
                End If
                vLenSale = Len(Me.TBSaleCode.Text)
                vSaleCode = vb6.Left(Me.TBSaleCode.Text, vInstrSale - 1)

                vMemberID = Me.TBMemberID.Text
                vTotalNetAmount = Me.TBItemAmount.Text
                vBeforeTaxAmount = (vTotalNetAmount * 100) / 107
                vTaxAmount = vTotalNetAmount - vBeforeTaxAmount
                vMyDescription = ""
                vReqTime = vb6.DateAdd(DateInterval.Minute, 15, Now)

                'vQuery = "exec dbo.usp_np_insertdriveinslip '" & vDocNo & "','" & vDocDate & "'," & vID & ",'" & vRefNo & "','" & vPickZone & "'," & vTotalNetAmount & ",'" & vUserID & "'"
                vQuery = "exec dbo.usp_np_insertpickingrequestmaster'" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalNetAmount & "," & vIsConditionSend & ",'" & vReqTime & "','" & vMyDescription & "','" & vPickZone & "','" & vUserID & "'"
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

                For i = 0 To Me.ListViewItem.Items.Count - 1
                    vItemCode = Me.ListViewItem.Items(i).SubItems(4).Text
                    vWHCode = Me.ListViewItem.Items(i).SubItems(7).Text
                    vShelfCode = Me.ListViewItem.Items(i).SubItems(8).Text
                    vQTY = Me.ListViewItem.Items(i).SubItems(2).Text
                    vPrice = Me.ListViewItem.Items(i).SubItems(5).Text
                    vNetAmount = Me.ListViewItem.Items(i).SubItems(6).Text
                    vUnitCode = Me.ListViewItem.Items(i).SubItems(3).Text
                    vBarCode = Me.ListViewItem.Items(i).SubItems(9).Text
                    vDiscountWord = ""
                    vDiscountAmount = 0
                    vShelfID = Me.ListViewItem.Items(i).SubItems(10).Text
                    vLineNumber = i

                    'vQuery = "exec dbo.usp_np_insertdriveinslipsub '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "'," & vQTY & "," & vQTY & ",'" & vUnitCode & "'," & vPrice & "," & vAmount & ",'" & vBarCode & "'," & vLineNumber & " "
                    vQuery = "exec dbo.usp_np_insertpickingrequestsub '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vNetAmount & ",'" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vPickZone & "','" & vBarCode & "'," & vLineNumber & " "
                    Dim vService2 As New WebReference.WebServiceCalc
                    Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                Next

                If Me.TBDocNo.Text = "" Then
                    vQuery = "exec dbo.usp_np_updatenewdocno 26"
                    Dim vService3 As New WebReference.WebServiceCalc
                    Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

                    MsgBox("ได้เลขที่เอกสารเลขที่" & vDocNo & " ", MsgBoxStyle.Information, "Send Information Message")
                    Me.LBLSaveMessage.Text = ""

                    vAnswer = MsgBox("คุณต้องการส่งคิวจัดสินค้าหรือไม่?", MsgBoxStyle.YesNo, "Send Question Message")
                    If vAnswer = 6 Then
                        Call SendCheckQue(vDocNo, vDocDate, vPickZone)
                    End If
                Else
                    MsgBox("แก้ไขเลขที่เอกสาร " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")
                    Me.LBLSaveMessage.Text = ""

                    vAnswer = MsgBox("คุณต้องการส่งคิวจัดสินค้าหรือไม่?", MsgBoxStyle.YesNo, "Send Question Message")
                    If vAnswer = 6 Then
                        Call SendCheckQue(vDocNo, vDocDate, vPickZone)
                    End If
                End If

                Call AfterSave()

                Me.ListViewItem.Items.Clear()
                Me.TBRefNo.Text = ""
                Me.TBItemAmount.Text = ""
                Me.TBDocNo.Text = ""
                Me.TBBarCode.Text = ""
                Call CallIDNumber()
                Me.TBRefNo.Focus()
            End If
        End If
    End Sub

    Public Sub BeforeSave()
        Me.BTNBack.Enabled = False
        Me.BTNClearPickUp.Enabled = False
        Me.BTNSave.Enabled = False
        Me.BTNSearch.Enabled = False
        Me.BTNClosePickup.Enabled = False
    End Sub

    Public Sub AfterSave()
        Me.BTNBack.Enabled = True
        Me.BTNClearPickUp.Enabled = True
        Me.BTNSave.Enabled = True
        Me.BTNSearch.Enabled = True
        Me.BTNClosePickup.Enabled = True
    End Sub

    Private Sub SendCheckQue(ByVal vDocNo As String, ByVal vDocDate As String, ByVal vPickZone As String)
        Dim vSendCountID As Integer
        Dim vLastCountID As Integer
        Dim vType As Integer
        Dim i As Integer
        Dim vGroupZone(4) As String
        Dim n As Integer
        Dim vPrinterName As String

        On Error GoTo ErrDescription

        If Me.ListViewItem.Items.Count > 0 Then
            vType = 1
            vQuery = "exec dbo.USP_NP_CheckQuePickCenter1 '" & vDocNo & "','" & vDocDate & "' "
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            If ds.Tables(0).Rows.Count > 0 Then
                vLastCountID = Trim(ds.Tables(0).Rows(0)("max1").ToString)
            End If

            vSendCountID = vLastCountID + 1

            vQuery = "exec dbo.USP_NP_SearchGroupPicking1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'"
            Dim vService1 As New WebReference.WebServiceCalc
            Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)

            If ds1.Tables(0).Rows.Count > 0 Then
                n = ds1.Tables(0).Rows.Count
                For i = 0 To ds1.Tables(0).Rows.Count - 1
                    vGroupZone(i) = Trim(ds1.Tables(0).Rows(i)("zoneid").ToString)
                Next
            End If

            For i = 0 To n - 1
                If vGroupZone(i) = "A" Then
                    Call InsertQueZoneA(vDocNo, vDocDate, vSendCountID, vType)
                    vPrinterName = "เอกสารพิมพ์ออกที่จุด A"
                End If

                If vGroupZone(i) = "B" Then
                    Call InsertQueZoneA(vDocNo, vDocDate, vSendCountID, vType)
                    vPrinterName = "เอกสารพิมพ์ออกที่จุด B"
                End If
            Next

            vQuery = "exec dbo.USP_NP_UpdateSendQueuePicking1 " & vType & ",'" & vDocNo & "'"
            Dim vService2 As New WebReference.WebServiceCalc
            Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

            vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto1 1,'" & vDocNo & "','" & vPickZone & "','" & vUserName & "'"
            Dim vService3 As New WebReference.WebServiceCalc
            Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

            MsgBox("ส่งรายการสินค้าไปทำการ CheckOut เรียบร้อยแล้ว  " & vPrinterName & " ", MsgBoxStyle.Information, "Send Information Message")
            Me.TBRefNo.Focus()
            Me.TBRefNo.SelectAll()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub InsertQueZoneA(ByVal vDocNo As String, ByVal vDocDate As String, ByVal vTimeID As Integer, ByVal vType As Integer)
        Dim vQueID As Integer
        Dim vQueDocDate As String
        Dim vARCode As String
        Dim vSaleCode As String
        Dim vRefNo As String
        Dim vMemberID As String
        Dim vSourceID As Integer
        Dim vQueZone As String
        Dim vQueReqTime As String
        Dim vIsConditionSend As Integer
        Dim vPickZone As String

        Dim vItemCode As String
        Dim vItemName As String
        Dim vQTY As Double
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vBarCode As String
        Dim vLineNumber As Integer

        Dim vInstrSale As Integer
        Dim vLenSale As Integer
        Dim vAddTime1 As Date
        Dim vAddTime As String

        Dim i As Integer

        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_NP_SearchNewDocNo 31"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vQueID = Trim(ds.Tables(0).Rows(0)("autonumber").ToString)
        End If

        vPickZone = "01"
        vQueZone = "B"
        vQueDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        vARCode = Me.TBARCode.Text
        vInstrSale = InStr(Me.TBSaleCode.Text, "/")
        vLenSale = Len(Me.TBSaleCode.Text)
        vSaleCode = vb6.Left(Me.TBSaleCode.Text, vInstrSale - 1)
        vRefNo = Me.TBRefNo.Text
        vMemberID = Me.TBMemberID.Text
        vSourceID = vType
        vAddTime1 = vb6.DateAdd(DateInterval.Minute, 15, Now)
        vAddTime = vAddTime1.Hour & ":" & vAddTime1.Minute
        vQueReqTime = vAddTime
        vIsConditionSend = Me.CMBConditionSend.SelectedIndex

        vQuery = "exec dbo.USP_NP_InsertQuePickCenterMasterPickRequest '" & vQueID & "','" & vQueDocDate & "','" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "','" & vSourceID & "','" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "'"
        Dim vService1 As New WebReference.WebServiceCalc
        Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

        vQuery = "exec dbo.USP_NP_SearchReqPickingItemZone1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'," & vTimeID & ""
        Dim vService2 As New WebReference.WebServiceCalc
        Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)

        If ds2.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds2.Tables(0).Rows.Count - 1
                vItemCode = Trim(ds2.Tables(0).Rows(i)("itemcode").ToString)
                vItemName = Trim(ds2.Tables(0).Rows(i)("itemname").ToString)
                vQTY = Trim(ds2.Tables(0).Rows(i)("qty").ToString)
                vUnitCode = Trim(ds2.Tables(0).Rows(i)("unitcode").ToString)
                vWHCode = Trim(ds2.Tables(0).Rows(i)("whcode").ToString)
                vShelfCode = Trim(ds2.Tables(0).Rows(i)("shelfcode").ToString)
                vShelfID = Trim(ds2.Tables(0).Rows(i)("shelfid").ToString)
                vBarCode = Trim(ds2.Tables(0).Rows(i)("barcode").ToString)
                vLineNumber = i

                vQuery = "exec dbo.USP_NP_InsertQuePickCenterPickRequestSub '" & vQueID & "','" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vQueZone & "','" & vPickZone & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & ""
                Dim vService3 As New WebReference.WebServiceCalc
                Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)
            Next
        End If

        vQuery = "exec dbo.USP_NP_UpdateNewDocNo 31"
        Dim vService4 As New WebReference.WebServiceCalc
        Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

    End Sub

    Private Sub InsertQueZoneB(ByVal vDocNo As String, ByVal vDocDate As String, ByVal vTimeID As Integer, ByVal vType As Integer)
        Dim vQueID As Integer
        Dim vQueDocDate As String
        Dim vARCode As String
        Dim vSaleCode As String
        Dim vRefNo As String
        Dim vMemberID As String
        Dim vSourceID As Integer
        Dim vQueZone As String
        Dim vQueReqTime As String
        Dim vIsConditionSend As Integer
        Dim vPickZone As String

        Dim vItemCode As String
        Dim vItemName As String
        Dim vQTY As Double
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vBarCode As String
        Dim vLineNumber As Integer

        Dim vInstrSale As Integer
        Dim vLenSale As Integer
        Dim vAddTime1 As Date
        Dim vAddTime As String

        Dim i As Integer

        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_NP_SearchNewDocNo 31"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vQueID = Trim(ds.Tables(0).Rows(0)("autonumber").ToString)
        End If

        vPickZone = "01"
        vQueZone = "B"
        vQueDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        vARCode = Me.TBARCode.Text
        vInstrSale = InStr(Me.TBSaleCode.Text, "/")
        vLenSale = Len(Me.TBSaleCode.Text)
        vSaleCode = vb6.Left(Me.TBSaleCode.Text, vInstrSale - 1)
        vRefNo = Me.TBRefNo.Text
        vMemberID = Me.TBMemberID.Text
        vSourceID = vType
        vAddTime1 = vb6.DateAdd(DateInterval.Minute, 15, Now)
        vAddTime = vAddTime1.Hour & ":" & vAddTime1.Minute
        vQueReqTime = vAddTime
        vIsConditionSend = 0

        vQuery = "exec dbo.USP_NP_InsertQuePickCenterMasterDriveIn1 '" & vQueID & "','" & vQueDocDate & "','" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "','" & vSourceID & "','" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "'"
        Dim vService1 As New WebReference.WebServiceCalc
        Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

        vQuery = "exec dbo.USP_NP_SearchReqPickingItemZone1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'," & vTimeID & ""
        Dim vService2 As New WebReference.WebServiceCalc
        Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)

        If ds2.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds2.Tables(0).Rows.Count - 1
                vItemCode = Trim(ds2.Tables(0).Rows(i)("itemcode").ToString)
                vItemName = Trim(ds2.Tables(0).Rows(i)("itemname").ToString)
                vQTY = Trim(ds2.Tables(0).Rows(i)("qty").ToString)
                vUnitCode = Trim(ds2.Tables(0).Rows(i)("unitcode").ToString)
                vWHCode = Trim(ds2.Tables(0).Rows(i)("whcode").ToString)
                vShelfCode = Trim(ds2.Tables(0).Rows(i)("shelfcode").ToString)
                vShelfID = Trim(ds2.Tables(0).Rows(i)("shelfid").ToString)
                vBarCode = Trim(ds2.Tables(0).Rows(i)("barcode").ToString)
                vLineNumber = i

                vQuery = "exec dbo.USP_NP_InsertQuePickCenterDriveInSub1 '" & vQueID & "','" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vQueZone & "','" & vPickZone & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & ""
                Dim vService3 As New WebReference.WebServiceCalc
                Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)
            Next
        End If

        vQuery = "exec dbo.USP_NP_UpdateNewDocNo 31"
        Dim vService4 As New WebReference.WebServiceCalc
        Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

    End Sub

    Private Sub SavePickRequest()
        Dim vHeader As String
        Dim vNumber As Integer
        Dim vDocNumber As String

        Dim vDocNo As String
        Dim vDocDate As String
        Dim vARCode As String
        Dim vMemberID As String
        Dim vSaleCode As String
        Dim vBeforeTaxAmount As Double
        Dim vTaxAmount As Double
        Dim vRefNo As String
        Dim vPickZone As String
        Dim vTotalNetAmount As Double
        Dim vMyDescription As String
        Dim vIsConditionSend As Integer
        Dim vReqTime As String

        Dim vItemCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vDiscountWord As String
        Dim vDiscountAmount As Double
        Dim vNetAmount As Double
        Dim vUnitCode As String
        Dim vUserID As String
        Dim i As Integer
        Dim vBarCode As String
        Dim vLineNumber As Integer

        Dim vInstrSale As Integer
        Dim vLenSale As Integer

        Dim vAnswer As Integer


        If Me.ListViewItem.Items.Count > 0 And Me.TBItemAmount.Text <> "" Then
            vUserID = Me.TBUserID.Text

            If Me.TBRefNo.Text = "" Then
                Me.TBRefNo.Text = "N/A"
            End If

            If Me.TBDocNo.Text = "" Then
                vQuery = "exec dbo.usp_np_searchnewdocno 26"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count > 0 Then
                    vHeader = Trim(ds.Tables(0).Rows(0)("header").ToString)
                    vNumber = ds.Tables(0).Rows(0)("autonumber").ToString
                    vDocNumber = Trim(ds.Tables(0).Rows(0)("docnumber").ToString)
                End If

                vDocNo = Trim(vDocNumber & vHeader & "-" & Format(vNumber, "0000"))
            Else
                vDocNo = Me.TBDocNo.Text
            End If

            If vDocNo <> "" Then

                Call BeforeSave()

                Me.LBLSaveMessage.Text = "กำลังบันทึกแก้ไขข้อมูล กรุณารอสักครู่"

                vDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year

                vQuery = "exec dbo.USP_NP_CheckDocDate"
                Dim vService7 As New WebReference.WebServiceCalc
                Dim ds7 As DataSet = vService7.vGetQueryAnlyzer(vQuery)
                If ds7.Tables(0).Rows.Count > 0 Then
                    vDocDate = ds7.Tables(0).Rows(0)("vdocdate").ToString
                End If

                vRefNo = Me.TBRefNo.Text

                If Me.RDZone1.Checked = True Then
                    vPickZone = "01"
                ElseIf Me.RDZone2.Checked = True Then
                    vPickZone = "02"
                ElseIf Me.RDZone3.Checked = True Then
                    vPickZone = "03"
                End If

                vConnectZone = vPickZone

                If Me.TBARCode.Text = "1" Then
                    vARCode = "99999"
                Else
                    vARCode = Me.TBARCode.Text
                End If

                vInstrSale = InStr(Me.TBSaleCode.Text, "/")
                If vInstrSale = 0 Then
                    MsgBox("กรุณากรอกรหัสพนักงานให้ถูกต้องตามโปรแกรมที่ระบุไว้ คีย์รหัสพนักงานแล้วกด Enter อีกครั้ง", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBSaleCode.Focus()
                    Me.TBSaleCode.SelectAll()
                    Exit Sub
                End If
                vLenSale = Len(Me.TBSaleCode.Text)
                vSaleCode = vb6.Left(Me.TBSaleCode.Text, vInstrSale - 1)

                vMemberID = Me.TBMemberID.Text
                vTotalNetAmount = Me.TBItemAmount.Text
                vBeforeTaxAmount = (vTotalNetAmount * 100) / 107
                vTaxAmount = vTotalNetAmount - vBeforeTaxAmount
                vMyDescription = ""
                vReqTime = vb6.DateAdd(DateInterval.Minute, 15, Now)

                'vQuery = "exec dbo.usp_np_insertdriveinslip '" & vDocNo & "','" & vDocDate & "'," & vID & ",'" & vRefNo & "','" & vPickZone & "'," & vTotalNetAmount & ",'" & vUserID & "'"
                vQuery = "exec dbo.usp_np_insertpickingrequestmaster'" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalNetAmount & "," & vIsConditionSend & ",'" & vReqTime & "','" & vMyDescription & "','" & vPickZone & "','" & vUserID & "'"
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

                For i = 0 To Me.ListViewItem.Items.Count - 1
                    vItemCode = Me.ListViewItem.Items(i).SubItems(4).Text
                    vWHCode = Me.ListViewItem.Items(i).SubItems(7).Text
                    vShelfCode = Me.ListViewItem.Items(i).SubItems(8).Text
                    vQTY = Me.ListViewItem.Items(i).SubItems(2).Text
                    vPrice = Me.ListViewItem.Items(i).SubItems(5).Text
                    vNetAmount = Me.ListViewItem.Items(i).SubItems(6).Text
                    vUnitCode = Me.ListViewItem.Items(i).SubItems(3).Text
                    vBarCode = Me.ListViewItem.Items(i).SubItems(9).Text
                    vDiscountWord = ""
                    vDiscountAmount = 0
                    vShelfID = Me.ListViewItem.Items(i).SubItems(10).Text
                    vLineNumber = i

                    'vQuery = "exec dbo.usp_np_insertdriveinslipsub '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "'," & vQTY & "," & vQTY & ",'" & vUnitCode & "'," & vPrice & "," & vAmount & ",'" & vBarCode & "'," & vLineNumber & " "
                    vQuery = "exec dbo.usp_np_insertpickingrequestsub '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vNetAmount & ",'" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vPickZone & "','" & vBarCode & "'," & vLineNumber & " "
                    Dim vService2 As New WebReference.WebServiceCalc
                    Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                Next

                If Me.TBDocNo.Text = "" Then
                    vQuery = "exec dbo.usp_np_updatenewdocno 26"
                    Dim vService3 As New WebReference.WebServiceCalc
                    Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

                    MsgBox("ได้เลขที่เอกสารเลขที่" & vDocNo & " ", MsgBoxStyle.Information, "Send Information Message")
                    Me.LBLSaveMessage.Text = ""

                    vAnswer = MsgBox("คุณต้องการส่งคิวจัดสินค้าหรือไม่?", MsgBoxStyle.YesNo, "Send Question Message")
                    If vAnswer = 6 Then
                        Call SendCheckQue(vDocNo, vDocDate, vPickZone)
                    End If
                Else
                    MsgBox("แก้ไขเลขที่เอกสาร " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")
                    Me.LBLSaveMessage.Text = ""

                    vAnswer = MsgBox("คุณต้องการส่งคิวจัดสินค้าหรือไม่?", MsgBoxStyle.YesNo, "Send Question Message")
                    If vAnswer = 6 Then
                        Call SendCheckQue(vDocNo, vDocDate, vPickZone)
                    End If
                End If

                Call AfterSave()

                Me.ListViewItem.Items.Clear()
                Me.TBRefNo.Text = ""
                Me.TBItemAmount.Text = ""
                Me.TBDocNo.Text = ""
                Me.TBBarCode.Text = ""
                Call CallIDNumber()
                Me.TBRefNo.Focus()
            End If
        End If
    End Sub
    Private Sub frmProgram1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PNDriveIn.Visible = False

        Me.PNLogIn.Visible = True
        Me.PNLogIn.BringToFront()
        Me.RDZone1.Focus()
    End Sub

    Private Sub CallIDNumber()
        Me.TBARCode.Text = "99999"
    End Sub

    Private Sub TBRefNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBRefNo.KeyDown

        'If e.KeyCode = Keys.Enter Then
        If e.KeyCode = 0 Then
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBBarCode.Focus()
        End If

    End Sub

    Private Sub TBQTY_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBQTY.KeyDown
        Dim vBarCode As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vIndex As Integer
        Dim vCheckExist As Integer

        Dim vCheckShelf As String
        Dim vCheckUnit As String
        Dim vCheckWHCode As String
        Dim v As Integer
        Dim vShelfQTY As Double
        Dim vShelfUnit As String
        Dim vListWHCode As String
        Dim vListShelf As String
        Dim vListUnit As String
        Dim vRate As Integer
        Dim vTotalQTY As Double

        Dim vAnswer As Integer

        If e.KeyCode = Keys.Escape Then
            Me.PNItemDetails.Visible = False
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBBarCode.Focus()
        End If


        'If e.KeyCode = Keys.Enter Then
        If e.KeyCode = 0 Then
            If Me.ListViewStock.Items.Count > 0 And Me.TBItem.Text <> "" Then
                vCheckWHCode = Me.TBWHCode.Text
                vCheckShelf = Me.TBShelfCode.Text
                vCheckUnit = Me.TBUnit.Text
                If Me.ListViewStock.Items.Count > 0 Then
                    For v = 0 To Me.ListViewStock.Items.Count - 1
                        vListWHCode = Me.ListViewStock.Items(v).Text
                        vListShelf = Me.ListViewStock.Items(v).SubItems(1).Text
                        vListUnit = Me.ListViewStock.Items(v).SubItems(3).Text
                        If vCheckWHCode = vListWHCode And vCheckShelf = vListShelf And vCheckUnit = vListUnit Then
                            vShelfQTY = Me.ListViewStock.Items(v).SubItems(2).Text
                            vShelfUnit = Me.ListViewStock.Items(v).SubItems(3).Text
                            GoTo Line1
                        End If
                    Next
                End If

Line1:
                vCheckExist = 0
                vBarCode = Me.TBMemBarCode.Text
                vItemCode = Me.TBItem.Text
                vItemName = Me.TBItemName.Text
                vWHCode = Me.TBWHCode.Text
                vShelfCode = Me.TBShelfCode.Text
                vUnitCode = Me.TBUnit.Text
                vRate = Me.TBRate.Text
                vShelfID = Me.TBShelfID.Text

                If Me.TBQTY.Text <> "" Then
                    vQTY = Me.TBQTY.Text
                End If

                If vShelfUnit <> vUnitCode Then
                    vTotalQTY = vShelfQTY / vRate
                    If vQTY > vTotalQTY Then
                        vAnswer = MsgBox("สินค้ารหัส " & vItemCode & " STOCK ไม่พอขาย ต้องการขายสินค้านี้ ใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message ")
                        If vAnswer = 7 Then
                            Me.TBQTY.SelectAll()
                            Exit Sub
                        End If
                    End If
                End If

                If vQTY > vShelfQTY And vShelfUnit = vUnitCode Then
                    vAnswer = MsgBox("สินค้ารหัส " & vItemCode & " STOCK ไม่พอขาย ต้องการขายสินค้านี้ ใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message ")
                    If vAnswer = 7 Then
                        Me.TBQTY.SelectAll()
                        Exit Sub
                    End If
                End If

                If Me.TBPrice.Text <> "" Then
                    vPrice = Me.TBPrice.Text
                End If
                vAmount = vQTY * vPrice

                vIndex = Me.ListViewItem.Items.Count + 1

                If vQTY = 0 Then
                    MsgBox("ไม่ได้ระบุจำนวนของสินค้าที่ต้องการ หรือต้องระบุจำนวนสินค้าที่ต้องการมากกว่า 0", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                Dim n As Integer
                Dim vCheckItemCode As String
                Dim vEditQTY As Double
                Dim vEditPrice As Double
                Dim vItemAmount As Double


                If Me.ListViewItem.Items.Count > 0 Then
                    For n = 0 To Me.ListViewItem.Items.Count - 1
                        vCheckItemCode = Me.ListViewItem.Items(n).SubItems(4).Text

                        If vItemCode = vCheckItemCode Then
                            vEditPrice = Me.TBPrice.Text
                            vEditQTY = Me.TBQTY.Text
                            vItemAmount = vEditQTY * vEditPrice
                            Me.ListViewItem.Items(n).SubItems(2).Text = Format(vEditQTY, "##,##0.00")
                            Me.ListViewItem.Items(n).SubItems(6).Text = Format(vItemAmount, "##,##0.00")
                            vCheckExist = 1
                            GoTo line2
                        End If
                    Next
                End If

line2:

                If vCheckExist = 0 Then
                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(vItemName)
                    listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                    listItem.SubItems.Add(vUnitCode)
                    listItem.SubItems.Add(vItemCode)
                    listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                    listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                    listItem.SubItems.Add(vWHCode)
                    listItem.SubItems.Add(vShelfCode)
                    listItem.SubItems.Add(vBarCode)
                    listItem.SubItems.Add(vShelfID)
                    Me.ListViewItem.Items.Add(listItem)
                End If

                Call CalcItemAmount()

                Me.TBItem.Text = ""
                Me.TBMemBarCode.Text = ""
                Me.TBItemName.Text = ""
                Me.TBPrice.Text = ""
                Me.TBUnit.Text = ""
                Me.TBWHCode.Text = ""
                Me.TBShelfCode.Text = ""
                Me.TBShelfID.Text = ""
                Me.TBQTY.Text = ""
                Me.ListViewStock.Items.Clear()
                Me.ListViewWareHouse.Items.Clear()
                Me.PNItemDetails.Visible = False
                Me.BTNSave.Visible = True
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
            Else
                MsgBox("ไม่มีรายการสินค้าไม่สามารถเพิ่ม รายการสินค้าลงตะกร้าได้", MsgBoxStyle.Critical, "Send Error Message")
            End If
        End If
    End Sub

    Private Sub BTNLogIN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim vUserCode As String
        'Dim vPassWord As String
        'Dim vCheckTypeLogIn As String

        'vUserCode = Me.TBUserCode.Text
        'vPassWord = Me.TBPassword.Text

        'Dim vService As New WebReference.WebServiceCalc
        'vCheckLogIn = vService.vLogIn(vUserCode, vPassWord)

        'If vCheckLogIn <> "" Then
        '    Me.PNLogIn.Visible = False
        '    Me.PNDriveIn.Visible = False
        '    Me.TBUserID.Text = vCheckLogIn
        '    Call CallIDNumber()

        '    If Me.RDZone1.Checked = True Then
        '        vConnectZone = "01"
        '        vCheckTypeLogIn = "จุดจ่ายที่1"
        '    ElseIf Me.RDZone2.Checked = True Then
        '        vConnectZone = "02"
        '        vCheckTypeLogIn = "จุดจ่ายที่2"
        '    ElseIf Me.RDZone3.Checked = True Then
        '        vConnectZone = "03"
        '        vCheckTypeLogIn = "จุดจ่ายที่3"
        '    ElseIf Me.RDZone4.Checked = True Then
        '        vConnectZone = "04"
        '        vCheckTypeLogIn = "จุดจ่ายที่4"
        '    End If



        '    If vCheckTypeLogIn <> "05-Checker" Then
        '        Me.PNLogIn.Visible = False
        '        Me.PNDriveIn.Visible = True
        '        Me.PNDriveIn.BringToFront()
        '        Me.TBRefNo.Focus()
        '    Else
        '        Me.PNLogIn.Visible = False
        '        Me.PNDriveIn.Visible = False
        '    End If

        'Else
        '    MsgBox("ไม่สามารถเข้าใช้งานโปรแกรมได้ กรุณาตรวจสอบชื่อและรหัสผ่าน", MsgBoxStyle.Critical, "Send Error Message")
        '    Me.TBPassword.Text = ""
        'End If
    End Sub

    Private Sub TBUserCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBUserCode.KeyDown

        'If e.KeyCode = Keys.Enter Then
        If e.KeyCode = 0 Then
            Me.TBPassword.Focus()
        End If
        If e.KeyCode = Keys.Down Then
            Me.TBPassword.Focus()
        End If
        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub TBPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBPassword.KeyDown
        If e.KeyCode = Keys.Up Then
            Me.TBUserCode.Focus()
        End If
        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub MenuDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuDelete.Click
        Dim i As Integer

        i = Me.ListViewItem.FocusedItem.Index
        Me.ListViewItem.Items.RemoveAt(i)
        Call GenIDNumber()
        Call CalcItemAmount()
        Me.TBBarCode.Focus()
    End Sub
    Private Sub GenIDNumber()
        Dim i As Integer
        Dim j As Integer

        If Me.ListViewItem.Items.Count > 0 Then
            j = 0
            For i = 0 To Me.ListViewItem.Items.Count - 1
                j = j + 1
                Me.ListViewItem.Items(i).SubItems(0).Text = j
            Next
        End If
    End Sub

    Private Sub CMBZone_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Me.TBUserCode.Focus()
    End Sub


    Private Sub BTNCloseLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If vCheckLogIn = "" Then
        '    Application.Exit()
        'Else
        '    Me.PNLogIn.Visible = False
        'End If
    End Sub

    Private Sub MenuSearchPickUp()
        Me.PNLogIn.Visible = False
        Me.PNDriveIn.Visible = False

        Me.PNSearchPickUp.Visible = True
        Me.PNSearchPickUp.BringToFront()
        Me.TBSearchPickup.Focus()
    End Sub

    Private Sub BTNClosePickup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClosePickup.Click
        Application.Exit()
    End Sub

    Public Sub ExitProgram()
        Application.Exit()
    End Sub

    Private Sub TBSearchPickup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchPickup.KeyDown
        ''If e.KeyCode = Keys.Enter Then
        If e.KeyCode = 0 Then
            Call SearchPickRequest()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.ListViewSearhPickup.Items.Clear()
            Me.TBSearchPickup.Text = ""
            Me.PNSearchPickUp.Visible = False
            Me.PNDriveIn.Visible = True
            Me.PNDriveIn.BringToFront()
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Down Then
            If Me.ListViewSearhPickup.Items.Count > 0 Then
                Me.ListViewSearhPickup.Focus()
                Me.ListViewSearhPickup.Items(0).Focused = True
                Me.ListViewSearhPickup.Items(0).Selected = True
            End If
        End If
    End Sub

    Public Sub SearchPickRequest()
        Dim vSearch As String
        Dim i As Integer
        Dim vDocno As String
        Dim vRefNo As String
        Dim vAmount As Double
        Dim vIndex As Integer
        Dim vPointID As String

        vSearch = Me.TBSearchPickup.Text

        If Me.RDZone1.Checked = True Then
            vPointID = "01"
        ElseIf Me.RDZone2.Checked = True Then
            vPointID = "02"
        ElseIf Me.RDZone3.Checked = True Then
            vPointID = "03"
        End If

        vQuery = "exec dbo.USP_NP_SearchReqPicking '" & vPointID & "','" & vSearch & "'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        Me.ListViewSearhPickup.Items.Clear()
        vIndex = 0
        If ds.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds.Tables(0).Rows.Count - 1
                vDocno = ds.Tables(0).Rows(i)("docno").ToString
                vRefNo = ds.Tables(0).Rows(i)("refno").ToString
                vAmount = ds.Tables(0).Rows(i)("netdebtamount").ToString

                vIndex = vIndex + 1
                Dim listItem As New ListViewItem(vIndex)
                listItem.SubItems.Add(vRefNo)
                listItem.SubItems.Add(vDocno)
                listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                Me.ListViewSearhPickup.Items.Add(listItem)

            Next
            Me.ListViewSearhPickup.Focus()
            Me.ListViewSearhPickup.Items(0).Selected = True
        Else
            Me.TBSearchPickup.Focus()
        End If

    End Sub

    Private Sub TBSearchPickup_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchPickup.TextChanged

    End Sub

    Private Sub BTNSearchPickup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Call SearchPickRequest()

        'Dim vSearch As String
        'Dim i As Integer
        'Dim vDocno As String
        'Dim vDocDate As String
        'Dim vRefID As String
        'Dim vPickZone As String
        'Dim vAmount As Double
        'Dim vIndex As Integer

        'vSearch = Me.TBSearchPickup.Text
        'vQuery = "exec dbo.usp_np_SearchDriveInMaster '" & vSearch & "'"
        'Dim vService As New WebReference.WebServiceCalc
        'Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        'Me.ListViewSearhPickup.Items.Clear()
        'vIndex = 0
        'If ds.Tables(0).Rows.Count > 0 Then
        '    For i = 0 To ds.Tables(0).Rows.Count - 1
        '        vDocno = ds.Tables(0).Rows(i)("docno").ToString
        '        vDocDate = ds.Tables(0).Rows(i)("docdate").ToString
        '        vRefID = ds.Tables(0).Rows(i)("refid").ToString
        '        vPickZone = ds.Tables(0).Rows(i)("pickzone").ToString
        '        vAmount = ds.Tables(0).Rows(i)("totalnetamount").ToString

        '        If vPickZone = vConnectZone Then
        '            vIndex = vIndex + 1
        '            Dim listItem As New ListViewItem(vIndex)
        '            listItem.SubItems.Add(vRefID)
        '            listItem.SubItems.Add(vDocno)
        '            listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
        '            Me.ListViewSearhPickup.Items.Add(listItem)
        '        End If

        '    Next
        '    Me.ListViewSearhPickup.Focus()
        'Else
        '    Me.TBSearchPickup.Focus()
        'End If
    End Sub

    Private Sub LBLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim vUserCode As String
        'Dim vPassWord As String
        'Dim vCheckTypeLogIn As String

        'vUserCode = Me.TBUserCode.Text
        'vPassWord = Me.TBPassword.Text

        'Dim vService As New WebReference.WebServiceCalc
        'vCheckLogIn = vService.vLogIn(vUserCode, vPassWord)

        'If vCheckLogIn <> "" Then
        '    Me.PNLogIn.Visible = False
        '    Me.PNDriveIn.Visible = False

        '    Me.TBUserID.Text = vCheckLogIn
        '    Call CallIDNumber()

        '    If Me.RDZone1.Checked = True Then
        '        vConnectZone = "01"
        '        vCheckTypeLogIn = "จุดจ่ายที่1"
        '    ElseIf Me.RDZone2.Checked = True Then
        '        vConnectZone = "02"
        '        vCheckTypeLogIn = "จุดจ่ายที่2"
        '    ElseIf Me.RDZone3.Checked = True Then
        '        vConnectZone = "03"
        '        vCheckTypeLogIn = "จุดจ่ายที่3"
        '    ElseIf Me.RDZone4.Checked = True Then
        '        vConnectZone = "04"
        '        vCheckTypeLogIn = "จุดจ่ายที่4"
        '    End If



        '    If vCheckTypeLogIn <> "05-Checker" Then
        '        Me.PNLogIn.Visible = False
        '        Me.PNDriveIn.Visible = True
        '        Me.PNDriveIn.BringToFront()
        '        Me.TBRefNo.Focus()
        '    Else
        '        Me.PNLogIn.Visible = False
        '        Me.PNDriveIn.Visible = False
        '    End If

        'Else
        '    MsgBox("ไม่สามารถเข้าใช้งานโปรแกรมได้ กรุณาตรวจสอบชื่อและรหัสผ่าน", MsgBoxStyle.Critical, "Send Error Message")
        '    Me.TBPassword.Text = ""
        'End If
    End Sub

    Private Sub LBCloseLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim vAnswer As Integer

        'If vCheckLogIn = "" Then
        '    vAnswer = MsgBox("คุณต้องการออกโปรแกรมใช่หรือไม่ ?", MsgBoxStyle.YesNo, "Send Question Information")
        '    If vAnswer = 6 Then
        '        Application.Exit()
        '    Else
        '        Exit Sub
        '    End If
        'Else
        '    Me.PNLogIn.Visible = False
        'End If
    End Sub

    Private Sub LBAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '        Dim vItemCode As String
        '        Dim vItemName As String
        '        Dim vWHCode As String
        '        Dim vShelfCode As String
        '        Dim vQTY As Double
        '        Dim vPrice As Double
        '        Dim vAmount As Double
        '        Dim vUnitCode As String
        '        Dim vIndex As Integer
        '        Dim vCheckExist As Integer

        '        If Me.ListViewStock.Items.Count > 0 And Me.TBItem.Text <> "" Then
        '            vCheckExist = 0
        '            vItemCode = Me.TBItem.Text
        '            vItemName = Me.TBItemName.Text
        '            vWHCode = Me.TBWHCode.Text
        '            vShelfCode = Me.TBShelfCode.Text
        '            If Me.TBQTY.Text <> "" Then
        '                vQTY = Me.TBQTY.Text
        '            End If
        '            If Me.TBPrice.Text <> "" Then
        '                vPrice = Me.TBPrice.Text
        '            End If
        '            vAmount = vQTY * vPrice
        '            vUnitCode = Me.TBUnit.Text
        '            vIndex = Me.ListViewItem.Items.Count + 1

        '            If vQTY = 0 Then
        '                MsgBox("ไม่ได้ระบุจำนวนของสินค้าที่ต้องการ หรือต้องระบุจำนวนสินค้าที่ต้องการมากกว่า 0", MsgBoxStyle.Critical, "Send Error Message")
        '                Exit Sub
        '            End If


        '            Dim n As Integer
        '            Dim vCheckItemCode As String
        '            Dim vEditQTY As Double


        '            If Me.ListViewItem.Items.Count > 0 Then
        '                For n = 0 To Me.ListViewItem.Items.Count - 1
        '                    vCheckItemCode = Me.ListViewItem.Items(n).SubItems(4).Text
        '                    vEditQTY = Me.TBQTY.Text
        '                    If vItemCode = vCheckItemCode Then
        '                        Me.ListViewItem.Items(n).SubItems(2).Text = Format(vEditQTY, "##,##0.00")
        '                        Call CalcItemAmount()
        '                        vCheckExist = 1
        '                        GoTo line1
        '                    End If
        '                Next
        '            End If

        'line1:

        '            If vCheckExist = 0 Then
        '                Dim listItem As New ListViewItem(vIndex)
        '                listItem.SubItems.Add(vItemName)
        '                listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
        '                listItem.SubItems.Add(vUnitCode)
        '                listItem.SubItems.Add(vItemCode)
        '                listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
        '                listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
        '                listItem.SubItems.Add(vWHCode)
        '                listItem.SubItems.Add(vShelfCode)
        '                'listItem.SubItems.Add(vBarCode)
        '                Me.ListViewItem.Items.Add(listItem)
        '            End If

        '            Call CalcItemAmount()

        '            Me.TBItem.Text = ""
        '            Me.TBItemName.Text = ""
        '            Me.TBPrice.Text = ""
        '            Me.TBReserve.Text = ""
        '            Me.TBUnit.Text = ""
        '            Me.TBWHCode.Text = ""
        '            Me.TBShelfCode.Text = ""
        '            Me.TBQTY.Text = ""
        '            Me.ListViewStock.Items.Clear()
        '            Me.PNItemDetails.Visible = False
        '            Me.BTNSave.Visible = True
        '            Me.TBBarCode.Text = ""
        '            Me.TBBarCode.Focus()
        '        Else
        '            MsgBox("ไม่มีรายการสินค้าไม่สามารถเพิ่ม รายการสินค้าลงตะกร้าได้", MsgBoxStyle.Critical, "Send Error Message")
        '        End If
    End Sub

    Private Sub TBQTY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBQTY.TextChanged

    End Sub

    Private Sub MenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEdit.Click
        Dim vBarCode As String
        Dim vRate As Integer
        Dim vDefShelfCode As String
        Dim vStockUnit As String
        Dim i As Integer
        Dim vStore As String
        Dim vStkQTY As Double

        vSelectLineEdit = Me.ListViewItem.FocusedItem.Index
        vBarCode = Me.ListViewItem.Items(vSelectLineEdit).SubItems(9).Text
        vDefShelfCode = Me.ListViewItem.Items(vSelectLineEdit).SubItems(8).Text
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetDataBarCode(vBarCode)
        Me.ListViewStock.Items.Clear()
        Me.ListViewWareHouse.Items.Clear()


        If ds.Tables(0).Rows.Count > 0 Then
            vRate = ds.Tables(0).Rows(0)("rate").ToString

            For i = 0 To ds.Tables(0).Rows.Count - 1
                vStore = ds.Tables(0).Rows(i)("shelfcode").ToString
                vStkQTY = ds.Tables(0).Rows(i)("stock").ToString
                vStockUnit = ds.Tables(0).Rows(i)("stkunitcode").ToString

                If vDefShelfCode = vStore Then
                    Me.TBEditStock.Text = Format(vStkQTY, "##,##0.00")
                    Me.TBEditStockUnit.Text = vStockUnit
                End If
            Next
        End If

        Me.PNItemEdit.Visible = True
        Me.TBEditCode.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(4).Text
        Me.TBEditName.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(1).Text
        Me.TBEditUnit.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(3).Text
        Me.TBEditPrice.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(5).Text
        Me.TBEditQty.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(2).Text
        Me.TBEditRate.Text = Format(vRate, "##,##0.00")
        Me.TBDefSaleUnitCode.Text = vDefShelfCode
        Me.TBEditQty.Focus()
        Me.TBEditQty.SelectAll()
    End Sub

    Private Sub LBItemEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim vQTY As Double
        'Dim vPrice As Double
        'Dim vAmount As Double

        'If Me.TBEditQty.Text <> "" Then
        '    vQTY = Me.TBEditQty.Text
        'End If
        'If Me.TBEditPrice.Text <> "" Then
        '    vPrice = Me.TBEditPrice.Text
        'End If
        'vAmount = vQTY * vPrice

        'Me.ListViewItem.Items(vSelectLineEdit).SubItems(2).Text = Format(vQTY, "##,##0.00")
        'Me.ListViewItem.Items(vSelectLineEdit).SubItems(6).Text = Format(vAmount, "##,##0.00")
        'Call CalcItemAmount()
        'Me.TBEditQty.Text = ""
        'Me.PNItemEdit.Visible = False
        'Me.TBBarCode.Focus()
    End Sub

    Private Sub LBCloseEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Me.PNItemEdit.Visible = False
        'Me.TBBarCode.Focus()
    End Sub

    Private Sub MenuSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuSelect.Click
        Dim i As Integer
        Dim vDocno As String
        Dim n As Integer
        Dim vNetItemAmount As Double

        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vPickZone As String
        Dim vBarCode As String
        Dim vIndex As Integer

        n = Me.ListViewSearhPickup.FocusedItem.Index
        vDocno = Me.ListViewSearhPickup.Items(n).SubItems(2).Text

        vQuery = "exec dbo.usp_np_SearchPickUp '" & vDocno & "'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        Me.ListViewItem.Items.Clear()
        If ds.Tables(0).Rows.Count > 0 Then
            Me.TBRefNo.Text = ds.Tables(0).Rows(i)("refid").ToString
            vNetItemAmount = ds.Tables(0).Rows(i)("totalnetamount").ToString
            Me.TBItemAmount.Text = Format(vNetItemAmount, "##,##0.00")
            Me.TBDocNo.Text = ds.Tables(0).Rows(i)("docno").ToString

            vIndex = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1

                vPickZone = ds.Tables(0).Rows(i)("pickzone").ToString
                vItemCode = ds.Tables(0).Rows(i)("itemcode").ToString
                vItemName = ds.Tables(0).Rows(i)("itemname").ToString
                vWHCode = ds.Tables(0).Rows(i)("whcode").ToString
                vShelfCode = ds.Tables(0).Rows(i)("shelfcode").ToString
                vQTY = ds.Tables(0).Rows(i)("qty").ToString
                vUnitCode = ds.Tables(0).Rows(i)("unitcode").ToString
                vPrice = ds.Tables(0).Rows(i)("price").ToString
                vAmount = ds.Tables(0).Rows(i)("amount").ToString
                vBarCode = ds.Tables(0).Rows(i)("barcode").ToString

                vIndex = vIndex + 1
                Dim listItem As New ListViewItem(vIndex)
                listItem.SubItems.Add(vItemName)
                listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                listItem.SubItems.Add(vUnitCode)
                listItem.SubItems.Add(vItemCode)
                listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                listItem.SubItems.Add(vWHCode)
                listItem.SubItems.Add(vShelfCode)
                listItem.SubItems.Add(vBarCode)
                Me.ListViewItem.Items.Add(listItem)
            Next
        End If
        Me.ListViewSearhPickup.Items.Clear()
        Me.TBSearchPickup.Text = ""
        Me.PNSearchPickUp.Visible = False
        Me.PNDriveIn.Visible = True
        Me.PNDriveIn.BringToFront()
        Me.TBBarCode.Focus()
    End Sub

    Private Sub BTNCloseSelectPickup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Me.ListViewSearhPickup.Items.Clear()
        'Me.TBSearchPickup.Text = ""
        'Me.PNSearchPickUp.Visible = False
    End Sub

    Private Sub TBEditQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBEditQty.KeyDown
        Dim vItemCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String

        Dim vShelfUnit As String
        Dim vShelfQTY As Double
        Dim vTotalQTY As Double
        Dim vRate As Integer

        Dim vAnswer As Integer

        Dim vEditIndex As Integer

        If e.KeyCode = Keys.Escape Then
            vEditIndex = Me.TBEditIndex.Text
            Me.PNItemEdit.Visible = False
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
        End If

        'If e.KeyCode = Keys.Enter Then
        If e.KeyCode = 0 Then
            If Me.TBEditQty.Text <> "" Then
                vQTY = Me.TBEditQty.Text
            End If

            vEditIndex = Me.TBEditIndex.Text
            vItemCode = Me.TBEditCode.Text
            vUnitCode = Me.TBEditUnit.Text
            vShelfUnit = Me.TBEditStockUnit.Text
            If Me.TBEditRate.Text <> "" Then
                vRate = Me.TBEditRate.Text
            End If
            If Me.TBEditStock.Text <> "" Then
                vShelfQTY = Me.TBEditStock.Text
            End If

            If vShelfUnit <> vUnitCode Then
                vTotalQTY = vShelfQTY / vRate
                If vQTY > vTotalQTY Then
                    vAnswer = MsgBox("สินค้ารหัส " & vItemCode & " STOCK ไม่พอขาย ต้องการขายสินค้านี้ ใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message ")
                    If vAnswer = 7 Then
                        Me.TBEditQty.SelectAll()
                        Exit Sub
                    End If
                End If
            End If

            If vQTY > vShelfQTY And vShelfUnit = vUnitCode Then
                vAnswer = MsgBox("สินค้ารหัส " & vItemCode & " STOCK ไม่พอขาย ต้องการขายสินค้านี้ ใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message ")
                If vAnswer = 7 Then
                    Me.TBEditQty.SelectAll()
                    Exit Sub
                End If
            End If


            If Me.TBEditPrice.Text <> "" Then
                vPrice = Me.TBEditPrice.Text
            End If
            vAmount = vQTY * vPrice

            Me.ListViewItem.Items(vSelectLineEdit).SubItems(2).Text = Format(vQTY, "##,##0.00")
            Me.ListViewItem.Items(vSelectLineEdit).SubItems(6).Text = Format(vAmount, "##,##0.00")
            Call CalcItemAmount()
            Me.TBEditQty.Text = ""
            Me.PNItemEdit.Visible = False
            If Me.ListViewItem.Items.Count = 1 Then
                Me.TBBarCode.Focus()
            ElseIf vEditIndex = Me.ListViewItem.Items.Count - 1 And Me.ListViewItem.Items.Count > 1 Then
                Me.ListViewItem.Focus()
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            ElseIf vEditIndex < Me.ListViewItem.Items.Count - 1 And Me.ListViewItem.Items.Count > 1 Then
                Me.ListViewItem.Focus()
                Me.ListViewItem.Items(vEditIndex + 1).Selected = True
                Me.ListViewItem.Items(vEditIndex + 1).Focused = True
            Else
                Me.ListViewItem.Focus()
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
        End If
    End Sub

    Private Sub BTNClearPickUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearPickUp.Click
        Me.TBRefNo.Text = ""
        Me.TBBarCode.Text = ""
        Me.TBItemAmount.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.LBLSaveMessage.Text = ""
    End Sub

    Public Sub ClearScreen()
        Me.TBRefNo.Text = ""
        Me.TBBarCode.Text = ""
        Me.TBItemAmount.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.LBLSaveMessage.Text = ""
    End Sub

    Private Sub RDZone1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDZone1.KeyDown
        'If e.KeyCode = Keys.Enter Then
        If e.KeyCode = 0 Then
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D1 Then
            Me.RDZone1.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D2 Then
            Me.RDZone2.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D3 Then
            Me.RDZone3.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub RDZone2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDZone2.KeyDown
        'If e.KeyCode = Keys.Enter Then
        If e.KeyCode = 0 Then
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D1 Then
            Me.RDZone1.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D2 Then
            Me.RDZone2.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D3 Then
            Me.RDZone3.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub RDZone3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDZone3.KeyDown
        'If e.KeyCode = Keys.Enter Then
        If e.KeyCode = 0 Then
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D1 Then
            Me.RDZone1.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D2 Then
            Me.RDZone2.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D3 Then
            Me.RDZone3.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub RDZone4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        'If e.KeyCode = Keys.Enter Then
        '    Me.TBUserCode.Focus()
        'End If

        'If e.KeyCode = Keys.D1 Then
        '    Me.RDZone1.Checked = True
        '    Me.TBUserCode.Focus()
        'End If

        'If e.KeyCode = Keys.D2 Then
        '    Me.RDZone2.Checked = True
        '    Me.TBUserCode.Focus()
        'End If

        'If e.KeyCode = Keys.D3 Then
        '    Me.RDZone3.Checked = True
        '    Me.TBUserCode.Focus()
        'End If

        'If e.KeyCode = Keys.Escape Then
        '    Dim vAnswer As Integer
        '    vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
        '    If vAnswer = 6 Then
        '        Application.Exit()
        '    End If
        'End If
    End Sub

    Private Sub TBPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPassword.TextChanged
        Dim vLenPassword As Integer
        Dim vCheckTypeLogIn As String

        On Error GoTo ErrDescription

        vLenPassword = Len(Me.TBPassword.Text)
        If vLenPassword = 4 And Me.TBUserCode.Text <> "" Then

            Me.TBPassword.Visible = False
            vUserCode = Me.TBUserCode.Text
            vPassWord = Me.TBPassword.Text


            Dim vService1 As New WebReference.WebServiceCalc
            Dim ds1 As DataSet = vService1.vLogIn(vUserCode, vPassWord)

            If ds1.Tables(0).Rows.Count > 0 Then
                vCheckLogIn = ds1.Tables(0).Rows(0)("username").ToString
                vUserName = ds1.Tables(0).Rows(0)("username").ToString
                vDuty = ds1.Tables(0).Rows(0)("duty").ToString
                vLevelID = ds1.Tables(0).Rows(0)("levelid").ToString
                vMemSaleName = ds1.Tables(0).Rows(0)("salename").ToString
            Else
                vCheckLogIn = ""
                vUserName = ""
                vDuty = ""
                vLevelID = 0
                vMemSaleName = ""
            End If

            If vCheckLogIn <> "" Then

                Me.PNLogIn.Visible = False
                Me.TBUserID.Text = vCheckLogIn
                Call CallIDNumber()
                Call SearchConditionSend()

                If Me.RDZone1.Checked = True Then
                    vConnectZone = "01"
                    vCheckTypeLogIn = "จุดจ่ายที่1"
                    'Me.LBLZoneID.Text = "01"
                ElseIf Me.RDZone2.Checked = True Then
                    vConnectZone = "02"
                    vCheckTypeLogIn = "จุดจ่ายที่2"
                    'Me.LBLZoneID.Text = "02"
                ElseIf Me.RDZone3.Checked = True Then
                    vConnectZone = "03"
                    vCheckTypeLogIn = "จุดจ่ายที่3"
                    'Me.LBLZoneID.Text = "03"
                End If
                Me.TBSaleCode.Text = vMemSaleName
                Me.PNDriveIn.Visible = True
                Me.PNDriveIn.BringToFront()
                Me.TBRefNo.Focus()
            Else
                MsgBox("ไม่สามารถเข้าใช้งานโปรแกรมได้ กรุณาตรวจสอบชื่อและรหัสผ่าน", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBPassword.Visible = True
                Me.TBPassword.Text = ""
                Me.TBSaleCode.Text = ""
                Me.TBPassword.Focus()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub ListViewItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItem.KeyDown
        Dim vItemCode As String
        Dim vIndex As Integer
        Dim vAnswerDelete As Integer


        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = Keys.Back Then
            If Me.ListViewItem.Items.Count > 0 Then
                vIndex = Me.ListViewItem.FocusedItem.Index
                vItemCode = Me.ListViewItem.Items(vIndex).SubItems(1).Text
                vAnswerDelete = MsgBox("คุณต้องการลบรายการ รหัส " & vItemCode & " นี้ใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
                If vAnswerDelete = 6 Then
                    Me.ListViewItem.Items.RemoveAt(vIndex)
                    Call GenIDNumber()
                    Call CalcItemAmount()
                    Me.TBBarCode.Focus()
                End If
            End If
        End If

        'If e.KeyCode = Keys.Enter Then
        If e.KeyCode = 0 Then
            If Me.ListViewItem.Items.Count > 0 Then
                Dim vBarCode As String
                Dim vRate As Integer
                Dim vDefWHCode As String
                Dim vDefShelfCode As String
                Dim vStockUnit As String
                Dim i As Integer
                Dim vWHCode As String
                Dim vStore As String
                Dim vStkQTY As Double

                On Error Resume Next

                vSelectLineEdit = Me.ListViewItem.FocusedItem.Index
                vBarCode = Me.ListViewItem.Items(vSelectLineEdit).SubItems(9).Text
                vDefWHCode = Me.ListViewItem.Items(vSelectLineEdit).SubItems(7).Text
                vDefShelfCode = Me.ListViewItem.Items(vSelectLineEdit).SubItems(8).Text
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetDataBarCode(vBarCode)
                Me.ListViewStock.Items.Clear()
                Me.ListViewWareHouse.Items.Clear()


                If ds.Tables(0).Rows.Count > 0 Then
                    vRate = ds.Tables(0).Rows(0)("rate").ToString

                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        vWHCode = ds.Tables(0).Rows(i)("whcode").ToString
                        vStore = ds.Tables(0).Rows(i)("shelfcode").ToString
                        vStkQTY = ds.Tables(0).Rows(i)("stock").ToString
                        vStockUnit = ds.Tables(0).Rows(i)("stkunitcode").ToString

                        If vDefWHCode = vWHCode And vDefShelfCode = vStore Then
                            Me.TBEditStock.Text = Format(vStkQTY, "##,##0.00")
                            Me.TBEditStockUnit.Text = vStockUnit
                        End If
                    Next
                End If

                Me.PNItemEdit.Visible = True
                Me.TBEditCode.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(4).Text
                Me.TBEditName.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(1).Text
                Me.TBEditUnit.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(3).Text
                Me.TBEditPrice.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(5).Text
                Me.TBEditQty.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(2).Text
                Me.TBEditRate.Text = Format(vRate, "##,##0.00")
                Me.TBDefSaleUnitCode.Text = vDefShelfCode
                Me.TBEditIndex.Text = vSelectLineEdit
                Me.TBEditQty.Focus()
                Me.TBEditQty.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.Up Then
            Dim vCount As Integer
            Dim vSelectID As Integer
            Dim i As Integer

            If Me.ListViewItem.Items.Count > 0 Then
                vCount = Me.ListViewItem.Items.Count
                For i = 0 To Me.ListViewItem.Items.Count - 1
                    If Me.ListViewItem.Items(i).Selected = True Then
                        vSelectID = i + 1
                        GoTo Line2
                    Else
                        vSelectID = 0
                    End If
                Next

            End If
Line2:
            If vSelectID = 0 Then
                Me.TBBarCode.Focus()
            ElseIf vSelectID = 1 Then
                Me.TBBarCode.Focus()
            End If
        End If


        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub BTNBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNBack.Click
        Me.TBRefNo.Text = ""
        Me.TBBarCode.Text = ""
        Me.TBItemAmount.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.PNDriveIn.Visible = False
        Me.PNLogIn.Visible = True
        Me.PNLogIn.BringToFront()
        Me.TBPassword.Text = ""
        Me.TBPassword.Visible = True
        Me.RDZone1.Focus()
    End Sub

    Public Sub PageLogIn()
        Me.TBRefNo.Text = ""
        Me.TBBarCode.Text = ""
        Me.TBItemAmount.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.PNDriveIn.Visible = False
        Me.PNLogIn.Visible = True
        Me.PNLogIn.BringToFront()
        Me.TBPassword.Text = ""
        Me.TBPassword.Visible = True
        Me.RDZone1.Focus()
    End Sub

    Private Sub TBEditQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBEditQty.TextChanged

    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        Me.PNLogIn.Visible = False
        Me.PNDriveIn.Visible = False
        Me.PNSearchPickUp.Visible = True
        Call SearchPickRequest()
        Me.PNSearchPickUp.BringToFront()
        Me.TBSearchPickup.Text = ""
        Me.TBSearchPickup.Focus()
    End Sub

    Public Sub SearchDoc()
        Me.PNLogIn.Visible = False
        Me.PNDriveIn.Visible = False
        Me.PNSearchPickUp.Visible = True
        Call SearchPickRequest()
        Me.PNSearchPickUp.BringToFront()
        Me.TBSearchPickup.Text = ""
        Me.TBSearchPickup.Focus()
    End Sub


    Private Sub BTNSave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSave.KeyDown
        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub ListViewItem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewItem.SelectedIndexChanged

    End Sub

    Private Sub TBDefSaleUnitCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBDefSaleUnitCode.TextChanged

    End Sub

    Private Sub ListViewSearhPickup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearhPickup.KeyDown
        Dim i As Integer
        Dim vDocno As String
        Dim n As Integer
        Dim vNetItemAmount As Double
        Dim vIsConditionSend As Integer

        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vPickZone As String
        Dim vBarCode As String
        Dim vShelfID As String
        Dim vIndex As Integer


        If e.KeyCode = Keys.Escape Then
            Me.ListViewSearhPickup.Items.Clear()
            Me.TBSearchPickup.Text = ""
            Me.PNSearchPickUp.Visible = False
        End If


        'If e.KeyCode = Keys.Enter Then
        If e.KeyCode = 0 Then
            On Error Resume Next
            If Me.ListViewSearhPickup.Items.Count > 0 Then
                n = Me.ListViewSearhPickup.FocusedItem.Index
                vDocno = Me.ListViewSearhPickup.Items(n).SubItems(2).Text

                'vQuery = "exec dbo.usp_np_SearchPickUp '" & vDocno & "'"

                vQuery = "exec dbo.usp_np_SearchReqPickingDetails '" & vDocno & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                Me.ListViewItem.Items.Clear()
                If ds.Tables(0).Rows.Count > 0 Then
                    Me.TBRefNo.Text = ds.Tables(0).Rows(0)("refno").ToString
                    vNetItemAmount = ds.Tables(0).Rows(0)("netdebtamount").ToString
                    vIsConditionSend = ds.Tables(0).Rows(0)("isconditionsend").ToString
                    Me.TBItemAmount.Text = Format(vNetItemAmount, "##,##0.00")
                    Me.TBDocNo.Text = ds.Tables(0).Rows(0)("docno").ToString
                    Me.TBARCode.Text = ds.Tables(0).Rows(0)("arcode").ToString
                    Me.TBSaleCode.Text = ds.Tables(0).Rows(0)("salecode").ToString & "/" & ds.Tables(0).Rows(0)("salename").ToString
                    Me.CMBConditionSend.SelectedIndex = vIsConditionSend

                    vIndex = 0
                    For i = 0 To ds.Tables(0).Rows.Count - 1

                        vPickZone = ds.Tables(0).Rows(i)("pointid").ToString
                        vItemCode = ds.Tables(0).Rows(i)("itemcode").ToString
                        vItemName = ds.Tables(0).Rows(i)("itemname").ToString
                        vWHCode = ds.Tables(0).Rows(i)("whcode").ToString
                        vShelfCode = ds.Tables(0).Rows(i)("shelfcode").ToString
                        vQTY = ds.Tables(0).Rows(i)("qty").ToString
                        vUnitCode = ds.Tables(0).Rows(i)("unitcode").ToString
                        vPrice = ds.Tables(0).Rows(i)("price").ToString
                        vAmount = ds.Tables(0).Rows(i)("amount").ToString
                        vBarCode = ds.Tables(0).Rows(i)("barcode").ToString
                        vShelfID = ds.Tables(0).Rows(i)("shelfid").ToString

                        vIndex = vIndex + 1
                        Dim listItem As New ListViewItem(vIndex)
                        listItem.SubItems.Add(vItemName)
                        listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                        listItem.SubItems.Add(vUnitCode)
                        listItem.SubItems.Add(vItemCode)
                        listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                        listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                        listItem.SubItems.Add(vWHCode)
                        listItem.SubItems.Add(vShelfCode)
                        listItem.SubItems.Add(vBarCode)
                        listItem.SubItems.Add(vShelfID)
                        Me.ListViewItem.Items.Add(listItem)
                    Next
                End If

                Me.ListViewSearhPickup.Items.Clear()
                Me.TBSearchPickup.Text = ""
                Me.PNSearchPickUp.Visible = False
                Me.PNDriveIn.Visible = True
                Me.PNDriveIn.BringToFront()
                Me.TBBarCode.Focus()
            End If
        End If
    End Sub


    Private Sub PNLogIn_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PNLogIn.GotFocus

    End Sub

    Private Sub TBARCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBARCode.KeyDown
        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub TBARCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBARCode.TextChanged
        Dim vQuery As String
        Dim vSearchAR As String

        On Error GoTo ErrDescription

        If Me.TBARCode.Text <> "" Then
            vSearchAR = Me.TBARCode.Text

            vQuery = "exec dbo.usp_ar_searchar1 '" & vSearchAR & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            If ds.Tables(0).Rows.Count > 0 Then
                Me.TBArName.Text = ds.Tables(0).Rows(0)("arname").ToString()
                Me.TBMemberID.Text = ds.Tables(0).Rows(0)("memberid").ToString
                Me.TBSaleCode.Focus()
            Else
                Me.TBArName.Text = ""
                Me.TBMemBarCode.Text = ""
                Me.TBARCode.Focus()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Public Sub SearchConditionSend()
        Me.CMBConditionSend.Items.Clear()
        Me.CMBConditionSend.Items.Add("รับเอง")
        Me.CMBConditionSend.Items.Add("ส่งให้")
        Me.CMBConditionSend.SelectedIndex = 0
    End Sub

    Private Sub TBSaleCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSaleCode.KeyDown
        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub TBSaleCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSaleCode.TextChanged

    End Sub

    Private Sub BTNSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSearch.KeyDown
        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub BTNClearPickUp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClearPickUp.KeyDown
        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub BTNBack_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNBack.KeyDown
        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub BTNClosePickup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClosePickup.KeyDown
        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub TBItemAmount_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBItemAmount.ParentChanged

    End Sub

    Private Sub TBRefNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBRefNo.TextChanged

    End Sub

    Private Sub TBMemberID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBMemberID.KeyDown
        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub TBMemberID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBMemberID.TextChanged

    End Sub

    Private Sub TBArName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBArName.KeyDown
        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub TBArName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBArName.TextChanged

    End Sub

    Private Sub CMBConditionSend_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBConditionSend.KeyDown
        If e.KeyCode = 34 Then
            Call SavePickRequest()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDoc()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub CMBConditionSend_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBConditionSend.SelectedIndexChanged

    End Sub

    Private Sub TBUserID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBUserID.KeyDown

    End Sub

    Private Sub TBUserID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBUserID.TextChanged

    End Sub
End Class
