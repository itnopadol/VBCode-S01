Imports System.IO
Imports Symbol
Imports Symbol.Barcode
Imports Symbol.Barcode.Reader
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.ComponentModel
Imports System.Windows.Forms
Imports vb6 = Microsoft.VisualBasic
Public Class FormPOS
    Private MyScanner As Symbol.Barcode.Device = Nothing

    Private MyReader As Symbol.Barcode.Reader = Nothing
    Private MyReaderData As Symbol.Barcode.ReaderData = Nothing
    Private MyEventHandler As System.EventHandler = Nothing

    Private MyReadNotifyHander As System.EventHandler = Nothing
    Private MyStatusNotifyHandler As System.EventHandler = Nothing
    Private MyActivateHandler As System.EventHandler = Nothing
    Private MyDeActivateHandler As System.EventHandler = Nothing

    Dim vQuery As String
    Dim vMemDocDate As String
    Dim vCalcStock As Integer
    Dim vInvoiceBillStatus As Integer
    Dim vInvoiceIsCancel As Integer
    Dim vInvoiceIsConfirm As Integer
    Dim vInvoiceIsOpen As Integer

    Dim vMemPayReceiptAmount As Double


    Private Sub FormPOS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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
        Dim vStkUnit As String
        Dim vStkQTY As Double
        Dim vDefWHCode As String
        Dim vDefShelfCode As String
        Dim vBarCode1 As String


        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            vBarCode = Me.TBBarCode.Text

            Me.TBBarCode.Text = vBarCode

            vQuery = "exec dbo.USP_MB_SearchBarCode '" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)
            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vPrice = pds.Tables(0).Rows(0)("price").ToString
                vRate = pds.Tables(0).Rows(0)("rate").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vDefWHCode = pds.Tables(0).Rows(0)("defsalewhcode").ToString
                vDefShelfCode = pds.Tables(0).Rows(0)("defsaleshelf").ToString
                vBarCode1 = pds.Tables(0).Rows(0)("barcode").ToString
                vStkUnit = pds.Tables(0).Rows(0)("stkunitcode").ToString
                vStkQTY = pds.Tables(0).Rows(0)("stock").ToString

                Dim n As Integer
                Dim vCheckItemCode As String
                Dim vCheckQTY As Double
                Dim vCheckWHCode As String
                Dim vCheckShelfCode As String
                Dim vCheckUnitCode As String

                If Me.ListViewItem.Items.Count > 0 Then
                    For n = 0 To Me.ListViewItem.Items.Count - 1
                        vCheckItemCode = Me.ListViewItem.Items(n).SubItems(6).Text
                        vCheckUnitCode = Me.ListViewItem.Items(n).SubItems(5).Text
                        vCheckWHCode = Me.ListViewItem.Items(n).SubItems(8).Text
                        vCheckShelfCode = Me.ListViewItem.Items(n).SubItems(9).Text
                        vCheckQTY = Me.ListViewItem.Items(n).SubItems(2).Text

                        If vItemCode = vCheckItemCode And vUnitCode = vCheckUnitCode And vDefWHCode = vCheckWHCode And vDefShelfCode = vCheckShelfCode Then
                            Me.TBQty.Text = Format(vCheckQTY, "##,##0.00")
                            GoTo Line1
                        End If
                    Next
                End If

Line1:
                If Me.TBQty.Text = "" Then
                    Me.TBQty.Text = 1
                End If
                Me.TBQty.Focus()
                Me.TBQty.SelectAll()
            Else
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                MsgBox("This item find not found ", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If

            Me.TBItemCode.Text = vItemCode
            Me.TBItemName.Text = vItemName
            Me.TBPrice.Text = Format(vPrice, "##,##0.00")
            Me.TBRate1.Text = Format(vRate, "##,##0.00")
            Me.TBUnit.Text = vUnitCode
            Me.TBWHCode.Text = vDefWHCode
            Me.TBShelfCode.Text = vDefShelfCode
            Me.TBRemainQty.Text = Format(vStkQTY, "##,##0.00")
            Me.TBStkUnit.Text = vStkUnit
        End If



ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBBarCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBBarCode.KeyPress

    End Sub

    Private Sub TBBarCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged
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
        Dim vDefWHCode As String
        Dim vDefShelfCode As String
        Dim vBarCode1 As String


        'On Error Resume Next

        If vb6.InStr(Me.TBBarCode.Text, "@") > 0 Then
            vBarCode = vb6.Left(Me.TBBarCode.Text, vb6.Len(Me.TBBarCode.Text) - 1)

            Me.TBBarCode.Text = vBarCode

            vQuery = "exec dbo.USP_MB_SearchBarCode '" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)
            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vPrice = pds.Tables(0).Rows(0)("price").ToString
                vRate = pds.Tables(0).Rows(0)("rate").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vDefWHCode = pds.Tables(0).Rows(0)("defsalewhcode").ToString
                vDefShelfCode = pds.Tables(0).Rows(0)("defsaleshelf").ToString
                vBarCode1 = pds.Tables(0).Rows(0)("barcode").ToString
                vStkUnit = pds.Tables(0).Rows(0)("stkunitcode").ToString
                vStkQTY = pds.Tables(0).Rows(0)("stock").ToString

                Dim n As Integer
                Dim vCheckItemCode As String
                Dim vCheckQTY As Double
                Dim vCheckWHCode As String
                Dim vCheckShelfCode As String
                Dim vCheckUnitCode As String

                If Me.ListViewItem.Items.Count > 0 Then
                    For n = 0 To Me.ListViewItem.Items.Count - 1
                        vCheckItemCode = Me.ListViewItem.Items(n).SubItems(6).Text
                        vCheckUnitCode = Me.ListViewItem.Items(n).SubItems(5).Text
                        vCheckWHCode = Me.ListViewItem.Items(n).SubItems(8).Text
                        vCheckShelfCode = Me.ListViewItem.Items(n).SubItems(9).Text
                        vCheckQTY = Me.ListViewItem.Items(n).SubItems(2).Text

                        If vItemCode = vCheckItemCode And vUnitCode = vCheckUnitCode And vDefWHCode = vCheckWHCode And vDefShelfCode = vCheckShelfCode Then
                            Me.TBQty.Text = Format(vCheckQTY, "##,##0.00")
                            GoTo Line1
                        End If
                    Next
                End If

Line1:
                If Me.TBQty.Text = "" Then
                    Me.TBQty.Text = 1
                End If
                Me.TBQty.Focus()
                Me.TBQty.SelectAll()
            Else
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                MsgBox("This item find not found ", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If

            Me.TBItemCode.Text = vItemCode
            Me.TBItemName.Text = vItemName
            Me.TBPrice.Text = Format(vPrice, "##,##0.00")
            Me.TBRate1.Text = Format(vRate, "##,##0.00")
            Me.TBUnit.Text = vUnitCode
            Me.TBWHCode.Text = vDefWHCode
            Me.TBShelfCode.Text = vDefShelfCode
        End If



ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        FormMainApplication.Show()
        Me.Hide()
    End Sub

    Private Sub TBQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBQty.KeyDown
        Dim vBarCode As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vIndex As Integer
        Dim vItemLine As Integer
        Dim vCheckExist As Integer

        Dim vShelfQTY As Double
        Dim vShelfUnit As String
        Dim vRate1 As Integer
        Dim vTotalQTY As Double

        Dim vAnswer As Integer
        Dim vCheckPrice As Double

        If e.KeyCode = Keys.Enter And Me.TBItemName.Text <> "" Then

            If Me.TBQty.Text = "" Or Me.TBQty.Text = "0" Then
                MsgBox("No item for sale", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                Exit Sub
            End If

            If Me.CMBMemCalcStock.Text <> "" Then
                vCalcStock = 0
            End If

            If Me.TBPrice.Text <> "" Then
                vCheckPrice = Me.TBPrice.Text
            End If

            If vCheckPrice = 0 Then
                MsgBox("This item is not set saleprice", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                Exit Sub
            End If

            vCheckExist = 0
            vBarCode = Me.TBBarCode.Text
            vItemCode = Me.TBItemCode.Text
            vItemName = Me.TBItemName.Text
            vWHCode = Me.TBWHCode.Text
            vShelfCode = Me.TBShelfCode.Text
            vUnitCode = Me.TBUnit.Text
            vRate1 = Me.TBRate1.Text
            If Me.TBRemainQty.Text <> "" Then
                vShelfQTY = Me.TBRemainQty.Text
            Else
                vShelfQTY = 0
            End If
            vShelfUnit = Me.TBStkUnit.Text


            If Me.TBQty.Text <> "" Then
                vQTY = Me.TBQty.Text
            End If

            If vShelfUnit <> vUnitCode Then
                vTotalQTY = vShelfQTY / vRate1
                If vQTY > vTotalQTY Then
                    vAnswer = MsgBox("This item qty less than ,Do you want sale this item ?", MsgBoxStyle.YesNo, "Send Question Message ")
                    If vAnswer = 7 Then
                        Me.TBQty.SelectAll()
                        Exit Sub
                    End If
                End If
            End If

            If vQTY > vShelfQTY And vShelfUnit = vUnitCode Then
                vAnswer = MsgBox("This item qty less than ,Do you want sale this item ?", MsgBoxStyle.YesNo, "Send Question Message ")
                If vAnswer = 7 Then
                    Me.TBQty.SelectAll()
                    Exit Sub
                End If
            End If

            If Me.TBPrice.Text <> "" Then
                vPrice = Me.TBPrice.Text
            End If
            vAmount = vQTY * vPrice

            vIndex = Me.ListViewItem.Items.Count + 1
            vItemLine = Me.ListViewItem.Items.Count

            If vQTY = 0 Then
                MsgBox("Please insert qty more than 0", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBQty.Focus()
                Me.TBQty.SelectAll()
                Exit Sub
            End If

            Dim n As Integer
            Dim vCheckItemCode As String
            Dim vCheckWHCode As String
            Dim vCheckUnitCode As String
            Dim vCheckShelfCode As String

            Dim vEditQTY As Double
            Dim vEditPrice As Double
            Dim vItemAmount As Double


            If Me.ListViewItem.Items.Count > 0 Then
                For n = 0 To Me.ListViewItem.Items.Count - 1
                    vCheckItemCode = Me.ListViewItem.Items(n).SubItems(6).Text
                    vCheckUnitCode = Me.ListViewItem.Items(n).SubItems(5).Text
                    vCheckWHCode = Me.ListViewItem.Items(n).SubItems(8).Text
                    vCheckShelfCode = Me.ListViewItem.Items(n).SubItems(9).Text

                    If vItemCode = vCheckItemCode And vUnitCode = vCheckUnitCode And vWHCode = vCheckWHCode And vShelfCode = vCheckShelfCode Then
                        vEditPrice = Me.TBPrice.Text
                        vEditQTY = Me.TBQty.Text
                        vItemAmount = vEditQTY * vEditPrice
                        Me.ListViewItem.Items(n).SubItems(2).Text = Format(vEditQTY, "##,##0.00")
                        Me.ListViewItem.Items(n).SubItems(3).Text = Format(vEditPrice, "##,##0.00")
                        Me.ListViewItem.Items(n).SubItems(4).Text = Format(vItemAmount, "##,##0.00")
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
                listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                listItem.SubItems.Add(vUnitCode)
                listItem.SubItems.Add(vItemCode)
                listItem.SubItems.Add(vBarCode)
                listItem.SubItems.Add(vWHCode)
                listItem.SubItems.Add(vShelfCode)
                listItem.SubItems.Add(vRate1)
                listItem.SubItems.Add(0)
                Me.ListViewItem.Items.Add(listItem)

                Me.ListViewItem.Items.Item(vItemLine).BackColor = Color.White

            End If

            Call CalcItemAmount()

            If vQTY >= 10000 Then
                MsgBox("This qty more than 10,000.Please check qty again", MsgBoxStyle.Information, "Send Error Message")
            End If

            Me.TBItemCode.Text = ""
            Me.TBBarCode.Text = ""
            Me.TBItemName.Text = ""
            Me.TBPrice.Text = ""
            Me.TBUnit.Text = ""
            Me.TBWHCode.Text = ""
            Me.TBShelfCode.Text = ""
            Me.TBQty.Text = ""
            Me.TBStkUnit.Text = ""
            Me.TBRemainQty.Text = ""
            Me.BTNSave.Visible = True
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()

        End If

    End Sub

    Private Sub CalcItemAmount()
        Dim i As Integer
        Dim vAmount As Double
        Dim vSumAmount As Double

        On Error GoTo ErrDescription

        If Me.ListViewItem.Items.Count > 0 Then
            For i = 0 To Me.ListViewItem.Items.Count - 1
                vAmount = Me.ListViewItem.Items(i).SubItems(4).Text
                vSumAmount = vSumAmount + vAmount
            Next
            Me.TBItemAmount.Text = Format(vSumAmount, "##,##0.00")
            Me.TBBillBalance.Text = Format(vSumAmount, "##,##0.00")
            Me.TBBalanceAmount.Text = Format(vSumAmount, "##,##0.00")
            Me.TBPayAmount.Text = Format(vSumAmount, "##,##0.00")
        Else
            Me.TBItemAmount.Text = Format(0, "##,##0.00")
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBQty.TextChanged

    End Sub


    Public Sub CalcPayAmount()
        Dim vMemTotalAmount As Double
        Dim vRemainAmount As Double
        Dim vMemCashLine As Double
        Dim vMemCashAmount As Double
        Dim vMemTransAmount As Double
        Dim vMemChqAmount As Double
        Dim vMemCreditAmount As Double
        Dim vMemOtherInCome As Double
        Dim vMemOtherExpense As Double

        Dim vOtherDebt As Double
        Dim vOtherExpense As Double

        'On Error Resume Next

        If Me.TBPayAmount.Text <> "" Then
            If Me.TBOtherDebt.Text <> "" Then
                vOtherDebt = Me.TBOtherDebt.Text
            End If
            If Me.TBOtherExpense.Text <> "" Then
                vOtherExpense = Me.TBOtherExpense.Text
            End If
            Me.TBBillAmount.Text = Me.TBPayAmount.Text
            vMemTotalAmount = Me.TBPayAmount.Text
            vMemTotalAmount = vMemTotalAmount - vOtherDebt + vOtherExpense
        Else
            Me.TBBillAmount.Text = 0
            vMemTotalAmount = 0
        End If

        If Me.TBCashAmount.Text = "" And Me.ListViewCreditCard.Items.Count = 0 Then
            Me.ListViewPayDetails.Items.Clear()

            Dim listItem As New ListViewItem("ยอดเงินที่จะได้รับ :")
            listItem.SubItems.Add(Format(vMemTotalAmount, "##,##0.00"))
            listItem.SubItems.Add(0)
            Me.ListViewPayDetails.Items.Add(listItem)

            Dim listItem1 As New ListViewItem("ยอดผลต่าง :")
            listItem1.SubItems.Add(Format(vMemTotalAmount, "##,##0.00"))
            listItem1.SubItems.Add(6)
            Me.ListViewPayDetails.Items.Add(listItem1)

            vRemainAmount = vMemTotalAmount

            If Me.TBOtherExpense.Text <> "" Then
                vMemOtherExpense = Me.TBOtherExpense.Text

                Dim listItem2 As New ListViewItem("รายการค่าใช้จ่ายอื่นๆ :")
                listItem2.SubItems.Add(Format(vMemOtherExpense, "##,##0.00"))
                listItem2.SubItems.Add(7)
                Me.ListViewPayDetails.Items.Add(listItem2)
            End If

            If Me.TBOtherDebt.Text <> "" Then
                vMemOtherInCome = Me.TBOtherDebt.Text

                Dim listItem3 As New ListViewItem("รายการรายได้อื่นๆ :")
                listItem3.SubItems.Add(Format(vMemOtherInCome, "##,##0.00"))
                listItem3.SubItems.Add(8)
                Me.ListViewPayDetails.Items.Add(listItem3)

            End If

        Else
            Me.ListViewPayDetails.Items.Clear()

            Dim listItem4 As New ListViewItem("ยอดเงินที่จะได้รับ :")
            listItem4.SubItems.Add(Format(vMemTotalAmount, "##,##0.00"))
            listItem4.SubItems.Add(0)
            Me.ListViewPayDetails.Items.Add(listItem4)

            If Me.TBCashAmount.Text <> "" Then
                vMemCashAmount = Me.TBCashAmount.Text

                Dim listItem5 As New ListViewItem("ยอดเงินรับ :")
                listItem5.SubItems.Add(Format(vMemCashAmount, "##,##0.00"))
                listItem5.SubItems.Add(2)
                Me.ListViewPayDetails.Items.Add(listItem5)

            End If

            If Me.ListViewCreditCard.Items.Count > 0 Then
                Dim n As Integer
                Dim vCreditTotalAmount As Double
                Dim vCreditAmount As Double

                For n = 0 To Me.ListViewCreditCard.Items.Count - 1
                    vCreditAmount = Me.ListViewCreditCard.Items(n).SubItems(4).Text()
                    vCreditTotalAmount = vCreditTotalAmount + vCreditAmount
                Next
                vMemCreditAmount = vCreditTotalAmount

                Dim listItem6 As New ListViewItem("ยอดบัตรเครดิตรับ :")
                listItem6.SubItems.Add(Format(vMemCreditAmount, "##,##0.00"))
                listItem6.SubItems.Add(5)
                Me.ListViewPayDetails.Items.Add(listItem6)

            End If

            vRemainAmount = vMemTotalAmount - vMemCashLine - vMemCashAmount - vMemTransAmount - vMemChqAmount - vMemCreditAmount

            Dim listItem7 As New ListViewItem("ยอดผลต่าง :")
            listItem7.SubItems.Add(Format(vRemainAmount, "##,##0.00"))
            listItem7.SubItems.Add(6)
            Me.ListViewPayDetails.Items.Add(listItem7)

            If Me.TBOtherExpense.Text <> "" Then
                vMemOtherExpense = Me.TBOtherExpense.Text

                Dim listItem8 As New ListViewItem("รายการค่าใช้จ่ายอื่นๆ :")
                listItem8.SubItems.Add(Format(vMemOtherExpense, "##,##0.00"))
                listItem8.SubItems.Add(7)
                Me.ListViewPayDetails.Items.Add(listItem8)

            End If

            If Me.TBOtherDebt.Text <> "" And Me.TBOtherDebt.Text <> "." Then
                vMemOtherInCome = Me.TBOtherDebt.Text

                Dim listItem9 As New ListViewItem("รายการรายได้อื่นๆ :")
                listItem9.SubItems.Add(Format(vMemOtherInCome, "##,##0.00"))
                listItem9.SubItems.Add(8)
                Me.ListViewPayDetails.Items.Add(listItem9)
            End If
        End If

        vMemPayReceiptAmount = vRemainAmount
        Me.TBBalanceAmount.Text = vRemainAmount
    End Sub

    Public Sub SaveData()
        Dim vDocNo As String
        Dim vDocDate As String
        Dim vSaleType As Integer
        Dim vTaxType As Integer
        Dim vArCode As String
        Dim vPassBillTo As String
        Dim vDueDate As String
        Dim vPayBillDate As String
        Dim vSaleCode As String
        Dim vTaxRate As Integer
        Dim vMyDescription As String
        Dim vBillType As Integer
        Dim vSumOfItemAmount As Double
        Dim vDiscountAmount As Double
        Dim vAfterDiscount As Double
        Dim vBeforeTaxAmount As Double
        Dim vTaxAmount As Double
        Dim vTotalAmount As Double
        Dim vNetDebtAmount As Double
        Dim vHomeAmount As Double
        Dim vBillBalance As Double
        Dim vGLFormat As String
        Dim vPayBillStatus As Integer
        Dim vProjectCode As String
        Dim vAllocateCode As String
        Dim vDepartCode As String
        Dim vContactCode As String
        Dim vMemBillBalance As Double
        Dim vMemDepBillBalance As Double
        Dim vMemPeriod As Integer
        Dim vSaveForm As Integer
        Dim vDepSaveForm As Integer
        Dim vMemNetDebtAmount As Double

        Dim i As Integer
        Dim n As Integer

        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vCNQty As Double
        Dim vQty As Double
        Dim vPrice As Double
        Dim vPriceAmount As Double
        Dim vPriceUnit As Double
        Dim vDiscountWord As String
        Dim vDiscountAmountSub As Double
        Dim vAmount As Double
        Dim vNetAmount As Double
        Dim vHomeAmountSub As Double
        Dim vSumOfCost As Double
        Dim vBalanceAmount As Double
        Dim vUnitCode As String
        Dim vStockType As Integer
        Dim vLineNumber As Integer
        Dim vPackingRate1 As Integer
        Dim vPackingRate2 As Integer
        Dim vItemDesc As String

        Dim vTaxDate As String
        Dim vTaxNo As String
        Dim vPeriod As Integer
        Dim vChqOnHand As Double
        Dim vCreditOnHand As Double
        Dim vChqReturn As Double
        Dim vDebtAmount As Double
        Dim vMyDescriptionTrans As String
        Dim vSource As Integer

        Dim vSumCashAmount As Double
        Dim vSumChqAmount As Double
        Dim vSumCreditAmount As Double
        Dim vChargeAmount As Double
        Dim vChangeAmount As Double
        Dim vSumBankAmount As Double
        Dim vSumOfWTax As Double
        Dim vOtherIncome As Double
        Dim vOtherExpense As Double
        Dim vExcessAmount1 As Double
        Dim vExcessAmount2 As Double
        Dim vSumOfDeposit1 As Double
        Dim vSumOfDeposit2 As Double
        Dim vSORefNo As String

        Dim vIssueType As String
        Dim vSumOfAmount As Double

        Dim vDepLineAmount As Double

        Dim vChqAmount As Double
        Dim d As Integer
        Dim vCreditAmount As Double
        Dim vSumChargeamount As Double
        Dim c As Integer

        Dim vCheckDepExist As Integer
        Dim vDepAnswer As Integer

        Dim vAccess As Integer
        Dim vCheckTax As Integer



        If vInvoiceBillStatus > 0 Then
            MsgBox("เอกสารเลขที่นี้ ถูกอ้างอิงแล้วไม่สามารถบันทึกการเปลี่ยนแปลงข้อมูลได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBItemCode.Focus()
            'Exit Sub
        End If


        If vInvoiceIsCancel = 1 Then
            MsgBox("เอกสารเลขที่นี้ ถูกยกเลิกแล้วไม่สามารถบันทึกการเปลี่ยนแปลงข้อมูลได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBItemCode.Focus()
            'Exit Sub
        End If

        If vInvoiceIsConfirm = 1 Then
            MsgBox("เอกสารเลขที่นี้ ถูกอ้างอิงแล้วไม่สามารถบันทึกการเปลี่ยนแปลงข้อมูลได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBItemCode.Focus()
            'Exit Sub
        End If

        If Me.TBMemArCode.Text = "" Then
            MsgBox("กรุณา กรอกข้อมูลลูกค้า ก่อนบันทึกข้อมูล", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBItemCode.Focus()
            'Exit Sub
        End If

        If Me.TBItemAmount.Text = "" Then
            MsgBox("มูลค่าของสินค้าต้องไม่น้อยกว่า 0 กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBItemCode.Focus()
            'Exit Sub
        End If


        If Me.ListViewItem.Items.Count = 0 Then
            MsgBox("ไม่มีรายการสินค้าในการบันทึกขาย กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBItemCode.Focus()
            'Exit Sub
        End If

        vSaleType = 0
        vBillType = 0
        vSaveForm = 1
        vDepSaveForm = 8
        vGLFormat = "B01"


        Call CalcPayAmount()

        If Me.TBBalanceAmount.Text <> "" Then
            vMemBillBalance = Me.TBBalanceAmount.Text
        End If


        vCheckTax = 1

        vArCode = Me.TBMemArCode.Text

        If vMemBillBalance <> 0 Then
            MsgBox("มูลค่าจ่ายไม่เท่ากับมูลค่าที่จะได้รับ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.PNReceiveMoney.Visible = True
            Me.TBCashAmount.Focus()
            Me.TBCashAmount.SelectAll()
            Exit Sub
        End If


        If vInvoiceIsOpen = 0 Then
            'Call GenNewDocNo()
        End If

        Try

            vQuery = "begin tran"
            Call vGetData1(vMemProfit, vQuery)


            vDocNo = Me.TBDocNo.Text
            If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                vDocDate = vMemDocDate
            Else
                vDocDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
            End If
            vTaxType = 1
            vArCode = Me.TBMemArCode.Text
            vPassBillTo = Me.TBMemArCode.Text
            If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                vDueDate = vMemDocDate
                vPayBillDate = vMemDocDate
            Else
                vDueDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                vPayBillDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
            End If

            vSaleCode = vUserID

            vQuery = "exec dbo.USP_NP_SearchCheckPeriod '" & vDocDate & "'"
            Call vGetData1(vMemProfit, vQuery)
            If pds1.Tables(0).Rows.Count > 0 Then
                vMemPeriod = pds1.Tables(0).Rows(0)("periodno").ToString
            End If

            vTaxRate = Me.TBMemTaxRate.Text
            vMyDescription = ""

            If Me.TBItemAmount.Text <> "" Then
                vSumOfItemAmount = Me.TBItemAmount.Text
            End If


            vDiscountAmount = 0
            vDiscountWord = ""
            vSumOfDeposit1 = 0

            vAfterDiscount = Me.TBItemAmount.Text
            vNetDebtAmount = Me.TBItemAmount.Text
            vHomeAmount = Me.TBItemAmount.Text
            vBillBalance = 0
            vMemNetDebtAmount = Me.TBItemAmount.Text

            If Me.TBCashAmount.Text <> "" Then
                vSumCashAmount = Me.TBCashAmount.Text
            End If
            vSumChqAmount = 0
            If Me.ListViewCreditCard.Items.Count > 0 Then
                For c = 0 To Me.ListViewCreditCard.Items.Count - 1
                    vCreditAmount = Me.ListViewCreditCard.Items(c).SubItems(4).Text
                    vSumCreditAmount = vSumCreditAmount + vCreditAmount
                    If Me.ListViewCreditCard.Items(c).SubItems(8).Text <> "" Then
                        vChargeAmount = Me.ListViewCreditCard.Items(c).SubItems(8).Text
                        vSumChargeamount = vSumChargeamount + vChargeAmount
                    End If
                Next
            End If
            vChangeAmount = 0
            vSumBankAmount = 0
            vSumOfWTax = 0
            If Me.TBOtherDebt.Text <> "" Then
                vOtherIncome = Me.TBOtherDebt.Text
            End If
            If Me.TBOtherExpense.Text <> "" Then
                vOtherExpense = Me.TBOtherExpense.Text
            End If
            If Me.TBOverMoneyInv.Text <> "" Then
                vExcessAmount1 = Me.TBOverMoneyInv.Text
            End If
            If Me.TBOverMoney.Text <> "" Then
                vExcessAmount2 = Me.TBOverMoney.Text
            End If
            vPayBillStatus = 0
            vBeforeTaxAmount = ((Me.TBItemAmount.Text * 100) / (100 + vTaxRate))
            vTaxAmount = vNetDebtAmount - ((Me.TBItemAmount.Text * 100) / (100 + vTaxRate))
            vTotalAmount = Me.TBItemAmount.Text
            vSumOfDeposit2 = ((vSumOfDeposit1 * 100) / (100 + vTaxRate))

            vProjectCode = ""
            If Me.CMBMemDepartment.Text <> "" Then
                vDepartCode = vb6.Left(Me.CMBMemDepartment.Text, vb6.InStr(Me.CMBMemDepartment.Text, "/") - 1)
            End If
            vContactCode = ""
            vAllocateCode = ""

            vQuery = "exec dbo.USP_NP_InsertUpdateARInvoice '" & vDocNo & "','','" & vDocDate & "','" & vUserID & "'," & vTaxType & ",'" & vArCode & "','" & vPassBillTo & "','" & vDueDate & "','" & vPayBillDate & "','" & vSaleCode & "'," & vTaxRate & ",'" & vMyDescription & "'," & vBillType & "," & vSumOfItemAmount & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vAfterDiscount & "," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalAmount & "," & vNetDebtAmount & "," & vHomeAmount & "," & vBillBalance & ",'" & vGLFormat & "','" & vProjectCode & "','" & vAllocateCode & "','" & vDepartCode & "','" & vContactCode & "'," & vSumCashAmount & "," & vSumChqAmount & "," & vSumCreditAmount & "," & vChargeAmount & "," & vChangeAmount & "," & vSumBankAmount & "," & vSumOfWTax & "," & vOtherIncome & "," & vOtherExpense & "," & vExcessAmount1 & "," & vExcessAmount2 & "," & vSumOfDeposit1 & "," & vSumOfDeposit2 & ",'" & vSORefNo & "' "
            Call vGetData1(vMemProfit, vQuery)


            For n = 0 To Me.ListViewItem.Items.Count - 1
                vItemCode = Me.ListViewItem.Items(n).SubItems(1).Text
                vItemName = Me.ListViewItem.Items(n).SubItems(2).Text
                vWHCode = Me.ListViewItem.Items(n).SubItems(8).Text
                vShelfCode = Me.ListViewItem.Items(n).SubItems(9).Text
                If Me.ListViewItem.Items(n).SubItems(3).Text <> "" Then
                    vCNQty = Me.ListViewItem.Items(n).SubItems(3).Text
                    vQty = Me.ListViewItem.Items(n).SubItems(3).Text
                End If
                If Me.ListViewItem.Items(n).SubItems(7).Text <> "" Then
                    vPriceAmount = Me.ListViewItem.Items(n).SubItems(7).Text
                End If
                If vQty <> 0 Then
                    vPriceUnit = vPriceAmount / vQty
                End If
                vPrice = vPriceUnit
                vDiscountWord = ""
                vDiscountAmountSub = 0
                If Me.ListViewItem.Items(n).SubItems(7).Text <> "" Then
                    vAmount = Me.ListViewItem.Items(n).SubItems(7).Text
                End If
                If vTaxType = 0 Then
                    vNetAmount = Me.ListViewItem.Items(n).SubItems(7).Text
                    vHomeAmountSub = Me.ListViewItem.Items(n).SubItems(7).Text
                ElseIf vTaxType = 1 Then
                    vNetAmount = ((vAmount * 100) / (100 + vTaxRate))
                    vHomeAmountSub = ((vAmount * 100) / (100 + vTaxRate))
                Else
                    vNetAmount = Me.ListViewItem.Items(n).SubItems(7).Text
                    vHomeAmountSub = Me.ListViewItem.Items(n).SubItems(7).Text
                End If

                vSumOfCost = 0
                If Me.ListViewItem.Items(n).SubItems(7).Text <> "" Then
                    vBalanceAmount = Me.ListViewItem.Items(n).SubItems(7).Text
                End If
                vUnitCode = Me.ListViewItem.Items(n).SubItems(5).Text
                vStockType = Me.ListViewItem.Items(n).SubItems(12).Text
                vLineNumber = n
                vPackingRate1 = Me.ListViewItem.Items(n).SubItems(10).Text
                vPackingRate2 = Me.ListViewItem.Items(n).SubItems(11).Text
                vItemDesc = ""

                vQuery = "exec dbo.USP_NP_InsertBCARInvoiceSub '" & vDocNo & "'," & vTaxType & ",'" & vItemCode & "','" & vDocDate & "','" & vArCode & "','" & vSaleCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "'," & vCNQty & "," & vQty & "," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmountSub & "," & vAmount & "," & vNetAmount & "," & vHomeAmountSub & "," & vSumOfCost & "," & vBalanceAmount & ",'" & vUnitCode & "'," & vStockType & "," & vLineNumber & ",'" & vProjectCode & "','" & vAllocateCode & "','" & vDepartCode & "'," & vPackingRate1 & "," & vPackingRate2 & "," & vTaxRate & ",'" & vItemDesc & "' "
                Call vGetData1(vMemProfit, vQuery)

                vQuery = "exec dbo.USP_NP_InsertItemProcessStock '" & vItemCode & "'," & vInvoiceIsOpen & ""
                Call vGetData1(vMemProfit, vQuery)

            Next

            Dim vGLBookName As String

            vTaxDate = vDocDate
            vTaxNo = vDocNo
            vPeriod = Month(vMemDocDate)

            vMyDescriptionTrans = "ขายสดให้กับ" & Me.TBMemArName.Text
            vGLBookName = "ขายสินค้าสด"
            vSource = 6

            'vQuery = "exec  dbo.USP_PC_InsertUpdateOutPutTax '" & vDocNo & "','" & vDocDate & "'," & vSource & ",'" & vTaxDate & "','" & vTaxNo & "','" & vArCode & "'," & vTaxRate & "," & vBeforeTaxAmount & "," & vTaxAmount & ",'" & vUserID & "','" & vGLBookName & "' "
            'Call vGetData1(vMemProfit, vQuery)

            If vInvoiceIsOpen = 0 Then
                If Me.ListViewCreditCard.Items.Count > 0 Then 'Invoice Receive CreditCard
                    Dim vBankCode As String
                    Dim vCreditCardNo As String
                    Dim vReceiveDate As String
                    Dim vCreditCardDueDate As String
                    Dim vStatus As Integer
                    Dim vBankBranchCode As String
                    Dim vCreditCardAmount As Double
                    Dim vCreditCardLineAmount As Double
                    Dim vMyDescriptionCredit As String
                    Dim vCreditType As String
                    Dim vConfirmNo As String
                    Dim vChargeWord As Double
                    Dim vCreditChargeAmount As Double

                    Dim vLineNumberCredit As Integer
                    Dim vPAYAMOUNTCredit As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoCredit As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String

                    Dim t As Integer

                    For t = 0 To Me.ListViewCreditCard.Items.Count - 1
                        vBankCode = Me.ListViewCreditCard.Items(t).SubItems(2).Text
                        vCreditCardNo = Me.ListViewCreditCard.Items(t).SubItems(0).Text

                        If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                            vReceiveDate = vMemDocDate
                        Else
                            vReceiveDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                        End If

                        vCreditCardDueDate = Me.ListViewCreditCard.Items(t).SubItems(5).Text
                        vStatus = 0
                        vBankBranchCode = Me.ListViewCreditCard.Items(t).SubItems(3).Text
                        vCreditCardLineAmount = Me.ListViewCreditCard.Items(t).SubItems(4).Text
                        vMyDescriptionCredit = "รับชำระหนี้"
                        vCreditType = Microsoft.VisualBasic.Left(Me.ListViewCreditCard.Items(t).SubItems(6).Text, InStr(Me.ListViewCreditCard.Items(t).SubItems(6).Text, "/") - 1)
                        vConfirmNo = Me.ListViewCreditCard.Items(t).SubItems(1).Text
                        vCreditChargeAmount = Me.ListViewCreditCard.Items(t).SubItems(8).Text
                        vCreditCardAmount = vCreditCardLineAmount + vCreditChargeAmount

                        If Me.ListViewCreditCard.Items(t).SubItems(7).Text <> "" Then
                            vChargeWord = Me.ListViewCreditCard.Items(t).SubItems(7).Text
                        End If

                        vLineNumberCredit = t
                        vPAYAMOUNTCredit = Me.ListViewCreditCard.Items(t).SubItems(4).Text
                        vPaymentType = 1
                        vRefNoCredit = Me.ListViewCreditCard.Items(t).SubItems(0).Text

                        If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                            vRefDate = vMemDocDate
                        Else
                            vRefDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                        End If

                        vCHQTOTALAMOUNT = Me.ListViewCreditCard.Items(t).SubItems(4).Text
                        vPayChqState = 0
                        vTransBankDate = "NULL"

                        vQuery = "exec dbo.USP_NP_InsertCreditCard   '" & vBankCode & "','" & vCreditCardNo & "','" & vDocNo & "','" & vArCode & "','" & vReceiveDate & "','" & vCreditCardDueDate & "'," & vStatus & ",'" & vBankBranchCode & "'," & vCreditCardAmount & ",'" & vMyDescriptionCredit & "','" & vCreditType & "','" & vConfirmNo & "'," & vCreditChargeAmount & " "
                        Call vGetData1(vMemProfit, vQuery)

                        vQuery = "exec dbo.USP_NP_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumberCredit & "," & vPAYAMOUNTCredit & "," & vPaymentType & ",'" & vGLBookName & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "','" & vRefNoCredit & "','" & vRefDate & "'," & vCreditCardAmount & "," & vPayChqState & ",'" & vBankCode & "','" & vBankBranchCode & "'," & vTransBankDate & ",'" & vCreditType & "'," & vCreditChargeAmount & ",'" & vConfirmNo & "','" & vChargeWord & "'," & vSaveForm & " "
                        Call vGetData1(vMemProfit, vQuery)

                    Next

                End If


                If Me.TBOverMoneyInv.Text <> "" And vExcessAmount1 <> 0 Then 'Invoice Receive OverMoney
                    Dim vLineNumberOver As Integer
                    Dim vPAYAMOUNTOver As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoOver As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String
                    Dim vCreditChargeAmount As Double
                    Dim vCreditType As String
                    Dim vBankCode As String
                    Dim vBankBranchCode As String
                    Dim vMyDescriptionOver As String
                    Dim vConfirmNo As String

                    vLineNumberOver = 4
                    vPAYAMOUNTOver = Me.TBOverMoneyInv.Text
                    vPaymentType = 5
                    vRefNoOver = "NULL"

                    If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                        vRefDate = vMemDocDate
                    Else
                        vRefDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                    End If

                    vCHQTOTALAMOUNT = 0
                    vPayChqState = 0
                    vTransBankDate = "NULL"
                    vCreditChargeAmount = 0
                    vCreditType = "NULL"
                    vBankCode = "NULL"
                    vBankBranchCode = "NULL"
                    vMyDescriptionOver = "รับชำระหนี้"
                    vConfirmNo = "NULL"

                    vQuery = "exec dbo.USP_NP_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumber & "," & vPAYAMOUNTOver & "," & vPaymentType & ",'" & vGLBookName & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "'," & vRefNoOver & ",'" & vRefDate & "'," & vCHQTOTALAMOUNT & "," & vPayChqState & "," & vBankCode & "," & vBankBranchCode & "," & vTransBankDate & "," & vCreditType & "," & vCreditChargeAmount & "," & vConfirmNo & ",''," & vSaveForm & " "
                    Call vGetData1(vMemProfit, vQuery)

                End If

                If Me.TBOverMoney.Text <> "" And vExcessAmount2 <> 0 Then 'Invoice Receive OverMoney
                    Dim vLineNumberOver As Integer
                    Dim vPAYAMOUNTOver As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoOver As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String
                    Dim vCreditChargeAmount As Double
                    Dim vCreditType As String
                    Dim vBankCode As String
                    Dim vBankBranchCode As String
                    Dim vMyDescriptionOver As String
                    Dim vConfirmNo As String

                    vLineNumberOver = 2
                    vPAYAMOUNTOver = Me.TBOverMoney.Text
                    vPaymentType = 6
                    vRefNoOver = "NULL"

                    If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                        vRefDate = vMemDocDate
                    Else
                        vRefDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                    End If

                    vCHQTOTALAMOUNT = 0
                    vPayChqState = 0
                    vTransBankDate = "NULL"
                    vCreditChargeAmount = 0
                    vCreditType = "NULL"
                    vBankCode = "NULL"
                    vBankBranchCode = "NULL"
                    vMyDescriptionOver = "รับชำระหนี้"
                    vConfirmNo = "NULL"

                    vQuery = "exec dbo.USP_NP_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumber & "," & vPAYAMOUNTOver & "," & vPaymentType & ",'" & vGLBookName & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "'," & vRefNoOver & ",'" & vRefDate & "'," & vCHQTOTALAMOUNT & "," & vPayChqState & "," & vBankCode & "," & vBankBranchCode & "," & vTransBankDate & "," & vCreditType & "," & vCreditChargeAmount & "," & vConfirmNo & ",''," & vSaveForm & " "
                    Call vGetData1(vMemProfit, vQuery)

                End If

                If Me.TBOtherExpense.Text <> "" And vOtherExpense <> 0 Then 'Invoice Receive OtherExpense
                    Dim vLineNumberExpense As Integer
                    Dim vPAYAMOUNTExpense As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoExpense As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String
                    Dim vCreditChargeAmount As Double
                    Dim vCreditType As String
                    Dim vBankCode As String
                    Dim vBankBranchCode As String
                    Dim vMyDescriptionExpense As String
                    Dim vConfirmNo As String

                    vLineNumberExpense = 0
                    vPAYAMOUNTExpense = Me.TBOtherExpense.Text
                    vPaymentType = 8
                    vRefNoExpense = "NULL"

                    If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                        vRefDate = vMemDocDate
                    Else
                        vRefDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                    End If

                    vCHQTOTALAMOUNT = 0
                    vPayChqState = 0
                    vTransBankDate = "NULL"
                    vCreditChargeAmount = 0
                    vCreditType = "NULL"
                    vBankCode = "NULL"
                    vBankBranchCode = "NULL"
                    vMyDescriptionExpense = "รับชำระหนี้"
                    vConfirmNo = "NULL"

                    vQuery = "exec dbo.USP_NP_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumber & "," & vPAYAMOUNTExpense & "," & vPaymentType & ",'" & vGLBookName & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "'," & vRefNoExpense & ",'" & vRefDate & "'," & vCHQTOTALAMOUNT & "," & vPayChqState & "," & vBankCode & "," & vBankBranchCode & "," & vTransBankDate & "," & vCreditType & "," & vCreditChargeAmount & "," & vConfirmNo & ",''," & vSaveForm & " "
                    Call vGetData1(vMemProfit, vQuery)
                End If

                If Me.TBOtherDebt.Text <> "" And vOtherIncome <> 0 Then 'Invoice Receive OtherDebt
                    Dim vLineNumberDebt As Integer
                    Dim vPAYAMOUNTDebt As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoDebt As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String
                    Dim vCreditChargeAmount As Double
                    Dim vCreditType As String
                    Dim vBankCode As String
                    Dim vBankBranchCode As String
                    Dim vMyDescriptionDebt As String
                    Dim vConfirmNo As String

                    vLineNumberDebt = 1
                    vPAYAMOUNTDebt = Me.TBOtherDebt.Text
                    vPaymentType = 7
                    vRefNoDebt = "NULL"
                    vRefDate = vMemDocDate

                    If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                        vRefDate = vMemDocDate
                    Else
                        vRefDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                    End If

                    vCHQTOTALAMOUNT = 0
                    vPayChqState = 0
                    vTransBankDate = "NULL"
                    vCreditChargeAmount = 0
                    vCreditType = "NULL"
                    vBankCode = "NULL"
                    vBankBranchCode = "NULL"
                    vMyDescriptionDebt = "รับชำระหนี้"
                    vConfirmNo = "NULL"

                    vQuery = "exec dbo.USP_NP_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumber & "," & vPAYAMOUNTDebt & "," & vPaymentType & ",'" & vGLBookName & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "'," & vRefNoDebt & ",'" & vRefDate & "'," & vCHQTOTALAMOUNT & "," & vPayChqState & "," & vBankCode & "," & vBankBranchCode & "," & vTransBankDate & "," & vCreditType & "," & vCreditChargeAmount & "," & vConfirmNo & ",''," & vSaveForm & " "
                    Call vGetData1(vMemProfit, vQuery)

                End If

                If Me.TBCashAmount.Text <> "" And vSumCashAmount <> 0 Then 'Invoice Receive Cash
                    Dim vLineNumberCash As Integer
                    Dim vPAYAMOUNTCash As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoCash As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String
                    Dim vCreditChargeAmount As Double
                    Dim vCreditType As String
                    Dim vBankCode As String
                    Dim vBankBranchCode As String
                    Dim vMyDescriptionCash As String
                    Dim vConfirmNo As String

                    vLineNumberCash = 2
                    vPAYAMOUNTCash = Me.TBCashAmount.Text
                    vPaymentType = 0
                    vRefNoCash = "NULL"
                    vRefDate = "NULL"
                    vCHQTOTALAMOUNT = 0
                    vPayChqState = 0
                    vTransBankDate = "NULL"
                    vCreditChargeAmount = 0
                    vCreditType = "NULL"
                    vBankCode = "NULL"
                    vBankBranchCode = "NULL"
                    vMyDescriptionCash = "รับชำระหนี้"
                    vConfirmNo = "NULL"

                    vQuery = "exec dbo.USP_NP_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumber & "," & vPAYAMOUNTCash & "," & vPaymentType & ",'" & vGLBookName & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "'," & vRefNoCash & "," & vRefDate & "," & vCHQTOTALAMOUNT & "," & vPayChqState & "," & vBankCode & "," & vBankBranchCode & "," & vTransBankDate & "," & vCreditType & "," & vCreditChargeAmount & "," & vConfirmNo & ",'' ," & vSaveForm & ""
                    Call vGetData1(vMemProfit, vQuery)

                End If

                vChqOnHand = vSumChqAmount
                vCreditOnHand = vSumCreditAmount
                vChqReturn = 0
                vDebtAmount = 0

            End If


            If vInvoiceIsOpen = 1 Then

                Dim vPayCash As Double
                Dim vPayCredit As Double
                Dim vPayChq As Double
                Dim vPayTrans As Double
                Dim vPayIncome As Double
                Dim vPayExpense As Double
                Dim vPayExcess1 As Double
                Dim vPayExcess2 As Double

                Dim vCheckCash As Double
                Dim vCheckCredit As Double
                Dim vCheckChq As Double
                Dim vCheckTrans As Double
                Dim vCheckIncome As Double
                Dim vCheckExpense As Double
                Dim vCheckExcess1 As Double
                Dim vCheckExcess2 As Double

                If Me.ListViewCreditCard.Items.Count > 0 Then
                    Dim vBankCode As String
                    Dim vCreditCardNo As String
                    Dim vReceiveDate As String
                    Dim vCreditCardDueDate As String
                    Dim vStatus As Integer
                    Dim vBankBranchCode As String
                    Dim vCreditCardAmount As Double
                    Dim vCreditCardLineAmount As Double
                    Dim vMyDescriptionCredit As String
                    Dim vCreditType As String
                    Dim vConfirmNo As String
                    Dim vChargeWord As Double
                    Dim vCreditChargeAmount As Double

                    Dim vLineNumberCredit As Integer
                    Dim vPAYAMOUNTCredit As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoCredit As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String

                    Dim t As Integer

                    For t = 0 To Me.ListViewCreditCard.Items.Count - 1
                        vBankCode = Me.ListViewCreditCard.Items(t).SubItems(2).Text
                        vCreditCardNo = Me.ListViewCreditCard.Items(t).SubItems(0).Text

                        If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                            vReceiveDate = vMemDocDate
                        Else
                            vReceiveDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                        End If

                        vCreditCardDueDate = Me.ListViewCreditCard.Items(t).SubItems(5).Text
                        vStatus = 0
                        vBankBranchCode = Me.ListViewCreditCard.Items(t).SubItems(3).Text
                        vCreditCardLineAmount = Me.ListViewCreditCard.Items(t).SubItems(4).Text
                        vMyDescriptionCredit = "รับชำระหนี้"
                        vCreditType = Microsoft.VisualBasic.Left(Me.ListViewCreditCard.Items(t).SubItems(6).Text, InStr(Me.ListViewCreditCard.Items(t).SubItems(6).Text, "/") - 1)
                        vConfirmNo = Me.ListViewCreditCard.Items(t).SubItems(1).Text
                        vCreditChargeAmount = Me.ListViewCreditCard.Items(t).SubItems(8).Text
                        vCreditCardAmount = vCreditCardLineAmount + vCreditChargeAmount

                        If Me.ListViewCreditCard.Items(t).SubItems(7).Text <> "" Then
                            vChargeWord = Me.ListViewCreditCard.Items(t).SubItems(7).Text
                        End If

                        vLineNumberCredit = 6
                        vPAYAMOUNTCredit = Me.ListViewCreditCard.Items(t).SubItems(4).Text
                        vPaymentType = 1
                        vRefNoCredit = Me.ListViewCreditCard.Items(t).SubItems(0).Text
                        vRefDate = vMemDocDate

                        If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                            vRefDate = vMemDocDate
                        Else
                            vRefDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                        End If

                        vCHQTOTALAMOUNT = Me.ListViewCreditCard.Items(t).SubItems(4).Text
                        vPayChqState = 0
                        vTransBankDate = "NULL"

                        vQuery = "exec dbo.USP_NP_DeleteCreditCard   '" & vDocNo & "' "
                        Call vGetData1(vMemProfit, vQuery)

                        vQuery = "exec dbo.USP_NP_InsertCreditCard   '" & vBankCode & "','" & vCreditCardNo & "','" & vDocNo & "','" & vArCode & "','" & vReceiveDate & "','" & vCreditCardDueDate & "'," & vStatus & ",'" & vBankBranchCode & "'," & vCreditCardAmount & ",'" & vMyDescriptionCredit & "','" & vCreditType & "','" & vConfirmNo & "'," & vCreditChargeAmount & " "
                        Call vGetData1(vMemProfit, vQuery)

                        vQuery = "exec dbo.USP_NP_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumber & "," & vPAYAMOUNTCredit & "," & vPaymentType & ",'" & vMyDescriptionCredit & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "','" & vRefNoCredit & "','" & vRefDate & "'," & vCreditCardAmount & "," & vPayChqState & ",'" & vBankCode & "','" & vBankBranchCode & "'," & vTransBankDate & ",'" & vCreditType & "'," & vCreditChargeAmount & ",'" & vConfirmNo & "','" & vChargeWord & "'," & vSaveForm & ""
                        Call vGetData1(vMemProfit, vQuery)
                    Next

                End If

                If Me.TBOverMoneyInv.Text <> "" And vExcessAmount1 <> 0 Then
                    Dim vLineNumberOver As Integer
                    Dim vPAYAMOUNTOver As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoOver As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String
                    Dim vCreditChargeAmount As Double
                    Dim vCreditType As String
                    Dim vBankCode As String
                    Dim vBankBranchCode As String
                    Dim vMyDescriptionOver As String
                    Dim vConfirmNo As String

                    vLineNumberOver = 4
                    vPAYAMOUNTOver = Me.TBOverMoneyInv.Text
                    vPaymentType = 5
                    vRefNoOver = "NULL"

                    If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                        vRefDate = vMemDocDate
                    Else
                        vRefDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                    End If

                    vCHQTOTALAMOUNT = 0
                    vPayChqState = 0
                    vTransBankDate = "NULL"
                    vCreditChargeAmount = 0
                    vCreditType = "NULL"
                    vBankCode = "NULL"
                    vBankBranchCode = "NULL"
                    vMyDescriptionOver = "รับชำระหนี้"
                    vConfirmNo = "NULL"

                    vQuery = "exec dbo.USP_PC_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumber & "," & vPAYAMOUNTOver & "," & vPaymentType & ",'" & vMyDescriptionOver & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "'," & vRefNoOver & ",'" & vRefDate & "'," & vCHQTOTALAMOUNT & "," & vPayChqState & "," & vBankCode & "," & vBankBranchCode & "," & vTransBankDate & "," & vCreditType & "," & vCreditChargeAmount & "," & vConfirmNo & ",''," & vSaveForm & " "
                    Call vGetData1(vMemProfit, vQuery)

                End If

                If Me.TBOverMoney.Text <> "" And vExcessAmount2 <> 0 Then
                    Dim vLineNumberOver As Integer
                    Dim vPAYAMOUNTOver As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoOver As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String
                    Dim vCreditChargeAmount As Double
                    Dim vCreditType As String
                    Dim vBankCode As String
                    Dim vBankBranchCode As String
                    Dim vMyDescriptionOver As String
                    Dim vConfirmNo As String

                    vLineNumberOver = 2
                    vPAYAMOUNTOver = Me.TBOverMoney.Text
                    vPaymentType = 6
                    vRefNoOver = "NULL"

                    If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                        vRefDate = vMemDocDate
                    Else
                        vRefDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                    End If

                    vCHQTOTALAMOUNT = 0
                    vPayChqState = 0
                    vTransBankDate = "NULL"
                    vCreditChargeAmount = 0
                    vCreditType = "NULL"
                    vBankCode = "NULL"
                    vBankBranchCode = "NULL"
                    vMyDescriptionOver = "รับชำระหนี้"
                    vConfirmNo = "NULL"

                    vQuery = "exec dbo.USP_PC_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumber & "," & vPAYAMOUNTOver & "," & vPaymentType & ",'" & vMyDescriptionOver & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "'," & vRefNoOver & ",'" & vRefDate & "'," & vCHQTOTALAMOUNT & "," & vPayChqState & "," & vBankCode & "," & vBankBranchCode & "," & vTransBankDate & "," & vCreditType & "," & vCreditChargeAmount & "," & vConfirmNo & ",''," & vSaveForm & " "
                    Call vGetData1(vMemProfit, vQuery)

                End If

                If Me.TBOtherExpense.Text <> "" And vOtherExpense <> 0 Then
                    Dim vLineNumberExpense As Integer
                    Dim vPAYAMOUNTExpense As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoExpense As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String
                    Dim vCreditChargeAmount As Double
                    Dim vCreditType As String
                    Dim vBankCode As String
                    Dim vBankBranchCode As String
                    Dim vMyDescriptionExpense As String
                    Dim vConfirmNo As String

                    vLineNumberExpense = 0
                    vPAYAMOUNTExpense = Me.TBOtherExpense.Text
                    vPaymentType = 8
                    vRefNoExpense = "NULL"

                    If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                        vRefDate = vMemDocDate
                    Else
                        vRefDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                    End If

                    vCHQTOTALAMOUNT = 0
                    vPayChqState = 0
                    vTransBankDate = "NULL"
                    vCreditChargeAmount = 0
                    vCreditType = "NULL"
                    vBankCode = "NULL"
                    vBankBranchCode = "NULL"
                    vMyDescriptionExpense = "รับชำระหนี้"
                    vConfirmNo = "NULL"

                    vQuery = "exec dbo.USP_PC_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumber & "," & vPAYAMOUNTExpense & "," & vPaymentType & ",'" & vMyDescriptionExpense & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "'," & vRefNoExpense & ",'" & vRefDate & "'," & vCHQTOTALAMOUNT & "," & vPayChqState & "," & vBankCode & "," & vBankBranchCode & "," & vTransBankDate & "," & vCreditType & "," & vCreditChargeAmount & "," & vConfirmNo & ",''," & vSaveForm & " "
                    Call vGetData1(vMemProfit, vQuery)

                End If

                If Me.TBOtherDebt.Text <> "" And vOtherIncome <> 0 Then
                    Dim vLineNumberDebt As Integer
                    Dim vPAYAMOUNTDebt As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoDebt As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String
                    Dim vCreditChargeAmount As Double
                    Dim vCreditType As String
                    Dim vBankCode As String
                    Dim vBankBranchCode As String
                    Dim vMyDescriptionDebt As String
                    Dim vConfirmNo As String

                    vLineNumberDebt = 1
                    vPAYAMOUNTDebt = Me.TBOtherDebt.Text
                    vPaymentType = 7
                    vRefNoDebt = "NULL"

                    If vb6.Left(vb6.Year(vMemDocDate), 2) = "20" Then
                        vRefDate = vMemDocDate
                    Else
                        vRefDate = vb6.Day(vMemDocDate) & "/" & vb6.Month(vMemDocDate) & "/" & vb6.Year(vMemDocDate) - 543
                    End If

                    vCHQTOTALAMOUNT = 0
                    vPayChqState = 0
                    vTransBankDate = "NULL"
                    vCreditChargeAmount = 0
                    vCreditType = "NULL"
                    vBankCode = "NULL"
                    vBankBranchCode = "NULL"
                    vMyDescriptionDebt = "รับชำระหนี้"
                    vConfirmNo = "NULL"

                    vQuery = "exec dbo.USP_PC_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumber & "," & vPAYAMOUNTDebt & "," & vPaymentType & ",'" & vMyDescriptionDebt & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "'," & vRefNoDebt & ",'" & vRefDate & "'," & vCHQTOTALAMOUNT & "," & vPayChqState & "," & vBankCode & "," & vBankBranchCode & "," & vTransBankDate & "," & vCreditType & "," & vCreditChargeAmount & "," & vConfirmNo & ",''," & vSaveForm & " "
                    Call vGetData1(vMemProfit, vQuery)

                End If


                If Me.TBCashAmount.Text <> "" And vSumCashAmount <> 0 Then
                    Dim vLineNumberCash As Integer
                    Dim vPAYAMOUNTCash As Double
                    Dim vPaymentType As Integer
                    Dim vRefNoCash As String
                    Dim vRefDate As String
                    Dim vCHQTOTALAMOUNT As Double
                    Dim vPayChqState As Integer
                    Dim vTransBankDate As String
                    Dim vCreditChargeAmount As Double
                    Dim vCreditType As String
                    Dim vBankCode As String
                    Dim vBankBranchCode As String
                    Dim vMyDescriptionCash As String
                    Dim vConfirmNo As String

                    vLineNumberCash = 2
                    vPAYAMOUNTCash = Me.TBCashAmount.Text
                    vPaymentType = 0
                    vRefNoCash = "NULL"
                    vRefDate = "NULL"
                    vCHQTOTALAMOUNT = 0
                    vPayChqState = 0
                    vTransBankDate = "NULL"
                    vCreditChargeAmount = 0
                    vCreditType = "NULL"
                    vBankCode = "NULL"
                    vBankBranchCode = "NULL"
                    vMyDescriptionCash = "รับชำระหนี้"
                    vConfirmNo = "NULL"

                    vQuery = "exec dbo.USP_PC_InsertBCRecMoney '" & vDocNo & "','" & vDocDate & "','" & vArCode & "'," & vLineNumber & "," & vPAYAMOUNTCash & "," & vPaymentType & ",'" & vMyDescriptionCash & "','" & vDepartCode & "','" & vSaleCode & "','" & vProjectCode & "'," & vRefNoCash & "," & vRefDate & "," & vCHQTOTALAMOUNT & "," & vPayChqState & "," & vBankCode & "," & vBankBranchCode & "," & vTransBankDate & "," & vCreditType & "," & vCreditChargeAmount & "," & vConfirmNo & ",''," & vSaveForm & " "
                    Call vGetData1(vMemProfit, vQuery)

                End If


                vChqOnHand = vSumChqAmount
                vCreditOnHand = vSumCreditAmount
                vChqReturn = 0
                vDebtAmount = 0

                vQuery = "exec dbo.USP_PC_UpdateBCARDebtTable '" & vArCode & "'," & vMemPeriod & "," & -1 * vMemNetDebtAmount & "," & vChqOnHand & "," & vCreditOnHand & "," & vChqReturn & " "
                Call vGetData1(vMemProfit, vQuery)
                vQuery = "exec dbo.USP_PC_UpdateBCAR '" & vArCode & "'," & -1 * vMemNetDebtAmount & "," & vChqOnHand & "," & vCreditOnHand & "," & vChqReturn & " "
                Call vGetData1(vMemProfit, vQuery)

                vQuery = "exec dbo.USP_PC_UpdateBCARDebtTable '" & vArCode & "'," & vMemPeriod & "," & vDebtAmount & "," & vChqOnHand & "," & vCreditOnHand & "," & vChqReturn & " "
                Call vGetData1(vMemProfit, vQuery)

                vQuery = "exec dbo.USP_PC_UpdateBCAR '" & vArCode & "'," & vDebtAmount & "," & vChqOnHand & "," & vCreditOnHand & "," & vChqReturn & " "
                Call vGetData1(vMemProfit, vQuery)
            End If


            vQuery = "commit tran"
            Call vGetData1(vMemProfit, vQuery)

            MsgBox("บันทึกข้อมูลเลขที่เอกสาร " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")

            'Call PrintInvoice(vDocNo)
            'Call ClearScreen()
            'Call ClearItem()
            'Me.TBItemCode.Text = ""
            'Call NewDocNo()
            'Call GenNewDocNo()


        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            vQuery = "rollback tran"
            Call vGetData1(vMemProfit, vQuery)
        End Try

    End Sub

    Private Sub BTNCreditUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCreditUpdate.Click
        Dim vCreditPayAmount As Double
        Dim i As Integer
        Dim vAnswer As Integer
        Dim vCreditBalance As Double
        Dim vCreditConfirm As String
        Dim vChargeAmount As Double

        On Error GoTo ErrDescription

        If Me.TBCreditCard.Text <> "" And Me.TBConfirmNo.Text <> "" And Me.TBCreditAmount.Text <> "" Then 'And Me.CMBBank.Text <> "" And Me.CMBBranch.Text <> "" And Me.CMBCreditType.Text <> "" Then

            Dim vChqMultiAmount As Double
            Dim vPayAmount As Double

            If Me.ListViewCreditCard.Items.Count > 0 Then
                For i = 0 To Me.ListViewCreditCard.Items.Count - 1
                    Dim vCheckCredit As String
                    Dim vCheckBank As String

                    vCheckCredit = Me.ListViewCreditCard.Items(i).SubItems(0).Text
                    vCheckBank = Me.ListViewCreditCard.Items(i).SubItems(2).Text

                    If Me.TBCreditCard.Text = vCheckCredit And Me.CMBBank.Text = vCheckBank Then
                        vAnswer = MsgBox("มีเลขที่บัตรเครดิตของธนาคารนี้อยู่แล้ว ต้องการปรับมูลค่าบัตรเครดิตหรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
                        If vAnswer = 6 Then
                            Me.ListViewCreditCard.Items(i).SubItems(4).Text = Me.TBCreditAmount.Text
                            GoTo LineCalAmount
                        Else
                            Exit Sub
                        End If
                    End If
                    vCreditPayAmount = Me.TBCreditAmount.Text
                    If Me.TBCreditAmount.Text <> "" Then
                        vCreditBalance = Me.TBCreditAmount.Text
                    Else
                        vCreditBalance = 0
                    End If
                    If Me.TBCharge.Text = "" Then
                        vChargeAmount = 0
                    Else
                        Dim vCharge As Integer
                        If Microsoft.VisualBasic.InStr(Me.TBCharge.Text, "%") > 0 Then
                            vCharge = Microsoft.VisualBasic.Left(Me.TBCharge.Text, Microsoft.VisualBasic.InStr(Me.TBCharge.Text, "%") - 1)
                            vChargeAmount = (vCreditPayAmount * vCharge) / 100
                        Else
                            vChargeAmount = Me.TBCharge.Text
                        End If
                    End If
                    vCreditConfirm = Me.TBConfirmNo.Text


                    Dim vCreditList As New ListViewItem(Me.TBCreditCard.Text)
                    vCreditList.SubItems.Add(vCreditConfirm)
                    vCreditList.SubItems.Add(Me.CMBBank.Text)
                    vCreditList.SubItems.Add(Me.CMBBranch.Text)
                    vCreditList.SubItems.Add(vCreditPayAmount)
                    vCreditList.SubItems.Add(vMemDocDate)
                    vCreditList.SubItems.Add(Me.CMBCreditType.Text)
                    vCreditList.SubItems.Add(Me.TBCharge.Text)
                    vCreditList.SubItems.Add(vChargeAmount)
                    vCreditList.SubItems.Add(vCreditPayAmount + vChargeAmount)
                    Me.ListViewCreditCard.Items.Add(vCreditList)
                Next
            Else
                vCreditPayAmount = Me.TBCreditAmount.Text
                If Me.TBCreditAmount.Text <> "" Then
                    vCreditBalance = Me.TBCreditAmount.Text
                Else
                    vCreditBalance = 0
                End If
                If Me.TBCharge.Text = "" Then
                    vChargeAmount = 0
                Else
                    Dim vCharge As Integer
                    If Microsoft.VisualBasic.InStr(Me.TBCharge.Text, "%") > 0 Then
                        vCharge = Microsoft.VisualBasic.Left(Me.TBCharge.Text, Microsoft.VisualBasic.InStr(Me.TBCharge.Text, "%") - 1)
                        vChargeAmount = (vCreditPayAmount * vCharge) / 100
                    Else
                        vChargeAmount = Me.TBCharge.Text
                    End If
                End If
                vCreditConfirm = Me.TBConfirmNo.Text

                Dim vCreditList As New ListViewItem(Me.TBCreditCard.Text)
                vCreditList.SubItems.Add(vCreditConfirm)
                vCreditList.SubItems.Add(Me.CMBBank.Text)
                vCreditList.SubItems.Add(Me.CMBBranch.Text)
                vCreditList.SubItems.Add(vCreditPayAmount)
                vCreditList.SubItems.Add(vMemDocDate)
                vCreditList.SubItems.Add(Me.CMBCreditType.Text)
                vCreditList.SubItems.Add(Me.TBCharge.Text)
                vCreditList.SubItems.Add(vChargeAmount)
                vCreditList.SubItems.Add(vCreditPayAmount + vChargeAmount)
                Me.ListViewCreditCard.Items.Add(vCreditList)
            End If


LineCalAmount:

            'Call CreditClear()
            Call CalcPayAmount()
        Else
            MsgBox("กรุณากรอก เลขที่บัตรเครดิต ธนาคาร สาขาของธนาคาร ประเภทบัตร และมูลค่าบัตรเครดิตให้ครบถ้วนก่อนบันทึกชำระบัตรเครดิต", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBCreditCard.Focus()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TCPayMoney_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TCPayMoney.SelectedIndexChanged

        If Me.TCPayMoney.SelectedIndex = 0 Then
            Me.BTNCreditUpdate.Visible = False
            Me.BTNCreditDelete.Visible = False
        End If

        If Me.TCPayMoney.SelectedIndex = 1 Then
            Me.BTNCreditUpdate.Visible = True
            Me.BTNCreditDelete.Visible = True
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Call savedata()
    End Sub

    Private Sub BTNPayOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPayOK.Click
        'On Error Resume Next


        Call CalcPayAmount()
        If vMemPayReceiptAmount <> 0 Then
            MsgBox("มูลค่าจ่ายไม่เท่ากับมูลค่าที่จะได้รับ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBCashAmount.Focus()
            Me.TBCashAmount.SelectAll()
            Exit Sub
        End If

        Me.PNReceiveMoney.Visible = False
    End Sub

    Private Sub TBCashAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBCashAmount.TextChanged
        Call CalcPayAmount()
    End Sub

    Private Sub BTNPayCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPayCancel.Click
        Me.PNReceiveMoney.Visible = False
    End Sub
End Class