Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic


Public Class FormPriceStructure
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vDocno As String
    Dim vBillDiscAmount As Double
    Dim vFollowDisc1Amount As Double
    Dim vFollowDisc2Amount As Double
    Dim vFollowDisc3Amount As Double
    Dim vRebateAmount As Double
    Dim vSpecialDiscAmount As Double
    Dim vMissProfitAmount As Double
    Dim vSendAmount As Double
    Dim vCustSendAmount As Double
    Dim vAdvertiseAmount As Double
    Dim vMarketAmount As Double
    Dim vTaxAmount As Double
    Dim vInstallAmount As Double
    Dim vServiceAmount As Double
    Dim vPointAmount As Double
    Dim vTargetAmount As Double
    Dim vGiftAmount As Double
    Dim vCommissionAmount As Double
    Dim vPriceSetAmount As Double
    Dim vBegProfitAmount As Double
    Dim vDiscMemberAmount As Double


    Dim vBillDisc As Double
    Dim vDiscFollow1 As Double
    Dim vDiscFollow2 As Double
    Dim vCommission As Double
    Dim vGift As Double
    Dim vTarget As Double
    Dim vPoint As Double
    Dim vService As Double
    Dim vInstall As Double
    Dim vTax As Double
    Dim vMarket As Double
    Dim vAdvertise As Double
    Dim vCustSend As Double
    Dim vSend As Double
    Dim vDiscMissProfit As Double
    Dim vDiscSpecial As Double
    Dim vDiscRebate As Double
    Dim vDiscFollow3 As Double
    Dim vDiscMember As Double

    Dim vBillDisc1 As Double
    Dim vDiscFollow11 As Double
    Dim vDiscFollow21 As Double
    Dim vCommission1 As Double
    Dim vGift1 As Double
    Dim vTarget1 As Double
    Dim vPoint1 As Double
    Dim vService1 As Double
    Dim vInstall1 As Double
    Dim vTax1 As Double
    Dim vMarket1 As Double
    Dim vAdvertise1 As Double
    Dim vCustSend1 As Double
    Dim vSend1 As Double
    Dim vDiscMissProfit1 As Double
    Dim vDiscSpecial1 As Double
    Dim vDiscRebate1 As Double
    Dim vDiscFollow31 As Double
    Dim vDiscMember1 As Double

    Dim vBegProfit1 As Double
    Dim vBegProfit2 As Double


    Dim vCheckItem As String
    Dim vCheckUnitCode As String
    Dim vSelectPriceListIndex As Integer

    Dim vIsOpen As Integer
    Dim vIsConfirm As Integer
    Dim vIsCancel As Integer

    Dim vCheckDocnoExist As Integer

    Dim vDepartmentCode As String
    Dim vBrandCode As String
    Dim vSearchItemCode As String
    Dim vCheckBuyer As String
    Dim vCheckUserID As String

    Dim vOldDocNO As String


    Private Sub BTN101Click()
        GB101.Visible = False
        GB103.Top = 103
        GB103.Height = 381
        ListView103.Height = 350
        Me.CBAll.Checked = False
        Me.CBProduct.Checked = False
        If GB102.Visible = True Then
            Me.CBStructure.Checked = True
        End If
    End Sub
    Private Sub BTN101Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN101Close.Click
        On Error GoTo ErrDescription

        Call BTN101Click()

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub CBAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBAll.CheckedChanged
        On Error GoTo ErrDescription

        If CBAll.Checked = True Then
            Call GB101Form()
            Call GB102Form()
            Call GB103Form()

            Me.CBProduct.Checked = True
            Me.CBStructure.Checked = True
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub BTN102Click()
        GB102.Visible = False
        GB101.Width = 1006
        ListView102.Width = 995

        BTN101Close.Top = 11
        BTN101Close.Left = 975

        GB103.Width = 1006
        ListView103.Width = 995

        Me.CBAll.Checked = False
        Me.CBStructure.Checked = False
        If GB101.Visible = True Then
            Me.CBProduct.Checked = True
        End If
    End Sub
    Private Sub BTN102Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN102Close.Click
        On Error GoTo ErrDescription

        Call BTN102Click()

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub CBProduct_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBProduct.CheckedChanged
        On Error GoTo ErrDescription

        If CBProduct.Checked = True Then
            If GB102.Visible = True Then
                Call GB101Form()
                Call GB103Form()
            Else
                Call GB101NewForm()
                Call GB103NewForm1()
            End If
        Else
            Call BTN101Click()
        End If
        If CBProduct.Checked = True And CBStructure.Checked = True Then
            CBAll.Checked = True
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub CBStructure_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBStructure.CheckedChanged
        On Error GoTo ErrDescription

        If CBStructure.Checked = True Then
            If GB101.Visible = True Then
                Call GB101Form()
                Call GB102Form()
                Call GB103Form()
            Else
                GB102.Visible = True
                Call GB103NewForm()
            End If
        Else
            Call BTN102Click()
        End If
        If CBStructure.Checked = True And CBProduct.Checked = True Then
            CBAll.Checked = True
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub FormPriceStructure_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CBAll.Checked = False
        Me.CBProduct.Checked = False
        Me.CBStructure.Checked = True
        Me.DocDate.Text = Now.Date
        Call InitializeDataBase()
        Call InitializeDataBaseBPlus()
        Call GetDepartment()
        Call GetBrand()
        Call GetUserLogIN()
        Me.DateUpdate.Text = DateAdd(DateInterval.Day, 1, Date.Now)

        'vQuery = "select isnull(emp_name,'')+'     '+isnull(emp_surnme,'') as PersonName from bcnp.dbo.vw_np_UserAutorityProgram where code = '" & vUserID & "' "
        'da = New SqlDataAdapter(vQuery, vConnectionBPlus)
        'ds = New DataSet
        'da.Fill(ds, "UserID")
        'dt = ds.Tables("UserID")


        'If dt.Rows.Count > 0 Then
        Me.TextBoxBuyer.Text = vUserID
        vCheckUserID = vUserID
        'Else
        '    MsgBox("ไม่มีชื่อพนักงานของ  UserID = " & vUserID & " นี้ กรุณาตรวจสอบหรือแจ้งแผนกคอมฯ", MsgBoxStyle.Critical, "Send Error ")
        '    Exit Sub
        'End If

        Me.BTNGenDocno.Focus()
    End Sub

    Public Sub GetDepartment()
        Dim i As Integer

        On Error GoTo ErrDescription

        vQuery = "exec dbo.usp_ps_departmentlist"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Department")
        dt = ds.Tables("Department")

        For i = 0 To dt.Rows.Count - 1
            Me.CMBDepartment.Items.Add(Trim(dt.Rows(i).Item("DepartmentCode")) & "//" & Trim(dt.Rows(i).Item("Department")))
        Next

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Public Sub GetBrand()
        Dim i As Integer

        On Error GoTo ErrDescription

        vQuery = "exec dbo.usp_ps_brandlist"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Brand")
        dt = ds.Tables("Brand")

        For i = 0 To dt.Rows.Count - 1
            Me.CMBBrand.Items.Add(Trim(dt.Rows(i).Item("BrandCode")) & "//" & Trim(dt.Rows(i).Item("Brand")))
        Next

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Public Sub GB101Form()
        GB101.Visible = True
        GB101.Top = 103
        GB101.Left = 6
        GB101.Height = 208
        GB101.Width = 612
        ListView102.Top = 69
        ListView102.Left = 6
        ListView102.Height = 131
        ListView102.Width = 600
        BTN101Close.Top = 11
        BTN101Close.Left = 581
    End Sub
    Public Sub GB101NewForm()
        GB101.Visible = True
        GB101.Top = 103
        GB101.Left = 6
        GB101.Height = 208
        GB101.Width = 1006
        ListView102.Top = 69
        ListView102.Left = 6
        ListView102.Height = 131
        ListView102.Width = 995
        BTN101Close.Top = 11
        BTN101Close.Left = 975
    End Sub

    Public Sub GB102Form()
        GB102.Visible = True
        GB102.Top = 68
        GB102.Left = 624
        GB102.Height = 454
        GB102.Width = 388
    End Sub
    Public Sub GB103Form()
        GB103.Visible = True
        GB103.Top = 317
        GB103.Left = 6
        GB103.Height = 174
        GB103.Width = 612
        ListView103.Top = 25
        ListView103.Left = 6
        ListView103.Height = 142
        ListView103.Width = 600
    End Sub
    Public Sub GB103NewForm()
        GB101.Visible = False
        GB103.Top = 103
        GB103.Height = 381
        ListView103.Height = 350
        GB103.Width = 612
        ListView103.Width = 600
        Me.CBAll.Checked = False
        Me.CBProduct.Checked = False
        If GB102.Visible = True Then
            Me.CBStructure.Checked = True
        End If
    End Sub
    Public Sub GB103NewForm1()
        GB103.Visible = True
        GB103.Top = 309
        GB103.Left = 6
        GB103.Height = 174
        GB103.Width = 1006
        ListView103.Top = 25
        ListView103.Left = 6
        ListView103.Height = 142
        ListView103.Width = 995
    End Sub

    Private Sub ClearItemData()
        Me.TXTItemCode.Text = ""
        Me.LBLItemName.Text = ""
        Me.CMBUnit.Items.Clear()
        Me.TXTDO.Text = ""
        Me.TXTPriceSet.Text = ""
    End Sub


    Public Sub GetDocNo()

        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_PS_NewDocno"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "NewDocno")
        dt = ds.Tables("NewDocno")

        If dt.Rows.Count > 0 Then
            vDocno = Trim(dt.Rows(0).Item("newdocno"))
            vIsOpen = 0
            vIsCancel = 0
            vIsConfirm = 0
            Me.DocDate.Text = Date.Now
            Me.TextBoxDocno.Text = vDocno


            Me.PB101.Visible = True
            Me.PB102.Visible = False
            Me.PB103.Visible = False
        Else
            MsgBox("ไม่สามารถรันเลขที่เอกสารได้ กรุณาลองกดปุ่ม รันเลขที่เอกสารอีกครั้ง", MsgBoxStyle.Critical, "Send Error")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub GetUserLogIN()

        On Error GoTo ErrDescription

        'vQuery = "select isnull(emp_name,'')+'     '+isnull(emp_surnme,'') as PersonName from bcnp.dbo.vw_np_UserAutorityProgram where code = '" & vUserID & "' "
        'da = New SqlDataAdapter(vQuery, vConnectionBPlus)
        'ds = New DataSet
        'da.Fill(ds, "UserID")
        'dt = ds.Tables("UserID")

        'If dt.Rows.Count > 0 Then
        Me.TextBoxBuyer.Text = vUserID
        'Else
        'MsgBox("ไม่มีชื่อพนักงานของ  UserID = " & vUserID & " นี้ กรุณาตรวจสอบหรือแจ้งแผนกคอมฯ", MsgBoxStyle.Critical, "Send Error ")
        'Exit Sub
        'End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub BTNGenDocno_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNGenDocno.Click
        On Error GoTo ErrDescription

        If Me.TextBoxDocno.Text <> "" Then
            vOldDocNO = Me.TextBoxDocno.Text
        End If

        Call GetDocNo()
        Me.TextBoxDocno.Text = vDocno

        If vOldDocNO <> "" And Me.ListView103.Items.Count > 0 Then
            Call SaveData()
            Me.TXTItemCode.Text = ""
            Me.LBLItemName.Text = ""
            Me.TXTDO.Text = ""
            Me.TXTPriceSet.Text = ""
            Me.CMBUnit.Items.Clear()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub CMBDepartment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBDepartment.SelectedIndexChanged
        Dim vTemplateList As ListViewItem
        Dim i As Integer
        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vBillDiscRes As Double
        Dim vFollowDisc1Res As Double
        Dim vFollowDisc2Res As Double
        Dim vFollowDisc3Res As Double
        Dim vRebateRes As Double
        Dim vSpecialDiscRes As Double
        Dim vMissProfitRes As Double
        Dim vSendRes As Double
        Dim vCustSendRes As Double
        Dim vAdvertiseRes As Double
        Dim vMarketRes As Double
        Dim vTaxRes As Double
        Dim vInstallRes As Double
        Dim vServiceRes As Double
        Dim vPointRes As Double
        Dim vTargetRes As Double
        Dim vGiftRes As Double
        Dim vCommissionRes As Double
        Dim vBegProfitAmountRes As Double
        Dim vInterestsAmountRes As Double
        Dim vProfitAmountRes As Double
        Dim vAddCashRes As Double
        Dim vAddCreditRes As Double
        Dim vDiscount1Res As Double
        Dim vDiscMemberRes As Double

        On Error GoTo ErrDescription

        If Me.CMBDepartment.Text <> "" Then
            If InStr(Trim(Me.CMBDepartment.Text), "//") > 0 Then
                vDepartmentCode = Microsoft.VisualBasic.Left(Trim(Me.CMBDepartment.Text), InStr(Trim(Me.CMBDepartment.Text), "//") - 1)
            Else
                vDepartmentCode = Trim(Me.CMBDepartment.Text)
            End If
            If Me.CMBBrand.Text <> "" And Me.TXTSearchItemCode.Text = "" Then
                If InStr(Trim(Me.CMBBrand.Text), "//") > 0 Then
                    vBrandCode = Microsoft.VisualBasic.Left(Trim(Me.CMBBrand.Text), InStr(Trim(Me.CMBBrand.Text), "//") - 1)
                Else
                    vBrandCode = Trim(Me.CMBBrand.Text)
                End If
                vQuery = "exec dbo.USP_PS_Template '" & vDepartmentCode & "','" & vBrandCode & "','',0 "
            ElseIf Me.CMBBrand.Text = "" And Me.TXTSearchItemCode.Text <> "" Then
                vSearchItemCode = TXTSearchItemCode.Text
                vQuery = "exec dbo.USP_PS_Template '" & vDepartmentCode & "','','" & vSearchItemCode & "',0 "
            ElseIf Me.CMBBrand.Text <> "" And Me.TXTSearchItemCode.Text <> "" Then
                If InStr(Trim(Me.CMBBrand.Text), "//") > 0 Then
                    vBrandCode = Microsoft.VisualBasic.Left(Trim(Me.CMBBrand.Text), InStr(Trim(Me.CMBBrand.Text), "//") - 1)
                Else
                    vBrandCode = Trim(Me.CMBBrand.Text)
                End If
                vSearchItemCode = TXTSearchItemCode.Text
                vQuery = "exec dbo.USP_PS_Template '" & vDepartmentCode & "','" & vBrandCode & "','" & vSearchItemCode & "',0 "
                Else
                    vQuery = ""
                End If
                If vQuery = "" Then
                    Exit Sub
                End If

                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "Template")
                dt = ds.Tables("Template")
                If dt.Rows.Count > 0 Then
                    Me.ListView102.Items.Clear()
                    Me.Cursor = Cursors.WaitCursor
                    For i = 0 To dt.Rows.Count - 1
                        vTemplateList = ListView102.Items.Add(Trim(dt.Rows(i).Item("ItemCode")))
                        vTemplateList.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("ItemName"))
                        vTemplateList.SubItems.Add(2).Text = Trim(dt.Rows(i).Item("BuyUnitSale"))

                        vDO = Trim(dt.Rows(i).Item("D/O"))
                        vTemplateList.SubItems.Add(3).Text = Format(vDO, "##,##0.00")

                        vPriceSet = Trim(dt.Rows(i).Item("PriceSet"))
                        vTemplateList.SubItems.Add(4).Text = Format(vPriceSet, "##,##0.00")

                        vTemplateList.SubItems.Add(5).Text = Trim(dt.Rows(i).Item("DiscountBillWord"))
                        vBillDiscRes = Trim(dt.Rows(i).Item("DiscountBillAmount"))
                        vTemplateList.SubItems.Add(6).Text = Format(vBillDiscRes, "##,##0.00")

                        vTemplateList.SubItems.Add(7).Text = Trim(dt.Rows(i).Item("DiscountFollow1Word"))
                        vFollowDisc1Res = Trim(dt.Rows(i).Item("DiscountFollow1After"))
                        vTemplateList.SubItems.Add(8).Text = Format(vFollowDisc1Res, "##,##0.00")

                        vTemplateList.SubItems.Add(9).Text = Trim(dt.Rows(i).Item("DiscountFollow2Word"))
                        vFollowDisc2Res = Trim(dt.Rows(i).Item("DiscountFollow2After"))
                        vTemplateList.SubItems.Add(10).Text = Format(vFollowDisc2Res, "##,##0.00")

                        vTemplateList.SubItems.Add(11).Text = Trim(dt.Rows(i).Item("DiscountFollow3Word"))
                        vFollowDisc3Res = Trim(dt.Rows(i).Item("DiscountFollow3After"))
                        vTemplateList.SubItems.Add(12).Text = Format(vFollowDisc3Res, "##,##0.00")

                        vTemplateList.SubItems.Add(13).Text = Trim(dt.Rows(i).Item("DiscountRebateWord"))
                        vRebateRes = Trim(dt.Rows(i).Item("DiscountRebateAfter"))
                        vTemplateList.SubItems.Add(14).Text = Format(vRebateRes, "##,##0.00")

                        vTemplateList.SubItems.Add(15).Text = Trim(dt.Rows(i).Item("DiscountSpecialword"))
                        vSpecialDiscRes = Trim(dt.Rows(i).Item("NetCost"))
                        vTemplateList.SubItems.Add(16).Text = Format(vSpecialDiscRes, "##,##0.00")

                        vTemplateList.SubItems.Add(17).Text = Trim(dt.Rows(i).Item("LossBudgetWord"))
                        vMissProfitRes = Trim(dt.Rows(i).Item("LossBudgetAfter"))
                        vTemplateList.SubItems.Add(18).Text = Format(vMissProfitRes, "##,##0.00")

                        vTemplateList.SubItems.Add(19).Text = Trim(dt.Rows(i).Item("TransferInWord"))
                        vSendRes = Trim(dt.Rows(i).Item("TransferInAfter"))
                        vTemplateList.SubItems.Add(20).Text = Format(vSendRes, "##,##0.00")

                        vTemplateList.SubItems.Add(21).Text = Trim(dt.Rows(i).Item("TransferOutWord"))
                        vCustSendRes = Trim(dt.Rows(i).Item("TransferOutAfter"))
                        vTemplateList.SubItems.Add(22).Text = Format(vCustSendRes, "##,##0.00")

                        vTemplateList.SubItems.Add(23).Text = Trim(dt.Rows(i).Item("AdvertiseWord"))
                        vAdvertiseRes = Trim(dt.Rows(i).Item("AdvertiseAfter"))
                        vTemplateList.SubItems.Add(24).Text = Format(vAdvertiseRes, "##,##0.00")

                        vTemplateList.SubItems.Add(25).Text = Trim(dt.Rows(i).Item("MarketingBudgetWord"))
                        vMarketRes = Trim(dt.Rows(i).Item("MarketingBudgetAfter"))
                        vTemplateList.SubItems.Add(26).Text = Format(vMarketRes, "##,##0.00")

                        vTemplateList.SubItems.Add(27).Text = Trim(dt.Rows(i).Item("VatWord"))
                        vTaxRes = Trim(dt.Rows(i).Item("VatAfter"))
                        vTemplateList.SubItems.Add(28).Text = Format(vTaxRes, "##,##0.00")

                        vTemplateList.SubItems.Add(29).Text = Trim(dt.Rows(i).Item("SetupWord"))
                        vInstallRes = Trim(dt.Rows(i).Item("SetupAfter"))
                        vTemplateList.SubItems.Add(30).Text = Format(vInstallRes, "##,##0.00")

                        vTemplateList.SubItems.Add(31).Text = Trim(dt.Rows(i).Item("ServiceWord"))
                        vServiceRes = Trim(dt.Rows(i).Item("MarketCost"))
                        vTemplateList.SubItems.Add(32).Text = Format(vServiceRes, "##,##0.00")

                        vTemplateList.SubItems.Add(33).Text = Trim(dt.Rows(i).Item("Pointword"))
                        vPointRes = Trim(dt.Rows(i).Item("PointAfter"))
                        vTemplateList.SubItems.Add(34).Text = Format(vPointRes, "##,##0.00")

                        vTemplateList.SubItems.Add(53).Text = Trim(dt.Rows(0).Item("MemberDiscountWord"))
                        vDiscMemberRes = Trim(dt.Rows(0).Item("MemberDiscountAfter"))
                        vTemplateList.SubItems.Add(54).Text = Format(vDiscMemberRes, "##,##0.00")

                        vTemplateList.SubItems.Add(35).Text = Trim(dt.Rows(i).Item("TargetWord"))
                        vTargetRes = Trim(dt.Rows(i).Item("TargetAfter"))
                        vTemplateList.SubItems.Add(36).Text = Format(vTargetRes, "##,##0.00")

                        vTemplateList.SubItems.Add(37).Text = Trim(dt.Rows(i).Item("PremiumWord"))
                        vGiftRes = Trim(dt.Rows(i).Item("PremiumAfter"))
                        vTemplateList.SubItems.Add(38).Text = Format(vGiftRes, "##,##0.00")

                        vTemplateList.SubItems.Add(39).Text = Trim(dt.Rows(i).Item("CommissionWord"))
                        vCommissionRes = Trim(dt.Rows(i).Item("CommissionAfter"))
                        vTemplateList.SubItems.Add(40).Text = Format(vCommissionRes, "##,##0.00")

                        vTemplateList.SubItems.Add(41).Text = Format(Int(Trim(dt.Rows(i).Item("GrossProfitPercent"))), "##,##0.00")
                        vBegProfitAmountRes = Trim(dt.Rows(i).Item("GrossProfitAmount"))
                        vTemplateList.SubItems.Add(42).Text = Format(vBegProfitAmountRes, "##,##0.00")

                        vTemplateList.SubItems.Add(43).Text = Format(Int(Trim(dt.Rows(i).Item("InterestStockPercent"))), "##,##0.00")
                        vInterestsAmountRes = Trim(dt.Rows(i).Item("InterestStockAmount"))
                        vTemplateList.SubItems.Add(44).Text = Format(vInterestsAmountRes, "##,##0.00")

                        vTemplateList.SubItems.Add(45).Text = Format(Int(Trim(dt.Rows(i).Item("ProfitPercent"))), "##,##0.00")
                        vProfitAmountRes = Trim(dt.Rows(i).Item("ProfitAmount"))
                        vTemplateList.SubItems.Add(46).Text = Format(vProfitAmountRes, "##,##0.00")

                        vTemplateList.SubItems.Add(47).Text = Trim(dt.Rows(i).Item("TransferValueWord"))
                        vAddCashRes = Trim(dt.Rows(i).Item("TransferValueAmount"))
                        vTemplateList.SubItems.Add(48).Text = Format(vAddCashRes, "##,##0.00")

                        vTemplateList.SubItems.Add(49).Text = Trim(dt.Rows(i).Item("CreditValueWord"))
                        vAddCreditRes = Trim(dt.Rows(i).Item("CreditValueAmount"))
                        vTemplateList.SubItems.Add(50).Text = Format(vAddCreditRes, "##,##0.00")

                        vTemplateList.SubItems.Add(51).Text = Trim(dt.Rows(i).Item("SpecialValueWord"))
                        vDiscount1Res = Trim(dt.Rows(i).Item("SpecialValueAmount"))
                        vTemplateList.SubItems.Add(52).Text = Format(vDiscount1Res, "##,##0.00")
                    Next
                    Me.Cursor = Cursors.Arrow
                    Me.ListView102.Focus()
                Else
                    MsgBox("ไม่พบข้อมูล Template ของแผนกและยี่ห้อที่ต้องการดูข้อมูล", MsgBoxStyle.Information, "send Information")
                End If
            End If

ErrDescription:
            If Err.Description <> "" Then
                MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            End If
    End Sub

    Private Sub CMBBrand_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBBrand.SelectedIndexChanged
        Dim vTemplateList As ListViewItem
        Dim i As Integer
        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vBillDiscRes As Double
        Dim vFollowDisc1Res As Double
        Dim vFollowDisc2Res As Double
        Dim vFollowDisc3Res As Double
        Dim vRebateRes As Double
        Dim vSpecialDiscRes As Double
        Dim vMissProfitRes As Double
        Dim vSendRes As Double
        Dim vCustSendRes As Double
        Dim vAdvertiseRes As Double
        Dim vMarketRes As Double
        Dim vTaxRes As Double
        Dim vInstallRes As Double
        Dim vServiceRes As Double
        Dim vPointRes As Double
        Dim vTargetRes As Double
        Dim vGiftRes As Double
        Dim vCommissionRes As Double
        Dim vBegProfitAmountRes As Double
        Dim vInterestsAmountRes As Double
        Dim vProfitAmountRes As Double
        Dim vAddCashRes As Double
        Dim vAddCreditRes As Double
        Dim vDiscount1Res As Double
        Dim vDiscMemberRes As Double


        On Error GoTo ErrDescription

        If Me.CMBBrand.Text <> "" Then
            If InStr(Trim(Me.CMBBrand.Text), "//") > 0 Then
                vBrandCode = Microsoft.VisualBasic.Left(Trim(Me.CMBBrand.Text), InStr(Trim(Me.CMBBrand.Text), "//") - 1)
            Else
                vBrandCode = Trim(Me.CMBBrand.Text)
            End If
            If Me.CMBDepartment.Text <> "" And Me.TXTSearchItemCode.Text = "" Then
                If InStr(Trim(Me.CMBDepartment.Text), "//") > 0 Then
                    vDepartmentCode = Microsoft.VisualBasic.Left(Trim(Me.CMBDepartment.Text), InStr(Trim(Me.CMBDepartment.Text), "//") - 1)
                Else
                    vDepartmentCode = Trim(Me.CMBDepartment.Text)
                End If
                vQuery = "exec dbo.USP_PS_Template '" & vDepartmentCode & "','" & vBrandCode & "','',0 "
            ElseIf Me.CMBDepartment.Text = "" And Me.TXTSearchItemCode.Text <> "" Then
                vSearchItemCode = TXTSearchItemCode.Text
                vQuery = "exec dbo.USP_PS_Template '','" & vBrandCode & "','" & vSearchItemCode & "',0 "
            ElseIf Me.CMBDepartment.Text <> "" And Me.TXTSearchItemCode.Text <> "" Then
                If InStr(Trim(Me.CMBDepartment.Text), "//") > 0 Then
                    vDepartmentCode = Microsoft.VisualBasic.Left(Trim(Me.CMBDepartment.Text), InStr(Trim(Me.CMBDepartment.Text), "//") - 1)
                Else
                    vDepartmentCode = Trim(Me.CMBDepartment.Text)
                End If
                vSearchItemCode = TXTSearchItemCode.Text
                vQuery = "exec dbo.USP_PS_Template '" & vDepartmentCode & "','" & vBrandCode & "','" & vSearchItemCode & "',0 "
            End If

                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "Template")
                dt = ds.Tables("Template")
                If dt.Rows.Count > 0 Then
                    Me.ListView102.Items.Clear()
                    Me.Cursor = Cursors.WaitCursor
                    For i = 0 To dt.Rows.Count - 1
                        vTemplateList = ListView102.Items.Add(Trim(dt.Rows(i).Item("ItemCode")))
                        vTemplateList.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("ItemName"))
                        vTemplateList.SubItems.Add(2).Text = Trim(dt.Rows(i).Item("BuyUnitSale"))

                        vDO = Trim(dt.Rows(i).Item("D/O"))
                        vTemplateList.SubItems.Add(3).Text = Format(vDO, "##,##0.00")

                        vPriceSet = Trim(dt.Rows(i).Item("PriceSet"))
                        vTemplateList.SubItems.Add(4).Text = Format(vPriceSet, "##,##0.00")

                        vTemplateList.SubItems.Add(5).Text = Trim(dt.Rows(i).Item("DiscountBillWord"))
                        vBillDiscRes = Trim(dt.Rows(i).Item("DiscountBillAmount"))
                        vTemplateList.SubItems.Add(6).Text = Format(vBillDiscRes, "##,##0.00")

                        vTemplateList.SubItems.Add(7).Text = Trim(dt.Rows(i).Item("DiscountFollow1Word"))
                        vFollowDisc1Res = Trim(dt.Rows(i).Item("DiscountFollow1After"))
                        vTemplateList.SubItems.Add(8).Text = Format(vFollowDisc1Res, "##,##0.00")

                        vTemplateList.SubItems.Add(9).Text = Trim(dt.Rows(i).Item("DiscountFollow2Word"))
                        vFollowDisc2Res = Trim(dt.Rows(i).Item("DiscountFollow2After"))
                        vTemplateList.SubItems.Add(10).Text = Format(vFollowDisc2Res, "##,##0.00")

                        vTemplateList.SubItems.Add(11).Text = Trim(dt.Rows(i).Item("DiscountFollow3Word"))
                        vFollowDisc3Res = Trim(dt.Rows(i).Item("DiscountFollow3After"))
                        vTemplateList.SubItems.Add(12).Text = Format(vFollowDisc3Res, "##,##0.00")

                        vTemplateList.SubItems.Add(13).Text = Trim(dt.Rows(i).Item("DiscountRebateWord"))
                        vRebateRes = Trim(dt.Rows(i).Item("DiscountRebateAfter"))
                        vTemplateList.SubItems.Add(14).Text = Format(vRebateRes, "##,##0.00")

                        vTemplateList.SubItems.Add(15).Text = Trim(dt.Rows(i).Item("DiscountSpecialword"))
                        vSpecialDiscRes = Trim(dt.Rows(i).Item("NetCost"))
                        vTemplateList.SubItems.Add(16).Text = Format(vSpecialDiscRes, "##,##0.00")

                        vTemplateList.SubItems.Add(17).Text = Trim(dt.Rows(i).Item("LossBudgetWord"))
                        vMissProfitRes = Trim(dt.Rows(i).Item("LossBudgetAfter"))
                        vTemplateList.SubItems.Add(18).Text = Format(vMissProfitRes, "##,##0.00")

                        vTemplateList.SubItems.Add(19).Text = Trim(dt.Rows(i).Item("TransferInWord"))
                        vSendRes = Trim(dt.Rows(i).Item("TransferInAfter"))
                        vTemplateList.SubItems.Add(20).Text = Format(vSendRes, "##,##0.00")

                        vTemplateList.SubItems.Add(21).Text = Trim(dt.Rows(i).Item("TransferOutWord"))
                        vCustSendRes = Trim(dt.Rows(i).Item("TransferOutAfter"))
                        vTemplateList.SubItems.Add(22).Text = Format(vCustSendRes, "##,##0.00")

                        vTemplateList.SubItems.Add(23).Text = Trim(dt.Rows(i).Item("AdvertiseWord"))
                        vAdvertiseRes = Trim(dt.Rows(i).Item("AdvertiseAfter"))
                        vTemplateList.SubItems.Add(24).Text = Format(vAdvertiseRes, "##,##0.00")

                        vTemplateList.SubItems.Add(25).Text = Trim(dt.Rows(i).Item("MarketingBudgetWord"))
                        vMarketRes = Trim(dt.Rows(i).Item("MarketingBudgetAfter"))
                        vTemplateList.SubItems.Add(26).Text = Format(vMarketRes, "##,##0.00")

                        vTemplateList.SubItems.Add(27).Text = Trim(dt.Rows(i).Item("VatWord"))
                        vTaxRes = Trim(dt.Rows(i).Item("VatAfter"))
                        vTemplateList.SubItems.Add(28).Text = Format(vTaxRes, "##,##0.00")

                        vTemplateList.SubItems.Add(29).Text = Trim(dt.Rows(i).Item("SetupWord"))
                        vInstallRes = Trim(dt.Rows(i).Item("SetupAfter"))
                        vTemplateList.SubItems.Add(30).Text = Format(vInstallRes, "##,##0.00")

                        vTemplateList.SubItems.Add(31).Text = Trim(dt.Rows(i).Item("ServiceWord"))
                        vServiceRes = Trim(dt.Rows(i).Item("MarketCost"))
                        vTemplateList.SubItems.Add(32).Text = Format(vServiceRes, "##,##0.00")

                        vTemplateList.SubItems.Add(33).Text = Trim(dt.Rows(i).Item("PointWord"))
                        vPointRes = Trim(dt.Rows(i).Item("PointAfter"))
                        vTemplateList.SubItems.Add(34).Text = Format(vPointRes, "##,##0.00")

                        vTemplateList.SubItems.Add(53).Text = Trim(dt.Rows(0).Item("MemberDiscountWord"))
                        vDiscMemberRes = Trim(dt.Rows(0).Item("MemberDiscountAfter"))
                        vTemplateList.SubItems.Add(54).Text = Format(vDiscMemberRes, "##,##0.00")

                        vTemplateList.SubItems.Add(35).Text = Trim(dt.Rows(i).Item("TargetWord"))
                        vTargetRes = Trim(dt.Rows(i).Item("TargetAfter"))
                        vTemplateList.SubItems.Add(36).Text = Format(vTargetRes, "##,##0.00")

                        vTemplateList.SubItems.Add(37).Text = Trim(dt.Rows(i).Item("PremiumWord"))
                        vGiftRes = Trim(dt.Rows(i).Item("PremiumAfter"))
                        vTemplateList.SubItems.Add(38).Text = Format(vGiftRes, "##,##0.00")

                        vTemplateList.SubItems.Add(39).Text = Trim(dt.Rows(i).Item("CommissionWord"))
                        vCommissionRes = Trim(dt.Rows(i).Item("CommissionAfter"))
                        vTemplateList.SubItems.Add(40).Text = Format(vCommissionRes, "##,##0.00")

                        vTemplateList.SubItems.Add(41).Text = Format(Int(Trim(dt.Rows(i).Item("GrossProfitPercent"))), "##,##0.00")
                        vBegProfitAmountRes = Trim(dt.Rows(i).Item("GrossProfitAmount"))
                        vTemplateList.SubItems.Add(42).Text = Format(vBegProfitAmountRes, "##,##0.00")

                        vTemplateList.SubItems.Add(43).Text = Format(Int(Trim(dt.Rows(i).Item("InterestStockPercent"))), "##,##0.00")
                        vInterestsAmountRes = Trim(dt.Rows(i).Item("InterestStockAmount"))
                        vTemplateList.SubItems.Add(44).Text = Format(vInterestsAmountRes, "##,##0.00")

                        vTemplateList.SubItems.Add(45).Text = Format(Int(Trim(dt.Rows(i).Item("ProfitPercent"))), "##,##0.00")
                        vProfitAmountRes = Trim(dt.Rows(i).Item("ProfitAmount"))
                        vTemplateList.SubItems.Add(46).Text = Format(vProfitAmountRes, "##,##0.00")

                        vTemplateList.SubItems.Add(47).Text = Trim(dt.Rows(i).Item("TransferValueWord"))
                        vAddCashRes = Trim(dt.Rows(i).Item("TransferValueAmount"))
                        vTemplateList.SubItems.Add(48).Text = Format(vAddCashRes, "##,##0.00")

                        vTemplateList.SubItems.Add(49).Text = Trim(dt.Rows(i).Item("CreditValueWord"))
                        vAddCreditRes = Trim(dt.Rows(i).Item("CreditValueAmount"))
                        vTemplateList.SubItems.Add(50).Text = Format(vAddCreditRes, "##,##0.00")

                        vTemplateList.SubItems.Add(51).Text = Trim(dt.Rows(i).Item("SpecialValueWord"))
                        vDiscount1Res = Trim(dt.Rows(i).Item("SpecialValueAmount"))
                        vTemplateList.SubItems.Add(52).Text = Format(vDiscount1Res, "##,##0.00")
                    Next
                    Me.Cursor = Cursors.Arrow
                    Me.ListView102.Focus()
                Else
                    MsgBox("ไม่พบข้อมูล Template ของแผนกและยี่ห้อที่ต้องการดูข้อมูล", MsgBoxStyle.Information, "send Information")
                End If
            End If

ErrDescription:
            If Err.Description <> "" Then
                MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            End If
    End Sub

    Private Sub TXTItemCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTItemCode.Click
        'Dim vItemcode As String
        'Dim i As Integer

        'If Me.TXTItemCode.Text <> "" Then
        '    vItemcode = Trim(Me.TXTItemCode.Text)
        '    vQuery = "exec dbo.USP_PS_SearchItemDescription '" & vItemcode & "' "
        '    da = New SqlDataAdapter(vQuery, vConnection)
        '    ds = New DataSet
        '    da.Fill(ds, "Itemcode")
        '    dt = ds.Tables("Itemcode")

        '    Me.CMBUnit.Items.Clear()
        '    If dt.Rows.Count > 0 Then
        '        Me.LBLItemName.Text = Trim(dt.Rows(0).Item("itemname"))
        '        For i = 0 To dt.Rows.Count - 1
        '            Me.CMBUnit.Items.Add(Trim(dt.Rows(i).Item("unitcode")))
        '        Next
        '    Else
        '        Me.CMBUnit.Items.Clear()
        '        Me.LBLItemName.Text = ""
        '        Me.TXTDO.Text = ""
        '    End If
        'End If
    End Sub

    Private Sub TXTItemCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTItemCode.KeyDown
        'Dim vItemcode As String
        'Dim vDO As Double
        'Dim vPriceSet As Double
        'Dim vBillDiscRes As Double
        'Dim vFollowDisc1Res As Double
        'Dim vFollowDisc2Res As Double
        'Dim vFollowDisc3Res As Double
        'Dim vRebateRes As Double
        'Dim vSpecialDiscRes As Double
        'Dim vMissProfitRes As Double
        'Dim vSendRes As Double
        'Dim vCustSendRes As Double
        'Dim vAdvertiseRes As Double
        'Dim vMarketRes As Double
        'Dim vTaxRes As Double
        'Dim vInstallRes As Double
        'Dim vServiceRes As Double
        'Dim vPointRes As Double
        'Dim vTargetRes As Double
        'Dim vGiftRes As Double
        'Dim vCommissionRes As Double
        'Dim vBegProfitAmountRes As Double
        'Dim vInterestsAmountRes As Double
        'Dim vProfitAmountRes As Double
        'Dim vAddCashRes As Double
        'Dim vAddCreditRes As Double
        'Dim vDiscount1Res As Double

        'If e.KeyCode = Keys.Enter Then
        '    If Me.TXTItemCode.Text <> "" Then
        '        vItemcode = Trim(Me.TXTItemCode.Text)
        '        vQuery = "exec dbo.USP_PS_Template '','','" & vItemcode & "',1 "
        '        da = New SqlDataAdapter(vQuery, vConnection)
        '        ds = New DataSet
        '        da.Fill(ds, "Template")
        '        dt = ds.Tables("Template")
        '        Me.Cursor = Cursors.WaitCursor
        '        If dt.Rows.Count > 0 Then
        '            Me.TXTItemCode.Text = Trim(dt.Rows(0).Item("ItemCode"))
        '            Me.LBLItemName.Text = Trim(dt.Rows(0).Item("ItemName"))

        '            vDO = Trim(dt.Rows(0).Item("D/O"))
        '            Me.TXTDO.Text = Format(vDO, "##,##0.00")
        '            vPriceSet = Trim(dt.Rows(0).Item("PriceSet"))
        '            Me.TXTPriceSet.Text = Format(vPriceSet, "##,##0.00")

        '            Me.TXTBillDisc.Text = Trim(dt.Rows(0).Item("DiscountBillWord"))
        '            vBillDiscRes = Trim(dt.Rows(0).Item("DiscountBillAmount"))
        '            Me.TXTBillDiscRes.Text = Format(vBillDiscRes, "##,##0.00")

        '            Me.TXTFollowDisc1.Text = Trim(dt.Rows(0).Item("DiscountFollow1Word"))
        '            vFollowDisc1Res = Trim(dt.Rows(0).Item("DiscountFollow1After"))
        '            Me.TXTFollowDisc1Res.Text = Format(vFollowDisc1Res, "##,##0.00")

        '            Me.TXTFollowDisc2.Text = Trim(dt.Rows(0).Item("DiscountFollow2Word"))
        '            vFollowDisc2Res = Trim(dt.Rows(0).Item("DiscountFollow2After"))
        '            Me.TXTFollowDisc2Res.Text = Format(vFollowDisc2Res, "##,##0.00")

        '            Me.TXTFollowDisc3.Text = Trim(dt.Rows(0).Item("DiscountFollow3Word"))
        '            vFollowDisc3Res = Trim(dt.Rows(0).Item("DiscountFollow3After"))
        '            Me.TXTFollowDisc3Res.Text = Format(vFollowDisc3Res, "##,##0.00")

        '            Me.TXTRebate.Text = Trim(dt.Rows(0).Item("DiscountRebateWord"))
        '            vRebateRes = Trim(dt.Rows(0).Item("DiscountRebateAfter"))
        '            Me.TXTRebateRes.Text = Format(vRebateRes, "##,##0.00")

        '            Me.TXTSpecialDisc.Text = Trim(dt.Rows(0).Item("DiscountSpecialword"))
        '            vSpecialDiscRes = Trim(dt.Rows(0).Item("NetCost"))
        '            Me.TXTSpecialDiscRes.Text = Format(vSpecialDiscRes, "##,##0.00")

        '            Me.TXTMissProfit.Text = Trim(dt.Rows(0).Item("LossBudgetWord"))
        '            vMissProfitRes = Trim(dt.Rows(0).Item("LossBudgetAfter"))
        '            Me.TXTMissProfitRes.Text = Format(vMissProfitRes, "##,##0.00")

        '            Me.TXTSend.Text = Trim(dt.Rows(0).Item("TransferInWord"))
        '            vSendRes = Trim(dt.Rows(0).Item("TransferInAfter"))
        '            Me.TXTSendRes.Text = Format(vSendRes, "##,##0.00")

        '            Me.TXTCustSend.Text = Trim(dt.Rows(0).Item("TransferOutWord"))
        '            vCustSendRes = Trim(dt.Rows(0).Item("TransferOutAfter"))
        '            Me.TXTCustSendRes.Text = Format(vCustSendRes, "##,##0.00")

        '            Me.TXTAdvertise.Text = Trim(dt.Rows(0).Item("AdvertiseWord"))
        '            vAdvertiseRes = Trim(dt.Rows(0).Item("AdvertiseAfter"))
        '            Me.TXTAdvertiseRes.Text = Format(vAdvertiseRes, "##,##0.00")

        '            Me.TXTMarket.Text = Trim(dt.Rows(0).Item("MarketingBudgetWord"))
        '            vMarketRes = Trim(dt.Rows(0).Item("MarketingBudgetAfter"))
        '            Me.TXTMarketRes.Text = Format(vMarketRes, "##,##0.00")

        '            Me.TXTTax.Text = Trim(dt.Rows(0).Item("VatWord"))
        '            vTaxRes = Trim(dt.Rows(0).Item("VatAfter"))
        '            Me.TXTTaxRes.Text = Format(vTaxRes, "##,##0.00")

        '            Me.TXTInstall.Text = Trim(dt.Rows(0).Item("SetupWord"))
        '            vInstallRes = Trim(dt.Rows(0).Item("SetupAfter"))
        '            Me.TXTInstallRes.Text = Format(vInstallRes, "##,##0.00")

        '            Me.TXTService.Text = Trim(dt.Rows(0).Item("ServiceWord"))
        '            vServiceRes = Trim(dt.Rows(0).Item("MarketCost"))
        '            Me.TXTServiceRes.Text = Format(vServiceRes, "##,##0.00")

        '            Me.TXTPoint.Text = Trim(dt.Rows(0).Item("Pointword"))
        '            vPointRes = Trim(dt.Rows(0).Item("PointAfter"))
        '            Me.TXTPointRes.Text = Format(vPointRes, "##,##0.00")

        '            Me.TXTTarget.Text = Trim(dt.Rows(0).Item("TargetWord"))
        '            vTargetRes = Trim(dt.Rows(0).Item("TargetAfter"))
        '            Me.TXTTargetRes.Text = Format(vTargetRes, "##,##0.00")

        '            Me.TXTGift.Text = Trim(dt.Rows(0).Item("PremiumWord"))
        '            vGiftRes = Trim(dt.Rows(0).Item("PremiumAfter"))
        '            Me.TXTGiftRes.Text = Format(vGiftRes, "##,##0.00")

        '            Me.TXTCommission.Text = Trim(dt.Rows(0).Item("CommissionWord"))
        '            vCommissionRes = Trim(dt.Rows(0).Item("CommissionAfter"))
        '            Me.TXTCommissionRes.Text = Format(vCommissionRes, "##,##0.00")

        '            Me.TXTBegProfit.Text = Format(Int(Trim(dt.Rows(0).Item("GrossProfitPercent"))), "##,##0.00")
        '            vBegProfitAmountRes = Trim(dt.Rows(0).Item("GrossProfitAmount"))
        '            Me.TXTBegProfitAmount.Text = Format(vBegProfitAmountRes, "##,##0.00")

        '            Me.TXTInterests.Text = Format(Int(Trim(dt.Rows(0).Item("InterestStockPercent"))), "##,##0.00")
        '            vInterestsAmountRes = Trim(dt.Rows(0).Item("InterestStockAmount"))
        '            Me.TXTInterestsAmount.Text = Format(vInterestsAmountRes, "##,##0.00")

        '            Me.TXTProfit.Text = Format(Int(Trim(dt.Rows(0).Item("ProfitPercent"))), "##,##0.00")
        '            vProfitAmountRes = Trim(dt.Rows(0).Item("ProfitAmount"))
        '            Me.TXTProfitAmount.Text = Format(vProfitAmountRes, "##,##0.00")

        '            Me.TXTAddCash.Text = Trim(dt.Rows(0).Item("TransferValueWord"))
        '            vAddCashRes = Trim(dt.Rows(0).Item("TransferValueAmount"))
        '            Me.TXTAddCashRes.Text = Format(Int(Trim(dt.Rows(0).Item("TransferValueAmount"))), "##,##0.00")

        '            Me.TXTAddCredit.Text = Trim(dt.Rows(0).Item("CreditValueWord"))
        '            vAddCreditRes = Trim(dt.Rows(0).Item("CreditValueAmount"))
        '            Me.TXTAddCreditRes.Text = Format(vAddCreditRes, "##,##0.00")

        '            Me.TXTDiscount1.Text = Trim(dt.Rows(0).Item("SpecialValueWord"))
        '            vDiscount1Res = Trim(dt.Rows(0).Item("SpecialValueAmount"))
        '            Me.TXTDiscount1Res.Text = Format(vDiscount1Res, "##,##0.00")

        '            Me.CMBUnit.Items.Add(Trim(dt.Rows(0).Item("BuyUnitSale")))
        '            Me.CMBUnit.Text = Me.CMBUnit.Items(0)
        '            Me.TXTDO.Focus()
        '        Else
        '            MsgBox("ไม่พบข้อมูล Template ที่ต้องการดูข้อมูล", MsgBoxStyle.Information, "send Information")
        '        End If
        '        Me.Cursor = Cursors.Arrow
        '    End If
        'End If

    End Sub

    Private Sub TXTItemCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TXTItemCode.KeyPress

    End Sub

    Private Sub TXTItemCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTItemCode.TextChanged
        Dim vItemcode As String
        Dim vListItemCode As String
        Dim i As Integer

        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vBillDiscRes As Double
        Dim vFollowDisc1Res As Double
        Dim vFollowDisc2Res As Double
        Dim vFollowDisc3Res As Double
        Dim vRebateRes As Double
        Dim vSpecialDiscRes As Double
        Dim vMissProfitRes As Double
        Dim vSendRes As Double
        Dim vCustSendRes As Double
        Dim vAdvertiseRes As Double
        Dim vMarketRes As Double
        Dim vTaxRes As Double
        Dim vInstallRes As Double
        Dim vServiceRes As Double
        Dim vPointRes As Double
        Dim vTargetRes As Double
        Dim vGiftRes As Double
        Dim vCommissionRes As Double
        Dim vBegProfitAmountRes As Double
        Dim vInterestsAmountRes As Double
        Dim vProfitAmountRes As Double
        Dim vAddCashRes As Double
        Dim vAddCreditRes As Double
        Dim vDiscount1Res As Double
        Dim vDiscMemberRes As Double


        On Error GoTo ErrDescription

        If Me.TXTItemCode.Text <> "" Then
            vItemcode = Trim(Me.TXTItemCode.Text)
            vQuery = "exec dbo.USP_PS_Template '','','" & vItemcode & "',1 "
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Template")
            dt = ds.Tables("Template")

            Me.Cursor = Cursors.WaitCursor
            If dt.Rows.Count > 0 Then
                Me.TXTItemCode.Text = Trim(dt.Rows(0).Item("ItemCode"))
                Me.LBLItemName.Text = Trim(dt.Rows(0).Item("ItemName"))

                vDO = Trim(dt.Rows(0).Item("D/O"))
                Me.TXTDO.Text = Format(vDO, "##,##0.00")
                vPriceSet = Trim(dt.Rows(0).Item("PriceSet"))
                Me.TXTPriceSet.Text = Format(vPriceSet, "##,##0.00")

                Me.TXTBillDisc.Text = Trim(dt.Rows(0).Item("DiscountBillWord"))
                vBillDiscRes = Trim(dt.Rows(0).Item("DiscountBillAmount"))
                Me.TXTBillDiscRes.Text = Format(vBillDiscRes, "##,##0.00")

                Me.TXTFollowDisc1.Text = Trim(dt.Rows(0).Item("DiscountFollow1Word"))
                vFollowDisc1Res = Trim(dt.Rows(0).Item("DiscountFollow1After"))
                Me.TXTFollowDisc1Res.Text = Format(vFollowDisc1Res, "##,##0.00")

                Me.TXTFollowDisc2.Text = Trim(dt.Rows(0).Item("DiscountFollow2Word"))
                vFollowDisc2Res = Trim(dt.Rows(0).Item("DiscountFollow2After"))
                Me.TXTFollowDisc2Res.Text = Format(vFollowDisc2Res, "##,##0.00")

                Me.TXTFollowDisc3.Text = Trim(dt.Rows(0).Item("DiscountFollow3Word"))
                vFollowDisc3Res = Trim(dt.Rows(0).Item("DiscountFollow3After"))
                Me.TXTFollowDisc3Res.Text = Format(vFollowDisc3Res, "##,##0.00")

                Me.TXTRebate.Text = Trim(dt.Rows(0).Item("DiscountRebateWord"))
                vRebateRes = Trim(dt.Rows(0).Item("DiscountRebateAfter"))
                Me.TXTRebateRes.Text = Format(vRebateRes, "##,##0.00")

                Me.TXTSpecialDisc.Text = Trim(dt.Rows(0).Item("DiscountSpecialword"))
                vSpecialDiscRes = Trim(dt.Rows(0).Item("NetCost"))
                Me.TXTSpecialDiscRes.Text = Format(vSpecialDiscRes, "##,##0.00")

                Me.TXTMissProfit.Text = Trim(dt.Rows(0).Item("LossBudgetWord"))
                vMissProfitRes = Trim(dt.Rows(0).Item("LossBudgetAfter"))
                Me.TXTMissProfitRes.Text = Format(vMissProfitRes, "##,##0.00")

                Me.TXTSend.Text = Trim(dt.Rows(0).Item("TransferInWord"))
                vSendRes = Trim(dt.Rows(0).Item("TransferInAfter"))
                Me.TXTSendRes.Text = Format(vSendRes, "##,##0.00")

                Me.TXTCustSend.Text = Trim(dt.Rows(0).Item("TransferOutWord"))
                vCustSendRes = Trim(dt.Rows(0).Item("TransferOutAfter"))
                Me.TXTCustSendRes.Text = Format(vCustSendRes, "##,##0.00")

                Me.TXTAdvertise.Text = Trim(dt.Rows(0).Item("AdvertiseWord"))
                vAdvertiseRes = Trim(dt.Rows(0).Item("AdvertiseAfter"))
                Me.TXTAdvertiseRes.Text = Format(vAdvertiseRes, "##,##0.00")

                Me.TXTMarket.Text = Trim(dt.Rows(0).Item("MarketingBudgetWord"))
                vMarketRes = Trim(dt.Rows(0).Item("MarketingBudgetAfter"))
                Me.TXTMarketRes.Text = Format(vMarketRes, "##,##0.00")

                Me.TXTTax.Text = Trim(dt.Rows(0).Item("VatWord"))
                vTaxRes = Trim(dt.Rows(0).Item("VatAfter"))
                Me.TXTTaxRes.Text = Format(vTaxRes, "##,##0.00")

                Me.TXTInstall.Text = Trim(dt.Rows(0).Item("SetupWord"))
                vInstallRes = Trim(dt.Rows(0).Item("SetupAfter"))
                Me.TXTInstallRes.Text = Format(vInstallRes, "##,##0.00")

                Me.TXTService.Text = Trim(dt.Rows(0).Item("ServiceWord"))
                vServiceRes = Trim(dt.Rows(0).Item("MarketCost"))
                Me.TXTServiceRes.Text = Format(vServiceRes, "##,##0.00")

                Me.TXTPoint.Text = Trim(dt.Rows(0).Item("Pointword"))
                vPointRes = Trim(dt.Rows(0).Item("PointAfter"))
                Me.TXTPointRes.Text = Format(vPointRes, "##,##0.00")

                Me.TXTDiscMember.Text = Trim(dt.Rows(0).Item("MemberDiscountWord"))
                vDiscMemberRes = Trim(dt.Rows(0).Item("MemberDiscountAfter"))
                Me.TXTDiscMemberRes.Text = Format(vDiscMemberRes, "##,##0.00")

                Me.TXTTarget.Text = Trim(dt.Rows(0).Item("TargetWord"))
                vTargetRes = Trim(dt.Rows(0).Item("TargetAfter"))
                Me.TXTTargetRes.Text = Format(vTargetRes, "##,##0.00")

                Me.TXTGift.Text = Trim(dt.Rows(0).Item("PremiumWord"))
                vGiftRes = Trim(dt.Rows(0).Item("PremiumAfter"))
                Me.TXTGiftRes.Text = Format(vGiftRes, "##,##0.00")

                Me.TXTCommission.Text = Trim(dt.Rows(0).Item("CommissionWord"))
                vCommissionRes = Trim(dt.Rows(0).Item("CommissionAfter"))
                Me.TXTCommissionRes.Text = Format(vCommissionRes, "##,##0.00")

                Me.TXTBegProfit.Text = Format(Int(Trim(dt.Rows(0).Item("GrossProfitPercent"))), "##,##0.00")
                vBegProfitAmountRes = Trim(dt.Rows(0).Item("GrossProfitAmount"))
                Me.TXTBegProfitAmount.Text = Format(vBegProfitAmountRes, "##,##0.00")

                Me.TXTInterests.Text = Format(Int(Trim(dt.Rows(0).Item("InterestStockPercent"))), "##,##0.00")
                vInterestsAmountRes = Trim(dt.Rows(0).Item("InterestStockAmount"))
                Me.TXTInterestsAmount.Text = Format(vInterestsAmountRes, "##,##0.00")

                Me.TXTProfit.Text = Format(Int(Trim(dt.Rows(0).Item("ProfitPercent"))), "##,##0.00")
                vProfitAmountRes = Trim(dt.Rows(0).Item("ProfitAmount"))
                Me.TXTProfitAmount.Text = Format(vProfitAmountRes, "##,##0.00")

                Me.TXTAddCash.Text = Trim(dt.Rows(0).Item("TransferValueWord"))
                vAddCashRes = Trim(dt.Rows(0).Item("TransferValueAmount"))
                Me.TXTAddCashRes.Text = Format(Int(Trim(dt.Rows(0).Item("TransferValueAmount"))), "##,##0.00")

                Me.TXTAddCredit.Text = Trim(dt.Rows(0).Item("CreditValueWord"))
                vAddCreditRes = Trim(dt.Rows(0).Item("CreditValueAmount"))
                Me.TXTAddCreditRes.Text = Format(vAddCreditRes, "##,##0.00")

                Me.TXTDiscount1.Text = Trim(dt.Rows(0).Item("SpecialValueWord"))
                vDiscount1Res = Trim(dt.Rows(0).Item("SpecialValueAmount"))
                Me.TXTDiscount1Res.Text = Format(vDiscount1Res, "##,##0.00")

                Me.CMBUnit.Items.Add(Trim(dt.Rows(0).Item("BuyUnitSale")))
                Me.CMBUnit.Text = Me.CMBUnit.Items(0)
                Me.TXTDO.Focus()


                vQuery = "exec dbo.USP_PS_SearchMultiUnitCode '" & vItemcode & "' "
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "MultiUnitCode")
                dt = ds.Tables("MultiUnitCode")

                Me.CMBMultiUnitCode.Items.Clear()
                If dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        Me.CMBMultiUnitCode.Items.Add(dt.Rows(i).Item("unitcode"))
                    Next
                    Me.CMBMultiUnitCode.Text = Me.CMBMultiUnitCode.Items(0)
                End If

            Else
                Call Me.ClearItemDescription()
            End If
            Me.Cursor = Cursors.Arrow



            If Me.ListView101.Items.Count > 0 Then
                vItemcode = Trim(Me.TXTItemCode.Text)
                vListItemCode = Me.ListView101.Items(0).SubItems(13).Text

                If vItemcode <> vListItemCode Then
                    Me.ListView101.Items.Clear()
                End If

            End If
        Else
            Call Me.ClearItemDescription()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Public Sub ClearItemDescription()

        On Error GoTo ErrDescription

        Me.CMBUnit.Items.Clear()
        Me.LBLItemName.Text = ""
        Me.TXTDO.Text = "0"
        Me.TXTPriceSet.Text = "0"
        'Me.TXTBillDisc.Text = "0"
        Me.TXTBillDiscRes.Text = ""
        'Me.TXTFollowDisc1.Text = "0"
        Me.TXTFollowDisc1Res.Text = ""
        'Me.TXTFollowDisc2.Text = "0"
        Me.TXTFollowDisc2Res.Text = ""
        'Me.TXTFollowDisc3.Text = "0"
        Me.TXTFollowDisc3Res.Text = ""
        'Me.TXTRebate.Text = "0"
        Me.TXTRebateRes.Text = ""
        'Me.TXTSpecialDisc.Text = "0"
        Me.TXTSpecialDiscRes.Text = ""
        'Me.TXTMissProfit.Text = "0"
        Me.TXTMissProfitRes.Text = ""
        'Me.TXTSend.Text = "0"
        Me.TXTSendRes.Text = ""
        'Me.TXTCustSend.Text = "0"
        Me.TXTCustSendRes.Text = ""
        'Me.TXTAdvertise.Text = "0"
        Me.TXTAdvertiseRes.Text = ""
        'Me.TXTMarket.Text = "0"
        Me.TXTMarketRes.Text = ""
        'Me.TXTTax.Text = "0"
        Me.TXTTaxRes.Text = ""
        'Me.TXTInstall.Text = "0"
        Me.TXTInstallRes.Text = ""
        'Me.TXTService.Text = "0"
        Me.TXTServiceRes.Text = ""
        'Me.TXTDiscMember.Text = "0"
        Me.TXTDiscMemberRes.Text = ""
        'Me.TXTPoint.Text = "0"
        Me.TXTPointRes.Text = ""
        'Me.TXTTarget.Text = "0"
        Me.TXTTargetRes.Text = ""
        'Me.TXTGift.Text = "0"
        Me.TXTGiftRes.Text = ""
        'Me.TXTCommission.Text = "0"
        Me.TXTCommissionRes.Text = ""
        Me.TXTBegProfit.Text = "0"
        Me.TXTBegProfitAmount.Text = ""
        Me.TXTInterests.Text = "0"
        Me.TXTInterestsAmount.Text = ""
        Me.TXTProfit.Text = "0"
        Me.TXTProfitAmount.Text = ""
        Me.TXTAddCash.Text = "0"
        Me.TXTAddCashRes.Text = "0"
        Me.TXTAddCredit.Text = "0"
        Me.TXTAddCreditRes.Text = "0"
        Me.TXTDiscount1.Text = "0"
        Me.TXTDiscount1Res.Text = "0"
        Me.ListView101.Items.Clear()
        Me.CMBMultiUnitCode.Items.Clear()

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub
    Public Sub ClearItemCode()

        On Error GoTo ErrDescription

        Me.TXTBillDisc.Text = "0"
        Me.TXTBillDiscRes.Text = ""
        Me.TXTFollowDisc1.Text = "0"
        Me.TXTFollowDisc1Res.Text = ""
        Me.TXTFollowDisc2.Text = "0"
        Me.TXTFollowDisc2Res.Text = ""
        Me.TXTFollowDisc3.Text = "0"
        Me.TXTFollowDisc3Res.Text = ""
        Me.TXTRebate.Text = "0"
        Me.TXTRebateRes.Text = ""
        Me.TXTSpecialDisc.Text = "0"
        Me.TXTSpecialDiscRes.Text = ""
        Me.TXTMissProfit.Text = "0"
        Me.TXTMissProfitRes.Text = ""
        Me.TXTSend.Text = "0"
        Me.TXTSendRes.Text = ""
        Me.TXTCustSend.Text = "0"
        Me.TXTCustSendRes.Text = ""
        Me.TXTAdvertise.Text = "0"
        Me.TXTAdvertiseRes.Text = ""
        Me.TXTMarket.Text = "0"
        Me.TXTMarketRes.Text = ""
        Me.TXTTax.Text = "0"
        Me.TXTTaxRes.Text = ""
        Me.TXTInstall.Text = "0"
        Me.TXTInstallRes.Text = ""
        Me.TXTService.Text = "0"
        Me.TXTServiceRes.Text = ""
        Me.TXTDiscMember.Text = "0"
        Me.TXTDiscMemberRes.Text = ""
        Me.TXTPoint.Text = "0"
        Me.TXTPointRes.Text = ""
        Me.TXTTarget.Text = "0"
        Me.TXTTargetRes.Text = ""
        Me.TXTGift.Text = "0"
        Me.TXTGiftRes.Text = ""
        Me.TXTCommission.Text = "0"
        Me.TXTCommissionRes.Text = ""
        Me.TXTBegProfit.Text = "0"
        Me.TXTBegProfitAmount.Text = ""
        Me.TXTInterests.Text = "0"
        Me.TXTInterestsAmount.Text = ""
        Me.TXTProfit.Text = "0"
        Me.TXTProfitAmount.Text = ""
        Me.TXTAddCash.Text = "0"
        Me.TXTAddCashRes.Text = "0"
        Me.TXTAddCredit.Text = "0"
        Me.TXTAddCreditRes.Text = "0"
        Me.TXTDiscount1.Text = "0"
        Me.TXTDiscount1Res.Text = "0"
        Me.TXTDO.Focus()
        Me.ListView101.Items.Clear()
        Me.CMBMultiUnitCode.Items.Clear()

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Public Sub ClearDocument()

        On Error GoTo ErrDescription

        Me.CMBUnit.Items.Clear()
        Me.LBLItemName.Text = ""
        Me.TXTDO.Text = "0"
        Me.TXTPriceSet.Text = "0"
        Me.TXTBillDisc.Text = "0"
        Me.TXTBillDiscRes.Text = ""
        Me.TXTFollowDisc1.Text = "0"
        Me.TXTFollowDisc1Res.Text = ""
        Me.TXTFollowDisc2.Text = "0"
        Me.TXTFollowDisc2Res.Text = ""
        Me.TXTFollowDisc3.Text = "0"
        Me.TXTFollowDisc3Res.Text = ""
        Me.TXTRebate.Text = "0"
        Me.TXTRebateRes.Text = ""
        Me.TXTSpecialDisc.Text = "0"
        Me.TXTSpecialDiscRes.Text = ""
        Me.TXTMissProfit.Text = "0"
        Me.TXTMissProfitRes.Text = ""
        Me.TXTSend.Text = "0"
        Me.TXTSendRes.Text = ""
        Me.TXTCustSend.Text = "0"
        Me.TXTCustSendRes.Text = ""
        Me.TXTAdvertise.Text = "0"
        Me.TXTAdvertiseRes.Text = ""
        Me.TXTMarket.Text = "0"
        Me.TXTMarketRes.Text = ""
        Me.TXTTax.Text = "0"
        Me.TXTTaxRes.Text = ""
        Me.TXTInstall.Text = "0"
        Me.TXTInstallRes.Text = ""
        Me.TXTService.Text = "0"
        Me.TXTServiceRes.Text = ""
        Me.TXTDiscMember.Text = "0"
        Me.TXTDiscMemberRes.Text = ""
        Me.TXTPoint.Text = "0"
        Me.TXTPointRes.Text = ""
        Me.TXTTarget.Text = "0"
        Me.TXTTargetRes.Text = ""
        Me.TXTGift.Text = "0"
        Me.TXTGiftRes.Text = ""
        Me.TXTCommission.Text = "0"
        Me.TXTCommissionRes.Text = ""
        Me.TXTBegProfit.Text = "0"
        Me.TXTBegProfitAmount.Text = ""
        Me.TXTInterests.Text = "0"
        Me.TXTInterestsAmount.Text = ""
        Me.TXTProfit.Text = "0"
        Me.TXTProfitAmount.Text = ""
        Me.TXTAddCash.Text = "0"
        Me.TXTAddCashRes.Text = "0"
        Me.TXTAddCredit.Text = "0"
        Me.TXTAddCreditRes.Text = "0"
        Me.TXTDiscount1.Text = "0"
        Me.TXTDiscount1Res.Text = "0"
        Me.ListView101.Items.Clear()
        Me.CMBMultiUnitCode.Items.Clear()
        Me.TextBoxDocno.Text = ""
        Me.ListView103.Items.Clear()

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub
    Private Sub CMBUnit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBUnit.SelectedIndexChanged
        Dim vItemCode As String
        Dim vUnitCode As String

        On Error GoTo ErrDescription

        If Me.CMBUnit.Text <> "" And Me.TXTItemCode.Text <> "" Then
            vItemCode = Trim(Me.TXTItemCode.Text)
            vUnitCode = Trim(Me.CMBUnit.Text)
            vQuery = "exec dbo.USP_PS_LastPurchase '" & vItemCode & "','" & vUnitCode & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "LastPurchase")
            dt = ds.Tables("LastPurchase")

            If dt.Rows.Count > 0 Then
                Me.TXTDO.Text = Format(dt.Rows(0).Item("lastpurchaseprice"), "##,##0.00")
            Else
                Me.TXTDO.Text = "0.00"
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub ListView102_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView102.DoubleClick
        Dim vIndex As Integer
        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vBillDiscRes As Double
        Dim vFollowDisc1Res As Double
        Dim vFollowDisc2Res As Double
        Dim vFollowDisc3Res As Double
        Dim vRebateRes As Double
        Dim vSpecialDiscRes As Double
        Dim vMissProfitRes As Double
        Dim vSendRes As Double
        Dim vCustSendRes As Double
        Dim vAdvertiseRes As Double
        Dim vMarketRes As Double
        Dim vTaxRes As Double
        Dim vInstallRes As Double
        Dim vServiceRes As Double
        Dim vPointRes As Double
        Dim vTargetRes As Double
        Dim vGiftRes As Double
        Dim vCommissionRes As Double
        Dim vBegProfitAmountRes As Double
        Dim vInterestsAmountRes As Double
        Dim vProfitAmountRes As Double
        Dim vAddCashRes As Double
        Dim vAddCreditRes As Double
        Dim vDiscount1Res As Double
        Dim vCreditCardRes As Double

        On Error GoTo ErrDescription

        If Me.ListView102.Items.Count > 0 Then

            Me.Cursor = Cursors.WaitCursor
            vIndex = Me.ListView102.SelectedItems(0).Index
            Me.TXTItemCode.Text = Trim(Me.ListView102.Items(vIndex).SubItems(0).Text)
            Me.LBLItemName.Text = Trim(Me.ListView102.Items(vIndex).SubItems(1).Text)
            Me.CMBUnit.Items.Clear()
            Me.CMBUnit.Items.Add(Me.ListView102.Items(vIndex).SubItems(2).Text)
            Me.CMBUnit.Text = Me.CMBUnit.Items.Item(0)

            vDO = Me.ListView102.Items(vIndex).SubItems(3).Text
            Me.TXTDO.Text = Format(vDO, "##,##0.00")

            vPriceSet = Me.ListView102.Items(vIndex).SubItems(4).Text
            Me.TXTPriceSet.Text = Format(vPriceSet, "##,##0.00")

            Me.TXTBillDisc.Text = Me.ListView102.Items(vIndex).SubItems(5).Text
            vBillDiscRes = Me.ListView102.Items(vIndex).SubItems(6).Text
            Me.TXTBillDiscRes.Text = Format(vBillDiscRes, "##,##0.00")

            Me.TXTFollowDisc1.Text = Me.ListView102.Items(vIndex).SubItems(7).Text
            vFollowDisc1Res = Me.ListView102.Items(vIndex).SubItems(8).Text
            Me.TXTFollowDisc1Res.Text = Format(vFollowDisc1Res, "##,##0.00")

            Me.TXTFollowDisc2.Text = Me.ListView102.Items(vIndex).SubItems(9).Text
            vFollowDisc2Res = Me.ListView102.Items(vIndex).SubItems(10).Text
            Me.TXTFollowDisc2Res.Text = Format(vFollowDisc2Res, "##,##0.00")

            Me.TXTFollowDisc3.Text = Me.ListView102.Items(vIndex).SubItems(11).Text
            vFollowDisc3Res = Me.ListView102.Items(vIndex).SubItems(12).Text
            Me.TXTFollowDisc3Res.Text = Format(vFollowDisc3Res, "##,##0.00")

            Me.TXTRebate.Text = Me.ListView102.Items(vIndex).SubItems(13).Text
            vRebateRes = Me.ListView102.Items(vIndex).SubItems(14).Text
            Me.TXTRebateRes.Text = Format(vRebateRes, "##,##0.00")

            Me.TXTSpecialDisc.Text = Me.ListView102.Items(vIndex).SubItems(15).Text
            vSpecialDiscRes = Me.ListView102.Items(vIndex).SubItems(16).Text
            Me.TXTSpecialDiscRes.Text = Format(vSpecialDiscRes, "##,##0.00")

            Me.TXTMissProfit.Text = Me.ListView102.Items(vIndex).SubItems(17).Text
            vMissProfitRes = Me.ListView102.Items(vIndex).SubItems(18).Text
            Me.TXTMissProfitRes.Text = Format(vMissProfitRes, "##,##0.00")

            Me.TXTSend.Text = Me.ListView102.Items(vIndex).SubItems(19).Text
            vSendRes = Me.ListView102.Items(vIndex).SubItems(20).Text
            Me.TXTSendRes.Text = Format(vSendRes, "##,##0.00")

            Me.TXTCustSend.Text = Me.ListView102.Items(vIndex).SubItems(21).Text
            vCustSendRes = Me.ListView102.Items(vIndex).SubItems(22).Text
            Me.TXTCustSendRes.Text = Format(vCustSendRes, "##,##0.00")

            Me.TXTAdvertise.Text = Me.ListView102.Items(vIndex).SubItems(23).Text
            vAdvertiseRes = Me.ListView102.Items(vIndex).SubItems(24).Text
            Me.TXTAdvertiseRes.Text = Format(vAdvertiseRes, "##,##0.00")

            Me.TXTMarket.Text = Me.ListView102.Items(vIndex).SubItems(25).Text
            vMarketRes = Me.ListView102.Items(vIndex).SubItems(26).Text
            Me.TXTMarketRes.Text = Format(vMarketRes, "##,##0.00")

            Me.TXTTax.Text = Me.ListView102.Items(vIndex).SubItems(27).Text
            vTaxRes = Me.ListView102.Items(vIndex).SubItems(28).Text
            Me.TXTTaxRes.Text = Format(vTaxRes, "##,##0.00")

            Me.TXTInstall.Text = Me.ListView102.Items(vIndex).SubItems(29).Text
            vInstallRes = Me.ListView102.Items(vIndex).SubItems(30).Text
            Me.TXTInstallRes.Text = Format(vInstallRes, "##,##0.00")

            Me.TXTService.Text = Me.ListView102.Items(vIndex).SubItems(31).Text
            vServiceRes = Me.ListView102.Items(vIndex).SubItems(32).Text
            Me.TXTServiceRes.Text = Format(vServiceRes, "##,##0.00")

            Me.TXTPoint.Text = Me.ListView102.Items(vIndex).SubItems(33).Text
            vPointRes = Me.ListView102.Items(vIndex).SubItems(34).Text
            Me.TXTPointRes.Text = Format(vPointRes, "##,##0.00")

            Me.TXTDiscMember.Text = Me.ListView102.Items(vIndex).SubItems(35).Text
            vCreditCardRes = Me.ListView102.Items(vIndex).SubItems(36).Text
            Me.TXTDiscMemberRes.Text = Format(vCreditCardRes, "##,##0.00")

            Me.TXTTarget.Text = Me.ListView102.Items(vIndex).SubItems(37).Text
            vTargetRes = Me.ListView102.Items(vIndex).SubItems(38).Text
            Me.TXTTargetRes.Text = Format(vTargetRes, "##,##0.00")

            Me.TXTGift.Text = Me.ListView102.Items(vIndex).SubItems(39).Text
            vGiftRes = Me.ListView102.Items(vIndex).SubItems(40).Text
            Me.TXTGiftRes.Text = Format(vGiftRes, "##,##0.00")

            Me.TXTCommission.Text = Me.ListView102.Items(vIndex).SubItems(41).Text
            vCommissionRes = Me.ListView102.Items(vIndex).SubItems(42).Text
            Me.TXTCommissionRes.Text = Format(vCommissionRes, "##,##0.00")

            Me.TXTBegProfit.Text = Me.ListView102.Items(vIndex).SubItems(43).Text
            vBegProfitAmountRes = Me.ListView102.Items(vIndex).SubItems(44).Text
            Me.TXTBegProfitAmount.Text = Format(vBegProfitAmountRes, "##,##0.00")

            Me.TXTInterests.Text = Me.ListView102.Items(vIndex).SubItems(45).Text
            vInterestsAmountRes = Me.ListView102.Items(vIndex).SubItems(46).Text
            Me.TXTInterestsAmount.Text = Format(vInterestsAmountRes, "##,##0.00")

            Me.TXTProfit.Text = Format(Me.ListView102.Items(vIndex).SubItems(47).Text)
            vProfitAmountRes = Me.ListView102.Items(vIndex).SubItems(48).Text
            Me.TXTProfitAmount.Text = Format(vProfitAmountRes, "##,##0.00")

            Me.TXTAddCash.Text = Me.ListView102.Items(vIndex).SubItems(49).Text
            vAddCashRes = Me.ListView102.Items(vIndex).SubItems(50).Text
            Me.TXTAddCashRes.Text = Format(vAddCashRes, "##,##0.00")

            Me.TXTAddCredit.Text = Me.ListView102.Items(vIndex).SubItems(51).Text
            vAddCreditRes = Me.ListView102.Items(vIndex).SubItems(52).Text
            Me.TXTAddCreditRes.Text = Format(vAddCreditRes, "##,##0.00")

            Me.TXTDiscount1.Text = Me.ListView102.Items(vIndex).SubItems(53).Text
            vDiscount1Res = Me.ListView102.Items(vIndex).SubItems(54).Text
            Me.TXTDiscount1Res.Text = Format(vDiscount1Res, "##,##0.00")

            Me.Cursor = Cursors.Arrow
            Me.TXTDO.Focus()
            Me.ListView101.Items.Clear()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTDO_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTDO.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTPriceSet.Focus()
        End If
        If Asc(e.KeyCode) = 52 Then
            Me.TXTPriceSet.Focus()
        End If
    End Sub

    Private Sub TXTDO_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TXTDO.KeyPress, TXTPriceSet.KeyPress, TXTPrice1.KeyPress, TXTPrice2.KeyPress
        On Error Resume Next

        Select Case Asc(e.KeyChar)
            Case 48 To 58, 8, 44, 46
            Case Else
                e.Handled = True
        End Select
    End Sub

    Private Sub TXTDO_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTDO.LostFocus
        Dim vCheckDO As Double

        On Error GoTo ErrDescription

        If Me.LBLItemName.Text <> "" Then
            If Me.TXTDO.Text <> "" And Me.TXTDO.Text <> "." And InStr(Me.TXTDO.Text, ".,") = 0 And InStr(Me.TXTDO.Text, ",.") = 0 Then
                vCheckDO = Me.TXTDO.Text
                Me.TXTDO.Text = Format(vCheckDO, "##,##0.00")
                Call Me.TXTBillDiscCalc()
            Else
                Me.TXTDO.Focus()
            End If
            If Me.TXTDO.Text = "." And Len(Me.TXTDO.Text) = 1 Then
                Me.TXTDO.Focus()
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub BTNDOHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNDOHistory.Click
        Dim vItemCode As String
        Dim vUnitCode As String
        Dim i As Integer
        Dim vListItem As ListViewItem

        On Error GoTo ErrDescription

        If Me.TXTItemCode.Text <> "" And Me.CMBUnit.Text <> "" Then
            Me.GBDO.Visible = True
            Me.GBDO.BringToFront()
            vItemCode = Trim(Me.TXTItemCode.Text)
            vUnitCode = Trim(Me.CMBUnit.Text)
            vQuery = "exec dbo.USP_PS_PurchaseHistory '" & vItemCode & "','" & vUnitCode & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "PurchaseHistory")
            dt = ds.Tables("PurchaseHistory")

            Me.ListViewDOHistory.Items.Clear()
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    vListItem = Me.ListViewDOHistory.Items.Add(dt.Rows(i).Item("purchasedate"))
                    vListItem.SubItems.Add(1).Text = Format(Int(dt.Rows(i).Item("qty")), "##,##0.00")
                    vListItem.SubItems.Add(2).Text = Format(Int(dt.Rows(i).Item("price")), "##,##0.00")
                    vListItem.SubItems.Add(3).Text = Trim(dt.Rows(i).Item("discountword"))
                    vListItem.SubItems.Add(4).Text = Format(Int(dt.Rows(i).Item("discountamount")), "##,##0.00")
                    vListItem.SubItems.Add(5).Text = Format(Int(dt.Rows(i).Item("amount")), "##,##0.00")
                Next
            End If
        Else
            MsgBox("ไม่สามารถแสดงข้อมูลประวัติราคาตั้งจาก Vendor ได้ เนื่องจากไม่มีรหัสสินค้าและหน่วยนับ", MsgBoxStyle.Critical, "Send Error ")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub BTNDOClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNDOClose.Click
        On Error GoTo ErrDescription

        Me.GBDO.Visible = False

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTBillDisc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTBillDisc.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTFollowDisc1.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTPriceSet.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTFollowDisc1.Focus()
        End If
    End Sub

    Private Sub TXTBillDisc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TXTBillDisc.KeyPress, TXTFollowDisc1.KeyPress, TXTFollowDisc2.KeyPress, TXTFollowDisc3.KeyPress, TXTRebate.KeyPress, TXTSpecialDisc.KeyPress, TXTMissProfit.KeyPress, TXTSend.KeyPress, TXTCustSend.KeyPress, TXTAdvertise.KeyPress, TXTMarket.KeyPress, TXTTax.KeyPress, TXTInstall.KeyPress, TXTService.KeyPress, TXTPoint.KeyPress, TXTTarget.KeyPress, TXTGift.KeyPress, TXTCommission.KeyPress, TXTAddCash.KeyPress, TXTAddCredit.KeyPress, TXTDiscount1.KeyPress
        On Error GoTo ErrDescription

        Select Case Asc(e.KeyChar)
            Case 48 To 58, 8, 44, 46, 37
            Case Else
                e.Handled = True
        End Select

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TXTBillDisc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTBillDisc.LostFocus
        Dim vCheckBillDisc As Double
        Dim vCheckBillDiscRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" Then

            vCheckCountDot = CheckDot(TXTBillDisc.Text)
            If vCheckCountDot > 1 Then
                Me.TXTBillDisc.Focus()
                Exit Sub
            End If

            If InStr(TXTBillDisc.Text, "0") = 1 And InStr(TXTBillDisc.Text, ".") = 2 And Len(TXTBillDisc.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTBillDisc.Text, ".") = 1 And Len(TXTBillDisc.Text) = 1 Then
                Me.TXTBillDisc.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTBillDisc.Text, "%") = 1 Then
                Me.TXTBillDisc.Focus()
            End If
            If InStr(TXTBillDisc.Text, ".%") > 0 Or InStr(TXTBillDisc.Text, "%.") > 0 Then
                Me.TXTBillDisc.Focus()
                Exit Sub
            End If

            If Me.TXTBillDisc.Text = "" Then
                Me.TXTBillDisc.Text = Format(0, "##,##0.00")
                Exit Sub
            End If
            If InStr(TXTBillDisc.Text, "%") = 0 Then
                vCheckBillDisc = TXTBillDisc.Text
                If vCheckBillDisc = 0 Then
                    vCheckBillDiscRes = Me.TXTDO.Text
                    Me.TXTBillDiscRes.Text = Format(vCheckBillDiscRes, "##,##0.00")
                End If
                Me.TXTBillDisc.Text = Format(vCheckBillDisc, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub
    Function CheckDot(ByVal vNumber As String) As Integer
        Dim i As Integer
        Dim vCount As Integer

        For i = 1 To Len(vNumber)
            If Microsoft.VisualBasic.Mid(vNumber, i, 1) = "." Then
                vCount = vCount + 1
            End If
        Next
        CheckDot = vCount
    End Function

    Private Sub TXTBillDisc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTBillDisc.TextChanged
        Call TXTBillDiscCalc()
    End Sub

    Private Sub TXTBillDiscCalc()
        Dim vDO As Double
        Dim vBillDiscWord As String
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And InStr(Me.TXTDO.Text, ".,") = 0 And InStr(Me.TXTDO.Text, ",.") = 0 And Me.TXTDO.Text <> "." Then
            If Me.TXTBillDisc.Text <> "" Then
                vBillDiscWord = Me.TXTBillDisc.Text
                vDO = Me.TXTDO.Text
                vCheckCountDot = CheckDot(TXTBillDisc.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTBillDisc.Focus()
                    Exit Sub
                End If

                If InStr(vBillDiscWord, "%") > 0 Then
                    If Len(vBillDiscWord) > 1 Then
                        If InStr(vBillDiscWord, "0") = 1 Then
                            Me.TXTBillDisc.Text = Microsoft.VisualBasic.Right(vBillDiscWord, Len(vBillDiscWord) - InStr(vBillDiscWord, "0"))
                            vBillDiscWord = Me.TXTBillDisc.Text
                        End If
                    End If
                    If InStr(vBillDiscWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTBillDisc.Focus()
                        Exit Sub
                    End If
                    If InStr(vBillDiscWord, ".%") > 0 Or InStr(vBillDiscWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTBillDisc.Focus()
                        Exit Sub
                    End If

                    vBillDiscWord = Microsoft.VisualBasic.Left(vBillDiscWord, InStr(vBillDiscWord, "%"))
                    Me.TXTBillDisc.Text = vBillDiscWord
                    vBillDisc = Microsoft.VisualBasic.Left(vBillDiscWord, InStr(vBillDiscWord, "%") - 1)
                    vBillDisc1 = ((vDO * vBillDisc) / 100)
                    vBillDiscAmount = vDO - vBillDisc1
                Else
                    If Len(vBillDiscWord) > 0 Then
                        If InStr(vBillDiscWord, ".") = 1 And Len(vBillDiscWord) = 1 Then
                            Me.TXTBillDisc.Focus()
                            Exit Sub
                        End If

                        If InStr(vBillDiscWord, "0") = 1 And InStr(vBillDiscWord, ".") = 2 And Len(vBillDiscWord) = 2 Then
                            Me.TXTBillDisc.Focus()
                            Exit Sub
                        End If

                        If (InStr(vBillDiscWord, ".") = 1) And Len(vBillDiscWord) = 1 Then
                            Me.TXTBillDisc.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vBillDiscWord, "0") = 1 And Len(vBillDiscWord) > 1 Then
                            Me.TXTBillDisc.Text = Microsoft.VisualBasic.Right(vBillDiscWord, Len(vBillDiscWord) - InStr(vBillDiscWord, "0"))
                        Else
                            Me.TXTBillDisc.Text = vBillDiscWord
                        End If

                    End If

                    vBillDisc = (vBillDiscWord)
                    vBillDisc1 = vBillDisc
                    vBillDiscAmount = vDO - vBillDisc1
                End If
                If vBillDiscAmount < 0 Then
                    Me.TXTBillDiscRes.ForeColor = Color.Red
                Else
                    Me.TXTBillDiscRes.ForeColor = Color.Black
                End If
                Me.TXTBillDiscRes.Text = Format(vBillDiscAmount, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message BillDisc")
        End If
    End Sub

    Private Sub TXTPriceSet_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTPriceSet.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTBillDisc.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTDO.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTBillDisc.Focus()
        End If
    End Sub

    Private Sub TXTFollowDisc1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTFollowDisc1.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTFollowDisc2.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTBillDisc.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTFollowDisc2.Focus()
        End If
    End Sub

    Private Sub TXTFollowDisc2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTFollowDisc2.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTFollowDisc3.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTFollowDisc1.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTFollowDisc3.Focus()
        End If
    End Sub

    Private Sub TXTFollowDisc3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTFollowDisc3.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTRebate.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTFollowDisc2.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTRebate.Focus()
        End If
    End Sub

    Private Sub TXTRebate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTRebate.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTSpecialDisc.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTFollowDisc3.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTSpecialDisc.Focus()
        End If
    End Sub

    Private Sub TXTSpecialDisc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTSpecialDisc.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTMissProfit.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTRebate.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTMissProfit.Focus()
        End If
    End Sub

    Private Sub TXTMissProfit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTMissProfit.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTSend.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTSpecialDisc.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTSend.Focus()
        End If
    End Sub

    Private Sub TXTSend_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTSend.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTCustSend.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTMissProfit.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTCustSend.Focus()
        End If
    End Sub

    Private Sub TXTCustSend_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTCustSend.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTAdvertise.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTSend.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTAdvertise.Focus()
        End If
    End Sub

    Private Sub TXTAdvertise_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTAdvertise.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTMarket.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTCustSend.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTMarket.Focus()
        End If
    End Sub

    Private Sub TXTMarket_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTMarket.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTTax.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTAdvertise.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTTax.Focus()
        End If
    End Sub

    Private Sub TXTTax_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTTax.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTInstall.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTMarket.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTInstall.Focus()
        End If
    End Sub

    Private Sub TXTInstall_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTInstall.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTService.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTTax.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTService.Focus()
        End If
    End Sub

    Private Sub TXTService_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTService.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTPoint.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTInstall.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTPoint.Focus()
        End If
    End Sub

    Private Sub TXTPoint_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTPoint.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTDiscMember.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTService.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTDiscMember.Focus()
        End If
    End Sub

    Private Sub TXTTarget_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTTarget.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTGift.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTDiscMember.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTGift.Focus()
        End If
    End Sub

    Private Sub TXTGift_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTGift.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTCommission.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTTarget.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTCommission.Focus()
        End If
    End Sub

    Private Sub TXTCommission_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTCommission.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTAddCash.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTGift.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTAddCash.Focus()
        End If
    End Sub

    Private Sub TXTAddCash_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTAddCash.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTAddCredit.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTCommission.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTAddCredit.Focus()
        End If
    End Sub

    Private Sub TXTAddCredit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTAddCredit.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTDiscount1.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTAddCash.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTDiscount1.Focus()
        End If
    End Sub

    Private Sub TXTDiscount1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTDiscount1.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.BTNReCommend.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTAddCredit.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.BTNReCommend.Focus()
        End If
    End Sub

    Private Sub TXTPriceSet_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTPriceSet.LostFocus
        Dim vCheckPriceSet As Double

        On Error GoTo ErrDescription

        If Me.LBLItemName.Text <> "" Then
            If Me.TXTPriceSet.Text <> "" And InStr(Me.TXTPriceSet.Text, ".,") = 0 And InStr(Me.TXTPriceSet.Text, ",.") = 0 Then
                vCheckPriceSet = Me.TXTPriceSet.Text
                Me.TXTPriceSet.Text = Format(vCheckPriceSet, "##,##0.00")
            Else
                Me.TXTPriceSet.Focus()
            End If

            If Me.TXTPriceSet.Text = "." And Len(Me.TXTPriceSet.Text) = 1 Then
                Me.TXTPriceSet.Focus()
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTFollowDisc1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTFollowDisc1.LostFocus
        Dim vCheckFollowDisc1 As Double
        Dim vCheckFollowDisc1Res As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTBillDiscRes.Text <> "" Then

            vCheckCountDot = CheckDot(TXTFollowDisc1.Text)
            If vCheckCountDot > 1 Then
                Me.TXTFollowDisc1.Focus()
                Exit Sub
            End If

            If InStr(TXTFollowDisc1.Text, "0") = 1 And InStr(TXTFollowDisc1.Text, ".") = 2 And Len(TXTFollowDisc1.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTFollowDisc1.Text, ".") = 1 And Len(TXTFollowDisc1.Text) = 1 Then
                Me.TXTFollowDisc1.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTFollowDisc1.Text, "%") = 1 Then
                Me.TXTFollowDisc1.Focus()
            End If

            If InStr(TXTFollowDisc1.Text, ".%") > 0 Or InStr(TXTFollowDisc1.Text, "%.") > 0 Then
                Me.TXTFollowDisc1.Focus()
                Exit Sub
            End If

            If Me.TXTFollowDisc1.Text = "" Then
                Me.TXTFollowDisc1.Text = Format(0, "##,##0.00")
                Exit Sub
            End If
            If InStr(TXTFollowDisc1.Text, "%") = 0 Then
                vCheckFollowDisc1 = TXTFollowDisc1.Text
                If vCheckFollowDisc1 = 0 Then
                    vCheckFollowDisc1Res = Me.TXTBillDiscRes.Text
                    Me.TXTFollowDisc1Res.Text = Format(vCheckFollowDisc1Res, "##,##0.00")
                End If
                Me.TXTFollowDisc1.Text = Format(vCheckFollowDisc1, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTFollowDisc1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTFollowDisc1.TextChanged
        Call TXTFollowDisc1Calc()
    End Sub
    Private Sub TXTFollowDisc1Calc()
        Dim vAccCost As Double
        Dim vDiscFollow1Word As String
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTBillDiscRes.Text <> "" Then
            If Me.TXTFollowDisc1.Text <> "" Then
                vDiscFollow1Word = Me.TXTFollowDisc1.Text
                vAccCost = Me.TXTBillDiscRes.Text

                vCheckCountDot = CheckDot(TXTFollowDisc1.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTFollowDisc1.Focus()
                    Exit Sub
                End If

                If InStr(vDiscFollow1Word, "%") > 0 Then
                    If Len(vDiscFollow1Word) > 1 Then
                        If InStr(vDiscFollow1Word, "0") = 1 Then
                            Me.TXTFollowDisc1.Text = Microsoft.VisualBasic.Right(vDiscFollow1Word, Len(vDiscFollow1Word) - InStr(vDiscFollow1Word, "0"))
                            vDiscFollow1Word = Me.TXTFollowDisc1.Text
                        End If
                    End If
                    If InStr(vDiscFollow1Word, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTFollowDisc1.Focus()
                        Exit Sub
                    End If
                    If InStr(vDiscFollow1Word, ".%") > 0 Or InStr(vDiscFollow1Word, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTFollowDisc1.Focus()
                        Exit Sub
                    End If
                    vDiscFollow1Word = Microsoft.VisualBasic.Left(vDiscFollow1Word, InStr(vDiscFollow1Word, "%"))
                    Me.TXTFollowDisc1.Text = vDiscFollow1Word
                    vDiscFollow1 = Microsoft.VisualBasic.Left(vDiscFollow1Word, InStr(vDiscFollow1Word, "%") - 1)
                    vDiscFollow11 = ((vAccCost * vDiscFollow1) / 100)
                    vFollowDisc1Amount = vAccCost - vDiscFollow11
                Else
                    If Len(vDiscFollow1Word) > 0 Then
                        If InStr(vDiscFollow1Word, ".") = 1 And Len(vDiscFollow1Word) = 1 Then
                            Me.TXTFollowDisc1.Focus()
                            Exit Sub
                        End If

                        If InStr(vDiscFollow1Word, "0") = 1 And InStr(vDiscFollow1Word, ".") = 2 And Len(vDiscFollow1Word) = 2 Then
                            Me.TXTFollowDisc1.Focus()
                            Exit Sub
                        End If

                        If (InStr(vDiscFollow1Word, ".") = 1) And Len(vDiscFollow1Word) = 1 Then
                            Me.TXTFollowDisc1.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vDiscFollow1Word, "0") = 1 And Len(vDiscFollow1Word) > 1 Then
                            Me.TXTFollowDisc1.Text = Microsoft.VisualBasic.Right(vDiscFollow1Word, Len(vDiscFollow1Word) - InStr(vDiscFollow1Word, "0"))
                        Else
                            Me.TXTFollowDisc1.Text = vDiscFollow1Word
                        End If

                    End If
                    vDiscFollow1 = vDiscFollow1Word
                    vDiscFollow11 = vDiscFollow1
                    vFollowDisc1Amount = vAccCost - vDiscFollow11
                End If

                If vFollowDisc1Amount < 0 Then
                    Me.TXTFollowDisc1Res.ForeColor = Color.Red
                Else
                    Me.TXTFollowDisc1Res.ForeColor = Color.Black
                End If

                Me.TXTFollowDisc1Res.Text = Format(vFollowDisc1Amount, "##,##0.00")

            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message FollowDisc1")
        End If
    End Sub

    Private Sub TXTFollowDisc2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTFollowDisc2.LostFocus
        Dim vCheckFollowDisc2 As Double
        Dim vCheckFollowDisc2Res As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTFollowDisc1Res.Text <> "" Then
            vCheckCountDot = CheckDot(TXTFollowDisc2.Text)
            If vCheckCountDot > 1 Then
                Me.TXTFollowDisc2.Focus()
                Exit Sub
            End If

            If InStr(TXTFollowDisc2.Text, "0") = 1 And InStr(TXTFollowDisc2.Text, ".") = 2 And Len(TXTFollowDisc2.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTFollowDisc2.Text, ".") = 1 And Len(TXTFollowDisc2.Text) = 1 Then
                Me.TXTFollowDisc2.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTFollowDisc2.Text, "%") = 1 Then
                Me.TXTFollowDisc2.Focus()
            End If

            If InStr(TXTFollowDisc2.Text, ".%") > 0 Or InStr(TXTFollowDisc2.Text, "%.") > 0 Then
                Me.TXTFollowDisc2.Focus()
                Exit Sub
            End If

            If Me.TXTFollowDisc2.Text = "" Then
                Me.TXTFollowDisc2.Text = Format(0, "##,##0.00")
                Exit Sub
            End If
            If InStr(TXTFollowDisc2.Text, "%") = 0 Then
                vCheckFollowDisc2 = TXTFollowDisc2.Text
                If vCheckFollowDisc2 = 0 Then
                    vCheckFollowDisc2Res = Me.TXTFollowDisc1Res.Text
                    Me.TXTFollowDisc2Res.Text = Format(vCheckFollowDisc2Res, "##,##0.00")
                End If
                Me.TXTFollowDisc2.Text = Format(vCheckFollowDisc2, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTFollowDisc2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTFollowDisc2.TextChanged
        Call TXTFollowDisc2Calc()
    End Sub

    Private Sub TXTFollowDisc2Calc()
        Dim vFollow1 As Double
        Dim vDiscFollow2Word As String
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTFollowDisc1Res.Text <> "" Then
            If Me.TXTFollowDisc2.Text <> "" Then
                vDiscFollow2Word = Me.TXTFollowDisc2.Text
                vFollow1 = Me.TXTFollowDisc1Res.Text

                vCheckCountDot = CheckDot(TXTFollowDisc2.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTFollowDisc2.Focus()
                    Exit Sub
                End If

                If InStr(vDiscFollow2Word, "%") > 0 Then
                    If Len(vDiscFollow2Word) > 1 Then
                        If InStr(vDiscFollow2Word, "0") = 1 Then
                            Me.TXTFollowDisc2.Text = Microsoft.VisualBasic.Right(vDiscFollow2Word, Len(vDiscFollow2Word) - InStr(vDiscFollow2Word, "0"))
                            vDiscFollow2Word = Me.TXTFollowDisc2.Text
                        End If
                    End If
                    If InStr(vDiscFollow2Word, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTFollowDisc2.Focus()
                        Exit Sub
                    End If

                    If InStr(vDiscFollow2Word, ".%") > 0 Or InStr(vDiscFollow2Word, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTFollowDisc2.Focus()
                        Exit Sub
                    End If

                    vDiscFollow2Word = Microsoft.VisualBasic.Left(vDiscFollow2Word, InStr(vDiscFollow2Word, "%"))
                    Me.TXTFollowDisc2.Text = vDiscFollow2Word
                    vDiscFollow2 = Microsoft.VisualBasic.Left(vDiscFollow2Word, InStr(vDiscFollow2Word, "%") - 1)
                    vDiscFollow21 = ((vFollow1 * vDiscFollow2) / 100)
                    vFollowDisc2Amount = vFollow1 - vDiscFollow21
                Else
                    If Len(vDiscFollow2Word) > 0 Then
                        If InStr(vDiscFollow2Word, ".") = 1 And Len(vDiscFollow2Word) = 1 Then
                            Me.TXTFollowDisc2.Focus()
                            Exit Sub
                        End If

                        If InStr(vDiscFollow2Word, "0") = 1 And InStr(vDiscFollow2Word, ".") = 2 And Len(vDiscFollow2Word) = 2 Then
                            Me.TXTFollowDisc2.Focus()
                            Exit Sub
                        End If

                        If (InStr(vDiscFollow2Word, ".") = 1) And Len(vDiscFollow2Word) = 1 Then
                            Me.TXTFollowDisc2.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vDiscFollow2Word, "0") = 1 And Len(vDiscFollow2Word) > 1 Then
                            Me.TXTFollowDisc2.Text = Microsoft.VisualBasic.Right(vDiscFollow2Word, Len(vDiscFollow2Word) - InStr(vDiscFollow2Word, "0"))
                        Else
                            Me.TXTFollowDisc2.Text = vDiscFollow2Word
                        End If
                    End If
                    vDiscFollow2 = vDiscFollow2Word
                    vDiscFollow21 = vDiscFollow2
                    vFollowDisc2Amount = vFollow1 - vDiscFollow21
                End If

                If vFollowDisc2Amount < 0 Then
                    Me.TXTFollowDisc2Res.ForeColor = Color.Red
                Else
                    Me.TXTFollowDisc2Res.ForeColor = Color.Black
                End If

                Me.TXTFollowDisc2Res.Text = Format(vFollowDisc2Amount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message FollowDisc2")
        End If
    End Sub

    Private Sub TXTFollowDisc3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTFollowDisc3.LostFocus
        Dim vCheckFollowDisc3 As Double
        Dim vCheckFollowDisc3Res As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTFollowDisc2Res.Text <> "" Then
            vCheckCountDot = CheckDot(TXTFollowDisc3.Text)
            If vCheckCountDot > 1 Then
                Me.TXTFollowDisc3.Focus()
                Exit Sub
            End If

            If InStr(TXTFollowDisc3.Text, "0") = 1 And InStr(TXTFollowDisc3.Text, ".") = 2 And Len(TXTFollowDisc3.Text) = 2 Then
                Exit Sub
            End If

            If InStr(TXTFollowDisc3.Text, ".") = 1 And Len(TXTFollowDisc3.Text) = 1 Then
                Me.TXTFollowDisc3.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTFollowDisc3.Text, "%") = 1 Then
                Me.TXTFollowDisc3.Focus()
            End If

            If InStr(TXTFollowDisc3.Text, ".%") > 0 Or InStr(TXTFollowDisc3.Text, "%.") > 0 Then
                Me.TXTFollowDisc3.Focus()
                Exit Sub
            End If

            If Me.TXTFollowDisc3.Text = "" Then
                Me.TXTFollowDisc3.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTFollowDisc3.Text, "%") = 0 Then
                vCheckFollowDisc3 = TXTFollowDisc3.Text
                If vCheckFollowDisc3 = 0 Then
                    vCheckFollowDisc3Res = Me.TXTFollowDisc2Res.Text
                    Me.TXTFollowDisc3Res.Text = Format(vCheckFollowDisc3Res, "##,##0.00")
                End If
                Me.TXTFollowDisc3.Text = Format(vCheckFollowDisc3, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTFollowDisc3Calc()
        Dim vFollow2 As Double
        Dim vDiscFollow3Word As String
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTFollowDisc2Res.Text <> "" Then
            If Me.TXTFollowDisc3.Text <> "" Then
                vDiscFollow3Word = Me.TXTFollowDisc3.Text
                vFollow2 = Me.TXTFollowDisc2Res.Text

                vCheckCountDot = CheckDot(TXTFollowDisc3.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTFollowDisc3.Focus()
                    Exit Sub
                End If

                If InStr(vDiscFollow3Word, "%") > 0 Then
                    If Len(vDiscFollow3Word) > 1 Then
                        If InStr(vDiscFollow3Word, "0") = 1 Then
                            Me.TXTFollowDisc3.Text = Microsoft.VisualBasic.Right(vDiscFollow3Word, Len(vDiscFollow3Word) - InStr(vDiscFollow3Word, "0"))
                            vDiscFollow3Word = Me.TXTFollowDisc3.Text
                        End If
                    End If
                    If InStr(vDiscFollow3Word, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTFollowDisc3.Focus()
                        Exit Sub
                    End If

                    If InStr(vDiscFollow3Word, ".%") > 0 Or InStr(vDiscFollow3Word, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTFollowDisc3.Focus()
                        Exit Sub
                    End If

                    vDiscFollow3Word = Microsoft.VisualBasic.Left(vDiscFollow3Word, InStr(vDiscFollow3Word, "%"))
                    Me.TXTFollowDisc3.Text = vDiscFollow3Word
                    vDiscFollow3 = Microsoft.VisualBasic.Left(vDiscFollow3Word, InStr(vDiscFollow3Word, "%") - 1)
                    vDiscFollow31 = ((vFollow2 * vDiscFollow3) / 100)
                    vFollowDisc3Amount = vFollow2 - vDiscFollow31
                Else
                    If Len(vDiscFollow3Word) > 0 Then
                        If InStr(vDiscFollow3Word, ".") = 1 And Len(vDiscFollow3Word) = 1 Then
                            Me.TXTFollowDisc3.Focus()
                            Exit Sub
                        End If

                        If InStr(vDiscFollow3Word, "0") = 1 And InStr(vDiscFollow3Word, ".") = 2 And Len(vDiscFollow3Word) = 2 Then
                            Me.TXTFollowDisc3.Focus()
                            Exit Sub
                        End If

                        If (InStr(vDiscFollow3Word, ".") = 1) And Len(vDiscFollow3Word) = 1 Then
                            Me.TXTFollowDisc3.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vDiscFollow3Word, "0") = 1 And Len(vDiscFollow3Word) > 1 Then
                            Me.TXTFollowDisc3.Text = Microsoft.VisualBasic.Right(vDiscFollow3Word, Len(vDiscFollow3Word) - InStr(vDiscFollow3Word, "0"))
                        Else
                            Me.TXTFollowDisc3.Text = vDiscFollow3Word
                        End If
                    End If

                    vDiscFollow3 = vDiscFollow3Word
                    vDiscFollow31 = vDiscFollow3
                    vFollowDisc3Amount = vFollow2 - vDiscFollow31
                End If

                If vFollowDisc3Amount < 0 Then
                    Me.TXTFollowDisc3Res.ForeColor = Color.Red
                Else
                    Me.TXTFollowDisc3Res.ForeColor = Color.Black
                End If

                Me.TXTFollowDisc3Res.Text = Format(vFollowDisc3Amount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message FollowDisc3")
        End If
    End Sub

    Private Sub TXTRebate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTRebate.LostFocus
        Dim vCheckRebate As Double
        Dim vCheckRebateRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If TXTFollowDisc3Res.Text <> "" Then
            vCheckCountDot = CheckDot(TXTRebate.Text)
            If vCheckCountDot > 1 Then
                Me.TXTRebate.Focus()
                Exit Sub
            End If

            If InStr(TXTRebate.Text, "0") = 1 And InStr(TXTRebate.Text, ".") = 2 And Len(TXTRebate.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTRebate.Text, ".") = 1 And Len(TXTRebate.Text) = 1 Then
                Me.TXTRebate.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTRebate.Text, "%") = 1 Then
                Me.TXTRebate.Focus()
            End If

            If InStr(TXTRebate.Text, ".%") > 0 Or InStr(TXTRebate.Text, "%.") > 0 Then
                Me.TXTRebate.Focus()
                Exit Sub
            End If

            If Me.TXTRebate.Text = "" Then
                Me.TXTRebate.Text = Format(0, "##,##0.00")
                Exit Sub
            End If
            If InStr(TXTRebate.Text, "%") = 0 Then
                vCheckRebate = TXTRebate.Text
                If vCheckRebate = 0 Then
                    vCheckRebateRes = Me.TXTFollowDisc3Res.Text
                    Me.TXTRebateRes.Text = Format(vCheckRebateRes, "##,##0.00")
                End If
                Me.TXTRebate.Text = Format(vCheckRebate, "##,##0.00")
            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTRebateCalc()
        Dim vFollow3 As Double
        Dim vDiscRebateWord As String
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTFollowDisc3Res.Text <> "" Then
            If Me.TXTRebate.Text <> "" Then
                vDiscRebateWord = Me.TXTRebate.Text
                vFollow3 = Me.TXTFollowDisc3Res.Text

                vCheckCountDot = CheckDot(TXTRebate.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTRebate.Focus()
                    Exit Sub
                End If

                If InStr(vDiscRebateWord, "%") > 0 Then
                    If Len(vDiscRebateWord) > 1 Then
                        If InStr(vDiscRebateWord, "0") = 1 Then
                            Me.TXTRebate.Text = Microsoft.VisualBasic.Right(vDiscRebateWord, Len(vDiscRebateWord) - InStr(vDiscRebateWord, "0"))
                            vDiscRebateWord = Me.TXTRebate.Text
                        End If
                    End If
                    If InStr(vDiscRebateWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTRebate.Focus()
                        Exit Sub
                    End If

                    If InStr(vDiscRebateWord, ".%") > 0 Or InStr(vDiscRebateWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTRebate.Focus()
                        Exit Sub
                    End If

                    vDiscRebateWord = Microsoft.VisualBasic.Left(vDiscRebateWord, InStr(vDiscRebateWord, "%"))
                    Me.TXTRebate.Text = vDiscRebateWord
                    vDiscRebate = Microsoft.VisualBasic.Left(vDiscRebateWord, InStr(vDiscRebateWord, "%") - 1)
                    vDiscRebate1 = ((vFollow3 * vDiscRebate) / 100)
                    vRebateAmount = vFollow3 - vDiscRebate1
                Else
                    If Len(vDiscRebateWord) > 0 Then
                        If InStr(vDiscRebateWord, ".") = 1 And Len(vDiscRebateWord) = 1 Then
                            Me.TXTRebate.Focus()
                            Exit Sub
                        End If

                        If InStr(vDiscRebateWord, "0") = 1 And InStr(vDiscRebateWord, ".") = 2 And Len(vDiscRebateWord) = 2 Then
                            Me.TXTRebate.Focus()
                            Exit Sub
                        End If

                        If (InStr(vDiscRebateWord, ".") = 1) And Len(vDiscRebateWord) = 1 Then
                            Me.TXTRebate.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vDiscRebateWord, "0") = 1 And Len(vDiscRebateWord) > 1 Then
                            Me.TXTRebate.Text = Microsoft.VisualBasic.Right(vDiscRebateWord, Len(vDiscRebateWord) - InStr(vDiscRebateWord, "0"))
                        Else
                            Me.TXTRebate.Text = vDiscRebateWord
                        End If
                    End If
                    vDiscRebate = vDiscRebateWord
                    vDiscRebate1 = vDiscRebate
                    vRebateAmount = vFollow3 - vDiscRebate1
                End If

                If vRebateAmount < 0 Then
                    Me.TXTRebateRes.ForeColor = Color.Red
                Else
                    Me.TXTRebateRes.ForeColor = Color.Black
                End If

                Me.TXTRebateRes.Text = Format(vRebateAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Rebate")
        End If
    End Sub

    Private Sub TXTSpecialDisc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTSpecialDisc.LostFocus
        Dim vCheckSpecialDisc As Double
        Dim vCheckSpecialDiscRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTRebateRes.Text <> "" Then
            vCheckCountDot = CheckDot(TXTSpecialDisc.Text)
            If vCheckCountDot > 1 Then
                Me.TXTSpecialDisc.Focus()
                Exit Sub
            End If

            If InStr(TXTSpecialDisc.Text, "0") = 1 And InStr(TXTSpecialDisc.Text, ".") = 2 And Len(TXTSpecialDisc.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTSpecialDisc.Text, ".") = 1 And Len(TXTSpecialDisc.Text) = 1 Then
                Me.TXTSpecialDisc.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTSpecialDisc.Text, "%") = 1 Then
                Me.TXTSpecialDisc.Focus()
            End If
            If InStr(TXTSpecialDisc.Text, ".%") > 0 Or InStr(TXTSpecialDisc.Text, "%.") > 0 Then
                Me.TXTSpecialDisc.Focus()
                Exit Sub
            End If

            If Me.TXTSpecialDisc.Text = "" Then
                Me.TXTSpecialDisc.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTSpecialDisc.Text, "%") = 0 Then
                vCheckSpecialDisc = TXTSpecialDisc.Text
                If vCheckSpecialDisc = 0 Then
                    vCheckSpecialDiscRes = Me.TXTRebateRes.Text
                    Me.TXTSpecialDiscRes.Text = Format(vCheckSpecialDiscRes, "##,##0.00")
                End If
                Me.TXTSpecialDisc.Text = Format(vCheckSpecialDisc, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTSpecialDiscCalc()
        Dim vRebate As Double
        Dim vDiscSpecialWord As String
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTRebateRes.Text <> "" Then
            If Me.TXTSpecialDisc.Text <> "" Then
                vDiscSpecialWord = Me.TXTSpecialDisc.Text
                vRebate = Me.TXTRebateRes.Text

                vCheckCountDot = CheckDot(TXTSpecialDisc.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTSpecialDisc.Focus()
                    Exit Sub
                End If

                If InStr(vDiscSpecialWord, "%") > 0 Then
                    If Len(vDiscSpecialWord) > 1 Then
                        If InStr(vDiscSpecialWord, "0") = 1 Then
                            Me.TXTSpecialDisc.Text = Microsoft.VisualBasic.Right(vDiscSpecialWord, Len(vDiscSpecialWord) - InStr(vDiscSpecialWord, "0"))
                            vDiscSpecialWord = Me.TXTSpecialDisc.Text
                        End If
                    End If

                    If InStr(vDiscSpecialWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTSpecialDisc.Focus()
                        Exit Sub
                    End If

                    If InStr(vDiscSpecialWord, ".%") > 0 Or InStr(vDiscSpecialWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTSpecialDisc.Focus()
                        Exit Sub
                    End If

                    vDiscSpecialWord = Microsoft.VisualBasic.Left(vDiscSpecialWord, InStr(vDiscSpecialWord, "%"))
                    Me.TXTSpecialDisc.Text = vDiscSpecialWord
                    vDiscSpecial = Microsoft.VisualBasic.Left(vDiscSpecialWord, InStr(vDiscSpecialWord, "%") - 1)
                    vDiscSpecial1 = ((vRebate * vDiscSpecial) / 100)
                    vRebateAmount = vRebate - vDiscSpecial1
                Else
                    If Len(vDiscSpecialWord) > 0 Then
                        If InStr(vDiscSpecialWord, ".") = 1 And Len(vDiscSpecialWord) = 1 Then
                            Me.TXTSpecialDisc.Focus()
                            Exit Sub
                        End If

                        If InStr(vDiscSpecialWord, "0") = 1 And InStr(vDiscSpecialWord, ".") = 2 And Len(vDiscSpecialWord) = 2 Then
                            Me.TXTSpecialDisc.Focus()
                            Exit Sub
                        End If

                        If (InStr(vDiscSpecialWord, ".") = 1) And Len(vDiscSpecialWord) = 1 Then
                            Me.TXTSpecialDisc.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vDiscSpecialWord, "0") = 1 And Len(vDiscSpecialWord) > 1 Then
                            Me.TXTSpecialDisc.Text = Microsoft.VisualBasic.Right(vDiscSpecialWord, Len(vDiscSpecialWord) - InStr(vDiscSpecialWord, "0"))
                        Else
                            Me.TXTSpecialDisc.Text = vDiscSpecialWord
                        End If
                    End If
                    vDiscSpecial = vDiscSpecialWord
                    vDiscSpecial1 = vDiscSpecial
                    vRebateAmount = vRebate - vDiscSpecial1
                End If

                If vRebateAmount < 0 Then
                    Me.TXTSpecialDiscRes.ForeColor = Color.Red
                Else
                    Me.TXTSpecialDiscRes.ForeColor = Color.Black
                End If

                Me.TXTSpecialDiscRes.Text = Format(vRebateAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message SpecialDisc")
        End If
    End Sub

    Private Sub TXTMissProfit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTMissProfit.LostFocus
        Dim vCheckMissProfit As Double
        Dim vCheckMissProfitRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTSpecialDiscRes.Text <> "" Then
            vCheckCountDot = CheckDot(TXTMissProfit.Text)
            If vCheckCountDot > 1 Then
                Me.TXTMissProfit.Focus()
                Exit Sub
            End If

            If InStr(TXTMissProfit.Text, "0") = 1 And InStr(TXTMissProfit.Text, ".") = 2 And Len(TXTMissProfit.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTMissProfit.Text, ".") = 1 And Len(TXTMissProfit.Text) = 1 Then
                Me.TXTMissProfit.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTMissProfit.Text, "%") = 1 Then
                Me.TXTMissProfit.Focus()
            End If
            If InStr(TXTMissProfit.Text, ".%") > 0 Or InStr(TXTMissProfit.Text, "%.") > 0 Then
                Me.TXTMissProfit.Focus()
                Exit Sub
            End If

            If Me.TXTMissProfit.Text = "" Then
                Me.TXTMissProfit.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTMissProfit.Text, "%") = 0 Then
                vCheckMissProfit = TXTMissProfit.Text
                If vCheckMissProfit = 0 Then
                    vCheckMissProfitRes = Me.TXTSpecialDiscRes.Text
                    Me.TXTMissProfitRes.Text = Format(vCheckMissProfitRes, "##,##0.00")
                End If
                Me.TXTMissProfit.Text = Format(vCheckMissProfit, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTMissProfitCalc()
        Dim vSpecial As Double
        Dim vDiscMissProfitWord As String
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTSpecialDiscRes.Text <> "" Then
            If Me.TXTMissProfit.Text <> "" Then
                vDiscMissProfitWord = Me.TXTMissProfit.Text
                vSpecial = Me.TXTSpecialDiscRes.Text

                vCheckCountDot = CheckDot(TXTMissProfit.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTMissProfit.Focus()
                    Exit Sub
                End If

                If InStr(vDiscMissProfitWord, "%") > 0 Then
                    If Len(vDiscMissProfitWord) > 1 Then
                        If InStr(vDiscMissProfitWord, "0") = 1 Then
                            Me.TXTMissProfit.Text = Microsoft.VisualBasic.Right(vDiscMissProfitWord, Len(vDiscMissProfitWord) - InStr(vDiscMissProfitWord, "0"))
                            vDiscMissProfitWord = Me.TXTMissProfit.Text
                        End If
                    End If

                    If InStr(vDiscMissProfitWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTMissProfit.Focus()
                        Exit Sub
                    End If

                    If InStr(vDiscMissProfitWord, ".%") > 0 Or InStr(vDiscMissProfitWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTMissProfit.Focus()
                        Exit Sub
                    End If

                    vDiscMissProfitWord = Microsoft.VisualBasic.Left(vDiscMissProfitWord, InStr(vDiscMissProfitWord, "%"))
                    Me.TXTMissProfit.Text = vDiscMissProfitWord
                    vDiscMissProfit = Microsoft.VisualBasic.Left(vDiscMissProfitWord, InStr(vDiscMissProfitWord, "%") - 1)
                    vDiscMissProfit1 = ((vSpecial * vDiscMissProfit) / 100)
                    vMissProfitAmount = vSpecial - vDiscMissProfit1
                Else
                    If Len(vDiscMissProfitWord) > 0 Then
                        If InStr(vDiscMissProfitWord, ".") = 1 And Len(vDiscMissProfitWord) = 1 Then
                            Me.TXTMissProfit.Focus()
                            Exit Sub
                        End If

                        If InStr(vDiscMissProfitWord, "0") = 1 And InStr(vDiscMissProfitWord, ".") = 2 And Len(vDiscMissProfitWord) = 2 Then
                            Me.TXTMissProfit.Focus()
                            Exit Sub
                        End If

                        If (InStr(vDiscMissProfitWord, ".") = 1) And Len(vDiscMissProfitWord) = 1 Then
                            Me.TXTMissProfit.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vDiscMissProfitWord, "0") = 1 And Len(vDiscMissProfitWord) > 1 Then
                            Me.TXTMissProfit.Text = Microsoft.VisualBasic.Right(vDiscMissProfitWord, Len(vDiscMissProfitWord) - InStr(vDiscMissProfitWord, "0"))
                        Else
                            Me.TXTMissProfit.Text = vDiscMissProfitWord
                        End If
                    End If
                    vDiscMissProfit = vDiscMissProfitWord
                    vDiscMissProfit1 = vDiscMissProfit
                    vMissProfitAmount = vSpecial - vDiscMissProfit1
                End If

                If vMissProfitAmount < 0 Then
                    Me.TXTMissProfitRes.ForeColor = Color.Red
                Else
                    Me.TXTMissProfitRes.ForeColor = Color.Black
                End If

                Me.TXTMissProfitRes.Text = Format(vMissProfitAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message MissProfit")
        End If
    End Sub

    Private Sub TXTSend_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTSend.LostFocus
        Dim vCheckSend As Double
        Dim vCheckSendRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTMissProfitRes.Text <> "" Then

            vCheckCountDot = CheckDot(TXTSend.Text)
            If vCheckCountDot > 1 Then
                Me.TXTSend.Focus()
                Exit Sub
            End If

            If InStr(TXTSend.Text, "0") = 1 And InStr(TXTSend.Text, ".") = 2 And Len(TXTSend.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTSend.Text, ".") = 1 And Len(TXTSend.Text) = 1 Then
                Me.TXTSend.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTSend.Text, "%") = 1 Then
                Me.TXTSend.Focus()
            End If
            If InStr(TXTSend.Text, ".%") > 0 Or InStr(TXTSend.Text, "%.") > 0 Then
                Me.TXTSend.Focus()
                Exit Sub
            End If

            If Me.TXTSend.Text = "" Then
                Me.TXTSend.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTSend.Text, "%") = 0 Then
                vCheckSend = TXTSend.Text
                If vCheckSend = 0 Then
                    vCheckSendRes = Me.TXTMissProfitRes.Text
                    Me.TXTSendRes.Text = Format(vCheckSendRes, "##,##0.00")
                End If
                Me.TXTSend.Text = Format(vCheckSend, "##,##0.00")
            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTSendCalc()
        Dim vMissProfit As Double
        Dim vSendWord As String
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTMissProfitRes.Text <> "" Then
            If Me.TXTSend.Text <> "" Then
                vSendWord = Me.TXTSend.Text
                vMissProfit = Me.TXTMissProfitRes.Text

                vCheckCountDot = CheckDot(TXTSend.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTSend.Focus()
                    Exit Sub
                End If

                If InStr(vSendWord, "%") > 0 Then
                    If Len(vSendWord) > 1 Then
                        If InStr(vSendWord, "0") = 1 Then
                            Me.TXTSend.Text = Microsoft.VisualBasic.Right(vSendWord, Len(vSendWord) - InStr(vSendWord, "0"))
                            vSendWord = Me.TXTSend.Text
                        End If
                    End If
                    If InStr(vSendWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTSend.Focus()
                        Exit Sub
                    End If

                    If InStr(vSendWord, ".%") > 0 Or InStr(vSendWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTSend.Focus()
                        Exit Sub
                    End If

                    vSendWord = Microsoft.VisualBasic.Left(vSendWord, InStr(vSendWord, "%"))
                    Me.TXTSend.Text = vSendWord
                    vSend = Microsoft.VisualBasic.Left(vSendWord, InStr(vSendWord, "%") - 1)
                    vSend1 = ((vMissProfit * vSend) / 100)
                    vSendAmount = vMissProfit + vSend1
                Else
                    If Len(vSendWord) > 0 Then
                        If InStr(vSendWord, ".") = 1 And Len(vSendWord) = 1 Then
                            Me.TXTSend.Focus()
                            Exit Sub
                        End If

                        If InStr(vSendWord, "0") = 1 And InStr(vSendWord, ".") = 2 And Len(vSendWord) = 2 Then
                            Me.TXTSend.Focus()
                            Exit Sub
                        End If

                        If (InStr(vSendWord, ".") = 1) And Len(vSendWord) = 1 Then
                            Me.TXTSend.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vSendWord, "0") = 1 And Len(vSendWord) > 1 Then
                            Me.TXTSend.Text = Microsoft.VisualBasic.Right(vSendWord, Len(vSendWord) - InStr(vSendWord, "0"))
                        Else
                            Me.TXTSend.Text = vSendWord
                        End If
                    End If
                    vSend = vSendWord
                    vSend1 = vSend
                    vSendAmount = vMissProfit + vSend1
                End If

                If vSendAmount < 0 Then
                    Me.TXTSendRes.ForeColor = Color.Red
                Else
                    Me.TXTSendRes.ForeColor = Color.Black
                End If

                Me.TXTSendRes.Text = Format(vSendAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Send")
        End If
    End Sub

    Private Sub TXTCustSend_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTCustSend.LostFocus
        Dim vCheckCustSend As Double
        Dim vCheckCustSendRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTSendRes.Text <> "" Then

            vCheckCountDot = CheckDot(TXTCustSend.Text)
            If vCheckCountDot > 1 Then
                Me.TXTCustSend.Focus()
                Exit Sub
            End If

            If InStr(TXTCustSend.Text, "0") = 1 And InStr(TXTCustSend.Text, ".") = 2 And Len(TXTCustSend.Text) = 2 Then
                Exit Sub
            End If

            If InStr(TXTCustSend.Text, ".") = 1 And Len(TXTCustSend.Text) = 1 Then
                Me.TXTCustSend.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTCustSend.Text, "%") = 1 Then
                Me.TXTCustSend.Focus()
            End If

            If InStr(TXTCustSend.Text, ".%") > 0 Or InStr(TXTCustSend.Text, "%.") > 0 Then
                Me.TXTCustSend.Focus()
                Exit Sub
            End If

            If Me.TXTCustSend.Text = "" Then
                Me.TXTCustSend.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTCustSend.Text, "%") = 0 Then
                vCheckCustSend = TXTCustSend.Text
                If vCheckCustSend = 0 Then
                    vCheckCustSendRes = Me.TXTCustSendRes.Text
                    Me.TXTCustSendRes.Text = Format(vCheckCustSendRes, "##,##0.00")
                End If
                Me.TXTCustSend.Text = Format(vCheckCustSend, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTCustSendCalc()
        Dim vSend As Double
        Dim vCustSendWord As String
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTSendRes.Text <> "" Then
            If Me.TXTCustSend.Text <> "" Then
                vCustSendWord = Me.TXTCustSend.Text
                vSend = Me.TXTSendRes.Text

                vCheckCountDot = CheckDot(TXTCustSend.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTCustSend.Focus()
                    Exit Sub
                End If

                If InStr(vCustSendWord, "%") > 0 Then
                    If Len(vCustSendWord) > 1 Then
                        If InStr(vCustSendWord, "0") = 1 Then
                            Me.TXTCustSend.Text = Microsoft.VisualBasic.Right(vCustSendWord, Len(vCustSendWord) - InStr(vCustSendWord, "0"))
                            vCustSendWord = Me.TXTCustSend.Text
                        End If
                    End If
                    If InStr(vCustSendWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTCustSend.Focus()
                        Exit Sub
                    End If

                    If InStr(vCustSendWord, ".%") > 0 Or InStr(vCustSendWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTCustSend.Focus()
                        Exit Sub
                    End If

                    vCustSendWord = Microsoft.VisualBasic.Left(vCustSendWord, InStr(vCustSendWord, "%"))
                    Me.TXTCustSend.Text = vCustSendWord
                    vCustSend = Microsoft.VisualBasic.Left(vCustSendWord, InStr(vCustSendWord, "%") - 1)
                    vCustSend1 = ((vSend * vCustSend) / 100)
                    vCustSendAmount = vSend + vCustSend1
                Else
                    If Len(vCustSendWord) > 0 Then
                        If InStr(vCustSendWord, ".") = 1 And Len(vCustSendWord) = 1 Then
                            Me.TXTCustSend.Focus()
                            Exit Sub
                        End If

                        If InStr(vCustSendWord, "0") = 1 And InStr(vCustSendWord, ".") = 2 And Len(vCustSendWord) = 2 Then
                            Me.TXTCustSend.Focus()
                            Exit Sub
                        End If

                        If (InStr(vCustSendWord, ".") = 1) And Len(vCustSendWord) = 1 Then
                            Me.TXTCustSend.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vCustSendWord, "0") = 1 And Len(vCustSendWord) > 1 Then
                            Me.TXTCustSend.Text = Microsoft.VisualBasic.Right(vCustSendWord, Len(vCustSendWord) - InStr(vCustSendWord, "0"))
                        Else
                            Me.TXTCustSend.Text = vCustSendWord
                        End If
                    End If
                    vCustSend = vCustSendWord
                    vCustSend1 = vCustSend
                    vCustSendAmount = vSend + vCustSend1
                End If

                If vCustSendAmount < 0 Then
                    Me.TXTCustSendRes.ForeColor = Color.Red
                Else
                    Me.TXTCustSendRes.ForeColor = Color.Black
                End If

                Me.TXTCustSendRes.Text = Format(vCustSendAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message CustSend")
        End If
    End Sub

    Private Sub TXTAdvertise_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTAdvertise.LostFocus
        Dim vCheckAdvertise As Double
        Dim vCheckAdvertiseRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTCustSendRes.Text <> "" Then
            vCheckCountDot = CheckDot(TXTAdvertise.Text)
            If vCheckCountDot > 1 Then
                Me.TXTAdvertise.Focus()
                Exit Sub
            End If

            If InStr(TXTAdvertise.Text, "0") = 1 And InStr(TXTAdvertise.Text, ".") = 2 And Len(TXTAdvertise.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTAdvertise.Text, ".") = 1 And Len(TXTAdvertise.Text) = 1 Then
                Me.TXTAdvertise.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTAdvertise.Text, "%") = 1 Then
                Me.TXTAdvertise.Focus()
            End If
            If InStr(TXTAdvertise.Text, ".%") > 0 Or InStr(TXTAdvertise.Text, "%.") > 0 Then
                Me.TXTAdvertise.Focus()
                Exit Sub
            End If

            If Me.TXTAdvertise.Text = "" Then
                Me.TXTAdvertise.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTAdvertise.Text, "%") = 0 Then
                vCheckAdvertise = TXTAdvertise.Text
                If vCheckAdvertise = 0 Then
                    vCheckAdvertiseRes = Me.TXTCustSendRes.Text
                    Me.TXTAdvertiseRes.Text = Format(vCheckAdvertiseRes, "##,##0.00")
                End If
                Me.TXTAdvertise.Text = Format(vCheckAdvertise, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTAdvertiseCalc()
        Dim vCustSend As Double
        Dim vAdvertiseWord As String
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTCustSendRes.Text <> "" Then
            If Me.TXTAdvertise.Text <> "" Then
                vAdvertiseWord = Me.TXTAdvertise.Text
                vCustSend = Me.TXTCustSendRes.Text

                vCheckCountDot = CheckDot(TXTAdvertise.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTAdvertise.Focus()
                    Exit Sub
                End If

                If InStr(vAdvertiseWord, "%") > 0 Then
                    If Len(vAdvertiseWord) > 1 Then
                        If InStr(vAdvertiseWord, "0") = 1 Then
                            Me.TXTAdvertise.Text = Microsoft.VisualBasic.Right(vAdvertiseWord, Len(vAdvertiseWord) - InStr(vAdvertiseWord, "0"))
                            vAdvertiseWord = Me.TXTAdvertise.Text
                        End If
                    End If
                    If InStr(vAdvertiseWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTAdvertise.Focus()
                        Exit Sub
                    End If

                    If InStr(vAdvertiseWord, ".%") > 0 Or InStr(vAdvertiseWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTAdvertise.Focus()
                        Exit Sub
                    End If

                    vAdvertiseWord = Microsoft.VisualBasic.Left(vAdvertiseWord, InStr(vAdvertiseWord, "%"))
                    Me.TXTAdvertise.Text = vAdvertiseWord
                    vAdvertise = Microsoft.VisualBasic.Left(vAdvertiseWord, InStr(vAdvertiseWord, "%") - 1)
                    vAdvertise1 = ((vCustSend * vAdvertise) / 100)
                    vAdvertiseAmount = vCustSend + vAdvertise1
                Else
                    If Len(vAdvertiseWord) > 0 Then
                        If InStr(vAdvertiseWord, ".") = 1 And Len(vAdvertiseWord) = 1 Then
                            Me.TXTAdvertise.Focus()
                            Exit Sub
                        End If

                        If InStr(vAdvertiseWord, "0") = 1 And InStr(vAdvertiseWord, ".") = 2 And Len(vAdvertiseWord) = 2 Then
                            Me.TXTAdvertise.Focus()
                            Exit Sub
                        End If

                        If (InStr(vAdvertiseWord, "0") = 1 And InStr(vAdvertiseWord, ".") = 1) And Len(vAdvertiseWord) = 1 Then
                            Me.TXTAdvertise.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vAdvertiseWord, "0") = 1 And Len(vAdvertiseWord) > 1 Then
                            Me.TXTAdvertise.Text = Microsoft.VisualBasic.Right(vAdvertiseWord, Len(vAdvertiseWord) - InStr(vAdvertiseWord, "0"))
                        Else
                            Me.TXTAdvertise.Text = vAdvertiseWord
                        End If
                    End If
                    vAdvertise = vAdvertiseWord
                    vAdvertise1 = vAdvertise
                    vAdvertiseAmount = vCustSend + vAdvertise1
                End If
                If vAdvertiseAmount < 0 Then
                    Me.TXTAdvertiseRes.ForeColor = Color.Red
                Else
                    Me.TXTAdvertiseRes.ForeColor = Color.Black
                End If

                Me.TXTAdvertiseRes.Text = Format(vAdvertiseAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Advertise")
        End If
    End Sub

    Private Sub TXTMarket_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTMarket.LostFocus
        Dim vCheckMarket As Double
        Dim vCheckMarketRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTAdvertiseRes.Text <> "" Then
            vCheckCountDot = CheckDot(TXTMarket.Text)
            If vCheckCountDot > 1 Then
                Me.TXTMarket.Focus()
                Exit Sub
            End If

            If InStr(TXTMarket.Text, "0") = 1 And InStr(TXTMarket.Text, ".") = 2 And Len(TXTMarket.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTMarket.Text, ".") = 1 And Len(TXTMarket.Text) = 1 Then
                Me.TXTMarket.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTMarket.Text, "%") = 1 Then
                Me.TXTMarket.Focus()
            End If

            If InStr(TXTMarket.Text, ".%") > 0 Or InStr(TXTMarket.Text, "%.") > 0 Then
                Me.TXTMarket.Focus()
                Exit Sub
            End If

            If Me.TXTMarket.Text = "" Then
                Me.TXTMarket.Text = Format(0, "##,##0.00")
                Exit Sub
            End If


            If InStr(TXTMarket.Text, "%") = 0 Then
                vCheckMarket = TXTMarket.Text
                If vCheckMarket = 0 Then
                    vCheckMarketRes = Me.TXTAdvertiseRes.Text
                    Me.TXTMarketRes.Text = Format(vCheckMarketRes, "##,##0.00")
                End If
                Me.TXTMarket.Text = Format(vCheckMarket, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTMarketCalc()
        Dim vMarketWord As String
        Dim vAdvertise As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTAdvertiseRes.Text <> "" Then
            If Me.TXTMarket.Text <> "" Then
                vMarketWord = Me.TXTMarket.Text
                vAdvertise = Me.TXTAdvertiseRes.Text

                vCheckCountDot = CheckDot(TXTMarket.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTMarket.Focus()
                    Exit Sub
                End If

                If InStr(vMarketWord, "%") > 0 Then
                    If Len(vMarketWord) > 1 Then
                        If InStr(vMarketWord, "0") = 1 Then
                            Me.TXTMarket.Text = Microsoft.VisualBasic.Right(vMarketWord, Len(vMarketWord) - InStr(vMarketWord, "0"))
                            vMarketWord = Me.TXTMarket.Text
                        End If
                    End If
                    If InStr(vMarketWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTMarket.Focus()
                        Exit Sub
                    End If

                    If InStr(vMarketWord, ".%") > 0 Or InStr(vMarketWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTMarket.Focus()
                        Exit Sub
                    End If

                    vMarketWord = Microsoft.VisualBasic.Left(vMarketWord, InStr(vMarketWord, "%"))
                    Me.TXTMarket.Text = vMarketWord
                    vMarket = Microsoft.VisualBasic.Left(vMarketWord, InStr(vMarketWord, "%") - 1)
                    vMarket1 = ((vAdvertise * vMarket) / 100)
                    vMarketAmount = vAdvertise + vMarket1
                Else
                    If Len(vMarketWord) > 0 Then
                        If InStr(vMarketWord, ".") = 1 And Len(vMarketWord) = 1 Then
                            Me.TXTMarket.Focus()
                            Exit Sub
                        End If

                        If InStr(vMarketWord, "0") = 1 And InStr(vMarketWord, ".") = 2 And Len(vMarketWord) = 2 Then
                            Me.TXTMarket.Focus()
                            Exit Sub
                        End If

                        If (InStr(vMarketWord, ".") = 1) And Len(vMarketWord) = 1 Then
                            Me.TXTMarket.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vMarketWord, "0") = 1 And Len(vMarketWord) > 1 Then
                            Me.TXTMarket.Text = Microsoft.VisualBasic.Right(vMarketWord, Len(vMarketWord) - InStr(vMarketWord, "0"))
                        Else
                            Me.TXTMarket.Text = vMarketWord
                        End If
                    End If
                    vMarket = vMarketWord
                    vMarket1 = vMarket
                    vMarketAmount = vAdvertise + vMarket1
                End If

                If vMarketAmount < 0 Then
                    Me.TXTMarketRes.ForeColor = Color.Red
                Else
                    Me.TXTMarketRes.ForeColor = Color.Black
                End If

                Me.TXTMarketRes.Text = Format(vMarketAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Market")
        End If
    End Sub

    Private Sub TXTTax_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTTax.LostFocus
        Dim vCheckTax As Double
        Dim vCheckTaxRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTMarketRes.Text <> "" Then

            vCheckCountDot = CheckDot(TXTTax.Text)
            If vCheckCountDot > 1 Then
                Me.TXTTax.Focus()
                Exit Sub
            End If

            If InStr(TXTTax.Text, "0") = 1 And InStr(TXTTax.Text, ".") = 2 And Len(TXTTax.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTTax.Text, ".") = 1 And Len(TXTTax.Text) = 1 Then
                Me.TXTTax.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTTax.Text, "%") = 1 Then
                Me.TXTTax.Focus()
            End If
            If InStr(TXTTax.Text, ".%") > 0 Or InStr(TXTTax.Text, "%.") > 0 Then
                Me.TXTTax.Focus()
                Exit Sub
            End If

            If Me.TXTTax.Text = "" Then
                Me.TXTTax.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTTax.Text, "%") = 0 Then
                vCheckTax = TXTTax.Text
                If vCheckTax = 0 Then
                    vCheckTaxRes = Me.TXTMarketRes.Text
                    Me.TXTTaxRes.Text = Format(vCheckTaxRes, "##,##0.00")
                End If
                Me.TXTTax.Text = Format(vCheckTax, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTTaxCalc()
        Dim vMarket As Double
        Dim vTaxWord As String
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTMarketRes.Text <> "" Then
            If Me.TXTTax.Text <> "" Then
                vTaxWord = Me.TXTTax.Text
                vMarket = Me.TXTMarketRes.Text

                vCheckCountDot = CheckDot(TXTTax.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTTax.Focus()
                    Exit Sub
                End If

                If InStr(vTaxWord, "%") > 0 Then
                    If Len(vTaxWord) > 1 Then
                        If InStr(vTaxWord, "0") = 1 Then
                            Me.TXTTax.Text = Microsoft.VisualBasic.Right(vTaxWord, Len(vTaxWord) - InStr(vTaxWord, "0"))
                            vTaxWord = Me.TXTTax.Text
                        End If
                    End If
                    If InStr(vTaxWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTTax.Focus()
                        Exit Sub
                    End If

                    If InStr(vTaxWord, ".%") > 0 Or InStr(vTaxWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTTax.Focus()
                        Exit Sub
                    End If

                    vTaxWord = Microsoft.VisualBasic.Left(vTaxWord, InStr(vTaxWord, "%"))
                    Me.TXTTax.Text = vTaxWord
                    vTax = Microsoft.VisualBasic.Left(vTaxWord, InStr(vTaxWord, "%") - 1)
                    vTax1 = ((vMarket * vTax) / 100)
                    vTaxAmount = vMarket + vTax1
                Else
                    If Len(vTaxWord) > 0 Then
                        If InStr(vTaxWord, ".") = 1 And Len(vTaxWord) = 1 Then
                            Me.TXTTax.Focus()
                            Exit Sub
                        End If

                        If InStr(vTaxWord, "0") = 1 And InStr(vTaxWord, ".") = 2 And Len(vTaxWord) = 2 Then
                            Me.TXTTax.Focus()
                            Exit Sub
                        End If

                        If (InStr(vTaxWord, ".") = 1) And Len(vTaxWord) = 1 Then
                            Me.TXTTax.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vTaxWord, "0") = 1 And Len(vTaxWord) > 1 Then
                            Me.TXTTax.Text = Microsoft.VisualBasic.Right(vTaxWord, Len(vTaxWord) - InStr(vTaxWord, "0"))
                        Else
                            Me.TXTTax.Text = vTaxWord
                        End If
                    End If
                    vTax = vTaxWord
                    vTax1 = vTax
                    vTaxAmount = vMarket + vTax1
                End If

                If vTaxAmount < 0 Then
                    Me.TXTTaxRes.ForeColor = Color.Red
                Else
                    Me.TXTTaxRes.ForeColor = Color.Black
                End If

                Me.TXTTaxRes.Text = Format(vTaxAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Tax")
        End If
    End Sub

    Private Sub TXTInstall_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTInstall.LostFocus
        Dim vCheckInstall As Double
        Dim vCheckInstallRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTTaxRes.Text <> "" Then
            vCheckCountDot = CheckDot(TXTInstall.Text)
            If vCheckCountDot > 1 Then
                Me.TXTInstall.Focus()
                Exit Sub
            End If

            If InStr(TXTInstall.Text, "0") = 1 And InStr(TXTInstall.Text, ".") = 2 And Len(TXTInstall.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTInstall.Text, ".") = 1 And Len(TXTInstall.Text) = 1 Then
                Me.TXTInstall.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTInstall.Text, "%") = 1 Then
                Me.TXTInstall.Focus()
            End If
            If InStr(TXTInstall.Text, ".%") > 0 Or InStr(TXTInstall.Text, "%.") > 0 Then
                Me.TXTInstall.Focus()
                Exit Sub
            End If

            If Me.TXTInstall.Text = "" Then
                Me.TXTInstall.Text = Format(0, "##,##0.00")
                Exit Sub
            End If


            If InStr(TXTInstall.Text, "%") = 0 Then
                vCheckInstall = TXTInstall.Text
                If vCheckInstall = 0 Then
                    vCheckInstallRes = Me.TXTTaxRes.Text
                    Me.TXTInstallRes.Text = Format(vCheckInstallRes, "##,##0.00")

                End If
                Me.TXTInstall.Text = Format(vCheckInstall, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTInstallCalc()
        Dim vInstallWord As String
        Dim vTax As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTTaxRes.Text <> "" Then
            If Me.TXTInstall.Text <> "" Then
                vInstallWord = Me.TXTInstall.Text
                vTax = Me.TXTTaxRes.Text

                vCheckCountDot = CheckDot(TXTInstall.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTInstall.Focus()
                    Exit Sub
                End If

                If InStr(vInstallWord, "%") > 0 Then
                    If Len(vInstallWord) > 1 Then
                        If InStr(vInstallWord, "0") = 1 Then
                            Me.TXTInstall.Text = Microsoft.VisualBasic.Right(vInstallWord, Len(vInstallWord) - InStr(vInstallWord, "0"))
                            vInstallWord = Me.TXTInstall.Text
                        End If
                    End If
                    If InStr(vInstallWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTInstall.Focus()
                        Exit Sub
                    End If

                    If InStr(vInstallWord, ".%") > 0 Or InStr(vInstallWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTInstall.Focus()
                        Exit Sub
                    End If

                    vInstallWord = Microsoft.VisualBasic.Left(vInstallWord, InStr(vInstallWord, "%"))
                    Me.TXTInstall.Text = vInstallWord
                    vInstall = Microsoft.VisualBasic.Left(vInstallWord, InStr(vInstallWord, "%") - 1)
                    vInstall1 = ((vTax * vInstall) / 100)
                    vInstallAmount = vTax + vInstall1
                Else
                    If Len(vInstallWord) > 0 Then
                        If InStr(vInstallWord, ".") = 1 And Len(vInstallWord) = 1 Then
                            Me.TXTInstall.Focus()
                            Exit Sub
                        End If

                        If InStr(vInstallWord, "0") = 1 And InStr(vInstallWord, ".") = 2 And Len(vInstallWord) = 2 Then
                            Me.TXTInstall.Focus()
                            Exit Sub
                        End If

                        If (InStr(vInstallWord, ".") = 1) And Len(vInstallWord) = 1 Then
                            Me.TXTInstall.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vInstallWord, "0") = 1 And Len(vInstallWord) > 1 Then
                            Me.TXTInstall.Text = Microsoft.VisualBasic.Right(vInstallWord, Len(vInstallWord) - InStr(vInstallWord, "0"))
                        Else
                            Me.TXTInstall.Text = vInstallWord
                        End If
                    End If
                    vInstall = vInstallWord
                    vInstall1 = vInstall
                    vInstallAmount = vTax + vInstall1
                End If

                If vInstallAmount < 0 Then
                    Me.TXTInstallRes.ForeColor = Color.Red
                Else
                    Me.TXTInstallRes.ForeColor = Color.Black
                End If

                Me.TXTInstallRes.Text = Format(vInstallAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Install")
        End If
    End Sub

    Private Sub TXTService_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTService.LostFocus
        Dim vCheckService As Double
        Dim vCheckServiceRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTInstallRes.Text <> "" Then
            vCheckCountDot = CheckDot(TXTService.Text)
            If vCheckCountDot > 1 Then
                Me.TXTService.Focus()
                Exit Sub
            End If

            If InStr(TXTService.Text, "0") = 1 And InStr(TXTService.Text, ".") = 2 And Len(TXTService.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTService.Text, ".") = 1 And Len(TXTService.Text) = 1 Then
                Me.TXTService.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTService.Text, "%") = 1 Then
                Me.TXTService.Focus()
            End If
            If InStr(TXTService.Text, ".%") > 0 Or InStr(TXTService.Text, "%.") > 0 Then
                Me.TXTService.Focus()
                Exit Sub
            End If

            If Me.TXTService.Text = "" Then
                Me.TXTService.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTService.Text, "%") = 0 Then
                If Int(Me.TXTService.Text) = 0 Then
                    Me.TXTServiceRes.Text = Format(Int(Me.TXTInstallRes.Text), "##,##0.00")
                End If
            End If

            If InStr(TXTService.Text, "%") = 0 Then
                vCheckService = TXTService.Text
                If vCheckService = 0 Then
                    vCheckServiceRes = Me.TXTInstallRes.Text
                    Me.TXTServiceRes.Text = Format(vCheckServiceRes, "##,##0.00")

                End If

                Me.TXTService.Text = Format(vCheckService, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTServiceCalc()
        Dim vServiceWord As String
        Dim vInstall As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTInstallRes.Text <> "" Then
            If Me.TXTService.Text <> "" Then
                vServiceWord = Me.TXTService.Text
                vInstall = Me.TXTInstallRes.Text

                vCheckCountDot = CheckDot(TXTService.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTService.Focus()
                    Exit Sub
                End If

                If InStr(vServiceWord, "%") > 0 Then
                    If Len(vServiceWord) > 1 Then
                        If InStr(vServiceWord, "0") = 1 Then
                            Me.TXTService.Text = Microsoft.VisualBasic.Right(vServiceWord, Len(vServiceWord) - InStr(vServiceWord, "0"))
                            vServiceWord = Me.TXTService.Text
                        End If
                    End If
                    If InStr(vServiceWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTService.Focus()
                        Exit Sub
                    End If

                    If InStr(vServiceWord, ".%") > 0 Or InStr(vServiceWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTService.Focus()
                        Exit Sub
                    End If

                    vServiceWord = Microsoft.VisualBasic.Left(vServiceWord, InStr(vServiceWord, "%"))
                    Me.TXTService.Text = vServiceWord
                    vService = Microsoft.VisualBasic.Left(vServiceWord, InStr(vServiceWord, "%") - 1)
                    vService1 = ((vInstall * vService) / 100)
                    vServiceAmount = vInstall + vService1
                Else
                    If Len(vServiceWord) > 0 Then
                        If InStr(vServiceWord, ".") = 1 And Len(vServiceWord) = 1 Then
                            Me.TXTService.Focus()
                            Exit Sub
                        End If

                        If InStr(vServiceWord, "0") = 1 And InStr(vServiceWord, ".") = 2 And Len(vServiceWord) = 2 Then
                            Me.TXTService.Focus()
                            Exit Sub
                        End If

                        If (InStr(vServiceWord, ".") = 1) And Len(vServiceWord) = 1 Then
                            Me.TXTService.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vServiceWord, "0") = 1 And Len(vServiceWord) > 1 Then
                            Me.TXTService.Text = Microsoft.VisualBasic.Right(vServiceWord, Len(vServiceWord) - InStr(vServiceWord, "0"))
                        Else
                            Me.TXTService.Text = vServiceWord
                        End If
                    End If
                    vService = vServiceWord
                    vService1 = vService
                    vServiceAmount = vInstall + vService1
                End If

                If vServiceAmount < 0 Then
                    Me.TXTServiceRes.ForeColor = Color.Red
                Else
                    Me.TXTServiceRes.ForeColor = Color.Black
                End If

                Me.TXTServiceRes.Text = Format(vServiceAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Service")
        End If
    End Sub

    Private Sub TXTPoint_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTPoint.LostFocus
        Dim vCheckPoint As Double
        Dim vCheckPointRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTServiceRes.Text <> "" Then
            vCheckCountDot = CheckDot(TXTPoint.Text)
            If vCheckCountDot > 1 Then
                Me.TXTPoint.Focus()
                Exit Sub
            End If

            If InStr(TXTPoint.Text, "0") = 1 And InStr(TXTPoint.Text, ".") = 2 And Len(TXTPoint.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTPoint.Text, ".") = 1 And Len(TXTPoint.Text) = 1 Then
                Me.TXTPoint.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTPoint.Text, "%") = 1 Then
                Me.TXTPoint.Focus()
            End If
            If InStr(TXTPoint.Text, ".%") > 0 Or InStr(TXTPoint.Text, "%.") > 0 Then
                Me.TXTPoint.Focus()
                Exit Sub
            End If

            If Me.TXTPoint.Text = "" Then
                Me.TXTPoint.Text = Format(0, "##,##0.00")
                Exit Sub
            End If


            If InStr(TXTPoint.Text, "%") = 0 Then
                If Int(Me.TXTPoint.Text) = 0 Then
                    Me.TXTPointRes.Text = Format(Int(Me.TXTServiceRes.Text), "##,##0.00")
                End If
            End If

            If InStr(TXTPoint.Text, "%") = 0 Then
                vCheckPoint = TXTPoint.Text
                If vCheckPoint = 0 Then
                    vCheckPointRes = Me.TXTServiceRes.Text
                    Me.TXTPointRes.Text = Format(vCheckPointRes, "##,##0.00")

                End If

                Me.TXTPoint.Text = Format(vCheckPoint, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTPointCalc()
        Dim vPointWord As String
        Dim vMarketCost As Double

        On Error GoTo ErrDescription

        If Me.TXTPriceSet.Text <> "" And Me.TXTDO.Text <> "" And Me.TXTServiceRes.Text <> "" Then
            If Me.TXTPoint.Text <> "" Then
                vPointWord = Me.TXTPoint.Text
                vMarketCost = Me.TXTServiceRes.Text
                vPriceSetAmount = Me.TXTPriceSet.Text

                If InStr(Me.TXTPoint.Text, "%") > 0 Then
                    vPoint = Microsoft.VisualBasic.Left(vPointWord, InStr(vPointWord, "%") - 1)
                    vPoint1 = ((vPriceSetAmount * vPoint) / 100)
                    vPointAmount = vMarketCost + vPoint1
                Else
                    vPoint = vPointWord
                    vPoint1 = vPoint
                    vPointAmount = vPoint1
                End If

                If vPointAmount < 0 Then
                    Me.TXTPointRes.ForeColor = Color.Red
                Else
                    Me.TXTPointRes.ForeColor = Color.Black
                End If
                Me.TXTPointRes.Text = Format(vPointAmount, "##,##0.00")
            End If
            End If


ErrDescription:
            If Err.Description <> "" Then
                MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Point")
            End If
    End Sub
    Private Sub TXTDiscMemberCalc()
        Dim vDiscMemberWord As String
        Dim vPoint As Double

        On Error GoTo ErrDescription

        If Me.TXTPriceSet.Text <> "" And Me.TXTDO.Text <> "" And Me.TXTPointRes.Text <> "" Then
            If Me.TXTDiscMember.Text <> "" Then

                vDiscMemberWord = Me.TXTDiscMember.Text
                vPriceSetAmount = Me.TXTPriceSet.Text
                vPoint = Me.TXTPointRes.Text

                If InStr(Me.TXTDiscMember.Text, "%") > 0 Then

                    vDiscMember = Microsoft.VisualBasic.Left(vDiscMemberWord, InStr(vDiscMemberWord, "%") - 1)
                    vDiscMember1 = ((vPriceSetAmount * vDiscMember) / 100)
                    vDiscMemberAmount = vPoint + vDiscMember1
                Else
                    vDiscMember = vDiscMemberWord
                    vDiscMember1 = vDiscMember
                    vDiscMemberAmount = vDiscMember1
                End If

                If vDiscMemberAmount < 0 Then
                    Me.TXTDiscMemberRes.ForeColor = Color.Red
                Else
                    Me.TXTDiscMemberRes.ForeColor = Color.Black
                End If
                Me.TXTDiscMemberRes.Text = Format(vDiscMemberAmount, "##,##0.00")

            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message DiscMember")
        End If
    End Sub
    Private Sub TXTCreditCardCalc()
        Dim vCreditCardWord As String
        Dim vService As Double

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTPriceSet.Text <> "" And Me.TXTServiceRes.Text <> "" Then
            'If Me.TXTDiscMember.Text <> "" Then
            '    vCreditCardWord = Me.TXTDiscMember.Text
            '    vService = Me.TXTServiceRes.Text
            '    vPriceSetAmount = Me.TXTPriceSet.Text

            '    vCreditCard = vCreditCardWord
            '    vCreditCard1 = ((vPriceSetAmount * vCreditCard) / 100)
            '    vCreditCardAmount = vService + vCreditCard1

            '    If vCreditCardAmount < 0 Then
            '        Me.TXTDiscMemberRes.ForeColor = Color.Red
            '    Else
            '        Me.TXTDiscMemberRes.ForeColor = Color.Black
            '    End If

            '    Me.TXTDiscMemberRes.Text = Format(vCreditCardAmount, "##,##0.00")

            'End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub
    Private Sub TXTTarget_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTTarget.LostFocus
        Dim vCheckTarget As Double
        Dim vCheckTargetRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTPointRes.Text <> "" Then

            vCheckCountDot = CheckDot(Me.TXTTarget.Text)
            If vCheckCountDot > 1 Then
                Me.TXTTarget.Focus()
                Exit Sub
            End If

            If InStr(TXTTarget.Text, "0") = 1 And InStr(TXTTarget.Text, ".") = 2 And Len(TXTTarget.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTTarget.Text, ".") = 1 And Len(TXTTarget.Text) = 1 Then
                Me.TXTTarget.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTTarget.Text, "%") = 1 Then
                Me.TXTTarget.Focus()
            End If
            If InStr(TXTTarget.Text, ".%") > 0 Or InStr(TXTTarget.Text, "%.") > 0 Then
                Me.TXTTarget.Focus()
                Exit Sub
            End If

            If Me.TXTTarget.Text = "" Then
                Me.TXTTarget.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTTarget.Text, "%") = 0 Then
                vCheckTarget = Me.TXTTarget.Text
                If vCheckTarget = 0 Then
                    vCheckTargetRes = Me.TXTPointRes.Text
                    Me.TXTTargetRes.Text = Format(vCheckTargetRes, "##,##0.00")

                End If

                Me.TXTTarget.Text = Format(vCheckTarget, "##,##0.00")
            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTTargetCalc()
        Dim vTargetWord As String
        Dim vMarketCost As Double
        Dim vCheckCountDot As Integer
        Dim vPoint As Double

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTPointRes.Text <> "" And Me.TXTServiceRes.Text <> "" Then
            If Me.TXTTarget.Text <> "" Then
                vTargetWord = Me.TXTTarget.Text
                vMarketCost = Me.TXTServiceRes.Text
                vPoint = Me.TXTPointRes.Text

                vCheckCountDot = CheckDot(TXTTarget.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTTarget.Focus()
                    Exit Sub
                End If

                If InStr(vTargetWord, "%") > 0 Then
                    If Len(vTargetWord) > 1 Then
                        If InStr(vTargetWord, "0") = 1 Then
                            Me.TXTTarget.Text = Microsoft.VisualBasic.Right(vTargetWord, Len(vTargetWord) - InStr(vTargetWord, "0"))
                            vTargetWord = Me.TXTTarget.Text
                        End If
                    End If
                    If InStr(vTargetWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTTarget.Focus()
                        Exit Sub
                    End If

                    If InStr(vTargetWord, ".%") > 0 Or InStr(vTargetWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTTarget.Focus()
                        Exit Sub
                    End If

                    vTargetWord = Microsoft.VisualBasic.Left(vTargetWord, InStr(vTargetWord, "%"))
                    Me.TXTTarget.Text = vTargetWord
                    vTarget = Microsoft.VisualBasic.Left(vTargetWord, InStr(vTargetWord, "%") - 1)
                    vTarget1 = ((vMarketCost * vTarget) / 100)
                    vTargetAmount = vPoint + vTarget1
                Else
                    If Len(vTargetWord) > 0 Then
                        If InStr(vTargetWord, ".") = 1 And Len(vTargetWord) = 1 Then
                            Me.TXTTarget.Focus()
                            Exit Sub
                        End If

                        If InStr(vTargetWord, "0") = 1 And InStr(vTargetWord, ".") = 2 And Len(vTargetWord) = 2 Then
                            Me.TXTTarget.Focus()
                            Exit Sub
                        End If

                        If (InStr(vTargetWord, ".") = 1) And Len(vTargetWord) = 1 Then
                            Me.TXTTarget.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vTargetWord, "0") = 1 And Len(vTargetWord) > 1 Then
                            Me.TXTTarget.Text = Microsoft.VisualBasic.Right(vTargetWord, Len(vTargetWord) - InStr(vTargetWord, "0"))
                        Else
                            Me.TXTTarget.Text = vTargetWord
                        End If
                    End If
                    vTarget = vTargetWord
                    vTarget1 = vTarget
                    vTargetAmount = vPoint + vTarget1
                End If

                If vTargetAmount < 0 Then
                    Me.TXTTargetRes.ForeColor = Color.Red
                Else
                    Me.TXTTargetRes.ForeColor = Color.Black
                End If

                Me.TXTTargetRes.Text = Format(vTargetAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Target")
        End If
    End Sub

    Private Sub TXTGift_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTGift.LostFocus
        Dim vCheckGift As Double
        Dim vCheckGiftRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTTargetRes.Text <> "" Then

            vCheckCountDot = CheckDot(Me.TXTGift.Text)
            If vCheckCountDot > 1 Then
                Me.TXTGift.Focus()
                Exit Sub
            End If

            If InStr(TXTGift.Text, "0") = 1 And InStr(TXTGift.Text, ".") = 2 And Len(TXTGift.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTGift.Text, ".") = 1 And Len(TXTGift.Text) = 1 Then
                Me.TXTGift.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTGift.Text, "%") = 1 Then
                Me.TXTGift.Focus()
            End If

            If InStr(TXTGift.Text, ".%") > 0 Or InStr(TXTGift.Text, "%.") > 0 Then
                Me.TXTGift.Focus()
                Exit Sub
            End If

            If Me.TXTGift.Text = "" Then
                Me.TXTGift.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTGift.Text, "%") = 0 Then
                vCheckGift = TXTGift.Text
                If vCheckGift = 0 Then
                    vCheckGiftRes = Me.TXTTargetRes.Text
                    Me.TXTBillDiscRes.Text = Format(vCheckGiftRes, "##,##0.00")

                End If
                Me.TXTGift.Text = Format(vCheckGift, "##,##0.00")
            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTGiftCalc()
        Dim vGiftWord As String
        Dim vTarget As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTTargetRes.Text <> "" Then
            If Me.TXTGift.Text <> "" Then
                vGiftWord = Me.TXTGift.Text
                vTarget = Me.TXTTargetRes.Text

                vCheckCountDot = CheckDot(TXTGift.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTGift.Focus()
                    Exit Sub
                End If

                If InStr(vGiftWord, "%") > 0 Then
                    If Len(vGiftWord) > 1 Then
                        If InStr(vGiftWord, "0") = 1 Then
                            Me.TXTGift.Text = Microsoft.VisualBasic.Right(vGiftWord, Len(vGiftWord) - InStr(vGiftWord, "0"))
                            vGiftWord = Me.TXTGift.Text
                        End If
                    End If
                    If InStr(vGiftWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTGift.Focus()
                        Exit Sub
                    End If

                    If InStr(vGiftWord, ".%") > 0 Or InStr(vGiftWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTGift.Focus()
                        Exit Sub
                    End If

                    vGiftWord = Microsoft.VisualBasic.Left(vGiftWord, InStr(vGiftWord, "%"))
                    Me.TXTGift.Text = vGiftWord
                    vGift = Microsoft.VisualBasic.Left(vGiftWord, InStr(vGiftWord, "%") - 1)
                    vGift1 = ((vTarget * vGift) / 100)
                    vGiftAmount = vTarget + vGift1
                Else
                    If Len(vGiftWord) > 0 Then
                        If InStr(vGiftWord, ".") = 1 And Len(vGiftWord) = 1 Then
                            Me.TXTGift.Focus()
                            Exit Sub
                        End If

                        If InStr(vGiftWord, "0") = 1 And InStr(vGiftWord, ".") = 2 And Len(vGiftWord) = 2 Then
                            Me.TXTGift.Focus()
                            Exit Sub
                        End If

                        If (InStr(vGiftWord, ".") = 1) And Len(vGiftWord) = 1 Then
                            Me.TXTGift.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vGiftWord, "0") = 1 And Len(vGiftWord) > 1 Then
                            Me.TXTGift.Text = Microsoft.VisualBasic.Right(vGiftWord, Len(vGiftWord) - InStr(vGiftWord, "0"))
                        Else
                            Me.TXTGift.Text = vGiftWord
                        End If
                    End If
                    vGift = vGiftWord
                    vGift1 = vGift
                    vGiftAmount = vTarget + vGift1
                End If

                If vGiftAmount < 0 Then
                    Me.TXTGiftRes.ForeColor = Color.Red
                Else
                    Me.TXTGiftRes.ForeColor = Color.Black
                End If

                Me.TXTGiftRes.Text = Format(vGiftAmount, "##,##0.00")

            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Gift")
        End If
    End Sub

    Private Sub TXTCommission_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTCommission.LostFocus
        Dim vCheckCommission As Double
        Dim vCheckCommissionRes As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTGiftRes.Text <> "" Then

            vCheckCountDot = CheckDot(TXTCommission.Text)
            If vCheckCountDot > 1 Then
                Me.TXTCommission.Focus()
                Exit Sub
            End If

            If InStr(TXTCommission.Text, "0") = 1 And InStr(TXTCommission.Text, ".") = 2 And Len(TXTCommission.Text) = 2 Then
                Exit Sub
            End If
            If InStr(TXTCommission.Text, ".") = 1 And Len(TXTCommission.Text) = 1 Then
                Me.TXTCommission.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTCommission.Text, "%") = 1 Then
                Me.TXTCommission.Focus()
            End If

            If InStr(TXTCommission.Text, ".%") > 0 Or InStr(TXTCommission.Text, "%.") > 0 Then
                Me.TXTCommission.Focus()
                Exit Sub
            End If

            If Me.TXTCommission.Text = "" Then
                Me.TXTCommission.Text = Format(0, "##,##0.00")
                Exit Sub
            End If

            If InStr(TXTCommission.Text, "%") = 0 Then
                vCheckCommission = TXTCommission.Text
                If vCheckCommission = 0 Then
                    vCheckCommissionRes = Me.TXTGiftRes.Text
                    Me.TXTCommissionRes.Text = Format(vCheckCommissionRes, "##,##0.00")

                End If
                Me.TXTCommission.Text = Format(vCheckCommission, "##,##0.00")
            End If

            Call TXTBegProfitCalc()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Public Sub TXTBegProfitCalc()

        On Error GoTo ErrDescription

        If Me.TXTPriceSet.Text <> "" And InStr(Me.TXTPriceSet.Text, ".,") = 0 And InStr(Me.TXTPriceSet.Text, ",.") = 0 And Me.TXTServiceRes.Text <> "" Then
            vPriceSetAmount = Me.TXTPriceSet.Text
            vServiceAmount = Me.TXTServiceRes.Text
            vBegProfit1 = ((vPriceSetAmount - vCommission1 - vGift1 - vTarget1 - vPoint1 - vDiscMember1 - vServiceAmount) * 100) / vServiceAmount
            vBegProfit2 = (vPriceSetAmount - vCommission1 - vGift1 - vTarget1 - vPoint1 - vDiscMember1 - vServiceAmount)
            Me.TXTBegProfit.Text = Format(vBegProfit1, "##,##0.00")
            If vBegProfit2 < 0 Then
                Me.TXTBegProfitAmount.ForeColor = Color.Red
            Else
                Me.TXTBegProfitAmount.ForeColor = Color.Black
            End If
            Me.TXTBegProfitAmount.Text = Format(vBegProfit2, "##,##0.00")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message BegProfit")
        End If
    End Sub
    Public Sub TXTInterestCalc()
        Dim vInterest1 As Double
        Dim vInterest2 As Double

        On Error GoTo ErrDescription

        If vBegProfit1 < 1 Then
            vInterest1 = 0
        ElseIf vBegProfit1 >= 1 And vBegProfit1 < 5 Then
            vInterest1 = Format((vBegProfit1 / 2), "##,##0.00")
        ElseIf vBegProfit1 >= 5 And vBegProfit1 < 6 Then
            vInterest1 = Format((vBegProfit1 * 0.6), "##,##0.00")
        Else
            vInterest1 = Format(3.5, "##,##0.00")
        End If


        vInterest2 = Format(((vInterest1 * vServiceAmount) / 100), "##,##0.00")

        Me.TXTInterests.Text = Format(vInterest1, "##,##0.00")

        If vInterest2 < 0 Then
            Me.TXTInterestsAmount.ForeColor = Color.Red
        Else
            Me.TXTInterestsAmount.ForeColor = Color.Black
        End If

        Me.TXTInterestsAmount.Text = Format(vInterest2, "##,##0.00")

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Interest")
        End If
    End Sub

    Public Sub TXTProfitCalc()
        Dim vProfit1 As Double
        Dim vProfit2 As Double
        Dim vInterestAmount1 As Double

        On Error GoTo ErrDescription

        If Me.TXTPriceSet.Text <> "" And InStr(Me.TXTPriceSet.Text, ".,") = 0 And InStr(Me.TXTPriceSet.Text, ",.") = 0 Then
            vPriceSetAmount = Me.TXTPriceSet.Text
            vInterestAmount1 = Me.TXTInterestsAmount.Text
            vProfit2 = vPriceSetAmount - vCommission1 - vGift1 - vTarget1 - vPoint1 - vDiscMember1 - vServiceAmount - vInterestAmount1
            vProfit1 = vProfit2 / vServiceAmount
            Me.TXTProfit.Text = Format(vProfit1, "##,##0.00")

            If vProfit2 < 0 Then
                Me.TXTProfitAmount.ForeColor = Color.Red
            Else
                Me.TXTProfitAmount.ForeColor = Color.Black
            End If

            Me.TXTProfitAmount.Text = Format(vProfit2, "##,##0.00")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Profit")
        End If
    End Sub

    Private Sub TXTCommissionCalc()
        Dim vCommissionWord As String
        Dim vGift As Double
        Dim vCheckCountDot As Integer

        On Error GoTo ErrDescription

        If Me.TXTDO.Text <> "" And Me.TXTGiftRes.Text <> "" Then
            If Me.TXTCommission.Text <> "" And Me.TXTCommission.Text <> "." Then
                vCommissionWord = Me.TXTCommission.Text
                vGift = Me.TXTGiftRes.Text

                vCheckCountDot = CheckDot(TXTCommission.Text)
                If vCheckCountDot > 1 Then
                    MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                    Me.TXTCommission.Focus()
                    Exit Sub
                End If

                If InStr(vCommissionWord, "%") > 0 Then
                    If Len(vCommissionWord) > 1 Then
                        If InStr(vCommissionWord, "0") = 1 Then
                            Me.TXTCommission.Text = Microsoft.VisualBasic.Right(vCommissionWord, Len(vCommissionWord) - InStr(vCommissionWord, "0"))
                            vCommissionWord = Me.TXTCommission.Text
                        End If
                    End If
                    If InStr(vCommissionWord, "%") = 1 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTCommission.Focus()
                        Exit Sub
                    End If

                    If InStr(vCommissionWord, ".%") > 0 Or InStr(vCommissionWord, "%.") > 0 Then
                        MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
                        Me.TXTCommission.Focus()
                        Exit Sub
                    End If

                    vCommissionWord = Microsoft.VisualBasic.Left(vCommissionWord, InStr(vCommissionWord, "%"))
                    Me.TXTCommission.Text = vCommissionWord
                    vCommission = Microsoft.VisualBasic.Left(vCommissionWord, InStr(vCommissionWord, "%") - 1)
                    vCommission1 = ((vGift * vCommission) / 100)
                    vCommissionAmount = vGift + vCommission1
                Else
                    If Len(vCommissionWord) > 0 Then
                        If InStr(vCommissionWord, ".") = 1 And Len(vCommissionWord) = 1 Then
                            Me.TXTCommission.Focus()
                            Exit Sub
                        End If

                        If InStr(vCommissionWord, "0") = 1 And InStr(vCommissionWord, ".") = 2 And Len(vCommissionWord) = 2 Then
                            Me.TXTCommission.Focus()
                            Exit Sub
                        End If

                        If (InStr(vCommissionWord, ".") = 1) And Len(vCommissionWord) = 1 Then
                            Me.TXTCommission.Text = Format(0, "##,##0.00")
                        ElseIf InStr(vCommissionWord, "0") = 1 And Len(vCommissionWord) > 1 Then
                            Me.TXTCommission.Text = Microsoft.VisualBasic.Right(vCommissionWord, Len(vCommissionWord) - InStr(vCommissionWord, "0"))
                        Else
                            Me.TXTCommission.Text = vCommissionWord
                        End If
                    End If
                    vCommission = vCommissionWord
                    vCommission1 = vCommission
                    vCommissionAmount = vGift + vCommission1
                End If
                If vCommissionAmount < 0 Then
                    Me.TXTCommissionRes.ForeColor = Color.Red
                Else
                    Me.TXTCommissionRes.ForeColor = Color.Black
                End If
                Me.TXTCommissionRes.Text = Format(vCommissionAmount, "##,##0.00")
            Else
                Me.TXTCommissionRes.Text = Format(0, "##,##0.00")
            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Commission")
        End If
    End Sub

    Private Sub TXTBegProfit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTBegProfit.TextChanged
        Dim vCheckBegProfit As Double

        On Error GoTo ErrDescription

        If Me.TXTBegProfit.Text <> "" Then
            vCheckBegProfit = Me.TXTBegProfit.Text
            If vCheckBegProfit < 0 Then
                Me.TXTBegProfit.ForeColor = Color.Red
            Else
                Me.TXTBegProfit.ForeColor = Color.Black
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTFollowDisc3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTFollowDisc3.TextChanged
        Call TXTFollowDisc3Calc()
    End Sub

    Private Sub TXTRebate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTRebate.TextChanged
        Call TXTRebateCalc()
    End Sub

    Private Sub TXTSpecialDisc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTSpecialDisc.TextChanged
        Call TXTSpecialDiscCalc()
    End Sub

    Private Sub TXTMissProfit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTMissProfit.TextChanged
        Call TXTMissProfitCalc()
    End Sub

    Private Sub TXTSend_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTSend.TextChanged
        Call TXTSendCalc()
    End Sub

    Private Sub TXTCustSend_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTCustSend.TextChanged
        Call TXTCustSendCalc()
    End Sub

    Private Sub TXTAdvertise_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTAdvertise.TextChanged
        Call TXTAdvertiseCalc()
    End Sub

    Private Sub TXTMarket_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTMarket.TextChanged
        Call TXTMarketCalc()
    End Sub

    Private Sub TXTTax_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTTax.TextChanged
        Call TXTTaxCalc()
    End Sub

    Private Sub TXTInstall_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTInstall.TextChanged
        Call TXTInstallCalc()
    End Sub

    Private Sub TXTService_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTService.TextChanged
        Call TXTServiceCalc()
    End Sub

    Private Sub TXTPoint_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTPoint.TextChanged
        Call TXTPointCalc()
    End Sub

    Private Sub TXTTarget_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTTarget.TextChanged
        Call TXTTargetCalc()
    End Sub

    Private Sub TXTGift_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTGift.TextChanged
        Call TXTGiftCalc()
    End Sub

    Private Sub TXTCommission_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTCommission.TextChanged
        Call TXTCommissionCalc()
    End Sub

    Private Sub TXTDO_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTDO.TextChanged
        'If Me.TXTBillDiscRes.Text <> "" And Me.TXTBillDiscRes.Text <> "0.00" Then
        Call TXTBillDiscCalc()
        'End If     
    End Sub

    Private Sub TXTCommissionRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTCommissionRes.TextChanged
        On Error GoTo ErrDescription

        If Me.TXTPriceSet.Text <> "" And Me.TXTPriceSet.Text <> "0.00" Then
            Call TXTBegProfitCalc()
            Call TXTinterestCalc()
            Call TXTProfitCalc()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTBillDiscRes_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTBillDiscRes.LostFocus
        Call TXTFollowDisc1Calc()
    End Sub

    Private Sub TXTBillDiscRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTBillDiscRes.TextChanged
        'If Me.TXTFollowDisc1Res.Text <> "" And Me.TXTFollowDisc1Res.Text <> "0.00" Then
        Call TXTFollowDisc1Calc()
        'End If
    End Sub

    Private Sub TXTFollowDisc1Res_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTFollowDisc1Res.TextChanged
        'If Me.TXTFollowDisc2Res.Text <> "" And Me.TXTFollowDisc2Res.Text <> "0.00" Then
        Call TXTFollowDisc2Calc()
        'End If
    End Sub

    Private Sub TXTFollowDisc2Res_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTFollowDisc2Res.TextChanged
        'If Me.TXTFollowDisc3Res.Text <> "" And Me.TXTFollowDisc3Res.Text <> "0.00" Then
        Call TXTFollowDisc3Calc()
        'End If
    End Sub

    Private Sub TXTFollowDisc3Res_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTFollowDisc3Res.TextChanged
        'If Me.TXTRebateRes.Text <> "" And Me.TXTRebateRes.Text <> "0.00" Then
        Call TXTRebateCalc()
        'End If
    End Sub

    Private Sub TXTRebateRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTRebateRes.TextChanged
        'If Me.TXTSpecialDiscRes.Text <> "" And Me.TXTSpecialDiscRes.Text <> "0.00" Then
        Call TXTSpecialDiscCalc()
        'End If
    End Sub

    Private Sub TXTSpecialDiscRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTSpecialDiscRes.TextChanged
        'If Me.TXTMissProfitRes.Text <> "" And Me.TXTMissProfitRes.Text <> "0.00" Then
        Call TXTMissProfitCalc()
        'End If
    End Sub

    Private Sub TXTMissProfitRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTMissProfitRes.TextChanged
        'If Me.TXTSendRes.Text <> "" And Me.TXTSendRes.Text <> "0.00" Then
        Call TXTSendCalc()
        'End If
    End Sub

    Private Sub TXTSendRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTSendRes.TextChanged
        'If Me.TXTCustSendRes.Text <> "" And Me.TXTCustSendRes.Text <> "0.00" Then
        Call TXTCustSendCalc()
        'End If
    End Sub

    Private Sub TXTCustSendRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTCustSendRes.TextChanged
        'If Me.TXTAdvertiseRes.Text <> "" And Me.TXTAdvertiseRes.Text <> "0.00" Then
        Call TXTAdvertiseCalc()
        'End If
    End Sub

    Private Sub TXTAdvertiseRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTAdvertiseRes.TextChanged
        'If Me.TXTMarketRes.Text <> "" And Me.TXTMarketRes.Text <> "0.00" Then
        Call TXTMarketCalc()
        'End If
    End Sub

    Private Sub TXTMarketRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTMarketRes.TextChanged
        'If Me.TXTTaxRes.Text <> "" And Me.TXTTaxRes.Text <> "0.00" Then
        Call TXTTaxCalc()
        'End If
    End Sub

    Private Sub TXTTaxRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTTaxRes.TextChanged
        'If Me.TXTInstallRes.Text <> "" And Me.TXTInstallRes.Text <> "0.00" Then
        Call TXTInstallCalc()
        'End If
    End Sub

    Private Sub TXTInstallRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTInstallRes.TextChanged
        'If Me.TXTServiceRes.Text <> "" And Me.TXTServiceRes.Text <> "0.00" Then
        Call TXTServiceCalc()
        'End If
    End Sub

    Private Sub TXTServiceRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTServiceRes.TextChanged
        'If Me.TXTPointRes.Text <> "" And Me.TXTPointRes.Text <> "0.00" Then
        Call TXTCreditCardCalc()
        'End If
    End Sub

    Private Sub TXTPointRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTPointRes.TextChanged
        'If Me.TXTTargetRes.Text <> "" And Me.TXTTargetRes.Text <> "0.00" Then
        Call TXTTargetCalc()
        'End If
    End Sub

    Private Sub TXTTargetRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTTargetRes.TextChanged
        'If Me.TXTGiftRes.Text <> "" And Me.TXTGiftRes.Text <> "0.00" Then
        Call TXTGiftCalc()
        'End If
    End Sub

    Private Sub TXTGiftRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTGiftRes.TextChanged
        'If Me.TXTCommissionRes.Text <> "" And Me.TXTCommissionRes.Text <> "0.00" Then
        Call TXTCommissionCalc()
        'End If
    End Sub

    Private Sub TXTAddCashCalc()
        Dim vPriceSetAmount As Double
        Dim vAddCashWord As String
        Dim vAddCash As Double

        On Error GoTo ErrDescription

        If Me.TXTPriceSet.Text <> "" And Me.TXTAddCash.Text <> "" Then
            vAddCashWord = Me.TXTAddCash.Text
            vPriceSetAmount = Me.TXTPriceSet.Text
            If InStr(vAddCashWord, "%") > 0 Then
                vAddCashWord = Microsoft.VisualBasic.Left(vAddCashWord, InStr(vAddCashWord, "%") - 1)
                vAddCash = (vPriceSetAmount * vAddCashWord) / 100
            Else
                vAddCash = vAddCashWord
            End If
            If vAddCash < 0 Then
                Me.TXTAddCashRes.ForeColor = Color.Red
            Else
                Me.TXTAddCashRes.ForeColor = Color.Black
            End If

            Me.TXTAddCashRes.Text = Format(vAddCash, "##,##0.00")
        Else
            Me.TXTAddCashRes.Text = Format(0, "##,##0.00")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message AddCash")
        End If
    End Sub
    Private Sub TXTAddCreditCalc()
        Dim vPriceSetAmount As Double
        Dim vAddCreditWord As String
        Dim vAddCredit As Double

        On Error GoTo ErrDescription

        If Me.TXTPriceSet.Text <> "" And Me.TXTAddCredit.Text <> "" Then
            vAddCreditWord = Me.TXTAddCredit.Text
            vPriceSetAmount = Me.TXTPriceSet.Text
            If InStr(vAddCreditWord, "%") > 0 Then
                vAddCreditWord = Microsoft.VisualBasic.Left(vAddCreditWord, InStr(vAddCreditWord, "%") - 1)
                vAddCredit = (vPriceSetAmount * vAddCreditWord) / 100
            Else
                vAddCredit = vAddCreditWord
            End If
            If vAddCredit < 0 Then
                Me.TXTAddCreditRes.ForeColor = Color.Red
            Else
                Me.TXTAddCreditRes.ForeColor = Color.Black
            End If
            Me.TXTAddCreditRes.Text = Format(vAddCredit, "##,##0.00")
        Else
            Me.TXTAddCreditRes.Text = Format(0, "##,##0.00")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message AddCredit")
        End If
    End Sub
    Private Sub TXTDiscount1Calc()
        Dim vPriceSetAmount As Double
        Dim vDiscount1Word As String
        Dim vDiscount1 As Double

        On Error GoTo ErrDescription

        If Me.TXTPriceSet.Text <> "" And Me.TXTDiscount1.Text <> "" Then
            vDiscount1Word = Me.TXTDiscount1.Text
            vPriceSetAmount = Me.TXTPriceSet.Text
            If InStr(vDiscount1Word, "%") > 0 Then
                vDiscount1Word = Microsoft.VisualBasic.Left(vDiscount1Word, InStr(vDiscount1Word, "%") - 1)
                vDiscount1 = (vPriceSetAmount * vDiscount1Word) / 100
            Else
                vDiscount1 = vDiscount1Word
            End If

            If vDiscount1 < 0 Then
                Me.TXTDiscount1Res.ForeColor = Color.Red
            Else
                Me.TXTDiscount1Res.ForeColor = Color.Black
            End If

            Me.TXTDiscount1Res.Text = Format(vDiscount1, "##,##0.00")
        Else
            Me.TXTDiscount1Res.Text = Format(0, "##,##0.00")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message Discount1")
        End If
    End Sub

    Private Sub TXTPriceSet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTPriceSet.TextChanged
        On Error GoTo ErrDescription

        If Me.TXTAddCash.Text <> "" And Me.TXTAddCash.Text <> "0.00" And InStr(Me.TXTAddCash.Text, "%") > 0 Then
            Call TXTAddCashCalc()
        End If
        If Me.TXTAddCredit.Text <> "" And Me.TXTAddCredit.Text <> "0.00" And InStr(Me.TXTAddCredit.Text, "%") > 0 Then
            Call TXTAddCreditCalc()
        End If
        If Me.TXTDiscount1.Text <> "" And Me.TXTDiscount1.Text <> "0.00" And InStr(Me.TXTDiscount1.Text, "%") > 0 Then
            Call TXTDiscount1Calc()
        End If

        Call TXTBegProfitCalc()
        Call TXTinterestCalc()
        Call TXTProfitCalc()
        Call TXTDiscMemberCalc()
        Call TXTPointCalc()

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTAddCash_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTAddCash.LostFocus
        Dim vCheckAddCash As Double

        On Error GoTo ErrDescription


        If InStr(Me.TXTAddCash.Text, "%") = 0 Then
            If Me.TXTAddCash.Text = "" Then
                Me.TXTAddCash.Text = Format(0, "##,##0.00")
            Else
                vCheckAddCash = Me.TXTAddCash.Text
                Me.TXTAddCash.Text = Format(vCheckAddCash, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTAddCash_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTAddCash.TextChanged
        Call TXTAddCashCalc()
    End Sub

    Private Sub TXTAddCredit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTAddCredit.LostFocus
        Dim vCheckAddCredit As Double

        On Error GoTo ErrDescription

        If InStr(Me.TXTAddCredit.Text, "%") = 0 Then
            If Me.TXTAddCredit.Text = "" Then
                Me.TXTAddCredit.Text = Format(0, "##,##0.00")
            Else
                vCheckAddCredit = Me.TXTAddCredit.Text
                Me.TXTAddCredit.Text = Format(vCheckAddCredit, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTAddCredit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTAddCredit.TextChanged
        Call TXTAddCreditCalc()
    End Sub

    Private Sub TXTDiscount1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTDiscount1.LostFocus
        Dim vCheckDiscount1 As Double

        On Error GoTo ErrDescription

        If InStr(Me.TXTDiscount1.Text, "%") = 0 Then
            If Me.TXTDiscount1.Text = "" Then
                Me.TXTDiscount1.Text = Format(0, "##,##0.00")
            Else
                vCheckDiscount1 = Me.TXTDiscount1.Text
                Me.TXTDiscount1.Text = Format(vCheckDiscount1, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTDiscount1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTDiscount1.TextChanged
        Call TXTDiscount1Calc()
    End Sub

    Private Sub BTNReCommend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNReCommend.Click
        Dim vItemCode As String
        Dim vUnitCode As String
        Dim vPriceSet As Double
        Dim vSendItemAmount As Double
        Dim vCreditAmount As Double
        Dim vDiscSpecialAmount As Double
        Dim i As Integer
        Dim vListPrice As ListViewItem
        Dim vCount As Integer
        Dim vDefSaleUnitCode As String

        Dim vUnit As String
        Dim vSaleType As Integer
        Dim vTransferType As Integer
        Dim vCheckItemCode As String
        Dim vCheckUnit As String
        Dim vCheckSaleType As Integer
        Dim vCheckTransferType As Integer

        On Error GoTo ErrDescription

        If Me.TXTItemCode.Text <> "" And Me.TXTPriceSet.Text <> "" And Me.CMBUnit.Text <> "" And Me.TXTAddCashRes.Text <> "" And Me.TXTAddCreditRes.Text <> "" And Me.TXTDiscount1Res.Text <> "" And Me.TXTPriceSet.Text <> "0.00" And Me.TXTPriceSet.Text <> "0" Then

            vItemCode = Trim(Me.TXTItemCode.Text)
            vUnitCode = Trim(Me.CMBUnit.Text)
            vDefSaleUnitCode = Trim(Me.CMBMultiUnitCode.Text)
            vPriceSet = Me.TXTPriceSet.Text
            vSendItemAmount = Me.TXTAddCashRes.Text
            vCreditAmount = Me.TXTAddCreditRes.Text
            vDiscSpecialAmount = Me.TXTDiscount1Res.Text
            vQuery = " exec dbo.usp_ps_pricerecommend '" & vItemCode & "','" & vUnitCode & "','" & vDefSaleUnitCode & "'," & vPriceSet & "," & vSendItemAmount & "," & vCreditAmount & "," & vDiscSpecialAmount & " "
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "PriceRecCommend")
            dt = ds.Tables("PriceRecCommend")
            If Me.ListView101.Items.Count > 0 Then
                vCheckUnitCode = Me.ListView101.Items(0).SubItems(0).Text
                vCheckItem = Me.ListView101.Items(0).SubItems(13).Text
            End If
            'If vCheckItem = vItemCode And vUnitCode = vCheckUnitCode Then
            '    'If MsgBox("มีระดับราคาของหน่วยนับ " & vUnitCode & " นี้อยู่แล้ว   ต้องการเคลียร์ข้อมูลก่อนหน้านี้หรือไม่", MsgBoxStyle.YesNo, "Send Question ?") = MsgBoxResult.Yes Then
            '    Me.ListView101.Items.Clear()
            '    'End If
            'End If
            If dt.Rows.Count > 0 Then

                If Me.ListView101.Items.Count = 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        vListPrice = Me.ListView101.Items.Add(dt.Rows(i).Item("defsaleunitcode"))
                        If dt.Rows(i).Item("saletype") = 0 Then
                            vListPrice.SubItems.Add(1).Text = "ขายเงินสด"
                        ElseIf dt.Rows(i).Item("saletype") = 1 Then
                            vListPrice.SubItems.Add(1).Text = "ขายเงินเชื่อ"
                        End If
                        If dt.Rows(i).Item("transporttype") = 0 Then
                            vListPrice.SubItems.Add(2).Text = "รับเอง"
                        ElseIf dt.Rows(i).Item("transporttype") = 1 Then
                            vListPrice.SubItems.Add(2).Text = "ส่งให้"
                        End If
                        vListPrice.SubItems.Add(3).Text = Format(dt.Rows(i).Item("fromqty"), "##,##0.000")
                        vListPrice.SubItems.Add(4).Text = Format(dt.Rows(i).Item("toqty"), "##,##0.000")
                        vListPrice.SubItems.Add(5).Text = Format(dt.Rows(i).Item("priceset1"), "##,##0.00")
                        vListPrice.SubItems.Add(6).Text = Format(dt.Rows(i).Item("priceset2"), "##,##0.00")
                        vListPrice.SubItems.Add(7).Text = "0"
                        vListPrice.SubItems.Add(8).Text = ""
                        vListPrice.SubItems.Add(9).Text = Format(DateAdd(DateInterval.Day, 3650, Now), "dd/MM/yyyy")
                        vListPrice.SubItems.Add(10).Text = ""
                        vListPrice.SubItems.Add(11).Text = ""
                        vListPrice.SubItems.Add(12).Text = ""
                        vListPrice.SubItems.Add(13).Text = dt.Rows(i).Item("itemcode")
                    Next
                Else
                    For i = 0 To dt.Rows.Count - 1


                        vUnit = dt.Rows(i).Item("defsaleunitcode")


                        vSaleType = dt.Rows(i).Item("saletype")

                        vTransferType = dt.Rows(i).Item("transporttype")


                        For vCount = 0 To Me.ListView101.Items.Count - 1
                            vCheckItemCode = Me.ListView101.Items(vCount).SubItems(13).Text
                            vCheckUnit = Me.ListView101.Items(vCount).SubItems(0).Text
                            If Me.ListView101.Items(vCount).SubItems(1).Text = "ขายเงินสด" Then
                                vCheckSaleType = 0
                            Else
                                vCheckSaleType = 1
                            End If
                            If Me.ListView101.Items(vCount).SubItems(2).Text = "รับเอง" Then
                                vCheckTransferType = 0
                            Else
                                vCheckTransferType = 1
                            End If

                            If vItemCode = vCheckItemCode Then
                                If vUnit = vCheckUnit Then
                                    If vSaleType = vCheckSaleType And vTransferType = vCheckTransferType Then
                                        Me.ListView101.Items(vCount).SubItems(5).Text = Format(dt.Rows(i).Item("priceset1"), "##,##0.00")
                                        Me.ListView101.Items(vCount).SubItems(6).Text = Format(dt.Rows(i).Item("priceset2"), "##,##0.00")
                                        GoTo Line1
                                    End If

                                End If
                            End If

                        Next

                        vListPrice = Me.ListView101.Items.Add(dt.Rows(i).Item("defsaleunitcode"))
                        If dt.Rows(i).Item("saletype") = 0 Then
                            vListPrice.SubItems.Add(1).Text = "ขายเงินสด"
                        ElseIf dt.Rows(i).Item("saletype") = 1 Then
                            vListPrice.SubItems.Add(1).Text = "ขายเงินเชื่อ"
                        End If
                        If dt.Rows(i).Item("transporttype") = 0 Then
                            vListPrice.SubItems.Add(2).Text = "รับเอง"
                        ElseIf dt.Rows(i).Item("transporttype") = 1 Then
                            vListPrice.SubItems.Add(2).Text = "ส่งให้"
                        End If
                        vListPrice.SubItems.Add(3).Text = Format(dt.Rows(i).Item("fromqty"), "##,##0.000")
                        vListPrice.SubItems.Add(4).Text = Format(dt.Rows(i).Item("toqty"), "##,##0.000")
                        vListPrice.SubItems.Add(5).Text = Format(dt.Rows(i).Item("priceset1"), "##,##0.00")
                        vListPrice.SubItems.Add(6).Text = Format(dt.Rows(i).Item("priceset2"), "##,##0.00")
                        vListPrice.SubItems.Add(7).Text = "0"
                        vListPrice.SubItems.Add(8).Text = ""
                        vListPrice.SubItems.Add(9).Text = Format(DateAdd(DateInterval.Day, 3650, Now), "dd/MM/yyyy")
                        vListPrice.SubItems.Add(10).Text = ""
                        vListPrice.SubItems.Add(11).Text = ""
                        vListPrice.SubItems.Add(12).Text = ""
                        vListPrice.SubItems.Add(13).Text = dt.Rows(i).Item("itemcode")
Line1:
                    Next
                End If
            End If

            Me.ListView101.Focus()
        Else
            MsgBox("ไม่สามารถกำหนดราคาได้ เนื่องจากข้อมูลไม่ครบตามที่ระบุไว้", MsgBoxStyle.Critical, "Send Error Message")
            Me.TXTPriceSet.Focus()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub ListView101_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView101.DoubleClick

        On Error GoTo ErrDescription

        If Me.ListView101.Items.Count > 0 Then
            vSelectPriceListIndex = Me.ListView101.SelectedItems(0).Index
            Me.NMFromNumber.Value = Me.ListView101.Items(vSelectPriceListIndex).SubItems(3).Text
            Me.NMToNumber.Value = Me.ListView101.Items(vSelectPriceListIndex).SubItems(4).Text
            Me.TXTPrice1.Text = Format(Int(Me.ListView101.Items(vSelectPriceListIndex).SubItems(5).Text), "##,##0.00")
            Me.TXTPrice2.Text = Format(Int(Me.ListView101.Items(vSelectPriceListIndex).SubItems(6).Text), "##,##0.00")
            If Me.ListView101.Items(vSelectPriceListIndex).SubItems(8).Text <> "" Then
                Me.DateUpdate.Text = Me.ListView101.Items(vSelectPriceListIndex).SubItems(8).Text
            Else
                Me.DateUpdate.Text = DateAdd(DateInterval.Day, 1, Date.Now)
            End If
            If Me.ListView101.Items(vSelectPriceListIndex).SubItems(9).Text <> "" Then
                Me.DateExpire.Text = Me.ListView101.Items(vSelectPriceListIndex).SubItems(9).Text
            Else
                Me.DateExpire.Text = DateAdd(DateInterval.Day, 3650, Now)
            End If
            If Me.ListView101.Items(vSelectPriceListIndex).SubItems(7).Text = "1" Then
                Me.CBUpdatePrice.Checked = True
            Else
                Me.CBUpdatePrice.Checked = False
            End If
            Me.GBPriceList.Visible = True
            Me.GBPriceList.BringToFront()
        End If


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub
    Public Sub ClearPriceList()
        On Error Resume Next

        Me.NMFromNumber.Value = 0
        Me.NMToNumber.Value = 0
        Me.TXTPrice1.Text = Format(0, "##,##0.00")
        Me.TXTPrice2.Text = Format(0, "##,##0.00")
        Me.DateUpdate.Text = DateAdd(DateInterval.Day, 1, Date.Now)
        Me.DateExpire.Text = Now
        Me.CBUpdatePrice.Checked = False
    End Sub

    Private Sub BTNBasket_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNBasket.Click
        Dim vDateUpDate As String

        On Error GoTo ErrDescription

        vDateUpDate = Me.DateUpdate.Text
        If DateDiff(DateInterval.Day, Date.Now, CDate(vDateUpDate)) < 0 Then
            MsgBox("วันที่ปรับราคาต้องไม่น้อยกว่าวันที่ปัจจุบัน")
            Exit Sub
        End If
        Me.ListView101.Items(vSelectPriceListIndex).SubItems(3).Text = Format(Me.NMFromNumber.Value, "##,##0.000")
        Me.ListView101.Items(vSelectPriceListIndex).SubItems(4).Text = Format(Me.NMToNumber.Value, "##,##0.000")
        Me.ListView101.Items(vSelectPriceListIndex).SubItems(5).Text = Format(Int(Me.TXTPrice1.Text), "##,##0.00")
        Me.ListView101.Items(vSelectPriceListIndex).SubItems(6).Text = Format(Int(Me.TXTPrice2.Text), "##,##0.00")
        If Me.CBUpdatePrice.Checked = True Then
            Me.ListView101.Items(vSelectPriceListIndex).SubItems(7).Text = 1
        Else
            Me.ListView101.Items(vSelectPriceListIndex).SubItems(7).Text = 0
        End If
        If Me.CBUpdatePrice.Checked = True Then
            Me.ListView101.Items(vSelectPriceListIndex).SubItems(8).Text = Me.DateUpdate.Text
        Else
            Me.ListView101.Items(vSelectPriceListIndex).SubItems(8).Text = ""
        End If
        Me.ListView101.Items(vSelectPriceListIndex).SubItems(9).Text = Me.DateExpire.Text
        Me.GBPriceList.Visible = False

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vDocno As String
        Dim vDocdate As String
        Dim vItemcode As String
        Dim vItemName As String
        Dim vBuyUnitCode As String
        Dim vBuyer As String
        Dim vDO As Double
        Dim vPriceSet As Double

        Dim vBillDiscWord As String
        Dim vBillDiscAmount1 As Double
        Dim vBillDiscRes As Double

        Dim vFollowDisc1Word As String
        Dim vFollowDisc1Amount1 As Double
        Dim vFollowDisc1Res As Double

        Dim vFollowDisc2Word As String
        Dim vFollowDisc2Amount1 As Double
        Dim vFollowDisc2Res As Double

        Dim vFollowDisc3Word As String
        Dim vFollowDisc3Amount1 As Double
        Dim vFollowDisc3Res As Double

        Dim vRebateWord As String
        Dim vRebateAmount1 As Double
        Dim vRebateRes As Double

        Dim vSpecialDiscWord As String
        Dim vSpecialDiscAmount1 As Double
        Dim vSpecialDiscRes As Double

        Dim vMissProfitWord As String
        Dim vMissProfitAmount1 As Double
        Dim vMissProfitRes As Double

        Dim vSendWord As String
        Dim vSendAmount1 As Double
        Dim vSendRes As Double

        Dim vCustSendWord As String
        Dim vCustSendAmount1 As Double
        Dim vCustSendRes As Double

        Dim vAdvertiseWord As String
        Dim vAdvertiseAmount1 As Double
        Dim vAdvertiseRes As Double

        Dim vMarketWord As String
        Dim vMarketAmount1 As Double
        Dim vMarketRes As Double

        Dim vTaxWord As String
        Dim vTaxAmount1 As Double
        Dim vTaxRes As Double

        Dim vInstallWord As String
        Dim vInstallAmount1 As Double
        Dim vInstallRes As Double

        Dim vServiceWord As String
        Dim vServiceAmount1 As Double
        Dim vServiceRes As Double

        Dim vDiscMemberWord As String
        Dim vDiscMemberAmount1 As Double
        Dim vDiscMemberRes As Double

        Dim vPointWord As String
        Dim vPointAmount1 As Double
        Dim vPointRes As Double

        Dim vTargetWord As String
        Dim vTargetAmount1 As Double
        Dim vTargetRes As Double

        Dim vGiftWord As String
        Dim vGiftAmount1 As Double
        Dim vGiftRes As Double

        Dim vCommissionWord As String
        Dim vCommissionAmount1 As Double
        Dim vCommissionRes As Double

        Dim vBegProfitPercent As Double
        Dim vBegProfitAmount1 As Double
        Dim vInterestPercent As Double
        Dim vInterestAmount1 As Double
        Dim vProfitPercent As Double
        Dim vProfitAmount1 As Double
        Dim vAddCashWord As String
        Dim vAddCashAmount1 As Double
        Dim vAddCreditWord As String
        Dim vAddCreditAmount1 As Double
        Dim vSpecialDisc1Word As String
        Dim vSpecialDisc1Amount1 As String
        Dim vFromQTY As Double
        Dim vToQTY As Double
        Dim vSaleType As Integer
        Dim vTransSportType As Integer
        Dim vPriceLevel1 As Double
        Dim vPriceLevel2 As Double
        Dim vLineNumber As Integer
        Dim vIsPriceUpdate As Integer
        Dim vDatePriceUpdate As String
        Dim vDateExpire As String
        Dim i As Integer
        Dim vItemUnitCode As String

        If Me.TextBoxDocno.Text <> "" And Me.TXTItemCode.Text <> "" And Me.ListView101.Items.Count > 0 And Me.CMBUnit.Text <> "" And Me.TXTPriceSet.Text <> "" Then
            If Me.ListView103.Items.Count = 0 Then
                Call GetDocNo()
            End If
            vDocno = Me.TextBoxDocno.Text
            If vIsCancel = 0 Then
                If vIsConfirm = 0 Then
                    vCheckBuyer = Me.TextBoxBuyer.Text
                    If Me.TextBoxDocno.Text <> "" Then
                        If vCheckUserID <> vCheckBuyer Then
                            MsgBox("ผู้เข้ามาใช้งานโปรแกรมกับผู้สร้างเอกสารไม่ใช่คนเดียวกัน ไม่สามารถเปลี่ยนแปลงข้อมูลของเอกสารนี้ได้", MsgBoxStyle.Critical, "Send Error Message")
                            Exit Sub
                        End If
                        Me.Cursor = Cursors.WaitCursor
                        Try
                            vQuery = "begin tran"
                            vCMD = New SqlCommand(vQuery, vConnection)
                            vCMD.ExecuteNonQuery()

                            vDocno = Me.TextBoxDocno.Text
                            vDocdate = Me.DocDate.Text

                            vQuery = "exec dbo.USP_PS_PriceStructureSet '" & vDocno & "','" & vDocdate & "' "
                            vCMD = New SqlCommand(vQuery, vConnection)
                            vCMD.ExecuteNonQuery()


                            vItemcode = Me.TXTItemCode.Text
                            vItemName = Me.LBLItemName.Text
                            vBuyUnitCode = Me.CMBUnit.Text
                            vBuyer = Me.TextBoxBuyer.Text
                            vDO = Me.TXTDO.Text
                            vPriceSet = Me.TXTPriceSet.Text

                            vBillDiscWord = Me.TXTBillDisc.Text
                            vBillDiscAmount1 = vBillDisc1
                            vBillDiscRes = Me.TXTBillDiscRes.Text

                            vFollowDisc1Word = Me.TXTFollowDisc1.Text
                            vFollowDisc1Amount1 = vDiscFollow11
                            vFollowDisc1Res = Me.TXTFollowDisc1Res.Text

                            vFollowDisc2Word = Me.TXTFollowDisc2.Text
                            vFollowDisc2Amount1 = vDiscFollow21
                            vFollowDisc2Res = Me.TXTFollowDisc2Res.Text

                            vFollowDisc3Word = Me.TXTFollowDisc3.Text
                            vFollowDisc3Amount1 = vDiscFollow31
                            vFollowDisc3Res = Me.TXTFollowDisc3Res.Text

                            vRebateWord = Me.TXTRebate.Text
                            vRebateAmount1 = vDiscRebate1
                            vRebateRes = Me.TXTRebateRes.Text

                            vSpecialDiscWord = Me.TXTSpecialDisc.Text
                            vSpecialDiscAmount1 = vDiscSpecial1
                            vSpecialDiscRes = Me.TXTSpecialDiscRes.Text

                            vMissProfitWord = Me.TXTMissProfit.Text
                            vMissProfitAmount1 = vDiscMissProfit1
                            vMissProfitRes = Me.TXTMissProfitRes.Text

                            vSendWord = Me.TXTSend.Text
                            vSendAmount1 = vSend1
                            vSendRes = Me.TXTSendRes.Text

                            vCustSendWord = Me.TXTCustSend.Text
                            vCustSendAmount1 = vCustSend1
                            vCustSendRes = Me.TXTCustSendRes.Text

                            vAdvertiseWord = Me.TXTAdvertise.Text
                            vAdvertiseAmount1 = vAdvertise1
                            vAdvertiseRes = Me.TXTAdvertiseRes.Text

                            vMarketWord = Me.TXTMarket.Text
                            vMarketAmount1 = vMarket1
                            vMarketRes = Me.TXTMarketRes.Text

                            vTaxWord = Me.TXTTax.Text
                            vTaxAmount1 = vTax1
                            vTaxRes = Me.TXTTaxRes.Text

                            vInstallWord = Me.TXTInstall.Text
                            vInstallAmount1 = vInstall1
                            vInstallRes = Me.TXTInstallRes.Text

                            vServiceWord = Me.TXTService.Text
                            vServiceAmount1 = vService1
                            vServiceRes = Me.TXTServiceRes.Text

                            vDiscMemberWord = Me.TXTDiscMember.Text
                            vDiscMemberAmount1 = vDiscMember1
                            vDiscMemberRes = Me.TXTDiscMemberRes.Text

                            vPointWord = Me.TXTPoint.Text
                            vPointAmount1 = vPoint1
                            vPointRes = Me.TXTPointRes.Text

                            vTargetWord = Me.TXTTarget.Text
                            vTargetAmount1 = vTarget1
                            vTargetRes = Me.TXTTargetRes.Text

                            vGiftWord = Me.TXTGift.Text
                            vGiftAmount1 = vGift1
                            vGiftRes = Me.TXTGiftRes.Text

                            vCommissionWord = Me.TXTCommission.Text
                            vCommissionAmount1 = vCommission1
                            vCommissionRes = Me.TXTCommissionRes.Text

                            vBegProfitPercent = Me.TXTBegProfit.Text
                            vBegProfitAmount1 = Me.TXTBegProfitAmount.Text
                            vInterestPercent = Me.TXTInterests.Text
                            vInterestAmount1 = Me.TXTInterestsAmount.Text
                            vProfitPercent = Me.TXTProfit.Text
                            vProfitAmount1 = Me.TXTProfitAmount.Text
                            vAddCashWord = Me.TXTAddCash.Text
                            vAddCashAmount1 = Me.TXTAddCashRes.Text
                            vAddCreditWord = Me.TXTAddCredit.Text
                            vAddCreditAmount1 = Me.TXTAddCreditRes.Text
                            vSpecialDisc1Word = Me.TXTDiscount1.Text
                            vSpecialDisc1Amount1 = Me.TXTDiscount1Res.Text

                            vQuery = "exec dbo.USP_PS_PriceStructureSubSet '" & vDocno & "','" & vItemcode & "','" & vItemName & "','" & vBuyUnitCode & "', " _
                            & " '" & vBuyer & "'," & vDO & "," & vPriceSet & ",'" & vBillDiscWord & "'," & vBillDiscAmount1 & "," & vBillDiscRes & ", " _
                            & " '" & vFollowDisc1Word & "'," & vFollowDisc1Amount1 & "," & vFollowDisc1Res & ",'" & vFollowDisc2Word & "'," & vFollowDisc2Amount1 & "," & vFollowDisc2Res & ", " _
                            & " '" & vFollowDisc3Word & "'," & vFollowDisc3Amount1 & "," & vFollowDisc3Res & ",'" & vRebateWord & "'," & vRebateAmount1 & "," & vRebateRes & ", " _
                            & " '" & vSpecialDiscWord & "'," & vSpecialDiscAmount1 & "," & vSpecialDiscRes & ",'" & vMissProfitWord & "'," & vMissProfitAmount1 & "," & vMissProfitRes & ", " _
                            & " '" & vSendWord & "'," & vSendAmount1 & "," & vSendRes & ",'" & vCustSendWord & "'," & vCustSendAmount & "," & vCustSendRes & ", " _
                            & " '" & vAdvertiseWord & "'," & vAdvertiseAmount1 & "," & vAdvertiseRes & ",'" & vMarketWord & "'," & vMarketAmount1 & "," & vMarketRes & ", " _
                            & " '" & vTaxWord & "'," & vTaxAmount1 & "," & vTaxRes & ",'" & vInstallWord & "'," & vInstallAmount1 & "," & vInstallRes & ",'" & vServiceWord & "'," & vServiceAmount1 & "," & vServiceRes & ", " _
                            & " '" & vPointWord & "'," & vPointAmount1 & "," & vPointRes & ",'" & vDiscMemberWord & "'," & vDiscMemberAmount1 & "," & vDiscMemberRes & ",'" & vTargetWord & "'," & vTargetAmount1 & "," & vTargetRes & ",'" & vGiftWord & "'," & vGiftAmount1 & "," & vGiftRes & ", " _
                            & " '" & vCommissionWord & "'," & vCommissionAmount1 & "," & vCommissionRes & "," & vBegProfitPercent & "," & vBegProfitAmount1 & "," & vInterestPercent & "," & vInterestAmount1 & ", " _
                            & " " & vProfitPercent & "," & vProfitAmount1 & ",'" & vAddCashWord & "'," & vAddCashAmount1 & ",'" & vAddCreditWord & "'," & vAddCreditAmount1 & ",'" & vSpecialDisc1Word & "'," & vSpecialDisc1Amount1 & " "
                            vCMD = New SqlCommand(vQuery, vConnection)
                            vCMD.ExecuteNonQuery()

                            vQuery = "exec dbo.USP_PS_PriceUpdateClear '" & vDocno & "','" & vItemcode & "'"
                            vCMD = New SqlCommand(vQuery, vConnection)
                            vCMD.ExecuteNonQuery()

                            For i = 0 To Me.ListView101.Items.Count - 1

                                vItemUnitCode = ListView101.Items(i).SubItems(0).Text
                                vFromQTY = ListView101.Items(i).SubItems(3).Text
                                vToQTY = ListView101.Items(i).SubItems(4).Text
                                If ListView101.Items(i).SubItems(1).Text = "ขายเงินสด" Then
                                    vSaleType = 0
                                Else
                                    vSaleType = 1
                                End If
                                If ListView101.Items(i).SubItems(2).Text = "รับเอง" Then
                                    vTransSportType = 0
                                Else
                                    vTransSportType = 1
                                End If
                                vPriceLevel1 = ListView101.Items(i).SubItems(5).Text
                                vPriceLevel2 = ListView101.Items(i).SubItems(6).Text

                                If vPriceLevel1 < 0 Or vPriceLevel2 < 0 Then
                                    MsgBox("การกำหนดราคาสินค้าไม่ควรต่ำกว่า 0 ", MsgBoxStyle.Critical, "Send Error Message")
                                    vQuery = "rollback tran"
                                    vCMD = New SqlCommand(vQuery, vConnection)
                                    vCMD.ExecuteNonQuery()
                                    Exit Sub
                                End If
                                vLineNumber = i
                                If ListView101.Items(i).SubItems(7).Text = "" Then
                                    vIsPriceUpdate = 0
                                Else
                                    vIsPriceUpdate = ListView101.Items(i).SubItems(7).Text
                                End If

                                vDatePriceUpdate = ListView101.Items(i).SubItems(8).Text
                                vDateExpire = ListView101.Items(i).SubItems(9).Text

                                vQuery = "exec dbo.USP_PS_PriceUpdateSet '" & vDocno & "','" & vItemcode & "','" & vItemName & "','" & vItemUnitCode & "', " _
                                & " " & vFromQTY & "," & vToQTY & "," & vSaleType & "," & vTransSportType & "," & vPriceLevel1 & "," & vPriceLevel2 & ", " _
                                & " " & vLineNumber & "," & vIsPriceUpdate & ",'" & vDatePriceUpdate & "','" & vDateExpire & "' "
                                vCMD = New SqlCommand(vQuery, vConnection)
                                vCMD.ExecuteNonQuery()


                            Next

                            vQuery = "commit tran"
                            vCMD = New SqlCommand(vQuery, vConnection)
                            vCMD.ExecuteNonQuery()
                            Me.Cursor = Cursors.Arrow

                            MsgBox("บันทึกข้อมูลเลขที่เอกสาร " & vDocno & " และ รหัสสินค้า  " & vItemcode & " นี้เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information")
                            Call ClearItemDescription()
                            Call ClearItemData()
                            Me.TXTItemCode.Text = ""
                            Call RefreshData(vDocno)


                        Catch ex As Exception
                            vQuery = "rollback tran"
                            vCMD = New SqlCommand(vQuery, vConnection)
                            vCMD.ExecuteNonQuery()
                            MsgBox(Err.Description)
                            Me.Cursor = Cursors.Arrow
                        End Try
                    Else
                        MsgBox("ไม่มีเลขที่เอกสาร กรุณารันเลขที่เอกสารด้วย", MsgBoxStyle.Critical, "Send Error")
                    End If
                Else
                    MsgBox("เอกสารเลขที่ " & vDocno & " นี้ได้ถูกอนุมัติแล้ว  ถ้าต้องการบันทึกต้องเปลี่ยนเป็นเลขที่เอกสารเลขที่ใหม่", MsgBoxStyle.Critical, "Send Error Message")
                End If
            Else
                MsgBox("เอกสารเลขที่ " & vDocno & " นี้ได้ถูกยกเลิกแล้ว  ถ้าต้องการบันทึกต้องเปลี่ยนเป็นเลขที่เอกสารเลขที่ใหม่", MsgBoxStyle.Critical, "Send Error Message")
            End If
        Else
            MsgBox("ต้องมีเลขที่เอกสาร ,รหัสสินค้า ,ราคาที่กำหนด ,รายการกำหนดราคา ถึงจะบันทึกข้อมูลได้ กรุณาตรวจสอบข้อมูลด้วย ", MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub SaveData()
        Dim vDocno As String
        Dim vDocdate As String
        Dim vItemcode As String
        Dim vItemName As String
        Dim vBuyUnitCode As String
        Dim vBuyer As String
        Dim vDO As Double
        Dim vPriceSet As Double

        Dim vBillDiscWord As String
        Dim vBillDiscAmount1 As Double
        Dim vBillDiscRes As Double

        Dim vFollowDisc1Word As String
        Dim vFollowDisc1Amount1 As Double
        Dim vFollowDisc1Res As Double

        Dim vFollowDisc2Word As String
        Dim vFollowDisc2Amount1 As Double
        Dim vFollowDisc2Res As Double

        Dim vFollowDisc3Word As String
        Dim vFollowDisc3Amount1 As Double
        Dim vFollowDisc3Res As Double

        Dim vRebateWord As String
        Dim vRebateAmount1 As Double
        Dim vRebateRes As Double

        Dim vSpecialDiscWord As String
        Dim vSpecialDiscAmount1 As Double
        Dim vSpecialDiscRes As Double

        Dim vMissProfitWord As String
        Dim vMissProfitAmount1 As Double
        Dim vMissProfitRes As Double

        Dim vSendWord As String
        Dim vSendAmount1 As Double
        Dim vSendRes As Double

        Dim vCustSendWord As String
        Dim vCustSendAmount1 As Double
        Dim vCustSendRes As Double

        Dim vAdvertiseWord As String
        Dim vAdvertiseAmount1 As Double
        Dim vAdvertiseRes As Double

        Dim vMarketWord As String
        Dim vMarketAmount1 As Double
        Dim vMarketRes As Double

        Dim vTaxWord As String
        Dim vTaxAmount1 As Double
        Dim vTaxRes As Double

        Dim vInstallWord As String
        Dim vInstallAmount1 As Double
        Dim vInstallRes As Double

        Dim vServiceWord As String
        Dim vServiceAmount1 As Double
        Dim vServiceRes As Double

        Dim vDiscMemberWord As String
        Dim vDiscMemberAmount1 As Double
        Dim vDiscMemberRes As Double

        Dim vPointWord As String
        Dim vPointAmount1 As Double
        Dim vPointRes As Double

        Dim vTargetWord As String
        Dim vTargetAmount1 As Double
        Dim vTargetRes As Double

        Dim vGiftWord As String
        Dim vGiftAmount1 As Double
        Dim vGiftRes As Double

        Dim vCommissionWord As String
        Dim vCommissionAmount1 As Double
        Dim vCommissionRes As Double

        Dim vBegProfitPercent As Double
        Dim vBegProfitAmount1 As Double
        Dim vInterestPercent As Double
        Dim vInterestAmount1 As Double
        Dim vProfitPercent As Double
        Dim vProfitAmount1 As Double
        Dim vAddCashWord As String
        Dim vAddCashAmount1 As Double
        Dim vAddCreditWord As String
        Dim vAddCreditAmount1 As Double
        Dim vSpecialDisc1Word As String
        Dim vSpecialDisc1Amount1 As String
        Dim vFromQTY As Double
        Dim vToQTY As Double
        Dim vSaleType As Integer
        Dim vTransSportType As Integer
        Dim vPriceLevel1 As Double
        Dim vPriceLevel2 As Double
        Dim vLineNumber As Integer
        Dim vIsPriceUpdate As Integer
        Dim vDatePriceUpdate As String
        Dim vDateExpire As String
        Dim i As Integer
        Dim vItemUnitCode As String
        Dim n As Integer
        Dim vds As DataSet
        Dim vdt As DataTable


        If vOldDocNO <> "" And Me.TextBoxDocno.Text <> "" And Me.ListView103.Items.Count > 0 Then
            vDocno = Me.TextBoxDocno.Text
            If vIsCancel = 0 Then
                If vIsConfirm = 0 Then
                    vCheckBuyer = Me.TextBoxBuyer.Text
                    Me.Cursor = Cursors.WaitCursor
                    Try

                        vQuery = "begin tran"
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vCMD.ExecuteNonQuery()

                        vDocno = Me.TextBoxDocno.Text
                        vDocdate = Me.DocDate.Text

                        vQuery = "exec dbo.USP_PS_PriceStructureSet '" & vDocno & "','" & vDocdate & "' "
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vCMD.ExecuteNonQuery()


                        vQuery = "exec dbo.USP_PS_PriceStructureSubList '" & vOldDocNO & "'"
                        da = New SqlDataAdapter(vQuery, vConnection)
                        ds = New DataSet
                        da.Fill(ds, "SaveData")
                        dt = ds.Tables("SaveData")

                        If dt.Rows.Count > 0 Then
                            For i = 0 To dt.Rows.Count - 1
                                vItemcode = dt.Rows(i).Item("itemcode")
                                vItemName = dt.Rows(i).Item("itemname")
                                vBuyUnitCode = dt.Rows(i).Item("buyunitcode")
                                vBuyer = Me.TextBoxBuyer.Text
                                vDO = dt.Rows(i).Item("d/o")
                                vPriceSet = dt.Rows(i).Item("priceset")

                                vBillDiscWord = dt.Rows(i).Item("discountbillword")
                                vBillDiscAmount1 = dt.Rows(i).Item("discountbillamount")
                                vBillDiscRes = dt.Rows(i).Item("acccost")

                                vFollowDisc1Word = dt.Rows(i).Item("discountfollow1word")
                                vFollowDisc1Amount1 = dt.Rows(i).Item("discountfollow1amount")
                                vFollowDisc1Res = dt.Rows(i).Item("discountfollow1after")

                                vFollowDisc2Word = dt.Rows(i).Item("discountfollow2word")
                                vFollowDisc2Amount1 = dt.Rows(i).Item("discountfollow2amount")
                                vFollowDisc2Res = dt.Rows(i).Item("discountfollow2after")

                                vFollowDisc3Word = dt.Rows(i).Item("discountfollow3word")
                                vFollowDisc3Amount1 = dt.Rows(i).Item("discountfollow3amount")
                                vFollowDisc3Res = dt.Rows(i).Item("discountfollow3after")

                                vRebateWord = dt.Rows(i).Item("discountrebateword")
                                vRebateAmount1 = dt.Rows(i).Item("discountrebateamount")
                                vRebateRes = dt.Rows(i).Item("discountrebateafter")

                                vSpecialDiscWord = dt.Rows(i).Item("discountspecialword")
                                vSpecialDiscAmount1 = dt.Rows(i).Item("discountspecialamount")
                                vSpecialDiscRes = dt.Rows(i).Item("netcost")

                                vMissProfitWord = dt.Rows(i).Item("lossbudgetword")
                                vMissProfitAmount1 = dt.Rows(i).Item("lossbudgetamount")
                                vMissProfitRes = dt.Rows(i).Item("lossbudgetafter")

                                vSendWord = dt.Rows(i).Item("transferinword")
                                vSendAmount1 = dt.Rows(i).Item("transferinamount")
                                vSendRes = dt.Rows(i).Item("transferinafter")

                                vCustSendWord = dt.Rows(i).Item("transferoutword")
                                vCustSendAmount1 = dt.Rows(i).Item("transferoutamount")
                                vCustSendRes = dt.Rows(i).Item("transferoutafter")

                                vAdvertiseWord = dt.Rows(i).Item("advertiseword")
                                vAdvertiseAmount1 = dt.Rows(i).Item("advertiseamount")
                                vAdvertiseRes = dt.Rows(i).Item("advertiseafter")

                                vMarketWord = dt.Rows(i).Item("marketingbudgetword")
                                vMarketAmount1 = dt.Rows(i).Item("marketingbudgetamount")
                                vMarketRes = dt.Rows(i).Item("marketingbudgetafter")

                                vTaxWord = dt.Rows(i).Item("vatword")
                                vTaxAmount1 = dt.Rows(i).Item("vatamount")
                                vTaxRes = dt.Rows(i).Item("vatafter")

                                vInstallWord = dt.Rows(i).Item("setupword")
                                vInstallAmount1 = dt.Rows(i).Item("setupamount")
                                vInstallRes = dt.Rows(i).Item("setupafter")

                                vServiceWord = dt.Rows(i).Item("serviceword")
                                vServiceAmount1 = dt.Rows(i).Item("serviceamount")
                                vServiceRes = dt.Rows(i).Item("marketcost")

                                vDiscMemberWord = dt.Rows(i).Item("memberdiscountword")
                                vDiscMemberAmount1 = dt.Rows(i).Item("memberdiscountamount")
                                vDiscMemberRes = dt.Rows(i).Item("memberdiscountafter")

                                vPointWord = dt.Rows(i).Item("pointword")
                                vPointAmount1 = dt.Rows(i).Item("pointafter")
                                vPointRes = dt.Rows(i).Item("pointamount")

                                vTargetWord = dt.Rows(i).Item("targetword")
                                vTargetAmount1 = dt.Rows(i).Item("targetamount")
                                vTargetRes = dt.Rows(i).Item("targetafter")

                                vGiftWord = dt.Rows(i).Item("premiumword")
                                vGiftAmount1 = dt.Rows(i).Item("premiumamount")
                                vGiftRes = dt.Rows(i).Item("premiumafter")

                                vCommissionWord = dt.Rows(i).Item("commissionword")
                                vCommissionAmount1 = dt.Rows(i).Item("commissionamount")
                                vCommissionRes = dt.Rows(i).Item("commissionafter")

                                vBegProfitPercent = dt.Rows(i).Item("grossprofitpercent")
                                vBegProfitAmount1 = dt.Rows(i).Item("grossprofitamount")
                                vInterestPercent = dt.Rows(i).Item("intereststockpercent")
                                vInterestAmount1 = dt.Rows(i).Item("intereststockamount")
                                vProfitPercent = dt.Rows(i).Item("profitpercent")
                                vProfitAmount1 = dt.Rows(i).Item("profitamount")
                                vAddCashWord = dt.Rows(i).Item("transfervalueword")
                                vAddCashAmount1 = dt.Rows(i).Item("transfervalueamount")
                                vAddCreditWord = dt.Rows(i).Item("creditvalueword")
                                vAddCreditAmount1 = dt.Rows(i).Item("creditvalueamount")
                                vSpecialDisc1Word = dt.Rows(i).Item("specialvalueword")
                                vSpecialDisc1Amount1 = dt.Rows(i).Item("specialvalueword")

                                vQuery = "exec dbo.USP_PS_PriceStructureSubSet '" & vDocno & "','" & vItemcode & "','" & vItemName & "','" & vBuyUnitCode & "', " _
                                & " '" & vBuyer & "'," & vDO & "," & vPriceSet & ",'" & vBillDiscWord & "'," & vBillDiscAmount1 & "," & vBillDiscRes & ", " _
                                & " '" & vFollowDisc1Word & "'," & vFollowDisc1Amount1 & "," & vFollowDisc1Res & ",'" & vFollowDisc2Word & "'," & vFollowDisc2Amount1 & "," & vFollowDisc2Res & ", " _
                                & " '" & vFollowDisc3Word & "'," & vFollowDisc3Amount1 & "," & vFollowDisc3Res & ",'" & vRebateWord & "'," & vRebateAmount1 & "," & vRebateRes & ", " _
                                & " '" & vSpecialDiscWord & "'," & vSpecialDiscAmount1 & "," & vSpecialDiscRes & ",'" & vMissProfitWord & "'," & vMissProfitAmount1 & "," & vMissProfitRes & ", " _
                                & " '" & vSendWord & "'," & vSendAmount1 & "," & vSendRes & ",'" & vCustSendWord & "'," & vCustSendAmount & "," & vCustSendRes & ", " _
                                & " '" & vAdvertiseWord & "'," & vAdvertiseAmount1 & "," & vAdvertiseRes & ",'" & vMarketWord & "'," & vMarketAmount1 & "," & vMarketRes & ", " _
                                & " '" & vTaxWord & "'," & vTaxAmount1 & "," & vTaxRes & ",'" & vInstallWord & "'," & vInstallAmount1 & "," & vInstallRes & ",'" & vServiceWord & "'," & vServiceAmount1 & "," & vServiceRes & ", " _
                                & " '" & vPointWord & "'," & vPointAmount1 & "," & vPointRes & ",'" & vDiscMemberWord & "'," & vDiscMemberAmount1 & "," & vDiscMemberRes & ",'" & vTargetWord & "'," & vTargetAmount1 & "," & vTargetRes & ",'" & vGiftWord & "'," & vGiftAmount1 & "," & vGiftRes & ", " _
                                & " '" & vCommissionWord & "'," & vCommissionAmount1 & "," & vCommissionRes & "," & vBegProfitPercent & "," & vBegProfitAmount1 & "," & vInterestPercent & "," & vInterestAmount1 & ", " _
                                & " " & vProfitPercent & "," & vProfitAmount1 & ",'" & vAddCashWord & "'," & vAddCashAmount1 & ",'" & vAddCreditWord & "'," & vAddCreditAmount1 & ",'" & vSpecialDisc1Word & "'," & vSpecialDisc1Amount1 & " "
                                vCMD = New SqlCommand(vQuery, vConnection)
                                vCMD.ExecuteNonQuery()

                                vQuery = "exec dbo.USP_PS_PriceUpdateClear '" & vDocno & "','" & vItemcode & "'"
                                vCMD = New SqlCommand(vQuery, vConnection)
                                vCMD.ExecuteNonQuery()

                                vQuery = "exec dbo.USP_PS_PriceList '" & vOldDocNO & "','" & vItemcode & "'"
                                da = New SqlDataAdapter(vQuery, vConnection)
                                vds = New DataSet
                                da.Fill(vds, "PriceList")
                                vdt = vds.Tables("PriceList")

                                If vdt.Rows.Count > 0 Then
                                    For n = 0 To vdt.Rows.Count - 1

                                        vItemUnitCode = vdt.Rows(n).Item("saleunitcode")
                                        vFromQTY = vdt.Rows(n).Item("fromqty")
                                        vToQTY = vdt.Rows(n).Item("toqty")
                                        vSaleType = vdt.Rows(n).Item("saletype")
                                        vTransSportType = vdt.Rows(n).Item("transporttype")
                                        vPriceLevel1 = vdt.Rows(n).Item("priceset")
                                        vPriceLevel2 = vdt.Rows(n).Item("priceset2")
                                        vLineNumber = vdt.Rows(n).Item("linenumber")
                                        vIsPriceUpdate = vdt.Rows(n).Item("isupdate")
                                        If Microsoft.VisualBasic.IsDBNull(vdt.Rows(n).Item("updatedate")) Then
                                            vDatePriceUpdate = ""
                                        Else
                                            vDatePriceUpdate = vdt.Rows(n).Item("updatedate")
                                        End If
                                        If Microsoft.VisualBasic.IsDBNull(vdt.Rows(n).Item("stopdate")) Then
                                            vDateExpire = ""
                                        Else
                                            vDateExpire = vdt.Rows(n).Item("stopdate")
                                        End If
                                        vQuery = "exec dbo.USP_PS_PriceUpdateSet '" & vDocno & "','" & vItemcode & "','" & vItemName & "','" & vItemUnitCode & "', " _
                                        & " " & vFromQTY & "," & vToQTY & "," & vSaleType & "," & vTransSportType & "," & vPriceLevel1 & "," & vPriceLevel2 & ", " _
                                        & " " & vLineNumber & "," & vIsPriceUpdate & ",'" & vDatePriceUpdate & "','" & vDateExpire & "' "
                                        vCMD = New SqlCommand(vQuery, vConnection)
                                        vCMD.ExecuteNonQuery()
                                    Next
                                End If
                            Next
                        End If
                        vQuery = "commit tran"
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vCMD.ExecuteNonQuery()
                        Me.Cursor = Cursors.Arrow

                        vOldDocNO = ""
                        Call ClearItemDescription()
                        Call RefreshData(vDocno)

                    Catch ex As Exception
                        vQuery = "rollback tran"
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vCMD.ExecuteNonQuery()
                        MsgBox(Err.Description)
                        Me.Cursor = Cursors.Arrow
                    End Try


                Else
                    MsgBox("เอกสารเลขที่ " & vDocno & " นี้ได้ถูกอนุมัติแล้ว  ถ้าต้องการบันทึกต้องเปลี่ยนเป็นเลขที่เอกสารเลขที่ใหม่", MsgBoxStyle.Critical, "Send Error Message")
                End If
            Else
                MsgBox("เอกสารเลขที่ " & vDocno & " นี้ได้ถูกยกเลิกแล้ว  ถ้าต้องการบันทึกต้องเปลี่ยนเป็นเลขที่เอกสารเลขที่ใหม่", MsgBoxStyle.Critical, "Send Error Message")
            End If
        End If
    End Sub
    Private Sub RefreshData(ByVal vDocno As String)
        Dim vItemPriceList As ListViewItem
        Dim i As Integer

        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vBillDiscRes As Double
        Dim vFollowDisc1Res As Double
        Dim vFollowDisc2Res As Double
        Dim vFollowDisc3Res As Double
        Dim vRebateRes As Double
        Dim vSpecialDiscRes As Double
        Dim vMissProfitRes As Double
        Dim vSendRes As Double
        Dim vCustSendRes As Double
        Dim vAdvertiseRes As Double
        Dim vMarketRes As Double
        Dim vTaxRes As Double
        Dim vInstallRes As Double
        Dim vServiceRes As Double
        Dim vPointRes As Double
        Dim vTargetRes As Double
        Dim vGiftRes As Double
        Dim vCommissionRes As Double
        Dim vBegProfitAmountRes As Double
        Dim vInterestsAmountRes As Double
        Dim vProfitAmountRes As Double
        Dim vAddCashRes As Double
        Dim vAddCreditRes As Double
        Dim vDiscount1Res As Double
        Dim vDiscMemberRes As Double

        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_PS_PriceStructureSubList '" & vDocno & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "RefreshData")
        dt = ds.Tables("RefreshData")

        Me.ListView103.Items.Clear()

        If dt.Rows.Count > 0 Then
            Me.Cursor = Cursors.WaitCursor
            vIsCancel = dt.Rows(i).Item("iscancel")
            vIsConfirm = dt.Rows(i).Item("isconfirm")
            vIsOpen = 1
            Me.TextBoxBuyer.Text = dt.Rows(i).Item("buyer")
            vCheckBuyer = dt.Rows(i).Item("buyer")

            If vIsCancel = 1 Then
                Me.PB101.Visible = False
                Me.PB102.Visible = False
                Me.PB103.Visible = True
            ElseIf vIsConfirm = 1 Then
                Me.PB101.Visible = False
                Me.PB102.Visible = True
                Me.PB103.Visible = False
            Else
                Me.PB101.Visible = True
                Me.PB102.Visible = False
                Me.PB103.Visible = False
            End If

            For i = 0 To dt.Rows.Count - 1
                vItemPriceList = Me.ListView103.Items.Add(dt.Rows(i).Item("itemcode"))
                vItemPriceList.SubItems.Add(1).Text = dt.Rows(i).Item("itemname")
                vItemPriceList.SubItems.Add(2).Text = dt.Rows(i).Item("buyunitcode")

                vDO = dt.Rows(i).Item("d/o")
                vItemPriceList.SubItems.Add(3).Text = Format(vDO, "##,##0.00")

                vPriceSet = dt.Rows(i).Item("priceset")
                vItemPriceList.SubItems.Add(4).Text = Format(vPriceSet, "##,##0.00")

                vItemPriceList.SubItems.Add(5).Text = dt.Rows(i).Item("discountbillword")
                vBillDiscRes = dt.Rows(i).Item("acccost")
                vItemPriceList.SubItems.Add(6).Text = Format(vBillDiscRes, "##,##0.00")

                vItemPriceList.SubItems.Add(7).Text = dt.Rows(i).Item("discountfollow1word")
                vFollowDisc1Res = dt.Rows(i).Item("discountfollow1after")
                vItemPriceList.SubItems.Add(8).Text = Format(vFollowDisc1Res, "##,##0.00")

                vItemPriceList.SubItems.Add(9).Text = dt.Rows(i).Item("discountfollow2word")
                vFollowDisc2Res = dt.Rows(i).Item("discountfollow2after")
                vItemPriceList.SubItems.Add(10).Text = Format(vFollowDisc2Res, "##,##0.00")

                vItemPriceList.SubItems.Add(11).Text = dt.Rows(i).Item("discountfollow3word")
                vFollowDisc3Res = dt.Rows(i).Item("discountfollow3after")
                vItemPriceList.SubItems.Add(12).Text = Format(vFollowDisc3Res, "##,##0.00")

                vItemPriceList.SubItems.Add(13).Text = dt.Rows(i).Item("discountrebateword")
                vRebateRes = dt.Rows(i).Item("discountrebateafter")
                vItemPriceList.SubItems.Add(14).Text = Format(vRebateRes, "##,##0.00")

                vItemPriceList.SubItems.Add(15).Text = dt.Rows(i).Item("discountspecialword")
                vSpecialDiscRes = dt.Rows(i).Item("netcost")
                vItemPriceList.SubItems.Add(16).Text = Format(vSpecialDiscRes, "##,##0.00")

                vItemPriceList.SubItems.Add(17).Text = dt.Rows(i).Item("lossbudgetword")
                vMissProfitRes = dt.Rows(i).Item("lossbudgetafter")
                vItemPriceList.SubItems.Add(18).Text = Format(vMissProfitRes, "##,##0.00")

                vItemPriceList.SubItems.Add(19).Text = dt.Rows(i).Item("transferinword")
                vSendRes = dt.Rows(i).Item("transferinafter")
                vItemPriceList.SubItems.Add(20).Text = Format(vSendRes, "##,##0.00")

                vItemPriceList.SubItems.Add(21).Text = dt.Rows(i).Item("transferoutword")
                vCustSendRes = dt.Rows(i).Item("transferoutafter")
                vItemPriceList.SubItems.Add(22).Text = Format(vCustSendRes, "##,##0.00")

                vItemPriceList.SubItems.Add(23).Text = dt.Rows(i).Item("advertiseword")
                vAdvertiseRes = dt.Rows(i).Item("advertiseafter")
                vItemPriceList.SubItems.Add(24).Text = Format(vAdvertiseRes, "##,##0.00")

                vItemPriceList.SubItems.Add(25).Text = dt.Rows(i).Item("marketingbudgetword")
                vMarketRes = dt.Rows(i).Item("marketingbudgetafter")
                vItemPriceList.SubItems.Add(26).Text = Format(vMarketRes, "##,##0.00")

                vItemPriceList.SubItems.Add(27).Text = dt.Rows(i).Item("vatword")
                vTaxRes = dt.Rows(i).Item("vatafter")
                vItemPriceList.SubItems.Add(28).Text = Format(vTaxRes, "##,##0.00")

                vItemPriceList.SubItems.Add(29).Text = dt.Rows(i).Item("setupword")
                vInstallRes = dt.Rows(i).Item("setupafter")
                vItemPriceList.SubItems.Add(30).Text = Format(vInstallRes, "##,##0.00")

                vItemPriceList.SubItems.Add(31).Text = dt.Rows(i).Item("serviceword")
                vServiceRes = dt.Rows(i).Item("marketcost")
                vItemPriceList.SubItems.Add(32).Text = Format(vServiceRes, "##,##0.00")

                vItemPriceList.SubItems.Add(33).Text = dt.Rows(i).Item("pointword")
                vPointRes = dt.Rows(i).Item("pointafter")
                vItemPriceList.SubItems.Add(34).Text = Format(vPointRes, "##,##0.00")

                vItemPriceList.SubItems.Add(35).Text = dt.Rows(i).Item("MemberDiscountword")
                vDiscMemberRes = dt.Rows(i).Item("MemberDiscountafter")
                vItemPriceList.SubItems.Add(36).Text = Format(vDiscMemberRes, "##,##0.00")

                vItemPriceList.SubItems.Add(37).Text = dt.Rows(i).Item("targetword")
                vTargetRes = dt.Rows(i).Item("targetafter")
                vItemPriceList.SubItems.Add(38).Text = Format(vTargetRes, "##,##0.00")

                vItemPriceList.SubItems.Add(39).Text = dt.Rows(i).Item("premiumword")
                vGiftRes = dt.Rows(i).Item("premiumafter")
                vItemPriceList.SubItems.Add(40).Text = Format(vGiftRes, "##,##0.00")

                vItemPriceList.SubItems.Add(41).Text = dt.Rows(i).Item("commissionword")
                vCommissionRes = dt.Rows(i).Item("commissionafter")
                vItemPriceList.SubItems.Add(42).Text = Format(vCommissionRes, "##,##0.00")

                vItemPriceList.SubItems.Add(43).Text = dt.Rows(i).Item("grossprofitpercent")
                vBegProfitAmountRes = dt.Rows(i).Item("grossprofitamount")
                vItemPriceList.SubItems.Add(44).Text = Format(vBegProfitAmountRes, "##,##0.00")

                vItemPriceList.SubItems.Add(45).Text = dt.Rows(i).Item("intereststockpercent")
                vInterestsAmountRes = dt.Rows(i).Item("intereststockamount")
                vItemPriceList.SubItems.Add(46).Text = Format(vInterestsAmountRes, "##,##0.00")

                vItemPriceList.SubItems.Add(47).Text = dt.Rows(i).Item("profitpercent")
                vProfitAmountRes = dt.Rows(i).Item("profitamount")
                vItemPriceList.SubItems.Add(48).Text = Format(vProfitAmountRes, "##,##0.00")

                vItemPriceList.SubItems.Add(49).Text = dt.Rows(i).Item("transfervalueword")
                vAddCashRes = dt.Rows(i).Item("transfervalueamount")
                vItemPriceList.SubItems.Add(50).Text = Format(vAddCashRes, "##,##0.00")

                vItemPriceList.SubItems.Add(51).Text = dt.Rows(i).Item("creditvalueword")
                vAddCreditRes = dt.Rows(i).Item("creditvalueamount")
                vItemPriceList.SubItems.Add(52).Text = Format(vAddCreditRes, "##,##0.00")

                vItemPriceList.SubItems.Add(53).Text = dt.Rows(i).Item("specialvalueword")
                vDiscount1Res = dt.Rows(i).Item("specialvalueamount")
                vItemPriceList.SubItems.Add(54).Text = Format(vDiscount1Res, "##,##0.00")
            Next
            vIsOpen = 1
            Me.Cursor = Cursors.Arrow
            Me.TextBoxDocno.Text = UCase(vDocno)
            Me.ListView101.Focus()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub ListView102_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView102.KeyDown
        Dim vIndex As Integer
        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vBillDiscRes As Double
        Dim vFollowDisc1Res As Double
        Dim vFollowDisc2Res As Double
        Dim vFollowDisc3Res As Double
        Dim vRebateRes As Double
        Dim vSpecialDiscRes As Double
        Dim vMissProfitRes As Double
        Dim vSendRes As Double
        Dim vCustSendRes As Double
        Dim vAdvertiseRes As Double
        Dim vMarketRes As Double
        Dim vTaxRes As Double
        Dim vInstallRes As Double
        Dim vServiceRes As Double
        Dim vPointRes As Double
        Dim vTargetRes As Double
        Dim vGiftRes As Double
        Dim vCommissionRes As Double
        Dim vBegProfitAmountRes As Double
        Dim vInterestsAmountRes As Double
        Dim vProfitAmountRes As Double
        Dim vAddCashRes As Double
        Dim vAddCreditRes As Double
        Dim vDiscount1Res As Double
        Dim vCreditCardRes As Double

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then

            If Me.ListView102.Items.Count > 0 Then

                Me.Cursor = Cursors.WaitCursor
                vIndex = Me.ListView102.SelectedItems(0).Index
                Me.TXTItemCode.Text = Trim(Me.ListView102.Items(vIndex).SubItems(0).Text)
                Me.LBLItemName.Text = Trim(Me.ListView102.Items(vIndex).SubItems(1).Text)
                Me.CMBUnit.Items.Clear()
                Me.CMBUnit.Items.Add(Me.ListView102.Items(vIndex).SubItems(2).Text)
                Me.CMBUnit.Text = Me.CMBUnit.Items.Item(0)

                vDO = Me.ListView102.Items(vIndex).SubItems(3).Text
                Me.TXTDO.Text = Format(vDO, "##,##0.00")

                vPriceSet = Me.ListView102.Items(vIndex).SubItems(4).Text
                Me.TXTPriceSet.Text = Format(vPriceSet, "##,##0.00")

                Me.TXTBillDisc.Text = Me.ListView102.Items(vIndex).SubItems(5).Text
                vBillDiscRes = Me.ListView102.Items(vIndex).SubItems(6).Text
                Me.TXTBillDiscRes.Text = Format(vBillDiscRes, "##,##0.00")

                Me.TXTFollowDisc1.Text = Me.ListView102.Items(vIndex).SubItems(7).Text
                vFollowDisc1Res = Me.ListView102.Items(vIndex).SubItems(8).Text
                Me.TXTFollowDisc1Res.Text = Format(vFollowDisc1Res, "##,##0.00")

                Me.TXTFollowDisc2.Text = Me.ListView102.Items(vIndex).SubItems(9).Text
                vFollowDisc2Res = Me.ListView102.Items(vIndex).SubItems(10).Text
                Me.TXTFollowDisc2Res.Text = Format(vFollowDisc2Res, "##,##0.00")

                Me.TXTFollowDisc3.Text = Me.ListView102.Items(vIndex).SubItems(11).Text
                vFollowDisc3Res = Me.ListView102.Items(vIndex).SubItems(12).Text
                Me.TXTFollowDisc3Res.Text = Format(vFollowDisc3Res, "##,##0.00")

                Me.TXTRebate.Text = Me.ListView102.Items(vIndex).SubItems(13).Text
                vRebateRes = Me.ListView102.Items(vIndex).SubItems(14).Text
                Me.TXTRebateRes.Text = Format(vRebateRes, "##,##0.00")

                Me.TXTSpecialDisc.Text = Me.ListView102.Items(vIndex).SubItems(15).Text
                vSpecialDiscRes = Me.ListView102.Items(vIndex).SubItems(16).Text
                Me.TXTSpecialDiscRes.Text = Format(vSpecialDiscRes, "##,##0.00")

                Me.TXTMissProfit.Text = Me.ListView102.Items(vIndex).SubItems(17).Text
                vMissProfitRes = Me.ListView102.Items(vIndex).SubItems(18).Text
                Me.TXTMissProfitRes.Text = Format(vMissProfitRes, "##,##0.00")

                Me.TXTSend.Text = Me.ListView102.Items(vIndex).SubItems(19).Text
                vSendRes = Me.ListView102.Items(vIndex).SubItems(20).Text
                Me.TXTSendRes.Text = Format(vSendRes, "##,##0.00")

                Me.TXTCustSend.Text = Me.ListView102.Items(vIndex).SubItems(21).Text
                vCustSendRes = Me.ListView102.Items(vIndex).SubItems(22).Text
                Me.TXTCustSendRes.Text = Format(vCustSendRes, "##,##0.00")

                Me.TXTAdvertise.Text = Me.ListView102.Items(vIndex).SubItems(23).Text
                vAdvertiseRes = Me.ListView102.Items(vIndex).SubItems(24).Text
                Me.TXTAdvertiseRes.Text = Format(vAdvertiseRes, "##,##0.00")

                Me.TXTMarket.Text = Me.ListView102.Items(vIndex).SubItems(25).Text
                vMarketRes = Me.ListView102.Items(vIndex).SubItems(26).Text
                Me.TXTMarketRes.Text = Format(vMarketRes, "##,##0.00")

                Me.TXTTax.Text = Me.ListView102.Items(vIndex).SubItems(27).Text
                vTaxRes = Me.ListView102.Items(vIndex).SubItems(28).Text
                Me.TXTTaxRes.Text = Format(vTaxRes, "##,##0.00")

                Me.TXTInstall.Text = Me.ListView102.Items(vIndex).SubItems(29).Text
                vInstallRes = Me.ListView102.Items(vIndex).SubItems(30).Text
                Me.TXTInstallRes.Text = Format(vInstallRes, "##,##0.00")

                Me.TXTService.Text = Me.ListView102.Items(vIndex).SubItems(31).Text
                vServiceRes = Me.ListView102.Items(vIndex).SubItems(32).Text
                Me.TXTServiceRes.Text = Format(vServiceRes, "##,##0.00")

                Me.TXTPoint.Text = Me.ListView102.Items(vIndex).SubItems(33).Text
                vPointRes = Me.ListView102.Items(vIndex).SubItems(34).Text
                Me.TXTPointRes.Text = Format(vPointRes, "##,##0.00")

                Me.TXTDiscMember.Text = Me.ListView102.Items(vIndex).SubItems(35).Text
                vCreditCardRes = Me.ListView102.Items(vIndex).SubItems(36).Text
                Me.TXTDiscMemberRes.Text = Format(vCreditCardRes, "##,##0.00")

                Me.TXTTarget.Text = Me.ListView102.Items(vIndex).SubItems(37).Text
                vTargetRes = Me.ListView102.Items(vIndex).SubItems(38).Text
                Me.TXTTargetRes.Text = Format(vTargetRes, "##,##0.00")

                Me.TXTGift.Text = Me.ListView102.Items(vIndex).SubItems(39).Text
                vGiftRes = Me.ListView102.Items(vIndex).SubItems(40).Text
                Me.TXTGiftRes.Text = Format(vGiftRes, "##,##0.00")

                Me.TXTCommission.Text = Me.ListView102.Items(vIndex).SubItems(41).Text
                vCommissionRes = Me.ListView102.Items(vIndex).SubItems(42).Text
                Me.TXTCommissionRes.Text = Format(vCommissionRes, "##,##0.00")

                Me.TXTBegProfit.Text = Me.ListView102.Items(vIndex).SubItems(43).Text
                vBegProfitAmountRes = Me.ListView102.Items(vIndex).SubItems(44).Text
                Me.TXTBegProfitAmount.Text = Format(vBegProfitAmountRes, "##,##0.00")

                Me.TXTInterests.Text = Me.ListView102.Items(vIndex).SubItems(45).Text
                vInterestsAmountRes = Me.ListView102.Items(vIndex).SubItems(46).Text
                Me.TXTInterestsAmount.Text = Format(vInterestsAmountRes, "##,##0.00")

                Me.TXTProfit.Text = Format(Me.ListView102.Items(vIndex).SubItems(47).Text)
                vProfitAmountRes = Me.ListView102.Items(vIndex).SubItems(48).Text
                Me.TXTProfitAmount.Text = Format(vProfitAmountRes, "##,##0.00")

                Me.TXTAddCash.Text = Me.ListView102.Items(vIndex).SubItems(49).Text
                vAddCashRes = Me.ListView102.Items(vIndex).SubItems(50).Text
                Me.TXTAddCashRes.Text = Format(vAddCashRes, "##,##0.00")

                Me.TXTAddCredit.Text = Me.ListView102.Items(vIndex).SubItems(51).Text
                vAddCreditRes = Me.ListView102.Items(vIndex).SubItems(52).Text
                Me.TXTAddCreditRes.Text = Format(vAddCreditRes, "##,##0.00")

                Me.TXTDiscount1.Text = Me.ListView102.Items(vIndex).SubItems(53).Text
                vDiscount1Res = Me.ListView102.Items(vIndex).SubItems(54).Text
                Me.TXTDiscount1Res.Text = Format(vDiscount1Res, "##,##0.00")

                Me.Cursor = Cursors.Arrow
                Me.TXTDO.Focus()
                Me.ListView101.Items.Clear()
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub ListView102_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView102.SelectedIndexChanged

    End Sub

    Private Sub BTNSearchItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TXTSearchItemCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTSearchItemCode.KeyDown
        Dim vTemplateList As ListViewItem
        Dim i As Integer
        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vBillDiscRes As Double
        Dim vFollowDisc1Res As Double
        Dim vFollowDisc2Res As Double
        Dim vFollowDisc3Res As Double
        Dim vRebateRes As Double
        Dim vSpecialDiscRes As Double
        Dim vMissProfitRes As Double
        Dim vSendRes As Double
        Dim vCustSendRes As Double
        Dim vAdvertiseRes As Double
        Dim vMarketRes As Double
        Dim vTaxRes As Double
        Dim vInstallRes As Double
        Dim vServiceRes As Double
        Dim vPointRes As Double
        Dim vTargetRes As Double
        Dim vGiftRes As Double
        Dim vCommissionRes As Double
        Dim vBegProfitAmountRes As Double
        Dim vInterestsAmountRes As Double
        Dim vProfitAmountRes As Double
        Dim vAddCashRes As Double
        Dim vAddCreditRes As Double
        Dim vDiscount1Res As Double
        Dim vDiscMemberRes As Double

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.TXTSearchItemCode.Text <> "" Then
                vSearchItemCode = Me.TXTSearchItemCode.Text

                If Me.CMBDepartment.Text <> "" And Me.CMBBrand.Text <> "" Then

                    If InStr(Trim(Me.CMBDepartment.Text), "//") > 0 Then
                        vDepartmentCode = Microsoft.VisualBasic.Left(Trim(Me.CMBDepartment.Text), InStr(Trim(Me.CMBDepartment.Text), "//") - 1)
                    Else
                        vDepartmentCode = Trim(Me.CMBDepartment.Text)
                    End If
                    If InStr(Trim(Me.CMBBrand.Text), "//") > 0 Then
                        vBrandCode = Microsoft.VisualBasic.Left(Trim(Me.CMBBrand.Text), InStr(Trim(Me.CMBBrand.Text), "//") - 1)
                    Else
                        vBrandCode = Trim(Me.CMBBrand.Text)
                    End If
                    vQuery = "exec dbo.USP_PS_Template '" & vDepartmentCode & "','" & vBrandCode & "','" & vSearchItemCode & "',0 "
                ElseIf Me.CMBDepartment.Text = "" And Me.CMBBrand.Text <> "" Then
                    If InStr(Trim(Me.CMBBrand.Text), "//") > 0 Then
                        vBrandCode = Microsoft.VisualBasic.Left(Trim(Me.CMBBrand.Text), InStr(Trim(Me.CMBBrand.Text), "//") - 1)
                    Else
                        vBrandCode = Trim(Me.CMBBrand.Text)
                    End If
                    vQuery = "exec dbo.USP_PS_Template '','" & vBrandCode & "','" & vSearchItemCode & "',0 "

                Else
                    vQuery = "exec dbo.USP_PS_Template '','','" & vSearchItemCode & "',0 "
                End If

                    da = New SqlDataAdapter(vQuery, vConnection)
                    ds = New DataSet
                    da.Fill(ds, "Template")
                    dt = ds.Tables("Template")
                    If dt.Rows.Count > 0 Then
                        Me.ListView102.Items.Clear()
                        Me.Cursor = Cursors.WaitCursor
                        For i = 0 To dt.Rows.Count - 1
                            vTemplateList = ListView102.Items.Add(Trim(dt.Rows(i).Item("ItemCode")))
                            vTemplateList.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("ItemName"))
                            vTemplateList.SubItems.Add(2).Text = Trim(dt.Rows(i).Item("BuyUnitSale"))

                            vDO = Trim(dt.Rows(i).Item("D/O"))
                            vTemplateList.SubItems.Add(3).Text = Format(vDO, "##,##0.00")

                            vPriceSet = Trim(dt.Rows(i).Item("PriceSet"))
                            vTemplateList.SubItems.Add(4).Text = Format(vPriceSet, "##,##0.00")

                            vTemplateList.SubItems.Add(5).Text = Trim(dt.Rows(i).Item("DiscountBillWord"))
                            vBillDiscRes = Trim(dt.Rows(i).Item("DiscountBillAmount"))
                            vTemplateList.SubItems.Add(6).Text = Format(vBillDiscRes, "##,##0.00")

                            vTemplateList.SubItems.Add(7).Text = Trim(dt.Rows(i).Item("DiscountFollow1Word"))
                            vFollowDisc1Res = Trim(dt.Rows(i).Item("DiscountFollow1After"))
                            vTemplateList.SubItems.Add(8).Text = Format(vFollowDisc1Res, "##,##0.00")

                            vTemplateList.SubItems.Add(9).Text = Trim(dt.Rows(i).Item("DiscountFollow2Word"))
                            vFollowDisc2Res = Trim(dt.Rows(i).Item("DiscountFollow2After"))
                            vTemplateList.SubItems.Add(10).Text = Format(vFollowDisc2Res, "##,##0.00")

                            vTemplateList.SubItems.Add(11).Text = Trim(dt.Rows(i).Item("DiscountFollow3Word"))
                            vFollowDisc3Res = Trim(dt.Rows(i).Item("DiscountFollow3After"))
                            vTemplateList.SubItems.Add(12).Text = Format(vFollowDisc3Res, "##,##0.00")

                            vTemplateList.SubItems.Add(13).Text = Trim(dt.Rows(i).Item("DiscountRebateWord"))
                            vRebateRes = Trim(dt.Rows(i).Item("DiscountRebateAfter"))
                            vTemplateList.SubItems.Add(14).Text = Format(vRebateRes, "##,##0.00")

                            vTemplateList.SubItems.Add(15).Text = Trim(dt.Rows(i).Item("DiscountSpecialword"))
                            vSpecialDiscRes = Trim(dt.Rows(i).Item("NetCost"))
                            vTemplateList.SubItems.Add(16).Text = Format(vSpecialDiscRes, "##,##0.00")

                            vTemplateList.SubItems.Add(17).Text = Trim(dt.Rows(i).Item("LossBudgetWord"))
                            vMissProfitRes = Trim(dt.Rows(i).Item("LossBudgetAfter"))
                            vTemplateList.SubItems.Add(18).Text = Format(vMissProfitRes, "##,##0.00")

                            vTemplateList.SubItems.Add(19).Text = Trim(dt.Rows(i).Item("TransferInWord"))
                            vSendRes = Trim(dt.Rows(i).Item("TransferInAfter"))
                            vTemplateList.SubItems.Add(20).Text = Format(vSendRes, "##,##0.00")

                            vTemplateList.SubItems.Add(21).Text = Trim(dt.Rows(i).Item("TransferOutWord"))
                            vCustSendRes = Trim(dt.Rows(i).Item("TransferOutAfter"))
                            vTemplateList.SubItems.Add(22).Text = Format(vCustSendRes, "##,##0.00")

                            vTemplateList.SubItems.Add(23).Text = Trim(dt.Rows(i).Item("AdvertiseWord"))
                            vAdvertiseRes = Trim(dt.Rows(i).Item("AdvertiseAfter"))
                            vTemplateList.SubItems.Add(24).Text = Format(vAdvertiseRes, "##,##0.00")

                            vTemplateList.SubItems.Add(25).Text = Trim(dt.Rows(i).Item("MarketingBudgetWord"))
                            vMarketRes = Trim(dt.Rows(i).Item("MarketingBudgetAfter"))
                            vTemplateList.SubItems.Add(26).Text = Format(vMarketRes, "##,##0.00")

                            vTemplateList.SubItems.Add(27).Text = Trim(dt.Rows(i).Item("VatWord"))
                            vTaxRes = Trim(dt.Rows(i).Item("VatAfter"))
                            vTemplateList.SubItems.Add(28).Text = Format(vTaxRes, "##,##0.00")

                            vTemplateList.SubItems.Add(29).Text = Trim(dt.Rows(i).Item("SetupWord"))
                            vInstallRes = Trim(dt.Rows(i).Item("SetupAfter"))
                            vTemplateList.SubItems.Add(30).Text = Format(vInstallRes, "##,##0.00")

                            vTemplateList.SubItems.Add(31).Text = Trim(dt.Rows(i).Item("ServiceWord"))
                            vServiceRes = Trim(dt.Rows(i).Item("Marketcost"))
                            vTemplateList.SubItems.Add(32).Text = Format(vServiceRes, "##,##0.00")

                            vTemplateList.SubItems.Add(33).Text = Trim(dt.Rows(i).Item("Pointword"))
                            vPointRes = Trim(dt.Rows(i).Item("PointAfter"))
                            vTemplateList.SubItems.Add(34).Text = Format(vPointRes, "##,##0.00")

                            vTemplateList.SubItems.Add(53).Text = Trim(dt.Rows(0).Item("MemberDiscountWord"))
                            vDiscMemberRes = Trim(dt.Rows(0).Item("MemberDiscountAfter"))
                            vTemplateList.SubItems.Add(54).Text = Format(vDiscMemberRes, "##,##0.00")

                            vTemplateList.SubItems.Add(35).Text = Trim(dt.Rows(i).Item("TargetWord"))
                            vTargetRes = Trim(dt.Rows(i).Item("TargetAfter"))
                            vTemplateList.SubItems.Add(36).Text = Format(vTargetRes, "##,##0.00")

                            vTemplateList.SubItems.Add(37).Text = Trim(dt.Rows(i).Item("PremiumWord"))
                            vGiftRes = Trim(dt.Rows(i).Item("PremiumAfter"))
                            vTemplateList.SubItems.Add(38).Text = Format(vGiftRes, "##,##0.00")

                            vTemplateList.SubItems.Add(39).Text = Trim(dt.Rows(i).Item("CommissionWord"))
                            vCommissionRes = Trim(dt.Rows(i).Item("CommissionAfter"))
                            vTemplateList.SubItems.Add(40).Text = Format(vCommissionRes, "##,##0.00")

                            vTemplateList.SubItems.Add(41).Text = Format(Int(Trim(dt.Rows(i).Item("GrossProfitPercent"))), "##,##0.00")
                            vBegProfitAmountRes = Trim(dt.Rows(i).Item("GrossProfitAmount"))
                            vTemplateList.SubItems.Add(42).Text = Format(vBegProfitAmountRes, "##,##0.00")

                            vTemplateList.SubItems.Add(43).Text = Format(Int(Trim(dt.Rows(i).Item("InterestStockPercent"))), "##,##0.00")
                            vInterestsAmountRes = Trim(dt.Rows(i).Item("InterestStockAmount"))
                            vTemplateList.SubItems.Add(44).Text = Format(vInterestsAmountRes, "##,##0.00")

                            vTemplateList.SubItems.Add(45).Text = Format(Int(Trim(dt.Rows(i).Item("ProfitPercent"))), "##,##0.00")
                            vProfitAmountRes = Trim(dt.Rows(i).Item("ProfitAmount"))
                            vTemplateList.SubItems.Add(46).Text = Format(vProfitAmountRes, "##,##0.00")

                            vTemplateList.SubItems.Add(47).Text = Trim(dt.Rows(i).Item("TransferValueWord"))
                            vAddCashRes = Trim(dt.Rows(i).Item("TransferValueAmount"))
                            vTemplateList.SubItems.Add(48).Text = Format(vAddCashRes, "##,##0.00")

                            vTemplateList.SubItems.Add(49).Text = Trim(dt.Rows(i).Item("CreditValueWord"))
                            vAddCreditRes = Trim(dt.Rows(i).Item("CreditValueAmount"))
                            vTemplateList.SubItems.Add(50).Text = Format(vAddCreditRes, "##,##0.00")

                            vTemplateList.SubItems.Add(51).Text = Trim(dt.Rows(i).Item("SpecialValueWord"))
                            vDiscount1Res = Trim(dt.Rows(i).Item("SpecialValueAmount"))
                            vTemplateList.SubItems.Add(52).Text = Format(vDiscount1Res, "##,##0.00")

                        Next
                        Me.Cursor = Cursors.Arrow
                        Me.ListView102.Focus()
                    Else
                        MsgBox("ไม่พบข้อมูล Template ที่ต้องการดูข้อมูล", MsgBoxStyle.Information, "send Information")
                    End If
                End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub


    Private Sub TXTSearchItemCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTSearchItemCode.TextChanged

    End Sub

    Private Sub TextBoxDocno_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxDocno.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.DocDate.Focus()
        End If
    End Sub

    Private Sub TextBoxDocno_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBoxDocno.LostFocus
        Me.TextBoxDocno.Text = UCase(Me.TextBoxDocno.Text)
    End Sub

    Private Sub TextBoxDocno_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxDocno.TextChanged
        Dim vCheckDocno As Integer

        On Error GoTo ErrDescription

        vDocno = UCase(Me.TextBoxDocno.Text)
        vQuery = "select isnull(count(docno),0) as vCount from npmaster.dbo.TB_PS_PriceStructure where docno = '" & vDocno & "' "
        vCMD = New SqlCommand(vQuery, vConnection)
        vReadQuery = vCMD.ExecuteReader()
        While vReadQuery.Read
            vCheckDocno = vReadQuery(0)
        End While
        vReadQuery.Close()

        If vCheckDocno = 1 Then
            Call RefreshData(vDocno)
        Else
            'Me.ListView103.Items.Clear()
            vIsCancel = 0
            vIsConfirm = 0
            vIsOpen = 0
            Call GetUserLogIN()
            Me.DocDate.Value = Date.Now

            If vIsCancel = 1 Then
                Me.PB101.Visible = False
                Me.PB102.Visible = False
                Me.PB103.Visible = True
            ElseIf vIsConfirm = 1 Then
                Me.PB101.Visible = False
                Me.PB102.Visible = True
                Me.PB103.Visible = False
            Else
                Me.PB101.Visible = True
                Me.PB102.Visible = False
                Me.PB103.Visible = False
            End If

        End If
        'If Len(Me.TextBoxDocno.Text) > 10 Then
        'If ListView103.Items.Count = 0 Then
        '    Me.TextBoxDocno.Text = Microsoft.VisualBasic.Left(Me.TextBoxDocno.Text, 10)
        '    Me.TextBoxDocno.Focus()
        'End If
        'End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub BTNClearItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearItem.Click
        ClearItemCode()
    End Sub

    Private Sub ListView101_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView101.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.ListView101.Items.Count > 0 Then
                vSelectPriceListIndex = Me.ListView101.SelectedItems(0).Index
                Me.NMFromNumber.Value = Me.ListView101.Items(vSelectPriceListIndex).SubItems(3).Text
                Me.NMToNumber.Value = Me.ListView101.Items(vSelectPriceListIndex).SubItems(4).Text
                Me.TXTPrice1.Text = Format(Int(Me.ListView101.Items(vSelectPriceListIndex).SubItems(5).Text), "##,##0.00")
                Me.TXTPrice2.Text = Format(Int(Me.ListView101.Items(vSelectPriceListIndex).SubItems(6).Text), "##,##0.00")
                If Me.ListView101.Items(vSelectPriceListIndex).SubItems(8).Text <> "" Then
                    Me.DateUpdate.Text = Me.ListView101.Items(vSelectPriceListIndex).SubItems(8).Text
                Else
                    Me.DateUpdate.Text = DateAdd(DateInterval.Day, 1, Date.Now)
                End If
                If Me.ListView101.Items(vSelectPriceListIndex).SubItems(9).Text <> "" Then
                    Me.DateExpire.Text = Me.ListView101.Items(vSelectPriceListIndex).SubItems(9).Text
                Else
                    Me.DateExpire.Text = DateAdd(DateInterval.Day, 3650, Now)
                End If
                If Me.ListView101.Items(vSelectPriceListIndex).SubItems(7).Text = "1" Then
                    Me.CBUpdatePrice.Checked = True
                Else
                    Me.CBUpdatePrice.Checked = False
                End If
                Me.GBPriceList.Visible = True
                Me.GBPriceList.BringToFront()
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub ListView101_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView101.SelectedIndexChanged

    End Sub

    Private Sub TXTSearchDocno_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTSearchDocno.DoubleClick

    End Sub

    Private Sub TXTSearchDocno_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTSearchDocno.KeyDown
        Dim vSearch As String
        Dim i As Integer
        Dim vListDocno As ListViewItem

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            vSearch = Trim(Me.TXTSearchDocno.Text)
            If TXTSearchDocno.Text <> "" Then
                vQuery = "exec dbo.USP_PS_SearchPriceStructureDocno '" & vSearch & "' "
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "Search")
                dt = ds.Tables("Search")

                Me.ListViewDocno.Items.Clear()
                If dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        vListDocno = Me.ListViewDocno.Items.Add(i + 1)
                        vListDocno.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("Docno"))
                        vListDocno.SubItems.Add(2).Text = Trim(dt.Rows(i).Item("Docdate"))
                        vListDocno.SubItems.Add(3).Text = Trim(dt.Rows(i).Item("creatorcode"))
                        vListDocno.SubItems.Add(4).Text = Trim(dt.Rows(i).Item("createdate"))
                        vListDocno.SubItems.Add(5).Text = Trim(dt.Rows(i).Item("isconfirm"))
                        vListDocno.SubItems.Add(6).Text = Trim(dt.Rows(i).Item("iscancel"))

                        If Trim(dt.Rows(i).Item("iscancel")) = 1 Then
                            Me.ListViewDocno.Items(i).ForeColor = Color.Red
                        End If
                        If Trim(dt.Rows(i).Item("isconfirm")) = 1 And Trim(dt.Rows(i).Item("iscancel")) = 0 Then
                            Me.ListViewDocno.Items(i).ForeColor = Color.Green
                        End If
                    Next
                End If
            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTSearchDocno_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTSearchDocno.TextChanged

    End Sub

    Private Sub BTNNoSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNNoSelect.Click
        Me.GBSearchDocno.Visible = False
    End Sub

    Private Sub BTNSeacrhDocno_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSeacrhDocno.Click
        On Error Resume Next

        Me.GBSearchDocno.BringToFront()
        Me.GBSearchDocno.Visible = True
        Me.TXTSearchDocno.Text = ""
        Me.ListViewDocno.Items.Clear()
        Me.TXTSearchDocno.Focus()
    End Sub

    Private Sub ListViewDocno_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewDocno.DoubleClick
        Dim vDocno As String
        Dim vIndex As Integer

        On Error GoTo ErrDescription

        If Me.ListViewDocno.Items.Count > 0 Then
            vIndex = Me.ListViewDocno.SelectedItems(0).Index
            vDocno = UCase(Me.ListViewDocno.Items(vIndex).SubItems(1).Text)

            Call RefreshData(vDocno)
            Me.GBSearchDocno.Visible = False
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub ListViewDocno_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewDocno.SelectedIndexChanged

    End Sub

    Private Sub ListView103_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView103.DoubleClick
        Dim vIndex As Integer
        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vBillDiscRes As Double
        Dim vFollowDisc1Res As Double
        Dim vFollowDisc2Res As Double
        Dim vFollowDisc3Res As Double
        Dim vRebateRes As Double
        Dim vSpecialDiscRes As Double
        Dim vMissProfitRes As Double
        Dim vSendRes As Double
        Dim vCustSendRes As Double
        Dim vAdvertiseRes As Double
        Dim vMarketRes As Double
        Dim vTaxRes As Double
        Dim vInstallRes As Double
        Dim vServiceRes As Double
        Dim vPointRes As Double
        Dim vTargetRes As Double
        Dim vGiftRes As Double
        Dim vCommissionRes As Double
        Dim vBegProfitAmountRes As Double
        Dim vInterestsAmountRes As Double
        Dim vProfitAmountRes As Double
        Dim vAddCashRes As Double
        Dim vAddCreditRes As Double
        Dim vDiscount1Res As Double
        Dim vDiscMemberRes As Double

        Dim vItemCode As String
        Dim vListPriceList As ListViewItem
        Dim i As Integer

        On Error GoTo ErrDescription

        If Me.ListView103.Items.Count > 0 Then
            Me.Cursor = Cursors.WaitCursor
            vIndex = Me.ListView103.SelectedItems(0).Index
            Me.TXTItemCode.Text = Trim(Me.ListView103.Items(vIndex).SubItems(0).Text)
            Me.LBLItemName.Text = Trim(Me.ListView103.Items(vIndex).SubItems(1).Text)
            Me.CMBUnit.Items.Clear()
            Me.CMBUnit.Items.Add(Me.ListView103.Items(vIndex).SubItems(2).Text)
            Me.CMBUnit.Text = Me.CMBUnit.Items.Item(0)

            vDO = Me.ListView103.Items(vIndex).SubItems(3).Text
            Me.TXTDO.Text = Format(vDO, "##,##0.00")

            vPriceSet = Me.ListView103.Items(vIndex).SubItems(4).Text
            Me.TXTPriceSet.Text = Format(vPriceSet, "##,##0.00")

            Me.TXTBillDisc.Text = Me.ListView103.Items(vIndex).SubItems(5).Text
            vBillDiscRes = Me.ListView103.Items(vIndex).SubItems(6).Text
            Me.TXTBillDiscRes.Text = Format(vBillDiscRes, "##,##0.00")

            Me.TXTFollowDisc1.Text = Me.ListView103.Items(vIndex).SubItems(7).Text
            vFollowDisc1Res = Me.ListView103.Items(vIndex).SubItems(8).Text
            Me.TXTFollowDisc1Res.Text = Format(vFollowDisc1Res, "##,##0.00")

            Me.TXTFollowDisc2.Text = Me.ListView103.Items(vIndex).SubItems(9).Text
            vFollowDisc2Res = Me.ListView103.Items(vIndex).SubItems(10).Text
            Me.TXTFollowDisc2Res.Text = Format(vFollowDisc2Res, "##,##0.00")

            Me.TXTFollowDisc3.Text = Me.ListView103.Items(vIndex).SubItems(11).Text
            vFollowDisc3Res = Me.ListView103.Items(vIndex).SubItems(12).Text
            Me.TXTFollowDisc3Res.Text = Format(vFollowDisc3Res, "##,##0.00")

            Me.TXTRebate.Text = Me.ListView103.Items(vIndex).SubItems(13).Text
            vRebateRes = Me.ListView103.Items(vIndex).SubItems(14).Text
            Me.TXTRebateRes.Text = Format(vRebateRes, "##,##0.00")

            Me.TXTSpecialDisc.Text = Me.ListView103.Items(vIndex).SubItems(15).Text
            vSpecialDiscRes = Me.ListView103.Items(vIndex).SubItems(16).Text
            Me.TXTSpecialDiscRes.Text = Format(vSpecialDiscRes, "##,##0.00")

            Me.TXTMissProfit.Text = Me.ListView103.Items(vIndex).SubItems(17).Text
            vMissProfitRes = Me.ListView103.Items(vIndex).SubItems(18).Text
            Me.TXTMissProfitRes.Text = Format(vMissProfitRes, "##,##0.00")

            Me.TXTSend.Text = Me.ListView103.Items(vIndex).SubItems(19).Text
            vSendRes = Me.ListView103.Items(vIndex).SubItems(20).Text
            Me.TXTSendRes.Text = Format(vSendRes, "##,##0.00")

            Me.TXTCustSend.Text = Me.ListView103.Items(vIndex).SubItems(21).Text
            vCustSendRes = Me.ListView103.Items(vIndex).SubItems(22).Text
            Me.TXTCustSendRes.Text = Format(vCustSendRes, "##,##0.00")

            Me.TXTAdvertise.Text = Me.ListView103.Items(vIndex).SubItems(23).Text
            vAdvertiseRes = Me.ListView103.Items(vIndex).SubItems(24).Text
            Me.TXTAdvertiseRes.Text = Format(vAdvertiseRes, "##,##0.00")

            Me.TXTMarket.Text = Me.ListView103.Items(vIndex).SubItems(25).Text
            vMarketRes = Me.ListView103.Items(vIndex).SubItems(26).Text
            Me.TXTMarketRes.Text = Format(vMarketRes, "##,##0.00")

            Me.TXTTax.Text = Me.ListView103.Items(vIndex).SubItems(27).Text
            vTaxRes = Me.ListView103.Items(vIndex).SubItems(28).Text
            Me.TXTTaxRes.Text = Format(vTaxRes, "##,##0.00")

            Me.TXTInstall.Text = Me.ListView103.Items(vIndex).SubItems(29).Text
            vInstallRes = Me.ListView103.Items(vIndex).SubItems(30).Text
            Me.TXTInstallRes.Text = Format(vInstallRes, "##,##0.00")

            Me.TXTService.Text = Me.ListView103.Items(vIndex).SubItems(31).Text
            vServiceRes = Me.ListView103.Items(vIndex).SubItems(32).Text
            Me.TXTServiceRes.Text = Format(vServiceRes, "##,##0.00")

            Me.TXTPoint.Text = Me.ListView103.Items(vIndex).SubItems(33).Text
            vPointRes = Me.ListView103.Items(vIndex).SubItems(34).Text
            Me.TXTPointRes.Text = Format(vPointRes, "##,##0.00")

            Me.TXTDiscMember.Text = Me.ListView103.Items(vIndex).SubItems(35).Text
            vDiscMemberRes = Me.ListView103.Items(vIndex).SubItems(36).Text
            Me.TXTDiscMemberRes.Text = Format(vDiscMemberRes, "##,##0.00")

            Me.TXTTarget.Text = Me.ListView103.Items(vIndex).SubItems(37).Text
            vTargetRes = Me.ListView103.Items(vIndex).SubItems(38).Text
            Me.TXTTargetRes.Text = Format(vTargetRes, "##,##0.00")

            Me.TXTGift.Text = Me.ListView103.Items(vIndex).SubItems(39).Text
            vGiftRes = Me.ListView103.Items(vIndex).SubItems(40).Text
            Me.TXTGiftRes.Text = Format(vGiftRes, "##,##0.00")

            Me.TXTCommission.Text = Me.ListView103.Items(vIndex).SubItems(41).Text
            vCommissionRes = Me.ListView103.Items(vIndex).SubItems(42).Text
            Me.TXTCommissionRes.Text = Format(vCommissionRes, "##,##0.00")

            Me.TXTBegProfit.Text = Me.ListView103.Items(vIndex).SubItems(43).Text
            vBegProfitAmountRes = Me.ListView103.Items(vIndex).SubItems(44).Text
            Me.TXTBegProfitAmount.Text = Format(vBegProfitAmountRes, "##,##0.00")

            Me.TXTInterests.Text = Me.ListView103.Items(vIndex).SubItems(45).Text
            vInterestsAmountRes = Me.ListView103.Items(vIndex).SubItems(46).Text
            Me.TXTInterestsAmount.Text = Format(vInterestsAmountRes, "##,##0.00")

            Me.TXTProfit.Text = Format(Me.ListView103.Items(vIndex).SubItems(47).Text)
            vProfitAmountRes = Me.ListView103.Items(vIndex).SubItems(48).Text
            Me.TXTProfitAmount.Text = Format(vProfitAmountRes, "##,##0.00")

            Me.TXTAddCash.Text = Me.ListView103.Items(vIndex).SubItems(49).Text
            vAddCashRes = Me.ListView103.Items(vIndex).SubItems(50).Text
            Me.TXTAddCashRes.Text = Format(vAddCashRes, "##,##0.00")

            Me.TXTAddCredit.Text = Me.ListView103.Items(vIndex).SubItems(51).Text
            vAddCreditRes = Me.ListView103.Items(vIndex).SubItems(52).Text
            Me.TXTAddCreditRes.Text = Format(vAddCreditRes, "##,##0.00")

            Me.TXTDiscount1.Text = Me.ListView103.Items(vIndex).SubItems(53).Text
            vDiscount1Res = Me.ListView103.Items(vIndex).SubItems(54).Text
            Me.TXTDiscount1Res.Text = Format(vDiscount1Res, "##,##0.00")


            vItemCode = Trim(Me.ListView103.Items(vIndex).SubItems(0).Text)

            vQuery = "exec dbo.USP_PS_PriceList '" & vDocno & "','" & vItemCode & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "PriceList")
            dt = ds.Tables("PriceList")

            Me.ListView101.Items.Clear()
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    vListPriceList = Me.ListView101.Items.Add(dt.Rows(i).Item("saleunitcode"))
                    If dt.Rows(i).Item("saletype") = 0 Then
                        vListPriceList.SubItems.Add(1).Text = "ขายเงินสด"
                    Else
                        vListPriceList.SubItems.Add(1).Text = "ขายเงินเชื่อ"
                    End If
                    If dt.Rows(i).Item("transporttype") = 0 Then
                        vListPriceList.SubItems.Add(2).Text = "รับเอง"
                    Else
                        vListPriceList.SubItems.Add(2).Text = "ส่งให้"
                    End If
                    vListPriceList.SubItems.Add(3).Text = Format(dt.Rows(i).Item("fromqty"), "##,##0.00")
                    vListPriceList.SubItems.Add(4).Text = Format(dt.Rows(i).Item("toqty"), "##,##0.00")
                    vListPriceList.SubItems.Add(5).Text = Format(dt.Rows(i).Item("priceset"), "##,##0.00")
                    vListPriceList.SubItems.Add(6).Text = Format(dt.Rows(i).Item("priceset2"), "##,##0.00")
                    vListPriceList.SubItems.Add(7).Text = dt.Rows(i).Item("ispriceupdate")
                    vListPriceList.SubItems.Add(8).Text = dt.Rows(i).Item("toupdatedate")
                    vListPriceList.SubItems.Add(9).Text = dt.Rows(i).Item("stopdate")
                    vListPriceList.SubItems.Add(10).Text = dt.Rows(i).Item("isupdate")
                    vListPriceList.SubItems.Add(11).Text = dt.Rows(i).Item("isprintlabel")
                    If IsDBNull(dt.Rows(i).Item("printdatetime")) Then
                        vListPriceList.SubItems.Add(12).Text = ""
                    Else
                        vListPriceList.SubItems.Add(12).Text = dt.Rows(i).Item("printdatetime")
                    End If

                    vListPriceList.SubItems.Add(13).Text = dt.Rows(i).Item("itemcode")

                Next
            End If


            Me.Cursor = Cursors.Arrow
            Me.TXTDO.Focus()
        End If


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub


    Private Sub BTNConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNConfirm.Click
        Dim vDocNo As String

        On Error GoTo ErrDescription

        If vIsOpen = 1 Then
            vDocNo = Trim(Me.TextBoxDocno.Text)
            If vIsCancel = 0 Then
                Call ChekAuthorityAccess()

                If vDepartment = "MG" And vLevelID = 0 Then

                    If MessageBox.Show("คุณต้องการอนุมัติเอกสาร ใช่หรือไม่", "ข้อความสอบถาม", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                        If vIsConfirm = 1 Then
                            MsgBox("ไม่สามารถอนุมัติเลขที่เอกสาร " & vDocNo & " ได้เนื่องจากได้ถูกอนุมัติเรียบร้อยแล้ว กรุณาตรวจสอบเลขที่เอกสารด้วย ", MsgBoxStyle.Critical, "Send Error")
                            Exit Sub
                        Else

                            vQuery = "exec dbo.USP_PS_PriceStructureConfirm '" & vDocNo & "' "
                            vCMD = New SqlCommand(vQuery, vConnection)
                            vCMD.ExecuteNonQuery()
                            Call ClearDocument()
                            Call ClearItemData()

                            MsgBox("อนุมัติเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information")

                        End If
                    End If
                Else
                    MsgBox("" & vUserID & " ไม่มีสิทธิ์ในการอนุมัติเอกสารโครงสร้างราคา", MsgBoxStyle.Critical, "Send Error Message")
                End If
            Else
                MsgBox("เอกสารเลขที่ " & vDocNo & " ถูกยกเลิกไปแล้วไม่สามารถอนุมัติได้  กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
            End If
        End If


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub BTNBasketCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNBasketCancel.Click
        Me.GBPriceList.Visible = False
    End Sub

    Private Sub TXTInterests_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTInterests.TextChanged
        Dim vCheckInterests As Double

        On Error GoTo ErrDescription

        If Me.TXTInterests.Text <> "" Then
            vCheckInterests = Me.TXTInterests.Text
            If vCheckInterests < 0 Then
                Me.TXTInterests.ForeColor = Color.Red
            Else
                Me.TXTInterests.ForeColor = Color.Black
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTProfit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTProfit.TextChanged
        Dim vCheckProfit As Double

        On Error GoTo ErrDescription

        If Me.TXTProfit.Text <> "" And Me.TXTBegProfit.Text <> "" Then
            vCheckProfit = Me.TXTBegProfit.Text
            If vCheckProfit < 0 Then
                Me.TXTProfit.ForeColor = Color.Red
            Else
                Me.TXTProfit.ForeColor = Color.Black
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTDiscMember_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TXTDiscMember.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TXTTarget.Focus()
        End If
        If Asc(e.KeyCode) = 51 Then
            Me.TXTPoint.Focus()
        ElseIf Asc(e.KeyCode) = 52 Then
            Me.TXTTarget.Focus()
        End If
    End Sub

    Private Sub TXTDiscMember_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTDiscMember.LostFocus
        Dim vCheckDiscMember As Double

        On Error GoTo ErrDescription

        If InStr(Me.TXTDiscMember.Text, "%") = 0 Then
            vCheckDiscMember = Me.TXTDiscMember.Text
            Me.TXTDiscMember.Text = Format(vCheckDiscMember, "##,##0.00")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTDiscMember_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTDiscMember.TextChanged
        Call Me.TXTDiscMemberCalc()
    End Sub

    Private Sub TXTDiscMemberRes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTDiscMemberRes.TextChanged
        Call Me.TXTTargetCalc()
    End Sub

    Private Sub BTNCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCancel.Click
        Dim vDocNo As String

        On Error GoTo ErrDescription

        If vIsOpen = 1 Then
            vDocNo = Trim(Me.TextBoxDocno.Text)
            If vIsConfirm = 0 Then
                Call ChekAuthorityAccess()

                If vCheckUserID = vCheckBuyer Then

                    If MessageBox.Show("คุณต้องการยกเลิกเอกสาร ใช่หรือไม่", "ข้อความสอบถาม", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                        If vIsCancel = 1 Then
                            MsgBox("ไม่สามารถยกเลิกเลขที่เอกสาร " & vDocNo & " ได้เนื่องจากได้ถูกยกเลิกเรียบร้อยแล้ว กรุณาตรวจสอบเลขที่เอกสารด้วย ", MsgBoxStyle.Critical, "Send Error")
                            Exit Sub
                        Else

                            vQuery = "exec dbo.USP_PS_CancelPriceStructure '" & vDocNo & "' "
                            vCMD = New SqlCommand(vQuery, vConnection)
                            vCMD.ExecuteNonQuery()
                            Call ClearDocument()
                            Call ClearItemData()

                            MsgBox("ยกเลิกเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information")

                        End If
                    End If
                Else
                    MsgBox("" & vUserID & " ไม่มีสิทธิ์ในการยกเลิกเอกสารโครงสร้างราคา เพราะผู้ที่ยกเลิกได้ต้องเป็นเจ้าของเอกสารเท่านั้น", MsgBoxStyle.Critical, "Send Error Message")
                End If
            Else
                MsgBox("เอกสารเลขที่ " & vDocNo & " ถูกอนุมัติไปแล้วไม่สามารถยกเลิกได้  กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
            End If
        End If


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub ListView103_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView103.KeyDown
        Dim vIndex As Integer
        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vBillDiscRes As Double
        Dim vFollowDisc1Res As Double
        Dim vFollowDisc2Res As Double
        Dim vFollowDisc3Res As Double
        Dim vRebateRes As Double
        Dim vSpecialDiscRes As Double
        Dim vMissProfitRes As Double
        Dim vSendRes As Double
        Dim vCustSendRes As Double
        Dim vAdvertiseRes As Double
        Dim vMarketRes As Double
        Dim vTaxRes As Double
        Dim vInstallRes As Double
        Dim vServiceRes As Double
        Dim vPointRes As Double
        Dim vTargetRes As Double
        Dim vGiftRes As Double
        Dim vCommissionRes As Double
        Dim vBegProfitAmountRes As Double
        Dim vInterestsAmountRes As Double
        Dim vProfitAmountRes As Double
        Dim vAddCashRes As Double
        Dim vAddCreditRes As Double
        Dim vDiscount1Res As Double
        Dim vDiscMemberRes As Double

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.ListView103.Items.Count > 0 Then
                Me.Cursor = Cursors.WaitCursor
                vIndex = Me.ListView103.SelectedItems(0).Index
                Me.TXTItemCode.Text = Trim(Me.ListView103.Items(vIndex).SubItems(0).Text)
                Me.LBLItemName.Text = Trim(Me.ListView103.Items(vIndex).SubItems(1).Text)
                Me.CMBUnit.Items.Clear()
                Me.CMBUnit.Items.Add(Me.ListView103.Items(vIndex).SubItems(2).Text)
                Me.CMBUnit.Text = Me.CMBUnit.Items.Item(0)

                vDO = Me.ListView103.Items(vIndex).SubItems(3).Text
                Me.TXTDO.Text = Format(vDO, "##,##0.00")

                vPriceSet = Me.ListView103.Items(vIndex).SubItems(4).Text
                Me.TXTPriceSet.Text = Format(vPriceSet, "##,##0.00")

                Me.TXTBillDisc.Text = Me.ListView103.Items(vIndex).SubItems(5).Text
                vBillDiscRes = Me.ListView103.Items(vIndex).SubItems(6).Text
                Me.TXTBillDiscRes.Text = Format(vBillDiscRes, "##,##0.00")

                Me.TXTFollowDisc1.Text = Me.ListView103.Items(vIndex).SubItems(7).Text
                vFollowDisc1Res = Me.ListView103.Items(vIndex).SubItems(8).Text
                Me.TXTFollowDisc1Res.Text = Format(vFollowDisc1Res, "##,##0.00")

                Me.TXTFollowDisc2.Text = Me.ListView103.Items(vIndex).SubItems(9).Text
                vFollowDisc2Res = Me.ListView103.Items(vIndex).SubItems(10).Text
                Me.TXTFollowDisc2Res.Text = Format(vFollowDisc2Res, "##,##0.00")

                Me.TXTFollowDisc3.Text = Me.ListView103.Items(vIndex).SubItems(11).Text
                vFollowDisc3Res = Me.ListView103.Items(vIndex).SubItems(12).Text
                Me.TXTFollowDisc3Res.Text = Format(vFollowDisc3Res, "##,##0.00")

                Me.TXTRebate.Text = Me.ListView103.Items(vIndex).SubItems(13).Text
                vRebateRes = Me.ListView103.Items(vIndex).SubItems(14).Text
                Me.TXTRebateRes.Text = Format(vRebateRes, "##,##0.00")

                Me.TXTSpecialDisc.Text = Me.ListView103.Items(vIndex).SubItems(15).Text
                vSpecialDiscRes = Me.ListView103.Items(vIndex).SubItems(16).Text
                Me.TXTSpecialDiscRes.Text = Format(vSpecialDiscRes, "##,##0.00")

                Me.TXTMissProfit.Text = Me.ListView103.Items(vIndex).SubItems(17).Text
                vMissProfitRes = Me.ListView103.Items(vIndex).SubItems(18).Text
                Me.TXTMissProfitRes.Text = Format(vMissProfitRes, "##,##0.00")

                Me.TXTSend.Text = Me.ListView103.Items(vIndex).SubItems(19).Text
                vSendRes = Me.ListView103.Items(vIndex).SubItems(20).Text
                Me.TXTSendRes.Text = Format(vSendRes, "##,##0.00")

                Me.TXTCustSend.Text = Me.ListView103.Items(vIndex).SubItems(21).Text
                vCustSendRes = Me.ListView103.Items(vIndex).SubItems(22).Text
                Me.TXTCustSendRes.Text = Format(vCustSendRes, "##,##0.00")

                Me.TXTAdvertise.Text = Me.ListView103.Items(vIndex).SubItems(23).Text
                vAdvertiseRes = Me.ListView103.Items(vIndex).SubItems(24).Text
                Me.TXTAdvertiseRes.Text = Format(vAdvertiseRes, "##,##0.00")

                Me.TXTMarket.Text = Me.ListView103.Items(vIndex).SubItems(25).Text
                vMarketRes = Me.ListView103.Items(vIndex).SubItems(26).Text
                Me.TXTMarketRes.Text = Format(vMarketRes, "##,##0.00")

                Me.TXTTax.Text = Me.ListView103.Items(vIndex).SubItems(27).Text
                vTaxRes = Me.ListView103.Items(vIndex).SubItems(28).Text
                Me.TXTTaxRes.Text = Format(vTaxRes, "##,##0.00")

                Me.TXTInstall.Text = Me.ListView103.Items(vIndex).SubItems(29).Text
                vInstallRes = Me.ListView103.Items(vIndex).SubItems(30).Text
                Me.TXTInstallRes.Text = Format(vInstallRes, "##,##0.00")

                Me.TXTService.Text = Me.ListView103.Items(vIndex).SubItems(31).Text
                vServiceRes = Me.ListView103.Items(vIndex).SubItems(32).Text
                Me.TXTServiceRes.Text = Format(vServiceRes, "##,##0.00")

                Me.TXTPoint.Text = Me.ListView103.Items(vIndex).SubItems(33).Text
                vPointRes = Me.ListView103.Items(vIndex).SubItems(34).Text
                Me.TXTPointRes.Text = Format(vPointRes, "##,##0.00")

                Me.TXTDiscMember.Text = Me.ListView103.Items(vIndex).SubItems(35).Text
                vDiscMemberRes = Me.ListView103.Items(vIndex).SubItems(36).Text
                Me.TXTDiscMemberRes.Text = Format(vDiscMemberRes, "##,##0.00")

                Me.TXTTarget.Text = Me.ListView103.Items(vIndex).SubItems(37).Text
                vTargetRes = Me.ListView103.Items(vIndex).SubItems(38).Text
                Me.TXTTargetRes.Text = Format(vTargetRes, "##,##0.00")

                Me.TXTGift.Text = Me.ListView103.Items(vIndex).SubItems(39).Text
                vGiftRes = Me.ListView103.Items(vIndex).SubItems(40).Text
                Me.TXTGiftRes.Text = Format(vGiftRes, "##,##0.00")

                Me.TXTCommission.Text = Me.ListView103.Items(vIndex).SubItems(41).Text
                vCommissionRes = Me.ListView103.Items(vIndex).SubItems(42).Text
                Me.TXTCommissionRes.Text = Format(vCommissionRes, "##,##0.00")

                Me.TXTBegProfit.Text = Me.ListView103.Items(vIndex).SubItems(43).Text
                vBegProfitAmountRes = Me.ListView103.Items(vIndex).SubItems(44).Text
                Me.TXTBegProfitAmount.Text = Format(vBegProfitAmountRes, "##,##0.00")

                Me.TXTInterests.Text = Me.ListView103.Items(vIndex).SubItems(45).Text
                vInterestsAmountRes = Me.ListView103.Items(vIndex).SubItems(46).Text
                Me.TXTInterestsAmount.Text = Format(vInterestsAmountRes, "##,##0.00")

                Me.TXTProfit.Text = Format(Me.ListView103.Items(vIndex).SubItems(47).Text)
                vProfitAmountRes = Me.ListView103.Items(vIndex).SubItems(48).Text
                Me.TXTProfitAmount.Text = Format(vProfitAmountRes, "##,##0.00")

                Me.TXTAddCash.Text = Me.ListView103.Items(vIndex).SubItems(49).Text
                vAddCashRes = Me.ListView103.Items(vIndex).SubItems(50).Text
                Me.TXTAddCashRes.Text = Format(vAddCashRes, "##,##0.00")

                Me.TXTAddCredit.Text = Me.ListView103.Items(vIndex).SubItems(51).Text
                vAddCreditRes = Me.ListView103.Items(vIndex).SubItems(52).Text
                Me.TXTAddCreditRes.Text = Format(vAddCreditRes, "##,##0.00")

                Me.TXTDiscount1.Text = Me.ListView103.Items(vIndex).SubItems(53).Text
                vDiscount1Res = Me.ListView103.Items(vIndex).SubItems(54).Text
                Me.TXTDiscount1Res.Text = Format(vDiscount1Res, "##,##0.00")

                Me.Cursor = Cursors.Arrow
                Me.TXTDO.Focus()
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub ListView103_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView103.SelectedIndexChanged

    End Sub

    Private Sub BTNClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClear.Click
        On Error GoTo ErrDescription

        Me.CMBUnit.Items.Clear()
        Me.LBLItemName.Text = ""
        Me.TXTDO.Text = "0"
        Me.TXTPriceSet.Text = "0"
        Me.TXTBillDisc.Text = "0"
        Me.TXTBillDiscRes.Text = ""
        Me.TXTFollowDisc1.Text = "0"
        Me.TXTFollowDisc1Res.Text = ""
        Me.TXTFollowDisc2.Text = "0"
        Me.TXTFollowDisc2Res.Text = ""
        Me.TXTFollowDisc3.Text = "0"
        Me.TXTFollowDisc3Res.Text = ""
        Me.TXTRebate.Text = "0"
        Me.TXTRebateRes.Text = ""
        Me.TXTSpecialDisc.Text = "0"
        Me.TXTSpecialDiscRes.Text = ""
        Me.TXTMissProfit.Text = "0"
        Me.TXTMissProfitRes.Text = ""
        Me.TXTSend.Text = "0"
        Me.TXTSendRes.Text = ""
        Me.TXTCustSend.Text = "0"
        Me.TXTCustSendRes.Text = ""
        Me.TXTAdvertise.Text = "0"
        Me.TXTAdvertiseRes.Text = ""
        Me.TXTMarket.Text = "0"
        Me.TXTMarketRes.Text = ""
        Me.TXTTax.Text = "0"
        Me.TXTTaxRes.Text = ""
        Me.TXTInstall.Text = "0"
        Me.TXTInstallRes.Text = ""
        Me.TXTService.Text = "0"
        Me.TXTServiceRes.Text = ""
        Me.TXTDiscMember.Text = "0"
        Me.TXTDiscMemberRes.Text = ""
        Me.TXTPoint.Text = "0"
        Me.TXTPointRes.Text = ""
        Me.TXTTarget.Text = "0"
        Me.TXTTargetRes.Text = ""
        Me.TXTGift.Text = "0"
        Me.TXTGiftRes.Text = ""
        Me.TXTCommission.Text = "0"
        Me.TXTCommissionRes.Text = ""
        Me.TXTBegProfit.Text = "0"
        Me.TXTBegProfitAmount.Text = ""
        Me.TXTInterests.Text = "0"
        Me.TXTInterestsAmount.Text = ""
        Me.TXTProfit.Text = "0"
        Me.TXTProfitAmount.Text = ""
        Me.TXTAddCash.Text = "0"
        Me.TXTAddCashRes.Text = "0"
        Me.TXTAddCredit.Text = "0"
        Me.TXTAddCreditRes.Text = "0"
        Me.TXTDiscount1.Text = "0"
        Me.TXTDiscount1Res.Text = "0"

        Me.TXTItemCode.Text = ""
        Me.LBLItemName.Text = ""
        Me.TXTDO.Text = "0.00"
        Me.TXTPriceSet.Text = "0.00"
        Me.CMBUnit.Items.Clear()
        Me.TXTSearchItemCode.Text = ""
        Me.ListView101.Items.Clear()
        Me.CMBMultiUnitCode.Items.Clear()
        Me.TextBoxDocno.Text = ""
        Me.ListView103.Items.Clear()
        Me.ListView102.Items.Clear()
        vIsOpen = 0
        vIsCancel = 0
        vIsConfirm = 0
        Me.TextBoxDocno.Focus()
        Me.DocDate.Value = Date.Now


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TXTPrice1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTPrice1.LostFocus
        Dim vCheckCountDot As Integer

        vCheckCountDot = CheckDot(TXTPrice1.Text)
        If vCheckCountDot > 1 Then
            Me.TXTPrice1.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub TXTPrice1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTPrice1.TextChanged
        Dim vCheckCountDot As Integer

        vCheckCountDot = CheckDot(TXTPrice1.Text)
        If vCheckCountDot > 1 Then
            MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
            Me.TXTPrice1.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub TXTPrice2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TXTPrice2.LostFocus
        Dim vCheckCountDot As Integer

        vCheckCountDot = CheckDot(TXTPrice2.Text)
        If vCheckCountDot > 1 Then
            Me.TXTPrice2.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub TXTPrice2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TXTPrice2.TextChanged
        Dim vCheckCountDot As Integer

        vCheckCountDot = CheckDot(TXTPrice2.Text)
        If vCheckCountDot > 1 Then
            MsgBox("ไวยกรณ์ผิดครับ กรุณาแก้ไขก่อน", MsgBoxStyle.Critical, "Send Error")
            Me.TXTPrice2.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub Label61_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label61.Click

    End Sub
End Class