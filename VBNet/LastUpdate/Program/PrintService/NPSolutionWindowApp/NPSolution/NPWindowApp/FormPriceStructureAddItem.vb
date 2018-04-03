Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports CrystalDecisions
Imports System
Imports Microsoft

Public Class FormPriceStructureAddItem

    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable

    Dim ds1 As DataSet
    Dim da1 As SqlDataAdapter
    Dim dt1 As DataTable

    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vIsOpen As Integer
    Dim vMemIsCancel As Integer
    Dim vMemIsConfirm As Integer

    Dim vIsNumber As Integer

    Dim vMemColumn As Integer
    Dim vMemRow As Integer

    Dim vMemStartDate As Date

    Private Sub FormPriceStructureAddItem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()
        Call vGetBeginData()
        Call SearchItemBrand()
        Call NewDoc()
        Call vGendocNoAuto()
    End Sub

    Public Sub NewDoc()
        Me.PBNew.Visible = True
        Me.PBCancel.Visible = False
        Me.PBConfirm.Visible = False
    End Sub

    Public Sub ConfirmDoc()
        Me.PBNew.Visible = False
        Me.PBCancel.Visible = False
        Me.PBConfirm.Visible = True
    End Sub

    Public Sub CancelDoc()
        Me.PBNew.Visible = False
        Me.PBCancel.Visible = True
        Me.PBConfirm.Visible = False
    End Sub

    Public Sub vGetBeginData()
        Dim i As Integer
        Dim n As Integer

        'On Error Resume Next

        Me.DGVItemDetails.Rows.Add(9999)
        For i = 0 To 9999 - 1
            n = n + 1
            Me.DGVItemDetails.Item(0, i).Value = n
        Next

        Me.DGVItemDetails.CurrentCell = Me.DGVItemDetails.Item(1, 0)
    End Sub

    Public Sub vGenDocNoAuto()
        Dim vNow As Date

        'On Error Resume Next

        vQuery = "set dateformat dmy"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vNow = Now.Day & "/" & Now.Month & "/" & Now.Year
        vQuery = "select dbo.ft_com_newrequest ('" & vNow & "') as docno"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchDocno")
        dt = ds.Tables("SearchDocno")
        If dt.Rows.Count > 0 Then
            Me.TBDocNo.Text = dt.Rows(0).Item("docno")
        Else
            Me.TBDocNo.Text = ""
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Public Sub SearchItemBrand()
        Dim i As Integer

        vQuery = "exec dbo.USP_PS_BrandList"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchBrand")
        dt = ds.Tables("SearchBrand")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                Me.CMBBrandCode.Items.Add(dt.Rows(i).Item("brandname"))
            Next
        End If
    End Sub

    Private Sub DGVItemDetails_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVItemDetails.CellBeginEdit
        Dim vRow As Integer
        Dim vLine As Integer

        'On Error Resume Next

        vRow = Me.DGVItemDetails.CurrentCell.RowIndex
        vLine = Me.DGVItemDetails.Item(0, vRow).Value

        If vLine = 0 Then
            Me.DGVItemDetails.Columns(0).ReadOnly = False
            Me.DGVItemDetails.Item(0, vRow).Value = vRow + 1
        End If
        Me.DGVItemDetails.Columns(0).ReadOnly = True
    End Sub

    Private Sub DGVItemDetails_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItemDetails.CellEndEdit
        Dim vCampaignCode As String
        Dim vDocNo As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vColumn As Integer
        Dim vRow As Integer
        Dim i As Integer

        Dim vCheckItemCode As String
        Dim vMemCountCheck As Integer
        Dim vCheckItemDup As Integer

        Dim vCheckNoDup As String
        Dim vCheckCampaign As String
        Dim vCheckCampaignName As String

        Dim vDateDiff As Integer
        Dim vNowDate As Date

        Dim vDOAmount As Double
        Dim vBillDisc As Integer
        Dim vBillDiscAmount As Double
        Dim vAccCost As Double
        Dim vDisc1 As Integer
        Dim vDiscAmount1 As Double
        Dim vAfterDiscAmount1 As Double
        Dim vUpdateDate As String
        Dim vMyDescription As String

        On Error Resume Next

        If Me.TBDocNo.Text = "" Then
            MsgBox("กรุณา กรอกเลขที่เอกสาร ก่อนเลือกสินค้า", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBDocNo.Focus()
            Exit Sub
        End If


        vNowDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        vDateDiff = vb6.DateDiff(DateInterval.Day, vMemStartDate, vNowDate)
        vUpdateDate = vb6.DateDiff(DateInterval.Day, vMemStartDate, vNowDate)

        vDocNo = Me.TBDocNo.Text

        vColumn = Me.DGVItemDetails.CurrentCell.ColumnIndex
        vRow = Me.DGVItemDetails.CurrentCell.RowIndex
        vItemCode = Me.DGVItemDetails.CurrentCell.Value

        If vColumn = 1 Then
            If vItemCode <> "" Then
                For i = 0 To Me.DGVItemDetails.Rows.Count - 1
                    vCheckItemCode = Me.DGVItemDetails.Item(1, i).Value

                    If vCheckItemCode = vItemCode Then
                        vMemCountCheck = vMemCountCheck + 1
                    End If
                Next

                If vMemCountCheck > 1 Then
                    MsgBox("สินค้า รหัส " & vItemCode & " มีอยู่แล้วในรายการเสนอขอคิดค่าคอมฯ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(vColumn, vRow).Value = ""
                    Exit Sub
                End If

                vQuery = "exec dbo.usp_np_searchitemdescription '" & vItemCode & "' "
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "CheckItem")
                dt = ds.Tables("CheckItem")
                If dt.Rows.Count > 0 Then

                    vUnitCode = dt.Rows(0).Item("unitcode")

                    'vQuery = "exec dbo.USP_COM_RequestDupCK '" & vDocNo & "','" & vCampaignCode & "','" & vItemCode & "','" & vUnitCode & "'"
                    'da1 = New SqlDataAdapter(vQuery, vConnection)
                    'ds1 = New DataSet
                    'da1.Fill(ds1, "CheckItemDup")
                    'dt1 = ds1.Tables("CheckItemDup")
                    'If dt1.Rows.Count > 0 Then
                    'vCheckItemDup = dt1.Rows(0).Item("duplicateItem")
                    'vCheckNoDup = dt1.Rows(0).Item("requestno_dup")
                    'vCheckCampaign = dt1.Rows(0).Item("campaigncode_dup")
                    'vCheckCampaignName = dt1.Rows(0).Item("campaignname_dup")
                    'End If

                    'If vCheckItemDup > 0 Then
                    'MsgBox("สินค้าซ้ำ ในแคมเปญ " & vCheckCampaign & "/" & vCheckCampaignName & " และเลขที่ " & vCheckNoDup & " ไม่สามารถเพิ่มได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    'Exit Sub
                    'End If

                    vItemName = dt.Rows(0).Item("itemname")
                    Me.DGVItemDetails.Item(2, vRow).Value = dt.Rows(0).Item("itemname")
                    Me.DGVItemDetails.Item(3, vRow).Value = dt.Rows(0).Item("unitcode")
                    Me.DGVItemDetails.Item(4, vRow).Value = 0

                    Dim dgvCmb As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn()

                    dgvCmb.HeaderText = "Name"

                    dgvCmb.Items.Add("Ghanashyam")
                    dgvCmb.Items.Add("Jignesh")
                    dgvCmb.Items.Add("Ishver")
                    dgvCmb.Items.Add("Anand")

                    dgvCmb.Name = "cmbName"

                    'DGVItemDetails.Columns.Add(dgvCmb)
                    'Me.DGVItemDetails.Item(3, 0).DataGridView.Columns.Add(dgvCmb)
                    'Me.DGVItemDetails.Columns.Item(3).DataGridView.Columns.Add(dgvCmb)

                    'Dim ComboColumn As DataGridViewComboBoxCell = (DataGridViewComboBoxCell)(DGVItemDetails.Rows(vRow).Cells(0)) 'Me.DGVItemDetails.Rows(3).Cells(0)

                    'ComboColumn.DataSource = dt.Rows(0).Item("unitcode")
                    'ComboColumn.DisplayMember = "Employee_Names"



                Else
                    MsgBox("สินค้า รหัส " & vItemCode & " ไม่มีข้อมูลในระบบ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")

                    Me.DGVItemDetails.Item(1, vRow).Value = ""
                    Me.DGVItemDetails.Item(2, vRow).Value = ""
                    Me.DGVItemDetails.Item(3, vRow).Value = ""
                    Me.DGVItemDetails.Item(4, vRow).Value = ""
                    Me.DGVItemDetails.Item(5, vRow).Value = ""
                    Me.DGVItemDetails.Item(6, vRow).Value = ""
                End If
            End If
        End If

        Dim vCharStr As String
        If e.ColumnIndex = 4 Then
            vCharStr = Me.DGVItemDetails.Item(4, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(4, e.RowIndex).Value = ""
                    MsgBox("ช่องจำนวนที่กำหนด ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(4, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 5 Then
            vCharStr = Me.DGVItemDetails.Item(5, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(5, e.RowIndex).Value = ""
                    MsgBox("ช่องค่าคอมฯขายสด ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(5, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 7 Then
            vCharStr = Me.DGVItemDetails.Item(7, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(7, e.RowIndex).Value = ""
                    MsgBox("ช่องค่าคอมฯขายเชื่อ ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(7, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 9 Then
            vCharStr = Me.DGVItemDetails.Item(9, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(7, e.RowIndex).Value = ""
                    MsgBox("ช่องค่าคอมฯขายเชื่อ ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(9, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        Dim vDisc2 As Integer
        Dim vDiscAmount2 As Double
        Dim vAfterDiscAmount2 As Double

        Dim vDisc3 As Integer
        Dim vDiscAmount3 As Double
        Dim vAfterDiscAmount3 As Double

        Dim vDisc4 As Integer
        Dim vDiscAmount4 As Double
        Dim vAfterDiscAmount4 As Double

        Dim vRebate As Integer
        Dim vRebateAmount As Double
        Dim vAfterRebateAmount As Double

        Dim vNetCost As Double
        Dim vDiscSpecial As Double

        Dim vLose As Double
        Dim vLoseAmount As Double
        Dim vAfterLoseAmount As Double

        Dim vTransInAmount As Double
        Dim vTransOutAmount As Double
        Dim vAdvertise As Double
        Dim vAdvertiseAmount As Double
        Dim vAfterAdvertiseAmount As Double

        Dim vVatCost As Double

        Dim vInstallAmount As Double
        Dim vServiceAmount As Double
        Dim vMarketCost As Double

        Dim vRelateStockPercent As Double
        Dim vRelateStockAmount As Double

        Dim vSmartPoint As Integer
        Dim vSmartPointAmount As Double
        Dim vAfterSmartPointAmount As Double

        Dim vCashProfit As Double
        Dim vCashProfitAmount As Double
        Dim vAfterCashProfitAmount As Double

        Dim vCreditProfit As Double
        Dim vCreditProfitAmount As Double
        Dim vAfterCreditProfitAmount As Double

        Dim vTotalPrice As Double
        Dim vCashPriceOwn1 As Double
        Dim vCashPriceSend1 As Double
        Dim vCreditPriceOwn1 As Double
        Dim vCreditPriceSend1 As Double
        Dim vSalePrice2 As Double

        Dim vTargetAmount As Double
        Dim vPremiumAmount As Double

        Dim vComm As Double
        Dim vCommAmount As Double
        Dim vAfterCommAmount As Double

        Dim vBaseProfitPercent As Double
        Dim vBaseProfit As Double


        vDOAmount = Me.DGVItemDetails.Item(4, e.RowIndex).Value
        vBillDisc = Me.DGVItemDetails.Item(5, e.RowIndex).Value
        vBillDiscAmount = (vDOAmount * (vBillDisc / 100))
        vAccCost = vDOAmount - vBillDiscAmount

        vDisc1 = Me.DGVItemDetails.Item(7, e.RowIndex).Value
        vDiscAmount1 = (vAccCost * (vDisc1 / 100))
        vAfterDiscAmount1 = vAccCost - vDiscAmount1

        vDisc2 = Me.DGVItemDetails.Item(9, e.RowIndex).Value
        vDiscAmount2 = (vAfterDiscAmount1 * (vdisc2 / 100))
        vAfterDiscAmount2 = vAfterDiscAmount1 - vDiscAmount2

        vDisc3 = Me.DGVItemDetails.Item(11, e.RowIndex).Value
        vDiscAmount3 = (vAfterDiscAmount2 * (vDisc3 / 100))
        vAfterDiscAmount3 = vAfterDiscAmount2 - vDiscAmount3

        vDisc4 = Me.DGVItemDetails.Item(13, e.RowIndex).Value
        vDiscAmount4 = (vAfterDiscAmount3 * (vDisc4 / 100))
        vAfterDiscAmount4 = vAfterDiscAmount3 - vDiscAmount4

        vRebate = Me.DGVItemDetails.Item(15, e.RowIndex).Value
        vRebateAmount = (vAfterDiscAmount4 * (vRebate / 100))
        vAfterRebateAmount = vAfterDiscAmount4 - vRebateAmount

        vDiscSpecial = Me.DGVItemDetails.Item(17, e.RowIndex).Value
        vNetCost = vAfterRebateAmount - vDiscSpecial

        vLose = Me.DGVItemDetails.Item(19, e.RowIndex).Value
        vLoseAmount = (vNetCost * (vLose / 100))
        vAfterLoseAmount = vNetCost - vLoseAmount

        vTransInAmount = Me.DGVItemDetails.Item(21, e.RowIndex).Value
        vTransOutAmount = Me.DGVItemDetails.Item(22, e.RowIndex).Value
        vAdvertise = Me.DGVItemDetails.Item(23, e.RowIndex).Value
        vAdvertiseAmount = (vAfterLoseAmount * (vAdvertise / 100))
        vAfterAdvertiseAmount = (vAfterLoseAmount + vTransInAmount + vTransOutAmount) - vAdvertiseAmount

        vVatCost = (vAfterAdvertiseAmount * 0.07) + vAfterAdvertiseAmount

        vInstallAmount = Me.DGVItemDetails.Item(26, e.RowIndex).Value
        vServiceAmount = Me.DGVItemDetails.Item(27, e.RowIndex).Value
        vMarketCost = vVatCost + vInstallAmount + vServiceAmount

        vCashPriceOwn1 = Me.DGVItemDetails.Item(42, e.RowIndex).Value
        vCashPriceSend1 = vCashPriceOwn1
        vCreditPriceOwn1 = Me.DGVItemDetails.Item(44, e.RowIndex).Value
        vCreditPriceSend1 = vCreditPriceOwn1

        vSmartPoint = Me.DGVItemDetails.Item(35, e.RowIndex).Value
        vSmartPointAmount = vSmartPoint
        vAfterSmartPointAmount = vSmartPointAmount * vCashPriceOwn1

        vTargetAmount = Me.DGVItemDetails.Item(37, e.RowIndex).Value
        vPremiumAmount = Me.DGVItemDetails.Item(38, e.RowIndex).Value

        vComm = Me.DGVItemDetails.Item(39, e.RowIndex).Value
        vCommAmount = vComm / 100
        vAfterCommAmount = vCommAmount * vCashPriceOwn1

        vTotalPrice = vMarketCost + vRelateStockAmount

        vAfterCashProfitAmount = vCashPriceOwn1 - vAfterCommAmount - vPremiumAmount - vTargetAmount - vAfterSmartPointAmount - vRelateStockAmount - vMarketCost
        vCashProfit = vAfterCashProfitAmount / vMarketCost

        vAfterCreditProfitAmount = vCreditPriceOwn1 - vMarketCost - vRelateStockAmount
        vCreditProfit = vAfterCreditProfitAmount / vMarketCost

        vBaseProfitPercent = ((vCashPriceOwn1 - vAfterCommAmount - vPremiumAmount - vAfterSmartPointAmount - vTargetAmount - vMarketCost) / vMarketCost) * 100
        vBaseProfit = ((vCashPriceOwn1 - vAfterCommAmount - vPremiumAmount - vAfterSmartPointAmount - vTargetAmount - vMarketCost) / vMarketCost)

        If vBaseProfit < 0.01 Then
            vRelateStockPercent = 0
        ElseIf vBaseProfit >= 0.01 And vBaseProfit < 0.05 Then
            vRelateStockPercent = vBaseProfit / 2
        ElseIf vBaseProfit >= 0.05 And vBaseProfit < 0.06 Then
            vRelateStockPercent = vBaseProfit * 0.6
        Else
            vRelateStockPercent = 0.035
        End If

        vRelateStockAmount = vRelateStockPercent * vMarketCost

        vItemName = Me.DGVItemDetails.Item(2, e.RowIndex).Value

        If vItemName <> "Nothing" And vItemName <> "" Then
            Me.DGVItemDetails.Item(6, e.RowIndex).Value = Format(vAccCost, "##,##0.00")
            Me.DGVItemDetails.Item(8, e.RowIndex).Value = Format(vAfterDiscAmount1, "##,##0.00")
            Me.DGVItemDetails.Item(10, e.RowIndex).Value = Format(vAfterDiscAmount2, "##,##0.00")
            Me.DGVItemDetails.Item(12, e.RowIndex).Value = Format(vAfterDiscAmount3, "##,##0.00")
            Me.DGVItemDetails.Item(14, e.RowIndex).Value = Format(vAfterDiscAmount4, "##,##0.00")
            Me.DGVItemDetails.Item(16, e.RowIndex).Value = Format(vAfterRebateAmount, "##,##0.00")
            Me.DGVItemDetails.Item(18, e.RowIndex).Value = Format(vNetCost, "##,##0.00")
            Me.DGVItemDetails.Item(20, e.RowIndex).Value = Format(vAfterLoseAmount, "##,##0.00")
            Me.DGVItemDetails.Item(24, e.RowIndex).Value = Format(vAfterAdvertiseAmount, "##,##0.00")
            Me.DGVItemDetails.Item(25, e.RowIndex).Value = Format(vVatCost, "##,##0.00")
            Me.DGVItemDetails.Item(28, e.RowIndex).Value = Format(vMarketCost, "##,##0.00")
            Me.DGVItemDetails.Item(30, e.RowIndex).Value = Format(vRelateStockAmount, "##,##0.00")
            Me.DGVItemDetails.Item(31, e.RowIndex).Value = Format(vCashProfit, "##,##0.00")
            Me.DGVItemDetails.Item(32, e.RowIndex).Value = Format(vAfterCashProfitAmount, "##,##0.00")
            Me.DGVItemDetails.Item(33, e.RowIndex).Value = Format(vCreditProfit, "##,##0.00")
            Me.DGVItemDetails.Item(34, e.RowIndex).Value = Format(vAfterCreditProfitAmount, "##,##0.00")
            Me.DGVItemDetails.Item(36, e.RowIndex).Value = Format(vAfterSmartPointAmount, "##,##0.00")
            Me.DGVItemDetails.Item(40, e.RowIndex).Value = Format(vAfterCommAmount, "##,##0.00")
            Me.DGVItemDetails.Item(41, e.RowIndex).Value = Format(vTotalPrice, "##,##0.00")
            Me.DGVItemDetails.Item(42, e.RowIndex).Value = Format(vCashPriceOwn1, "##,##0.00")
            Me.DGVItemDetails.Item(43, e.RowIndex).Value = Format(vCashPriceSend1, "##,##0.00")
            Me.DGVItemDetails.Item(44, e.RowIndex).Value = Format(vCreditPriceOwn1, "##,##0.00")
            Me.DGVItemDetails.Item(45, e.RowIndex).Value = Format(vCreditPriceSend1, "##,##0.00")
            Me.DGVItemDetails.Item(49, e.RowIndex).Value = Format(vBaseProfit, "##,##0.00")
        End If

    End Sub

    Public Sub vCheckNumber(ByVal vNumber As String)
        Dim vLen As Integer
        Dim vChar As String
        Dim i As Integer
        Dim vString As String

        On Error Resume Next

        vString = vNumber
        vLen = vb6.Len(vString)
        For i = 1 To vLen
            vChar = Mid(vString, i, 1)

            If vChar = "1" Or vChar = "2" Or vChar = "3" Or vChar = "4" Or vChar = "5" Or vChar = "6" Or vChar = "7" Or vChar = "8" Or vChar = "9" Or vChar = "0" Or vChar = "," Or vChar = "." Or vChar = "%" Then
                vIsNumber = 1
            Else
                vIsNumber = 0
                GoTo Line1
            End If
        Next
Line1:

    End Sub

    Private Sub DGVItemDetails_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItemDetails.CellContentClick

    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Me.Close()
    End Sub

    Private Sub BTNSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelect.Click
        Dim i As Integer
        Dim n As Integer
        Dim m As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vCheckItemCode As String
        Dim vCheckUnitCode As String
        Dim vCheckAdd As Integer

        If Me.ListViewSearch.Items.Count > 0 Then
            For i = 0 To Me.ListViewSearch.Items.Count - 1
                If Me.ListViewSearch.Items(i).Checked = True Then

                    vItemCode = Me.ListViewSearch.Items(i).SubItems(1).Text
                    vItemName = Me.ListViewSearch.Items(i).SubItems(2).Text
                    vUnitCode = Me.ListViewSearch.Items(i).SubItems(3).Text

                    For n = 0 To Me.DGVItemDetails.RowCount - 1
                        vCheckItemCode = Me.DGVItemDetails.Item(1, n).Value
                        vCheckUnitCode = Me.DGVItemDetails.Item(3, n).Value

                        If vItemCode = vCheckItemCode And vUnitCode = vCheckUnitCode Then
                            vCheckAdd = 1
                            GoTo Line1
                        Else
                            vCheckAdd = 0
                        End If
                    Next

                    If vCheckAdd = 0 Then
                        For m = 0 To Me.DGVItemDetails.RowCount - 1
                            If Me.DGVItemDetails.Item(1, m).Value = Nothing Then
                                Me.DGVItemDetails.Item(1, m).Value = vItemCode
                                Me.DGVItemDetails.Item(2, m).Value = vItemName
                                Me.DGVItemDetails.Item(3, m).Value = vUnitCode
                                GoTo Line1
                            End If
                        Next
                    End If

                End If
Line1:
            Next
            Me.PNSearch.Visible = False
        End If
    End Sub

    Private Sub BTNClickSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClickSearch.Click
        Call SearchItem()
    End Sub

    Public Sub SearchItem()
        Dim vSearch As String
        Dim vType As Integer
        Dim vBrandCode As String
        Dim vListItem As ListViewItem
        Dim i As Integer
        Dim n As Integer

        Dim vCashSalePrice As Double
        Dim vCreditSalePrice As Double


        If Me.CBNotAddPriceStructure.Checked = True Then
            vType = 1
        ElseIf Me.CBItemSaleLose.Checked = True Then
            vType = 2
        Else
            vType = 0
        End If
        Me.ListViewSearch.Items.Clear()
        vSearch = Me.TBSearch.Text
        If Me.CMBBrandCode.Text <> "" Then
            vBrandCode = vb6.Left(Me.CMBBrandCode.Text, vb6.InStr(Me.CMBBrandCode.Text, "/") - 1)
        Else
            vBrandCode = ""
        End If

        vQuery = "exec dbo.USP_NP_SearchItemPriceStructure " & vType & ",'" & vBrandCode & "','" & vSearch & "'"

        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchPaidNo")
        dt = ds.Tables("SearchPaidNo")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vCashSalePrice = dt.Rows(i).Item("cashsaleprice")
                vCreditSalePrice = dt.Rows(i).Item("creditsaleprice")

                vListItem = Me.ListViewSearch.Items.Add(n)
                vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("itemcode")
                vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("itemname")
                vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("unitcode")
                vListItem.SubItems.Add(3).Text = Format(vCashSalePrice, "##,##0.00")
                vListItem.SubItems.Add(4).Text = Format(vCreditSalePrice, "##,##0.00")

            Next
        End If
    End Sub

    Private Sub CBItemSaleLose_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBItemSaleLose.CheckedChanged
        If Me.CBItemSaleLose.Checked = True Then
            Me.CBNotAddPriceStructure.Checked = False
        End If
    End Sub

    Private Sub CBNotAddPriceStructure_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBNotAddPriceStructure.CheckedChanged
        If Me.CBNotAddPriceStructure.Checked = True Then
            Me.CBItemSaleLose.Checked = False
        End If
    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        Me.PNSearch.Visible = False
        Me.TBDocNo.Focus()
    End Sub

    Private Sub BTNSearchOldData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchOldData.Click
        Me.PNSearch.Visible = True
        Me.TBSearch.Focus()
    End Sub

    Private Sub TBSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call SearchItem()
        End If
    End Sub

    Private Sub TBSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearch.TextChanged

    End Sub

    Public Function CalcAmount(ByVal vPriceSetAmount As Double, ByVal vPercent As Double) As Double
        CalcAmount = (vPriceSetAmount * vPercent) / 100
    End Function

    Public Function CalcAmountAfterAdd(ByVal vPriceSetAmount As Double, ByVal vPercent As Double) As Double
        CalcAmountAfterAdd = vPriceSetAmount + (vPriceSetAmount * vPercent) / 100
    End Function
    Public Function CalcAmountAfterDelete(ByVal vPriceSetAmount As Double, ByVal vPercent As Double) As Double
        CalcAmountAfterDelete = vPriceSetAmount - (vPriceSetAmount * vPercent) / 100
    End Function

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vDocno As String
        Dim vDocDate As String
        Dim vTargetProfit As String
        Dim vProfit As String
        Dim vSmartPoint As String
        Dim vMemberDiscount As String
        Dim vFileDataSource As String
        Dim vPathFile As String
        '------------------------------------------
        Dim vItemCode As String
        Dim vItemName As String
        Dim vSaleUnitCode As String
        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vDiscountBillWord As String
        Dim vDiscountBill1 As Double
        Dim vDiscountBillAmount As Double
        Dim vAccCost As Double
        Dim vDiscountFollow1Word As String
        Dim vDiscountFollow11 As Double
        Dim vDiscountFollow1Amount As Double
        Dim vDiscountFollow1After As Double
        Dim vDiscountFollow2Word As String
        Dim vDiscountFollow21 As Double
        Dim vDiscountFollow2Amount As Double
        Dim vDiscountFollow2After As Double
        Dim vDiscountFollow3Word As String
        Dim vDiscountFollow31 As Double
        Dim vDiscountFollow3Amount As Double
        Dim vDiscountFollow3After As Double
        Dim vDiscountFollow4Word As String
        Dim vDiscountFollow41 As Double
        Dim vDiscountFollow4Amount As Double
        Dim vDiscountFollow4After As Double
        Dim vDiscountRebateWord As String
        Dim vDiscountRebate1 As Double
        Dim vDiscountRebateAmount As Double
        Dim vDiscountRebateAfter As Double
        Dim vDiscountSpecialWord As String
        Dim vDiscountSpecial1 As Double
        Dim vDiscountSpecialAmount As Double
        Dim vNetCost As Double
        Dim vLossBudgetWord As String
        Dim vLossBudget1 As Double
        Dim vLossBudgetAmount As Double
        Dim vLossBudgetAfter As Double
        Dim vTransferInWord As String
        Dim vTransferIn1 As Double
        Dim vTransferOutWord As String
        Dim vTransferOut1 As Double
        Dim vAdvertiseWord As String
        Dim vAdvertise1 As Double
        Dim vAdvertiseAmount As Double
        Dim vAdvertiseAfter As Double
        Dim vVatCost As Double
        Dim vVatAmount As Double
        Dim vSetupWord As String
        Dim vSetupAmount As Double
        Dim vServiceWord As String
        Dim vServiceAmount As Double
        Dim vMarketCost As Double
        Dim vPointWord As String
        Dim vPoint1 As Double
        Dim vPointAmount As Double
        Dim vPointAfter As Double
        Dim vTargetWord As String
        Dim vTargetAmount As Double
        Dim vPremiumWord As String
        Dim vPremiumAmount As Double
        Dim vCommissionWord As String
        Dim vCommission1 As Double
        Dim vCommissionAmount As Double
        Dim vCommissionAfter As Double
        Dim vGrossProfitPercent As String
        Dim vGrossProfitAmount As Double
        Dim vInterestStockPercent As String
        Dim vInterestStockAmount As Double
        Dim vProfitPercent As String
        Dim vProfitAmount As Double
        Dim vProfitPercent_W As String
        Dim vProfitAmount_W As Double
        Dim vMyDescription As String
        Dim vMyDescriptionSub As String
        Dim vTransferInAfter As Double
        Dim vTransferOutAfter As Double
        Dim vVatWord As String
        Dim vSetupAfter As Double
        Dim vTargetAfter As Double
        Dim vPremiumAfter As Double
        '--------------------------------------------------------------

        '---------------------------------------------------------------
        Dim i As Integer
        Dim vLenFilePath As Integer
        Dim vLenMark As Integer



        Me.PB101.Value = 1
        If Me.TBDocNo.Text = "" Then
            vQuery = "exec dbo.USP_PS_NewDocno"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "NewDocno")
            vdt = ds.Tables("NewDocno")
            vDocno = vdt.Rows(0).Item("newdocno")
            vDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        Else
            vDocno = Me.TBDocNo.Text
            vDocDate = Me.DTPDocDate.Value
        End If


        Try
            vQuery = "begin tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            If Me.TBTarget.Text <> "" Then
                vTargetProfit = Me.TBTarget.Text
            Else
                vTargetProfit = 0
            End If
            If Me.TBTargetAverage.Text <> "" Then
                vProfit = Me.TBTargetAverage.Text
            Else
                vProfit = 0
            End If
            If Me.TBSmartPoint.Text <> "" Then
                vSmartPoint = Me.TBSmartPoint.Text
            Else
                vSmartPoint = 0
            End If
            If Me.TBSmartPoint.Text <> "" Then
                vMemberDiscount = 0
            Else
                vMemberDiscount = 0
            End If
            vMyDescription = Me.TBMyDescription.Text
            vPathFile = ""
            vLenMark = ""
            vLenFilePath = ""
            vFileDataSource = "" 'Microsoft.VisualBasic.Right(Me.LBLFileName.Text, vLenFilePath - vLenMark)
            vQuery = "exec dbo.USP_PS_InsertPriceStructureSet1 '" & vDocno & "','" & vDocDate & "','" & vTargetProfit & "','" & vProfit & "','" & vSmartPoint & "','" & vMemberDiscount & "','" & vFileDataSource & "','" & vPathFile & "','" & vMyDescription & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            Me.PB101.Maximum = Me.DGVItemDetails.RowCount - 2

            For i = 0 To Me.DGVItemDetails.RowCount - 2

                vItemCode = Trim(Me.DGVItemDetails.Rows(i).Cells(0).Value)
                vItemName = Trim(Me.DGVItemDetails.Rows(i).Cells(1).Value)
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(2).Value) Then
                    vSaleUnitCode = Trim(Me.DGVItemDetails.Rows(i).Cells(2).Value)
                Else
                    vSaleUnitCode = ""
                End If
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(3).Value) Then
                    vDO = Trim(Me.DGVItemDetails.Rows(i).Cells(3).Value)
                Else
                    vDO = 0
                End If
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(42).Value) Then
                    vPriceSet = Me.DGVItemDetails.Rows(i).Cells(42).Value
                Else
                    vPriceSet = 0
                End If
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(4).Value) Then
                    vDiscountBillWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(4).Value, Double) * 100), String) & "%"
                    vDiscountBill1 = (CType(Me.DGVItemDetails.Rows(i).Cells(4).Value, Double) * 100)
                Else
                    vDiscountBillWord = ""
                    vDiscountBill1 = 0
                End If
                vDiscountBillAmount = CalcAmount(vDO, vDiscountBill1)
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(5).Value) Then
                    vAccCost = Me.DGVItemDetails.Rows(i).Cells(5).Value
                Else
                    vAccCost = 0
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(6).Value) Then
                    vDiscountFollow1Word = CType((CType(Me.DGVItemDetails.Rows(i).Cells(6).Value, Double) * 100), String) & "%"
                    vDiscountFollow11 = (CType(Me.DGVItemDetails.Rows(i).Cells(6).Value, Double) * 100)
                Else
                    vDiscountFollow1Word = ""
                    vDiscountFollow11 = 0
                End If
                vDiscountFollow1Amount = CalcAmount(vAccCost, vDiscountFollow11)
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(7).Value) Then
                    vDiscountFollow1After = Me.DGVItemDetails.Rows(i).Cells(7).Value
                Else
                    vDiscountFollow1After = 0
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(8).Value) Then
                    vDiscountFollow2Word = CType((CType(Me.DGVItemDetails.Rows(i).Cells(8).Value, Double) * 100), String) & "%"
                    vDiscountFollow21 = (CType(Me.DGVItemDetails.Rows(i).Cells(8).Value, Double) * 100)
                Else
                    vDiscountFollow2Word = ""
                    vDiscountFollow21 = 0
                End If
                vDiscountFollow2Amount = CalcAmount(vDiscountFollow1After, vDiscountFollow21)
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(9).Value) Then
                    vDiscountFollow2After = Me.DGVItemDetails.Rows(i).Cells(9).Value
                Else
                    vDiscountFollow2After = 0
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(10).Value) Then
                    vDiscountFollow3Word = CType((CType(Me.DGVItemDetails.Rows(i).Cells(10).Value, Double) * 100), String) & "%"
                    vDiscountFollow31 = (CType(Me.DGVItemDetails.Rows(i).Cells(10).Value, Double) * 100)
                Else
                    vDiscountFollow3Word = ""
                    vDiscountFollow31 = 0
                End If
                vDiscountFollow3Amount = CalcAmount(vDiscountFollow2After, vDiscountFollow31)
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(11).Value) Then
                    vDiscountFollow3After = Me.DGVItemDetails.Rows(i).Cells(11).Value
                Else
                    vDiscountFollow3After = 0
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(12).Value) Then
                    vDiscountFollow4Word = CType((CType(Me.DGVItemDetails.Rows(i).Cells(12).Value, Double) * 100), String) & "%"
                    vDiscountFollow41 = (CType(Me.DGVItemDetails.Rows(i).Cells(12).Value, Double) * 100)
                Else
                    vDiscountFollow4Word = ""
                    vDiscountFollow41 = 0
                End If
                vDiscountFollow4Amount = CalcAmount(vDiscountFollow3After, vDiscountFollow41)
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(13).Value) Then
                    vDiscountFollow4After = Me.DGVItemDetails.Rows(i).Cells(13).Value
                Else
                    vDiscountFollow4After = 0
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(14).Value) Then
                    vDiscountRebateWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(14).Value, Double) * 100), String) & "%"
                    vDiscountRebate1 = (CType(Me.DGVItemDetails.Rows(i).Cells(14).Value, Double) * 100)
                Else
                    vDiscountRebateWord = ""
                    vDiscountRebate1 = 0
                End If
                vDiscountRebateAmount = CalcAmount(vDiscountFollow4After, vDiscountRebate1)
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(15).Value) Then
                    vDiscountRebateAfter = Me.DGVItemDetails.Rows(i).Cells(15).Value
                Else
                    vDiscountRebateAfter = 0
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(16).Value) Then
                    vDiscountSpecialWord = Me.DGVItemDetails.Rows(i).Cells(16).Value
                    vDiscountSpecial1 = Me.DGVItemDetails.Rows(i).Cells(16).Value
                Else
                    vDiscountSpecialWord = 0
                    vDiscountSpecial1 = 0
                End If
                vDiscountSpecialAmount = vDiscountSpecial1
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(17).Value) Then
                    vNetCost = Me.DGVItemDetails.Rows(i).Cells(17).Value
                Else
                    vNetCost = 0
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(18).Value) Then
                    vLossBudgetWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(18).Value, Double) * 100), String) & "%"
                    vLossBudget1 = (CType(Me.DGVItemDetails.Rows(i).Cells(18).Value, Double) * 100)
                Else
                    vLossBudgetWord = ""
                    vLossBudget1 = 0
                End If
                vLossBudgetAmount = CalcAmount(vNetCost, vLossBudget1)
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(19).Value) Then
                    vLossBudgetAfter = Me.DGVItemDetails.Rows(i).Cells(19).Value
                Else
                    vLossBudgetAfter = 0
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(20).Value) Then
                    vTransferInWord = Me.DGVItemDetails.Rows(i).Cells(20).Value
                    vTransferIn1 = Me.DGVItemDetails.Rows(i).Cells(20).Value
                    vTransferInAfter = vLossBudgetAfter + vTransferIn1
                Else
                    vTransferIn1 = 0
                    vTransferInWord = ""
                    vTransferInAfter = vLossBudgetAfter + vTransferIn1
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(21).Value) Then
                    vTransferOutWord = Me.DGVItemDetails.Rows(i).Cells(21).Value
                    vTransferOut1 = Me.DGVItemDetails.Rows(i).Cells(21).Value
                    vTransferOutAfter = vTransferInAfter + vTransferOut1
                Else
                    vTransferOutWord = ""
                    vTransferOut1 = 0
                    vTransferOutAfter = vTransferInAfter + vTransferOut1
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(22).Value) Then
                    vAdvertiseWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(22).Value, Double) * 100), String) & "%"
                    vAdvertise1 = (CType(Me.DGVItemDetails.Rows(i).Cells(22).Value, Double) * 100)
                Else
                    vAdvertiseWord = ""
                    vAdvertise1 = 0
                End If
                vAdvertiseAmount = CalcAmount(vTransferOutAfter, vAdvertise1)
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(23).Value) Then
                    vAdvertiseAfter = Me.DGVItemDetails.Rows(i).Cells(23).Value
                Else
                    vAdvertiseAfter = 0
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(24).Value) Then
                    vVatCost = Me.DGVItemDetails.Rows(i).Cells(24).Value
                    vVatAmount = (vAdvertiseAfter * 7) / 100
                    vVatWord = "7%"
                Else
                    vVatCost = 0
                    vVatAmount = 0
                    vVatWord = 0
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(25).Value) Then
                    vSetupWord = Me.DGVItemDetails.Rows(i).Cells(25).Value
                    vSetupAmount = Me.DGVItemDetails.Rows(i).Cells(25).Value
                    vSetupAfter = vVatCost + vSetupAmount
                Else
                    vSetupWord = ""
                    vSetupAmount = 0
                    vSetupAfter = vVatCost + vSetupAmount
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(26).Value) Then
                    vServiceWord = Me.DGVItemDetails.Rows(i).Cells(26).Value
                    vServiceAmount = Me.DGVItemDetails.Rows(i).Cells(26).Value
                Else
                    vServiceWord = ""
                    vServiceAmount = 0
                End If
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(27).Value) Then
                    vMarketCost = Me.DGVItemDetails.Rows(i).Cells(27).Value
                Else
                    vMarketCost = 0
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(34).Value) Then
                    vPointWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(34).Value, Double) * 100), String) & "%"
                Else
                    vPointWord = ""
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(35).Value) Then
                    vPoint1 = Me.DGVItemDetails.Rows(i).Cells(35).Value
                Else
                    vPoint1 = 0
                End If
                vPointAmount = vPoint1
                vPointAfter = vMarketCost + vPointAmount

                If vPointWord = "" And vPointAmount = 0 Then
                    MsgBox("รหัสสินค้า " & vItemCode & "   " & vItemName & " ไม่มีค่าของ Smart Point ไม่สามารถบันทึกข้อมูลได้  กรุณาแก้ไขข้อมูลก่อนบันทึกใหม่")
                    vQuery = "rollback tran"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()
                    Exit Sub
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(36).Value) Then
                    vTargetWord = Me.DGVItemDetails.Rows(i).Cells(36).Value
                    vTargetAmount = Me.DGVItemDetails.Rows(i).Cells(36).Value
                    vTargetAfter = vPointAfter + vTargetAmount
                Else
                    vTargetWord = ""
                    vTargetAmount = 0
                    vTargetAfter = vPointAfter + vTargetAmount
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(37).Value) Then
                    vPremiumWord = Me.DGVItemDetails.Rows(i).Cells(37).Value
                    vPremiumAmount = Me.DGVItemDetails.Rows(i).Cells(37).Value
                    vPremiumAfter = vTargetAfter + vPremiumAmount
                Else
                    vPremiumWord = ""
                    vPremiumAmount = 0
                    vPremiumAfter = vTargetAfter + vPremiumAmount
                End If

                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(38).Value) Then
                    vCommissionWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(38).Value, Double) * 100), String) & "%"
                Else
                    vCommissionWord = ""
                End If
                vCommissionAmount = vCommission1
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(39).Value) Then
                    vCommission1 = Me.DGVItemDetails.Rows(i).Cells(39).Value
                Else
                    vCommission1 = 0
                End If
                vCommissionAmount = vCommission1
                vCommissionAfter = vPremiumAfter + vCommissionAmount

                vGrossProfitPercent = (Me.DGVItemDetails.Rows(i).Cells(48).Value * 100)
                vGrossProfitAmount = ((Me.DGVItemDetails.Rows(i).Cells(48).Value * 100) * vMarketCost) / 100
                vInterestStockPercent = (Me.DGVItemDetails.Rows(i).Cells(28).Value * 100)
                vInterestStockAmount = Me.DGVItemDetails.Rows(i).Cells(29).Value
                vProfitPercent = (Me.DGVItemDetails.Rows(i).Cells(30).Value * 100)
                vProfitAmount = Me.DGVItemDetails.Rows(i).Cells(31).Value
                vProfitPercent_W = (Me.DGVItemDetails.Rows(i).Cells(32).Value * 100)
                vProfitAmount_W = Me.DGVItemDetails.Rows(i).Cells(33).Value
                If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(47).Value) Then
                    vMyDescriptionSub = Me.DGVItemDetails.Rows(i).Cells(47).Value
                Else
                    vMyDescriptionSub = ""
                End If
                '---------------------------------------------------------------------------------------------------
                Dim vFromQTY As Double
                Dim vToQTY As Double
                Dim vPriceSet2 As Double
                Dim vIsPriceUpdate As Integer = 1
                Dim vToUpdateDate As String = Me.DGVItemDetails.Rows(i).Cells(46).Value
                Dim vIsUpdate As Integer = 0
                Dim vIsPrintLabel As Integer = 0
                Dim vPrice1CashRec As Double
                Dim vPrice1CashDel As Double
                Dim vPrice1CreditRec As Double
                Dim vPrice1CreditDel As Double
                '---------------------------------------------------------------------------------------------------

                vSaleUnitCode = Me.DGVItemDetails.Rows(i).Cells(2).Value
                vFromQTY = 1
                vToQTY = 99999
                vPrice1CashRec = Me.DGVItemDetails.Rows(i).Cells(41).Value
                vPrice1CashDel = Me.DGVItemDetails.Rows(i).Cells(42).Value
                vPrice1CreditRec = Me.DGVItemDetails.Rows(i).Cells(43).Value
                vPrice1CreditDel = Me.DGVItemDetails.Rows(i).Cells(44).Value
                vPriceSet2 = Me.DGVItemDetails.Rows(i).Cells(45).Value


                vQuery = "exec dbo.USP_PS_InsertPriceStructureSubSet '" & vDocno & "','" & vItemCode & "','" & vItemName & "','" & vSaleUnitCode & "'," & vDO & ", " _
                & "" & vPriceSet & ",'" & vDiscountBillWord & "'," & vDiscountBillAmount & "," & vAccCost & "," _
                & "'" & vDiscountFollow1Word & "'," & vDiscountFollow1Amount & "," & vDiscountFollow1After & "," _
                & "'" & vDiscountFollow2Word & "'," & vDiscountFollow2Amount & "," & vDiscountFollow2After & "," _
                & "'" & vDiscountFollow3Word & "'," & vDiscountFollow3Amount & "," & vDiscountFollow3After & "," _
                & "'" & vDiscountFollow4Word & "'," & vDiscountFollow4Amount & "," & vDiscountFollow4After & "," _
                & "'" & vDiscountRebateWord & "'," & vDiscountRebateAmount & "," & vDiscountRebateAfter & "," _
                & "'" & vDiscountSpecialWord & "'," & vDiscountSpecialAmount & "," & vNetCost & "," _
                & "'" & vLossBudgetWord & "'," & vLossBudgetAmount & "," & vLossBudgetAfter & "," _
                & "'" & vTransferInWord & "'," & vTransferIn1 & "," & vTransferInAfter & "," _
                & "'" & vTransferOutWord & "'," & vTransferOut1 & "," & vTransferOutAfter & "," _
                & "'" & vAdvertiseWord & "'," & vAdvertiseAmount & "," & vAdvertiseAfter & "," _
                & "'" & vVatWord & "'," & vVatCost & "," & vVatAmount & "," _
                & "'" & vSetupWord & "'," & vSetupAmount & "," & vSetupAfter & "," _
                & "'" & vServiceWord & "'," & vServiceAmount & "," & vMarketCost & "," _
                & "'" & vPointWord & "'," & vPointAmount & "," & vPointAfter & "," _
                & "'" & vTargetWord & "'," & vTargetAmount & "," & vTargetAfter & "," _
                & "'" & vPremiumWord & "'," & vPremiumAmount & "," & vPremiumAfter & "," _
                & "'" & vCommissionWord & "'," & vCommissionAmount & "," & vCommissionAfter & "," _
                & "" & vGrossProfitPercent & "," & vGrossProfitAmount & "," & vInterestStockPercent & "," & vInterestStockAmount & "," _
                & "" & vProfitPercent & "," & vProfitAmount & "," & vProfitPercent_W & "," & vProfitAmount_W & ",'" & vMyDescriptionSub & "'," & vFromQTY & "," & vToQTY & "," & vPrice1CashRec & "," & vPrice1CashDel & "," & vPrice1CreditRec & "," & vPrice1CreditDel & "," & vPriceSet2 & ", " _
                & "" & vIsPriceUpdate & ",'" & vToUpdateDate & "'," & vIsUpdate & " "
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()
                Me.PB101.Value = i
            Next


            vQuery = "commit tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            vQuery = "exec dbo.USP_PS_DeliverySendMail '" & vDocno & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            MsgBox("บันทึกข้อมูลโครงสร้างราคาเลขที่ " & vDocno & " เรียบร้อยแล้วครับ")
            Me.PB101.Value = 0
            Me.DGVItemDetails.DataSource = Nothing
            Me.PB101.Value = 0
            Me.TBDocNo.Text = ""
            Me.TBMyDescription.Text = ""
            Me.DTPDocDate.Text = Now.Date
            Me.TBDocNo.Text = ""
            Me.PBNew.Visible = True
            Me.PBConfirm.Visible = False

            vPriceStructureDocNo = Trim(vDocno)

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            vQuery = "rollback tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()
        End Try


    End Sub

    Private Sub DTPUpdate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTPUpdate.ValueChanged
        Dim i As Integer
        Dim vItemCode As String


        For i = 0 To Me.DGVItemDetails.RowCount - 1
            vItemCode = Me.DGVItemDetails.Item(1, i).Value
            If vItemCode <> "Nothing" And vItemCode <> "" Then
                If vb6.Year(Me.DTPUpdate.Value) > 2500 Then
                    Me.DGVItemDetails.Item(47, i).Value = vb6.Day(Me.DTPUpdate.Value) & "/" & vb6.Month(Me.DTPUpdate.Value) & "/" & vb6.Year(Me.DTPUpdate.Value) - 543
                Else
                    Me.DGVItemDetails.Item(47, i).Value = vb6.Day(Me.DTPUpdate.Value) & "/" & vb6.Month(Me.DTPUpdate.Value) & "/" & vb6.Year(Me.DTPUpdate.Value)
                End If
            End If
        Next
    End Sub
End Class