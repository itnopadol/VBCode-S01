Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class FormBudgetTargetDepartmentConfig
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand
    Dim vDepartmentClick As String
    Dim vProgressBarClick As Double
    Dim vRemainProgress As Double
    Dim vSelectItemBudget As Integer

    Private Sub FormBudgetTargetDepartmentConfig_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer

        On Error GoTo ErrDescription

        Call InitializeDataBase()
        Me.CMBSaleType.Text = Me.CMBSaleType.Items(0)
        vQuery = "exec dbo.USP_ICT_FiscalYearList"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "FiscalYearList")
        dt = ds.Tables("FiscalYearList")
        For i = 0 To dt.Rows.Count - 1
            Me.CMBFiscalYear.Items.Add(dt.Rows(i).Item("fiscalyear"))
        Next
        Me.CMBFiscalYear.Text = Now.Year

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub CMBFiscalYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBFiscalYear.SelectedIndexChanged
        Dim vFisCalYear As String
        Dim i As Integer

        On Error GoTo ErrDescription

        vFisCalYear = Me.CMBFiscalYear.SelectedItem

        vQuery = "exec dbo.USP_ICT_PeriodOf4WeekList '" & vFisCalYear & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Period")
        dt = ds.Tables("Period")

        Me.CMBPeriod.Items.Clear()
        For i = 0 To dt.Rows.Count - 1
            Me.CMBPeriod.Items.Add(dt.Rows(i).Item("fiscalyear"))
        Next
        Me.CMBPeriod.Text = Now.Month

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub CMBPeriod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBPeriod.SelectedIndexChanged
        Dim vSaleType As Integer
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer
        Dim vListBudget As ListViewItem
        Dim i As Integer

        On Error GoTo ErrDescription

        vSaleType = Me.CMBSaleType.SelectedIndex
        vFiscalYear = Me.CMBFiscalYear.Text
        vPeriodOf4Week = Me.CMBPeriod.Text

        vQuery = "exec dbo.USP_ICT_DebtBudgetPlan '" & vSaleType & "'," & vFiscalYear & "," & vPeriodOf4Week & " "
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "DebtBudgetPlan")
        dt = ds.Tables("DebtBudgetPlan")
        ListView101.Items.Clear()
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                vListBudget = ListView101.Items.Add(Trim(dt.Rows(i).Item("departmentcode")))
                vListBudget.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("department"))
                vListBudget.SubItems.Add(2).Text = Format(dt.Rows(i).Item("targetsale"), "##,##0.00")
                vListBudget.SubItems.Add(3).Text = Format(dt.Rows(i).Item("targetgp"), "##,##0.00")
                vListBudget.SubItems.Add(4).Text = Format(dt.Rows(i).Item("budgetpercent"), "##,##0.00")
                vListBudget.SubItems.Add(5).Text = Format(dt.Rows(i).Item("budgetsalemin"), "##,##0.00")
                vListBudget.SubItems.Add(6).Text = Format(dt.Rows(i).Item("budgetsalemax"), "##,##0.00")
                vListBudget.SubItems.Add(7).Text = Format(dt.Rows(i).Item("budgetgpmin"), "##,##0.00")
                vListBudget.SubItems.Add(8).Text = Format(dt.Rows(i).Item("budgetgpmax"), "##,##0.00")
            Next
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub ListView101_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView101.DoubleClick
        On Error GoTo ErrDescription

        If ListView101.Items.Count > 0 Then
            vDepartmentClick = Trim(ListView101.SelectedItems(0).SubItems(0).Text)
            Me.LBLDepartment.Text = "แผนก : " & Trim(ListView101.SelectedItems(0).SubItems(1).Text)
            GB101.Visible = True
            Call ClearData()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCancel.Click
        Call ClearData()
        GB101.Visible = False
    End Sub

    Private Sub BTNCal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCal.Click
        Dim vSaleType As Integer
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer
        Dim vListBudget As ListViewItem
        Dim i As Integer

        On Error GoTo ErrDescription

        Me.Cursor = Cursors.WaitCursor

        vSaleType = Me.CMBSaleType.SelectedIndex
        vFiscalYear = Me.CMBFiscalYear.Text
        vPeriodOf4Week = Me.CMBPeriod.Text
        vDepartmentClick = Trim(ListView101.SelectedItems(0).SubItems(0).Text)

        vQuery = "exec dbo.USP_ICT_BudgetTargetSetList '" & vSaleType & "'," & vFiscalYear & "," & vPeriodOf4Week & ",'" & vDepartmentClick & "' "
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "BudgetTargetSetList")
        dt = ds.Tables("BudgetTargetSetList")
        If dt.Rows.Count > 0 Then
            Me.TextLastYearSale.Text = Format(Int(dt.Rows(0).Item("lastyearsale")), "##,##0.00")
            Me.TextLastYearGP.Text = Format(Int(dt.Rows(0).Item("lastyeargp")), "##,##0.00")
            Me.TextTargetSale.Text = Format(Int(dt.Rows(0).Item("targetsale")), "##,##0.00")
            Me.TextTargetGP.Text = Format(Int(dt.Rows(0).Item("targetgp")), "##,##0.00")
            Me.NTargetSale.Value = Format(dt.Rows(0).Item("targetsale%"), "##,##0.00")
            Me.NTargetGP.Value = Format(dt.Rows(0).Item("targetgp%"), "##,##0.00")
        End If

        vQuery = "exec dbo.USP_ICT_BudgetSetList '" & vSaleType & "'," & vFiscalYear & "," & vPeriodOf4Week & ",'" & vDepartmentClick & "' "
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "BudgetSetList")
        dt = ds.Tables("BudgetSetList")
        If dt.Rows.Count > 0 Then
            PB101.Value = dt.Rows(0).Item("Period_BudgetRemainPercent")
            vRemainProgress = dt.Rows(0).Item("Period_BudgetRemainPercent")
            Me.TextPeriodSaleMin.Text = Format(Int(dt.Rows(0).Item("Period_BudgetSaleMin")), "##,##0.00")
            Me.TextPeriodSaleMax.Text = Format(Int(dt.Rows(0).Item("Period_BudgetSaleMax")), "##,##0.00")
            Me.TextPeriodGPMin.Text = Format(Int(dt.Rows(0).Item("Period_BudgetGPMin")), "##,##0.00")
            Me.TextPeriodGPMax.Text = Format(Int(dt.Rows(0).Item("Period_BudgetGPMax")), "##,##0.00")
            Me.TextBudgetSaleMin.Text = Format(dt.Rows(0).Item("Dept_BudgetSaleMin"), "##,##0.00")
            Me.TextBudgetSaleMax.Text = Format(dt.Rows(0).Item("Dept_BudgetSaleMax"), "##,##0.00")
            Me.TextBudgetGPMin.Text = Format(dt.Rows(0).Item("Dept_BudgetGPMin"), "##,##0.00")
            Me.TextBudgetGPMax.Text = Format(dt.Rows(0).Item("Dept_BudgetGPMax"), "##,##0.00")
            vProgressBarClick = dt.Rows(0).Item("Dept_BudgetPercent")
            Me.NBudgetSaleMin.Value = Format(dt.Rows(0).Item("Dept_BudgetPercent"), "##,##0.00")
            LBLRemainPercent.Text = "งบประมาณ Period คงเหลือ     " & "" & Format(dt.Rows(0).Item("Period_BudgetRemainPercent"), "##,##0.00") & "%"
        End If

        vQuery = "exec dbo.USP_ICT_StepSetList '" & vSaleType & "'," & vFiscalYear & "," & vPeriodOf4Week & ",'" & vDepartmentClick & "' "
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "StepSetList")
        dt = ds.Tables("StepSetList")
        ListView102.Items.Clear()
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                vListBudget = ListView102.Items.Add(dt.Rows(i).Item("type"))
                vListBudget.SubItems.Add(0).Text = dt.Rows(i).Item("status")
                vListBudget.SubItems.Add(1).Text = Format(dt.Rows(i).Item("growth"), "##,##0.00")
                vListBudget.SubItems.Add(2).Text = Format(dt.Rows(i).Item("growthrate"), "##,##0.00")
            Next
        End If
        Me.Cursor = Cursors.Default

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Public Sub ClearData()
        On Error Resume Next

        Me.TextLastYearGP.Text = ""
        Me.TextLastYearSale.Text = ""
        Me.TextBudgetGPMax.Text = ""
        Me.TextBudgetGPMin.Text = ""
        Me.TextBudgetSaleMax.Text = ""
        Me.TextBudgetSaleMin.Text = ""
        Me.TextPeriodGPMax.Text = ""
        Me.TextPeriodGPMin.Text = ""
        Me.TextPeriodSaleMax.Text = ""
        Me.TextPeriodSaleMin.Text = ""
        Me.NTargetGP.Value = 0.0
        Me.NTargetSale.Value = 0.0
        Me.NBudgetSaleMin.Value = 0.0
        Me.ListView102.Items.Clear()
        Me.PB101.Value = 0
        Me.TextTargetGP.Text = ""
        Me.TextTargetSale.Text = ""

    End Sub

    Private Sub NTargetSale_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NTargetSale.ValueChanged
        'Dim vNTargetSale As Double
        'Dim vTargetSale As Double
        'Dim vLastYearSale As Double

        'vNTargetSale = Me.NTargetSale.Value
        'If Me.TextLastYearSale.Text <> "" Then
        '    vLastYearSale = Me.TextLastYearSale.Text
        'Else
        '    vLastYearSale = 0
        'End If
        'vTargetSale = (vNTargetSale * vLastYearSale) / 100
        'Me.TextTargetSale.Text = Format(vTargetSale, "##,##0.00")
    End Sub

    Private Sub NTargetGP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NTargetGP.ValueChanged
        'Dim vNTargetGP As Double
        'Dim vTargetGP As Double
        'Dim vLastYearGP As Double

        'vNTargetGP = Me.NTargetGP.Value
        'If Me.TextLastYearGP.Text <> "" Then
        '    vLastYearGP = Me.TextLastYearGP.Text
        'Else
        '    vLastYearGP = 0
        'End If
        'vTargetGP = (vNTargetGP * vLastYearGP) / 100
        'Me.TextTargetGP.Text = Format(vTargetGP, "##,##0.00")
    End Sub

    Private Sub NBudgetSaleMin_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NBudgetSaleMin.ValueChanged
        Dim vNBudgetSaleMin As Double
        Dim vPeriodSaleMin As Double
        Dim vBudgetSaleMin As Double
        Dim vPeriodSaleMax As Double
        Dim vBudgetSaleMax As Double
        Dim vPeriodGPMin As Double
        Dim vBudgetGPMin As Double
        Dim vPeriodGPMax As Double
        Dim vBudgetGPMax As Double
        Dim vCheckPercent As Double
        Dim vTotalRemain As Double

        On Error GoTo ErrDescription

        If Me.TextPeriodSaleMin.Text <> "" Then
            vCheckPercent = Me.NBudgetSaleMin.Value
            If vRemainProgress = 0 And vCheckPercent > vProgressBarClick Then
                MsgBox("ไม่สามารถเพิ่ม % ได้ เนื่องจาก % ในการกำหนดครบ 100% แล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Error")
                Me.NBudgetSaleMin.Value = Format(vProgressBarClick, "##,##0.00")
                Me.PB101.Value = vRemainProgress
                LBLRemainPercent.Text = "งบประมาณ Period คงเหลือ     " & "" & Format(vRemainProgress, "##,##0.00") & "%"
                Exit Sub
            End If

            If vRemainProgress = 0 And vCheckPercent < vProgressBarClick Then
                Me.PB101.Value = (vProgressBarClick - vCheckPercent)
                LBLRemainPercent.Text = "งบประมาณ Period คงเหลือ     " & "" & Format((vTotalRemain - vCheckPercent), "##,##0.00") & "%"
            End If

            If vRemainProgress <> 0 Then
                vTotalRemain = vProgressBarClick + vRemainProgress
            End If
            If vRemainProgress <> 0 And vCheckPercent > vTotalRemain Then
                MsgBox("ไม่สามารถเพิ่ม % ได้ เนื่องจาก % ในการกำหนดครบ 100% แล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Error")
                Me.NBudgetSaleMin.Value = Format(vProgressBarClick, "##,##0.00")
                Me.PB101.Value = vRemainProgress
                LBLRemainPercent.Text = "งบประมาณ Period คงเหลือ     " & "" & Format(vRemainProgress, "##,##0.00") & "%"
                Exit Sub
            End If

            If vRemainProgress <> 0 And vCheckPercent < vTotalRemain Then
                Me.PB101.Value = (vTotalRemain - vCheckPercent)
                LBLRemainPercent.Text = "งบประมาณ Period คงเหลือ     " & "" & Format((vTotalRemain - vCheckPercent), "##,##0.00") & "%"
            End If


            vNBudgetSaleMin = Me.NBudgetSaleMin.Value
            If Me.TextPeriodSaleMin.Text <> "" Then
                vPeriodSaleMin = Me.TextPeriodSaleMin.Text
            Else
                vPeriodSaleMin = 0
            End If
            vBudgetSaleMin = (vPeriodSaleMin * vNBudgetSaleMin) / 100
            Me.TextBudgetSaleMin.Text = Format(vBudgetSaleMin, "##,##0.00")

            If Me.TextPeriodSaleMax.Text <> "" Then
                vPeriodSaleMax = Me.TextPeriodSaleMax.Text
            Else
                vPeriodSaleMax = 0
            End If
            vBudgetSaleMax = (vPeriodSaleMax * vNBudgetSaleMin) / 100
            Me.TextBudgetSaleMax.Text = Format(vBudgetSaleMax, "##,##0.00")

            If Me.TextPeriodGPMin.Text <> "" Then
                vPeriodGPMin = Me.TextPeriodGPMin.Text
            Else
                vPeriodGPMin = 0
            End If
            vBudgetGPMin = (vPeriodGPMin * vNBudgetSaleMin) / 100
            Me.TextBudgetGPMin.Text = Format(vBudgetGPMin, "##,##0.00")

            If Me.TextPeriodGPMax.Text <> "" Then
                vPeriodGPMax = Me.TextPeriodGPMax.Text
            Else
                vPeriodGPMax = 0
            End If
            vBudgetGPMax = (vPeriodGPMax * vNBudgetSaleMin) / 100
            Me.TextBudgetGPMax.Text = Format(vBudgetGPMax, "##,##0.00")
        Else
            NBudgetSaleMin.Value = 0.0
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub ListView102_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView102.DoubleClick
        On Error GoTo ErrDescription

        If ListView102.SelectedItems(0).SubItems(0).Text = 1 Then
            GB102.Visible = True
            vSelectItemBudget = ListView102.SelectedItems(0).Index
            NTargetBudgetPercent.Value = ListView102.SelectedItems(0).SubItems(2).Text
            NTargetBudgetPercent.Focus()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNKeyPercent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNKeyPercent.Click
        On Error GoTo ErrDescription

        ListView102.Items(vSelectItemBudget).SubItems(2).Text = Format(NTargetBudgetPercent.Value, "##,##0.00")
        GB102.Visible = False

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNOK.Click
        Dim vSaleType As Integer
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer
        Dim vDepartmentCode As String
        Dim vLastYearSale As Double
        Dim vLastYearGP As Double
        Dim vTargetSalePercent As Double
        Dim vTargetGPPercent As Double
        Dim vTargetSale As Double
        Dim vTargetGP As Double
        Dim vBudgetPercent As Double
        Dim vBudgetSaleMin As Double
        Dim vBudgetSaleMax As Double
        Dim vBudgetGPMin As Double
        Dim vBudgetGPMax As Double
        Dim vType As Integer
        Dim vStep As Integer
        Dim vPercentTarget As Double
        Dim vRate As Double
        Dim i As Integer


        If Me.TextTargetSale.Text = "" And Me.TextTargetGP.Text = "" Then
            MsgBox("ไม่ได้กำหนดเป้าแผนก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
        If Me.TextBudgetSaleMin.Text = "" And Me.TextBudgetSaleMax.Text = "" And Me.TextBudgetGPMin.Text = "" And Me.TextBudgetGPMax.Text = "" Then
            MsgBox("ไม่ได้กำหนดงบประมาณแผนก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If

        If Me.ListView102.Items.Count > 0 Then
            On Error GoTo ErrDescription
            'Try


            'vQuery = "begin tran"
            'vCMD = New SqlCommand(vQuery, vConnection)
            'vCMD.ExecuteNonQuery()


            vSaleType = Me.CMBSaleType.SelectedIndex
            vFiscalYear = Me.CMBFiscalYear.Text
            vPeriodOf4Week = Me.CMBPeriod.Text
            vDepartmentCode = Trim(ListView101.SelectedItems(0).SubItems(0).Text)

            vLastYearSale = Me.TextLastYearSale.Text
            vLastYearGP = Me.TextLastYearGP.Text
            vTargetSalePercent = Me.NTargetSale.Value
            vTargetGPPercent = Me.NTargetGP.Value
            vTargetSale = Me.TextTargetSale.Text
            vTargetGP = Me.TextTargetGP.Text
            vBudgetPercent = Me.NBudgetSaleMin.Value
            vBudgetSaleMin = Me.TextBudgetSaleMin.Text
            vBudgetSaleMax = Me.TextBudgetSaleMax.Text
            vBudgetGPMin = Me.TextBudgetGPMin.Text
            vBudgetGPMax = Me.TextBudgetGPMax.Text

            vQuery = "exec dbo.USP_ICT_DeptBudgetPlanSet " & vSaleType & "," & vFiscalYear & "," & vPeriodOf4Week & ",'" & vDepartmentCode & "'," & vLastYearSale & "," & vLastYearGP & "," & vTargetSalePercent & "," & vTargetGPPercent & "," & vTargetSale & "," & vTargetGP & "," & vBudgetPercent & "," & vBudgetSaleMin & "," & vBudgetSaleMax & "," & vBudgetGPMin & "," & vBudgetGPMax & " "
            vCMD = New SqlCommand(vQuery, vConnection)
            vCMD.ExecuteNonQuery()

            For i = 0 To ListView102.Items.Count - 1
                vType = ListView102.Items(i).SubItems(0).Text
                vStep = ListView102.Items(i).SubItems(1).Text
                vPercentTarget = ListView102.Items(i).SubItems(2).Text
                vRate = ListView102.Items(i).SubItems(3).Text

                vQuery = "exec dbo.USP_ICT_StepSet " & vSaleType & "," & vFiscalYear & "," & vPeriodOf4Week & ",'" & vDepartmentCode & "'," & vType & "," & vStep & "," & vPercentTarget & "," & vRate & " "
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()
            Next

            'vQuery = "commit tran"
            'vCMD = New SqlCommand(vQuery, vConnection)
            'vCMD.ExecuteNonQuery()

            Call ClearData()
            Call RefreshBudgetTarget()
            MsgBox("บันทึกข้อมูล Budget Config เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information")
            GB101.Visible = False

            'Catch
            '    MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            '    vQuery = "rollback tran"
            '    vCMD = New SqlCommand(vQuery, vConnection)
            '    vCMD.ExecuteNonQuery()
            'End Try
        Else
            MsgBox("ไม่มีรายการกำหนด Step กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If

    End Sub
    Public Sub RefreshBudgetTarget()
        Dim vSaleType As Integer
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer
        Dim vListBudget As ListViewItem
        Dim i As Integer

        On Error GoTo ErrDescription


        vSaleType = Me.CMBSaleType.SelectedIndex
        vFiscalYear = Me.CMBFiscalYear.Text
        vPeriodOf4Week = Me.CMBPeriod.Text

        vQuery = "exec dbo.USP_ICT_DebtBudgetPlan '" & vSaleType & "'," & vFiscalYear & "," & vPeriodOf4Week & " "
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "DebtBudgetPlan")
        dt = ds.Tables("DebtBudgetPlan")
        ListView101.Items.Clear()
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                vListBudget = ListView101.Items.Add(Trim(dt.Rows(i).Item("departmentcode")))
                vListBudget.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("department"))
                vListBudget.SubItems.Add(2).Text = Format(dt.Rows(i).Item("targetsale"), "##,##0.00")
                vListBudget.SubItems.Add(3).Text = Format(dt.Rows(i).Item("targetgp"), "##,##0.00")
                vListBudget.SubItems.Add(4).Text = Format(dt.Rows(i).Item("budgetpercent"), "##,##0.00")
                vListBudget.SubItems.Add(5).Text = Format(dt.Rows(i).Item("budgetsalemin"), "##,##0.00")
                vListBudget.SubItems.Add(6).Text = Format(dt.Rows(i).Item("budgetsalemax"), "##,##0.00")
                vListBudget.SubItems.Add(7).Text = Format(dt.Rows(i).Item("budgetgpmin"), "##,##0.00")
                vListBudget.SubItems.Add(8).Text = Format(dt.Rows(i).Item("budgetgpmax"), "##,##0.00")
            Next
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TextTargetSale_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextTargetSale.KeyPress, TextTargetGP.KeyPress
        On Error GoTo ErrDescription

        Select Case Asc(e.KeyChar)
            Case 47 To 58, 8, 44, 46
            Case Else
                e.Handled = True
        End Select

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TextTargetSale_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextTargetSale.LostFocus, TextTargetGP.LostFocus
        On Error GoTo ErrDescription
        If Me.TextTargetSale.Text <> "" Then
            Me.TextTargetSale.Text = Format(Int(Me.TextTargetSale.Text), "##,##0.00")
        End If
        If Me.TextTargetGP.Text <> "" Then
            Me.TextTargetGP.Text = Format(Int(Me.TextTargetGP.Text), "##,##0.00")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TextTargetSale_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextTargetSale.TextChanged
        Dim vNTargetSale As Double
        Dim vTargetSale As Double
        Dim vLastYearSale As Double

        On Error GoTo ErrDescription

        If Me.TextLastYearSale.Text <> "" Then
            vTargetSale = Me.TextTargetSale.Text
            If Me.TextLastYearSale.Text <> "" Then
                vLastYearSale = Me.TextLastYearSale.Text
            Else
                vLastYearSale = 0
            End If
            vNTargetSale = (vTargetSale * 100) / vLastYearSale
            Me.NTargetSale.Value = Format(vNTargetSale, "##,##0.00")
        Else
            Me.TextTargetSale.Text = ""
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TextTargetGP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextTargetGP.TextChanged
        Dim vNTargetGP As Double
        Dim vTargetGP As Double
        Dim vLastYearGP As Double

        On Error GoTo ErrDescription

        If Me.TextLastYearGP.Text <> "" Then
            vTargetGP = Me.TextTargetGP.Text
            If Me.TextLastYearGP.Text <> "" Then
                vLastYearGP = Me.TextLastYearGP.Text
            Else
                vLastYearGP = 0
            End If
            vNTargetGP = (vTargetGP * 100) / vLastYearGP
            Me.NTargetGP.Value = vNTargetGP
        Else
            Me.TextTargetGP.Text = ""
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        Me.Close()
    End Sub

    Private Sub CMBSaleType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBSaleType.SelectedIndexChanged
        Dim vSaleType As Integer
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer
        Dim vListBudget As ListViewItem
        Dim i As Integer

        On Error GoTo ErrDescription

        If Me.CMBFiscalYear.Items.Count > 0 Then
            vSaleType = Me.CMBSaleType.SelectedIndex
            vFiscalYear = Me.CMBFiscalYear.Text
            vPeriodOf4Week = Me.CMBPeriod.Text

            vQuery = "exec dbo.USP_ICT_DebtBudgetPlan '" & vSaleType & "'," & vFiscalYear & "," & vPeriodOf4Week & " "
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "DebtBudgetPlan")
            dt = ds.Tables("DebtBudgetPlan")
            ListView101.Items.Clear()
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    vListBudget = ListView101.Items.Add(Trim(dt.Rows(i).Item("departmentcode")))
                    vListBudget.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("department"))
                    vListBudget.SubItems.Add(2).Text = Format(dt.Rows(i).Item("targetsale"), "##,##0.00")
                    vListBudget.SubItems.Add(3).Text = Format(dt.Rows(i).Item("targetgp"), "##,##0.00")
                    vListBudget.SubItems.Add(4).Text = Format(dt.Rows(i).Item("budgetpercent"), "##,##0.00")
                    vListBudget.SubItems.Add(5).Text = Format(dt.Rows(i).Item("budgetsalemin"), "##,##0.00")
                    vListBudget.SubItems.Add(6).Text = Format(dt.Rows(i).Item("budgetsalemax"), "##,##0.00")
                    vListBudget.SubItems.Add(7).Text = Format(dt.Rows(i).Item("budgetgpmin"), "##,##0.00")
                    vListBudget.SubItems.Add(8).Text = Format(dt.Rows(i).Item("budgetgpmax"), "##,##0.00")
                Next
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

End Class