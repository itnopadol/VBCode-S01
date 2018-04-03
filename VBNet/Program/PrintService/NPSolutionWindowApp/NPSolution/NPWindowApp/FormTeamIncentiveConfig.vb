Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Public Class FormTeamIncentiveConfig
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand
    Dim vReadQuery As SqlDataReader
    Dim vDepartmentClick As String
    Dim vTempTotalPercent As Double
    Dim vTempBudgetSaleMin As Double
    Dim vTempBudgetSaleMax As Double
    Dim vTempBudgetGPMin As Double
    Dim vTempBudgetGPMax As Double


    Private Sub FormTeamIncentiveConfig_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
                vListBudget.SubItems.Add(4).Text = Format(dt.Rows(i).Item("budgetsalemin"), "##,##0.00")
                vListBudget.SubItems.Add(5).Text = Format(dt.Rows(i).Item("budgetsalemax"), "##,##0.00")
                vListBudget.SubItems.Add(6).Text = Format(dt.Rows(i).Item("budgetgpmin"), "##,##0.00")
                vListBudget.SubItems.Add(7).Text = Format(dt.Rows(i).Item("budgetgpmax"), "##,##0.00")
                vListBudget.SubItems.Add(8).Text = Format(dt.Rows(i).Item("budgetremain"), "##,##0.00")
            Next
        End If


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub ListView101_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView101.DoubleClick
        Dim vSaleType As Integer
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer
        Dim vListBudget As ListViewItem
        Dim i As Integer
        Dim vIndex As Integer

        On Error GoTo ErrDescription

        vIndex = ListView101.SelectedItems(0).Index

        If ListView101.Items(vIndex).SubItems(3).Text <> 0 And ListView101.Items(vIndex).SubItems(4).Text <> 0 And ListView101.Items(vIndex).SubItems(5).Text <> 0 Then

            vSaleType = Me.CMBSaleType.SelectedIndex
            vFiscalYear = Me.CMBFiscalYear.Text
            vPeriodOf4Week = Me.CMBPeriod.Text
            vDepartmentClick = ListView101.SelectedItems(0).SubItems(0).Text
            vTempBudgetSaleMin = ListView101.SelectedItems(0).SubItems(4).Text
            vTempBudgetSaleMax = ListView101.SelectedItems(0).SubItems(5).Text
            vTempBudgetGPMin = ListView101.SelectedItems(0).SubItems(6).Text
            vTempBudgetGPMax = ListView101.SelectedItems(0).SubItems(7).Text
            vQuery = "exec dbo.USP_ICT_TeamBudgetDept '" & vDepartmentClick & "','" & vSaleType & "'," & vFiscalYear & "," & vPeriodOf4Week & " "
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "TeamBudgetDept")
            dt = ds.Tables("TeamBudgetDept")
            ListView102.Items.Clear()
            vTempTotalPercent = 0
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    vListBudget = ListView102.Items.Add(Trim(dt.Rows(i).Item("team")))
                    vListBudget.SubItems.Add(1).Text = Format(dt.Rows(i).Item("budgetpercent"), "##,##0.00")
                    vListBudget.SubItems.Add(2).Text = Format(dt.Rows(i).Item("budgetsalemin"), "##,##0.00")
                    vListBudget.SubItems.Add(3).Text = Format(dt.Rows(i).Item("budgetsalemax"), "##,##0.00")
                    vListBudget.SubItems.Add(4).Text = Format(dt.Rows(i).Item("budgetgpmin"), "##,##0.00")
                    vListBudget.SubItems.Add(5).Text = Format(dt.Rows(i).Item("budgetgpmax"), "##,##0.00")
                    vListBudget.SubItems.Add(6).Text = (dt.Rows(i).Item("id"))
                    vTempTotalPercent = vTempTotalPercent + dt.Rows(i).Item("budgetpercent")
                Next
            End If
            GB101.Visible = True
            Me.LBLDepartment.Text = Trim(ListView101.SelectedItems(0).SubItems(1).Text)
            Me.NTeamBudget.Value = 0.0
            Me.TextTotalPercent.Text = Format(vTempTotalPercent, "##,##0.00")
            vQuery = "select * from  dbo.VW_ICT_TeamList"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Team")
            dt = ds.Tables("Team")
            Me.CMBTeam.Items.Clear()
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    CMBTeam.Items.Add(dt.Rows(i).Item("team"))
                Next
            End If
        Else
            MsgBox("ไม่สามารถกำหนด Budget ของทีมได้ เนื่องจากยังไม่ได้กำหนด Budget และ Target ของ Department", MsgBoxStyle.Critical, "Send Error")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub


    Private Sub BTNCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCancel.Click
        GB101.Visible = False
    End Sub

    Private Sub BTNInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNInsert.Click
        Dim i As Integer
        Dim vCheckTeam As String
        Dim vTeam As String
        Dim vListItem As ListViewItem
        Dim vAddPercent As Double

        On Error GoTo ErrDescription

        If vTempTotalPercent < 100 Then
            vAddPercent = Me.NTeamBudget.Value
            If vTempTotalPercent + vAddPercent > 100 Then
                MsgBox("ไม่สามารถกำหนดทีมเพิ่มได้ เนื่องจาก % ของทีมใหม่รวมกับ % เดิมเกิน 100 % แล้ว กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error")
                Exit Sub
            End If
            If Me.NTeamBudget.Value > 0.0 Then
                vTeam = Trim(Me.CMBTeam.Text)
                For i = 0 To ListView102.Items.Count - 1
                    vCheckTeam = Trim(ListView102.Items(i).SubItems(0).Text)
                    If vTeam = vCheckTeam Then
                        MsgBox("มีข้อมูลการกำหนด TeamBudget ของทีม " & vTeam & " นี้อยู่แล้วในรายการที่ " & i + 1 & " กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
                        Exit Sub
                    End If
                Next

                vListItem = ListView102.Items.Add(Trim(Me.CMBTeam.Text))
                vListItem.SubItems.Add(1).Text = Format(Me.NTeamBudget.Value, "##,##0.00")
                vListItem.SubItems.Add(2).Text = Format(((Me.NTeamBudget.Value * vTempBudgetSaleMin) / 100), "##,##0.00")
                vListItem.SubItems.Add(3).Text = Format(((Me.NTeamBudget.Value * vTempBudgetSaleMax) / 100), "##,##0.00")
                vListItem.SubItems.Add(4).Text = Format(((Me.NTeamBudget.Value * vTempBudgetGPMin) / 100), "##,##0.00")
                vListItem.SubItems.Add(5).Text = Format(((Me.NTeamBudget.Value * vTempBudgetGPMax) / 100), "##,##0.00")
                vTempTotalPercent = vTempTotalPercent + vAddPercent
                Me.TextTotalPercent.Text = Format(vTempTotalPercent, "##,##0.00")
                vTempTotalPercent = Me.TextTotalPercent.Text
                Me.NTeamBudget.Value = 0.0
                Me.CMBTeam.Text = ""
            End If
        Else
            MsgBox("ไม่สามารถกำหนดทีมเพิ่มได้ เนื่องจาก % ของทีมครบ 100 % เรียบร้อยแล้ว", MsgBoxStyle.Critical, "Send Error")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim i As Integer
        Dim vTeamCode As String
        Dim vDepartmentCode As String
        Dim vSaleType As Integer
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer
        Dim vBudgetPercent As Double
        Dim vBudgetSaleMin As Double
        Dim vBudgetSaleMax As Double
        Dim vBudgetGPMin As Double
        Dim vBudgetGPMax As Double

        If ListView102.Items.Count > 0 Then

            vSaleType = Me.CMBSaleType.SelectedIndex
            vFiscalYear = Me.CMBFiscalYear.Text
            vPeriodOf4Week = Me.CMBPeriod.Text
            vDepartmentCode = vDepartmentClick
            Try
                vQuery = "begin tran"
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()

                vQuery = "exec dbo.USP_ICT_IncentiveSetClear '" & vDepartmentCode & "'," & vSaleType & "," & vFiscalYear & "," & vPeriodOf4Week & " "
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()

                For i = 0 To ListView102.Items.Count - 1
                    vTeamCode = Microsoft.VisualBasic.Left(ListView102.Items(i).SubItems(0).Text, InStr(ListView102.Items(i).SubItems(0).Text, "/") - 1)
                    vBudgetPercent = ListView102.Items(i).SubItems(1).Text
                    vBudgetSaleMin = ListView102.Items(i).SubItems(2).Text
                    vBudgetSaleMax = ListView102.Items(i).SubItems(3).Text
                    vBudgetGPMin = ListView102.Items(i).SubItems(4).Text
                    vBudgetGPMax = ListView102.Items(i).SubItems(5).Text

                    vQuery = "exec dbo.USP_ICT_IncentiveSet '" & vTeamCode & "','" & vDepartmentCode & "'," & vSaleType & "," & vFiscalYear & "," & vPeriodOf4Week & "," & vBudgetPercent & "," & vBudgetSaleMin & "," & vBudgetSaleMax & "," & vBudgetGPMin & "," & vBudgetGPMax & " "
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vCMD.ExecuteNonQuery()
                Next i

                vQuery = "exec dbo.USP_ICT_QueueCalBudgetRemain '" & vDepartmentCode & "'," & vSaleType & "," & vFiscalYear & "," & vPeriodOf4Week & " "
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()

                vQuery = "commit tran"
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()
                MsgBox("บันทึกข้อมูล แผนก " & Me.LBLDepartment.Text & " เรียบร้อยแล้ว ", MsgBoxStyle.Information, "Send Information")
                Me.GB101.Visible = False
            Catch
                MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
                vQuery = "rollback tran"
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()
            End Try
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
                    vListBudget.SubItems.Add(4).Text = Format(dt.Rows(i).Item("budgetsalemin"), "##,##0.00")
                    vListBudget.SubItems.Add(5).Text = Format(dt.Rows(i).Item("budgetsalemax"), "##,##0.00")
                    vListBudget.SubItems.Add(6).Text = Format(dt.Rows(i).Item("budgetgpmin"), "##,##0.00")
                    vListBudget.SubItems.Add(7).Text = Format(dt.Rows(i).Item("budgetgpmax"), "##,##0.00")
                    vListBudget.SubItems.Add(8).Text = Format(dt.Rows(i).Item("budgetremain"), "##,##0.00")
                Next
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub ListView102_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView102.KeyDown
        Dim i As Integer
        Dim vBudgetPercent As Double
        Dim vID As Integer
        Dim vCountRef As Integer
        Dim vDepartmentCode As String
        Dim vSaleType As Integer
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer

        'On Error GoTo ErrDescription

        If ListView101.Items.Count > 0 Then

            If e.KeyCode = Keys.Delete Then
                If MessageBox.Show("คุณต้องการลบกรายการนี้ใช่หรือไม่", "Send Question ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                    i = ListView102.SelectedItems(0).Index
                    vBudgetPercent = ListView102.Items(i).SubItems(1).Text
                    vID = ListView102.Items(i).SubItems(6).Text

                    vQuery = "select count(teambudgetdebtid) as vCount from npmaster.dbo.tb_ict_requestsub where teambudgetdebtid = " & vID & " and iscancel = 0 "
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vReadQuery = vCMD.ExecuteReader
                    While vReadQuery.Read
                        vCountRef = vReadQuery(0)
                    End While
                    vReadQuery.Close()

                    If vCountRef = 0 Then
                        vTempTotalPercent = Me.TextTotalPercent.Text - vBudgetPercent

                        Me.TextTotalPercent.Text = Format(vTempTotalPercent, "##,##0.00")

                        If Me.ListView102.Items.Count = 1 Then

                            Try
                                vSaleType = Me.CMBSaleType.SelectedIndex
                                vFiscalYear = Me.CMBFiscalYear.Text
                                vPeriodOf4Week = Me.CMBPeriod.Text
                                vDepartmentCode = vDepartmentClick
                                vQuery = "begin tran"
                                vCMD = New SqlCommand(vQuery, vConnection)
                                vCMD.ExecuteNonQuery()

                                vQuery = "exec dbo.USP_ICT_IncentiveSetClear '" & vDepartmentCode & "'," & vSaleType & "," & vFiscalYear & "," & vPeriodOf4Week & " "
                                vCMD = New SqlCommand(vQuery, vConnection)
                                vCMD.ExecuteNonQuery()

                                vQuery = "exec dbo.USP_ICT_QueueCalBudgetRemain '" & vDepartmentCode & "'," & vSaleType & "," & vFiscalYear & "," & vPeriodOf4Week & " "
                                vCMD = New SqlCommand(vQuery, vConnection)
                                vCMD.ExecuteNonQuery()

                                vQuery = "commit tran"
                                vCMD = New SqlCommand(vQuery, vConnection)
                                vCMD.ExecuteNonQuery()
                                Me.GB101.Visible = False

                            Catch ex As Exception
                                MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
                                vQuery = "rollback tran"
                                vCMD = New SqlCommand(vQuery, vConnection)
                                vCMD.ExecuteNonQuery()
                            End Try

                        End If
                        ListView102.Items.RemoveAt(i)

                    Else
                        MsgBox("ไม่สามารถลบรายการนี้ได้ เนื่องจากถูกอ้างไปเสนอจ่ายเรียบร้อยแล้ว", MsgBoxStyle.Critical, "Send Error")
                    End If



                End If
            End If
        End If


        'ErrDescription:
        '        If Err.Description <> "" Then
        '            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
        '            Exit Sub
        '        End If
    End Sub

    Private Sub ListView102_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView102.SelectedIndexChanged

    End Sub
End Class