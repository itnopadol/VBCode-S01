Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Public Class FormSaleHistory
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand
    Private Sub FormSaleHistory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.CBPeriod.Checked = True

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
                vListBudget.SubItems.Add(2).Text = dt.Rows(i).Item("iscapturehistory")
                ListView101.Items(i).Checked = True
            Next
        End If
        Me.CBPeriod.Checked = True
        Me.PGBar101.Value = 0

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub CBPeriod_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBPeriod.CheckedChanged
        Dim i As Integer

        On Error GoTo ErrDescription

        If CBPeriod.Checked = True Then
            For i = 0 To ListView101.Items.Count - 1
                ListView101.Items(i).Checked = True
            Next
            Me.CBPeriod.Checked = True
        End If

        If CBPeriod.Checked = False Then
            For i = 0 To ListView101.Items.Count - 1
                ListView101.Items(i).Checked = False
            Next

        End If


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim i As Integer
        Dim vSaleType As Integer
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer
        Dim vDepartmentCode As String

        On Error GoTo ErrDescription


        vSaleType = Me.CMBSaleType.SelectedIndex
        vFiscalYear = Me.CMBFiscalYear.Text
        vPeriodOf4Week = Me.CMBPeriod.Text

        PGBar101.Maximum = ListView101.Items.Count
        For i = 0 To ListView101.Items.Count - 1
            If ListView101.Items(i).Checked = True Then
                vDepartmentCode = ListView101.Items(i).SubItems(0).Text
                vQuery = "exec dbo.USP_ICT_CaptureProductSale '" & vDepartmentCode & "'," & vSaleType & "," & vFiscalYear & "," & vPeriodOf4Week & " "
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()
            End If
            PGBar101.Value = i + 1
        Next

        Me.PGBar101.Value = 0
        ListView101.Items.Clear()

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
                    vListBudget.SubItems.Add(2).Text = dt.Rows(i).Item("iscapturehistory")
                    ListView101.Items(i).Checked = True
                Next
            End If
            Me.CBPeriod.Checked = True
            Me.PGBar101.Value = 0
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub
End Class