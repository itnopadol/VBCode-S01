Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class FormBudgetConfig
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand

    Private Sub FormBudgetConfig_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Dim vKPIType As String
        Dim vPeriodOf4Week As Integer
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



        If Me.CMBSaleType.Text <> "" And Me.CMBPeriod.Text <> "" Then
            vKPIType = Me.CMBSaleType.SelectedIndex
            vFisCalYear = Int(Me.CMBFiscalYear.Text)
            vPeriodOf4Week = Int(Me.CMBPeriod.Text)
            vQuery = "exec dbo.USP_ICT_SearchOfBudgetConfig " & vKPIType & "," & vFisCalYear & "," & vPeriodOf4Week & " "
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Budget")
            dt = ds.Tables("Budget")

            If dt.Rows.Count > 0 Then
                Me.TextSaleMin.Text = Format(dt.Rows(0).Item("budgetsalemin"), "##,##0.00")
                Me.TextSaleMax.Text = Format(dt.Rows(0).Item("budgetsalemax"), "##,##0.00")
                Me.TextGPMin.Text = Format(dt.Rows(0).Item("budgetgpmin"), "##,##0.00")
                Me.TextGPMax.Text = Format(dt.Rows(0).Item("budgetgpmax"), "##,##0.00")
                Me.NUDReturnItem.Value = dt.Rows(0).Item("returnday")
            Else
                Me.TextSaleMin.Text = ""
                Me.TextSaleMax.Text = ""
                Me.TextGPMin.Text = ""
                Me.TextGPMax.Text = ""
                Me.NUDReturnItem.Value = 28
            End If
        End If




ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vKPIType As String
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer
        Dim vReturnDay As Integer
        Dim vBudgetSaleMin As Object
        Dim vBudgetSaleMax As Object
        Dim vBudgetGPMin As Object
        Dim vBudgetGPMax As Object

        If Me.TextSaleMin.Text <> "" And Me.TextSaleMax.Text <> "" And Me.TextGPMin.Text <> "" And Me.TextGPMax.Text <> "" Then

            vKPIType = Me.CMBSaleType.SelectedIndex
            vFiscalYear = Int(Me.CMBFiscalYear.Text)
            vPeriodOf4Week = Int(Me.CMBPeriod.Text)
            vReturnDay = Int(Me.NUDReturnItem.Value)
            vBudgetSaleMin = Int(Me.TextSaleMin.Text)
            vBudgetSaleMax = Int(Me.TextSaleMax.Text)
            vBudgetGPMin = Int(Me.TextGPMin.Text)
            vBudgetGPMax = Int(Me.TextGPMax.Text)

            Try

                vQuery = "begin tran"
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()

                vQuery = "exec dbo.USP_ICT_UpdateBudgetConfig '" & vKPIType & "'," & vFiscalYear & "," & vPeriodOf4Week & "," & vReturnDay & "," & vBudgetSaleMin & "," & vBudgetSaleMax & "," & vBudgetGPMin & "," & vBudgetGPMax & " "
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()

                vQuery = "commit tran"
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()

                Call ClearData()
                MsgBox("บันทึกข้อมูล Budget Config เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information")

            Catch
                MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
                vQuery = "rollback tran"
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()
            End Try

        Else
            MsgBox("กรุณากรอกข้อมูลเกี่ยวกับ Budget ให้ครบด้วย", MsgBoxStyle.Critical, "Send Error ")
        End If
    End Sub

    Private Sub TextSaleMin_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextSaleMin.KeyPress, TextSaleMax.KeyPress, TextGPMin.KeyPress, TextGPMax.KeyPress
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

    Private Sub TextSaleMin_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextSaleMin.LostFocus, TextSaleMax.LostFocus, TextGPMin.LostFocus, TextGPMax.LostFocus
        On Error GoTo ErrDescription

        If Me.TextSaleMin.Text <> "" Then
            Me.TextSaleMin.Text = Format(Int(Me.TextSaleMin.Text), "##,##0.00")
        End If
        If Me.TextSaleMax.Text <> "" Then
            Me.TextSaleMax.Text = Format(Int(Me.TextSaleMax.Text), "##,##0.00")
        End If
        If Me.TextGPMin.Text <> "" Then
            Me.TextGPMin.Text = Format(Int(Me.TextGPMin.Text), "##,##0.00")
        End If
        If Me.TextGPMax.Text <> "" Then
            Me.TextGPMax.Text = Format(Int(Me.TextGPMax.Text), "##,##0.00")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub ClearData()
        Me.TextSaleMin.Text = ""
        Me.TextSaleMax.Text = ""
        Me.TextGPMin.Text = ""
        Me.TextGPMax.Text = ""
        Me.NUDReturnItem.Value = 28
        Me.CMBFiscalYear.Text = Now.Year
        Me.CMBSaleType.Text = Me.CMBSaleType.Items(0)
        Me.CMBPeriod.Text = Now.Month
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Me.Close()
    End Sub

    Private Sub CMBPeriod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBPeriod.SelectedIndexChanged
        Dim vKPIType As String
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer

        On Error GoTo ErrDescription

        If Me.CMBSaleType.Text <> "" And Me.CMBFiscalYear.Text <> "" Then
            vKPIType = Me.CMBSaleType.SelectedIndex
            vFiscalYear = Int(Me.CMBFiscalYear.Text)
            vPeriodOf4Week = Int(Me.CMBPeriod.Text)
            vQuery = "exec dbo.USP_ICT_SearchOfBudgetConfig " & vKPIType & "," & vFiscalYear & "," & vPeriodOf4Week & " "
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Budget")
            dt = ds.Tables("Budget")

            If dt.Rows.Count > 0 Then
                Me.TextSaleMin.Text = Format(dt.Rows(0).Item("budgetsalemin"), "##,##0.00")
                Me.TextSaleMax.Text = Format(dt.Rows(0).Item("budgetsalemax"), "##,##0.00")
                Me.TextGPMin.Text = Format(dt.Rows(0).Item("budgetgpmin"), "##,##0.00")
                Me.TextGPMax.Text = Format(dt.Rows(0).Item("budgetgpmax"), "##,##0.00")
                Me.NUDReturnItem.Value = dt.Rows(0).Item("returnday")
            Else
                Me.TextSaleMin.Text = ""
                Me.TextSaleMax.Text = ""
                Me.TextGPMin.Text = ""
                Me.TextGPMax.Text = ""
                Me.NUDReturnItem.Value = 28
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub CMBSaleType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBSaleType.SelectedIndexChanged
        Dim vKPIType As String
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer

        On Error GoTo ErrDescription

        If Me.CMBFiscalYear.Text <> "" And Me.CMBPeriod.Text <> "" Then
            vKPIType = Me.CMBSaleType.SelectedIndex
            vFiscalYear = Int(Me.CMBFiscalYear.Text)
            vPeriodOf4Week = Int(Me.CMBPeriod.Text)
            vQuery = "exec dbo.USP_ICT_SearchOfBudgetConfig " & vKPIType & "," & vFiscalYear & "," & vPeriodOf4Week & " "
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Budget")
            dt = ds.Tables("Budget")

            If dt.Rows.Count > 0 Then
                Me.TextSaleMin.Text = Format(dt.Rows(0).Item("budgetsalemin"), "##,##0.00")
                Me.TextSaleMax.Text = Format(dt.Rows(0).Item("budgetsalemax"), "##,##0.00")
                Me.TextGPMin.Text = Format(dt.Rows(0).Item("budgetgpmin"), "##,##0.00")
                Me.TextGPMax.Text = Format(dt.Rows(0).Item("budgetgpmax"), "##,##0.00")
                Me.NUDReturnItem.Value = dt.Rows(0).Item("returnday")
            Else
                Me.TextSaleMin.Text = ""
                Me.TextSaleMax.Text = ""
                Me.TextGPMin.Text = ""
                Me.TextGPMax.Text = ""
                Me.NUDReturnItem.Value = 28
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub
End Class