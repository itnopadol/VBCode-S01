Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class FormPayIncentive
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand
    Dim vTotalAmount As Double
    Dim vReadQuery As SqlDataReader
    Dim vIsOpen As Integer
    Dim vIsConfirm As Integer
    Dim frmIncentiveRequest As FormIncentiveRequest
    Dim frmIncentiveRequestDetails As FormIncentiveRequestDetails
    Dim vIsCancel As Integer
    Dim vCreatorCode As String


    Private Sub FormPayIncentive_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrDescription

        Call InitializeDataBase()

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

        vQuery = "exec dbo.USP_ICT_RequestSelectList '" & vSaleType & "'," & vFiscalYear & "," & vPeriodOf4Week & " "
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "RequestSelectList")
        dt = ds.Tables("RequestSelectList")
        ListView102.Items.Clear()
        If dt.Rows.Count > 0 Then

            For i = 0 To dt.Rows.Count - 1
                vListBudget = ListView102.Items.Add("")
                vListBudget.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("departmentcode"))
                vListBudget.SubItems.Add(2).Text = dt.Rows(i).Item("departmentname")
                vListBudget.SubItems.Add(2).Text = dt.Rows(i).Item("team")
                vListBudget.SubItems.Add(2).Text = Format(dt.Rows(i).Item("dept_incentive"), "##,##0.00")
                vListBudget.SubItems.Add(3).Text = Format(dt.Rows(i).Item("team_incentive"), "##,##0.00")
                vListBudget.SubItems.Add(3).Text = Format(dt.Rows(i).Item("requestremain"), "##,##0.00")
                vListBudget.SubItems.Add(3).Text = dt.Rows(i).Item("team_isprocess")
                vListBudget.SubItems.Add(3).Text = dt.Rows(i).Item("team_processdate")
            Next

            Me.CH101.Checked = True
        Else
            Me.CH101.Checked = False
        End If


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If

    End Sub

    Private Sub BTNAddList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddList.Click
        Dim i As Integer

        On Error GoTo ErrDescription

        Me.CMBSaleType.Text = Me.CMBSaleType.Items(0)
        vQuery = "exec dbo.USP_ICT_FiscalYearList"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "FiscalYearList")
        dt = ds.Tables("FiscalYearList")
        Me.CMBFiscalYear.Items.Clear()
        For i = 0 To dt.Rows.Count - 1
            Me.CMBFiscalYear.Items.Add(dt.Rows(i).Item("fiscalyear"))
        Next
        Me.CMBFiscalYear.Text = Now.Year
        Me.DocDate.Text = Now.Date
        GB101.Visible = True

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNGen.Click
        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_ICT_RequestNewDoc"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then
            Me.TextDocNo.Text = Trim(dt.Rows(0).Item("NewDoc"))
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNFind.Click
        Dim vListItem As ListViewItem
        Dim i As Integer

        On Error GoTo ErrDescription


        ListView103.Items.Clear()
        GB102.Visible = True
        vQuery = "exec dbo.USP_ICT_RequestMasterList "
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Search")
        dt = ds.Tables("Search")

        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                vListItem = ListView103.Items.Add(dt.Rows(i).Item("docno"))
                vListItem.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("docdate"))
                vListItem.SubItems.Add(2).Text = Format(dt.Rows(i).Item("totalamount"), "##,##0.00")
                vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("isconfirm")
            Next
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        GB102.Visible = False
    End Sub

    Private Sub TextSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextSearchDocNo.KeyDown
        Dim vSearchDocNo As String
        Dim vListItem As ListViewItem
        Dim i As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.TextSearchDocNo.Text <> "" Then
                vSearchDocNo = Trim(Me.TextSearchDocNo.Text)
                vQuery = "exec dbo.USP_ICT_RequestMasterList '" & vSearchDocNo & "' "
            Else
                vQuery = "exec dbo.USP_ICT_RequestMasterList "
            End If

            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Search")
            dt = ds.Tables("Search")

            ListView103.Items.Clear()

            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    vListItem = ListView103.Items.Add(dt.Rows(i).Item("docno"))
                    vListItem.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("docdate"))
                    vListItem.SubItems.Add(2).Text = Format(dt.Rows(i).Item("totalamount"), "##,##0.00")
                    vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("isconfirm")
                Next
            End If
        End If


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub


    Private Sub BTNBasketSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNBasketSelect.Click
        Dim vSaleType As String
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer
        Dim vListSelect As ListViewItem
        Dim i As Integer
        Dim n As Integer
        Dim vCheckSaleType As String
        Dim vCheckFiscalYear As Integer
        Dim vCheckPeriod As Integer
        Dim vCheckDepartmentCode As String
        Dim vCheckTeam As String
        Dim vCheckAdd As Integer
        Dim vDepartmentCode As String
        Dim vTeam As String

        On Error GoTo ErrDescription

        vSaleType = Me.CMBSaleType.Text
        vFiscalYear = Me.CMBFiscalYear.Text
        vPeriodOf4Week = Me.CMBPeriod.Text

        For n = 0 To ListView102.Items.Count - 1
            If ListView102.Items(n).Checked = True Then

                If ListView101.Items.Count > 0 Then
                    i = ListView101.Items.Count

                    For i = 0 To ListView101.Items.Count - 1
                        vCheckSaleType = Trim(ListView101.Items(i).SubItems(0).Text)
                        vCheckFiscalYear = Trim(ListView101.Items(i).SubItems(1).Text)
                        vCheckPeriod = Trim(ListView101.Items(i).SubItems(2).Text)
                        vCheckDepartmentCode = Trim(ListView101.Items(i).SubItems(3).Text)
                        vCheckTeam = Trim(ListView101.Items(i).SubItems(5).Text)
                        vDepartmentCode = Trim(ListView102.Items(n).SubItems(1).Text)
                        vTeam = Trim(ListView102.Items(n).SubItems(3).Text)

                        If vCheckSaleType = vSaleType And vCheckFiscalYear = vFiscalYear And vCheckPeriod = vPeriodOf4Week And vCheckDepartmentCode = vDepartmentCode And vCheckTeam = vTeam Then
                            MsgBox("ไม่สามารถเพิ่มรายการได้ เนื่องจากมีรายการซ้ำกันเกิดขึ้น", MsgBoxStyle.Critical, "Send Error ")
                            vCheckAdd = 1
                            Exit For
                        Else
                            vCheckAdd = 0
                        End If

                    Next
                End If

                If vCheckAdd = 0 Then
                    vTotalAmount = Me.LBLTotalAmount.Text
                    vListSelect = ListView101.Items.Add(Trim(vSaleType))
                    vListSelect.SubItems.Add(1).Text = vFiscalYear
                    vListSelect.SubItems.Add(2).Text = vPeriodOf4Week
                    vListSelect.SubItems.Add(3).Text = Trim(ListView102.Items(n).SubItems(1).Text)
                    vListSelect.SubItems.Add(4).Text = Trim(ListView102.Items(n).SubItems(2).Text)
                    vListSelect.SubItems.Add(5).Text = Trim(ListView102.Items(n).SubItems(3).Text)
                    vListSelect.SubItems.Add(3).Text = Trim(ListView102.Items(n).SubItems(5).Text)
                    vTotalAmount = vTotalAmount + ListView102.Items(n).SubItems(5).Text
                    Me.LBLTotalAmount.Text = Format(vTotalAmount, "##,##0.00")
                End If
            End If
        Next
        GB101.Visible = False
        ListView102.Items.Clear()

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub ListView101_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView101.KeyDown
        Dim i As Integer
        Dim vIncentive As Double

        On Error GoTo ErrDescription

        If Me.PB101.Visible = True Then
            If ListView101.Items.Count > 0 Then
                If e.KeyCode = Keys.Delete Then
                    If MessageBox.Show("คุณต้องการลบกรายการนี้ใช่หรือไม่", "Send Question ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                        i = ListView101.SelectedItems(0).Index
                        vIncentive = ListView101.Items(i).SubItems(6).Text
                        vTotalAmount = Me.LBLTotalAmount.Text - vIncentive

                        ListView101.Items.RemoveAt(i)
                        Me.LBLTotalAmount.Text = Format(vTotalAmount, "##,##0.00")
                    End If
                End If
            End If
        End If


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vSaleType As Integer
        Dim vFiscalYear As Integer
        Dim vPeriodOf4Week As Integer
        Dim vDepartmentCode As String
        Dim vTeamCode As String
        Dim vDocNo As String
        Dim vIncentive As Double
        Dim i As Integer
        Dim vIncentiveTotalAmount As Double
        Dim vDocDate As String
        Dim vInsertType As Integer
        Dim vExist As Integer
        Dim vCheckDocno As Integer

        If vIsConfirm = 0 And vIsCancel = 0 Then
            If Me.TextDocNo.Text <> "" And ListView101.Items.Count > 0 Then
                vDocNo = Trim(Me.TextDocNo.Text)

                If vIsOpen = 0 Then
                    vQuery = "select count(docno) as vCount from npmaster.dbo.TB_ICT_Request where docno = '" & vDocNo & "'"
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vReadQuery = vCMD.ExecuteReader()
                    While vReadQuery.Read
                        vCheckDocno = vReadQuery(0)
                    End While
                    vReadQuery.Close()
                End If

                If vCheckDocno = 1 Then
                    MsgBox("มีเลขที่เอกสารเลขที่ " & vDocNo & " นี้อยู่แล้วในระบบ กรุณากดรันเลขที่เอกสารใหม่อีกครั้ง", MsgBoxStyle.Critical, "Send Error")
                    Exit Sub
                End If

                vQuery = "select count(docno) as vCount from npmaster.dbo.TB_ICT_Request where docno = '" & vDocNo & "'"
                vCMD = New SqlCommand(vQuery, vConnection)
                vReadQuery = vCMD.ExecuteReader()
                While vReadQuery.Read
                    vExist = vReadQuery(0)
                End While
                vReadQuery.Close()

                If vExist = 1 Then
                    vInsertType = 0
                Else
                    vInsertType = 1
                End If
                vDocDate = Me.DocDate.Text
                vIncentiveTotalAmount = Me.LBLTotalAmount.Text

                Try
                    vQuery = "begin tran"
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vCMD.ExecuteNonQuery()

                    vQuery = "exec dbo.USP_ICT_RequestHeadUpdate " & vInsertType & ",'" & vDocNo & "','" & vDocDate & "'," & vIncentiveTotalAmount & " "
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vCMD.ExecuteNonQuery()

                    vQuery = "exec dbo.USP_ICT_RequestDetailClear '" & vDocNo & "' "
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vCMD.ExecuteNonQuery()

                    For i = 0 To ListView101.Items.Count - 1

                        Select Case Trim(ListView101.Items(i).SubItems(0).Text)
                            Case Trim("ขายเงินเชื่อ")
                                vSaleType = 0
                            Case Trim("ขายเงินสด")
                                vSaleType = 1
                        End Select
                        vFiscalYear = Trim(ListView101.Items(i).SubItems(1).Text)
                        vPeriodOf4Week = Trim(ListView101.Items(i).SubItems(2).Text)
                        vDepartmentCode = Trim(ListView101.Items(i).SubItems(3).Text)
                        vTeamCode = Microsoft.VisualBasic.Left(ListView101.Items(i).SubItems(5).Text, InStr(ListView101.Items(i).SubItems(5).Text, "/") - 1)
                        vIncentive = ListView101.Items(i).SubItems(6).Text

                        vQuery = "exec dbo.USP_ICT_RequestDetailInsert '" & vDocNo & "'," & vSaleType & "," & vFiscalYear & "," & vPeriodOf4Week & ",'" & vDepartmentCode & "','" & vTeamCode & "'," & vIncentive & " "
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vCMD.ExecuteNonQuery()
                    Next

                    vQuery = "commit tran"
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vCMD.ExecuteNonQuery()

                    MsgBox("บันทึกเลขที่เอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Error")
                    Call ClearData()
                Catch ex As Exception
                    MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
                    vQuery = "rollback tran"
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vCMD.ExecuteNonQuery()
                End Try
            Else
                MsgBox("กรุณากรอกเลขที่เอกสารด้วย", MsgBoxStyle.Critical, "Send Error ")
            End If
        End If
    End Sub
    Private Sub ClearData()
        On Error Resume Next

        Me.TextDocNo.Text = ""
        Me.DocDate.Text = Now.Date
        Me.ListView101.Items.Clear()
        Me.LBLTotalAmount.Text = "0.00"
        Me.PB101.Visible = True
        Me.PB102.Visible = False
        Me.PB103.Visible = False
        vIsOpen = 0
        vIsConfirm = 0
        vIsCancel = 0
    End Sub

    Private Sub ListView103_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView103.DoubleClick
        Dim vDocno As String
        Dim i As Integer
        Dim vListItem As ListViewItem

        On Error GoTo ErrDescription

        If ListView103.Items.Count > 0 Then
            Call ClearData()
            vIsOpen = 1
            i = ListView103.SelectedItems(0).Index
            vDocno = ListView103.Items(i).SubItems(0).Text

            vQuery = "exec dbo.USP_ICT_RequestList '" & vDocno & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Docno")
            dt = ds.Tables("Docno")

            If dt.Rows.Count > 0 Then
                vIsConfirm = dt.Rows(0).Item("isconfirm")
                vIsCancel = dt.Rows(0).Item("iscancel")
                vCreatorCode = dt.Rows(0).Item("creatorcode")
                Me.TextDocNo.Text = Trim(dt.Rows(0).Item("docno"))
                Me.DocDate.Text = Trim(dt.Rows(0).Item("docdate"))
                Me.LBLTotalAmount.Text = Format(dt.Rows(0).Item("totalamount"), "##,##0.00")

                For i = 0 To dt.Rows.Count - 1
                    If dt.Rows(i).Item("kpitype") = 0 Then
                        vListItem = ListView101.Items.Add("ขายเงินเชื่อ")
                    Else
                        vListItem = ListView101.Items.Add("ขายเงินสด")
                    End If
                    vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("fiscalyear")
                    vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("period")
                    vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("departmentcode")
                    vListItem.SubItems.Add(4).Text = dt.Rows(i).Item("departmentname")
                    vListItem.SubItems.Add(5).Text = dt.Rows(i).Item("team")
                    vListItem.SubItems.Add(6).Text = dt.Rows(i).Item("paid")
                Next
            End If

            If vIsConfirm = 0 And vIsCancel = 0 Then
                Me.PB101.Visible = True
                Me.PB102.Visible = False
                Me.PB103.Visible = False
            ElseIf vIsConfirm = 1 And vIsCancel = 0 Then
                Me.PB101.Visible = False
                Me.PB102.Visible = True
                Me.PB103.Visible = False
            ElseIf vIsCancel = 1 Then
                Me.PB101.Visible = False
                Me.PB102.Visible = False
                Me.PB103.Visible = True
            End If

            GB102.Visible = False
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClear.Click
        Call ClearData()
    End Sub

    Private Sub TextSearchDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextSearchDocNo.TextChanged

    End Sub

    Private Sub BTNConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNConfirm.Click
        Dim vDocNo As String

        On Error GoTo ErrDescription

        If vIsOpen = 1 Then
            vDocNo = Trim(Me.TextDocNo.Text)
            If vIsCancel = 0 Then
                Call ChekAuthorityAccess()

                If vDepartment = "MG" And vLevelID = 0 Then

                    If MessageBox.Show("คุณต้องการอนุมัติเอกสาร ใช่หรือไม่", "ข้อความสอบถาม", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                        If vIsConfirm = 1 Then
                            MsgBox("ไม่สามารถอนุมัติเลขที่เอกสาร " & vDocNo & " ได้เนื่องจากได้ถูกอนุมัติเรียบร้อยแล้ว กรุณาตรวจสอบเลขที่เอกสารด้วย ", MsgBoxStyle.Critical, "Send Error")
                            Exit Sub
                        Else
                            vQuery = "exec dbo.USP_ICT_UpdateConfirmIncentiveDocNo '" & vDocNo & "' "
                            vCMD = New SqlCommand(vQuery, vConnection)
                            vCMD.ExecuteNonQuery()

                            MsgBox("อนุมัติเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information")
                            Call ClearData()
                        End If
                    End If
                End If
            Else
                MsgBox("เอกสารเลขที่ " & vDocNo & " ถูกยกเลิกไปแล้วไม่สามารถอนุมัติได้  กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
            End If
        End If
ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
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

            vQuery = "exec dbo.USP_ICT_RequestSelectList '" & vSaleType & "'," & vFiscalYear & "," & vPeriodOf4Week & " "
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "RequestSelectList")
            dt = ds.Tables("RequestSelectList")
            ListView102.Items.Clear()
            If dt.Rows.Count > 0 Then

                For i = 0 To dt.Rows.Count - 1
                    vListBudget = ListView102.Items.Add("")
                    vListBudget.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("departmentcode"))
                    vListBudget.SubItems.Add(2).Text = dt.Rows(i).Item("departmentname")
                    vListBudget.SubItems.Add(2).Text = dt.Rows(i).Item("team")
                    vListBudget.SubItems.Add(2).Text = Format(dt.Rows(i).Item("dept_incentive"), "##,##0.00")
                    vListBudget.SubItems.Add(3).Text = Format(dt.Rows(i).Item("team_incentive"), "##,##0.00")
                    vListBudget.SubItems.Add(3).Text = Format(dt.Rows(i).Item("requestremain"), "##,##0.00")
                    vListBudget.SubItems.Add(3).Text = dt.Rows(i).Item("team_isprocess")
                    vListBudget.SubItems.Add(3).Text = dt.Rows(i).Item("team_processdate")
                Next
                Me.CH101.Checked = True
            Else
                Me.CH101.Checked = False
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If

    End Sub

    Private Sub BTNPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPrint.Click

        If Me.TextDocNo.Text <> "" And vIsOpen = 1 And vIsCancel = 0 Then
            'Master Report
            If Me.frmIncentiveRequest Is Nothing Then
                frmIncentiveRequest = New FormIncentiveRequest
            Else
                If frmIncentiveRequest.IsDisposed Then
                    frmIncentiveRequest = New FormIncentiveRequest
                End If
            End If

            vIncentiveDocNo = Trim(Me.TextDocNo.Text)
            frmIncentiveRequest.Show()
            frmIncentiveRequest.BringToFront()

            'Details Report
            If Me.frmIncentiveRequestDetails Is Nothing Then
                frmIncentiveRequestDetails = New FormIncentiveRequestDetails
            Else
                If frmIncentiveRequestDetails.IsDisposed Then
                    frmIncentiveRequestDetails = New FormIncentiveRequestDetails
                End If
            End If

            vIncentiveDocNo = Trim(Me.TextDocNo.Text)
            frmIncentiveRequestDetails.Show()
            frmIncentiveRequestDetails.BringToFront()

        Else
            MsgBox("ไม่สามารถพิมพ์เอกสารได้", MsgBoxStyle.Critical, "Send Error")
        End If
    End Sub

    Private Sub BTNCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCancel.Click
        Dim vDocNo As String

        On Error GoTo ErrDescription

        If vIsOpen = 1 Then
            vDocNo = Trim(Me.TextDocNo.Text)
            If vIsConfirm = 0 Then
                If vUserID = vCreatorCode Then
                    If MessageBox.Show("คุณต้องการยกเลิกเอกสาร ใช่หรือไม่", "ข้อความสอบถาม", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                        If vIsCancel = 1 Then
                            MsgBox("ไม่สามารถยกเลิกเลขที่เอกสาร " & vDocNo & " ได้เนื่องจากได้ถูกยกเลิกเรียบร้อยแล้ว กรุณาตรวจสอบเลขที่เอกสารด้วย ", MsgBoxStyle.Critical, "Send Error")
                            Exit Sub
                        Else
                            vQuery = "exec dbo.USP_ICT_CancelIncentiveDocNo '" & vDocNo & "' "
                            vCMD = New SqlCommand(vQuery, vConnection)
                            vCMD.ExecuteNonQuery()

                            MsgBox("ยกเลิกเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information")
                            Call ClearData()
                        End If
                    End If
                Else
                    MsgBox("ไม่สามารถยกเลิกเอกสารเลขที่ " & vDocNo & " ได้เนื่องจาก ชื่อผู้สร้างเอกสารกับชื่อผู้ที่จะยกเลิกคนละชื่อกัน กรุณาแจ้งผู้สร้างเอกสารดำเนินการต่อไป", MsgBoxStyle.Critical, "Send Information")
                End If
            Else
                MsgBox("ไม่สามารถยกเลิกเอกสารเลขที่ " & vDocNo & " ได้เนื่องจาก ถูกอนุมัติเรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub CH101_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CH101.CheckedChanged
        Dim i As Integer
        If ListView102.Items.Count > 0 Then
            If CH101.Checked = True Then
                For i = 0 To ListView102.Items.Count - 1
                    ListView102.Items(i).Checked = True
                Next
            Else
                For i = 0 To ListView102.Items.Count - 1
                    ListView102.Items(i).Checked = False
                Next
            End If
        End If
    End Sub
End Class