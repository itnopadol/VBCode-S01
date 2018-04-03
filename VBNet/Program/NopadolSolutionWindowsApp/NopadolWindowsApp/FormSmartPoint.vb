Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization

Public Class FormSmartPoint
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim dt1 As DataTable
    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vMemIsMember As Integer
    Dim vMemSpecialCancel As Integer
    Dim vMemSpecialConfirm As Integer
    Dim vMemSpecialOpen As Integer

    Dim vMemIssueCancel As Integer
    Dim vMemIssueConfirm As Integer
    Dim vMemIssueOpen As Integer
    Dim vMemIssuePrintCount As Integer

    Dim vMemCampaignOpen As Integer

    Dim vMemCheckCouponAmount As Double

    Private Sub BTNAccumulateScore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAccumulateScore.Click
        On Error Resume Next

        Me.PNAccumulateScore.Visible = True
        Me.PNIssueScoreHistory.Visible = False
        Me.PNAddSpecialScore.Visible = False
        Me.PNIssueScore.Visible = False
        '=====================================================
        vMemIssueCancel = 0
        vMemIssueConfirm = 0
        vMemIssueOpen = 0
        vMemIssuePrintCount = 0
        Me.TBIssueNo.Text = ""
        Me.DTPIssueDate.Value = Now
        Me.TBIssueScore.Text = ""
        Me.TBMemOldScore.Text = ""
        Me.TBCouponAmount.Text = ""
        Me.ListViewCoupon.Items.Clear()
        Me.ND1000.Value = 0
        Me.ND500.Value = 0
        Me.ND200.Value = 0
        Me.ND100.Value = 0
        '=====================================================
        vMemSpecialCancel = 0
        vMemSpecialConfirm = 0
        vMemSpecialOpen = 0

        Me.TBSpecialNo.Text = ""
        Me.DTPSpecialDate.Value = Now
        If Me.CMBSpecialCampaign.Items.Count > 0 Then
            Me.CMBSpecialCampaign.SelectedIndex = 0
        Else
            Me.CMBSpecialCampaign.Text = ""
        End If
        Me.TBSpecialReason.Text = ""
        Me.TBSpecialScore.Text = ""

        If Me.ListViewAccumulateScore.Items.Count > 0 Then
            Me.ListViewAccumulateScore.Focus()
            Me.ListViewAccumulateScore.Items(0).Selected = True
            Me.ListViewAccumulateScore.Items(0).Focused = True
        Else
            Me.TBMemberID.Focus()
            Me.TBMemberID.SelectAll()
        End If
        '=====================================================
    End Sub

    Private Sub BTNIssueScoreHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNIssueScoreHistory.Click
        On Error Resume Next

        If vMemIsMember <> 0 Then
            Me.PNAccumulateScore.Visible = False
            Me.PNIssueScoreHistory.Visible = True

            Me.PNAddSpecialScore.Visible = False
            Me.PNIssueScore.Visible = False

            If Me.ListViewIssueScore.Items.Count > 0 Then
                Me.ListViewIssueScore.Focus()
                Me.ListViewIssueScore.Items(0).Selected = True
                Me.ListViewIssueScore.Items(0).Focused = True
            Else
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If
            '=====================================================
            vMemIssueCancel = 0
            vMemIssueConfirm = 0
            vMemIssueOpen = 0
            vMemIssuePrintCount = 0
            Me.TBIssueNo.Text = ""
            Me.DTPIssueDate.Value = Now
            Me.TBIssueScore.Text = ""
            Me.TBMemOldScore.Text = ""
            Me.TBCouponAmount.Text = ""
            Me.ListViewCoupon.Items.Clear()
            Me.ND1000.Value = 0
            Me.ND500.Value = 0
            Me.ND200.Value = 0
            Me.ND100.Value = 0
            '=====================================================
            vMemSpecialCancel = 0
            vMemSpecialConfirm = 0
            vMemSpecialOpen = 0

            Me.TBSpecialNo.Text = ""
            Me.DTPSpecialDate.Value = Now
            If Me.CMBSpecialCampaign.Items.Count > 0 Then
                Me.CMBSpecialCampaign.SelectedIndex = 0
            Else
                Me.CMBSpecialCampaign.Text = ""
            End If
            Me.TBSpecialReason.Text = ""
            Me.TBSpecialScore.Text = ""
            '=====================================================
        End If
    End Sub

    Private Sub ListViewAccumulateScore_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewAccumulateScore.Click

    End Sub

    Private Sub ListViewAccumulateScore_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewAccumulateScore.DoubleClick
        Dim vIndex As Integer
        Dim vInvoiceNo As String
        Dim vCampaignCode As String
        Dim vARCode As String
        Dim vStartDate As String
        Dim vStopDate As String

        On Error Resume Next

        If Me.ListViewAccumulateScore.Items.Count > 0 Then
            vIndex = Me.ListViewAccumulateScore.SelectedItems(0).Index
            Me.TBInvoiceNo.Text = Me.ListViewAccumulateScore.Items(vIndex).SubItems(2).Text

            If Me.TBInvoiceNo.Text = "" Then
                MsgBox("กรุณา กรอกเลขที่เอกสารขาย", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBInvoiceNo.Focus()
                Me.TBInvoiceNo.SelectAll()
                Exit Sub
            End If

            If Me.CMBCampaign.Text = "" Then
                MsgBox("กรุณา กรอกรหัสแคมเปญ", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBCampaign.Focus()
                Me.CMBCampaign.SelectAll()
                Exit Sub
            End If

            If Me.TBArCode.Text = "" Then
                MsgBox("กรุณา กรอกรหัสลูกค้า", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBArCode.Focus()
                Me.TBArCode.SelectAll()
                Exit Sub
            End If

            vInvoiceNo = Me.TBInvoiceNo.Text
            vARCode = Me.TBArCode.Text
            vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
            vStartDate = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
            vStopDate = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)
            Call CheckScoreHistFilter(vARCode, vCampaignCode, vMemIsMember, vStartDate, vStopDate, vInvoiceNo)
            Call vCalcArScore()
            Call vCalcArIssueScoreHist()

            Me.TBInvoiceNo.Focus()
            Me.TBInvoiceNo.SelectAll()
        End If
    End Sub

    Private Sub ListViewAccumulateScore_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListViewAccumulateScore.MouseMove

    End Sub

    Private Sub ListViewAccumulateScore_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewAccumulateScore.SelectedIndexChanged

    End Sub

    Private Sub TBInvoiceNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBInvoiceNo.KeyDown
        Dim vInvoiceNo As String
        Dim vCampaignCode As String
        Dim vARCode As String
        Dim vStartDate As String
        Dim vStopDate As String

        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.PNAccumulateScore.Visible = True
            Me.PNIssueScoreHistory.Visible = False
            Me.PNAddSpecialScore.Visible = False
            Me.PNIssueScore.Visible = False
        End If

        If e.KeyCode = Keys.Enter Then
            If Me.TBInvoiceNo.Text = "" Then
                MsgBox("กรุณา กรอกเลขที่เอกสารขาย", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBInvoiceNo.Focus()
                Me.TBInvoiceNo.SelectAll()
                Exit Sub
            End If

            If Me.CMBCampaign.Text = "" Then
                MsgBox("กรุณา กรอกรหัสแคมเปญ", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBCampaign.Focus()
                Me.CMBCampaign.SelectAll()
                Exit Sub
            End If

            If Me.TBArCode.Text = "" Then
                MsgBox("กรุณา กรอกรหัสลูกค้า", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBArCode.Focus()
                Me.TBArCode.SelectAll()
                Exit Sub
            End If

            vInvoiceNo = Me.TBInvoiceNo.Text
            vARCode = Me.TBArCode.Text
            vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
            vStartDate = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
            vStopDate = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)
            Call CheckScoreHistFilter(vARCode, vCampaignCode, vMemIsMember, vStartDate, vStopDate, vInvoiceNo)
            Call vCalcArScore()
            Call vCalcArIssueScoreHist()
        End If
    End Sub

    Private Sub TBInvoiceNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBInvoiceNo.TextChanged
        Dim vARCode As String
        Dim vCampaignCode As String
        Dim vCampaignStart As Date
        Dim vCampaignStop As Date
        Dim vSelectStart As Date
        Dim vSelectStop As Date
        Dim vInvoiceNo As String
        Dim vDate1 As String
        Dim vDate2 As String

        On Error Resume Next

        If Me.TBInvoiceNo.Text = "" Then
            vCampaignStart = Me.TBStartDate.Text
            vCampaignStop = Me.TBExpireDate.Text
            vSelectStart = Me.DTPDate1.Text
            vSelectStop = Me.DTPDate2.Text

            If Me.CMBCampaign.Text = "" Then
                Exit Sub
            End If

            If Me.TBArCode.Text = "" Then
                Exit Sub
            End If

            If vCampaignStart = vSelectStart And vCampaignStart = vSelectStart Then
                vARCode = Me.TBArCode.Text
                vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)

                Call CheckScoreHist(vARCode, vCampaignCode, vMemIsMember)
            Else
                If Me.CMBCampaign.Text = "" Then
                    Exit Sub
                End If

                If Me.TBArCode.Text = "" Then
                    Exit Sub
                End If
                vARCode = Me.TBArCode.Text
                vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
                vInvoiceNo = ""
                vDate1 = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
                vDate2 = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)
                Call CheckScoreHistFilter1(vARCode, vCampaignCode, vMemIsMember, vDate1, vDate2)
                Call vCalcArScore()
                Call vCalcArIssueScoreHist()
            End If
            Call vCalcArScore()
            Call vCalcArIssueScoreHist()
        End If
    End Sub

    Private Sub ListViewItemDetails_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.PNAccumulateScore.Visible = True
            Me.PNIssueScoreHistory.Visible = False
            Me.PNAddSpecialScore.Visible = False
            Me.PNIssueScore.Visible = False
        End If
    End Sub

    Private Sub ListViewItemDetails_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BTNAddSpecialScore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddSpecialScore.Click
        On Error Resume Next

        Call ChekAuthorityAccess()
        If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
            If vMemIsMember <> 0 Then
                Me.DTPSpecialDate.Value = Now
                Call vGetCampaignSpecial("")
                Call vGetSpecialDocNo()
                Call vGetNewSpecial()

                '=====================================================
                vMemIssueCancel = 0
                vMemIssueConfirm = 0
                vMemIssueOpen = 0
                vMemIssuePrintCount = 0
                Me.TBIssueNo.Text = ""
                Me.DTPIssueDate.Value = Now
                Me.TBIssueScore.Text = ""
                Me.TBMemOldScore.Text = ""
                Me.TBCouponAmount.Text = ""
                Me.ListViewCoupon.Items.Clear()
                Me.ND1000.Value = 0
                Me.ND500.Value = 0
                Me.ND200.Value = 0
                Me.ND100.Value = 0
                '=====================================================

                Me.TBSpecialScore.Focus()
                Me.TBSpecialScore.SelectAll()
            End If
        Else
            MsgBox("คุณไม่มีสิทธิ์ในการใช้งานในส่วนนี้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Public Sub vGetSpecialDocNo()
        Dim vDocNo As String
        Dim vNow As String

        On Error Resume Next

        vNow = vb6.Day(Now) & "/" & vb6.Month(Now) & "/" & vb6.Year(Now)
        vQuery = "set dateformat dmy"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vQuery = "select dbo.FT_VP_NewPointSpecialSet ('" & vNow & "') as docno"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then
            vDocNo = dt.Rows(0).Item("Docno")
            Me.TBSpecialNo.Text = vDocNo
            Me.PNAccumulateScore.Visible = False
            Me.PNIssueScoreHistory.Visible = False
            Me.PNAddSpecialScore.Visible = True
            Me.PNIssueScore.Visible = False
            Call AddScoreNew()
        Else
            Me.TBSpecialNo.Text = ""
        End If

        vMemSpecialCancel = 0
        vMemSpecialConfirm = 0
        vMemSpecialOpen = 0
    End Sub

    Public Sub vGetIssueDocNo()
        Dim vDocNo As String
        Dim vNow As String

        On Error Resume Next

        If vb6.Year(Now) >= 2500 Then
            vNow = vb6.Day(Now) & "/" & vb6.Month(Now) & "/" & vb6.Year(Now) - 543
        Else
            vNow = vb6.Day(Now) & "/" & vb6.Month(Now) & "/" & vb6.Year(Now)
        End If
        vQuery = "set dateformat dmy"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vQuery = "select dbo.FT_VP_NewWithdrawSet ('" & vNow & "') as docno"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then
            vDocNo = dt.Rows(0).Item("Docno")
            Me.TBIssueNo.Text = vDocNo
            Me.PNAccumulateScore.Visible = False
            Me.PNIssueScoreHistory.Visible = False
            Me.PNAddSpecialScore.Visible = False
            Me.PNIssueScore.Visible = True
        Else
            Me.TBIssueNo.Text = ""
        End If

        vMemIssueCancel = 0
        vMemIssueConfirm = 0
        vMemIssueOpen = 0
        vMemIssuePrintCount = 0
    End Sub

    Public Sub AddScoreNew()
        On Error Resume Next

        Me.PBAddNew.Visible = True
        Me.PBAddCancel.Visible = False
        Me.PBAddConfirm.Visible = False
    End Sub

    Private Sub BTNIssueScore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNIssueScore.Click
        On Error Resume Next

        Call ChekAuthorityAccess()
        If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
            If vMemIsMember <> 0 Then
                Me.PNAccumulateScore.Visible = False
                Me.PNIssueScoreHistory.Visible = False
                Me.PNAddSpecialScore.Visible = False
                Me.PNIssueScore.Visible = True
                Call IssueNew()
                Me.DTPIssueDate.Value = Now
                Call vGetIssueDocNo()

                '=====================================================
                vMemSpecialCancel = 0
                vMemSpecialConfirm = 0
                vMemSpecialOpen = 0

                Me.TBSpecialNo.Text = ""
                Me.DTPSpecialDate.Value = Now
                If Me.CMBSpecialCampaign.Items.Count > 0 Then
                    Me.CMBSpecialCampaign.SelectedIndex = 0
                Else
                    Me.CMBSpecialCampaign.Text = ""
                End If
                Me.TBSpecialReason.Text = ""
                Me.TBSpecialScore.Text = ""
                '=====================================================

                Me.TBIssueScore.Focus()
                Me.TBIssueScore.SelectAll()
            End If

        Else
            MsgBox("คุณไม่มีสิทธิ์ในการใช้งานในส่วนนี้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Public Sub IssueNew()
        On Error Resume Next

        Me.PBIssueNew.Visible = True
        Me.PBIssueCancel.Visible = False
        Me.PBIssueConfirm.Visible = False
    End Sub

    Public Sub IssueCancel()
        On Error Resume Next

        Me.PBIssueNew.Visible = False
        Me.PBIssueCancel.Visible = True
        Me.PBIssueConfirm.Visible = False
    End Sub

    Public Sub IssueConfirm()
        On Error Resume Next

        Me.PBIssueNew.Visible = False
        Me.PBIssueCancel.Visible = False
        Me.PBIssueConfirm.Visible = True
    End Sub

    Private Sub BTNSearchSpecialScore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSpecialSearch.Click
        Dim vARCode As String

        On Error Resume Next

        If Me.TBArName.Text <> "" Then
            vARCode = Me.TBArCode.Text
            Call SearchSpecialScore(vARCode, "")
            If Me.ListViewSearchSpecialScore.Items.Count > 0 Then
                Me.PNSearchSpecialScore.Visible = True
                Me.ListViewSearchSpecialScore.Focus()
                Me.ListViewSearchSpecialScore.Items(0).Focused = True
                Me.ListViewSearchSpecialScore.Items(0).Selected = True
            Else
                MsgBox("ลูกค้ารหัส " & vARCode & " ไม่มีข้อมูลการเพิ่มแต้มพิเศษ ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBSpecialNo.Focus()
                Me.TBSpecialNo.SelectAll()
            End If
        End If
    End Sub

    Public Sub SearchSpecialScore(ByVal vArCode As String, ByVal vDocno As String)
        Dim i As Integer
        Dim n As Integer
        Dim vScore As Double
        Dim vListDocNo As ListViewItem
        Dim vIsCancel As Integer
        Dim vIsconfirm As Integer

        On Error Resume Next

        Me.ListViewSearchSpecialScore.Items.Clear()
        vQuery = "exec dbo.USP_VP_PointSpecialSearchAR '" & vDocno & "','" & vArCode & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Member")
        dt = ds.Tables("member")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vIsCancel = dt.Rows(i).Item("iscancel")
                vIsconfirm = dt.Rows(i).Item("isconfirm")
                vListDocNo = Me.ListViewSearchSpecialScore.Items.Add(n)
                vListDocNo.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                vListDocNo.SubItems.Add(1).Text = dt.Rows(i).Item("docdate")
                vScore = dt.Rows(i).Item("point")
                vListDocNo.SubItems.Add(2).Text = Format(vScore, "##,##0.00")
                vListDocNo.SubItems.Add(3).Text = dt.Rows(i).Item("reason")
                If vIsCancel = 1 Then
                    vListDocNo.SubItems.Add(4).Text = "ยกเลิก"
                ElseIf vIsconfirm = 1 Then
                    vListDocNo.SubItems.Add(4).Text = "อนุมติแล้ว"
                Else
                    vListDocNo.SubItems.Add(4).Text = "ยังไม่ได้อนุมติ"
                End If
            Next

            Dim a As Integer
            If Me.ListViewSearchSpecialScore.Items.Count > 0 Then
                For a = 0 To Me.ListViewSearchSpecialScore.Items.Count - 1
                    If a Mod 2 = 0 Then
                        Me.ListViewSearchSpecialScore.Items(a).BackColor = Color.AliceBlue
                    Else
                        Me.ListViewSearchSpecialScore.Items(a).BackColor = Color.LavenderBlush
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub BTNCloseSearchSpecialScore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSearchSpecialScore.Click
        Me.PNSearchSpecialScore.Visible = False
    End Sub

    Private Sub BTNCloseSearchIssueScore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSearchIssueScore.Click
        Me.PNSearchIssueScore.Visible = False
    End Sub

    Private Sub BTNSearchIssueScore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNIssueSearch.Click
        Dim vARCode As String

        On Error Resume Next

        If Me.TBArName.Text <> "" Then
            vARCode = Me.TBArCode.Text
            Call SearchIssueScore(vARCode, "")
            If Me.ListViewSearchIssueScore.Items.Count > 0 Then
                Me.PNSearchIssueScore.Visible = True
                Me.ListViewSearchIssueScore.Focus()
                Me.ListViewSearchIssueScore.Items(0).Focused = True
                Me.ListViewSearchIssueScore.Items(0).Selected = True
            Else
                MsgBox("ลูกค้ารหัส " & vARCode & " ไม่มีข้อมูลการเบิกแต้ม ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBIssueNo.Focus()
                Me.TBIssueNo.SelectAll()
            End If
        End If
    End Sub

    Public Sub SearchIssueScore(ByVal vARCode As String, ByVal vDocNo As String)
        Dim i As Integer
        Dim n As Integer
        Dim vScore As Double
        Dim vListDocNo As ListViewItem
        Dim vIsCancel As Integer
        Dim vIsconfirm As Integer
        Dim vPrintCount As Integer

        On Error Resume Next

        Me.ListViewSearchIssueScore.Items.Clear()
        vQuery = "exec dbo.USP_VP_withdrawSearchNew '" & vARCode & "','" & vDocNo & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "IssueNo")
        dt = ds.Tables("IssueNo")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vIsCancel = dt.Rows(i).Item("iscancel")
                vIsconfirm = dt.Rows(i).Item("isconfirm")
                vPrintCount = dt.Rows(i).Item("printcount")

                vListDocNo = Me.ListViewSearchIssueScore.Items.Add(n)
                vListDocNo.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                vListDocNo.SubItems.Add(1).Text = dt.Rows(i).Item("docdate")
                vScore = dt.Rows(i).Item("point")
                vListDocNo.SubItems.Add(2).Text = Format(vScore, "##,##0.00")
                vListDocNo.SubItems.Add(3).Text = ""
                If vIsCancel = 1 Then
                    vListDocNo.SubItems.Add(4).Text = "ยกเลิก"
                ElseIf vPrintCount > 0 Then
                    vListDocNo.SubItems.Add(4).Text = "พิมพ์คูปองแล้ว"
                ElseIf vPrintCount = 0 Then
                    vListDocNo.SubItems.Add(4).Text = "ยังไม่ได้พิมพ์"
                End If
            Next

            Dim a As Integer
            If Me.ListViewSearchIssueScore.Items.Count > 0 Then
                For a = 0 To Me.ListViewSearchIssueScore.Items.Count - 1
                    If a Mod 2 = 0 Then
                        Me.ListViewSearchIssueScore.Items(a).BackColor = Color.AliceBlue
                    Else
                        Me.ListViewSearchIssueScore.Items(a).BackColor = Color.LavenderBlush
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub FormSmartPoint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        On Error Resume Next

        Call InitializeDataBase()
        Call vGetNewSpecial()
        Call vGetCampaign("")
        Call vGetAcculateLine()
        Call vGetIssueLine()

        Me.PNCheckScore.Visible = True
        Me.PNCampaign.Visible = False
        Me.PNSpecialApprove.Visible = False
        Me.PNSpecialCancel.Visible = False

        Me.TBMemberID.Focus()
        Me.TBMemberID.SelectAll()
    End Sub

    Public Sub vGetCampaign(ByVal vSearch As String)
        Dim i As Integer

        On Error Resume Next

        Me.CMBCampaign.Items.Clear()
        vQuery = "exec dbo.USP_VP_CampaignSearch '" & vSearch & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Campaign")
        dt = ds.Tables("Campaign")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                Me.CMBCampaign.Items.Add(dt.Rows(i).Item("code") & "/" & dt.Rows(i).Item("nameth"))
            Next
        End If

        If Me.CMBCampaign.Items.Count > 0 Then
            Me.CMBCampaign.SelectedIndex = 0
        End If
    End Sub

    Public Sub vGetCampaignSpecial(ByVal vSearch As String)
        Dim i As Integer

        On Error Resume Next

        Me.CMBSpecialCampaign.Items.Clear()
        vQuery = "exec dbo.USP_VP_CampaignSearch '" & vSearch & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Campaign")
        dt = ds.Tables("Campaign")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                Me.CMBSpecialCampaign.Items.Add(dt.Rows(i).Item("code") & "/" & dt.Rows(i).Item("nameth"))
            Next
        End If

        If Me.CMBSpecialCampaign.Items.Count > 0 Then
            Me.CMBSpecialCampaign.SelectedIndex = 0
        End If
    End Sub

    Private Sub BTNCloseSearchMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSearchMember.Click
        On Error Resume Next

        Me.PNSearchMember.Visible = False

        If Me.TBArName.Text = "" Then
            Me.TBMemberID.Focus()
            Me.TBMemberID.SelectAll()
        End If

        If Me.TBArName.Text <> "" And Me.TBArCode.Text <> "" Then
            Me.CMBCampaign.Focus()
        End If
    End Sub

    Private Sub TBMemberID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBMemberID.KeyDown ', BTNSearchMember.KeyDown, TBArName.KeyDown, TBArCode.KeyDown, TBArAddress.KeyDown, TBApplyDate.KeyDown, TBExpireDate.KeyDown, CMBCampaign.KeyDown, TBStartDate.KeyDown, TBEndDate.KeyDown, TBScoreRemain.KeyDown, BTNAccumulateScore.KeyDown, BTNIssueScoreHistory.KeyDown, BTNAddSpecialScore.KeyDown, BTNCalcScore.KeyDown, BTNIssueScore.KeyDown
        If e.KeyCode = Keys.Enter Then

        End If
    End Sub

    Public Sub ClearMember()
        Me.TBMemberID.Text = ""
    End Sub


    Public Sub vSearchMember(ByVal vSearch As String)
        Dim i As Integer
        Dim n As Integer
        Dim vListAr As ListViewItem
        Dim vMemberStatus As Integer

        On Error Resume Next

        If Me.ListViewSearchMember.Items.Count = 0 Then
            vQuery = "exec dbo.USP_AR_SearchMember '" & vSearch & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Member")
            dt = ds.Tables("member")
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    n = n + 1
                    vMemberStatus = dt.Rows(i).Item("memberstatus")
                    vListAr = Me.ListViewSearchMember.Items.Add(n)
                    vListAr.SubItems.Add(0).Text = dt.Rows(i).Item("memberid")
                    vListAr.SubItems.Add(1).Text = dt.Rows(i).Item("name1")
                    vListAr.SubItems.Add(2).Text = dt.Rows(i).Item("code")
                    If vMemberStatus = 1 Then
                        vListAr.SubItems.Add(3).Text = dt.Rows(i).Item("begindate")
                        vListAr.SubItems.Add(4).Text = dt.Rows(i).Item("expiredate")
                    Else
                        vListAr.SubItems.Add(3).Text = ""
                        vListAr.SubItems.Add(4).Text = ""
                    End If
                Next

                Dim a As Integer
                If Me.ListViewSearchMember.Items.Count > 0 Then
                    For a = 0 To Me.ListViewSearchMember.Items.Count - 1
                        If a Mod 2 = 0 Then
                            Me.ListViewSearchMember.Items(a).BackColor = Color.AliceBlue
                        Else
                            Me.ListViewSearchMember.Items(a).BackColor = Color.LavenderBlush
                        End If
                    Next
                End If

                If Me.ListViewSearchMember.Items.Count > 0 Then
                    Me.PNSearchMember.Visible = True
                    Me.ListViewSearchMember.Focus()
                    Me.ListViewSearchMember.Items(0).Selected = True
                    Me.ListViewSearchMember.Items(0).Focused = True
                End If

            End If
        ElseIf Me.ListViewSearchMember.Items.Count > 0 And Me.TBSearchMember.Text = "" Then
            Me.PNSearchMember.Visible = True
            Call vGetSearchMemberLine()
            Me.TBSearchMember.Focus()
            Me.TBSearchMember.SelectAll()

        ElseIf Me.ListViewSearchMember.Items.Count > 0 And Me.TBSearchMember.Text <> "" Then
            Me.ListViewSearchMember.Items.Clear()
            vQuery = "exec dbo.USP_AR_SeeMember '" & vSearch & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Member")
            dt = ds.Tables("member")
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    n = n + 1
                    vMemberStatus = dt.Rows(i).Item("memberstatus")
                    vListAr = Me.ListViewSearchMember.Items.Add(n)
                    vListAr.SubItems.Add(0).Text = dt.Rows(i).Item("memberid")
                    vListAr.SubItems.Add(1).Text = dt.Rows(i).Item("name1")
                    vListAr.SubItems.Add(2).Text = dt.Rows(i).Item("code")
                    If vMemberStatus = 1 Then
                        vListAr.SubItems.Add(3).Text = dt.Rows(i).Item("begindate")
                        vListAr.SubItems.Add(4).Text = dt.Rows(i).Item("expiredate")
                    Else
                        vListAr.SubItems.Add(3).Text = ""
                        vListAr.SubItems.Add(4).Text = ""
                    End If
                Next

                Dim a As Integer
                If Me.ListViewSearchMember.Items.Count > 0 Then
                    For a = 0 To Me.ListViewSearchMember.Items.Count - 1
                        If a Mod 2 = 0 Then
                            Me.ListViewSearchMember.Items(a).BackColor = Color.AliceBlue
                        Else
                            Me.ListViewSearchMember.Items(a).BackColor = Color.LavenderBlush
                        End If
                    Next
                End If

                If Me.ListViewSearchMember.Items.Count > 0 Then
                    Me.PNSearchMember.Visible = True
                    Me.ListViewSearchMember.Focus()
                    Me.ListViewSearchMember.Items(0).Selected = True
                    Me.ListViewSearchMember.Items(0).Focused = True
                End If
            End If
        End If
    End Sub


    Private Sub TBMemberID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TBMemberID.LostFocus
        'Dim vMemberID As String

        'vMemberID = Me.TBMemberID.Text
        'Call SearchMemberDetails(vMemberID)

        'Dim vMemberID As String
        'Dim vArCode As String
        'Dim vMemberStatus As Integer
        'Dim vCampaignCode As String
        'Dim vScore As Double
        'Dim vMemBeginDate As Date
        'Dim vMemExpireDate As Date
        'Dim vCheckYearBegin As Integer
        'Dim vCheckYearExpire As Integer
        'Dim vNewBeginDate As Date
        'Dim vNewExpireDate As Date

        'If Me.TBMemberID.Text <> "" Then
        '    vMemberID = Me.TBMemberID.Text
        '    vQuery = "exec dbo.usp_ar_searchmemberdetails '" & vMemberID & "'"
        '    da = New SqlDataAdapter(vQuery, vConnection)
        '    ds = New DataSet
        '    da.Fill(ds, "Campaign")
        '    dt = ds.Tables("Campaign")
        '    If dt.Rows.Count > 0 Then
        '        Me.TBArName.Text = dt.Rows(0).Item("name1")
        '        Me.TBArCode.Text = dt.Rows(0).Item("code")
        '        Me.TBArAddress.Text = dt.Rows(0).Item("billaddress")
        '        vMemberStatus = dt.Rows(0).Item("memberstatus")
        '        vMemIsMember = dt.Rows(0).Item("memberstatus")
        '        If vMemberStatus = 0 Then
        '            Me.TBApplyDate.Text = "ยังไม่ได้สมัคร"
        '            Me.TBExpireDate.Text = "ยังไม่ได้สมัคร"
        '        ElseIf vMemberStatus = 1 Then
        '            vMemBeginDate = dt.Rows(0).Item("beginmember")
        '            vMemExpireDate = dt.Rows(0).Item("expiredate")
        '            vCheckYearBegin = vb6.Year(vMemBeginDate)
        '            vCheckYearExpire = vb6.Year(vMemExpireDate)

        '            If vCheckYearBegin > 2000 Then
        '                vNewBeginDate = vb6.DateAdd(DateInterval.Year, 543, vMemBeginDate)
        '                Me.TBApplyDate.Text = vNewBeginDate 'dt.Rows(0).Item("beginmember")
        '            End If

        '            If vCheckYearExpire > 2000 Then
        '                vNewExpireDate = vb6.DateAdd(DateInterval.Year, 543, vMemExpireDate)
        '                Me.TBExpireDate.Text = vNewExpireDate 'dt.Rows(0).Item("beginmember")
        '            End If

        '        ElseIf vMemberStatus = 2 Then
        '            vMemBeginDate = dt.Rows(0).Item("beginmember")
        '            vMemExpireDate = dt.Rows(0).Item("expiredate")
        '            vCheckYearBegin = vb6.Year(vMemBeginDate)
        '            vCheckYearExpire = vb6.Year(vMemExpireDate)

        '            If vCheckYearBegin > 2000 Then
        '                vNewBeginDate = vb6.DateAdd(DateInterval.Year, 543, vMemBeginDate)
        '                Me.TBApplyDate.Text = vNewBeginDate 'dt.Rows(0).Item("beginmember")
        '            End If

        '            If vCheckYearExpire > 2000 Then
        '                vNewExpireDate = vb6.DateAdd(DateInterval.Year, 543, vMemExpireDate)
        '                Me.TBExpireDate.Text = vNewExpireDate 'dt.Rows(0).Item("beginmember")
        '            Else
        '                Me.TBExpireDate.Text = "หมดอายุ"
        '            End If
        '        End If

        '        vArCode = Me.TBArCode.Text
        '        vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
        '        Call CheckScoreHist(vArCode, vCampaignCode, vMemberStatus)
        '        Call vCheckIssueSore(vArCode, vCampaignCode)
        '        Call vCalcArScore()
        '    Else
        '        Me.TBArName.Text = ""
        '        Me.TBArCode.Text = ""
        '        Me.TBArAddress.Text = ""
        '        Me.TBApplyDate.Text = ""
        '        Me.TBExpireDate.Text = ""
        '        Me.TBArScore.Text = ""
        '        vMemberStatus = 0
        '        vScore = 0
        '        Me.ListViewAccumulateScore.Items.Clear()
        '        Me.ListViewIssueScore.Items.Clear()
        '        Call vGetAcculateLine()
        '        Call vGetIssueLine()
        '        Call ClearSpecialScore()
        '        Me.TBScoreRemain.Text = Format(vScore, "##,##0.00")
        '    End If

        '    If vMemberStatus <> 0 Then
        '        If Me.CMBCampaign.Text <> "" And Me.TBArCode.Text <> "" Then
        '            vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
        '            vArCode = Me.TBArCode.Text
        '            vQuery = "exec dbo.usp_vp_checkpoint '" & vCampaignCode & "','" & vArCode & "'"
        '            da = New SqlDataAdapter(vQuery, vConnection)
        '            ds = New DataSet
        '            da.Fill(ds, "Score")
        '            dt = ds.Tables("Score")
        '            If dt.Rows.Count > 0 Then
        '                vScore = dt.Rows(0).Item("point")
        '            Else
        '                vScore = 0
        '            End If
        '            Me.TBScoreRemain.Text = Format(vScore, "##,##0.00")
        '            Me.CMBSpecialCampaign.Focus()
        '        End If
        '    Else
        '        Me.TBArCode.Focus()
        '        Me.TBArCode.SelectAll()
        '    End If
        'Else
        '    Me.TBArName.Text = ""
        '    Me.TBArCode.Text = ""
        '    Me.TBArAddress.Text = ""
        '    Me.TBApplyDate.Text = ""
        '    Me.TBExpireDate.Text = ""
        '    Me.TBArScore.Text = ""
        '    vMemberStatus = 0
        '    vScore = 0
        '    Me.ListViewAccumulateScore.Items.Clear()
        '    Me.ListViewIssueScore.Items.Clear()
        '    Call vGetAcculateLine()
        '    Call vGetIssueLine()
        '    Call ClearSpecialScore()
        '    Me.TBScoreRemain.Text = Format(vScore, "##,##0.00")
        'End If

        'Me.TBMemberID.Text = UCase(Me.TBMemberID.Text)


    End Sub
    Public Sub SearchMemberDetails(ByVal vMemberID As String)
        Dim vArCode As String
        Dim vMemberStatus As Integer
        Dim vCampaignCode As String
        Dim vScore As Double
        Dim vMemBeginDate As Date
        Dim vMemExpireDate As Date
        Dim vCheckYearBegin As Integer
        Dim vCheckYearExpire As Integer
        Dim vNewBeginDate As Date
        Dim vNewExpireDate As Date

        On Error Resume Next

        If Me.TBMemberID.Text <> "" Then
            vMemberID = Me.TBMemberID.Text
            vQuery = "exec dbo.usp_ar_searchmemberdetails '" & vMemberID & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "MemberDetails")
            dt = ds.Tables("MemberDetails")
            If dt.Rows.Count > 0 Then
                Me.TBArName.Text = dt.Rows(0).Item("name1")
                Me.TBArCode.Text = dt.Rows(0).Item("code")
                Me.TBArAddress.Text = dt.Rows(0).Item("billaddress")
                vMemberStatus = dt.Rows(0).Item("memberstatus")
                vMemIsMember = dt.Rows(0).Item("memberstatus")
                If vMemberStatus = 0 Then
                    Me.TBApplyDate.Text = "ยังไม่ได้สมัคร"
                    Me.TBExpireDate.Text = "ยังไม่ได้สมัคร"
                ElseIf vMemberStatus = 1 Then
                    vMemBeginDate = dt.Rows(0).Item("beginmember")
                    vMemExpireDate = dt.Rows(0).Item("expiredate")
                    vCheckYearBegin = vb6.Year(vMemBeginDate)
                    vCheckYearExpire = vb6.Year(vMemExpireDate)

                    If vCheckYearBegin > 2000 Then
                        vNewBeginDate = vb6.DateAdd(DateInterval.Year, 543, vMemBeginDate)
                        Me.TBApplyDate.Text = vNewBeginDate 'dt.Rows(0).Item("beginmember")
                    End If

                    If vCheckYearExpire > 2000 Then
                        vNewExpireDate = vb6.DateAdd(DateInterval.Year, 543, vMemExpireDate)
                        Me.TBExpireDate.Text = vNewExpireDate 'dt.Rows(0).Item("beginmember")
                    End If

                ElseIf vMemberStatus = 2 Then
                    vMemBeginDate = dt.Rows(0).Item("beginmember")
                    vMemExpireDate = dt.Rows(0).Item("expiredate")
                    vCheckYearBegin = vb6.Year(vMemBeginDate)
                    vCheckYearExpire = vb6.Year(vMemExpireDate)

                    If vCheckYearBegin > 2000 Then
                        vNewBeginDate = vb6.DateAdd(DateInterval.Year, 543, vMemBeginDate)
                        Me.TBApplyDate.Text = vNewBeginDate 'dt.Rows(0).Item("beginmember")
                    End If

                    If vCheckYearExpire > 2000 Then
                        vNewExpireDate = vb6.DateAdd(DateInterval.Year, 543, vMemExpireDate)
                        Me.TBExpireDate.Text = vNewExpireDate 'dt.Rows(0).Item("beginmember")
                    Else
                        Me.TBExpireDate.Text = "หมดอายุ"
                    End If
                End If

                vArCode = Me.TBArCode.Text
                vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
                Call CheckScoreHist(vArCode, vCampaignCode, vMemberStatus)
                Call vCheckIssueSore(vArCode, vCampaignCode)
                Call vCalcArScore()
                Call vCalcArIssueScoreHist()
                Call IssueClearScreen()
                Call ClearSpecialScore()
                Me.PNAccumulateScore.Visible = True
                Me.PNIssueScoreHistory.Visible = False
                Me.PNAddSpecialScore.Visible = False
                Me.PNIssueScore.Visible = False

            Else
                Me.TBArName.Text = ""
                Me.TBArCode.Text = ""
                Me.TBArAddress.Text = ""
                Me.TBApplyDate.Text = ""
                Me.TBExpireDate.Text = ""
                Me.TBArScore.Text = ""
                vMemberStatus = 0
                vScore = 0
                vMemIsMember = 0
                Me.ListViewAccumulateScore.Items.Clear()
                Me.ListViewIssueScore.Items.Clear()
                Call vGetAcculateLine()
                Call vGetIssueLine()
                Call IssueClearScreen()
                Call ClearSpecialScore()
                Me.PNAccumulateScore.Visible = True
                Me.PNIssueScoreHistory.Visible = False
                Me.PNAddSpecialScore.Visible = False
                Me.PNIssueScore.Visible = False
                Me.TBScoreRemain.Text = Format(vScore, "##,##0.00")
            End If

            If vMemberStatus <> 0 Then
                If Me.CMBCampaign.Text <> "" And Me.TBArCode.Text <> "" Then
                    vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
                    vArCode = Me.TBArCode.Text
                    vQuery = "exec dbo.usp_vp_checkpoint '" & vCampaignCode & "','" & vArCode & "'"
                    da = New SqlDataAdapter(vQuery, vConnection)
                    ds = New DataSet
                    da.Fill(ds, "Score")
                    dt = ds.Tables("Score")
                    If dt.Rows.Count > 0 Then
                        vScore = dt.Rows(0).Item("point")
                    Else
                        vScore = 0
                    End If
                    Me.TBScoreRemain.Text = Format(vScore, "##,##0.00")
                End If

                If Me.ListViewAccumulateScore.Items.Count > 0 And Me.ListViewAccumulateScore.Items(0).SubItems(1).Text <> "" Then
                    Me.ListViewAccumulateScore.Focus()
                    Me.ListViewAccumulateScore.Items(0).Focused = True
                    Me.ListViewAccumulateScore.Items(0).Selected = True
                Else
                    Me.TBMemberID.Focus()
                    Me.TBMemberID.SelectAll()
                End If
            End If
        Else
            Me.TBArName.Text = ""
            Me.TBArCode.Text = ""
            Me.TBArAddress.Text = ""
            Me.TBApplyDate.Text = ""
            Me.TBExpireDate.Text = ""
            Me.TBArScore.Text = ""
            vMemberStatus = 0
            vScore = 0
            vMemIsMember = 0
            Me.ListViewAccumulateScore.Items.Clear()
            Me.ListViewIssueScore.Items.Clear()
            Call vGetAcculateLine()
            Call vGetIssueLine()
            Call IssueClearScreen()
            Call ClearSpecialScore()
            Me.PNAccumulateScore.Visible = True
            Me.PNIssueScoreHistory.Visible = False
            Me.PNAddSpecialScore.Visible = False
            Me.PNIssueScore.Visible = False
            Me.TBScoreRemain.Text = Format(vScore, "##,##0.00")
        End If
    End Sub

    Private Sub TBMemberID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBMemberID.TextChanged
        Dim vMemberID As String
        Dim vMemberStatus As Integer
        Dim vScore As Integer

        On Error Resume Next

        If Me.TBMemberID.Text <> "" Then
            vMemberID = Me.TBMemberID.Text
            If vMemberID = "99999" Or vMemberID = "999999" Or vMemberID = "1" Then
                MsgBox("รหัสเงินสดไม่สามารถดูข้อมูลได้", MsgBoxStyle.Critical, "Send Error Message")
            Else
                Call SearchMemberDetails(vMemberID)
            End If
        Else
            Me.TBArName.Text = ""
            Me.TBArCode.Text = ""
            Me.TBArAddress.Text = ""
            Me.TBApplyDate.Text = ""
            Me.TBExpireDate.Text = ""
            Me.TBArScore.Text = ""
            vMemberStatus = 0
            vScore = 0
            vMemIsMember = 0
            Me.ListViewAccumulateScore.Items.Clear()
            Me.ListViewIssueScore.Items.Clear()
            Call vGetAcculateLine()
            Call vGetIssueLine()
            Call IssueClearScreen()
            Call ClearSpecialScore()
            Me.PNAccumulateScore.Visible = True
            Me.PNIssueScoreHistory.Visible = False
            Me.PNAddSpecialScore.Visible = False
            Me.PNIssueScore.Visible = False
            Me.TBScoreRemain.Text = Format(vScore, "##,##0.00")
        End If

        'Dim vMemberID As String
        'Dim vArCode As String
        'Dim vMemberStatus As Integer
        'Dim vCampaignCode As String
        'Dim vScore As Double
        'Dim vMemBeginDate As Date
        'Dim vMemExpireDate As Date
        'Dim vCheckYearBegin As Integer
        'Dim vCheckYearExpire As Integer
        'Dim vNewBeginDate As Date
        'Dim vNewExpireDate As Date

        'If Me.TBMemberID.Text <> "" Then
        '    vMemberID = Me.TBMemberID.Text
        '    vQuery = "exec dbo.usp_ar_searchmemberdetails '" & vMemberID & "'"
        '    da = New SqlDataAdapter(vQuery, vConnection)
        '    ds = New DataSet
        '    da.Fill(ds, "Campaign")
        '    dt = ds.Tables("Campaign")
        '    If dt.Rows.Count > 0 Then
        '        Me.TBArName.Text = dt.Rows(0).Item("name1")
        '        Me.TBArCode.Text = dt.Rows(0).Item("code")
        '        Me.TBArAddress.Text = dt.Rows(0).Item("billaddress")
        '        vMemberStatus = dt.Rows(0).Item("memberstatus")
        '        vMemIsMember = dt.Rows(0).Item("memberstatus")
        '        If vMemberStatus = 0 Then
        '            Me.TBApplyDate.Text = "ยังไม่ได้สมัคร"
        '            Me.TBExpireDate.Text = "ยังไม่ได้สมัคร"
        '        ElseIf vMemberStatus = 1 Then
        '            vMemBeginDate = dt.Rows(0).Item("beginmember")
        '            vMemExpireDate = dt.Rows(0).Item("expiredate")
        '            vCheckYearBegin = vb6.Year(vMemBeginDate)
        '            vCheckYearExpire = vb6.Year(vMemExpireDate)

        '            If vCheckYearBegin > 2000 Then
        '                vNewBeginDate = vb6.DateAdd(DateInterval.Year, 543, vMemBeginDate)
        '                Me.TBApplyDate.Text = vNewBeginDate 'dt.Rows(0).Item("beginmember")
        '            End If

        '            If vCheckYearExpire > 2000 Then
        '                vNewExpireDate = vb6.DateAdd(DateInterval.Year, 543, vMemExpireDate)
        '                Me.TBExpireDate.Text = vNewExpireDate 'dt.Rows(0).Item("beginmember")
        '            End If

        '        ElseIf vMemberStatus = 2 Then
        '            vMemBeginDate = dt.Rows(0).Item("beginmember")
        '            vMemExpireDate = dt.Rows(0).Item("expiredate")
        '            vCheckYearBegin = vb6.Year(vMemBeginDate)
        '            vCheckYearExpire = vb6.Year(vMemExpireDate)

        '            If vCheckYearBegin > 2000 Then
        '                vNewBeginDate = vb6.DateAdd(DateInterval.Year, 543, vMemBeginDate)
        '                Me.TBApplyDate.Text = vNewBeginDate 'dt.Rows(0).Item("beginmember")
        '            End If

        '            If vCheckYearExpire > 2000 Then
        '                vNewExpireDate = vb6.DateAdd(DateInterval.Year, 543, vMemExpireDate)
        '                Me.TBExpireDate.Text = vNewExpireDate 'dt.Rows(0).Item("beginmember")
        '            Else
        '                Me.TBExpireDate.Text = "หมดอายุ"
        '            End If
        '        End If

        '        vArCode = Me.TBArCode.Text
        '        vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
        '        Call CheckScoreHist(vArCode, vCampaignCode, vMemberStatus)
        '        Call vCheckIssueSore(vArCode, vCampaignCode)
        '        Call vCalcArScore()
        '    Else
        '        Me.TBArName.Text = ""
        '        Me.TBArCode.Text = ""
        '        Me.TBArAddress.Text = ""
        '        Me.TBApplyDate.Text = ""
        '        Me.TBExpireDate.Text = ""
        '        Me.TBArScore.Text = ""
        '        vMemberStatus = 0
        '        vMemIsMember = 0
        '        vScore = 0
        '        Me.ListViewAccumulateScore.Items.Clear()
        '        Me.ListViewIssueScore.Items.Clear()
        '        Call vGetAcculateLine()
        '        Call vGetIssueLine()
        '        Call ClearSpecialScore()
        '        Me.PNAddSpecialScore.Visible = False
        '        Me.PNAccumulateScore.Visible = True
        '        Me.TBScoreRemain.Text = Format(vScore, "##,##0.00")
        '    End If

        '    If vMemberStatus <> 0 Then
        '        If Me.CMBCampaign.Text <> "" And Me.TBArCode.Text <> "" Then
        '            vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
        '            vArCode = Me.TBArCode.Text
        '            vQuery = "exec dbo.usp_vp_checkpoint '" & vCampaignCode & "','" & vArCode & "'"
        '            da = New SqlDataAdapter(vQuery, vConnection)
        '            ds = New DataSet
        '            da.Fill(ds, "Score")
        '            dt = ds.Tables("Score")
        '            If dt.Rows.Count > 0 Then
        '                vScore = dt.Rows(0).Item("point")
        '            Else
        '                vScore = 0
        '            End If
        '            Me.TBScoreRemain.Text = Format(vScore, "##,##0.00")
        '        End If
        '    End If
        'Else
        '    Me.TBArName.Text = ""
        '    Me.TBArCode.Text = ""
        '    Me.TBArAddress.Text = ""
        '    Me.TBApplyDate.Text = ""
        '    Me.TBExpireDate.Text = ""
        '    Me.TBArScore.Text = ""
        '    vMemberStatus = 0
        '    vMemIsMember = 0
        '    vScore = 0
        '    Me.ListViewAccumulateScore.Items.Clear()
        '    Me.ListViewIssueScore.Items.Clear()
        '    Call vGetAcculateLine()
        '    Call vGetIssueLine()
        '    Call ClearSpecialScore()
        '    Me.PNAddSpecialScore.Visible = False
        '    Me.PNAccumulateScore.Visible = True
        '    Me.TBScoreRemain.Text = Format(vScore, "##,##0.00")
        'End If
    End Sub

    Private Sub CMBCampaign_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBCampaign.SelectedIndexChanged
        Dim vCampaignCode As String
        Dim vARCode As String
        Dim vScore As Double

        On Error Resume Next

        If Me.CMBCampaign.Text <> "" Then
            vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)

            vQuery = "exec dbo.usp_vp_campaignsearch '" & vCampaignCode & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Campaign")
            dt = ds.Tables("Campaign")
            If dt.Rows.Count > 0 Then
                Me.TBStartDate.Text = dt.Rows(0).Item("startdate")
                Me.TBEndDate.Text = dt.Rows(0).Item("stopdate")
                Me.DTPDate1.Text = dt.Rows(0).Item("startdate")
                Me.DTPDate2.Text = dt.Rows(0).Item("stopdate")
            End If

            If vMemIsMember <> 0 Then
                If Me.CMBCampaign.Text <> "" And Me.TBArCode.Text <> "" Then
                    vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
                    vArCode = Me.TBArCode.Text
                    vQuery = "exec dbo.usp_vp_checkpoint '" & vCampaignCode & "','" & vArCode & "'"
                    da = New SqlDataAdapter(vQuery, vConnection)
                    ds = New DataSet
                    da.Fill(ds, "Score")
                    dt = ds.Tables("Score")
                    If dt.Rows.Count > 0 Then
                        vScore = dt.Rows(0).Item("point")
                    Else
                        vScore = 0
                    End If
                    Me.TBScoreRemain.Text = Format(vScore, "##,##0.00")
                End If
            End If

        End If
    End Sub

    Private Sub BTNSearchMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchMember.Click
        Call vSearchMember("")
    End Sub

    Private Sub TBSearchMember_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchMember.KeyDown
        Dim vSearch As String

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.TBSearchMember.Text <> "" Then
                vSearch = Me.TBSearchMember.Text
                Call vSearchMember(vSearch)
            Else
                Call vSearchMember("")
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchMember.Visible = False

            If Me.TBArName.Text = "" Then
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If

            If Me.TBArName.Text <> "" And Me.TBArCode.Text <> "" Then
                Me.CMBCampaign.Focus()
            End If
        End If

        If e.KeyCode = Keys.Down Then
            If Me.ListViewSearchMember.Items.Count > 0 Then
                Me.ListViewSearchMember.Focus()
                Me.ListViewSearchMember.Items(0).Focused = True
                Me.ListViewSearchMember.Items(0).Selected = True
            End If
        End If
    End Sub

    Private Sub TBSearchMember_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchMember.TextChanged

    End Sub

    Private Sub ListViewSearchMember_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewSearchMember.DoubleClick
        Dim vIndex As Integer
        Dim vMemberID As String

        On Error Resume Next

        If Me.ListViewSearchMember.Items.Count > 0 Then
            vIndex = Me.ListViewSearchMember.SelectedItems(0).Index
            Me.TBMemberID.Text = Me.ListViewSearchMember.Items(vIndex).SubItems(1).Text
            Me.TBSearchMember.Text = ""
            Me.ListViewSearchMember.Items.Clear()
            Me.PNSearchMember.Visible = False
            Call ClearSpecialScore()
            Me.PNAddSpecialScore.Visible = False
            Me.PNAccumulateScore.Visible = True
            vMemberID = Me.TBMemberID.Text
            'Call SearchMemberDetails(vMemberID)
            Me.TBMemberID.Focus()
            Me.TBMemberID.SelectAll()
        End If
    End Sub

    Private Sub ListViewSearchMember_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearchMember.KeyDown
        Dim vIndex As Integer
        Dim vMemberID As String

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewSearchMember.Items.Count > 0 Then
                vIndex = Me.ListViewSearchMember.SelectedItems(0).Index
                Me.TBMemberID.Text = Me.ListViewSearchMember.Items(vIndex).SubItems(1).Text
                Me.TBSearchMember.Text = ""
                Me.ListViewSearchMember.Items.Clear()
                Me.PNSearchMember.Visible = False
                Call ClearSpecialScore()
                Me.PNAddSpecialScore.Visible = False
                Me.PNAccumulateScore.Visible = True
                vMemberID = Me.TBMemberID.Text
                Call SearchMemberDetails(vMemberID)
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchMember.Visible = False

            If Me.TBArName.Text = "" Then
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If

            If Me.TBArName.Text <> "" And Me.TBArCode.Text <> "" Then
                Me.CMBCampaign.Focus()
            End If
        End If

        Dim vCheckLine As Integer

        If Me.ListViewSearchMember.Items.Count > 0 Then
            If e.KeyCode = Keys.Up Then
                vCheckLine = Me.ListViewSearchMember.SelectedItems(0).Index
                If vCheckLine = 0 Then
                    Me.TBSearchMember.Focus()
                    Me.TBSearchMember.SelectAll()
                End If
            End If
        End If
    End Sub

    Private Sub ListViewSearchMember_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewSearchMember.SelectedIndexChanged

    End Sub

    Public Sub CheckScoreHist(ByVal vARCode As String, ByVal vCampaign As String, ByVal vIsMember As Integer)
        Dim i As Integer
        Dim vListItem As ListViewItem
        Dim n As Integer
        Dim vQty As Double
        Dim vItemPoint As Double
        Dim vPrice As Double
        Dim vDiscountTotal As Double
        Dim vAmount As Double

        On Error Resume Next

        If Me.TBArName.Text <> "" Then
            If vIsMember <> 0 Then
                Me.ListViewAccumulateScore.Items.Clear()
                vQuery = "exec dbo.usp_vp_pointdescsub1 '" & vCampaign & "','" & vARCode & "'"
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "ScoreHist")
                dt = ds.Tables("ScoreHist")
                If dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        n = n + 1
                        vListItem = Me.ListViewAccumulateScore.Items.Add(n)
                        vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("docdate")
                        vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("docno")
                        vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("itemcode")
                        vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("itemname")
                        vQty = dt.Rows(i).Item("qty")
                        vItemPoint = dt.Rows(i).Item("pointfinal")
                        vPrice = dt.Rows(i).Item("price")
                        vDiscountTotal = dt.Rows(i).Item("discounttotal")
                        vAmount = dt.Rows(i).Item("amount")
                        vListItem.SubItems.Add(4).Text = Format(vQty, "##,##0.00")
                        vListItem.SubItems.Add(5).Text = Format(vItemPoint, "##,##0.00")
                        vListItem.SubItems.Add(6).Text = dt.Rows(i).Item("unitcode")
                        vListItem.SubItems.Add(7).Text = Format(vPrice, "##,##0.00")
                        vListItem.SubItems.Add(8).Text = Format(vDiscountTotal, "##,##0.00")
                        vListItem.SubItems.Add(9).Text = Format(vAmount, "##,##0.00")
                    Next

                    Dim a As Integer
                    If Me.ListViewAccumulateScore.Items.Count > 0 Then
                        For a = 0 To Me.ListViewAccumulateScore.Items.Count - 1
                            If a Mod 2 = 0 Then
                                Me.ListViewAccumulateScore.Items(a).BackColor = Color.AliceBlue
                            Else
                                Me.ListViewAccumulateScore.Items(a).BackColor = Color.LightGreen
                            End If
                        Next
                    End If
                Else
                    Call vGetAcculateLine()
                End If

            End If
        End If

    End Sub

    Public Sub CheckScoreHistFilter(ByVal vARCode As String, ByVal vCampaign As String, ByVal vIsMember As Integer, ByVal vStartDate As String, ByVal vStopDate As String, ByVal vInvoiceNo As String)
        Dim i As Integer
        Dim vListItem As ListViewItem
        Dim n As Integer
        Dim vQty As Double
        Dim vItemPoint As Double
        Dim vPrice As Double
        Dim vDiscountTotal As Double
        Dim vAmount As Double

        On Error Resume Next

        If Me.TBArName.Text <> "" Then
            If vIsMember <> 0 Then
                Me.ListViewAccumulateScore.Items.Clear()
                vQuery = "exec dbo.usp_vp_pointdescsub2 '" & vCampaign & "','" & vARCode & "','" & vStartDate & "','" & vStopDate & "','" & vInvoiceNo & "'"
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "ScoreHist")
                dt = ds.Tables("ScoreHist")
                If dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        n = n + 1
                        vListItem = Me.ListViewAccumulateScore.Items.Add(n)
                        vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("docdate")
                        vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("docno")
                        vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("itemcode")
                        vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("itemname")
                        vQty = dt.Rows(i).Item("qty")
                        vItemPoint = dt.Rows(i).Item("pointfinal")
                        vPrice = dt.Rows(i).Item("price")
                        vDiscountTotal = dt.Rows(i).Item("discounttotal")
                        vAmount = dt.Rows(i).Item("amount")
                        vListItem.SubItems.Add(4).Text = Format(vQty, "##,##0.00")
                        vListItem.SubItems.Add(5).Text = Format(vItemPoint, "##,##0.00")
                        vListItem.SubItems.Add(6).Text = dt.Rows(i).Item("unitcode")
                        vListItem.SubItems.Add(7).Text = Format(vPrice, "##,##0.00")
                        vListItem.SubItems.Add(8).Text = Format(vDiscountTotal, "##,##0.00")
                        vListItem.SubItems.Add(9).Text = Format(vAmount, "##,##0.00")
                    Next

                    Dim a As Integer
                    If Me.ListViewAccumulateScore.Items.Count > 0 Then
                        For a = 0 To Me.ListViewAccumulateScore.Items.Count - 1
                            If a Mod 2 = 0 Then
                                Me.ListViewAccumulateScore.Items(a).BackColor = Color.AliceBlue
                            Else
                                Me.ListViewAccumulateScore.Items(a).BackColor = Color.LightGreen
                            End If
                        Next
                    End If
                Else
                    Call vGetAcculateLine()
                End If

            End If
        End If

    End Sub

    Public Sub CheckScoreHistFilter1(ByVal vARCode As String, ByVal vCampaign As String, ByVal vIsMember As Integer, ByVal vStartDate As String, ByVal vStopDate As String)
        Dim i As Integer
        Dim vListItem As ListViewItem
        Dim n As Integer
        Dim vQty As Double
        Dim vItemPoint As Double
        Dim vPrice As Double
        Dim vDiscountTotal As Double
        Dim vAmount As Double

        On Error Resume Next

        If Me.TBArName.Text <> "" Then
            If vIsMember <> 0 Then
                Me.ListViewAccumulateScore.Items.Clear()
                vQuery = "exec dbo.usp_vp_pointdescsub3 '" & vCampaign & "','" & vARCode & "','" & vStartDate & "','" & vStopDate & "'"
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "ScoreHist")
                dt = ds.Tables("ScoreHist")
                If dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        n = n + 1
                        vListItem = Me.ListViewAccumulateScore.Items.Add(n)
                        vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("docdate")
                        vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("docno")
                        vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("itemcode")
                        vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("itemname")
                        vQty = dt.Rows(i).Item("qty")
                        vItemPoint = dt.Rows(i).Item("pointfinal")
                        vPrice = dt.Rows(i).Item("price")
                        vDiscountTotal = dt.Rows(i).Item("discounttotal")
                        vAmount = dt.Rows(i).Item("amount")
                        vListItem.SubItems.Add(4).Text = Format(vQty, "##,##0.00")
                        vListItem.SubItems.Add(5).Text = Format(vItemPoint, "##,##0.00")
                        vListItem.SubItems.Add(6).Text = dt.Rows(i).Item("unitcode")
                        vListItem.SubItems.Add(7).Text = Format(vPrice, "##,##0.00")
                        vListItem.SubItems.Add(8).Text = Format(vDiscountTotal, "##,##0.00")
                        vListItem.SubItems.Add(9).Text = Format(vAmount, "##,##0.00")
                    Next

                    Dim a As Integer
                    If Me.ListViewAccumulateScore.Items.Count > 0 Then
                        For a = 0 To Me.ListViewAccumulateScore.Items.Count - 1
                            If a Mod 2 = 0 Then
                                Me.ListViewAccumulateScore.Items(a).BackColor = Color.AliceBlue
                            Else
                                Me.ListViewAccumulateScore.Items(a).BackColor = Color.LightGreen
                            End If
                        Next
                    End If
                Else
                    Call vGetAcculateLine()
                End If

            End If
        End If

    End Sub

    Public Sub vCheckIssueSore(ByVal vARCode As String, ByVal vCampaignCode As String)
        Dim i As Integer
        Dim n As Integer
        Dim vListItem As ListViewItem
        Dim vAmount As Double
        Dim vPoint As Double

        On Error Resume Next

        Me.ListViewIssueScore.Items.Clear()
        vQuery = "exec dbo.usp_vp_checkpointwithdraw2010 '" & vCampaignCode & "','" & vARCode & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "IssueScore")
        dt = ds.Tables("IssueScore")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vListItem = Me.ListViewIssueScore.Items.Add(n)
                vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("docdate")
                vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("docno")
                vAmount = dt.Rows(i).Item("amount")
                vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("couponno")
                vListItem.SubItems.Add(3).Text = Format(vAmount, "##,##0.00")
                vListItem.SubItems.Add(4).Text = dt.Rows(i).Item("creatorname")
            Next

            Dim a As Integer
            If Me.ListViewIssueScore.Items.Count > 0 Then
                For a = 0 To Me.ListViewIssueScore.Items.Count - 1
                    If a Mod 2 = 0 Then
                        Me.ListViewIssueScore.Items(a).BackColor = Color.AliceBlue
                    Else
                        Me.ListViewIssueScore.Items(a).BackColor = Color.LightBlue
                    End If
                Next
            End If
        Else
            Call vGetIssueLine()
        End If
    End Sub

    Public Sub vGetAcculateLine()
        Dim i As Integer
        Dim vListItem As ListViewItem

        On Error Resume Next

        For i = 0 To 8
            vListItem = Me.ListViewAccumulateScore.Items.Add("")
            vListItem.SubItems.Add(0).Text = ""
            vListItem.SubItems.Add(1).Text = ""
            vListItem.SubItems.Add(2).Text = ""
            vListItem.SubItems.Add(3).Text = ""
            vListItem.SubItems.Add(4).Text = ""
            vListItem.SubItems.Add(5).Text = ""
            vListItem.SubItems.Add(6).Text = ""
            vListItem.SubItems.Add(7).Text = ""
            vListItem.SubItems.Add(8).Text = ""
        Next

        Dim a As Integer
        If Me.ListViewAccumulateScore.Items.Count > 0 Then
            For a = 0 To Me.ListViewAccumulateScore.Items.Count - 1
                If a Mod 2 = 0 Then
                    Me.ListViewAccumulateScore.Items(a).BackColor = Color.AliceBlue
                Else
                    Me.ListViewAccumulateScore.Items(a).BackColor = Color.LightGreen
                End If
            Next
        End If
    End Sub

    Public Sub vGetIssueLine()
        Dim i As Integer
        Dim vListItem As ListViewItem

        On Error Resume Next

        For i = 0 To 14
            vListItem = Me.ListViewIssueScore.Items.Add("")
            vListItem.SubItems.Add(0).Text = ""
            vListItem.SubItems.Add(1).Text = ""
            vListItem.SubItems.Add(2).Text = ""
            vListItem.SubItems.Add(3).Text = ""
            vListItem.SubItems.Add(4).Text = ""
            vListItem.SubItems.Add(5).Text = ""
        Next

        Dim a As Integer
        If Me.ListViewIssueScore.Items.Count > 0 Then
            For a = 0 To Me.ListViewIssueScore.Items.Count - 1
                If a Mod 2 = 0 Then
                    Me.ListViewIssueScore.Items(a).BackColor = Color.AliceBlue
                Else
                    Me.ListViewIssueScore.Items(a).BackColor = Color.LightBlue
                End If
            Next
        End If
    End Sub

    Public Sub vGetSearchMemberLine()
        Dim i As Integer
        Dim vListItem As ListViewItem

        On Error Resume Next

        For i = 0 To 16
            vListItem = Me.ListViewSearchMember.Items.Add("")
            vListItem.SubItems.Add(0).Text = ""
            vListItem.SubItems.Add(1).Text = ""
            vListItem.SubItems.Add(2).Text = ""
            vListItem.SubItems.Add(3).Text = ""
            vListItem.SubItems.Add(4).Text = ""
        Next

        Dim a As Integer
        If Me.ListViewSearchMember.Items.Count > 0 Then
            For a = 0 To Me.ListViewSearchMember.Items.Count - 1
                If a Mod 2 = 0 Then
                    Me.ListViewSearchMember.Items(a).BackColor = Color.AliceBlue
                Else
                    Me.ListViewSearchMember.Items(a).BackColor = Color.LavenderBlush
                End If
            Next
        End If
    End Sub

    Public Sub vCalcArScore()
        Dim i As Integer
        Dim vScore As Double
        Dim vSumScore As Double

        On Error Resume Next

        If Me.ListViewAccumulateScore.Items.Count > 0 Then
            For i = 0 To Me.ListViewAccumulateScore.Items.Count - 1
                If Me.ListViewAccumulateScore.Items(i).SubItems(6).Text <> "" Then
                    vScore = Me.ListViewAccumulateScore.Items(i).SubItems(6).Text
                Else
                    vScore = 0
                End If
                vSumScore = vSumScore + vScore
            Next
        End If

        Me.TBArScore.Text = Format(vSumScore, "##,##0.00")
    End Sub

    Public Sub vCalcArIssueScoreHist()
        'Dim i As Integer
        'Dim vScore As Double
        'Dim vSumScore As Double

        'If Me.ListViewIssueScore.Items.Count > 0 Then
        '    For i = 0 To Me.ListViewIssueScore.Items.Count - 1
        '        If Me.ListViewIssueScore.Items(i).SubItems(4).Text <> "" Then
        '            vScore = Me.ListViewIssueScore.Items(i).SubItems(4).Text
        '        Else
        '            vScore = 0
        '        End If
        '        vSumScore = vSumScore + vScore
        '    Next
        'End If

        'Me.TBIssueScoreHist.Text = Format(vSumScore, "##,##0.00")
    End Sub

    Private Sub BTNFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNFilter.Click
        Dim vInvoiceNo As String
        Dim vCampaignCode As String
        Dim vARCode As String
        Dim vStartDate As String
        Dim vStopDate As String
        Dim vCheckStart As Date
        Dim vCheckStop As Date

        On Error Resume Next

        If Me.TBInvoiceNo.Text <> "" Then
            If Me.CMBCampaign.Text = "" Then
                MsgBox("กรุณา กรอกรหัสแคมเปญ", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBCampaign.Focus()
                Me.CMBCampaign.SelectAll()
                Exit Sub
            End If

            If Me.TBArCode.Text = "" Then
                MsgBox("กรุณา กรอกรหัสลูกค้า", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBArCode.Focus()
                Me.TBArCode.SelectAll()
                Exit Sub
            End If

            vInvoiceNo = Me.TBInvoiceNo.Text
            vARCode = Me.TBArCode.Text
            vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
            vStartDate = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
            vStopDate = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)
            Call CheckScoreHistFilter(vARCode, vCampaignCode, vMemIsMember, vStartDate, vStopDate, vInvoiceNo)
            Call vCalcArScore()
            Call vCalcArIssueScoreHist()
        End If

        If Me.TBInvoiceNo.Text = "" Then
            If Me.CMBCampaign.Text = "" Then
                MsgBox("กรุณา กรอกรหัสแคมเปญ", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBCampaign.Focus()
                Me.CMBCampaign.SelectAll()
                Exit Sub
            End If

            If Me.TBArCode.Text = "" Then
                MsgBox("กรุณา กรอกรหัสลูกค้า", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBArCode.Focus()
                Me.TBArCode.SelectAll()
                Exit Sub
            End If

            vCheckStart = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
            vCheckStop = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)

            If vCheckStart > vCheckStop Then
                MsgBox("ไม่สามารถเลือกวันที่เริ่มค้นหา มากกว่าวันที่สิ้นสุด", MsgBoxStyle.Critical, "Send Error Message")
                Me.ListViewAccumulateScore.Items.Clear()
                Call vGetAcculateLine()
                Me.DTPDate1.Focus()
                Exit Sub
            End If

            vInvoiceNo = ""
            vARCode = Me.TBArCode.Text
            vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
            vStartDate = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
            vStopDate = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)

            Call CheckScoreHistFilter1(vARCode, vCampaignCode, vMemIsMember, vStartDate, vStopDate)
            Call vCalcArScore()
            Call vCalcArIssueScoreHist()
        End If

    End Sub

    Private Sub DTPDate2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTPDate2.LostFocus
        'Dim vARCode As String
        'Dim vCampaignCode As String
        'Dim vStartDate As String
        'Dim vStopDate As String
        'Dim vInvoiceNo As String
        'Dim vCheckStart As Date
        'Dim vCheckStop As Date

        'If Me.TBInvoiceNo.Text = "" Then

        '    If Me.CMBCampaign.Text = "" Then
        '        Exit Sub
        '    End If

        '    If Me.TBArCode.Text = "" Then
        '        Exit Sub
        '    End If

        '    If vCheckStart > vCheckStop Then
        '        Me.ListViewAccumulateScore.Items.Clear()
        '        Call vGetAcculateLine()
        '        Me.DTPDate2.Focus()
        '        'MsgBox("ไม่สามารถเลือกวันที่เริ่มค้นหา มากกว่าวันที่สิ้นสุด", MsgBoxStyle.Critical, "Send Error Message")
        '        Exit Sub
        '    End If

        '    vInvoiceNo = ""
        '    vARCode = Me.TBArCode.Text
        '    vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
        '    vStartDate = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
        '    vStopDate = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)

        '    Call CheckScoreHistFilter1(vARCode, vCampaignCode, vMemIsMember, vStartDate, vStopDate)
        '    Call vCalcArScore()
        'End If
    End Sub

    Private Sub DTPDate2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTPDate2.ValueChanged
        'Me.TBSpecialNo.Focus()

        'Dim vARCode As String
        'Dim vCampaignCode As String
        'Dim vStartDate As String
        'Dim vStopDate As String
        'Dim vInvoiceNo As String
        'Dim vCheckStart As Date
        'Dim vCheckStop As Date

        'If Me.TBInvoiceNo.Text = "" Then

        '    If Me.CMBCampaign.Text = "" Then
        '        Exit Sub
        '    End If

        '    If Me.TBArCode.Text = "" Then
        '        Exit Sub
        '    End If

        '    If vCheckStart > vCheckStop Then
        '        Me.ListViewAccumulateScore.Items.Clear()
        '        Call vGetAcculateLine()
        '        Me.DTPDate2.Focus()
        '        'MsgBox("ไม่สามารถเลือกวันที่เริ่มค้นหา มากกว่าวันที่สิ้นสุด", MsgBoxStyle.Critical, "Send Error Message")
        '        Exit Sub
        '    End If

        '    vInvoiceNo = ""
        '    vARCode = Me.TBArCode.Text
        '    vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
        '    vStartDate = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
        '    vStopDate = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)

        '    Call CheckScoreHistFilter1(vARCode, vCampaignCode, vMemIsMember, vStartDate, vStopDate)
        '    Call vCalcArScore()
        'End If
    End Sub

    Private Sub DTPDate1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTPDate1.LostFocus
        'Dim vARCode As String
        'Dim vCampaignCode As String
        'Dim vStartDate As String
        'Dim vStopDate As String
        'Dim vInvoiceNo As String
        'Dim vCheckStart As Date
        'Dim vCheckStop As Date

        'If Me.TBInvoiceNo.Text = "" Then

        '    If Me.CMBCampaign.Text = "" Then
        '        Exit Sub
        '    End If

        '    If Me.TBArCode.Text = "" Then
        '        Exit Sub
        '    End If

        '    vCheckStart = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
        '    vCheckStop = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)

        '    If vCheckStart > vCheckStop Then
        '        'MsgBox("ไม่สามารถเลือกวันที่เริ่มค้นหา มากกว่าวันที่สิ้นสุด", MsgBoxStyle.Critical, "Send Error Message")
        '        Me.ListViewAccumulateScore.Items.Clear()
        '        Call vGetAcculateLine()
        '        Me.DTPDate1.Focus()
        '        Exit Sub
        '    End If

        '    vInvoiceNo = ""
        '    vARCode = Me.TBArCode.Text
        '    vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
        '    vStartDate = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
        '    vStopDate = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)

        '    Call CheckScoreHistFilter1(vARCode, vCampaignCode, vMemIsMember, vStartDate, vStopDate)
        '    Call vCalcArScore()
        'End If
    End Sub

    Private Sub DTPDate1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTPDate1.ValueChanged
        'Me.DTPDate2.Focus()
        'Dim vARCode As String
        'Dim vCampaignCode As String
        'Dim vStartDate As String
        'Dim vStopDate As String
        'Dim vInvoiceNo As String
        'Dim vCheckStart As Date
        'Dim vCheckStop As Date

        'If Me.TBInvoiceNo.Text = "" Then

        '    If Me.CMBCampaign.Text = "" Then
        '        Exit Sub
        '    End If

        '    If Me.TBArCode.Text = "" Then
        '        Exit Sub
        '    End If

        '    vCheckStart = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
        '    vCheckStop = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)

        '    If vCheckStart > vCheckStop Then
        '        'MsgBox("ไม่สามารถเลือกวันที่เริ่มค้นหา มากกว่าวันที่สิ้นสุด", MsgBoxStyle.Critical, "Send Error Message")
        '        Me.ListViewAccumulateScore.Items.Clear()
        '        Call vGetAcculateLine()
        '        Me.DTPDate1.Focus()
        '        Exit Sub
        '    End If

        '    vInvoiceNo = ""
        '    vARCode = Me.TBArCode.Text
        '    vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
        '    vStartDate = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
        '    vStopDate = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)

        '    Call CheckScoreHistFilter1(vARCode, vCampaignCode, vMemIsMember, vStartDate, vStopDate)
        '    Call vCalcArScore()
        'End If
    End Sub

    Private Sub BTNSpecialScoreClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSpecialScoreClear.Click
        Call ClearSpecialScore()
    End Sub

    Public Sub ClearSpecialScore()
        On Error Resume Next

        vMemSpecialCancel = 0
        vMemSpecialConfirm = 0
        vMemSpecialOpen = 0

        Me.TBSpecialNo.Text = ""
        Me.DTPSpecialDate.Value = Now
        If Me.CMBSpecialCampaign.Items.Count > 0 Then
            Me.CMBSpecialCampaign.SelectedIndex = 0
        Else
            Me.CMBSpecialCampaign.Text = ""
        End If
        Me.PNSearchSpecialScore.Visible = False
        Me.TBSpecialReason.Text = ""
        Me.TBSpecialScore.Text = ""
        Call vGetCampaignSpecial("")
        Call vGetSpecialDocNo()
        Call vGetNewSpecial()
    End Sub

    Public Sub vGetNewSpecial()
        On Error Resume Next

        Me.PBAddCancel.Visible = False
        Me.PBAddConfirm.Visible = False
        Me.PBAddNew.Visible = True
    End Sub

    Public Sub vGetConfirmSpecial()
        On Error Resume Next

        Me.PBAddCancel.Visible = False
        Me.PBAddConfirm.Visible = True
        Me.PBAddNew.Visible = False
    End Sub

    Public Sub vGetCancelSpecial()
        On Error Resume Next

        Me.PBAddCancel.Visible = True
        Me.PBAddConfirm.Visible = False
        Me.PBAddNew.Visible = False
    End Sub


    Private Sub BTNSpecialSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSpecialSave.Click
        Dim vDocno As String
        Dim vARCode As String
        Dim vDocDate As String
        Dim vCampaignCode As String
        Dim vScore As String
        Dim vReason As String
        Dim vIsInsert As Integer
        Dim vMemBeginTran As Integer

        If vMemIsMember <> 0 Then
            If Me.TBArName.Text = "" Then
                MsgBox("ยังไม่ได้ระบุรหัสสมาชิก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
                Exit Sub
            End If

            If Me.TBSpecialNo.Text = "" Then
                MsgBox("ยังไม่ได้ระบุ เลขที่เบิกแต้ม กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Call vGetSpecialDocNo()
                Exit Sub
            End If

            vDocno = Me.TBSpecialNo.Text

            If Me.TBSpecialScore.Text = "" Then
                MsgBox("ยังไม่ได้ระบุแต้มที่จะเพิ่ม กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBSpecialScore.Focus()
                Me.TBSpecialScore.SelectAll()
                Exit Sub
            End If

            If vMemSpecialCancel = 1 Then
                MsgBox("เอกสารเพิ่มแต้มพิเศษเลขที่ " & vDocno & " ได้ถูกยกเลิกไปแล้ว ไม่สามารถบันทึกแก้ไขได้อีก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBSpecialNo.Focus()
                Me.TBSpecialNo.SelectAll()
                Exit Sub
            End If

            If vMemSpecialConfirm = 1 Then
                MsgBox("เอกสารเพิ่มแต้มพิเศษเลขที่ " & vDocno & " ได้ถูกอนุมัติไปแล้ว ไม่สามารถบันทึกแก้ไขได้อีก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBSpecialNo.Focus()
                Me.TBSpecialNo.SelectAll()
                Exit Sub
            End If

            If vMemSpecialOpen = 0 Then
                Call vGetSpecialDocNo()
                vIsInsert = 1
            Else
                vIsInsert = 0
            End If

            vARCode = Me.TBArCode.Text
            vDocno = Me.TBSpecialNo.Text
            If vb6.Year(Me.DTPSpecialDate.Value) >= 2500 Then
                vDocDate = vb6.Day(Me.DTPSpecialDate.Value) & "/" & vb6.Month(Me.DTPSpecialDate.Value) & "/" & vb6.Year(Me.DTPSpecialDate.Value) - 543
            Else
                vDocDate = vb6.Day(Me.DTPSpecialDate.Value) & "/" & vb6.Month(Me.DTPSpecialDate.Value) & "/" & vb6.Year(Me.DTPSpecialDate.Value)
            End If
            vCampaignCode = vb6.Left(Me.CMBSpecialCampaign.Text, vb6.InStr(Me.CMBSpecialCampaign.Text, "/") - 1)
            vScore = Me.TBSpecialScore.Text
            vReason = Me.TBSpecialReason.Text

            vMemBeginTran = 1
            vQuery = "begin tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            vQuery = "exec dbo.USP_VP_PointSpecialSet " & vIsInsert & ",'" & vCampaignCode & "','" & vDocno & "','" & vDocDate & "','" & vARCode & "'," & vScore & ",'" & vReason & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            vQuery = "commit tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            MsgBox("บันทึกข้อมูลเบิกแต้มพิเศษเรียบร้อย กรุณาตรวจสอบ", MsgBoxStyle.Information, "send information Message")
            vMemBeginTran = 0

            vARCode = Me.TBArCode.Text
            vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)

            vQuery = "exec dbo.USP_VP_CalMemberPointFromInvoiceHistory2010 '" & vCampaignCode & "','" & vARCode & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            Call ClearSpecialScore()
            Me.TBSpecialNo.Focus()
            Me.TBSpecialNo.SelectAll()

        Else
            MsgBox("ลูกค้ายังไม่ได้เป็นสมาชิกไม่สามารถเบิกแต้มพิเศษได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

ErrDescription:
            If Err.Description <> "" Then
                If vMemBeginTran = 1 Then
                    vQuery = "rollback tran"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()
                End If
                MsgBox(Err.Description, MsgBoxStyle.Critical, "Error")
                Exit Sub
            End If
    End Sub

    Private Sub BTNCalcScore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCalcScore.Click
        Dim vARCode As String
        Dim vCampaignCode As String
        Dim vAnswer As Integer

        On Error Resume Next

        If Me.TBArName.Text <> "" Then
            If Me.CMBCampaign.Text <> "" Then
                If vMemIsMember <> 0 Then
                    vAnswer = MsgBox("คุณต้องการคำนวณแต้ม ใช่หรือไม่ ?", MsgBoxStyle.YesNo, "Send Question Message")
                    If vAnswer = 6 Then
                        vARCode = Me.TBArCode.Text
                        vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)

                        vQuery = "exec dbo.USP_VP_CalMemberPointFromInvoiceHistory2010 '" & vCampaignCode & "','" & vARCode & "'"
                        cmd = New SqlCommand(vQuery, vConnection)
                        cmd.ExecuteNonQuery()

                        MsgBox("ส่งรหัสสมาชิกไปคำนวณแต้มเรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")
                        Me.TBMemberID.Focus()
                        Me.TBMemberID.SelectAll()
                    Else
                        Me.TBMemberID.Focus()
                        Me.TBMemberID.SelectAll()
                    End If
                Else
                    MsgBox("ลูกค้าที่เป็นสมาชิกเท่านั้นถึงจะคำนวณแต้มได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
            Else
                MsgBox("กรุณากรอกรหัสแคมเปญ ที่ต้องการคำนวณแต้ม กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBCampaign.Focus()
            End If
        Else

            MsgBox("กรุณากรอกรหัสสมาชิก ที่ต้องการคำนวณแต้ม กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBMemberID.Focus()
            Me.TBMemberID.SelectAll()
        End If
    End Sub

    Public Sub CalcMemberRemain()
        Dim vARCode As String
        Dim vCampaignCode As String
        Dim vAnswer As Integer

        On Error Resume Next

        If Me.TBArName.Text <> "" Then
            If Me.CMBCampaign.Text <> "" Then
                If vMemIsMember <> 0 Then
                    vAnswer = MsgBox("คุณต้องการคำนวณแต้ม ใช่หรือไม่ ?", MsgBoxStyle.YesNo, "Send Question Message")
                    If vAnswer = 6 Then
                        vARCode = Me.TBArCode.Text
                        vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)

                        vQuery = "exec dbo.USP_VP_CalMemberPointFromInvoiceHistory2010 '" & vCampaignCode & "','" & vARCode & "'"
                        cmd = New SqlCommand(vQuery, vConnection)
                        cmd.ExecuteNonQuery()

                        MsgBox("ส่งรหัสสมาชิกไปคำนวณแต้มเรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")
                        Me.TBMemberID.Focus()
                        Me.TBMemberID.SelectAll()
                    Else
                        Me.TBMemberID.Focus()
                        Me.TBMemberID.SelectAll()
                    End If
                Else
                    MsgBox("ลูกค้าที่เป็นสมาชิกเท่านั้นถึงจะคำนวณแต้มได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
            Else
                MsgBox("กรุณากรอกรหัสแคมเปญ ที่ต้องการคำนวณแต้ม กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBCampaign.Focus()
            End If
        Else
            MsgBox("กรุณากรอกรหัสสมาชิก ที่ต้องการคำนวณแต้ม กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBMemberID.Focus()
            Me.TBMemberID.SelectAll()
        End If
    End Sub

    Private Sub ListViewSearchSpecialScore_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewSearchSpecialScore.DoubleClick
        Dim vIndex As Integer
        Dim vDocNo As String
        Dim vARCode As String

        On Error Resume Next

        If Me.TBArName.Text = "" Then
            MsgBox("", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
        vARCode = Me.TBArCode.Text
        If Me.ListViewSearchSpecialScore.Items.Count > 0 Then
            vIndex = Me.ListViewSearchSpecialScore.SelectedItems(0).Index
            vDocNo = Me.ListViewSearchSpecialScore.Items(vIndex).SubItems(1).Text
            Call vGetSpecialScoreDetails(vARCode, vDocNo)
        End If
    End Sub

    Public Sub vGetSpecialScoreDetails(ByVal vARCode As String, ByVal vDocNo As String)
        Dim vScore As Double

        On Error Resume Next

        vQuery = "exec dbo.USP_VP_PointSpecialSearchAr '" & vDocNo & "','" & vARCode & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "IssueScore")
        dt = ds.Tables("IssueScore")
        If dt.Rows.Count > 0 Then
            vMemSpecialCancel = dt.Rows(0).Item("iscancel")
            vMemSpecialConfirm = dt.Rows(0).Item("isconfirm")
            vMemSpecialOpen = 1
            vScore = dt.Rows(0).Item("point")
            Me.TBSpecialNo.Text = dt.Rows(0).Item("docno")
            Me.DTPSpecialDate.Value = dt.Rows(0).Item("docdate")
            Me.CMBSpecialCampaign.Text = dt.Rows(0).Item("campaigncode") & "/" & dt.Rows(0).Item("campaign")
            Me.TBSpecialScore.Text = Format(vScore, "##,##0.00")
            Me.TBSpecialReason.Text = dt.Rows(0).Item("reason")

            If vMemSpecialCancel = 1 Then
                Call vGetCancelSpecial()
            End If

            If vMemSpecialConfirm = 1 Then
                Call vGetConfirmSpecial()
            End If

            If vMemSpecialCancel = 0 And vMemSpecialConfirm = 0 Then
                Call vGetNewSpecial()
            End If

            Me.PNSearchSpecialScore.Visible = False
        Else
            Call vGetSpecialDocNo()
            vMemSpecialCancel = 0
            vMemSpecialConfirm = 0
            vMemSpecialOpen = 0
            Me.TBSpecialNo.Text = ""
            Me.DTPSpecialDate.Value = Now
            Me.CMBSpecialCampaign.SelectedIndex = 0
            Me.TBSpecialScore.Text = ""
            Me.TBSpecialReason.Text = ""
        End If
    End Sub

    Private Sub ListViewSearchSpecialScore_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewSearchSpecialScore.SelectedIndexChanged

    End Sub

    Private Sub BTNClickSearchMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClickSearchMember.Click
        Dim vSearch As String
        On Error Resume Next

        If Me.TBSearchMember.Text <> "" Then
            vSearch = Me.TBSearchMember.Text
            Call vSearchMember(vSearch)
        Else
            Call vSearchMember("")
        End If
    End Sub

    Private Sub BTNSelectSearchMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectSearchMember.Click
        Dim vIndex As Integer
        Dim vMemberID As String

        On Error Resume Next

        If Me.ListViewSearchMember.Items.Count > 0 Then
            vIndex = Me.ListViewSearchMember.SelectedItems(0).Index
            Me.TBMemberID.Text = Me.ListViewSearchMember.Items(vIndex).SubItems(1).Text
            Me.TBSearchMember.Text = ""
            Me.ListViewSearchMember.Items.Clear()
            Me.PNSearchMember.Visible = False
            Call ClearSpecialScore()
            Me.PNAddSpecialScore.Visible = False
            Me.PNAccumulateScore.Visible = True
            vMemberID = Me.TBMemberID.Text
            Call SearchMemberDetails(vMemberID)
            Me.TBMemberID.Focus()
            Me.TBMemberID.SelectAll()
        End If
    End Sub

    Private Sub BTNCloseSearchMember_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCloseSearchMember.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchMember.Visible = False

            If Me.TBArName.Text = "" Then
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If

            If Me.TBArName.Text <> "" And Me.TBArCode.Text <> "" Then
                Me.CMBCampaign.Focus()
            End If
        End If
    End Sub

    Private Sub BTNSelectSearchMember_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSelectSearchMember.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchMember.Visible = False

            If Me.TBArName.Text = "" Then
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If

            If Me.TBArName.Text <> "" And Me.TBArCode.Text <> "" Then
                Me.CMBCampaign.Focus()
            End If
        End If
    End Sub

    Private Sub BTNClickSearchMember_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClickSearchMember.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchMember.Visible = False

            If Me.TBArName.Text = "" Then
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If

            If Me.TBArName.Text <> "" And Me.TBArCode.Text <> "" Then
                Me.CMBCampaign.Focus()
            End If
        End If
    End Sub

    Private Sub ND1000_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ND1000.ValueChanged
        Call SumCounponAmount()
    End Sub

    Public Sub SumCounponAmount()
        Dim vAmount1000 As Double
        Dim vAmount500 As Double
        Dim vAmount200 As Double
        Dim vAmount100 As Double
        Dim vCount1000 As Integer
        Dim vCount500 As Integer
        Dim vCount200 As Integer
        Dim vCount100 As Integer
        Dim vSumCouponAmount As Double

        On Error Resume Next

        If Me.TBIssueScore.Text <> "" Then
            vCount1000 = Me.ND1000.Value
            vCount500 = Me.ND500.Value
            vCount200 = Me.ND200.Value
            vCount100 = Me.ND100.Value

            vAmount1000 = vCount1000 * 1000
            vAmount500 = vCount500 * 500
            vAmount200 = vCount200 * 200
            vAmount100 = vCount100 * 100

            vSumCouponAmount = vAmount1000 + vAmount500 + vAmount200 + vAmount100

            Me.TBCouponAmount.Text = Format(vSumCouponAmount, "##,##0.00")
        End If
    End Sub

    Public Sub vCheckCouponAmount()
        Dim vSumAmount As Double
        Dim vAmount1000 As Double
        Dim vAmount500 As Double
        Dim vAmount200 As Double
        Dim vAmount100 As Double
        Dim vCount1000 As Integer
        Dim vCount500 As Integer
        Dim vCount200 As Integer
        Dim vCount100 As Integer
        Dim vSumCouponAmount As Double
        Dim vCountCoupon As Integer

        On Error Resume Next

        If Me.TBIssueScore.Text <> "" Then
            vCount1000 = Me.ND1000.Value
            vCount500 = Me.ND500.Value
            vCount200 = Me.ND200.Value
            vCount100 = Me.ND100.Value

            vCountCoupon = vCount1000 + vCount500 + vCount200 + vCount100
            vAmount1000 = vCount1000 * 1000
            vAmount500 = vCount500 * 500
            vAmount200 = vCount200 * 200
            vAmount100 = vCount100 * 100

            vSumCouponAmount = vAmount1000 + vAmount500 + vAmount200 + vAmount100
            vSumAmount = Me.TBIssueScore.Text

            If vSumCouponAmount <> vSumAmount Then
                vMemCheckCouponAmount = 1
                Me.TBIssueScore.Focus()
                Me.TBIssueScore.SelectAll()
                MsgBox("มูลค่าแต้มที่เบิก ไม่เท่ากับ มูลค่าคูปองเงินสด กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            Else
                vMemCheckCouponAmount = 0
            End If

            If vSumCouponAmount <= 7200 And vCountCoupon > 8 Then
                Me.TBIssueScore.Focus()
                Me.TBIssueScore.SelectAll()
                MsgBox("เอกสารพิมพ์คูปอง  1 แผ่น สามารถพิมพ์คูปองได้ 8 ดวง ขั้นต่ำคือ มูลค่า 100 บาท มากสุด มูลค่า 1,000 บาท ซึ่งมูลค่าที่ไม่เกิน 7,200 บาท ยังจัดสรรการพิมพ์ให้อยู่ใน  1 แผ่นได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            End If

        End If
    End Sub

    Private Sub TBIssueScore_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBIssueScore.KeyDown
        Dim vSumAmount As Double
        Dim vMod100 As Double
        Dim vScore As Double
        Dim vMemberScore As Double
        Dim vOldScore As Double

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.TBArName.Text <> "" Then
                If vMemIssueOpen = 0 Then
                    If Me.TBScoreRemain.Text = "" Or Me.TBScoreRemain.Text = "0.00" Then
                        MsgBox("ไม่สามารถเบิกแต้มได้ เนื่องจากสมาชิกไม่มีแต้มคงเหลือ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "")
                        Call IssueClearScreen()
                        Me.TBIssueScore.Text = ""
                        Me.TBIssueScore.Focus()
                        Me.TBIssueScore.SelectAll()
                        Exit Sub
                    End If

                    vMemberScore = Me.TBScoreRemain.Text
                    If Me.TBIssueScore.Text <> "" Then
                        vScore = Me.TBIssueScore.Text
                    End If

                    If vScore > vMemberScore Then
                        MsgBox("ไม่สามารถเบิกคูปองเงินสดเกิน แต้มที่มีอยู่จริง กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                        Me.TBIssueScore.Text = ""
                        Me.TBIssueScore.Focus()
                        Me.TBIssueScore.SelectAll()
                        Exit Sub
                    End If
                End If

                If vMemIssueOpen = 1 Then
                    If vMemIssuePrintCount = 0 Then

                        If Me.TBScoreRemain.Text <> "" Then
                            vMemberScore = Me.TBScoreRemain.Text
                        Else
                            vMemberScore = 0
                        End If

                        If Me.TBIssueScore.Text <> "" Then
                            vScore = Me.TBIssueScore.Text
                        End If

                        If Me.TBMemOldScore.Text <> "" Then
                            vOldScore = Me.TBMemOldScore.Text
                        End If

                        vMemberScore = vMemberScore + vOldScore

                        If vScore > vMemberScore Then
                            MsgBox("ไม่สามารถเบิกแต้มมากกว่า แต้มคงเหลือ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "")
                            Me.TBIssueScore.Text = ""
                            Me.TBIssueScore.Focus()
                            Me.TBIssueScore.SelectAll()
                            Exit Sub
                        End If
                    Else
                        MsgBox("ไม่สามารถแก้ไขเอกสารได้ เนื่องจากพิมพ์คูปองไปแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Critical, "")
                        Call IssueClearScreen()
                        Exit Sub
                    End If
                End If

                If Me.TBIssueScore.Text <> "" Then
                    vSumAmount = Me.TBIssueScore.Text
                    vMod100 = vSumAmount Mod 100

                    If vMod100 > 0 Then
                        MsgBox("มูลค่าการเบิกคูปอง มีเศษของการเบิกอยู่ " & vMod100 & " บาท ซึ่งคูปองมีมูลค่าขั้นต่ำ 100 บาท ไม่สามารถพิมพ์คูปองได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                        Me.TBIssueScore.Text = ""
                        Me.TBIssueScore.Focus()
                        Me.TBIssueScore.SelectAll()
                        Exit Sub
                    End If

                    If vSumAmount >= 1000 Then
                        Me.ND1000.Focus()
                    ElseIf vSumAmount < 1000 And vSumAmount >= 500 Then
                        Me.ND500.Focus()
                    ElseIf vSumAmount < 500 And vSumAmount >= 200 Then
                        Me.ND200.Focus()
                    Else
                        Me.ND100.Focus()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TBIssueScore_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBIssueScore.KeyPress
        On Error GoTo ErrDescription

        Select Case Asc(e.KeyChar)
            Case 48 To 58, 8, 44, 45, 46, 37
            Case Else
                e.Handled = True
        End Select

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TBIssueScore_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TBIssueScore.LostFocus
        Dim vScore As Double
        Dim vMemberScore As Double
        Dim vOldScore As Double
        Dim vSumAmount As Double
        Dim vMod100 As Double

        On Error Resume Next

        If Me.TBIssueScore.Text <> "" Then
            If Me.TBArName.Text <> "" Then
                If vMemIssueOpen = 0 Then
                    If Me.TBScoreRemain.Text = "" Or Me.TBScoreRemain.Text = "0.00" Then
                        MsgBox("ไม่สามารถเบิกแต้มได้ เนื่องจากสมาชิกไม่มีแต้มคงเหลือ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "")
                        Call IssueClearScreen()
                        Me.TBIssueScore.Text = ""
                        Me.TBIssueScore.Focus()
                        Me.TBIssueScore.SelectAll()
                        Exit Sub
                    End If

                    vMemberScore = Me.TBScoreRemain.Text
                    If Me.TBIssueScore.Text <> "" Then
                        vScore = Me.TBIssueScore.Text
                    End If

                    If vScore > vMemberScore Then
                        MsgBox("ไม่สามารถเบิกคูปองเงินสดเกิน แต้มที่มีอยู่จริง กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                        Me.TBIssueScore.Text = ""
                        Me.TBIssueScore.Focus()
                        Me.TBIssueScore.SelectAll()
                        Exit Sub
                    End If
                End If

                If vMemIssueOpen = 1 Then
                    If vMemIssuePrintCount = 0 Then

                        If Me.TBScoreRemain.Text <> "" Then
                            vMemberScore = Me.TBScoreRemain.Text
                        Else
                            vMemberScore = 0
                        End If

                        If Me.TBIssueScore.Text <> "" Then
                            vScore = Me.TBIssueScore.Text
                        End If

                        If Me.TBMemOldScore.Text <> "" Then
                            vOldScore = Me.TBMemOldScore.Text
                        End If

                        vMemberScore = vMemberScore + vOldScore

                        If vScore > vMemberScore Then
                            MsgBox("ไม่สามารถเบิกแต้มมากกว่า แต้มคงเหลือ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "")
                            Me.TBIssueScore.Text = ""
                            Me.TBIssueScore.Focus()
                            Me.TBIssueScore.SelectAll()
                            Exit Sub
                        End If
                    Else
                        MsgBox("ไม่สามารถแก้ไขเอกสารได้ เนื่องจากพิมพ์คูปองไปแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Critical, "")
                        Call IssueClearScreen()
                        Exit Sub
                    End If
                End If

                If Me.TBIssueScore.Text <> "" Then
                    vSumAmount = Me.TBIssueScore.Text
                    vMod100 = vSumAmount Mod 100

                    If vMod100 > 0 Then
                        MsgBox("มูลค่าการเบิกคูปอง มีเศษของการเบิกอยู่ " & vMod100 & " บาท ซึ่งคูปองมีมูลค่าขั้นต่ำ 100 บาท ไม่สามารถพิมพ์คูปองได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                        Me.TBIssueScore.Text = ""
                        Me.TBIssueScore.Focus()
                        Me.TBIssueScore.SelectAll()
                        Exit Sub
                    End If

                    If vSumAmount >= 1000 Then
                        Me.ND1000.Focus()
                    ElseIf vSumAmount < 1000 And vSumAmount >= 500 Then
                        Me.ND500.Focus()
                    ElseIf vSumAmount < 500 And vSumAmount >= 200 Then
                        Me.ND200.Focus()
                    Else
                        Me.ND100.Focus()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub BTNIssueSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNIssueSave.Click
        Dim vDocno As String
        Dim vARCode As String
        Dim vDocDate As String
        Dim vCampaignCode As String
        Dim vScore As String
        Dim vIsInsert As Integer
        Dim i As Integer
        Dim vLineNumber As Integer
        Dim vCouponAmount As Double
        Dim vMemBeginTran As Integer
        Dim vMemberID As String

        'MsgBox("ยังไม่เปิดให้ใช้งาน ", MsgBoxStyle.Critical, "")
        'Exit Sub

        Call ChekAuthorityAccess()
        If vDepartment <> "CH" And vDepartment <> "IT" Then
            MsgBox("คุณไม่มีสิทธิ์ในการ บันทึกการเบิกแต้ม กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

        If vMemIsMember <> 0 Then
            If vMemIssuePrintCount = 0 Then
                If Me.TBArName.Text = "" Then
                    MsgBox("ไม่มีข้อมูลลูกค้า กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBMemberID.Focus()
                    Me.TBMemberID.SelectAll()
                    Exit Sub
                End If

                If Me.TBIssueNo.Text = "" Then
                    MsgBox("ไม่มีเลขที่เอกสาร กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Call vGetIssueDocNo()
                    Exit Sub
                End If

                If Me.TBIssueScore.Text = "" Then
                    MsgBox("ไม่ได้กรอก มูลค่าแต้ม กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBIssueScore.Focus()
                    Me.TBIssueScore.SelectAll()
                    Exit Sub
                End If

                If Me.TBCouponAmount.Text = "" Then
                    MsgBox("ยังไม่ได้เลือก จำนวนคูปองที่จะพิมพ์", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBCouponAmount.Focus()
                    Me.TBCouponAmount.SelectAll()
                    Exit Sub
                End If

                If vMemIssueCancel = 1 Then
                    MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                If vMemIssuePrintCount > 0 Then
                    MsgBox("เอกสารถูกพิมพ์คูปองไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                If vMemIssueOpen = 0 Then
                    Call vGetIssueDocNo()
                    vIsInsert = 1
                Else
                    vIsInsert = 0
                End If

                Call vCheckCouponAmount()

                If vMemCheckCouponAmount = 1 Then
                    MsgBox("มูลค่าคูปองเงินสด ไม่ตรงกับมูลค่าแต้มที่จะเบิก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBCouponAmount.Focus()
                    Me.TBCouponAmount.SelectAll()
                    Exit Sub
                End If

                Call vInsertCoupon()

                vARCode = Me.TBArCode.Text
                vDocno = Me.TBIssueNo.Text
                If vb6.Year(Me.DTPIssueDate.Value) >= 2500 Then
                    vDocDate = vb6.Day(Me.DTPIssueDate.Value) & "/" & vb6.Month(Me.DTPIssueDate.Value) & "/" & vb6.Year(Me.DTPIssueDate.Value) - 543
                Else
                    vDocDate = vb6.Day(Me.DTPIssueDate.Value) & "/" & vb6.Month(Me.DTPIssueDate.Value) & "/" & vb6.Year(Me.DTPIssueDate.Value)
                End If
                vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
                vScore = Me.TBIssueScore.Text

                vMemBeginTran = 1

                vQuery = "begin tran"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                vQuery = "exec dbo.USP_VP_WithdrawSetNew " & vIsInsert & ",'" & vDocno & "','" & vDocDate & "','" & vCampaignCode & "','" & vARCode & "'," & vScore & ",'" & vWindowsName & "' "
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                For i = 0 To Me.ListViewCoupon.Items.Count - 1
                    vLineNumber = i + 1
                    vCouponAmount = Me.ListViewCoupon.Items(i).SubItems(1).Text

                    vQuery = "exec dbo.USP_VP_WithdrawSubSetNew '" & vDocno & "','" & vDocDate & "'," & vLineNumber & "," & vCouponAmount & ""
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()
                Next

                vQuery = "commit tran"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                vMemCheckCouponAmount = 0

                MsgBox("เอกสารเลขที่ " & vDocno & " บันทึกเรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")
                Call IssueClearScreen()

                vMemBeginTran = 0

                vARCode = Me.TBArCode.Text
                vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)

                vQuery = "exec dbo.USP_VP_CalMemberPointFromInvoiceHistory2010 '" & vCampaignCode & "','" & vARCode & "'"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                Call CheckScoreHist(vARCode, vCampaignCode, vMemIsMember)
                Call vCheckIssueSore(vARCode, vCampaignCode)
                Call vCalcArScore()
                Call vCalcArIssueScoreHist()
            Else
                MsgBox("เอกสารเบิกคูปองไม่สามารถแก้ไขได้อีก หลังจากได้พิมพ์คูปองไปแล้ว ต้องเรียกคูปองเก่ากลับมาและยกเลิกเอกสารดังกล่าว แล้วค่อยทำเอกสารใบใหม่", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBIssueNo.Focus()
                Me.TBIssueNo.SelectAll()
            End If
            Else
                MsgBox("ลูกค้าท่านนี้ยังไม่ได้ สมัครสมาชิก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            End If

ErrDescription:
        If Err.Description <> "" Then
            If vMemBeginTran = 1 Then
                vQuery = "rollback tran"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()
            End If
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If
    End Sub

    Public Sub vInsertCoupon()
        Dim i As Integer
        Dim vListCoupon As ListViewItem
        Dim vCount1000 As Integer
        Dim vCount500 As Integer
        Dim vCount200 As Integer
        Dim vCount100 As Integer
        Dim vCountCoupon As Integer
        Dim n As Integer

        On Error Resume Next

        vCount1000 = Me.ND1000.Value
        vCount500 = Me.ND500.Value
        vCount200 = Me.ND200.Value
        vCount100 = Me.ND100.Value

        Me.ListViewCoupon.Items.Clear()

        vCountCoupon = vCount1000 + vCount500 + vCount200 + vCount100

        n = Me.ListViewCoupon.Items.Count + 1
        For i = 1 To vCount1000
            vListCoupon = Me.ListViewCoupon.Items.Add(n)
            vListCoupon.SubItems.Add(0).Text = Format(1000, "##,##0.00")
            n = n + 1
        Next

        For i = 1 To vCount500
            vListCoupon = Me.ListViewCoupon.Items.Add(n)
            vListCoupon.SubItems.Add(0).Text = Format(500, "##,##0.00")
            n = n + 1
        Next

        For i = 1 To vCount200
            vListCoupon = Me.ListViewCoupon.Items.Add(n)
            vListCoupon.SubItems.Add(0).Text = Format(200, "##,##0.00")
            n = n + 1
        Next

        For i = 1 To vCount100
            vListCoupon = Me.ListViewCoupon.Items.Add(n)
            vListCoupon.SubItems.Add(0).Text = Format(100, "##,##0.00")
            n = n + 1
        Next


    End Sub

    Private Sub ND500_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ND500.ValueChanged
        Call SumCounponAmount()
    End Sub

    Private Sub ND200_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ND200.ValueChanged
        Call SumCounponAmount()
    End Sub

    Private Sub ND100_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ND100.ValueChanged
        Call SumCounponAmount()
    End Sub

    Private Sub ListViewSearchIssueScore_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewSearchIssueScore.DoubleClick
        Dim vIndex As Integer
        Dim vDocNo As String
        Dim vARCode As String

        On Error Resume Next

        If Me.TBArName.Text = "" Then
            MsgBox("กรุณากรอก รหัสสมาชิกเพื่อค้นหาเลขที่เอกสารเบิกแต้ม", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBMemberID.Focus()
            Me.TBMemberID.SelectAll()
            Me.PNSearchIssueScore.Visible = False
            Exit Sub
        End If
        vARCode = Me.TBArCode.Text
        If Me.ListViewSearchIssueScore.Items.Count > 0 Then
            vIndex = Me.ListViewSearchIssueScore.SelectedItems(0).Index
            vDocNo = Me.ListViewSearchIssueScore.Items(vIndex).SubItems(1).Text
            Call vGetIssueScoreDetails(vDocNo)
        End If

    End Sub

    Private Sub ListViewSearchIssueScore_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearchIssueScore.KeyDown
        Dim vIndex As Integer
        Dim vDocNo As String
        Dim vARCode As String

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.TBArName.Text = "" Then
                MsgBox("", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If
            vARCode = Me.TBArCode.Text
            If Me.ListViewSearchIssueScore.Items.Count > 0 Then
                vIndex = Me.ListViewSearchIssueScore.SelectedItems(0).Index
                vDocNo = Me.ListViewSearchIssueScore.Items(vIndex).SubItems(1).Text
                Call vGetIssueScoreDetails(vDocNo)
            End If
        End If
    End Sub

    Private Sub ListViewSpecialScore_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewSearchIssueScore.SelectedIndexChanged

    End Sub

    Public Sub vGetIssueScoreDetails(ByVal vDocNo As String)
        Dim i As Integer
        Dim vScore As Double
        Dim vCount1000 As Integer
        Dim vCount500 As Integer
        Dim vCount200 As Integer
        Dim vCount100 As Integer
        Dim vCouponAmount As Double

        On Error Resume Next

        vQuery = "exec dbo.usp_vp_withdrawsearchsubnew '" & vDocNo & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "IssueScore")
        dt = ds.Tables("IssueScore")
        If dt.Rows.Count > 0 Then
            vMemIssueCancel = dt.Rows(0).Item("iscancel")
            vMemIssueConfirm = dt.Rows(0).Item("isconfirm")
            vMemIssuePrintCount = dt.Rows(0).Item("printcount")
            vMemIssueOpen = 1
            vScore = dt.Rows(0).Item("point")
            Me.TBIssueNo.Text = dt.Rows(0).Item("docno")
            Me.DTPIssueDate.Value = dt.Rows(0).Item("docdate")
            Me.TBIssueScore.Text = Format(vScore, "##,##0.00")
            Me.TBMemOldScore.Text = Format(vScore, "##,##0.00")

            If vMemIssueCancel = 1 Then
                Call IssueCancel()
                Me.BTNIssueCancel.Enabled = False
                Me.BTNIssuePrint.Enabled = False
            End If

            If vMemIssuePrintCount > 0 Then
                Call IssueConfirm()
                Me.BTNIssuePrint.Enabled = False
            Else
                Me.BTNIssuePrint.Enabled = True
            End If

            If vMemIssueCancel = 0 And vMemIssueConfirm = 0 Then
                Call IssueNew()
                Me.BTNIssuePrint.Enabled = True
            End If

            For i = 0 To dt.Rows.Count - 1
                vCouponAmount = dt.Rows(i).Item("amount")

                If vCouponAmount = 1000 Then
                    vCount1000 = vCount1000 + 1
                End If

                If vCouponAmount = 500 Then
                    vCount500 = vCount500 + 1
                End If

                If vCouponAmount = 200 Then
                    vCount200 = vCount200 + 1
                End If

                If vCouponAmount = 100 Then
                    vCount100 = vCount100 + 1
                End If
            Next

            Me.ND1000.Value = vCount1000
            Me.ND500.Value = vCount500
            Me.ND200.Value = vCount200
            Me.ND100.Value = vCount100

            Me.PNSearchIssueScore.Visible = False

        Else
            Call vGetSpecialDocNo()
            vMemIssueCancel = 0
            vMemIssueConfirm = 0
            vMemIssuePrintCount = 0
            vMemIssueOpen = 0
            Me.TBIssueNo.Text = ""
            Me.DTPIssueDate.Value = Now
            Me.TBIssueScore.Text = ""
            Me.TBMemOldScore.Text = ""
        End If
    End Sub

    Private Sub BTNPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNIssuePrint.Click
        Dim vDocNo As String
        Dim vPrinterName As String

        On Error Resume Next

        'MsgBox("ยังไม่เปิดให้ใช้งาน ", MsgBoxStyle.Critical, "")
        'Exit Sub

        Call ChekAuthorityAccess()
        If vDepartment <> "CH" And vDepartment <> "IT" Then
            MsgBox("คุณไม่มีสิทธิ์ในการ พิมพ์คูปองเงินสด กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
        If vMemIssueOpen = 1 Then
            If vMemIssuePrintCount = 0 Then

                If Me.TBArName.Text = "" Then
                    MsgBox("ไม่มีข้อมูลลูกค้า กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                If Me.TBIssueNo.Text = "" Then
                    MsgBox("ไม่มีเลขที่เอกสาร กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                vDocNo = Me.TBIssueNo.Text

                Call vCheckCouponAmount()
                Call vInsertCoupon()

                vQuery = "exec dbo.USP_VP_WithdrawUpdatePrint '" & vDocNo & "','" & vWindowsName & "'"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()


                'vQuery = "Exec dbo.USP_NP_SearchPrinterName '02','01'"
                'da = New SqlDataAdapter(vQuery, vConnection)
                'ds = New DataSet
                'da.Fill(ds, "Search")
                'dt1 = ds.Tables("Search")
                'If dt1.Rows.Count > 0 Then
                '    vPrinterName = dt1.Rows(0).Item("printername")
                'End If

                'vQuery = "Exec dbo.USP_VP_WithdrawSearchSubNew '" & vDocNo & "'"
                'da = New SqlDataAdapter(vQuery, vConnection)
                'ds = New DataSet
                'da.Fill(ds, "Search")
                'dt = ds.Tables("Search")

                'Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                'Dim frmObj As New FormReportReqItemComm
                'Dim FileName As New String("V:\Reports\NopadolCoupon\RP_NP_CashMemberCoupon.rpt")

                'rpt.Load(FileName)

                'Dim Params As New CrystalDecisions.Shared.ParameterField
                'Dim ParamCollection As New CrystalDecisions.Shared.ParameterFields
                'Dim ParamDisVal As New CrystalDecisions.Shared.ParameterDiscreteValue()
                'Params.ParameterFieldName = "@DocNo"
                'ParamDisVal.Value = vDocNo
                'Params.CurrentValues.Add(ParamDisVal)
                'ParamCollection.Add(Params)

                'rpt.Load(FileName)
                'rpt.SetDataSource(ds.Tables("Search"))
                'rpt.SetParameterValue("@DocNo", ParamDisVal)
                'rpt.PrintOptions.PrinterName = ""
                'rpt.PrintToPrinter(1, False, 0, 0)


                If frmCashCoupon Is Nothing Then
                    frmCashCoupon = New FormCashCoupon
                Else
                    If frmCashCoupon.IsDisposed Then
                        frmCashCoupon = New FormCashCoupon
                    End If
                End If

                frmCashCoupon.MdiParent = FormMain
                frmCashCoupon.Show()
                frmCashCoupon.BringToFront()


                MsgBox("พิมพ์คูปองเงินสดเรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")

                Call IssueClearScreen()

            Else
                MsgBox("เอกสารเบิกคูปองไม่สามารถแก้ไขได้อีก หลังจากได้พิมพ์คูปองไปแล้ว ต้องเรียกคูปองเก่ากลับมาและยกเลิกเอกสารดังกล่าว แล้วค่อยทำเอกสารใบใหม่", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBIssueNo.Focus()
                Me.TBIssueNo.SelectAll()
            End If
        Else
            MsgBox("เอกสารยังไม่ได้บันทึก ข้อมูล ไม่สามารถพิมพ์คูปองได้", MsgBoxStyle.Information, "Send Error Message")
        End If


        'vQuery = "Exec dbo.USP_COM_RequestSearch2 'RQ5301-0016'"
        'da = New SqlDataAdapter(vQuery, vConnection)
        'ds = New DataSet
        'da.Fill(ds, "Search")
        'dt = ds.Tables("Search")

        'Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        'Dim frmObj As New FormReportReqItemComm
        'Dim FileName As New String("W:\External\Reports\Commission\RP_COM_ReqComm.rpt")

        'rpt.Load(FileName)

        'Dim Params As New CrystalDecisions.Shared.ParameterField
        'Dim ParamCollection As New CrystalDecisions.Shared.ParameterFields
        'Dim ParamDisVal As New CrystalDecisions.Shared.ParameterDiscreteValue()
        'Params.ParameterFieldName = "@DocNo"
        'ParamDisVal.Value = "RQ5301-0016"
        'Params.CurrentValues.Add(ParamDisVal)
        'ParamCollection.Add(Params)

        'rpt.Load(FileName)
        'rpt.SetDataSource(ds.Tables("Search"))
        'rpt.SetParameterValue("@DocNo", ParamDisVal)
        'rpt.PrintOptions.PrinterName = "\\galaxy\HP Laser 2420 WS"
        'rpt.PrintToPrinter(1, False, 0, 0)

        ''With frmObj
        ''    .Crystal101.ReportSource = rpt
        ''    'ParameterInfo has been given all the information needed
        ''    .Crystal101.ParameterFieldInfo = ParamCollection
        ''    .ShowDialog()
        ''End With
    End Sub

    Private Sub BTNIssueClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNIssueClear.Click
        On Error Resume Next

        vMemIssueCancel = 0
        vMemIssueConfirm = 0
        vMemIssueOpen = 0
        vMemIssuePrintCount = 0
        vMemCheckCouponAmount = 0
        Me.TBIssueNo.Text = ""
        Me.DTPIssueDate.Value = Now
        Me.TBIssueScore.Text = ""
        Me.TBMemOldScore.Text = ""
        Me.TBCouponAmount.Text = ""
        Me.ListViewCoupon.Items.Clear()
        Me.PNSearchIssueScore.Visible = False
        Me.ND1000.Value = 0
        Me.ND500.Value = 0
        Me.ND200.Value = 0
        Me.ND100.Value = 0
        Call IssueNew()
        Call vGetIssueDocNo()
        Me.BTNIssuePrint.Enabled = False
    End Sub

    Public Sub IssueClearScreen()
        On Error Resume Next

        vMemIssueCancel = 0
        vMemIssueConfirm = 0
        vMemIssueOpen = 0
        vMemIssuePrintCount = 0
        vMemCheckCouponAmount = 0
        Me.TBIssueNo.Text = ""
        Me.DTPIssueDate.Value = Now
        Me.TBIssueScore.Text = ""
        Me.TBMemOldScore.Text = ""
        Me.TBCouponAmount.Text = ""
        Me.ListViewCoupon.Items.Clear()
        Me.PNSearchIssueScore.Visible = False
        Me.ND1000.Value = 0
        Me.ND500.Value = 0
        Me.ND200.Value = 0
        Me.ND100.Value = 0
        Call IssueNew()
        Call vGetIssueDocNo()
        Me.BTNIssuePrint.Enabled = False
    End Sub

    Private Sub BTNIssueCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNIssueCancel.Click
        'MsgBox("ยังไม่เปิดให้ใช้งาน ", MsgBoxStyle.Critical, "")
        'Exit Sub
        On Error Resume Next

        Call ChekAuthorityAccess()
        If vDepartment <> "CH" And vDepartment <> "AC" And vDepartment <> "IT" Then
            MsgBox("คุณไม่มีสิทธิ์ในการ ยกเลิกคูปองเงินสด กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
        Dim vDocno As String
        Dim vAnswer As Integer
        Dim vArCode As String
        Dim vCampaignCode As String

        If vMemIssueOpen = 1 Then
            If vMemIsMember <> 0 Then
                If Me.TBArName.Text = "" Then
                    MsgBox("ไม่มีข้อมูลลูกค้า กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                If Me.TBIssueNo.Text = "" Then
                    MsgBox("ไม่มีเลขที่เอกสาร กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                If Me.TBIssueScore.Text = "" Then
                    MsgBox("ไม่ได้กรอก มูลค่าแต้ม กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                If Me.TBCouponAmount.Text = "" Then
                    MsgBox("ยังไม่ได้เลือก จำนวนคูปองที่จะพิมพ์", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                If vMemIssueCancel = 1 Then
                    MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถยกเลิกเอกสารได้อีก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                vDocno = Me.TBIssueNo.Text

                vAnswer = MsgBox("คุณต้องการยกเลิกเอกสารเบิกแต้มเลขที่ " & vDocno & " ใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")

                If vAnswer = 6 Then
                    vQuery = "exec dbo.USP_VP_WithdrawCancelNew '" & vDocno & "','" & vWindowsName & "' "
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    vArCode = Me.TBArCode.Text
                    vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)

                    vQuery = "exec dbo.USP_VP_CalMemberPointFromInvoiceHistory2010 '" & vCampaignCode & "','" & vArCode & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    MsgBox("เอกสารเลขที่ " & vDocno & " ถูกยกเลิกเรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")
                    Call IssueClearScreen()
                    Call CheckScoreHist(vArCode, vCampaignCode, vMemIsMember)
                    Call vCheckIssueSore(vArCode, vCampaignCode)
                    Call vCalcArScore()
                    Call vCalcArIssueScoreHist()
                Else
                    Me.TBIssueNo.Focus()
                    Me.TBIssueNo.SelectAll()
                    Exit Sub
                End If

            Else
                MsgBox("ลูกค้าท่านนี้ยังไม่ได้ สมัครสมาชิก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            End If
        Else
            MsgBox("เอกสารยังไม่ได้บันทึกไม่สามารถยกเลิกเอกสารได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub LLBCampaign_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LLBCampaign.LinkClicked
        On Error Resume Next

        Call ChekAuthorityAccess()
        If vDepartment = "MK" Or vDepartment = "IT" Then
            Me.PNCheckScore.Visible = False
            Me.PNCampaign.Visible = True
            Me.PNSpecialApprove.Visible = False
            Me.PNSpecialCancel.Visible = False

            Call vGetCampaignNo()
            Me.DTPBeginDate.Value = Now
            Me.DTPEndDate.Value = Now
            Me.TBCampaignName.Focus()
            Me.TBCampaignName.SelectAll()
        Else
            MsgBox("คุณไม่มีสิทธิ์ในการใช้งานในส่วนนี้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Public Sub vGetCampaignNo()
        Dim vDocNo As String
        Dim vNow As String

        On Error Resume Next

        vNow = vb6.Day(Now) & "/" & vb6.Month(Now) & "/" & vb6.Year(Now)
        vQuery = "set dateformat dmy"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vQuery = "select dbo.FT_VP_NewCampaign ('" & vNow & "') as docno"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then
            vDocNo = dt.Rows(0).Item("Docno")
            Me.TBCampaignCode.Text = vDocNo
        Else
            Me.TBCampaignCode.Text = ""
        End If

        Me.TBCampaignName.Text = ""
        Me.TBCampaignEng.Text = ""
        Me.DTPBeginDate.Value = Now
        Me.DTPEndDate.Value = Now
        vMemCampaignOpen = 0
    End Sub

    Private Sub LLBCheckScore_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LLBCheckScore.LinkClicked
        On Error Resume Next

        Me.PNCheckScore.Visible = True
        Me.PNCampaign.Visible = False
        Me.PNSpecialApprove.Visible = False
        Me.PNSpecialCancel.Visible = False

        Me.TBMemberID.Focus()
        Me.TBMemberID.SelectAll()
    End Sub

    Private Sub BTNCampaignSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCampaignSave.Click
        Dim vIsInsert As Integer
        Dim vCampaignCode As String
        Dim vCampaignName As String
        Dim vCampaignEng As String
        Dim vCampaignBegin As String
        Dim vCampaignEnd As String

        On Error Resume Next

        If Me.TBCampaignCode.Text <> "" And Me.TBCampaignName.Text <> "" Then
            If vMemCampaignOpen = 0 Then
                vIsInsert = 1
            Else
                vIsInsert = 0
            End If

            vCampaignCode = Me.TBCampaignCode.Text
            vCampaignName = Me.TBCampaignName.Text
            vCampaignEng = Me.TBCampaignEng.Text
            vCampaignBegin = vb6.Day(Me.DTPBeginDate.Value) & "/" & vb6.Month(Me.DTPBeginDate.Value) & "/" & vb6.Year(Me.DTPBeginDate.Value)
            vCampaignEnd = vb6.Day(Me.DTPEndDate.Value) & "/" & vb6.Month(Me.DTPEndDate.Value) & "/" & vb6.Year(Me.DTPEndDate.Value)

            vQuery = "exec dbo.USP_VP_CampaignAssign " & vIsInsert & ",'" & vCampaignCode & "','" & vCampaignName & "','" & vCampaignEng & "','" & vCampaignBegin & "','" & vCampaignEnd & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            MsgBox("บันทึกข้อมูลแคมเปญ เรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")

            Call vCampaignClear()
        Else
            MsgBox("กรอกข้อมูลแคมเปญไม่ครบไม่สามารถบันทึกข้อมูลได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "send Error Message")
            Me.TBCampaignName.Focus()
            Me.TBCampaignName.SelectAll()
        End If
    End Sub

    Private Sub BTNCampaignClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCampaignClear.Click
        On Error Resume Next

        Me.TBCampaignCode.Enabled = True
        Me.TBCampaignCode.Text = ""
        Me.TBCampaignName.Text = ""
        Me.TBCampaignEng.Text = ""
        Me.DTPBeginDate.Value = Now
        Me.DTPEndDate.Value = Now
        vMemCampaignOpen = 0
        Call vGetCampaignNo()
        Me.TBCampaignName.Focus()
        Me.TBCampaignName.SelectAll()
    End Sub

    Public Sub vCampaignClear()
        On Error Resume Next

        Me.TBCampaignCode.Enabled = True
        Me.TBCampaignCode.Text = ""
        Me.TBCampaignName.Text = ""
        Me.TBCampaignEng.Text = ""
        Me.DTPBeginDate.Value = Now
        Me.DTPEndDate.Value = Now
        vMemCampaignOpen = 0
        Call vGetCampaignNo()
        Me.TBCampaignName.Focus()
        Me.TBCampaignName.SelectAll()
    End Sub

    Private Sub BTNCampaignSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCampaignSearch.Click
        Dim vSearch As String

        On Error Resume Next

        vSearch = ""
        Call vSearchCampaign(vSearch)
        If Me.ListViewSearchCampaign.Items.Count > 0 Then
            Me.PNSearchCampaign.Visible = True

            Me.ListViewSearchCampaign.Focus()
            Me.ListViewSearchCampaign.Items(0).Selected = True
            Me.ListViewSearchCampaign.Items(0).Focused = True
        Else
            MsgBox("ไม่มีทะเบียนแคมเปญ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.PNSearchCampaign.Visible = False
            Me.TBCampaignCode.Focus()
            Me.TBCampaignCode.SelectAll()
        End If

    End Sub

    Public Sub vSearchCampaign(ByVal vSearch As String)
        Dim i As Integer
        Dim vListItem As ListViewItem
        Dim n As Integer

        On Error Resume Next

        Me.ListViewSearchCampaign.Items.Clear()
        vQuery = "exec dbo.USP_VP_CampaignSearch '" & vSearch & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vListItem = Me.ListViewSearchCampaign.Items.Add(n)
                vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("code")
                vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("nameth")
                vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("nameen")
                vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("startdate")
                vListItem.SubItems.Add(4).Text = dt.Rows(i).Item("stopdate")
            Next
        End If
    End Sub

    Private Sub TBSearchCampaign_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchCampaign.KeyDown
        Dim vSearch As String

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            vSearch = Me.TBSearchCampaign.Text
            Call vSearchCampaign(vSearch)
            If Me.ListViewSearchCampaign.Items.Count > 0 Then
                Me.ListViewSearchCampaign.Focus()
                Me.ListViewSearchCampaign.Items(0).Selected = True
                Me.ListViewSearchCampaign.Items(0).Focused = True
            Else
                MsgBox("ไม่มีทะเบียนแคมเปญที่ค้นหา กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBSearchCampaign.Focus()
                Me.TBSearchCampaign.SelectAll()
            End If
        End If
    End Sub

    Private Sub TBSearchCampaign_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchCampaign.TextChanged

    End Sub

    Private Sub ListViewSearchCampaign_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewSearchCampaign.DoubleClick
        Dim vIndex As Integer

        On Error Resume Next

        If Me.ListViewSearchCampaign.Items.Count > 0 Then
            vIndex = Me.ListViewSearchCampaign.SelectedItems(0).Index

            Me.TBCampaignCode.Text = Me.ListViewSearchCampaign.Items(vIndex).SubItems(1).Text
            Me.TBCampaignName.Text = Me.ListViewSearchCampaign.Items(vIndex).SubItems(2).Text
            Me.TBCampaignEng.Text = Me.ListViewSearchCampaign.Items(vIndex).SubItems(3).Text
            Me.DTPBeginDate.Value = Me.ListViewSearchCampaign.Items(vIndex).SubItems(4).Text
            Me.DTPEndDate.Value = Me.ListViewSearchCampaign.Items(vIndex).SubItems(5).Text

            vMemCampaignOpen = 1
            Me.TBCampaignCode.Enabled = False
            Me.PNSearchCampaign.Visible = False
            Me.TBCampaignName.Focus()
            Me.TBCampaignName.SelectAll()
        End If
    End Sub

    Private Sub ListViewSearchCampaign_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearchCampaign.KeyDown
        Dim vIndex As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewSearchCampaign.Items.Count > 0 Then
                vIndex = Me.ListViewSearchCampaign.SelectedItems(0).Index

                Me.TBCampaignCode.Text = Me.ListViewSearchCampaign.Items(vIndex).SubItems(1).Text
                Me.TBCampaignName.Text = Me.ListViewSearchCampaign.Items(vIndex).SubItems(2).Text
                Me.TBCampaignEng.Text = Me.ListViewSearchCampaign.Items(vIndex).SubItems(3).Text
                Me.DTPBeginDate.Value = Me.ListViewSearchCampaign.Items(vIndex).SubItems(4).Text
                Me.DTPEndDate.Value = Me.ListViewSearchCampaign.Items(vIndex).SubItems(5).Text

                vMemCampaignOpen = 1
                Me.TBCampaignCode.Enabled = False
                Me.PNSearchCampaign.Visible = False
                Me.TBCampaignName.Focus()
                Me.TBCampaignName.SelectAll()
            End If
        End If
    End Sub

    Private Sub ListViewSearchCampaign_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewSearchCampaign.SelectedIndexChanged

    End Sub

    Private Sub BTNSearchCampaign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchCampaign.Click
        Dim vSearch As String

        On Error Resume Next

        vSearch = Me.TBSearchCampaign.Text
        Call vSearchCampaign(vSearch)
        If Me.ListViewSearchCampaign.Items.Count > 0 Then
            Me.ListViewSearchCampaign.Focus()
            Me.ListViewSearchCampaign.Items(0).Selected = True
            Me.ListViewSearchCampaign.Items(0).Focused = True
        Else
            MsgBox("ไม่มีทะเบียนแคมเปญที่ค้นหา กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBSearchCampaign.Focus()
            Me.TBSearchCampaign.SelectAll()
        End If
    End Sub

    Private Sub BTNSelectCampaign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectCampaign.Click
        Dim vIndex As Integer

        On Error Resume Next

        If Me.ListViewSearchCampaign.Items.Count > 0 Then
            vIndex = Me.ListViewSearchCampaign.SelectedItems(0).Index

            Me.TBCampaignCode.Text = Me.ListViewSearchCampaign.Items(vIndex).SubItems(1).Text
            Me.TBCampaignName.Text = Me.ListViewSearchCampaign.Items(vIndex).SubItems(2).Text
            Me.TBCampaignEng.Text = Me.ListViewSearchCampaign.Items(vIndex).SubItems(3).Text
            Me.DTPBeginDate.Value = Me.ListViewSearchCampaign.Items(vIndex).SubItems(4).Text
            Me.DTPEndDate.Value = Me.ListViewSearchCampaign.Items(vIndex).SubItems(5).Text

            vMemCampaignOpen = 1
            Me.TBCampaignCode.Enabled = False
            Me.PNSearchCampaign.Visible = False
            Me.TBCampaignName.Focus()
            Me.TBCampaignName.SelectAll()
        End If
    End Sub

    Private Sub BTNCloseSelectCampaign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSelectCampaign.Click
        Me.PNSearchCampaign.Visible = False
        Me.TBCampaignName.Focus()
        Me.TBCampaignName.SelectAll()
    End Sub

    Private Sub LLBSpecialApprove_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LLBSpecialApprove.LinkClicked
        On Error Resume Next

        Call ChekAuthorityAccess()
        If vDepartment = "MC" Then
            Me.PNCheckScore.Visible = False
            Me.PNCampaign.Visible = False
            Me.PNSpecialApprove.Visible = True
            Me.PNSpecialCancel.Visible = False

            Call vGetCampaignApprove()
            Me.ListViewCampaignApprove.Items.Clear()
            Me.CMBCampaignApprove.Focus()
        Else
            MsgBox("คุณไม่มีสิทธิ์ในการใช้งานในส่วนนี้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Public Sub vGetCampaignApprove()
        Dim i As Integer

        On Error Resume Next

        Me.CMBCampaignApprove.Items.Clear()
        vQuery = "exec dbo.USP_VP_CampaignList"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Campaign")
        dt = ds.Tables("Campaign")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                Me.CMBCampaignApprove.Items.Add(dt.Rows(i).Item("code") & "/" & dt.Rows(i).Item("nameth"))
            Next
        End If
    End Sub

    Public Sub vSearchSpecialApprove(ByVal vCampaignCode As String)
        Dim i As Integer
        Dim vListItem As ListViewItem
        Dim n As Integer
        Dim vScore As Double

        On Error Resume Next

        Me.ListViewCampaignApprove.Items.Clear()
        vQuery = "exec dbo.USP_VP_PointSpecialHolding'" & vCampaignCode & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vListItem = Me.ListViewCampaignApprove.Items.Add(n)
                vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("arname")
                vScore = dt.Rows(i).Item("point")
                vListItem.SubItems.Add(2).Text = Format(vScore, "##,##0.00")
                vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("docdate")
                vListItem.SubItems.Add(4).Text = dt.Rows(i).Item("memberid")
                vListItem.SubItems.Add(5).Text = dt.Rows(i).Item("arcode")
                vListItem.SubItems.Add(6).Text = dt.Rows(i).Item("reason")
            Next
        End If
    End Sub

    Public Sub vSearchSpecialCancel(ByVal vCampaignCode As String)
        Dim i As Integer
        Dim vListItem As ListViewItem
        Dim n As Integer
        Dim vScore As Double

        On Error Resume Next

        Me.ListViewCampaignCancel.Items.Clear()
        vQuery = "exec dbo.USP_VP_PointSpecialHolding'" & vCampaignCode & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vListItem = Me.ListViewCampaignCancel.Items.Add(n)
                vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("arname")
                vScore = dt.Rows(i).Item("point")
                vListItem.SubItems.Add(2).Text = Format(vScore, "##,##0.00")
                vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("docdate")
                vListItem.SubItems.Add(4).Text = dt.Rows(i).Item("memberid")
                vListItem.SubItems.Add(5).Text = dt.Rows(i).Item("arcode")
                vListItem.SubItems.Add(6).Text = dt.Rows(i).Item("reason")
            Next
        End If
    End Sub

    Private Sub LLBSpecialCancel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LLBSpecialCancel.LinkClicked
        On Error Resume Next

        Call ChekAuthorityAccess()
        If vDepartment = "MC" Then
            Me.PNCheckScore.Visible = False
            Me.PNCampaign.Visible = False
            Me.PNSpecialApprove.Visible = False
            Me.PNSpecialCancel.Visible = True

            Call vGetCampaignCancel()
            Me.ListViewCampaignCancel.Items.Clear()
            Me.CMBCampaignCancel.Focus()
        Else
            MsgBox("คุณไม่มีสิทธิ์ในการใช้งานในส่วนนี้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Public Sub vGetCampaignCancel()
        Dim i As Integer

        On Error Resume Next

        Me.CMBCampaignCancel.Items.Clear()
        vQuery = "exec dbo.USP_VP_CampaignList"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Campaign")
        dt = ds.Tables("Campaign")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                Me.CMBCampaignCancel.Items.Add(dt.Rows(i).Item("code") & "/" & dt.Rows(i).Item("nameth"))
            Next
        End If
    End Sub


    Private Sub CMBCampaignApprove_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBCampaignApprove.SelectedIndexChanged
        Dim vCampaignCode As String

        On Error Resume Next

        If Me.CMBCampaignApprove.Items.Count > 0 Then
            vCampaignCode = vb6.Left(Me.CMBCampaignApprove.SelectedItem, vb6.InStr(Me.CMBCampaignApprove.SelectedItem, "/") - 1)
            Call vSearchSpecialApprove(vCampaignCode)

            If Me.ListViewCampaignApprove.Items.Count > 0 Then
                Me.ListViewCampaignApprove.Focus()
                Me.ListViewCampaignApprove.Items(0).Selected = True
                Me.ListViewCampaignApprove.Items(0).Focused = True
            Else
                Me.CMBCampaignApprove.Focus()
            End If
        End If
    End Sub

    Private Sub CBApproveAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBApproveAll.CheckedChanged
        Dim i As Integer

        On Error Resume Next

        If Me.ListViewCampaignApprove.Items.Count > 0 Then
            If Me.CBApproveAll.Checked = True Then
                For i = 0 To Me.ListViewCampaignApprove.Items.Count - 1
                    Me.ListViewCampaignApprove.Items(i).Checked = True
                Next
            End If

            If Me.CBApproveAll.Checked = False Then
                For i = 0 To Me.ListViewCampaignApprove.Items.Count - 1
                    Me.ListViewCampaignApprove.Items(i).Checked = False
                Next
            End If
        End If
    End Sub

    Private Sub BTNApproveSpecial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNApproveSpecial.Click
        Dim i As Integer
        Dim n As Integer
        Dim vDocNo As String
        Dim vCountSelect As Integer

        On Error Resume Next

        If Me.ListViewCampaignApprove.Items.Count > 0 Then
            For i = 0 To Me.ListViewCampaignApprove.Items.Count - 1
                If Me.ListViewCampaignApprove.Items(i).Checked = True Then
                    vCountSelect = vCountSelect + 1
                End If
            Next

            If vCountSelect > 0 Then
                For n = 0 To Me.ListViewCampaignApprove.Items.Count - 1
                    If Me.ListViewCampaignApprove.Items(n).Checked = True Then
                        vDocNo = Me.ListViewCampaignApprove.Items(n).SubItems(1).Text

                        vQuery = "exec dbo.USP_VP_PointSpecialConfirm '" & vDocNo & "'"
                        cmd = New SqlCommand(vQuery, vConnection)
                        cmd.ExecuteNonQuery()

                    End If
                Next

                MsgBox("อนุมัติ รายการเอกสาร เรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Information Message")
                Me.ListViewCampaignApprove.Items.Clear()
                Me.CBApproveAll.Checked = False
                Me.CMBCampaignApprove.Focus()
            Else
                MsgBox("ยังไม่ได้เลือก รายการเอกสารที่ต้องการอนุมัติ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                If Me.ListViewCampaignApprove.Items.Count > 0 Then
                    Me.ListViewCampaignApprove.Focus()
                    Me.ListViewCampaignApprove.Items(0).Selected = True
                    Me.ListViewCampaignApprove.Items(0).Focused = True
                Else
                    Me.CMBCampaignApprove.Focus()
                End If
            End If
        End If

    End Sub

    Private Sub CBCancelAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBCancelAll.CheckedChanged
        Dim i As Integer

        On Error Resume Next

        If Me.ListViewCampaignCancel.Items.Count > 0 Then
            If Me.CBCancelAll.Checked = True Then
                For i = 0 To Me.ListViewCampaignCancel.Items.Count - 1
                    Me.ListViewCampaignCancel.Items(i).Checked = True
                Next
            End If

            If Me.CBCancelAll.Checked = False Then
                For i = 0 To Me.ListViewCampaignCancel.Items.Count - 1
                    Me.ListViewCampaignCancel.Items(i).Checked = False
                Next
            End If
        End If
    End Sub

    Private Sub CMBCampaignCancel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBCampaignCancel.SelectedIndexChanged
        Dim vCampaignCode As String

        On Error Resume Next

        If Me.CMBCampaignApprove.Items.Count > 0 Then
            vCampaignCode = vb6.Left(Me.CMBCampaignCancel.SelectedItem, vb6.InStr(Me.CMBCampaignCancel.SelectedItem, "/") - 1)
            Call vSearchSpecialCancel(vCampaignCode)

            If Me.ListViewCampaignCancel.Items.Count > 0 Then
                Me.ListViewCampaignCancel.Focus()
                Me.ListViewCampaignCancel.Items(0).Selected = True
                Me.ListViewCampaignCancel.Items(0).Focused = True
            Else
                Me.CMBCampaignCancel.Focus()
            End If
        End If
    End Sub

    Private Sub BTNCampaignCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCampaignCancel.Click
        Dim i As Integer
        Dim n As Integer
        Dim vDocNo As String
        Dim vCountSelect As Integer

        On Error Resume Next

        If Me.ListViewCampaignCancel.Items.Count > 0 Then
            For i = 0 To Me.ListViewCampaignCancel.Items.Count - 1
                If Me.ListViewCampaignCancel.Items(i).Checked = True Then
                    vCountSelect = vCountSelect + 1
                End If
            Next

            If vCountSelect > 0 Then
                For n = 0 To Me.ListViewCampaignCancel.Items.Count - 1
                    If Me.ListViewCampaignCancel.Items(n).Checked = True Then
                        vDocNo = Me.ListViewCampaignCancel.Items(n).SubItems(1).Text

                        vQuery = "exec dbo.USP_VP_PointSpecialCancel '" & vDocNo & "'"
                        cmd = New SqlCommand(vQuery, vConnection)
                        cmd.ExecuteNonQuery()

                    End If
                Next

                MsgBox("ยกเลิก รายการเอกสาร เรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Information Message")
                Me.ListViewCampaignCancel.Items.Clear()
                Me.CBCancelAll.Checked = False
                Me.CMBCampaignCancel.Focus()
            Else
                MsgBox("ยังไม่ได้เลือก รายการเอกสารที่ต้องการยกเลิก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                If Me.ListViewCampaignCancel.Items.Count > 0 Then
                    Me.ListViewCampaignCancel.Focus()
                    Me.ListViewCampaignCancel.Items(0).Selected = True
                    Me.ListViewCampaignCancel.Items(0).Focused = True
                Else
                    Me.CMBCampaignCancel.Focus()
                End If
            End If
        End If

    End Sub

    Private Sub TBSpecialScore_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBSpecialScore.KeyPress
        On Error GoTo ErrDescription

        Select Case Asc(e.KeyChar)
            Case 48 To 58, 8, 44, 45, 46, 37
            Case Else
                e.Handled = True
        End Select

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TBSpecialScore_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSpecialScore.TextChanged

    End Sub

    Private Sub TBIssueNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBIssueNo.TextChanged

    End Sub

    Private Sub TBArName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBArName.TextChanged

    End Sub

    Public Sub ClearFilterScoreHist()
        Dim vARCode As String
        Dim vCampaignCode As String
        Dim vCampaignStart As Date
        Dim vCampaignStop As Date
        Dim vSelectStart As Date
        Dim vSelectStop As Date
        Dim vInvoiceNo As String
        Dim vDate1 As String
        Dim vDate2 As String

        On Error Resume Next


        If Me.TBInvoiceNo.Text <> "" Then
            Me.TBInvoiceNo.Text = ""
            vCampaignStart = Me.TBStartDate.Text
            vCampaignStop = Me.TBExpireDate.Text
            vSelectStart = Me.DTPDate1.Text
            vSelectStop = Me.DTPDate2.Text

            If Me.CMBCampaign.Text = "" Then
                Exit Sub
            End If

            If Me.TBArCode.Text = "" Then
                Exit Sub
            End If

            If vCampaignStart = vSelectStart And vCampaignStart = vSelectStart Then
                vARCode = Me.TBArCode.Text
                vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)

                Call CheckScoreHist(vARCode, vCampaignCode, vMemIsMember)
            Else
                If Me.CMBCampaign.Text = "" Then
                    Exit Sub
                End If

                If Me.TBArCode.Text = "" Then
                    Exit Sub
                End If
                vARCode = Me.TBArCode.Text
                vCampaignCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
                vInvoiceNo = ""
                vDate1 = vb6.Day(Me.DTPDate1.Text) & "/" & vb6.Month(Me.DTPDate1.Text) & "/" & vb6.Year(Me.DTPDate1.Text)
                vDate2 = vb6.Day(Me.DTPDate2.Text) & "/" & vb6.Month(Me.DTPDate2.Text) & "/" & vb6.Year(Me.DTPDate2.Text)
                Call CheckScoreHistFilter1(vARCode, vCampaignCode, vMemIsMember, vDate1, vDate2)
                Call vCalcArScore()
                Call vCalcArIssueScoreHist()
            End If
            Call vCalcArScore()
            Call vCalcArIssueScoreHist()

            If Me.ListViewAccumulateScore.Items.Count > 0 Then
                Me.ListViewAccumulateScore.Focus()
                Me.ListViewAccumulateScore.Items(0).Selected = True
                Me.ListViewAccumulateScore.Items(0).Focused = True
            End If

        End If
    End Sub

    Private Sub BTNFilter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNFilter.KeyDown, DTPDate1.KeyDown, DTPDate2.KeyDown, TBInvoiceNo.KeyDown, ListViewAccumulateScore.KeyDown, TBArScore.KeyDown
        Dim vSearch As String
        vSearch = Me.TBSearchMember.Text

        On Error Resume Next

        If e.Shift = True And e.KeyCode = Keys.F1 Then
            Call vSearchMember(vSearch)
            If Me.ListViewSearchMember.Items.Count > 0 Then
                Me.ListViewSearchMember.Focus()
                Me.ListViewSearchMember.Items(0).Selected = True
                Me.ListViewSearchMember.Items(0).Focused = True
            End If
        End If

        If e.Shift = False And e.KeyCode = Keys.F1 Then
            Me.PNAccumulateScore.Visible = True
            Me.PNIssueScoreHistory.Visible = False
            Me.PNAddSpecialScore.Visible = False
            Me.PNIssueScore.Visible = False
            If Me.ListViewAccumulateScore.Items.Count > 0 Then
                Me.ListViewAccumulateScore.Focus()
                Me.ListViewAccumulateScore.Items(0).Selected = True
                Me.ListViewAccumulateScore.Items(0).Focused = True
            Else
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.F2 Then
            If vMemIsMember <> 0 Then
                Me.PNAccumulateScore.Visible = False
                Me.PNIssueScoreHistory.Visible = True

                Me.PNAddSpecialScore.Visible = False
                Me.PNIssueScore.Visible = False

                If Me.ListViewIssueScore.Items.Count > 0 Then
                    Me.ListViewIssueScore.Focus()
                    Me.ListViewIssueScore.Items(0).Selected = True
                    Me.ListViewIssueScore.Items(0).Focused = True
                Else
                    Me.TBMemberID.Focus()
                    Me.TBMemberID.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.F3 Then
            Call ChekAuthorityAccess()
            If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
                If vMemIsMember <> 0 Then
                    Me.DTPSpecialDate.Value = Now
                    Call vGetCampaignSpecial("")
                    Call vGetSpecialDocNo()
                    Call vGetNewSpecial()
                    Me.TBSpecialScore.Focus()
                    Me.TBSpecialScore.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.F4 Then
            If Me.TBArName.Text <> "" Then
                If vMemIsMember <> 0 Then
                    Call CalcMemberRemain()
                Else
                    MsgBox("ลูกค้าที่เป็นสมาชิกเท่านั้นถึงจะคำนวณแต้มได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
            End If
        End If

        If e.KeyCode = Keys.F5 Then
            Call ChekAuthorityAccess()
            If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
                If vMemIsMember <> 0 Then
                    Me.PNAccumulateScore.Visible = False
                    Me.PNIssueScoreHistory.Visible = False
                    Me.PNAddSpecialScore.Visible = False
                    Me.PNIssueScore.Visible = True

                    Call IssueNew()
                    Me.DTPIssueDate.Value = Now
                    Call vGetIssueDocNo()

                    '=====================================================
                    vMemSpecialCancel = 0
                    vMemSpecialConfirm = 0
                    vMemSpecialOpen = 0

                    Me.TBSpecialNo.Text = ""
                    Me.DTPSpecialDate.Value = Now
                    If Me.CMBSpecialCampaign.Items.Count > 0 Then
                        Me.CMBSpecialCampaign.SelectedIndex = 0
                    Else
                        Me.CMBSpecialCampaign.Text = ""
                    End If
                    Me.TBSpecialReason.Text = ""
                    Me.TBSpecialScore.Text = ""
                    '=====================================================

                    Me.TBIssueScore.Focus()
                    Me.TBIssueScore.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearFilterScoreHist()
        End If
    End Sub

    Private Sub BTNSpecialApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ListViewIssueScore_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewIssueScore.KeyDown
        Dim vSearch As String

        On Error Resume Next

        vSearch = Me.TBSearchMember.Text

        If e.Shift = True And e.KeyCode = Keys.F1 Then
            Call vSearchMember(vSearch)
            If Me.ListViewSearchMember.Items.Count > 0 Then
                Me.ListViewSearchMember.Focus()
                Me.ListViewSearchMember.Items(0).Selected = True
                Me.ListViewSearchMember.Items(0).Focused = True
            End If
        End If

        If e.Shift = False And e.KeyCode = Keys.F1 Then
            Me.PNAccumulateScore.Visible = True
            Me.PNIssueScoreHistory.Visible = False
            Me.PNAddSpecialScore.Visible = False
            Me.PNIssueScore.Visible = False
            If Me.ListViewAccumulateScore.Items.Count > 0 Then
                Me.ListViewAccumulateScore.Focus()
                Me.ListViewAccumulateScore.Items(0).Selected = True
                Me.ListViewAccumulateScore.Items(0).Focused = True
            Else
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.F2 Then
            If vMemIsMember <> 0 Then
                Me.PNAccumulateScore.Visible = False
                Me.PNIssueScoreHistory.Visible = True

                Me.PNAddSpecialScore.Visible = False
                Me.PNIssueScore.Visible = False

                If Me.ListViewIssueScore.Items.Count > 0 Then
                    Me.ListViewIssueScore.Focus()
                    Me.ListViewIssueScore.Items(0).Selected = True
                    Me.ListViewIssueScore.Items(0).Focused = True
                Else
                    Me.TBMemberID.Focus()
                    Me.TBMemberID.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.F3 Then
            Call ChekAuthorityAccess()
            If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
                If vMemIsMember <> 0 Then
                    Me.DTPSpecialDate.Value = Now
                    Call vGetCampaignSpecial("")
                    Call vGetSpecialDocNo()
                    Call vGetNewSpecial()
                    Me.TBSpecialScore.Focus()
                    Me.TBSpecialScore.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.F4 Then
            If Me.TBArName.Text <> "" Then
                If vMemIsMember <> 0 Then
                    Call CalcMemberRemain()
                Else
                    MsgBox("ลูกค้าที่เป็นสมาชิกเท่านั้นถึงจะคำนวณแต้มได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
            End If
        End If

        If e.KeyCode = Keys.F5 Then
            Call ChekAuthorityAccess()
            If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
                If vMemIsMember <> 0 Then
                    Me.PNAccumulateScore.Visible = False
                    Me.PNIssueScoreHistory.Visible = False
                    Me.PNAddSpecialScore.Visible = False
                    Me.PNIssueScore.Visible = True

                    Call IssueNew()
                    Me.DTPIssueDate.Value = Now
                    Call vGetIssueDocNo()

                    '=====================================================
                    vMemSpecialCancel = 0
                    vMemSpecialConfirm = 0
                    vMemSpecialOpen = 0

                    Me.TBSpecialNo.Text = ""
                    Me.DTPSpecialDate.Value = Now
                    If Me.CMBSpecialCampaign.Items.Count > 0 Then
                        Me.CMBSpecialCampaign.SelectedIndex = 0
                    Else
                        Me.CMBSpecialCampaign.Text = ""
                    End If
                    Me.TBSpecialReason.Text = ""
                    Me.TBSpecialScore.Text = ""
                    '=====================================================

                    Me.TBIssueScore.Focus()
                    Me.TBIssueScore.SelectAll()
                End If
            End If
        End If

    End Sub

    Private Sub ListViewIssueScore_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewIssueScore.SelectedIndexChanged

    End Sub

    Private Sub BTNIssueSave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNIssueSave.KeyDown, TBIssueNo.KeyDown, DTPIssueDate.KeyDown, TBIssueScore.KeyDown, TBMemOldScore.KeyDown, ListViewCoupon.KeyDown, TBCouponAmount.KeyDown, ND1000.KeyDown, ND500.KeyDown, ND200.KeyDown, ND100.KeyDown, BTNIssueClear.KeyDown, BTNIssueSearch.KeyDown, BTNIssuePrint.KeyDown, BTNIssueCancel.KeyDown
        Dim vSearch As String

        On Error Resume Next

        vSearch = Me.TBSearchMember.Text

        If e.Shift = True And e.KeyCode = Keys.F1 Then
            Call vSearchMember(vSearch)
            If Me.ListViewSearchMember.Items.Count > 0 Then
                Me.ListViewSearchMember.Focus()
                Me.ListViewSearchMember.Items(0).Selected = True
                Me.ListViewSearchMember.Items(0).Focused = True
            End If
        End If

        If e.Shift = False And e.KeyCode = Keys.F1 Then
            Me.PNAccumulateScore.Visible = True
            Me.PNIssueScoreHistory.Visible = False
            Me.PNAddSpecialScore.Visible = False
            Me.PNIssueScore.Visible = False
            If Me.ListViewAccumulateScore.Items.Count > 0 Then
                Me.ListViewAccumulateScore.Focus()
                Me.ListViewAccumulateScore.Items(0).Selected = True
                Me.ListViewAccumulateScore.Items(0).Focused = True
            Else
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.F2 Then
            If vMemIsMember <> 0 Then
                Me.PNAccumulateScore.Visible = False
                Me.PNIssueScoreHistory.Visible = True

                Me.PNAddSpecialScore.Visible = False
                Me.PNIssueScore.Visible = False

                If Me.ListViewIssueScore.Items.Count > 0 Then
                    Me.ListViewIssueScore.Focus()
                    Me.ListViewIssueScore.Items(0).Selected = True
                    Me.ListViewIssueScore.Items(0).Focused = True
                Else
                    Me.TBMemberID.Focus()
                    Me.TBMemberID.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.F3 Then
            Call ChekAuthorityAccess()
            If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
                If vMemIsMember <> 0 Then
                    Me.DTPSpecialDate.Value = Now
                    Call vGetCampaignSpecial("")
                    Call vGetSpecialDocNo()
                    Call vGetNewSpecial()
                    Me.TBSpecialScore.Focus()
                    Me.TBSpecialScore.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.F4 Then
            If Me.TBArName.Text <> "" Then
                If vMemIsMember <> 0 Then
                    Call CalcMemberRemain()
                Else
                    MsgBox("ลูกค้าที่เป็นสมาชิกเท่านั้นถึงจะคำนวณแต้มได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
            End If
        End If

        If e.KeyCode = Keys.F5 Then
            Call ChekAuthorityAccess()
            If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
                If vMemIsMember <> 0 Then
                    Me.PNAccumulateScore.Visible = False
                    Me.PNIssueScoreHistory.Visible = False
                    Me.PNAddSpecialScore.Visible = False
                    Me.PNIssueScore.Visible = True

                    Call IssueNew()
                    Me.DTPIssueDate.Value = Now
                    Call vGetIssueDocNo()

                    '=====================================================
                    vMemSpecialCancel = 0
                    vMemSpecialConfirm = 0
                    vMemSpecialOpen = 0

                    Me.TBSpecialNo.Text = ""
                    Me.DTPSpecialDate.Value = Now
                    If Me.CMBSpecialCampaign.Items.Count > 0 Then
                        Me.CMBSpecialCampaign.SelectedIndex = 0
                    Else
                        Me.CMBSpecialCampaign.Text = ""
                    End If
                    Me.TBSpecialReason.Text = ""
                    Me.TBSpecialScore.Text = ""
                    '=====================================================

                    Me.TBIssueScore.Focus()
                    Me.TBIssueScore.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Call IssueClearScreen()
        End If
    End Sub

    Private Sub ListViewCoupon_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewCoupon.SelectedIndexChanged

    End Sub

    Private Sub BTNSpecialSave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSpecialSave.KeyDown, TBSpecialNo.KeyDown, DTPSpecialDate.KeyDown, CMBSpecialCampaign.KeyDown, TBSpecialScore.KeyDown, TBSpecialReason.KeyDown, BTNSpecialScoreClear.KeyDown, BTNSpecialSearch.KeyDown
        Dim vSearch As String

        On Error Resume Next

        vSearch = Me.TBSearchMember.Text

        If e.Shift = True And e.KeyCode = Keys.F1 Then
            Call vSearchMember(vSearch)
            If Me.ListViewSearchMember.Items.Count > 0 Then
                Me.ListViewSearchMember.Focus()
                Me.ListViewSearchMember.Items(0).Selected = True
                Me.ListViewSearchMember.Items(0).Focused = True
            End If
        End If

        If e.Shift = False And e.KeyCode = Keys.F1 Then
            Me.PNAccumulateScore.Visible = True
            Me.PNIssueScoreHistory.Visible = False
            Me.PNAddSpecialScore.Visible = False
            Me.PNIssueScore.Visible = False
            If Me.ListViewAccumulateScore.Items.Count > 0 Then
                Me.ListViewAccumulateScore.Focus()
                Me.ListViewAccumulateScore.Items(0).Selected = True
                Me.ListViewAccumulateScore.Items(0).Focused = True
            Else
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.F2 Then
            If vMemIsMember <> 0 Then
                Me.PNAccumulateScore.Visible = False
                Me.PNIssueScoreHistory.Visible = True

                Me.PNAddSpecialScore.Visible = False
                Me.PNIssueScore.Visible = False

                If Me.ListViewIssueScore.Items.Count > 0 Then
                    Me.ListViewIssueScore.Focus()
                    Me.ListViewIssueScore.Items(0).Selected = True
                    Me.ListViewIssueScore.Items(0).Focused = True
                Else
                    Me.TBMemberID.Focus()
                    Me.TBMemberID.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.F3 Then
            Call ChekAuthorityAccess()
            If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
                If vMemIsMember <> 0 Then
                    Me.DTPSpecialDate.Value = Now
                    Call vGetCampaignSpecial("")
                    Call vGetSpecialDocNo()
                    Call vGetNewSpecial()
                    Me.TBSpecialScore.Focus()
                    Me.TBSpecialScore.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.F4 Then
            If Me.TBArName.Text <> "" Then
                If vMemIsMember <> 0 Then
                    Call CalcMemberRemain()
                Else
                    MsgBox("ลูกค้าที่เป็นสมาชิกเท่านั้นถึงจะคำนวณแต้มได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
            End If
        End If

        If e.KeyCode = Keys.F5 Then
            Call ChekAuthorityAccess()
            If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
                If vMemIsMember <> 0 Then
                    Me.PNAccumulateScore.Visible = False
                    Me.PNIssueScoreHistory.Visible = False
                    Me.PNAddSpecialScore.Visible = False
                    Me.PNIssueScore.Visible = True

                    Call IssueNew()
                    Me.DTPIssueDate.Value = Now
                    Call vGetIssueDocNo()

                    '=====================================================
                    vMemSpecialCancel = 0
                    vMemSpecialConfirm = 0
                    vMemSpecialOpen = 0

                    Me.TBSpecialNo.Text = ""
                    Me.DTPSpecialDate.Value = Now
                    If Me.CMBSpecialCampaign.Items.Count > 0 Then
                        Me.CMBSpecialCampaign.SelectedIndex = 0
                    Else
                        Me.CMBSpecialCampaign.Text = ""
                    End If
                    Me.TBSpecialReason.Text = ""
                    Me.TBSpecialScore.Text = ""
                    '=====================================================

                    Me.TBIssueScore.Focus()
                    Me.TBIssueScore.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearSpecialScore()
        End If
    End Sub

    Private Sub BTNSearchMember_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSearchMember.KeyDown, TBMemberID.KeyDown, TBArName.KeyDown, TBArCode.KeyDown, TBArAddress.KeyDown, TBApplyDate.KeyDown, TBExpireDate.KeyDown, CMBCampaign.KeyDown, TBStartDate.KeyDown, TBEndDate.KeyDown, TBScoreRemain.KeyDown, BTNAccumulateScore.KeyDown, BTNIssueScoreHistory.KeyDown, BTNAddSpecialScore.KeyDown, BTNCalcScore.KeyDown, BTNIssueScore.KeyDown
        Dim vSearch As String
        On Error Resume Next

        vSearch = Me.TBSearchMember.Text

        If e.Shift = True And e.KeyCode = Keys.F1 Then
            Call vSearchMember(vSearch)
            If Me.ListViewSearchMember.Items.Count > 0 Then
                Me.ListViewSearchMember.Focus()
                Me.ListViewSearchMember.Items(0).Selected = True
                Me.ListViewSearchMember.Items(0).Focused = True
            End If
        End If

        If e.Shift = False And e.KeyCode = Keys.F1 Then
            Me.PNAccumulateScore.Visible = True
            Me.PNIssueScoreHistory.Visible = False
            Me.PNAddSpecialScore.Visible = False
            Me.PNIssueScore.Visible = False
            If Me.ListViewAccumulateScore.Items.Count > 0 Then
                Me.ListViewAccumulateScore.Focus()
                Me.ListViewAccumulateScore.Items(0).Selected = True
                Me.ListViewAccumulateScore.Items(0).Focused = True
            Else
                Me.TBMemberID.Focus()
                Me.TBMemberID.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.F2 Then
            If vMemIsMember <> 0 Then
                Me.PNAccumulateScore.Visible = False
                Me.PNIssueScoreHistory.Visible = True

                Me.PNAddSpecialScore.Visible = False
                Me.PNIssueScore.Visible = False

                If Me.ListViewIssueScore.Items.Count > 0 Then
                    Me.ListViewIssueScore.Focus()
                    Me.ListViewIssueScore.Items(0).Selected = True
                    Me.ListViewIssueScore.Items(0).Focused = True
                Else
                    Me.TBMemberID.Focus()
                    Me.TBMemberID.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.F3 Then
            Call ChekAuthorityAccess()
            If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
                If vMemIsMember <> 0 Then
                    Me.DTPSpecialDate.Value = Now
                    Call vGetCampaignSpecial("")
                    Call vGetSpecialDocNo()
                    Call vGetNewSpecial()
                    Me.TBSpecialScore.Focus()
                    Me.TBSpecialScore.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.F4 Then
            If Me.TBArName.Text <> "" Then
                If vMemIsMember <> 0 Then
                    Call CalcMemberRemain()
                Else
                    MsgBox("ลูกค้าที่เป็นสมาชิกเท่านั้นถึงจะคำนวณแต้มได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If
            End If
        End If

        If e.KeyCode = Keys.F5 Then
            Call ChekAuthorityAccess()
            If vDepartment = "CH" Or vDepartment = "IT" Or vDepartment = "MC" Then
                If vMemIsMember <> 0 Then
                    Me.PNAccumulateScore.Visible = False
                    Me.PNIssueScoreHistory.Visible = False
                    Me.PNAddSpecialScore.Visible = False
                    Me.PNIssueScore.Visible = True

                    Call IssueNew()
                    Me.DTPIssueDate.Value = Now
                    Call vGetIssueDocNo()

                    '=====================================================
                    vMemSpecialCancel = 0
                    vMemSpecialConfirm = 0
                    vMemSpecialOpen = 0

                    Me.TBSpecialNo.Text = ""
                    Me.DTPSpecialDate.Value = Now
                    If Me.CMBSpecialCampaign.Items.Count > 0 Then
                        Me.CMBSpecialCampaign.SelectedIndex = 0
                    Else
                        Me.CMBSpecialCampaign.Text = ""
                    End If
                    Me.TBSpecialReason.Text = ""
                    Me.TBSpecialScore.Text = ""
                    '=====================================================

                    Me.TBIssueScore.Focus()
                    Me.TBIssueScore.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearMember()
        End If

        'Dim vMemberID As String
        'If e.KeyCode = Keys.Enter Then
        '    vMemberID = Me.TBMemberID.Text
        '    Call SearchMemberDetails(vMemberID)
        'End If
    End Sub
End Class