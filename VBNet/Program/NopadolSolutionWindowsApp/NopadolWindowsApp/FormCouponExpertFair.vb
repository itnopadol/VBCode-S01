Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization

Public Class FormCouponExpertFair
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim dt1 As DataTable
    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vCheckExist As Integer

    Private Sub FormCouponExpertFair_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click

    End Sub

    Private Sub MEID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MEID.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBMember.Focus()
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBMember.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.MEID.Text = ""
            Me.TBMember.Text = ""
            Me.TBCouponAmount.Text = ""
            Me.MEID.Focus()
            Me.MEID.SelectAll()
        End If
    End Sub

    Private Sub MEID_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs) Handles MEID.MaskInputRejected

    End Sub

    Private Sub TBMember_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBMember.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBCouponAmount.Focus()
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBCouponAmount.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.MEID.Focus()
            Me.MEID.SelectAll()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.MEID.Text = ""
            Me.TBMember.Text = ""
            Me.TBCouponAmount.Text = ""
            Me.MEID.Focus()
            Me.MEID.SelectAll()
        End If
    End Sub

    Private Sub TBMember_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TBMember.LostFocus
        Dim vMemberID As String
        Dim vCheckExist As Integer

        On Error Resume Next

        If Me.TBMember.Text <> "" Then
            vMemberID = Me.TBMember.Text
            vQuery = "select memberid from dbo.bcar where memberid = '" & vMemberID & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Search")
            dt = ds.Tables("Search")
            If dt.Rows.Count > 0 Then
                vCheckExist = 1
            Else
                vCheckExist = 0
            End If

            If vCheckExist = 0 Then
                MsgBox("ไม่มีรหัส สมาชิกที่กรอกไว้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBMember.Text = ""
            End If
        End If
    End Sub

    Private Sub TBMember_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBMember.TextChanged

    End Sub

    Private Sub TBCouponAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBCouponAmount.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.BTNSave.Focus()
            Me.BTNSave.Select()
        End If

        If e.KeyCode = Keys.Down Then
            Me.BTNSave.Focus()
            Me.BTNSave.Select()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBMember.Focus()
            Me.TBMember.SelectAll()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.MEID.Text = ""
            Me.TBMember.Text = ""
            Me.TBCouponAmount.Text = ""
            Me.MEID.Focus()
            Me.MEID.SelectAll()
        End If
    End Sub

    Private Sub TBCouponAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBCouponAmount.KeyPress
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

    Private Sub TBCouponAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBCouponAmount.TextChanged

    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vID As String
        Dim vMemberID As String
        Dim vCouponAmount As Double
        Dim vMemCouponAmount As Double

        Dim vAnswer As Integer
        Dim vAnswerData As Integer

        On Error Resume Next

        If Me.MEID.Text <> "" Or Me.TBMember.Text <> "" Then
            If Me.TBCouponAmount.Text = "" Then
                MsgBox("กรุณา กรอกมูลค่าคูปองที่จ่ายไปด้วย", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBCouponAmount.Focus()
                Me.TBCouponAmount.SelectAll()
                Exit Sub
            End If

            vID = Me.MEID.Text
            vMemberID = Me.TBMember.Text
            vCouponAmount = Me.TBCouponAmount.Text

            vQuery = "exec dbo.USP_NP_SearchCouponExpertFairDetails '" & vID & "','" & vMemberID & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Search")
            dt = ds.Tables("Search")
            If dt.Rows.Count > 0 Then
                vMemCouponAmount = dt.Rows(0).Item("couponamount")
                vCheckExist = 1
            Else
                vCheckExist = 0
            End If

            If vCheckExist = 1 Then
                vAnswer = MsgBox("บุคคลท่านนี้ เคยรับคูปองเงินสด มูลค่า " & vMemCouponAmount & " บาท ก่อนหน้านี้แล้ว คุณต้องการทำรายการต่อหรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
                If vAnswer = 6 Then
                    vAnswerData = MsgBox("คุณต้องการให้เพิ่มข้อมูลใช่หรือไม่ กด Yes = เพิ่ม กด No = แก้ไข", MsgBoxStyle.YesNo, "Send Question Message")
                    If vAnswerData = 6 Then
                        vQuery = "exec dbo.USP_NP_InsertCouponExpertFair 1,'" & vID & "','" & vMemberID & "'," & vCouponAmount & ",'" & vUserID & "'"
                        cmd = New SqlCommand(vQuery, vConnection)
                        cmd.ExecuteNonQuery()
                    Else
                        vQuery = "exec dbo.USP_NP_InsertCouponExpertFair 2,'" & vID & "','" & vMemberID & "'," & vCouponAmount & ",'" & vUserID & "'"
                        cmd = New SqlCommand(vQuery, vConnection)
                        cmd.ExecuteNonQuery()
                    End If

                    Me.MEID.Text = ""
                    Me.TBMember.Text = ""
                    Me.TBCouponAmount.Text = ""
                    Me.MEID.Focus()
                    Me.MEID.SelectAll()
                Else
                    Me.MEID.Text = ""
                    Me.TBMember.Text = ""
                    Me.TBCouponAmount.Text = ""
                    Me.MEID.Focus()
                    Me.MEID.SelectAll()
                    Exit Sub
                End If
            Else
                vQuery = "exec dbo.USP_NP_InsertCouponExpertFair 1,'" & vID & "','" & vMemberID & "'," & vCouponAmount & ",'" & vUserID & "'"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()
                Me.MEID.Text = ""
                Me.TBMember.Text = ""
                Me.TBCouponAmount.Text = ""
                Me.MEID.Focus()
                Me.MEID.SelectAll()
            End If
        Else
            MsgBox("กรุณากรอก เลขที่บัตรประชาชนหรือรหัสสมาชิก ด้วยก่อนบันทึกการจ่ายคูปอง", MsgBoxStyle.Critical, "Send Error Message")
            Me.MEID.Focus()
            Me.MEID.SelectAll()
        End If

    End Sub

    Private Sub TBSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call Search(Me.TBSearch.Text)
        End If
    End Sub

    Private Sub TBSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearch.TextChanged

    End Sub


    Public Sub SearchMember(ByVal vSearch As String)
        Dim i As Integer
        Dim n As Integer
        Dim vAmount As Double
        Dim vList As ListViewItem

        On Error Resume Next

        Me.ListViewMember.Items.Clear()
        vQuery = "select * from (select memberid,name1 as arname from dbo.bcar where memberid like '%'+ '" & vSearch & "' +'%' and memberid is not null union select memberid,name1 as arname from dbo.bcar where name1 like '%'+ '" & vSearch & "' +'%' and memberid is not null union select memberid,name1 as arname from dbo.bcar where code like '%'+ '" & vSearch & "' +'%' and memberid is not null) as result order by memberid"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Search")
        dt = ds.Tables("Search")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vList = Me.ListViewMember.Items.Add(n)
                vList.SubItems.Add(1).Text = dt.Rows(i).Item("memberid")
                vList.SubItems.Add(2).Text = dt.Rows(i).Item("arname")

            Next
        End If

        Dim a As Integer
        If Me.ListViewMember.Items.Count > 0 Then
            For a = 0 To Me.ListViewMember.Items.Count - 1
                If a Mod 2 = 0 Then
                    Me.ListViewMember.Items(a).BackColor = Color.White
                Else
                    Me.ListViewMember.Items(a).BackColor = Color.LightBlue
                End If
            Next
        End If

    End Sub

    Public Sub Search(ByVal vSearch As String)
        Dim i As Integer
        Dim n As Integer
        Dim vAmount As Double
        Dim vList As ListViewItem

        On Error Resume Next

        Me.ListViewSearch.Items.Clear()
        vQuery = "exec dbo.USP_NP_SearchCouponExpertFair '" & vSearch & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Search")
        dt = ds.Tables("Search")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vAmount = dt.Rows(i).Item("couponamount")
                vList = Me.ListViewSearch.Items.Add(n)
                vList.SubItems.Add(0).Text = dt.Rows(i).Item("id")
                vList.SubItems.Add(1).Text = dt.Rows(i).Item("memberid")
                vList.SubItems.Add(2).Text = dt.Rows(i).Item("arname")
                vList.SubItems.Add(3).Text = Format(vAmount, "##,##0.00")
                vList.SubItems.Add(4).Text = dt.Rows(i).Item("numberid")
            Next
        End If

        Dim a As Integer
        If Me.ListViewSearch.Items.Count > 0 Then
            For a = 0 To Me.ListViewSearch.Items.Count - 1
                If a Mod 2 = 0 Then
                    Me.ListViewSearch.Items(a).BackColor = Color.White
                Else
                    Me.ListViewSearch.Items(a).BackColor = Color.LightBlue
                End If
            Next
        End If

    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click

        Call Search(Me.TBSearch.Text)

    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        Me.TBSearch.Text = ""
        Me.ListViewSearch.Items.Clear()
        Me.PNSearch.Visible = False
        Me.MEID.Focus()
        Me.MEID.SelectAll()
    End Sub

    Private Sub BTNSearchID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchID.Click
        Call Search("")
        Me.PNSearch.Visible = True
        Me.TBSearch.Focus()
        Me.TBSearch.SelectAll()
    End Sub

    Private Sub BTNClose_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClose.KeyDown, BTNSearch.KeyDown, ListViewSearch.KeyDown, TBSearch.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.TBSearch.Text = ""
            Me.ListViewSearch.Items.Clear()
            Me.PNSearch.Visible = False
            Me.MEID.Focus()
            Me.MEID.SelectAll()
        End If
    End Sub

    Private Sub BTNMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNMember.Click
        Me.PNMember.Visible = True
        Me.TBSearchMember.Focus()
        If Me.TBSearchMember.Text <> "" Then
            Call SearchMember(Me.TBSearchMember.Text)
        End If
    End Sub

    Private Sub BTNSearchMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchMember.Click
        If Me.TBSearchMember.Text <> "" Then
            Call SearchMember(Me.TBSearchMember.Text)
        End If
    End Sub

    Private Sub BTNMemberClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNMemberClose.Click
        Me.TBSearchMember.Text = ""
        Me.ListViewMember.Items.Clear()
        Me.PNMember.Visible = False
        Me.MEID.Focus()
        Me.MEID.SelectAll()
    End Sub

    Private Sub TBSearchMember_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchMember.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Me.TBSearchMember.Text <> "" Then
                Call SearchMember(Me.TBSearchMember.Text)
            End If
        End If
    End Sub

    Private Sub TBSearchMember_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchMember.TextChanged

    End Sub

    Private Sub ListViewMember_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewMember.DoubleClick
        Dim vIndex As Integer

        On Error Resume Next

        If Me.ListViewMember.Items.Count > 0 Then
            vIndex = Me.ListViewMember.SelectedItems(0).Index
            Me.TBMember.Text = Me.ListViewMember.Items(vIndex).SubItems(1).Text
            Me.PNMember.Visible = False
            Me.TBCouponAmount.Focus()
            Me.TBCouponAmount.SelectAll()
        End If
    End Sub

    Private Sub BTNMemberClose_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNMemberClose.KeyDown, BTNSearchMember.KeyDown, ListViewMember.KeyDown, TBSearchMember.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.TBSearchMember.Text = ""
            Me.ListViewMember.Items.Clear()
            Me.PNMember.Visible = False
            Me.MEID.Focus()
            Me.MEID.SelectAll()
        End If
    End Sub

    Private Sub ListViewMember_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewMember.SelectedIndexChanged

    End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    Dim vPrinterName As String
    '    Dim vDocNo As String

    '    vQuery = "Exec dbo.USP_NP_SearchPrinterName '02','01'"
    '    da = New SqlDataAdapter(vQuery, vConnection)
    '    ds = New DataSet
    '    da.Fill(ds, "Search")
    '    dt1 = ds.Tables("Search")
    '    If dt1.Rows.Count > 0 Then
    '        vPrinterName = dt1.Rows(0).Item("printername")
    '    End If


    '    vQuery = "Exec dbo.USP_VP_WithdrawSearchSubNew '" & vDocNo & "'"
    '    da = New SqlDataAdapter(vQuery, vConnection)
    '    ds = New DataSet
    '    da.Fill(ds, "Search")
    '    dt = ds.Tables("Search")

    '    Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    '    Dim frmObj As New FormReportReqItemComm
    '    Dim FileName As New String("V:\Reports\NopadolCoupon\RP_NP_CashMemberCoupon.rpt")

    '    rpt.Load(FileName)

    '    Dim Params As New CrystalDecisions.Shared.ParameterField
    '    Dim ParamCollection As New CrystalDecisions.Shared.ParameterFields
    '    Dim ParamDisVal As New CrystalDecisions.Shared.ParameterDiscreteValue()
    '    Params.ParameterFieldName = "@DocNo"
    '    ParamDisVal.Value = vDocNo
    '    Params.CurrentValues.Add(ParamDisVal)
    '    ParamCollection.Add(Params)

    '    rpt.Load(FileName)
    '    rpt.SetDataSource(ds.Tables("Search"))
    '    rpt.SetParameterValue("@DocNo", ParamDisVal)
    '    rpt.PrintOptions.PrinterName = vPrinterName
    '    rpt.PrintToPrinter(1, False, 0, 0)
    'End Sub

    Private Sub BTNSearchID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSearchID.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.MEID.Text = ""
            Me.TBMember.Text = ""
            Me.TBCouponAmount.Text = ""
            Me.MEID.Focus()
            Me.MEID.SelectAll()
        End If
    End Sub

    Private Sub BTNMember_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNMember.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.MEID.Text = ""
            Me.TBMember.Text = ""
            Me.TBCouponAmount.Text = ""
            Me.MEID.Focus()
            Me.MEID.SelectAll()
        End If
    End Sub

    Private Sub BTNSave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSave.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.MEID.Text = ""
            Me.TBMember.Text = ""
            Me.TBCouponAmount.Text = ""
            Me.MEID.Focus()
            Me.MEID.SelectAll()
        End If
    End Sub
End Class