Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic

Public Class dgvSearchSPPoint
    Dim fDoctext As String
    Dim fQry As String
    Dim dt As New DataTable
    Dim iLVitem As ListViewItem
    Dim i As Integer
    Dim icf As Integer
    Dim icl As Integer
    
    Private Sub dgvSearchSPPoint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
        Me.btnSPfind.Enabled = False
    End Sub

    Private Sub btnFsp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFsp.Click
        Call viewspData()
    End Sub
    Private Sub viewspData()
        If Me.txtFspDocNo.Text <> "" Then
            fDoctext = Me.txtFspDocNo.Text
            fQry = "exec dbo.USP_VP_PointSpecialSearch '" & fDoctext & "'"
            da = New SqlDataAdapter(fQry, vConnection)
            ds = New DataSet
            da.Fill(ds, "vSP")
            dt = ds.Tables("vSP")
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    iLVitem = Me.LVfSpPoint.Items.Add(dt.Rows(i).Item("docno"))
                    iLVitem.SubItems.Add(0).Text = dt.Rows(i).Item("docdate")
                    iLVitem.SubItems.Add(0).Text = dt.Rows(i).Item("arcode")
                    iLVitem.SubItems.Add(0).Text = dt.Rows(i).Item("arName")
                    iLVitem.SubItems.Add(0).Text = dt.Rows(i).Item("memberid")
                    iLVitem.SubItems.Add(0).Text = dt.Rows(i).Item("point")
                    iLVitem.SubItems.Add(0).Text = dt.Rows(i).Item("campaigncode")
                    iLVitem.SubItems.Add(0).Text = dt.Rows(i).Item("Reason")
                    iLVitem.SubItems.Add(0).Text = dt.Rows(i).Item("isconfirm")
                    iLVitem.SubItems.Add(0).Text = dt.Rows(i).Item("iscancel")
                Next
            Else
                MsgBox("ไม่พบข้อมูลเอกสารที่ต้องการ", MsgBoxStyle.Information, "Information")
            End If

        End If
    End Sub

    Private Sub txtFspDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFspDocNo.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call viewspData()
        End If
    End Sub

    Private Sub btnSPfind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSPfind.Click
        Call SLspPointDocno()
    End Sub
    Private Sub SLspPointDocno()
        For i = 0 To Me.LVfSpPoint.Items.Count - 1
            If Me.LVfSpPoint.Items(i).Selected = True Then
                frmPoint01.txtSPDocno.Text = Me.LVfSpPoint.Items(i).SubItems(0).Text
                frmPoint01.dtpSPDocdate.Text = Me.LVfSpPoint.Items(i).SubItems(1).Text
                frmPoint01.txtSPARCode.Text = Me.LVfSpPoint.Items(i).SubItems(2).Text
                frmPoint01.txtSParName.Text = Me.LVfSpPoint.Items(i).SubItems(3).Text
                frmPoint01.txtSPmemberid.Text = Me.LVfSpPoint.Items(i).SubItems(4).Text
                frmPoint01.txtSPpoint.Text = Me.LVfSpPoint.Items(i).SubItems(5).Text
                frmPoint01.cbxSPCampaign.SelectedValue = Me.LVfSpPoint.Items(i).SubItems(6).Text
                frmPoint01.txtIssue.Text = Me.LVfSpPoint.Items(i).SubItems(7).Text
                icf = Me.LVfSpPoint.Items(i).SubItems(8).Text
                If icf = 0 Then
                    frmPoint01.lblConfirm.Text = "--N--"
                ElseIf icf = 1 Then
                    frmPoint01.lblConfirm.Text = "--CF--"
                End If
                icl = Me.LVfSpPoint.Items(i).SubItems(9).Text
                If icl = 0 Then
                    frmPoint01.lblcancel.Visible = True
                ElseIf icl = 1 Then
                    frmPoint01.lblcancel.Text = "Cancel"
                    frmPoint01.lblcancel.Visible = True
                End If
            End If
        Next
        PsaveSPpoint = 0
        frmPoint01.btnSaveSP.Enabled = True
        Me.Close()
    End Sub

    Private Sub btnExitSP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExitSP.Click
        Me.Close()
    End Sub

    Private Sub LVfSpPoint_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles LVfSpPoint.ItemSelectionChanged
        Me.btnSPfind.Enabled = True
    End Sub

    Private Sub LVfSpPoint_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LVfSpPoint.MouseDoubleClick
        Call SLspPointDocno()
    End Sub
End Class
