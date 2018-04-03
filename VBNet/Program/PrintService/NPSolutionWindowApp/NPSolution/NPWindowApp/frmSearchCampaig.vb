Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic
Public Class frmSearchCampaig
    Dim iCPno As String
    Dim QryString As String
    Dim da As New SqlDataAdapter
    Dim ds As New DataSet
    Dim dt As New DataTable
    Dim iLVW As New ListViewItem

    Private Sub txtFindCP_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFindCP.KeyDown
        Call InitializeDataBase()
        Dim i As Integer
        Dim nl As Integer
        iCPno = Me.txtFindCP.Text
        If e.KeyCode = Keys.Enter Then
            QryString = "exec dbo.USP_VP_CampaignSearch '" & iCPno & "'"
            da = New SqlDataAdapter(QryString, vConnection)
            ds = New DataSet
            da.Fill(ds, "cpData")
            dt = ds.Tables("cpData")
            If dt.Rows.Count > 0 Then
                Me.LVFindCP.Items.Clear()
                For i = 0 To dt.Rows.Count - 1
                    nl = nl + 1
                    iLVW = Me.LVFindCP.Items.Add(nl)
                    iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("code")
                    iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("NameTH")
                    iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("NameEN")
                    iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("StartDate")
                    iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("StopDate")
                Next
            End If
        End If
    End Sub

    Private Sub frmSearchCampaig_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
    End Sub

    Private Sub btnFinCP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinCP.Click
        Call InitializeDataBase()
        Dim i As Integer
        Dim nl As Integer
        iCPno = Me.txtFindCP.Text
        QryString = "exec dbo.USP_VP_CampaignSearch '" & iCPno & "'"
        da = New SqlDataAdapter(QryString, vConnection)
        ds = New DataSet
        da.Fill(ds, "cpData")
        dt = ds.Tables("cpData")
        If dt.Rows.Count > 0 Then
            Me.LVFindCP.Items.Clear()
            For i = 0 To dt.Rows.Count - 1
                nl = nl + 1
                iLVW = Me.LVFindCP.Items.Add(nl)
                iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("code")
                iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("NameTH")
                iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("NameEN")
                iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("StartDate")
                iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("StopDate")
            Next
        Else
            MsgBox("ไม่พบข้อมูลที่ค้นหา", MsgBoxStyle.Information, "Information")
        End If
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim x As Integer
        PsaveCPstatus = 0
        For x = 0 To Me.LVFindCP.Items.Count - 1
            If Me.LVFindCP.Items(x).Selected = True Then
                frmPoint01.txtCPno.Text = Me.LVFindCP.Items(x).SubItems(1).Text
                frmPoint01.txtCPthName.Text = Me.LVFindCP.Items(x).SubItems(2).Text
                frmPoint01.txtCPenName.Text = Me.LVFindCP.Items(x).SubItems(3).Text
                frmPoint01.dtpCPStartDate.Value = Me.LVFindCP.Items(x).SubItems(4).Text
                frmPoint01.dtpCPendDate.Value = Me.LVFindCP.Items(x).SubItems(5).Text
            Else
                MsgBox("คุณไม่ได้เลือกรายการ แล้วจะกดทำไม", MsgBoxStyle.Critical, "Error")
            End If
        Next
        Me.Close()
        frmPoint01.txtCPno.ReadOnly = True
        frmPoint01.txtCPName.ReadOnly = True
        frmPoint01.txtCPenName.ReadOnly = True
        frmPoint01.dtpCPStartDate.Enabled = False
        frmPoint01.dtpCPendDate.Enabled = False
        frmPoint01.btnCPnew.Enabled = False
        frmPoint01.btnExitCP.Enabled = True
        frmPoint01.btnClearCP.Enabled = True
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class
