Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Public Class FormApproveCommission
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Private Sub FormApproveCommission_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()

        Dim n As Integer
        Dim i As Integer
        Dim vListDoc As ListViewItem

        On Error Resume Next

        Me.ListViewReqComm.Items.Clear()
        vQuery = "exec dbo.USP_COM_RequestConfirmWaiting"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchReqApprove")
        dt = ds.Tables("SearchReqApprove")
        If dt.Rows.Count > 0 Then

            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vListDoc = Me.ListViewReqComm.Items.Add(n)
                vListDoc.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                vListDoc.SubItems.Add(1).Text = dt.Rows(i).Item("docdate")
                vListDoc.SubItems.Add(2).Text = dt.Rows(i).Item("campaign")
                vListDoc.SubItems.Add(3).Text = dt.Rows(i).Item("mydescription")
            Next
            Me.CBSelectAll.Checked = False
            If Me.ListViewReqComm.Items.Count > 0 Then
                Me.ListViewReqComm.Focus()
                Me.ListViewReqComm.Items(0).Selected = True
                Me.ListViewReqComm.Items(0).Focused = True
            End If
        Else
            Me.BTNRefresh.Focus()
        End If
    End Sub

    Private Sub BTNRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNRefresh.Click
        Call GetRequest()
    End Sub

    Public Sub GetRequest()
        Dim n As Integer
        Dim i As Integer
        Dim vListDoc As ListViewItem

        On Error Resume Next

        Me.ListViewReqComm.Items.Clear()
        vQuery = "exec dbo.USP_COM_RequestConfirmWaiting"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchReqApprove")
        dt = ds.Tables("SearchReqApprove")
        If dt.Rows.Count > 0 Then

            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vListDoc = Me.ListViewReqComm.Items.Add(n)
                vListDoc.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                vListDoc.SubItems.Add(1).Text = dt.Rows(i).Item("docdate")
                vListDoc.SubItems.Add(2).Text = dt.Rows(i).Item("campaign")
                vListDoc.SubItems.Add(3).Text = dt.Rows(i).Item("mydescription")
            Next
            Me.CBSelectAll.Checked = False
            If Me.ListViewReqComm.Items.Count > 0 Then
                Me.ListViewReqComm.Focus()
                Me.ListViewReqComm.Items(0).Selected = True
                Me.ListViewReqComm.Items(0).Focused = True
            End If
        Else
            MsgBox("ไม่มีเอกสารขอเสนอสินค้าคิดค่าคอม ฯ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNRefresh.Focus()
        End If

    End Sub

    Private Sub BTNApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNApprove.Click
        Dim i As Integer
        Dim vDocNo As String
        Dim vCountItemCheck As Integer

        If Me.ListViewReqComm.Items.Count > 0 Then

            For i = 0 To Me.ListViewReqComm.Items.Count - 1
                If Me.ListViewReqComm.Items(i).Checked = True Then
                    vCountItemCheck = vCountItemCheck + 1
                End If
            Next

            If vCountItemCheck > 0 Then

                On Error GoTo ErrDescription

                vQuery = "begin tran"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()
                For i = 0 To Me.ListViewReqComm.Items.Count - 1
                    If Me.ListViewReqComm.Items(i).Checked = True Then
                        vDocNo = Me.ListViewReqComm.Items(i).SubItems(1).Text

                        vQuery = "exec dbo.USP_COM_RequestConfirm '" & vDocNo & "'"
                        cmd = New SqlCommand(vQuery, vConnection)
                        cmd.ExecuteNonQuery()

                    End If
                Next

                vQuery = "commit tran"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                MsgBox("อนุมัติใบเสนอสินค้าขอคิดค่าคอมฯ เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")
                Call GetRequest()
            Else
                MsgBox("ยังไม่ได้เลือกรายการเอกสารที่จะอนุมัติ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                If Me.ListViewReqComm.Items.Count > 0 Then
                    Me.CBSelectAll.Focus()
                End If
            End If
        Else
            MsgBox("ไม่มีรายการเอกสาร กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
        End If

ErrDescription:
        If Err.Description <> "" Then
            vQuery = "rollback tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()
            MsgBox("ไม่สามารถอนุมัติเอกสารได้ เกิดปัญหา จาก Store : dbo.USP_COM_RequestConfirm  ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If
    End Sub

    Public Sub ApproveDoc()
        Dim i As Integer
        Dim vDocNo As String
        Dim vCountItemCheck As Integer

        If Me.ListViewReqComm.Items.Count > 0 Then

            For i = 0 To Me.ListViewReqComm.Items.Count - 1
                If Me.ListViewReqComm.Items(i).Checked = True Then
                    vCountItemCheck = vCountItemCheck + 1
                End If
            Next

            If vCountItemCheck > 0 Then

                On Error GoTo ErrDescription

                vQuery = "begin tran"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()
                For i = 0 To Me.ListViewReqComm.Items.Count - 1
                    If Me.ListViewReqComm.Items(i).Checked = True Then
                        vDocNo = Me.ListViewReqComm.Items(i).SubItems(1).Text

                        vQuery = "exec dbo.USP_COM_RequestConfirm '" & vDocNo & "'"
                        cmd = New SqlCommand(vQuery, vConnection)
                        cmd.ExecuteNonQuery()

                    End If
                Next

                vQuery = "commit tran"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                MsgBox("อนุมัติใบเสนอสินค้าขอคิดค่าคอมฯ เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")
                Call GetRequest()
            Else
                MsgBox("ยังไม่ได้เลือกรายการเอกสารที่จะอนุมัติ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                If Me.ListViewReqComm.Items.Count > 0 Then
                    Me.CBSelectAll.Focus()
                End If
            End If
        Else
            MsgBox("ไม่มีรายการเอกสาร กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
        End If

ErrDescription:
        If Err.Description <> "" Then
            vQuery = "rollback tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()
            MsgBox("ไม่สามารถอนุมัติเอกสารได้ เกิดปัญหา จาก Store : dbo.USP_COM_RequestConfirm  ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If
    End Sub

    Private Sub CBSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBSelectAll.CheckedChanged
        Dim i As Integer


        If Me.ListViewReqComm.Items.Count > 0 Then
            If Me.CBSelectAll.Checked = True Then
                For i = 0 To Me.ListViewReqComm.Items.Count - 1
                    Me.ListViewReqComm.Items(i).Checked = True
                Next
            Else
                For i = 0 To Me.ListViewReqComm.Items.Count - 1
                    Me.ListViewReqComm.Items(i).Checked = False
                Next
            End If
        End If
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Me.Close()
    End Sub

    Private Sub CBSelectAll_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CBSelectAll.KeyDown, ListViewReqComm.KeyDown, BTNRefresh.KeyDown, BTNApprove.KeyDown, BTNExit.KeyDown
        Dim vCheckLine As Integer

        If Me.ListViewReqComm.Items.Count > 0 Then
            If e.KeyCode = Keys.Up Then
                vCheckLine = Me.ListViewReqComm.SelectedItems(0).Index
                If vCheckLine = 0 Then
                    Me.CBSelectAll.Focus()
                End If
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If

        If e.KeyCode = Keys.F1 Then
            Call GetRequest()
        End If

        If e.KeyCode = Keys.F5 Then
            Call ApproveDoc()
        End If
    End Sub

    Private Sub ListViewReqComm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewReqComm.SelectedIndexChanged

    End Sub
End Class