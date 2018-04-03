Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports CrystalDecisions
Imports System
Imports Microsoft

Public Class FormReqCommission
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vIsOpen As Integer
    Dim vMemIsCancel As Integer
    Dim vMemIsConfirm As Integer

    Dim vIsNumber As Integer

    Dim vMemColumn As Integer
    Dim vMemRow As Integer

    Private Sub FormReqCommission_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()
        Call vGetBeginData()
        Call NewDoc()
        Call vGendocNoAuto()
    End Sub

    Public Sub NewDoc()
        Me.PBNew.Visible = True
        Me.PBCancel.Visible = False
        Me.PBConfirm.Visible = False
    End Sub

    Public Sub ConfirmDoc()
        Me.PBNew.Visible = False
        Me.PBCancel.Visible = False
        Me.PBConfirm.Visible = True
    End Sub

    Public Sub CancelDoc()
        Me.PBNew.Visible = False
        Me.PBCancel.Visible = True
        Me.PBConfirm.Visible = False
    End Sub

    Public Sub vGetBeginData()
        Dim i As Integer
        Dim n As Integer

        On Error Resume Next

        Me.DGVItemDetails.Rows.Add(300)
        For i = 0 To 300 - 1
            n = n + 1
            Me.DGVItemDetails.Item(0, i).Value = n
        Next

        Me.DGVItemDetails.CurrentCell = Me.DGVItemDetails.Item(1, 0)
    End Sub

    Private Sub DGVItemDetails_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVItemDetails.CellBeginEdit
        Dim vRow As Integer
        Dim vLine As Integer

        On Error Resume Next

        vRow = Me.DGVItemDetails.CurrentCell.RowIndex
        vLine = Me.DGVItemDetails.Item(0, vRow).Value

        If vLine = 0 Then
            Me.DGVItemDetails.Columns(0).ReadOnly = False
            Me.DGVItemDetails.Item(0, vRow).Value = vRow + 1
        End If
        Me.DGVItemDetails.Columns(0).ReadOnly = True

    End Sub

    Private Sub DGVItemDetails_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItemDetails.CellEndEdit
        Dim vItemCode As String
        Dim vColumn As Integer
        Dim vRow As Integer
        Dim i As Integer

        Dim vCheckItemCode As String
        Dim vMemCountCheck As Integer

        On Error Resume Next

        vColumn = Me.DGVItemDetails.CurrentCell.ColumnIndex
        vRow = Me.DGVItemDetails.CurrentCell.RowIndex
        vItemCode = Me.DGVItemDetails.CurrentCell.Value

        If vColumn = 1 Then
            If vItemCode <> "" Then
                For i = 0 To Me.DGVItemDetails.Rows.Count - 1
                    vCheckItemCode = Me.DGVItemDetails.Item(1, i).Value

                    If vCheckItemCode = vItemCode Then
                        vMemCountCheck = vMemCountCheck + 1
                    End If
                Next

                If vMemCountCheck > 1 Then
                    MsgBox("สินค้า รหัส " & vItemCode & " มีอยู่แล้วในรายการเสนอขอคิดค่าคอมฯ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(vColumn, vRow).Value = ""
                    Exit Sub
                End If

                vQuery = "exec dbo.usp_np_searchitemdescription '" & vItemCode & "' "
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "CheckItem")
                dt = ds.Tables("CheckItem")
                If dt.Rows.Count > 0 Then
                    Me.DGVItemDetails.Item(2, vRow).Value = dt.Rows(0).Item("itemname")
                    Me.DGVItemDetails.Item(3, vRow).Value = dt.Rows(0).Item("unitcode")
                    Me.DGVItemDetails.Item(4, vRow).Value = 1
                Else
                    MsgBox("สินค้า รหัส " & vItemCode & " ไม่มีข้อมูลในระบบ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")

                    Me.DGVItemDetails.Item(1, vRow).Value = ""
                    Me.DGVItemDetails.Item(2, vRow).Value = ""
                    Me.DGVItemDetails.Item(3, vRow).Value = ""
                    Me.DGVItemDetails.Item(4, vRow).Value = ""
                    Me.DGVItemDetails.Item(5, vRow).Value = ""
                    Me.DGVItemDetails.Item(6, vRow).Value = ""
                End If
            End If
        End If

        Dim vCharStr As String
        If e.ColumnIndex = 4 Then
            vCharStr = Me.DGVItemDetails.Item(4, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(4, e.RowIndex).Value = ""
                    MsgBox("ช่องจำนวนที่กำหนด ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                End If
            End If
        End If

        If e.ColumnIndex = 5 Then
            vCharStr = Me.DGVItemDetails.Item(5, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(5, e.RowIndex).Value = ""
                    MsgBox("ช่องค่าคอมฯขายสด ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                End If
            End If
        End If

        If e.ColumnIndex = 6 Then
            vCharStr = Me.DGVItemDetails.Item(6, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(6, e.RowIndex).Value = ""
                    MsgBox("ช่องค่าคอมฯขายเชื่อ ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                End If
            End If
        End If
    End Sub

    Private Sub BTNDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNDocNo.Click
        If vIsOpen = 0 Then
            Call vGenDocNoAuto()
        Else
            Call ClearScreen()
            Call vGenDocNoAuto()
        End If
    End Sub

    Public Sub vGenDocNoAuto()
        Dim vNow As Date

        On Error Resume Next

        vQuery = "set dateformat dmy"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vNow = Now.Day & "/" & Now.Month & "/" & Now.Year
        vQuery = "select dbo.ft_com_newrequest ('" & vNow & "') as docno"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchDocno")
        dt = ds.Tables("SearchDocno")
        If dt.Rows.Count > 0 Then
            Me.TBDocNo.Text = dt.Rows(0).Item("docno")
        Else
            Me.TBDocNo.Text = ""
            MsgBox("กำหนด เลขที่เอกสารไม่ได้ เกิดปัญหา จาก Store : dbo.ft_com_newrequest  ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub BTNCampaign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCampaign.Click
        Call SearchCampaign()
    End Sub

    Public Sub SearchCampaign()
        Dim vSearch As String
        Dim i As Integer
        Dim n As Integer
        Dim vListItem As ListViewItem

        On Error Resume Next

        vSearch = Me.TBSearch.Text
        Me.ListViewSearch.Items.Clear()
        vQuery = "exec dbo.USP_COM_CampaignSearch '" & vSearch & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Search")
        dt = ds.Tables("Search")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vListItem = Me.ListViewSearch.Items.Add(n)
                vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("code")
                vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("name1")
                vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("begindate")
                vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("enddate")
                vListItem.SubItems.Add(4).Text = dt.Rows(i).Item("creatorcode")
            Next

            Me.PNSearch.Visible = True
            If Me.ListViewSearch.Items.Count > 0 Then
                Me.ListViewSearch.Focus()
                Me.ListViewSearch.Items(0).Focused = True
                Me.ListViewSearch.Items(0).Selected = True
            Else
                Me.TBSearch.Text = ""
                Me.TBSearch.Focus()
                Me.TBSearch.SelectAll()
            End If

        End If
    End Sub

    Private Sub ListViewSearch_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewSearch.DoubleClick
        Dim vIndex As Integer

        On Error Resume Next

        If Me.ListViewSearch.Items.Count > 0 Then
            vIndex = Me.ListViewSearch.SelectedItems(0).Index
            Me.TBCampaignNo.Text = Me.ListViewSearch.Items(vIndex).SubItems(1).Text
            Me.LBLCampaignName.Text = Me.ListViewSearch.Items(vIndex).SubItems(2).Text

            Me.PNSearch.Visible = False
            Me.TBMyDescription.Focus()
            Me.TBMyDescription.SelectAll()
        End If
    End Sub

    Private Sub ListViewSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearch.KeyDown
        Dim vIndex As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewSearch.Items.Count > 0 Then
                vIndex = Me.ListViewSearch.SelectedItems(0).Index
                Me.TBCampaignNo.Text = Me.ListViewSearch.Items(vIndex).SubItems(1).Text
                Me.LBLCampaignName.Text = Me.ListViewSearch.Items(vIndex).SubItems(2).Text

                Me.PNSearch.Visible = False
                Me.TBMyDescription.Focus()
                Me.TBMyDescription.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSearch.Visible = False
            Me.BTNCampaign.Focus()
        End If

        Dim vCheckLine As Integer

        If Me.ListViewSearch.Items.Count > 0 Then
            If e.KeyCode = Keys.Up Then
                vCheckLine = Me.ListViewSearch.SelectedItems(0).Index
                If vCheckLine = 0 Then
                    Me.TBSearch.Focus()
                    Me.TBSearch.SelectAll()
                End If
            End If
        End If
    End Sub

    Private Sub ListViewSearch_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewSearch.SelectedIndexChanged

    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        Me.PNSearch.Visible = False
        Me.BTNCampaign.Focus()
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vDocNo As String
        Dim vDocDate As String
        Dim vCampaignCode As String
        Dim vMyDescription As String
        Dim i As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vQty As Double
        Dim vCommCash As Double
        Dim vCommCredit As Double
        Dim vMemItemCount As Integer
        Dim vMemItemNotKeyComm As Integer

        Dim vMemBeginTran As Integer
        Dim vTypeInsert As Integer
        Dim vCharStr As String

        Dim vCheckQty As Double
        Dim vCash As Double
        Dim vCredit As Double

        On Error GoTo ErrDescription

        If Me.TBDocNo.Text = "" Then
            MsgBox("ยังไม่ได้ กำหนดเลขที่เอกสาร", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNDocNo.Focus()
            Exit Sub
        End If

        If Me.TBCampaignNo.Text = "" Then
            MsgBox("ยังไม่ได้ กำหนดเลขที่แคมเปญ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNCampaign.Focus()
            Exit Sub
        End If

        If vMemIsConfirm = 1 Then
            MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNDocNo.Focus()
            Exit Sub
        End If

        If vMemIsCancel = 1 Then
            MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNDocNo.Focus()
            Exit Sub
        End If

        If Me.DGVItemDetails.RowCount > 0 Then
            For i = 0 To Me.DGVItemDetails.RowCount - 1

                vItemCode = Me.DGVItemDetails.Item(1, i).Value
                vItemName = Me.DGVItemDetails.Item(2, i).Value
                vUnitCode = Me.DGVItemDetails.Item(3, i).Value

                If vItemCode <> "" And vItemName <> "" And vUnitCode <> "" Then
                    vMemItemCount = vMemItemCount + 1
                End If
            Next
        Else
            MsgBox("ยังไม่ได้ กำหนดรหัสสินค้าที่จะเสนอขอคิดค่าคอมฯ", MsgBoxStyle.Critical, "Send Error Message")
            Me.DGVItemDetails.Focus()
            Exit Sub
        End If

        For i = 0 To Me.DGVItemDetails.RowCount - 1
            vItemCode = Me.DGVItemDetails.Item(1, i).Value
            vItemName = Me.DGVItemDetails.Item(2, i).Value
            vUnitCode = Me.DGVItemDetails.Item(3, i).Value

            If Me.DGVItemDetails.Item(4, i).Value <> "" Then
                vQty = Me.DGVItemDetails.Item(4, i).Value
            End If

            If Me.DGVItemDetails.Item(5, i).Value <> "" Then
                vCash = Me.DGVItemDetails.Item(5, i).Value
            End If

            If Me.DGVItemDetails.Item(6, i).Value <> "" Then
                vCredit = Me.DGVItemDetails.Item(5, i).Value
            End If

            If vItemCode <> "" And vItemName <> "" And vUnitCode <> "" And vQty = 0 Then
                vMemItemNotKeyComm = vMemItemNotKeyComm + 1
            End If
        Next

        If vMemItemNotKeyComm > 0 Then
            MsgBox("ยังไม่ได้ กำหนดจำนวนหรือค่าคอมของแต่ละรหัสสินค้าที่จะเสนอขอคิดค่าคอมฯ", MsgBoxStyle.Critical, "Send Error Message")
            Me.DGVItemDetails.Focus()
            Exit Sub
        End If

        If vMemItemCount > 0 Then
            vDocNo = Me.TBDocNo.Text
            vDocDate = vb6.Day(Me.DTPDocDate.Value) & "/" & vb6.Month(Me.DTPDocDate.Value) & "/" & vb6.Year(Me.DTPDocDate.Value)
            vCampaignCode = Me.TBCampaignNo.Text
            vMyDescription = Me.TBMyDescription.Text
            If vIsOpen = 0 Then
                vTypeInsert = 1
            Else
                vTypeInsert = 0
            End If


            vQuery = "begin tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            vMemBeginTran = 1

            vQuery = "exec dbo.USP_COM_RequestSave '" & vTypeInsert & "','" & vDocNo & "','" & vDocDate & "','" & vCampaignCode & "','" & vMyDescription & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                vItemCode = Me.DGVItemDetails.Item(1, i).Value
                vItemName = Me.DGVItemDetails.Item(2, i).Value
                vUnitCode = Me.DGVItemDetails.Item(3, i).Value

                vCharStr = Me.DGVItemDetails.Item(4, i).Value
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    MsgBox("รายการที่ " & i + 1 & " ในช่องจำนวนที่กำหนด กรอกตัวเลขได้เท่านั้น", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(4, i).Selected = True

                    vQuery = "rollback tran"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Exit Sub
                End If

                vCharStr = Me.DGVItemDetails.Item(5, i).Value
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    MsgBox("รายการที่ " & i + 1 & " ในช่องค่าคอมฯขายสด กรอกตัวเลขได้เท่านั้น", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(5, i).Selected = True

                    vQuery = "rollback tran"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Exit Sub
                End If

                vCharStr = Me.DGVItemDetails.Item(6, i).Value
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    MsgBox("รายการที่ " & i + 1 & " ในช่องค่าคอมฯขายเชื่อ กรอกตัวเลขได้เท่านั้น", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(6, i).Selected = True

                    vQuery = "rollback tran"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Exit Sub
                End If

                If Me.DGVItemDetails.Item(4, i).Value <> "" Then
                    vQty = Me.DGVItemDetails.Item(4, i).Value
                End If
                If Me.DGVItemDetails.Item(5, i).Value <> "" Then
                    vCommCash = Me.DGVItemDetails.Item(5, i).Value
                End If
                If Me.DGVItemDetails.Item(6, i).Value <> "" Then
                    vCommCredit = Me.DGVItemDetails.Item(6, i).Value
                End If

                If vItemCode <> "" And vItemName <> "" And vQty <> 0 And vUnitCode <> "" Then
                    vQuery = "exec dbo.USP_COM_RequestSubSave '" & vDocNo & "'," & i & ",'" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vQty & "," & vCommCash & "," & vCommCredit & " "
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()
                End If
            Next

            vQuery = "exec dbo.USP_COM_RequestCompleteSave '" & vDocNo & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            vQuery = "commit tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()
            MsgBox("บันทึกเอกสาร เลขที่ " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")




            Call ClearScreen()
            Call vGendocNoAuto()

        End If

ErrDescription:
        If Err.Description <> "" Then
            If vMemBeginTran = 1 Then
                vQuery = "rollback tran"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()
            End If
            MsgBox("ไม่สามารถบันทึกเอกสารได้ เกิดปัญหา จาก Store : dbo.USP_COM_CampaignSave  ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Error")
            Me.BTNSave.Focus()
            Exit Sub
        End If
    End Sub



    Public Sub SaveData()
        Dim vDocNo As String
        Dim vDocDate As String
        Dim vCampaignCode As String
        Dim vMyDescription As String
        Dim i As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vQty As Double
        Dim vCommCash As Double
        Dim vCommCredit As Double
        Dim vMemItemCount As Integer
        Dim vMemItemNotKeyComm As Integer

        Dim vMemBeginTran As Integer
        Dim vTypeInsert As Integer
        Dim vCharStr As String

        On Error GoTo ErrDescription

        If Me.TBDocNo.Text = "" Then
            MsgBox("ยังไม่ได้ กำหนดเลขที่เอกสาร", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNDocNo.Focus()
            Exit Sub
        End If

        If Me.TBCampaignNo.Text = "" Then
            MsgBox("ยังไม่ได้ กำหนดเลขที่แคมเปญ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNCampaign.Focus()
            Exit Sub
        End If

        If vMemIsConfirm = 1 Then
            MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNDocNo.Focus()
            Exit Sub
        End If

        If vMemIsCancel = 1 Then
            MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNDocNo.Focus()
            Exit Sub
        End If

        If Me.DGVItemDetails.RowCount > 0 Then
            For i = 0 To Me.DGVItemDetails.RowCount - 1
                If Me.DGVItemDetails.Item(2, i).Value <> "" Then
                    vMemItemCount = vMemItemCount + 1
                End If
            Next
        Else
            MsgBox("ยังไม่ได้ กำหนดรหัสสินค้าที่จะเสนอขอคิดค่าคอมฯ", MsgBoxStyle.Critical, "Send Error Message")
            Me.DGVItemDetails.Focus()
            Exit Sub
        End If

        For i = 0 To Me.DGVItemDetails.RowCount - 1
            If Me.DGVItemDetails.Item(2, i).Value <> "" And Me.DGVItemDetails.Item(3, i).Value <> "" And (Me.DGVItemDetails.Item(4, i).Value = "" Or Me.DGVItemDetails.Item(5, i).Value = "" Or Me.DGVItemDetails.Item(6, i).Value = "") Then
                vMemItemNotKeyComm = vMemItemNotKeyComm + 1
            End If
        Next

        If vMemItemNotKeyComm > 0 Then
            MsgBox("ยังไม่ได้ กำหนดจำนวนหรือค่าคอมของแต่ละรหัสสินค้าที่จะเสนอขอคิดค่าคอมฯ", MsgBoxStyle.Critical, "Send Error Message")
            Me.DGVItemDetails.Focus()
            Exit Sub
        End If

        If vMemItemCount > 0 Then
            vDocNo = Me.TBDocNo.Text
            vDocDate = vb6.Day(Me.DTPDocDate.Value) & "/" & vb6.Month(Me.DTPDocDate.Value) & "/" & vb6.Year(Me.DTPDocDate.Value)
            vCampaignCode = Me.TBCampaignNo.Text
            vMyDescription = Me.TBMyDescription.Text
            If vIsOpen = 0 Then
                vTypeInsert = 1
            Else
                vTypeInsert = 0
            End If


            vQuery = "begin tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            vMemBeginTran = 1

            vQuery = "exec dbo.USP_COM_RequestSave '" & vTypeInsert & "','" & vDocNo & "','" & vDocDate & "','" & vCampaignCode & "','" & vMyDescription & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                vItemCode = Me.DGVItemDetails.Item(1, i).Value
                vItemName = Me.DGVItemDetails.Item(2, i).Value
                vUnitCode = Me.DGVItemDetails.Item(3, i).Value

                vCharStr = Me.DGVItemDetails.Item(4, i).Value
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(4, i).Selected = True

                    vQuery = "rollback tran"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Exit Sub
                End If

                vCharStr = Me.DGVItemDetails.Item(5, i).Value
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(5, i).Selected = True

                    vQuery = "rollback tran"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Exit Sub
                End If

                vCharStr = Me.DGVItemDetails.Item(6, i).Value
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(6, i).Selected = True

                    vQuery = "rollback tran"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Exit Sub
                End If

                If Me.DGVItemDetails.Item(4, i).Value <> "" Then
                    vQty = Me.DGVItemDetails.Item(4, i).Value
                End If
                If Me.DGVItemDetails.Item(5, i).Value <> "" Then
                    vCommCash = Me.DGVItemDetails.Item(5, i).Value
                End If
                If Me.DGVItemDetails.Item(6, i).Value <> "" Then
                    vCommCredit = Me.DGVItemDetails.Item(6, i).Value
                End If

                If vItemCode <> "" And vItemName <> "" And vQty <> 0 And vUnitCode <> "" And vCommCash <> 0 And vCommCredit <> 0 Then
                    vQuery = "exec dbo.USP_COM_RequestSubSave '" & vDocNo & "'," & i & ",'" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vQty & "," & vCommCash & "," & vCommCredit & " "
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()
                End If
            Next

            vQuery = "exec dbo.USP_COM_RequestCompleteSave '" & vDocNo & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            vQuery = "commit tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            MsgBox("บันทึกเอกสาร เลขที่ " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")

            Call ClearScreen()
            Call vGenDocNoAuto()

        End If

ErrDescription:
        If Err.Description <> "" Then
            If vMemBeginTran = 1 Then
                vQuery = "rollback tran"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()
            End If
            MsgBox("ไม่สามารถบันทึกเอกสารได้ เกิดปัญหา จาก Store : dbo.USP_COM_CampaignSave  ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Error")
            Me.BTNSave.Focus()
            Exit Sub
        End If
    End Sub

    Public Sub ClearScreen()
        On Error Resume Next

        vIsOpen = 0
        vMemIsCancel = 0
        vMemIsConfirm = 0
        Me.TBDocNo.Text = ""
        Me.TBCampaignNo.Text = ""
        Me.DTPDocDate.Value = Now
        Me.LBLCampaignName.Text = ""
        Me.TBMyDescription.Text = ""
        Call NewDoc()
        Call ClearDataDGV()
        Me.BTNDocNo.Focus()
    End Sub

    Public Sub vCheckNumber(ByVal vNumber As String)
        Dim vLen As Integer
        Dim vChar As String
        Dim i As Integer
        Dim vString As String

        On Error Resume Next

        vString = vNumber
        vLen = vb6.Len(vString)
        For i = 1 To vLen
            vChar = Mid(vString, i, 1)

            If vChar = "1" Or vChar = "2" Or vChar = "3" Or vChar = "4" Or vChar = "5" Or vChar = "6" Or vChar = "7" Or vChar = "8" Or vChar = "9" Or vChar = "0" Or vChar = "," Or vChar = "." Or vChar = "%" Then
                vIsNumber = 1
            Else
                vIsNumber = 0
                GoTo Line1
            End If
        Next
Line1:

    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        Call SearchDocNo()
    End Sub

    Public Sub SearchDocNo()
        Dim i As Integer
        Dim n As Integer
        Dim vListItem As ListViewItem
        Dim vSearch As String

        On Error Resume Next

        vSearch = Me.TBSearchDocNo.Text
        Me.ListViewSearchDocNo.Items.Clear()
        vQuery = "exec dbo.USP_COM_RequestSearch '" & vSearch & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Search")
        dt = ds.Tables("Search")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vListItem = Me.ListViewSearchDocNo.Items.Add(n)
                vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("docdate")
                vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("campaigncode")
                vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("campaignname")
                vListItem.SubItems.Add(4).Text = dt.Rows(i).Item("creatorcode")
            Next

            Me.PNSearchDocNo.Visible = True
            If Me.ListViewSearchDocNo.Items.Count > 0 Then
                Me.ListViewSearchDocNo.Focus()
                Me.ListViewSearchDocNo.Items(0).Focused = True
                Me.ListViewSearchDocNo.Items(0).Selected = True
            Else
                Me.TBSearchDocNo.Text = ""
                Me.TBSearchDocNo.Focus()
                Me.TBSearchDocNo.SelectAll()
            End If
        Else
            Me.TBDocNo.Text = ""
            Me.TBDocNo.Focus()
        End If
    End Sub

    Private Sub BTNCloseSearchDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSearchDocNo.Click
        Me.PNSearchDocNo.Visible = False
        Me.TBDocNo.Focus()
        Me.TBDocNo.SelectAll()
    End Sub

    Public Sub ClearDataDGV()
        Dim i As Integer

        On Error Resume Next

        For i = 0 To Me.DGVItemDetails.Rows.Count - 1
            Me.DGVItemDetails.Item(1, i).Value = ""
            Me.DGVItemDetails.Item(2, i).Value = ""
            Me.DGVItemDetails.Item(3, i).Value = ""
            Me.DGVItemDetails.Item(4, i).Value = ""
            Me.DGVItemDetails.Item(5, i).Value = ""
            Me.DGVItemDetails.Item(6, i).Value = ""
        Next
    End Sub

    Private Sub ListViewSearchDocNo_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewSearchDocNo.DoubleClick
        Dim vDocNo As String
        Dim i As Integer
        Dim vQty As Double
        Dim vRetailCom As Double
        Dim vWholeSaleCom As Double

        On Error Resume Next

        If Me.ListViewSearchDocNo.Items.Count > 0 Then
            vDocNo = Me.ListViewSearchDocNo.SelectedItems(0).SubItems(1).Text

            Call ClearDataDGV()
            vQuery = "exec dbo.USP_COM_RequestSearch2 '" & vDocNo & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Search")
            dt = ds.Tables("Search")
            If dt.Rows.Count > 0 Then
                vIsOpen = 1
                vMemIsConfirm = dt.Rows(i).Item("isconfirm")
                vMemIsCancel = dt.Rows(i).Item("iscancel")
                Me.TBDocNo.Text = dt.Rows(i).Item("docno")
                Me.DTPDocDate.Text = dt.Rows(i).Item("docdate")
                Me.TBCampaignNo.Text = dt.Rows(i).Item("campaigncode")
                Me.LBLCampaignName.Text = dt.Rows(i).Item("campaignname")
                Me.TBMyDescription.Text = dt.Rows(i).Item("mydesc")

                For i = 0 To dt.Rows.Count - 1

                    vQty = dt.Rows(i).Item("conditionqty")
                    vRetailCom = dt.Rows(i).Item("retailcom")
                    vWholeSaleCom = dt.Rows(i).Item("wholesalecom")

                    Me.DGVItemDetails.Item(1, i).Value = dt.Rows(i).Item("itemcode")
                    Me.DGVItemDetails.Item(2, i).Value = dt.Rows(i).Item("itemname")
                    Me.DGVItemDetails.Item(3, i).Value = dt.Rows(i).Item("unitcode")
                    Me.DGVItemDetails.Item(4, i).Value = Format(vQty, "##,##0.00")
                    Me.DGVItemDetails.Item(5, i).Value = Format(vRetailCom, "##,##0.00")
                    Me.DGVItemDetails.Item(6, i).Value = Format(vWholeSaleCom, "##,##0.00")
                Next

                If vMemIsConfirm = 1 Then
                    Call ConfirmDoc()
                End If

                If vMemIsCancel = 1 Then
                    Call CancelDoc()
                End If

                If vMemIsCancel = 0 And vMemIsConfirm = 0 Then
                    Call NewDoc()
                End If

                Me.PNSearchDocNo.Visible = False
                Me.TBDocNo.Focus()
                Me.TBDocNo.SelectAll()
            Else
                MsgBox("ไม่พบข้อมูลเอกสารที่ต้องการ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                If Me.ListViewSearchDocNo.Items.Count > 0 Then
                    Me.ListViewSearchDocNo.Focus()
                    Me.ListViewSearchDocNo.Items(0).Focused = True
                    Me.ListViewSearchDocNo.Items(0).Selected = True
                End If
            End If
        End If
    End Sub

    Private Sub ListViewSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearchDocNo.KeyDown
        Dim vDocNo As String
        Dim i As Integer
        Dim vQty As Double
        Dim vRetailCom As Double
        Dim vWholeSaleCom As Double

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewSearchDocNo.Items.Count > 0 Then
                vDocNo = Me.ListViewSearchDocNo.SelectedItems(0).SubItems(1).Text

                Call ClearDataDGV()
                vQuery = "exec dbo.USP_COM_RequestSearch2 '" & vDocNo & "'"
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "Search")
                dt = ds.Tables("Search")
                If dt.Rows.Count > 0 Then
                    vIsOpen = 1
                    vMemIsConfirm = dt.Rows(i).Item("isconfirm")
                    vMemIsCancel = dt.Rows(i).Item("iscancel")
                    Me.TBDocNo.Text = dt.Rows(i).Item("docno")
                    Me.DTPDocDate.Text = dt.Rows(i).Item("docdate")
                    Me.TBCampaignNo.Text = dt.Rows(i).Item("campaigncode")
                    Me.LBLCampaignName.Text = dt.Rows(i).Item("campaignname")
                    Me.TBMyDescription.Text = dt.Rows(i).Item("mydesc")

                    For i = 0 To dt.Rows.Count - 1

                        vQty = dt.Rows(i).Item("conditionqty")
                        vRetailCom = dt.Rows(i).Item("retailcom")
                        vWholeSaleCom = dt.Rows(i).Item("wholesalecom")

                        Me.DGVItemDetails.Item(1, i).Value = dt.Rows(i).Item("itemcode")
                        Me.DGVItemDetails.Item(2, i).Value = dt.Rows(i).Item("itemname")
                        Me.DGVItemDetails.Item(3, i).Value = dt.Rows(i).Item("unitcode")
                        Me.DGVItemDetails.Item(4, i).Value = Format(vQty, "##,##0.00")
                        Me.DGVItemDetails.Item(5, i).Value = Format(vRetailCom, "##,##0.00")
                        Me.DGVItemDetails.Item(6, i).Value = Format(vWholeSaleCom, "##,##0.00")
                    Next

                    If vMemIsConfirm = 1 Then
                        Call ConfirmDoc()
                    End If

                    If vMemIsCancel = 1 Then
                        Call CancelDoc()
                    End If

                    If vMemIsCancel = 0 And vMemIsConfirm = 0 Then
                        Call NewDoc()
                    End If

                    Me.PNSearchDocNo.Visible = False
                    Me.TBDocNo.Focus()
                    Me.TBDocNo.SelectAll()
                Else
                    MsgBox("ไม่พบข้อมูลเอกสารที่ต้องการ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    If Me.ListViewSearchDocNo.Items.Count > 0 Then
                        Me.ListViewSearchDocNo.Focus()
                        Me.ListViewSearchDocNo.Items(0).Focused = True
                        Me.ListViewSearchDocNo.Items(0).Selected = True
                    End If
                End If
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchDocNo.Visible = False
            Me.TBDocNo.Focus()
            Me.TBDocNo.SelectAll()
        End If

        Dim vCheckLine As Integer

        If Me.ListViewSearchDocNo.Items.Count > 0 Then
            If e.KeyCode = Keys.Up Then
                vCheckLine = Me.ListViewSearchDocNo.SelectedItems(0).Index
                If vCheckLine = 0 Then
                    Me.TBSearchDocNo.Focus()
                    Me.TBSearchDocNo.SelectAll()
                End If
            End If
        End If
    End Sub

    Private Sub ListViewSearchDocNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewSearchDocNo.SelectedIndexChanged

    End Sub

    Private Sub BTNClearScreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearScreen.Click

        Call ClearScreen()
        Call vGenDocNoAuto()

        'For i = 0 To Me.DGVItemDetails.Rows.Count - 1
        '    Me.DGVItemDetails.Rows.RemoveAt(i)
        'Next
        'Call vGetBeginData()
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Me.Close()
    End Sub

    Private Sub BTNPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNPrint.Click
        If vIsOpen = 1 Then
            vMemReqCommNo = Me.TBDocNo.Text

            If vMemIsConfirm = 1 Then
                MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNDocNo.Focus()
                Exit Sub
            End If

            If vMemIsCancel = 1 Then
                MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNDocNo.Focus()
                Exit Sub
            End If

            If frmReportReqItemComm Is Nothing Then
                frmReportReqItemComm = New FormReportReqItemComm
            Else
                If frmReportReqItemComm.IsDisposed Then
                    frmReportReqItemComm = New FormReportReqItemComm
                End If
            End If

            frmReportReqItemComm.Show()
            frmReportReqItemComm.BringToFront()
        Else
            MsgBox("เอกสารยังไม่ได้บันทึกไม่สามารถพิมพ์ได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBDocNo.Focus()
        End If
    End Sub


    Public Sub PrintDocument()
        If vIsOpen = 1 Then
            vMemReqCommNo = Me.TBDocNo.Text

            If vMemIsConfirm = 1 Then
                MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNDocNo.Focus()
                Exit Sub
            End If

            If vMemIsCancel = 1 Then
                MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNDocNo.Focus()
                Exit Sub
            End If

            If frmReportReqItemComm Is Nothing Then
                frmReportReqItemComm = New FormReportReqItemComm
            Else
                If frmReportReqItemComm.IsDisposed Then
                    frmReportReqItemComm = New FormReportReqItemComm
                End If
            End If

            frmReportReqItemComm.Show()
            frmReportReqItemComm.BringToFront()
        End If
    End Sub

    Private Sub BTNCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNCancel.Click
        If vIsOpen = 1 Then
            If vMemIsConfirm = 1 Then
                MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNDocNo.Focus()
                Exit Sub
            End If

            If vMemIsCancel = 1 Then
                MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNDocNo.Focus()
                Exit Sub
            End If

            MsgBox("เมนูยกเลิก ยังไม่ได้เปิดใช้งาน", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBDocNo.Focus()
        End If
    End Sub

    Private Sub TBDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBDocNo.KeyDown, BTNDocNo.KeyDown, DTPDocDate.KeyDown, BTNCampaign.KeyDown, TBCampaignNo.KeyDown, TBMyDescription.KeyDown, DGVItemDetails.KeyDown, BTNClearScreen.KeyDown, BTNSave.KeyDown, BTNSearch.KeyDown, BTNCancel.KeyDown, BTNPrint.KeyDown, BTNExit.KeyDown
        If e.KeyCode = Keys.F1 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = Keys.F4 Then
            Call ClearScreen()
            Call vGenDocNoAuto()
        End If

        If e.KeyCode = Keys.F5 Then
            Call SaveData()
        End If

        If e.KeyCode = Keys.F9 Then
            Call PrintDocument()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub TBSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchDocNo.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call SearchDocNo()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchDocNo.Visible = False
            Me.TBDocNo.Focus()
            Me.TBDocNo.SelectAll()
        End If
    End Sub

    Private Sub TBSearchDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchDocNo.TextChanged

    End Sub

    Private Sub BTNSearchDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchDocNo.Click

        Call SearchDocNo()

    End Sub

    Private Sub BTNSelectSearchDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectSearchDocNo.Click
        Dim vDocNo As String
        Dim i As Integer
        Dim vQty As Double
        Dim vRetailCom As Double
        Dim vWholeSaleCom As Double

        On Error Resume Next

        If Me.ListViewSearchDocNo.Items.Count > 0 Then
            vDocNo = Me.ListViewSearchDocNo.SelectedItems(0).SubItems(1).Text

            Call ClearDataDGV()
            vQuery = "exec dbo.USP_COM_RequestSearch2 '" & vDocNo & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Search")
            dt = ds.Tables("Search")
            If dt.Rows.Count > 0 Then
                vIsOpen = 1
                vMemIsConfirm = dt.Rows(i).Item("isconfirm")
                vMemIsCancel = dt.Rows(i).Item("iscancel")
                Me.TBDocNo.Text = dt.Rows(i).Item("docno")
                Me.DTPDocDate.Text = dt.Rows(i).Item("docdate")
                Me.TBCampaignNo.Text = dt.Rows(i).Item("campaigncode")
                Me.LBLCampaignName.Text = dt.Rows(i).Item("campaignname")
                Me.TBMyDescription.Text = dt.Rows(i).Item("mydesc")

                For i = 0 To dt.Rows.Count - 1

                    vQty = dt.Rows(i).Item("conditionqty")
                    vRetailCom = dt.Rows(i).Item("retailcom")
                    vWholeSaleCom = dt.Rows(i).Item("wholesalecom")

                    Me.DGVItemDetails.Item(1, i).Value = dt.Rows(i).Item("itemcode")
                    Me.DGVItemDetails.Item(2, i).Value = dt.Rows(i).Item("itemname")
                    Me.DGVItemDetails.Item(3, i).Value = dt.Rows(i).Item("unitcode")
                    Me.DGVItemDetails.Item(4, i).Value = Format(vQty, "##,##0.00")
                    Me.DGVItemDetails.Item(5, i).Value = Format(vRetailCom, "##,##0.00")
                    Me.DGVItemDetails.Item(6, i).Value = Format(vWholeSaleCom, "##,##0.00")
                Next

                Me.PNSearchDocNo.Visible = False
                Me.TBDocNo.Focus()
                Me.TBDocNo.SelectAll()
            Else
                MsgBox("ไม่พบข้อมูลเอกสารที่ต้องการ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                If Me.ListViewSearchDocNo.Items.Count > 0 Then
                    Me.ListViewSearchDocNo.Focus()
                    Me.ListViewSearchDocNo.Items(0).Focused = True
                    Me.ListViewSearchDocNo.Items(0).Selected = True
                End If
            End If
        End If
    End Sub

    Private Sub TBSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call SearchCampaign()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSearch.Visible = False
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub TBSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearch.TextChanged

    End Sub

    Private Sub BTNClickSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClickSearch.Click
        Call SearchCampaign()
    End Sub

    Private Sub BTNSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelect.Click
        Dim vIndex As Integer

        On Error Resume Next

        If Me.ListViewSearch.Items.Count > 0 Then
            vIndex = Me.ListViewSearch.SelectedItems(0).Index
            Me.TBCampaignNo.Text = Me.ListViewSearch.Items(vIndex).SubItems(1).Text
            Me.LBLCampaignName.Text = Me.ListViewSearch.Items(vIndex).SubItems(2).Text

            Me.PNSearch.Visible = False
            Me.TBMyDescription.Focus()
            Me.TBMyDescription.SelectAll()
        End If
    End Sub

    Private Sub BTNCloseSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCloseSearchDocNo.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.PNSearch.Visible = False
            Me.TBDocNo.Focus()
            Me.TBDocNo.SelectAll()
        End If
    End Sub

    Private Sub BTNSelectSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSelectSearchDocNo.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.PNSearchDocNo.Visible = False
            Me.TBDocNo.Focus()
            Me.TBDocNo.SelectAll()
        End If
    End Sub

    Private Sub PNSearchDocNo_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PNSearchDocNo.Paint

    End Sub

    Private Sub BTNSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSearchDocNo.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.PNSearchDocNo.Visible = False
            Me.TBDocNo.Focus()
            Me.TBDocNo.SelectAll()
        End If
    End Sub

    Private Sub BTNClose_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClose.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.PNSearch.Visible = False
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub BTNSelect_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSelect.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.PNSearch.Visible = False
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub BTNClickSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClickSearch.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.PNSearch.Visible = False
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub MenuDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuDelete.Click
        Dim vRowID As Integer
        Dim vColumnID As Integer
        Dim vCellID As Integer
        Dim i As Integer
        Dim n As Integer
        Dim vItemCode As String
        Dim vAnswer As Integer

        On Error Resume Next

        If Me.DGVItemDetails.RowCount > 0 Then

            If vMemIsConfirm = 1 Then
                MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNDocNo.Focus()
                Exit Sub
            End If

            If vMemIsCancel = 1 Then
                MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNDocNo.Focus()
                Exit Sub
            End If


            vRowID = vMemRow
            vItemCode = Me.DGVItemDetails.Item(1, vRowID).Value
            If vItemCode <> "" Then

                vAnswer = MsgBox("คุณต้องการรายการที่ " & vRowID + 1 & " ใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
                If vAnswer = 6 Then
                    Me.DGVItemDetails.Rows.RemoveAt(vRowID)

                    n = 1
                    For i = 0 To Me.DGVItemDetails.Rows.Count - 1
                        Me.DGVItemDetails.Item(0, i).Value = n
                        n = n + 1
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub DGVItemDetails_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItemDetails.CellContentClick

    End Sub

    Private Sub DGVItemDetails_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItemDetails.CellClick
        vMemRow = Me.DGVItemDetails.CurrentCell.RowIndex
    End Sub
End Class