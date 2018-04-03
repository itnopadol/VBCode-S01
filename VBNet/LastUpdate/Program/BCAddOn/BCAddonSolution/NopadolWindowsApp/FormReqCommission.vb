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

    Dim ds1 As DataSet
    Dim da1 As SqlDataAdapter
    Dim dt1 As DataTable

    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vIsOpen As Integer
    Dim vMemIsCancel As Integer
    Dim vMemIsConfirm As Integer

    Dim vIsNumber As Integer

    Dim vMemColumn As Integer
    Dim vMemRow As Integer

    Dim vMemStartDate As Date

    Private Sub FormReqCommission_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()
        Call SearchItemBrand()
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

        Me.DGVItemDetails.Rows.Add(9999)
        For i = 0 To 9999 - 1
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
        Dim vCampaignCode As String
        Dim vDocNo As String
        Dim vItemCode As String
        Dim vUnitCode As String
        Dim vColumn As Integer
        Dim vRow As Integer
        Dim i As Integer

        Dim vCheckItemCode As String
        Dim vMemCountCheck As Integer
        Dim vCheckItemDup As Integer

        Dim vCheckNoDup As String
        Dim vCheckCampaign As String
        Dim vCheckCampaignName As String

        Dim vDateDiff As Integer
        Dim vNowDate As Date
        Dim vPromoPrice As Double
        Dim vBeginDate As String
        Dim vEndDate As String


        On Error Resume Next

        If Me.TBCampaignNo.Text = "" Then
            MsgBox("กรุณา กรอกรหัสแคมเปญ ก่อนเลือกสินค้า", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBCampaignNo.Focus()
            Exit Sub
        End If

        If Me.TBDocNo.Text = "" Then
            MsgBox("กรุณา กรอกเลขที่เอกสาร ก่อนเลือกสินค้า", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBDocNo.Focus()
            Exit Sub
        End If


        vNowDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        vDateDiff = vb6.DateDiff(DateInterval.Day, vMemStartDate, vNowDate)

        'If vDateDiff > -7 Then
        '    MsgBox("ไม่สามารถเพิ่มรายการสินค้าเสนอค่าคอมได้ในแคมเปญนี้ เพราะต้องทำก่อนแคมเปญจะเริ่ม 7 วัน กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")
        '    Me.TBCampaignNo.Focus()
        '    Exit Sub
        'End If


        vDocNo = Me.TBDocNo.Text
        vCampaignCode = Me.TBCampaignNo.Text

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

                    vUnitCode = dt.Rows(0).Item("unitcode")

                    vQuery = "exec dbo.USP_COM_RequestDupCK '" & vDocNo & "','" & vCampaignCode & "','" & vItemCode & "','" & vUnitCode & "'"
                    da1 = New SqlDataAdapter(vQuery, vConnection)
                    ds1 = New DataSet
                    da1.Fill(ds1, "CheckItemDup")
                    dt1 = ds1.Tables("CheckItemDup")
                    If dt1.Rows.Count > 0 Then
                        vCheckItemDup = dt1.Rows(0).Item("duplicateItem")
                        vCheckNoDup = dt1.Rows(0).Item("requestno_dup")
                        vCheckCampaign = dt1.Rows(0).Item("campaigncode_dup")
                        vCheckCampaignName = dt1.Rows(0).Item("campaignname_dup")
                    End If

                    If vCheckItemDup > 0 Then
                        MsgBox("สินค้าซ้ำ ในแคมเปญ " & vCheckCampaign & "/" & vCheckCampaignName & " และเลขที่ " & vCheckNoDup & " ไม่สามารถเพิ่มได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                        Exit Sub
                    End If

                    Me.DGVItemDetails.Item(2, vRow).Value = dt.Rows(0).Item("itemname")
                    Me.DGVItemDetails.Item(3, vRow).Value = dt.Rows(0).Item("unitcode")
                    Me.DGVItemDetails.Item(4, vRow).Value = 1
                    Me.DGVItemDetails.Item(5, vRow).Value = ""
                    Me.DGVItemDetails.Item(6, vRow).Value = ""

                    vPromoPrice = dt.Rows(0).Item("PromoPrice")
                    vBeginDate = dt.Rows(0).Item("probegindate")
                    vEndDate = dt.Rows(0).Item("proenddate")

                    If vPromoPrice > 0 Then
                        Me.DGVItemDetails.Item(7, vRow).Value = Format(vPromoPrice, "##,##0.00")
                        Me.DGVItemDetails.Item(8, vRow).Value = dt.Rows(0).Item("proname")
                        Me.DGVItemDetails.Item(9, vRow).Value = vBeginDate
                        Me.DGVItemDetails.Item(10, vRow).Value = vEndDate
                    End If
                Else
                    MsgBox("สินค้า รหัส " & vItemCode & " ไม่มีข้อมูลในระบบ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")

                    Me.DGVItemDetails.Item(1, vRow).Value = ""
                    Me.DGVItemDetails.Item(2, vRow).Value = ""
                    Me.DGVItemDetails.Item(3, vRow).Value = ""
                    Me.DGVItemDetails.Item(4, vRow).Value = ""
                    Me.DGVItemDetails.Item(5, vRow).Value = ""
                    Me.DGVItemDetails.Item(6, vRow).Value = ""
                    Me.DGVItemDetails.Item(7, vRow).Value = ""
                    Me.DGVItemDetails.Item(8, vRow).Value = ""
                    Me.DGVItemDetails.Item(9, vRow).Value = ""
                    Me.DGVItemDetails.Item(10, vRow).Value = ""
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
            MsgBox(Err.Description & "ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Send Error Message")
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
            vMemStartDate = Me.ListViewSearch.Items(vIndex).SubItems(3).Text

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
                vMemStartDate = Me.ListViewSearch.Items(vIndex).SubItems(3).Text

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
        Dim vCommCash As String
        Dim vCommCredit As String
        Dim vMemItemCount As Integer
        Dim vMemItemNotKeyComm As Integer

        Dim vMemBeginTran As Integer
        Dim vTypeInsert As Integer
        Dim vCharStr As String

        Dim vCheckQty As Double
        Dim vCash As String
        Dim vCredit As String

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

            If vItemCode <> "" Then 'and Me.DGVItemDetails.Item(4, i).Value <> "0" And Me.DGVItemDetails.Item(4, i).Value <> "" Then
                vQty = Me.DGVItemDetails.Item(4, i).Value
            End If

            'MsgBox(Me.DGVItemDetails.Item(5, i).Value)
            If vItemCode <> "" Then
                vCash = Me.DGVItemDetails.Item(5, i).Value
            ElseIf Me.DGVItemDetails.Item(5, i).Value = "" And vItemCode <> "" Then
                MsgBox("ยังไม่ได้กำหนดค่าคอมฯ ขายสด", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If

            If vItemCode <> "" Then
                vCredit = Me.DGVItemDetails.Item(6, i).Value
            ElseIf Me.DGVItemDetails.Item(6, i).Value = "" And vItemCode <> "" Then
                MsgBox("ยังไม่ได้กำหนดค่าคอมฯ ขายเชื่อ", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If

            vCharStr = Me.DGVItemDetails.Item(4, i).Value
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                MsgBox("รายการที่ " & i + 1 & " ในช่องจำนวนที่กำหนด กรอกตัวเลขได้เท่านั้น", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(4, i).Selected = True
                Exit Sub
            End If

            vCharStr = Me.DGVItemDetails.Item(5, i).Value
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                MsgBox("รายการที่ " & i + 1 & " ในช่องค่าคอมฯขายสด กรอกตัวเลขได้เท่านั้น", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(5, i).Selected = True
                Exit Sub
            End If

            vCharStr = Me.DGVItemDetails.Item(6, i).Value
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                MsgBox("รายการที่ " & i + 1 & " ในช่องค่าคอมฯขายเชื่อ กรอกตัวเลขได้เท่านั้น", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(6, i).Selected = True
                Exit Sub
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

            Call vGenDocNoAuto()

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

                If vItemCode <> "" Then
                    vQty = Me.DGVItemDetails.Item(4, i).Value
                    vCommCash = Me.DGVItemDetails.Item(5, i).Value
                    vCommCredit = Me.DGVItemDetails.Item(6, i).Value
                End If

                If vItemCode <> "" And vItemName <> "" And vQty <> 0 And vUnitCode <> "" Then
                    vQuery = "exec dbo.USP_COM_RequestSubSave '" & vDocNo & "'," & i & ",'" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vQty & ",'" & vCommCash & "','" & vCommCredit & "' "
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
            'MsgBox("ไม่สามารถบันทึกเอกสารได้ เกิดปัญหา จาก Store : dbo.USP_COM_CampaignSave  ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Error")
            MsgBox(Err.Description & "ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Send Error Message")
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

            Call vGenDocNoAuto()

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
            'MsgBox("ไม่สามารถบันทึกเอกสารได้ เกิดปัญหา จาก Store : dbo.USP_COM_CampaignSave  ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Error")
            MsgBox(Err.Description & "ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Send Error Message")
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
                vListItem.SubItems.Add(5).Text = dt.Rows(i).Item("mydesc")
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
            Me.DGVItemDetails.Item(7, i).Value = ""
            Me.DGVItemDetails.Item(8, i).Value = ""
            Me.DGVItemDetails.Item(9, i).Value = ""
            Me.DGVItemDetails.Item(10, i).Value = ""
        Next
    End Sub

    Private Sub ListViewSearchDocNo_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewSearchDocNo.DoubleClick
        Dim vDocNo As String
        Dim i As Integer
        Dim vQty As Double
        Dim vRetailCom As Double
        Dim vWholeSaleCom As Double

        Dim vPromoPrice As Double
        Dim vBeginDate As String
        Dim vEndDate As String

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

                    vPromoPrice = dt.Rows(i).Item("promoprice")
                    vBeginDate = dt.Rows(i).Item("probegindate")
                    vEndDate = dt.Rows(i).Item("proenddate")

                    Me.DGVItemDetails.Item(1, i).Value = dt.Rows(i).Item("itemcode")
                    Me.DGVItemDetails.Item(2, i).Value = dt.Rows(i).Item("itemname")
                    Me.DGVItemDetails.Item(3, i).Value = dt.Rows(i).Item("unitcode")
                    Me.DGVItemDetails.Item(4, i).Value = Format(vQty, "##,##0.00")
                    Me.DGVItemDetails.Item(5, i).Value = Format(vRetailCom, "##,##0.00")
                    Me.DGVItemDetails.Item(6, i).Value = Format(vWholeSaleCom, "##,##0.00")
                    Me.DGVItemDetails.Item(7, i).Value = Format(vPromoPrice, "##,##0.00")
                    Me.DGVItemDetails.Item(8, i).Value = dt.Rows(i).Item("proname")
                    Me.DGVItemDetails.Item(9, i).Value = vBeginDate
                    Me.DGVItemDetails.Item(10, i).Value = vEndDate
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

        Dim vPromoPrice As Double
        Dim vBeginDate As String
        Dim vEndDate As String

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

                        vPromoPrice = dt.Rows(i).Item("promoprice")
                        vBeginDate = dt.Rows(i).Item("probegindate")
                        vEndDate = dt.Rows(i).Item("proenddate")

                        Me.DGVItemDetails.Item(1, i).Value = dt.Rows(i).Item("itemcode")
                        Me.DGVItemDetails.Item(2, i).Value = dt.Rows(i).Item("itemname")
                        Me.DGVItemDetails.Item(3, i).Value = dt.Rows(i).Item("unitcode")
                        Me.DGVItemDetails.Item(4, i).Value = Format(vQty, "##,##0.00")
                        Me.DGVItemDetails.Item(5, i).Value = Format(vRetailCom, "##,##0.00")
                        Me.DGVItemDetails.Item(6, i).Value = Format(vWholeSaleCom, "##,##0.00")
                        Me.DGVItemDetails.Item(7, i).Value = Format(vPromoPrice, "##,##0.00")
                        Me.DGVItemDetails.Item(8, i).Value = dt.Rows(i).Item("proname")
                        Me.DGVItemDetails.Item(9, i).Value = vBeginDate
                        Me.DGVItemDetails.Item(10, i).Value = vEndDate
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
        On Error Resume Next

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
        On Error Resume Next

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
        Dim vDocNo As String
        Dim vAnswer As Integer


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

            vDocNo = Me.TBDocNo.Text
            vAnswer = MsgBox("คุณต้องการยกเลิกเอกสารนี้ใช่หรือไม่ ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 Then

                vQuery = "begin tran"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                vQuery = "exec dbo.USP_Com_CancelRequest '" & vDocNo & "','" & vUserID & "'"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                vQuery = "commit tran"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                MsgBox("ยกเลิกเอกสาร เลขที่ " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")

                Call ClearScreen()
                Call vGenDocNoAuto()

            End If
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

        Dim vPromoPrice As Double
        Dim vBeginDate As String
        Dim vEndDate As String

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

                    vPromoPrice = dt.Rows(i).Item("promoprice")
                    vBeginDate = dt.Rows(i).Item("probegindate")
                    vEndDate = dt.Rows(i).Item("proenddate")

                    Me.DGVItemDetails.Item(1, i).Value = dt.Rows(i).Item("itemcode")
                    Me.DGVItemDetails.Item(2, i).Value = dt.Rows(i).Item("itemname")
                    Me.DGVItemDetails.Item(3, i).Value = dt.Rows(i).Item("unitcode")
                    Me.DGVItemDetails.Item(4, i).Value = Format(vQty, "##,##0.00")
                    Me.DGVItemDetails.Item(5, i).Value = Format(vRetailCom, "##,##0.00")
                    Me.DGVItemDetails.Item(6, i).Value = Format(vWholeSaleCom, "##,##0.00")
                    Me.DGVItemDetails.Item(7, i).Value = Format(vPromoPrice, "##,##0.00")
                    Me.DGVItemDetails.Item(8, i).Value = dt.Rows(i).Item("proname")
                    Me.DGVItemDetails.Item(9, i).Value = vBeginDate
                    Me.DGVItemDetails.Item(10, i).Value = vEndDate

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
            vMemStartDate = Me.ListViewSearch.Items(vIndex).SubItems(3).Text

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

    Private Sub DGVItemDetails_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItemDetails.CellClick
        vMemRow = Me.DGVItemDetails.CurrentCell.RowIndex
    End Sub

    Private Sub BTNSearchListOldData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchListOldData.Click
        Dim vDocNo As String
        Dim i As Integer
        Dim n As Integer
        Dim vListItem As ListViewItem
        Dim vQty As Double
        Dim vRetailCom As Double
        Dim vWholeSaleCom As Double
        Dim vPromoPrice As Double
        Dim vBeginDate As String
        Dim vEndDate As String

        On Error Resume Next


        Me.ListViewOldData.Items.Clear()

        If Me.TBSearchOldData.Text <> "" Then

            vDocNo = Me.TBSearchOldData.Text
            vQuery = "exec dbo.USP_COM_SearchReqCommOldData '" & vDocNo & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Search")
            dt = ds.Tables("Search")
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    n = n + 1
                    vListItem = Me.ListViewOldData.Items.Add(n)
                    vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                    vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("itemcode")
                    vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("itemname")
                    vQty = dt.Rows(i).Item("conditionqty")
                    vRetailCom = dt.Rows(i).Item("retailcom")
                    vWholeSaleCom = dt.Rows(i).Item("wholesalecom")

                    vPromoPrice = dt.Rows(i).Item("promoprice")
                    vBeginDate = dt.Rows(i).Item("probegindate")
                    vEndDate = dt.Rows(i).Item("proenddate")


                    vListItem.SubItems.Add(3).Text = Format(vQty, "##,##0.00")
                    vListItem.SubItems.Add(4).Text = dt.Rows(i).Item("unitcode")
                    vListItem.SubItems.Add(5).Text = Format(vRetailCom, "##,##0.00")
                    vListItem.SubItems.Add(6).Text = Format(vWholeSaleCom, "##,##0.00")
                    vListItem.SubItems.Add(7).Text = Format(vPromoPrice, "##,##0.00")
                    vListItem.SubItems.Add(8).Text = dt.Rows(i).Item("proname")
                    vListItem.SubItems.Add(9).Text = vBeginDate
                    vListItem.SubItems.Add(10).Text = vEndDate
                Next
            End If

        End If
    End Sub

    Private Sub BTNSearchOldData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchOldData.Click
        If Me.TBCampaignNo.Text <> "" Then
            Me.PNOldDocNo.Visible = True
            Me.TBSearchOldData.Focus()
        Else
            MsgBox("กรุณาเลือก แคมเปญคอมมิชชั่นก่อนใช้ข้อมูลเก่า กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBCampaignNo.Focus()
        End If
    End Sub

    Private Sub BTNCloseOldData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseOldData.Click
        Me.PNOldDocNo.Visible = False
        Me.TBDocNo.Focus()
    End Sub

    Private Sub BTNSelectOldData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectOldData.Click
        Dim i As Integer
        Dim n As Integer
        Dim vCountSelect As Integer
        Dim vCampaignCode As String
        Dim vItemCode As String
        Dim vUnitCode As String
        Dim vCheckUnitCode As String
        Dim vItemName As String
        Dim vConditionQty As Double
        Dim vRetailComm As Double
        Dim vWholeSaleComm As Double
        Dim vCheckItemCode As String
        Dim vMemCountCheck As Integer
        Dim vCheckItemDup As String
        Dim vCheckNoDup As String
        Dim vCheckCampaign As String
        Dim vCheckCampaignName As String
        Dim vCheckLineAdd As Integer

        Dim vNowDate As Date
        Dim vDateDiff As Integer

        Dim vPromoPrice As Double
        Dim vProName As String
        Dim vProBeginDate As String
        Dim vProEndDate As String

        On Error Resume Next

        vCountSelect = 0
        vCampaignCode = Me.TBCampaignNo.Text


        vNowDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        vDateDiff = vb6.DateDiff(DateInterval.Day, vMemStartDate, vNowDate)

        'If vDateDiff > -7 Then
        '    MsgBox("ไม่สามารถเพิ่มรายการสินค้าเสนอค่าคอมได้ในแคมเปญนี้ เพราะต้องทำก่อนแคมเปญจะเริ่ม 7 วัน กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")
        '    Me.TBCampaignNo.Focus()
        '    Exit Sub
        'End If


        If Me.ListViewOldData.Items.Count > 0 Then
            For i = 0 To Me.ListViewOldData.Items.Count - 1
                If Me.ListViewOldData.Items(i).Checked = True Then
                    vCountSelect = vCountSelect + 1
                End If
            Next

            If vCountSelect = 0 Then
                MsgBox("กรุณาเลือกรายการที่ต้องการใช้งาน", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBSearchOldData.Focus()
                Exit Sub
            End If

            For n = 0 To Me.ListViewOldData.Items.Count - 1
                If Me.ListViewOldData.Items(n).Checked = True Then
                    vItemCode = Me.ListViewOldData.Items(n).SubItems(2).Text
                    vUnitCode = Me.ListViewOldData.Items(n).SubItems(5).Text
                    vItemName = Me.ListViewOldData.Items(n).SubItems(3).Text
                    vConditionQty = Me.ListViewOldData.Items(n).SubItems(4).Text
                    vRetailComm = Me.ListViewOldData.Items(n).SubItems(6).Text
                    vWholeSaleComm = Me.ListViewOldData.Items(n).SubItems(7).Text
                    vPromoPrice = Me.ListViewOldData.Items(n).SubItems(8).Text
                    vProName = Me.ListViewOldData.Items(n).SubItems(9).Text
                    vProBeginDate = Me.ListViewOldData.Items(n).SubItems(10).Text
                    vProEndDate = Me.ListViewOldData.Items(n).SubItems(11).Text


                    For i = 0 To Me.DGVItemDetails.Rows.Count - 1
                        vCheckItemCode = Me.DGVItemDetails.Item(1, i).Value
                        vCheckUnitCode = Me.DGVItemDetails.Item(3, i).Value

                        vCheckLineAdd = i
                        If vCheckItemCode = Nothing Then
                            GoTo Line1
                        End If
                        If vCheckItemCode = vItemCode And vUnitCode = vCheckUnitCode Then
                            vMemCountCheck = vMemCountCheck + 1
                            GoTo Line1
                        End If
                    Next
Line1:
                    If vMemCountCheck >= 1 Then
                        MsgBox("สินค้า รหัส " & vItemCode & " มีอยู่แล้วในรายการเสนอขอคิดค่าคอมฯ บรรทัดที่ " & vCheckLineAdd + 1 & " กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                        Exit Sub
                    End If

                    vQuery = "exec dbo.USP_COM_RequestDupCK '','" & vCampaignCode & "','" & vItemCode & "','" & vUnitCode & "'"
                    da1 = New SqlDataAdapter(vQuery, vConnection)
                    ds1 = New DataSet
                    da1.Fill(ds1, "CheckItemDup")
                    dt1 = ds1.Tables("CheckItemDup")
                    If dt1.Rows.Count > 0 Then
                        vCheckItemDup = dt1.Rows(0).Item("duplicateItem")
                        vCheckNoDup = dt1.Rows(0).Item("requestno_dup")
                        vCheckCampaign = dt1.Rows(0).Item("campaigncode_dup")
                        vCheckCampaignName = dt1.Rows(0).Item("campaignname_dup")
                    End If

                    If vCheckItemDup > 0 Then
                        MsgBox("สินค้า รหัส " & vItemCode & " ซ้ำกับแคมเปญ " & vCheckCampaign & "/" & vCheckCampaignName & " และเลขที่ " & vCheckNoDup & " ไม่สามารถเพิ่มได้ อยู่ในบรรทัดที่ " & n + 1 & " กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                        Exit Sub
                    End If

                    Me.DGVItemDetails.Item(1, vCheckLineAdd).Value = vItemCode
                    Me.DGVItemDetails.Item(2, vCheckLineAdd).Value = vItemName
                    Me.DGVItemDetails.Item(3, vCheckLineAdd).Value = vUnitCode
                    Me.DGVItemDetails.Item(4, vCheckLineAdd).Value = Format(vConditionQty, "##,##0.00")
                    Me.DGVItemDetails.Item(5, vCheckLineAdd).Value = Format(vRetailComm, "##,##0.00")
                    Me.DGVItemDetails.Item(6, vCheckLineAdd).Value = Format(vWholeSaleComm, "##,##0.00")
                    If vPromoPrice > 0 Then
                        Me.DGVItemDetails.Item(7, vCheckLineAdd).Value = Format(vPromoPrice, "##,##0.00")
                        Me.DGVItemDetails.Item(8, vCheckLineAdd).Value = vProName
                        Me.DGVItemDetails.Item(9, vCheckLineAdd).Value = vProBeginDate
                        Me.DGVItemDetails.Item(10, vCheckLineAdd).Value = vProEndDate
                    Else
                        Me.DGVItemDetails.Item(7, vCheckLineAdd).Value = ""
                        Me.DGVItemDetails.Item(8, vCheckLineAdd).Value = ""
                        Me.DGVItemDetails.Item(9, vCheckLineAdd).Value = ""
                        Me.DGVItemDetails.Item(10, vCheckLineAdd).Value = ""
                    End If

                End If
            Next
        End If

        Me.PNOldDocNo.Visible = False
        Me.TBDocNo.Focus()

    End Sub

    Private Sub CBSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBSelectAll.CheckedChanged
        Dim i As Integer

        On Error Resume Next

        If Me.ListViewOldData.Items.Count > 0 Then
            If Me.CBSelectAll.Checked = True Then
                For i = 0 To Me.ListViewOldData.Items.Count - 1
                    Me.ListViewOldData.Items(i).Checked = True
                Next
            End If
            If Me.CBSelectAll.Checked = False Then
                For i = 0 To Me.ListViewOldData.Items.Count - 1
                    Me.ListViewOldData.Items(i).Checked = False
                Next
            End If
        End If

    End Sub

    Private Sub DGVItemDetails_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItemDetails.CellContentClick

    End Sub

    Private Sub TBSearchOldData_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchOldData.KeyDown
        Dim vDocNo As String
        Dim i As Integer
        Dim n As Integer
        Dim vListItem As ListViewItem
        Dim vQty As Double
        Dim vRetailCom As Double
        Dim vWholeSaleCom As Double

        Dim vPromoPrice As Double
        Dim vBeginDate As String
        Dim vEndDate As String


        'On Error Resume Next

        If e.KeyCode = Keys.Enter Then

            Me.ListViewOldData.Items.Clear()

            If Me.TBSearchOldData.Text <> "" Then

                vDocNo = Me.TBSearchOldData.Text
                vQuery = "exec dbo.USP_COM_SearchReqCommOldData '" & vDocNo & "'"
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "Search")
                dt = ds.Tables("Search")
                If dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        n = n + 1
                        vListItem = Me.ListViewOldData.Items.Add(n)
                        vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                        vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("itemcode")
                        vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("itemname")
                        vQty = dt.Rows(i).Item("conditionqty")
                        vRetailCom = dt.Rows(i).Item("retailcom")
                        vWholeSaleCom = dt.Rows(i).Item("wholesalecom")

                        vPromoPrice = dt.Rows(i).Item("promoprice")
                        vBeginDate = dt.Rows(i).Item("probegindate")
                        vEndDate = dt.Rows(i).Item("proenddate")

                        vListItem.SubItems.Add(3).Text = Format(vQty, "##,##0.00")
                        vListItem.SubItems.Add(4).Text = dt.Rows(i).Item("unitcode")
                        vListItem.SubItems.Add(5).Text = Format(vRetailCom, "##,##0.00")
                        vListItem.SubItems.Add(6).Text = Format(vWholeSaleCom, "##,##0.00")
                        vListItem.SubItems.Add(7).Text = Format(vPromoPrice, "##,##0.00")
                        vListItem.SubItems.Add(8).Text = dt.Rows(i).Item("proname")
                        vListItem.SubItems.Add(9).Text = vBeginDate
                        vListItem.SubItems.Add(10).Text = vEndDate
                    Next
                End If

            End If
        End If
    End Sub

    Private Sub TBSearchOldData_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBSearchOldData.KeyPress

    End Sub

    Private Sub TBSearchOldData_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchOldData.TextChanged

    End Sub

    Private Sub BTNSearchItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchItem.Click
        Call SearchItem()
    End Sub

    Public Sub SearchItem()
        Dim vSearch As String
        Dim vType As Integer
        Dim vBrandCode As String
        Dim vListItem As ListViewItem
        Dim i As Integer
        Dim n As Integer

        Dim vCashSalePrice As Double
        Dim vCreditSalePrice As Double

        On Error Resume Next

        'If Me.TBSearch.Text = "" Then
        '    MsgBox("กรุณา กรอกรหัสหรือชื่อสินค้าที่ต้องการค้นหา", MsgBoxStyle.Critical, "Send Information Message")
        '    Me.TBSearch.Focus()
        '    Exit Sub
        'End If

        If Me.CBNotAddPriceStructure.Checked = True Then
            vType = 1
        ElseIf Me.CBItemSaleLose.Checked = True Then
            vType = 2
        Else
            vType = 0
        End If
        Me.ListViewSearchItem.Items.Clear()
        vSearch = Me.TBSearchItem.Text
        If Me.CMBBrandCode.Text <> "" Then
            vBrandCode = vb6.Left(Me.CMBBrandCode.Text, vb6.InStr(Me.CMBBrandCode.Text, "/") - 1)
        Else
            vBrandCode = ""
        End If

        vQuery = "exec dbo.USP_NP_SearchItemPriceStructure " & vType & ",'" & vBrandCode & "','" & vSearch & "'"

        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchPaidNo")
        dt = ds.Tables("SearchPaidNo")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vCashSalePrice = dt.Rows(i).Item("cashsaleprice")
                vCreditSalePrice = dt.Rows(i).Item("creditsaleprice")

                vListItem = Me.ListViewSearchItem.Items.Add(n)
                vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("itemcode")
                vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("itemname")
                vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("unitcode")
                vListItem.SubItems.Add(3).Text = Format(vCashSalePrice, "##,##0.00")
                vListItem.SubItems.Add(4).Text = Format(vCreditSalePrice, "##,##0.00")

            Next
        End If
    End Sub

    Private Sub BTNItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNItem.Click
        Me.PNSearchItem.Visible = True
        Me.TBSearchItem.Focus()
    End Sub

    Private Sub BTNCloseSearchItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSearchItem.Click
        Me.PNSearchItem.Visible = False
        Me.TBDocNo.Focus()
    End Sub

    Public Sub SearchItemBrand()
        Dim i As Integer

        On Error Resume Next

        vQuery = "exec dbo.USP_PS_BrandList"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchBrand")
        dt = ds.Tables("SearchBrand")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                Me.CMBBrandCode.Items.Add(dt.Rows(i).Item("brandname"))
            Next
        End If
    End Sub

    Private Sub BTNSelectItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectItem.Click
        Dim i As Integer
        Dim n As Integer
        Dim m As Integer
        Dim a As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vCheckItemCode As String
        Dim vCheckUnitCode As String
        Dim vCheckAdd As Integer

        On Error Resume Next

        If Me.ListViewSearchItem.Items.Count > 0 Then
            For i = 0 To Me.ListViewSearchItem.Items.Count - 1
                If Me.ListViewSearchItem.Items(i).Checked = True Then

                    vItemCode = Me.ListViewSearchItem.Items(i).SubItems(1).Text
                    vItemName = Me.ListViewSearchItem.Items(i).SubItems(2).Text
                    vUnitCode = Me.ListViewSearchItem.Items(i).SubItems(3).Text

                    For n = 0 To Me.DGVItemDetails.RowCount - 1
                        vCheckItemCode = Me.DGVItemDetails.Item(1, n).Value
                        vCheckUnitCode = Me.DGVItemDetails.Item(3, n).Value

                        If vItemCode = vCheckItemCode And vUnitCode = vCheckUnitCode Then
                            vCheckAdd = 1
                            GoTo Line1
                        Else
                            vCheckAdd = 0
                        End If
                    Next

                    If vCheckAdd = 0 Then
                        For m = 0 To Me.DGVItemDetails.RowCount - 1
                            If Me.DGVItemDetails.Item(1, m).Value = Nothing Then
                                Me.DGVItemDetails.Item(1, m).Value = vItemCode
                                Me.DGVItemDetails.Item(2, m).Value = vItemName
                                Me.DGVItemDetails.Item(3, m).Value = vUnitCode
                                Me.DGVItemDetails.Item(4, m).Value = Format(1, "##,##0.00")
                                GoTo Line1
                            End If
                        Next
                    End If

                End If
Line1:
            Next


            For a = 0 To Me.ListViewSearchItem.Items.Count - 1
                Me.ListViewSearchItem.Items(a).Checked = False
            Next

            Me.PNSearchItem.Visible = False
        End If
    End Sub

    Public Sub AddCashComm()
        Dim i As Integer
        Dim vItemCode As String
        Dim vCashComm As Double

        On Error Resume Next

        If Me.CBCashComm.Checked = True Then
            vCashComm = Me.NMCashComm.Value

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                vItemCode = Me.DGVItemDetails.Item(1, i).Value
                If vItemCode <> "Nothing" And vItemCode <> "" Then

                    If vCashComm <> 0 Then
                        Me.DGVItemDetails.Item(5, i).Value = Format(vCashComm, "##,##0.00")
                    End If

                End If
            Next
        End If
    End Sub

    Public Sub AddCreditComm()
        Dim i As Integer
        Dim vItemCode As String
        Dim vCreditComm As Double

        On Error Resume Next

        If Me.CBCreditComm.Checked = True Then
            vCreditComm = Me.NMCreditComm.Value

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                vItemCode = Me.DGVItemDetails.Item(1, i).Value
                If vItemCode <> "Nothing" And vItemCode <> "" Then

                    If vCreditComm <> 0 Then
                        Me.DGVItemDetails.Item(6, i).Value = Format(vCreditComm, "##,##0.00")
                    End If

                End If
            Next
        End If
    End Sub

    Private Sub NMCashComm_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NMCashComm.ValueChanged
        Call AddCashComm()
    End Sub

    Private Sub NMCreditComm_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NMCreditComm.ValueChanged
        Call AddCreditComm()
    End Sub

    Private Sub CBCreditComm_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBCreditComm.CheckedChanged
        If Me.CBCreditComm.Checked = True Then
            Me.NMCreditComm.Enabled = True
        End If

        If Me.CBCreditComm.Checked = False Then
            Me.NMCreditComm.Enabled = False
        End If
    End Sub

    Private Sub CBCashComm_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBCashComm.CheckedChanged
        If Me.CBCashComm.Checked = True Then
            Me.NMCashComm.Enabled = True
        End If

        If Me.CBCashComm.Checked = False Then
            Me.NMCashComm.Enabled = False
        End If
    End Sub

    Private Sub CBSelectItemAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBSelectItemAll.CheckedChanged
        Dim i As Integer

        On Error Resume Next

        If Me.ListViewSearchItem.Items.Count > 0 Then
            If Me.CBSelectItemAll.Checked = True Then
                For i = 0 To Me.ListViewSearchItem.Items.Count - 1
                    Me.ListViewSearchItem.Items(i).Checked = True
                Next
            End If

            If Me.CBSelectItemAll.Checked = False Then
                For i = 0 To Me.ListViewSearchItem.Items.Count - 1
                    Me.ListViewSearchItem.Items(i).Checked = False
                Next
            End If
        End If
    End Sub

    Private Sub TBSearchItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchItem.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call SearchItem()
        End If
    End Sub

    Private Sub TBSearchItem_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchItem.TextChanged

    End Sub

    Private Sub DGVItemDetails_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItemDetails.CellEnter

    End Sub
End Class