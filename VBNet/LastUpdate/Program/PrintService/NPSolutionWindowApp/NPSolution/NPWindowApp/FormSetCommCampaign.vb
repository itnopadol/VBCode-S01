Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports vb6 = Microsoft.VisualBasic
Imports System.IO
Imports System.Globalization

Public Class FormSetCommCampaign
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vIsOpen As Integer
    Dim vMemIsCancel As Integer
    Dim vMemIsConfirm As Integer

    Private Sub FormSetCommCampaign_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()
        Call NewLoad()
        Call vGenDocNoAuto()
        Call NewDoc()
    End Sub
    Public Sub NewLoad()
        On Error Resume Next

        Me.TBCampaignCode.Text = ""
        Me.TBCampaignName.Text = ""
        Me.DTPStart.Value = Now
        Me.DTPStop.Value = Now
        Me.RBNotTarget.Checked = True
        Me.RBTarget.Checked = False
        Me.TBCash.Text = ""
        Me.TBCredit.Text = ""
        Me.TBSaleCash.Text = ""
        Me.TBSaleCredit.Text = ""
        Me.TBCash.Enabled = False
        Me.TBCredit.Enabled = False
        Me.TBSaleCash.Enabled = False
        Me.TBSaleCredit.Enabled = False
        Me.CMBPriceLevel.SelectedIndex = 1
        Me.RBCommission.Checked = True
        Me.RBNotCommission.Checked = False
        Me.PBNew.Visible = True
        Me.PBCancel.Visible = False
        vIsOpen = 0
        vMemIsCancel = 0
        vMemIsConfirm = 0
        Me.BTNCampaign.Focus()

    End Sub

    Private Sub BTNCampaign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCampaign.Click
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
        vQuery = "select dbo.ft_com_newcampaign ('" & vNow & "') as docno"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchCode")
        dt = ds.Tables("SearchCode")
        If dt.Rows.Count > 0 Then
            Me.TBCampaignCode.Text = dt.Rows(0).Item("docno")
        Else
            Me.TBCampaignCode.Text = ""
            MsgBox("กำหนด รหัสแคมเปญไม่ได้ เกิดปัญหา จาก Store : dbo.ft_com_newcampaign  ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vCheckStartDate As Date
        Dim vCheckStopDate As Date
        Dim vIsInsert As Integer
        Dim vStartDate As String
        Dim vStopDate As String

        Dim vCampaignCode As String
        Dim vCampaignName As String
        Dim vIsTarget As Integer
        Dim vTargetCash As Double
        Dim vTargetCredit As Double
        Dim vPriceLevel As Integer
        Dim vIsCommission As Integer

        Dim vMemBeginTran As Integer

        On Error GoTo ErrDescription

        vCheckStartDate = vb6.Day(Me.DTPStart.Text) & "/" & vb6.Month(Me.DTPStart.Text) & "/" & vb6.Year(Me.DTPStart.Text)
        vCheckStopDate = vb6.Day(Me.DTPStop.Text) & "/" & vb6.Month(Me.DTPStop.Text) & "/" & vb6.Year(Me.DTPStop.Text)

        vStartDate = vb6.Day(Me.DTPStart.Text) & "/" & vb6.Month(Me.DTPStart.Text) & "/" & vb6.Year(Me.DTPStart.Text)
        vStopDate = vb6.Day(Me.DTPStop.Text) & "/" & vb6.Month(Me.DTPStop.Text) & "/" & vb6.Year(Me.DTPStop.Text)

        If vMemIsConfirm = 1 Then
            MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNCampaign.Focus()
            Exit Sub
        End If

        If vMemIsCancel = 1 Then
            MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNCampaign.Focus()
            Exit Sub
        End If

        If vCheckStopDate < vCheckStartDate Then
            MsgBox("กำหนดวันหมดอายุของแคมเปญ น้อยกว่าวันเริ่มแคมเปญ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.DTPStop.Focus()
            Exit Sub
        End If

        If vIsOpen = 0 Then
            vIsInsert = 1
        Else
            vIsInsert = 0
        End If

        If Me.TBCampaignCode.Text = "" Then
            MsgBox("ยังไม่ได้กำหนด รหัสแคมเปญ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNCampaign.Focus()
            Exit Sub
        End If

        If Me.TBCampaignName.Text = "" Then
            MsgBox("ยังไม่ได้กำหนด ชื่อแคมเปญ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBCampaignName.Focus()
            Me.TBCampaignName.SelectAll()
            Exit Sub
        End If

        If Me.RBTarget.Checked = True Then
            vIsTarget = 1

            If Me.TBCash.Text = "" Then
                MsgBox("ยังไม่ได้กำหนด เป้าฝั่งขายเงินสด กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBCash.Focus()
                Me.TBCash.SelectAll()
                Exit Sub
            End If

            If Me.TBCredit.Text = "" Then
                MsgBox("ยังไม่ได้กำหนด เป้าฝั่งขายเงินเชื่อ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBCredit.Focus()
                Me.TBCredit.SelectAll()
                Exit Sub
            End If

            vTargetCash = Me.TBCash.Text
            vTargetCredit = Me.TBCredit.Text
        Else
            vIsTarget = 0
        End If

        vCampaignCode = Me.TBCampaignCode.Text
        vCampaignName = Me.TBCampaignName.Text
        vPriceLevel = Me.CMBPriceLevel.SelectedIndex + 1
        If Me.RBCommission.Checked = True Then
            vIsCommission = 1
        Else
            vIsCommission = 0
        End If


        vQuery = "begin tran"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vMemBeginTran = 1

        vQuery = "exec dbo.USP_COM_CampaignSave " & vIsInsert & ",'" & vCampaignCode & "','" & vCampaignName & "','" & vStartDate & "','" & vStopDate & "'," & vIsTarget & "," & vTargetCash & "," & vTargetCredit & "," & vPriceLevel & "," & vIsCommission & " "
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vQuery = "commit tran"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vMemBeginTran = 0
        MsgBox("บันทึกแคมเปญ " & vCampaignCode & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")
        Call NewLoad()
        Call vGenDocNoAuto()

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
        Dim vCheckStartDate As Date
        Dim vCheckStopDate As Date
        Dim vIsInsert As Integer
        Dim vStartDate As String
        Dim vStopDate As String

        Dim vCampaignCode As String
        Dim vCampaignName As String
        Dim vIsTarget As Integer
        Dim vTargetCash As Double
        Dim vTargetCredit As Double
        Dim vPriceLevel As Integer
        Dim vIsCommission As Integer

        Dim vMemBeginTran As Integer

        On Error GoTo ErrDescription

        vCheckStartDate = vb6.Day(Me.DTPStart.Text) & "/" & vb6.Month(Me.DTPStart.Text) & "/" & vb6.Year(Me.DTPStart.Text)
        vCheckStopDate = vb6.Day(Me.DTPStop.Text) & "/" & vb6.Month(Me.DTPStop.Text) & "/" & vb6.Year(Me.DTPStop.Text)

        vStartDate = vb6.Day(Me.DTPStart.Text) & "/" & vb6.Month(Me.DTPStart.Text) & "/" & vb6.Year(Me.DTPStart.Text)
        vStopDate = vb6.Day(Me.DTPStop.Text) & "/" & vb6.Month(Me.DTPStop.Text) & "/" & vb6.Year(Me.DTPStop.Text)

        If vMemIsConfirm = 1 Then
            MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNCampaign.Focus()
            Exit Sub
        End If

        If vMemIsCancel = 1 Then
            MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNCampaign.Focus()
            Exit Sub
        End If

        If vCheckStopDate < vCheckStartDate Then
            MsgBox("กำหนดวันหมดอายุของแคมเปญ น้อยกว่าวันเริ่มแคมเปญ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.DTPStop.Focus()
            Exit Sub
        End If


        If vIsOpen = 0 Then
            vIsInsert = 1
        Else
            vIsInsert = 0
        End If

        If Me.TBCampaignCode.Text = "" Then
            MsgBox("ยังไม่ได้กำหนด รหัสแคมเปญ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNCampaign.Focus()
            Exit Sub
        End If

        If Me.TBCampaignName.Text = "" Then
            MsgBox("ยังไม่ได้กำหนด ชื่อแคมเปญ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBCampaignName.Focus()
            Me.TBCampaignName.SelectAll()
            Exit Sub
        End If

        If Me.RBTarget.Checked = True Then
            vIsTarget = 1

            If Me.TBCash.Text = "" Then
                MsgBox("ยังไม่ได้กำหนด เป้าฝั่งขายเงินสด กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBCash.Focus()
                Me.TBCash.SelectAll()
                Exit Sub
            End If

            If Me.TBCredit.Text = "" Then
                MsgBox("ยังไม่ได้กำหนด เป้าฝั่งขายเงินเชื่อ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBCredit.Focus()
                Me.TBCredit.SelectAll()
                Exit Sub
            End If

            vTargetCash = Me.TBCash.Text
            vTargetCredit = Me.TBCredit.Text
        Else
            vIsTarget = 0
        End If

        vCampaignCode = Me.TBCampaignCode.Text
        vCampaignName = Me.TBCampaignName.Text
        vPriceLevel = Me.CMBPriceLevel.SelectedIndex + 1
        If Me.RBCommission.Checked = True Then
            vIsCommission = 1
        Else
            vIsCommission = 0
        End If


        vQuery = "begin tran"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vMemBeginTran = 1

        vQuery = "exec dbo.USP_COM_CampaignSave " & vIsInsert & ",'" & vCampaignCode & "','" & vCampaignName & "','" & vStartDate & "','" & vStopDate & "'," & vIsTarget & "," & vTargetCash & "," & vTargetCredit & "," & vPriceLevel & "," & vIsCommission & " "
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vQuery = "commit tran"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vMemBeginTran = 0
        MsgBox("บันทึกแคมเปญ " & vCampaignCode & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")
        Call NewLoad()
        Call vGenDocNoAuto()

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

    Private Sub RBTarget_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBTarget.CheckedChanged
        On Error Resume Next

        If Me.RBTarget.Checked = True Then
            Me.TBCash.Enabled = True
            Me.TBCredit.Enabled = True
            Me.TBCash.Focus()
            Me.TBCash.SelectAll()
        Else
            Me.TBCash.Enabled = False
            Me.TBCredit.Enabled = False
            Me.TBCash.Text = ""
            Me.TBCredit.Text = ""
            Me.TBCash.Focus()
            Me.TBCash.SelectAll()
        End If
    End Sub

    Private Sub RBNotTarget_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBNotTarget.CheckedChanged

    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        On Error Resume Next

        Me.TBSearch.Text = ""
        Me.ListViewSearch.Items.Clear()
        Me.PNSearch.Visible = False
        Me.BTNCampaign.Focus()
    End Sub

    Private Sub BTNClickSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClickSearch.Click
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
        Else
            Me.TBCampaignCode.Text = ""
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub BTNSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        Call SearchCampaign()
    End Sub

    Private Sub ListViewSearch_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewSearch.DoubleClick
        Dim vCode As String
        Dim vIsTarget As Integer
        Dim vIsCommission As Integer
        Dim vSaleCash As Double
        Dim vSaleCredit As Double
        Dim vTargetCash As Double
        Dim vTargetCredit As Double
        Dim vPriceLevel As Integer
        Dim vIndex As Integer

        On Error Resume Next

        If Me.ListViewSearch.Items.Count > 0 Then
            vIndex = Me.ListViewSearch.SelectedItems(0).Index
            vCode = Me.ListViewSearch.Items(vIndex).SubItems(1).Text
            vQuery = "exec dbo.USP_COM_CampaignSearch '" & vCode & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Search")
            dt = ds.Tables("Search")
            If dt.Rows.Count > 0 Then

                vIsOpen = 1
                vMemIsCancel = dt.Rows(0).Item("iscancel")
                vMemIsConfirm = 0
                vIsTarget = dt.Rows(0).Item("istarget")
                vIsCommission = dt.Rows(0).Item("ispromopaid")
                vSaleCash = dt.Rows(0).Item("retailamount")
                vSaleCredit = dt.Rows(0).Item("wholeamount")
                vTargetCash = dt.Rows(0).Item("retailtarget")
                vTargetCredit = dt.Rows(0).Item("wholesaletarget")
                vPriceLevel = dt.Rows(0).Item("pricelevelup")

                Me.TBCampaignCode.Text = dt.Rows(0).Item("code")
                Me.TBCampaignName.Text = dt.Rows(0).Item("name1")
                Me.DTPStart.Value = dt.Rows(0).Item("begindate")
                Me.DTPStop.Value = dt.Rows(0).Item("enddate")
                If vIsTarget = 1 Then
                    Me.RBTarget.Checked = True
                    Me.TBCash.Text = Format(vTargetCash, "##,##0.00")
                    Me.TBCredit.Text = Format(vTargetCredit, "##,##0.00")
                    Me.TBSaleCash.Text = Format(vSaleCash, "##,##0.00")
                    Me.TBSaleCredit.Text = Format(vSaleCredit, "##,##0.00")
                Else
                    Me.RBNotTarget.Checked = True
                    Me.TBCash.Text = ""
                    Me.TBCredit.Text = ""
                    Me.TBSaleCash.Text = ""
                    Me.TBSaleCredit.Text = ""
                End If

                If vIsCommission = 1 Then
                    Me.RBCommission.Checked = True
                Else
                    Me.RBNotCommission.Checked = True
                End If

                If vMemIsConfirm = 1 Then
                    Call ConfirmDoc()
                End If

                If vMemIsCancel = 1 Then
                    Call CancelDoc()
                End If

                If vMemIsCancel = 0 And vMemIsConfirm = 0 Then
                    Call NewDoc()
                End If

                Me.CMBPriceLevel.SelectedIndex = vPriceLevel - 1
                Me.TBSearch.Text = ""
                Me.ListViewSearch.Items.Clear()
                Me.PNSearch.Visible = False
                Me.BTNCampaign.Focus()
            End If
        End If
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

    Private Sub TBSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearch.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Call SearchCampaign()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBSearch.Text = ""
            Me.ListViewSearch.Items.Clear()
            Me.PNSearch.Visible = False
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub BTNSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelect.Click
        Dim vCode As String
        Dim vIsTarget As Integer
        Dim vIsCommission As Integer
        Dim vSaleCash As Double
        Dim vSaleCredit As Double
        Dim vTargetCash As Double
        Dim vTargetCredit As Double
        Dim vPriceLevel As Integer
        Dim vIndex As Integer

        On Error Resume Next

        If Me.ListViewSearch.Items.Count > 0 Then
            vIndex = Me.ListViewSearch.SelectedItems(0).Index
            vCode = Me.ListViewSearch.Items(vIndex).SubItems(1).Text
            vQuery = "exec dbo.USP_COM_CampaignSearch '" & vCode & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Search")
            dt = ds.Tables("Search")
            If dt.Rows.Count > 0 Then

                vIsOpen = 1
                vMemIsCancel = dt.Rows(0).Item("iscancel")
                vMemIsConfirm = 0
                vIsTarget = dt.Rows(0).Item("istarget")
                vIsCommission = dt.Rows(0).Item("ispromopaid")
                vSaleCash = dt.Rows(0).Item("retailamount")
                vSaleCredit = dt.Rows(0).Item("wholeamount")
                vTargetCash = dt.Rows(0).Item("retailtarget")
                vTargetCredit = dt.Rows(0).Item("wholesaletarget")
                vPriceLevel = dt.Rows(0).Item("pricelevelup")

                Me.TBCampaignCode.Text = dt.Rows(0).Item("code")
                Me.TBCampaignName.Text = dt.Rows(0).Item("name1")
                Me.DTPStart.Value = dt.Rows(0).Item("begindate")
                Me.DTPStop.Value = dt.Rows(0).Item("enddate")
                If vIsTarget = 1 Then
                    Me.RBTarget.Checked = True
                    Me.TBCash.Text = Format(vTargetCash, "##,##0.00")
                    Me.TBCredit.Text = Format(vTargetCredit, "##,##0.00")
                    Me.TBSaleCash.Text = Format(vSaleCash, "##,##0.00")
                    Me.TBSaleCredit.Text = Format(vSaleCredit, "##,##0.00")
                Else
                    Me.RBNotTarget.Checked = True
                    Me.TBCash.Text = ""
                    Me.TBCredit.Text = ""
                    Me.TBSaleCash.Text = ""
                    Me.TBSaleCredit.Text = ""
                End If

                If vIsCommission = 1 Then
                    Me.RBCommission.Checked = True
                Else
                    Me.RBNotCommission.Checked = True
                End If

                Me.CMBPriceLevel.SelectedIndex = vPriceLevel - 1
                Me.TBSearch.Text = ""
                Me.ListViewSearch.Items.Clear()
                Me.PNSearch.Visible = False
                Me.BTNCampaign.Focus()
            End If
        End If
    End Sub

    Private Sub ListViewSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearch.KeyDown
        Dim vCode As String
        Dim vIsTarget As Integer
        Dim vIsCommission As Integer
        Dim vSaleCash As Double
        Dim vSaleCredit As Double
        Dim vTargetCash As Double
        Dim vTargetCredit As Double
        Dim vPriceLevel As Integer
        Dim vIndex As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.TBSearch.Text = ""
            Me.ListViewSearch.Items.Clear()
            Me.PNSearch.Visible = False
            Me.BTNCampaign.Focus()
        End If

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewSearch.Items.Count > 0 Then
                vIndex = Me.ListViewSearch.SelectedItems(0).Index
                vCode = Me.ListViewSearch.Items(vIndex).SubItems(1).Text
                vQuery = "exec dbo.USP_COM_CampaignSearch '" & vCode & "'"
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "Search")
                dt = ds.Tables("Search")
                If dt.Rows.Count > 0 Then

                    vIsOpen = 1
                    vMemIsCancel = dt.Rows(0).Item("iscancel")
                    vMemIsConfirm = 0
                    vIsTarget = dt.Rows(0).Item("istarget")
                    vIsCommission = dt.Rows(0).Item("ispromopaid")
                    vSaleCash = dt.Rows(0).Item("retailamount")
                    vSaleCredit = dt.Rows(0).Item("wholeamount")
                    vTargetCash = dt.Rows(0).Item("retailtarget")
                    vTargetCredit = dt.Rows(0).Item("wholesaletarget")
                    vPriceLevel = dt.Rows(0).Item("pricelevelup")

                    Me.TBCampaignCode.Text = dt.Rows(0).Item("code")
                    Me.TBCampaignName.Text = dt.Rows(0).Item("name1")
                    Me.DTPStart.Value = dt.Rows(0).Item("begindate")
                    Me.DTPStop.Value = dt.Rows(0).Item("enddate")
                    If vIsTarget = 1 Then
                        Me.RBTarget.Checked = True
                        Me.TBCash.Text = Format(vTargetCash, "##,##0.00")
                        Me.TBCredit.Text = Format(vTargetCredit, "##,##0.00")
                        Me.TBSaleCash.Text = Format(vSaleCash, "##,##0.00")
                        Me.TBSaleCredit.Text = Format(vSaleCredit, "##,##0.00")
                    Else
                        Me.RBNotTarget.Checked = True
                        Me.TBCash.Text = ""
                        Me.TBCredit.Text = ""
                        Me.TBSaleCash.Text = ""
                        Me.TBSaleCredit.Text = ""
                    End If

                    If vIsCommission = 1 Then
                        Me.RBCommission.Checked = True
                    Else
                        Me.RBNotCommission.Checked = True
                    End If

                    If vMemIsConfirm = 1 Then
                        Call ConfirmDoc()
                    End If

                    If vMemIsCancel = 1 Then
                        Call CancelDoc()
                    End If

                    If vMemIsCancel = 0 And vMemIsConfirm = 0 Then
                        Call NewDoc()
                    End If

                    Me.CMBPriceLevel.SelectedIndex = vPriceLevel - 1
                    Me.TBSearch.Text = ""
                    Me.ListViewSearch.Items.Clear()
                    Me.PNSearch.Visible = False
                    Me.BTNCampaign.Focus()
                End If
            End If
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

    Private Sub BTNClearScreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearScreen.Click
        Call ClearScreen()
    End Sub

    Public Sub ClearScreen()
        On Error Resume Next

        vIsOpen = 0
        vMemIsCancel = 0
        vMemIsConfirm = 0
        Me.TBCampaignCode.Text = ""
        Me.TBCampaignName.Text = ""
        Me.DTPStart.Value = Now
        Me.DTPStop.Value = Now
        Me.RBNotTarget.Checked = True
        Me.TBCash.Text = ""
        Me.TBCredit.Text = ""
        Me.TBSaleCash.Text = ""
        Me.TBSaleCredit.Text = ""
        Me.RBCommission.Checked = True
        Me.CMBPriceLevel.SelectedIndex = 1
        Call NewDoc()
        Call vGenDocNoAuto()
        Me.BTNCampaign.Focus()
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Me.Close()
    End Sub

    Private Sub BTNPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNPrint.Click
        If vIsOpen = 1 Then
            vMemCampaignNo = Me.TBCampaignCode.Text


            If vMemIsConfirm = 1 Then
                MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNCampaign.Focus()
                Exit Sub
            End If

            If vMemIsCancel = 1 Then
                MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNCampaign.Focus()
                Exit Sub
            End If

            If frmReportSetConditionComm Is Nothing Then
                frmReportSetConditionComm = New FormReportSetConditionComm
            Else
                If frmReportSetConditionComm.IsDisposed Then
                    frmReportSetConditionComm = New FormReportSetConditionComm
                End If
            End If

            frmReportSetConditionComm.Show()
            frmReportSetConditionComm.BringToFront()
        End If
    End Sub

    Public Sub PrintDocument()
        If vIsOpen = 1 Then
            vMemCampaignNo = Me.TBCampaignCode.Text

            If vMemIsConfirm = 1 Then
                MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNCampaign.Focus()
                Exit Sub
            End If

            If vMemIsCancel = 1 Then
                MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNCampaign.Focus()
                Exit Sub
            End If

            If frmReportSetConditionComm Is Nothing Then
                frmReportSetConditionComm = New FormReportSetConditionComm
            Else
                If frmReportSetConditionComm.IsDisposed Then
                    frmReportSetConditionComm = New FormReportSetConditionComm
                End If
            End If

            frmReportSetConditionComm.Show()
            frmReportSetConditionComm.BringToFront()
        End If
    End Sub

    Private Sub BTNCampaign_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCampaign.KeyDown, TBCampaignCode.KeyDown, TBCampaignName.KeyDown, DTPStart.KeyDown, DTPStop.KeyDown, RBTarget.KeyDown, RBNotTarget.KeyDown, TBCash.KeyDown, TBCredit.KeyDown, TBSaleCash.KeyDown, TBSaleCredit.KeyDown, CMBPriceLevel.KeyDown, RBCommission.KeyDown, RBNotCommission.KeyDown, BTNClearScreen.KeyDown, BTNSave.KeyDown, BTNSearch.KeyDown, BTNPrint.KeyDown, BTNExit.KeyDown
        If e.KeyCode = Keys.F1 Then
            Call SearchCampaign()
        End If

        If e.KeyCode = Keys.F4 Then
            Call ClearScreen()
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

    Private Sub TBSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearch.TextChanged

    End Sub

    Private Sub BTNClickSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClickSearch.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.TBSearch.Text = ""
            Me.ListViewSearch.Items.Clear()
            Me.PNSearch.Visible = False
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub BTNSelect_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSelect.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.TBSearch.Text = ""
            Me.ListViewSearch.Items.Clear()
            Me.PNSearch.Visible = False
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub BTNClose_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClose.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.TBSearch.Text = ""
            Me.ListViewSearch.Items.Clear()
            Me.PNSearch.Visible = False
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub TBCredit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBCredit.KeyDown
        If e.KeyCode = Keys.Up Then
            Me.TBCash.Focus()
            Me.TBCash.SelectAll()
        End If
    End Sub

    Private Sub TBCash_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBCash.KeyDown
        If e.KeyCode = Keys.Up Then
            Me.RBTarget.Focus()
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBCredit.Focus()
            Me.TBCredit.SelectAll()
        End If
    End Sub

    Private Sub TBCash_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBCash.KeyPress, TBCredit.KeyPress
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

    Private Sub TBCash_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBCash.TextChanged

    End Sub

    Private Sub BTNCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNCancel.Click
        If vIsOpen = 1 Then
            If vMemIsConfirm = 1 Then
                MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNCampaign.Focus()
                Exit Sub
            End If

            If vMemIsCancel = 1 Then
                MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.BTNCampaign.Focus()
                Exit Sub
            End If

            MsgBox("เมนูยกเลิก ยังไม่ได้เปิดใช้งาน", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNCampaign.Focus()
        End If
    End Sub

    Private Sub RBNotTarget_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RBNotTarget.KeyDown
        If e.KeyCode = Keys.Left Then
            Me.RBTarget.Checked = True
        End If
    End Sub

    Private Sub RBTarget_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RBTarget.KeyDown
        If e.KeyCode = Keys.Right Then
            Me.RBNotTarget.Checked = True
        End If
    End Sub

    Private Sub TBCampaignCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBCampaignCode.TextChanged

    End Sub

    Private Sub TBCredit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBCredit.TextChanged

    End Sub
End Class