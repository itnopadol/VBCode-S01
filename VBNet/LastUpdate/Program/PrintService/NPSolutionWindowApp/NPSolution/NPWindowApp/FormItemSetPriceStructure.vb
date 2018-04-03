Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports CrystalDecisions
Imports System
Imports Microsoft
Public Class FormItemSetPriceStructure
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

    Private Sub FormItemSetPriceStructure_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()
        Call vGetBeginData()
        Call SearchItemBrand()
        Call NewDoc()
        Call vGenDocNoAuto()
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

    Public Sub vGenDocNoAuto()

        On Error Resume Next

        vQuery = "set dateformat dmy"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vQuery = "exec dbo.USP_PS_NewDocno"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "NewDocno")
        dt = ds.Tables("NewDocno")
        If dt.Rows.Count > 0 Then
            Me.TBDocNo.Text = dt.Rows(0).Item("newdocno")
        Else
            Me.TBDocNo.Text = ""
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End If
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
        Dim vItemName As String
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
        Dim vAddDate As Date
        Dim vCheckUpdate As Date
        Dim vDocDate As String

        Dim vDOAmount As Double
        Dim vBillDisc As Double
        Dim vBillDiscAmount As Double
        Dim vAccCost As Double
        Dim vDisc1 As Double
        Dim vDiscAmount1 As Double
        Dim vAfterDiscAmount1 As Double

        Dim vGetBillDisc As String
        Dim vGetFollowDisc1 As String
        Dim vGetFollowDisc2 As String
        Dim vGetFollowDisc3 As String


        On Error Resume Next

        If Me.TBDocNo.Text = "" Then
            MsgBox("กรุณา กรอกเลขที่เอกสาร ก่อนเลือกสินค้า", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBDocNo.Focus()
            Exit Sub
        End If

        vNowDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        vCheckUpdate = vb6.Day(Me.DTPUpdate.Value) & "/" & vb6.Month(Me.DTPUpdate.Value) & "/" & vb6.Year(Me.DTPUpdate.Value)

        vAddDate = vb6.DateAdd(DateInterval.Day, 1, vNowDate)

        If vCheckUpdate > vNowDate Then
            If vb6.Left(vCheckUpdate.Year, 2) = "20" Then
                vDocDate = vCheckUpdate
            Else
                vDocDate = vb6.Day(vCheckUpdate) & "/" & vb6.Month(vCheckUpdate) & "/" & vb6.Year(vCheckUpdate) - 543
            End If
        Else
            If vb6.Left(vAddDate.Year, 2) = "20" Then
                vDocDate = vAddDate
            Else
                vDocDate = vb6.Day(vAddDate) & "/" & vb6.Month(vAddDate) & "/" & vb6.Year(vAddDate) - 543
            End If
        End If

        vDocNo = Me.TBDocNo.Text

        If vMemIsCancel = 1 Then
            MsgBox("เอกสารเลขที่ " & vDocNo & " ถูกยกเลิกไปแล้ว ไม่สามารถแก้ไขข้อมูลได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Information Message")
            Me.DGVItemDetails.Item(1, e.RowIndex).Value = ""
            Me.TBDocNo.Focus()
            Exit Sub
        End If

        If vMemIsConfirm = 1 Then
            MsgBox("เอกสารเลขที่ " & vDocNo & " ถูกอนุมัติไปแล้ว ไม่สามารถแก้ไขข้อมูลได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Information Message")
            Me.DGVItemDetails.Item(1, e.RowIndex).Value = ""
            Me.TBDocNo.Focus()
            Exit Sub
        End If

        vColumn = Me.DGVItemDetails.CurrentCell.ColumnIndex
        vRow = Me.DGVItemDetails.CurrentCell.RowIndex
        vItemCode = Me.DGVItemDetails.Item(1, e.RowIndex).Value '

        If vItemCode = "" Then
            Me.DGVItemDetails.Item(2, e.RowIndex).Value = "" '
            Me.DGVItemDetails.Item(3, e.RowIndex).Value = "" '
            Me.DGVItemDetails.Item(4, e.RowIndex).Value = "" '
            Me.DGVItemDetails.Item(5, e.RowIndex).Value = "" '
            Me.DGVItemDetails.Item(6, e.RowIndex).Value = "" '
            Me.DGVItemDetails.Item(7, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(8, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(9, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(10, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(11, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(12, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(13, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(14, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(15, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(16, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(17, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(18, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(19, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(20, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(21, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(22, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(23, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(24, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(25, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(26, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(27, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(28, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(29, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(30, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(31, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(32, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(33, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(34, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(35, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(36, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(37, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(38, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(39, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(40, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(41, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(42, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(43, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(44, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(45, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(46, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(47, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(48, e.RowIndex).Value = ""
            Me.DGVItemDetails.Item(49, e.RowIndex).Value = ""
        End If

        If vColumn = 1 Then
            If vItemCode <> "" Then
                For i = 0 To Me.DGVItemDetails.Rows.Count - 1
                    vCheckItemCode = Me.DGVItemDetails.Item(1, i).Value '

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
                    vItemName = dt.Rows(0).Item("itemname")
                    Me.DGVItemDetails.Item(2, vRow).Value = dt.Rows(0).Item("itemname") '
                    Me.DGVItemDetails.Item(3, vRow).Value = dt.Rows(0).Item("unitcode") '
                    Me.DGVItemDetails.Item(47, vRow).Value = vDocDate '

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

        If Me.CBBillDisc.Checked = True And Me.TBBillDisc.Text <> "" Then
            vGetBillDisc = Me.TBBillDisc.Text

            Me.DGVItemDetails.Item(5, e.RowIndex).Value = vGetBillDisc
        End If

        If Me.CBFollowDisc1.Checked = True And Me.TBFollowDisc1.Text <> "" Then
            vGetFollowDisc1 = Me.TBFollowDisc1.Text

            Me.DGVItemDetails.Item(7, e.RowIndex).Value = vGetFollowDisc1
        End If

        If Me.CBFollowDisc2.Checked = True And Me.TBFollowDisc2.Text <> "" Then
            vGetFollowDisc2 = Me.TBFollowDisc2.Text

            Me.DGVItemDetails.Item(9, e.RowIndex).Value = vGetFollowDisc2
        End If

        If Me.CBFollowDisc3.Checked = True And Me.TBFollowDisc3.Text <> "" Then
            vGetFollowDisc3 = Me.TBFollowDisc3.Text

            Me.DGVItemDetails.Item(11, e.RowIndex).Value = vGetFollowDisc3
        End If

        Dim vCharStr As String
        If e.ColumnIndex = 4 Then
            vCharStr = Me.DGVItemDetails.Item(4, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(4, e.RowIndex).Value = ""
                    MsgBox("ช่อง D/O ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(4, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 5 Then
            vCharStr = Me.DGVItemDetails.Item(5, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(5, e.RowIndex).Value = ""
                    MsgBox("ช่องส่วนลดหน้าบิล% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(5, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 7 Then
            vCharStr = Me.DGVItemDetails.Item(7, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(7, e.RowIndex).Value = ""
                    MsgBox("ช่องส่วนลดตาม1% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(7, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 9 Then
            vCharStr = Me.DGVItemDetails.Item(9, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(9, e.RowIndex).Value = ""
                    MsgBox("ช่องส่วนลดตาม2% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(9, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 11 Then
            vCharStr = Me.DGVItemDetails.Item(11, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(11, e.RowIndex).Value = ""
                    MsgBox("ช่องส่วนลดตาม3% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(11, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 13 Then
            vCharStr = Me.DGVItemDetails.Item(13, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(13, e.RowIndex).Value = ""
                    MsgBox("ช่องส่วนลดตาม4% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(13, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 15 Then
            vCharStr = Me.DGVItemDetails.Item(15, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(15, e.RowIndex).Value = ""
                    MsgBox("ช่องส่วนลดRebate% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(15, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 17 Then
            vCharStr = Me.DGVItemDetails.Item(17, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(17, e.RowIndex).Value = ""
                    MsgBox("ช่องส่วนลดพิเศษ ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(17, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 19 Then
            vCharStr = Me.DGVItemDetails.Item(19, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(19, e.RowIndex).Value = ""
                    MsgBox("ช่องงบขาดทุน ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(19, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 21 Then
            vCharStr = Me.DGVItemDetails.Item(21, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(21, e.RowIndex).Value = ""
                    MsgBox("ช่องค่าขนส่งเข้า ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(21, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 22 Then
            vCharStr = Me.DGVItemDetails.Item(22, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(22, e.RowIndex).Value = ""
                    MsgBox("ช่องค่าขนส่งให้ลูกค้า ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(22, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 23 Then
            vCharStr = Me.DGVItemDetails.Item(23, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(23, e.RowIndex).Value = ""
                    MsgBox("ช่องโฆษณา% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(23, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 26 Then
            vCharStr = Me.DGVItemDetails.Item(26, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(26, e.RowIndex).Value = ""
                    MsgBox("ช่องค่าแรงค่าติดตั้ง ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(26, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 27 Then
            vCharStr = Me.DGVItemDetails.Item(27, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(27, e.RowIndex).Value = ""
                    MsgBox("ช่องบริการ ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(27, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 35 Then
            vCharStr = Me.DGVItemDetails.Item(35, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(35, e.RowIndex).Value = ""
                    MsgBox("ช่องSmartPoint% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(35, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
            'MsgBox(Me.DGVItemDetails.Item(35, e.RowIndex).Value)
        End If

        If e.ColumnIndex = 37 Then
            vCharStr = Me.DGVItemDetails.Item(37, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(37, e.RowIndex).Value = ""
                    MsgBox("ช่องเป้า ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(37, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 38 Then
            vCharStr = Me.DGVItemDetails.Item(38, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(38, e.RowIndex).Value = ""
                    MsgBox("ช่องของแถม ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(38, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 39 Then
            vCharStr = Me.DGVItemDetails.Item(39, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(39, e.RowIndex).Value = ""
                    MsgBox("ช่องคอม% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(39, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 42 Then
            vCharStr = Me.DGVItemDetails.Item(42, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(42, e.RowIndex).Value = ""
                    MsgBox("ช่องราคา1เงินสดรับเอง ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(42, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 44 Then
            vCharStr = Me.DGVItemDetails.Item(44, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(44, e.RowIndex).Value = ""
                    MsgBox("ช่องราคา1เงินเชื่อรับเอง ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(44, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        If e.ColumnIndex = 46 Then
            vCharStr = Me.DGVItemDetails.Item(46, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVItemDetails.Item(46, e.RowIndex).Value = ""
                    MsgBox("ช่องราคา2 ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(46, e.RowIndex).Selected = True
                    Exit Sub
                End If
            End If
        End If

        Dim vDisc2 As Double
        Dim vDiscAmount2 As Double
        Dim vAfterDiscAmount2 As Double

        Dim vDisc3 As Double
        Dim vDiscAmount3 As Double
        Dim vAfterDiscAmount3 As Double

        Dim vDisc4 As Double
        Dim vDiscAmount4 As Double
        Dim vAfterDiscAmount4 As Double

        Dim vRebate As Double
        Dim vRebateAmount As Double
        Dim vAfterRebateAmount As Double

        Dim vNetCost As Double
        Dim vDiscSpecial As Double

        Dim vLose As Double
        Dim vLoseAmount As Double
        Dim vAfterLoseAmount As Double

        Dim vTransInAmount As Double
        Dim vTransOutAmount As Double
        Dim vAdvertise As Double
        Dim vAdvertiseAmount As Double
        Dim vAfterAdvertiseAmount As Double

        Dim vVatCost As Double

        Dim vInstallAmount As Double
        Dim vServiceAmount As Double
        Dim vMarketCost As Double

        Dim vRelateStockPercent As Double
        Dim vRelateStockAmount As Double

        Dim vSmartPoint As Double
        Dim vSmartPointAmount As Double
        Dim vAfterSmartPointAmount As Double

        Dim vCashProfit As Double
        Dim vCashProfitAmount As Double
        Dim vAfterCashProfitAmount As Double

        Dim vCreditProfit As Double
        Dim vCreditProfitAmount As Double
        Dim vAfterCreditProfitAmount As Double

        Dim vTotalPrice As Double
        Dim vCashPriceOwn1 As Double
        Dim vCashPriceSend1 As Double
        Dim vCreditPriceOwn1 As Double
        Dim vCreditPriceSend1 As Double
        Dim vSalePrice2 As Double

        Dim vTargetAmount As Double
        Dim vPremiumAmount As Double

        Dim vComm As Double
        Dim vCommAmount As Double
        Dim vAfterCommAmount As Double

        Dim vBaseProfitPercent As Double
        Dim vBaseProfit As Double
        Dim vRelateStockPercentShow As Double

        vDOAmount = Me.DGVItemDetails.Item(4, e.RowIndex).Value
        vBillDisc = Me.DGVItemDetails.Item(5, e.RowIndex).Value
        vBillDiscAmount = (vDOAmount * (vBillDisc / 100))
        vAccCost = vDOAmount - vBillDiscAmount

        vDisc1 = Me.DGVItemDetails.Item(7, e.RowIndex).Value
        vDiscAmount1 = (vAccCost * (vDisc1 / 100))
        vAfterDiscAmount1 = vAccCost - vDiscAmount1

        vDisc2 = Me.DGVItemDetails.Item(9, e.RowIndex).Value
        vDiscAmount2 = (vAfterDiscAmount1 * (vDisc2 / 100))
        vAfterDiscAmount2 = vAfterDiscAmount1 - vDiscAmount2

        vDisc3 = Me.DGVItemDetails.Item(11, e.RowIndex).Value
        vDiscAmount3 = (vAfterDiscAmount2 * (vDisc3 / 100))
        vAfterDiscAmount3 = vAfterDiscAmount2 - vDiscAmount3

        vDisc4 = Me.DGVItemDetails.Item(13, e.RowIndex).Value
        vDiscAmount4 = (vAfterDiscAmount3 * (vDisc4 / 100))
        vAfterDiscAmount4 = vAfterDiscAmount3 - vDiscAmount4

        vRebate = Me.DGVItemDetails.Item(15, e.RowIndex).Value
        vRebateAmount = (vAfterDiscAmount4 * (vRebate / 100))
        vAfterRebateAmount = vAfterDiscAmount4 - vRebateAmount

        vDiscSpecial = Me.DGVItemDetails.Item(17, e.RowIndex).Value
        vNetCost = vAfterRebateAmount - vDiscSpecial

        vLose = Me.DGVItemDetails.Item(19, e.RowIndex).Value
        vLoseAmount = (vNetCost * (vLose / 100))
        vAfterLoseAmount = vNetCost - vLoseAmount

        vTransInAmount = Me.DGVItemDetails.Item(21, e.RowIndex).Value
        vTransOutAmount = Me.DGVItemDetails.Item(22, e.RowIndex).Value
        vAdvertise = Me.DGVItemDetails.Item(23, e.RowIndex).Value
        vAdvertiseAmount = (vAfterLoseAmount * (vAdvertise / 100))
        vAfterAdvertiseAmount = (vAfterLoseAmount + vTransInAmount + vTransOutAmount) + vAdvertiseAmount

        vVatCost = (vAfterAdvertiseAmount * 0.07) + vAfterAdvertiseAmount

        vInstallAmount = Me.DGVItemDetails.Item(26, e.RowIndex).Value
        vServiceAmount = Me.DGVItemDetails.Item(27, e.RowIndex).Value
        vMarketCost = vVatCost + vInstallAmount + vServiceAmount

        vCashPriceOwn1 = Me.DGVItemDetails.Item(42, e.RowIndex).Value
        vCashPriceSend1 = vCashPriceOwn1
        vCreditPriceOwn1 = Me.DGVItemDetails.Item(44, e.RowIndex).Value
        vCreditPriceSend1 = vCreditPriceOwn1

        If Me.TBSmartPoint.Text = "" Then
            vSmartPoint = Me.DGVItemDetails.Item(35, e.RowIndex).Value
        Else
            vSmartPoint = Me.TBSmartPoint.Text
            Me.DGVItemDetails.Item(35, e.RowIndex).Value = Format(vSmartPoint, "##,##0.0000")
        End If

        vSmartPointAmount = vSmartPoint
        vAfterSmartPointAmount = (vSmartPointAmount * vCashPriceOwn1) / 100

        vTargetAmount = Me.DGVItemDetails.Item(37, e.RowIndex).Value
        vPremiumAmount = Me.DGVItemDetails.Item(38, e.RowIndex).Value

        vComm = Me.DGVItemDetails.Item(39, e.RowIndex).Value
        vCommAmount = vComm / 100
        vAfterCommAmount = vCommAmount * vCashPriceOwn1

        vBaseProfitPercent = ((vCashPriceOwn1 - vAfterCommAmount - vPremiumAmount - vAfterSmartPointAmount - vTargetAmount - vMarketCost) / vMarketCost) * 100
        vBaseProfit = ((vCashPriceOwn1 - vAfterCommAmount - vPremiumAmount - vAfterSmartPointAmount - vTargetAmount - vMarketCost) / vMarketCost)

        If vBaseProfit < 0.01 Then
            vRelateStockPercent = 0
        ElseIf vBaseProfit >= 0.01 And vBaseProfit < 0.05 Then
            vRelateStockPercent = vBaseProfit / 2
        ElseIf vBaseProfit >= 0.05 And vBaseProfit < 0.06 Then
            vRelateStockPercent = vBaseProfit * 0.6
        Else
            vRelateStockPercent = 0.035
        End If

        vRelateStockPercentShow = vRelateStockPercent * 100
        vRelateStockAmount = vRelateStockPercent * vMarketCost

        vAfterCashProfitAmount = vCashPriceOwn1 - vAfterCommAmount - vPremiumAmount - vAfterSmartPointAmount - vTargetAmount - vRelateStockAmount - vMarketCost
        vCashProfit = (vAfterCashProfitAmount / vMarketCost) * 100

        vAfterCreditProfitAmount = vCreditPriceOwn1 - vMarketCost - vRelateStockAmount
        vCreditProfit = (vAfterCreditProfitAmount / vMarketCost) * 100

        vTotalPrice = vMarketCost + vRelateStockAmount + vAfterCashProfitAmount + vAfterSmartPointAmount + vTargetAmount + vPremiumAmount + vAfterCommAmount

        vItemName = Me.DGVItemDetails.Item(2, e.RowIndex).Value
        vSalePrice2 = Me.DGVItemDetails.Item(46, e.RowIndex).Value

        If vItemName <> "Nothing" And vItemName <> "" Then
            If vDOAmount <> 0 Then
                Me.DGVItemDetails.Item(4, e.RowIndex).Value = Format(vDOAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(4, e.RowIndex).Value = ""
            End If

            If vBillDisc = 0 Then
                Me.DGVItemDetails.Item(5, e.RowIndex).Value = ""
            End If

            If vAccCost <> 0 Then
                Me.DGVItemDetails.Item(6, e.RowIndex).Value = Format(vAccCost, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(6, e.RowIndex).Value = ""
            End If

            If vDisc1 = 0 Then
                Me.DGVItemDetails.Item(7, e.RowIndex).Value = ""
            End If

            If vAfterDiscAmount1 <> 0 Then
                Me.DGVItemDetails.Item(8, e.RowIndex).Value = Format(vAfterDiscAmount1, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(8, e.RowIndex).Value = ""
            End If

            If vDisc2 = 0 Then
                Me.DGVItemDetails.Item(9, e.RowIndex).Value = ""
            End If

            If vAfterDiscAmount2 <> 0 Then
                Me.DGVItemDetails.Item(10, e.RowIndex).Value = Format(vAfterDiscAmount2, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(10, e.RowIndex).Value = ""
            End If

            If vDisc3 = 0 Then
                Me.DGVItemDetails.Item(11, e.RowIndex).Value = ""
            End If

            If vAfterDiscAmount3 <> 0 Then
                Me.DGVItemDetails.Item(12, e.RowIndex).Value = Format(vAfterDiscAmount3, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(12, e.RowIndex).Value = ""
            End If

            If vDisc4 = 0 Then
                Me.DGVItemDetails.Item(13, e.RowIndex).Value = ""
            End If

            If vAfterDiscAmount4 <> 0 Then
                Me.DGVItemDetails.Item(14, e.RowIndex).Value = Format(vAfterDiscAmount4, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(14, e.RowIndex).Value = ""
            End If

            If vRebate = 0 Then
                Me.DGVItemDetails.Item(15, e.RowIndex).Value = ""
            End If

            If vAfterRebateAmount <> 0 Then
                Me.DGVItemDetails.Item(16, e.RowIndex).Value = Format(vAfterRebateAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(16, e.RowIndex).Value = ""
            End If

            If vDiscSpecial <> 0 Then
                Me.DGVItemDetails.Item(17, e.RowIndex).Value = Format(vDiscSpecial, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(17, e.RowIndex).Value = ""
            End If

            If vNetCost <> 0 Then
                Me.DGVItemDetails.Item(18, e.RowIndex).Value = Format(vNetCost, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(18, e.RowIndex).Value = ""
            End If

            If vLose = 0 Then
                Me.DGVItemDetails.Item(19, e.RowIndex).Value = ""
            End If

            If vAfterLoseAmount <> 0 Then
                Me.DGVItemDetails.Item(20, e.RowIndex).Value = Format(vAfterLoseAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(20, e.RowIndex).Value = ""
            End If

            If vTransInAmount <> 0 Then
                Me.DGVItemDetails.Item(21, e.RowIndex).Value = Format(vTransInAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(21, e.RowIndex).Value = ""
            End If

            If vTransOutAmount <> 0 Then
                Me.DGVItemDetails.Item(22, e.RowIndex).Value = Format(vTransOutAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(22, e.RowIndex).Value = ""
            End If

            If vAdvertise = 0 Then
                Me.DGVItemDetails.Item(23, e.RowIndex).Value = ""
            End If

            If vAfterAdvertiseAmount <> 0 Then
                Me.DGVItemDetails.Item(24, e.RowIndex).Value = Format(vAfterAdvertiseAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(24, e.RowIndex).Value = ""
            End If

            If vVatCost <> 0 Then
                Me.DGVItemDetails.Item(25, e.RowIndex).Value = Format(vVatCost, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(25, e.RowIndex).Value = ""
            End If

            If vInstallAmount <> 0 Then
                Me.DGVItemDetails.Item(26, e.RowIndex).Value = Format(vInstallAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(26, e.RowIndex).Value = ""
            End If

            If vServiceAmount <> 0 Then
                Me.DGVItemDetails.Item(27, e.RowIndex).Value = Format(vServiceAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(27, e.RowIndex).Value = ""
            End If

            If vMarketCost <> 0 Then
                Me.DGVItemDetails.Item(28, e.RowIndex).Value = Format(vMarketCost, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(28, e.RowIndex).Value = ""
            End If

            If vRelateStockPercentShow <> 0 Then
                Me.DGVItemDetails.Item(29, e.RowIndex).Value = Format(vRelateStockPercentShow, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(29, e.RowIndex).Value = ""
            End If

            If vRelateStockAmount <> 0 Then
                Me.DGVItemDetails.Item(30, e.RowIndex).Value = Format(vRelateStockAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(30, e.RowIndex).Value = ""
            End If

            If vCashProfit <> 0 Then
                Me.DGVItemDetails.Item(31, e.RowIndex).Value = Format(vCashProfit, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(31, e.RowIndex).Value = ""
            End If

            If vAfterCashProfitAmount <> 0 Then
                Me.DGVItemDetails.Item(32, e.RowIndex).Value = Format(vAfterCashProfitAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(32, e.RowIndex).Value = ""
            End If

            If vCreditProfit <> 0 Then
                Me.DGVItemDetails.Item(33, e.RowIndex).Value = Format(vCreditProfit, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(33, e.RowIndex).Value = ""
            End If

            If vAfterCreditProfitAmount <> 0 Then
                Me.DGVItemDetails.Item(34, e.RowIndex).Value = Format(vAfterCreditProfitAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(34, e.RowIndex).Value = ""
            End If

            If vSmartPoint = 0 Then
                Me.DGVItemDetails.Item(35, e.RowIndex).Value = ""
            End If

            If vAfterSmartPointAmount <> 0 Then
                Me.DGVItemDetails.Item(36, e.RowIndex).Value = Format(vAfterSmartPointAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(36, e.RowIndex).Value = ""
            End If

            If vTargetAmount <> 0 Then
                Me.DGVItemDetails.Item(37, e.RowIndex).Value = Format(vTargetAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(37, e.RowIndex).Value = ""
            End If

            If vPremiumAmount <> 0 Then
                Me.DGVItemDetails.Item(38, e.RowIndex).Value = Format(vPremiumAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(38, e.RowIndex).Value = ""
            End If

            If vComm = 0 Then
                Me.DGVItemDetails.Item(39, e.RowIndex).Value = ""
            End If

            If vAfterCommAmount <> 0 Then
                Me.DGVItemDetails.Item(40, e.RowIndex).Value = Format(vAfterCommAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(40, e.RowIndex).Value = ""
            End If

            If vTotalPrice <> 0 Then
                Me.DGVItemDetails.Item(41, e.RowIndex).Value = Format(vTotalPrice, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(41, e.RowIndex).Value = ""
            End If

            If vCashPriceOwn1 <> 0 Then
                Me.DGVItemDetails.Item(42, e.RowIndex).Value = Format(vCashPriceOwn1, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(42, e.RowIndex).Value = ""
            End If

            If vCashPriceSend1 <> 0 Then
                Me.DGVItemDetails.Item(43, e.RowIndex).Value = Format(vCashPriceSend1, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(43, e.RowIndex).Value = ""
            End If

            If vCreditPriceOwn1 <> 0 Then
                Me.DGVItemDetails.Item(44, e.RowIndex).Value = Format(vCreditPriceOwn1, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(44, e.RowIndex).Value = ""
            End If

            If vCreditPriceSend1 <> 0 Then
                Me.DGVItemDetails.Item(45, e.RowIndex).Value = Format(vCreditPriceSend1, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(45, e.RowIndex).Value = ""
            End If

            If vSalePrice2 <> 0 Then
                Me.DGVItemDetails.Item(46, e.RowIndex).Value = Format(vSalePrice2, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(46, e.RowIndex).Value = ""
            End If

            If vBaseProfitPercent <> 0 Then
                Me.DGVItemDetails.Item(49, e.RowIndex).Value = Format(vBaseProfitPercent, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(49, e.RowIndex).Value = ""
            End If
            If vDOAmount > 0 Then
                Call CalcCashPriceLine(e.RowIndex)
                Call CalcCreditPriceLine(e.RowIndex)
                Call CalcPrice2Line(e.RowIndex)

                Call CalcDOCashPriceLine(e.RowIndex)
                Call CalcDOCreditPriceLine(e.RowIndex)
                Call CalcDOPrice2Line(e.RowIndex)
            End If

            If vAfterCashProfitAmount < 0 And Me.DGVItemDetails.Item(42, e.RowIndex).Value <> "" Then
                Me.DGVItemDetails.Item(42, e.RowIndex).Style.BackColor = Color.Red
                MsgBox("กำไรขายปลีกติดลบ", MsgBoxStyle.Critical, "Send Information Message")
            Else
                Me.DGVItemDetails.Item(42, e.RowIndex).Style.BackColor = Color.Linen
            End If

            If vCreditPriceOwn1 < vCashPriceOwn1 And Me.DGVItemDetails.Item(44, e.RowIndex).Value <> "" Then
                Me.DGVItemDetails.Item(44, e.RowIndex).Style.BackColor = Color.Red
                MsgBox("กำหนดราคาเงินเชื่อต่ำกว่าราคาเงินสด", MsgBoxStyle.Critical, "Send Information Message")
            Else
                Me.DGVItemDetails.Item(44, e.RowIndex).Style.BackColor = Color.Linen
            End If

            If vSalePrice2 < vMarketCost And Me.DGVItemDetails.Item(46, e.RowIndex).Value <> "" Then
                Me.DGVItemDetails.Item(46, e.RowIndex).Style.BackColor = Color.Red
                MsgBox("กำหนดราคา2 ต่ำกว่าทุนตลาด", MsgBoxStyle.Critical, "Send Information Message")
            Else
                Me.DGVItemDetails.Item(46, e.RowIndex).Style.BackColor = Color.Linen
            End If
        End If

    End Sub

    Public Sub vCalcItemLine(ByVal vLine As Integer)
        Dim vDocNo As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vColumn As Integer
        Dim vRow As Integer
        Dim i As Integer

        Dim vCheckItemCode As String
        Dim vMemCountCheck As Integer

        Dim vNowDate As Date
        Dim vAddDate As Date
        Dim vCheckUpdate As Date
        Dim vDocDate As String

        Dim vDOAmount As Double
        Dim vBillDisc As Double
        Dim vBillDiscAmount As Double
        Dim vAccCost As Double
        Dim vDisc1 As Double
        Dim vDiscAmount1 As Double
        Dim vAfterDiscAmount1 As Double


        On Error Resume Next

        If Me.TBDocNo.Text = "" Then
            Exit Sub
        End If

        vNowDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        vCheckUpdate = vb6.Day(Me.DTPUpdate.Value) & "/" & vb6.Month(Me.DTPUpdate.Value) & "/" & vb6.Year(Me.DTPUpdate.Value)

        vAddDate = vb6.DateAdd(DateInterval.Day, 1, vNowDate)

        If vCheckUpdate > vNowDate Then
            If vb6.Left(vCheckUpdate.Year, 2) = "20" Then
                vDocDate = vCheckUpdate
            Else
                vDocDate = vb6.Day(vCheckUpdate) & "/" & vb6.Month(vCheckUpdate) & "/" & vb6.Year(vCheckUpdate) - 543
            End If
        Else
            If vb6.Left(vAddDate.Year, 2) = "20" Then
                vDocDate = vAddDate
            Else
                vDocDate = vb6.Day(vAddDate) & "/" & vb6.Month(vAddDate) & "/" & vb6.Year(vAddDate) - 543
            End If
        End If

        vDocNo = Me.TBDocNo.Text

        If vMemIsCancel = 1 Then
            MsgBox("เอกสารเลขที่ " & vDocNo & " ถูกยกเลิกไปแล้ว ไม่สามารถแก้ไขข้อมูลได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Information Message")
            Me.DGVItemDetails.Item(1, vLine).Value = ""
            Me.TBDocNo.Focus()
            Exit Sub
        End If

        If vMemIsConfirm = 1 Then
            MsgBox("เอกสารเลขที่ " & vDocNo & " ถูกอนุมัติไปแล้ว ไม่สามารถแก้ไขข้อมูลได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Information Message")
            Me.DGVItemDetails.Item(1, vLine).Value = ""
            Me.TBDocNo.Focus()
            Exit Sub
        End If

        vColumn = Me.DGVItemDetails.CurrentCell.ColumnIndex
        vRow = Me.DGVItemDetails.CurrentCell.RowIndex
        vItemCode = Me.DGVItemDetails.Item(1, vLine).Value

        If vItemCode = "" Then
            Me.DGVItemDetails.Item(2, vLine).Value = ""
            Me.DGVItemDetails.Item(3, vLine).Value = ""
            Me.DGVItemDetails.Item(4, vLine).Value = ""
            Me.DGVItemDetails.Item(5, vLine).Value = ""
            Me.DGVItemDetails.Item(6, vLine).Value = ""
            Me.DGVItemDetails.Item(7, vLine).Value = ""
            Me.DGVItemDetails.Item(8, vLine).Value = ""
            Me.DGVItemDetails.Item(9, vLine).Value = ""
            Me.DGVItemDetails.Item(10, vLine).Value = ""
            Me.DGVItemDetails.Item(11, vLine).Value = ""
            Me.DGVItemDetails.Item(12, vLine).Value = ""
            Me.DGVItemDetails.Item(13, vLine).Value = ""
            Me.DGVItemDetails.Item(14, vLine).Value = ""
            Me.DGVItemDetails.Item(15, vLine).Value = ""
            Me.DGVItemDetails.Item(16, vLine).Value = ""
            Me.DGVItemDetails.Item(17, vLine).Value = ""
            Me.DGVItemDetails.Item(18, vLine).Value = ""
            Me.DGVItemDetails.Item(19, vLine).Value = ""
            Me.DGVItemDetails.Item(20, vLine).Value = ""
            Me.DGVItemDetails.Item(21, vLine).Value = ""
            Me.DGVItemDetails.Item(22, vLine).Value = ""
            Me.DGVItemDetails.Item(23, vLine).Value = ""
            Me.DGVItemDetails.Item(24, vLine).Value = ""
            Me.DGVItemDetails.Item(25, vLine).Value = ""
            Me.DGVItemDetails.Item(26, vLine).Value = ""
            Me.DGVItemDetails.Item(27, vLine).Value = ""
            Me.DGVItemDetails.Item(28, vLine).Value = ""
            Me.DGVItemDetails.Item(29, vLine).Value = ""
            Me.DGVItemDetails.Item(30, vLine).Value = ""
            Me.DGVItemDetails.Item(31, vLine).Value = ""
            Me.DGVItemDetails.Item(32, vLine).Value = ""
            Me.DGVItemDetails.Item(33, vLine).Value = ""
            Me.DGVItemDetails.Item(34, vLine).Value = ""
            Me.DGVItemDetails.Item(35, vLine).Value = ""
            Me.DGVItemDetails.Item(36, vLine).Value = ""
            Me.DGVItemDetails.Item(37, vLine).Value = ""
            Me.DGVItemDetails.Item(38, vLine).Value = ""
            Me.DGVItemDetails.Item(39, vLine).Value = ""
            Me.DGVItemDetails.Item(40, vLine).Value = ""
            Me.DGVItemDetails.Item(41, vLine).Value = ""
            Me.DGVItemDetails.Item(42, vLine).Value = ""
            Me.DGVItemDetails.Item(43, vLine).Value = ""
            Me.DGVItemDetails.Item(44, vLine).Value = ""
            Me.DGVItemDetails.Item(45, vLine).Value = ""
            Me.DGVItemDetails.Item(46, vLine).Value = ""
            Me.DGVItemDetails.Item(47, vLine).Value = ""
            Me.DGVItemDetails.Item(48, vLine).Value = ""
            Me.DGVItemDetails.Item(49, vLine).Value = ""
        End If

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
                    vItemName = dt.Rows(0).Item("itemname")
                    Me.DGVItemDetails.Item(2, vRow).Value = dt.Rows(0).Item("itemname")
                    Me.DGVItemDetails.Item(3, vRow).Value = dt.Rows(0).Item("unitcode")
                    Me.DGVItemDetails.Item(47, vRow).Value = vDocDate

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

        vCharStr = Me.DGVItemDetails.Item(4, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(4, vLine).Value = ""
                MsgBox("ช่อง D/O ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(4, vLine).Selected = True
                Exit Sub
            End If
        End If


        vCharStr = Me.DGVItemDetails.Item(5, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(5, vLine).Value = ""
                MsgBox("ช่องส่วนลดหน้าบิล% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(5, vLine).Selected = True
                Exit Sub
            End If
        End If


        vCharStr = Me.DGVItemDetails.Item(7, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(7, vLine).Value = ""
                MsgBox("ช่องส่วนลดตาม1% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(7, vLine).Selected = True
                Exit Sub
            End If
        End If


        vCharStr = Me.DGVItemDetails.Item(9, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(9, vLine).Value = ""
                MsgBox("ช่องส่วนลดตาม2% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(9, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(11, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(11, vLine).Value = ""
                MsgBox("ช่องส่วนลดตาม3% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(11, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(13, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(13, vLine).Value = ""
                MsgBox("ช่องส่วนลดตาม4% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(13, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(15, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(15, vLine).Value = ""
                MsgBox("ช่องส่วนลดRebate% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(15, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(17, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(17, vLine).Value = ""
                MsgBox("ช่องส่วนลดพิเศษ ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(17, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(19, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(19, vLine).Value = ""
                MsgBox("ช่องงบขาดทุน ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(19, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(21, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(21, vLine).Value = ""
                MsgBox("ช่องค่าขนส่งเข้า ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(21, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(22, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(22, vLine).Value = ""
                MsgBox("ช่องค่าขนส่งให้ลูกค้า ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(22, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(23, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(23, vLine).Value = ""
                MsgBox("ช่องโฆษณา% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(23, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(26, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(26, vLine).Value = ""
                MsgBox("ช่องค่าแรงค่าติดตั้ง ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(26, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(27, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(27, vLine).Value = ""
                MsgBox("ช่องบริการ ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(27, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(35, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(35, vLine).Value = ""
                MsgBox("ช่องSmartPoint% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(35, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(37, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(37, vLine).Value = ""
                MsgBox("ช่องเป้า ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(37, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(38, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(38, vLine).Value = ""
                MsgBox("ช่องของแถม ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(38, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(39, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(39, vLine).Value = ""
                MsgBox("ช่องคอม% ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(39, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(42, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(42, vLine).Value = ""
                MsgBox("ช่องราคา1เงินสดรับเอง ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(42, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(44, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(44, vLine).Value = ""
                MsgBox("ช่องราคา1เงินเชื่อรับเอง ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(44, vLine).Selected = True
                Exit Sub
            End If
        End If

        vCharStr = Me.DGVItemDetails.Item(46, vLine).Value
        If vCharStr <> "" Then
            Call vCheckNumber(vCharStr)
            If vIsNumber = 0 Then
                Me.DGVItemDetails.Item(46, vLine).Value = ""
                MsgBox("ช่องราคา2 ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.DGVItemDetails.Item(46, vLine).Selected = True
                Exit Sub
            End If
        End If

        Dim vDisc2 As Double
        Dim vDiscAmount2 As Double
        Dim vAfterDiscAmount2 As Double

        Dim vDisc3 As Double
        Dim vDiscAmount3 As Double
        Dim vAfterDiscAmount3 As Double

        Dim vDisc4 As Double
        Dim vDiscAmount4 As Double
        Dim vAfterDiscAmount4 As Double

        Dim vRebate As Double
        Dim vRebateAmount As Double
        Dim vAfterRebateAmount As Double

        Dim vNetCost As Double
        Dim vDiscSpecial As Double

        Dim vLose As Double
        Dim vLoseAmount As Double
        Dim vAfterLoseAmount As Double

        Dim vTransInAmount As Double
        Dim vTransOutAmount As Double
        Dim vAdvertise As Double
        Dim vAdvertiseAmount As Double
        Dim vAfterAdvertiseAmount As Double

        Dim vVatCost As Double

        Dim vInstallAmount As Double
        Dim vServiceAmount As Double
        Dim vMarketCost As Double

        Dim vRelateStockPercent As Double
        Dim vRelateStockAmount As Double

        Dim vSmartPoint As Double
        Dim vSmartPointAmount As Double
        Dim vAfterSmartPointAmount As Double

        Dim vCashProfit As Double
        Dim vCashProfitAmount As Double
        Dim vAfterCashProfitAmount As Double

        Dim vCreditProfit As Double
        Dim vCreditProfitAmount As Double
        Dim vAfterCreditProfitAmount As Double

        Dim vTotalPrice As Double
        Dim vCashPriceOwn1 As Double
        Dim vCashPriceSend1 As Double
        Dim vCreditPriceOwn1 As Double
        Dim vCreditPriceSend1 As Double
        Dim vSalePrice2 As Double

        Dim vTargetAmount As Double
        Dim vPremiumAmount As Double

        Dim vComm As Double
        Dim vCommAmount As Double
        Dim vAfterCommAmount As Double

        Dim vBaseProfitPercent As Double
        Dim vBaseProfit As Double
        Dim vRelateStockPercentShow As Double

        vDOAmount = Me.DGVItemDetails.Item(4, vLine).Value
        If Me.DGVItemDetails.Item(5, vLine).Value <> "" Then
            vBillDisc = Me.DGVItemDetails.Item(5, vLine).Value
        Else
            vBillDisc = 0
        End If
        vBillDiscAmount = (vDOAmount * (vBillDisc / 100))
        vAccCost = vDOAmount - vBillDiscAmount

        If Me.DGVItemDetails.Item(7, vLine).Value <> "" Then
            vDisc1 = Me.DGVItemDetails.Item(7, vLine).Value
        Else
            vDisc1 = 0
        End If
        vDiscAmount1 = (vAccCost * (vDisc1 / 100))
        vAfterDiscAmount1 = vAccCost - vDiscAmount1

        If Me.DGVItemDetails.Item(9, vLine).Value <> "" Then
            vDisc2 = Me.DGVItemDetails.Item(9, vLine).Value
        Else
            vDisc2 = 0
        End If

        vDiscAmount2 = (vAfterDiscAmount1 * (vDisc2 / 100))
        vAfterDiscAmount2 = vAfterDiscAmount1 - vDiscAmount2

        If Me.DGVItemDetails.Item(11, vLine).Value <> "" Then
            vDisc3 = Me.DGVItemDetails.Item(11, vLine).Value
        Else
            vDisc3 = 0
        End If

        vDiscAmount3 = (vAfterDiscAmount2 * (vDisc3 / 100))
        vAfterDiscAmount3 = vAfterDiscAmount2 - vDiscAmount3

        If Me.DGVItemDetails.Item(13, vLine).Value <> "" Then
            vDisc4 = Me.DGVItemDetails.Item(13, vLine).Value
        Else
            vDisc4 = 0
        End If
        vDiscAmount4 = (vAfterDiscAmount3 * (vDisc4 / 100))
        vAfterDiscAmount4 = vAfterDiscAmount3 - vDiscAmount4

        If Me.DGVItemDetails.Item(15, vLine).Value <> "" Then
            vRebate = Me.DGVItemDetails.Item(15, vLine).Value
        Else
            vRebate = 0
        End If
        vRebateAmount = (vAfterDiscAmount4 * (vRebate / 100))
        vAfterRebateAmount = vAfterDiscAmount4 - vRebateAmount

        If Me.DGVItemDetails.Item(17, vLine).Value <> "" Then
            vDiscSpecial = Me.DGVItemDetails.Item(17, vLine).Value
        Else
            vDiscSpecial = 0
        End If
        vNetCost = vAfterRebateAmount - vDiscSpecial

        If Me.DGVItemDetails.Item(19, vLine).Value <> "" Then
            vLose = Me.DGVItemDetails.Item(19, vLine).Value
        Else
            vLose = 0
        End If

        vLoseAmount = (vNetCost * (vLose / 100))
        vAfterLoseAmount = vNetCost - vLoseAmount

        If Me.DGVItemDetails.Item(21, vLine).Value <> "" Then
            vTransInAmount = Me.DGVItemDetails.Item(21, vLine).Value
        Else
            vTransInAmount = 0
        End If

        If Me.DGVItemDetails.Item(22, vLine).Value <> "" Then
            vTransOutAmount = Me.DGVItemDetails.Item(22, vLine).Value
        Else
            vTransOutAmount = 0
        End If

        If Me.DGVItemDetails.Item(23, vLine).Value <> "" Then
            vAdvertise = Me.DGVItemDetails.Item(23, vLine).Value
        Else
            vAdvertise = 0
        End If

        vAdvertiseAmount = (vAfterLoseAmount * (vAdvertise / 100))
        vAfterAdvertiseAmount = (vAfterLoseAmount + vTransInAmount + vTransOutAmount) + vAdvertiseAmount

        vVatCost = (vAfterAdvertiseAmount * 0.07) + vAfterAdvertiseAmount

        If Me.DGVItemDetails.Item(26, vLine).Value <> "" Then
            vInstallAmount = Me.DGVItemDetails.Item(26, vLine).Value
        Else
            vInstallAmount = 0
        End If

        If Me.DGVItemDetails.Item(27, vLine).Value <> "" Then
            vServiceAmount = Me.DGVItemDetails.Item(27, vLine).Value
        Else
            vServiceAmount = 0
        End If

        vMarketCost = vVatCost + vInstallAmount + vServiceAmount

        If Me.DGVItemDetails.Item(42, vLine).Value <> "" Then
            vCashPriceOwn1 = Me.DGVItemDetails.Item(42, vLine).Value
        Else
            vCashPriceOwn1 = 0
        End If

        vCashPriceSend1 = vCashPriceOwn1

        If Me.DGVItemDetails.Item(44, vLine).Value <> "" Then
            vCreditPriceOwn1 = Me.DGVItemDetails.Item(44, vLine).Value
        Else
            vCreditPriceOwn1 = 0
        End If

        vCreditPriceSend1 = vCreditPriceOwn1

        If Me.DGVItemDetails.Item(35, vLine).Value <> "" Then
            vSmartPoint = Me.DGVItemDetails.Item(35, vLine).Value
        Else
            vSmartPoint = 0
        End If

        vSmartPointAmount = vSmartPoint
        vAfterSmartPointAmount = (vSmartPointAmount * vCashPriceOwn1) / 100

        If Me.DGVItemDetails.Item(37, vLine).Value <> "" Then
            vTargetAmount = Me.DGVItemDetails.Item(37, vLine).Value
        Else
            vTargetAmount = 0
        End If

        If Me.DGVItemDetails.Item(38, vLine).Value <> "" Then
            vPremiumAmount = Me.DGVItemDetails.Item(38, vLine).Value
        Else
            vPremiumAmount = 0
        End If

        If Me.DGVItemDetails.Item(39, vLine).Value <> "" Then
            vComm = Me.DGVItemDetails.Item(39, vLine).Value
        Else
            vComm = 0
        End If

        vCommAmount = vComm / 100
        vAfterCommAmount = vCommAmount * vCashPriceOwn1

        vBaseProfitPercent = ((vCashPriceOwn1 - vAfterCommAmount - vPremiumAmount - vAfterSmartPointAmount - vTargetAmount - vMarketCost) / vMarketCost) * 100
        vBaseProfit = ((vCashPriceOwn1 - vAfterCommAmount - vPremiumAmount - vAfterSmartPointAmount - vTargetAmount - vMarketCost) / vMarketCost)

        If vBaseProfit < 0.01 Then
            vRelateStockPercent = 0
        ElseIf vBaseProfit >= 0.01 And vBaseProfit < 0.05 Then
            vRelateStockPercent = vBaseProfit / 2
        ElseIf vBaseProfit >= 0.05 And vBaseProfit < 0.06 Then
            vRelateStockPercent = vBaseProfit * 0.6
        Else
            vRelateStockPercent = 0.035
        End If

        vRelateStockPercentShow = vRelateStockPercent * 100
        vRelateStockAmount = vRelateStockPercent * vMarketCost

        vAfterCashProfitAmount = vCashPriceOwn1 - vAfterCommAmount - vPremiumAmount - vAfterSmartPointAmount - vTargetAmount - vRelateStockAmount - vMarketCost
        vCashProfit = (vAfterCashProfitAmount / vMarketCost) * 100

        vAfterCreditProfitAmount = vCreditPriceOwn1 - vMarketCost - vRelateStockAmount
        vCreditProfit = (vAfterCreditProfitAmount / vMarketCost) * 100

        vTotalPrice = vMarketCost + vRelateStockAmount + vAfterCashProfitAmount + vAfterSmartPointAmount + vTargetAmount + vPremiumAmount + vAfterCommAmount

        vItemName = Me.DGVItemDetails.Item(2, vLine).Value
        If Me.DGVItemDetails.Item(46, vLine).Value <> "" Then
            vSalePrice2 = Me.DGVItemDetails.Item(46, vLine).Value
        Else
            vSalePrice2 = 0
        End If


        If vItemName <> "Nothing" And vItemName <> "" Then
            If vDOAmount <> 0 Then
                Me.DGVItemDetails.Item(4, vLine).Value = Format(vDOAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(4, vLine).Value = ""
            End If

            If vBillDisc = 0 Then
                Me.DGVItemDetails.Item(5, vLine).Value = ""
            End If

            If vAccCost <> 0 Then
                Me.DGVItemDetails.Item(6, vLine).Value = Format(vAccCost, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(6, vLine).Value = ""
            End If

            If vDisc1 = 0 Then
                Me.DGVItemDetails.Item(7, vLine).Value = ""
            End If

            If vAfterDiscAmount1 <> 0 Then
                Me.DGVItemDetails.Item(8, vLine).Value = Format(vAfterDiscAmount1, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(8, vLine).Value = ""
            End If

            If vDisc2 = 0 Then
                Me.DGVItemDetails.Item(9, vLine).Value = ""
            End If

            If vAfterDiscAmount2 <> 0 Then
                Me.DGVItemDetails.Item(10, vLine).Value = Format(vAfterDiscAmount2, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(10, vLine).Value = ""
            End If

            If vDisc3 = 0 Then
                Me.DGVItemDetails.Item(11, vLine).Value = ""
            End If

            If vAfterDiscAmount3 <> 0 Then
                Me.DGVItemDetails.Item(12, vLine).Value = Format(vAfterDiscAmount3, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(12, vLine).Value = ""
            End If

            If vDisc4 = 0 Then
                Me.DGVItemDetails.Item(13, vLine).Value = ""
            End If

            If vAfterDiscAmount4 <> 0 Then
                Me.DGVItemDetails.Item(14, vLine).Value = Format(vAfterDiscAmount4, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(14, vLine).Value = ""
            End If

            If vRebate = 0 Then
                Me.DGVItemDetails.Item(15, vLine).Value = ""
            End If

            If vAfterRebateAmount <> 0 Then
                Me.DGVItemDetails.Item(16, vLine).Value = Format(vAfterRebateAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(16, vLine).Value = ""
            End If

            If vDiscSpecial <> 0 Then
                Me.DGVItemDetails.Item(17, vLine).Value = Format(vDiscSpecial, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(17, vLine).Value = ""
            End If

            If vNetCost <> 0 Then
                Me.DGVItemDetails.Item(18, vLine).Value = Format(vNetCost, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(18, vLine).Value = ""
            End If

            If vLose = 0 Then
                Me.DGVItemDetails.Item(19, vLine).Value = ""
            End If

            If vAfterLoseAmount <> 0 Then
                Me.DGVItemDetails.Item(20, vLine).Value = Format(vAfterLoseAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(20, vLine).Value = ""
            End If

            If vTransInAmount <> 0 Then
                Me.DGVItemDetails.Item(21, vLine).Value = Format(vTransInAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(21, vLine).Value = ""
            End If

            If vTransOutAmount <> 0 Then
                Me.DGVItemDetails.Item(22, vLine).Value = Format(vTransOutAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(22, vLine).Value = ""
            End If

            If vAdvertise = 0 Then
                Me.DGVItemDetails.Item(23, vLine).Value = ""
            End If

            If vAfterAdvertiseAmount <> 0 Then
                Me.DGVItemDetails.Item(24, vLine).Value = Format(vAfterAdvertiseAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(24, vLine).Value = ""
            End If

            If vVatCost <> 0 Then
                Me.DGVItemDetails.Item(25, vLine).Value = Format(vVatCost, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(25, vLine).Value = ""
            End If

            If vInstallAmount <> 0 Then
                Me.DGVItemDetails.Item(26, vLine).Value = Format(vInstallAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(26, vLine).Value = ""
            End If

            If vServiceAmount <> 0 Then
                Me.DGVItemDetails.Item(27, vLine).Value = Format(vServiceAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(27, vLine).Value = ""
            End If

            If vMarketCost <> 0 Then
                Me.DGVItemDetails.Item(28, vLine).Value = Format(vMarketCost, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(28, vLine).Value = ""
            End If

            If vRelateStockPercentShow <> 0 Then
                Me.DGVItemDetails.Item(29, vLine).Value = Format(vRelateStockPercentShow, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(29, vLine).Value = ""
            End If

            If vRelateStockAmount <> 0 Then
                Me.DGVItemDetails.Item(30, vLine).Value = Format(vRelateStockAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(30, vLine).Value = ""
            End If

            If vCashProfit <> 0 Then
                Me.DGVItemDetails.Item(31, vLine).Value = Format(vCashProfit, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(31, vLine).Value = ""
            End If

            If vAfterCashProfitAmount <> 0 Then
                Me.DGVItemDetails.Item(32, vLine).Value = Format(vAfterCashProfitAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(32, vLine).Value = ""
            End If

            If vCreditProfit <> 0 Then
                Me.DGVItemDetails.Item(33, vLine).Value = Format(vCreditProfit, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(33, vLine).Value = ""
            End If

            If vAfterCreditProfitAmount <> 0 Then
                Me.DGVItemDetails.Item(34, vLine).Value = Format(vAfterCreditProfitAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(34, vLine).Value = ""
            End If

            If vSmartPoint = 0 Then
                Me.DGVItemDetails.Item(35, vLine).Value = ""
            End If

            If vAfterSmartPointAmount <> 0 Then
                Me.DGVItemDetails.Item(36, vLine).Value = Format(vAfterSmartPointAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(36, vLine).Value = ""
            End If

            If vTargetAmount <> 0 Then
                Me.DGVItemDetails.Item(37, vLine).Value = Format(vTargetAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(37, vLine).Value = ""
            End If

            If vPremiumAmount <> 0 Then
                Me.DGVItemDetails.Item(38, vLine).Value = Format(vPremiumAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(38, vLine).Value = ""
            End If

            If vComm = 0 Then
                Me.DGVItemDetails.Item(39, vLine).Value = ""
            End If

            If vAfterCommAmount <> 0 Then
                Me.DGVItemDetails.Item(40, vLine).Value = Format(vAfterCommAmount, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(40, vLine).Value = ""
            End If

            If vTotalPrice <> 0 Then
                Me.DGVItemDetails.Item(41, vLine).Value = Format(vTotalPrice, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(41, vLine).Value = ""
            End If

            If vCashPriceOwn1 <> 0 Then
                Me.DGVItemDetails.Item(42, vLine).Value = Format(vCashPriceOwn1, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(42, vLine).Value = ""
            End If

            If vCashPriceSend1 <> 0 Then
                Me.DGVItemDetails.Item(43, vLine).Value = Format(vCashPriceSend1, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(43, vLine).Value = ""
            End If

            If vCreditPriceOwn1 <> 0 Then
                Me.DGVItemDetails.Item(44, vLine).Value = Format(vCreditPriceOwn1, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(44, vLine).Value = ""
            End If

            If vCreditPriceSend1 <> 0 Then
                Me.DGVItemDetails.Item(45, vLine).Value = Format(vCreditPriceSend1, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(45, vLine).Value = ""
            End If

            If vSalePrice2 <> 0 Then
                Me.DGVItemDetails.Item(46, vLine).Value = Format(vSalePrice2, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(46, vLine).Value = ""
            End If

            If vBaseProfitPercent <> 0 Then
                Me.DGVItemDetails.Item(49, vLine).Value = Format(vBaseProfitPercent, "##,##0.000")
            Else
                Me.DGVItemDetails.Item(49, vLine).Value = ""
            End If
        End If
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

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Me.Close()
    End Sub

    Private Sub BTNSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelect.Click
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

        If Me.ListViewSearch.Items.Count > 0 Then
            For i = 0 To Me.ListViewSearch.Items.Count - 1
                If Me.ListViewSearch.Items(i).Checked = True Then

                    vItemCode = Me.ListViewSearch.Items(i).SubItems(1).Text
                    vItemName = Me.ListViewSearch.Items(i).SubItems(2).Text
                    vUnitCode = Me.ListViewSearch.Items(i).SubItems(3).Text

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
                                GoTo Line1
                            End If
                        Next
                    End If

                End If
Line1:
            Next


            For a = 0 To Me.ListViewSearch.Items.Count - 1
                Me.ListViewSearch.Items(a).Checked = False
            Next

            Me.PNSearch.Visible = False
        End If
    End Sub

    Private Sub BTNClickSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClickSearch.Click
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
        Me.ListViewSearch.Items.Clear()
        vSearch = Me.TBSearch.Text
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

                vListItem = Me.ListViewSearch.Items.Add(n)
                vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("itemcode")
                vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("itemname")
                vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("unitcode")
                vListItem.SubItems.Add(3).Text = Format(vCashSalePrice, "##,##0.000")
                vListItem.SubItems.Add(4).Text = Format(vCreditSalePrice, "##,##0.000")

            Next
        End If
    End Sub

    Private Sub CBItemSaleLose_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBItemSaleLose.CheckedChanged
        If Me.CBItemSaleLose.Checked = True Then
            Me.CBNotAddPriceStructure.Checked = False
        End If
    End Sub

    Private Sub CBNotAddPriceStructure_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBNotAddPriceStructure.CheckedChanged
        If Me.CBNotAddPriceStructure.Checked = True Then
            Me.CBItemSaleLose.Checked = False
        End If
    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        Me.PNSearch.Visible = False
        Me.TBDocNo.Focus()
    End Sub

    Private Sub BTNSearchOldData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchOldData.Click
        Me.PNSearch.Visible = True
        Me.TBSearch.Focus()
    End Sub

    Private Sub TBSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call SearchItem()
        End If
    End Sub

    Private Sub TBSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearch.TextChanged

    End Sub

    Public Function CalcAmount(ByVal vPriceSetAmount As Double, ByVal vPercent As Double) As Double
        CalcAmount = (vPriceSetAmount * vPercent) / 100
    End Function

    Public Function CalcAmountAfterAdd(ByVal vPriceSetAmount As Double, ByVal vPercent As Double) As Double
        CalcAmountAfterAdd = vPriceSetAmount + (vPriceSetAmount * vPercent) / 100
    End Function
    Public Function CalcAmountAfterDelete(ByVal vPriceSetAmount As Double, ByVal vPercent As Double) As Double
        CalcAmountAfterDelete = vPriceSetAmount - (vPriceSetAmount * vPercent) / 100
    End Function

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vDocno As String
        Dim vDocDate As String
        Dim vTargetProfit As String
        Dim vProfit As String
        Dim vSmartPoint As String
        Dim vMemberDiscount As String
        Dim vFileDataSource As String
        Dim vPathFile As String
        '------------------------------------------
        Dim vItemCode As String
        Dim vItemName As String
        Dim vSaleUnitCode As String
        Dim vDO As Double
        Dim vPriceSet As Double
        Dim vDiscountBillWord As String
        Dim vDiscountBill1 As Double
        Dim vDiscountBillAmount As Double
        Dim vAccCost As Double
        Dim vDiscountFollow1Word As String
        Dim vDiscountFollow11 As Double
        Dim vDiscountFollow1Amount As Double
        Dim vDiscountFollow1After As Double
        Dim vDiscountFollow2Word As String
        Dim vDiscountFollow21 As Double
        Dim vDiscountFollow2Amount As Double
        Dim vDiscountFollow2After As Double
        Dim vDiscountFollow3Word As String
        Dim vDiscountFollow31 As Double
        Dim vDiscountFollow3Amount As Double
        Dim vDiscountFollow3After As Double
        Dim vDiscountFollow4Word As String
        Dim vDiscountFollow41 As Double
        Dim vDiscountFollow4Amount As Double
        Dim vDiscountFollow4After As Double
        Dim vDiscountRebateWord As String
        Dim vDiscountRebate1 As Double
        Dim vDiscountRebateAmount As Double
        Dim vDiscountRebateAfter As Double
        Dim vDiscountSpecialWord As String
        Dim vDiscountSpecial1 As Double
        Dim vDiscountSpecialAmount As Double
        Dim vNetCost As Double
        Dim vLossBudgetWord As String
        Dim vLossBudget1 As Double
        Dim vLossBudgetAmount As Double
        Dim vLossBudgetAfter As Double
        Dim vTransferInWord As String
        Dim vTransferIn1 As Double
        Dim vTransferOutWord As String
        Dim vTransferOut1 As Double
        Dim vAdvertiseWord As String
        Dim vAdvertise1 As Double
        Dim vAdvertiseAmount As Double
        Dim vAdvertiseAfter As Double
        Dim vVatCost As Double
        Dim vVatAmount As Double
        Dim vSetupWord As String
        Dim vSetupAmount As Double
        Dim vServiceWord As String
        Dim vServiceAmount As Double
        Dim vMarketCost As Double
        Dim vPointWord As String
        Dim vPoint1 As Double
        Dim vPointAmount As Double
        Dim vPointAfter As Double
        Dim vTargetWord As String
        Dim vTargetAmount As Double
        Dim vPremiumWord As String
        Dim vPremiumAmount As Double
        Dim vCommissionWord As String
        Dim vCommission1 As Double
        Dim vCommissionAmount As Double
        Dim vCommissionAfter As Double
        Dim vGrossProfitPercent As String
        Dim vGrossProfitAmount As Double
        Dim vInterestStockPercent As String
        Dim vInterestStockAmount As Double
        Dim vProfitPercent As String
        Dim vProfitAmount As Double
        Dim vProfitPercent_W As String
        Dim vProfitAmount_W As Double
        Dim vMyDescription As String
        Dim vMyDescriptionSub As String
        Dim vTransferInAfter As Double
        Dim vTransferOutAfter As Double
        Dim vVatWord As String
        Dim vSetupAfter As Double
        Dim vTargetAfter As Double
        Dim vPremiumAfter As Double
        '--------------------------------------------------------------

        '---------------------------------------------------------------
        Dim i As Integer
        Dim n As Integer


        Me.PB101.Value = 1
        If vIsOpen = 0 Then
            vQuery = "exec dbo.USP_PS_NewDocno"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "NewDocno")
            vdt = ds.Tables("NewDocno")

            vDocno = vdt.Rows(0).Item("newdocno")

            If vb6.Left(vb6.Year(Me.DTPDocDate.Value), 2) = "20" Then
                vDocDate = Me.DTPDocDate.Text
            Else
                vDocDate = vb6.Day(Me.DTPDocDate.Value) & "/" & vb6.Month(Me.DTPDocDate.Value) & "/" & vb6.Year(Me.DTPDocDate.Value) - 543
            End If

        Else
            vDocno = Me.TBDocNo.Text
            If vb6.Left(vb6.Year(Me.DTPDocDate.Value), 2) = "20" Then
                vDocDate = Me.DTPDocDate.Text
            Else
                vDocDate = vb6.Day(Me.DTPDocDate.Value) & "/" & vb6.Month(Me.DTPDocDate.Value) & "/" & vb6.Year(Me.DTPDocDate.Value) - 543
            End If
        End If

        If vMemIsCancel = 1 Then
            MsgBox("เอกสารเลขที่ " & vDocno & " ถูกยกเลิกไปแล้ว ไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Information Message")
            Me.TBDocNo.Focus()
            Exit Sub
        End If

        If vMemIsConfirm = 1 Then
            MsgBox("เอกสารเลขที่ " & vDocno & " ถูกอนุมัติไปแล้ว ไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Information Message")
            Me.TBDocNo.Focus()
            Exit Sub
        End If

        If Me.DGVItemDetails.Item(1, 0).Value = "" And Me.DGVItemDetails.Item(2, 0).Value = "" Then
            MsgBox("ตรวจสอบรายการสินค้าว่า กรอกข้อมูลรหัส ชื่อ ราคาเงินสดแล้วหรือยัง ไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Information Message")
            Me.TBDocNo.Focus()
            Exit Sub
        End If

        For n = 0 To Me.DGVItemDetails.RowCount - 1
            If Me.DGVItemDetails.Item(1, n).Value <> "" Then
                If Me.DGVItemDetails.Item(4, n).Value = "" Then
                    MsgBox("ไม่ได้กำหนด D/O", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(4, n).Selected = True
                    Exit Sub
                End If

                If Me.DGVItemDetails.Item(35, n).Value = "" Then
                    MsgBox("ไม่ได้กำหนด SmartPoint", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(35, n).Selected = True
                    Exit Sub
                End If

                If Me.DGVItemDetails.Item(42, n).Value = "" Then
                    MsgBox("ไม่ได้กำหนดราคาขายเงินสด", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(42, n).Selected = True
                    Exit Sub
                End If

                If Me.DGVItemDetails.Item(44, n).Value = "" Then
                    MsgBox("ไม่ได้กำหนดราคาขายเงินเชื่อ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(44, n).Selected = True
                    Exit Sub
                End If

                If Me.DGVItemDetails.Item(46, n).Value = "" Then
                    MsgBox("ไม่ได้กำหนดราคาขาย 2", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(46, n).Selected = True
                    Exit Sub
                End If

                If Me.DGVItemDetails.Item(47, n).Value = "" Then
                    MsgBox("ไม่ได้กำหนด วันที่จะปรับราคา", MsgBoxStyle.Critical, "Send Error Message")
                    Me.DGVItemDetails.Item(47, n).Selected = True
                    Exit Sub
                End If
            End If
        Next


        Try
            vQuery = "begin tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            If Me.TBTarget.Text <> "" Then
                vTargetProfit = Me.TBTarget.Text
            Else
                vTargetProfit = 0
            End If
            If Me.TBTargetAverage.Text <> "" Then
                vProfit = Me.TBTargetAverage.Text
            Else
                vProfit = 0
            End If
            If Me.TBSmartPoint.Text <> "" Then
                vSmartPoint = Me.TBSmartPoint.Text
            Else
                vSmartPoint = 0
            End If
            If Me.TBSmartPoint.Text <> "" Then
                vMemberDiscount = 0
            Else
                vMemberDiscount = 0
            End If
            vMyDescription = Me.TBMyDescription.Text
            vPathFile = ""
            vFileDataSource = ""

            vQuery = "exec dbo.USP_PS_InsertPriceStructureSet1 '" & vDocno & "','" & vDocDate & "','" & vTargetProfit & "','" & vProfit & "','" & vSmartPoint & "','" & vMemberDiscount & "','" & vFileDataSource & "','" & vPathFile & "','" & vMyDescription & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            Me.PB101.Maximum = Me.DGVItemDetails.RowCount - 2

            For i = 0 To Me.DGVItemDetails.RowCount - 2

                vItemCode = Trim(Me.DGVItemDetails.Rows(i).Cells(1).Value)
                If vItemCode <> "" Then
                    vItemName = Trim(Me.DGVItemDetails.Rows(i).Cells(2).Value)
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(3).Value) Then
                        vSaleUnitCode = Trim(Me.DGVItemDetails.Rows(i).Cells(3).Value)
                    Else
                        vSaleUnitCode = ""
                    End If
                    'MsgBox(Me.DGVItemDetails.Rows(i).Cells(4).Value)
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(4).Value) Then
                        vDO = Trim(Me.DGVItemDetails.Rows(i).Cells(4).Value)
                    Else
                        vDO = 0
                    End If
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(42).Value) Then
                        vPriceSet = Me.DGVItemDetails.Rows(i).Cells(42).Value
                    Else
                        vPriceSet = 0
                    End If
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(5).Value) And Me.DGVItemDetails.Rows(i).Cells(5).Value <> "" Then
                        vDiscountBillWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(5).Value, Double)), String) & "%"
                        vDiscountBill1 = (CType(Me.DGVItemDetails.Rows(i).Cells(5).Value, Double))
                    Else
                        vDiscountBillWord = ""
                        vDiscountBill1 = 0
                    End If
                    If vDiscountBillWord = "0%" Then
                        vDiscountBillWord = ""
                    End If

                    vDiscountBillAmount = CalcAmount(vDO, vDiscountBill1)
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(6).Value) And Me.DGVItemDetails.Rows(i).Cells(6).Value <> "" Then
                        vAccCost = Me.DGVItemDetails.Rows(i).Cells(6).Value
                    Else
                        vAccCost = 0
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(7).Value) And Me.DGVItemDetails.Rows(i).Cells(7).Value <> "" Then
                        vDiscountFollow1Word = CType((CType(Me.DGVItemDetails.Rows(i).Cells(7).Value, Double)), String) & "%"
                        vDiscountFollow11 = (CType(Me.DGVItemDetails.Rows(i).Cells(7).Value, Double))
                    Else
                        vDiscountFollow1Word = ""
                        vDiscountFollow11 = 0
                    End If
                    If vDiscountFollow1Word = "0%" Then
                        vDiscountFollow1Word = ""
                    End If

                    vDiscountFollow1Amount = CalcAmount(vAccCost, vDiscountFollow11)
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(8).Value) And Me.DGVItemDetails.Rows(i).Cells(8).Value <> "" Then
                        vDiscountFollow1After = Me.DGVItemDetails.Rows(i).Cells(8).Value
                    Else
                        vDiscountFollow1After = 0
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(9).Value) And Me.DGVItemDetails.Rows(i).Cells(9).Value <> "" Then
                        vDiscountFollow2Word = CType((CType(Me.DGVItemDetails.Rows(i).Cells(9).Value, Double)), String) & "%"
                        vDiscountFollow21 = (CType(Me.DGVItemDetails.Rows(i).Cells(9).Value, Double))
                    Else
                        vDiscountFollow2Word = ""
                        vDiscountFollow21 = 0
                    End If
                    If vDiscountFollow2Word = "0%" Then
                        vDiscountFollow2Word = ""
                    End If

                    vDiscountFollow2Amount = CalcAmount(vDiscountFollow1After, vDiscountFollow21)
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(10).Value) And Me.DGVItemDetails.Rows(i).Cells(10).Value <> "" Then
                        vDiscountFollow2After = Me.DGVItemDetails.Rows(i).Cells(10).Value
                    Else
                        vDiscountFollow2After = 0
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(11).Value) And Me.DGVItemDetails.Rows(i).Cells(11).Value <> "" Then
                        vDiscountFollow3Word = CType((CType(Me.DGVItemDetails.Rows(i).Cells(11).Value, Double)), String) & "%"
                        vDiscountFollow31 = (CType(Me.DGVItemDetails.Rows(i).Cells(11).Value, Double))
                    Else
                        vDiscountFollow3Word = ""
                        vDiscountFollow31 = 0
                    End If
                    If vDiscountFollow3Word = "0%" Then
                        vDiscountFollow3Word = ""
                    End If

                    vDiscountFollow3Amount = CalcAmount(vDiscountFollow2After, vDiscountFollow31)
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(12).Value) And Me.DGVItemDetails.Rows(i).Cells(12).Value <> "" Then
                        vDiscountFollow3After = Me.DGVItemDetails.Rows(i).Cells(12).Value
                    Else
                        vDiscountFollow3After = 0
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(13).Value) And Me.DGVItemDetails.Rows(i).Cells(13).Value <> "" Then
                        vDiscountFollow4Word = CType((CType(Me.DGVItemDetails.Rows(i).Cells(13).Value, Double)), String) & "%"
                        vDiscountFollow41 = (CType(Me.DGVItemDetails.Rows(i).Cells(13).Value, Double))
                    Else
                        vDiscountFollow4Word = ""
                        vDiscountFollow41 = 0
                    End If
                    If vDiscountFollow4Word = "0%" Then
                        vDiscountFollow4Word = ""
                    End If

                    vDiscountFollow4Amount = CalcAmount(vDiscountFollow3After, vDiscountFollow41)
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(14).Value) And Me.DGVItemDetails.Rows(i).Cells(14).Value <> "" Then
                        vDiscountFollow4After = Me.DGVItemDetails.Rows(i).Cells(14).Value
                    Else
                        vDiscountFollow4After = 0
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(15).Value) And Me.DGVItemDetails.Rows(i).Cells(15).Value <> "" Then
                        vDiscountRebateWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(15).Value, Double)), String) & "%"
                        vDiscountRebate1 = (CType(Me.DGVItemDetails.Rows(i).Cells(15).Value, Double))
                    Else
                        vDiscountRebateWord = ""
                        vDiscountRebate1 = 0
                    End If
                    If vDiscountRebateWord = "0%" Then
                        vDiscountRebateWord = ""
                    End If

                    vDiscountRebateAmount = CalcAmount(vDiscountFollow4After, vDiscountRebate1)
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(16).Value) And Me.DGVItemDetails.Rows(i).Cells(16).Value <> "" Then
                        vDiscountRebateAfter = Me.DGVItemDetails.Rows(i).Cells(16).Value
                    Else
                        vDiscountRebateAfter = 0
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(17).Value) And Me.DGVItemDetails.Rows(i).Cells(17).Value <> "" Then
                        vDiscountSpecialWord = Me.DGVItemDetails.Rows(i).Cells(17).Value
                        vDiscountSpecial1 = Me.DGVItemDetails.Rows(i).Cells(17).Value
                    Else
                        vDiscountSpecialWord = 0
                        vDiscountSpecial1 = 0
                    End If

                    vDiscountSpecialAmount = vDiscountSpecial1
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(18).Value) And Me.DGVItemDetails.Rows(i).Cells(18).Value <> "" Then
                        vNetCost = Me.DGVItemDetails.Rows(i).Cells(18).Value
                    Else
                        vNetCost = 0
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(19).Value) And Me.DGVItemDetails.Rows(i).Cells(19).Value <> "" Then
                        vLossBudgetWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(19).Value, Double)), String) & "%"
                        vLossBudget1 = (CType(Me.DGVItemDetails.Rows(i).Cells(19).Value, Double))
                    Else
                        vLossBudgetWord = ""
                        vLossBudget1 = 0
                    End If
                    If vLossBudgetWord = "0%" Then
                        vLossBudgetWord = ""
                    End If

                    vLossBudgetAmount = CalcAmount(vNetCost, vLossBudget1)
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(20).Value) And Me.DGVItemDetails.Rows(i).Cells(20).Value <> "" Then
                        vLossBudgetAfter = Me.DGVItemDetails.Rows(i).Cells(20).Value
                    Else
                        vLossBudgetAfter = 0
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(21).Value) And Me.DGVItemDetails.Rows(i).Cells(21).Value <> "" Then
                        vTransferInWord = Me.DGVItemDetails.Rows(i).Cells(21).Value
                        vTransferIn1 = Me.DGVItemDetails.Rows(i).Cells(21).Value
                        vTransferInAfter = vLossBudgetAfter + vTransferIn1
                    Else
                        vTransferIn1 = 0
                        vTransferInWord = ""
                        vTransferInAfter = vLossBudgetAfter + vTransferIn1
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(22).Value) And Me.DGVItemDetails.Rows(i).Cells(22).Value <> "" Then
                        vTransferOutWord = Me.DGVItemDetails.Rows(i).Cells(22).Value
                        vTransferOut1 = Me.DGVItemDetails.Rows(i).Cells(22).Value
                        vTransferOutAfter = vTransferInAfter + vTransferOut1
                    Else
                        vTransferOutWord = ""
                        vTransferOut1 = 0
                        vTransferOutAfter = vTransferInAfter + vTransferOut1
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(23).Value) And Me.DGVItemDetails.Rows(i).Cells(23).Value <> "" Then
                        vAdvertiseWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(23).Value, Double)), String) & "%"
                        vAdvertise1 = (CType(Me.DGVItemDetails.Rows(i).Cells(23).Value, Double))
                    Else
                        vAdvertiseWord = ""
                        vAdvertise1 = 0
                    End If
                    If vAdvertiseWord = "0%" Then
                        vAdvertiseWord = ""
                    End If

                    vAdvertiseAmount = CalcAmount(vTransferOutAfter, vAdvertise1)
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(24).Value) And Me.DGVItemDetails.Rows(i).Cells(24).Value <> "" Then
                        vAdvertiseAfter = Me.DGVItemDetails.Rows(i).Cells(24).Value
                    Else
                        vAdvertiseAfter = 0
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(25).Value) And Me.DGVItemDetails.Rows(i).Cells(25).Value <> "" Then
                        vVatCost = Me.DGVItemDetails.Rows(i).Cells(25).Value
                        vVatAmount = (vAdvertiseAfter * 7) / 100
                        vVatWord = "7%"
                    Else
                        vVatCost = 0
                        vVatAmount = 0
                        vVatWord = 0
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(26).Value) And Me.DGVItemDetails.Rows(i).Cells(26).Value <> "" Then
                        vSetupWord = Me.DGVItemDetails.Rows(i).Cells(26).Value
                        vSetupAmount = Me.DGVItemDetails.Rows(i).Cells(26).Value
                        vSetupAfter = vVatCost + vSetupAmount
                    Else
                        vSetupWord = ""
                        vSetupAmount = 0
                        vSetupAfter = vVatCost + vSetupAmount
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(27).Value) And Me.DGVItemDetails.Rows(i).Cells(27).Value <> "" Then
                        vServiceWord = Me.DGVItemDetails.Rows(i).Cells(27).Value
                        vServiceAmount = Me.DGVItemDetails.Rows(i).Cells(27).Value
                    Else
                        vServiceWord = ""
                        vServiceAmount = 0
                    End If
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(28).Value) And Me.DGVItemDetails.Rows(i).Cells(28).Value <> "" Then
                        vMarketCost = Me.DGVItemDetails.Rows(i).Cells(28).Value
                    Else
                        vMarketCost = 0
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(35).Value) And Me.DGVItemDetails.Rows(i).Cells(35).Value <> "" Then
                        vPointWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(35).Value, Double)), String) & "%"
                    Else
                        vPointWord = ""
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(36).Value) And Me.DGVItemDetails.Rows(i).Cells(36).Value <> "" Then
                        vPoint1 = Me.DGVItemDetails.Rows(i).Cells(36).Value
                    Else
                        vPoint1 = 0
                    End If
                    vPointAmount = vPoint1
                    vPointAfter = vMarketCost + vPointAmount

                    If vPointWord = "" And vPointAmount = 0 Then
                        MsgBox("รหัสสินค้า " & vItemCode & "   " & vItemName & " ไม่มีค่าของ Smart Point ไม่สามารถบันทึกข้อมูลได้  กรุณาแก้ไขข้อมูลก่อนบันทึกใหม่")
                        vQuery = "rollback tran"
                        cmd = New SqlCommand(vQuery, vConnection)
                        cmd.ExecuteNonQuery()
                        Me.DGVItemDetails.Item(35, i).Selected = True
                        Exit Sub
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(37).Value) And Me.DGVItemDetails.Rows(i).Cells(37).Value <> "" Then
                        vTargetWord = Me.DGVItemDetails.Rows(i).Cells(37).Value
                        vTargetAmount = Me.DGVItemDetails.Rows(i).Cells(37).Value
                        vTargetAfter = vPointAfter + vTargetAmount
                    Else
                        vTargetWord = ""
                        vTargetAmount = 0
                        vTargetAfter = vPointAfter + vTargetAmount
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(38).Value) And Me.DGVItemDetails.Rows(i).Cells(38).Value <> "" Then
                        vPremiumWord = Me.DGVItemDetails.Rows(i).Cells(38).Value
                        vPremiumAmount = Me.DGVItemDetails.Rows(i).Cells(38).Value
                        vPremiumAfter = vTargetAfter + vPremiumAmount
                    Else
                        vPremiumWord = ""
                        vPremiumAmount = 0
                        vPremiumAfter = vTargetAfter + vPremiumAmount
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(39).Value) And Me.DGVItemDetails.Rows(i).Cells(39).Value <> "" Then
                        vCommissionWord = CType((CType(Me.DGVItemDetails.Rows(i).Cells(39).Value, Double)), String) & "%"
                    Else
                        vCommissionWord = ""
                    End If
                    If vCommissionWord = "0%" Then
                        vCommissionWord = ""
                    End If

                    vCommissionAmount = vCommission1
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(40).Value) And Me.DGVItemDetails.Rows(i).Cells(40).Value <> "" Then
                        vCommission1 = Me.DGVItemDetails.Rows(i).Cells(40).Value
                    Else
                        vCommission1 = 0
                    End If
                    vCommissionAmount = vCommission1
                    vCommissionAfter = vPremiumAfter + vCommissionAmount

                    vGrossProfitPercent = (Me.DGVItemDetails.Rows(i).Cells(49).Value)
                    vGrossProfitAmount = ((Me.DGVItemDetails.Rows(i).Cells(49).Value) * vMarketCost) / 100

                    If vGrossProfitPercent <= 0 Then
                        MsgBox("รหัสสินค้า " & vItemCode & "   " & vItemName & " ไม่มีกำไรขั้นต้น ไม่สามารถบันทึกข้อมูลได้  กรุณาแก้ไขข้อมูลก่อนบันทึกใหม่")
                        vQuery = "rollback tran"
                        cmd = New SqlCommand(vQuery, vConnection)
                        cmd.ExecuteNonQuery()
                        Me.DGVItemDetails.Item(49, i).Selected = True
                        Exit Sub
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(29).Value) Then
                        vInterestStockPercent = (Me.DGVItemDetails.Rows(i).Cells(29).Value)
                    Else
                        vInterestStockPercent = ""
                    End If

                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(30).Value) And Me.DGVItemDetails.Rows(i).Cells(30).Value <> "" Then
                        vInterestStockAmount = Me.DGVItemDetails.Rows(i).Cells(30).Value
                    End If

                    vProfitPercent = (Me.DGVItemDetails.Rows(i).Cells(31).Value)
                    vProfitAmount = Me.DGVItemDetails.Rows(i).Cells(32).Value
                    vProfitPercent_W = (Me.DGVItemDetails.Rows(i).Cells(33).Value)
                    vProfitAmount_W = Me.DGVItemDetails.Rows(i).Cells(34).Value
                    If Not IsDBNull(Me.DGVItemDetails.Rows(i).Cells(48).Value) And Me.DGVItemDetails.Rows(i).Cells(48).Value <> "" Then
                        vMyDescriptionSub = Me.DGVItemDetails.Rows(i).Cells(48).Value
                    Else
                        vMyDescriptionSub = ""
                    End If
                    '---------------------------------------------------------------------------------------------------
                    Dim vFromQTY As Double
                    Dim vToQTY As Double
                    Dim vPriceSet2 As Double
                    Dim vIsPriceUpdate As Integer = 1
                    Dim vToUpdateDate As String = Me.DGVItemDetails.Rows(i).Cells(47).Value
                    Dim vIsUpdate As Integer = 0
                    Dim vIsPrintLabel As Integer = 0
                    Dim vPrice1CashRec As Double
                    Dim vPrice1CashDel As Double
                    Dim vPrice1CreditRec As Double
                    Dim vPrice1CreditDel As Double

                    Dim vDateNow As Date
                    Dim vUpdateDate As Date
                    Dim vItemDocDate As String
                    Dim vCheckDateDiff As Integer
                    '---------------------------------------------------------------------------------------------------

                    vDateNow = Now.Day & "/" & Now.Month & "/" & Now.Year
                    'vUpdateDate = vb6.Day(Me.DTPUpdate.Value) & "/" & vb6.Month(Me.DTPUpdate.Value) & "/" & vb6.Year(Me.DTPUpdate.Value)

                    vItemDocDate = Me.DGVItemDetails.Rows(i).Cells(47).Value
                    vUpdateDate = vItemDocDate
                    vCheckDateDiff = vb6.DateDiff(DateInterval.Day, vDateNow, vUpdateDate)


                    If vCheckDateDiff < 1 Then
                        MsgBox("รหัสสินค้า " & vItemCode & "   " & vItemName & " กำหนดวันที่ปรับต้องมากวันที่ปัจจุบัน ไม่สามารถบันทึกข้อมูลได้  กรุณาแก้ไขข้อมูลก่อนบันทึกใหม่")
                        vQuery = "rollback tran"
                        cmd = New SqlCommand(vQuery, vConnection)
                        cmd.ExecuteNonQuery()
                        Me.DGVItemDetails.Item(47, i).Selected = True
                        Exit Sub
                    End If
                    vSaleUnitCode = Me.DGVItemDetails.Rows(i).Cells(3).Value
                    vFromQTY = 1
                    vToQTY = 99999
                    vPrice1CashRec = Me.DGVItemDetails.Rows(i).Cells(42).Value
                    vPrice1CashDel = Me.DGVItemDetails.Rows(i).Cells(43).Value
                    vPrice1CreditRec = Me.DGVItemDetails.Rows(i).Cells(44).Value
                    vPrice1CreditDel = Me.DGVItemDetails.Rows(i).Cells(45).Value
                    vPriceSet2 = Me.DGVItemDetails.Rows(i).Cells(46).Value


                    vQuery = "exec dbo.USP_PS_InsertPriceStructureSubSet '" & vDocno & "','" & vItemCode & "','" & vItemName & "','" & vSaleUnitCode & "'," & vDO & ", " _
                    & "" & vPriceSet & ",'" & vDiscountBillWord & "'," & vDiscountBillAmount & "," & vAccCost & "," _
                    & "'" & vDiscountFollow1Word & "'," & vDiscountFollow1Amount & "," & vDiscountFollow1After & "," _
                    & "'" & vDiscountFollow2Word & "'," & vDiscountFollow2Amount & "," & vDiscountFollow2After & "," _
                    & "'" & vDiscountFollow3Word & "'," & vDiscountFollow3Amount & "," & vDiscountFollow3After & "," _
                    & "'" & vDiscountFollow4Word & "'," & vDiscountFollow4Amount & "," & vDiscountFollow4After & "," _
                    & "'" & vDiscountRebateWord & "'," & vDiscountRebateAmount & "," & vDiscountRebateAfter & "," _
                    & "'" & vDiscountSpecialWord & "'," & vDiscountSpecialAmount & "," & vNetCost & "," _
                    & "'" & vLossBudgetWord & "'," & vLossBudgetAmount & "," & vLossBudgetAfter & "," _
                    & "'" & vTransferInWord & "'," & vTransferIn1 & "," & vTransferInAfter & "," _
                    & "'" & vTransferOutWord & "'," & vTransferOut1 & "," & vTransferOutAfter & "," _
                    & "'" & vAdvertiseWord & "'," & vAdvertiseAmount & "," & vAdvertiseAfter & "," _
                    & "'" & vVatWord & "'," & vVatCost & "," & vVatAmount & "," _
                    & "'" & vSetupWord & "'," & vSetupAmount & "," & vSetupAfter & "," _
                    & "'" & vServiceWord & "'," & vServiceAmount & "," & vMarketCost & "," _
                    & "'" & vPointWord & "'," & vPointAmount & "," & vPointAfter & "," _
                    & "'" & vTargetWord & "'," & vTargetAmount & "," & vTargetAfter & "," _
                    & "'" & vPremiumWord & "'," & vPremiumAmount & "," & vPremiumAfter & "," _
                    & "'" & vCommissionWord & "'," & vCommissionAmount & "," & vCommissionAfter & "," _
                    & "" & vGrossProfitPercent & "," & vGrossProfitAmount & ",'" & vInterestStockPercent & "'," & vInterestStockAmount & "," _
                    & "" & vProfitPercent & "," & vProfitAmount & "," & vProfitPercent_W & "," & vProfitAmount_W & ",'" & vMyDescriptionSub & "'," & vFromQTY & "," & vToQTY & "," & vPrice1CashRec & "," & vPrice1CashDel & "," & vPrice1CreditRec & "," & vPrice1CreditDel & "," & vPriceSet2 & ", " _
                    & "" & vIsPriceUpdate & ",'" & vToUpdateDate & "'," & vIsUpdate & " "
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()
                    Me.PB101.Value = i
                End If
            Next

            vQuery = "commit tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            'vQuery = "exec dbo.USP_PS_DeliverySendMail '" & vDocno & "'"
            'cmd = New SqlCommand(vQuery, vConnection)
            'cmd.ExecuteNonQuery()

            MsgBox("บันทึกข้อมูลโครงสร้างราคาเลขที่ " & vDocno & " เรียบร้อยแล้วครับ")
            Me.PB101.Value = 0
            Me.DGVItemDetails.DataSource = Nothing
            Me.PB101.Value = 0
            Me.TBDocNo.Text = ""
            Me.TBMyDescription.Text = ""
            Me.DTPDocDate.Text = Now.Date
            Me.TBDocNo.Text = ""
            Me.PBNew.Visible = True
            Me.PBConfirm.Visible = False
            Call ClearScreen()
            Call ClearDataDGV()
            Call NewDoc()
            Call vGenDocNoAuto()

            vPriceStructureDocNo = Trim(vDocno)

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            vQuery = "rollback tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()
        End Try


    End Sub

    Private Sub DTPUpdate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTPUpdate.ValueChanged
        Dim i As Integer
        Dim vItemCode As String

        On Error Resume Next

        For i = 0 To Me.DGVItemDetails.RowCount - 1
            vItemCode = Me.DGVItemDetails.Item(1, i).Value
            If vItemCode <> "Nothing" And vItemCode <> "" Then
                If vb6.Year(Me.DTPUpdate.Value) > 2500 Then
                    Me.DGVItemDetails.Item(47, i).Value = vb6.Day(Me.DTPUpdate.Value) & "/" & vb6.Month(Me.DTPUpdate.Value) & "/" & vb6.Year(Me.DTPUpdate.Value) - 543
                Else
                    Me.DGVItemDetails.Item(47, i).Value = vb6.Day(Me.DTPUpdate.Value) & "/" & vb6.Month(Me.DTPUpdate.Value) & "/" & vb6.Year(Me.DTPUpdate.Value)
                End If
            End If
        Next
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click

    End Sub

    Private Sub BTNDocNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNDocNo.Click
        Call vGenDocNoAuto()
        Call NewDoc()
        vIsOpen = 0
        vMemIsCancel = 0
        vMemIsConfirm = 0
        Me.TBDocNo.Focus()
    End Sub

    Private Sub CBSelectItemAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBSelectItemAll.CheckedChanged
        Dim i As Integer

        On Error Resume Next

        If Me.ListViewSearch.Items.Count > 0 Then
            If Me.CBSelectItemAll.Checked = True Then
                For i = 0 To Me.ListViewSearch.Items.Count - 1
                    Me.ListViewSearch.Items(i).Checked = True
                Next
            End If

            If Me.CBSelectItemAll.Checked = False Then
                For i = 0 To Me.ListViewSearch.Items.Count - 1
                    Me.ListViewSearch.Items(i).Checked = False
                Next
            End If
        End If
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
        vQuery = "exec dbo.USP_PS_SearchListPriceStructure '" & vSearch & "'"
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
                vListItem.SubItems.Add(2).Text = dt.Rows(i).Item("creatorcode")
                vListItem.SubItems.Add(3).Text = dt.Rows(i).Item("mydescription")
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

    Private Sub BTNSearchDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchDocNo.Click
        Call SearchDocNo()
    End Sub

    Private Sub BTNCloseSearchDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSearchDocNo.Click
        Me.PNSearchDocNo.Visible = False
        Me.TBDocNo.Focus()
    End Sub

    Private Sub ListViewSearchDocNo_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewSearchDocNo.DoubleClick
        Call SearchPriceStructureDetails()
    End Sub

    Public Sub SearchPriceStructureDetails()
        Dim i As Integer
        Dim n As Integer
        Dim vDocNo As String
        Dim vDocdate As String
        Dim vTargetProfit As String
        Dim vProfit As String
        Dim vSmartpoint As String
        Dim vMyDescription As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vSaleUnitCode As String
        Dim vDOAmount As Double
        Dim vPriceSet As Double
        Dim vDiscountBillWord As String
        Dim vDiscountBillAmount As Double
        Dim vAccCost As Double
        Dim vDiscountFollow1Word As String
        Dim vDiscountFollow1Amount As Double
        Dim vDiscountFollow1After As Double
        Dim vDiscountFollow2Word As String
        Dim vDiscountFollow2Amount As Double
        Dim vDiscountFollow2After As Double
        Dim vDiscountFollow3Word As String
        Dim vDiscountFollow3Amount As Double
        Dim vDiscountFollow3After As Double
        Dim vDiscountFollow4Word As String
        Dim vDiscountFollow4Amount As Double
        Dim vDiscountFollow4After As Double
        Dim vDiscountRebateWord As String
        Dim vDiscountRebateAmount As Double
        Dim vDiscountRebateAfter As Double
        Dim vDiscountSpecialWord As String
        Dim vDiscountSpecialAmount As Double
        Dim vNetCost As Double
        Dim vLossBudgetWord As String
        Dim vLossBudgetAmount As Double
        Dim vLossBudgetAfter As Double
        Dim vTransferInWord As String
        Dim vTransferInAmount As Double
        Dim vTransferInAfter As Double
        Dim vTransferOutWord As String
        Dim vTransferOutAmount As Double
        Dim vTransferOutAfter As Double
        Dim vAdvertiseWord As String
        Dim vAdvertiseAmount As Double
        Dim vAdvertiseAfter As Double
        Dim vMarketingBudgetWord As String
        Dim vMarketingBudgetAmount As Double
        Dim vMarketingBudgetAfter As Double
        Dim vVatWord As String
        Dim vVatAmount As Double
        Dim vVatCost As Double
        Dim vSetupWord As String
        Dim vSetupAmount As Double
        Dim vSetupAfter As Double
        Dim vServiceWord As String
        Dim vServiceAmount As Double
        Dim vMarketCost As Double
        Dim vPointWord As String
        Dim vPointAmount As Double
        Dim vPointAfter As Double
        Dim vMemberDiscountWord As String
        Dim vMemberDiscountAmount As Double
        Dim vMemberDiscountAfter As Double
        Dim vTargetWord As String
        Dim vTargetAmount As Double
        Dim vTargetAfter As Double
        Dim vPremiumWord As String
        Dim vPremiumAmount As Double
        Dim vPremiumAfter As Double
        Dim vCommissionWord As String
        Dim vCommissionAmount As Double
        Dim vCommissionAfter As Double
        Dim vGrossProfitPercent As Double
        Dim vGrossProfitAmount As Double
        Dim vInterestStockPercent As Double
        Dim vInterestStockAmount As Double
        Dim vProfitPercent As Double
        Dim vProfitAmount As Double
        Dim vProfitPercent_W As Double
        Dim vProfitAmount_W As Double

        Dim vPrice1CashRec As Double
        Dim vPrice1CashDel As Double
        Dim vPrice1CreditRec As Double
        Dim vPrice1CreditDel As Double
        Dim vPriceSet2 As Double
        Dim vUpdateDate As String
        Dim vMyDescriptionSub As String


        On Error Resume Next

        If Me.ListViewSearchDocNo.Items.Count > 0 Then
            vDocNo = Me.ListViewSearchDocNo.SelectedItems(0).SubItems(1).Text

            Call ClearDataDGV()
            vQuery = "exec dbo.USP_PS_PriceStructureDetails '" & vDocNo & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Search")
            dt = ds.Tables("Search")
            If dt.Rows.Count > 0 Then
                vIsOpen = 1
                vMemIsConfirm = dt.Rows(0).Item("isconfirm")
                vMemIsCancel = dt.Rows(0).Item("iscancel")
                vDocNo = dt.Rows(0).Item("docno")
                vDocdate = dt.Rows(0).Item("docdate")
                vTargetProfit = dt.Rows(0).Item("TargetProfit")
                vProfit = dt.Rows(0).Item("Profit")
                vSmartpoint = dt.Rows(0).Item("Smartpoint")
                vMyDescription = dt.Rows(0).Item("MyDescription")

                'vMemberDiscount 
                Me.TBDocNo.Text = vDocNo
                Me.DTPDocDate.Text = vDocdate
                If InStr(vTargetProfit, "%") > 0 Then
                    Me.TBTarget.Text = vb6.Left(vTargetProfit, InStr(vTargetProfit, "%") - 1)
                Else
                    Me.TBTarget.Text = vTargetProfit
                End If
                If InStr(vProfit, "%") > 0 Then
                    Me.TBTargetAverage.Text = vb6.Left(vProfit, InStr(vProfit, "%") - 1)
                Else
                    Me.TBTargetAverage.Text = vProfit
                End If
                If InStr(vSmartpoint, "%") > 0 Then
                    Me.TBSmartPoint.Text = vb6.Left(vSmartpoint, InStr(vSmartpoint, "%") - 1)
                Else
                    Me.TBSmartPoint.Text = vSmartpoint
                End If

                Me.TBMyDescription.Text = vMyDescription

                For i = 0 To dt.Rows.Count - 1
                    n = n + 1
                    vItemCode = dt.Rows(i).Item("itemcode")
                    vItemName = dt.Rows(i).Item("itemname")
                    vSaleUnitCode = dt.Rows(i).Item("saleunitcode")


                    vDOAmount = dt.Rows(i).Item("D/O")
                    vPriceSet = dt.Rows(i).Item("PriceSet")
                    If dt.Rows(i).Item("DiscountBillWord") <> "" Then
                        vDiscountBillWord = vb6.Left(dt.Rows(i).Item("DiscountBillWord"), InStr(dt.Rows(i).Item("DiscountBillWord"), "%") - 1)
                    Else
                        vDiscountBillWord = ""
                    End If
                    vDiscountBillAmount = dt.Rows(i).Item("DiscountBillAmount")
                    vAccCost = dt.Rows(i).Item("AccCost")
                    If dt.Rows(i).Item("DiscountFollow1Word") <> "" Then
                        vDiscountFollow1Word = vb6.Left(dt.Rows(i).Item("DiscountFollow1Word"), InStr(dt.Rows(i).Item("DiscountFollow1Word"), "%") - 1)
                    Else
                        vDiscountFollow1Word = ""
                    End If

                    vDiscountFollow1Amount = dt.Rows(i).Item("DiscountFollow1Amount")
                    vDiscountFollow1After = dt.Rows(i).Item("DiscountFollow1After")
                    If dt.Rows(i).Item("DiscountFollow2Word") <> "" Then
                        vDiscountFollow2Word = vb6.Left(dt.Rows(i).Item("DiscountFollow2Word"), InStr(dt.Rows(i).Item("DiscountFollow2Word"), "%") - 1)
                    Else
                        vDiscountFollow2Word = ""
                    End If

                    vDiscountFollow2Amount = dt.Rows(i).Item("DiscountFollow2Amount")
                    vDiscountFollow2After = dt.Rows(i).Item("DiscountFollow2After")
                    If dt.Rows(i).Item("DiscountFollow3Word") <> "" Then
                        vDiscountFollow3Word = vb6.Left(dt.Rows(i).Item("DiscountFollow3Word"), InStr(dt.Rows(i).Item("DiscountFollow3Word"), "%") - 1)
                    Else
                        vDiscountFollow3Word = ""
                    End If

                    vDiscountFollow3Amount = dt.Rows(i).Item("DiscountFollow3Amount")
                    vDiscountFollow3After = dt.Rows(i).Item("DiscountFollow3After")
                    If dt.Rows(i).Item("DiscountFollow4Word") <> "" Then
                        vDiscountFollow4Word = vb6.Left(dt.Rows(i).Item("DiscountFollow4Word"), InStr(dt.Rows(i).Item("DiscountFollow4Word"), "%") - 1)
                    Else
                        vDiscountFollow4Word = ""
                    End If

                    vDiscountFollow4Amount = dt.Rows(i).Item("DiscountFollow4Amount")
                    vDiscountFollow4After = dt.Rows(i).Item("DiscountFollow4After")
                    If dt.Rows(i).Item("DiscountRebateWord") <> "" Then
                        vDiscountRebateWord = vb6.Left(dt.Rows(i).Item("DiscountRebateWord"), InStr(dt.Rows(i).Item("DiscountRebateWord"), "%") - 1)
                    Else
                        vDiscountRebateWord = ""
                    End If

                    vDiscountRebateAmount = dt.Rows(i).Item("DiscountRebateAmount")
                    vDiscountRebateAfter = dt.Rows(i).Item("DiscountRebateAfter")
                    If dt.Rows(i).Item("DiscountSpecialWord") <> "" And dt.Rows(i).Item("DiscountSpecialWord") <> "0" Then
                        vDiscountSpecialWord = dt.Rows(i).Item("DiscountSpecialWord")
                    Else
                        vDiscountSpecialWord = ""
                    End If

                    vDiscountSpecialAmount = dt.Rows(i).Item("DiscountSpecialAmount")
                    vNetCost = dt.Rows(i).Item("NetCost")
                    If dt.Rows(i).Item("LossBudgetWord") <> "" Then
                        vLossBudgetWord = vb6.Left(dt.Rows(i).Item("LossBudgetWord"), InStr(dt.Rows(i).Item("LossBudgetWord"), "%") - 1)
                    Else
                        vLossBudgetWord = ""
                    End If

                    vLossBudgetAmount = dt.Rows(i).Item("LossBudgetAmount")
                    vLossBudgetAfter = dt.Rows(i).Item("LossBudgetAfter")
                    'MsgBox(dt.Rows(i).Item("TransferInWord"))
                    If dt.Rows(i).Item("TransferInWord") <> "" Then
                        vTransferInWord = dt.Rows(i).Item("TransferInWord")
                    Else
                        vTransferInWord = ""
                    End If

                    vTransferInAmount = dt.Rows(i).Item("TransferInAmount")
                    vTransferInAfter = dt.Rows(i).Item("TransferInAfter")
                    If dt.Rows(i).Item("TransferOutWord") <> "" Then
                        vTransferOutWord = dt.Rows(i).Item("TransferOutWord")
                    Else
                        vTransferOutWord = ""
                    End If

                    vTransferOutAmount = dt.Rows(i).Item("TransferOutAmount")
                    vTransferOutAfter = dt.Rows(i).Item("TransferOutAfter")
                    If dt.Rows(i).Item("AdvertiseWord") <> "" Then
                        vAdvertiseWord = vb6.Left(dt.Rows(i).Item("AdvertiseWord"), InStr(dt.Rows(i).Item("AdvertiseWord"), "%") - 1)
                    Else
                        vAdvertiseWord = ""
                    End If

                    vAdvertiseAmount = dt.Rows(i).Item("AdvertiseAmount")
                    vAdvertiseAfter = dt.Rows(i).Item("AdvertiseAfter")
                    If dt.Rows(i).Item("MarketingBudgetWord") <> "" Then
                        vMarketingBudgetWord = dt.Rows(i).Item("MarketingBudgetWord")
                    Else
                        vMarketingBudgetWord = ""
                    End If

                    vMarketingBudgetAmount = dt.Rows(i).Item("MarketingBudgetAmount")
                    vMarketingBudgetAfter = dt.Rows(i).Item("MarketingBudgetAfter")
                    If dt.Rows(i).Item("VatWord") <> "" Then
                        vVatWord = vb6.Left(dt.Rows(i).Item("VatWord"), InStr(dt.Rows(i).Item("VatWord"), "%") - 1)
                    Else
                        vVatWord = ""
                    End If

                    vVatAmount = dt.Rows(i).Item("VatAmount")
                    vVatCost = dt.Rows(i).Item("VatCost")
                    If dt.Rows(i).Item("SetupWord") <> "" Then
                        vSetupWord = dt.Rows(i).Item("SetupWord")
                    Else
                        vSetupWord = ""
                    End If

                    vSetupAmount = dt.Rows(i).Item("SetupAmount")
                    vSetupAfter = dt.Rows(i).Item("SetupAfter")
                    If dt.Rows(i).Item("ServiceWord") <> "" Then
                        vServiceWord = dt.Rows(i).Item("ServiceWord")
                    Else
                        vServiceWord = ""
                    End If

                    vServiceAmount = dt.Rows(i).Item("ServiceAmount")
                    vMarketCost = dt.Rows(i).Item("MarketCost")
                    If dt.Rows(i).Item("PointWord") <> "" Then
                        vPointWord = vb6.Left(dt.Rows(i).Item("PointWord"), InStr(dt.Rows(i).Item("PointWord"), "%") - 1)
                    Else
                        vPointWord = ""
                    End If

                    vPointAmount = dt.Rows(i).Item("PointAmount")
                    vPointAfter = dt.Rows(i).Item("PointAfter")
                    If dt.Rows(i).Item("MemberDiscountWord") <> "" Then
                        vMemberDiscountWord = vb6.Left(dt.Rows(i).Item("MemberDiscountWord"), InStr(dt.Rows(i).Item("MemberDiscountWord"), "%") - 1)
                    Else
                        vMemberDiscountWord = ""
                    End If

                    vMemberDiscountAmount = dt.Rows(i).Item("MemberDiscountAmount")
                    vMemberDiscountAfter = dt.Rows(i).Item("MemberDiscountAfter")
                    If dt.Rows(i).Item("TargetWord") <> "" Then
                        vTargetWord = dt.Rows(i).Item("TargetWord")
                    Else
                        vTargetWord = ""
                    End If

                    vTargetAmount = dt.Rows(i).Item("TargetAmount")
                    vTargetAfter = dt.Rows(i).Item("TargetAfter")
                    If dt.Rows(i).Item("PremiumWord") <> "" Then
                        vPremiumWord = dt.Rows(i).Item("PremiumWord")
                    Else
                        vPremiumWord = ""
                    End If

                    vPremiumAmount = dt.Rows(i).Item("PremiumAmount")
                    vPremiumAfter = dt.Rows(i).Item("PremiumAfter")
                    If dt.Rows(i).Item("CommissionWord") <> "" Then
                        vCommissionWord = vb6.Left(dt.Rows(i).Item("CommissionWord"), InStr(dt.Rows(i).Item("CommissionWord"), "%") - 1)
                    Else
                        vCommissionWord = ""
                    End If

                    vCommissionAmount = dt.Rows(i).Item("CommissionAmount")
                    vCommissionAfter = dt.Rows(i).Item("CommissionAfter")

                    vGrossProfitPercent = dt.Rows(i).Item("GrossProfitPercent")
                    vGrossProfitAmount = dt.Rows(i).Item("GrossProfitAmount")
                    vInterestStockPercent = dt.Rows(i).Item("InterestStockPercent")
                    vInterestStockAmount = dt.Rows(i).Item("InterestStockAmount")
                    vProfitPercent = dt.Rows(i).Item("ProfitPercent")
                    vProfitAmount = dt.Rows(i).Item("ProfitAmount")
                    vProfitPercent_W = dt.Rows(i).Item("ProfitPercent_W")
                    vProfitAmount_W = dt.Rows(i).Item("ProfitAmount_W")

                    vPrice1CashRec = dt.Rows(i).Item("Price1CashRec")
                    vPrice1CashDel = dt.Rows(i).Item("Price1CashRec")
                    vPrice1CreditRec = dt.Rows(i).Item("Price1CreditRec")
                    vPrice1CreditDel = dt.Rows(i).Item("Price1CreditRec")
                    vPriceSet2 = dt.Rows(i).Item("PriceSet2")

                    vMyDescriptionSub = dt.Rows(i).Item("MyDescriptionSub")
                    vUpdateDate = dt.Rows(i).Item("toupdatedate")

                    Me.DGVItemDetails.Item(1, i).Value = vItemCode
                    Me.DGVItemDetails.Item(2, i).Value = vItemName
                    Me.DGVItemDetails.Item(3, i).Value = vSaleUnitCode
                    Me.DGVItemDetails.Item(4, i).Value = Format(vDOAmount, "##,##0.000")
                    Me.DGVItemDetails.Item(5, i).Value = vDiscountBillWord
                    Me.DGVItemDetails.Item(6, i).Value = Format(vAccCost, "##,##0.000")
                    Me.DGVItemDetails.Item(7, i).Value = vDiscountFollow1Word
                    Me.DGVItemDetails.Item(8, i).Value = Format(vDiscountFollow1After, "##,##0.000")
                    Me.DGVItemDetails.Item(9, i).Value = vDiscountFollow2Word
                    Me.DGVItemDetails.Item(10, i).Value = Format(vDiscountFollow2After, "##,##0.000")
                    Me.DGVItemDetails.Item(11, i).Value = vDiscountFollow3Word
                    Me.DGVItemDetails.Item(12, i).Value = Format(vDiscountFollow3After, "##,##0.000")
                    Me.DGVItemDetails.Item(13, i).Value = vDiscountFollow4Word
                    Me.DGVItemDetails.Item(14, i).Value = Format(vDiscountFollow4After, "##,##0.000")
                    Me.DGVItemDetails.Item(15, i).Value = vDiscountRebateWord
                    Me.DGVItemDetails.Item(16, i).Value = Format(vDiscountRebateAfter, "##,##0.000")
                    Me.DGVItemDetails.Item(17, i).Value = vDiscountSpecialWord
                    Me.DGVItemDetails.Item(18, i).Value = Format(vNetCost, "##,##0.000")
                    Me.DGVItemDetails.Item(19, i).Value = vLossBudgetWord
                    Me.DGVItemDetails.Item(20, i).Value = Format(vLossBudgetAfter, "##,##0.000")
                    Me.DGVItemDetails.Item(21, i).Value = vTransferInWord
                    Me.DGVItemDetails.Item(22, i).Value = vTransferOutWord
                    Me.DGVItemDetails.Item(23, i).Value = vAdvertiseWord
                    Me.DGVItemDetails.Item(24, i).Value = Format(vAdvertiseAfter, "##,##0.000")
                    Me.DGVItemDetails.Item(25, i).Value = Format(vVatCost, "##,##0.000")
                    Me.DGVItemDetails.Item(26, i).Value = vSetupWord
                    Me.DGVItemDetails.Item(27, i).Value = vServiceWord
                    Me.DGVItemDetails.Item(28, i).Value = Format(vMarketCost, "##,##0.000")
                    Me.DGVItemDetails.Item(29, i).Value = Format(vInterestStockPercent, "##,##0.000")
                    Me.DGVItemDetails.Item(30, i).Value = Format(vInterestStockAmount, "##,##0.000")
                    Me.DGVItemDetails.Item(31, i).Value = Format(vProfitPercent, "##,##0.000")
                    Me.DGVItemDetails.Item(32, i).Value = Format(vProfitAmount, "##,##0.000")
                    Me.DGVItemDetails.Item(33, i).Value = Format(vProfitPercent_W, "##,##0.000")
                    Me.DGVItemDetails.Item(34, i).Value = Format(vProfitAmount_W, "##,##0.000")
                    Me.DGVItemDetails.Item(35, i).Value = vPointWord
                    Me.DGVItemDetails.Item(36, i).Value = Format(vPointAmount, "##,##0.000")
                    Me.DGVItemDetails.Item(37, i).Value = vTargetWord
                    Me.DGVItemDetails.Item(38, i).Value = vPremiumWord
                    Me.DGVItemDetails.Item(39, i).Value = vCommissionWord
                    Me.DGVItemDetails.Item(40, i).Value = Format(vCommissionAmount, "##,##0.000")
                    Me.DGVItemDetails.Item(41, i).Value = Format(vPrice1CashRec, "##,##0.000")
                    Me.DGVItemDetails.Item(42, i).Value = Format(vPrice1CashRec, "##,##0.000")
                    Me.DGVItemDetails.Item(43, i).Value = Format(vPrice1CashRec, "##,##0.000")
                    Me.DGVItemDetails.Item(44, i).Value = Format(vPrice1CreditRec, "##,##0.000")
                    Me.DGVItemDetails.Item(45, i).Value = Format(vPrice1CreditRec, "##,##0.000")
                    Me.DGVItemDetails.Item(46, i).Value = Format(vPriceSet2, "##,##0.000")
                    Me.DGVItemDetails.Item(47, i).Value = vUpdateDate
                    Me.DGVItemDetails.Item(48, i).Value = vMyDescriptionSub
                    Me.DGVItemDetails.Item(49, i).Value = Format(vProfitPercent_W, "##,##0.000")

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
                Me.DGVItemDetails.Focus()
                Me.DGVItemDetails.Item(1, 0).Selected = True

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

    Private Sub ListViewSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearchDocNo.KeyDown
        If ListViewSearchDocNo.Items.Count > 0 Then
            If e.KeyCode = Keys.Enter Then
                Call SearchPriceStructureDetails()
            End If
        End If
    End Sub

    Private Sub ListViewSearchDocNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewSearchDocNo.SelectedIndexChanged

    End Sub

    Private Sub BTNPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPrint.Click
        On Error Resume Next

        If Me.TBDocNo.Text <> "" And vIsOpen = 1 Then
            If Me.DGVItemDetails.Rows.Count > 0 Then

                If frmPriceStructureRequest Is Nothing Then
                    frmPriceStructureRequest = New FormPriceStructureRequest
                Else
                    If frmPriceStructureRequest.IsDisposed Then
                        frmPriceStructureRequest = New FormPriceStructureRequest
                    End If
                End If

                vPriceStructureDocNo = Trim(Me.TBDocNo.Text)
                frmPriceStructureRequest.Show()
                frmPriceStructureRequest.BringToFront()
            Else
                MsgBox("ไม่สามารถพิมพ์เอกสารได้ เนื่องจากยังไม่ได้ Generate ข้อมูล", MsgBoxStyle.Critical, "Send Error")
            End If
        Else
            MsgBox("ไม่สามารถพิมพ์เอกสารได้ เนื่องจากยังไม่มีการบันทึกข้อมูล", MsgBoxStyle.Critical, "Send Error")
        End If
    End Sub

    Private Sub BTNClearScreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearScreen.Click
        Call ClearScreen()
        Call vGenDocNoAuto()
    End Sub

    Public Sub ClearScreen()
        On Error Resume Next

        vIsOpen = 0
        vMemIsCancel = 0
        vMemIsConfirm = 0
        Me.TBSmartPoint.Text = ""
        Me.TBTarget.Text = ""
        Me.TBTargetAverage.Text = ""
        Me.TBDocNo.Text = ""
        Me.DTPDocDate.Value = Now
        Me.TBMyDescription.Text = ""
        Call NewDoc()
        Call ClearDataDGV()
        Me.BTNDocNo.Focus()
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
            Me.DGVItemDetails.Item(11, i).Value = ""
            Me.DGVItemDetails.Item(12, i).Value = ""
            Me.DGVItemDetails.Item(13, i).Value = ""
            Me.DGVItemDetails.Item(14, i).Value = ""
            Me.DGVItemDetails.Item(15, i).Value = ""
            Me.DGVItemDetails.Item(16, i).Value = ""
            Me.DGVItemDetails.Item(17, i).Value = ""
            Me.DGVItemDetails.Item(18, i).Value = ""
            Me.DGVItemDetails.Item(19, i).Value = ""
            Me.DGVItemDetails.Item(20, i).Value = ""
            Me.DGVItemDetails.Item(21, i).Value = ""
            Me.DGVItemDetails.Item(22, i).Value = ""
            Me.DGVItemDetails.Item(23, i).Value = ""
            Me.DGVItemDetails.Item(24, i).Value = ""
            Me.DGVItemDetails.Item(25, i).Value = ""
            Me.DGVItemDetails.Item(26, i).Value = ""
            Me.DGVItemDetails.Item(27, i).Value = ""
            Me.DGVItemDetails.Item(28, i).Value = ""
            Me.DGVItemDetails.Item(29, i).Value = ""
            Me.DGVItemDetails.Item(30, i).Value = ""
            Me.DGVItemDetails.Item(31, i).Value = ""
            Me.DGVItemDetails.Item(32, i).Value = ""
            Me.DGVItemDetails.Item(33, i).Value = ""
            Me.DGVItemDetails.Item(34, i).Value = ""
            Me.DGVItemDetails.Item(35, i).Value = ""
            Me.DGVItemDetails.Item(36, i).Value = ""
            Me.DGVItemDetails.Item(37, i).Value = ""
            Me.DGVItemDetails.Item(38, i).Value = ""
            Me.DGVItemDetails.Item(39, i).Value = ""
            Me.DGVItemDetails.Item(40, i).Value = ""
            Me.DGVItemDetails.Item(41, i).Value = ""
            Me.DGVItemDetails.Item(42, i).Value = ""
            Me.DGVItemDetails.Item(43, i).Value = ""
            Me.DGVItemDetails.Item(44, i).Value = ""
            Me.DGVItemDetails.Item(45, i).Value = ""
            Me.DGVItemDetails.Item(46, i).Value = ""
            Me.DGVItemDetails.Item(47, i).Value = ""
            Me.DGVItemDetails.Item(48, i).Value = ""
            Me.DGVItemDetails.Item(49, i).Value = ""
        Next
    End Sub

    Private Sub BTNCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCancel.Click
        Dim vDocNo As String
        Dim vAnswer As Integer

        On Error Resume Next

        If Me.TBDocNo.Text <> "" And vIsOpen = 1 Then

            vDocNo = Me.TBDocNo.Text

            If vMemIsCancel = 1 Then
                MsgBox("เอกสารเลขที่ " & vDocNo & " ถูกยกเลิกไปแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Information Message")
                Me.TBDocNo.Focus()
                Exit Sub
            End If

            If vMemIsConfirm = 1 Then
                MsgBox("เอกสารเลขที่ " & vDocNo & " ถูกอนุมัติไปแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Information Message")
                Me.TBDocNo.Focus()
                Exit Sub
            End If

            vAnswer = MsgBox("คุณต้องการยกเลิกเอกสารเลขที่ " & vDocNo & " นี้ใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 Then
                vQuery = "exec dbo.USP_PS_CancelPriceStructureSet '" & vDocNo & "','" & vUserID & "'"
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                MsgBox("ยกเลิกเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")
                Call ClearScreen()
                Call NewDoc()
            End If

            Me.TBDocNo.Focus()
        End If
    End Sub

    Private Sub BTNSelectSearchDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectSearchDocNo.Click
        If ListViewSearchDocNo.Items.Count > 0 Then
            Call SearchPriceStructureDetails()
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

                vAnswer = MsgBox("คุณต้องการลบรายการที่ " & vRowID + 1 & " ใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
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

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NDCash.ValueChanged
        Call CalcCashPrice()
    End Sub

    Private Sub CBCash_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBCash.CheckedChanged
        If Me.CBCash.Checked = True Then
            Call CalcCashPrice()
            Me.NDCash.Enabled = True
            Me.CBDOCash.Enabled = False
            Me.NDCash.Focus()
        Else
            Me.CBDOCash.Enabled = True
            Me.NDCash.Enabled = False
        End If
    End Sub

    Public Sub CalcCashPrice()
        Dim i As Integer
        Dim vItemCode As String
        Dim vMarketCost As Double
        Dim vProfit As Double
        Dim vCashAmount As Double

        On Error Resume Next


        If Me.CBCash.Checked = True Then
            vProfit = Me.NDCash.Value

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                If Me.DGVItemDetails.Item(28, i).Value <> "" Then
                    vMarketCost = Me.DGVItemDetails.Item(28, i).Value
                Else
                    vMarketCost = 0
                End If
                vItemCode = Me.DGVItemDetails.Item(1, i).Value

                If vItemCode <> "Nothing" And vItemCode <> "" Then

                    If vMarketCost <> 0 Then
                        vCashAmount = Math.Round((vMarketCost * 100) / (100 - vProfit))
                        Me.DGVItemDetails.Item(42, i).Value = Format(vCashAmount, "##,##0.000")
                        Me.DGVItemDetails.Item(43, i).Value = Format(vCashAmount, "##,##0.000")
                        Call vCalcItemLine(i)
                    End If
                End If
            Next
        End If
    End Sub

    Public Sub CalcDOCashPrice()
        Dim i As Integer
        Dim vItemCode As String
        Dim vDOPrice As Double
        Dim vProfit As Double
        Dim vCashAmount As Double

        On Error Resume Next


        If Me.CBDOCash.Checked = True Then
            vProfit = Me.NDDOCash.Value

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                If Me.DGVItemDetails.Item(4, i).Value <> "" Then
                    vDOPrice = Me.DGVItemDetails.Item(4, i).Value
                Else
                    vDOPrice = 0
                End If
                vItemCode = Me.DGVItemDetails.Item(1, i).Value

                If vItemCode <> "Nothing" And vItemCode <> "" Then

                    If vDOPrice <> 0 Then
                        'vCashAmount = vDOPrice - Math.Round((vDOPrice * vProfit) / 100)
                        vCashAmount = vDOPrice - ((vDOPrice * vProfit) / 100)

                        vCashAmount = Math.Ceiling(vCashAmount)
                        Me.DGVItemDetails.Item(42, i).Value = Format(vCashAmount, "##,##0.000")
                        Me.DGVItemDetails.Item(43, i).Value = Format(vCashAmount, "##,##0.000")
                        Call vCalcItemLine(i)
                    End If
                End If
            Next
        End If
    End Sub

    Public Sub CalcDOCashPriceLine(ByVal i As Integer)
        Dim vItemCode As String
        Dim vDOPrice As Double
        Dim vProfit As Double
        Dim vCashAmount As Double

        On Error Resume Next


        If Me.CBDOCash.Checked = True Then
            vProfit = Me.NDDOCash.Value

            If Me.DGVItemDetails.Item(4, i).Value <> "" Then
                vDOPrice = Me.DGVItemDetails.Item(4, i).Value
            Else
                vDOPrice = 0
            End If
            vItemCode = Me.DGVItemDetails.Item(1, i).Value

            If vItemCode <> "Nothing" And vItemCode <> "" Then

                If vDOPrice <> 0 Then
                    'vCashAmount = vDOPrice - Math.Round((vDOPrice * vProfit) / 100)
                    vCashAmount = vDOPrice - ((vDOPrice * vProfit) / 100)

                    vCashAmount = Math.Ceiling(vCashAmount)

                    Me.DGVItemDetails.Item(42, i).Value = Format(vCashAmount, "##,##0.000")
                    Me.DGVItemDetails.Item(43, i).Value = Format(vCashAmount, "##,##0.000")
                    Call vCalcItemLine(i)
                End If
            End If
        End If
    End Sub

    Public Sub CalcDOCreditPrice()
        Dim i As Integer
        Dim vItemCode As String
        Dim vDOPrice As Double
        Dim vProfit As Double
        Dim vCreditAmount As Double

        On Error Resume Next


        If Me.CBDOCredit.Checked = True Then
            vProfit = Me.NDDOCredit.Value

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                If Me.DGVItemDetails.Item(4, i).Value <> "" Then
                    vDOPrice = Me.DGVItemDetails.Item(4, i).Value
                Else
                    vDOPrice = 0
                End If
                vItemCode = Me.DGVItemDetails.Item(1, i).Value

                If vItemCode <> "Nothing" And vItemCode <> "" Then

                    If vDOPrice <> 0 Then
                        'vCreditAmount = vDOPrice - Math.Round((vDOPrice * vProfit) / 100)
                        vCreditAmount = vDOPrice - ((vDOPrice * vProfit) / 100)

                        vCreditAmount = Math.Ceiling(vCreditAmount)

                        Me.DGVItemDetails.Item(44, i).Value = Format(vCreditAmount, "##,##0.000")
                        Me.DGVItemDetails.Item(45, i).Value = Format(vCreditAmount, "##,##0.000")
                        Call vCalcItemLine(i)
                    End If
                End If
            Next
        End If
    End Sub

    Public Sub CalcDOCreditPriceLine(ByVal i As Integer)
        Dim vItemCode As String
        Dim vDOPrice As Double
        Dim vProfit As Double
        Dim vCreditAmount As Double

        On Error Resume Next


        If Me.CBDOCredit.Checked = True Then
            vProfit = Me.NDDOCredit.Value

            If Me.DGVItemDetails.Item(4, i).Value <> "" Then
                vDOPrice = Me.DGVItemDetails.Item(4, i).Value
            Else
                vDOPrice = 0
            End If
            vItemCode = Me.DGVItemDetails.Item(1, i).Value

            If vItemCode <> "Nothing" And vItemCode <> "" Then

                If vDOPrice <> 0 Then
                    'vCreditAmount = vDOPrice - Math.Round((vDOPrice * vProfit) / 100)
                    vCreditAmount = vDOPrice - ((vDOPrice * vProfit) / 100)

                    vCreditAmount = Math.Ceiling(vCreditAmount)

                    Me.DGVItemDetails.Item(44, i).Value = Format(vCreditAmount, "##,##0.000")
                    Me.DGVItemDetails.Item(45, i).Value = Format(vCreditAmount, "##,##0.000")
                    Call vCalcItemLine(i)
                End If
            End If
        End If
    End Sub

    Public Sub CalcDOPrice2()
        Dim i As Integer
        Dim vItemCode As String
        Dim vDOPrice As Double
        Dim vProfit As Double
        Dim vPrice2 As Double

        On Error Resume Next


        If Me.CBDOPrice2.Checked = True Then
            vProfit = Me.NDDOPrice2.Value

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                If Me.DGVItemDetails.Item(4, i).Value <> "" Then
                    vDOPrice = Me.DGVItemDetails.Item(4, i).Value
                Else
                    vDOPrice = 0
                End If
                vItemCode = Me.DGVItemDetails.Item(1, i).Value

                If vItemCode <> "Nothing" And vItemCode <> "" Then

                    If vDOPrice <> 0 Then
                        vPrice2 = vDOPrice - ((vDOPrice * vProfit) / 100)
                        Me.DGVItemDetails.Item(46, i).Value = Format(vPrice2, "##,##0.000")
                        Call vCalcItemLine(i)
                    End If
                End If
            Next
        End If
    End Sub

    Public Sub CalcDOPrice2Line(ByVal i As Integer)
        Dim vItemCode As String
        Dim vDOPrice As Double
        Dim vProfit As Double
        Dim vPrice2 As Double

        On Error Resume Next


        If Me.CBDOPrice2.Checked = True Then
            vProfit = Me.NDDOPrice2.Value

            If Me.DGVItemDetails.Item(4, i).Value <> "" Then
                vDOPrice = Me.DGVItemDetails.Item(4, i).Value
            Else
                vDOPrice = 0
            End If
            vItemCode = Me.DGVItemDetails.Item(1, i).Value

            If vItemCode <> "Nothing" And vItemCode <> "" Then

                If vDOPrice <> 0 Then
                    vPrice2 = vDOPrice - Math.Round((vDOPrice * vProfit) / 100)
                    Me.DGVItemDetails.Item(46, i).Value = Format(vPrice2, "##,##0.000")
                    Call vCalcItemLine(i)
                End If
            End If
        End If
    End Sub

    Public Sub CalcCashPriceLine(ByVal i As Integer)
        Dim vItemCode As String
        Dim vMarketCost As Double
        Dim vProfit As Double
        Dim vCashAmount As Double

        On Error Resume Next

        If Me.CBCash.Checked = True Then
            vProfit = Me.NDCash.Value

            If Me.DGVItemDetails.Item(28, i).Value <> "" Then
                vMarketCost = Me.DGVItemDetails.Item(28, i).Value
            Else
                vMarketCost = 0
            End If
            vItemCode = Me.DGVItemDetails.Item(1, i).Value

            If vItemCode <> "Nothing" And vItemCode <> "" Then

                If vMarketCost <> 0 Then
                    vCashAmount = Math.Round((vMarketCost * 100) / (100 - vProfit))
                    Me.DGVItemDetails.Item(42, i).Value = Format(vCashAmount, "##,##0.000")
                    Me.DGVItemDetails.Item(43, i).Value = Format(vCashAmount, "##,##0.000")
                    Call vCalcItemLine(i)
                End If

            End If

        End If
    End Sub


    Public Sub CalcCreditPrice()
        Dim i As Integer
        Dim vItemCode As String
        Dim vCashAmount As Double
        Dim vAddPrice As Double
        Dim vCreditAmount As Double

        On Error Resume Next

        If Me.CBCredit.Checked = True Then
            vAddPrice = Me.NDCredit.Value

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                If Me.DGVItemDetails.Item(42, i).Value <> "" Then
                    vCashAmount = Me.DGVItemDetails.Item(42, i).Value
                Else
                    vCashAmount = 0
                End If
                vItemCode = Me.DGVItemDetails.Item(1, i).Value

                If vItemCode <> "Nothing" And vItemCode <> "" Then

                    If vCashAmount <> 0 Then
                        vCreditAmount = vCashAmount + vAddPrice
                        Me.DGVItemDetails.Item(44, i).Value = Format(vCreditAmount, "##,##0.000")
                        Me.DGVItemDetails.Item(45, i).Value = Format(vCreditAmount, "##,##0.000")
                        Call vCalcItemLine(i)
                    End If

                End If
            Next
        End If
    End Sub

    Public Sub CalcCreditPriceLine(ByVal i As Integer)
        Dim vItemCode As String
        Dim vCashAmount As Double
        Dim vAddPrice As Double
        Dim vCreditAmount As Double

        On Error Resume Next

        If Me.CBCredit.Checked = True Then
            vAddPrice = Me.NDCredit.Value

            If Me.DGVItemDetails.Item(42, i).Value <> "" Then
                vCashAmount = Me.DGVItemDetails.Item(42, i).Value
            Else
                vCashAmount = 0
            End If
            vItemCode = Me.DGVItemDetails.Item(1, i).Value

            If vItemCode <> "Nothing" And vItemCode <> "" Then

                If vCashAmount <> 0 Then
                    vCreditAmount = vCashAmount + vAddPrice
                    Me.DGVItemDetails.Item(44, i).Value = Format(vCreditAmount, "##,##0.000")
                    Me.DGVItemDetails.Item(45, i).Value = Format(vCreditAmount, "##,##0.000")
                    Call vCalcItemLine(i)
                End If

            End If

        End If
    End Sub

    Private Sub CBCredit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBCredit.CheckedChanged
        If Me.CBCredit.Checked = True Then
            Call CalcCreditPrice()
            Me.NDCredit.Enabled = True
            Me.CBDOCredit.Enabled = False
            Me.NDCredit.Focus()
        Else
            Me.CBDOCredit.Enabled = True
            Me.NDCredit.Enabled = False
        End If
    End Sub

    Private Sub NDCredit_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NDCredit.ValueChanged
        Call CalcCreditPrice()
    End Sub

    Public Sub CalcPrice2()
        Dim i As Integer
        Dim vItemCode As String
        Dim vCashAmount As Double
        Dim vDownPrice As Double
        Dim vPrice2Amount As Double

        On Error Resume Next

        If Me.CBPrice2.Checked = True Then
            vDownPrice = Me.NDPrice2.Value

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                If Me.DGVItemDetails.Item(42, i).Value <> "" Then
                    vCashAmount = Me.DGVItemDetails.Item(42, i).Value
                Else
                    vCashAmount = 0
                End If
                vItemCode = Me.DGVItemDetails.Item(1, i).Value

                If vItemCode <> "Nothing" And vItemCode <> "" Then

                    If vCashAmount <> 0 Then
                        vPrice2Amount = vCashAmount - vDownPrice
                        Me.DGVItemDetails.Item(46, i).Value = Format(vPrice2Amount, "##,##0.000")
                        Call vCalcItemLine(i)
                    End If

                End If
            Next
        End If
    End Sub

    Public Sub CalcPrice2Line(ByVal i As Integer)
        Dim vItemCode As String
        Dim vCashAmount As Double
        Dim vDownPrice As Double
        Dim vPrice2Amount As Double

        On Error Resume Next

        If Me.CBPrice2.Checked = True Then
            vDownPrice = Me.NDPrice2.Value

            If Me.DGVItemDetails.Item(42, i).Value <> "" Then
                vCashAmount = Me.DGVItemDetails.Item(42, i).Value
            Else
                vCashAmount = 0
            End If
            vItemCode = Me.DGVItemDetails.Item(1, i).Value

            If vItemCode <> "Nothing" And vItemCode <> "" Then

                If vCashAmount <> 0 Then
                    vPrice2Amount = vCashAmount - vDownPrice
                    Me.DGVItemDetails.Item(46, i).Value = Format(vPrice2Amount, "##,##0.000")
                    Call vCalcItemLine(i)
                End If

            End If
        End If
    End Sub

    Private Sub CBPrice2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBPrice2.CheckedChanged
        If Me.CBPrice2.Checked = True Then
            Call CalcPrice2()
            Me.NDPrice2.Enabled = True
            Me.CBDOPrice2.Enabled = False
            Me.NDPrice2.Focus()
        Else
            Me.CBDOPrice2.Enabled = True
            Me.NDPrice2.Enabled = False
        End If
    End Sub

    Private Sub NDPrice2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NDPrice2.ValueChanged
        Call CalcPrice2()
    End Sub

    Private Sub BTNShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNShow.Click
        Me.DGVItemDetails.Columns(6).Visible = True
        Me.DGVItemDetails.Columns(7).Visible = True
        Me.DGVItemDetails.Columns(8).Visible = True
        Me.DGVItemDetails.Columns(9).Visible = True
        Me.DGVItemDetails.Columns(10).Visible = True
        Me.DGVItemDetails.Columns(11).Visible = True
        Me.DGVItemDetails.Columns(12).Visible = True
        Me.DGVItemDetails.Columns(13).Visible = True
        Me.DGVItemDetails.Columns(14).Visible = True
        Me.DGVItemDetails.Columns(15).Visible = True
        Me.DGVItemDetails.Columns(16).Visible = True
        Me.DGVItemDetails.Columns(17).Visible = True
        Me.DGVItemDetails.Columns(18).Visible = True
        Me.DGVItemDetails.Columns(19).Visible = True
        Me.DGVItemDetails.Columns(20).Visible = True
        Me.DGVItemDetails.Columns(23).Visible = True
        Me.DGVItemDetails.Columns(24).Visible = True
        Me.DGVItemDetails.Columns(25).Visible = True
        Me.DGVItemDetails.Columns(26).Visible = True
        Me.DGVItemDetails.Columns(27).Visible = True
        Me.DGVItemDetails.Columns(28).Visible = True
        Me.DGVItemDetails.Columns(29).Visible = True
        Me.DGVItemDetails.Columns(30).Visible = True
        Me.DGVItemDetails.Columns(31).Visible = True
        Me.DGVItemDetails.Columns(32).Visible = True
        Me.DGVItemDetails.Columns(33).Visible = True
        Me.DGVItemDetails.Columns(34).Visible = True
        Me.DGVItemDetails.Columns(35).Visible = True
        Me.DGVItemDetails.Columns(36).Visible = True
        Me.DGVItemDetails.Columns(37).Visible = True
        Me.DGVItemDetails.Columns(38).Visible = True
        Me.DGVItemDetails.Columns(39).Visible = True
        Me.DGVItemDetails.Columns(40).Visible = True
        Me.DGVItemDetails.Columns(41).Visible = True
        Me.DGVItemDetails.Columns(43).Visible = True
        Me.DGVItemDetails.Columns(45).Visible = True
    End Sub

    Private Sub BTNHideNormal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNHideNormal.Click
        Me.DGVItemDetails.Columns(6).Visible = True
        Me.DGVItemDetails.Columns(7).Visible = True
        Me.DGVItemDetails.Columns(8).Visible = False
        Me.DGVItemDetails.Columns(9).Visible = True
        Me.DGVItemDetails.Columns(10).Visible = False
        Me.DGVItemDetails.Columns(11).Visible = True
        Me.DGVItemDetails.Columns(12).Visible = False
        Me.DGVItemDetails.Columns(13).Visible = True
        Me.DGVItemDetails.Columns(14).Visible = False
        Me.DGVItemDetails.Columns(15).Visible = True
        Me.DGVItemDetails.Columns(16).Visible = False
        Me.DGVItemDetails.Columns(17).Visible = False
        Me.DGVItemDetails.Columns(18).Visible = False
        Me.DGVItemDetails.Columns(19).Visible = False
        Me.DGVItemDetails.Columns(20).Visible = False
        Me.DGVItemDetails.Columns(23).Visible = False
        Me.DGVItemDetails.Columns(24).Visible = False
        Me.DGVItemDetails.Columns(25).Visible = False
        Me.DGVItemDetails.Columns(26).Visible = False
        Me.DGVItemDetails.Columns(27).Visible = False
        Me.DGVItemDetails.Columns(28).Visible = True
        Me.DGVItemDetails.Columns(29).Visible = False
        Me.DGVItemDetails.Columns(30).Visible = False
        Me.DGVItemDetails.Columns(31).Visible = False
        Me.DGVItemDetails.Columns(32).Visible = False
        Me.DGVItemDetails.Columns(33).Visible = False
        Me.DGVItemDetails.Columns(34).Visible = False
        Me.DGVItemDetails.Columns(35).Visible = True
        Me.DGVItemDetails.Columns(37).Visible = False
        Me.DGVItemDetails.Columns(38).Visible = False
        Me.DGVItemDetails.Columns(39).Visible = False
        Me.DGVItemDetails.Columns(40).Visible = False
        Me.DGVItemDetails.Columns(41).Visible = False
        Me.DGVItemDetails.Columns(43).Visible = False
        Me.DGVItemDetails.Columns(45).Visible = False
    End Sub

    Private Sub BTNHideAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNHideAll.Click
        Me.DGVItemDetails.Columns(6).Visible = False
        Me.DGVItemDetails.Columns(7).Visible = False
        Me.DGVItemDetails.Columns(8).Visible = False
        Me.DGVItemDetails.Columns(9).Visible = False
        Me.DGVItemDetails.Columns(10).Visible = False
        Me.DGVItemDetails.Columns(11).Visible = False
        Me.DGVItemDetails.Columns(12).Visible = False
        Me.DGVItemDetails.Columns(13).Visible = False
        Me.DGVItemDetails.Columns(14).Visible = False
        Me.DGVItemDetails.Columns(15).Visible = False
        Me.DGVItemDetails.Columns(16).Visible = False
        Me.DGVItemDetails.Columns(17).Visible = False
        Me.DGVItemDetails.Columns(18).Visible = False
        Me.DGVItemDetails.Columns(19).Visible = False
        Me.DGVItemDetails.Columns(20).Visible = False
        Me.DGVItemDetails.Columns(22).Visible = False
        Me.DGVItemDetails.Columns(23).Visible = False
        Me.DGVItemDetails.Columns(24).Visible = False
        Me.DGVItemDetails.Columns(25).Visible = False
        Me.DGVItemDetails.Columns(26).Visible = False
        Me.DGVItemDetails.Columns(27).Visible = False
        Me.DGVItemDetails.Columns(28).Visible = True
        Me.DGVItemDetails.Columns(29).Visible = False
        Me.DGVItemDetails.Columns(30).Visible = False
        Me.DGVItemDetails.Columns(31).Visible = False
        Me.DGVItemDetails.Columns(32).Visible = False
        Me.DGVItemDetails.Columns(33).Visible = False
        Me.DGVItemDetails.Columns(34).Visible = False
        Me.DGVItemDetails.Columns(35).Visible = True
        Me.DGVItemDetails.Columns(36).Visible = False
        Me.DGVItemDetails.Columns(37).Visible = False
        Me.DGVItemDetails.Columns(38).Visible = False
        Me.DGVItemDetails.Columns(39).Visible = False
        Me.DGVItemDetails.Columns(40).Visible = False
        Me.DGVItemDetails.Columns(41).Visible = False
        Me.DGVItemDetails.Columns(43).Visible = False
        Me.DGVItemDetails.Columns(45).Visible = False
    End Sub

    Private Sub BTNHideDiscount1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNHideDiscount1.Click
        Me.DGVItemDetails.Columns(6).Visible = False
        Me.DGVItemDetails.Columns(7).Visible = True
        Me.DGVItemDetails.Columns(8).Visible = False
        Me.DGVItemDetails.Columns(9).Visible = False
        Me.DGVItemDetails.Columns(10).Visible = False
        Me.DGVItemDetails.Columns(11).Visible = False
        Me.DGVItemDetails.Columns(12).Visible = False
        Me.DGVItemDetails.Columns(13).Visible = False
        Me.DGVItemDetails.Columns(14).Visible = False
        Me.DGVItemDetails.Columns(15).Visible = False
        Me.DGVItemDetails.Columns(16).Visible = False
        Me.DGVItemDetails.Columns(17).Visible = False
        Me.DGVItemDetails.Columns(18).Visible = False
        Me.DGVItemDetails.Columns(19).Visible = False
        Me.DGVItemDetails.Columns(20).Visible = False
        Me.DGVItemDetails.Columns(22).Visible = False
        Me.DGVItemDetails.Columns(23).Visible = False
        Me.DGVItemDetails.Columns(24).Visible = False
        Me.DGVItemDetails.Columns(25).Visible = False
        Me.DGVItemDetails.Columns(26).Visible = False
        Me.DGVItemDetails.Columns(27).Visible = False
        Me.DGVItemDetails.Columns(28).Visible = True
        Me.DGVItemDetails.Columns(29).Visible = False
        Me.DGVItemDetails.Columns(30).Visible = False
        Me.DGVItemDetails.Columns(31).Visible = False
        Me.DGVItemDetails.Columns(32).Visible = False
        Me.DGVItemDetails.Columns(33).Visible = False
        Me.DGVItemDetails.Columns(34).Visible = False
        Me.DGVItemDetails.Columns(35).Visible = True
        Me.DGVItemDetails.Columns(36).Visible = False
        Me.DGVItemDetails.Columns(37).Visible = False
        Me.DGVItemDetails.Columns(38).Visible = False
        Me.DGVItemDetails.Columns(39).Visible = False
        Me.DGVItemDetails.Columns(40).Visible = False
        Me.DGVItemDetails.Columns(41).Visible = False
        Me.DGVItemDetails.Columns(43).Visible = False
        Me.DGVItemDetails.Columns(45).Visible = False
    End Sub

    Private Sub BTNHideDiscount2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNHideDiscount2.Click
        Me.DGVItemDetails.Columns(6).Visible = False
        Me.DGVItemDetails.Columns(7).Visible = True
        Me.DGVItemDetails.Columns(8).Visible = False
        Me.DGVItemDetails.Columns(9).Visible = True
        Me.DGVItemDetails.Columns(10).Visible = False
        Me.DGVItemDetails.Columns(11).Visible = False
        Me.DGVItemDetails.Columns(12).Visible = False
        Me.DGVItemDetails.Columns(13).Visible = False
        Me.DGVItemDetails.Columns(14).Visible = False
        Me.DGVItemDetails.Columns(15).Visible = False
        Me.DGVItemDetails.Columns(16).Visible = False
        Me.DGVItemDetails.Columns(17).Visible = False
        Me.DGVItemDetails.Columns(18).Visible = False
        Me.DGVItemDetails.Columns(19).Visible = False
        Me.DGVItemDetails.Columns(20).Visible = False
        Me.DGVItemDetails.Columns(22).Visible = False
        Me.DGVItemDetails.Columns(23).Visible = False
        Me.DGVItemDetails.Columns(24).Visible = False
        Me.DGVItemDetails.Columns(25).Visible = False
        Me.DGVItemDetails.Columns(26).Visible = False
        Me.DGVItemDetails.Columns(27).Visible = False
        Me.DGVItemDetails.Columns(28).Visible = True
        Me.DGVItemDetails.Columns(29).Visible = False
        Me.DGVItemDetails.Columns(30).Visible = False
        Me.DGVItemDetails.Columns(31).Visible = False
        Me.DGVItemDetails.Columns(32).Visible = False
        Me.DGVItemDetails.Columns(33).Visible = False
        Me.DGVItemDetails.Columns(34).Visible = False
        Me.DGVItemDetails.Columns(35).Visible = True
        Me.DGVItemDetails.Columns(36).Visible = False
        Me.DGVItemDetails.Columns(37).Visible = False
        Me.DGVItemDetails.Columns(38).Visible = False
        Me.DGVItemDetails.Columns(39).Visible = False
        Me.DGVItemDetails.Columns(40).Visible = False
        Me.DGVItemDetails.Columns(41).Visible = False
        Me.DGVItemDetails.Columns(43).Visible = False
        Me.DGVItemDetails.Columns(45).Visible = False
    End Sub

    Private Sub BTNHideDiscount3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNHideDiscount3.Click
        Me.DGVItemDetails.Columns(6).Visible = False
        Me.DGVItemDetails.Columns(7).Visible = True
        Me.DGVItemDetails.Columns(8).Visible = False
        Me.DGVItemDetails.Columns(9).Visible = True
        Me.DGVItemDetails.Columns(10).Visible = False
        Me.DGVItemDetails.Columns(11).Visible = True
        Me.DGVItemDetails.Columns(12).Visible = False
        Me.DGVItemDetails.Columns(13).Visible = False
        Me.DGVItemDetails.Columns(14).Visible = False
        Me.DGVItemDetails.Columns(15).Visible = False
        Me.DGVItemDetails.Columns(16).Visible = False
        Me.DGVItemDetails.Columns(17).Visible = False
        Me.DGVItemDetails.Columns(18).Visible = False
        Me.DGVItemDetails.Columns(19).Visible = False
        Me.DGVItemDetails.Columns(20).Visible = False
        Me.DGVItemDetails.Columns(22).Visible = False
        Me.DGVItemDetails.Columns(23).Visible = False
        Me.DGVItemDetails.Columns(24).Visible = False
        Me.DGVItemDetails.Columns(25).Visible = False
        Me.DGVItemDetails.Columns(26).Visible = False
        Me.DGVItemDetails.Columns(27).Visible = False
        Me.DGVItemDetails.Columns(28).Visible = True
        Me.DGVItemDetails.Columns(29).Visible = False
        Me.DGVItemDetails.Columns(30).Visible = False
        Me.DGVItemDetails.Columns(31).Visible = False
        Me.DGVItemDetails.Columns(32).Visible = False
        Me.DGVItemDetails.Columns(33).Visible = False
        Me.DGVItemDetails.Columns(34).Visible = False
        Me.DGVItemDetails.Columns(35).Visible = True
        Me.DGVItemDetails.Columns(36).Visible = False
        Me.DGVItemDetails.Columns(37).Visible = False

        Me.DGVItemDetails.Columns(38).Visible = False
        Me.DGVItemDetails.Columns(39).Visible = False
        Me.DGVItemDetails.Columns(40).Visible = False
        Me.DGVItemDetails.Columns(41).Visible = False
        Me.DGVItemDetails.Columns(43).Visible = False
        Me.DGVItemDetails.Columns(45).Visible = False
    End Sub

    Private Sub CMDelete_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles CMDelete.Opening

    End Sub

    Private Sub TBSmartPoint_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBSmartPoint.KeyPress, TBTargetAverage.KeyPress, TBTarget.KeyPress
        On Error GoTo ErrDescription

        Select Case Asc(e.KeyChar)
            Case 48 To 58, 8, 44, 46, 37
            Case Else
                e.Handled = True
        End Select

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TBSmartPoint_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSmartPoint.TextChanged
        Call CalcSmartPoint()
    End Sub

    Public Sub CalcSmartPoint()
        Dim i As Integer
        Dim vItemCode As String
        Dim vSmartPoint As Decimal
        Dim vSmartPointAmount As Double
        Dim vCashPriceOwn1 As Double

        On Error Resume Next

        If Me.TBSmartPoint.Text <> "" Then
            vSmartPoint = Me.TBSmartPoint.Text
        Else
            vSmartPoint = 0
        End If

        For i = 0 To Me.DGVItemDetails.RowCount - 1
            If Me.DGVItemDetails.Item(42, i).Value <> "" Then
                vCashPriceOwn1 = Me.DGVItemDetails.Item(42, i).Value
            Else
                vCashPriceOwn1 = 0
            End If
            vItemCode = Me.DGVItemDetails.Item(1, i).Value

            If vItemCode <> "Nothing" And vItemCode <> "" Then

                If vCashPriceOwn1 <> 0 Then
                    vSmartPointAmount = (vSmartPoint * vCashPriceOwn1) / 100

                    Me.DGVItemDetails.Item(35, i).Value = Format(vSmartPoint, "##,##0.0000")
                    Me.DGVItemDetails.Item(36, i).Value = Format(vSmartPointAmount, "##,##0.000")
                    Call vCalcItemLine(i)
                End If
            End If
        Next

    End Sub

    Private Sub CBDOCash_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBDOCash.CheckedChanged
        On Error Resume Next

        If Me.CBDOCash.Checked = True Then
            Call CalcDOCashPrice()
            Me.CBCash.Enabled = False
            Me.NDDOCash.Enabled = True
            Me.NDDOCash.Focus()
        Else
            Me.CBCash.Enabled = True
            Me.NDDOCash.Enabled = False
        End If
    End Sub

    Private Sub CBDOCredit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBDOCredit.CheckedChanged
        On Error Resume Next

        If Me.CBDOCredit.Checked = True Then
            Call CalcDOCreditPrice()
            Me.NDDOCredit.Enabled = True
            Me.CBCredit.Enabled = False
            Me.NDDOCredit.Focus()
        Else
            Me.CBCredit.Enabled = True
            Me.NDDOCredit.Enabled = False
        End If
    End Sub

    Private Sub CBDOPrice2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBDOPrice2.CheckedChanged
        On Error Resume Next

        If Me.CBDOPrice2.Checked = True Then
            Call CalcDOPrice2()
            Me.NDDOPrice2.Enabled = True
            Me.CBPrice2.Enabled = False
            Me.NDDOPrice2.Focus()
        Else
            Me.CBPrice2.Enabled = True
            Me.NDDOPrice2.Enabled = False
        End If
    End Sub

    Private Sub NDDOCash_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NDDOCash.ValueChanged
        Call CalcDOCashPrice()
    End Sub

    Private Sub NDDOCredit_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NDDOCredit.ValueChanged
        Call CalcDOCreditPrice()
    End Sub

    Private Sub NDDOPrice2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NDDOPrice2.ValueChanged
        Call CalcDOPrice2()
    End Sub

    Private Sub CBBillDisc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBBillDisc.CheckedChanged
        On Error Resume Next

        If Me.CBBillDisc.Checked = True Then
            Me.TBBillDisc.Enabled = True
            Me.TBBillDisc.Focus()
        Else
            Me.TBBillDisc.Enabled = False
        End If
    End Sub

    Private Sub TBBillDisc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBBillDisc.KeyPress, TBFollowDisc1.KeyPress, TBFollowDisc2.KeyPress, TBFollowDisc3.KeyPress
        On Error GoTo ErrDescription

        Select Case Asc(e.KeyChar)
            Case 48 To 58, 8, 46, 37
            Case Else
                e.Handled = True
        End Select

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TBBillDisc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBillDisc.TextChanged
        Dim vDocNo As String
        Dim vItemCode As String
        Dim vBillDisc1 As String
        Dim i As Integer
        Dim vGetDisc As Double


        On Error Resume Next


        If Me.TBBillDisc.Text <> "" Then
            vGetDisc = Me.TBBillDisc.Text

            If vGetDisc >= 100 Then
                MsgBox("กรุณาตรวจสอบส่วนลดหน้าบิล", MsgBoxStyle.Information, "Send Information")
                Me.TBBillDisc.Text = ""
                Me.TBBillDisc.Focus()
                Exit Sub
            End If
        End If

        If Me.DGVItemDetails.RowCount > 0 Then

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                vItemCode = Me.DGVItemDetails.Item(1, i).Value
                vBillDisc1 = Me.TBBillDisc.Text
                If vItemCode <> "Nothing" And vItemCode <> "" Then
                    Me.DGVItemDetails.Item(5, i).Value = vBillDisc1
                    Call vCalcItemLine(i)
                End If
            Next
        End If
    End Sub

    Private Sub CBFollowDisc1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBFollowDisc1.CheckedChanged
        On Error Resume Next

        If Me.CBFollowDisc1.Checked = True Then
            Me.TBFollowDisc1.Enabled = True
            Me.TBFollowDisc1.Focus()
        Else
            Me.TBFollowDisc1.Enabled = False
        End If
    End Sub

    Private Sub CBFollowDisc2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBFollowDisc2.CheckedChanged
        On Error Resume Next

        If Me.CBFollowDisc2.Checked = True Then
            Me.TBFollowDisc2.Enabled = True
            Me.TBFollowDisc2.Focus()
        Else
            Me.TBFollowDisc2.Enabled = False
        End If
    End Sub

    Private Sub CBFollowDisc3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBFollowDisc3.CheckedChanged
        On Error Resume Next

        If Me.CBFollowDisc3.Checked = True Then
            Me.TBFollowDisc3.Enabled = True
            Me.TBFollowDisc3.Focus()
        Else
            Me.TBFollowDisc3.Enabled = False
        End If
    End Sub

    Private Sub TBFollowDisc1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBFollowDisc1.TextChanged
        Dim vDocNo As String
        Dim vItemCode As String
        Dim vFollowDisc1 As String
        Dim i As Integer
        Dim vGetDisc As Double


        On Error Resume Next

        If Me.TBFollowDisc1.Text <> "" Then
            vGetDisc = Me.TBFollowDisc1.Text

            If vGetDisc >= 100 Then
                MsgBox("กรุณาตรวจสอบส่วนลดตาม1", MsgBoxStyle.Information, "Send Information")
                Me.TBFollowDisc1.Text = ""
                Me.TBFollowDisc1.Focus()
                Exit Sub
            End If
        End If

        If Me.DGVItemDetails.RowCount > 0 Then

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                vItemCode = Me.DGVItemDetails.Item(1, i).Value
                vFollowDisc1 = Me.TBFollowDisc1.Text
                If vItemCode <> "Nothing" And vItemCode <> "" Then
                    Me.DGVItemDetails.Item(7, i).Value = vFollowDisc1
                    Call vCalcItemLine(i)
                End If
            Next
        End If
    End Sub

    Private Sub TBFollowDisc2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBFollowDisc2.TextChanged
        Dim vDocNo As String
        Dim vItemCode As String
        Dim vFollowDisc2 As String
        Dim i As Integer
        Dim vGetDisc As Double

        On Error Resume Next

        If Me.TBFollowDisc2.Text <> "" Then
            vGetDisc = Me.TBFollowDisc2.Text

            If vGetDisc >= 100 Then
                MsgBox("กรุณาตรวจสอบส่วนลดตาม2", MsgBoxStyle.Information, "Send Information")
                Me.TBFollowDisc2.Text = ""
                Me.TBFollowDisc2.Focus()
                Exit Sub
            End If
        End If

        If Me.DGVItemDetails.RowCount > 0 Then

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                vItemCode = Me.DGVItemDetails.Item(1, i).Value
                vFollowDisc2 = Me.TBFollowDisc2.Text
                If vItemCode <> "Nothing" And vItemCode <> "" Then
                    Me.DGVItemDetails.Item(9, i).Value = vFollowDisc2
                    Call vCalcItemLine(i)
                End If
            Next
        End If
    End Sub

    Private Sub TBFollowDisc3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBFollowDisc3.TextChanged
        Dim vDocNo As String
        Dim vItemCode As String
        Dim vFollowDisc3 As String
        Dim i As Integer
        Dim vGetDisc As Double


        On Error Resume Next

        If Me.TBFollowDisc3.Text <> "" Then
            vGetDisc = Me.TBFollowDisc3.Text

            If vGetDisc >= 100 Then
                MsgBox("กรุณาตรวจสอบส่วนลดตาม3", MsgBoxStyle.Information, "Send Information")
                Me.TBFollowDisc3.Text = ""
                Me.TBFollowDisc3.Focus()
                Exit Sub
            End If
        End If

        If Me.DGVItemDetails.RowCount > 0 Then

            For i = 0 To Me.DGVItemDetails.RowCount - 1
                vItemCode = Me.DGVItemDetails.Item(1, i).Value
                vFollowDisc3 = Me.TBFollowDisc3.Text
                If vItemCode <> "Nothing" And vItemCode <> "" Then
                    Me.DGVItemDetails.Item(11, i).Value = vFollowDisc3
                    Call vCalcItemLine(i)
                End If
            Next
        End If
    End Sub

    Private Sub DGVItemDetails_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItemDetails.CellContentClick

    End Sub
End Class