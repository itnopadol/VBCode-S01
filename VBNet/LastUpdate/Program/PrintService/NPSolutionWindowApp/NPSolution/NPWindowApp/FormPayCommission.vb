Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Public Class FormPayCommission
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vIsOpen As Integer
    Dim vMemIsCancel As Integer
    Dim vMemIsConfirm As Integer
    Dim vMemBeginTran As Integer

    Private Sub FormPayCommission_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()
        Call NewDoc()
        Call vGenDocNoAuto()
    End Sub

    Public Sub CalcCommissionAmount()
        Dim i As Integer
        Dim vComAmount As Double
        Dim vSumComAmount As Double
        Dim vTotalComAmount As Double

        On Error Resume Next

        For i = 0 To Me.ListViewPayCommission.Items.Count - 1
            vComAmount = Me.ListViewPayCommission.Items(i).SubItems(12).Text
            vSumComAmount = vSumComAmount + vComAmount
        Next

        vTotalComAmount = vSumComAmount
        Me.TBComAmount.Text = Format(vTotalComAmount, "##,##0.00")
    End Sub

    Public Sub vGenLineID()
        Dim i As Integer
        Dim n As Integer

        For i = 0 To Me.ListViewPayCommission.Items.Count - 1
            n = n + 1
            Me.ListViewPayCommission.Items(i).SubItems(0).Text = n
        Next
    End Sub

    Private Sub BTNSearchRequest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchRequest.Click
        Me.PNSearchRemainPaid.Visible = True
        Me.BTNRemainPaid.Focus()

        'Dim i As Integer
        'Dim n As Integer
        'Dim vListDoc As ListViewItem
        'Dim vConditionQty As Double
        'Dim vQty As Double
        'Dim vPriceSet As Double
        'Dim vPrice As Double
        'Dim vComSet As Double
        'Dim vCom As Double
        'Dim vProfitCenter As String
        'Dim vPayCondition As Integer
        'Dim vDocDate1 As String
        'Dim vDocDate2 As String

        'On Error Resume Next

        'If Me.CMBAgent.Text <> "" Then
        '    vProfitCenter = vb6.Left(Me.CMBAgent.Text, vb6.InStr(Me.CMBAgent.Text, "/") - 1)
        '    If Me.CBPayCondition.Checked = True Then
        '        vPayCondition = 1
        '    Else
        '        vPayCondition = 0
        '    End If

        '    Me.DTPDate1.Value = DateAdd(DateInterval.Day, 30, Me.DTPDate1.Value)
        '    vDocDate1 = Day(Me.DTPDate1.Value) & "/" & Month(Me.DTPDate1.Value) & "/" & Year(Me.DTPDate1.Value)
        '    vDocDate2 = Day(Me.DTPDate2.Value) & "/" & Month(Me.DTPDate2.Value) & "/" & Year(Me.DTPDate2.Value)

        '    Me.ListViewRemainPaid.Items.Clear()
        '    vQuery = "exec dbo.USP_COM_PaidWaiting '" & vProfitCenter & "'," & vPayCondition & ",'" & vDocDate1 & "','" & vDocDate2 & "'"
        '    da = New SqlDataAdapter(vQuery, vConnection)
        '    ds = New DataSet
        '    da.Fill(ds, "SearchPaidWaiting")
        '    dt = ds.Tables("SearchPaidWaiting")
        '    If dt.Rows.Count > 0 Then

        '        For i = 0 To dt.Rows.Count - 1
        '            n = n + 1
        '            vListDoc = Me.ListViewRemainPaid.Items.Add(n)
        '            vListDoc.SubItems.Add(0).Text = dt.Rows(i).Item("campaign")
        '            vListDoc.SubItems.Add(1).Text = dt.Rows(i).Item("salename")
        '            If dt.Rows(i).Item("saletype") = 0 Then
        '                vListDoc.SubItems.Add(2).Text = "ขายเงินสด"
        '            Else
        '                vListDoc.SubItems.Add(2).Text = "ขายเงินเชื่อ"
        '            End If
        '            vListDoc.SubItems.Add(3).Text = dt.Rows(i).Item("saledocno")
        '            vListDoc.SubItems.Add(4).Text = dt.Rows(i).Item("refdocdate")
        '            vListDoc.SubItems.Add(6).Text = dt.Rows(i).Item("item")
        '            vListDoc.SubItems.Add(7).Text = dt.Rows(i).Item("unitcode")
        '            vConditionQty = dt.Rows(i).Item("ConditionQty")
        '            vQty = dt.Rows(i).Item("qty")
        '            vPriceSet = dt.Rows(i).Item("priceset")
        '            vPrice = dt.Rows(i).Item("price")
        '            vComSet = dt.Rows(i).Item("comset")
        '            vCom = dt.Rows(i).Item("com")
        '            vListDoc.SubItems.Add(8).Text = Format(vConditionQty, "##,##0.00")
        '            vListDoc.SubItems.Add(9).Text = Format(vQty, "##,##0.00")
        '            vListDoc.SubItems.Add(10).Text = Format(vPriceSet, "##,##0.00")
        '            vListDoc.SubItems.Add(11).Text = Format(vPrice, "##,##0.00")
        '            vListDoc.SubItems.Add(12).Text = Format(vComSet, "##,##0.00")
        '            vListDoc.SubItems.Add(13).Text = Format(vCom, "##,##0.00")
        '            vListDoc.SubItems.Add(14).Text = dt.Rows(i).Item("linenumber")
        '        Next
        '    End If

        '    If Me.ListViewRemainPaid.Items.Count > 0 Then
        '        Me.PNSearchRemainPaid.Visible = True
        '        Me.ListViewRemainPaid.Focus()
        '        Me.ListViewRemainPaid.Items(0).Selected = True
        '        Me.ListViewRemainPaid.Items(0).Focused = True
        '    End If
        'Else
        '    MsgBox("", MsgBoxStyle.Critical, "Send Error Message")
        '    Me.BTNSearchRequest.Focus()
        'End If
    End Sub

    Public Sub SearchReqComm()
        Me.PNSearchRemainPaid.Visible = True
        Me.BTNRemainPaid.Focus()
    End Sub

    Public Sub SearchCommRemainPaid()
        Dim i As Integer
        Dim n As Integer
        Dim vListDoc As ListViewItem
        Dim vConditionQty As Double
        Dim vQty As Double
        Dim vPriceSet As Double
        Dim vPrice As Double
        Dim vComSet As Double
        Dim vCom As Double
        Dim vProfitCenter As String
        Dim vPayCondition As Integer
        Dim vDocDate1 As String
        Dim vDocDate2 As String

        On Error Resume Next

        If Me.CMBAgent.Text <> "" Then
            vProfitCenter = vb6.Left(Me.CMBAgent.Text, vb6.InStr(Me.CMBAgent.Text, "/") - 1)
            If Me.CBPayCondition.Checked = True Then
                vPayCondition = 1
            Else
                vPayCondition = 0
            End If

            vDocDate1 = Day(Me.DTPDate1.Value) & "/" & Month(Me.DTPDate1.Value) & "/" & Year(Me.DTPDate1.Value)
            vDocDate2 = Day(Me.DTPDate2.Value) & "/" & Month(Me.DTPDate2.Value) & "/" & Year(Me.DTPDate2.Value)

            Me.ListViewRemainPaid.Items.Clear()
            vQuery = "exec dbo.USP_COM_PaidWaiting '" & vProfitCenter & "'," & vPayCondition & ",'" & vDocDate1 & "','" & vDocDate2 & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "SearchPaidWaiting")
            dt = ds.Tables("SearchPaidWaiting")
            If dt.Rows.Count > 0 Then

                For i = 0 To dt.Rows.Count - 1
                    n = n + 1
                    vListDoc = Me.ListViewRemainPaid.Items.Add(n)
                    vListDoc.SubItems.Add(0).Text = dt.Rows(i).Item("campaign")
                    vListDoc.SubItems.Add(1).Text = dt.Rows(i).Item("salename")
                    If dt.Rows(i).Item("saletype") = 0 Then
                        vListDoc.SubItems.Add(2).Text = "ขายเงินสด"
                    Else
                        vListDoc.SubItems.Add(2).Text = "ขายเงินเชื่อ"
                    End If
                    vListDoc.SubItems.Add(3).Text = dt.Rows(i).Item("saledocno")
                    vListDoc.SubItems.Add(4).Text = dt.Rows(i).Item("refdocdate")
                    vListDoc.SubItems.Add(6).Text = dt.Rows(i).Item("item")
                    vListDoc.SubItems.Add(7).Text = dt.Rows(i).Item("unitcode")
                    vConditionQty = dt.Rows(i).Item("ConditionQty")
                    vQty = dt.Rows(i).Item("qty")
                    vPriceSet = dt.Rows(i).Item("priceset")
                    vPrice = dt.Rows(i).Item("price")
                    vComSet = dt.Rows(i).Item("comset")
                    vCom = dt.Rows(i).Item("com")
                    vListDoc.SubItems.Add(8).Text = Format(vConditionQty, "##,##0.00")
                    vListDoc.SubItems.Add(9).Text = Format(vQty, "##,##0.00")
                    vListDoc.SubItems.Add(10).Text = Format(vPriceSet, "##,##0.00")
                    vListDoc.SubItems.Add(11).Text = Format(vPrice, "##,##0.00")
                    vListDoc.SubItems.Add(12).Text = Format(vComSet, "##,##0.00")
                    vListDoc.SubItems.Add(13).Text = Format(vCom, "##,##0.00")
                    vListDoc.SubItems.Add(14).Text = dt.Rows(i).Item("linenumber")
                Next
            End If

            If Me.ListViewRemainPaid.Items.Count > 0 Then
                Me.PNSearchRemainPaid.Visible = True
                Me.ListViewRemainPaid.Focus()
                Me.ListViewRemainPaid.Items(0).Selected = True
                Me.ListViewRemainPaid.Items(0).Focused = True
            End If
        Else
            MsgBox("กรุณาเลือก ProfitCenter ก่อนเลือกเอกสารค้างจ่าย กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNSearchRequest.Focus()
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vDocNo As String
        Dim vDocDate As String
        Dim vItemCode As String
        Dim vIsInsert As Integer
        Dim i As Integer
        Dim vReqNo As String
        Dim vLineDoc As Integer
        Dim vProfitCenter As String


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

        If Me.TBDocNo.Text = "" Then
            MsgBox("กรุณากรอกเลขที่เอกสารจ่ายค่าคอมฯ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNDocNo.Focus()
            Exit Sub
        End If

        If Me.CMBAgent.Text = "" Then
            MsgBox("กรุณาเลือกศูนย์ธุรกิจในการจ่ายค่าคอมฯ", MsgBoxStyle.Critical, "Send Error Message")
            Me.CMBAgent.Focus()
            Exit Sub
        End If

        If Me.ListViewPayCommission.Items.Count = 0 Then
            MsgBox("ไม่มีรายการเอกสารเสนอสินค้าคิดค่าคอมฯ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNSearchRequest.Focus()
            Exit Sub
        End If

        If Me.ListViewPayCommission.Items.Count > 0 Then
            If vIsOpen = 0 Then
                vIsInsert = 1
            Else
                vIsInsert = 0
            End If

            vProfitCenter = vb6.Left(Me.CMBAgent.Text, vb6.InStr(Me.CMBAgent.Text, "/") - 1)

            vDocNo = Me.TBDocNo.Text
            vDocDate = vb6.Day(Me.DTPDocDate.Value) & "/" & vb6.Month(Me.DTPDocDate.Value) & "/" & vb6.Year(Me.DTPDocDate.Value)

            On Error GoTo ErrDescription

            vQuery = "begin tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            vMemBeginTran = 1

            vQuery = "exec dbo.usp_com_paidsave " & vIsInsert & ",'" & vDocNo & "','" & vDocDate & "','" & vProfitCenter & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            For i = 0 To Me.ListViewPayCommission.Items.Count - 1
                vItemCode = vb6.Left(Me.ListViewPayCommission.Items(i).SubItems(5).Text, InStr(Me.ListViewPayCommission.Items(i).SubItems(5).Text, "/") - 1)
                vReqNo = Me.ListViewPayCommission.Items(i).SubItems(3).Text
                vLineDoc = Me.ListViewPayCommission.Items(i).SubItems(13).Text
                vQuery = "exec dbo.usp_com_paidsubsave '" & vDocNo & "','" & vReqNo & "'," & vLineDoc & ",'" & vItemCode & "' "
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()
            Next

            vQuery = "exec dbo.usp_com_paidcompletesave '" & vDocNo & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            vQuery = "commit tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            MsgBox("บันทึกข้อมูลเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")

            Call ClearScreen()
            Call vGenDocNoAuto()
            Me.DTPDocDate.Focus()
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
        Dim vIsInsert As Integer
        Dim i As Integer
        Dim vReqNo As String
        Dim vLineDoc As Integer
        Dim vProfitCenter As String


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

        If Me.TBDocNo.Text = "" Then
            MsgBox("กรุณากรอกเลขที่เอกสารจ่ายค่าคอมฯ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNDocNo.Focus()
            Exit Sub
        End If

        If Me.CMBAgent.Text = "" Then
            MsgBox("กรุณาเลือกศูนย์ธุรกิจในการจ่ายค่าคอมฯ", MsgBoxStyle.Critical, "Send Error Message")
            Me.CMBAgent.Focus()
            Exit Sub
        End If

        If Me.ListViewPayCommission.Items.Count = 0 Then
            MsgBox("ไม่มีรายการเอกสารเสนอสินค้าคิดค่าคอมฯ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNSearchRequest.Focus()
            Exit Sub
        End If

        If Me.ListViewPayCommission.Items.Count > 0 Then
            If vIsOpen = 0 Then
                vIsInsert = 1
            Else
                vIsInsert = 0
            End If

            vProfitCenter = vb6.Left(Me.CMBAgent.Text, vb6.InStr(Me.CMBAgent.Text, "/") - 1)

            vDocNo = Me.TBDocNo.Text
            vDocDate = vb6.Day(Me.DTPDocDate.Value) & "/" & vb6.Month(Me.DTPDocDate.Value) & "/" & vb6.Year(Me.DTPDocDate.Value)

            On Error GoTo ErrDescription

            vQuery = "begin tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            vMemBeginTran = 1

            vQuery = "exec dbo.usp_com_paidsave " & vIsInsert & ",'" & vDocNo & "','" & vDocDate & "','" & vProfitCenter & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            For i = 0 To Me.ListViewPayCommission.Items.Count - 1
                vReqNo = Me.ListViewPayCommission.Items(i).SubItems(3).Text
                vLineDoc = Me.ListViewPayCommission.Items(i).SubItems(13).Text
                vQuery = "exec dbo.usp_com_paidsubsave '" & vDocNo & "','" & vReqNo & "'," & vLineDoc & " "
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()
            Next

            vQuery = "exec dbo.usp_com_paidcompletesave '" & vDocNo & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            vQuery = "commit tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            MsgBox("บันทึกข้อมูลเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")

            Call ClearScreen()
            Call vGenDocNoAuto()
            Me.DTPDocDate.Focus()
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

    Public Sub NewDoc()
        Me.PBNew.Visible = True
        Me.PBCancel.Visible = False
        Me.PBConfirm.Visible = False
        Me.CMBAgent.SelectedIndex = 0
        Me.DTPDocDate.Value = Now
        Me.DTPDate1.Value = Now
        Me.DTPDate2.Value = Now
    End Sub

    Public Sub CancelDoc()
        Me.PBNew.Visible = False
        Me.PBCancel.Visible = True
        Me.PBConfirm.Visible = False
    End Sub

    Public Sub ConfirmDoc()
        Me.PBNew.Visible = False
        Me.PBCancel.Visible = False
        Me.PBConfirm.Visible = True
    End Sub

    Private Sub BTNSelectRemainPaid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectRemainPaid.Click
        Dim i As Integer
        Dim n As Integer
        Dim vListDoc As ListViewItem
        Dim vConditionQty As Double
        Dim vQty As Double
        Dim vPriceSet As Double
        Dim vPrice As Double
        Dim vComSet As Double
        Dim vCom As Double

        Dim vDocNo As String
        Dim vItemCode As String
        Dim vLineItem As Integer
        Dim vUnitCode As String

        Dim vCheckDocNo As String
        Dim vCheckItemCode As String
        Dim vCheckLineItem As Integer
        Dim vCheckUnitCode As String
        Dim a As Integer

        On Error Resume Next

        If Me.ListViewRemainPaid.Items.Count > 0 Then

            For i = 0 To Me.ListViewRemainPaid.Items.Count - 1
                If Me.ListViewRemainPaid.Items(i).Checked = True Then
                    n = Me.ListViewPayCommission.Items.Count + 1

                    vDocNo = Me.ListViewRemainPaid.Items(i).SubItems(4).Text
                    vItemCode = vb6.Left(Me.ListViewRemainPaid.Items(i).SubItems(6).Text, vb6.InStr(Me.ListViewRemainPaid.Items(i).SubItems(6).Text, "/") - 1)
                    vLineItem = Me.ListViewRemainPaid.Items(i).SubItems(14).Text
                    vUnitCode = Me.ListViewRemainPaid.Items(i).SubItems(7).Text

                    If Me.ListViewPayCommission.Items.Count Then
                        For a = 0 To Me.ListViewPayCommission.Items.Count - 1
                            vCheckDocNo = Me.ListViewPayCommission.Items(a).SubItems(3).Text
                            vCheckItemCode = vb6.Left(Me.ListViewPayCommission.Items(a).SubItems(5).Text, vb6.InStr(Me.ListViewPayCommission.Items(a).SubItems(5).Text, "/") - 1)
                            vCheckLineItem = Me.ListViewPayCommission.Items(a).SubItems(13).Text
                            vCheckUnitCode = Me.ListViewPayCommission.Items(a).SubItems(6).Text

                            If vDocNo = vCheckDocNo And vItemCode = vCheckItemCode And vLineItem = vCheckLineItem And vUnitCode = vCheckUnitCode Then
                                GoTo Line1
                            End If
                        Next
                    End If

                    vListDoc = Me.ListViewPayCommission.Items.Add(n)
                    vListDoc.SubItems.Add(0).Text = Me.ListViewRemainPaid.Items(i).SubItems(2).Text
                    vListDoc.SubItems.Add(1).Text = Me.ListViewRemainPaid.Items(i).SubItems(3).Text

                    vListDoc.SubItems.Add(2).Text = Me.ListViewRemainPaid.Items(i).SubItems(4).Text

                    vListDoc.SubItems.Add(3).Text = Me.ListViewRemainPaid.Items(i).SubItems(5).Text
                    vListDoc.SubItems.Add(4).Text = Me.ListViewRemainPaid.Items(i).SubItems(6).Text
                    vListDoc.SubItems.Add(5).Text = Me.ListViewRemainPaid.Items(i).SubItems(7).Text
                    vListDoc.SubItems.Add(6).Text = Me.ListViewRemainPaid.Items(i).SubItems(8).Text
                    vConditionQty = Me.ListViewRemainPaid.Items(i).SubItems(9).Text
                    vQty = Me.ListViewRemainPaid.Items(i).SubItems(10).Text
                    vPriceSet = Me.ListViewRemainPaid.Items(i).SubItems(11).Text
                    vPrice = Me.ListViewRemainPaid.Items(i).SubItems(12).Text
                    vComSet = Me.ListViewRemainPaid.Items(i).SubItems(13).Text
                    vCom = Me.ListViewRemainPaid.Items(i).SubItems(14).Text
                    vListDoc.SubItems.Add(7).Text = Format(vConditionQty, "##,##0.00")
                    vListDoc.SubItems.Add(8).Text = Format(vQty, "##,##0.00")
                    vListDoc.SubItems.Add(9).Text = Format(vPriceSet, "##,##0.00")
                    vListDoc.SubItems.Add(10).Text = Format(vPrice, "##,##0.00")
                    vListDoc.SubItems.Add(11).Text = Format(vComSet, "##,##0.00")
                    vListDoc.SubItems.Add(12).Text = Format(vCom, "##,##0.00")
                    vListDoc.SubItems.Add(13).Text = vLineItem
                End If
Line1:
            Next

            Me.PNSearchRemainPaid.Visible = False
        End If

        Call CalcCommissionAmount()
    End Sub

    Private Sub BTNCloseRemainPaid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseRemainPaid.Click
        On Error Resume Next

        Me.CBSelectAll.Checked = False
        Me.ListViewRemainPaid.Items.Clear()
        Me.PNSearchRemainPaid.Visible = False
        Me.DTPDocDate.Focus()
    End Sub

    Private Sub ListViewPayCommission_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewPayCommission.KeyDown
        Dim vIndex As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Delete Then
            If Me.ListViewPayCommission.Items.Count > 0 Then
                vIndex = Me.ListViewPayCommission.SelectedItems(0).Index
                Me.ListViewPayCommission.Items.RemoveAt(vIndex)
                Call CalcCommissionAmount()
                Call vGenLineID()
            End If
        End If
    End Sub

    Private Sub ListViewPayCommission_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewPayCommission.SelectedIndexChanged

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
        vQuery = "select dbo.ft_com_newpaid ('" & vNow & "') as docno"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchDocno")
        dt = ds.Tables("SearchDocno")
        If dt.Rows.Count > 0 Then
            Me.TBDocNo.Text = dt.Rows(0).Item("docno")
            Me.DTPDocDate.Focus()
        Else
            Me.TBDocNo.Text = ""
            MsgBox("กำหนด เลขที่เอกสารไม่ได้ เกิดปัญหา จาก Store : dbo.ft_com_newpaid  ติดต่อแผนกคอมพิวเตอร์ เบอร์ 702", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNDocNo.Focus()
        End If
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Me.Close()
    End Sub

    Private Sub BTNClearScreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearScreen.Click
        Call ClearScreen()
        Call vGenDocNoAuto()
        Me.DTPDocDate.Focus()
    End Sub

    Public Sub ClearScreen()
        On Error Resume Next

        vIsOpen = 0
        vMemIsCancel = 0
        vMemIsConfirm = 0

        Call vGenDocNoAuto()
        Call NewDoc()
        Me.DTPDocDate.Value = Now
        Me.CMBAgent.SelectedIndex = 0
        Me.ListViewPayCommission.Items.Clear()
        Me.TBComAmount.Text = ""
        Me.BTNDocNo.Focus()
    End Sub

    Private Sub CBSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBSelectAll.CheckedChanged
        Dim i As Integer

        On Error Resume Next

        If Me.ListViewRemainPaid.Items.Count > 0 Then
            If Me.CBSelectAll.Checked = True Then
                For i = 0 To Me.ListViewRemainPaid.Items.Count - 1
                    Me.ListViewRemainPaid.Items(i).Checked = True
                Next
            End If

            If Me.CBSelectAll.Checked = False Then
                For i = 0 To Me.ListViewRemainPaid.Items.Count - 1
                    Me.ListViewRemainPaid.Items(i).Checked = False
                Next
            End If
        End If
    End Sub

    Private Sub CBSelectAll_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CBSelectAll.KeyDown, ListViewRemainPaid.KeyDown, BTNSelectRemainPaid.KeyDown, BTNCloseRemainPaid.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.CBSelectAll.Checked = False
            Me.ListViewRemainPaid.Items.Clear()
            Me.PNSearchRemainPaid.Visible = False
            Me.DTPDocDate.Focus()
        End If
    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        Call SearchPaidNo()
    End Sub

    Private Sub BTNSearchPaidNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchPaidNo.Click
        Dim vSearch As String

        vSearch = Me.TBSearchPaidNo.Text
        Call SearchPaidNo()
    End Sub

    Public Sub SearchPaidNo()
        Dim i As Integer
        Dim n As Integer
        Dim vListDoc As ListViewItem
        Dim vSearch As String

        On Error Resume Next

        Me.ListViewSearchPaidNo.Items.Clear()
        vSearch = Me.TBSearchPaidNo.Text
        vQuery = "exec dbo.USP_COM_PaidSearch '" & vSearch & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchPaidNo")
        dt = ds.Tables("SearchPaidNo")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vListDoc = Me.ListViewSearchPaidNo.Items.Add(n)
                vListDoc.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                vListDoc.SubItems.Add(1).Text = dt.Rows(i).Item("docdate")
                vListDoc.SubItems.Add(2).Text = dt.Rows(i).Item("totalamount")
                vListDoc.SubItems.Add(3).Text = dt.Rows(i).Item("Profitcenter")
                vListDoc.SubItems.Add(4).Text = ""
            Next
        End If

        If Me.ListViewSearchPaidNo.Items.Count > 0 Then
            Me.PNSearchPaidNo.Visible = True
            Me.ListViewSearchPaidNo.Focus()
            Me.ListViewSearchPaidNo.Items(0).Selected = True
            Me.ListViewSearchPaidNo.Items(0).Focused = True
        End If
    End Sub

    Private Sub BTNCloseSearchPaidNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSearchPaidNo.Click
        Me.PNSearchPaidNo.Visible = False
        Me.DTPDocDate.Focus()
    End Sub

    Private Sub ListViewSearchPaidNo_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewSearchPaidNo.DoubleClick
        Dim vIndex As Integer
        Dim vDocNo As String
        Dim vProfitCenter As String

        On Error Resume Next

        If Me.ListViewSearchPaidNo.Items.Count > 0 Then
            vIndex = Me.ListViewSearchPaidNo.SelectedItems(0).Index
            vDocNo = Me.ListViewSearchPaidNo.Items(vIndex).SubItems(1).Text
            vProfitCenter = Me.ListViewSearchPaidNo.Items(vIndex).SubItems(4).Text
            Call SearchPaidNoDetails(vDocNo, vProfitCenter)

            If Me.ListViewPayCommission.Items.Count > 0 Then
                Me.PNSearchPaidNo.Visible = False
                Me.ListViewPayCommission.Focus()
                Me.ListViewPayCommission.Items(0).Selected = True
                Me.ListViewPayCommission.Items(0).Focused = True
            Else
                MsgBox("ไม่มีข้อมูลของเอกสารเลขที่ " & vDocNo & " กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.ListViewSearchPaidNo.Focus()
                Me.ListViewSearchPaidNo.Items(vIndex).Selected = True
                Me.ListViewSearchPaidNo.Items(vIndex).Focused = True
            End If
        End If
    End Sub

    Public Sub SearchPaidNoDetails(ByVal vDocNo As String, ByVal vProfitCenter As String)
        Dim i As Integer
        Dim n As Integer
        Dim vListDoc As ListViewItem
        Dim vConditionQty As Double
        Dim vQty As Double
        Dim vPriceSet As Double
        Dim vPrice As Double
        Dim vComSet As Double
        Dim vCom As Double
        Dim vSaleType As Integer

        On Error Resume Next

        If Me.ListViewSearchPaidNo.Items.Count > 0 Then
            Me.ListViewPayCommission.Items.Clear()
            vQuery = "exec dbo.USP_COM_PaidSearch2 '" & vDocNo & "','" & vProfitCenter & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "PaidNo")
            dt = ds.Tables("PaidNo")
            If dt.Rows.Count > 0 Then
                Me.TBDocNo.Text = dt.Rows(i).Item("docno")
                Me.DTPDocDate.Text = dt.Rows(i).Item("docdate")

                vMemIsConfirm = dt.Rows(i).Item("isconfirm")
                vMemIsCancel = dt.Rows(i).Item("iscancel")

                For i = 0 To dt.Rows.Count - 1
                    n = n + 1
                    vListDoc = Me.ListViewPayCommission.Items.Add(n)
                    vListDoc.SubItems.Add(0).Text = dt.Rows(i).Item("salename")
                    If dt.Rows(i).Item("saletype") = 0 Then
                        vListDoc.SubItems.Add(1).Text = "ขายเงินสด"
                    ElseIf dt.Rows(i).Item("saletype") = 1 Then
                        vListDoc.SubItems.Add(1).Text = "ขายเงินเชื่อ"
                    End If

                    If dt.Rows(i).Item("profitcenter") <> "" Then
                        If UCase(dt.Rows(i).Item("profitcenter")) = "S01" Then
                            Me.CMBAgent.SelectedIndex = 0
                        End If
                        If UCase(dt.Rows(i).Item("profitcenter")) = "S02" Then
                            Me.CMBAgent.SelectedIndex = 1
                        End If
                        If UCase(dt.Rows(i).Item("profitcenter")) = "W01" Then
                            Me.CMBAgent.SelectedIndex = 3
                        End If
                    End If

                    vListDoc.SubItems.Add(2).Text = dt.Rows(i).Item("saledocno")

                    vListDoc.SubItems.Add(3).Text = dt.Rows(i).Item("refdocdate")
                    vListDoc.SubItems.Add(4).Text = dt.Rows(i).Item("item")
                    vListDoc.SubItems.Add(5).Text = dt.Rows(i).Item("unitcode")
                    vConditionQty = dt.Rows(i).Item("conditionqty")
                    vQty = dt.Rows(i).Item("qty")
                    vPriceSet = dt.Rows(i).Item("priceset")
                    vPrice = dt.Rows(i).Item("price")
                    vComSet = dt.Rows(i).Item("comset")
                    vCom = dt.Rows(i).Item("com")
                    vListDoc.SubItems.Add(6).Text = Format(vConditionQty, "##,##0.00")
                    vListDoc.SubItems.Add(7).Text = Format(vQty, "##,##0.00")
                    vListDoc.SubItems.Add(8).Text = Format(vPriceSet, "##,##0.00")
                    vListDoc.SubItems.Add(9).Text = Format(vPrice, "##,##0.00")
                    vListDoc.SubItems.Add(10).Text = Format(vComSet, "##,##0.00")
                    vListDoc.SubItems.Add(11).Text = Format(vCom, "##,##0.00")
                    vListDoc.SubItems.Add(12).Text = dt.Rows(i).Item("linenumber")
                Next

                vIsOpen = 1
                Call CalcCommissionAmount()
                Me.TBDocNo.Focus()
                Me.TBDocNo.SelectAll()
            End If
        End If

    End Sub

    Private Sub ListViewSearchPaidNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearchPaidNo.KeyDown
        Dim vIndex As Integer
        Dim vDocNo As String
        Dim vProfitCenter As String

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewSearchPaidNo.Items.Count > 0 Then
                vIndex = Me.ListViewSearchPaidNo.SelectedItems(0).Index
                vDocNo = Me.ListViewSearchPaidNo.Items(vIndex).SubItems(1).Text
                vProfitCenter = Me.ListViewSearchPaidNo.Items(vIndex).SubItems(4).Text
                Call SearchPaidNoDetails(vDocNo, vProfitCenter)

                If Me.ListViewPayCommission.Items.Count > 0 Then
                    Me.PNSearchPaidNo.Visible = False
                    Me.ListViewPayCommission.Focus()
                    Me.ListViewPayCommission.Items(0).Selected = True
                    Me.ListViewPayCommission.Items(0).Focused = True
                Else
                    MsgBox("ไม่มีข้อมูลของเอกสารเลขที่ " & vDocNo & " กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.ListViewSearchPaidNo.Focus()
                    Me.ListViewSearchPaidNo.Items(vIndex).Selected = True
                    Me.ListViewSearchPaidNo.Items(vIndex).Focused = True
                End If
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchPaidNo.Visible = False
            Me.DTPDocDate.Focus()
        End If

        Dim vCheckLine As Integer

        If Me.ListViewSearchPaidNo.Items.Count > 0 Then
            If e.KeyCode = Keys.Up Then
                vCheckLine = Me.ListViewSearchPaidNo.SelectedItems(0).Index
                If vCheckLine = 0 Then
                    Me.TBSearchPaidNo.Focus()
                    Me.TBSearchPaidNo.SelectAll()
                End If
            End If
        End If
    End Sub

    Private Sub ListViewSearchPaidNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewSearchPaidNo.SelectedIndexChanged

    End Sub

    Private Sub BTNPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNPrint.Click
        On Error Resume Next

        If vIsOpen = 1 Then
            vMemPaidCommNo = Me.TBDocNo.Text
            vMemProfitCenter = vb6.Left(Me.CMBAgent.Text, vb6.InStr(Me.CMBAgent.Text, "/") - 1)


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

            If frmReportPaidComm Is Nothing Then
                frmReportPaidComm = New FormReportPaidComm
            Else
                If frmReportPaidComm.IsDisposed Then
                    frmReportPaidComm = New FormReportPaidComm
                End If
            End If

            frmReportPaidComm.Show()
            frmReportPaidComm.BringToFront()
        End If
    End Sub

    Public Sub PrintDocument()
        On Error Resume Next

        If vIsOpen = 1 Then
            vMemPaidCommNo = Me.TBDocNo.Text
            vMemProfitCenter = vb6.Left(Me.CMBAgent.Text, vb6.InStr(Me.CMBAgent.Text, "/") - 1)


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

            If frmReportPaidComm Is Nothing Then
                frmReportPaidComm = New FormReportPaidComm
            Else
                If frmReportPaidComm.IsDisposed Then
                    frmReportPaidComm = New FormReportPaidComm
                End If
            End If

            frmReportPaidComm.Show()
            frmReportPaidComm.BringToFront()
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
        End If
    End Sub

    Private Sub TBDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBDocNo.KeyDown, BTNDocNo.KeyDown, DTPDocDate.KeyDown, CMBAgent.KeyDown, BTNSearchRequest.KeyDown, ListViewPayCommission.KeyDown, BTNClearScreen.KeyDown, BTNSave.KeyDown, BTNCancel.KeyDown, BTNPrint.KeyDown, BTNSearch.KeyDown, BTNExit.KeyDown, TBComAmount.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If

        If e.KeyCode = Keys.F1 Then
            Call SearchPaidNo()
        End If

        If e.KeyCode = Keys.F2 Then
            Call SearchReqComm()
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
    End Sub

    Private Sub TBDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBDocNo.TextChanged

    End Sub

    Private Sub BTNSearchPaidNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSearchPaidNo.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.PNSearchPaidNo.Visible = False
            Me.DTPDocDate.Focus()
        End If
    End Sub

    Private Sub BTNCloseSearchPaidNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCloseSearchPaidNo.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.PNSearchPaidNo.Visible = False
            Me.DTPDocDate.Focus()
        End If
    End Sub

    Private Sub BTNSelectSearchPaidNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectSearchPaidNo.Click
        Dim vIndex As Integer
        Dim vDocNo As String
        Dim vProfitCenter As String

        On Error Resume Next

        If Me.ListViewSearchPaidNo.Items.Count > 0 Then
            vIndex = Me.ListViewSearchPaidNo.SelectedItems(0).Index
            vDocNo = Me.ListViewSearchPaidNo.Items(vIndex).SubItems(1).Text
            vProfitCenter = Me.ListViewSearchPaidNo.Items(vIndex).SubItems(4).Text
            Call SearchPaidNoDetails(vDocNo, vProfitCenter)

            If Me.ListViewPayCommission.Items.Count > 0 Then
                Me.PNSearchPaidNo.Visible = False
                Me.ListViewPayCommission.Focus()
                Me.ListViewPayCommission.Items(0).Selected = True
                Me.ListViewPayCommission.Items(0).Focused = True
            Else
                MsgBox("ไม่มีข้อมูลของเอกสารเลขที่ " & vDocNo & " กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.ListViewSearchPaidNo.Focus()
                Me.ListViewSearchPaidNo.Items(vIndex).Selected = True
                Me.ListViewSearchPaidNo.Items(vIndex).Focused = True
            End If
        End If
    End Sub

    Private Sub BTNSelectSearchPaidNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSelectSearchPaidNo.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.PNSearchPaidNo.Visible = False
            Me.DTPDocDate.Focus()
        End If
    End Sub

    Private Sub TBSearchPaidNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchPaidNo.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.PNSearchPaidNo.Visible = False
            Me.DTPDocDate.Focus()
        End If

        If e.KeyCode = Keys.Enter Then
            Call SearchPaidNo()
        End If
    End Sub

    Private Sub TBSearchPaidNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchPaidNo.TextChanged

    End Sub

    Private Sub ListViewRemainPaid_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewRemainPaid.SelectedIndexChanged

    End Sub

    Private Sub BTNRemainPaid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNRemainPaid.Click
        Call SearchCommRemainPaid()
    End Sub

    Private Sub BTNPrintNetAmount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPrintNetAmount.Click
        On Error Resume Next


        'vMemPaidCommNo = Me.TBDocNo.Text
        'vMemProfitCenter = vb6.Left(Me.CMBAgent.Text, vb6.InStr(Me.CMBAgent.Text, "/") - 1)


        'If vMemIsConfirm = 1 Then
        '    MsgBox("เอกสารถูกอ้างอิงไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
        '    Me.BTNDocNo.Focus()
        '    Exit Sub
        'End If

        'If vMemIsCancel = 1 Then
        '    MsgBox("เอกสารถูกยกเลิกไปแล้วไม่สามารถบันทึกได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
        '    Me.BTNDocNo.Focus()
        '    Exit Sub
        'End If

        If frmReportPaidCommNet Is Nothing Then
            frmReportPaidCommNet = New FormReportPaidCommNet
        Else
            If frmReportPaidCommNet.IsDisposed Then
                frmReportPaidCommNet = New FormReportPaidCommNet
            End If
        End If

        frmReportPaidCommNet.Show()
        frmReportPaidCommNet.BringToFront()

    End Sub
End Class