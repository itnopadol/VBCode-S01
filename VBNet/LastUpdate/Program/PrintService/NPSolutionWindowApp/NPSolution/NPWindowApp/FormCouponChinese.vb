Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Public Class FormCouponChinese
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim dt1 As DataTable
    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vCheckExist As Integer
    Private Sub FormCouponChinese_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()
        'Call InitializeDataBaseBranch()

        Me.TBDocNo.Focus()
    End Sub

    Private Sub TBDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBDocNo.KeyDown
        Dim vDocNo As String

        If Me.TBDocNo.Text <> "" Then
            If e.KeyCode = Keys.Enter Then
                vDocNo = Me.TBDocNo.Text
                Call vGetInvoiceData(vDocNo)
            Else
            End If
        End If
    End Sub

    Private Sub TBDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBDocNo.TextChanged
        Dim vDocno As String

        On Error Resume Next

        If Me.TBDocNo.Text = "" Then
            Me.TBArName.Text = ""
            Me.ListViewItem.Items.Clear()
            Me.TBSumItemAmount.Text = ""
        End If
    End Sub

    Public Sub vGetInvoiceData(ByVal vDocNo As String)
        Dim i As Integer
        Dim vListItem As ListViewItem
        Dim n As Integer
        Dim vItemType As String
        Dim vItemAmount As Double
        Dim vBillAmount As Double
        Dim vDisCountAmount As Double
        Dim vCoupongAmount As Double
        Dim vNetAmount As Double

        On Error Resume Next

        Me.TBArName.Text = ""
        Me.TBSumItemAmount.Text = ""

        Me.ListViewItem.Items.Clear()
        vQuery = "exec dbo.USP_NP_InvoiceItemCheckCoupon '" & vDocNo & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then

            Me.TBArName.Text = dt.Rows(i).Item("arname")
            Me.TBSumItemAmount.Text = dt.Rows(i).Item("family")
            vDisCountAmount = dt.Rows(i).Item("discount")
            vCoupongAmount = dt.Rows(i).Item("coupongamount")
            Me.TBDiscountAmount.Text = Format(vDisCountAmount, "##,##0.00")
            Me.TBDisCountCoupon.Text = Format(vCoupongAmount, "##,##0.00")

            For i = 0 To dt.Rows.Count - 1
                n = n + 1
                vItemType = dt.Rows(i).Item("family")

                vListItem = Me.ListViewItem.Items.Add(n)
                vListItem.SubItems.Add(0).Text = dt.Rows(i).Item("family")
                vListItem.SubItems.Add(1).Text = dt.Rows(i).Item("itemname")
                vItemAmount = dt.Rows(i).Item("amount")
                vListItem.SubItems.Add(2).Text = Format(vItemAmount, "##,##0.00")

                If vItemType = "หมวดสินค้าทั่วไป" Then
                    vBillAmount = vBillAmount + vItemAmount
                End If
            Next

            vNetAmount = (vBillAmount - vDisCountAmount) - vCoupongAmount

            If vNetAmount > 0 Then
                Me.TBSumItemAmount.Text = Format(vNetAmount, "##,##0.00")
            End If
        End If
    End Sub

    Private Sub BTNAddBill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddBill.Click
        Dim i As Integer
        Dim vBillAmount As Double
        Dim vListItem As ListViewItem
        Dim n As Integer
        Dim vMemDocNo As String
        Dim vMemAr As String
        Dim vArCode As String
        Dim vInvoice As String
        Dim vCheckCount As Integer
        Dim vItemAmount As Double

        On Error Resume Next

        If Me.TBArName.Text <> "" And Me.ListViewItem.Items.Count > 0 And Me.TBSumItemAmount.Text <> "" Then

            vItemAmount = Me.TBSumItemAmount.Text

            If vItemAmount <= 0 Then
                MsgBox("เอกสารนี้ มูลค่าบิลไม่สามารถแลกคูปองได้  กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBDocNo.Focus()
                Exit Sub
            End If

            vInvoice = Me.TBDocNo.Text

            vQuery = "select isnull(count(docno),0)as vCount from npmaster.dbo.TB_NP_ChineseCouponLogs where docno = '" & vInvoice & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "Docno")
            dt = ds.Tables("Docno")
            If dt.Rows.Count > 0 Then
                vCheckCount = dt.Rows(i).Item("vCount")
            End If

            If vCheckCount > 0 Then
                MsgBox("เอกสารนี้ ได้บันทึกการเบิกคูปองแล้ว  กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBDocNo.Focus()
                Exit Sub
            End If

            vArCode = vb6.Left(Me.TBArName.Text, vb6.InStr(Me.TBArName.Text, "/") - 1)
            If Me.TBSumItemAmount.Text <> "" Then
                vBillAmount = Me.TBSumItemAmount.Text
            End If

            If vBillAmount = 0 Then
                MsgBox("เอกสารนี้ ไม่มีสินค้าหมวดทั่วไป กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBDocNo.Focus()
                Exit Sub
            End If

            If Me.ListViewInvoice.Items.Count = 0 Then
                n = 1
                vListItem = Me.ListViewInvoice.Items.Add(n)
                vListItem.SubItems.Add(0).Text = vInvoice
                vListItem.SubItems.Add(1).Text = vArCode
                vListItem.SubItems.Add(2).Text = Format(vBillAmount, "##,##0.00")

                Call vCalcCoupon()

                Me.TBDocNo.Text = ""
                Me.TBSumItemAmount.Text = ""
                Exit Sub
            End If

            If Me.ListViewInvoice.Items.Count > 0 Then

                For i = 0 To Me.ListViewInvoice.Items.Count - 1
                    vMemDocNo = Me.ListViewInvoice.Items(i).SubItems(1).Text
                    vMemAr = Me.ListViewInvoice.Items(i).SubItems(2).Text
                    If vArCode = vMemAr Then
                        If vInvoice = vMemDocNo Then
                            MsgBox("มีเอกสารอยู่แล้ว กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                            Me.TBDocNo.Focus()
                            Exit Sub
                        End If
                    Else
                        MsgBox("ลูกค้าคนละรหัส กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                        Me.TBDocNo.Focus()
                        Exit Sub
                    End If
                Next

                n = Me.ListViewInvoice.Items.Count + 1
                vListItem = Me.ListViewInvoice.Items.Add(n)
                vListItem.SubItems.Add(0).Text = Me.TBDocNo.Text
                vListItem.SubItems.Add(1).Text = vb6.Left(Me.TBArName.Text, vb6.InStr(Me.TBArName.Text, "/") - 1)
                vListItem.SubItems.Add(2).Text = Format(vBillAmount, "##,##0.00")
            End If

            Call vCalcCoupon()

            Me.TBDocNo.Text = ""
            Me.TBSumItemAmount.Text = ""

        End If
    End Sub

    Public Sub vCalcCoupon()
        Dim i As Integer
        Dim vTotalAmount As Double
        Dim vBillAmount As Double
        Dim vCouponCount As Integer
        Dim vRemainAmout As Double

        On Error Resume Next

        If Me.ListViewInvoice.Items.Count > 0 Then
            For i = 0 To Me.ListViewInvoice.Items.Count - 1
                vBillAmount = Me.ListViewInvoice.Items(i).SubItems(3).Text
                vTotalAmount = vTotalAmount + vBillAmount
            Next

            vCouponCount = Math.Truncate(vTotalAmount / 2000)


            vRemainAmout = 2000 - (vTotalAmount Mod 2000)


            Me.TBCouponAmount.Text = Format(vTotalAmount, "##,##0.00")
            Me.TBCoupon.Text = vCouponCount
            If vRemainAmout > 0 Then
                Me.TBRemainAmount.Text = Format(vRemainAmout, "##,##0.00")
            Else
                Me.TBRemainAmount.Text = ""
            End If
        End If

        If Me.ListViewInvoice.Items.Count = 0 Then
            Me.TBCouponAmount.Text = ""
            Me.TBCoupon.Text = ""
            Me.TBRemainAmount.Text = ""
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vDocno As String
        Dim vNetAmount As Double
        Dim vARCode As String
        Dim i As Integer

        On Error Resume Next

        If Me.ListViewInvoice.Items.Count > 0 Then
            For i = 0 To Me.ListViewInvoice.Items.Count - 1

                vDocno = Me.ListViewInvoice.Items(i).SubItems(1).Text
                vARCode = Me.ListViewInvoice.Items(i).SubItems(2).Text
                vNetAmount = Me.ListViewInvoice.Items(i).SubItems(3).Text

                vQuery = ("insert npmaster.dbo.TB_NP_ChineseCouponLogs (docno,arcode,netamount,creatorcode,createdatetime) values('" & vDocno & "','" & vARCode & "'," & vNetAmount & ",'" & vUserID & "',getdate())")
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()
            Next

            MsgBox("บันทึกเรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")
            Me.TBDocNo.Text = ""
            Me.ListViewInvoice.Items.Clear()
            Me.TBCoupon.Text = ""
            Me.TBCouponAmount.Text = ""
            Me.TBRemainAmount.Text = ""
            Me.TBDocNo.Focus()
        End If
    End Sub

    Private Sub ListViewInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewInvoice.KeyDown
        Dim vIndex As Integer

        On Error Resume Next

        If Me.ListViewInvoice.Items.Count > 0 Then
            vIndex = Me.ListViewInvoice.SelectedItems(0).Index

            Me.ListViewInvoice.Items.RemoveAt(vIndex)
            Call vCalcCoupon()
            Call GenLineNumber()
            Me.TBDocNo.Focus()

        End If
    End Sub

    Public Sub GenLineNumber()
        Dim i As Integer
        Dim n As Integer

        On Error Resume Next

        If Me.ListViewInvoice.Items.Count > 0 Then
            n = 1

            For i = 0 To Me.ListViewInvoice.Items.Count - 1
                Me.ListViewInvoice.Items(i).Text = n
                n = n + 1
            Next
        End If
    End Sub

    Private Sub ListViewInvoice_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewInvoice.SelectedIndexChanged

    End Sub

    Private Sub BTNClearScreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearScreen.Click
        Me.TBDocNo.Text = ""
        Me.TBArName.Text = ""
        Me.TBCoupon.Text = ""
        Me.TBCouponAmount.Text = ""
        Me.TBRemainAmount.Text = ""
        Me.TBSumItemAmount.Text = ""
        Me.ListViewInvoice.Items.Clear()
        Me.ListViewItem.Items.Clear()
    End Sub
End Class