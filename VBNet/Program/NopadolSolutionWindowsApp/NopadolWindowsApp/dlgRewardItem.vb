Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic

Public Class dlgRewardItem
    Dim QryString As String
    Dim dt As New DataTable
    Dim i As Integer
    Dim iRewardItem As String
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Call GetData()
        Call calcTotalDW()
        Call dgvReadonly()
        'frmPoint01.Show()
        PsaveDWstatus = 0
        Me.Close()
    End Sub
    Private Sub CheckRwPoint()
        Dim x As Integer
        Dim NPoint As Double
        Dim vSumPoint As Double
        Dim vPointAmount As Double
        vPointAmount = PtotalPoint
        For x = 0 To Me.LVrwd.Items.Count - 1
            If Me.LVrwd.Items(x).Checked = True Then
                NPoint = Me.LVrwd.Items(x).SubItems(6).Text
                vSumPoint = vSumPoint + NPoint
                If vSumPoint > vPointAmount Then
                    MsgBox("แต้มสมาชิกที่เหลือ ไม่พอเบิกรายการนี้เพิ่มได้", MsgBoxStyle.Critical, "Warning")
                    Me.LVrwd.Items(x).Checked = False
                End If
            End If
        Next
    End Sub
    Private Sub GetData()
        Dim i As Integer
        Dim dt1 As New DataTable("vData")
        Dim dr As DataRow
        Dim inum As Integer
        Dim x As Integer
        dt1.Columns.Add("ลำดับ", GetType(Integer))
        dt1.Columns.Add("รหัสสินค้า", GetType(String))
        dt1.Columns.Add("ชื่อสินค้า", GetType(String))
        dt1.Columns.Add("จำนวน", GetType(String))
        dt1.Columns.Add("หน่วยนับ", GetType(String))
        dt1.Columns.Add("มูลค่า:หน่วย", GetType(String))
        dt1.Columns.Add("แต้ม:หน่วย", GetType(String))
        dt1.Columns.Add("มูลค่า", GetType(String))
        dt1.Columns.Add("แต้ม", GetType(String))
        For x = 0 To Me.LVrwd.Items.Count - 1
            If Me.LVrwd.Items(x).Checked = True Then
                iRewardItem = Me.LVrwd.Items(x).SubItems(1).Text
                QryString = "exec dbo.USP_VP_ItemReward '" & iRewardItem & "'"
                da = New SqlDataAdapter(QryString, vConnection)
                ds = New DataSet
                da.Fill(ds, "rwList")
                dt = ds.Tables("rwList")
                If dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        inum = inum + 1
                        dr = dt1.NewRow
                        dr("ลำดับ") = inum
                        dr("รหัสสินค้า") = dt.Rows(i).Item("itemcode")
                        dr("ชื่อสินค้า") = dt.Rows(i).Item("itemname")
                        dr("จำนวน") = 1
                        dr("หน่วยนับ") = dt.Rows(i).Item("unitcode")
                        dr("มูลค่า:หน่วย") = dt.Rows(i).Item("amount")
                        dr("แต้ม:หน่วย") = dt.Rows(i).Item("point")
                        dr("มูลค่า") = dt.Rows(i).Item("amount")
                        dr("แต้ม") = dt.Rows(i).Item("point")
                        dt1.Rows.Add(dr)
                    Next
                    frmPoint01.dgvDWreward.DataSource = dt1
                End If
            End If
        Next
    End Sub
    Private Sub calcTotalDW()
        Dim nx As Integer
        Dim TotalAmount As Integer
        Dim TotalPoint As Integer
        For nx = 0 To frmPoint01.dgvDWreward.Rows.Count - 1
            TotalAmount = (TotalAmount + frmPoint01.dgvDWreward.Item(7, frmPoint01.dgvDWreward.Rows(nx).Index).Value)
            TotalPoint = (TotalPoint + frmPoint01.dgvDWreward.Item(8, frmPoint01.dgvDWreward.Rows(nx).Index).Value)
        Next
        frmPoint01.txtTotaldwAmount.Text = Format(TotalAmount, "##,##0.00")
        frmPoint01.txtTotaldwPoint.Text = Format(TotalPoint, "##,##0.00")
    End Sub
    Private Sub dgvReadonly()
        frmPoint01.dgvDWreward.Columns(0).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(1).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(2).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(3).DefaultCellStyle.BackColor = Color.YellowGreen
        frmPoint01.dgvDWreward.Columns(4).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(5).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(6).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(7).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(8).ReadOnly = True

    End Sub
 

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.Close()
    End Sub

    Private Sub dlgRewardItem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
        Call PointListforChange()
        Me.txtmemberp.Text = Format(CDbl(PtotalPoint), "##,##0.00")
    End Sub
    Private Sub PointListforChange()
        Dim m As Integer
        Dim n As Integer
        Dim itmCode As String
        Dim lvItmcode As String

        Call InitializeDataBase()
        Dim xLV As ListViewItem
        Dim Pqry As String
        Dim xnum As Integer
        Dim PointAM As Integer
        PointAM = PtotalPoint
        Pqry = "exec dbo.USP_VP_CheckAVLReward '" & PointAM & "'"
        da = New SqlDataAdapter(Pqry, vConnection)
        ds = New DataSet
        da.Fill(ds, "pchg")
        dt = ds.Tables("pchg")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                xnum = xnum + 1
                xLV = Me.LVrwd.Items.Add(xnum)
                xLV.SubItems.Add(0).Text = dt.Rows(i).Item("itemcode")
                xLV.SubItems.Add(0).Text = dt.Rows(i).Item("itemname")
                xLV.SubItems.Add(0).Text = dt.Rows(i).Item("stockqty")
                xLV.SubItems.Add(0).Text = dt.Rows(i).Item("unitcode")
                xLV.SubItems.Add(0).Text = dt.Rows(i).Item("Amount")
                xLV.SubItems.Add(0).Text = dt.Rows(i).Item("Point")
            Next
        End If
        If frmPoint01.dgvDWreward.Rows.Count > 0 Then
            For m = 0 To frmPoint01.dgvDWreward.Rows.Count - 1
                itmCode = frmPoint01.dgvDWreward.Rows(m).Cells(1).Value
                For n = 0 To Me.LVrwd.Items.Count - 1
                    lvItmcode = Me.LVrwd.Items(n).SubItems(1).Text
                    If itmCode = lvItmcode Then
                        Me.LVrwd.Items(n).Checked = True
                    End If
                Next
            Next

        End If
    End Sub


    Private Sub LVrwd_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles LVrwd.ItemChecked
        Call CheckRwPoint()
    End Sub

   
End Class
