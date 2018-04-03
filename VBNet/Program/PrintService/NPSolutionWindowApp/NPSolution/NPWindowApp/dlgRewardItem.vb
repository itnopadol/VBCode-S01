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
                    MsgBox("�����Ҫԡ�������� �����ԡ��¡�ù��������", MsgBoxStyle.Critical, "Warning")
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
        dt1.Columns.Add("�ӴѺ", GetType(Integer))
        dt1.Columns.Add("�����Թ���", GetType(String))
        dt1.Columns.Add("�����Թ���", GetType(String))
        dt1.Columns.Add("�ӹǹ", GetType(String))
        dt1.Columns.Add("˹��¹Ѻ", GetType(String))
        dt1.Columns.Add("��Ť��:˹���", GetType(String))
        dt1.Columns.Add("���:˹���", GetType(String))
        dt1.Columns.Add("��Ť��", GetType(String))
        dt1.Columns.Add("���", GetType(String))
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
                        dr("�ӴѺ") = inum
                        dr("�����Թ���") = dt.Rows(i).Item("itemcode")
                        dr("�����Թ���") = dt.Rows(i).Item("itemname")
                        dr("�ӹǹ") = 1
                        dr("˹��¹Ѻ") = dt.Rows(i).Item("unitcode")
                        dr("��Ť��:˹���") = dt.Rows(i).Item("amount")
                        dr("���:˹���") = dt.Rows(i).Item("point")
                        dr("��Ť��") = dt.Rows(i).Item("amount")
                        dr("���") = dt.Rows(i).Item("point")
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
