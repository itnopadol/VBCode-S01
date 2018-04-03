Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic
Public Class FormApproveVolumeSet
    Dim QryString As String
    Dim i As Integer
    Dim da As SqlDataAdapter
    Dim ds As DataSet
    Dim dt As DataTable
    Dim cmd As New SqlCommand

    Private Sub FormApproveVolumeSet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
        Call getConfirmPriceVolume()
        Me.Text = "อนุมัติกำหนดราคาตามจำนวน"
    End Sub
    Private Sub getConfirmPriceVolume()
        Dim iCFLV As New ListViewItem
        Dim vMKcostGP As Integer
        Dim vAVcostGP As Integer
        Me.ListQue.Items.Clear()
        QryString = "exec dbo.USP_PS_PriceVolumeSetNotconfirm"
        da = New SqlDataAdapter(QryString, vConnectionString)
        ds = New DataSet
        da.Fill(ds, "cfVM")
        dt = ds.Tables("cfVM")
        If dt.Rows.Count <> 0 Then
            For i = 0 To dt.Rows.Count - 1
                iCFLV = Me.ListQue.Items.Add(dt.Rows(i).Item("DocNo"))
                iCFLV.SubItems.Add(0).Text = dt.Rows(i).Item("DocDate")
                vMKcostGP = dt.Rows(i).Item("marketcostGP")
                vAVcostGP = dt.Rows(i).Item("LotAverageCostGP")
                iCFLV.SubItems.Add(0).Text = Format(Int(vMKcostGP), "##,##0.00")
                iCFLV.SubItems.Add(0).Text = Format(Int(vAVcostGP), "##,##0.00")
                iCFLV.SubItems.Add(0).Text = dt.Rows(i).Item("creator")
            Next
        End If
    End Sub

    Private Sub chkAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAll.CheckedChanged
        Dim xi As Integer
        If Me.ListQue.Items.Count > 0 Then
            If Me.chkAll.Checked = True Then
                For xi = 0 To Me.ListQue.Items.Count - 1
                    Me.ListQue.Items(xi).Checked = True
                Next
            Else
                For xi = 0 To Me.ListQue.Items.Count - 1
                    Me.ListQue.Items(xi).Checked = False
                Next
            End If

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Call getConfirmPriceVolume()
    End Sub

    Private Sub btnApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApprove.Click
        Dim idocno As String

        QryString = "begin tran"
        With cmd
            .CommandType = CommandType.Text
            .CommandText = QryString
            .Connection = vConnection
            .ExecuteNonQuery()
        End With
        On Error GoTo ErrDesc
        '----------------
        If Me.ListQue.Items.Count > 0 Then
            For i = 0 To Me.ListQue.Items.Count - 1
                If Me.ListQue.Items(i).Checked = True Then
                    idocno = Me.ListQue.Items(i).SubItems(0).Text
                    QryString = "exec dbo.usp_ps_pricevolumesetconfirm2'" & idocno & "'"
                    With cmd
                        .CommandType = CommandType.Text
                        .CommandText = QryString
                        .Connection = vConnection
                        .ExecuteNonQuery()
                    End With
                End If
            Next
            QryString = "commit tran"
            With cmd
                .CommandType = CommandType.Text
                .CommandText = QryString
                .Connection = vConnection
                .ExecuteNonQuery()
            End With
            MsgBox("อนุมัติเอกสารเรียบร้อยแล้ว..", MsgBoxStyle.Information, "Information")
            Call getConfirmPriceVolume()
        End If

ErrDesc:
        If Err.Description <> "" Then
            QryString = "roll back"
            With cmd
                .CommandType = CommandType.Text
                .CommandText = QryString
                .Connection = vConnection
                .ExecuteNonQuery()
            End With
            MsgBox("ไม่สามารถอนุมัติเอกสารได้.", MsgBoxStyle.Critical, "Error")
        End If

    End Sub
End Class