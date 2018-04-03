Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Public Class FormUpdateItemPriceDocNo
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand
    Dim vReadQuery As SqlDataReader
    Dim vCountItem As Integer

    Private Sub FormUpdateItemPriceDocNo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim vListItem As ListViewItem
        Dim i As Integer

        On Error GoTo ErrDescription

        Call InitializeDataBase()
        vQuery = "exec dbo.USP_NP_ItemChagePriceLevel"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "ItemLabel")
        dt = ds.Tables("ItemLabel")
        If dt.Rows.Count > 0 Then
            'vCountItem = dt.Rows.Count
            For i = 0 To dt.Rows.Count - 1
                vListItem = ListView101.Items.Add(dt.Rows(i).Item("docno"))
                vListItem.SubItems.Add(0).Text = Trim(dt.Rows(i).Item("itemcode"))
                vListItem.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("pricelevel"))
                vListItem.SubItems.Add(2).Text = Trim(dt.Rows(i).Item("type"))
                vListItem.SubItems.Add(3).Text = Format(Int(dt.Rows(i).Item("oldprice")), "##,##0.00")
                vListItem.SubItems.Add(4).Text = Format(Int(dt.Rows(i).Item("newprice")), "##,##0.00")
                vListItem.SubItems.Add(5).Text = Trim(dt.Rows(i).Item("unitcode"))
                vListItem.SubItems.Add(6).Text = Trim(dt.Rows(i).Item("itemname"))
            Next
        Else
            MsgBox("ไม่มีรายการสินค้าที่จะปรับราคาสินค้าในวันนี้ ", MsgBoxStyle.Critical, "Send Information")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub


    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Me.Close()
    End Sub

    Private Sub BTNUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNUpdate.Click
        Dim i As Integer
        Dim vDocNo As String

        vQuery = "exec dbo.USP_NP_SearchDocnoUpdatePrice"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        vCountItem = dt.Rows.Count - 1
        If dt.Rows.Count > 0 Then
            PGB101.Maximum = vCountItem
            For i = 0 To dt.Rows.Count - 1
                vDocNo = dt.Rows(i).Item("docno")
                vQuery = "exec dbo.USP_BS_PriceUpdateDaily '" & vDocNo & "'"
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()
                PGB101.Value = i
            Next
            MsgBox("ปรับข้อมูลราคาสินค้าเรียบร้อยแล้วครับ", MsgBoxStyle.Information, "Send Information")
            ListView101.Items.Clear()
        End If
    End Sub
End Class