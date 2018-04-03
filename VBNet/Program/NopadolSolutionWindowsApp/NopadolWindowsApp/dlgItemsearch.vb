Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic

Public Class dlgItemsearch
    Dim itmQry As String
    Dim dt As New DataTable
    Dim i As Integer
    Dim iLVitm As ListViewItem
    Dim x As Integer
    Dim txtSCH As String

    Private Sub dlgItemsearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
    End Sub
    Private Sub ListReward()
        Dim avgCost As Double
        If Me.txtSCHitm.Text <> "" Then
            txtSCH = Me.txtSCHitm.Text
            itmQry = "exec dbo.USP_VP_Itemsearch '" & txtSCH & "'"
            da = New SqlDataAdapter(itmQry, vConnection)
            ds = New DataSet
            da.Fill(ds, "itm")
            dt = ds.Tables("itm")
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    iLVitm = Me.LVSCHitm.Items.Add(dt.Rows(i).Item("itemcode"))
                    iLVitm.SubItems.Add(0).Text = dt.Rows(i).Item("itemName")
                    iLVitm.SubItems.Add(0).Text = dt.Rows(i).Item("defsaleUnitcode")
                    If dt.Rows(i).Item("Averagecost") IsNot DBNull.Value Then
                        iLVitm.SubItems.Add(0).Text = Format(dt.Rows(i).Item("Averagecost"), "##,##0.00")
                    Else
                        iLVitm.SubItems.Add(0).Text = 0.0
                    End If
                Next
            End If
        Else
            MsgBox("คุณยังไม่ใส่รหัสสินค้าหรือชื่อสินค้า", MsgBoxStyle.Critical, "Error")
        End If
    End Sub

    Private Sub btnSCHitm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSCHitm.Click
        Call ListReward()
    End Sub

    Private Sub txtSCHitm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSCHitm.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call ListReward()
        End If
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Call SelectItemRw()
    End Sub
    Private Sub SelectItemRw()
        If Me.LVSCHitm.Items.Count > 0 Then
            For x = 0 To Me.LVSCHitm.Items.Count - 1
                If Me.LVSCHitm.Items(x).Selected = True Then
                    frmPoint01.txtNewitmcode.Text = Me.LVSCHitm.Items(x).SubItems(0).Text
                    frmPoint01.txtItmName.Text = Me.LVSCHitm.Items(x).SubItems(1).Text
                    frmPoint01.txtItmUnitcode.Text = Me.LVSCHitm.Items(x).SubItems(2).Text
                    frmPoint01.txtItmAmount.Text = Me.LVSCHitm.Items(x).SubItems(3).Text
                End If
            Next
        End If
        frmPoint01.txtNewitmcode.ReadOnly = True
        frmPoint01.txtItmName.ReadOnly = True
        ' frmPoint01.MdiParent = frmMainMember
        ' frmPoint01.Show()
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub LVSCHitm_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LVSCHitm.MouseDoubleClick
        SelectItemRw()
    End Sub

   
End Class
