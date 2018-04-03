Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic

Public Class dlgSearchRwItem
    Dim vQry As String
    Dim vSearch As String
    Dim i As Integer
    Dim dt As New DataTable
    Dim iLVW As New ListViewItem
    Private Sub dlgSearchRwItem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
    End Sub
    Private Sub LoadRW()
        If Me.txtFrwitm.Text <> "" Then
            Me.LVrwItem.Items.Clear()
            vSearch = Me.txtFrwitm.Text
            vQry = "exec dbo.USP_VP_RewardSearch '" & vSearch & "'"
            da = New SqlDataAdapter(vQry, vConnection)
            ds = New DataSet
            da.Fill(ds, "rw")
            dt = ds.Tables("rw")
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    iLVW = Me.LVrwItem.Items.Add(dt.Rows(i).Item("itemcode")) '0
                    iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("itemname") '1
                    iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("Unitcode") '2
                    iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("stockqty") '3
                    iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("Point") '4
                    iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("PicturePath") '5
                    iLVW.SubItems.Add(0).Text = dt.Rows(i).Item("Amount") '6
                Next
            Else
                MsgBox("ไม่พบข้อมูลสินค้าที่ค้นหา", MsgBoxStyle.Information, "Information")
            End If
        End If
    End Sub

    Private Sub btnFrw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFrw.Click
        Call LoadRW()
    End Sub

    'Private Sub txtFrwitm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFrwitm.KeyDown
    '    If e.KeyCode = Keys.Enter Then
    '        Call LoadRW()
    '    End If
    'End Sub

    Private Sub btnRWFN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRWFN.Click
        Call SelectReward()

        '        Dim x As Integer
        '        Dim vPath As String
        '        On Error GoTo ErrDes
        '        For x = 0 To Me.LVrwItem.Items.Count - 1
        '            If Me.LVrwItem.Items(x).Selected = True Then
        '                frmPoint01.txtNewitmcode.Text = Me.LVrwItem.Items(x).SubItems(0).Text
        '                frmPoint01.txtItmName.Text = Me.LVrwItem.Items(x).SubItems(1).Text
        '                frmPoint01.txtItmUnitcode.Text = Me.LVrwItem.Items(x).SubItems(2).Text
        '                frmPoint01.txtItmQty.Text = Me.LVrwItem.Items(x).SubItems(3).Text
        '                frmPoint01.txtItmPoint.Text = Me.LVrwItem.Items(x).SubItems(4).Text
        '                frmPoint01.txtItmPictPath.Text = Me.LVrwItem.Items(x).SubItems(5).Text
        '                frmPoint01.txtItmAmount.Text = Me.LVrwItem.Items(x).SubItems(6).Text
        '                frmPoint01.pbxItem.Image = Image.FromFile(frmPoint01.txtItmPictPath.Text)
        '                frmPoint01.pbxItem.Show()
        '            End If
        '        Next
        'errDes:
        '        If Err.Description <> "" Then
        '            vPath = "M:\ของรางวัล\noimage.jpg"
        '            frmPoint01.pbxItem.Image = Image.FromFile(vPath)
        '        End If
        '        frmPoint01.PNsaveSpecialPoint.Visible = True
        '        PsaveItmStatus = 0
        '        frmPoint01.Show()
        '        Me.Close()
    End Sub
    Private Sub SelectReward()
        Dim x As Integer
        Dim vPath As String
        On Error GoTo errDes
        For x = 0 To Me.LVrwItem.Items.Count - 1
            If Me.LVrwItem.Items(x).Selected = True Then
                frmPoint01.txtNewitmcode.Text = Me.LVrwItem.Items(x).SubItems(0).Text
                frmPoint01.txtItmName.Text = Me.LVrwItem.Items(x).SubItems(1).Text
                frmPoint01.txtItmUnitcode.Text = Me.LVrwItem.Items(x).SubItems(2).Text
                frmPoint01.txtItmPoint.Text = Me.LVrwItem.Items(x).SubItems(4).Text
                frmPoint01.txtItmPictPath.Text = Me.LVrwItem.Items(x).SubItems(5).Text
                frmPoint01.txtItmAmount.Text = Me.LVrwItem.Items(x).SubItems(6).Text
                frmPoint01.pbxItem.Image = Image.FromFile(frmPoint01.txtItmPictPath.Text)
                frmPoint01.pbxItem.Show()
            End If
        Next
errDes:
        If Err.Description <> "" Then
            vPath = "M:\ของรางวัล\noimage.jpg"
            frmPoint01.pbxItem.Image = Image.FromFile(vPath)
        End If
        frmPoint01.PNsaveSpecialPoint.Visible = True
        PsaveItmStatus = 0
        Me.Close()
        frmPoint01.Show()

    End Sub
   

    Private Sub btnExitRw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExitRw.Click
        Me.Close()
    End Sub

    Private Sub txtFrwitm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFrwitm.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call LoadRW()
        End If
    End Sub

    Private Sub LVrwItem_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LVrwItem.MouseDoubleClick
        Call SelectReward()
    End Sub

   
End Class
