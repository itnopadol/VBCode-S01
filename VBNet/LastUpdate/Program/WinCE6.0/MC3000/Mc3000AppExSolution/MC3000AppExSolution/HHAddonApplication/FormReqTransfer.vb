Public Class FormReqTransfer

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Call ClearScreen()
        FormMainApplication.Show()
        Me.Hide()
    End Sub

    Public Sub ClearScreen()
        Me.TBDocNo.Text = ""
        Me.DTPDocDate.Value = Now
        Me.TBBarCode.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.TBBarCode.Focus()
    End Sub

    Public Sub ClearItem()
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBUnitCode.Text = ""
        Me.TBQty.Text = ""
        Me.ListViewStock.Items.Clear()
        Me.TBBarCode.Focus()
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBItemName.TextChanged

    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBUnitCode.TextChanged

    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        Me.PNItem.Visible = False
        Me.TBBarCode.Focus()
    End Sub

    Private Sub BTNAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAdd.Click

    End Sub

    Private Sub TBQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBQty.KeyDown
        Dim i As Integer
        Dim n As Integer
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vFromWH As String
        Dim vFromSH As String
        Dim vToWH As String
        Dim vToSH As String

        Dim vQty As Double
        Dim vRemainQty As Double

        Dim vChkItemCode As String
        Dim vChkUnitCode As String
        Dim vChkFromWH As String
        Dim vChkFromSH As String
        Dim vChkToWH As String
        Dim vChkToSH As String

        Dim vAnswer As Integer
        Dim vAnswer1 As Integer

        Dim vOldQty As Double
        Dim vAddQty As Double


        If e.KeyCode = Keys.Enter And Me.TBQty.Text <> "" And Me.TBItemCode.Text <> "" Then

            vItemCode = Me.TBItemCode.Text
            vBarCode = Me.TBBarCode.Text
            vItemName = Me.TBItemName.Text
            vUnitCode = Me.TBUnitCode.Text
            vFromWH = Me.CMBFromWH.Text
            vFromSH = Me.CMBFromSH.Text
            vToWH = Me.CMBToWH.Text
            vToSH = Me.CMBToSH.Text

            If Me.TBQty.Text <> "" Then
                vQty = Me.TBQty.Text
            Else
                vQty = 0
            End If

            If Me.ListViewItem.Items.Count > 0 Then

                For i = 0 To Me.ListViewItem.Items.Count - 1
                    vChkItemCode = Me.ListViewItem.Items(i).SubItems(1).Text
                    vChkUnitCode = Me.ListViewItem.Items(i).SubItems(4).Text
                    vChkFromWH = Me.ListViewItem.Items(i).SubItems(6).Text
                    vChkFromSH = Me.ListViewItem.Items(i).SubItems(7).Text
                    vChkToWH = Me.ListViewItem.Items(i).SubItems(6).Text
                    vChkToSH = Me.ListViewItem.Items(i).SubItems(7).Text
                    vOldQty = Me.ListViewItem.Items(i).SubItems(3).Text



                    If vItemCode = vChkItemCode And vUnitCode = vChkUnitCode And vFromWH = vChkFromWH And vFromSH = vChkFromSH And vToWH = vChkToWH And vToSH = vChkToSH Then

                        Exit Sub
                    End If
                Next

            End If

            n = Me.ListViewItem.Items.Count + 1
            Dim listItem As New ListViewItem(n)
            listItem.SubItems.Add(vItemCode)
            listItem.SubItems.Add(Format(vRemainQty, "##,##0.00"))
            listItem.SubItems.Add(Format(vQty, "##,##0.00"))
            listItem.SubItems.Add(vUnitCode)
            listItem.SubItems.Add(vBarCode)
            listItem.SubItems.Add(vFromWH)
            listItem.SubItems.Add(vFromSH)
            listItem.SubItems.Add(vToWH)
            listItem.SubItems.Add(vToSH)
            listItem.SubItems.Add(Now)
            listItem.SubItems.Add(0)
            listItem.SubItems.Add(vItemName)
            Me.ListViewItem.Items.Add(listItem)

            Me.PNItem.Visible = False
            Call ClearItem()

        End If

        If e.KeyCode = Keys.Up Then
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Down And Me.ListViewItem.Items.Count > 0 Then
            Me.ListViewItem.Items(0).Focused = True
            Me.ListViewItem.Items(0).Selected = True
            Me.ListViewItem.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearItem()
            Me.PNItem.Visible = False
            Me.TBBarCode.Focus()
        End If
    End Sub

    Private Sub TBQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBQty.TextChanged

    End Sub
End Class