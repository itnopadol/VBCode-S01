Public Class FormMainApplication

    Dim vQuery As String
    Dim vDuty As String

    Private Sub BTNBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNBack.Click
        FormLogIn.Show()
        FormLogIn.TBPassword.Text = ""
        FormLogIn.TBPassword.Focus()
        Me.Hide()
    End Sub

    Private Sub BTNReOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNReOrder.Click
        FormReOrder.Show()
        Call vReOrderClearScreen()
        Me.Hide()
    End Sub

    Public Sub vReOrderClearScreen()
        FormReOrder.TBBarCode.Text = ""
        FormReOrder.TBItemCode.Text = ""
        FormReOrder.TBItemName.Text = ""
        FormReOrder.TBQty.Text = ""
        FormReOrder.TBRemainQty.Text = ""
        FormReOrder.TBSuggest.Text = ""
        FormReOrder.TBOrderPoint.Text = ""
        FormReOrder.TBMin.Text = ""
        FormReOrder.TBMax.Text = ""
        FormReOrder.TBUnit.Text = ""
        FormReOrder.TBReOrder.Text = ""
        FormReOrder.TBPrice.Text = ""
        FormReOrder.TBItemStatus.Text = ""
        FormReOrder.TBPORemain.Text = ""
        FormReOrder.TBSale1M.Text = ""
        FormReOrder.TBFrequency.Text = ""
        FormReOrder.BTNRedDot.Visible = False
        FormReOrder.ListViewStock.Items.Clear()
        FormReOrder.ListViewStock.Visible = False
        FormReOrder.ListViewShelfID.Items.Clear()
        FormReOrder.ListViewShelfID.Visible = False
        FormReOrder.TBBarCode.Focus()
        FormReOrder.TBBarCode.SelectAll()
    End Sub

    Private Sub BTNReOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNReOrder.KeyDown ', BTNPickup.KeyDown, BTNCheckOut.KeyDown, BTNBack.KeyDown, BTNCountStockDaily.KeyDown, BTNPrintLabel.KeyDown
        If e.KeyCode = Keys.Enter Then
            FormReOrder.Show()
            Call vReOrderClearScreen()
            Me.Hide()
        End If
    End Sub

    Private Sub BTNPickup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPickup.Click
        Dim vLogInName As String

        If vMemProfit = "S02" Then

            vLogInName = Me.TBUserName.Text

            vQuery = "select dutycode from npmaster.dbo.TB_NP_DriveInAuthorizeUser where username = '" & vLogInName & "' and activestatus = 1"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vDuty = pds.Tables(0).Rows(0)("dutycode").ToString
            End If

            If UCase(vDuty) = "PICKUP" Or UCase(vDuty) = "ADMIN" Then
                FormPickUp.Show()
                Me.Hide()
            Else
                MsgBox("Your user can not pickup", MsgBoxStyle.Critical, "Send Error Message")
            End If
        End If
    End Sub

    Private Sub BTNCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCheckOut.Click
        Dim vLogInName As String

        If vMemProfit = "S02" Then

            vLogInName = Me.TBUserName.Text

            vQuery = "select dutycode from npmaster.dbo.TB_NP_DriveInAuthorizeUser where username = '" & vLogInName & "' and activestatus = 1"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vDuty = pds.Tables(0).Rows(0)("dutycode").ToString
            End If

            If UCase(vDuty) = "CHECKOUT" Or UCase(vDuty) = "ADMIN" Then
                FormCheckOut.Show()
                Me.Hide()
            Else
                MsgBox("Your user can not checkout", MsgBoxStyle.Critical, "Send Error Message")
            End If
        End If
    End Sub

    Private Sub BTNPrintLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPrintLabel.Click
        FormPrintLabel.Show()
        Me.Hide()
    End Sub

    Public Sub vPrintLabelClearScreen()
        On Error Resume Next

        FormPrintLabel.TBBarcode.Text = ""
        FormPrintLabel.TBItemCode.Text = ""
        FormPrintLabel.TBItemName.Text = ""
        FormPrintLabel.TBQty.Text = ""
        FormPrintLabel.TBUnitPrice.Text = ""
        FormPrintLabel.TBPrice.Text = ""
        FormPrintLabel.TBTypePrice.Text = ""
        FormPrintLabel.BTNRedDot.Visible = False
        FormPrintLabel.ListViewItem.Items.Clear()
        FormPrintLabel.CMBLabelType.Focus()
    End Sub

    Private Sub BTNCountStockDaily_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCountStockDaily.Click
        FormCountStockDaily.Show()
        FormCountStockDaily.PNSelectStkType.Visible = True
        FormCountStockDaily.PNSelectStkType.BringToFront()
        FormCountStockDaily.RDBDay.Focus()

        Me.Hide()
    End Sub

    Private Sub BTNCheckShelf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCheckShelf.Click
        FormTestScan.Show()
        Me.Hide()
    End Sub

    Private Sub BTNShelfAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNShelfAddItem.Click
        FormShelfAddItem.Show()
        Me.Hide()
    End Sub

    Private Sub BTNSlotTag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSlotTag.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FormPutAway.Show()
        Me.Hide()
    End Sub

    Private Sub BTNPickup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNPickup.KeyDown
        If e.KeyCode = Keys.Enter Then
            FormPickUp.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub BTNCheckOut_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCheckOut.KeyDown
        If e.KeyCode = Keys.Enter Then
            FormCheckOut.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub BTNPrintLabel_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNPrintLabel.KeyDown
        If e.KeyCode = Keys.Enter Then
            FormPrintLabel.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub BTNCountStockDaily_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCountStockDaily.KeyDown
        If e.KeyCode = Keys.Enter Then
            FormCountStockDaily.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub BTNBack_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNBack.KeyDown, BTNReOrder.KeyDown, BTNPickup.KeyDown, BTNCheckOut.KeyDown, BTNCountStockDaily.KeyDown, BTNPrintLabel.KeyDown

        If e.KeyCode = 49 Then
            FormReOrder.Show()
            Call vReOrderClearScreen()
            Me.Hide()
        End If

        If e.KeyCode = 50 Then
            FormPickUp.Show()
            Me.Hide()
        End If

        If e.KeyCode = 51 Then
            FormCheckOut.Show()
            Me.Hide()
        End If

        If e.KeyCode = 52 Then
            FormPrintLabel.Show()
            Me.Hide()
        End If

        If e.KeyCode = 53 Then
            FormCountStockDaily.Show()
            Me.Hide()
        End If

        If e.KeyCode = 54 Then
            FormShelfAddItem.Show()
            Me.Hide()
        End If

        If e.KeyCode = Keys.Escape Then
            FormLogIn.Show()
            FormLogIn.TBPassword.Text = ""
            FormLogIn.TBPassword.Focus()
            Me.Hide()
        End If
    End Sub

    Private Sub BTNPOS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPOS.Click

    End Sub

    Private Sub FormMainApplication_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BTNReOrder.Focus()
    End Sub

    Private Sub BTNRecMoney_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNRecMoney.Click

    End Sub

    Private Sub BTNReqPromo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNReqPromo.Click
        FormReqPromotion.Show()
        Me.Hide()
    End Sub
End Class