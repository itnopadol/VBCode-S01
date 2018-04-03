Imports System.data
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms
Public Class FrmItemPrint

    Dim vQuery As String


    Private Sub TBBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
        Dim vBarcode As String
        Dim vItemCode As String
        Dim vPrice As Double
        Dim i As Integer
        Dim vExpireDate As Date


        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.TBBarCode.Text <> "" Then
                vBarcode = Me.TBBarCode.Text
            Else
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                Exit Sub
            End If

            vQuery = "exec dbo.USP_NP_SearchBarcodeDetails '" & vBarcode & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            If ds.Tables(0).Rows.Count > 0 Then
                Me.TBMemBarCode.Text = ds.Tables(0).Rows(0)("barcode").ToString
                Me.TBItemCode.Text = ds.Tables(0).Rows(0)("itemcode").ToString
                Me.TBItemName.Text = ds.Tables(0).Rows(0)("itemname").ToString
                Me.TBUnitCode.Text = ds.Tables(0).Rows(0)("unitcode").ToString

                vPrice = ds.Tables(0).Rows(0)("price").ToString
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBPriceType.Text = ds.Tables(0).Rows(0)("pricedesc").ToString

                If ds.Tables(0).Rows(0)("pricedesc").ToString = "ราคาPromotion" Then
                    vExpireDate = ds.Tables(0).Rows(0)("dateend").ToString
                    Me.TBPromoExpire.Text = vExpireDate
                Else
                    Me.TBPromoExpire.Text = ""
                End If

                Me.PNShowDetails.Visible = True
                Me.PNShowDetails.BringToFront()
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
            Else
                MsgBox("ไม่มีข้อมูลของบาร์โค้ดที่ค้นหา กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Text = ""
                Exit Sub
            End If

        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call MainMenu()
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If


ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

    End Sub

    Public Sub ClearScreen()
        Me.TBMemBarCode.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBUnitCode.Text = ""
        Me.TBPrice.Text = ""
        Me.TBPriceType.Text = ""
        Me.TBPromoExpire.Text = ""
        Me.TBBarCode.Text = ""
        Me.PNShowDetails.Visible = False
        Me.PNKeyQty.Visible = False
        Me.TBQty.Text = ""
        Me.TBBarCode.Enabled = True
        Me.TBBarCode.Focus()
        Me.TBBarCode.SelectAll()
    End Sub

    Private Sub TBBarCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged

    End Sub

    Public Sub AddItem()
        Me.TBBarCode.Enabled = False
        Me.PNKeyQty.Visible = True
        Me.PNKeyQty.BringToFront()
        Me.TBQty.Focus()
        Me.TBQty.SelectAll()
    End Sub

    Private Sub BTNAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddItem.Click
        Me.TBBarCode.Enabled = False
        Me.PNKeyQty.Visible = True
        Me.PNKeyQty.BringToFront()
        Me.TBQty.Focus()
        Me.TBQty.SelectAll()
    End Sub

    Private Sub TBItemCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBItemCode.KeyDown, TBItemName.KeyDown, TBPrice.KeyDown, TBPriceType.KeyDown, TBPromoExpire.KeyDown, TBUnitCode.KeyDown
        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If
    End Sub

    Private Sub TBItemCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBItemCode.TextChanged

    End Sub

    Private Sub TBQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBQty.KeyDown
        Dim i As Integer
        Dim n As Integer
        Dim vQty As Double

        Dim vItemcode As String
        Dim vBarCode As String
        Dim vUnitCode As String

        Dim vCheckItemCode As String
        Dim vCheckBarCode As String
        Dim vCheckUnitCode As String
        Dim vCheckQty As Double
        Dim vAnswer As Integer

        If e.KeyCode = Keys.Enter Then

            If Me.TBQty.Text = "" Then
                MsgBox("กรณีที่ต้องการเพิ่มรายการสินค้าที่จะพิมพ์ป้าย ต้องกรอกจำนวนที่ต้องการพิมพ์ อย่างน้อย 1 ดวงขึ้นไป", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBQty.Focus()
                Me.TBQty.SelectAll()
                Exit Sub
            End If

            vItemcode = Me.TBItemCode.Text
            vBarCode = Me.TBMemBarCode.Text
            vUnitCode = Me.TBUnitCode.Text

            vQty = Me.TBQty.Text
            If Me.TBItemCode.Text <> "" Then

                If Me.ListViewItem.Items.Count > 0 Then
                    For n = 0 To Me.ListViewItem.Items.Count - 1
                        vCheckItemCode = Me.ListViewItem.Items(n).SubItems(1).Text
                        vCheckBarCode = Me.ListViewItem.Items(n).SubItems(4).Text
                        vCheckUnitCode = Me.ListViewItem.Items(n).SubItems(5).Text
                        vCheckQty = Me.ListViewItem.Items(n).SubItems(3).Text

                        If vItemcode = vCheckItemCode And vBarCode = vCheckBarCode And vUnitCode = vCheckUnitCode Then
                            vAnswer = MsgBox("มีรายการสินค้า อยู่แล้วในบรรทัด ที่ " & n + 1 & " จำนวนที่สั่งพิมพ์คือ " & vCheckQty & " ต้องการเปลี่ยนหรือไม่ ?", MsgBoxStyle.YesNo, "Send Question Message")
                            If vAnswer = 6 Then
                                Me.ListViewItem.Items(n).SubItems(3).Text = Format(vQty, "##,##0.00")
                                Call ClearScreen()
                                Exit Sub
                            Else
                                Call ClearScreen()
                                Exit Sub
                            End If
                        End If
                    Next
                End If

                i = Me.ListViewItem.Items.Count
                i = i + 1

                Dim listItem As New ListViewItem(i)
                listItem.SubItems.Add(Me.TBItemCode.Text)
                listItem.SubItems.Add(Me.TBItemName.Text)
                listItem.SubItems.Add(Format(vQty, "##,##0.00"))
                listItem.SubItems.Add(Me.TBMemBarCode.Text)
                listItem.SubItems.Add(Me.TBUnitCode.Text)
                Me.ListViewItem.Items.Add(listItem)

                Me.PNKeyQty.Visible = False
                Call ClearScreen()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If
    End Sub

    Private Sub TBQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBQty.KeyPress
        On Error GoTo ErrDescription

        Select Case Asc(e.KeyChar)
            Case 48 To 58, 8, 44, 45, 46, 37
            Case Else
                e.Handled = True
        End Select

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TBQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBQty.TextChanged

    End Sub

    Private Sub BTNSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim i As Integer
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vQty As Double
        Dim vUnitCode As String

        MsgBox("ครั้งต่อไป ให้กดปุ่ม สีส้ม + ปุ่มหมายเลข 8", MsgBoxStyle.Critical, "Send Error Message")
        If Me.ListViewItem.Items.Count > 0 Then
            For i = 0 To Me.ListViewItem.Items.Count - 1
                vItemCode = Me.ListViewItem.Items(i).SubItems(1).Text
                vBarCode = Me.ListViewItem.Items(i).SubItems(4).Text
                vQty = Me.ListViewItem.Items(i).SubItems(3).Text
                vUnitCode = Me.ListViewItem.Items(i).SubItems(5).Text

                vQuery = "exec dbo.USP_NP_InsertselectItemMobileForJob '" & vItemCode & "','" & vBarCode & "'," & vQty & ",'" & vUnitCode & "','" & vPersonName & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As Integer = vService.vExecuteQuery(vQuery)
            Next

            MsgBox("บันทึกรายการสินค้าเรียบร้อย", MsgBoxStyle.Information, "Send Information Message")
            Me.ListViewItem.Items.Clear()
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If
    End Sub

    Public Sub SaveData()
        Dim i As Integer
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vQty As Double
        Dim vUnitCode As String

        If Me.ListViewItem.Items.Count > 0 Then
            For i = 0 To Me.ListViewItem.Items.Count - 1
                vItemCode = Me.ListViewItem.Items(i).SubItems(1).Text
                vBarCode = Me.ListViewItem.Items(i).SubItems(4).Text
                vQty = Me.ListViewItem.Items(i).SubItems(3).Text
                vUnitCode = Me.ListViewItem.Items(i).SubItems(5).Text

                vQuery = "exec dbo.USP_NP_InsertselectItemMobileForJob '" & vItemCode & "','" & vBarCode & "'," & vQty & ",'" & vUnitCode & "','" & vPersonName & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As Integer = vService.vExecuteQuery(vQuery)
            Next

            MsgBox("บันทึกรายการสินค้าเรียบร้อย", MsgBoxStyle.Information, "Send Information Message")
            Me.ListViewItem.Items.Clear()
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If
    End Sub

    Private Sub ListViewItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItem.KeyDown
        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If


        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call MainMenu()
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If
    End Sub

    Private Sub BTNAddItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNAddItem.KeyDown
        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call MainMenu()
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If
    End Sub

    Private Sub BTNSave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSave.KeyDown
        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call MainMenu()
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If
    End Sub

    Public Sub MainMenu()
        Call ClearScreen()
        FrmMobileApp.Show()
        Me.Hide()
    End Sub

    Private Sub BTNMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNMenu.Click
        MsgBox("ครั้งต่อไปให้กดปุ่ม สีส้ม + ปุ่มเลข 7", MsgBoxStyle.Information, "Send Information Message")
        Call ClearScreen()
        FrmMobileApp.Show()
        Me.Hide()
    End Sub

    Private Sub BTNAddQty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddQty.Click
        MsgBox("ครั้งต่อไปให้กดปุ่ม สีส้ม + ปุ่มเลข 1", MsgBoxStyle.Information, "Send Information Message")
        Me.TBBarCode.Enabled = False
        Me.PNKeyQty.Visible = True
        Me.PNKeyQty.BringToFront()
        Me.TBQty.Focus()
        Me.TBQty.SelectAll()
    End Sub

    Private Sub BTNAddQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNAddQty.KeyDown
        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call MainMenu()
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If

    End Sub
End Class