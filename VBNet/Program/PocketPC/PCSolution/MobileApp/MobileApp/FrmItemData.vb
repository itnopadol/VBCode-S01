Imports System.data
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms

Public Class FrmItemData
    Dim vQuery As String

    Private Sub BTNMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNMain.Click
        Call ClearScreen()
        FrmMobileApp.Show()
        Me.Hide()
    End Sub

    Private Sub TBBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
        Dim vBarcode As String
        Dim vItemCode As String
        Dim vStkQty As Double
        Dim vResQty As Double
        Dim vPrice As Double
        Dim i As Integer
        Dim vQty As Double

        Dim vSaleUnit As String
        Dim vSaleType As String
        Dim vPrice1 As Double
        Dim vPrice2 As Double
        Dim vPrice3 As Double
        Dim vFromQty As Double
        Dim vToQty As Double
        Dim vFromDate As Date
        Dim vToDate As Date

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

            Me.ListViewItemPrice.Items.Clear()

            If ds.Tables(0).Rows.Count > 0 Then
                Me.TBItemCode.Text = ds.Tables(0).Rows(0)("itemcode").ToString
                Me.TBItemName.Text = ds.Tables(0).Rows(0)("itemname").ToString
                Me.TBItemUnit.Text = ds.Tables(0).Rows(0)("unitcode").ToString

                vPrice = ds.Tables(0).Rows(0)("price").ToString
                vStkQty = ds.Tables(0).Rows(0)("realqty").ToString
                vResQty = ds.Tables(0).Rows(0)("resqty").ToString

                Me.TBItemQty.Text = Format(vStkQty, "##,##0.00")
                Me.TBItemResQty.Text = Format(vResQty, "##,##0.00")
                Me.TBItemPrice.Text = Format(vPrice, "##,##0.00")
                Me.LBLQtyUnit1.Text = ds.Tables(0).Rows(0)("unitcode").ToString
                Me.LBLQtyUnit2.Text = ds.Tables(0).Rows(0)("unitcode").ToString

                For i = 0 To ds.Tables(0).Rows.Count - 1

                    vSaleUnit = ds.Tables(0).Rows(i)("unitsale").ToString
                    vSaleType = ds.Tables(0).Rows(i)("saletypedesc").ToString
                    vPrice1 = ds.Tables(0).Rows(i)("saleprice1").ToString
                    vPrice2 = ds.Tables(0).Rows(i)("saleprice2").ToString
                    vPrice3 = ds.Tables(0).Rows(i)("saleprice3").ToString
                    vFromQty = ds.Tables(0).Rows(i)("fromqty").ToString
                    vToQty = ds.Tables(0).Rows(i)("toqty").ToString
                    vFromDate = ds.Tables(0).Rows(i)("startdate").ToString
                    vToDate = ds.Tables(0).Rows(i)("stopdate").ToString

                    Dim listItem As New ListViewItem(vSaleUnit)
                    listItem.SubItems.Add(vSaleType)
                    listItem.SubItems.Add(vPrice1)
                    listItem.SubItems.Add(vPrice2)
                    listItem.SubItems.Add(Format(vPrice3, "##,##0.00"))
                    listItem.SubItems.Add(vFromDate)
                    listItem.SubItems.Add(vToDate)
                    listItem.SubItems.Add(vFromQty)
                    listItem.SubItems.Add(vToQty)
                    Me.ListViewItemPrice.Items.Add(listItem)
                Next i

                Me.PNItem1.Visible = True
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = False

                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
            Else
                MsgBox("ไม่มีข้อมูลของบาร์โค้ดที่ค้นหา กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Text = ""
                Call ClearScreen()
                Exit Sub
            End If

            vItemCode = Me.TBItemCode.Text
            vQuery = "exec dbo.USP_NP_ItemStockBranch '" & vItemCode & "'"
            Dim vService1 As New WebReference.WebServiceCalc
            Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)

            Me.ListViewItemStock.Items.Clear()

            If ds1.Tables(0).Rows.Count > 0 Then
                vQty = ds1.Tables(0).Rows(0)("qty").ToString
                Dim listItem As New ListViewItem(ds1.Tables(0).Rows(0)("whcode").ToString)
                listItem.SubItems.Add(ds1.Tables(0).Rows(0)("shelfcode").ToString)
                listItem.SubItems.Add(vQty)
                listItem.SubItems.Add(ds1.Tables(0).Rows(0)("unitcode").ToString)
                Me.ListViewItemStock.Items.Add(listItem)
            End If

        End If

            If e.KeyCode = 114 Then
                Me.PNItem1.Visible = True
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = False
            End If

            If e.KeyCode = 115 Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
            End If

            If e.KeyCode = 33 Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = True
            End If

            If e.KeyCode = 34 Then
                Call ClearScreen()
                FrmMobileApp.Show()
                Me.Hide()
            End If

            If e.KeyCode = Keys.Escape Then
                Call ClearScreen()
            End If

            If e.KeyCode = 37 Then
                If Me.PNItem2.Visible = True Then
                    Me.PNItem1.Visible = True
                    Me.PNItem2.Visible = False
                    Me.PNItem3.Visible = False
                    Exit Sub
                End If

                If Me.PNItem3.Visible = True Then
                    Me.PNItem1.Visible = False
                    Me.PNItem2.Visible = True
                    Me.PNItem3.Visible = False
                    Exit Sub
                End If
            End If

            If e.KeyCode = 39 Then
                If Me.PNItem1.Visible = True Then
                    Me.PNItem1.Visible = False
                    Me.PNItem2.Visible = True
                    Me.PNItem3.Visible = False
                    Exit Sub
                End If

                If Me.PNItem2.Visible = True Then
                    Me.PNItem1.Visible = False
                    Me.PNItem2.Visible = False
                    Me.PNItem3.Visible = True
                    Exit Sub
                End If
            End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

    End Sub

    Public Sub ClearScreen()
        On Error Resume Next

        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBItemPrice.Text = ""
        Me.TBItemQty.Text = ""
        Me.TBItemResQty.Text = ""
        Me.TBItemUnit.Text = ""
        Me.LBLQtyUnit1.Text = ""
        Me.LBLQtyUnit2.Text = ""
        Me.PNItem1.Visible = True
        Me.PNItem2.Visible = False
        Me.PNItem3.Visible = False
        Me.ListViewItemPrice.Items.Clear()
        Me.ListViewItemStock.Items.Clear()
        Me.TBBarCode.Focus()
        Me.TBBarCode.SelectAll()
    End Sub

    Private Sub LinkLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LLItem2.Click
        On Error Resume Next

        Me.PNItem2.Visible = True
        Me.PNItem1.Visible = False
        Me.PNItem3.Visible = False
    End Sub

    Private Sub FrmItemData_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        Me.TBBarCode.Text = ""
        Me.TBBarCode.Focus()
        Me.TBBarCode.SelectAll()
    End Sub

    Private Sub LLItem2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles LLItem2.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 114 Then
            Me.PNItem1.Visible = True
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 115 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = True
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 33 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = True
        End If

        If e.KeyCode = 34 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 37 Then
            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = True
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem3.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If
        End If

        If e.KeyCode = 39 Then
            If Me.PNItem1.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = True
                Exit Sub
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub LLItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LLItem3.Click
        On Error Resume Next

        Me.PNItem2.Visible = False
        Me.PNItem1.Visible = False
        Me.PNItem3.Visible = True
    End Sub

    Private Sub LLItem3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles LLItem3.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 114 Then
            Me.PNItem1.Visible = True
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 115 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = True
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 33 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = True
        End If

        If e.KeyCode = 34 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 37 Then
            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = True
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem3.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If
        End If

        If e.KeyCode = 39 Then
            If Me.PNItem1.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = True
                Exit Sub
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub LLItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LLItem1.Click
        On Error Resume Next

        Me.PNItem2.Visible = False
        Me.PNItem1.Visible = True
        Me.PNItem3.Visible = False
    End Sub

    Private Sub LLItem1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles LLItem1.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 114 Then
            Me.PNItem1.Visible = True
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 115 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = True
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 33 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = True
        End If

        If e.KeyCode = 34 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 37 Then
            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = True
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem3.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If
        End If

        If e.KeyCode = 39 Then
            If Me.PNItem1.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = True
                Exit Sub
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNMain.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 114 Then
            Me.PNItem1.Visible = True
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 115 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = True
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 33 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = True
        End If

        If e.KeyCode = 34 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 37 Then
            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = True
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem3.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If
        End If

        If e.KeyCode = 39 Then
            If Me.PNItem1.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = True
                Exit Sub
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub ListViewItemStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItemStock.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 114 Then
            Me.PNItem1.Visible = True
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 115 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = True
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 33 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = True
        End If

        If e.KeyCode = 34 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 37 Then
            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = True
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem3.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If
        End If

        If e.KeyCode = 39 Then
            If Me.PNItem1.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = True
                Exit Sub
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub ListViewItemPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItemPrice.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 114 Then
            Me.PNItem1.Visible = True
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 115 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = True
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 33 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = True
        End If

        If e.KeyCode = 34 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 37 Then
            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = True
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem3.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If
        End If

        If e.KeyCode = 39 Then
            If Me.PNItem1.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = True
                Exit Sub
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBItemCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBItemCode.KeyDown, TBItemName.KeyDown, TBItemPrice.KeyDown, TBItemQty.KeyDown, TBItemResQty.KeyDown, TBItemUnit.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 114 Then
            Me.PNItem1.Visible = True
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 115 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = True
            Me.PNItem3.Visible = False
        End If

        If e.KeyCode = 33 Then
            Me.PNItem1.Visible = False
            Me.PNItem2.Visible = False
            Me.PNItem3.Visible = True
        End If

        If e.KeyCode = 34 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 37 Then
            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = True
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem3.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If
        End If

        If e.KeyCode = 39 Then
            If Me.PNItem1.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = True
                Me.PNItem3.Visible = False
                Exit Sub
            End If

            If Me.PNItem2.Visible = True Then
                Me.PNItem1.Visible = False
                Me.PNItem2.Visible = False
                Me.PNItem3.Visible = True
                Exit Sub
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBBarCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged

    End Sub
End Class