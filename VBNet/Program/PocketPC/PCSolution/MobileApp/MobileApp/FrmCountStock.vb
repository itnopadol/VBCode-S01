Imports System.data
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms

Public Class FrmCountStock

    Dim vQuery As String

    Dim vGetDocNo As String
    Dim vNewDocNo As String
    Dim vDocno As String
    Dim vStkNo As String

    Private Sub TBBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
        Dim vBarcode As String
        Dim vItemCode As String
        Dim i As Integer
        Dim n As Integer
        Dim vQty As Double


        If e.KeyCode = Keys.Enter Then
            If Me.TBBarCode.Text <> "" Then
                vBarcode = Me.TBBarCode.Text
                vQuery = "exec dbo.USP_MB_SearchCheckBarcode '" & vBarcode & "' "
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count <= 0 Then
                    MsgBox("ไม่มีข้อมูลบาร์โค้ดที่ต้องการนับสต็อก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBBarCode.Text = ""
                    Me.TBBarCode.SelectAll()
                    Exit Sub
                End If

                vQuery = "exec dbo.USP_IV_CheckItemDescription '" & vBarcode & "'"
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)

                If ds1.Tables(0).Rows.Count > 0 Then
                    vItemCode = ds1.Tables(0).Rows(n)("code").ToString

                    Dim x As Integer
                    Dim y As Integer
                    Dim vCheckItemCode As String
                    Dim vCountItemExist As Integer

                    If Me.ListViewItem.Items.Count > 0 Then
                        For x = 0 To Me.ListViewItem.Items.Count - 1
                            vCheckItemCode = Me.ListViewItem.Items(x).SubItems(4).Text

                            If vItemCode = vCheckItemCode Then

                                Me.TBItemCode.Text = Me.ListViewItem.Items(x).SubItems(4).Text
                                Me.TBItemName.Text = Me.ListViewItem.Items(x).SubItems(1).Text
                                Me.TBItemUnit.Text = Me.ListViewItem.Items(x).SubItems(3).Text

                                y = y + 1
                                Dim listItem As New ListViewItem(y)
                                listItem.SubItems.Add(Me.ListViewItem.Items(x).SubItems(6).Text)
                                If Me.ListViewItem.Items(x).SubItems(2).Text <> "" Then
                                    vQty = Me.ListViewItem.Items(x).SubItems(2).Text
                                Else
                                    vQty = 0
                                End If
                                listItem.SubItems.Add(Format(vQty, "##,##0.00"))
                                listItem.SubItems.Add(Me.ListViewItem.Items(x).SubItems(3).Text)
                                listItem.SubItems.Add(Me.ListViewItem.Items(x).SubItems(9).Text)
                                listItem.SubItems.Add(Me.ListViewItem.Items(x).SubItems(5).Text)
                                listItem.SubItems.Add(Me.ListViewItem.Items(x).SubItems(7).Text)
                                listItem.SubItems.Add(Me.ListViewItem.Items(x).SubItems(8).Text)
                                listItem.SubItems.Add(1)
                                listItem.SubItems.Add(x)
                                Me.ListViewStock.Items.Add(listItem)

                                vCountItemExist = vCountItemExist + 1
                            End If
                        Next

                        If vCountItemExist > 0 Then
                            Me.PNKeyQty.Visible = True
                            Me.PNKeyQty.BringToFront()
                            Me.TBBarCode.Text = ""
                            Me.ListViewStock.Focus()
                            Me.ListViewStock.Items(0).Selected = True
                            Me.ListViewStock.Items(0).Focused = True
                            Exit Sub
                        End If
                    End If


                    Me.ListViewStock.Items.Clear()
                    vQuery = "exec dbo.USP_MB_ItemStockLocationAll '" & vItemCode & "'"
                    Dim vService2 As New WebReference.WebServiceCalc
                    Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)

                    If ds2.Tables(0).Rows.Count > 0 Then
                        Me.TBItemCode.Text = ds2.Tables(0).Rows(n)("itemcode").ToString
                        Me.TBItemName.Text = ds2.Tables(0).Rows(n)("itemname").ToString
                        Me.TBItemUnit.Text = ds2.Tables(0).Rows(n)("unitcode").ToString

                        For n = 0 To ds2.Tables(0).Rows.Count - 1
                            i = i + 1
                            Dim listItem As New ListViewItem(i)
                            listItem.SubItems.Add(ds2.Tables(0).Rows(n)("shelfcode").ToString)
                            vQty = ds2.Tables(0).Rows(n)("qty").ToString
                            listItem.SubItems.Add("")
                            listItem.SubItems.Add(ds2.Tables(0).Rows(n)("unitcode").ToString)
                            listItem.SubItems.Add(Format(vQty, "##,##0.00"))
                            listItem.SubItems.Add(ds2.Tables(0).Rows(n)("whcode").ToString)
                            listItem.SubItems.Add(ds2.Tables(0).Rows(n)("shelfid").ToString)
                            listItem.SubItems.Add(1)
                            listItem.SubItems.Add(0)
                            listItem.SubItems.Add(0)
                            Me.ListViewStock.Items.Add(listItem)
                        Next

                        Me.PNKeyQty.Visible = True
                        Me.PNKeyQty.BringToFront()
                        Me.TBBarCode.Text = ""
                        Me.ListViewStock.Focus()
                        Me.ListViewStock.Items(0).Selected = True
                        Me.ListViewStock.Items(0).Focused = True
                    Else
                        Me.PNKeyQty.Visible = True
                        Me.PNKeyQty.BringToFront()
                        Me.TBBarCode.Text = ""
                        Me.TBItemCode.Text = ds1.Tables(0).Rows(n)("code").ToString
                        Me.TBItemName.Text = ds1.Tables(0).Rows(n)("name1").ToString
                        Me.TBItemUnit.Text = ds1.Tables(0).Rows(n)("defstkunitcode").ToString
                        Me.BTNAddItem.Focus()
                    End If

                Else
                    Me.PNKeyQty.Visible = True
                    Me.PNKeyQty.BringToFront()
                    Me.TBItemCode.Text = ds1.Tables(0).Rows(n)("code").ToString
                    Me.TBItemName.Text = ds1.Tables(0).Rows(n)("name1").ToString
                    Me.TBItemUnit.Text = ds1.Tables(0).Rows(n)("unitcode").ToString

                End If
            End If
        End If

        If e.KeyCode = Keys.Up Then
            Me.CMBReason.Focus()
            Me.CMBReason.SelectedIndex = 0
        End If

        If e.KeyCode = Keys.Down Then
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Focus()
                Me.ListViewItem.Items(0).Selected = True
                Me.ListViewItem.Items(0).Focused = True
            End If
        End If

        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call ClearData()
        End If

        If e.KeyCode = 114 Then
            Call SearchData()
        End If

        If e.KeyCode = 115 Then
            Call PrintData()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If
    End Sub


    Private Sub TBBarCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged

    End Sub

    Private Sub BTNAddShelf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddShelf.Click
        Me.PNAddStock.Visible = True
        Me.PNAddStock.BringToFront()
        Call vGetShelf()
        Me.CMBShelf.Focus()
    End Sub

    Public Sub AddNewItemShelf()
        Me.PNAddStock.Visible = True
        Me.PNAddStock.BringToFront()
        Call vGetShelf()
        Me.CMBShelf.Focus()
    End Sub

    Private Sub BTNAddStockOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddStockOK.Click
        Dim vShelfCode As String
        Dim vAddQty As Double
        Dim i As Integer
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfID As String
        Dim vIndex As Integer
        Dim vCheckShelfCode As String
        Dim n As Integer

        If Me.CMBShelf.Items.Count > 0 And Me.CMBShelf.Text <> "" And Me.TBAddQty.Text <> "" Then
            Me.PNAddStock.Visible = False
            Me.PNKeyQty.BringToFront()
            vShelfCode = Me.CMBShelf.Text
            vAddQty = Me.TBAddQty.Text
            vUnitCode = Me.TBItemUnit.Text
            vWHCode = "S02"
            vShelfID = Me.TBAddShelfID.Text

            For n = 0 To Me.ListViewStock.Items.Count - 1
                vCheckShelfCode = Me.ListViewStock.Items(n).SubItems(1).Text

                If vShelfCode = vCheckShelfCode Then
                    MsgBox("ชั้นเก็บ " & vShelfCode & " มีอยู่แล้วไม่สามารถเพิ่มได้อีก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBAddQty.Text = ""
                    Me.ListViewStock.Focus()
                    Me.ListViewStock.Items(0).Selected = True
                    Me.ListViewStock.Items(0).Focused = True
                    Exit Sub
                End If
            Next


            i = Me.ListViewStock.Items.Count + 1
            Dim listItem As New ListViewItem(i)
            listItem.SubItems.Add(vShelfCode)
            listItem.SubItems.Add(Format(vAddQty, "##,##0.00"))
            listItem.SubItems.Add(vUnitCode)
            listItem.SubItems.Add(Format(0, "##,##0.00"))
            listItem.SubItems.Add(vWHCode)
            listItem.SubItems.Add(vShelfID)
            listItem.SubItems.Add(2)
            listItem.SubItems.Add(0)
            listItem.SubItems.Add(0)
            Me.ListViewStock.Items.Add(listItem)

            Me.TBAddShelfID.Text = ""
            Me.TBAddQty.Text = ""

            vIndex = Me.ListViewStock.Items.Count - 1
            If Me.ListViewStock.Items.Count > 0 Then
                Me.ListViewStock.Focus()
                Me.ListViewStock.Items(vIndex).Selected = True
                Me.ListViewStock.Items(vIndex).Focused = True
            End If
        Else
            MsgBox("เมื่อต้องการเพิ่มการนับของชั้นเก็บ ต้องระบุชั้นเก็บและจำนวนที่นับได้", MsgBoxStyle.Critical, "Send Error Message")
            Me.CMBShelf.Focus()
        End If
    End Sub

    Public Sub vGetShelf()
        Dim i As Integer

        vQuery = "exec dbo.USP_NP_SearchWarehouse 'S02'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        Me.CMBShelf.Items.Clear()
        If ds.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds.Tables(0).Rows.Count - 1
                Me.CMBShelf.Items.Add(ds.Tables(0).Rows(i)("code").ToString)
            Next
        End If
    End Sub

    Public Sub vGetReasonCountStock()
        Dim i As Integer

        vQuery = "exec dbo.USP_MB_SearchCauseProductNegative"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        Me.CMBReason.Items.Clear()

        If ds.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds.Tables(0).Rows.Count - 1
                Me.CMBReason.Items.Add(ds.Tables(0).Rows(i)("causename").ToString)
            Next
        End If
    End Sub

    Private Sub FrmCountStock_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call vGetReasonCountStock()
        If Me.CMBReason.Items.Count > 0 Then
            Me.CMBReason.SelectedIndex = 0
        End If
    End Sub

    Private Sub ListViewStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewStock.KeyDown
        Dim vIndex As Integer
        Dim vShelfCode As String
        Dim vQty As Double
        Dim vShelfID As String
        Dim vItemIndex As Integer


        If e.KeyCode = Keys.Enter Then
            If Me.ListViewStock.Items.Count > 0 Then

                vIndex = Me.ListViewStock.FocusedItem.Index

                vShelfCode = Me.ListViewStock.Items(vIndex).SubItems(1).Text
                If Me.ListViewStock.Items(vIndex).SubItems(2).Text <> "" Then
                    vQty = Me.ListViewStock.Items(vIndex).SubItems(2).Text
                Else
                    vQty = 0
                End If

                If vQty <> 0 Then
                    Me.TBQty.Text = Format(vQty, "##,##0.00")
                End If

                If Me.ListViewStock.Items(vIndex).SubItems(6).Text <> "" Then
                    vShelfID = Me.ListViewStock.Items(vIndex).SubItems(6).Text
                Else
                    vShelfID = ""
                End If

                Me.LBLIndex.Text = vIndex
                Me.TBItemShelf.Text = vShelfCode
                Me.TBShelfID.Text = vShelfID
                Me.PNItemQty.Visible = True
                Me.PNItemQty.BringToFront()

                If TBShelfID.Text = "" Then
                    Me.TBShelfID.Focus()
                    Me.TBShelfID.SelectAll()
                Else
                    Me.TBQty.Focus()
                    Me.TBQty.SelectAll()
                End If
            End If

        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBItemCode.Text = ""
            Me.TBItemName.Text = ""
            Me.TBItemUnit.Text = ""
            Me.TBQty.Text = ""
            Me.TBShelfID.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNKeyQty.Visible = False
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If

        If e.KeyCode = 114 Then
            Me.PNAddStock.Visible = True
            Me.PNAddStock.BringToFront()
            Call vGetShelf()
            Me.CMBShelf.Focus()
        End If

        If e.KeyCode = 34 Then
            Call AddItem()
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseAddItem()
        End If

        Dim vIndexDel As Integer
        Dim vAnswer As Integer
        Dim vCheckType As Integer

        If e.KeyCode = Keys.Back Then
            If Me.ListViewStock.Items.Count > 0 Then
                vIndexDel = Me.ListViewStock.FocusedItem.Index
                vCheckType = Me.ListViewStock.Items(vIndexDel).SubItems(7).Text

                vAnswer = MsgBox("คุณต้องการ ลบรายการตรวจนับ รายการที่ " & vIndexDel + 1 & " นี้ใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message ?")

                If vAnswer = 6 Then
                    If vCheckType = 2 Then
                        vItemIndex = Me.ListViewStock.Items(vIndexDel).SubItems(9).Text
                        If Me.ListViewItem.Items.Count > 0 Then
                            Dim vCheckItemCode As String
                            Dim vCheckshelfCode As String
                            Dim vItemCodeLine As String
                            Dim vShelfCodeLine As String
                            Dim i As Integer

                            vItemCodeLine = Me.TBItemCode.Text
                            vShelfCodeLine = Me.ListViewStock.Items(vIndexDel).SubItems(1).Text

                            For i = 0 To Me.ListViewItem.Items.Count - 1
                                vCheckItemCode = Me.ListViewItem.Items(i).SubItems(4).Text
                                vCheckshelfCode = Me.ListViewItem.Items(i).SubItems(6).Text
                                If vItemCodeLine = vCheckItemCode And vShelfCodeLine = vCheckshelfCode Then
                                    Me.ListViewItem.Items.RemoveAt(vItemIndex)
                                End If
                            Next
                        End If
                        Me.ListViewStock.Items.RemoveAt(vIndexDel)
                        Call UpdateLineNumber()

                        MsgBox("ลบรายการนับสต๊อก เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")

                        If Me.ListViewStock.Items.Count > 0 Then
                            Me.ListViewStock.Focus()
                            Me.ListViewStock.Items(0).Selected = True
                            Me.ListViewStock.Items(0).Focused = True
                        Else
                            Me.TBItemCode.Focus()
                            Me.TBItemCode.SelectAll()
                        End If
                    Else
                        MsgBox("รายการชั้นเก็บที่ได้จาก ข้อมูลสต๊อกจริง ๆ ไม่สามารถลบได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                        Me.ListViewStock.Focus()
                        Me.ListViewStock.Items(vIndexDel).Selected = True
                        Me.ListViewStock.Items(vIndexDel).Focused = True
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub UpdateLineNumber()
        Dim i As Integer
        Dim vLine As Integer

        For i = 0 To Me.ListViewStock.Items.Count - 1
            vLine = vLine + 1
            Me.ListViewStock.Items(i).SubItems(0).Text = vLine
        Next

    End Sub

    Private Sub BTNKeyQty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNKeyQty.Click
        Dim vQty As Double
        Dim vIndex As Integer
        Dim vShelfID As String
        Dim vAnswer As Integer
        Dim vCheckShelf As Integer
        Dim vItemIndex As Integer


        If Me.TBQty.Text <> "" Then
            vIndex = Me.LBLIndex.Text
            vQty = Me.TBQty.Text
            vShelfID = Me.TBShelfID.Text

            If Me.TBShelfID.Text <> "" Then
                vShelfID = Me.TBShelfID.Text

                vQuery = "exec dbo.USP_NP_CheckShelfID '" & vShelfID & "' "
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count > 0 Then
                    vCheckShelf = 1
                Else
                    vCheckShelf = 0
                End If
            End If

            If vCheckShelf = 0 Then
                vAnswer = MsgBox("ไม่มีทะเบียนที่เก็บ ที่ได้ระบุไว้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Question Message")

                If vAnswer = 6 Then
                    Me.TBShelfID.Focus()
                    Me.TBShelfID.SelectAll()
                    Exit Sub
                End If
            End If

            vItemIndex = vIndex + 1

            Me.ListViewStock.Items(vIndex).SubItems(2).Text = Format(vQty, "##,##0.00")

            If vShelfID <> "" Then
                Me.ListViewStock.Items(vIndex).SubItems(6).Text = vShelfID
            End If

            Me.TBQty.Text = ""
            Me.TBShelfID.Text = ""
            Me.PNItemQty.Visible = False

            If vIndex < Me.ListViewStock.Items.Count - 1 Then
                Me.ListViewStock.Focus()
                Me.ListViewStock.Items(vItemIndex).Selected = True
                Me.ListViewStock.Items(vItemIndex).Focused = True
            ElseIf vIndex = Me.ListViewStock.Items.Count - 1 Then
                Me.ListViewStock.Focus()
                Me.ListViewStock.Items(vIndex).Selected = True
                Me.ListViewStock.Items(vIndex).Focused = True
            End If
        Else
            MsgBox("กรณีสินค้าไม่มียอดที่นับได้ ให้กรอกเป็น 0 ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBQty.Focus()
            Me.TBQty.SelectAll()
        End If

    End Sub

    Private Sub TBQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBQty.KeyDown
        Dim vQty As Double
        Dim vIndex As Integer
        Dim vItemIndex As Integer
        Dim vShelfID As String
        Dim vCheckShelf As Integer
        Dim vAnswer As Integer

        If e.KeyCode = Keys.Up Then
            Me.TBShelfID.Focus()
            Me.TBShelfID.SelectAll()
        End If

        If e.KeyCode = Keys.Down Then
            Me.BTNKeyQty.Focus()
        End If

        If e.KeyCode = Keys.Enter Then
            If Me.TBQty.Text <> "" Then
                vIndex = Me.LBLIndex.Text
                vQty = Me.TBQty.Text
                vShelfID = Me.TBShelfID.Text

                If Me.TBShelfID.Text <> "" Then
                    vShelfID = Me.TBShelfID.Text

                    vQuery = "exec dbo.USP_NP_CheckShelfID '" & vShelfID & "' "
                    Dim vService As New WebReference.WebServiceCalc
                    Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                    If ds.Tables(0).Rows.Count > 0 Then
                        vCheckShelf = 1
                    Else
                        vCheckShelf = 0
                    End If
                End If

                If vCheckShelf = 0 Then
                    vAnswer = MsgBox("ไม่มีทะเบียนที่เก็บ ที่ได้ระบุไว้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Question Message")

                    If vAnswer = 6 Then
                        Me.TBShelfID.Focus()
                        Me.TBShelfID.SelectAll()
                        Exit Sub
                    End If
                End If

                vItemIndex = vIndex + 1

                Me.ListViewStock.Items(vIndex).SubItems(2).Text = Format(vQty, "##,##0.00")

                If vShelfID <> "" Then
                    Me.ListViewStock.Items(vIndex).SubItems(6).Text = vShelfID
                End If

                Me.TBQty.Text = ""
                Me.TBShelfID.Text = ""
                Me.PNItemQty.Visible = False

                If vIndex < Me.ListViewStock.Items.Count - 1 Then
                    Me.ListViewStock.Focus()
                    Me.ListViewStock.Items(vItemIndex).Selected = True
                    Me.ListViewStock.Items(vItemIndex).Focused = True
                ElseIf vIndex = Me.ListViewStock.Items.Count - 1 Then
                    Me.ListViewStock.Focus()
                    Me.ListViewStock.Items(vIndex).Selected = True
                    Me.ListViewStock.Items(vIndex).Focused = True
                End If
            Else
                MsgBox("กรณีสินค้าไม่มียอดที่นับได้ ให้กรอกเป็น 0 ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBQty.Focus()
                Me.TBQty.SelectAll()
            End If
        End If

        'If e.KeyCode = Keys.Escape Then
        '    vIndex = Me.LBLIndex.Text
        '    vItemIndex = vIndex + 1
        '    Me.PNItemQty.Visible = False
        '    If vIndex < Me.ListViewStock.Items.Count - 1 Then
        '        Me.ListViewStock.Focus()
        '        Me.ListViewStock.Items(vItemIndex).Selected = True
        '        Me.ListViewStock.Items(vItemIndex).Focused = True
        '    ElseIf vIndex = Me.ListViewStock.Items.Count - 1 Then
        '        Me.ListViewStock.Focus()
        '        Me.ListViewStock.Items(vIndex).Selected = True
        '        Me.ListViewStock.Items(vIndex).Focused = True
        '    End If
        'End If

        If e.KeyCode = Keys.Escape Then
            Call CloseKeyQty()
        End If
    End Sub

    Public Sub CloseKeyQty()
        Dim vIndex As Integer
        Dim vItemIndex As Integer

        vIndex = Me.LBLIndex.Text
        vItemIndex = vIndex + 1
        Me.PNItemQty.Visible = False
        If vIndex < Me.ListViewStock.Items.Count - 1 Then
            Me.ListViewStock.Focus()
            Me.ListViewStock.Items(vItemIndex).Selected = True
            Me.ListViewStock.Items(vItemIndex).Focused = True
        ElseIf vIndex = Me.ListViewStock.Items.Count - 1 Then
            Me.ListViewStock.Focus()
            Me.ListViewStock.Items(vIndex).Selected = True
            Me.ListViewStock.Items(vIndex).Focused = True
        End If
    End Sub

    Private Sub TBQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBQty.TextChanged

    End Sub

    Private Sub BTNKeyQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNKeyQty.KeyDown
        Dim vItemIndex As Integer
        Dim vIndex As Integer

        If e.KeyCode = Keys.Escape Then
            vIndex = Me.LBLIndex.Text
            vItemIndex = vIndex + 1
            Me.PNItemQty.Visible = False
            If vIndex < Me.ListViewStock.Items.Count - 1 Then
                Me.ListViewStock.Focus()
                Me.ListViewStock.Items(vItemIndex).Selected = True
                Me.ListViewStock.Items(vItemIndex).Focused = True
            ElseIf vIndex = Me.ListViewStock.Items.Count - 1 Then
                Me.ListViewStock.Focus()
                Me.ListViewStock.Items(vIndex).Selected = True
                Me.ListViewStock.Items(vIndex).Focused = True
            End If
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBQty.Focus()
            Me.TBQty.SelectAll()
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseKeyQty()
        End If
    End Sub

    Private Sub CMBReason_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBReason.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Me.CMBReason.Items.Count > 0 Then
                Me.CMBReason.SelectedItem = Me.CMBReason.SelectedIndex
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
            End If

        End If

        If e.KeyCode = Keys.Down Then
            If Me.CMBReason.Items.Count > 0 Then
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
            End If
        End If

        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call ClearData()
        End If

        If e.KeyCode = 114 Then
            Call SearchData()
        End If

        If e.KeyCode = 115 Then
            Call PrintData()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If
    End Sub

    Private Sub CMBReason_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBReason.SelectedIndexChanged

    End Sub

    Private Sub ListViewStock_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewStock.SelectedIndexChanged

    End Sub

    Private Sub BTNCloseAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseAddItem.Click
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBItemUnit.Text = ""
        Me.TBQty.Text = ""
        Me.TBShelfID.Text = ""
        Me.ListViewStock.Items.Clear()
        Me.PNKeyQty.Visible = False
        Me.TBBarCode.Focus()
        Me.TBBarCode.SelectAll()
    End Sub

    Public Sub CloseAddItem()
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBItemUnit.Text = ""
        Me.TBQty.Text = ""
        Me.TBShelfID.Text = ""
        Me.ListViewStock.Items.Clear()
        Me.PNKeyQty.Visible = False
        Me.TBBarCode.Focus()
        Me.TBBarCode.SelectAll()
    End Sub

    Private Sub BTNAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddItem.Click
        Dim i As Integer
        Dim n As Integer
        Dim x As Integer
        Dim vLineIndex As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vCountQty As Double
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vCountItem As Integer
        Dim vTypeAdd As Integer
        Dim vStkQty As Double
        Dim vCountAddQty As Integer

        Dim vCheckItemCode As String
        Dim vCheckWHCode As String
        Dim vCheckShelfCode As String
        Dim vCheckUnitCode As String
        Dim vEditIndex As Integer
        Dim vEdit As Integer
        Dim vReasonCode As String

        If Me.ListViewStock.Items.Count > 0 Then

            For x = 0 To Me.ListViewStock.Items.Count - 1
                If Me.ListViewStock.Items(x).SubItems(2).Text <> "" Then
                    vCountAddQty = vCountAddQty + 1
                End If
            Next

            If Me.CMBReason.Text <> "" Then
                vReasonCode = vb6.Left(Me.CMBReason.Text, 6)
            Else
                MsgBox("กรุณากรอก เหตุผลของการตรวจนับสินค้าด้วย", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBReason.Focus()
                Exit Sub
            End If

            If vCountAddQty < Me.ListViewStock.Items.Count Then
                MsgBox("กรุณา กรอกผลการนับสินค้าทุกรายการที่ได้ทำการนับ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.ListViewStock.Focus()
                Me.ListViewStock.Items(0).Selected = True
                Me.ListViewStock.Items(0).Focused = True
                Exit Sub
            End If


            vItemCode = Me.TBItemCode.Text
            vItemName = Me.TBItemName.Text

            If Me.ListViewItem.Items.Count = 0 Then
                vLineIndex = 1
                vShelfCode = Me.ListViewStock.Items(0).SubItems(1).Text
                vCountQty = Me.ListViewStock.Items(0).SubItems(2).Text
                vUnitCode = Me.ListViewStock.Items(0).SubItems(3).Text
                vWHCode = Me.ListViewStock.Items(0).SubItems(5).Text
                vShelfID = Me.ListViewStock.Items(0).SubItems(6).Text
                vTypeAdd = Me.ListViewStock.Items(0).SubItems(7).Text

                If Me.ListViewStock.Items(0).SubItems(4).Text <> "" Then
                    vStkQty = Me.ListViewStock.Items(0).SubItems(4).Text
                Else
                    vStkQty = 0
                End If

                Dim listItem As New ListViewItem(vLineIndex)
                listItem.SubItems.Add(vItemName)
                listItem.SubItems.Add(Format(vCountQty, "##,##0.00"))
                listItem.SubItems.Add(vUnitCode)
                listItem.SubItems.Add(vItemCode)
                listItem.SubItems.Add(vWHCode)
                listItem.SubItems.Add(vShelfCode)
                listItem.SubItems.Add(vShelfID)
                listItem.SubItems.Add(vTypeAdd)
                listItem.SubItems.Add(Format(vStkQty, "##,##0.00"))
                listItem.SubItems.Add(vReasonCode)
                Me.ListViewItem.Items.Add(listItem)
            End If

            vLineIndex = Me.ListViewItem.Items.Count + 1


            For i = 0 To Me.ListViewStock.Items.Count - 1
                vShelfCode = Me.ListViewStock.Items(i).SubItems(1).Text
                vCountQty = Me.ListViewStock.Items(i).SubItems(2).Text
                vUnitCode = Me.ListViewStock.Items(i).SubItems(3).Text
                vWHCode = Me.ListViewStock.Items(i).SubItems(5).Text
                vShelfID = Me.ListViewStock.Items(i).SubItems(6).Text
                vTypeAdd = Me.ListViewStock.Items(i).SubItems(7).Text

                If Me.ListViewStock.Items(i).SubItems(4).Text <> "" Then
                    vStkQty = Me.ListViewStock.Items(i).SubItems(4).Text
                Else
                    vStkQty = 0
                End If

                If Me.ListViewStock.Items(i).SubItems(8).Text <> 0 Then
                    vEdit = 1
                    vEditIndex = Me.ListViewStock.Items(i).SubItems(9).Text
                Else
                    vEdit = 0
                End If

                For n = 0 To Me.ListViewItem.Items.Count - 1
                    vCheckItemCode = Me.ListViewItem.Items(n).SubItems(4).Text
                    vCheckUnitCode = Me.ListViewItem.Items(n).SubItems(3).Text
                    vCheckWHCode = Me.ListViewItem.Items(n).SubItems(5).Text
                    vCheckShelfCode = Me.ListViewItem.Items(n).SubItems(6).Text

                    If vEdit = 0 Then
                        If vItemCode <> vCheckItemCode Or vShelfCode <> vCheckShelfCode Then

                            Dim listItem As New ListViewItem(vLineIndex)
                            listItem.SubItems.Add(vItemName)
                            listItem.SubItems.Add(Format(vCountQty, "##,##0.00"))
                            listItem.SubItems.Add(vUnitCode)
                            listItem.SubItems.Add(vItemCode)
                            listItem.SubItems.Add(vWHCode)
                            listItem.SubItems.Add(vShelfCode)
                            listItem.SubItems.Add(vShelfID)
                            listItem.SubItems.Add(vTypeAdd)
                            listItem.SubItems.Add(Format(vStkQty, "##,##0.00"))
                            listItem.SubItems.Add(vReasonCode)
                            Me.ListViewItem.Items.Add(listItem)

                            vLineIndex = vLineIndex + 1
                            GoTo Line1

                        End If

                    ElseIf vEdit = 1 Then
                        Me.ListViewItem.Items(vEditIndex).SubItems(2).Text = Format(vCountQty, "##,##0.00")
                    End If

                Next
Line1:
            Next

            vCountItem = Me.ListViewItem.Items.Count - 1
            Me.TBItemCode.Text = ""
            Me.TBItemName.Text = ""
            Me.TBItemUnit.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNKeyQty.Visible = False
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        Else
            MsgBox("ไม่มีรายการสินค้าให้เพิ่ม กรณีไม่เพิ่มก็ให้กดปุ่ม ESC ออกหน้านี้", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNAddShelf.Focus()
        End If
    End Sub

    Public Sub AddItem()
        Dim i As Integer
        Dim n As Integer
        Dim x As Integer
        Dim vLineIndex As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vCountQty As Double
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vCountItem As Integer
        Dim vTypeAdd As Integer
        Dim vStkQty As Double
        Dim vCountAddQty As Integer

        Dim vCheckItemCode As String
        Dim vCheckWHCode As String
        Dim vCheckShelfCode As String
        Dim vCheckUnitCode As String
        Dim vEditIndex As Integer
        Dim vEdit As Integer
        Dim vReasonCode As String

        If Me.ListViewStock.Items.Count > 0 Then

            For x = 0 To Me.ListViewStock.Items.Count - 1
                If Me.ListViewStock.Items(x).SubItems(2).Text <> "" Then
                    vCountAddQty = vCountAddQty + 1
                End If
            Next

            If Me.CMBReason.Text <> "" Then
                vReasonCode = vb6.Left(Me.CMBReason.Text, 6)
            Else
                MsgBox("กรุณากรอก เหตุผลของการตรวจนับสินค้าด้วย", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBReason.Focus()
                Exit Sub
            End If

            If vCountAddQty < Me.ListViewStock.Items.Count Then
                MsgBox("กรุณา กรอกผลการนับสินค้าทุกรายการที่ได้ทำการนับ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.ListViewStock.Focus()
                Me.ListViewStock.Items(0).Selected = True
                Me.ListViewStock.Items(0).Focused = True
                Exit Sub
            End If


            vItemCode = Me.TBItemCode.Text
            vItemName = Me.TBItemName.Text

            If Me.ListViewItem.Items.Count = 0 Then
                vLineIndex = 1
                vShelfCode = Me.ListViewStock.Items(0).SubItems(1).Text
                vCountQty = Me.ListViewStock.Items(0).SubItems(2).Text
                vUnitCode = Me.ListViewStock.Items(0).SubItems(3).Text
                vWHCode = Me.ListViewStock.Items(0).SubItems(5).Text
                vShelfID = Me.ListViewStock.Items(0).SubItems(6).Text
                vTypeAdd = Me.ListViewStock.Items(0).SubItems(7).Text

                If Me.ListViewStock.Items(0).SubItems(4).Text <> "" Then
                    vStkQty = Me.ListViewStock.Items(0).SubItems(4).Text
                Else
                    vStkQty = 0
                End If

                Dim listItem As New ListViewItem(vLineIndex)
                listItem.SubItems.Add(vItemName)
                listItem.SubItems.Add(Format(vCountQty, "##,##0.00"))
                listItem.SubItems.Add(vUnitCode)
                listItem.SubItems.Add(vItemCode)
                listItem.SubItems.Add(vWHCode)
                listItem.SubItems.Add(vShelfCode)
                listItem.SubItems.Add(vShelfID)
                listItem.SubItems.Add(vTypeAdd)
                listItem.SubItems.Add(Format(vStkQty, "##,##0.00"))
                listItem.SubItems.Add(vReasonCode)
                Me.ListViewItem.Items.Add(listItem)
            End If

            vLineIndex = Me.ListViewItem.Items.Count + 1


            For i = 0 To Me.ListViewStock.Items.Count - 1
                vShelfCode = Me.ListViewStock.Items(i).SubItems(1).Text
                vCountQty = Me.ListViewStock.Items(i).SubItems(2).Text
                vUnitCode = Me.ListViewStock.Items(i).SubItems(3).Text
                vWHCode = Me.ListViewStock.Items(i).SubItems(5).Text
                vShelfID = Me.ListViewStock.Items(i).SubItems(6).Text
                vTypeAdd = Me.ListViewStock.Items(i).SubItems(7).Text

                If Me.ListViewStock.Items(i).SubItems(4).Text <> "" Then
                    vStkQty = Me.ListViewStock.Items(i).SubItems(4).Text
                Else
                    vStkQty = 0
                End If

                If Me.ListViewStock.Items(i).SubItems(8).Text <> 0 Then
                    vEdit = 1
                    vEditIndex = Me.ListViewStock.Items(i).SubItems(9).Text
                Else
                    vEdit = 0
                End If

                For n = 0 To Me.ListViewItem.Items.Count - 1
                    vCheckItemCode = Me.ListViewItem.Items(n).SubItems(4).Text
                    vCheckUnitCode = Me.ListViewItem.Items(n).SubItems(3).Text
                    vCheckWHCode = Me.ListViewItem.Items(n).SubItems(5).Text
                    vCheckShelfCode = Me.ListViewItem.Items(n).SubItems(6).Text

                    If vEdit = 0 Then
                        If vItemCode <> vCheckItemCode Or vShelfCode <> vCheckShelfCode Then

                            Dim listItem As New ListViewItem(vLineIndex)
                            listItem.SubItems.Add(vItemName)
                            listItem.SubItems.Add(Format(vCountQty, "##,##0.00"))
                            listItem.SubItems.Add(vUnitCode)
                            listItem.SubItems.Add(vItemCode)
                            listItem.SubItems.Add(vWHCode)
                            listItem.SubItems.Add(vShelfCode)
                            listItem.SubItems.Add(vShelfID)
                            listItem.SubItems.Add(vTypeAdd)
                            listItem.SubItems.Add(Format(vStkQty, "##,##0.00"))
                            listItem.SubItems.Add(vReasonCode)
                            Me.ListViewItem.Items.Add(listItem)

                            vLineIndex = vLineIndex + 1
                            GoTo Line1

                        End If

                    ElseIf vEdit = 1 Then
                        Me.ListViewItem.Items(vEditIndex).SubItems(2).Text = Format(vCountQty, "##,##0.00")
                    End If

                Next
Line1:
            Next

            vCountItem = Me.ListViewItem.Items.Count - 1
            Me.TBItemCode.Text = ""
            Me.TBItemName.Text = ""
            Me.TBItemUnit.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNKeyQty.Visible = False
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        Else
            MsgBox("ไม่มีรายการสินค้าให้เพิ่ม กรณีไม่เพิ่มก็ให้กดปุ่ม ESC ออกหน้านี้", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNAddShelf.Focus()
        End If
    End Sub

    Private Sub TBShelfID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBShelfID.KeyDown
        Dim vShelfID As String
        Dim vCheckShelf As Integer

        If e.KeyCode = Keys.Down Then
            Me.TBQty.Focus()
            Me.TBQty.SelectAll()
        End If

        If e.KeyCode = Keys.Enter Then
            If Me.TBShelfID.Text <> "" Then
                vShelfID = Me.TBShelfID.Text

                vQuery = "exec dbo.USP_NP_CheckShelfID '" & vShelfID & "' "
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count > 0 Then
                    vCheckShelf = 1
                Else
                    vCheckShelf = 0
                End If
            End If

            If vCheckShelf = 0 Then
                MsgBox("ไม่มีทะเบียนชั้นเก็บ ที่ได้ระบุไว้ กรุณาแก้ไข", MsgBoxStyle.Critical, "Send Error Message")

                Me.TBShelfID.Focus()
                Me.TBShelfID.SelectAll()
            Else
                Me.TBShelfID.Text = UCase(Me.TBShelfID.Text)
                Me.TBQty.Focus()
                Me.TBQty.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseKeyQty()
        End If
    End Sub

    Private Sub TBShelfID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBShelfID.TextChanged

    End Sub

    Private Sub ListViewItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItem.KeyDown
        If e.KeyCode = Keys.Up Then
            If Me.ListViewItem.Items(0).Focused = True Then
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
            End If
        End If

        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call ClearData()
        End If

        If e.KeyCode = 114 Then
            Call SearchData()
        End If

        If e.KeyCode = 115 Then
            Call PrintData()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If
    End Sub

    Private Sub BTNCloseAddItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCloseAddItem.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.TBItemCode.Text = ""
            Me.TBItemName.Text = ""
            Me.TBItemUnit.Text = ""
            Me.TBQty.Text = ""
            Me.TBShelfID.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNKeyQty.Visible = False
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If

        If e.KeyCode = 114 Then
            Me.PNAddStock.Visible = True
            Me.PNAddStock.BringToFront()
            Call vGetShelf()
            Me.CMBShelf.Focus()
        End If

        If e.KeyCode = 34 Then
            Call AddItem()
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseAddItem()
        End If
    End Sub

    Private Sub BTNAddItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNAddItem.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.TBItemCode.Text = ""
            Me.TBItemName.Text = ""
            Me.TBItemUnit.Text = ""
            Me.TBQty.Text = ""
            Me.TBShelfID.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNKeyQty.Visible = False
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If

        If e.KeyCode = 114 Then
            Me.PNAddStock.Visible = True
            Me.PNAddStock.BringToFront()
            Call vGetShelf()
            Me.CMBShelf.Focus()
        End If

        If e.KeyCode = 34 Then
            Call AddItem()
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseAddItem()
        End If
    End Sub

    Private Sub BTNAddShelf_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNAddShelf.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.TBItemCode.Text = ""
            Me.TBItemName.Text = ""
            Me.TBItemUnit.Text = ""
            Me.TBQty.Text = ""
            Me.TBShelfID.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNKeyQty.Visible = False
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If

        If e.KeyCode = 114 Then
            Me.PNAddStock.Visible = True
            Me.PNAddStock.BringToFront()
            Call vGetShelf()
            Me.CMBShelf.Focus()
        End If

        If e.KeyCode = 34 Then
            Call AddItem()
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseAddItem()
        End If
    End Sub

    Private Sub TBItemCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBItemCode.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.TBItemCode.Text = ""
            Me.TBItemName.Text = ""
            Me.TBItemUnit.Text = ""
            Me.TBQty.Text = ""
            Me.TBShelfID.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNKeyQty.Visible = False
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If

        If e.KeyCode = 114 Then
            Me.PNAddStock.Visible = True
            Me.PNAddStock.BringToFront()
            Call vGetShelf()
            Me.CMBShelf.Focus()
        End If

        If e.KeyCode = 34 Then
            Call AddItem()
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseAddItem()
        End If
    End Sub

    Private Sub TBItemCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBItemCode.TextChanged

    End Sub

    Private Sub TBItemName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBItemName.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.TBItemCode.Text = ""
            Me.TBItemName.Text = ""
            Me.TBItemUnit.Text = ""
            Me.TBQty.Text = ""
            Me.TBShelfID.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNKeyQty.Visible = False
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If

        If e.KeyCode = 114 Then
            Me.PNAddStock.Visible = True
            Me.PNAddStock.BringToFront()
            Call vGetShelf()
            Me.CMBShelf.Focus()
        End If

        If e.KeyCode = 34 Then
            Call AddItem()
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseAddItem()
        End If
    End Sub

    Private Sub TBItemName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBItemName.TextChanged

    End Sub

    Private Sub TBItemUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBItemUnit.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.TBItemCode.Text = ""
            Me.TBItemName.Text = ""
            Me.TBItemUnit.Text = ""
            Me.TBQty.Text = ""
            Me.TBShelfID.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNKeyQty.Visible = False
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If

        If e.KeyCode = 114 Then
            Me.PNAddStock.Visible = True
            Me.PNAddStock.BringToFront()
            Call vGetShelf()
            Me.CMBShelf.Focus()
        End If

        If e.KeyCode = 34 Then
            Call AddItem()
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseAddItem()
        End If
    End Sub

    Private Sub TBItemUnit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBItemUnit.TextChanged

    End Sub

    Private Sub BTNAddStockOK_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNAddStockOK.KeyDown
        If e.KeyCode = Keys.Escape Then
            Call CloseAddStock()
        End If
    End Sub

    Public Sub CloseAddStock()
        Dim vIndex As Integer

        Me.PNAddStock.Visible = False
        Me.TBAddQty.Text = ""
        If Me.ListViewStock.Items.Count > 0 Then
            vIndex = Me.ListViewStock.Items.Count - 1
            If Me.ListViewStock.Items.Count > 0 Then
                Me.ListViewStock.Focus()
                Me.ListViewStock.Items(vIndex).Selected = True
                Me.ListViewStock.Items(vIndex).Focused = True
            End If
        Else
            Me.BTNAddItem.Focus()
        End If

    End Sub

    Private Sub TBAddQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBAddQty.KeyDown
        Dim vIndex As Integer
        Dim vShelfCode As String
        Dim vAddQty As Double
        Dim i As Integer
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfID As String
        Dim vCheckShelfCode As String
        Dim n As Integer

        If e.KeyCode = Keys.Enter Then
            If Me.CMBShelf.Items.Count > 0 And Me.CMBShelf.Text <> "" And Me.TBAddQty.Text <> "" Then
                Me.PNAddStock.Visible = False
                Me.PNKeyQty.BringToFront()
                vShelfCode = Me.CMBShelf.Text
                vAddQty = Me.TBAddQty.Text
                vUnitCode = Me.TBItemUnit.Text
                vWHCode = "S02"
                vShelfID = Me.TBAddShelfID.Text

                For n = 0 To Me.ListViewStock.Items.Count - 1
                    vCheckShelfCode = Me.ListViewStock.Items(n).SubItems(1).Text

                    If vShelfCode = vCheckShelfCode Then
                        MsgBox("ชั้นเก็บ " & vShelfCode & " มีอยู่แล้วไม่สามารถเพิ่มได้อีก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                        Me.TBAddQty.Text = ""
                        Me.ListViewStock.Focus()
                        Me.ListViewStock.Items(0).Selected = True
                        Me.ListViewStock.Items(0).Focused = True
                        Exit Sub
                    End If
                Next

                i = Me.ListViewStock.Items.Count + 1
                Dim listItem As New ListViewItem(i)
                listItem.SubItems.Add(vShelfCode)
                listItem.SubItems.Add(Format(vAddQty, "##,##0.00"))
                listItem.SubItems.Add(vUnitCode)
                listItem.SubItems.Add(Format(0, "##,##0.00"))
                listItem.SubItems.Add(vWHCode)
                listItem.SubItems.Add(vShelfID)
                listItem.SubItems.Add(2)
                listItem.SubItems.Add(0)
                listItem.SubItems.Add(0)
                Me.ListViewStock.Items.Add(listItem)

                Me.TBAddShelfID.Text = ""
                Me.TBAddQty.Text = ""

                vIndex = Me.ListViewStock.Items.Count - 1
                If Me.ListViewStock.Items.Count > 0 Then
                    Me.ListViewStock.Focus()
                    Me.ListViewStock.Items(vIndex).Selected = True
                    Me.ListViewStock.Items(vIndex).Focused = True
                End If
            Else
                MsgBox("เมื่อต้องการเพิ่มการนับของชั้นเก็บ ต้องระบุชั้นเก็บและจำนวนที่นับได้", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBShelf.Focus()
            End If
        End If


        If e.KeyCode = Keys.Escape Then
            Call CloseAddStock()
        End If

    End Sub

    Private Sub TBAddQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBAddQty.KeyPress, TBQty.KeyPress
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

    Private Sub TBAddQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBAddQty.TextChanged

    End Sub

    Private Sub CMBShelf_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBShelf.KeyDown

        If e.KeyCode = Keys.Escape Then
            Call CloseAddStock()
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBAddShelfID.Focus()
            Me.TBAddShelfID.SelectAll()
        End If

        If e.KeyCode = Keys.Enter Then
            If Me.CMBShelf.Items.Count > 0 Then
                Me.CMBShelf.Text = Me.CMBShelf.SelectedIndex
                Me.TBAddShelfID.Focus()
                Me.TBAddShelfID.SelectAll()
            End If
        End If
    End Sub

    Private Sub CMBShelf_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBShelf.SelectedIndexChanged

    End Sub

    Private Sub TBAddShelfID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBAddShelfID.KeyDown
        Dim vShelfID As String
        Dim vCheckShelf As Integer


        If e.KeyCode = Keys.Up Then
            Me.CMBShelf.Focus()
        End If


        If e.KeyCode = Keys.Down Then
            Me.TBAddQty.Focus()
            Me.TBAddQty.SelectAll()
        End If

        If e.KeyCode = Keys.Enter Then
            If Me.TBAddShelfID.Text <> "" Then
                vShelfID = Me.TBAddShelfID.Text

                vQuery = "exec dbo.USP_NP_CheckShelfID '" & vShelfID & "' "
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count > 0 Then
                    vCheckShelf = 1
                Else
                    vCheckShelf = 0
                End If
            End If

            If vCheckShelf = 0 Then
                MsgBox("ไม่มีทะเบียนชั้นเก็บ ที่ได้ระบุไว้ กรุณาแก้ไข", MsgBoxStyle.Critical, "Send Error Message")

                Me.TBAddShelfID.Focus()
                Me.TBAddShelfID.SelectAll()
            Else
                Me.TBAddShelfID.Text = UCase(Me.TBAddShelfID.Text)
                Me.TBAddQty.Focus()
                Me.TBAddQty.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseAddStock()
        End If
    End Sub

    Private Sub TBAddShelfID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBAddShelfID.TextChanged

    End Sub

    Private Sub BTNDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNDelete.Click
        Dim i As Integer
        Dim vIndex As Integer
        Dim vAnswer As Integer
        Dim vItemCode As String
        Dim vCheckEdit As Integer

        If Me.TBItemCode.Text <> "" And Me.TBItemName.Text <> "" And Me.TBItemUnit.Text <> "" And Me.ListViewStock.Items.Count > 0 Then
            vItemCode = Me.TBItemCode.Text
            vAnswer = MsgBox("คุณต้องการลบ รายการสินค้าที่ตรวจนับ รายการ " & vItemCode & " นี้ทิ้งใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 Then
                For i = 0 To Me.ListViewStock.Items.Count - 1
                    vIndex = Me.ListViewStock.Items(i).SubItems(9).Text
                    vCheckEdit = Me.ListViewStock.Items(i).SubItems(8).Text
                    If vCheckEdit = 1 Then
                        Me.ListViewItem.Items.RemoveAt(vIndex)
                    End If
                Next
                Me.ListViewStock.Items.Clear()
                Me.TBItemCode.Text = ""
                Me.TBItemName.Text = ""
                Me.TBItemUnit.Text = ""
                Me.PNKeyQty.Visible = False
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
            Else
                Me.ListViewStock.Focus()
                Me.ListViewStock.Items(0).Selected = True
                Me.ListViewStock.Items(0).Focused = True
            End If
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim i As Integer
        Dim vItem As String
        Dim vUnit As String
        Dim vWH As String
        Dim vShelf As String
        Dim vItemName As String
        Dim vQty As Double
        Dim vLineNumber As Integer
        Dim vShelfStock As String
        Dim vReasonCode As String

        Dim vServerDate As Date
        Dim vDocDate As Date


        If Me.CMBReason.Text <> "" And Me.ListViewItem.Items.Count > 0 Then
            MsgBox("ครั้งต่อไปให้กดปุ่มสีส้ม+ปุ่มเลข 8", MsgBoxStyle.Information, "Send Information Message")
            Call BeforeSave()
            If Me.TBDocNo.Text = "" And Me.TBStockNo.Text = "" Then
                vQuery = "begin tran"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As Integer = vService.vExecuteQuery(vQuery)

                Call GetDocNo()
                For i = 0 To Me.ListViewItem.Items.Count - 1
                    vItem = Me.ListViewItem.Items(i).SubItems(4).Text
                    vItemName = Me.ListViewItem.Items(i).SubItems(1).Text
                    vUnit = Me.ListViewItem.Items(i).SubItems(3).Text
                    vWH = Me.ListViewItem.Items(i).SubItems(5).Text
                    vShelfStock = Me.ListViewItem.Items(i).SubItems(6).Text
                    vShelf = Me.ListViewItem.Items(i).SubItems(7).Text
                    vQty = Me.ListViewItem.Items(i).SubItems(2).Text
                    vReasonCode = Me.ListViewItem.Items(i).SubItems(10).Text
                    vLineNumber = i
                    vQuery = "exec dbo.USP_NP_InsertInspectLog '" & vGetDocNo & "','" & vItem & "','" & vItemName & "','" & vWH & "','" & vShelf & "'," & vQty & ",'" & vUnit & "','" & vPersonName & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)
                Next

                Call vGetInspectNo()

                Call AddItemInspect()

                vQuery = "commit tran"
                Dim vService2 As New WebReference.WebServiceCalc
                Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)
            End If

            If Me.TBDocNo.Text <> "" And Me.TBStockNo.Text <> "" Then

                vQuery = "exec dbo.USP_CN2_CheckDateNow"
                Dim vService3 As New WebReference.WebServiceCalc
                Dim ds3 As DataSet = vService3.vGetQueryAnlyzer(vQuery)

                If ds3.Tables(0).Rows.Count > 0 Then
                    vServerDate = ds3.Tables(0).Rows(0)("docdate").ToString
                End If

                vDocDate = Me.TBDocDate.Text

                If vServerDate <> vDocDate Then
                    MsgBox("ไม่สามารถแก้ไขเอกสาร ย้อนหลังจากวันที่ปัจจุบันได้ เนื่องจากยอดสต๊อกอาจมีการเคลื่อนไหวไปเยอะแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBBarCode.Focus()
                    Me.TBBarCode.SelectAll()
                    Call AfterSave()
                    Exit Sub
                End If

                vQuery = "begin tran"
                Dim vService4 As New WebReference.WebServiceCalc
                Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)

                vStkNo = Me.TBStockNo.Text

                vQuery = "delete npmaster.dbo.TB_IV_StkInspect_Logs where docno = '" & vStkNo & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As Integer = vService.vExecuteQuery(vQuery)

                For i = 0 To Me.ListViewItem.Items.Count - 1
                    vItem = Me.ListViewItem.Items(i).SubItems(4).Text
                    vItemName = Me.ListViewItem.Items(i).SubItems(1).Text
                    vUnit = Me.ListViewItem.Items(i).SubItems(3).Text
                    vWH = Me.ListViewItem.Items(i).SubItems(5).Text
                    vShelfStock = Me.ListViewItem.Items(i).SubItems(6).Text
                    vShelf = Me.ListViewItem.Items(i).SubItems(7).Text
                    vQty = Me.ListViewItem.Items(i).SubItems(2).Text
                    vReasonCode = Me.ListViewItem.Items(i).SubItems(10).Text
                    vLineNumber = i
                    vQuery = "exec dbo.USP_NP_InsertInspectLog '" & vStkNo & "','" & vItem & "','" & vItemName & "','" & vWH & "','" & vShelf & "'," & vQty & ",'" & vUnit & "','" & vPersonName & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)
                Next

                Call UpdateAddItemInspect()

                vQuery = "commit tran"
                Dim vService2 As New WebReference.WebServiceCalc
                Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)
            End If

            Call AfterSave()
            vDocno = ""
            vNewDocNo = ""
            vGetDocNo = ""
            vStkNo = ""
            Me.TBDocNo.Text = ""
            Me.TBDocDate.Text = ""
            Me.TBBarCode.Text = ""
            Me.TBStockNo.Text = ""
            Me.ListViewItem.Items.Clear()
            Me.ListViewStock.Items.Clear()
            Me.CMBReason.Focus()

        Else
            Call AfterSave()
            MsgBox("ไม่มีรายการสินค้าให้บันทึกข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If
    End Sub

    Public Sub SaveData()
        Dim i As Integer
        Dim vItem As String
        Dim vUnit As String
        Dim vWH As String
        Dim vShelf As String
        Dim vItemName As String
        Dim vQty As Double
        Dim vLineNumber As Integer
        Dim vShelfStock As String
        Dim vReasonCode As String

        Dim vServerDate As Date
        Dim vDocDate As Date


        If Me.CMBReason.Text <> "" And Me.ListViewItem.Items.Count > 0 Then
            Call BeforeSave()
            If Me.TBDocNo.Text = "" And Me.TBStockNo.Text = "" Then
                vQuery = "begin tran"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As Integer = vService.vExecuteQuery(vQuery)

                Call GetDocNo()
                For i = 0 To Me.ListViewItem.Items.Count - 1
                    vItem = Me.ListViewItem.Items(i).SubItems(4).Text
                    vItemName = Me.ListViewItem.Items(i).SubItems(1).Text
                    vUnit = Me.ListViewItem.Items(i).SubItems(3).Text
                    vWH = Me.ListViewItem.Items(i).SubItems(5).Text
                    vShelfStock = Me.ListViewItem.Items(i).SubItems(6).Text
                    vShelf = Me.ListViewItem.Items(i).SubItems(7).Text
                    vQty = Me.ListViewItem.Items(i).SubItems(2).Text
                    vReasonCode = Me.ListViewItem.Items(i).SubItems(10).Text
                    vLineNumber = i
                    vQuery = "exec dbo.USP_NP_InsertInspectLog '" & vGetDocNo & "','" & vItem & "','" & vItemName & "','" & vWH & "','" & vShelf & "'," & vQty & ",'" & vUnit & "','" & vPersonName & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)
                Next

                Call vGetInspectNo()

                Call AddItemInspect()

                vQuery = "commit tran"
                Dim vService2 As New WebReference.WebServiceCalc
                Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)
            End If

            If Me.TBDocNo.Text <> "" And Me.TBStockNo.Text <> "" Then

                vQuery = "exec dbo.USP_CN2_CheckDateNow"
                Dim vService3 As New WebReference.WebServiceCalc
                Dim ds3 As DataSet = vService3.vGetQueryAnlyzer(vQuery)

                If ds3.Tables(0).Rows.Count > 0 Then
                    vServerDate = ds3.Tables(0).Rows(0)("docdate").ToString
                End If

                vDocDate = Me.TBDocDate.Text

                If vServerDate <> vDocDate Then
                    MsgBox("ไม่สามารถแก้ไขเอกสาร ย้อนหลังจากวันที่ปัจจุบันได้ เนื่องจากยอดสต๊อกอาจมีการเคลื่อนไหวไปเยอะแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBBarCode.Focus()
                    Me.TBBarCode.SelectAll()
                    Call AfterSave()
                    Exit Sub
                End If

                vQuery = "begin tran"
                Dim vService4 As New WebReference.WebServiceCalc
                Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)

                vStkNo = Me.TBStockNo.Text

                vQuery = "delete npmaster.dbo.TB_IV_StkInspect_Logs where docno = '" & vStkNo & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As Integer = vService.vExecuteQuery(vQuery)

                For i = 0 To Me.ListViewItem.Items.Count - 1
                    vItem = Me.ListViewItem.Items(i).SubItems(4).Text
                    vItemName = Me.ListViewItem.Items(i).SubItems(1).Text
                    vUnit = Me.ListViewItem.Items(i).SubItems(3).Text
                    vWH = Me.ListViewItem.Items(i).SubItems(5).Text
                    vShelfStock = Me.ListViewItem.Items(i).SubItems(6).Text
                    vShelf = Me.ListViewItem.Items(i).SubItems(7).Text
                    vQty = Me.ListViewItem.Items(i).SubItems(2).Text
                    vReasonCode = Me.ListViewItem.Items(i).SubItems(10).Text
                    vLineNumber = i
                    vQuery = "exec dbo.USP_NP_InsertInspectLog '" & vStkNo & "','" & vItem & "','" & vItemName & "','" & vWH & "','" & vShelf & "'," & vQty & ",'" & vUnit & "','" & vPersonName & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)
                Next

                Call UpdateAddItemInspect()

                ''vQuery = "commit tran"
                ''Dim vService2 As New WebReference.WebServiceCalc
                ''Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)
            End If

            Call AfterSave()
            vDocno = ""
            vNewDocNo = ""
            vGetDocNo = ""
            vStkNo = ""
            Me.TBDocNo.Text = ""
            Me.TBDocDate.Text = ""
            Me.TBBarCode.Text = ""
            Me.TBStockNo.Text = ""
            Me.ListViewItem.Items.Clear()
            Me.ListViewStock.Items.Clear()
            Me.CMBReason.Focus()
        Else
            Call AfterSave()
            MsgBox("ไม่มีรายการสินค้าให้บันทึกข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If
    End Sub

    Public Sub GetDocNo()
        Dim vYear As String
        Dim vYear1 As Integer
        Dim vYear2 As String
        Dim vMonth, vMonth1 As String
        Dim vHeader As String
        Dim vHeader1 As String
        Dim vAutoNumber As Integer

        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_NP_SearchNewDocNo 10 "
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vHeader = ds.Tables(0).Rows(0)("header").ToString
            vAutoNumber = ds.Tables(0).Rows(0)("AutoNumber").ToString
        End If

        vYear = Mid(Year(Now), 2, 3)
        vYear1 = vYear
        If vYear1 < 543 Then
            vYear1 = vYear1 + 43
        End If
        vYear2 = vYear1
        vHeader1 = Trim(RTrim(LTrim(vHeader)) + RTrim(LTrim(vYear2)))
        vGetDocNo = vHeader1 + "-" + Format(vAutoNumber, "0000")

        vQuery = "exec dbo.USP_NP_UpdateNewDocNo  10 "
        Dim vService1 As New WebReference.WebServiceCalc
        Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description)
            Exit Sub
        End If
    End Sub

    Public Sub vGetInspectNo()
        Dim vCheckDocno As String
        Dim vYear As String
        Dim vYear1 As Integer
        Dim vYear2 As String
        Dim vMonth As String
        Dim vMonth1 As String

        vQuery = "select top 1 docno from dbo.bcstkinspect  order by docno desc"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vCheckDocno = ds.Tables(0).Rows(0)("docno").ToString
        End If


        If vb6.Left(vCheckDocno, 2) = "IH" Then
            vYear = Mid(vCheckDocno, 2, 2)
            vMonth = Mid(vCheckDocno, 5, 2)
            vYear1 = Mid(Year(Now), 3, 2)
            vMonth1 = Month(Now)

            vYear1 = vYear
            If vYear1 < 43 Then
                vYear1 = vYear1 + 43
            End If
            vYear2 = vYear1
            If Len(vMonth1) <> 2 Then
                vMonth1 = "0" & vMonth1
            End If

            If vYear2 = vYear And vMonth1 = vMonth Then
                vQuery = "select * from V_WEB_IV_ItemCheck_NewDocNo"
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)

                If ds1.Tables(0).Rows.Count > 0 Then
                    vNewDocNo = ds1.Tables(0).Rows(0)("newdocno").ToString
                End If
            Else
                vNewDocNo = "S02-IH" & Trim(vYear2 & vMonth1 & "-0001")
            End If

        ElseIf vb6.Left(vCheckDocno, 3) = "S02" Then

            Dim vLen As Integer
            Dim vDocNo As String

            vLen = Len(vCheckDocno)
            vDocNo = vb6.Right(vCheckDocno, vLen - 4)

            vYear = Mid(vDocNo, 3, 2)
            vMonth = Mid(vDocNo, 5, 2)
            vMonth1 = Month(Now)

            vYear1 = vYear
            If vYear1 < 43 Then
                vYear1 = vYear1 + 43
            End If
            vYear2 = vYear1

            If Len(vMonth1) <> 2 Then
                vMonth1 = "0" & vMonth1
            End If
            If vYear2 = vYear And vMonth1 = vMonth Then
                vQuery = "select * from V_WEB_IV_ItemCheck_NewDocNo"
                Dim vService2 As New WebReference.WebServiceCalc
                Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)

                If ds2.Tables(0).Rows.Count > 0 Then
                    vNewDocNo = ds2.Tables(0).Rows(0)("newdocno").ToString
                End If

            Else
                vNewDocNo = "S02-IH" & Trim(vYear2 & vMonth1 & "-0001")
            End If
        Else
            vYear = Mid(Year(Now), 2, 3)
            vYear1 = vYear
            If vYear1 < 543 Then
                vYear1 = vYear1 + 43
            End If
            vMonth1 = Month(Now)

            If Len(vMonth1) <> 2 Then
                vMonth1 = "0" & vMonth1
            End If
            vYear2 = vYear1

            vNewDocNo = "S02-IH" & Trim(vYear2 & vMonth1 & "-0001")
        End If
    End Sub


    Public Sub AddItemInspect()
        Dim vItem(500) As String
        Dim vUnitCode(500) As String
        Dim vShelf As String
        Dim vItemName(500) As String
        Dim vQty(500) As Double
        Dim vDiff(500) As Double
        Dim vInspectQTY(500) As Double
        Dim vCountItem As Double
        Dim vSumItem(500) As Double
        Dim i As Double
        Dim j As Double
        Dim n As Double
        Dim vItemCode(500) As String
        Dim vShelfCode(500) As String
        Dim vWHCode(500) As String
        Dim vInSpectDesc(500) As String

        On Error GoTo ErrDescription

        vQuery = "select count(itemcode) as countitem from (select distinct itemcode,whcode,stockshelf from npmaster.dbo.TB_IV_StkInspect_Logs where docno = '" & vGetDocNo & "') as a "
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
        If ds.Tables(0).Rows.Count > 0 Then
            vCountItem = ds.Tables(0).Rows(0)("countitem").ToString
        End If


        vQuery = "exec dbo.USP_NP_InsertBCSTKInspect '" & vNewDocNo & "','" & vPersonName & "' "
        Dim vService1 As New WebReference.WebServiceCalc
        Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

        vQuery = "exec dbo.USP_NP_SelectItemInspect '" & vGetDocNo & "' "
        Dim vService2 As New WebReference.WebServiceCalc
        Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)
        If ds2.Tables(0).Rows.Count > 0 Then
            j = 0
            For n = 0 To ds2.Tables(0).Rows.Count - 1
                j = j + 1
                vItemCode(j) = ds2.Tables(0).Rows(n)("itemcode").ToString
                vWHCode(j) = ds2.Tables(0).Rows(n)("whcode").ToString
                vShelfCode(j) = ds2.Tables(0).Rows(n)("stockshelf").ToString
            Next
        End If

        For i = 1 To vCountItem
            vQuery = "exec dbo.USP_NP_SelectItemDetailsInspect '" & vGetDocNo & "' , '" & vItemCode(i) & "' ,'" & vShelfCode(i) & "' ,'" & vWHCode(i) & "' "
            Dim vService3 As New WebReference.WebServiceCalc
            Dim ds3 As DataSet = vService3.vGetQueryAnlyzer(vQuery)
            If ds3.Tables(0).Rows.Count > 0 Then
                vItemName(i) = ds3.Tables(0).Rows(0)("itemname").ToString
                vUnitCode(i) = ds3.Tables(0).Rows(0)("unitcode").ToString
                vInSpectDesc(i) = ds3.Tables(0).Rows(0)("reasoncode").ToString
            End If

            vQuery = "exec dbo.USP_NP_SelectSumItemQtyInspect '" & vGetDocNo & "','" & vItemCode(i) & "','" & vWHCode(i) & "','" & vShelfCode(i) & "' "
            Dim vService4 As New WebReference.WebServiceCalc
            Dim ds4 As DataSet = vService4.vGetQueryAnlyzer(vQuery)
            If ds4.Tables(0).Rows.Count > 0 Then
                vSumItem(i) = ds4.Tables(0).Rows(0)("qty").ToString
            Else
                vSumItem(i) = 0
            End If


            vQuery = "exec dbo.USP_NP_SelectItemQtySTKLocation '" & vItemCode(i) & "','" & vWHCode(i) & "','" & vShelfCode(i) & "' "
            Dim vService6 As New WebReference.WebServiceCalc
            Dim ds6 As DataSet = vService6.vGetQueryAnlyzer(vQuery)
            If ds6.Tables(0).Rows.Count > 0 Then
                vInspectQTY(i) = ds6.Tables(0).Rows(0)("qty").ToString
            End If

            vDiff(i) = vSumItem(i) - vInspectQTY(i)
        Next i

        Dim vLine As Integer
        For i = 1 To vCountItem
            vLine = i - 1
            vQuery = "exec dbo.USP_NP_InsertBCInspectSub '" & vNewDocNo & "','" & vItemCode(i) & "','" & vUnitCode(i) & "','" & vWHCode(i) & "','" & vShelfCode(i) & "'," & vInspectQTY(i) & "," & vSumItem(i) & "," & vDiff(i) & ",'" & vInSpectDesc(i) & "'," & vLine & " "
            Dim vService7 As New WebReference.WebServiceCalc
            Dim ds7 As Integer = vService7.vExecuteQuery(vQuery)
        Next i
        vQuery = "Update npmaster.dbo.TB_IV_StkInspect_Logs set Inspectno = '" & vNewDocNo & "' where docno = '" & vGetDocNo & "' "
        Dim vService8 As New WebReference.WebServiceCalc
        Dim ds8 As Integer = vService8.vExecuteQuery(vQuery)

        Call PrintData()
        MsgBox("ได้เอกสารตรวจนับเลขที่ " & vNewDocNo & " โปรแกรมได้ส่งเอกสารไปพิมพ์ให้เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description)
            Exit Sub
        End If
    End Sub

    Public Sub UpdateAddItemInspect()

        Dim vItem(500) As String
        Dim vUnitCode(500) As String
        Dim vShelf As String
        Dim vItemName(500) As String
        Dim vQty(500) As Double
        Dim vDiff(500) As Double
        Dim vInspectQTY(500) As Double
        Dim vCountItem As Double
        Dim vSumItem(500) As Double
        Dim i As Double
        Dim j As Double
        Dim n As Double
        Dim vItemCode(500) As String
        Dim vShelfCode(500) As String
        Dim vWHCode(500) As String
        Dim vInSpectDesc(500) As String

        On Error GoTo ErrDescription

        vDocno = Me.TBDocNo.Text
        vStkNo = Me.TBStockNo.Text
        vQuery = "select count(itemcode) as countitem from (select distinct itemcode,whcode,stockshelf from npmaster.dbo.TB_IV_StkInspect_Logs where docno = '" & vStkNo & "') as a "
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
        If ds.Tables(0).Rows.Count > 0 Then
            vCountItem = ds.Tables(0).Rows(0)("countitem").ToString
        End If

        vQuery = "exec dbo.USP_NP_UpdateBCSTKInspect '" & vDocno & "','" & vPersonName & "' "
        Dim vService1 As New WebReference.WebServiceCalc
        Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

        vQuery = "exec dbo.USP_NP_SelectItemInspect '" & vStkNo & "' "
        Dim vService2 As New WebReference.WebServiceCalc
        Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)
        If ds2.Tables(0).Rows.Count > 0 Then
            j = 0
            For n = 0 To ds2.Tables(0).Rows.Count - 1
                j = j + 1
                vItemCode(j) = ds2.Tables(0).Rows(n)("itemcode").ToString
                vWHCode(j) = ds2.Tables(0).Rows(n)("whcode").ToString
                vShelfCode(j) = ds2.Tables(0).Rows(n)("stockshelf").ToString
            Next
        End If

        For i = 1 To vCountItem
            vQuery = "exec dbo.USP_NP_SelectItemDetailsInspect '" & vStkNo & "' , '" & vItemCode(i) & "' ,'" & vShelfCode(i) & "' ,'" & vWHCode(i) & "' "
            Dim vService3 As New WebReference.WebServiceCalc
            Dim ds3 As DataSet = vService3.vGetQueryAnlyzer(vQuery)
            If ds3.Tables(0).Rows.Count > 0 Then
                vItemName(i) = ds3.Tables(0).Rows(0)("itemname").ToString
                vUnitCode(i) = ds3.Tables(0).Rows(0)("unitcode").ToString
                vInSpectDesc(i) = ds3.Tables(0).Rows(0)("reasoncode").ToString
            End If

            vQuery = "exec dbo.USP_NP_SelectSumItemQtyInspect '" & vStkNo & "','" & vItemCode(i) & "','" & vWHCode(i) & "','" & vShelfCode(i) & "' "
            Dim vService4 As New WebReference.WebServiceCalc
            Dim ds4 As DataSet = vService4.vGetQueryAnlyzer(vQuery)
            If ds4.Tables(0).Rows.Count > 0 Then
                vSumItem(i) = ds4.Tables(0).Rows(0)("qty").ToString
            Else
                vSumItem(i) = 0
            End If

            vQuery = "exec dbo.USP_NP_SelectItemQtySTKLocation '" & vItemCode(i) & "','" & vWHCode(i) & "','" & vShelfCode(i) & "' "
            Dim vService6 As New WebReference.WebServiceCalc
            Dim ds6 As DataSet = vService6.vGetQueryAnlyzer(vQuery)
            If ds6.Tables(0).Rows.Count > 0 Then
                vInspectQTY(i) = ds6.Tables(0).Rows(0)("qty").ToString
            End If

            vDiff(i) = vSumItem(i) - vInspectQTY(i)
        Next i

        vQuery = "delete dbo.BCStkInspectSub where docno = '" & vDocno & "'"
        Dim vService9 As New WebReference.WebServiceCalc
        Dim ds9 As Integer = vService9.vExecuteQuery(vQuery)

        Dim vLine As Integer
        For i = 1 To vCountItem
            vLine = i - 1
            vQuery = "exec dbo.USP_NP_InsertBCInspectSub '" & vDocno & "','" & vItemCode(i) & "','" & vUnitCode(i) & "','" & vWHCode(i) & "','" & vShelfCode(i) & "'," & vInspectQTY(i) & "," & vSumItem(i) & "," & vDiff(i) & ",'" & vInSpectDesc(i) & "'," & vLine & " "
            Dim vService7 As New WebReference.WebServiceCalc
            Dim ds7 As Integer = vService7.vExecuteQuery(vQuery)
        Next i

        vQuery = "Update npmaster.dbo.TB_IV_StkInspect_Logs set Inspectno = '" & vDocno & "' where docno = '" & vStkNo & "' "
        Dim vService8 As New WebReference.WebServiceCalc
        Dim ds8 As Integer = vService8.vExecuteQuery(vQuery)

        MsgBox("ปรับปรุงเอกสารตรวจนับเลขที่ " & vDocno & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description)
            Exit Sub
        End If
    End Sub
    Private Sub BTNClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClear.Click
        MsgBox("ครั้งต่อไปให้กดปุ่มสีส้ม+ปุ่มเลข 7", MsgBoxStyle.Information, "Send Information Message")
        vDocno = ""
        vNewDocNo = ""
        vGetDocNo = ""
        vStkNo = ""
        Me.TBDocNo.Text = ""
        Me.TBDocDate.Text = ""
        Me.TBBarCode.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.ListViewStock.Items.Clear()
        Me.CMBReason.Focus()
    End Sub

    Public Sub ClearData()
        vDocno = ""
        vNewDocNo = ""
        vGetDocNo = ""
        vStkNo = ""
        Me.TBDocNo.Text = ""
        Me.TBDocDate.Text = ""
        Me.TBBarCode.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.ListViewStock.Items.Clear()
        Me.CMBReason.Focus()
    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        Dim i As Integer
        Dim n As Integer

        MsgBox("ครั้งต่อไปให้กดปุ่มสีส้ม+ปุ่มเลข 1", MsgBoxStyle.Information, "Send Information Message")

        Me.ListViewSearch.Items.Clear()
        vQuery = "exec dbo.USP_NP_SearchInspectNo ''"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
        If ds.Tables(0).Rows.Count > 0 Then
            n = 1
            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim listItem As New ListViewItem(n)
                listItem.SubItems.Add(ds.Tables(0).Rows(i)("docno").ToString)
                listItem.SubItems.Add(ds.Tables(0).Rows(i)("creatorcode").ToString)
                Me.ListViewSearch.Items.Add(listItem)

                n = n + 1
            Next

            Me.PNSearch.Visible = True
            Me.PNSearch.BringToFront()
            Me.TBSearch.Text = ""
            Me.ListViewSearch.Focus()
            Me.ListViewSearch.Items(0).Selected = True
            Me.ListViewSearch.Items(0).Focused = True
        Else
            MsgBox("ไม่มีรายการเอกสารตรวจนับในระบบ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNSearch.Focus()
        End If

    End Sub

    Public Sub SearchData()
        Dim i As Integer
        Dim n As Integer

        Me.ListViewSearch.Items.Clear()
        vQuery = "exec dbo.USP_NP_SearchInspectNo ''"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
        If ds.Tables(0).Rows.Count > 0 Then
            n = 1
            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim listItem As New ListViewItem(n)
                listItem.SubItems.Add(ds.Tables(0).Rows(i)("docno").ToString)
                listItem.SubItems.Add(ds.Tables(0).Rows(i)("creatorcode").ToString)
                Me.ListViewSearch.Items.Add(listItem)

                n = n + 1
            Next

            Me.PNSearch.Visible = True
            Me.PNSearch.BringToFront()
            Me.TBSearch.Text = ""
            Me.ListViewSearch.Focus()
            Me.ListViewSearch.Items(0).Selected = True
            Me.ListViewSearch.Items(0).Focused = True
        Else
            MsgBox("ไม่มีรายการเอกสารตรวจนับในระบบ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.BTNSearch.Focus()
        End If

    End Sub

    Private Sub BTNSearchClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchClose.Click
        Me.PNSearch.Visible = False
        If Me.ListViewItem.Items.Count > 0 Then
            Me.ListViewItem.Focus()
            Me.ListViewItem.Items(0).Selected = True
            Me.ListViewItem.Items(0).Focused = True
        Else
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If
    End Sub

    Public Sub CloseSearch()
        Me.PNSearch.Visible = False
        If Me.ListViewItem.Items.Count > 0 Then
            Me.ListViewItem.Focus()
            Me.ListViewItem.Items(0).Selected = True
            Me.ListViewItem.Items(0).Focused = True
        Else
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If
    End Sub

    Private Sub TBSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearch.KeyDown
        Dim i As Integer
        Dim n As Integer
        Dim vSearch As String

        If e.KeyCode = Keys.Enter Then
            If Me.TBSearch.Text <> "" Then
                Me.ListViewSearch.Items.Clear()
                vSearch = Me.TBSearch.Text
                vQuery = "exec dbo.USP_NP_SearchInspectNo '" & vSearch & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
                If ds.Tables(0).Rows.Count > 0 Then
                    n = 1
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        Dim listItem As New ListViewItem(n)
                        listItem.SubItems.Add(ds.Tables(0).Rows(i)("docno").ToString)
                        listItem.SubItems.Add(ds.Tables(0).Rows(i)("creatorcode").ToString)
                        Me.ListViewSearch.Items.Add(listItem)

                        n = n + 1
                    Next

                    Me.PNSearch.Visible = True
                    Me.PNSearch.BringToFront()
                    Me.TBSearch.Text = ""
                    Me.ListViewSearch.Focus()
                    Me.ListViewSearch.Items(0).Selected = True
                    Me.ListViewSearch.Items(0).Focused = True
                Else
                    MsgBox("ไม่มีรายการเอกสารตรวจนับที่ค้นหาในระบบ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.BTNSearch.Focus()
                End If
            End If
        End If


        If e.KeyCode = Keys.Down Then
            If Me.ListViewSearch.Items.Count > 0 Then
                Me.ListViewSearch.Focus()
                Me.ListViewSearch.Items(0).Selected = True
                Me.ListViewSearch.Items(0).Focused = True
            Else
                Me.TBSearch.Focus()
                Me.TBSearch.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseSearch()
        End If
    End Sub

    Private Sub TBSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearch.TextChanged

    End Sub

    Private Sub ListViewSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearch.KeyDown
        Dim i As Integer
        Dim n As Integer
        Dim vDocno As String
        Dim vInspectQty As Double
        Dim vSTKQty As Double

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewSearch.Items.Count > 0 Then
                Me.ListViewItem.Items.Clear()
                vDocno = Me.ListViewSearch.Items(Me.ListViewSearch.FocusedItem.Index).SubItems(1).Text
                vQuery = "exec dbo.USP_NP_SearchInspectNoDetails '" & vDocno & "' "
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
                If ds.Tables(0).Rows.Count > 0 Then
                    Me.TBDocNo.Text = ds.Tables(0).Rows(i)("docno").ToString
                    Me.TBDocDate.Text = ds.Tables(0).Rows(i)("docdate").ToString
                    Me.TBStockNo.Text = ds.Tables(0).Rows(i)("stkno").ToString
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        n = n + 1
                        Dim listItem As New ListViewItem(n)
                        listItem.SubItems.Add(ds.Tables(0).Rows(i)("itemname").ToString)
                        vInspectQty = ds.Tables(0).Rows(i)("inspectqty").ToString
                        listItem.SubItems.Add(Format(vInspectQty, "##,##0.00"))
                        listItem.SubItems.Add(ds.Tables(0).Rows(i)("unitcode").ToString)
                        listItem.SubItems.Add(ds.Tables(0).Rows(i)("itemcode").ToString)
                        listItem.SubItems.Add(ds.Tables(0).Rows(i)("whcode").ToString)
                        listItem.SubItems.Add(ds.Tables(0).Rows(i)("shelfcode").ToString)
                        listItem.SubItems.Add(ds.Tables(0).Rows(i)("shelfid").ToString)
                        listItem.SubItems.Add(1)
                        vSTKQty = ds.Tables(0).Rows(i)("stkqty").ToString
                        listItem.SubItems.Add(Format(vSTKQty, "##,##0.00"))
                        listItem.SubItems.Add(ds.Tables(0).Rows(i)("inspectdesc").ToString)
                        Me.ListViewItem.Items.Add(listItem)
                    Next

                    Me.PNSearch.Visible = False
                    Me.TBBarCode.Focus()
                    Me.TBBarCode.SelectAll()
                End If
            End If
        End If

        Dim x As Integer

        If e.KeyCode = Keys.Down Then
            x = Me.ListViewSearch.FocusedItem.Index
            If Me.ListViewSearch.Items.Count = x Then
                Me.BTNSearchOK.Focus()
            End If
        End If

        If e.KeyCode = Keys.Up Then
            x = Me.ListViewSearch.FocusedItem.Index
            If x = 0 Then
                Me.TBSearch.Focus()
                Me.TBSearch.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseSearch()
        End If
    End Sub


    Private Sub BTNSearchOK_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSearchOK.KeyDown
        If e.KeyCode = Keys.Up Then
            If Me.ListViewSearch.Items.Count > 0 Then
                Me.ListViewSearch.Focus()
                Me.ListViewSearch.Items(0).Selected = True
                Me.ListViewSearch.Items(0).Focused = True
            End If
        Else
            Me.TBSearch.Focus()
            Me.TBSearch.SelectAll()
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseSearch()
        End If
    End Sub

    Private Sub BTNSearchClose_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSearchClose.KeyDown
        If e.KeyCode = Keys.Up Then
            If Me.ListViewSearch.Items.Count > 0 Then
                Me.ListViewSearch.Focus()
                Me.ListViewSearch.Items(0).Selected = True
                Me.ListViewSearch.Items(0).Focused = True
            End If
        Else
            Me.TBSearch.Focus()
            Me.TBSearch.SelectAll()
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseSearch()
        End If
    End Sub

    Private Sub ListViewSearch_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewSearch.SelectedIndexChanged

    End Sub

    Private Sub BTNClear_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClear.KeyDown
        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call ClearData()
        End If

        If e.KeyCode = 114 Then
            Call SearchData()
        End If

        If e.KeyCode = 115 Then
            Call PrintData()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If
    End Sub

    Private Sub BTNPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPrint.Click
        MsgBox("ครั้งต่อไปให้กดปุ่มสีส้ม+ปุ่มเลข 9", MsgBoxStyle.Information, "Send Information Message")
        Call PrintData()
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        MsgBox("ครั้งต่อไปให้กดปุ่ม ESC", MsgBoxStyle.Information, "Send Information Message")
        Call ClearData()
        FrmMobileApp.Show()
        Me.Hide()
    End Sub

    Public Sub ExitProgram()
        Call ClearData()
        FrmMobileApp.Show()
        Me.Hide()
    End Sub

    Private Sub TBDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBDocNo.KeyDown
        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call ClearData()
        End If

        If e.KeyCode = 114 Then
            Call SearchData()
        End If

        If e.KeyCode = 115 Then
            Call PrintData()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If
    End Sub

    Private Sub TBDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBDocNo.TextChanged

    End Sub

    Public Sub PrintData()
        Dim vDocNumber As String

        If vNewDocNo <> "" Then
            vDocNumber = vNewDocNo
        End If

        If vNewDocNo = "" Then
            vDocno = Me.TBDocNo.Text
        End If

        If vDocno <> "" Then
            vDocNumber = vDocno
        End If

        If vDocNumber = "" Then
            MsgBox("ไม่สามารถพิมพ์เอกสารตรวจนับได้ ยังไม่มีเลขที่เอกสาร กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
            Exit Sub
        End If

        vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto1 10,'" & vDocNumber & "','','" & vPersonName & "'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As Integer = vService.vExecuteQuery(vQuery)

        MsgBox("ส่งเอกสารพิมพ์เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")
    End Sub

    Private Sub TBDocDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBDocDate.KeyDown
        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call ClearData()
        End If

        If e.KeyCode = 114 Then
            Call SearchData()
        End If

        If e.KeyCode = 115 Then
            Call PrintData()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If
    End Sub

    Private Sub TBDocDate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBDocDate.TextChanged

    End Sub

    Private Sub TBStockNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBStockNo.KeyDown
        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call ClearData()
        End If

        If e.KeyCode = 114 Then
            Call SearchData()
        End If

        If e.KeyCode = 115 Then
            Call PrintData()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If
    End Sub

    Private Sub ListViewItem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewItem.SelectedIndexChanged

    End Sub

    Private Sub BTNSave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSave.KeyDown
        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call ClearData()
        End If

        If e.KeyCode = 114 Then
            Call SearchData()
        End If

        If e.KeyCode = 115 Then
            Call PrintData()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If
    End Sub

    Private Sub BTNSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSearch.KeyDown
        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call ClearData()
        End If

        If e.KeyCode = 114 Then
            Call SearchData()
        End If

        If e.KeyCode = 115 Then
            Call PrintData()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If
    End Sub

    Private Sub BTNPrint_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNPrint.KeyDown
        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call ClearData()
        End If

        If e.KeyCode = 114 Then
            Call SearchData()
        End If

        If e.KeyCode = 115 Then
            Call PrintData()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If
    End Sub

    Private Sub BTNExit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNExit.KeyDown
        If e.KeyCode = 16 Then
            Call SaveData()
        End If

        If e.KeyCode = 33 Then
            Call ClearData()
        End If

        If e.KeyCode = 114 Then
            Call SearchData()
        End If

        If e.KeyCode = 115 Then
            Call PrintData()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If
    End Sub

    Private Sub BTNDelete_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNDelete.KeyDown
        If e.KeyCode = 114 Then
            Me.PNAddStock.Visible = True
            Me.PNAddStock.BringToFront()
            Call vGetShelf()
            Me.CMBShelf.Focus()
        End If

        If e.KeyCode = 34 Then
            Call AddItem()
        End If

        If e.KeyCode = Keys.Escape Then
            Call CloseAddItem()
        End If
    End Sub

    Private Sub TBItemShelf_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBItemShelf.KeyDown

        If e.KeyCode = Keys.Escape Then
            Call CloseKeyQty()
        End If
    End Sub

    Public Sub BeforeSave()
        Me.TBDocNo.Enabled = False
        Me.TBBarCode.Enabled = False
        Me.CMBReason.Enabled = False
        Me.ListViewItem.Enabled = False
        Me.BTNClear.Enabled = False
        Me.BTNSave.Enabled = False
        Me.BTNSearch.Enabled = False
        Me.BTNPrint.Enabled = False
        Me.BTNExit.Enabled = False
    End Sub

    Public Sub AfterSave()
        Me.TBDocNo.Enabled = True
        Me.TBBarCode.Enabled = True
        Me.CMBReason.Enabled = True
        Me.ListViewItem.Enabled = True
        Me.BTNClear.Enabled = True
        Me.BTNSave.Enabled = True
        Me.BTNSearch.Enabled = True
        Me.BTNPrint.Enabled = True
        Me.BTNExit.Enabled = True
    End Sub
End Class