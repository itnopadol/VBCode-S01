Imports System.data
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms

Public Class FrmAddItemShelf
    Dim vQuery As String
    Private Sub TBShelf_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBShelf.KeyDown
        Dim vShelfCode As String

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            vShelfCode = Me.TBShelf.Text

            vQuery = "exec dbo.USP_NP_CheckWHShelfCode '" & vShelfCode & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox("ไม่มีรหัสชั้นเก็บ ที่ได้กรอกมา กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBShelf.Text = ""
                Me.TBShelf.Focus()
                Me.TBShelf.SelectAll()
            Else
                Me.TBShelf.Text = UCase(Me.TBShelf.Text)
                Me.TBZone.Text = ds.Tables(0).Rows(0)("fiscalshelf").ToString
                Me.TBShelf.Enabled = False
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
            End If
        End If

        If e.KeyCode = 40 Then
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call SaveItemShelf()
        End If

        If e.KeyCode = 114 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = 115 Then
            Call CheckItemShelf()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
        Dim n As Integer
        Dim vBarCode As String
        Dim i As Integer
        Dim vItemCode As String
        Dim vItemList As String

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then

            If Me.TBBarCode.Text <> "" Then
                vBarCode = Me.TBBarCode.Text
            Else
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                Exit Sub
            End If

            If Me.TBShelf.Text = "" Then
                MsgBox("ก่อนเพิ่มรหัสสินค้าเข้าที่เก็บ ต้องกรอกรหัสที่เก็บให้ถูกต้องก่อน กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBShelf.Focus()
                Me.TBShelf.SelectAll()
                Exit Sub
            End If

            n = Me.ListViewSelectItem.Items.Count
            n = n + 1

            vQuery = "exec dbo.USP_NP_SearchBarcodeDetails '" & vBarCode & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
            If ds.Tables(0).Rows.Count > 0 Then

                vItemCode = ds.Tables(0).Rows(0)("itemcode").ToString

                If Me.ListViewSelectItem.Items.Count > 0 Then
                    For i = 0 To Me.ListViewSelectItem.Items.Count - 1

                        vItemList = Me.ListViewSelectItem.Items(i).SubItems(2).Text
                        If vItemCode = vItemList Then
                            Me.TBBarCode.Text = ""
                            Me.TBBarCode.Focus()
                            Exit Sub
                        End If
                    Next
                End If

                Dim listItem As New ListViewItem(n)
                listItem.SubItems.Add(ds.Tables(0).Rows(0)("itemname").ToString)
                listItem.SubItems.Add(ds.Tables(0).Rows(0)("itemcode").ToString)
                listItem.SubItems.Add(Me.TBWHCode.Text)
                listItem.SubItems.Add(Me.TBShelf.Text)
                listItem.SubItems.Add(ds.Tables(0).Rows(0)("barcode").ToString)
                listItem.SubItems.Add(ds.Tables(0).Rows(0)("unitcode").ToString)
                listItem.SubItems.Add(Me.TBZone.Text)
                Me.ListViewSelectItem.Items.Add(listItem)

                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
            Else
                MsgBox("ไม่มีข้อมูลของบาร์โค้ดที่ค้นหา กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                Exit Sub
            End If
        End If

        If e.KeyCode = 40 Then
            If Me.ListViewSelectItem.Items.Count > 0 Then
                Me.ListViewSelectItem.Focus()
                Me.ListViewSelectItem.Items(0).Selected = True
                Me.ListViewSelectItem.Items(0).Focused = True
            Else
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
            End If
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call SaveItemShelf()
        End If

        If e.KeyCode = 114 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = 115 Then
            Call CheckItemShelf()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClear.Click
        Call ClearScreen()
    End Sub

    Public Sub ClearScreen()
        On Error Resume Next

        Me.TBShelf.Enabled = True
        Me.TBShelf.Text = ""
        Me.TBZone.Text = ""
        Me.TBBarCode.Text = ""
        Me.ListViewSelectItem.Items.Clear()
        Me.TBShelf.Focus()
        Me.TBShelf.SelectAll()
    End Sub

    Private Sub TBWHCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBWHCode.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 40 Then
            Me.TBShelf.Focus()
            Me.TBShelf.SelectAll()
        End If

        If e.KeyCode = Keys.Enter Then
            Me.TBShelf.Focus()
            Me.TBShelf.SelectAll()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call SaveItemShelf()
        End If

        If e.KeyCode = 114 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = 115 Then
            Call CheckItemShelf()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNMenu.Click
        On Error Resume Next

        Call ClearScreen()
        FrmMobileApp.Show()
        Me.Hide()
    End Sub

    Private Sub ListViewSelectItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSelectItem.KeyDown
        Dim vIndex As Integer
        Dim vAnswer As Integer
        Dim n As Integer
        Dim i As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Back Then
            If Me.ListViewSelectItem.Items.Count > 0 Then
                vIndex = Me.ListViewSelectItem.FocusedItem.Index
                vAnswer = MsgBox("คุณต้องการลบรายการสินค้านี้ออกการตารางใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message ?")
                If vAnswer = 6 Then
                    Me.ListViewSelectItem.Items.RemoveAt(vIndex)

                    For n = 0 To Me.ListViewSelectItem.Items.Count - 1
                        i = i + 1
                        Me.ListViewSelectItem.Items(n).SubItems(0).Text = i
                    Next

                Else
                    Me.ListViewSelectItem.Focus()
                    Me.ListViewSelectItem.Items(vIndex).Selected = True
                    Me.ListViewSelectItem.Items(vIndex).Focused = True
                    Exit Sub
                End If
            End If
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call SaveItemShelf()
        End If

        If e.KeyCode = 114 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = 115 Then
            Call CheckItemShelf()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vZoneCode As String
        Dim n As Integer

        On Error GoTo ErrDescription

        If Me.ListViewSelectItem.Items.Count > 0 Then
            Call BeforeSave()
            For n = 0 To Me.ListViewSelectItem.Items.Count - 1
                vItemCode = Me.ListViewSelectItem.Items(n).SubItems(2).Text
                vBarCode = UCase(Me.ListViewSelectItem.Items(n).SubItems(5).Text)
                vItemName = Me.ListViewSelectItem.Items(n).SubItems(1).Text
                vUnitCode = Me.ListViewSelectItem.Items(n).SubItems(6).Text
                vWHCode = UCase(Me.ListViewSelectItem.Items(n).SubItems(3).Text)
                vShelfCode = UCase(Me.ListViewSelectItem.Items(n).SubItems(4).Text)
                vZoneCode = Me.ListViewSelectItem.Items(n).SubItems(7).Text

                vQuery = "exec dbo.usp_np_insertscanitemshelfcode '" & vItemCode & "','" & vBarCode & "','" & vItemName & "','" & vUnitCode & "','" & vWHCode & "','" & vZoneCode & "','" & vShelfCode & "','" & vPersonName & "','บันทึกที่เก็บสินค้าจาก CN2' "
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

            Next

            MsgBox("บันทึกสินค้าเข้าที่เก็บ เรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")
            Call AfterSave()
            Me.ListViewSelectItem.Items.Clear()
            Me.TBShelf.Text = ""
            Me.TBZone.Text = ""
            Me.TBShelf.Focus()
            Me.TBShelf.SelectAll()

        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BeforeSave()
        On Error Resume Next

        Me.TBShelf.Enabled = False
        Me.ListViewSelectItem.Enabled = False
        Me.BTNCheckShelf.Enabled = False
        Me.BTNCHKExit.Enabled = False
        Me.BTNClear.Enabled = False
        Me.BTNMenu.Enabled = False
        Me.BTNSave.Enabled = False
    End Sub

    Private Sub AfterSave()
        On Error Resume Next

        Me.TBShelf.Enabled = True
        Me.ListViewSelectItem.Enabled = True
        Me.BTNCheckShelf.Enabled = True
        Me.BTNCHKExit.Enabled = True
        Me.BTNClear.Enabled = True
        Me.BTNMenu.Enabled = True
        Me.BTNSave.Enabled = True
    End Sub

    Private Sub SaveItemShelf()
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vZoneCode As String
        Dim n As Integer

        On Error GoTo ErrDescription

        If Me.ListViewSelectItem.Items.Count > 0 Then
            Call BeforeSave()
            For n = 0 To Me.ListViewSelectItem.Items.Count - 1
                vItemCode = Me.ListViewSelectItem.Items(n).SubItems(2).Text
                vBarCode = UCase(Me.ListViewSelectItem.Items(n).SubItems(5).Text)
                vItemName = Me.ListViewSelectItem.Items(n).SubItems(1).Text
                vUnitCode = Me.ListViewSelectItem.Items(n).SubItems(6).Text
                vWHCode = UCase(Me.ListViewSelectItem.Items(n).SubItems(3).Text)
                vShelfCode = UCase(Me.ListViewSelectItem.Items(n).SubItems(4).Text)
                vZoneCode = Me.ListViewSelectItem.Items(n).SubItems(7).Text

                vQuery = "exec dbo.usp_np_insertscanitemshelfcode '" & vItemCode & "','" & vBarCode & "','" & vItemName & "','" & vUnitCode & "','" & vWHCode & "','" & vZoneCode & "','" & vShelfCode & "','" & vPersonName & "','บันทึกที่เก็บสินค้าจาก CN2' "
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

            Next

            MsgBox("บันทึกสินค้าเข้าที่เก็บ เรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")
            Call AfterSave()
            Me.ListViewSelectItem.Items.Clear()
            Me.TBShelf.Text = ""
            Me.TBZone.Text = ""
            Me.TBShelf.Focus()
            Me.TBShelf.SelectAll()

        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Dim n As Integer
        'Dim vBarCode As String
        'Dim i As Integer
        'Dim vItemCode As String
        'Dim vItemList As String

        'n = Me.ListViewSelectItem.Items.Count
        'n = n + 1
        'i = Me.ListViewSelectItem.Items.Count

        'vBarCode = "2120250"

        'vQuery = "exec dbo.USP_NP_SearchBarcodeDetails '" & vBarCode & "'"
        'Dim vService As New WebReference.WebServiceCalc
        'Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
        'If ds.Tables(0).Rows.Count > 0 Then

        '    vItemCode = ds.Tables(0).Rows(0)("itemcode").ToString

        '    Dim listItem As New ListViewItem(n)
        '    listItem.SubItems.Add(ds.Tables(0).Rows(0)("itemname").ToString)
        '    listItem.SubItems.Add(ds.Tables(0).Rows(0)("itemcode").ToString)
        '    listItem.SubItems.Add("S02")
        '    listItem.SubItems.Add("01A01")
        '    listItem.SubItems.Add(ds.Tables(0).Rows(0)("barcode").ToString)
        '    listItem.SubItems.Add(ds.Tables(0).Rows(0)("unitcode").ToString)
        '    listItem.SubItems.Add("S02")
        '    Me.ListViewSelectItem.Items.Add(listItem)
        'End If

        'Me.ListViewSelectItem.Items(i).Selected = True
        'Me.ListViewSelectItem.Items(i).Focused = True

    End Sub

    Private Sub FrmAddItemShelf_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        Me.TBShelf.Text = ""
        Me.TBShelf.Focus()
        Me.TBShelf.SelectAll()
    End Sub

    Private Sub TBCHKShelf_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBCHKShelf.KeyDown
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim n As Integer
        Dim i As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.TBCHKShelf.Text <> "" Then
                vWHCode = Me.TBCHKWHCode.Text
                vShelfCode = Me.TBCHKShelf.Text

                Me.ListViewCHKItem.Items.Clear()
                vQuery = "exec dbo.USP_MB_SearchItemScanRecProduct '" & vWHCode & "','" & vShelfCode & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
                If ds.Tables(0).Rows.Count > 0 Then
                    For n = 0 To ds.Tables(0).Rows.Count - 1
                        i = i + 1
                        Dim listItem As New ListViewItem(i)
                        listItem.SubItems.Add(ds.Tables(0).Rows(n)("itemname").ToString)
                        listItem.SubItems.Add(ds.Tables(0).Rows(n)("itemcode").ToString)
                        listItem.SubItems.Add(ds.Tables(0).Rows(n)("unitcode").ToString)
                        Me.ListViewCHKItem.Items.Add(listItem)
                    Next

                    Me.TBCHKShelf.Enabled = False

                    If Me.ListViewCHKItem.Items.Count > 0 Then
                        Me.ListViewCHKItem.Focus()
                        Me.ListViewCHKItem.Items(0).Selected = True
                        Me.ListViewCHKItem.Items(0).Focused = True
                    Else
                        Me.TBCHKShelf.Focus()
                        Me.TBCHKShelf.SelectAll()
                    End If
                End If
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBCHKShelf.Enabled = True
            Me.TBCHKShelf.Text = ""
            Me.ListViewCHKItem.Items.Clear()
            Me.PNCheckShelf.Visible = False
            Me.TBShelf.Focus()
            Me.TBShelf.SelectAll()
        End If

        If e.KeyCode = 33 Then
            Me.TBCHKShelf.Text = ""
            Me.TBCHKShelf.Enabled = True
            Me.ListViewCHKItem.Items.Clear()
            Me.TBCHKShelf.Focus()
            Me.TBCHKShelf.SelectAll()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNCheckShelf_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNCheckShelf.Click
        On Error Resume Next

        Me.TBShelf.Enabled = True
        Me.TBShelf.Text = ""
        Me.PNCheckShelf.Visible = True
        Me.PNCheckShelf.BringToFront()
        Me.TBCHKShelf.Text = ""
        Me.TBCHKShelf.Focus()
        Me.TBCHKShelf.SelectAll()
    End Sub

    Private Sub CheckItemShelf()
        On Error Resume Next

        Me.TBShelf.Enabled = True
        Me.TBShelf.Text = ""
        Me.PNCheckShelf.Visible = True
        Me.PNCheckShelf.BringToFront()
        Me.TBCHKShelf.Text = ""
        Me.TBCHKShelf.Focus()
        Me.TBCHKShelf.SelectAll()
    End Sub

    Private Sub TBCHKWHCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBCHKWHCode.KeyDown, BTNCHKExit.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.TBCHKShelf.Enabled = True
            Me.TBCHKShelf.Text = ""
            Me.ListViewCHKItem.Items.Clear()
            Me.PNCheckShelf.Visible = False
            Me.TBShelf.Focus()
            Me.TBShelf.SelectAll()
        End If

        If e.KeyCode = 33 Then
            Me.TBCHKShelf.Text = ""
            Me.TBCHKShelf.Enabled = True
            Me.ListViewCHKItem.Items.Clear()
            Me.TBShelf.Focus()
            Me.TBShelf.SelectAll()
        End If
    End Sub

    Private Sub BTNCHKExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCHKExit.Click
        On Error Resume Next

        Me.TBCHKShelf.Enabled = True
        Me.TBCHKShelf.Text = ""
        Me.ListViewCHKItem.Items.Clear()
        Me.PNCheckShelf.Visible = False
        Me.TBShelf.Focus()
        Me.TBShelf.SelectAll()
    End Sub

    Private Sub ListViewCHKItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewCHKItem.KeyDown
        Dim vIndex As Integer
        Dim vAnswer As Integer
        Dim vItemCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim n As Integer
        Dim i As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Me.TBCHKShelf.Enabled = True
            Me.TBCHKShelf.Text = ""
            Me.ListViewCHKItem.Items.Clear()
            Me.PNCheckShelf.Visible = False
            Me.TBShelf.Focus()
            Me.TBShelf.SelectAll()
        End If

        If e.KeyCode = 33 Then
            Me.TBCHKShelf.Text = ""
            Me.TBCHKShelf.Enabled = True
            Me.ListViewCHKItem.Items.Clear()
            Me.TBShelf.Focus()
            Me.TBShelf.SelectAll()
        End If

        If e.KeyCode = Keys.Back Then
            If Me.TBCHKShelf.Text = "" Then
                MsgBox("กรุณากรอก รหัสชั้นเก็บ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBCHKShelf.Focus()
                Me.TBCHKShelf.SelectAll()
                Exit Sub
            End If

            If Me.ListViewCHKItem.Items.Count = 0 Then
                MsgBox("ไม่มีรายการสินค้าที่จะทำการลบ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBCHKShelf.Focus()
                Me.TBCHKShelf.SelectAll()
                Exit Sub
            End If

            vIndex = Me.ListViewCHKItem.FocusedItem.Index
            vAnswer = MsgBox("คุณต้องการลบรายการสินค้านี้ออกการตารางใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message ?")
            If vAnswer = 6 Then
                vItemCode = Me.ListViewCHKItem.Items(vIndex).SubItems(2).Text
                vWHCode = Me.TBCHKWHCode.Text
                vShelfCode = Me.TBCHKShelf.Text

                vQuery = "exec dbo.USP_MB_DeleteRecProductShelfCode '" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "'"
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

                Me.ListViewCHKItem.Items.RemoveAt(vIndex)

                For n = 0 To Me.ListViewCHKItem.Items.Count - 1
                    i = i + 1
                    Me.ListViewCHKItem.Items(n).SubItems(0).Text = i
                Next

            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBZone_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBZone.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call SaveItemShelf()
        End If

        If e.KeyCode = 114 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = 115 Then
            Call CheckItemShelf()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub


    Private Sub BTNMenu_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNMenu.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call SaveItemShelf()
        End If

        If e.KeyCode = 114 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = 115 Then
            Call CheckItemShelf()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNClear_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClear.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call SaveItemShelf()
        End If

        If e.KeyCode = 114 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = 115 Then
            Call CheckItemShelf()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSave.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call SaveItemShelf()
        End If

        If e.KeyCode = 114 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = 115 Then
            Call CheckItemShelf()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNCheckShelf_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCheckShelf.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call SaveItemShelf()
        End If

        If e.KeyCode = 114 Then
            Call ClearScreen()
            FrmMobileApp.Show()
            Me.Hide()
        End If

        If e.KeyCode = 115 Then
            Call CheckItemShelf()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNCHKClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCHKClear.Click
        On Error Resume Next

        Me.TBCHKShelf.Text = ""
        Me.TBCHKShelf.Enabled = True
        Me.ListViewCHKItem.Items.Clear()
        Me.TBShelf.Focus()
        Me.TBShelf.SelectAll()
    End Sub

    Private Sub BTNCHKClear_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCHKClear.KeyDown
        On Error Resume Next

        If e.KeyCode = 33 Then
            Me.TBCHKShelf.Text = ""
            Me.TBCHKShelf.Enabled = True
            Me.ListViewCHKItem.Items.Clear()
            Me.TBShelf.Focus()
            Me.TBShelf.SelectAll()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBCHKShelf.Enabled = True
            Me.TBCHKShelf.Text = ""
            Me.ListViewCHKItem.Items.Clear()
            Me.PNCheckShelf.Visible = False
            Me.TBShelf.Focus()
            Me.TBShelf.SelectAll()
        End If
    End Sub
End Class