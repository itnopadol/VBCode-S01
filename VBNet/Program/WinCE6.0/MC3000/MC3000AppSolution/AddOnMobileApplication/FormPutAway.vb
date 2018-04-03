Imports System.IO
Imports Symbol
Imports Symbol.Barcode
Imports Symbol.Barcode.Reader
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.ComponentModel
Imports System.Windows.Forms
Imports vb6 = Microsoft.VisualBasic
Public Class FormPutAway
    Dim vQuery As String
    Dim vMemShelf As String

    Private Sub TBShelfID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBShelfID.KeyDown
        Dim vShelfID As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vIndex As Integer
        Dim i As Integer

        If e.KeyCode = Keys.Enter Then
            i = Me.TBIDNumber.Text
            vIndex = i - 1
            vShelfID = Me.TBShelfID.Text

            vQuery = "exec dbo.USP_NP_SearchShelfMasterDetails '" & vShelfID & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vWHCode = pds.Tables(0).Rows(0)("whcode").ToString
                vShelfCode = pds.Tables(0).Rows(0)("fiscalshelf").ToString

                Me.ListViewItem.Items(vIndex).SubItems(1).Text = vShelfID
                Me.ListViewItem.Items(vIndex).SubItems(7).Text = vWHCode
                Me.ListViewItem.Items(vIndex).SubItems(8).Text = vShelfCode

                Me.TBIDNumber.Text = ""
                Me.TBItemCode.Text = ""
                Me.TBItemName.Text = ""
                Me.TBUnitCode.Text = ""
                Me.TBWHCode.Text = ""
                Me.TBShelfCode.Text = ""
                Me.TBShelfID.Text = ""
                Me.TBIDNumber.Focus()

            Else
                Me.TBWHCode.Text = ""
                Me.TBShelfCode.Text = ""
                MsgBox(vShelfID & "  " & "This shelf not exist", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBShelfID.Text = ""
                Me.TBShelfID.Focus()
                Exit Sub
            End If
        End If


        If e.KeyCode = Keys.Escape Then
            Me.TBShelfID.Text = ""
            Me.TBShelfID.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBIDNumber.Focus()
        End If

        If e.KeyCode = Keys.Down And Me.ListViewItem.Items.Count > 0 Then
            Me.ListViewItem.Focus()
            Me.ListViewItem.Items(0).Focused = True
            Me.ListViewItem.Items(0).Selected = True
        End If

        On Error Resume Next

        If e.KeyCode = Keys.Up And IsDBNull(Me.ListViewItem.FocusedItem.Index) = 0 Then
            Me.TBShelfID.Focus()
            Me.TBShelfID.SelectAll()
        End If
    End Sub

    Private Sub TBShelfID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBShelfID.TextChanged
        Dim vShelfID As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vIndex As Integer
        Dim i As Integer

        If vb6.InStr(Me.TBShelfID.Text, "@") > 0 Then
            i = Me.TBIDNumber.Text

            vIndex = i - 1
            vShelfID = vb6.Left(Me.TBShelfID.Text, vb6.Len(Me.TBShelfID.Text) - 1)

            Me.TBShelfID.Text = vShelfID

            vQuery = "exec dbo.USP_NP_SearchShelfMasterDetails '" & vShelfID & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vWHCode = pds.Tables(0).Rows(0)("whcode").ToString
                vShelfCode = pds.Tables(0).Rows(0)("fiscalshelf").ToString

                Me.ListViewItem.Items(vIndex).SubItems(1).Text = vShelfID
                Me.ListViewItem.Items(vIndex).SubItems(7).Text = vWHCode
                Me.ListViewItem.Items(vIndex).SubItems(8).Text = vShelfCode

                Me.TBIDNumber.Text = ""
                Me.TBItemCode.Text = ""
                Me.TBItemName.Text = ""
                Me.TBUnitCode.Text = ""
                Me.TBWHCode.Text = ""
                Me.TBShelfCode.Text = ""
                Me.TBShelfID.Text = ""
                Me.TBIDNumber.Focus()

            Else
                Me.TBWHCode.Text = ""
                Me.TBShelfCode.Text = ""
                MsgBox(vShelfID & "  " & "This shelf not exist", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBShelfID.Text = ""
                Me.TBShelfID.Focus()
                Exit Sub
            End If
        End If
    End Sub

    Private Sub FormReceiveItemAddShelfID_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TBRVNo.Text = "S01-RV5508-0022"
    End Sub

    Private Sub TBRVNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBRVNo.KeyDown
        Dim i As Integer
        Dim vIndex As Integer
        Dim vDocNo As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vPrice As Double
        Dim vRate As Integer
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vQTY As Double
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vStatus As Integer

        On Error Resume Next

        Me.ListViewItem.Items.Clear()
        If e.KeyCode = Keys.Enter Then
            vDocNo = Me.TBRVNo.Text

            vQuery = "exec dbo.USP_NP_SearchReceiveItemDetails '" & vDocNo & "'"
            Call vGetData(vMemProfit, vQuery)
            If pds.Tables(0).Rows.Count > 0 Then
                For i = 0 To pds.Tables(0).Rows.Count - 1

                    vItemCode = pds.Tables(0).Rows(i)("itemcode").ToString
                    vItemName = pds.Tables(0).Rows(i)("itemname").ToString
                    vUnitCode = pds.Tables(0).Rows(i)("unitcode").ToString
                    vShelfID = pds.Tables(0).Rows(i)("shelfid").ToString
                    vWHCode = pds.Tables(0).Rows(i)("whcode").ToString
                    vShelfCode = pds.Tables(0).Rows(i)("shelfcode").ToString
                    vUnitCode = pds.Tables(0).Rows(i)("unitcode").ToString
                    vQTY = pds.Tables(0).Rows(i)("qty").ToString
                    vStatus = pds.Tables(0).Rows(i)("status").ToString
                    vIndex = i + 1

                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(vShelfID)
                    listItem.SubItems.Add(vItemName)
                    listItem.SubItems.Add(vItemCode)
                    listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                    listItem.SubItems.Add(vUnitCode)
                    listItem.SubItems.Add(vStatus)
                    listItem.SubItems.Add(vWHCode)
                    listItem.SubItems.Add(vShelfCode)
                    Me.ListViewItem.Items.Add(listItem)
                Next

                Me.TBIDNumber.Focus()
            Else
                MsgBox("This item find not found ", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBRVNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBRVNo.TextChanged


    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        Me.PNAddShelfByReceipt.Visible = False
    End Sub

    Private Sub TBIDNumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBIDNumber.KeyDown
        Dim i As Integer
        Dim vIndex As Integer
        Dim vCheckID As Integer

        If e.KeyCode = Keys.Enter Then
            If Me.TBIDNumber.Text <> "" Then
                vIndex = Me.TBIDNumber.Text

                For i = 0 To Me.ListViewItem.Items.Count - 1
                    vCheckID = Me.ListViewItem.Items(i).SubItems(0).Text

                    If vIndex = vCheckID Then
                        Me.TBShelfID.Text = Me.ListViewItem.Items(i).SubItems(1).Text
                        Me.TBItemName.Text = Me.ListViewItem.Items(i).SubItems(2).Text
                        Me.TBItemCode.Text = Me.ListViewItem.Items(i).SubItems(3).Text
                        Me.TBUnitCode.Text = Me.ListViewItem.Items(i).SubItems(5).Text

                        Me.TBShelfID.Focus()
                        Me.TBShelfID.SelectAll()

                    End If
                Next
            Else
                Me.TBItemCode.Text = ""
                Me.TBItemName.Text = ""
                Me.TBUnitCode.Text = ""
                Me.TBWHCode.Text = ""
                Me.TBShelfCode.Text = ""
                Me.TBShelfID.Text = ""
                Me.TBShelfID.Focus()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBIDNumber.Text = ""
            Me.TBIDNumber.Focus()
        End If
    End Sub

    Public Sub ClearScreen()
        Me.TBIDNumber.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBRVNo.Text = ""
        Me.TBShelfCode.Text = ""
        Me.TBShelfID.Text = ""
        Me.TBUnitCode.Text = ""
        Me.TBWHCode.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.TBRVNo.Focus()
    End Sub


    Private Sub TBIDNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBIDNumber.TextChanged
        If Me.TBIDNumber.Text = "" Then
            Me.TBItemCode.Text = ""
            Me.TBItemName.Text = ""
            Me.TBShelfCode.Text = ""
            Me.TBShelfID.Text = ""
            Me.TBUnitCode.Text = ""
            Me.TBWHCode.Text = ""
        End If
    End Sub

    Private Sub BTNSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSaveData.Click
        Dim i As Integer
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vZoneCode As String
        Dim vShelfCode As String
        Dim vUserScan As String
        Dim vModeScan As String
        Dim vIsSave As Integer
        Dim vDocNo As String

        If Me.ListViewItem.Items.Count > 0 Then
            vDocNo = Me.TBRVNo.Text

            For i = 0 To Me.ListViewItem.Items.Count - 1
                vItemCode = Trim(ListViewItem.Items(i).SubItems(3).Text)
                vBarCode = Trim(ListViewItem.Items(i).SubItems(3).Text)
                vItemName = Trim(ListViewItem.Items(i).SubItems(2).Text)
                vUnitCode = Trim(ListViewItem.Items(i).SubItems(5).Text)
                vWHCode = Trim(ListViewItem.Items(i).SubItems(7).Text)
                vZoneCode = Trim(ListViewItem.Items(i).SubItems(8).Text)
                vShelfCode = vb6.UCase(Trim(ListViewItem.Items(i).SubItems(1).Text))
                vUserScan = vUserName
                vModeScan = "บันทึกที่เก็บของสินค้าจากใบรับเข้า MC3000"
                vIsSave = ListViewItem.Items(i).SubItems(6).Text

                If vIsSave = 0 Then
                    vQuery = "exec dbo.USP_NP_AddItemReceiveShelfCode  '" & vItemCode & "','" & vBarCode & "','" & vItemName & "','" & vUnitCode & "','" & vWHCode & "','" & vZoneCode & "','" & vShelfCode & "','" & vUserScan & "','" & vModeScan & "','" & vDocNo & "' "
                    Call vGetData(vMemProfit, vQuery)
                End If
            Next i

            MsgBox("Save Date Is Complete", MsgBoxStyle.Critical, "Send Error Message")
            Call ClearScreen()
        End If
    End Sub

    Private Sub BTNClearScreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearScreen.Click
        Call ClearScreen()
    End Sub

    Private Sub ListViewItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItem.KeyDown
        Dim vIndex As Integer
        Dim vAnswer As Integer
        Dim vIsSave As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Back And Me.ListViewItem.Items.Count > 0 Then
            vIndex = Me.ListViewItem.FocusedItem.Index
            vIsSave = Me.ListViewItem.Items(vIndex).SubItems(6).Text

            vAnswer = MsgBox("Do you want clear shelfid this item ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 And vIsSave = 0 Then
                Me.ListViewItem.Items(vIndex).SubItems(1).Text = ""
            End If

            If vAnswer = 6 And vIsSave = 1 Then
                Me.ListViewItem.Items(vIndex).SubItems(1).Text = ""
                Me.ListViewItem.Items(vIndex).SubItems(6).Text = 0
            End If

        End If

    End Sub

    Private Sub ListViewItem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewItem.SelectedIndexChanged

    End Sub

    Private Sub BTNSClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSClear.Click
        Call vGetZone()
        Call vGetRow()
        Call vGetBay()
        Call vGetShelf()

        Me.CMBSelectZone.Focus()
    End Sub

    Public Sub vGetZone()
        Dim i As Integer
        Dim vZone As String

        Me.CMBSelectZone.Items.Clear()
        vQuery = "exec dbo.USP_NP_SearchShelfID 0"
        Call vGetData(vMemProfit, vQuery)
        If pds.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds.Tables(0).Rows.Count - 1
                vZone = pds.Tables(0).Rows(i)("code").ToString

                Me.CMBSelectZone.Items.Add(vZone)
            Next
        End If
    End Sub


    Public Sub vGetRow()
        Dim i As Integer
        Dim vRow As String

        Me.CMBSelectRow.Items.Clear()
        vQuery = "exec dbo.USP_NP_SearchShelfID 1"
        Call vGetData(vMemProfit, vQuery)
        If pds.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds.Tables(0).Rows.Count - 1
                vRow = pds.Tables(0).Rows(i)("code").ToString

                Me.CMBSelectRow.Items.Add(vRow)
            Next
        End If
    End Sub

    Public Sub vGetBay()
        Dim i As Integer
        Dim vBay As String

        Me.CMBSelectBay.Items.Clear()
        vQuery = "exec dbo.USP_NP_SearchShelfID 2"
        Call vGetData(vMemProfit, vQuery)
        If pds.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds.Tables(0).Rows.Count - 1
                vBay = pds.Tables(0).Rows(i)("code").ToString

                Me.CMBSelectBay.Items.Add(vBay)
            Next
        End If
    End Sub

    Public Sub vGetShelf()
        Dim i As Integer
        Dim vShelf As String

        Me.CMBSelectShelf.Items.Clear()
        vQuery = "exec dbo.USP_NP_SearchShelfID 3"
        Call vGetData(vMemProfit, vQuery)
        If pds.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds.Tables(0).Rows.Count - 1
                vShelf = pds.Tables(0).Rows(i)("code").ToString

                Me.CMBSelectShelf.Items.Add(vShelf)
            Next
        End If
    End Sub

    Private Sub CMBSelectRow_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CMBSelectRow.LostFocus
        If Me.CMBSelectRow.Text <> "" Then
            If vb6.Len(Me.CMBSelectRow.Text) <> 2 Then
                Me.CMBSelectRow.Text = ""
            End If
        End If
    End Sub

    Private Sub CMBSelectRow_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBSelectRow.SelectedIndexChanged

    End Sub

    Public Sub ShelfID()
        Dim vShelfID As String
        Dim vZone As String
        Dim vRow As String
        Dim vBay As String
        Dim vID As String


        If Me.CMBSelectZone.Text <> "" And Me.CMBSelectRow.Text <> "" And Me.CMBSelectBay.Text <> "" And Me.CMBSelectShelf.Text <> "" Then

            vZone = Me.CMBSelectZone.Text
            vRow = Me.CMBSelectRow.Text
            vBay = Me.CMBSelectBay.Text
            vID = Me.CMBSelectShelf.Text

            vShelfID = vZone & vRow & vBay & vID

            Me.TBGenShelfID.Text = vShelfID
        End If
    End Sub

    Public Sub ShelfClearScreen()
        Me.CMBSelectZone.Text = ""
        Me.CMBSelectRow.Text = ""
        Me.CMBSelectBay.Text = ""
        Me.CMBSelectShelf.Text = ""
        Me.CMBSection.Text = ""
        Me.TBGetShelfID.Text = ""
        vMemShelf = ""

        Me.TBGenShelfID.Text = ""
        Me.CMBSelectZone.Focus()
    End Sub

    Private Sub CMBSelectZone_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBSelectZone.SelectedIndexChanged

    End Sub

    Private Sub CMBSelectZone_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CMBSelectZone.TextChanged
        Call ShelfID()
    End Sub

    Private Sub CMBSelectRow_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CMBSelectRow.TextChanged
        Call ShelfID()
    End Sub

    Private Sub CMBSelectBay_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBSelectBay.SelectedIndexChanged

    End Sub

    Private Sub CMBSelectBay_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CMBSelectBay.TextChanged
        Call ShelfID()
    End Sub

    Private Sub CMBSelectShelf_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBSelectShelf.SelectedIndexChanged

    End Sub

    Private Sub CMBSelectShelf_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CMBSelectShelf.TextChanged
        Call ShelfID()
    End Sub

    Private Sub TBGenShelfID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBGenShelfID.TextChanged
        Dim vShelfID As String
        Dim vGetStaff As Integer

        If Me.TBGenShelfID.Text <> "" Then
            vShelfID = Me.TBGenShelfID.Text

            vQuery = "exec dbo.USP_NP_CheckShelfID '" & vShelfID & "'"
            Call vGetData(vMemProfit, vQuery)
            If pds.Tables(0).Rows.Count > 0 Then
                vMemShelf = pds.Tables(0).Rows(0)("code").ToString
                vGetStaff = pds.Tables(0).Rows(0)("staff").ToString
                Me.TBGetShelfID.Text = vMemShelf
                Me.CMBSection.Text = vGetStaff
            Else
                vMemShelf = ""
                Me.TBGetShelfID.Text = vMemShelf
                Me.CMBSection.Text = ""
            End If

        End If
    End Sub

    Private Sub BTNSDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSDelete.Click
        Dim vShelfID As String

        If Me.TBGetShelfID.Text <> "" Then
            vShelfID = Me.TBGetShelfID.Text

            vQuery = "exec dbo.USP_NP_DeleteShelfID '" & vShelfID & "'"
            Call vExecData(vMemProfit, vQuery)

            MsgBox("Delete shelfid is complete", MsgBoxStyle.Critical, "Send Information Message")
            Call ShelfClearScreen()
        Else
            MsgBox("Can not delete this shelfid", MsgBoxStyle.Critical, "Send Error Message")
            Me.CMBSelectZone.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub BTNSExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSExit.Click
        Me.PNManageShelfID.Visible = False
        Me.BTNManageShelf.Focus()
    End Sub

    Private Sub BTNSSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSSave.Click
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vStaff As String


        vWHCode = vMemProfit
        vShelfCode = Me.CMBShelfCode.Text
        vShelfID = Me.TBGenShelfID.Text
        vStaff = Me.CMBSection.Text

        vQuery = "exec dbo.USP_NP_InsertShelfID '" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vStaff & "','" & vUserID & "'"
        Call vExecData(vMemProfit, vQuery)

        MsgBox("Save shelfid is complete", MsgBoxStyle.Critical, "Send Information Message")
        Call ShelfClearScreen()

    End Sub

    Private Sub BTNRVScanClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNRVScanClose.Click
        Me.BTNRVScanClose.Visible = False
        Me.ListViewItem.Focus()
    End Sub

    Private Sub BTNScanItemClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNScanItemClose.Click
        Me.PNScanItemShelf.Visible = False
        Me.BTNManageShelf.Focus()
    End Sub

    Private Sub BTNManageShelf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNManageShelf.Click
        Me.PNManageShelfID.Visible = True
        Me.PNAddShelfByReceipt.Visible = False
        Me.PNScanItemShelf.Visible = False
        Me.CMBSelectZone.Focus()
    End Sub

    Private Sub BTNAddShelfByReceipt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddShelfByReceipt.Click
        Me.PNManageShelfID.Visible = False
        Me.PNAddShelfByReceipt.Visible = True
        Me.PNScanItemShelf.Visible = False
        Me.TBRVNo.Focus()
    End Sub

    Private Sub BTNAddShlefByMoveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddShlefByMoveItem.Click
        Me.PNManageShelfID.Visible = False
        Me.PNAddShelfByReceipt.Visible = False
        Me.PNScanItemShelf.Visible = True
        Me.TBScanItemBar.Focus()
    End Sub
End Class