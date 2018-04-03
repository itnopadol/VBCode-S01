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

    Private Sub TBShelfID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
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


                Me.TBShelfID.Text = ""
                Me.TBIDNumber.Focus()

            Else

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

    Private Sub TBShelfID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
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


                Me.TBShelfID.Text = ""
                Me.TBIDNumber.Focus()

            Else

                MsgBox(vShelfID & "  " & "This shelf not exist", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBShelfID.Text = ""
                Me.TBShelfID.Focus()
                Exit Sub
            End If
        End If
    End Sub

    Private Sub FormReceiveItemAddShelfID_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ClearScreen()
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

                If Me.ListViewItem.Items.Count > 0 Then
                    Me.ListViewItem.Focus()
                    Me.ListViewItem.Items(0).Selected = True
                    Me.ListViewItem.Items(0).Focused = True
                Else
                    Me.TBRVNo.Focus()
                    Me.TBRVNo.SelectAll()
                End If

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


        If vb6.InStr(Me.TBRVNo.Text, "@") > 0 Then
            vDocNo = vb6.Left(Me.TBRVNo.Text, vb6.Len(Me.TBRVNo.Text) - 1)

            Me.TBRVNo.Text = vDocNo

            vDocNo = Me.TBRVNo.Text


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

                If Me.ListViewItem.Items.Count > 0 Then
                    Me.ListViewItem.Focus()
                    Me.ListViewItem.Items(0).Selected = True
                    Me.ListViewItem.Items(0).Focused = True
                Else
                    Me.TBRVNo.Focus()
                    Me.TBRVNo.SelectAll()
                End If

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

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        On Error Resume Next

        Me.PNAddShelfByReceipt.Visible = False
    End Sub

    Private Sub TBIDNumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
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


                        Me.TBShelfID.Focus()
                        Me.TBShelfID.SelectAll()

                    End If
                Next
            Else

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
        Dim vGetYear As Double
        Dim vGetMonth As Integer
        Dim vYear As String
        Dim vMonth As String
        Dim vYear1 As String
        Dim vMonth1 As String

        vGetYear = Now.Year
        vGetMonth = Now.Month

        If vGetYear > 2013 Then
            vYear = vGetYear
            vYear1 = vb6.Left(vYear, 2)
        Else
            vGetYear = vGetYear + 543
            vYear = vGetYear
            vYear1 = vb6.Right(vYear, 2)
        End If

        vMonth = vGetMonth
        If vb6.Len(vMonth) > 1 Then
            vMonth1 = vMonth
        Else
            vMonth1 = "0" & vMonth
        End If


        Me.TBIDNumber.Text = ""
        Me.TBRVNo.Text = Trim(vMemProfit & "-RV" & vYear1 & vMonth1 & "-")
        Me.ListViewItem.Items.Clear()
        Me.TBRVNo.Focus()
    End Sub


    Private Sub TBIDNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.TBIDNumber.Text = "" Then

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

        On Error Resume Next

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

            MsgBox("Save Date Is Complete", MsgBoxStyle.Information, "Send Error Message")
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
        Dim vID As Integer

        'On Error Resume Next

        If e.KeyCode = Keys.Enter And Me.ListViewItem.Items.Count > 0 Then
            vIndex = Me.ListViewItem.FocusedItem.Index

            vID = vIndex + 1
            Me.TBRVID.Text = vID
            Me.TBRVITemCode.Text = Me.ListViewItem.Items(vIndex).SubItems(3).Text
            Me.TBRVItemName.Text = Me.ListViewItem.Items(vIndex).SubItems(2).Text
            Me.TBScanShelf.Text = Me.ListViewItem.Items(vIndex).SubItems(1).Text
            Me.TBRVWHCode.Text = Me.ListViewItem.Items(vIndex).SubItems(7).Text
            Me.TBRVShelfCode.Text = Me.ListViewItem.Items(vIndex).SubItems(8).Text

            Me.PNScanShelf.Visible = True
            Me.TBScanShelf.Focus()
            Me.TBScanShelf.SelectAll()

        End If

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

    Public Sub vGetShelfCode()
        Dim i As Integer
        Dim vShelfCode As String

        Me.CMBShelfCode.Items.Clear()
        vQuery = "select code from dbo.bcshelf where whcode = '" & vMemProfit & "' order by code"
        Call vGetData(vMemProfit, vQuery)
        If pds.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds.Tables(0).Rows.Count - 1
                vShelfCode = pds.Tables(0).Rows(i)("code").ToString

                Me.CMBShelfCode.Items.Add(vShelfCode)
            Next
        End If
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

    Private Sub CMBSelectRow_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBSelectRow.KeyDown
        If e.KeyCode = Keys.Up And Me.CMBSelectRow.SelectedIndex = 0 Then
            Me.CMBSelectZone.Focus()
        End If

        If e.KeyCode = Keys.Enter Then
            Me.CMBSelectBay.Focus()
        End If

        If e.KeyCode = Keys.Down And Me.CMBSelectRow.SelectedIndex = Me.CMBSelectRow.Items.Count - 1 Then
            Me.CMBSelectBay.Focus()
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

    Public Function vCheckShelfID(ByVal vShelfID As String)
        Dim vZone As String
        Dim vRow As String
        Dim vBay As String
        Dim vID As String
        Dim vGetID As Integer
        Dim vCheckID As Integer
        Dim vCount As Integer

        If Me.CMBSelectZone.Text <> "" And Me.CMBSelectRow.Text <> "" And Me.CMBSelectBay.Text <> "" And Me.CMBSelectShelf.Text <> "" Then

            vZone = Me.CMBSelectZone.Text
            vRow = Me.CMBSelectRow.Text
            vBay = Me.CMBSelectBay.Text
            vGetID = Me.CMBSelectShelf.Text
            vCheckID = vGetID - 1
            vID = vCheckID

            If vID > 0 Then
                vCheckShelfID = vZone & vRow & vBay & vID

                vQuery = "select isnull(count(code),0) as countshelf from npmaster.dbo.tb_rc_shelf where code = '" & vCheckShelfID & "' "
                Call vGetData(vMemProfit, vQuery)
                If pds.Tables(0).Rows.Count > 0 Then
                    vCount = pds.Tables(0).Rows(0)("countshelf").ToString
                Else
                    vCount = 0
                End If
            ElseIf vID = 0 Then
                vCount = 1
            End If

            Return vCount
        End If
    End Function


    Public Function vExistShelfID(ByVal vShelfID As String)
        Dim vCountExist As Integer


        If Me.CMBSelectZone.Text <> "" And Me.CMBSelectRow.Text <> "" And Me.CMBSelectBay.Text <> "" And Me.CMBSelectShelf.Text <> "" Then

            vQuery = "select isnull(count(code),0) as countshelf from npmaster.dbo.tb_rc_shelf where code = '" & vShelfID & "' "
            Call vGetData(vMemProfit, vQuery)
            If pds.Tables(0).Rows.Count > 0 Then
                vCountExist = pds.Tables(0).Rows(0)("countshelf").ToString
            Else
                vCountExist = 0
            End If
        End If

        Return vCountExist
    End Function


    Public Function vCountItemShelfID(ByVal vWHCode As String, ByVal vZoneCode As String, ByVal vShelfID As String)
        Dim vCountItem As Integer

        If Me.CMBSelectZone.Text <> "" And Me.CMBSelectRow.Text <> "" And Me.CMBSelectBay.Text <> "" And Me.CMBSelectShelf.Text <> "" Then

            vQuery = "select isnull(count(itemcode),0) as vCount from npmaster.dbo.np_scanbarcode_logs  where whcode = '" & vWHCode & "' and zonecode = '" & vZoneCode & "' and shelfcode = '" & vShelfID & "' and isused = 1"
            Call vGetData(vMemProfit, vQuery)
            If pds.Tables(0).Rows.Count > 0 Then
                vCountItem = pds.Tables(0).Rows(0)("vCount").ToString
            Else
                vCountItem = 0
            End If
        End If

        Return vCountItem
    End Function

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

    Private Sub CMBSelectZone_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBSelectZone.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.CMBSelectRow.Focus()
        End If

        If e.KeyCode = Keys.Down And Me.CMBSelectZone.SelectedIndex = Me.CMBSelectZone.Items.Count Then
            Me.CMBSelectBay.Focus()
        End If
    End Sub

    Private Sub CMBSelectZone_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBSelectZone.SelectedIndexChanged

    End Sub

    Private Sub CMBSelectZone_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CMBSelectZone.TextChanged
        Call ShelfID()
    End Sub

    Private Sub CMBSelectRow_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CMBSelectRow.TextChanged
        Call ShelfID()
    End Sub

    Private Sub CMBSelectBay_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBSelectBay.KeyDown
        If e.KeyCode = Keys.Up And Me.CMBSelectBay.SelectedIndex = 0 Then
            Me.CMBSelectRow.Focus()
        End If

        If e.KeyCode = Keys.Enter Then
            Me.CMBSelectShelf.Focus()
        End If

        If e.KeyCode = Keys.Down And Me.CMBSelectBay.SelectedIndex = Me.CMBSelectBay.Items.Count - 1 Then
            Me.CMBSelectShelf.Focus()
        End If
    End Sub

    Private Sub CMBSelectBay_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBSelectBay.SelectedIndexChanged

    End Sub

    Private Sub CMBSelectBay_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CMBSelectBay.TextChanged
        Call ShelfID()
    End Sub

    Private Sub CMBSelectShelf_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBSelectShelf.KeyDown
        If e.KeyCode = Keys.Up And Me.CMBSelectShelf.SelectedIndex = 0 Then
            Me.CMBSelectBay.Focus()
        End If

        If e.KeyCode = Keys.Enter Then
            Me.CMBSection.Focus()
        End If

        If e.KeyCode = Keys.Down And Me.CMBSelectShelf.SelectedIndex = Me.CMBSelectShelf.Items.Count - 1 Then
            Me.CMBSection.Focus()
        End If
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
                vGetStaff = 0 'pds.Tables(0).Rows(0)("staff").ToString
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
        Dim vWHCode As String
        Dim vZoneCode As String
        Dim vCountItem As Double
        Dim vAnswer As Integer

        If Me.TBGenShelfID.Text <> "" Then
            vWHCode = vMemProfit
            vZoneCode = Me.CMBShelfCode.Text
            vShelfID = Me.TBGenShelfID.Text

            vCountItem = vCountItemShelfID(vWHCode, vZoneCode, vShelfID)

            If vCountItem > 0 Then
                Exit Sub
            End If

            vAnswer = MsgBox("Do you want delete new shelfid :" & vShelfID, MsgBoxStyle.YesNo, "Send Question Message ?")

            If vAnswer = 6 Then
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
            End If
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
        Dim vAnswer As Integer
        Dim vCheckShelf As Integer
        Dim vExistShelf As Integer

        If Me.TBGenShelfID.Text <> "" Then
            vAnswer = MsgBox("Do you want save new shelfid :" & vShelfID, MsgBoxStyle.YesNo, "Send Question Message ?")

            If vAnswer = 6 Then
                vWHCode = vMemProfit
                vShelfCode = Me.CMBShelfCode.Text
                vShelfID = Me.TBGenShelfID.Text
                vStaff = Me.CMBSection.Text

                vExistShelf = vExistShelfID(vShelfID)
                vCheckShelf = vCheckShelfID(vShelfID)

                If vExistShelf = 1 Then
                    Me.PNMsg.Visible = True
                    Me.PNMsg.BringToFront()
                    Me.TBMsg.Text = "ชั้นเก็บที่จะเพิ่มมีอยู่แล้ว กรุณาตรวจสอบ"
                    Exit Sub
                End If

                If vCheckShelf = 0 Then
                    Me.PNMsg.Visible = True
                    Me.PNMsg.BringToFront()
                    Me.TBMsg.Text = "ชั้นเก็บที่จะเพิ่มต้องเรียงจากชั้นเก็บที่มีอยู่แล้ว ห้ามเพิ่มข้ามชั้น กรุณาตรวจสอบ"
                    Exit Sub
                End If

                vQuery = "exec dbo.USP_NP_InsertShelfID '" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vStaff & "','" & vUserID & "'"
                Call vExecData(vMemProfit, vQuery)

                MsgBox("Save shelfid is complete", MsgBoxStyle.Information, "Send Information Message")
                Call ShelfClearScreen()
            End If
        End If
    End Sub

    Private Sub BTNRVScanClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNRVScanClose.Click
        On Error Resume Next

        Me.PNScanShelf.Visible = False
        Me.ListViewItem.Focus()
        If Me.ListViewItem.Items.Count > 0 Then
            Me.ListViewItem.Items(0).Selected = True
        Else
            Me.TBRVNo.Focus()
        End If

    End Sub

    Private Sub BTNScanItemClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNScanItemClose.Click
        Me.PNScanItemShelf.Visible = False
        Me.BTNManageShelf.Focus()
    End Sub

    Private Sub BTNManageShelf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNManageShelf.Click
        Me.PNManageShelfID.Visible = True
        Me.PNAddShelfByReceipt.Visible = False
        Me.PNScanItemShelf.Visible = False
        Me.PNSearchItemShelf.Visible = False

        Call Me.vGetZone()
        Call Me.vGetRow()
        Call Me.vGetBay()
        Call Me.vGetShelf()
        Call vGetShelfCode()

        Me.CMBSelectZone.Focus()
    End Sub

    Private Sub BTNAddShelfByReceipt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddShelfByReceipt.Click
        On Error Resume Next

        Me.PNManageShelfID.Visible = False
        Me.PNAddShelfByReceipt.Visible = True
        Me.PNScanItemShelf.Visible = False
        Me.PNSearchItemShelf.Visible = False
        Me.TBRVNo.Focus()
    End Sub

    Private Sub BTNAddShlefByMoveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAddShlefByMoveItem.Click
        'Me.PNManageShelfID.Visible = False
        'Me.PNAddShelfByReceipt.Visible = False
        'Me.PNScanItemShelf.Visible = True
        'Me.PNSearchItemShelf.Visible = False
        'Me.TBScanItemBar.Focus()

        On Error Resume Next

        FormShelfAddItem.Show()
        Me.Hide()
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        On Error Resume Next

        Call ClearScreen()
        FormMainApplication.Show()
        Me.Hide()
    End Sub

    Private Sub BTNCloseMsg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseMsg.Click
        Me.PNMsg.Visible = False
    End Sub

    Private Sub CMBSection_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBSection.KeyDown
        If e.KeyCode = Keys.Up And Me.CMBSection.SelectedIndex = 0 Then
            Me.CMBSelectShelf.Focus()
        End If

        If e.KeyCode = Keys.Enter Then
            Me.CMBShelfCode.Focus()
        End If

        If e.KeyCode = Keys.Down And Me.CMBSection.SelectedIndex = Me.CMBSection.Items.Count - 1 Then
            Me.CMBShelfCode.Focus()
        End If
    End Sub

    Private Sub CMBSection_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBSection.SelectedIndexChanged
        Dim vCountItem As Double
        Dim vWHCode As String
        Dim vZoneCode As String
        Dim vShelfID As String

        vWHCode = vMemProfit
        vZoneCode = Me.CMBShelfCode.Text
        vShelfID = Me.TBGenShelfID.Text

        vCountItem = vCountItemShelfID(vWHCode, vZoneCode, vShelfID)

        Me.TBCountItem.Text = vCountItem
    End Sub

    Private Sub CMBShelfCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBShelfCode.KeyDown
        If e.KeyCode = Keys.Up And Me.CMBShelfCode.SelectedIndex = 0 Then
            Me.CMBSection.Focus()
        End If
    End Sub

    Private Sub CMBShelfCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBShelfCode.SelectedIndexChanged

    End Sub

    Private Sub BTNRVBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNRVBack.Click
        Dim vIndex As Integer
        Dim vID As Integer
        Dim vShowID As Integer
        Dim vBackID As Integer
        Dim vAnswer As Integer
        Dim vIsSave As Integer

        On Error Resume Next

        If Me.ListViewItem.Items.Count > 0 And Me.TBRVID.Text <> "" Then
            vID = Me.TBRVID.Text

            vIndex = vID - 1
            If vIndex <> 0 Then
                vBackID = vIndex - 1
            Else
                Me.TBScanShelf.Focus()
                Exit Sub
            End If

            vShowID = vBackID + 1

            Me.TBRVID.Text = vShowID
            Me.TBRVITemCode.Text = Me.ListViewItem.Items(vBackID).SubItems(3).Text
            Me.TBRVItemName.Text = Me.ListViewItem.Items(vBackID).SubItems(2).Text
            Me.TBScanShelf.Text = Me.ListViewItem.Items(vBackID).SubItems(1).Text
            Me.TBRVWHCode.Text = Me.ListViewItem.Items(vBackID).SubItems(7).Text
            Me.TBRVShelfCode.Text = Me.ListViewItem.Items(vBackID).SubItems(8).Text

            Me.PNScanShelf.Visible = True
            Me.TBScanShelf.Focus()
            Me.TBScanShelf.SelectAll()

        End If
    End Sub

    Private Sub BTNRVNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNRVNext.Click
        Dim vIndex As Integer
        Dim vID As Integer
        Dim vNextID As Integer
        Dim vShowID As Integer
        Dim vAnswer As Integer
        Dim vIsSave As Integer

        On Error Resume Next

        If Me.ListViewItem.Items.Count > 0 And Me.TBRVID.Text <> "" Then
            vID = Me.TBRVID.Text

            vIndex = vID - 1
            If vIndex < (Me.ListViewItem.Items.Count - 1) Then
                vNextID = vIndex + 1
            Else
                Me.TBScanShelf.Focus()
                Exit Sub
            End If

            vShowID = vNextID + 1
            Me.TBRVID.Text = vShowID
            Me.TBRVITemCode.Text = Me.ListViewItem.Items(vNextID).SubItems(3).Text
            Me.TBRVItemName.Text = Me.ListViewItem.Items(vNextID).SubItems(2).Text
            Me.TBScanShelf.Text = Me.ListViewItem.Items(vNextID).SubItems(1).Text
            Me.TBRVWHCode.Text = Me.ListViewItem.Items(vNextID).SubItems(7).Text
            Me.TBRVShelfCode.Text = Me.ListViewItem.Items(vNextID).SubItems(8).Text

            Me.PNScanShelf.Visible = True
            Me.TBScanShelf.Focus()
            Me.TBScanShelf.SelectAll()

        End If
    End Sub

    Private Sub BTNRVInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNRVInsert.Click
        Dim vShelfID As String
        Dim vItemCode As String
        Dim vIndex As Integer
        Dim vID As Integer
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vCheckShelf As String

        On Error Resume Next

        If Me.TBRVID.Text <> "" Then
            vID = Me.TBRVID.Text
            vIndex = vID - 1
            vItemCode = Me.TBRVITemCode.Text
            vWHCode = Me.TBRVWHCode.Text
            vShelfCode = Me.TBRVShelfCode.Text
            vShelfID = Me.TBScanShelf.Text

            vQuery = "exec dbo.USP_MB_SearchShelfMasterDetails '" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vCheckShelf = pds.Tables(0).Rows(0)("shelf").ToString
            Else
                vCheckShelf = ""
                MsgBox(vShelfID & "  " & "This shelf not exist", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBShelfID.Text = ""
                Me.TBShelfID.Focus()
                Exit Sub
            End If

            Me.ListViewItem.Items(vIndex).SubItems(1).Text = vShelfID
            Me.ListViewItem.Items(vIndex).SubItems(6).Text = 0

            Me.TBRVID.Text = ""
            Me.TBRVITemCode.Text = ""
            Me.TBRVWHCode.Text = ""
            Me.TBRVShelfCode.Text = ""
            Me.TBScanShelf.Text = ""

            Me.PNScanShelf.Visible = False

            Me.ListViewItem.Focus()
            If vIndex < Me.ListViewItem.Items.Count - 1 Then
                Me.ListViewItem.Items(vIndex + 1).Focused = True
                Me.ListViewItem.Items(vIndex + 1).Selected = True
            Else
                Me.ListViewItem.Items(vIndex).Focused = True
                Me.ListViewItem.Items(vIndex).Selected = True
            End If


        End If
    End Sub

    Private Sub TBScanShelf_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBScanShelf.KeyDown
        Dim vShelfID As String
        Dim vItemCode As String
        Dim vIndex As Integer
        Dim vID As Integer
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vCheckShelf As String

        On Error Resume Next

        If e.KeyCode = Keys.Enter And vb6.InStr(Me.TBScanShelf.Text, "@") > 0 Then
            If vb6.InStr(Me.TBScanShelf.Text, "@") > 0 Then
                vShelfID = vb6.Left(Me.TBScanShelf.Text, vb6.Len(Me.TBScanShelf.Text) - 1)

                Me.TBScanShelf.Text = vShelfID

                vShelfID = Me.TBScanShelf.Text
            End If
        End If

        If e.KeyCode = Keys.Enter And vb6.InStr(Me.TBScanShelf.Text, "@") = 0 Then
            If Me.TBRVID.Text <> "" Then
                vID = Me.TBRVID.Text
                vIndex = vID - 1
                vItemCode = Me.TBRVITemCode.Text
                vWHCode = Me.TBRVWHCode.Text
                vShelfCode = Me.TBRVShelfCode.Text
                vShelfID = Me.TBScanShelf.Text

                vQuery = "exec dbo.USP_MB_SearchShelfMasterDetails '" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "'"
                Call vGetData(vMemProfit, vQuery)

                If pds.Tables(0).Rows.Count > 0 Then
                    vCheckShelf = pds.Tables(0).Rows(0)("shelf").ToString
                Else
                    vCheckShelf = ""
                    MsgBox(vShelfID & "  " & "This shelf not exist", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBShelfID.Text = ""
                    Me.TBShelfID.Focus()
                    Exit Sub
                End If

                Me.ListViewItem.Items(vIndex).SubItems(1).Text = vShelfID
                Me.ListViewItem.Items(vIndex).SubItems(6).Text = 0

                Me.TBRVID.Text = ""
                Me.TBRVITemCode.Text = ""
                Me.TBRVWHCode.Text = ""
                Me.TBRVShelfCode.Text = ""
                Me.TBScanShelf.Text = ""

                Me.PNScanShelf.Visible = False

                Me.ListViewItem.Focus()
                If vIndex < Me.ListViewItem.Items.Count - 1 Then
                    Me.ListViewItem.Items(vIndex + 1).Focused = True
                    Me.ListViewItem.Items(vIndex + 1).Selected = True
                Else
                    Me.ListViewItem.Items(vIndex).Focused = True
                    Me.ListViewItem.Items(vIndex).Selected = True
                End If
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBScanShelf.Text = ""
            Me.TBScanShelf.Focus()
        End If
    End Sub

    Private Sub TBScanShelf_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBScanShelf.TextChanged
        Dim vShelfID As String

        On Error Resume Next

        If vb6.InStr(Me.TBScanShelf.Text, "@") > 0 Then
            vShelfID = vb6.Left(Me.TBScanShelf.Text, vb6.Len(Me.TBScanShelf.Text) - 1)

            Me.TBScanShelf.Text = vShelfID

            vShelfID = Me.TBScanShelf.Text
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchItemCode.TextChanged

    End Sub

    Private Sub TBSearchBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchBarCode.KeyDown
        Dim vBarCode As String
        Dim i As Integer
        Dim vIndex As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vUserCode As String
        Dim vScanDateTime As String
        Dim vShelfID As String
        Dim vWHCode As String
        Dim vZoneCode As String

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            vBarCode = Me.TBSearchBarCode.Text
            Me.ListViewSearchShelfID.Items.Clear()

            vQuery = "exec dbo.USP_MB_SearchItemScanShelfCode '" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)
            If pds.Tables(0).Rows.Count > 0 Then

                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString

                Me.TBSearchItemCode.Text = vItemCode
                Me.TBSearchItemName.Text = vItemName
                Me.TBSearchUnit.Text = vUnitCode

                For i = 0 To pds.Tables(0).Rows.Count - 1
                    vShelfID = pds.Tables(0).Rows(i)("shelfcode").ToString
                    vWHCode = pds.Tables(0).Rows(i)("whcode").ToString
                    vZoneCode = pds.Tables(0).Rows(i)("zonecode").ToString
                    vUserCode = pds.Tables(0).Rows(i)("userscan").ToString
                    vScanDateTime = pds.Tables(0).Rows(i)("scandatetime").ToString

                    vIndex = i + 1

                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(vShelfID)
                    listItem.SubItems.Add(vWHCode)
                    listItem.SubItems.Add(vZoneCode)
                    listItem.SubItems.Add(vScanDateTime)
                    listItem.SubItems.Add(vUserCode)
                    Me.ListViewSearchShelfID.Items.Add(listItem)
                Next

                Me.TBSearchBarCode.Focus()

            Else
                Me.TBSearchItemName.Text = ""
                Me.TBSearchUnit.Text = ""
                Me.ListViewSearchShelfID.Items.Clear()
                MsgBox("This item not scan add shelf ", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBSearchBarCode.Text = ""
            Me.TBSearchItemName.Text = ""
            Me.TBSearchUnit.Text = ""
            Me.ListViewSearchShelfID.Items.Clear()
            Me.TBSearchBarCode.Focus()
        End If
    End Sub

    Private Sub TBSearchBarCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchBarCode.TextChanged
        Dim vBarCode As String
        Dim i As Integer
        Dim vIndex As Integer
        Dim vItemCode As String
        Dim vUnitCode As String
        Dim vItemName As String
        Dim vUserCode As String
        Dim vScanDateTime As String
        Dim vShelfID As String
        Dim vWHCode As String
        Dim vZoneCode As String

        On Error Resume Next

        If vb6.InStr(Me.TBSearchBarCode.Text, "@") > 0 Then
            vBarCode = vb6.Left(Me.TBSearchBarCode.Text, vb6.Len(Me.TBSearchBarCode.Text) - 1)

            Me.TBSearchBarCode.Text = vBarCode

            vBarCode = Me.TBSearchBarCode.Text
            Me.ListViewSearchShelfID.Items.Clear()

            vQuery = "exec dbo.USP_MB_SearchItemScanShelfCode '" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)
            If pds.Tables(0).Rows.Count > 0 Then

                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString

                Me.TBSearchItemCode.Text = vItemCode
                Me.TBSearchItemName.Text = vItemName
                Me.TBSearchUnit.Text = vUnitCode

                For i = 0 To pds.Tables(0).Rows.Count - 1
                    vShelfID = pds.Tables(0).Rows(i)("shelfcode").ToString
                    vWHCode = pds.Tables(0).Rows(i)("whcode").ToString
                    vZoneCode = pds.Tables(0).Rows(i)("zonecode").ToString
                    vUserCode = pds.Tables(0).Rows(i)("userscan").ToString
                    vScanDateTime = pds.Tables(0).Rows(i)("scandatetime").ToString

                    vIndex = i + 1

                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(vShelfID)
                    listItem.SubItems.Add(vWHCode)
                    listItem.SubItems.Add(vZoneCode)
                    listItem.SubItems.Add(vScanDateTime)
                    listItem.SubItems.Add(vUserCode)
                    Me.ListViewSearchShelfID.Items.Add(listItem)
                Next

                Me.TBSearchBarCode.Focus()

            Else
                Me.TBSearchItemName.Text = ""
                Me.TBSearchUnit.Text = ""
                Me.ListViewSearchShelfID.Items.Clear()
                MsgBox("This item find not found ", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If
        End If

        If Me.TBSearchBarCode.Text = "" Then
            Me.TBSearchItemCode.Text = ""
            Me.TBSearchItemName.Text = ""
            Me.TBSearchUnit.Text = ""
            Me.ListViewSearchShelfID.Items.Clear()
        End If
    End Sub

    Private Sub BTNSearchShelfExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchShelfExit.Click
        On Error Resume Next

        Me.PNSearchItemShelf.Visible = False
    End Sub

    Private Sub BTNSearchItemShelf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchItemShelf.Click
        On Error Resume Next

        Me.PNManageShelfID.Visible = False
        Me.PNAddShelfByReceipt.Visible = False
        Me.PNScanItemShelf.Visible = False
        Me.PNSearchItemShelf.Visible = True
        Me.TBSearchBarCode.Focus()
    End Sub

    Private Sub ListViewSearchShelfID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearchShelfID.KeyDown
        Dim vIndex As Integer
        Dim vAnswer As Integer
        Dim vItemCode As String
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vZoneCode As String
        Dim vShelfID As String
        Dim vIsSave As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Back And Me.ListViewSearchShelfID.Items.Count > 0 Then
            vIndex = Me.ListViewSearchShelfID.FocusedItem.Index
            vItemCode = Me.TBSearchItemCode.Text
            vUnitCode = Me.TBSearchUnit.Text
            vWHCode = Me.ListViewSearchShelfID.Items(vIndex).SubItems(2).Text
            vZoneCode = Me.ListViewSearchShelfID.Items(vIndex).SubItems(3).Text
            vShelfID = Me.ListViewSearchShelfID.Items(vIndex).SubItems(1).Text
            vIsSave = 1

            vAnswer = MsgBox("Do you want delete itemcode " & vItemCode & " ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 Then
                Me.ListViewSearchShelfID.Items.RemoveAt(vIndex)
                Call GenLineNumber()
                vQuery = "exec dbo.USP_NP_DeleteItemShelfIDLogs '" & vItemCode & "','" & vUnitCode & "','" & vWHCode & "','" & vZoneCode & "','" & vShelfID & "'"
                Call vGetData(vMemProfit, vQuery)
                Me.TBSearchBarCode.Focus()
            End If
        End If
    End Sub

    Public Sub GenLineNumber()
        Dim i As Integer
        Dim n As Integer

        On Error Resume Next

        If Me.ListViewSearchShelfID.Items.Count > 0 Then

            For i = 0 To Me.ListViewSearchShelfID.Items.Count - 1
                n = i + 1
                Me.ListViewSearchShelfID.Items(i).SubItems(0).Text = n
            Next
        End If
    End Sub

    Private Sub ListViewSearchShelfID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewSearchShelfID.SelectedIndexChanged

    End Sub

    Private Sub BTNUpdateAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNUpdateAll.Click
        Dim vShelfID As String
        Dim vIndex As Integer
        Dim vID As Integer
        Dim i As Integer
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vCheckShelf As String

        On Error Resume Next

        If Me.TBScanShelf.Text <> "" And Me.ListViewItem.Items.Count > 0 Then
            vWHCode = Me.TBRVWHCode.Text
            vShelfCode = Me.TBRVShelfCode.Text
            vShelfID = Me.TBScanShelf.Text

            vQuery = "exec dbo.USP_MB_SearchShelfMasterDetails '" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vCheckShelf = pds.Tables(0).Rows(0)("shelf").ToString
            Else
                vCheckShelf = ""
                MsgBox(vShelfID & "  " & "This shelf not exist", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBShelfID.Text = ""
                Me.TBShelfID.Focus()
                Exit Sub
            End If

            For i = 0 To Me.ListViewItem.Items.Count - 1
                Me.ListViewItem.Items(i).SubItems(1).Text = vShelfID
                Me.ListViewItem.Items(i).SubItems(6).Text = 0
            Next
            Me.TBRVID.Text = ""
            Me.TBRVITemCode.Text = ""
            Me.TBRVWHCode.Text = ""
            Me.TBRVShelfCode.Text = ""
            Me.TBScanShelf.Text = ""

            Me.PNScanShelf.Visible = False

            Me.ListViewItem.Focus()

        End If
    End Sub
End Class