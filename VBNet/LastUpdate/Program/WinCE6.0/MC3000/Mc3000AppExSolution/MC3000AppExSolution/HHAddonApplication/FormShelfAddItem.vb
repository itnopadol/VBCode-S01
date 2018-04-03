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

Public Class FormShelfAddItem

    Private MyEventHandler As System.EventHandler = Nothing
    Dim vQuery As String
    Dim vMemDocDate As String


    Private Sub FormShelfAddItem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        vQuery = "exec dbo.USP_NP_CheckDocDate"
        Call vGetData1(vMemProfit, vQuery)
        If pds1.Tables(0).Rows.Count > 0 Then
            vMemDocDate = pds1.Tables(0).Rows(0)("vdocdate").ToString
        End If

        Call vGetWareHouse()
        Call vGetShelf()
    End Sub

    Public Sub vGetWareHouse()
        Me.CMBWHCode.Items.Add(vMemProfit)
        Me.CMBWHCode.SelectedIndex = 0
    End Sub

    Public Sub vGetShelf()
        Dim i As Integer
        Dim vShelfCode As String

        On Error Resume Next

        vQuery = "exec dbo.USP_NP_SearchShelfByWareHouse"
        Call vGetData(vMemProfit, vQuery)

        If pds.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds.Tables(0).Rows.Count - 1
                vShelfCode = pds.Tables(0).Rows(i)("code").ToString
                Me.CMBZone.Items.Add(vShelfCode)
            Next
        End If

        If Me.CMBZone.Items.Count > 0 Then
            Me.CMBZone.SelectedIndex = 0
        End If
    End Sub

    Private Sub TBBarcode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarcode.KeyDown
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vBarCode As String
        Dim vUserScan As String
        Dim vModeScan As String

        Dim vWHCode As String
        Dim vZoneCode As String
        Dim vShelfID As String

        Dim i As Integer
        Dim n As Integer
        Dim vCheckLine As Integer
        Dim vDocDate As String

        Dim vCheckItemCode As String
        Dim vCheckShelf As String

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            vBarCode = Me.TBBarcode.Text
            vWHCode = Me.CMBWHCode.Text
            vZoneCode = Me.CMBZone.Text
            vShelfID = Me.TBShelfID.Text
            vDocDate = Now

            vQuery = "exec dbo.USP_MB_SearchShelfMasterDetails '" & vWHCode & "','" & vZoneCode & "','" & vShelfID & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vCheckShelf = pds.Tables(0).Rows(0)("shelf").ToString
            Else
                vCheckShelf = ""
                MsgBox(vShelfID & "  " & "This shelf not exist", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarcode.Text = ""
                Me.TBShelfID.Focus()
                Exit Sub
            End If

            vUserScan = vUserName
            vModeScan = "หน้ายิงบาร์โค้ดตรวจสอบชั้นเก็บ MC3000"

            vQuery = "exec dbo.usp_np_DataItemDetails '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString

                If Me.ListViewItem.Items.Count > 0 Then
                    For n = 0 To Me.ListViewItem.Items.Count - 1
                        vCheckItemCode = Me.ListViewItem.Items(n).SubItems(1).Text

                        If vItemCode = vCheckItemCode Then
                            Me.ListViewItem.Items(n).SubItems(6).Text = vDocDate
                            Me.ListViewItem.Items(n).SubItems(7).Text = 0
                            Me.TBBarcode.Text = ""
                            Me.TBBarcode.Focus()
                            Exit Sub
                        End If
                    Next
                End If

                i = Me.ListViewItem.Items.Count + 1

                Dim listItem As New ListViewItem(i)
                listItem.SubItems.Add(vItemCode)
                listItem.SubItems.Add(vItemName)
                listItem.SubItems.Add(vShelfID)
                listItem.SubItems.Add(vWHCode)
                listItem.SubItems.Add(vZoneCode)
                listItem.SubItems.Add(vDocDate)
                listItem.SubItems.Add(0)
                listItem.SubItems.Add(vUnitCode)
                Me.ListViewItem.Items.Add(listItem)

                vQuery = "exec dbo.USP_NP_InsertScanItemShelfCode  '" & vItemCode & "','" & vBarCode & "','" & vItemName & "','" & vUnitCode & "','" & vWHCode & "','" & vZoneCode & "','" & vShelfID & "','" & vUserScan & "','" & vModeScan & "' "
                Call vGetData(vMemProfit, vQuery)

                If Me.ListViewItem.Items.Count > 0 Then
                    vCheckLine = Me.ListViewItem.Items.Count
                    Me.ListViewItem.Focus()
                    VScrollBar1.Value = vCheckLine - 1
                End If

                Me.TBBarcode.Text = ""
                Me.TBBarcode.Focus()
            Else
                MsgBox(vBarCode & " " & "This barcode not exist", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarcode.Text = ""
                Me.TBBarcode.Focus()
            End If

        End If


        If e.KeyCode = Keys.Up Then
            Me.TBShelfID.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBBarcode.Text = ""
            Me.TBBarcode.Focus()
        End If

        If e.KeyCode = Keys.Down And Me.ListViewItem.Items.Count > 0 Then
            Me.ListViewItem.Items(0).Focused = True
            Me.ListViewItem.Items(0).Selected = True
            Me.ListViewItem.Focus()
        End If
    End Sub

    Private Sub TBBarcode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarcode.TextChanged
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vBarCode As String
        Dim vUserScan As String
        Dim vModeScan As String

        Dim vWHCode As String
        Dim vZoneCode As String
        Dim vShelfID As String

        Dim i As Integer
        Dim n As Integer
        Dim vCheckLine As Integer
        Dim vDocDate As String
        Dim vCheckShelf As String

        Dim vCheckItemCode As String

        On Error Resume Next

        If Me.TBShelfID.Text <> "" Then

            If vb6.InStr(Me.TBBarcode.Text, "@") > 0 Then
                vBarCode = vb6.Left(Me.TBBarcode.Text, vb6.Len(Me.TBBarcode.Text) - 1)
                vWHCode = Me.CMBWHCode.Text
                vZoneCode = Me.CMBZone.Text
                vShelfID = Me.TBShelfID.Text
                vDocDate = Now

                vQuery = "exec dbo.USP_MB_SearchShelfMasterDetails '" & vWHCode & "','" & vZoneCode & "','" & vShelfID & "'"
                Call vGetData(vMemProfit, vQuery)

                If pds.Tables(0).Rows.Count > 0 Then
                    vCheckShelf = pds.Tables(0).Rows(0)("shelf").ToString
                Else
                    vCheckShelf = ""
                    MsgBox(vShelfID & "  " & "This shelf not exist", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBBarcode.Text = ""
                    Me.TBShelfID.Focus()
                    Exit Sub
                End If


                Me.TBBarcode.Text = vBarCode
                vUserScan = vUserName
                vModeScan = "หน้ายิงบาร์โค้ดตรวจสอบชั้นเก็บ MC3000"

                vQuery = "exec dbo.usp_np_DataItemDetails '" & vMemProfit & "','" & vBarCode & "'"
                Call vGetData(vMemProfit, vQuery)

                If pds.Tables(0).Rows.Count > 0 Then
                    vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                    vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                    vBarCode = pds.Tables(0).Rows(0)("barcode").ToString
                    vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString

                    If Me.ListViewItem.Items.Count > 0 Then
                        For n = 0 To Me.ListViewItem.Items.Count - 1
                            vCheckItemCode = Me.ListViewItem.Items(n).SubItems(1).Text

                            If vItemCode = vCheckItemCode Then
                                Me.ListViewItem.Items(n).SubItems(6).Text = vDocDate
                                Me.ListViewItem.Items(n).SubItems(7).Text = 0
                                Me.TBBarcode.Text = ""
                                Me.TBBarcode.Focus()
                                Exit Sub
                            End If
                        Next
                    End If

                    i = Me.ListViewItem.Items.Count + 1

                    Dim listItem As New ListViewItem(i)
                    listItem.SubItems.Add(vItemCode)
                    listItem.SubItems.Add(vItemName)
                    listItem.SubItems.Add(vShelfID)
                    listItem.SubItems.Add(vWHCode)
                    listItem.SubItems.Add(vZoneCode)
                    listItem.SubItems.Add(vDocDate)
                    listItem.SubItems.Add(0)
                    listItem.SubItems.Add(vUnitCode)
                    Me.ListViewItem.Items.Add(listItem)


                    vQuery = "exec dbo.USP_NP_InsertScanItemShelfCode  '" & vItemCode & "','" & vBarCode & "','" & vItemName & "','" & vUnitCode & "','" & vWHCode & "','" & vZoneCode & "','" & vShelfID & "','" & vUserScan & "','" & vModeScan & "' "
                    Call vGetData(vMemProfit, vQuery)


                    If Me.ListViewItem.Items.Count > 0 Then
                        vCheckLine = Me.ListViewItem.Items.Count
                        Me.ListViewItem.Focus()
                        VScrollBar1.Value = vCheckLine - 1
                    End If

                    Me.TBBarcode.Text = ""
                    Me.TBBarcode.Focus()
                Else
                    MsgBox(vBarCode & " " & "This barcode not exist", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBBarcode.Text = ""
                    Me.TBBarcode.Focus()
                End If
            End If
        End If
    End Sub

    Public Sub SearchShelfItem()
        Dim vWHCode As String
        Dim vZoneCode As String
        Dim vShelfID As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vBarCode As String
        Dim vDocDate As String
        Dim vUnitCode As String

        Dim i As Integer
        Dim n As Integer

        On Error Resume Next

        Me.ListViewItem.Items.Clear()

        If vb6.InStr(Me.TBShelfID.Text, "@") > 0 Then
            vShelfID = vb6.Left(Me.TBShelfID.Text, vb6.Len(Me.TBShelfID.Text) - 1)
        Else
            vShelfID = Me.TBShelfID.Text
        End If

        vWHCode = Me.CMBWHCode.Text
        vZoneCode = Me.CMBZone.Text


        If vWHCode <> "" And vZoneCode <> "" And vShelfID <> "" Then

            vQuery = "exec dbo.USP_MB_SearchItemCodeZoneRecProduct '" & vWHCode & "','" & vZoneCode & "','" & vShelfID & "','' "
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                For i = 0 To pds.Tables(0).Rows.Count - 1
                    vItemCode = pds.Tables(0).Rows(i)("itemcode").ToString
                    vItemName = pds.Tables(0).Rows(i)("itemname").ToString
                    vBarCode = pds.Tables(0).Rows(i)("itemcode").ToString
                    vDocDate = pds.Tables(0).Rows(i)("scandatetime").ToString
                    vUnitCode = pds.Tables(0).Rows(i)("unitcode").ToString

                    n = Me.ListViewItem.Items.Count + 1

                    Dim listItem As New ListViewItem(n)
                    listItem.SubItems.Add(vItemCode)
                    listItem.SubItems.Add(vItemName)
                    listItem.SubItems.Add(vShelfID)
                    listItem.SubItems.Add(vWHCode)
                    listItem.SubItems.Add(vZoneCode)
                    listItem.SubItems.Add(vDocDate)
                    listItem.SubItems.Add(1)
                    listItem.SubItems.Add(vUnitCode)
                    Me.ListViewItem.Items.Add(listItem)

                    Me.ListViewItem.Items(i).BackColor = Color.LightGreen
                Next
            End If
        End If
    End Sub

    Private Sub CMBWHCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBWHCode.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.CMBZone.Focus()
        End If
    End Sub

    Private Sub CMBWHCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBWHCode.SelectedIndexChanged
        Call SearchShelfItem()
    End Sub

    Private Sub CMBZone_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBZone.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBShelfID.Focus()
        End If
    End Sub

    Private Sub CMBZone_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBZone.SelectedIndexChanged
        Call SearchShelfItem()
    End Sub

    Private Sub TBShelfID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBShelfID.KeyDown
        Dim vWHCode As String
        Dim vZoneCode As String
        Dim vShelfID As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vBarCode As String
        Dim vDocDate As String
        Dim vUnitCode As String
        Dim vShelfCode As String
        Dim vUserScan As String
        Dim vModeScan As String
        Dim vIsSave As Integer

        Dim i As Integer
        Dim n As Integer
        Dim vCheckShelf As String

        Dim m As Integer
        Dim v As Integer
        Dim vCountNotSave As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then

            'For m = 0 To Me.ListViewItem.Items.Count - 1
            '    If Me.ListViewItem.Items(m).SubItems(7).Text = 0 Then
            '        vCountNotSave = vCountNotSave + 1
            '    End If
            'Next

            'If Me.ListViewItem.Items.Count > 0 Then
            '    For v = 0 To Me.ListViewItem.Items.Count - 1
            '        vItemCode = Trim(ListViewItem.Items(v).SubItems(1).Text)
            '        vBarCode = Trim(ListViewItem.Items(v).SubItems(1).Text)
            '        vItemName = Trim(ListViewItem.Items(v).SubItems(2).Text)
            '        vUnitCode = Trim(ListViewItem.Items(v).SubItems(8).Text)
            '        vWHCode = Trim(ListViewItem.Items(v).SubItems(4).Text)
            '        vZoneCode = Trim(ListViewItem.Items(v).SubItems(5).Text)
            '        vShelfCode = vb6.UCase(Trim(ListViewItem.Items(v).SubItems(3).Text))
            '        vUserScan = vUserName
            '        vModeScan = "หน้ายิงบาร์โค้ดตรวจสอบชั้นเก็บ MC3000"
            '        vIsSave = ListViewItem.Items(v).SubItems(7).Text

            '        If vIsSave = 0 Then
            '            vQuery = "exec dbo.USP_NP_InsertScanItemShelfCode  '" & vItemCode & "','" & vBarCode & "','" & vItemName & "','" & vUnitCode & "','" & vWHCode & "','" & vZoneCode & "','" & vShelfCode & "','" & vUserScan & "','" & vModeScan & "' "
            '            Call vGetData(vMemProfit, vQuery)
            '        End If
            '    Next v
            'End If

            Me.ListViewItem.Items.Clear()
            vWHCode = Me.CMBWHCode.Text
            vZoneCode = Me.CMBZone.Text
            vShelfID = Me.TBShelfID.Text

            vQuery = "exec dbo.USP_MB_SearchShelfMasterDetails '" & vWHCode & "','" & vZoneCode & "','" & vShelfID & "'"
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

            If vCheckShelf <> "" Then

                If vWHCode <> "" And vZoneCode <> "" And vShelfID <> "" Then

                    vQuery = "exec dbo.USP_MB_SearchItemCodeZoneRecProduct '" & vWHCode & "','" & vZoneCode & "','" & vShelfID & "','' "
                    Call vGetData(vMemProfit, vQuery)

                    If pds.Tables(0).Rows.Count > 0 Then
                        For i = 0 To pds.Tables(0).Rows.Count - 1
                            vItemCode = pds.Tables(0).Rows(i)("itemcode").ToString
                            vItemName = pds.Tables(0).Rows(i)("itemname").ToString
                            vBarCode = pds.Tables(0).Rows(i)("itemcode").ToString
                            vDocDate = pds.Tables(0).Rows(i)("scandatetime").ToString
                            vUnitCode = pds.Tables(0).Rows(i)("unitcode").ToString

                            n = Me.ListViewItem.Items.Count + 1

                            Dim listItem As New ListViewItem(n)
                            listItem.SubItems.Add(vItemCode)
                            listItem.SubItems.Add(vItemName)
                            listItem.SubItems.Add(vShelfID)
                            listItem.SubItems.Add(vWHCode)
                            listItem.SubItems.Add(vZoneCode)
                            listItem.SubItems.Add(vDocDate)
                            listItem.SubItems.Add(1)
                            listItem.SubItems.Add(vUnitCode)
                            Me.ListViewItem.Items.Add(listItem)

                            Me.ListViewItem.Items(i).BackColor = Color.LightGreen
                        Next

                    End If
                End If
                Me.TBBarcode.Focus()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBShelfID.Text = ""
            Me.TBBarcode.Text = ""
            Me.ListViewItem.Items.Clear()
            Me.TBShelfID.Focus()
        End If

        If e.KeyCode = Keys.Down Or e.KeyCode = Keys.Right Then
            If Me.TBShelfID.Text <> "" Then
                Call SearchShelfItem()
            End If
            Me.TBBarcode.Focus()
        End If

        If e.KeyCode = Keys.Up Or e.KeyCode = Keys.Left Then
            If Me.TBShelfID.Text <> "" Then
                Call SearchShelfItem()
            End If
            Me.CMBZone.Focus()
        End If
    End Sub

    Private Sub TBShelfID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBShelfID.TextChanged
        Dim vWHCode As String
        Dim vZoneCode As String
        Dim vShelfID As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vBarCode As String
        Dim vDocDate As String
        Dim vUnitCode As String

        Dim i As Integer
        Dim n As Integer

        Dim vCheckShelf As String

        On Error Resume Next

        If vb6.InStr(Me.TBShelfID.Text, "@") > 0 Then
            Me.ListViewItem.Items.Clear()
            vWHCode = Me.CMBWHCode.Text
            vZoneCode = Me.CMBZone.Text
            vShelfID = vb6.Left(Me.TBShelfID.Text, vb6.Len(Me.TBShelfID.Text) - 1)

            Me.TBShelfID.Text = vShelfID

            vQuery = "exec dbo.USP_MB_SearchShelfMasterDetails '" & vWHCode & "','" & vZoneCode & "','" & vShelfID & "'"
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

            If vCheckShelf <> "" Then
                If vWHCode <> "" And vZoneCode <> "" And vShelfID <> "" Then

                    vQuery = "exec dbo.USP_MB_SearchItemCodeZoneRecProduct '" & vWHCode & "','" & vZoneCode & "','" & vShelfID & "','' "
                    Call vGetData(vMemProfit, vQuery)

                    If pds.Tables(0).Rows.Count > 0 Then
                        For i = 0 To pds.Tables(0).Rows.Count - 1
                            vItemCode = pds.Tables(0).Rows(i)("itemcode").ToString
                            vItemName = pds.Tables(0).Rows(i)("itemname").ToString
                            vBarCode = pds.Tables(0).Rows(i)("itemcode").ToString
                            vDocDate = pds.Tables(0).Rows(i)("scandatetime").ToString
                            vUnitCode = pds.Tables(0).Rows(i)("unitcode").ToString

                            n = Me.ListViewItem.Items.Count + 1

                            Dim listItem As New ListViewItem(n)
                            listItem.SubItems.Add(vItemCode)
                            listItem.SubItems.Add(vItemName)
                            listItem.SubItems.Add(vShelfID)
                            listItem.SubItems.Add(vWHCode)
                            listItem.SubItems.Add(vZoneCode)
                            listItem.SubItems.Add(vDocDate)
                            listItem.SubItems.Add(1)
                            listItem.SubItems.Add(vUnitCode)
                            Me.ListViewItem.Items.Add(listItem)

                            Me.ListViewItem.Items(i).BackColor = Color.LightGreen
                        Next
                    End If
                End If
                Me.TBBarcode.Focus()
            End If
        End If
    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        On Error Resume Next
        FormPutAway.Show()
        FormPutAway.PNAddShelfByReceipt.Visible = False
        Me.Hide()
    End Sub

    Public Sub ClearScreen()
        Me.TBShelfID.Text = ""
        Me.TBBarcode.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.TBShelfID.Focus()
    End Sub


    Private Sub BTNRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNRefresh.Click
        Me.TBShelfID.Text = ""
        Me.TBBarcode.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.TBShelfID.Focus()
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

        On Error GoTo ErrDescription

        If Me.ListViewItem.Items.Count > 0 Then
            For i = 0 To Me.ListViewItem.Items.Count - 1
                vItemCode = Trim(ListViewItem.Items(i).SubItems(1).Text)
                vBarCode = Trim(ListViewItem.Items(i).SubItems(1).Text)
                vItemName = Trim(ListViewItem.Items(i).SubItems(2).Text)
                vUnitCode = Trim(ListViewItem.Items(i).SubItems(8).Text)
                vWHCode = Trim(ListViewItem.Items(i).SubItems(4).Text)
                vZoneCode = Trim(ListViewItem.Items(i).SubItems(5).Text)
                vShelfCode = vb6.UCase(Trim(ListViewItem.Items(i).SubItems(3).Text))
                vUserScan = vUserName
                vModeScan = "หน้ายิงบาร์โค้ดตรวจสอบชั้นเก็บ MC3000"
                vIsSave = ListViewItem.Items(i).SubItems(7).Text

                If vIsSave = 0 Then
                    vQuery = "exec dbo.USP_NP_InsertScanItemShelfCode  '" & vItemCode & "','" & vBarCode & "','" & vItemName & "','" & vUnitCode & "','" & vWHCode & "','" & vZoneCode & "','" & vShelfCode & "','" & vUserScan & "','" & vModeScan & "' "
                    Call vGetData(vMemProfit, vQuery)
                End If
            Next i

            Call clearscreen()

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description)
            Exit Sub
        End If
    End Sub

    Private Sub ListViewItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItem.KeyDown
        Dim vIndex As Integer
        Dim vAnswer As Integer
        Dim vItemCode As String
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vZoneCode As String
        Dim vShelfID As String
        Dim vIsSave As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Back And Me.ListViewItem.Items.Count > 0 Then
            vIndex = Me.ListViewItem.FocusedItem.Index
            vItemCode = Me.ListViewItem.Items(vIndex).SubItems(1).Text
            vUnitCode = Me.ListViewItem.Items(vIndex).SubItems(8).Text
            vWHCode = Me.ListViewItem.Items(vIndex).SubItems(4).Text
            vZoneCode = Me.ListViewItem.Items(vIndex).SubItems(5).Text
            vShelfID = Me.ListViewItem.Items(vIndex).SubItems(3).Text
            vIsSave = Me.ListViewItem.Items(vIndex).SubItems(7).Text

            vAnswer = MsgBox("Do you want delete itemcode " & vItemCode & " ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 Then
                Me.ListViewItem.Items.RemoveAt(vIndex)
                Call GenLineNumber()
                vQuery = "exec dbo.USP_NP_DeleteItemShelfIDLogs '" & vItemCode & "','" & vUnitCode & "','" & vWHCode & "','" & vZoneCode & "','" & vShelfID & "'"
                Call vGetData(vMemProfit, vQuery)
                Me.TBBarcode.Focus()
            End If
        End If

        If e.KeyCode = Keys.Up And IsDBNull(Me.ListViewItem.FocusedItem.Index) = 0 And Me.ListViewItem.Items.Count > 0 Then
            Me.TBBarcode.Focus()
            Me.TBBarcode.SelectAll()
        End If
    End Sub

    Public Sub GenLineNumber()
        Dim i As Integer
        Dim n As Integer

        On Error Resume Next

        If Me.ListViewItem.Items.Count > 0 Then

            For i = 0 To Me.ListViewItem.Items.Count - 1
                n = i + 1
                Me.ListViewItem.Items(i).SubItems(0).Text = n
            Next
        End If
    End Sub

    Private Sub ListViewItem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewItem.SelectedIndexChanged

    End Sub
End Class