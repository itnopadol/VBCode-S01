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

Public Class FormCountStockDaily

    Private MyReader As Symbol.Barcode.Reader = Nothing
    Private MyReaderData As Symbol.Barcode.ReaderData = Nothing
    Private MyEventHandler As System.EventHandler = Nothing

    Dim vQuery As String
    Dim vMemDocDate As String
    Dim vGetNow As Date
    Dim vMemYear As Integer
    Dim vMemMonth As Integer
    Dim vGenDocNo As String
    Dim vGenInspectNo As String

    Dim vMemGetInspectNo As String

    Private Sub FormCountStockDaily_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        'If (Me.InitReader()) Then
        '    Me.StartRead()
        'Else
        '    Me.Close()
        '    Return
        'End If

        Me.PNSelectStkType.Visible = True
        Me.PNSelectStkType.BringToFront()

        Call GetCauseProductNegative()
        Call vGetWareHouse()
        Call vGetShelf()

        vQuery = "exec dbo.USP_NP_CheckDocDate"
        Call vGetData1(vMemProfit, vQuery)
        If pds1.Tables(0).Rows.Count > 0 Then
            vMemDocDate = pds1.Tables(0).Rows(0)("vdocdate").ToString
        End If

        Dim vLenDate As Integer
        Dim vCheckSlate1 As Integer
        Dim vCheckSlate2 As Integer
        Dim vGetCutMonth1 As String
        Dim vGetCutMonth2 As String

        vLenDate = vb6.Len(vMemDocDate)
        vCheckSlate1 = vb6.InStr(vMemDocDate, "/")
        vGetCutMonth1 = vb6.Right(vMemDocDate, vLenDate - vCheckSlate1)
        vCheckSlate2 = vb6.InStr(vGetCutMonth1, "/")
        vGetCutMonth2 = vb6.Left(vGetCutMonth1, vCheckSlate2 - 1)
        vMemYear = vb6.Right(vMemDocDate, 4)
        vMemMonth = vGetCutMonth2

        Me.RDBDay.Focus()
    End Sub

    Private Function InitReader() As Boolean

        On Error Resume Next

        If Not (Me.MyReader Is Nothing) Then
            Return False
        End If

        Me.MyReader = New Symbol.Barcode.Reader
        Me.MyReaderData = New Symbol.Barcode.ReaderData( _
                                     Symbol.Barcode.ReaderDataTypes.Text, _
                                     Symbol.Barcode.ReaderDataLengths.MaximumLabel)
        Me.MyEventHandler = New System.EventHandler(AddressOf MyReader_ReadNotify)
        Me.MyReader.Actions.Enable()

        AddHandler Me.Activated, New EventHandler(AddressOf ReaderForm_Activated)
        AddHandler Me.Deactivate, New EventHandler(AddressOf ReaderForm_Deactivate)

        Return True
    End Function

    Private Sub ReaderForm_Activated(ByVal sender As Object, ByVal e As EventArgs)
        On Error Resume Next

        If Not (Me.MyReaderData.IsPending) Then
            Me.StartRead()
        End If
    End Sub

    Private Sub ReaderForm_Deactivate(ByVal sender As Object, ByVal e As EventArgs)
        Me.StopRead()
    End Sub

    Private Sub MyReader_ReadNotify(ByVal o As Object, ByVal e As EventArgs)
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vBarCode As String

        Dim vStkUnit As String
        Dim vStockQty As Double

        Dim i As Integer
        Dim n As Integer
        Dim a As Integer
        Dim b As Integer
        Dim vSumQty As Double
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vCheckShelfID As String
        Dim vCheckWHCode As String
        Dim vCheckShelfCode As String

        Dim TheReaderData As Symbol.Barcode.ReaderData = Me.MyReader.GetNextReaderData()

        On Error Resume Next

        If (TheReaderData.Result = Symbol.Results.SUCCESS) Then
            Me.TBBarCode.Text = TheReaderData.Text
            Me.StartRead()

            vBarCode = TheReaderData.Text
            Me.ListViewStock.Items.Clear()
            Me.ListViewShelfID.Items.Clear()

            vQuery = "exec dbo.usp_np_DataItemDetails '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBUnitCode.Text = vUnitCode

                vSumQty = 0

                For i = 0 To pds.Tables(0).Rows.Count - 1
                    vWHCode = pds.Tables(0).Rows(i)("whcode").ToString
                    vShelfCode = pds.Tables(0).Rows(i)("shelfcode").ToString
                    vStkUnit = pds.Tables(0).Rows(i)("defstkunitcode").ToString
                    vStockQty = pds.Tables(0).Rows(i)("qty").ToString

                    If Me.ListViewStock.Items.Count > 0 Then
                        For n = 0 To Me.ListViewStock.Items.Count - 1
                            vCheckWHCode = Me.ListViewStock.Items(n).SubItems(0).Text
                            vCheckShelfCode = Me.ListViewStock.Items(n).SubItems(1).Text

                            If vWHCode = vCheckWHCode And vShelfCode = vCheckShelfCode Then
                                GoTo Line1
                            End If
                        Next
                    End If

                    If vWHCode <> "" And vShelfCode <> "" Then
                        Dim listItem As New ListViewItem(vWHCode)
                        listItem.SubItems.Add(vShelfCode)
                        listItem.SubItems.Add(Format(vStockQty, "##,##0.00"))
                        listItem.SubItems.Add(vStkUnit)
                        Me.ListViewStock.Items.Add(listItem)

                        If vWHCode = vMemProfit Then
                            vSumQty = vSumQty + vStockQty
                        End If
                    End If
                    Me.TBRemainQty.Text = Format(vSumQty, "##,##0.00")

Line1:
                Next

                For a = 0 To pds.Tables(0).Rows.Count - 1
                    vShelfID = pds.Tables(0).Rows(a)("shelfid").ToString
                    If Me.ListViewShelfID.Items.Count > 0 Then
                        For b = 0 To Me.ListViewShelfID.Items.Count - 1
                            vCheckShelfID = Me.ListViewShelfID.Items(b).SubItems(0).Text

                            If vShelfID = vCheckShelfID Then
                                GoTo Line2
                            End If
                        Next
                    End If

                    If vShelfID <> "" Then
                        Dim listItem As New ListViewItem(vShelfID)
                        Me.ListViewShelfID.Items.Add(listItem)
                    End If

Line2:
                Next

                Me.TBQty.Focus()

            Else
                Me.TBBarcode.Text = ""
                Me.TBBarcode.Focus()
                Me.TBBarcode.SelectAll()
                MsgBox("This barcode find not found !", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If

        End If
    End Sub

    Private Sub StartRead()
        If Not ((Me.MyReader Is Nothing) And (Me.MyReaderData Is Nothing)) Then
            AddHandler MyReader.ReadNotify, Me.MyEventHandler
            Me.MyReader.Actions.Read(Me.MyReaderData)
        End If
    End Sub

    Private Sub StopRead()
        If Not (Me.MyReader Is Nothing) Then
            RemoveHandler MyReader.ReadNotify, Me.MyEventHandler
            Me.MyReader.Actions.Flush()
        End If
    End Sub

    Public Sub GetCauseProductNegative()
        Dim i As Integer
        Dim vCauseName As String

        On Error Resume Next

        vQuery = "exec dbo.USP_MB_SearchCauseProductNegative"
        Call vGetData(vMemProfit, vQuery)

        If pds.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds.Tables(0).Rows.Count - 1
                vCauseName = pds.Tables(0).Rows(i)("causename").ToString
                Me.CMBReason.Items.Add(vCauseName)
            Next
        End If

        If Me.CMBReason.Items.Count > 0 Then
            Me.CMBReason.SelectedIndex = 0
        End If
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
                Me.CMBShelfCode.Items.Add(vShelfCode)
            Next
        End If

        If Me.CMBShelfCode.Items.Count > 0 Then
            Me.CMBShelfCode.SelectedIndex = 0
        End If
    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        FormMainApplication.Show()
        Me.Hide()
    End Sub

    Private Sub CMBShelfCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBShelfCode.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.CMBReason.Focus()
        End If
    End Sub

    Private Sub CMBShelfCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBShelfCode.SelectedIndexChanged

    End Sub

    Private Sub CMBReason_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBReason.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBBarCode.Focus()
        End If
    End Sub

    Private Sub CMBReason_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBReason.SelectedIndexChanged

    End Sub

    Private Sub TBBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vBarCode As String

        Dim vStkUnit As String
        Dim vStockQty As Double

        Dim i As Integer
        Dim n As Integer
        Dim a As Integer
        Dim b As Integer
        Dim vSumQty As Double
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vCheckShelfID As String
        Dim vCheckWHCode As String
        Dim vCheckShelfCode As String
        Dim vMemShelfCode As String
        Dim vMemWHCode As String
        Dim x As Integer
        Dim vLineWHCode As String
        Dim vLineShelfCode As String


        On Error Resume Next


        If e.KeyCode = Keys.Enter Then

            Me.ListViewStock.Items.Clear()
            Me.ListViewShelfID.Items.Clear()

            vBarCode = Me.TBBarCode.Text
            vMemWHCode = Me.CMBWHCode.Text
            vMemShelfCode = Me.CMBShelfCode.Text

            vQuery = "exec dbo.usp_np_DataItemDetails '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBUnitCode.Text = vUnitCode
                Me.TBMemBar.Text = vBarCode

                vSumQty = 0

                For i = 0 To pds.Tables(0).Rows.Count - 1
                    vWHCode = pds.Tables(0).Rows(i)("whcode").ToString
                    vShelfCode = pds.Tables(0).Rows(i)("shelfcode").ToString
                    vStkUnit = pds.Tables(0).Rows(i)("defstkunitcode").ToString
                    vStockQty = pds.Tables(0).Rows(i)("qty").ToString

                    If vWHCode = vMemProfit And Me.CMBShelfCode.Text = vShelfCode Then
                        Me.TBRemainQty.Text = Format(vStockQty, "##,##0.00")
                    End If

                    If Me.ListViewStock.Items.Count > 0 Then
                        For n = 0 To Me.ListViewStock.Items.Count - 1
                            vCheckWHCode = Me.ListViewStock.Items(n).SubItems(0).Text
                            vCheckShelfCode = Me.ListViewStock.Items(n).SubItems(1).Text

                            If vWHCode = vCheckWHCode And vShelfCode = vCheckShelfCode Then
                                GoTo Line1
                            End If
                        Next
                    End If

                    If vWHCode <> "" And vShelfCode <> "" Then
                        Dim listItem As New ListViewItem(vWHCode)
                        listItem.SubItems.Add(vShelfCode)
                        listItem.SubItems.Add(Format(vStockQty, "##,##0.00"))
                        listItem.SubItems.Add(vStkUnit)
                        Me.ListViewStock.Items.Add(listItem)

                        For x = 0 To Me.ListViewStock.Items.Count - 1
                            vLineWHCode = Me.ListViewStock.Items(x).SubItems(0).Text
                            vLineShelfCode = Me.ListViewStock.Items(x).SubItems(1).Text

                            If vMemWHCode = vLineWHCode And vMemShelfCode = vLineShelfCode Then
                                Me.ListViewStock.Items(x).ForeColor = Color.DarkGreen
                            End If
                        Next

                    End If

Line1:
                Next

                For a = 0 To pds.Tables(0).Rows.Count - 1
                    vShelfID = pds.Tables(0).Rows(a)("shelfid").ToString
                    If Me.ListViewShelfID.Items.Count > 0 Then
                        For b = 0 To Me.ListViewShelfID.Items.Count - 1
                            vCheckShelfID = Me.ListViewShelfID.Items(b).SubItems(0).Text

                            If vShelfID = vCheckShelfID Then
                                GoTo Line2
                            End If
                        Next
                    End If

                    If vShelfID <> "" Then
                        Dim listItem As New ListViewItem(vShelfID)
                        Me.ListViewShelfID.Items.Add(listItem)
                    End If

Line2:
                Next

                Me.ListViewStock.Visible = True
                Me.ListViewShelfID.Visible = True
                Me.PNItem.Visible = True
                Me.PNItem.BringToFront()
                Me.TBQty.Focus()

            Else
                Me.TBBarCode.Text = ""
                Me.ListViewStock.Visible = False
                Me.ListViewShelfID.Visible = False
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                MsgBox("This barcode find not found !", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If

        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearItem()
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBQty.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.CMBReason.Focus()
        End If
    End Sub

    Private Sub TBBarCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vBarCode As String

        Dim vStkUnit As String
        Dim vStockQty As Double

        Dim i As Integer
        Dim n As Integer
        Dim a As Integer
        Dim b As Integer
        Dim vSumQty As Double
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vCheckShelfID As String
        Dim vCheckWHCode As String
        Dim vCheckShelfCode As String

        On Error Resume Next

        If vb6.InStr(Me.TBBarCode.Text, "@") > 0 Then

            vBarCode = vb6.Left(Me.TBBarCode.Text, vb6.Len(Me.TBBarCode.Text) - 1)

            Me.TBBarCode.Text = vBarCode

            Me.ListViewStock.Items.Clear()
            Me.ListViewShelfID.Items.Clear()

            vQuery = "exec dbo.usp_np_DataItemDetails '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBUnitCode.Text = vUnitCode
                Me.TBMemBar.Text = vBarCode

                vSumQty = 0

                For i = 0 To pds.Tables(0).Rows.Count - 1
                    vWHCode = pds.Tables(0).Rows(i)("whcode").ToString
                    vShelfCode = pds.Tables(0).Rows(i)("shelfcode").ToString
                    vStkUnit = pds.Tables(0).Rows(i)("defstkunitcode").ToString
                    vStockQty = pds.Tables(0).Rows(i)("qty").ToString

                    If vWHCode = vMemProfit And Me.CMBShelfCode.Text = vShelfCode Then
                        Me.TBRemainQty.Text = Format(vStockQty, "##,##0.00")
                    End If

                    If Me.ListViewStock.Items.Count > 0 Then
                        For n = 0 To Me.ListViewStock.Items.Count - 1
                            vCheckWHCode = Me.ListViewStock.Items(n).SubItems(0).Text
                            vCheckShelfCode = Me.ListViewStock.Items(n).SubItems(1).Text

                            If vWHCode = vCheckWHCode And vShelfCode = vCheckShelfCode Then
                                GoTo Line1
                            End If
                        Next
                    End If

                    If vWHCode <> "" And vShelfCode <> "" Then
                        Dim listItem As New ListViewItem(vWHCode)
                        listItem.SubItems.Add(vShelfCode)
                        listItem.SubItems.Add(Format(vStockQty, "##,##0.00"))
                        listItem.SubItems.Add(vStkUnit)
                        Me.ListViewStock.Items.Add(listItem)
                    End If

Line1:
                Next

                For a = 0 To pds.Tables(0).Rows.Count - 1
                    vShelfID = pds.Tables(0).Rows(a)("shelfid").ToString
                    If Me.ListViewShelfID.Items.Count > 0 Then
                        For b = 0 To Me.ListViewShelfID.Items.Count - 1
                            vCheckShelfID = Me.ListViewShelfID.Items(b).SubItems(0).Text

                            If vShelfID = vCheckShelfID Then
                                GoTo Line2
                            End If
                        Next
                    End If

                    If vShelfID <> "" Then
                        Dim listItem As New ListViewItem(vShelfID)
                        Me.ListViewShelfID.Items.Add(listItem)
                    End If

Line2:
                Next

                Me.ListViewStock.Visible = True
                Me.ListViewShelfID.Visible = True
                Me.PNItem.Visible = True
                Me.PNItem.BringToFront()
                Me.TBQty.Focus()

            ElseIf Me.TBBarCode.Text = "" Then
                Call ClearItem()
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                Exit Sub
            End If

        End If
    End Sub

    Private Sub TBQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBQty.KeyDown
        Dim i As Integer
        Dim n As Integer
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vReasonCode As String

        Dim vQty As Double
        Dim vRemainQty As Double

        Dim vChkItemCode As String
        Dim vChkUnitCode As String
        Dim vChkWHCode As String
        Dim vChkShelfCode As String

        Dim vAnswer As Integer
        Dim vAnswer1 As Integer

        Dim vOldQty As Double
        Dim vAddQty As Double


        If e.KeyCode = Keys.Enter And Me.TBQty.Text <> "" And Me.TBItemCode.Text <> "" Then

            vItemCode = Me.TBItemCode.Text
            vBarCode = Me.TBBarCode.Text
            vItemName = Me.TBItemName.Text
            vUnitCode = Me.TBUnitCode.Text
            vWHCode = Me.CMBWHCode.Text
            vShelfCode = Me.CMBShelfCode.Text
            If Me.TBQty.Text <> "" Then
                vQty = Me.TBQty.Text
            Else
                vQty = 0
            End If
            If Me.TBRemainQty.Text <> "" Then
                vRemainQty = Me.TBRemainQty.Text
            Else
                vRemainQty = 0
            End If
            vReasonCode = vb6.Left(Me.CMBReason.Text, InStr(Me.CMBReason.Text, "//") - 1)

            If Me.ListViewItem.Items.Count > 0 Then

                For i = 0 To Me.ListViewItem.Items.Count - 1
                    vChkItemCode = Me.ListViewItem.Items(i).SubItems(1).Text
                    vChkUnitCode = Me.ListViewItem.Items(i).SubItems(4).Text
                    vChkWHCode = Me.ListViewItem.Items(i).SubItems(6).Text
                    vChkShelfCode = Me.ListViewItem.Items(i).SubItems(7).Text
                    vOldQty = Me.ListViewItem.Items(i).SubItems(3).Text

                    If vItemCode = vChkItemCode And vUnitCode = vChkUnitCode And vWHCode = vChkWHCode And vShelfCode = vChkShelfCode And vMemInspectIsOpen = 0 Then
                        Me.ListViewItem.Items(i).Selected = True
                        Me.ListViewItem.Items(i).Focused = True

                        Me.PNItem.Visible = False
                        Me.PNCheckCount.Visible = True
                        Me.PNCheckCount.BringToFront()
                        Me.TBADDNewQty.Focus()
                        Me.TBADDBarCode.Text = Me.TBMemBar.Text
                        Me.TBADDItemCode.Text = Me.TBItemCode.Text
                        Me.TBADDItemName.Text = Me.TBItemName.Text
                        Me.TBADDStkQty.Text = Me.TBRemainQty.Text
                        Me.TBADDUnit.Text = Me.TBUnitCode.Text
                        Me.TBADDOldQty.Text = Format(vOldQty, "##,##0.00")
                        Me.TBADDNewQty.Text = Me.TBQty.Text
                        Me.TBADDId.Text = i
                        Me.TBADDNewQty.Focus()
                        Exit Sub

                        'vAnswer = MsgBox("This item is exist at line " & i + 1 & ". Do you want edit qty ?", MsgBoxStyle.YesNo, "Send Question Message")
                        'If vAnswer = 6 Then
                        '    vAnswer1 = MsgBox("Click Yes is Add Qty or Click No is Edit Qty", MsgBoxStyle.YesNo, "Send Question Message")
                        '    If vAnswer1 = 6 Then
                        '        vAddQty = vQty + vOldQty
                        '        Me.ListViewItem.Items(i).SubItems(2).Text = Format(vRemainQty, "##,##0.00")
                        '        Me.ListViewItem.Items(i).SubItems(3).Text = Format(vAddQty, "##,##0.00")
                        '        Me.PNItem.Visible = False
                        '        Call ClearItem()

                        '        Exit Sub
                        '    Else
                        '        Me.ListViewItem.Items(i).SubItems(2).Text = Format(vRemainQty, "##,##0.00")
                        '        Me.ListViewItem.Items(i).SubItems(3).Text = Format(vQty, "##,##0.00")
                        '        Me.PNItem.Visible = False
                        '        Call ClearItem()
                        '        Exit Sub
                        '    End If
                        'Else
                        '    Me.PNItem.Visible = False
                        '    Call ClearItem()
                        '    Exit Sub
                        'End If
                    End If

                    If vItemCode = vChkItemCode And vUnitCode = vChkUnitCode And vWHCode = vChkWHCode And vShelfCode = vChkShelfCode And vMemInspectIsOpen = 1 Then

                        Me.PNItem.Visible = False
                        Me.PNCheckCount.Visible = True
                        Me.PNCheckCount.BringToFront()
                        Me.TBADDNewQty.Focus()
                        Me.TBADDBarCode.Text = Me.TBMemBar.Text
                        Me.TBADDItemCode.Text = Me.TBItemCode.Text
                        Me.TBADDItemName.Text = Me.TBItemName.Text
                        Me.TBADDStkQty.Text = Me.TBRemainQty.Text
                        Me.TBADDUnit.Text = Me.TBUnitCode.Text
                        Me.TBADDOldQty.Text = Format(vOldQty, "##,##0.00")
                        Me.TBADDNewQty.Text = Me.TBQty.Text
                        Me.TBADDId.Text = i
                        Me.TBADDNewQty.Focus()
                        Exit Sub

                        'vAnswer = MsgBox("Do you want edit qty ?", MsgBoxStyle.YesNo, "Send Question Message")
                        'If vAnswer = 6 Then
                        '    vAnswer1 = MsgBox("Click Yes is Add Qty or Click No is Edit Qty", MsgBoxStyle.YesNo, "Send Question Message")
                        '    If vAnswer1 = 6 Then
                        '        vAddQty = vQty + vOldQty
                        '        Me.ListViewItem.Items(i).SubItems(3).Text = Format(vAddQty, "##,##0.00")
                        '        Me.ListViewItem.Items(i).SubItems(10).Text = 0
                        '        Me.PNItem.Visible = False
                        '        Call ClearItem()
                        '        Exit Sub
                        '    Else
                        '        Me.ListViewItem.Items(i).SubItems(3).Text = Format(vQty, "##,##0.00")
                        '        Me.ListViewItem.Items(i).SubItems(10).Text = 0
                        '        Me.PNItem.Visible = False
                        '        Call ClearItem()
                        '        Exit Sub
                        '    End If

                        'Else
                        '    Me.PNItem.Visible = False
                        '    Call ClearItem()
                        '    Exit Sub
                        'End If
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
            listItem.SubItems.Add(vWHCode)
            listItem.SubItems.Add(vShelfCode)
            listItem.SubItems.Add(vReasonCode)
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

    Public Sub ClearScreen()
        'On Error Resume Next

        vQuery = "exec dbo.USP_NP_CheckDocDate"
        Call vGetData1(vMemProfit, vQuery)
        If pds1.Tables(0).Rows.Count > 0 Then
            vMemDocDate = pds1.Tables(0).Rows(0)("vdocdate").ToString
        End If

        Me.TBDocNo.Text = ""
        vMemGetInspectNo = ""
        Me.TBDocDate.Text = vMemDocDate

        vMemInspectIsOpen = 0
        Me.TBMemBar.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBBarCode.Text = ""
        Me.TBUnitCode.Text = ""
        Me.CMBWHCode.SelectedIndex = 0
        Me.CMBShelfCode.SelectedIndex = 0
        Me.CMBReason.SelectedIndex = 0
        Me.TBQty.Text = ""
        Me.TBRemainQty.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.ListViewStock.Items.Clear()
        Me.ListViewShelfID.Clear()
        Me.ListViewStock.Visible = False
        Me.ListViewShelfID.Visible = False
        Me.CMBReason.SelectedIndex = 0

        Me.PNSelectStkType.Visible = True
        Me.RDBDay.Focus()
    End Sub

    Public Sub ClearItem()
        Me.TBMemBar.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBBarCode.Text = ""
        Me.TBUnitCode.Text = ""
        Me.TBQty.Text = ""
        Me.TBRemainQty.Text = ""
        Me.ListViewShelfID.Clear()
        Me.ListViewStock.Visible = False
        Me.ListViewShelfID.Visible = False
        Me.TBBarCode.Focus()
    End Sub

    Private Sub TBQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBQty.KeyPress
        On Error GoTo ErrDescription

        Select Case Asc(e.KeyChar)
            Case 48 To 58, 8, 44, 45, 46, 37, 64
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
        If vb6.InStr(Me.TBQty.Text, "@") > 0 Then
            Me.TBQty.Text = ""
            Me.TBQty.Focus()
        End If
    End Sub

    Private Sub BTNSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSaveData.Click
        Dim vQuery As String
        Dim vItem(500), vUnitCode(500), vShelf, vItemName(500) As String
        Dim vQty(500), vDiff(500), vInspectQTY(500) As Double
        Dim vCountItem As Double
        Dim vSumItem(500), i, j As Double
        Dim vItemCode(500) As String
        Dim vShelfCode(500) As String
        Dim vWHCode(500) As String
        Dim vInSpectDesc(500) As String
        Dim x As Integer
        Dim a As Integer

        Dim n As Integer

        Dim vUnit, vWH As String
        Dim vLineNumber As Integer
        Dim vShelfStock As String
        Dim vReasonCode As String

        Dim vItmCode As String
        Dim vItmName As String
        Dim vItmQty As Double
        Dim vItmDiff As Double
        Dim vItmInspectQty As Double

        Dim vIsSave As Integer
        Dim vDocNo As String
        Dim vDocDate As String

        On Error Resume Next

        If vMemInspectIsOpen = 0 Then
            If Me.ListViewItem.Items.Count > 0 Then

                For i = 0 To Me.ListViewItem.Items.Count - 1

                    vLineNumber = i
                    vShelfStock = Trim(ListViewItem.Items(i).SubItems(7).Text)
                    vItmCode = Me.ListViewItem.Items(i).SubItems(1).Text
                    vItmName = Me.ListViewItem.Items(i).SubItems(11).Text
                    vWH = Trim(ListViewItem.Items(i).SubItems(6).Text)
                    vShelf = Me.ListViewItem.Items(i).SubItems(7).Text
                    vUnit = Me.ListViewItem.Items(i).SubItems(4).Text
                    vItmQty = Me.ListViewItem.Items(i).SubItems(3).Text
                    vItmInspectQty = Me.ListViewItem.Items(i).SubItems(2).Text
                    vItmDiff = vItmInspectQty - vItmQty
                    vReasonCode = Me.ListViewItem.Items(i).SubItems(8).Text

                    vQuery = "exec dbo.USP_NP_InsertInspectLog '','" & vItmCode & "','" & vItmName & "','" & vWH & "','" & vShelf & "'," & vItmQty & ",'" & vUnit & "','" & vUserName & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
                    Call vGetData(vMemProfit, vQuery)

                Next

                If Me.TBDocNo.Text = "" And vMemInspectIsOpen = 0 Then
                    Call vGetDocNo()
                End If

                If Me.ListViewItem.Items.Count > 0 Then
                    vQuery = "exec dbo.USP_NP_UpdateInspectNoLog '" & vGenDocNo & "','" & vUserName & "' "
                    Call vGetData(vMemProfit, vQuery)
                End If

                vQuery = "select count(itemcode) as countitem from (select distinct itemcode,whcode,stockshelf from npmaster.dbo.TB_IV_StkInspect_Logs where docno = '" & vGenDocNo & "') as a "
                Call vGetData(vMemProfit, vQuery)

                If pds.Tables(0).Rows.Count > 0 Then
                    vCountItem = pds.Tables(0).Rows(0)("countitem").ToString
                End If

                vQuery = "exec dbo.USP_HH_InsertBCSTKInspect '" & vGenInspectNo & "','" & vUserName & "' "
                Call vGetData1(vMemProfit, vQuery)

                vQuery = "exec dbo.USP_NP_SelectItemInspect '" & vGenDocNo & "' "
                Call vGetData2(vMemProfit, vQuery)

                If pds2.Tables(0).Rows.Count > 0 Then
                    j = 0
                    For n = 0 To pds2.Tables(0).Rows.Count - 1
                        vItemCode(j) = pds2.Tables(0).Rows(n)("itemcode").ToString
                        vWHCode(j) = pds2.Tables(0).Rows(n)("whcode").ToString
                        vShelfCode(j) = pds2.Tables(0).Rows(n)("stockshelf").ToString
                        j = j + 1
                    Next
                End If

                For a = 0 To vCountItem - 1
                    vQuery = "exec dbo.USP_NP_SelectItemDetailsInspect '" & vGenDocNo & "' , '" & vItemCode(a) & "' ,'" & vShelfCode(a) & "' ,'" & vWHCode(a) & "' "
                    Call vGetData3(vMemProfit, vQuery)

                    If pds3.Tables(0).Rows.Count > 0 Then
                        vItemName(a) = pds3.Tables(0).Rows(0)("itemname").ToString
                        vUnitCode(a) = pds3.Tables(0).Rows(0)("unitcode").ToString
                        vInSpectDesc(a) = pds3.Tables(0).Rows(0)("reasoncode").ToString
                    End If

                    vQuery = "exec dbo.USP_NP_SelectSumItemQtyInspect '" & vGenDocNo & "','" & vItemCode(a) & "','" & vWHCode(a) & "','" & vShelfCode(a) & "' "
                    Call vGetData4(vMemProfit, vQuery)

                    If pds4.Tables(0).Rows.Count > 0 Then
                        vSumItem(a) = pds4.Tables(0).Rows(0)("qty").ToString
                    Else
                        vSumItem(a) = 0
                    End If

                    vQuery = "exec dbo.USP_NP_SelectItemQtySTKLocation '" & vItemCode(a) & "','" & vWHCode(a) & "','" & vShelfCode(a) & "' "
                    Call vGetData5(vMemProfit, vQuery)

                    If pds5.Tables(0).Rows.Count > 0 Then
                        vInspectQTY(a) = pds5.Tables(0).Rows(0)("qty").ToString
                    End If
                    vDiff(a) = vSumItem(a) - vInspectQTY(a)
                Next

                For x = 0 To vCountItem - 1
                    vQuery = "exec dbo.USP_NP_InsertBCInspectSub '" & vGenInspectNo & "','" & vItemCode(x) & "','" & vUnitCode(x) & "','" & vWHCode(x) & "','" & vShelfCode(x) & "'," & vInspectQTY(x) & "," & vSumItem(x) & "," & vDiff(x) & ",'" & vInSpectDesc(x) & "' "
                    Call vGetData6(vMemProfit, vQuery)
                Next

                vQuery = "Update npmaster.dbo.TB_IV_StkInspect_Logs set Inspectno = '" & vGenInspectNo & "' where docno = '" & vGenDocNo & "' "
                Call vGetData7(vMemProfit, vQuery)
                MsgBox("Save data is complete docno is  " & vGenInspectNo & " ")
                Call ClearScreen()
            Else
                MsgBox("Can not Save data.Item not exist in table")
            End If

        ElseIf vMemInspectIsOpen = 1 Then
            vDocNo = Me.TBDocNo.Text
            vDocDate = Me.TBDocDate.Text

            If Me.ListViewItem.Items.Count > 0 Then

                For i = 0 To Me.ListViewItem.Items.Count - 1

                    vLineNumber = i
                    vShelfStock = Trim(ListViewItem.Items(i).SubItems(7).Text)
                    vItmCode = Me.ListViewItem.Items(i).SubItems(1).Text
                    vItmName = Me.ListViewItem.Items(i).SubItems(11).Text
                    vWH = Trim(ListViewItem.Items(i).SubItems(6).Text)
                    vShelf = Me.ListViewItem.Items(i).SubItems(7).Text
                    vUnit = Me.ListViewItem.Items(i).SubItems(4).Text
                    vItmQty = Me.ListViewItem.Items(i).SubItems(3).Text
                    vItmInspectQty = Me.ListViewItem.Items(i).SubItems(2).Text
                    vItmDiff = vItmInspectQty - vItmQty
                    vReasonCode = Me.ListViewItem.Items(i).SubItems(8).Text
                    vIsSave = Me.ListViewItem.Items(i).SubItems(10).Text

                    If vIsSave = 0 Then
                        vQuery = "exec dbo.USP_NP_InsertBCInspectSub '" & vDocNo & "','" & vItmCode & "','" & vUnit & "','" & vWH & "','" & vShelf & "'," & vItmInspectQty & "," & vItmQty & "," & vItmDiff & ",'" & vReasonCode & "' "
                        Call vGetData6(vMemProfit, vQuery)

                        vQuery = "exec dbo.USP_NP_InsertInspectLog '" & vMemGetInspectNo & "','" & vItmCode & "','" & vItmName & "','" & vWH & "','" & vShelf & "'," & vItmQty & ",'" & vUnit & "','" & vUserName & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
                        Call vGetData(vMemProfit, vQuery)

                        vQuery = "Update npmaster.dbo.TB_IV_StkInspect_Logs set Inspectno = '" & vDocNo & "' where docno = '" & vMemGetInspectNo & "' "
                        Call vGetData7(vMemProfit, vQuery)

                        'vQuery = "exec dbo.USP_NP_ '" & vMemGetInspectNo & "'"
                        'Call vGetData(vMemProfit, vQuery)
                    End If

                Next
                MsgBox("Save data is complete docno is  " & vGenInspectNo & " ")
                Call ClearScreen()
            End If

        End If

    End Sub

    Public Sub SaveData()
        Dim vQuery As String
        Dim vItem(500), vUnitCode(500), vShelf, vItemName(500) As String
        Dim vQty(500), vDiff(500), vInspectQTY(500) As Double
        Dim vCountItem As Double
        Dim vSumItem(500), i, j As Double
        Dim vItemCode(500) As String
        Dim vShelfCode(500) As String
        Dim vWHCode(500) As String
        Dim vInSpectDesc(500) As String
        Dim x As Integer
        Dim a As Integer

        Dim n As Integer

        Dim vUnit, vWH As String
        Dim vLineNumber As Integer
        Dim vShelfStock As String
        Dim vReasonCode As String

        Dim vItmCode As String
        Dim vItmName As String
        Dim vItmQty As Double
        Dim vItmDiff As Double
        Dim vItmInspectQty As Double

        Dim vIsSave As Integer
        Dim vDocNo As String
        Dim vDocDate As String


        On Error Resume Next

        If vMemInspectIsOpen = 0 Then
            If Me.ListViewItem.Items.Count > 0 Then

                For i = 0 To Me.ListViewItem.Items.Count - 1

                    vLineNumber = i
                    vShelfStock = Trim(ListViewItem.Items(i).SubItems(7).Text)
                    vItmCode = Me.ListViewItem.Items(i).SubItems(1).Text
                    vItmName = Me.ListViewItem.Items(i).SubItems(11).Text
                    vWH = Trim(ListViewItem.Items(i).SubItems(6).Text)
                    vShelf = Me.ListViewItem.Items(i).SubItems(7).Text
                    vUnit = Me.ListViewItem.Items(i).SubItems(4).Text
                    vItmQty = Me.ListViewItem.Items(i).SubItems(3).Text
                    vItmInspectQty = Me.ListViewItem.Items(i).SubItems(2).Text
                    vItmDiff = vItmInspectQty - vItmQty
                    vReasonCode = Me.ListViewItem.Items(i).SubItems(8).Text

                    vQuery = "exec dbo.USP_NP_InsertInspectLog '','" & vItmCode & "','" & vItmName & "','" & vWH & "','" & vShelf & "'," & vItmQty & ",'" & vUnit & "','" & vUserName & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
                    Call vGetData(vMemProfit, vQuery)

                Next

                If Me.TBDocNo.Text = "" And vMemInspectIsOpen = 0 Then
                    Call vGetDocNo()
                End If

                If Me.ListViewItem.Items.Count > 0 Then
                    vQuery = "exec dbo.USP_NP_UpdateInspectNoLog '" & vGenDocNo & "','" & vUserName & "' "
                    Call vGetData(vMemProfit, vQuery)
                End If

                vQuery = "select count(itemcode) as countitem from (select distinct itemcode,whcode,stockshelf from npmaster.dbo.TB_IV_StkInspect_Logs where docno = '" & vGenDocNo & "') as a "
                Call vGetData(vMemProfit, vQuery)

                If pds.Tables(0).Rows.Count > 0 Then
                    vCountItem = pds.Tables(0).Rows(0)("countitem").ToString
                End If

                vQuery = "exec dbo.USP_HH_InsertBCSTKInspect '" & vGenInspectNo & "','" & vUserName & "' "
                Call vGetData1(vMemProfit, vQuery)

                vQuery = "exec dbo.USP_NP_SelectItemInspect '" & vGenDocNo & "' "
                Call vGetData2(vMemProfit, vQuery)

                If pds2.Tables(0).Rows.Count > 0 Then
                    j = 0
                    For n = 0 To pds2.Tables(0).Rows.Count - 1
                        vItemCode(j) = pds2.Tables(0).Rows(n)("itemcode").ToString
                        vWHCode(j) = pds2.Tables(0).Rows(n)("whcode").ToString
                        vShelfCode(j) = pds2.Tables(0).Rows(n)("stockshelf").ToString
                        j = j + 1
                    Next
                End If

                For a = 0 To vCountItem - 1
                    vQuery = "exec dbo.USP_NP_SelectItemDetailsInspect '" & vGenDocNo & "' , '" & vItemCode(a) & "' ,'" & vShelfCode(a) & "' ,'" & vWHCode(a) & "' "
                    Call vGetData3(vMemProfit, vQuery)

                    If pds3.Tables(0).Rows.Count > 0 Then
                        vItemName(a) = pds3.Tables(0).Rows(0)("itemname").ToString
                        vUnitCode(a) = pds3.Tables(0).Rows(0)("unitcode").ToString
                        vInSpectDesc(a) = pds3.Tables(0).Rows(0)("reasoncode").ToString
                    End If

                    vQuery = "exec dbo.USP_NP_SelectSumItemQtyInspect '" & vGenDocNo & "','" & vItemCode(a) & "','" & vWHCode(a) & "','" & vShelfCode(a) & "' "
                    Call vGetData4(vMemProfit, vQuery)

                    If pds4.Tables(0).Rows.Count > 0 Then
                        vSumItem(a) = pds4.Tables(0).Rows(0)("qty").ToString
                    Else
                        vSumItem(a) = 0
                    End If

                    vQuery = "exec dbo.USP_NP_SelectItemQtySTKLocation '" & vItemCode(a) & "','" & vWHCode(a) & "','" & vShelfCode(a) & "' "
                    Call vGetData5(vMemProfit, vQuery)

                    If pds5.Tables(0).Rows.Count > 0 Then
                        vInspectQTY(a) = pds5.Tables(0).Rows(0)("qty").ToString
                    End If
                    vDiff(a) = vSumItem(a) - vInspectQTY(a)
                Next

                For x = 0 To vCountItem - 1
                    vQuery = "exec dbo.USP_NP_InsertBCInspectSub '" & vGenInspectNo & "','" & vItemCode(x) & "','" & vUnitCode(x) & "','" & vWHCode(x) & "','" & vShelfCode(x) & "'," & vInspectQTY(x) & "," & vSumItem(x) & "," & vDiff(x) & ",'" & vInSpectDesc(x) & "' "
                    Call vGetData6(vMemProfit, vQuery)
                Next

                vQuery = "Update npmaster.dbo.TB_IV_StkInspect_Logs set Inspectno = '" & vGenInspectNo & "' where docno = '" & vGenDocNo & "' "
                Call vGetData7(vMemProfit, vQuery)
                MsgBox("Save data is complete docno is  " & vGenInspectNo & " ")
                Call ClearScreen()
            Else
                MsgBox("Can not Save data.Item not exist in table")
            End If

        ElseIf vMemInspectIsOpen = 1 Then
            vDocNo = Me.TBDocNo.Text
            vDocDate = Me.TBDocDate.Text

            If Me.ListViewItem.Items.Count > 0 Then

                For i = 0 To Me.ListViewItem.Items.Count - 1

                    vLineNumber = i
                    vShelfStock = Trim(ListViewItem.Items(i).SubItems(7).Text)
                    vItmCode = Me.ListViewItem.Items(i).SubItems(1).Text
                    vItmName = Me.ListViewItem.Items(i).SubItems(11).Text
                    vWH = Trim(ListViewItem.Items(i).SubItems(6).Text)
                    vShelf = Me.ListViewItem.Items(i).SubItems(7).Text
                    vUnit = Me.ListViewItem.Items(i).SubItems(4).Text
                    vItmQty = Me.ListViewItem.Items(i).SubItems(3).Text
                    vItmInspectQty = Me.ListViewItem.Items(i).SubItems(2).Text
                    vItmDiff = vItmInspectQty - vItmQty
                    vReasonCode = Me.ListViewItem.Items(i).SubItems(8).Text
                    vIsSave = Me.ListViewItem.Items(i).SubItems(10).Text

                    If vIsSave = 0 Then
                        vQuery = "exec dbo.USP_NP_InsertBCInspectSub '" & vDocNo & "','" & vItmCode & "','" & vUnit & "','" & vWH & "','" & vShelf & "'," & vItmInspectQty & "," & vItmQty & "," & vItmDiff & ",'" & vReasonCode & "' "
                        Call vGetData6(vMemProfit, vQuery)

                        vQuery = "exec dbo.USP_NP_InsertInspectLog '" & vMemGetInspectNo & "','" & vItmCode & "','" & vItmName & "','" & vWH & "','" & vShelf & "'," & vItmQty & ",'" & vUnit & "','" & vUserName & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
                        Call vGetData(vMemProfit, vQuery)

                        vQuery = "Update npmaster.dbo.TB_IV_StkInspect_Logs set Inspectno = '" & vDocNo & "' where docno = '" & vMemGetInspectNo & "' "
                        Call vGetData7(vMemProfit, vQuery)

                        'vQuery = "exec dbo.USP_NP_ '" & vMemGetInspectNo & "'"
                        'Call vGetData(vMemProfit, vQuery)
                    End If

                Next
                MsgBox("Save data is complete docno is  " & vGenInspectNo & " ")
                Call ClearScreen()
            End If

        End If
    End Sub


    Public Sub vGetDocNo()
        Dim vYear As String
        Dim vYear1 As Integer
        Dim vYear2 As String
        Dim vMonth As String
        Dim vMonth1 As Integer
        Dim vMonth2 As String

        Dim vGetHeader As String
        Dim vHeader As String
        Dim vAutoNumber As Integer
        Dim vNoNumber As String

        Dim vCheckDocno As String
        Dim vGetMonth As String
        Dim vGetMonth1 As String
        Dim vGetYear As String
        Dim vGetYear1 As String

        Dim vNewDocNo As String
        Dim vMemDocNo As String

        On Error Resume Next


        vQuery = "exec dbo.USP_NP_SearchNewDocNo 10 "

        Call vGetData1(vMemProfit, vQuery)

        If pds1.Tables(0).Rows.Count > 0 Then
            vGetHeader = Trim(pds1.Tables(0).Rows(0)("header").ToString)
            vAutoNumber = pds1.Tables(0).Rows(0)("AutoNumber").ToString
        End If

        vYear = vb6.Right(vMemYear, 2)
        vYear1 = vYear
        If vYear1 < 54 Then
            vYear1 = vYear1 + 43
        End If
        vYear2 = vYear1

        vMonth = vMemMonth
        vMonth1 = vMonth
        If Len(vMonth1) < 2 Then
            vMonth1 = "0" & vMonth1
        End If
        vMonth2 = vMonth1

        vHeader = Trim(vGetHeader & vYear2 & vMonth2)
        vNoNumber = Format(vAutoNumber, "0000")
        vGenDocNo = Trim(vHeader & "-" & vNoNumber)

        vQuery = "exec dbo.USP_NP_UpdateNewDocNo  10 "
        Call vGetData2(vMemProfit, vQuery)

        'vQuery = "select top 1 docno from bcnp.dbo.bcstkinspect  where docno like '" & vMemProfit & "%' order by docno desc"

        If Me.RDBDay.Checked = True Then
            vMemDocNo = Trim(vMemProfit & "-ID")
        ElseIf Me.RDBBetweenDay.Checked = True Then
            vMemDocNo = Trim(vMemProfit & "-IH")
        End If

        vQuery = "select top 1 docno from bcnp.dbo.bcstkinspect  where docno like '" & vMemDocNo & "%' order by docno desc"
        Call vGetData3(vMemProfit, vQuery)

        If pds3.Tables(0).Rows.Count > 0 Then
            vCheckDocno = pds3.Tables(0).Rows(0)("docno").ToString
        End If

        If vb6.Left(vCheckDocno, 2) = "IH" Or vb6.Left(vCheckDocno, 2) = "ID" Then
            vGetYear = Mid(vCheckDocno, 3, 2)
            vGetMonth = Mid(vCheckDocno, 5, 2)
            vGetYear1 = vb6.Right(vMemYear, 2)
            vGetMonth1 = vMemMonth
            If vGetYear1 < 54 Then
                vGetYear1 = vYear1 + 43
            End If
            If Len(vGetMonth1) <> 2 Then
                vGetMonth1 = "0" & vGetMonth1
            End If
            If vGetYear1 = vGetYear And vGetMonth1 = vGetMonth Then
                vQuery = "select * from V_WEB_IV_ItemCheck_NewDocNo"
                Call vGetData4(vMemProfit, vQuery)
                If pds4.Tables(0).Rows.Count > 0 Then
                    vGenInspectNo = pds4.Tables(0).Rows(0)("newdocno").ToString
                End If

            Else
                If Me.RDBBetweenDay.Checked = True Then
                    vGenInspectNo = vMemProfit & "-" & Trim("IH" & vGetYear1 & vGetMonth1 & "-0001")
                ElseIf Me.RDBDay.Checked = True Then
                    vGenInspectNo = vMemProfit & "-" & Trim("ID" & vGetYear1 & vGetMonth1 & "-0001")
                End If

            End If
        ElseIf vb6.Left(vCheckDocno, 3) = vMemProfit Then

            Dim vLen As Integer
            Dim vDocNo As String

            vLen = Len(vCheckDocno)
            vDocNo = vb6.Right(vCheckDocno, vLen - 4)

            vGetYear = Mid(vDocNo, 3, 2)
            vGetMonth = Mid(vDocNo, 5, 2)
            vGetYear1 = vb6.Right(vMemYear, 2)
            vGetMonth1 = vMemMonth
            If vGetYear1 < 54 Then
                vGetYear1 = vGetYear1 + 43
            End If
            If Len(vGetMonth1) <> 2 Then
                vGetMonth1 = "0" & vGetMonth1
            End If
            If vGetYear1 = vGetYear And vGetMonth1 = vGetMonth Then
                If Me.RDBDay.Checked = True Then
                    vQuery = "select * from V_WEB_IV_ItemCheck_NewInspect"
                ElseIf Me.RDBBetweenDay.Checked = True Then
                    vQuery = "select * from V_WEB_IV_ItemCheck_NewDocNo"
                End If

                Call vGetData4(vMemProfit, vQuery)
                If pds4.Tables(0).Rows.Count > 0 Then
                    vGenInspectNo = pds4.Tables(0).Rows(0)("newdocno").ToString
                End If
            Else

                If Me.RDBBetweenDay.Checked = True Then
                    vGenInspectNo = vMemProfit & "-" & Trim("IH" & vGetYear1 & vGetMonth1 & "-0001")
                ElseIf Me.RDBDay.Checked = True Then
                    vGenInspectNo = vMemProfit & "-" & Trim("ID" & vGetYear1 & vGetMonth1 & "-0001")
                End If

            End If
        Else

            vYear = vb6.Right(vMemYear, 2)
            vYear1 = vYear
            vGetYear1 = vYear1
            vGetMonth1 = vMemMonth
            If vGetYear1 < 54 Then
                vGetYear1 = vGetYear1 + 43
            End If
            If Len(vGetMonth1) <> 2 Then
                vGetMonth1 = "0" & vGetMonth1
            End If

            If Me.RDBBetweenDay.Checked = True Then
                vGenInspectNo = vMemProfit & "-" & Trim("IH" & vGetYear1 & vGetMonth1 & "-0001")
            ElseIf Me.RDBDay.Checked = True Then
                vGenInspectNo = vMemProfit & "-" & Trim("ID" & vGetYear1 & vGetMonth1 & "-0001")
            End If

            End If
    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        Call vSearchStockInspect()
        Me.PNSearchDocNo.Visible = True
        Me.TBSearchDocNo.Focus()
    End Sub

    Public Sub vSearchStockInspect()
        Dim i As Integer
        Dim n As Integer
        Dim vGetDocDate As Date
        Dim vDocDate As String
        Dim vSearch As String

        On Error Resume Next

        vSearch = Me.TBSearchDocNo.Text
        Me.ListViewSearchDocNo.Items.Clear()
        vQuery = "exec dbo.USP_NP_SearchInspect '" & vSearch & "'"
        Call vGetData(vMemProfit, vQuery)

        If pds.Tables(0).Rows.Count > 0 Then

            For i = 0 To pds.Tables(0).Rows.Count - 1
                n = i + 1

                Dim listItem As New ListViewItem(n)
                listItem.SubItems.Add(pds.Tables(0).Rows(i)("docno").ToString)

                vGetDocDate = pds.Tables(0).Rows(i)("docdate").ToString
                vDocDate = vGetDocDate.Day & "/" & vGetDocDate.Month & "/" & vGetDocDate.Year

                listItem.SubItems.Add(vDocDate)
                listItem.SubItems.Add(pds.Tables(0).Rows(i)("creatorcode").ToString)
                Me.ListViewSearchDocNo.Items.Add(listItem)
            Next
        End If
    End Sub

    Private Sub BTNSearchDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchDocNo.Click
        On Error Resume Next

        Call vSearchStockInspect()
        Me.PNSearchDocNo.Visible = True
        Me.TBSearchDocNo.Focus()
    End Sub

    Private Sub ListViewSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearchDocNo.KeyDown
        Dim vIndex As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Enter And Me.ListViewSearchDocNo.Items.Count > 0 Then
            vIndex = Me.ListViewSearchDocNo.FocusedItem.Index
            Me.TBDocNo.Text = Me.ListViewSearchDocNo.Items(vIndex).SubItems(1).Text

            Me.PNSearchDocNo.Visible = False
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Up And Me.ListViewSearchDocNo.FocusedItem.Index = 0 Then
            Me.TBSearchDocNo.Focus()
            Me.TBSearchDocNo.SelectAll()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchDocNo.Visible = False
            Me.TBBarCode.Focus()
        End If
    End Sub

    Private Sub ListViewSearchDocNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewSearchDocNo.SelectedIndexChanged

    End Sub

    Private Sub TBDocNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TBDocNo.TextChanged

        'On Error Resume Next

        If Me.TBDocNo.Text <> "" Then
            Call StockInspectDetails(Me.TBDocNo.Text)
        End If
        Me.TBBarCode.Focus()
    End Sub

    Public Sub StockInspectDetails(ByVal vDocNo As String)
        Dim i As Integer
        Dim n As Integer
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vReasonCode As String
        Dim vQty As Double
        Dim vRemainQty As Double
        Dim vChkItemCode As String
        Dim vChkUnitCode As String
        Dim vChkWHCode As String
        Dim vChkShelfCode As String
        Dim vGetDocDate As Date
        Dim vDocDate As String
        Dim vCountQty As Double


        'On Error Resume Next

        Me.ListViewItem.Items.Clear()
        vQuery = "exec dbo.USP_NP_SearchInspectDetails '" & vDocNo & "'"
        Call vGetData(vMemProfit, vQuery)

        If pds.Tables(0).Rows.Count > 0 Then
            vMemInspectIsOpen = 1
            vGetDocDate = pds.Tables(0).Rows(i)("docdate").ToString
            vMemGetInspectNo = pds.Tables(0).Rows(i)("inspectno").ToString
            vDocDate = vGetDocDate.Day & "/" & vGetDocDate.Month & "/" & vGetDocDate.Year
            Me.TBDocDate.Text = vDocDate

            For i = 0 To pds.Tables(0).Rows.Count - 1

                vRemainQty = pds.Tables(0).Rows(i)("stkqty").ToString
                vQty = pds.Tables(0).Rows(i)("inspectqty").ToString
                vItemCode = pds.Tables(0).Rows(i)("itemcode").ToString
                vBarCode = pds.Tables(0).Rows(i)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(i)("itemname").ToString
                vUnitCode = pds.Tables(0).Rows(i)("unitcode").ToString
                vWHCode = pds.Tables(0).Rows(i)("whcode").ToString
                vShelfCode = pds.Tables(0).Rows(i)("shelfcode").ToString
                vReasonCode = pds.Tables(0).Rows(i)("reasoncode").ToString

                n = Me.ListViewItem.Items.Count + 1
                Dim listItem As New ListViewItem(n)
                listItem.SubItems.Add(vItemCode)
                listItem.SubItems.Add(Format(vRemainQty, "##,##0.00"))
                listItem.SubItems.Add(Format(vQty, "##,##0.00"))
                listItem.SubItems.Add(vUnitCode)
                listItem.SubItems.Add(vBarCode)
                listItem.SubItems.Add(vWHCode)
                listItem.SubItems.Add(vShelfCode)
                listItem.SubItems.Add(vReasonCode)
                listItem.SubItems.Add(Now)
                listItem.SubItems.Add(1)
                listItem.SubItems.Add(vItemName)
                Me.ListViewItem.Items.Add(listItem)

                Me.ListViewItem.Items(i).BackColor = Color.LightGreen
            Next
        End If
    End Sub

    Private Sub BTNClearScreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearScreen.Click
        Call ClearScreen()
    End Sub

    Private Sub BTNDeleteLine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNDeleteLine.Click
        Call DeleteData()
    End Sub

    Public Sub DeleteData()
        Dim vDocNo As String
        Dim vAnswer As Integer

        On Error GoTo ErrDescription

        If Me.TBDocNo.Text <> "" And vMemInspectIsOpen = 1 Then
            vDocNo = Me.TBDocNo.Text

            vAnswer = MsgBox("Do you want delete this docno ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 Then
                vQuery = "exec dbo.USP_HH_DeleteStockInspect '" & vDocNo & "'"
                Call vGetData(vMemProfit, vQuery)

                MsgBox("Delete this " & vDocNo & " is complete", MsgBoxStyle.Information, "Send Information Message")
                Call ClearScreen()
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub ListViewItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItem.KeyDown
        Dim vIndex As Integer
        Dim vAnswer As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Back And Me.ListViewItem.Items.Count > 0 Then
            If Me.ListViewItem.Items.Count > 0 Then
                vIndex = Me.ListViewItem.FocusedItem.Index

                vAnswer = MsgBox("Do you want delete this item ?", MsgBoxStyle.YesNo, "Send Question Message")

                If vAnswer = 6 Then
                    Me.ListViewItem.Items.RemoveAt(vIndex)
                    Call GenLineNumber()
                End If
                Me.TBBarCode.Focus()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
            FormMainApplication.Show()
            Me.Hide()
        End If

        If e.KeyCode = Keys.Up And IsDBNull(Me.ListViewItem.FocusedItem.Index) = 0 Then
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
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

    Private Sub BTNSelectDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectDocNo.Click
        Dim vIndex As Integer

        On Error Resume Next

        If Me.ListViewSearchDocNo.Items.Count > 0 Then
            vIndex = Me.ListViewSearchDocNo.FocusedItem.Index
            Me.TBDocNo.Text = Me.ListViewSearchDocNo.Items(vIndex).SubItems(1).Text

            Me.PNSearchDocNo.Visible = False
            Me.TBBarCode.Focus()
        End If
    End Sub

    Private Sub BTNCloseSearchDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSearchDoc.Click
        Me.PNSearchDocNo.Visible = False
        Me.TBItemCode.Focus()
    End Sub

    Private Sub TBSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchDocNo.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call vSearchStockInspect()
            Me.PNSearchDocNo.Visible = True
            Me.TBSearchDocNo.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchDocNo.Visible = False
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Down And Me.ListViewSearchDocNo.Items.Count > 0 Then
            Me.ListViewSearchDocNo.Items(0).Focused = True
            Me.ListViewSearchDocNo.Items(0).Selected = True
            Me.ListViewSearchDocNo.Focus()
        End If
    End Sub

    Private Sub TBSearchDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchDocNo.TextChanged

    End Sub

    Private Sub BTNClose_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClose.KeyDown, BTNClearScreen.KeyDown, BTNDeleteLine.KeyDown, BTNSaveData.KeyDown, BTNSearch.KeyDown
        Dim vAnswer As Integer

        On Error Resume Next

        If e.KeyCode = 113 Then
            vAnswer = MsgBox("Do you want clear screen ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 Then
                Call ClearScreen()
            End If
        End If

        If e.KeyCode = 116 Then
            Call SaveData()
        End If

        If e.KeyCode = 117 Then
            Call vSearchStockInspect()
            Me.PNSearchDocNo.Visible = True
            Me.TBSearchDocNo.Focus()
        End If

        If e.KeyCode = 119 Then
            Call DeleteData()
        End If

        If e.KeyCode = Keys.Escape Then
            vAnswer = MsgBox("Do you exit program ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 Then
                Call ClearScreen()
                FormMainApplication.Show()
                Me.Hide()
            End If
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBADDItemCode.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBADDStkQty.TextChanged

    End Sub

    Private Sub BTNADDQty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNADDQty.Click
        Dim vAddQty As Double
        Dim vNewQty As Double
        Dim vOldQty As Double
        Dim vIndex As Integer

        On Error Resume Next

        If Me.TBADDId.Text <> "" Then
            vIndex = Me.TBADDId.Text

            If Me.TBADDOldQty.Text <> "" Then
                vOldQty = Me.TBADDOldQty.Text
            End If

            If Me.TBADDNewQty.Text <> "" Then
                vNewQty = Me.TBADDNewQty.Text
            End If

            vAddQty = vNewQty + vOldQty
            Me.ListViewItem.Items(vIndex).SubItems(3).Text = Format(vAddQty, "##,##0.00")
            Me.ListViewItem.Items(vIndex).SubItems(10).Text = 0

            Me.PNCheckCount.Visible = False
            Call ClearItem()
            Me.TBBarCode.Focus()
        End If
    End Sub

    Private Sub BTNChangeQty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNChangeQty.Click
        Dim vAddQty As Double
        Dim vNewQty As Double
        Dim vOldQty As Double
        Dim vIndex As Integer

        On Error Resume Next

        If Me.TBADDId.Text <> "" Then
            vIndex = Me.TBADDId.Text

            If Me.TBADDOldQty.Text <> "" Then
                vOldQty = Me.TBADDOldQty.Text
            End If

            If Me.TBADDNewQty.Text <> "" Then
                vNewQty = Me.TBADDNewQty.Text
            End If

            vAddQty = vNewQty
            Me.ListViewItem.Items(vIndex).SubItems(3).Text = Format(vAddQty, "##,##0.00")
            Me.ListViewItem.Items(vIndex).SubItems(10).Text = 0

            Me.PNCheckCount.Visible = False
            Call ClearItem()
            Me.TBBarCode.Focus()
        End If
    End Sub

    Private Sub BTNADDClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNADDClose.Click
        Me.PNCheckCount.Visible = False
        Call ClearItem()
        Me.TBBarCode.Focus()
    End Sub

    Private Sub BTNNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.PNSelectStkType.Visible = False

    End Sub

    Private Sub BTNSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelect.Click
        Me.PNSelectStkType.Visible = False
        Me.CMBShelfCode.Focus()
    End Sub

    Private Sub RDBDay_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RDBDay.CheckedChanged

    End Sub

    Private Sub RDBDay_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDBDay.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.PNSelectStkType.Visible = False
            Me.CMBShelfCode.Focus()
        End If
    End Sub

    Private Sub RDBDay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RDBDay.KeyPress
    End Sub

    Private Sub RDBBetweenDay_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RDBBetweenDay.CheckedChanged

    End Sub

    Private Sub RDBBetweenDay_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDBBetweenDay.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.PNSelectStkType.Visible = False
            Me.CMBShelfCode.Focus()
        End If
    End Sub

    Private Sub PNSelectStkType_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PNSelectStkType.GotFocus

    End Sub
End Class