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

Public Class FormPrintLabel
    Private MyReader As Symbol.Barcode.Reader = Nothing
    Private MyReaderData As Symbol.Barcode.ReaderData = Nothing
    Private MyEventHandler As System.EventHandler = Nothing

    Dim vQuery As String
    Dim vMemDocDate As String

    Private Sub FormPrintLabel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        'If (Me.InitReader()) Then
        '    Me.StartRead()
        'Else
        '    Me.Close()
        '    Return
        'End If

        Call SearchLabelNotPrint()

        Call AddLabelType()

        vQuery = "exec dbo.USP_NP_CheckDocDate"
        Call vGetData1(vMemProfit, vQuery)
        If pds1.Tables(0).Rows.Count > 0 Then
            vMemDocDate = pds1.Tables(0).Rows(0)("vdocdate").ToString
        End If

        Me.CMBLabelType.SelectedIndex = 0
        Me.CMBLabelType.Focus()

    End Sub

    Public Sub SearchLabelNotPrint()
        Dim i As Integer
        Dim n As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vBarCode As String
        Dim vQty As Double
        Dim vType As String
        Dim vDocDate As String

        On Error Resume Next

        vQuery = "exec dbo.USP_NP_SearchLabelNotPrint '" & vUserName & "'"
        Call vGetData1(vMemProfit, vQuery)
        If pds1.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds1.Tables(0).Rows.Count - 1
                vItemCode = pds1.Tables(0).Rows(i)("itemcode").ToString
                vItemName = pds1.Tables(0).Rows(i)("itemname").ToString
                vQty = pds1.Tables(0).Rows(i)("qty").ToString
                vUnitCode = pds1.Tables(0).Rows(i)("unitcode").ToString
                vBarCode = pds1.Tables(0).Rows(i)("barcode").ToString
                vType = pds1.Tables(0).Rows(i)("labeltype").ToString
                vDocDate = pds1.Tables(0).Rows(i)("datetimestamp").ToString

                n = n + 1

                Dim listItem As New ListViewItem(n)
                listItem.SubItems.Add(vBarCode)
                listItem.SubItems.Add(vQty)
                listItem.SubItems.Add(vType)
                listItem.SubItems.Add(vDocDate)
                listItem.SubItems.Add(1)
                listItem.SubItems.Add(vItemCode)
                listItem.SubItems.Add(vUnitCode)
                Me.ListViewItem.Items.Add(listItem)
                Me.ListViewItem.Items(i).BackColor = Color.LightGreen

            Next
        End If
    End Sub

    Public Sub AddLabelType()
        On Error Resume Next

        Me.CMBLabelType.Items.Add("")
        Me.CMBLabelType.Items.Add("P1F1-»éÒÂ¸ÃÃÁ´Ò 21 ´Ç§/A4")
        Me.CMBLabelType.Items.Add("P2F1-»éÒÂ¸ÃÃÁ´Ò 3 ´Ç§/A4")
        Me.CMBLabelType.Items.Add("P3F1-»éÒÂ¸ÃÃÁ´Ò 2 ´Ç§/A4")
        Me.CMBLabelType.Items.Add("P4F1-»éÒÂ¸ÃÃÁ´Ò A4")
        Me.CMBLabelType.Items.Add("P5F1-»éÒÂ¸ÃÃÁ´Ò Cotto")
        Me.CMBLabelType.Items.Add("P1F2-»éÒÂÃÒ¤Ò¾ÔàÈÉ 21 ´Ç§/A4")
        Me.CMBLabelType.Items.Add("P2F2-»éÒÂÃÒ¤Ò¾ÔàÈÉ 3 ´Ç§/A4")
        Me.CMBLabelType.Items.Add("P3F2-»éÒÂÃÒ¤Ò¾ÔàÈÉ 2 ´Ç§/A4")
        Me.CMBLabelType.Items.Add("P4F2-»éÒÂÃÒ¤Ò¾ÔàÈÉ A4")
        Me.CMBLabelType.SelectedIndex = 1
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
        Dim vPrice As Double
        Dim vRedDot As Integer
        Dim vPriceType As String

        Dim TheReaderData As Symbol.Barcode.ReaderData = Me.MyReader.GetNextReaderData()

        On Error Resume Next

        If (TheReaderData.Result = Symbol.Results.SUCCESS) Then
            Me.TBBarCode.Text = TheReaderData.Text
            Me.StartRead()

            vBarCode = TheReaderData.Text

            Me.BTNRedDot.Visible = False

            vQuery = "exec dbo.usp_hh_SearchItemDataDetails '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vPrice = pds.Tables(0).Rows(0)("saleprice1").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString
                vRedDot = pds.Tables(0).Rows(0)("reddot").ToString
                vPriceType = pds.Tables(0).Rows(0)("remark").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBTypePrice.Text = vPriceType
                Me.TBUnitPrice.Text = vUnitCode
                If vRedDot = 1 Then
                    Me.BTNRedDot.Visible = True
                Else
                    Me.BTNRedDot.Visible = False
                End If

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

    Private Sub TBBarcode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarcode.KeyDown
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vBarCode As String
        Dim vPrice As Double
        Dim vRedDot As Integer
        Dim vPriceType As String

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            vBarCode = Me.TBBarcode.Text
            vQuery = "exec dbo.usp_hh_SearchItemDataDetails '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vPrice = pds.Tables(0).Rows(0)("saleprice1").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString
                vRedDot = pds.Tables(0).Rows(0)("reddot").ToString
                vPriceType = pds.Tables(0).Rows(0)("remark").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBTypePrice.Text = vPriceType
                Me.TBUnitPrice.Text = vUnitCode
                If vRedDot = 1 Then
                    Me.BTNRedDot.Visible = True
                Else
                    Me.BTNRedDot.Visible = False
                End If

                Me.TBQty.Focus()

            Else
                Me.TBBarcode.Text = ""
                Me.TBBarcode.Focus()
                Me.TBBarcode.SelectAll()
                MsgBox("This barcode find not found !", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If

        End If

        If e.KeyCode = Keys.Up Then
            Me.CMBLabelType.Focus()
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBQty.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearItem()
            Me.TBBarcode.Focus()
        End If

    End Sub

    Public Sub ClearItem()
        On Error Resume Next

        Me.TBBarcode.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBQty.Text = ""
        Me.TBUnitPrice.Text = ""
        Me.TBPrice.Text = ""
        Me.TBTypePrice.Text = ""
        Me.BTNRedDot.Visible = False

        Me.TBBarcode.Focus()
        Me.TBBarcode.SelectAll()
    End Sub

    Public Sub ClearScreen()
        On Error Resume Next

        Me.TBBarCode.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBQty.Text = ""
        Me.TBUnitPrice.Text = ""
        Me.TBPrice.Text = ""
        Me.TBTypePrice.Text = ""
        Me.BTNRedDot.Visible = False
        Me.ListViewItem.Items.Clear()

        Call SearchLabelNotPrint()

        Me.TBBarCode.Focus()
        Me.TBBarCode.SelectAll()
    End Sub

    Private Sub TBBarcode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarcode.TextChanged
        Dim vBarCode As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vPrice As Double
        Dim vRedDot As Integer
        Dim vPriceType As String

        If vb6.InStr(Me.TBBarcode.Text, "@") > 0 Then
            vBarCode = vb6.Left(Me.TBBarcode.Text, vb6.Len(Me.TBBarcode.Text) - 1)

            Me.TBBarcode.Text = vBarCode

            vBarCode = Me.TBBarcode.Text
            vQuery = "exec dbo.usp_hh_SearchItemDataDetails '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vPrice = pds.Tables(0).Rows(0)("saleprice1").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString
                vRedDot = pds.Tables(0).Rows(0)("reddot").ToString
                vPriceType = pds.Tables(0).Rows(0)("remark").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBTypePrice.Text = vPriceType
                Me.TBUnitPrice.Text = vUnitCode
                If vRedDot = 1 Then
                    Me.BTNRedDot.Visible = True
                Else
                    Me.BTNRedDot.Visible = False
                End If

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

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        Call ClearScreen()
        FormMainApplication.Show()
        Me.Hide()
    End Sub

    Private Sub TBQty_HandleDestroyed(ByVal sender As Object, ByVal e As System.EventArgs) Handles TBQty.HandleDestroyed

    End Sub

    Private Sub TBQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBQty.KeyDown
        Dim vBarCode As String
        Dim vQty As Double
        Dim i As Integer
        Dim vType As String
        Dim n As Integer
        Dim vCheckBarCode As String
        Dim vCheckLine As Integer
        Dim vCheckType As String
        Dim vDocDate As String
        Dim vAnswer As Integer
        Dim vItemCode As String
        Dim vUnitCode As String

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.CMBLabelType.Text = "" Then
                MsgBox("Please Select Label Type", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBLabelType.Focus()
                Exit Sub
            End If

            If Me.TBBarcode.Text <> "" And Me.TBQty.Text <> "" Then
                i = Me.ListViewItem.Items.Count + 1
                vBarCode = Me.TBBarcode.Text
                vItemCode = Me.TBItemCode.Text
                vUnitCode = Me.TBUnitPrice.Text
                vQty = Me.TBQty.Text
                vType = vb6.Left(Me.CMBLabelType.Text, 4)

                For n = 0 To Me.ListViewItem.Items.Count - 1
                    vCheckBarCode = Me.ListViewItem.Items(n).SubItems(1).Text
                    vCheckType = Me.ListViewItem.Items(n).SubItems(3).Text

                    If vBarCode = vCheckBarCode And vType = vCheckType Then
                        vAnswer = MsgBox("This barcode " & vBarCode & " aleady exist. Do you want replace qty this itemcode?", MsgBoxStyle.YesNo, "Send Error Message")
                        If vAnswer = 6 Then
                            Me.ListViewItem.Items(n).SubItems(2).Text = vQty
                            Me.TBBarcode.Text = ""
                            Me.TBItemCode.Text = ""
                            Me.TBQty.Text = ""
                            Me.TBItemName.Text = ""
                            Me.TBPrice.Text = ""
                            Me.TBUnitPrice.Text = ""
                            Me.TBTypePrice.Text = ""
                            Me.BTNRedDot.Visible = False
                            Me.TBBarcode.Focus()
                            Exit Sub
                        Else
                            Me.TBBarcode.Text = ""
                            Me.TBItemCode.Text = ""
                            Me.TBQty.Text = ""
                            Me.TBItemName.Text = ""
                            Me.TBPrice.Text = ""
                            Me.TBUnitPrice.Text = ""
                            Me.TBTypePrice.Text = ""
                            Me.BTNRedDot.Visible = False
                            Me.TBBarcode.Focus()
                            Exit Sub
                        End If
                    End If
                Next

                vDocDate = Now

                Dim listItem As New ListViewItem(i)
                listItem.SubItems.Add(vBarCode)
                listItem.SubItems.Add(vQty)
                listItem.SubItems.Add(vType)
                listItem.SubItems.Add(vDocDate)
                listItem.SubItems.Add(0)
                listItem.SubItems.Add(vItemCode)
                listItem.SubItems.Add(vUnitCode)
                Me.ListViewItem.Items.Add(listItem)

                If Me.ListViewItem.Items.Count > 0 Then
                    vCheckLine = Me.ListViewItem.Items.Count
                    Me.ListViewItem.Focus()
                    VScrollBar1.Value = vCheckLine - 1
                End If

                Me.TBBarcode.Text = ""
                Me.TBItemCode.Text = ""
                Me.TBItemName.Text = ""
                Me.TBQty.Text = ""
                Me.TBUnitPrice.Text = ""
                Me.TBPrice.Text = ""
                Me.TBTypePrice.Text = ""
                Me.BTNRedDot.Visible = False
                Me.TBBarcode.Focus()
            End If
        End If

        If e.KeyCode = Keys.Down And Me.ListViewItem.Items.Count > 0 Then
            Me.ListViewItem.Focus()
            Me.ListViewItem.Items(0).Focused = True
            Me.ListViewItem.Items(0).Selected = True
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearItem()
            Me.TBBarcode.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBBarcode.Focus()
        End If

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
        Dim i As Integer
        Dim vFileName As String
        Dim vItemCode As String
        Dim vUnitCode As String
        Dim vBarCode As String
        Dim vQty As Double
        Dim vReOrder As Double
        Dim vSuggest As Double
        Dim vPrice As Double
        Dim vLabelType As String
        Dim vWHCode As String
        Dim vZoneID As String
        Dim vShelfCode As String
        Dim vRowID As String
        Dim vShelfID As String
        Dim vJobType As Integer
        Dim vDocDate As String
        Dim vReasonCode As String

        'On Error Resume Next

        vJobType = 4

        If Me.ListViewItem.Items.Count > 0 Then

            For i = 0 To Me.ListViewItem.Items.Count - 1
                vBarCode = Me.ListViewItem.Items(i).SubItems(1).Text
                vQty = Me.ListViewItem.Items(i).SubItems(2).Text
                vLabelType = Me.ListViewItem.Items(i).SubItems(3).Text
                vDocDate = vMemDocDate
                vItemCode = Me.ListViewItem.Items(i).SubItems(6).Text
                vUnitCode = Me.ListViewItem.Items(i).SubItems(7).Text

                vReOrder = 0
                vSuggest = 0
                vPrice = 0
                vWHCode = ""
                vZoneID = ""
                vShelfCode = ""
                vRowID = ""
                vShelfID = ""
                vReasonCode = ""
                vFileName = ""

                vQuery = "exec dbo.USP_NP_InsertItemDataOfflineCenter " & vJobType & ",'" & vItemCode & "','" & vBarCode & "'," & vQty & "," & vReOrder & "," & vSuggest & ",'" & vWHCode & "','" & vZoneID & "','" & vShelfCode & "','" & vRowID & "','" & vShelfID & "'," & vPrice & ",'" & vLabelType & "','" & vDocDate & "','" & vUserName & "','" & vFileName & "','" & vUnitCode & "','" & vReasonCode & "'"
                Call vExecData(vMemProfit, vQuery)
            Next

            Call ClearScreen()
        End If
    End Sub

    Private Sub ListViewItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItem.KeyDown
        Dim vIndex As Integer
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vUnitCode As String
        Dim vLabelType As String
        Dim vAnswer As Integer
        Dim vIsSave As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Back And Me.ListViewItem.Items.Count > 0 Then
            vIndex = Me.ListViewItem.FocusedItem.Index
            vBarCode = Me.ListViewItem.Items(vIndex).SubItems(1).Text
            vItemCode = Me.ListViewItem.Items(vIndex).SubItems(6).Text
            vUnitCode = Me.ListViewItem.Items(vIndex).SubItems(7).Text
            vLabelType = Me.ListViewItem.Items(vIndex).SubItems(3).Text
            vIsSave = Me.ListViewItem.Items(vIndex).SubItems(5).Text

            vAnswer = MsgBox("Do you want delete barcode " & vBarCode & " ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 And vIsSave = 0 Then
                Me.ListViewItem.Items.RemoveAt(vIndex)
                Call GenLineNumber()
            End If

            If vAnswer = 6 And vIsSave = 1 Then
                Me.ListViewItem.Items.RemoveAt(vIndex)
                Call GenLineNumber()

                vQuery = "exec dbo.USP_NP_DeleteLabelNotPrint '" & vUserName & "','" & vItemCode & "','" & vUnitCode & "','" & vLabelType & "'"
                Call vExecData(vMemProfit, vQuery)
            End If

        End If

        If e.KeyCode = Keys.Up And IsDBNull(Me.ListViewItem.FocusedItem.Index) = 0 Then
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

    Private Sub CMBLabelType_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBLabelType.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBBarcode.Focus()
        End If
    End Sub

    Private Sub BTNClearScreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearScreen.Click
        Call ClearScreen()
        Me.CMBLabelType.Focus()
    End Sub
End Class