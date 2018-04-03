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

Public Class FormReOrder

    Private MyScanner As Symbol.Barcode.Device = Nothing

    Private MyReader As Symbol.Barcode.Reader = Nothing
    Private MyReaderData As Symbol.Barcode.ReaderData = Nothing
    Private MyEventHandler As System.EventHandler = Nothing

    Private MyReadNotifyHander As System.EventHandler = Nothing
    Private MyStatusNotifyHandler As System.EventHandler = Nothing
    Private MyActivateHandler As System.EventHandler = Nothing
    Private MyDeActivateHandler As System.EventHandler = Nothing



    Dim vQuery As String
    Dim vMemDocDate As String

    Private Sub TBBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
        Dim vAnswer As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vStkUnit As String
        Dim vBarCode As String
        Dim vPrice As Double
        Dim vStockQty As Double
        Dim vStockMax As Double
        Dim vStockMin As Double
        Dim vOrderPoint As Double
        Dim vItemStatus As String
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
        Dim vSumCashSale3Month As Double
        Dim vPORemainIn As Double
        Dim vRedDot As Integer

        Dim vFreq3Month As Double
        Dim vMyGrade As String
        Dim vExpertTeam As String

        Dim vPRDate As String
        Dim vAPDate As String
        Dim vPODate As String
        Dim vPOLeadDate As String

        Dim vGetPRDate As Date
        Dim vGetAPDate As Date
        Dim vGetPODate As Date
        Dim vGetPOLeadDate As Date

        Dim vPRNo As String
        Dim vPRQty As Double
        Dim vPRLine As Integer
        Dim vPRRemain As Double
        Dim vAPNo As String
        Dim vAPQty As Double
        Dim vAPRemain As Double
        Dim vPONo As String
        Dim vPOQty As Double
        Dim vPORemain As Double
        Dim vPRItemUnit As String
        Dim vRQStatus As Integer

        On Error GoTo ErrDescription


        If e.KeyCode = Keys.Escape Then
            Call ClearItem()
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Enter Then

            If vb6.InStr(Me.TBBarCode.Text, "@") > 0 Then

                vBarCode = vb6.Left(Me.TBBarCode.Text, vb6.Len(Me.TBBarCode.Text) - 1)

                Me.TBBarCode.Text = vBarCode
            Else
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                Exit Sub
            End If

            MsgBox(vBarCode)

            Me.BTNRedDot.Visible = False
            Me.ListViewStock.Items.Clear()
            Me.ListViewStock.Visible = False
            Me.ListViewShelfID.Items.Clear()
            Me.ListViewShelfID.Visible = False

            vQuery = "exec dbo.usp_hh_SearchItemDataDetails_Cat '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            vRQStatus = 0
            vPRNo = ""
            vAPNo = ""
            vPONo = ""

            vPRDate = ""
            vAPDate = ""
            vPODate = ""
            vPOLeadDate = ""

            vPRQty = 0
            vAPQty = 0
            vPOQty = 0
            vPRRemain = 0
            vAPRemain = 0
            vPORemain = 0
            vPRLine = 0
            vPRItemUnit = ""

            vItemCode = ""
            vItemName = ""
            vPrice = 0
            vUnitCode = ""
            vBarCode = ""

            vOrderPoint = 0
            vStockMin = 0
            vStockMax = 0
            vItemStatus = ""
            vPORemainIn = 0
            vSumCashSale3Month = 0
            vFreq3Month = 0
            vRedDot = 0
            vMyGrade = ""
            vExpertTeam = ""

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vPrice = pds.Tables(0).Rows(0)("saleprice1").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString
                vRQStatus = pds.Tables(0).Rows(0)("rqstatus").ToString
                vOrderPoint = pds.Tables(0).Rows(0)("orderpoint").ToString
                vStockMin = pds.Tables(0).Rows(0)("stockmin").ToString
                vStockMax = pds.Tables(0).Rows(0)("stockmax").ToString
                vItemStatus = pds.Tables(0).Rows(0)("itemstatus").ToString
                vPORemainIn = pds.Tables(0).Rows(0)("remaininqty").ToString
                vSumCashSale3Month = pds.Tables(0).Rows(0)("avgsale1Month").ToString
                vFreq3Month = pds.Tables(0).Rows(0)("avgcountbill1Month").ToString
                vRedDot = pds.Tables(0).Rows(0)("reddot").ToString
                vMyGrade = pds.Tables(0).Rows(0)("mygrade").ToString
                vExpertTeam = pds.Tables(0).Rows(0)("expertteam").ToString
                vPRNo = pds.Tables(0).Rows(0)("prno").ToString
                vAPNo = pds.Tables(0).Rows(0)("cfno").ToString
                vPONo = pds.Tables(0).Rows(0)("pono").ToString
                vGetPRDate = pds.Tables(0).Rows(0)("prdate").ToString
                vGetAPDate = pds.Tables(0).Rows(0)("cfdate").ToString
                vGetPODate = pds.Tables(0).Rows(0)("podate").ToString
                vGetPOLeadDate = pds.Tables(0).Rows(0)("poleaddate").ToString
                vPRQty = pds.Tables(0).Rows(0)("prqty").ToString
                vAPQty = pds.Tables(0).Rows(0)("cfqty").ToString
                vPOQty = pds.Tables(0).Rows(0)("poqty").ToString
                vPRRemain = pds.Tables(0).Rows(0)("prremain").ToString
                vAPRemain = pds.Tables(0).Rows(0)("cfremain").ToString
                vPORemain = pds.Tables(0).Rows(0)("poremian").ToString
                vPRLine = pds.Tables(0).Rows(0)("prlinenumber").ToString
                vPRItemUnit = pds.Tables(0).Rows(0)("pritemunit").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBGrade.Text = vMyGrade
                Me.TBQty.Text = ""
                Me.TBRemainQty.Text = Format(vStockQty, "##,##0.00")
                Me.TBSuggest.Text = ""
                Me.TBOrderPoint.Text = Format(vOrderPoint, "##,##0.00")
                Me.TBMin.Text = Format(vStockMin, "##,##0.00")
                Me.TBMax.Text = Format(vStockMax, "##,##0.00")
                Me.TBUnit.Text = vUnitCode
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBReOrder.Text = ""
                Me.TBItemStatus.Text = vItemStatus
                Me.TBPORemain.Text = Format(vPORemainIn, "##,##0.00")
                Me.TBSale1M.Text = Format(vSumCashSale3Month, "##,##0.00")
                Me.TBFrequency.Text = Format(vFreq3Month, "##,##0.00")
                Me.TBExpertTeam.Text = vExpertTeam

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

                If ((vPRNo <> "") Or (vAPNo <> "") Or (vPORemain > 0)) And vRQStatus <> 3 And vRQStatus <> 6 And vRQStatus <> 7 Then

                    vPRDate = vb6.Day(vGetPRDate) & "/" & vb6.Month(vGetPRDate) & "/" & vb6.Year(vGetPRDate)
                    vPOLeadDate = vb6.Day(vGetPOLeadDate) & "/" & vb6.Month(vGetPOLeadDate) & "/" & vb6.Year(vGetPOLeadDate)

                    Me.PNShowHist.Visible = True
                    Me.TBShowItem.Text = vItemCode & "//" & vItemName
                    Me.TBPRNo.Text = vPRNo
                    Me.TBPRDate.Text = vPRDate
                    Me.TBPRQty.Text = vPRQty
                    Me.TBConfirmQty.Text = vAPQty
                    Me.TBConfirmNo.Text = vAPNo
                    Me.TBPOQty.Text = vPOQty
                    Me.TBPONo.Text = vPONo
                    Me.TBLineNumber.Text = vPRLine
                    If vPOLeadDate <> "1/1/1900" Then
                        Me.TBPOLeadDate.Text = vPOLeadDate
                    Else
                        Me.TBPOLeadDate.Text = ""
                    End If
                    Me.TBPRUnit.Text = vPRItemUnit

                    If vAPNo <> "" Then
                        Me.TBEditQty.ReadOnly = True
                    End If

                    If vPORemain > 0 Then
                        Me.TBEditQty.ReadOnly = True
                    End If

                    If vPRNo <> "" And vAPNo = "" And vPONo = "" Then
                        Me.TBEditQty.ReadOnly = False
                        Me.TBEditQty.Focus()
                    End If
                    Me.TBEditQty.Focus()
                    Exit Sub
                End If

            Else
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                MsgBox("This barcode find not found !", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If


            If vRedDot > 0 Then
                Me.BTNRedDot.Visible = True
            Else
                Me.BTNRedDot.Visible = False
            End If

            Me.ListViewStock.Visible = True
            Me.ListViewShelfID.Visible = True

            Me.TBQty.Focus()
            Me.TBQty.SelectAll()

        End If


        If e.KeyCode = Keys.Down Then
            Me.TBQty.Focus()
            Me.TBQty.SelectAll()
        End If

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
            Call vSearchStockRequest()
            Me.PNSearchDocNo.Visible = True
            Me.TBSearchDocNo.Focus()
        End If

        If e.KeyCode = 119 Then
            Call CancelData()
        End If


ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

    End Sub


    Private Sub TBBarCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged
        Dim vAnswer As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vStkUnit As String
        Dim vBarCode As String
        Dim vPrice As Double
        Dim vStockQty As Double
        Dim vStockMax As Double
        Dim vStockMin As Double
        Dim vOrderPoint As Double
        Dim vItemStatus As String
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
        Dim vSumCashSale3Month As Double
        Dim vPORemainIn As Double
        Dim vRedDot As Integer

        Dim vFreq3Month As Double
        Dim vMyGrade As String
        Dim vExpertTeam As String

        Dim vPRDate As String
        Dim vAPDate As String
        Dim vPODate As String
        Dim vPOLeadDate As String

        Dim vGetPRDate As Date
        Dim vGetAPDate As Date
        Dim vGetPODate As Date
        Dim vGetPOLeadDate As Date

        Dim vPRNo As String
        Dim vPRQty As Double
        Dim vPRLine As Integer
        Dim vPRRemain As Double
        Dim vAPNo As String
        Dim vAPQty As Double
        Dim vAPRemain As Double
        Dim vPONo As String
        Dim vPOQty As Double
        Dim vPORemain As Double
        Dim vPRItemUnit As String
        Dim vRQStatus As Integer

        On Error GoTo ErrDescription

        If vb6.InStr(Me.TBBarCode.Text, "@") > 0 Then

            vBarCode = vb6.Left(Me.TBBarCode.Text, vb6.Len(Me.TBBarCode.Text) - 1)

            Me.TBBarCode.Text = vBarCode

            Me.BTNRedDot.Visible = False
            Me.ListViewStock.Items.Clear()
            Me.ListViewStock.Visible = False
            Me.ListViewShelfID.Items.Clear()
            Me.ListViewShelfID.Visible = False

            vQuery = "exec dbo.usp_hh_SearchItemDataDetails_Cat '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            vRQStatus = 0
            vPRNo = ""
            vAPNo = ""
            vPONo = ""

            vPRDate = ""
            vAPDate = ""
            vPODate = ""
            vPOLeadDate = ""

            vPRQty = 0
            vAPQty = 0
            vPOQty = 0
            vPRRemain = 0
            vAPRemain = 0
            vPORemain = 0
            vPRLine = 0
            vPRItemUnit = ""

            vItemCode = ""
            vItemName = ""
            vPrice = 0
            vUnitCode = ""
            vBarCode = ""

            vOrderPoint = 0
            vStockMin = 0
            vStockMax = 0
            vItemStatus = ""
            vPORemainIn = 0
            vSumCashSale3Month = 0
            vFreq3Month = 0
            vRedDot = 0
            vMyGrade = ""
            vExpertTeam = ""

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vPrice = pds.Tables(0).Rows(0)("saleprice1").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString
                vRQStatus = pds.Tables(0).Rows(0)("rqstatus").ToString
                vOrderPoint = pds.Tables(0).Rows(0)("orderpoint").ToString
                vStockMin = pds.Tables(0).Rows(0)("stockmin").ToString
                vStockMax = pds.Tables(0).Rows(0)("stockmax").ToString
                vItemStatus = pds.Tables(0).Rows(0)("itemstatus").ToString
                vPORemainIn = pds.Tables(0).Rows(0)("remaininqty").ToString
                vSumCashSale3Month = pds.Tables(0).Rows(0)("avgsale1Month").ToString
                vFreq3Month = pds.Tables(0).Rows(0)("avgcountbill1Month").ToString
                vRedDot = pds.Tables(0).Rows(0)("reddot").ToString
                vMyGrade = pds.Tables(0).Rows(0)("mygrade").ToString
                vExpertTeam = pds.Tables(0).Rows(0)("expertteam").ToString
                vPRNo = pds.Tables(0).Rows(0)("prno").ToString
                vAPNo = pds.Tables(0).Rows(0)("cfno").ToString
                vPONo = pds.Tables(0).Rows(0)("pono").ToString
                vGetPRDate = pds.Tables(0).Rows(0)("prdate").ToString
                vGetAPDate = pds.Tables(0).Rows(0)("cfdate").ToString
                vGetPODate = pds.Tables(0).Rows(0)("podate").ToString
                vGetPOLeadDate = pds.Tables(0).Rows(0)("poleaddate").ToString
                vPRQty = pds.Tables(0).Rows(0)("prqty").ToString
                vAPQty = pds.Tables(0).Rows(0)("cfqty").ToString
                vPOQty = pds.Tables(0).Rows(0)("poqty").ToString
                vPRRemain = pds.Tables(0).Rows(0)("prremain").ToString
                vAPRemain = pds.Tables(0).Rows(0)("cfremain").ToString
                vPORemain = pds.Tables(0).Rows(0)("poremian").ToString
                vPRLine = pds.Tables(0).Rows(0)("prlinenumber").ToString
                vPRItemUnit = pds.Tables(0).Rows(0)("pritemunit").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBGrade.Text = vMyGrade
                Me.TBQty.Text = ""
                Me.TBRemainQty.Text = Format(vStockQty, "##,##0.00")
                Me.TBSuggest.Text = ""
                Me.TBOrderPoint.Text = Format(vOrderPoint, "##,##0.00")
                Me.TBMin.Text = Format(vStockMin, "##,##0.00")
                Me.TBMax.Text = Format(vStockMax, "##,##0.00")
                Me.TBUnit.Text = vUnitCode
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBReOrder.Text = ""
                Me.TBItemStatus.Text = vItemStatus
                Me.TBPORemain.Text = Format(vPORemainIn, "##,##0.00")
                Me.TBSale1M.Text = Format(vSumCashSale3Month, "##,##0.00")
                Me.TBFrequency.Text = Format(vFreq3Month, "##,##0.00")
                Me.TBExpertTeam.Text = vExpertTeam

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

                If ((vPRNo <> "") Or (vAPNo <> "") Or (vPORemain > 0)) And vRQStatus <> 3 And vRQStatus <> 6 And vRQStatus <> 7 Then

                    vPRDate = vb6.Day(vGetPRDate) & "/" & vb6.Month(vGetPRDate) & "/" & vb6.Year(vGetPRDate)
                    vPOLeadDate = vb6.Day(vGetPOLeadDate) & "/" & vb6.Month(vGetPOLeadDate) & "/" & vb6.Year(vGetPOLeadDate)

                    Me.PNShowHist.Visible = True
                    Me.TBShowItem.Text = vItemCode & "//" & vItemName
                    Me.TBPRNo.Text = vPRNo
                    Me.TBPRDate.Text = vPRDate
                    Me.TBPRQty.Text = vPRQty
                    Me.TBConfirmQty.Text = vAPQty
                    Me.TBConfirmNo.Text = vAPNo
                    Me.TBPOQty.Text = vPOQty
                    Me.TBPONo.Text = vPONo
                    Me.TBLineNumber.Text = vPRLine
                    If vPOLeadDate <> "1/1/1900" Then
                        Me.TBPOLeadDate.Text = vPOLeadDate
                    Else
                        Me.TBPOLeadDate.Text = ""
                    End If
                    Me.TBPRUnit.Text = vPRItemUnit

                    If vAPNo <> "" Then
                        Me.TBEditQty.ReadOnly = True
                    End If

                    If vPORemain > 0 Then
                        Me.TBEditQty.ReadOnly = True
                    End If

                    If vPRNo <> "" And vAPNo = "" And vPONo = "" Then
                        Me.TBEditQty.ReadOnly = False
                        Me.TBEditQty.Focus()
                    End If
                    Me.TBEditQty.Focus()
                    Exit Sub
                End If

            ElseIf Me.TBBarCode.Text = "" Then
                Call ClearItem()
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
            End If


            If vRedDot > 0 Then
                Me.BTNRedDot.Visible = True
            Else
                Me.BTNRedDot.Visible = False
            End If

            Me.ListViewStock.Visible = True
            Me.ListViewShelfID.Visible = True

            Me.TBQty.Focus()
            Me.TBQty.SelectAll()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Function InitReader2() As Boolean

        ' If reader is already present then fail initialize
        If Not (Me.MyReader Is Nothing) Then

            Return False

        End If

        'Create new reader, first available reader will be used.
        Me.MyReader = New Symbol.Barcode.Reader

        'Create reader data
        Me.MyReaderData = New Symbol.Barcode.ReaderData(Symbol.Barcode.ReaderDataTypes.Text, Symbol.Barcode.ReaderDataLengths.DefaultText)


        ' create event handler delegate
        Me.MyEventHandler = New System.EventHandler(AddressOf MyReader_ReadNotify)

        'Enable reader, with wait cursor
        Me.MyReader.Actions.Enable()

        'Me.MyReader.Parameters.Feedback.Success.BeepTime = 0
        'Me.MyReader.Parameters.Feedback.Success.WaveFile = "\\windows\\alarm3.wav"

        AddHandler Me.Activated, New EventHandler(AddressOf ReaderForm_Activated)
        AddHandler Me.Deactivate, New EventHandler(AddressOf ReaderForm_Deactivate)


        Return True

    End Function


    Private Sub TermReader()

        'If we have a reader
        If Not (Me.MyReader Is Nothing) Then

            'Disable reader, with wait cursor
            Me.MyReader.Actions.Disable()

            'free it up
            Me.MyReader.Dispose()

            ' Indicate we no longer have one
            Me.MyReader = Nothing

        End If

        ' If we have a reader data
        If Not (Me.MyReaderData Is Nothing) Then

            'Free it up
            Me.MyReaderData.Dispose()

            'Indicate we no longer have one
            Me.MyReaderData = Nothing

        End If

    End Sub


    Private Sub StartRead2()

        'If we have both a reader and a reader data
        If Not ((Me.MyReader Is Nothing) And (Me.MyReaderData Is Nothing)) Then

            'Submit a read
            AddHandler MyReader.ReadNotify, Me.MyEventHandler

            Me.MyReader.Actions.Read(Me.MyReaderData)

        End If

    End Sub


    Private Sub StopRead2()

        'If we have a reader
        If Not (Me.MyReader Is Nothing) Then

            'Flush (Cancel all pending reads)
            RemoveHandler MyReader.ReadNotify, Me.MyEventHandler

            Me.MyReader.Actions.Flush()

        End If

    End Sub


    Private Sub MyReader_ReadNotify2(ByVal o As Object, ByVal e As EventArgs)

        Dim TheReaderData As Symbol.Barcode.ReaderData = Me.MyReader.GetNextReaderData()

        If (TheReaderData.Result = Symbol.Results.SUCCESS) Then

            Me.HandleData(TheReaderData)

            Me.StartRead()

        End If

    End Sub


    Private Sub HandleData(ByVal TheReaderData As Symbol.Barcode.ReaderData)
        Try

            Dim S As String
            S = TheReaderData.Text

        Catch ex As Exception
        End Try


    End Sub

    Private Sub FORM_Activated(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Activated
        'If there are no reads pending on MyReader start a new read
        'If Not Me.MyReaderData Is Nothing Then
        '    If Not (Me.MyReaderData.IsPending) Then

        '        Me.StartRead()

        '    End If
        'End If

    End Sub

    Private Sub FormReOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'On Error Resume Next

        'If (Me.InitReader()) Then
        '    Me.StartRead()
        'Else
        '    Me.Close()
        '    Return
        'End If



        'Try
        '    If (Me.InitReader()) Then
        '        Me.StartRead()
        '    End If
        'Catch ex As Exception
        '    MsgBox(Err.Description)
        'End Try


        vQuery = "exec dbo.USP_NP_CheckDocDate"
        Call vGetData1(vMemProfit, vQuery)
        If pds1.Tables(0).Rows.Count > 0 Then
            vMemDocDate = pds1.Tables(0).Rows(0)("vdocdate").ToString
        End If

        Me.TBDocDate.Text = vMemDocDate
        Me.TBBarCode.Focus()
    End Sub

    Private Function InitReader() As Boolean

        'On Error Resume Next

        If Not (Me.MyReader Is Nothing) Then
            Return False
        End If

        Me.MyReader = New Symbol.Barcode.Reader
        Me.MyReaderData = New Symbol.Barcode.ReaderData( _
                                     Symbol.Barcode.ReaderDataTypes.Text, _
                                     Symbol.Barcode.ReaderDataLengths.DefaultText)
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
        Dim vAnswer As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vStkUnit As String
        Dim vBarCode As String
        Dim vPrice As Double
        Dim vStockQty As Double
        Dim vStockMax As Double
        Dim vStockMin As Double
        Dim vOrderPoint As Double
        Dim vItemStatus As String
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
        Dim vSumCashSale3Month As Double
        Dim vPORemainIn As Double
        Dim vRedDot As Integer

        Dim vFreq3Month As Double
        Dim vMyGrade As String
        Dim vExpertTeam As String

        Dim vPRDate As String
        Dim vAPDate As String
        Dim vPODate As String
        Dim vPOLeadDate As String

        Dim vGetPRDate As Date
        Dim vGetAPDate As Date
        Dim vGetPODate As Date
        Dim vGetPOLeadDate As Date

        Dim vPRNo As String
        Dim vPRQty As Double
        Dim vPRLine As Integer
        Dim vPRRemain As Double
        Dim vAPNo As String
        Dim vAPQty As Double
        Dim vAPRemain As Double
        Dim vPONo As String
        Dim vPOQty As Double
        Dim vPORemain As Double
        Dim vPRItemUnit As String
        Dim vRQStatus As Integer

        Dim TheReaderData As Symbol.Barcode.ReaderData = Me.MyReader.GetNextReaderData()

        On Error Resume Next

        If (TheReaderData.Result = Symbol.Results.SUCCESS) Then
            Me.TBBarCode.Text = TheReaderData.Text
            Me.StartRead()

            vBarCode = TheReaderData.Text

            Me.BTNRedDot.Visible = False
            Me.ListViewStock.Items.Clear()
            Me.ListViewStock.Visible = False
            Me.ListViewShelfID.Items.Clear()
            Me.ListViewShelfID.Visible = False

            vQuery = "exec dbo.usp_hh_SearchItemDataDetails_Cat '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            vRQStatus = 0
            vPRNo = ""
            vAPNo = ""
            vPONo = ""

            vPRDate = ""
            vAPDate = ""
            vPODate = ""
            vPOLeadDate = ""

            vPRQty = 0
            vAPQty = 0
            vPOQty = 0
            vPRRemain = 0
            vAPRemain = 0
            vPORemain = 0
            vPRLine = 0
            vPRItemUnit = ""

            vItemCode = ""
            vItemName = ""
            vPrice = 0
            vUnitCode = ""
            vBarCode = ""

            vOrderPoint = 0
            vStockMin = 0
            vStockMax = 0
            vItemStatus = ""
            vPORemainIn = 0
            vSumCashSale3Month = 0
            vFreq3Month = 0
            vRedDot = 0
            vMyGrade = ""
            vExpertTeam = ""

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vPrice = pds.Tables(0).Rows(0)("saleprice1").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString
                vRQStatus = pds.Tables(0).Rows(0)("rqstatus").ToString
                vOrderPoint = pds.Tables(0).Rows(0)("orderpoint").ToString
                vStockMin = pds.Tables(0).Rows(0)("stockmin").ToString
                vStockMax = pds.Tables(0).Rows(0)("stockmax").ToString
                vItemStatus = pds.Tables(0).Rows(0)("itemstatus").ToString
                vPORemainIn = pds.Tables(0).Rows(0)("remaininqty").ToString
                vSumCashSale3Month = pds.Tables(0).Rows(0)("avgsale1Month").ToString
                vFreq3Month = pds.Tables(0).Rows(0)("avgcountbill1Month").ToString
                vRedDot = pds.Tables(0).Rows(0)("reddot").ToString
                vMyGrade = pds.Tables(0).Rows(0)("mygrade").ToString
                vExpertTeam = pds.Tables(0).Rows(0)("expertteam").ToString
                vPRNo = pds.Tables(0).Rows(0)("prno").ToString
                vAPNo = pds.Tables(0).Rows(0)("cfno").ToString
                vPONo = pds.Tables(0).Rows(0)("pono").ToString
                vGetPRDate = pds.Tables(0).Rows(0)("prdate").ToString
                vGetAPDate = pds.Tables(0).Rows(0)("cfdate").ToString
                vGetPODate = pds.Tables(0).Rows(0)("podate").ToString
                vGetPOLeadDate = pds.Tables(0).Rows(0)("poleaddate").ToString
                vPRQty = pds.Tables(0).Rows(0)("prqty").ToString
                vAPQty = pds.Tables(0).Rows(0)("cfqty").ToString
                vPOQty = pds.Tables(0).Rows(0)("poqty").ToString
                vPRRemain = pds.Tables(0).Rows(0)("prremain").ToString
                vAPRemain = pds.Tables(0).Rows(0)("cfremain").ToString
                vPORemain = pds.Tables(0).Rows(0)("poremian").ToString
                vPRLine = pds.Tables(0).Rows(0)("prlinenumber").ToString
                vPRItemUnit = pds.Tables(0).Rows(0)("pritemunit").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBGrade.Text = vMyGrade
                Me.TBQty.Text = ""
                Me.TBRemainQty.Text = Format(vStockQty, "##,##0.00")
                Me.TBSuggest.Text = ""
                Me.TBOrderPoint.Text = Format(vOrderPoint, "##,##0.00")
                Me.TBMin.Text = Format(vStockMin, "##,##0.00")
                Me.TBMax.Text = Format(vStockMax, "##,##0.00")
                Me.TBUnit.Text = vUnitCode
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBReOrder.Text = ""
                Me.TBItemStatus.Text = vItemStatus
                Me.TBPORemain.Text = Format(vPORemainIn, "##,##0.00")
                Me.TBSale1M.Text = Format(vSumCashSale3Month, "##,##0.00")
                Me.TBFrequency.Text = Format(vFreq3Month, "##,##0.00")
                Me.TBExpertTeam.Text = vExpertTeam

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

                If ((vPRNo <> "") Or (vAPNo <> "") Or (vPORemain > 0)) And vRQStatus <> 7 Then

                    vPRDate = vb6.Day(vGetPRDate) & "/" & vb6.Month(vGetPRDate) & "/" & vb6.Year(vGetPRDate)
                    vPOLeadDate = vb6.Day(vGetPOLeadDate) & "/" & vb6.Month(vGetPOLeadDate) & "/" & vb6.Year(vGetPOLeadDate)

                    Me.PNShowHist.Visible = True
                    Me.TBShowItem.Text = vItemCode & "//" & vItemName
                    Me.TBPRNo.Text = vPRNo
                    Me.TBPRDate.Text = vPRDate
                    Me.TBPRQty.Text = vPRQty
                    Me.TBConfirmQty.Text = vAPQty
                    Me.TBConfirmNo.Text = vAPNo
                    Me.TBPOQty.Text = vPOQty
                    Me.TBPONo.Text = vPONo
                    Me.TBLineNumber.Text = vPRLine
                    If vPOLeadDate <> "1/1/1900" Then
                        Me.TBPOLeadDate.Text = vPOLeadDate
                    Else
                        Me.TBPOLeadDate.Text = ""
                    End If
                    Me.TBPRUnit.Text = vPRItemUnit

                    If vAPNo <> "" Then
                        Me.TBEditQty.ReadOnly = True
                    End If

                    If vPORemain > 0 Then
                        Me.TBEditQty.ReadOnly = True
                    End If

                    If vPRNo <> "" And vAPNo = "" And vPONo = "" Then
                        Me.TBEditQty.ReadOnly = False
                        Me.TBEditQty.Focus()
                    End If
                    Me.TBEditQty.Focus()
                    Exit Sub
                End If

            Else
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                MsgBox("This barcode find not found !", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If


            If vRedDot > 0 Then
                Me.BTNRedDot.Visible = True
            Else
                Me.BTNRedDot.Visible = False
            End If

            Me.ListViewStock.Visible = True
            Me.ListViewShelfID.Visible = True

            Me.TBQty.Focus()
            Me.TBQty.SelectAll()

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

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Call ClearScreen()
        FormMainApplication.Show()
        Me.Hide()
    End Sub

    Private Sub TBReOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBReOrder.KeyDown
        Dim i As Integer
        Dim n As Integer
        Dim vCheckLine As Integer
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vQty As Double
        Dim vReOrder As Double
        Dim vSuggest As Double
        Dim vUnitCode As String
        Dim vOldReOrder As Double
        Dim vDocDate As String
        Dim vGetSale1Month As Double

        Dim vCheckItemCode As String

        Dim vAnswer As Integer
        Dim vAnswer1 As Integer
        Dim vAnswer2 As Integer
        Dim vNewQty As Double
        Dim vGetItemStatus As String
        Dim vItemStatus As Integer
        Dim vGrade As String

        Dim vOrderPoint As Double
        Dim vStockMax As Double
        Dim vStockMin As Double
        Dim vCountStkQty As Double
        Dim vExpertTeam As String

        Dim vAnswerBuyOverOrder As Integer
        Dim vSumQtyCheckOrder As Double


        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then

            If Me.TBBarCode.Text = "" Then
                MsgBox("Please Insert ItemCode", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Exit Sub
            End If

            If Me.TBReOrder.Text = "" Then
                MsgBox("Please Insert Re-Order Qty", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBReOrder.Focus()
                Exit Sub
            End If

            If Me.TBReOrder.Text <> "" Then
                vReOrder = Me.TBReOrder.Text
            Else
                vReOrder = 0
            End If

            If vReOrder = 0 Then
                MsgBox("Please Insert Re-Order Qty", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBReOrder.Focus()
                Exit Sub
            End If

            If Me.TBQty.Text <> "" Then
                vCountStkQty = Me.TBQty.Text
            Else
                vCountStkQty = 0
            End If

            If Me.TBOrderPoint.Text <> "" Then
                vOrderPoint = Me.TBOrderPoint.Text
            Else
                vOrderPoint = 0
            End If

            If Me.TBMax.Text <> "" Then
                vStockMax = Me.TBMax.Text
            Else
                vStockMax = 0
            End If

            If Me.TBMin.Text <> "" Then
                vStockMin = Me.TBMin.Text
            Else
                vStockMin = 0
            End If

            If Me.TBQty.Text = "" Then
                vQty = 0
            Else
                vQty = Me.TBQty.Text
            End If

            If Me.TBSale1M.Text <> "" Then
                vGetSale1Month = Me.TBSale1M.Text
            Else
                vGetSale1Month = 0
            End If

            vItemCode = Me.TBItemCode.Text
            vBarCode = Me.TBBarCode.Text
            If Me.TBSuggest.Text <> "" Then
                vSuggest = Me.TBSuggest.Text
            Else
                vSuggest = 0
            End If

            vReOrder = Me.TBReOrder.Text

            vUnitCode = Me.TBUnit.Text
            vDocDate = Now
            vGrade = Me.TBGrade.Text
            vExpertTeam = Me.TBExpertTeam.Text

            If Me.TBItemStatus.Text <> "" Then
                vGetItemStatus = vb6.Left(Me.TBItemStatus.Text, 1)
                vItemStatus = vGetItemStatus
                vItemStatus = vItemStatus - 1
            Else
                vItemStatus = 1
            End If

            vSumQtyCheckOrder = vReOrder + vCountStkQty

            If vSumQtyCheckOrder > vGetSale1Month Then
                vAnswerBuyOverOrder = MsgBox("This Order+StockOnHand is over AverageSale1Month ! Do you want buy this item ?", MsgBoxStyle.YesNo, "Send Question Message")
                'MsgBox("This item is over stock max " & vGrade & " unable to Re-Order", MsgBoxStyle.Critical, "Send Error Message")
                If vAnswerBuyOverOrder = 7 Then
                    Me.TBReOrder.Focus()
                    Exit Sub
                End If
            End If

            If vItemStatus = 0 Then
                MsgBox("This item is Stop Sale", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Exit Sub
            End If

            If vItemStatus = 2 Then
                MsgBox("This item is Stop Buy", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Exit Sub
            End If

            If vItemStatus = 3 Then
                MsgBox("This item is Special Order", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Exit Sub
            End If

            If vItemStatus = 4 Then
                MsgBox("This item is Free Item", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Exit Sub
            End If

            If vItemStatus = 5 Then
                MsgBox("This item is Assets", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Exit Sub
            End If

            If Me.ListViewItem.Items.Count > 0 Then
                For n = 0 To Me.ListViewItem.Items.Count - 1
                    vCheckItemCode = Me.ListViewItem.Items(n).SubItems(1).Text
                    vOldReOrder = Me.ListViewItem.Items(n).SubItems(3).Text

                    If vItemCode = vCheckItemCode Then
                        vAnswer = MsgBox("This item aleady exist at line " & n + 1 & " Do you want edit qty ?", MsgBoxStyle.YesNo, "Send Error Message")
                        If vAnswer = 6 Then
                            vAnswer1 = MsgBox("Click YES Replace Qty,Click No Add QTY", MsgBoxStyle.YesNo, "")
                            If vAnswer1 = 6 Then
                                vNewQty = Me.TBReOrder.Text
                                vSumQtyCheckOrder = vNewQty + vCountStkQty

                                If vSumQtyCheckOrder > vGetSale1Month Then
                                    vAnswerBuyOverOrder = MsgBox("This Order Over AverageSale1Month ! Do you want buy this item ?", MsgBoxStyle.YesNo, "Send Question Message")
                                    'MsgBox("This item is over stock for 3 month " & vGrade & " unable to Re-Order", MsgBoxStyle.Critical, "Send Error Message")
                                    If vAnswerBuyOverOrder = 7 Then
                                        Me.TBReOrder.Focus()
                                        Exit Sub
                                    End If
                                End If

                                Me.ListViewItem.Items(n).SubItems(2).Text = Format(vQty, "##,##0.00")
                                Me.ListViewItem.Items(n).SubItems(3).Text = Format(vNewQty, "##,##0.00")
                                Me.ListViewItem.Items(n).SubItems(7).Text = 0
                                Me.ListViewItem.Items(n).BackColor = Color.Yellow
                            Else
                                vSumQtyCheckOrder = vReOrder + vOldReOrder + vCountStkQty
                                vNewQty = vReOrder + vOldReOrder

                                If vSumQtyCheckOrder > vGetSale1Month Then
                                    vAnswerBuyOverOrder = MsgBox("This Order Over AverageSale1Month ! Do you want buy this item ?", MsgBoxStyle.YesNo, "Send Question Message")
                                    'MsgBox("This item is over stock for 3 month " & vGrade & " unable to Re-Order", MsgBoxStyle.Critical, "Send Error Message")
                                    If vAnswerBuyOverOrder = 7 Then
                                        Me.TBReOrder.Focus()
                                        Exit Sub
                                    End If
                                End If

                                Me.ListViewItem.Items(n).SubItems(2).Text = Format(vQty, "##,##0.00")
                                Me.ListViewItem.Items(n).SubItems(3).Text = Format(vNewQty, "##,##0.00")
                                Me.ListViewItem.Items(n).SubItems(7).Text = 0
                                Me.ListViewItem.Items(n).BackColor = Color.Yellow
                            End If

                        End If

                        Call ClearItem()
                        Exit Sub
                    End If
                Next
            End If

            i = Me.ListViewItem.Items.Count + 1

            Dim listItem As New ListViewItem(i)
            listItem.SubItems.Add(vItemCode)
            listItem.SubItems.Add(Format(vQty, "##,##0.00"))
            listItem.SubItems.Add(Format(vReOrder, "##,##0.00"))
            listItem.SubItems.Add(Format(vSuggest, "##,##0.00"))
            listItem.SubItems.Add(vUnitCode)
            listItem.SubItems.Add(vDocDate)
            listItem.SubItems.Add(0)
            listItem.SubItems.Add(vExpertTeam)

            Me.ListViewItem.Items.Add(listItem)

            If Me.ListViewItem.Items.Count > 0 Then
                vCheckLine = Me.ListViewItem.Items.Count
                Me.ListViewItem.Focus()
                VScrollBar1.Value = vCheckLine - 1
            End If

            Call ClearItem()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBQty.Focus()
            Me.TBQty.SelectAll()
        End If

        If e.KeyCode = Keys.Down And Me.ListViewItem.Items.Count > 0 Then
            Me.ListViewItem.Items(0).Focused = True
            Me.ListViewItem.Items(0).Selected = True
            Me.ListViewItem.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearItem()
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = 113 Then
            vAnswer2 = MsgBox("Do you want clear screen ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer2 = 6 Then
                Call ClearScreen()
            End If
        End If

        If e.KeyCode = 116 Then
            Call SaveData()
        End If

        If e.KeyCode = 117 Then
            Call vSearchStockRequest()
            Me.PNSearchDocNo.Visible = True
            Me.TBSearchDocNo.Focus()
        End If

        If e.KeyCode = 119 Then
            Call CancelData()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TBReOrder_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBReOrder.TextChanged
        If vb6.InStr(Me.TBReOrder.Text, "@") > 0 Then
            Me.TBReOrder.Text = ""
            Me.TBReOrder.Focus()
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Call SaveData()
    End Sub

    Public Sub SaveData()
        Dim vDocNo As String
        Dim vNewDocNo As String
        Dim vDocDate As String
        Dim vItemCode As String
        Dim vQty As Double
        Dim vUnitCode As String
        Dim vLineNumber As Integer
        Dim i As Integer
        Dim n As Integer
        Dim vAnswer As Integer
        Dim vJobID As Integer
        Dim vCountQty As Double

        On Error GoTo ErrDescription

        vJobID = 1

        If Me.ListViewItem.Items.Count > 0 Then
            vAnswer = MsgBox("Do you want save this docno ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 Then
                If vMemReOrderIsOpen = 0 Then

                    vQuery = "exec dbo.USP_PR_GetPRHandHeldNo '" & vMemProfit & "'"
                    Call vGetData(vMemProfit, vQuery)

                    If pds.Tables(0).Rows.Count > 0 Then
                        vNewDocNo = pds.Tables(0).Rows(0)("PRHandHeldNo").ToString
                    End If

                    If Me.ListViewItem.Items.Count > 0 Then
                        vDocNo = vNewDocNo
                        vDocDate = vMemDocDate
                        vQuery = "exec dbo.USP_NP_InsertSTKRequest '" & vDocNo & "','" & vMemDocDate & "','" & vUserID & "','" & vUserName & "'"
                        Call vExecData(vMemProfit, vQuery)

                        For i = 0 To Me.ListViewItem.Items.Count - 1

                            vItemCode = Me.ListViewItem.Items(i).SubItems(1).Text
                            vCountQty = Me.ListViewItem.Items(i).SubItems(2).Text
                            vQty = Me.ListViewItem.Items(i).SubItems(3).Text
                            vUnitCode = Me.ListViewItem.Items(i).SubItems(5).Text
                            vLineNumber = i

                            vQuery = "exec dbo.USP_NP_InsertSTKRequestSub '" & vDocNo & "','" & vItemCode & "','" & vMemDocDate & "'," & vQty & ",'" & vUnitCode & "'," & vLineNumber & " "
                            Call vExecData(vMemProfit, vQuery)

                            vQuery = "exec dbo.USP_HH_InsertDataUsedHandHeld " & vJobID & ",'" & vItemCode & "','" & vItemCode & "','',''," & vCountQty & ",'" & vUnitCode & "','','" & vDocNo & "','" & vUserName & "'"
                            Call vExecData(vMemProfit, vQuery)

                        Next

                    End If

                    MsgBox("Save this " & vDocNo & " is complete", MsgBoxStyle.Information, "Send Information Message")

                    Call ClearScreen()

                ElseIf vMemReOrderIsOpen = 1 Then

                    vDocNo = Me.TBDocNo.Text
                    vDocDate = Me.TBDocDate.Text

                    If Me.ListViewItem.Items.Count > 0 Then
                        vQuery = "exec dbo.USP_NP_InsertSTKRequest '" & vDocNo & "','" & vDocDate & "','" & vUserID & "','" & vUserName & "'"
                        Call vExecData(vMemProfit, vQuery)

                        For n = 0 To Me.ListViewItem.Items.Count - 1

                            vItemCode = Me.ListViewItem.Items(n).SubItems(1).Text
                            vCountQty = Me.ListViewItem.Items(n).SubItems(2).Text
                            vQty = Me.ListViewItem.Items(n).SubItems(3).Text
                            vUnitCode = Me.ListViewItem.Items(n).SubItems(5).Text
                            vLineNumber = n

                            vQuery = "exec dbo.USP_NP_InsertSTKRequestSub '" & vDocNo & "','" & vItemCode & "','" & vDocDate & "'," & vQty & ",'" & vUnitCode & "'," & vLineNumber & " "
                            Call vExecData(vMemProfit, vQuery)

                            vQuery = "exec dbo.USP_HH_InsertDataUsedHandHeld " & vJobID & ",'" & vItemCode & "','" & vItemCode & "','',''," & vCountQty & ",'" & vUnitCode & "','','" & vDocNo & "','" & vUserName & "'"
                            Call vExecData(vMemProfit, vQuery)
                        Next

                    End If

                    MsgBox("Update this " & vDocNo & " is complete", MsgBoxStyle.Information, "Send Information Message")

                    Call ClearScreen()

                End If
            End If

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TBDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBDocNo.TextChanged
        On Error Resume Next

        If Me.TBDocNo.Text <> "" Then
            Call StockRequestDetails(Me.TBDocNo.Text)
        End If
        Me.TBBarCode.Focus()
    End Sub

    Public Sub StockRequestDetails(ByVal vDocNo As String)
        Dim i As Integer
        Dim n As Integer
        Dim vGetDocDate As Date
        Dim vDocDate As String
        Dim vCountQty As Double
        Dim vQty As Double
        Dim vSuggest As Double

        On Error Resume Next

        vIsconfirm = 0
        vIsCancel = 0

        Me.ListViewItem.Items.Clear()
        vQuery = "exec dbo.USP_HH_SearchStockRequestDetails '" & vDocNo & "'"
        Call vGetData(vMemProfit, vQuery)

        If pds.Tables(0).Rows.Count > 0 Then
            vMemReOrderIsOpen = 1
            vIsconfirm = pds.Tables(0).Rows(0)("isconfirm").ToString
            vIsCancel = pds.Tables(0).Rows(0)("iscancel").ToString
            vGetDocDate = pds.Tables(0).Rows(0)("docdate").ToString
            vDocDate = vGetDocDate.Day & "/" & vGetDocDate.Month & "/" & vGetDocDate.Year
            Me.TBDocDate.Text = vDocDate

            For i = 0 To pds.Tables(0).Rows.Count - 1
                n = i + 1

                vCountQty = pds.Tables(0).Rows(i)("countqty").ToString
                vQty = pds.Tables(0).Rows(i)("qty").ToString
                vSuggest = 0

                Dim listItem As New ListViewItem(n)
                listItem.SubItems.Add(pds.Tables(0).Rows(i)("itemcode").ToString)
                listItem.SubItems.Add(Format(vCountQty, "##,##0.00"))
                listItem.SubItems.Add(Format(vQty, "##,##0.00"))
                listItem.SubItems.Add(Format(vSuggest, "##,##0.00"))
                listItem.SubItems.Add(pds.Tables(0).Rows(i)("unitcode").ToString)
                listItem.SubItems.Add(Now)
                listItem.SubItems.Add(1)
                Me.ListViewItem.Items.Add(listItem)
                Me.ListViewItem.Items(i).BackColor = Color.LightGreen
            Next
        End If
    End Sub


    Private Sub BTNClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNNew.Click
        Dim vAnswer As Integer

        On Error Resume Next

        vAnswer = MsgBox("Do you want clear screen ?", MsgBoxStyle.YesNo, "Send Question Message")

        If vAnswer = 6 Then
            Call ClearScreen()
        End If
    End Sub

    Public Sub ClearScreen()
        On Error Resume Next

        vQuery = "exec dbo.USP_NP_CheckDocDate"
        Call vGetData1(vMemProfit, vQuery)
        If pds1.Tables(0).Rows.Count > 0 Then
            vMemDocDate = pds1.Tables(0).Rows(0)("vdocdate").ToString
        End If

        Me.TBDocDate.Text = vMemDocDate
        vIsconfirm = 0
        vIsCancel = 0
        vMemReOrderIsOpen = 0
        Me.TBDocNo.Text = ""
        Me.TBDocDate.Text = vMemDocDate
        Me.TBBarCode.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBQty.Text = ""
        Me.TBRemainQty.Text = ""
        Me.TBSuggest.Text = ""
        Me.TBOrderPoint.Text = ""
        Me.TBMin.Text = ""
        Me.TBMax.Text = ""
        Me.TBUnit.Text = ""
        Me.TBReOrder.Text = ""
        Me.TBPrice.Text = ""
        Me.TBItemStatus.Text = ""
        Me.TBPORemain.Text = ""
        Me.TBSale1M.Text = ""
        Me.TBFrequency.Text = ""
        Me.TBGrade.Text = ""
        Me.TBExpertTeam.Text = ""

        Me.TBShowItem.Text = ""
        Me.TBPRNo.Text = ""
        Me.TBPRDate.Text = ""
        Me.TBPRQty.Text = ""
        Me.TBConfirmQty.Text = ""
        Me.TBConfirmNo.Text = ""
        Me.TBPOQty.Text = ""
        Me.TBPONo.Text = ""
        Me.TBPOLeadDate.Text = ""
        Me.TBPRUnit.Text = ""
        Me.TBLineNumber.Text = ""


        Me.BTNRedDot.Visible = False
        Me.ListViewStock.Items.Clear()
        Me.ListViewStock.Visible = False
        Me.ListViewShelfID.Items.Clear()
        Me.ListViewShelfID.Visible = False
        Me.ListViewItem.Items.Clear()
        Me.TBBarCode.Focus()
        Me.TBBarCode.SelectAll()
    End Sub

    Public Sub ClearItem()
        On Error Resume Next

        Me.TBBarCode.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBQty.Text = ""
        Me.TBRemainQty.Text = ""
        Me.TBSuggest.Text = ""
        Me.TBOrderPoint.Text = ""
        Me.TBMin.Text = ""
        Me.TBMax.Text = ""
        Me.TBUnit.Text = ""
        Me.TBReOrder.Text = ""
        Me.TBPrice.Text = ""
        Me.TBItemStatus.Text = ""
        Me.TBPORemain.Text = ""
        Me.TBSale1M.Text = ""
        Me.TBFrequency.Text = ""
        Me.TBGrade.Text = ""
        Me.TBExpertTeam.Text = ""
        Me.BTNRedDot.Visible = False
        Me.ListViewStock.Items.Clear()
        Me.ListViewStock.Visible = False
        Me.ListViewShelfID.Items.Clear()
        Me.ListViewShelfID.Visible = False
        Me.TBBarCode.Focus()
        Me.TBBarCode.SelectAll()
    End Sub

    Private Sub TBQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBQty.KeyDown
        Dim vAnswer As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TBReOrder.Focus()
            Me.TBReOrder.SelectAll()
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBReOrder.Focus()
            Me.TBReOrder.SelectAll()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearItem()
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

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
            Call vSearchStockRequest()
            Me.PNSearchDocNo.Visible = True
            Me.TBSearchDocNo.Focus()
        End If

        If e.KeyCode = 119 Then
            Call CancelData()
        End If
    End Sub

    Private Sub TBQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBQty.TextChanged
        If vb6.InStr(Me.TBQty.Text, "@") > 0 Then
            Me.TBQty.Text = ""
            Me.TBQty.Focus()
        End If
    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        Call vSearchStockRequest()
        Me.PNSearchDocNo.Visible = True
        Me.TBSearchDocNo.Focus()
    End Sub

    Public Sub vSearchStockRequest()
        Dim i As Integer
        Dim n As Integer
        Dim vGetDocDate As Date
        Dim vDocDate As String
        Dim vSearch As String

        On Error Resume Next

        vSearch = Me.TBSearchDocNo.Text
        Me.ListViewSearchDocNo.Items.Clear()
        vQuery = "exec dbo.USP_HH_SearchStockRequest '" & vMemProfit & "','" & vSearch & "'"
        Call vGetData(vMemProfit, vQuery)

        If pds.Tables(0).Rows.Count > 0 Then

            For i = 0 To pds.Tables(0).Rows.Count - 1
                n = i + 1

                Dim listItem As New ListViewItem(n)
                listItem.SubItems.Add(pds.Tables(0).Rows(i)("docno").ToString)

                vGetDocDate = pds.Tables(0).Rows(i)("docdate").ToString
                vDocDate = vGetDocDate.Day & "/" & vGetDocDate.Month & "/" & vGetDocDate.Year

                listItem.SubItems.Add(vDocDate)
                listItem.SubItems.Add(pds.Tables(0).Rows(i)("workmanname").ToString)
                Me.ListViewSearchDocNo.Items.Add(listItem)
            Next
        End If
    End Sub

    Private Sub BTNCloseSearchDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSearchDoc.Click
        Me.PNSearchDocNo.Visible = False
        Me.TBBarCode.Focus()
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

    Private Sub BTNCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCancel.Click
        Call CancelData()
    End Sub

    Public Sub CancelData()
        Dim vDocNo As String
        Dim vAnswer As Integer

        On Error GoTo ErrDescription

        If Me.TBDocNo.Text <> "" And vMemReOrderIsOpen = 1 And vIsconfirm = 0 And vIsCancel = 0 Then
            vDocNo = Me.TBDocNo.Text

            vAnswer = MsgBox("Do you want cancel this docno ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 Then
                vQuery = "exec dbo.USP_HH_CancelStockRequest '" & vDocNo & "','" & vUserName & "'"
                Call vGetData(vMemProfit, vQuery)

                MsgBox("Cancel this " & vDocNo & " is complete", MsgBoxStyle.Information, "Send Information Message")
                Call ClearScreen()
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSearchDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchDocNo.Click
        On Error Resume Next

        Call vSearchStockRequest()
        Me.PNSearchDocNo.Visible = True
        Me.TBSearchDocNo.Focus()
    End Sub

    Private Sub TBSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchDocNo.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Call vSearchStockRequest()
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

    Private Sub ListViewItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItem.KeyDown
        Dim vIndex As Integer

        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vStkUnit As String
        Dim vBarCode As String
        Dim vPrice As Double
        Dim vStockQty As Double
        Dim vStockMax As Double
        Dim vStockMin As Double
        Dim vOrderPoint As Double
        Dim vItemStatus As String
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
        Dim vSumCashSale3Month As Double
        Dim vPORemainIn As Double
        Dim vRedDot As Integer
        Dim vFreq3Month As Double
        Dim vCountQty As Double
        Dim vReOrderQty As Double
        Dim vMyGrade As String
        Dim vAnswer As Integer

        'On Error GoTo ErrDescription
        On Error Resume Next

        If e.KeyCode = Keys.Back Then
            If Me.ListViewItem.Items.Count > 0 Then
                vIndex = Me.ListViewItem.FocusedItem.Index

                vAnswer = MsgBox("Do you want delete line " & vIndex + 1 & " ?", MsgBoxStyle.YesNo, "Send Question Message")
                If vAnswer = 6 Then
                    Me.ListViewItem.Items.RemoveAt(vIndex)

                    Call GenLineNumber()
                    Me.TBBarCode.Focus()
                End If
            End If
        End If

        If e.KeyCode = Keys.Up And IsDBNull(Me.ListViewItem.FocusedItem.Index) = 0 Then
            Me.TBReOrder.Focus()
            Me.TBReOrder.SelectAll()
        End If

        If e.KeyCode = Keys.Enter Then
            vIndex = Me.ListViewItem.FocusedItem.Index

            vBarCode = Me.ListViewItem.Items(vIndex).SubItems(1).Text
            vCountQty = Me.ListViewItem.Items(vIndex).SubItems(2).Text
            vReOrderQty = Me.ListViewItem.Items(vIndex).SubItems(3).Text

            Me.TBBarCode.Text = vBarCode
            Me.BTNRedDot.Visible = False
            Me.ListViewStock.Items.Clear()
            Me.ListViewStock.Visible = False
            Me.ListViewShelfID.Items.Clear()
            Me.ListViewShelfID.Visible = False


            vQuery = "exec dbo.usp_hh_SearchItemDataDetails '" & vMemProfit & "','" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)


            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vPrice = pds.Tables(0).Rows(0)("saleprice1").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString

                vOrderPoint = pds.Tables(0).Rows(0)("orderpoint").ToString
                vStockMin = pds.Tables(0).Rows(0)("stockmin").ToString
                vStockMax = pds.Tables(0).Rows(0)("stockmax").ToString
                vItemStatus = pds.Tables(0).Rows(0)("itemstatus").ToString
                vPORemainIn = pds.Tables(0).Rows(0)("remaininqty").ToString
                vSumCashSale3Month = pds.Tables(0).Rows(0)("sumsale3month").ToString
                vFreq3Month = pds.Tables(0).Rows(0)("countbills").ToString
                vRedDot = pds.Tables(0).Rows(0)("reddot").ToString
                vMyGrade = pds.Tables(0).Rows(0)("mygrade").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBGrade.Text = vMyGrade
                Me.TBQty.Text = Format(vCountQty, "##,##0.00")
                Me.TBRemainQty.Text = Format(vStockQty, "##,##0.00")
                Me.TBSuggest.Text = ""
                Me.TBOrderPoint.Text = Format(vOrderPoint, "##,##0.00")
                Me.TBMin.Text = Format(vStockMin, "##,##0.00")
                Me.TBMax.Text = Format(vStockMax, "##,##0.00")
                Me.TBUnit.Text = vUnitCode
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBReOrder.Text = Format(vReOrderQty, "##,##0.00")
                Me.TBItemStatus.Text = vItemStatus
                Me.TBPORemain.Text = Format(vPORemainIn, "##,##0.00")
                Me.TBSale1M.Text = Format(vSumCashSale3Month, "##,##0.00")
                Me.TBFrequency.Text = Format(vFreq3Month, "##,##0.00")

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

            Else
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                MsgBox("This barcode find not found !", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If


            If vRedDot > 0 Then
                Me.BTNRedDot.Visible = True
            Else
                Me.BTNRedDot.Visible = False
            End If

            Me.ListViewStock.Visible = True
            Me.ListViewShelfID.Visible = True

            Me.TBQty.Focus()
            Me.TBQty.SelectAll()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
            FormMainApplication.Show()
            Me.Hide()
        End If

        'ErrDescription:
        '        If Err.Description <> "" Then
        '            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
        '            Exit Sub
        '        End If
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

    Private Sub BTNSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSearchDocNo.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchDocNo.Visible = False
            Me.TBBarCode.Focus()
        End If
    End Sub

    Private Sub BTNCloseSearchDoc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCloseSearchDoc.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchDocNo.Visible = False
            Me.TBBarCode.Focus()
        End If
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

    Private Sub BTNSelectDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSelectDocNo.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchDocNo.Visible = False
            Me.TBBarCode.Focus()
        End If
    End Sub

    Private Sub BTNExit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNExit.KeyDown, BTNCancel.KeyDown, BTNNew.KeyDown, BTNSave.KeyDown, BTNSearch.KeyDown, ListViewItem.KeyDown, TBDocDate.KeyDown, TBDocNo.KeyDown, TBFrequency.KeyDown, TBItemCode.KeyDown, TBItemName.KeyDown, TBItemStatus.KeyDown, TBMax.KeyDown, TBMin.KeyDown, TBOrderPoint.KeyDown, TBPORemain.KeyDown, TBPrice.KeyDown, TBRemainQty.KeyDown, TBSale1M.KeyDown, TBUnit.KeyDown ', TBBarCode.KeyDown, TBQty.KeyDown, TBReOrder.KeyDown
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
            Call vSearchStockRequest()
            Me.PNSearchDocNo.Visible = True
            Me.TBSearchDocNo.Focus()
        End If

        If e.KeyCode = 119 Then
            Call CancelData()
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

    Private Sub TBQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TBQty.LostFocus
        Dim vBarCode As String
        Dim vGetCountQty As Double
        Dim vGetStockMax As Double
        Dim vGetStockMin As Double
        Dim vGetOrderPoint As Double

        On Error GoTo ErrDescription

        If Me.TBQty.Text <> "" Then
            vBarCode = Me.TBBarCode.Text

            If Me.TBRemainQty.Text <> "" Then
                vGetCountQty = Me.TBQty.Text
            Else
                vGetCountQty = 0
            End If

            If Me.TBOrderPoint.Text <> "" Then
                vGetOrderPoint = Me.TBOrderPoint.Text
            Else
                vGetOrderPoint = 0
            End If

            If Me.TBMax.Text <> "" Then
                vGetStockMax = Me.TBMax.Text
            Else
                vGetStockMax = 0
            End If

            If Me.TBMin.Text <> "" Then
                vGetStockMin = Me.TBMin.Text
            Else
                vGetStockMin = 0
            End If

            If vGetCountQty < vGetOrderPoint Then
                Me.TBSuggest.Text = Format((vGetStockMax - vGetCountQty), "##,##0.00")
            Else
                vGetCountQty = 0
                Me.TBSuggest.Text = Format(vGetCountQty, "##,##0.00")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TBReOrder_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBReOrder.KeyPress
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

    Private Sub BTNReqConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNReqConfirm.Click
        Dim vPRNo As String
        Dim vPRDate As String
        Dim vItemCode As String
        Dim vEditQty As Double
        Dim vUnitCode As String
        Dim vPRLine As Integer

        On Error Resume Next

        If Me.TBEditQty.Text = "" And Me.TBEditQty.ReadOnly = False Then
            MsgBox("You are not Edit Qty", MsgBoxStyle.Critical, "Send Information Message")
            Me.TBQty.Focus()
            Exit Sub
        ElseIf Me.TBEditQty.ReadOnly = True Then
            Me.TBPRNo.Text = ""
            Me.TBPRDate.Text = ""
            Me.TBShowItem.Text = ""
            Me.TBEditQty.Text = ""
            Me.TBPRUnit.Text = ""
            Me.TBLineNumber.Text = ""
            Me.PNShowHist.Visible = False
            Me.TBQty.Focus()
            Exit Sub
        End If

        vPRNo = Me.TBPRNo.Text
        vPRDate = Me.TBPRDate.Text
        vItemCode = Me.TBItemCode.Text
        vEditQty = Me.TBEditQty.Text
        vUnitCode = Me.TBPRUnit.Text
        vPRLine = Me.TBLineNumber.Text

        vQuery = "exec dbo.USP_NP_UpdateSTKRequestSub '" & vPRNo & "','" & vItemCode & "','" & vPRDate & "'," & vEditQty & ",'" & vUnitCode & "'," & vPRLine & ",'" & vUserName & "' "
        Call vExecData(vMemProfit, vQuery)

        Me.TBPRNo.Text = ""
        Me.TBPRDate.Text = ""
        Me.TBShowItem.Text = ""
        Me.TBEditQty.Text = ""
        Me.TBPRUnit.Text = ""
        Me.TBLineNumber.Text = ""
        Me.PNShowHist.Visible = False

        MsgBox("Edit Qty For Request Complete", MsgBoxStyle.Information, "Send Information Message")
        Call ClearItem()
        Me.TBBarCode.Focus()
    End Sub

    Private Sub BTNReqCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNReqCancel.Click
        On Error Resume Next

        Me.PNShowHist.Visible = False
        If Me.TBEditQty.ReadOnly = True Then
            Me.TBQty.Focus()
            Exit Sub
        End If

        Call ClearItem()
        Me.TBBarCode.Focus()
    End Sub

    Private Sub TBEditQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBEditQty.KeyDown, TBShowItem.KeyDown, TBPRNo.KeyDown, TBPRDate.KeyDown, TBPRQty.KeyDown, TBPRUnit.KeyDown, TBConfirmNo.KeyDown, TBConfirmQty.KeyDown, TBPONo.KeyDown, TBPOQty.KeyDown, TBPOLeadDate.KeyDown
        Dim vPRNo As String
        Dim vPRDate As String
        Dim vItemCode As String
        Dim vEditQty As Double
        Dim vUnitCode As String
        Dim vPRLine As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then

            If Me.TBEditQty.Text = "" And Me.TBEditQty.ReadOnly = False Then
                MsgBox("You are not Edit Qty", MsgBoxStyle.Critical, "Send Information Message")
                Me.TBEditQty.Focus()
                Exit Sub
            ElseIf Me.TBEditQty.ReadOnly = True Then
                Me.TBPRNo.Text = ""
                Me.TBPRDate.Text = ""
                Me.TBShowItem.Text = ""
                Me.TBEditQty.Text = ""
                Me.TBPRUnit.Text = ""
                Me.TBLineNumber.Text = ""
                Me.PNShowHist.Visible = False
                Me.TBQty.Focus()
                Exit Sub
            End If

            vPRNo = Me.TBPRNo.Text
            vPRDate = Me.TBPRDate.Text
            vItemCode = Me.TBItemCode.Text
            vEditQty = Me.TBEditQty.Text
            vUnitCode = Me.TBPRUnit.Text
            vPRLine = Me.TBLineNumber.Text

            vQuery = "exec dbo.USP_NP_UpdateSTKRequestSub '" & vPRNo & "','" & vItemCode & "','" & vPRDate & "'," & vEditQty & ",'" & vUnitCode & "'," & vPRLine & ",'" & vUserName & "' "
            Call vExecData(vMemProfit, vQuery)

            Me.TBPRNo.Text = ""
            Me.TBPRDate.Text = ""
            Me.TBShowItem.Text = ""
            Me.TBEditQty.Text = ""
            Me.TBPRUnit.Text = ""
            Me.TBLineNumber.Text = ""
            Me.PNShowHist.Visible = False

            MsgBox("Edit Qty For Request Complete", MsgBoxStyle.Information, "Send Information Message")
            Call ClearItem()
            Me.TBBarCode.Focus()

        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNShowHist.Visible = False
            If Me.TBEditQty.ReadOnly = True Then
                Me.TBQty.Focus()
                Exit Sub
            End If

            Call ClearItem()
            Me.TBBarCode.Focus()
        End If
    End Sub

    Private Sub TBEditQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBEditQty.KeyPress
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

    Private Sub TBEditQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBEditQty.TextChanged
        If vb6.InStr(Me.TBEditQty.Text, "@") > 0 Then
            Me.TBEditQty.Text = ""
            Me.TBEditQty.Focus()
        End If
    End Sub

    Private Sub Label19_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label19.ParentChanged

    End Sub
End Class


