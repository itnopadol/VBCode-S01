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


Public Class FormReqPromotion
    Private MyEventHandler As System.EventHandler = Nothing

    Private MyReadNotifyHander As System.EventHandler = Nothing
    Private MyStatusNotifyHandler As System.EventHandler = Nothing
    Private MyActivateHandler As System.EventHandler = Nothing
    Private MyDeActivateHandler As System.EventHandler = Nothing

    Dim vQuery As String
    Dim vMemDocDate As String
    Dim vMemIsCancel As Integer
    Dim vMemIsConfirm As Integer

    Private Sub TBBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
        Dim vAnswer As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vStkUnit As String
        Dim vBarCode As String
        Dim vPrice As Double
        Dim vPrice2 As Double
        Dim vStockQty As Double
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


        Dim vPRNo As String
        Dim vQty As Double
        Dim vLine As Integer
        Dim vItemUnit As String
        Dim vPromotion As String
        Dim vPromoExpire As String

        On Error GoTo ErrDescription


        If e.KeyCode = Keys.Escape Then
            Call ClearItem()
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Enter Then

            If Me.TBBarCode.Text <> "" Then
                vBarCode = Me.TBBarCode.Text
            Else
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                Exit Sub
            End If

            Me.BTNRedDot.Visible = False
            Me.ListViewStock.Items.Clear()
            Me.ListViewStock.Visible = False
            Me.ListViewShelfID.Items.Clear()
            Me.ListViewShelfID.Visible = False

            vQuery = "exec dbo.USP_NP_SearchItemData '" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)

            vPRNo = ""

            vItemCode = ""
            vItemName = ""
            vPrice = 0
            vUnitCode = ""
            vBarCode = ""

            vOrderPoint = 0
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
                vPrice2 = pds.Tables(0).Rows(0)("saleprice2").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString
                vRedDot = pds.Tables(0).Rows(0)("reddot").ToString
                vMyGrade = pds.Tables(0).Rows(0)("mygrade").ToString
                vPromotion = pds.Tables(0).Rows(0)("pmcode").ToString
                vPromoExpire = pds.Tables(0).Rows(0)("pmexpire").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBGrade.Text = vMyGrade
                Me.TBDisCountWord.Text = ""
                Me.TBPromoPrice.Text = ""
                Me.TBRemainQty.Text = Format(vStockQty, "##,##0.00")
                Me.TBUnit.Text = vUnitCode
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBPromotion.Text = vItemStatus
                Me.TBPromotion.Text = vPromotion
                Me.TBPromoExpire.Text = vPromoExpire

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

                Me.TBItemCode.Text = ""
                Me.TBItemName.Text = ""
                Me.TBGrade.Text = ""
                Me.TBDisCountWord.Text = ""
                Me.TBPromoPrice.Text = ""
                Me.TBRemainQty.Text = ""
                Me.TBUnit.Text = ""
                Me.TBPrice.Text = ""
                Me.TBPromotion.Text = ""
                Me.TBPromotion.Text = ""
                Me.TBPromoExpire.Text = ""
                Me.ListViewShelfID.Items.Clear()
                Me.ListViewStock.Items.Clear()

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

            Me.TBDisCountWord.Focus()
            Me.TBDisCountWord.SelectAll()

        End If


        If e.KeyCode = Keys.Down Then
            Me.TBDisCountWord.Focus()
            Me.TBDisCountWord.SelectAll()
        End If

        'If e.KeyCode = 113 Then
        '    vAnswer = MsgBox("Do you want clear screen ?", MsgBoxStyle.YesNo, "Send Question Message")

        '    If vAnswer = 6 Then
        '        'Call ClearScreen()
        '    End If
        'End If

        'If e.KeyCode = 116 Then
        '    'Call SaveData()
        'End If

        'If e.KeyCode = 117 Then
        '    Call vSearchStockRequest()
        '    Me.PNSearchDocNo.Visible = True
        '    Me.TBSearchDocNo.Focus()
        'End If


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
        Dim vPrice2 As Double
        Dim vDisCount As Double
        Dim vStockQty As Double
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


        Dim vPRNo As String
        Dim vQty As Double
        Dim vLine As Integer
        Dim vItemUnit As String
        Dim vPromotion As String
        Dim vPromoExpire As String

        On Error GoTo ErrDescription

        If vb6.InStr(Me.TBBarCode.Text, "@") > 0 Then

            vBarCode = vb6.Left(Me.TBBarCode.Text, vb6.Len(Me.TBBarCode.Text) - 1)

            Me.TBBarCode.Text = vBarCode

            Me.BTNRedDot.Visible = False
            Me.ListViewStock.Items.Clear()
            Me.ListViewStock.Visible = False
            Me.ListViewShelfID.Items.Clear()
            Me.ListViewShelfID.Visible = False

            vQuery = "exec dbo.USP_NP_SearchItemData '" & vBarCode & "'"
            Call vGetData(vMemProfit, vQuery)


            vItemCode = ""
            vItemName = ""
            vPrice = 0
            vPrice2 = 0
            vUnitCode = ""
            vBarCode = ""
            vDisCount = 0

            vItemStatus = ""
            vPORemainIn = 0
            vSumCashSale3Month = 0
            vFreq3Month = 0
            vRedDot = 0
            vMyGrade = ""

            If pds.Tables(0).Rows.Count > 0 Then
                vItemCode = pds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(0)("itemname").ToString
                vPrice = pds.Tables(0).Rows(0)("saleprice1").ToString
                vPrice2 = pds.Tables(0).Rows(0)("saleprice2").ToString
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString
                vRedDot = pds.Tables(0).Rows(0)("reddot").ToString
                vMyGrade = pds.Tables(0).Rows(0)("mygrade").ToString
                vPromotion = pds.Tables(0).Rows(0)("pmcode").ToString
                vPromoExpire = pds.Tables(0).Rows(0)("pmexpire").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBGrade.Text = vMyGrade
                Me.TBDisCountWord.Text = ""
                Me.TBPromoPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBUnit.Text = vUnitCode
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBPrice2.Text = Format(vPrice2, "##,##0.00")
                Me.TBPromotion.Text = vItemStatus
                Me.TBPromotion.Text = vPromotion
                Me.TBPromoExpire.Text = vPromoExpire

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

                Me.TBItemCode.Text = ""
                Me.TBItemName.Text = ""
                Me.TBGrade.Text = ""
                Me.TBDisCountWord.Text = ""
                Me.TBPromoPrice.Text = ""
                Me.TBRemainQty.Text = ""
                Me.TBUnit.Text = ""
                Me.TBPrice.Text = ""
                Me.TBPromotion.Text = ""
                Me.TBPromotion.Text = ""
                Me.TBPromoExpire.Text = ""
                Me.ListViewShelfID.Items.Clear()
                Me.ListViewStock.Items.Clear()

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

            Me.TBDisCountWord.Focus()
            Me.TBDisCountWord.SelectAll()

        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Public Sub SearchPromoType()
        Dim i As Integer
        Dim n As Integer
        Dim vTypeCode As String
        Dim vTypeName As String

        On Error Resume Next

        Me.CMBPromoType.Items.Clear()
        vQuery = "exec dbo.USP_PM_FindPromoType"
        Call vGetData(vMemProfit, vQuery)
        If pds.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds.Tables(0).Rows.Count - 1
                Me.CMBPromoType.Items.Add(pds.Tables(0).Rows(i)("name1").ToString)
            Next
        End If
    End Sub

    Public Sub SearchCampaign()
        Dim i As Integer
        Dim n As Integer
        Dim vTypeCode As String
        Dim vTypeName As String

        On Error Resume Next

        Me.CMBCampaign.Items.Clear()
        vQuery = "exec dbo.USP_PM_Find"
        Call vGetData(vMemProfit, vQuery)
        If pds.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds.Tables(0).Rows.Count - 1
                Me.CMBCampaign.Items.Add(pds.Tables(0).Rows(i)("pmcode").ToString + "/" + pds.Tables(0).Rows(i)("pmname").ToString)
            Next
        End If
    End Sub


    Public Sub SearchSection()
        Dim i As Integer
        Dim n As Integer
        Dim vTypeCode As String
        Dim vTypeName As String

        On Error Resume Next

        Me.CMBSection.Items.Clear()
        vQuery = "exec dbo.USP_PM_FindSecMan"
        Call vGetData(vMemProfit, vQuery)
        If pds.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds.Tables(0).Rows.Count - 1
                Me.CMBSection.Items.Add(pds.Tables(0).Rows(i)("secmancode").ToString)
            Next
        End If
    End Sub


    Public Sub GenNewDocNo()
        Dim vGetNewDocNo As String

        On Error Resume Next

        vQuery = "execute dbo.USP_PM_RequestPromoDocNo"
        Call vGetData(vMemProfit, vQuery)
        If pds.Tables(0).Rows.Count > 0 Then
            vGetNewDocNo = pds.Tables(0).Rows(0)("newdocno").ToString
        End If
        Me.TBDocNo.Text = vGetNewDocNo
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Call ClearScreen()
        FormMainApplication.Show()
        Me.Hide()
    End Sub

    Private Sub FormReqPromo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        Call GenNewDocNo()

        vQuery = "exec dbo.USP_NP_CheckDocDate"
        Call vGetData1(vMemProfit, vQuery)
        If pds1.Tables(0).Rows.Count > 0 Then
            vMemDocDate = pds1.Tables(0).Rows(0)("vdocdate").ToString
        End If
        Me.TBDocDate.Text = vMemDocDate


        Call SearchCampaign()
        Call SearchSection()
        Call SearchPromoType()

        Me.CMBCampaign.Focus()
    End Sub

    Private Sub BTNNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNNext.Click
        On Error Resume Next

        If Me.CMBCampaign.Text = "" Then
            MsgBox("Please select campaign", MsgBoxStyle.Critical, "Send Error Message")
            Me.CMBCampaign.Focus()
            Exit Sub
        End If

        If Me.CMBSection.Text = "" Then
            MsgBox("Please select section manager", MsgBoxStyle.Critical, "Send Error Message")
            Me.CMBSection.Focus()
            Exit Sub
        End If

        If Me.CMBPromoType.Text = "" Then
            MsgBox("Please select promotion type", MsgBoxStyle.Critical, "Send Error Message")
            Me.CMBPromoType.Focus()
            Exit Sub
        End If

        Me.PNHeader.Visible = False
        Me.PNSearchDocNo.Visible = False
        Me.TBBarCode.Focus()
        Me.TBBarCode.SelectAll()
    End Sub

    Private Sub BTNNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNNew.Click
        Dim vAnswer As Integer

        On Error Resume Next

        vAnswer = MsgBox("Do you want clear screen ?", MsgBoxStyle.YesNo, "Send Question Message")

        If vAnswer = 6 Then
            Call ClearScreen()
            Me.PNHeader.Visible = True
            Me.PNHeader.BringToFront()
        End If


    End Sub

    Private Sub CMBCampaign_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBCampaign.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.CMBSection.Focus()
        End If
    End Sub

    Private Sub CMBCampaign_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBCampaign.SelectedIndexChanged
        Dim vTypeCode As String
        Dim vDateStart As Date
        Dim vDateStop As Date

        On Error Resume Next

        If Me.CMBCampaign.Items.Count > 0 And Me.CMBCampaign.Text <> "" Then
            vTypeCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
            vQuery = "exec dbo.USP_PM_Find '" & vTypeCode & "'"
            Call vGetData(vMemProfit, vQuery)
            If pds.Tables(0).Rows.Count > 0 Then
                vDateStart = pds.Tables(0).Rows(0)("DateStart").ToString
                vDateStop = pds.Tables(0).Rows(0)("DateEnd").ToString

                Me.TBDateStart.Text = vDateStart.Day & "/" & vDateStart.Month & "/" & vDateStart.Year
                Me.TBDateStop.Text = vDateStop.Day & "/" & vDateStop.Month & "/" & vDateStop.Year

            End If
        End If
    End Sub

    Public Sub ClearScreen()

        On Error Resume Next

        vQuery = "exec dbo.USP_NP_CheckDocDate"
        Call vGetData1(vMemProfit, vQuery)
        If pds1.Tables(0).Rows.Count > 0 Then
            vMemDocDate = pds1.Tables(0).Rows(0)("vdocdate").ToString
        End If

        Me.TBDocNo.Text = ""

        Call GenNewDocNo()

        vMemIsConfirm = 0
        vMemIsCancel = 0
        vMemReqProIsOpen = 0

        Me.TBDateStart.Text = ""
        Me.TBDateStop.Text = ""

        Me.TBDocDate.Text = vMemDocDate
        Me.TBBarCode.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBDisCountWord.Text = ""
        Me.TBDisCount.Text = ""
        Me.TBRemainQty.Text = ""
        Me.TBPromoPrice.Text = ""
        Me.TBUnit.Text = ""
        Me.TBPrice.Text = ""
        Me.TBPrice2.Text = ""
        Me.TBPromotion.Text = ""
        Me.TBPromotion.Text = ""
        Me.TBPromoExpire.Text = ""
        Me.TBGrade.Text = ""

        Me.BTNRedDot.Visible = False
        Me.ListViewStock.Items.Clear()
        Me.ListViewStock.Visible = False
        Me.ListViewShelfID.Items.Clear()
        Me.ListViewShelfID.Visible = False
        Me.ListViewItem.Items.Clear()

        Me.PNHeader.Visible = True
        Me.PNSearchDocNo.Visible = False
        Me.PNHeader.BringToFront()

        Me.CMBCampaign.Focus()
    End Sub

    Public Sub ClearItem()
        On Error Resume Next

        Me.TBBarCode.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBDisCountWord.Text = ""
        Me.TBDisCount.Text = ""
        Me.TBRemainQty.Text = ""
        Me.TBPromoPrice.Text = ""
        Me.TBUnit.Text = ""
        Me.TBPrice.Text = ""
        Me.TBPrice2.Text = ""
        Me.TBPromotion.Text = ""
        Me.TBPromotion.Text = ""
        Me.TBPromoExpire.Text = ""
        Me.TBGrade.Text = ""
        Me.CBPercent.Checked = False
        Me.BTNRedDot.Visible = False
        Me.ListViewStock.Items.Clear()
        Me.ListViewStock.Visible = False
        Me.ListViewShelfID.Items.Clear()
        Me.ListViewShelfID.Visible = False
        Me.TBBarCode.Focus()
        Me.TBBarCode.SelectAll()
    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        On Error Resume Next

        Call vSearchRequestPromo()
        Me.PNSearchDocNo.Visible = True
        Me.PNSearchDocNo.BringToFront()
        Me.TBSearchDocNo.Focus()
    End Sub


    Public Sub vSearchRequestPromo()
        Dim i As Integer
        Dim n As Integer
        Dim vGetDocDate As Date
        Dim vDocDate As String
        Dim vSearch As String

        On Error Resume Next

        vSearch = Me.TBSearchDocNo.Text
        Me.ListViewSearchDocNo.Items.Clear()

        vQuery = "exec dbo.USP_PM_RequestPromoSearch '" & vSearch & "'"
        Call vGetData(vMemProfit, vQuery)

        If pds.Tables(0).Rows.Count > 0 Then

            For i = 0 To pds.Tables(0).Rows.Count - 1
                n = i + 1

                vGetDocDate = pds.Tables(0).Rows(i)("docdate").ToString
                vDocDate = vGetDocDate.Day & "/" & vGetDocDate.Month & "/" & vGetDocDate.Year

                Dim listItem As New ListViewItem(n)
                listItem.SubItems.Add(pds.Tables(0).Rows(i)("docno").ToString)
                listItem.SubItems.Add(vDocDate)
                listItem.SubItems.Add(pds.Tables(0).Rows(i)("pmcode").ToString)
                listItem.SubItems.Add(pds.Tables(0).Rows(i)("secman").ToString)
                listItem.SubItems.Add(pds.Tables(0).Rows(i)("isconfirm").ToString)
                listItem.SubItems.Add(pds.Tables(0).Rows(i)("iscancel").ToString)
                Me.ListViewSearchDocNo.Items.Add(listItem)
            Next
        End If
    End Sub

    Private Sub BTNCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPrevious.Click
        Dim vAnswer As Integer

        On Error Resume Next

        Me.PNHeader.Visible = True
        Me.PNHeader.BringToFront()

    End Sub

    Public Sub CancelData()
        Dim vDocNo As String
        Dim vAnswer As Integer

        On Error GoTo ErrDescription

        If Me.TBDocNo.Text <> "" And vMemReqProIsOpen = 1 And vMemIsConfirm = 0 And vMemIsCancel = 0 Then
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

    Public Sub SaveData()
        Dim vStartPromo As Date
        Dim vCountItem As Integer
        Dim vSecName As String
        Dim vPromotionCode As String
        Dim i As Integer
        Dim a As Integer
        Dim vError As Integer
        Dim vIsCompleteSave As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vNormalPrice As Double
        Dim vFromQTY As Double, vToQty As Double
        Dim vDiscount As Double
        Dim vDiscountWord As String
        Dim vDiscountType As String
        Dim vPromotionPrice As Double
        Dim vMydescription As String
        Dim vLineNumber As Integer
        Dim vIsBrochure As String
        Dim vIsMember As String
        Dim vPromotionTypeCode As String
        Dim vItemIsCancel As Integer
        Dim vCheckDeleteDocno As String
        Dim vCheckDuplicatePromotion As Integer
        Dim vCheckDuplicateDocno As String
        Dim vHotPrice As String

        Dim vDocNo As String
        Dim vDocdate As String

        Dim vCheckItemCode As String
        Dim vCheckUnitCode As String
        Dim vIsSave As Integer
        Dim vIsInsert As Integer

        On Error GoTo ErrDescription


        If Trim(Me.TBDocNo.Text) <> "" And ListViewItem.Items.Count <> 0 Then
            vCountItem = ListViewItem.Items.Count
            If vCountItem > 0 Then

                If TBDocNo.Text <> "" Then

                    If vMemIsConfirm = 1 Then
                        MsgBox("This docno is confirm. Can not modify ", MsgBoxStyle.Critical, "Send Message Error")
                        Me.PNHeader.Visible = True
                        Me.PNHeader.BringToFront()
                        Me.TBDocNo.Focus()
                        Exit Sub
                    End If

                    If vMemIsCancel = 1 Then
                        MsgBox("This docno is cancel. Can not modify ", MsgBoxStyle.Critical, "Send Message Error")
                        Me.PNHeader.Visible = True
                        Me.PNHeader.BringToFront()
                        Me.TBDocNo.Focus()
                        Exit Sub
                    End If

                    vStartPromo = Me.TBDateStart.Text
                    vSecName = vb6.Left(Me.CMBSection.Text, vb6.InStr(Me.CMBSection.Text, "/") - 1)
                    vPromotionCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)
                    vPromotionTypeCode = vb6.Left(Me.CMBPromoType.Text, vb6.InStr(Me.CMBPromoType.Text, "/") - 1)

                    For a = 0 To Me.ListViewItem.Items.Count - 1

                        vCheckItemCode = Trim(ListViewItem.Items(a).SubItems(6).Text)
                        vCheckUnitCode = Trim(ListViewItem.Items(a).SubItems(5).Text)
                        vIsSave = Trim(ListViewItem.Items(a).SubItems(10).Text)

                        If vIsSave = 0 Then
                            vQuery = "exec USP_PM_ItemDuplicate  '" & Trim(vPromotionCode) & "', '" & Trim(vCheckItemCode) & "','" & vCheckUnitCode & "'  "
                            Call vGetData2(vMemProfit, vQuery)
                            If pds2.Tables(0).Rows.Count > 0 Then
                                vCheckDuplicatePromotion = pds2.Tables(0).Rows(0)("isduplicate").ToString
                                vCheckDuplicateDocno = pds2.Tables(0).Rows(0)("duplicate").ToString
                            End If


                            If vCheckDuplicatePromotion > 0 Then
                                MsgBox("This Item  " & vCheckItemCode & " have exist in docno  " & vCheckDuplicateDocno & " ", vbCritical, "Send Error")
                                Exit Sub
                            End If
                        End If
                    Next


                    If vMemReqProIsOpen = 0 Then
                        Call Me.GenNewDocNo()
                        vIsInsert = 1
                    Else
                        vIsInsert = 0
                    End If

                    vDocNo = Me.TBDocNo.Text
                    vDocdate = Me.TBDocDate.Text

                    vQuery = "execute USP_PM_InsertRequestPromo " & vIsInsert & ",'" & vDocNo & "','" & vDocdate & "','" & vSecName & "','" & vPromotionCode & "','" & vUserID & "' "
                    Call vExecData(vMemProfit, vQuery)

                    For i = 0 To ListViewItem.Items.Count - 1
                        vError = 0
                        If i = ListViewItem.Items.Count Then
                            vIsCompleteSave = 1
                        Else
                            vIsCompleteSave = 0
                        End If
                        vItemCode = Trim(ListViewItem.Items(i).SubItems(6).Text)
                        vIsCompleteSave = 1
                        vError = 0
                        vItemName = Trim(ListViewItem.Items(i).SubItems(1).Text)
                        vUnitCode = Trim(ListViewItem.Items(i).SubItems(5).Text)
                        vNormalPrice = Trim(ListViewItem.Items(i).SubItems(2).Text)
                        vFromQTY = 1
                        vToQty = 99999
                        If Trim(ListViewItem.Items(i).SubItems(4).Text) <> 2 Then
                            vDiscount = Trim(ListViewItem.Items(i).SubItems(4).Text)
                        Else
                            vDiscount = 0
                        End If
                        vDiscountType = 0
                        vDiscountWord = Trim(ListViewItem.Items(i).SubItems(9).Text)
                        vPromotionPrice = Trim(ListViewItem.Items(i).SubItems(3).Text)
                        vMydescription = ""
                        vLineNumber = i
                        vIsBrochure = 0
                        vIsMember = 0

                        If vPromotionTypeCode = "11" Then
                            vHotPrice = "S02"
                        Else
                            vHotPrice = ""
                        End If

                        vItemIsCancel = 0

                        vQuery = "execute USP_PM_InsertRequestPromoSub " & vError & "," & vIsCompleteSave & ",'" & vDocNo & "','" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vNormalPrice & "," & vFromQTY & "," & vToQty & "," & vDiscount & ",'" & vDiscountType & "','" & vDiscountWord & "'," & vPromotionPrice & ",'" & vMydescription & "','" & vItemIsCancel & "'," & vLineNumber & ",'" & vIsBrochure & "','" & vIsMember & "' ,'" & vPromotionTypeCode & "','" & vHotPrice & "' "
                        Call vExecData(vMemProfit, vQuery)

                    Next i

                    MsgBox("Save Complete DocNo is  " & vDocNo & " ")

                    vCheckDeleteDocno = Trim(TBDocNo.Text)
                    vQuery = "USP_PM_DeleteCheckDuplicatItemLine '" & vCheckDeleteDocno & "','" & vPromotionCode & "','" & vUserID & "' "
                    Call vExecData(vMemProfit, vQuery)

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

    Private Sub BTNAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAdd.Click
        Dim i As Integer
        Dim n As Integer
        Dim vCheckLine As Integer
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vItemName As String
        Dim vPrice As Double
        Dim vPromoPrice As Double
        Dim vPrice2 As Double
        Dim vUnitCode As String
        Dim vDocDate As String
        Dim vDisCount As Double
        Dim vDisCountWord As String
        Dim vCheckItemCode As String
        Dim vPromoType As String

        Dim vPromotionCode As String
        Dim vCheckDuplicatePromotion As Integer
        Dim vCheckDuplicateDocno As String

        Dim vAnswer As Integer


        On Error GoTo ErrDescription


        If Me.TBBarCode.Text = "" Then
            MsgBox("Please Insert ItemCode", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBBarCode.Focus()
            Exit Sub
        End If

        If Me.TBItemCode.Text = "" Then
            MsgBox("Please Select ItemCode", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBBarCode.Focus()
            Exit Sub
        End If

        If Me.TBPromoPrice.Text = "" Then
            MsgBox("Please Insert Promotion Price", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBDisCountWord.Focus()
            Exit Sub
        End If

        If Me.TBPrice.Text <> "" Then
            vPrice = Me.TBPrice.Text
        Else
            vPrice = 0
        End If

        If Me.TBPromoPrice.Text <> "" Then
            vPromoPrice = Me.TBPromoPrice.Text
        Else
            vPromoPrice = 0
        End If

        If Me.TBPrice2.Text <> "" Then
            vPrice2 = Me.TBPrice2.Text
        Else
            vPrice2 = 0
        End If

        If Me.TBDisCount.Text <> "" Then
            vDisCount = Me.TBDisCount.Text
        Else
            vDisCount = 0
        End If

        If Me.TBDisCountWord.Text <> "" Then
            vDisCountWord = Me.TBDisCountWord.Text
        Else
            vDisCountWord = ""
        End If

        vItemCode = Me.TBItemCode.Text
        vItemName = Me.TBItemName.Text
        vUnitCode = Me.TBUnit.Text
        vDocDate = Now
        vPromoType = vb6.Left(Me.CMBPromoType.Text, vb6.InStr(Me.CMBPromoType.Text, "/") - 1)

        vPromotionCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)


        vQuery = "exec USP_PM_ItemDuplicate  '" & Trim(vPromotionCode) & "', '" & Trim(vItemCode) & "','" & vUnitCode & "'  "
        Call vGetData2(vMemProfit, vQuery)
        If pds2.Tables(0).Rows.Count > 0 Then
            vCheckDuplicatePromotion = pds2.Tables(0).Rows(0)("isduplicate").ToString
            vCheckDuplicateDocno = pds2.Tables(0).Rows(0)("duplicate").ToString
        End If


        If vCheckDuplicatePromotion > 0 Then
            MsgBox("This Item  " & vItemCode & " have exist in docno  " & vCheckDuplicateDocno & " ", vbCritical, "Send Error")
            Me.TBBarCode.Focus()
            Exit Sub
        End If


        If Me.ListViewItem.Items.Count > 0 Then
            For n = 0 To Me.ListViewItem.Items.Count - 1
                vCheckItemCode = Me.ListViewItem.Items(n).SubItems(6).Text

                If vItemCode = vCheckItemCode Then
                    vAnswer = MsgBox("This item aleady exist at line " & n + 1 & " Do you want edit qty ?", MsgBoxStyle.YesNo, "Send Error Message")
                    If vAnswer = 6 Then
                        vAnswer = MsgBox("Click YES Replace Promotion Price", MsgBoxStyle.YesNo, "")
                        If vAnswer = 6 Then

                            Me.ListViewItem.Items(n).SubItems(3).Text = Format(vPromoPrice, "##,##0.00")
                            Me.ListViewItem.Items(n).SubItems(4).Text = Format(vDisCount, "##,##0.00")
                            Me.ListViewItem.Items(n).SubItems(9).Text = vDisCountWord
                            Me.ListViewItem.Items(n).SubItems(10).Text = 1
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
        listItem.SubItems.Add(vItemName)
        listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
        listItem.SubItems.Add(Format(vPromoPrice, "##,##0.00"))
        listItem.SubItems.Add(Format(vDisCount, "##,##0.00"))
        listItem.SubItems.Add(vUnitCode)
        listItem.SubItems.Add(vItemCode)
        listItem.SubItems.Add(vPromoType)
        listItem.SubItems.Add(vPrice2)
        listItem.SubItems.Add(vDisCountWord)
        listItem.SubItems.Add(0)

        Me.ListViewItem.Items.Add(listItem)

        If Me.ListViewItem.Items.Count > 0 Then
            vCheckLine = Me.ListViewItem.Items.Count
            Me.ListViewItem.Focus()
            VScrollBar1.Value = vCheckLine - 1
        End If

        Call ClearItem()



ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Call SaveData()
    End Sub

    Private Sub TBDisCountWord_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBDisCountWord.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.BTNAdd.Focus() 
        End If
    End Sub

    Private Sub TBDisCount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBDisCountWord.KeyPress
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

    Private Sub TBDisCount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBDisCountWord.TextChanged
        Dim vTextDiscount As String
        Dim vCutDiscount As String
        Dim vDiscountAmount As Double
        Dim vDiscount As Integer
        Dim vPrice As Double
        Dim vPrice2 As Double
        Dim vPromoPrice As Double

        On Error Resume Next


        If Me.TBDisCountWord.Text <> "" Then
            vTextDiscount = Me.TBDisCountWord.Text

            If Me.TBPrice.Text <> "" Then
                vPrice = Me.TBPrice.Text
            Else
                vPrice = 0
            End If

            If Me.TBPrice2.Text <> "" Then
                vPrice2 = Me.TBPrice2.Text
            Else
                vPrice2 = 0
            End If

            If vb6.Len(vTextDiscount) > 1 And vb6.InStr(vTextDiscount, "%") > 0 Then
                vCutDiscount = vb6.Left(vTextDiscount, vb6.InStr(vTextDiscount, "%") - 1)
                vDiscount = vCutDiscount                
                vPromoPrice = vPrice - ((vPrice * vDiscount) / 100)
                vDiscountAmount = vPrice - vPromoPrice
            Else
                vCutDiscount = vTextDiscount
                vDiscount = vCutDiscount
                vDiscountAmount = vDiscount
                vPromoPrice = vPrice - vDiscount
            End If

            'If vPromoPrice < vPrice2 Then
            '    MsgBox("Can not discount low than Price2", MsgBoxStyle.Critical, "Send Error Message")
            '    Me.TBDisCountWord.Focus()
            '    Me.TBDisCountWord.SelectAll()
            '    Exit Sub
            'End If

            Me.TBDisCount.Text = Format(vDiscountAmount, "##,##0.00")
            Me.TBPromoPrice.Text = Format(vPromoPrice, "##,##0.00")
        End If
    End Sub

    Private Sub ListViewItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItem.KeyDown
        Dim vIndex As Integer
        Dim vAnswer As Integer

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

    Private Sub ListViewSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearchDocNo.KeyDown
        Dim vIndex As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Enter And Me.ListViewSearchDocNo.Items.Count > 0 Then
            vIndex = Me.ListViewSearchDocNo.FocusedItem.Index
            Me.TBDocNo.Text = Me.ListViewSearchDocNo.Items(vIndex).SubItems(1).Text

            Me.PNSearchDocNo.Visible = False
            Me.PNHeader.Visible = True
            Me.PNHeader.BringToFront()
            Me.TBDocNo.Focus()
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

    Private Sub TBDocDate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBDocDate.TextChanged

    End Sub

    Public Sub RequestPromoDetails(ByVal vDocNo As String)
        Dim i As Integer
        Dim n As Integer
        Dim vGetDocDate As Date
        Dim vDocDate As String

        Dim vCampaign As String
        Dim vSection As String
        Dim vPromoType As String

        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vPrice As Double
        Dim vPromoPrice As Double
        Dim vDiscount As Double
        Dim vDiscountWord As String
        Dim vIsSave As Integer


        On Error Resume Next

        vMemIsConfirm = 0
        vMemIsCancel = 0

        Me.ListViewItem.Items.Clear()
        vQuery = "exec dbo.USP_PM_RequestPromoSubSearch '" & vDocNo & "'"
        Call vGetData(vMemProfit, vQuery)

        If pds.Tables(0).Rows.Count > 0 Then
            vMemReqProIsOpen = 1
            vMemIsConfirm = pds.Tables(0).Rows(0)("isconfirm").ToString
            vMemIsCancel = pds.Tables(0).Rows(0)("iscancel").ToString
            vGetDocDate = pds.Tables(0).Rows(0)("docdate").ToString
            vDocDate = vGetDocDate.Day & "/" & vGetDocDate.Month & "/" & vGetDocDate.Year
            Me.TBDocDate.Text = vDocDate

            vCampaign = pds.Tables(0).Rows(0)("pmname").ToString
            vSection = pds.Tables(0).Rows(0)("secname").ToString
            vPromoType = pds.Tables(0).Rows(0)("name1").ToString


            For i = 0 To pds.Tables(0).Rows.Count - 1
                n = i + 1

                vItemCode = pds.Tables(0).Rows(i)("itemcode").ToString
                vItemName = pds.Tables(0).Rows(i)("itemname").ToString
                vPrice = pds.Tables(0).Rows(i)("price").ToString
                vPromoPrice = pds.Tables(0).Rows(i)("promoprice").ToString
                vDiscount = pds.Tables(0).Rows(i)("discount").ToString
                vDiscountWord = pds.Tables(0).Rows(i)("discountword").ToString
                vUnitCode = pds.Tables(0).Rows(i)("unitcode").ToString
                vIsSave = 1

                Dim listItem As New ListViewItem(n)
                listItem.SubItems.Add(vItemName)
                listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                listItem.SubItems.Add(Format(vPromoPrice, "##,##0.00"))
                listItem.SubItems.Add(Format(vDiscount, "##,##0.00"))
                listItem.SubItems.Add(vUnitCode)
                listItem.SubItems.Add(vItemCode)
                listItem.SubItems.Add(vPromoType)
                listItem.SubItems.Add(Format(0, "##,##0.00"))
                listItem.SubItems.Add(vDiscountWord)
                listItem.SubItems.Add(1)
                Me.ListViewItem.Items.Add(listItem)
                Me.ListViewItem.Items(i).BackColor = Color.LightGreen
            Next
        End If

        Me.CMBCampaign.Text = vCampaign
        Me.CMBPromoType.Text = vPromoType
        Me.CMBSection.Text = vSection
        Me.CMBCampaign.Focus()
    End Sub

    Private Sub BTNCloseSearchDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSearchDoc.Click
        On Error Resume Next

        Me.PNSearchDocNo.Visible = False
        Me.PNHeader.Visible = True
        Me.PNHeader.BringToFront()
        Me.CMBCampaign.Focus()
    End Sub

    Private Sub CMBSection_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBSection.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.CMBPromoType.Focus()
        End If
    End Sub

    Private Sub CMBSection_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBSection.SelectedIndexChanged

    End Sub

    Private Sub TBDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBDocNo.TextChanged
        On Error Resume Next

        If Me.TBDocNo.Text <> "" Then
            Call RequestPromoDetails(Me.TBDocNo.Text)
        End If
        Me.TBDocNo.Focus()
    End Sub

    Private Sub CMBPromoType_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBPromoType.KeyDown

        On Error Resume Next


        If Me.CMBCampaign.Text <> "" And Me.CMBSection.Text <> "" And e.KeyCode = Keys.Enter Then
            Me.BTNNext.Focus()
        End If
    End Sub

    Private Sub CMBPromoType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBPromoType.SelectedIndexChanged

    End Sub

    Private Sub CBPercent_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBPercent.CheckStateChanged
        Dim vDiscount As String
        Dim vAddPercent As String

        On Error Resume Next

        If Me.CBPercent.Checked = True And Me.TBDisCountWord.Text <> "" Then
            vDiscount = Me.TBDisCountWord.Text
            If vb6.Len(vDiscount) > 1 And vb6.InStr(vDiscount, "%") > 0 Then
                Me.TBDisCountWord.Focus()
                Exit Sub
            Else
                vAddPercent = vDiscount + "%"
                Me.TBDisCountWord.Text = vAddPercent
                Me.TBDisCountWord.Focus()
            End If
        Else
            vDiscount = Me.TBDisCountWord.Text
            If vb6.Len(vDiscount) > 1 And vb6.InStr(vDiscount, "%") > 0 Then

                vAddPercent = vb6.Replace(vDiscount, "%", "")
                Me.TBDisCountWord.Text = vAddPercent
                Me.TBDisCountWord.Focus()
            Else
                Me.TBDisCountWord.Focus()
                Exit Sub
            End If
        End If
    End Sub

    Private Sub BTNSearchDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchDocNo.Click
        On Error Resume Next

        Call vSearchRequestPromo()
        Me.PNSearchDocNo.Visible = True
        Me.TBSearchDocNo.Focus()
    End Sub

    Private Sub BTNSelectDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectDocNo.Click
        Dim vIndex As Integer

        On Error Resume Next

        If Me.ListViewSearchDocNo.Items.Count > 0 Then
            vIndex = Me.ListViewSearchDocNo.FocusedItem.Index
            Me.TBDocNo.Text = Me.ListViewSearchDocNo.Items(vIndex).SubItems(1).Text

            Me.PNSearchDocNo.Visible = False
            Me.PNHeader.Visible = True
            Me.PNHeader.BringToFront()
            Me.TBDocNo.Focus()
        End If
    End Sub

    Private Sub ListViewSearchDocNo_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListViewSearchDocNo.Validated

    End Sub

    Private Sub TBSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchDocNo.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Call vSearchRequestPromo()
        End If

        If e.KeyCode = Keys.Down And Me.ListViewSearchDocNo.Items.Count > 0 Then
            Me.ListViewSearchDocNo.Focus()
            Me.ListViewSearchDocNo.Items(0).Selected = True
        End If
    End Sub

    Private Sub TBSearchDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchDocNo.TextChanged

    End Sub

    Private Sub BTNAdd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNAdd.KeyDown
        Dim i As Integer
        Dim n As Integer
        Dim vCheckLine As Integer
        Dim vItemCode As String
        Dim vBarCode As String
        Dim vItemName As String
        Dim vPrice As Double
        Dim vPromoPrice As Double
        Dim vPrice2 As Double
        Dim vUnitCode As String
        Dim vDocDate As String
        Dim vDisCount As Double
        Dim vDisCountWord As String
        Dim vCheckItemCode As String
        Dim vPromoType As String

        Dim vPromotionCode As String
        Dim vCheckDuplicatePromotion As Integer
        Dim vCheckDuplicateDocno As String

        Dim vAnswer As Integer


        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then

            If Me.TBBarCode.Text = "" Then
                MsgBox("Please Insert ItemCode", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Exit Sub
            End If

            If Me.TBItemCode.Text = "" Then
                MsgBox("Please Select ItemCode", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Exit Sub
            End If

            If Me.TBPromoPrice.Text = "" Then
                MsgBox("Please Insert Promotion Price", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBDisCountWord.Focus()
                Exit Sub
            End If

            If Me.TBPrice.Text <> "" Then
                vPrice = Me.TBPrice.Text
            Else
                vPrice = 0
            End If

            If Me.TBPromoPrice.Text <> "" Then
                vPromoPrice = Me.TBPromoPrice.Text
            Else
                vPromoPrice = 0
            End If

            If Me.TBPrice2.Text <> "" Then
                vPrice2 = Me.TBPrice2.Text
            Else
                vPrice2 = 0
            End If

            If Me.TBDisCount.Text <> "" Then
                vDisCount = Me.TBDisCount.Text
            Else
                vDisCount = 0
            End If

            If Me.TBDisCountWord.Text <> "" Then
                vDisCountWord = Me.TBDisCountWord.Text
            Else
                vDisCountWord = ""
            End If

            vItemCode = Me.TBItemCode.Text
            vItemName = Me.TBItemName.Text
            vUnitCode = Me.TBUnit.Text
            vDocDate = Now
            vPromoType = vb6.Left(Me.CMBPromoType.Text, vb6.InStr(Me.CMBPromoType.Text, "/") - 1)


            vPromotionCode = vb6.Left(Me.CMBCampaign.Text, vb6.InStr(Me.CMBCampaign.Text, "/") - 1)


            vQuery = "exec USP_PM_ItemDuplicate  '" & Trim(vPromotionCode) & "', '" & Trim(vItemCode) & "','" & vUnitCode & "'  "
            Call vGetData2(vMemProfit, vQuery)
            If pds2.Tables(0).Rows.Count > 0 Then
                vCheckDuplicatePromotion = pds2.Tables(0).Rows(0)("isduplicate").ToString
                vCheckDuplicateDocno = pds2.Tables(0).Rows(0)("duplicate").ToString
            End If


            If vCheckDuplicatePromotion > 0 Then
                MsgBox("This Item  " & vItemCode & " have exist in docno  " & vCheckDuplicateDocno & " ", vbCritical, "Send Error")
                Me.TBBarCode.Focus()
                Exit Sub
            End If


            If Me.ListViewItem.Items.Count > 0 Then
                For n = 0 To Me.ListViewItem.Items.Count - 1
                    vCheckItemCode = Me.ListViewItem.Items(n).SubItems(6).Text

                    If vItemCode = vCheckItemCode Then
                        vAnswer = MsgBox("This item aleady exist at line " & n + 1 & " Do you want edit qty ?", MsgBoxStyle.YesNo, "Send Error Message")
                        If vAnswer = 6 Then
                            vAnswer = MsgBox("Click YES Replace Promotion Price", MsgBoxStyle.YesNo, "")
                            If vAnswer = 6 Then

                                Me.ListViewItem.Items(n).SubItems(3).Text = Format(vPromoPrice, "##,##0.00")
                                Me.ListViewItem.Items(n).SubItems(4).Text = Format(vDisCount, "##,##0.00")
                                Me.ListViewItem.Items(n).SubItems(9).Text = vDisCountWord
                                Me.ListViewItem.Items(n).SubItems(10).Text = 1
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
            listItem.SubItems.Add(vItemName)
            listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
            listItem.SubItems.Add(Format(vPromoPrice, "##,##0.00"))
            listItem.SubItems.Add(Format(vDisCount, "##,##0.00"))
            listItem.SubItems.Add(vUnitCode)
            listItem.SubItems.Add(vItemCode)
            listItem.SubItems.Add(vPromoType)
            listItem.SubItems.Add(vPrice2)
            listItem.SubItems.Add(vDisCountWord)
            listItem.SubItems.Add(0)

            Me.ListViewItem.Items.Add(listItem)

            If Me.ListViewItem.Items.Count > 0 Then
                vCheckLine = Me.ListViewItem.Items.Count
                Me.ListViewItem.Focus()
                VScrollBar1.Value = vCheckLine - 1
            End If

            Call ClearItem()



ErrDescription:
            If Err.Description <> "" Then
                MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub BTNNext_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNNext.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.CMBCampaign.Text = "" Then
                MsgBox("Please select campaign", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBCampaign.Focus()
                Exit Sub
            End If

            If Me.CMBSection.Text = "" Then
                MsgBox("Please select section manager", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBSection.Focus()
                Exit Sub
            End If

            If Me.CMBPromoType.Text = "" Then
                MsgBox("Please select promotion type", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBPromoType.Focus()
                Exit Sub
            End If

            Me.PNHeader.Visible = False
            Me.PNSearchDocNo.Visible = False
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If
    End Sub
End Class