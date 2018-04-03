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

Public Class FormReqPromo
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

        On Error GoTo ErrDescription


        If e.KeyCode = Keys.Escape Then
            'Call ClearItem()
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Enter Then

            If vb6.InStr(Me.TBBarCode.Text, "@") <> 0 Then
                vBarCode = vb6.Left(Me.TBBarCode.Text, vb6.Len(Me.TBBarCode.Text) - 1)
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

            vQuery = "exec dbo.usp_hh_SearchItemDataDetails_Cat '" & vMemProfit & "','" & vBarCode & "'"
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
                vUnitCode = pds.Tables(0).Rows(0)("unitcode").ToString
                vBarCode = pds.Tables(0).Rows(0)("barcode").ToString
                vOrderPoint = pds.Tables(0).Rows(0)("orderpoint").ToString
                vItemStatus = pds.Tables(0).Rows(0)("itemstatus").ToString
                vPORemainIn = pds.Tables(0).Rows(0)("remaininqty").ToString
                vSumCashSale3Month = pds.Tables(0).Rows(0)("avgsale1Month").ToString
                vFreq3Month = pds.Tables(0).Rows(0)("avgcountbill1Month").ToString
                vRedDot = pds.Tables(0).Rows(0)("reddot").ToString
                vMyGrade = pds.Tables(0).Rows(0)("mygrade").ToString
                vExpertTeam = pds.Tables(0).Rows(0)("expertteam").ToString
                vPRNo = pds.Tables(0).Rows(0)("prno").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBGrade.Text = vMyGrade
                Me.TBDisCount.Text = ""
                Me.TBPromoPrice.Text = ""
                Me.TBRemainQty.Text = Format(vStockQty, "##,##0.00")
                Me.TBUnit.Text = vUnitCode
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBItemStatus.Text = vItemStatus
                Me.TBPORemain.Text = Format(vPORemainIn, "##,##0.00")
                Me.TBSale1M.Text = Format(vSumCashSale3Month, "##,##0.00")

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

            Me.TBDisCount.Focus()
            Me.TBDisCount.SelectAll()

        End If


        If e.KeyCode = Keys.Down Then
            Me.TBDisCount.Focus()
            Me.TBDisCount.SelectAll()
        End If

        If e.KeyCode = 113 Then
            vAnswer = MsgBox("Do you want clear screen ?", MsgBoxStyle.YesNo, "Send Question Message")

            If vAnswer = 6 Then
                'Call ClearScreen()
            End If
        End If

        If e.KeyCode = 116 Then
            'Call SaveData()
        End If

        If e.KeyCode = 117 Then
            'Call vSearchStockRequest()
            Me.PNSearchDocNo.Visible = True
            Me.TBSearchDocNo.Focus()
        End If

        If e.KeyCode = 119 Then
            'Call CancelData()
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

        On Error GoTo ErrDescription

        If Me.TBBarCode.Text <> 0 Then
            vBarCode = Me.TBBarCode.Text

            Me.BTNRedDot.Visible = False
            Me.ListViewStock.Items.Clear()
            Me.ListViewStock.Visible = False
            Me.ListViewShelfID.Items.Clear()
            Me.ListViewShelfID.Visible = False

            vQuery = "exec dbo.usp_hh_SearchItemDataDetails_Cat '" & vMemProfit & "','" & vBarCode & "'"
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
                vOrderPoint = pds.Tables(0).Rows(0)("orderpoint").ToString
                vItemStatus = pds.Tables(0).Rows(0)("itemstatus").ToString
                vPORemainIn = pds.Tables(0).Rows(0)("remaininqty").ToString
                vSumCashSale3Month = pds.Tables(0).Rows(0)("avgsale1Month").ToString
                vFreq3Month = pds.Tables(0).Rows(0)("avgcountbill1Month").ToString
                vRedDot = pds.Tables(0).Rows(0)("reddot").ToString
                vMyGrade = pds.Tables(0).Rows(0)("mygrade").ToString
                vExpertTeam = pds.Tables(0).Rows(0)("expertteam").ToString
                vPRNo = pds.Tables(0).Rows(0)("prno").ToString

                Me.TBItemCode.Text = vItemCode
                Me.TBItemName.Text = vItemName
                Me.TBGrade.Text = vMyGrade
                Me.TBDisCount.Text = ""
                Me.TBPromoPrice.Text = ""
                Me.TBRemainQty.Text = Format(vStockQty, "##,##0.00")
                Me.TBUnit.Text = vUnitCode
                Me.TBPrice.Text = Format(vPrice, "##,##0.00")
                Me.TBPrice2.Text = Format(vPrice2, "##,##0.00")
                Me.TBItemStatus.Text = vItemStatus
                Me.TBPORemain.Text = Format(vPORemainIn, "##,##0.00")
                Me.TBSale1M.Text = Format(vSumCashSale3Month, "##,##0.00")

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

            Me.TBDisCount.Focus()
            Me.TBDisCount.SelectAll()

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
        vQuery = "exec dbo.USP_PM_FindType"
        Call vGetData(vMemProfit, vQuery)
        If pds.Tables(0).Rows.Count > 0 Then
            For i = 0 To pds.Tables(0).Rows.Count - 1
                Me.CMBPromoType.Items.Add(pds.Tables(0).Rows(i)("code").ToString + "/" + pds.Tables(0).Rows(i)("name1").ToString)
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
                Me.CMBSection.Items.Add(pds.Tables(0).Rows(i)("secmancode").ToString + "/" + pds.Tables(0).Rows(i)("secmanname").ToString)
            Next
        End If
    End Sub


    Public Sub GenNewDocNo()
        Dim vGetNewDocNo As String

        On Error Resume Next

        vQuery = "execute dbo.USP_PM_RequestNewDocNo"
        Call vGetData(vMemProfit, vQuery)
        If pds.Tables(0).Rows.Count > 0 Then
            vGetNewDocNo = pds.Tables(0).Rows(0)("newdocno").ToString
        End If
        Me.TBDocNo.Text = vGetNewDocNo
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        'Call ClearScreen()
        FormMainApplication.Show()
        Me.Hide()
    End Sub

    Private Sub FormReqPromo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call SearchCampaign()
        Call SearchSection()
        Call SearchPromoType()
    End Sub

    Private Sub BTNNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNNext.Click
        If Me.CMBCampaign.Text = "" Then
            MsgBox("กรุณาเลือก ทะเบียนโปรโมชั่น", MsgBoxStyle.Critical, "Send Error Message")
            Me.CMBCampaign.Focus()
            Exit Sub
        End If

        If Me.CMBSection.Text = "" Then
            MsgBox("กรุณาเลือก Section Manager", MsgBoxStyle.Critical, "Send Error Message")
            Me.CMBSection.Focus()
            Exit Sub
        End If

        If Me.CMBPromoType.Text = "" Then
            MsgBox("กรุณาเลือก ประเภทเสนอสินค้าเข้าร่วมโปรโมชั่น", MsgBoxStyle.Critical, "Send Error Message")
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

    Private Sub CMBCampaign_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBCampaign.SelectedIndexChanged
        Dim vTypeCode As String
        Dim vDateStart As Date
        Dim vDateStop As Date

        On Error Resume Next

        If Me.CMBCampaign.Items.Count > 0 Then

            vTypeCode = vb6.Left(Me.CMBCampaign.SelectedItem, vb6.InStr(Me.CMBCampaign.SelectedItem, "/") - 1)
            vQuery = "exec dbo.USP_PM_Find '" & vTypeCode & "'"
            Call vGetData(vMemProfit, vQuery)
            If pds.Tables(0).Rows.Count > 0 Then
                vDateStart = pds.Tables(0).Rows(0)("DateStart").ToString
                vDateStop = pds.Tables(0).Rows(0)("DateEnd").ToString

                Me.TBDateStart.Text = vb6.Day(vDateStart) + "/" + vb6.Month(vDateStart) + "/" + vb6.Year(vDateStart)
                Me.TBDateStop.Text = vb6.Day(vDateStop) + "/" + vb6.Month(vDateStop) + "/" + vb6.Year(vDateStop)

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

        Me.TBDocDate.Text = vMemDocDate
        vIsconfirm = 0
        vIsCancel = 0
        vMemReOrderIsOpen = 0
        Me.TBDocNo.Text = ""
        Me.TBDocDate.Text = vMemDocDate
        Me.TBBarCode.Text = ""
        Me.TBItemCode.Text = ""
        Me.TBItemName.Text = ""
        Me.TBDisCount.Text = ""
        Me.TBRemainQty.Text = ""
        Me.TBPromoPrice.Text = ""
        Me.TBUnit.Text = ""
        Me.TBPrice.Text = ""
        Me.TBItemStatus.Text = ""
        Me.TBPORemain.Text = ""
        Me.TBSale1M.Text = ""
        Me.TBGrade.Text = ""

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
        Me.TBDisCount.Text = ""
        Me.TBRemainQty.Text = ""
        Me.TBPromoPrice.Text = ""
        Me.TBUnit.Text = ""
        Me.TBPrice.Text = ""
        Me.TBItemStatus.Text = ""
        Me.TBPORemain.Text = ""
        Me.TBSale1M.Text = ""
        Me.TBGrade.Text = ""
        Me.BTNRedDot.Visible = False
        Me.ListViewStock.Items.Clear()
        Me.ListViewStock.Visible = False
        Me.ListViewShelfID.Items.Clear()
        Me.ListViewShelfID.Visible = False
        Me.TBBarCode.Focus()
        Me.TBBarCode.SelectAll()
    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        'Call vSearchStockRequest()
        Me.PNSearchDocNo.Visible = True
        Me.TBSearchDocNo.Focus()
    End Sub

    Private Sub BTNCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCancel.Click
        Call CancelData()
    End Sub

    Public Sub CancelData()
        Dim vDocNo As String
        Dim vAnswer As Integer

        On Error GoTo ErrDescription

        If Me.TBDocNo.Text <> "" And vMemReqProIsOpen = 1 And vIsconfirm = 0 And vIsCancel = 0 Then
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
        Dim vPromoname As String
        Dim vStartPromo As Date
        Dim vIsCancel As String
        Dim vCountItem As Integer
        Dim vSecName As String
        Dim vPromotionCode As String
        Dim i As Integer
        Dim vError As Integer
        Dim vIsCompleteSave As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vNormalPrice As Double
        Dim vFromQTY As Long, vToQty As Long
        Dim vDiscount As Double
        Dim vDiscountWord As String
        Dim vDiscountType As String
        Dim vPromotionPrice As Double
        Dim vMydescription As String
        Dim vLineNumber As Integer
        Dim vIsBrochure As String
        Dim vIsMember As String
        Dim vIsConfirm As Integer
        Dim vPromotionType As String
        Dim vPromotionTypeCode As String
        Dim vItemIsCancel As Integer
        Dim vCheckDeleteDocno As String
        Dim vCheckDuplicatePromotion As Integer
        Dim vCheckDuplicateDocno As String
        Dim vHotPrice As String

        Dim vDocNo As String



        If Trim(Me.TBDocNo.Text) <> "" And ListViewItem.Items.Count <> 0 Then
            vCountItem = ListViewItem.Items.Count
            If vCountItem > 0 Then
                If TBDocNo.Text <> "" Then

                    vStartPromo = Me.TBDateStart.Text
                    vSecName = vb6.Left(Me.CMBSection.SelectedItem, vb6.InStr(Me.CMBSection.SelectedItem, "/") - 1)
                    vPromotionCode = vb6.Left(Me.CMBCampaign.SelectedItem, vb6.InStr(Me.CMBCampaign.SelectedItem, "/") - 1)

                    Call Me.GenNewDocNo()

                    vDocNo = Me.TBDocNo.Text

                    vQuery = "execute USP_PM_InsertRequest " & 1 & ",'" & vDocNo & "','" & vStartPromo & "','" & vSecName & "','" & vPromotionCode & "','" & vUserID & "' "
                    Call vExecData(vMemProfit, vQuery)

                    For i = 1 To ListViewItem.Items.Count
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
                        vDiscountWord = 0
                        vPromotionPrice = Trim(ListViewItem.Items(i).SubItems(3).Text)
                        vMydescription = ""
                        vLineNumber = i - 1
                        vIsBrochure = 0
                        vIsMember = 0
                        vPromotionTypeCode = Trim(ListViewItem.Items(i).SubItems(7).Text)

                        If vPromotionTypeCode = "11" Then
                            vHotPrice = "S02"
                        Else
                            vHotPrice = ""
                        End If

                        vItemIsCancel = 0


                        vQuery = "exec USP_PM_ItemDuplicate  '" & Trim(vPromotionCode) & "', '" & Trim(vItemCode) & "','" & vUnitCode & "'  "
                        Call vGetData2(vMemProfit, vQuery)
                        If pds2.Tables(0).Rows.Count > 0 Then
                            vCheckDuplicatePromotion = pds2.Tables(0).Rows(0)("isduplicate").ToString
                            vCheckDuplicateDocno = pds2.Tables(0).Rows(0)("duplicate").ToString
                        End If

                        If vCheckDuplicatePromotion = 0 Then
                            vQuery = "execute USP_PM_InsertRequestSub " & vError & "," & vIsCompleteSave & ",'" & vDocNo & "','" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vNormalPrice & "," & vFromQTY & "," & vToQty & "," & vDiscount & ",'" & vDiscountType & "','" & vDiscountWord & "'," & vPromotionPrice & ",'" & vMydescription & "','" & vItemIsCancel & "'," & vLineNumber & ",'" & vIsBrochure & "','" & vIsMember & "' ,'" & vPromotionTypeCode & "','" & vHotPrice & "' "
                            Call vExecData(vMemProfit, vQuery)
                        Else
                            vQuery = "execute USP_PM_InsertRequestSub 1," & vIsCompleteSave & ",'" & vDocNo & "','" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vNormalPrice & "," & vFromQTY & "," & vToQty & "," & vDiscount & ",'" & vDiscountType & "','" & vDiscountWord & "'," & vPromotionPrice & ",'" & vMydescription & "','" & vItemIsCancel & "'," & vLineNumber & ",'" & vIsBrochure & "','" & vIsMember & "' ,'" & vPromotionTypeCode & "','" & vHotPrice & "'  "
                            Call vExecData(vMemProfit, vQuery)
                            MsgBox("รายการสินค้า รหัส " & vItemCode & " ในโปรโมชั่นนี้มีอยู่แล้ว ในเอกสารเลขที่ " & vCheckDuplicateDocno & " กรุณาตรวจสอบ", vbCritical, "Send Error")
                            Exit Sub
                        End If

                    Next i
                    MsgBox("ได้เอกสารเลขที่  " & vDocNo & " ")

                    vQuery = "execute USP_PM_DeliverySendMail '" & vDocNo & "' "
                    Call vExecData(vMemProfit, vQuery)

                    vCheckDeleteDocno = Trim(TBDocNo.Text)
                    vQuery = "USP_PM_DeleteCheckDuplicatItemLine '" & vCheckDeleteDocno & "','" & vPromotionCode & "','" & vUserID & "' "
                    Call vExecData(vMemProfit, vQuery)


                End If
            End If
        End If



    End Sub

    Private Sub BTNAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNAdd.Click
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


        If Me.TBBarCode.Text = "" Then
            MsgBox("Please Insert ItemCode", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBBarCode.Focus()
            Exit Sub
        End If

        If Me.TBPromoPrice.Text = "" Then
            MsgBox("Please Insert Request Qty", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBDisCount.Focus()
            Exit Sub
        End If



        If Me.TBPrice.Text <> "" Then
            vCountStkQty = Me.TBPrice.Text
        Else
            vCountStkQty = 0
        End If


        If Me.TBSale1M.Text <> "" Then
            vGetSale1Month = Me.TBSale1M.Text
        Else
            vGetSale1Month = 0
        End If




        vUnitCode = Me.TBUnit.Text
        vDocDate = Now
        vGrade = Me.TBGrade.Text

        If vSumQtyCheckOrder > vGetSale1Month Then
            vAnswerBuyOverOrder = MsgBox("This Order+StockOnHand is over AverageSale1Month ! Do you want buy this item ?", MsgBoxStyle.YesNo, "Send Question Message")
            'MsgBox("This item is over stock max " & vGrade & " unable to Re-Order", MsgBoxStyle.Critical, "Send Error Message")
            If vAnswerBuyOverOrder = 7 Then
                Me.TBDisCount.Focus()
                Exit Sub
            End If
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
                            vNewQty = Me.TBDisCount.Text
                            vSumQtyCheckOrder = vNewQty + vCountStkQty

                            If vSumQtyCheckOrder > vGetSale1Month Then
                                vAnswerBuyOverOrder = MsgBox("This Order Over AverageSale1Month ! Do you want buy this item ?", MsgBoxStyle.YesNo, "Send Question Message")
                                'MsgBox("This item is over stock for 3 month " & vGrade & " unable to Re-Order", MsgBoxStyle.Critical, "Send Error Message")
                                If vAnswerBuyOverOrder = 7 Then
                                    Me.TBDisCount.Focus()
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
                                    Me.TBDisCount.Focus()
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



ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub
End Class