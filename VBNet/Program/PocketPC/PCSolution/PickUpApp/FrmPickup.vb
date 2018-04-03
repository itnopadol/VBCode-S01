Imports System.data
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms

Public Class FrmPickup
    Dim ds As DataSet
    Dim vQuery As String

    Dim vCountItemOld As Integer
    Dim vMemItemCodeOld(0) As String
    Dim vMemUnitCodeOld(0) As String
    Dim vMemWHCodeOld(0) As String
    Dim vMemShelfCodeOld(0) As String
    Dim vMemZoneIDOld(0) As String
    Dim vMemPickZoneOld(0) As String
    Dim vMemBarCodeOld(0) As String
    Dim vCountItemZoneOld As Integer

    Dim vUserCode As String
    Dim vPassWord As String

    Dim vMemSaleName As String

    Private Sub BeforeSaveData()
        On Error Resume Next

        Me.TBRefNo.Enabled = False
        Me.TBARCode.Enabled = False
        Me.TBSaleCode.Enabled = False
        Me.TBBarCode.Enabled = False
        Me.ListViewItem.Enabled = False
        Me.BTNBack.Enabled = False
        Me.BTNClearPickUp.Enabled = False
        Me.BTNSave.Enabled = False
        Me.BTNSearch.Enabled = False
        Me.BTNClosePickup.Enabled = False
        Me.BTNCancel.Enabled = False
    End Sub

    Private Sub AfterSaveData()
        On Error Resume Next

        Me.TBRefNo.Enabled = True
        Me.TBARCode.Enabled = True
        Me.TBSaleCode.Enabled = True
        Me.TBBarCode.Enabled = True
        Me.ListViewItem.Enabled = True
        Me.BTNBack.Enabled = True
        Me.BTNClearPickUp.Enabled = True
        Me.BTNSave.Enabled = True
        Me.BTNSearch.Enabled = True
        Me.BTNClosePickup.Enabled = True
        Me.BTNCancel.Enabled = True
        Me.TBSaleCode.Text = vMemSaleName
    End Sub

    Private Sub TBBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
        Dim vBarCode As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vPrice As Double
        Dim vRate As Integer
        Dim vUnitCode As String
        Dim i As Integer
        Dim vWHCode As String
        Dim vStore As String
        Dim vStkUnit As String
        Dim vStkQTY As Double
        Dim vReserveQTY As Double
        Dim vDefWHCode As String
        Dim vDefShelfCode As String
        Dim vShelfID As String
        Dim vZoneID As String
        Dim vBarCode1 As String
        Dim vLinePickZone As String

        On Error GoTo ErrDescription

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If


        If e.KeyCode = Keys.Up Then
            Me.TBSaleCode.Focus()
            Me.TBSaleCode.SelectAll()
        End If

        If e.KeyCode = Keys.Down Then
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(0).Selected = True
                Me.ListViewItem.Items(0).Focused = True
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBItem.Text = ""
            Me.TBItemName.Text = ""
            Me.TBPrice.Text = ""
            Me.TBUnit.Text = ""
            Me.TBWHCode.Text = ""
            Me.TBShelfCode.Text = ""
            Me.TBQTY.Text = ""
            Me.TBRate.Text = ""
            Me.TBReserve.Text = ""
            Me.TBMemBarCode.Text = ""
            Me.TBShelfID.Text = ""
            Me.TBZoneID.Text = ""
            Me.ListViewStock.Items.Clear()
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

            If Me.RDZone1.Checked = True Then
                vLinePickZone = "01"
            End If

            If Me.RDZone2.Checked = True Then
                vLinePickZone = "02"
            End If

            If Me.RDZone3.Checked = True Then
                vLinePickZone = "03"
            End If

            If Me.RDZone4.Checked = True Then
                vLinePickZone = "04"
            End If

            Me.ListViewStock.Items.Clear()

            vQuery = "exec dbo.USP_MB_SearchBarCode '" & vBarCode & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
            If ds.Tables(0).Rows.Count > 0 Then
                vItemCode = ds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = ds.Tables(0).Rows(0)("itemname").ToString
                vPrice = ds.Tables(0).Rows(0)("price").ToString
                vRate = ds.Tables(0).Rows(0)("rate").ToString
                vReserveQTY = ds.Tables(0).Rows(0)("reserveqty").ToString
                vUnitCode = ds.Tables(0).Rows(0)("unitcode").ToString
                vDefWHCode = ds.Tables(0).Rows(0)("defsalewhcode").ToString
                vDefShelfCode = ds.Tables(0).Rows(0)("defsaleshelf").ToString
                vShelfID = ds.Tables(0).Rows(0)("shelfid").ToString
                vZoneID = ds.Tables(0).Rows(0)("zoneid").ToString
                vBarCode1 = ds.Tables(0).Rows(0)("barcode").ToString

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    vWHCode = ds.Tables(0).Rows(i)("whcode").ToString
                    vStore = ds.Tables(0).Rows(i)("shelfcode").ToString
                    vStkUnit = ds.Tables(0).Rows(i)("stkunitcode").ToString
                    vStkQTY = ds.Tables(0).Rows(i)("stock").ToString

                    Dim listItem As New ListViewItem(vWHCode)
                    listItem.SubItems.Add(vStore)
                    listItem.SubItems.Add(Format(vStkQTY, "##,##0.00"))
                    listItem.SubItems.Add(vStkUnit)
                    Me.ListViewStock.Items.Add(listItem)
                Next

                Dim n As Integer
                Dim vCheckItemCode As String
                Dim vCheckQTY As Double
                Dim vCheckWHCode As String
                Dim vCheckShelfCode As String
                Dim vCheckZoneID As String
                Dim vCheckPickZone As String
                Dim vCheckUnitCode As String

                If Me.ListViewItem.Items.Count > 0 Then
                    For n = 0 To Me.ListViewItem.Items.Count - 1
                        vCheckItemCode = Me.ListViewItem.Items(n).SubItems(4).Text
                        vCheckUnitCode = Me.ListViewItem.Items(n).SubItems(3).Text
                        vCheckWHCode = Me.ListViewItem.Items(n).SubItems(7).Text
                        vCheckShelfCode = Me.ListViewItem.Items(n).SubItems(8).Text
                        vCheckZoneID = Me.ListViewItem.Items(n).SubItems(11).Text
                        vCheckPickZone = Me.ListViewItem.Items(n).SubItems(12).Text
                        vCheckQTY = Me.ListViewItem.Items(n).SubItems(2).Text

                        If vItemCode = vCheckItemCode And vUnitCode = vCheckUnitCode And vDefWHCode = vCheckWHCode And vDefShelfCode = vCheckShelfCode And vZoneID = vCheckZoneID And vLinePickZone = vCheckPickZone Then
                            Me.TBQTY.Text = Format(vCheckQTY, "##,##0.00")
                            GoTo Line1
                        End If
                    Next
                End If

Line1:
                If Me.TBQTY.Text = "" Then
                    Me.TBQTY.Text = 1
                End If
                Me.TBQTY.Focus()
                Me.TBQTY.SelectAll()
            Else
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                MsgBox("��辺���ʺ����� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If

            Me.TBItem.Text = vItemCode
            Me.TBItemName.Text = vItemName
            Me.TBPrice.Text = Format(vPrice, "##,##0.00")
            Me.TBRate.Text = Format(vRate, "##,##0.00")
            Me.TBReserve.Text = Format(vReserveQTY, "##,##0.00")
            Me.TBUnit.Text = vUnitCode
            Me.TBWHCode.Text = vDefWHCode
            Me.TBShelfCode.Text = vDefShelfCode
            Me.TBShelfID.Text = vShelfID
            Me.TBZoneID.Text = vZoneID
            Me.TBMemBarCode.Text = vBarCode1
        End If

        If e.KeyCode = Keys.Back Then
            Me.TBItem.Text = ""
            Me.TBItemName.Text = ""
            Me.TBPrice.Text = ""
            Me.TBUnit.Text = ""
            Me.TBWHCode.Text = ""
            Me.TBShelfCode.Text = ""
            Me.TBQTY.Text = ""
            Me.TBRate.Text = ""
            Me.TBReserve.Text = ""
            Me.TBMemBarCode.Text = ""
            Me.TBShelfID.Text = ""
            Me.TBZoneID.Text = ""
            Me.ListViewStock.Items.Clear()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBBarCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged
        On Error GoTo ErrDescription

        If Me.TBRefNo.Text = "" Then
            MsgBox("��سҡ�͡ ����¹ö��͹���С�͡��������� � �ء����", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBRefNo.Focus()
            Exit Sub
        End If

        If Me.TBBarCode.Text <> "" Then
            Me.TBRefNo.Enabled = False
            Me.TBARCode.Enabled = False
            Me.TBSaleCode.Enabled = False
            Me.PNItemDetails.Visible = True
            Me.PNItemDetails.BringToFront()
            Me.BTNSave.Visible = False
        Else
            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
            Me.PNItemDetails.Visible = False
            Me.PNDriveIn.Visible = True
            Me.PNDriveIn.BringToFront()
            Me.BTNSave.Visible = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vIndex As Integer

        On Error GoTo ErrDescription

        If Me.ListViewStock.Items.Count > 0 And Me.TBItem.Text <> "" And Me.TBQTY.Text <> "" Then
            vItemCode = Me.TBItem.Text
            vItemName = Me.TBItemName.Text
            vWHCode = Me.TBWHCode.Text
            vShelfCode = Me.TBShelfCode.Text
            vQTY = Me.TBQTY.Text
            vPrice = Me.TBPrice.Text
            vAmount = vQTY * vPrice
            vUnitCode = Me.TBUnit.Text
            vIndex = Me.ListViewItem.Items.Count + 1

            Dim listItem As New ListViewItem(vIndex)
            listItem.SubItems.Add(vItemName)
            listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
            listItem.SubItems.Add(vUnitCode)
            listItem.SubItems.Add(vItemCode)
            listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
            listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
            listItem.SubItems.Add(vWHCode)
            listItem.SubItems.Add(vShelfCode)
            Me.ListViewItem.Items.Add(listItem)

            Call CalcItemAmount()

            Me.TBItem.Text = ""
            Me.TBItemName.Text = ""
            Me.TBPrice.Text = ""
            Me.TBUnit.Text = ""
            Me.TBWHCode.Text = ""
            Me.TBShelfCode.Text = ""
            Me.TBQTY.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNItemDetails.Visible = False
            Me.TBBarCode.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub CalcItemAmount()
        Dim i As Integer
        Dim vAmount As Double
        Dim vSumAmount As Double

        On Error GoTo ErrDescription

        If Me.ListViewItem.Items.Count > 0 Then
            For i = 0 To Me.ListViewItem.Items.Count - 1
                vAmount = Me.ListViewItem.Items(i).SubItems(6).Text
                vSumAmount = vSumAmount + vAmount
            Next
            Me.TBItemAmount.Text = Format(vSumAmount, "##,##0.00")
        Else
            Me.TBItemAmount.Text = Format(0, "##,##0.00")
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub


    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vCountItem As Integer
        Dim vHeader As String
        Dim vNumber As Integer
        Dim vDocNumber As String

        Dim vDocNo As String
        Dim vDocDate As String
        Dim vARCode As String
        Dim vSaleCode As String
        Dim vMemberID As String
        Dim vRefNo As String
        Dim vTotalNetAmount As Double
        Dim vBeforeTaxAmount As Double
        Dim vTaxAmount As Double

        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vDiscountWord As String
        Dim vDiscountAmount As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vUserID As String
        Dim i As Integer
        Dim vBarCode As String
        Dim vShelfID As String
        Dim vZoneID As String
        Dim vLinePickZone As String
        Dim vLineNumber As Integer

        Dim a As Integer
        Dim b As Integer
        Dim vCheckItemCode As String
        Dim vCheckUnitCode As String
        Dim vCheckBarCode As String
        Dim vCheckWHCode As String
        Dim vCheckShelfCode As String
        Dim vCheckZoneID As String
        Dim vCheckPickZone As String

        Dim vOldItem As String
        Dim vOldUnit As String
        Dim vOldBar As String
        Dim vOldWH As String
        Dim vOldShelf As String
        Dim vOldZone As String
        Dim vOldPick As String
        Dim vOld As Integer

        Dim vCountItemPickZone As Integer
        Dim vItemPickZone As String
        Dim vCount As Integer
        Dim vQueZone As String

        Dim vCheckIsConfirm As Integer
        Dim vCheckHoldBillNo As String

        On Error GoTo ErrDescription

        If Me.ListViewItem.Items.Count > 0 And Me.TBItemAmount.Text <> "" Then
            vUserID = Me.TBUserID.Text

            MsgBox("���駵��� ��顴���� �����+�����Ţ 9 ���ͷӡ�úѹ�֡������", MsgBoxStyle.Information, "Send Information Message")

            If Me.TBRefNo.Text = "" Then
                MsgBox("��ͧ��͡�Ţ�����ҧ�ԧ��ǡ�͹�ѹ�֡�����ŷء����", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBRefNo.Focus()
                Me.TBRefNo.SelectAll()
                Exit Sub
            End If

            If Me.TBSaleCode.Text = "" Then
                MsgBox("��سҡ�͡ ���ʾ�ѡ�ҹ��� ��͹�ѹ�֡������", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBSaleCode.Focus()
                Me.TBSaleCode.SelectAll()
                Exit Sub
            End If


            If Me.TBDocNo.Text = "" Then
                vQuery = "exec dbo.usp_np_searchnewdocno 29"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count > 0 Then
                    vHeader = Trim(ds.Tables(0).Rows(0)("header").ToString)
                    vNumber = ds.Tables(0).Rows(0)("autonumber").ToString
                    vDocNumber = Trim(ds.Tables(0).Rows(0)("docnumber").ToString)
                End If

                vDocNo = Trim(vDocNumber & vHeader & "-" & Format(vNumber, "0000"))
            Else
                vDocNo = Me.TBDocNo.Text
            End If

            If vDocNo <> "" Then
                vDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year

                vQuery = "exec dbo.USP_NP_CheckDocDate"
                Dim vService7 As New WebReference.WebServiceCalc
                Dim ds7 As DataSet = vService7.vGetQueryAnlyzer(vQuery)
                If ds7.Tables(0).Rows.Count > 0 Then
                    vDocDate = ds7.Tables(0).Rows(0)("vdocdate").ToString
                End If


                vRefNo = Me.TBRefNo.Text

                If Me.RDZone1.Checked = True Then
                    vConnectZone = "01"
                    vQueZone = "A"
                ElseIf Me.RDZone2.Checked = True Then
                    vConnectZone = "02"
                    vQueZone = "B"
                ElseIf Me.RDZone3.Checked = True Then
                    vConnectZone = "03"
                    vQueZone = "C"
                ElseIf Me.RDZone4.Checked = True Then
                    vConnectZone = "04"
                    vQueZone = "D"
                End If

                For vCount = 0 To Me.ListViewItem.Items.Count - 1
                    vItemPickZone = Me.ListViewItem.Items(vCount).SubItems(12).Text
                    If vConnectZone = vItemPickZone Then
                        vCountItemPickZone = vCountItemPickZone + 1
                    End If
                Next

                If vCountItemPickZone = 0 Then
                    If vCountItemZoneOld = 0 Then
                        Call ClearSaveData()
                        Exit Sub
                    End If
                End If

                Dim vInstrSale As Integer
                Dim vLenSale As Integer

                If Me.TBARCode.Text = "1" Then
                    vARCode = "99999"
                Else
                    vARCode = Me.TBARCode.Text
                End If

                vInstrSale = InStr(Me.TBSaleCode.Text, "/")
                If vInstrSale = 0 Then
                    MsgBox("��سҡ�͡���ʾ�ѡ�ҹ���١��ͧ������������к���� �������ʾ�ѡ�ҹ���ǡ� Enter �ա����", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBSaleCode.Focus()
                    Me.TBSaleCode.SelectAll()
                    Exit Sub
                End If
                vLenSale = Len(Me.TBSaleCode.Text)
                vSaleCode = vb6.Left(Me.TBSaleCode.Text, vInstrSale - 1)

                vMemberID = Me.TBMemberID.Text
                vTotalNetAmount = Me.TBItemAmount.Text
                vBeforeTaxAmount = (vTotalNetAmount * 100) / 107
                vTaxAmount = vTotalNetAmount - vBeforeTaxAmount

                If vIsOpen = 0 Then
                    Call BeforeSaveData()
                    vQuery = "exec dbo.usp_np_insertdriveinslip1 '" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vMemberID & "','" & vSaleCode & "','" & vRefNo & "'," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalNetAmount & ",'" & vUserID & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

                    For i = 0 To Me.ListViewItem.Items.Count - 1
                        vItemCode = Me.ListViewItem.Items(i).SubItems(4).Text
                        vItemName = Me.ListViewItem.Items(i).SubItems(1).Text
                        vWHCode = Me.ListViewItem.Items(i).SubItems(7).Text
                        vShelfCode = Me.ListViewItem.Items(i).SubItems(8).Text
                        vQTY = Me.ListViewItem.Items(i).SubItems(2).Text
                        vPrice = Me.ListViewItem.Items(i).SubItems(5).Text
                        vAmount = Me.ListViewItem.Items(i).SubItems(6).Text
                        vUnitCode = Me.ListViewItem.Items(i).SubItems(3).Text
                        vBarCode = Me.ListViewItem.Items(i).SubItems(9).Text
                        vShelfID = Me.ListViewItem.Items(i).SubItems(10).Text
                        vZoneID = Me.ListViewItem.Items(i).SubItems(11).Text
                        vLinePickZone = Me.ListViewItem.Items(i).SubItems(12).Text
                        vDiscountWord = ""
                        vDiscountAmount = 0
                        vLineNumber = i

                        vQuery = "exec dbo.usp_np_insertdriveinslipsub1 '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "','" & vLinePickZone & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vAmount & ",'" & vBarCode & "','" & vSaleCode & "'," & vLineNumber & " "
                        Dim vService2 As New WebReference.WebServiceCalc
                        Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                    Next

                    vQuery = "exec dbo.usp_np_updatenewdocno 29"
                    Dim vService3 As New WebReference.WebServiceCalc
                    Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

                    MsgBox("�ѹ�֡���������Ţ����͡��� " & vDocNo & " ", MsgBoxStyle.Information, "Send Information Message")

                    Dim vAnswer As Integer

                    vAnswer = MsgBox("�س��ͧ����� �͡���价ӡ�� Check Out �������", MsgBoxStyle.YesNo, "Send Information Message")
                    If vAnswer = 6 Then
                        Call SendCheckQue(vDocNo, vDocDate, vConnectZone)
                        Call ClearSaveData()

                    Else
                        Call ClearSaveData()
                    End If


                End If


                If vIsOpen = 1 Then
                    Call BeforeSaveData()
                    vQuery = "exec dbo.USP_NP_CheckQueHoldBill1 '" & vDocNo & "','" & vARCode & "'"
                    Dim vService As New WebReference.WebServiceCalc
                    Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                    If ds.Tables(0).Rows.Count > 0 Then
                        vCheckIsConfirm = ds.Tables(0).Rows(0)("isconfirm").ToString()
                        vCheckHoldBillNo = ds.Tables(0).Rows(0)("holdbillno").ToString()
                    End If

                    If vCheckIsConfirm = 1 And vCheckHoldBillNo <> "" Then
                        MsgBox("�Ţ����͡�����ӡ�� ��Ǩ�ͺ�Թ�����з��͡��þѡ������º�������� �������ö�觤������͡����� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
                        Call ClearSaveData()
                        Call AfterSaveData()
                        Me.TBRefNo.Focus()
                        Me.TBRefNo.SelectAll()
                        Exit Sub
                    End If

                    vQuery = "exec dbo.usp_np_insertdriveinslip1 '" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vMemberID & "','" & vSaleCode & "','" & vRefNo & "'," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalNetAmount & ",'" & vUserID & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

                    vCountItem = Me.ListViewItem.Items.Count

                    For a = 0 To vCountItemOld
                        vOldItem = vMemItemCodeOld(a)
                        vOldUnit = vMemUnitCodeOld(a)
                        vOldBar = vMemBarCodeOld(a)
                        vOldWH = vMemWHCodeOld(a)
                        vOldShelf = vMemShelfCodeOld(a)
                        vOldZone = vMemZoneIDOld(a)
                        vOldPick = vMemPickZoneOld(a)

                        For b = 0 To Me.ListViewItem.Items.Count - 1
                            vCheckItemCode = Me.ListViewItem.Items(b).SubItems(4).Text
                            vCheckUnitCode = Me.ListViewItem.Items(b).SubItems(3).Text
                            vCheckBarCode = Me.ListViewItem.Items(b).SubItems(9).Text
                            vCheckWHCode = Me.ListViewItem.Items(b).SubItems(7).Text
                            vCheckShelfCode = Me.ListViewItem.Items(b).SubItems(8).Text
                            vCheckZoneID = Me.ListViewItem.Items(b).SubItems(11).Text
                            vCheckPickZone = Me.ListViewItem.Items(b).SubItems(12).Text

                            If vCheckItemCode = vOldItem And vCheckUnitCode = vOldUnit And vCheckBarCode = vOldBar And vCheckWHCode = vOldWH And vCheckShelfCode = vOldShelf And vCheckZoneID = vOldZone And vCheckPickZone = vOldPick Then
                                vOld = 1
                                GoTo Line1
                            Else
                                vOld = 0
                            End If
                        Next
Line1:

                        If vOld = 0 Then
                            vItemCode = vOldItem
                            vWHCode = vOldWH
                            vShelfCode = vOldShelf
                            vUnitCode = vOldUnit
                            vBarCode = vOldBar
                            vZoneID = vOldZone
                            vLinePickZone = vOldPick

                            vQuery = "exec dbo.usp_np_deletedriveinslipsub1 '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vZoneID & "','" & vLinePickZone & "','" & vUnitCode & "','" & vBarCode & "'," & vTotalNetAmount & " "
                            Dim vService2 As New WebReference.WebServiceCalc
                            Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                        End If
                    Next

                    For i = 0 To Me.ListViewItem.Items.Count - 1
                        vItemCode = Me.ListViewItem.Items(i).SubItems(4).Text
                        vItemName = Me.ListViewItem.Items(i).SubItems(1).Text
                        vWHCode = Me.ListViewItem.Items(i).SubItems(7).Text
                        vShelfCode = Me.ListViewItem.Items(i).SubItems(8).Text
                        vQTY = Me.ListViewItem.Items(i).SubItems(2).Text
                        vPrice = Me.ListViewItem.Items(i).SubItems(5).Text
                        vAmount = Me.ListViewItem.Items(i).SubItems(6).Text
                        vUnitCode = Me.ListViewItem.Items(i).SubItems(3).Text
                        vBarCode = Me.ListViewItem.Items(i).SubItems(9).Text
                        vShelfID = Me.ListViewItem.Items(i).SubItems(10).Text
                        vZoneID = Me.ListViewItem.Items(i).SubItems(11).Text
                        vLinePickZone = Me.ListViewItem.Items(i).SubItems(12).Text
                        vDiscountWord = ""
                        vDiscountAmount = 0
                        vLineNumber = i

                        If vConnectZone = vLinePickZone Then
                            vQuery = "exec dbo.usp_np_insertdriveinslipsub1 '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "','" & vLinePickZone & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vAmount & ",'" & vBarCode & "','" & vSaleCode & "'," & vLineNumber & " "
                            Dim vService3 As New WebReference.WebServiceCalc
                            Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)
                        End If
                    Next
                    MsgBox("����Ţ����͡��� " & vDocNo & " ���º��������", MsgBoxStyle.Information, "Send Information Message")

                    Me.TBRefNo.Enabled = True
                    Me.TBARCode.Enabled = True
                    Me.TBSaleCode.Enabled = True

                    Dim vAnswer As Integer

                    vAnswer = MsgBox("�س��ͧ����� �͡���价ӡ�� Check Out �������", MsgBoxStyle.YesNo, "Send Information Message")
                    If vAnswer = 6 Then

                        Dim m As Integer
                        Dim vQueItemCode As String
                        Dim vQueItemName As String
                        Dim vQueUnit As String
                        Dim vQueQty As Double
                        Dim vQueID As Integer
                        Dim vQueArName As String
                        Dim vQueSaleName As String
                        Dim vQueZoneID As String
                        Dim vQueRefNo As String
                        Dim vIndex As Integer
                        Dim vQueDocNo As String
                        Dim vQueWHCode As String
                        Dim vQueShelfCode As String
                        Dim vQueShelfID As String
                        Dim vQueBarCode As String
                        Dim vQuePickZone As String


                        vQuery = "exec dbo.USP_NP_CheckQueDriveIn1 '" & vDocNo & "','" & vDocDate & "','" & vQueZone & "' "
                        Dim vService4 As New WebReference.WebServiceCalc
                        Dim ds4 As DataSet = vService4.vGetQueryAnlyzer(vQuery)

                        If ds4.Tables(0).Rows.Count > 0 Then
                            Me.PNLastQueSend.Visible = True
                            Me.TBCarLicense.Text = Trim(ds4.Tables(0).Rows(0)("refno").ToString)
                            Me.TBQueAR.Text = Trim(ds4.Tables(0).Rows(0)("arcode").ToString) & "/" & Trim(ds4.Tables(0).Rows(0)("arname").ToString)

                            Me.ListViewItemLastSend.Items.Clear()
                            For m = 0 To ds4.Tables(0).Rows.Count - 1
                                vIndex = vIndex + 1
                                vQueItemCode = Trim(ds4.Tables(0).Rows(m)("itemcode").ToString)
                                vQueItemName = Trim(ds4.Tables(0).Rows(m)("itemname").ToString)
                                vQueUnit = Trim(ds4.Tables(0).Rows(m)("unitcode").ToString)
                                vQueQty = Trim(ds4.Tables(0).Rows(m)("qty").ToString)
                                vQueID = Trim(ds4.Tables(0).Rows(m)("queid").ToString)
                                vQueArName = Trim(ds4.Tables(0).Rows(m)("arname").ToString)
                                vQueSaleName = Trim(ds4.Tables(0).Rows(m)("salename").ToString)
                                vQueZoneID = Trim(ds4.Tables(0).Rows(m)("quezone").ToString)
                                vQueRefNo = Trim(ds4.Tables(0).Rows(m)("refno").ToString)
                                vQueDocNo = Trim(ds4.Tables(0).Rows(m)("docno").ToString)
                                vQueWHCode = Trim(ds4.Tables(0).Rows(m)("whcode").ToString)
                                vQueShelfCode = Trim(ds4.Tables(0).Rows(m)("shelfcode").ToString)
                                vQueShelfID = Trim(ds4.Tables(0).Rows(m)("shelfid").ToString)
                                vQueBarCode = Trim(ds4.Tables(0).Rows(m)("barcode").ToString)
                                vQuePickZone = Trim(ds4.Tables(0).Rows(m)("pickzone").ToString)

                                Dim listItem As New ListViewItem(vIndex)
                                listItem.SubItems.Add(vQueItemName)
                                listItem.SubItems.Add(Format(vQueQty, "##,##0.00"))
                                listItem.SubItems.Add(vQueUnit)
                                listItem.SubItems.Add(vQueID)
                                listItem.SubItems.Add(vQueZoneID)
                                listItem.SubItems.Add(vQueDocNo)
                                listItem.SubItems.Add(vQueItemCode)
                                listItem.SubItems.Add(vQueWHCode)
                                listItem.SubItems.Add(vQueShelfCode)
                                listItem.SubItems.Add(vQueBarCode)
                                listItem.SubItems.Add(vQuePickZone)
                                listItem.SubItems.Add(vQueShelfID)
                                Me.ListViewItemLastSend.Items.Add(listItem)
                            Next

                            Me.ListViewItemLastSend.Focus()
                            Me.ListViewItemLastSend.Items(0).Selected = True
                            Me.ListViewItemLastSend.Items(0).Focused = True
                        Else
                            Call SendCheckQue(vDocNo, vDocDate, vConnectZone)
                            Call ClearSaveData()
                        End If
                    Else
                        Call ClearSaveData()
                    End If

                End If

            End If
        Else
            MsgBox("�������¡���Թ������ѹ�֡������ DriveThru ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBRefNo.Focus()
            Me.TBRefNo.SelectAll()
        End If

        Call AfterSaveData()

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub ClearSaveData()
        On Error Resume Next

        Me.TBBarCode.Enabled = True
        Me.ListViewItem.Enabled = True
        Me.ListViewItem.Items.Clear()
        Me.TBSaleCode.Text = vMemSaleName
        Me.TBARCode.Text = ""
        Me.TBRefNo.Text = ""
        Me.TBItemAmount.Text = ""
        Me.TBDocNo.Text = ""
        Me.TBBarCode.Text = ""
        Me.TBARCode.Text = "99999"
        Me.TBRefNo.Enabled = True
        Me.TBRefNo.Focus()
        vIsOpen = 0
        vIsCancal = 0
        vIsconfirm = 0
        vIsSendQue = 0
    End Sub

    Private Sub SendCheckQue(ByVal vDocNo As String, ByVal vDocDate As String, ByVal vPickZone As String)
        Dim vSendCountID As Integer
        Dim vLastCountID As Integer
        Dim vType As Integer
        Dim i As Integer
        Dim vGroupZone(4) As String
        Dim n As Integer
        Dim vPrinterName As String

        On Error GoTo ErrDescription

        If Me.ListViewItem.Items.Count > 0 Then
            vType = 3
            vQuery = "exec dbo.USP_NP_CheckQuePickCenter1 '" & vDocNo & "','" & vDocDate & "' "
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            If ds.Tables(0).Rows.Count > 0 Then
                vLastCountID = Trim(ds.Tables(0).Rows(0)("max1").ToString)
            End If

            vSendCountID = vLastCountID + 1

            vQuery = "exec dbo.USP_NP_SearchGroupPicking1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'"
            Dim vService1 As New WebReference.WebServiceCalc
            Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)

            If ds1.Tables(0).Rows.Count > 0 Then
                n = ds1.Tables(0).Rows.Count
                For i = 0 To ds1.Tables(0).Rows.Count - 1
                    vGroupZone(i) = Trim(ds1.Tables(0).Rows(i)("zoneid").ToString)
                Next
            End If

            For i = 0 To n - 1
                If vGroupZone(i) = "A" Then
                    Call InsertQueZoneA(vDocNo, vDocDate, vSendCountID, vType)
                    vPrinterName = "�͡��þ�����͡���ش A"
                End If

                If vGroupZone(i) = "B" Then
                    Call InsertQueZoneB(vDocNo, vDocDate, vSendCountID, vType)
                    vPrinterName = "�͡��þ�����͡���ش B"
                End If

                If vGroupZone(i) = "C" Then
                    Call InsertQueZoneC(vDocNo, vDocDate, vSendCountID, vType)
                    vPrinterName = "�͡��þ�����͡���ش C"
                End If

                If vGroupZone(i) = "D" Then
                    Call InsertQueZoneD(vDocNo, vDocDate, vSendCountID, vType)
                    vPrinterName = "�͡��þ�����͡���ش D"
                End If
            Next

            vQuery = "exec dbo.USP_NP_UpdateSendQueuePicking1 " & vType & ",'" & vDocNo & "'"
            Dim vService2 As New WebReference.WebServiceCalc
            Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

            vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto1 3,'" & vDocNo & "','" & vPickZone & "','" & vUserName & "'"
            Dim vService3 As New WebReference.WebServiceCalc
            Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

            MsgBox("����¡���Թ���价ӡ�� CheckOut ���º��������  " & vPrinterName & " ", MsgBoxStyle.Information, "Send Information Message")
            Me.TBRefNo.Focus()
            Me.TBRefNo.SelectAll()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub InsertQueZoneA(ByVal vDocNo As String, ByVal vDocDate As String, ByVal vTimeID As Integer, ByVal vType As Integer)
        Dim vQueID As Integer
        Dim vQueDocDate As String
        Dim vARCode As String
        Dim vSaleCode As String
        Dim vRefNo As String
        Dim vMemberID As String
        Dim vSourceID As Integer
        Dim vQueZone As String
        Dim vQueReqTime As String
        Dim vIsConditionSend As Integer
        Dim vPickZone As String

        Dim vItemCode As String
        Dim vItemName As String
        Dim vQTY As Double
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vBarCode As String
        Dim vLineNumber As Integer

        Dim vInstrSale As Integer
        Dim vLenSale As Integer
        Dim vAddTime1 As Date
        Dim vAddTime As String

        Dim i As Integer

        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_NP_SearchNewDocNo 31"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vQueID = Trim(ds.Tables(0).Rows(0)("autonumber").ToString)
        End If

        vPickZone = "01"
        vQueZone = "A"
        vQueDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        vARCode = Me.TBARCode.Text
        vInstrSale = InStr(Me.TBSaleCode.Text, "/")
        vLenSale = Len(Me.TBSaleCode.Text)
        vSaleCode = vb6.Left(Me.TBSaleCode.Text, vInstrSale - 1)
        vRefNo = Me.TBRefNo.Text
        vMemberID = Me.TBMemberID.Text
        vSourceID = vType
        vAddTime1 = vb6.DateAdd(DateInterval.Minute, 15, Now)
        vAddTime = vAddTime1.Hour & ":" & vAddTime1.Minute
        vQueReqTime = vAddTime
        vIsConditionSend = 0

        vQuery = "exec dbo.USP_NP_InsertQuePickCenterMasterDriveIn1 '" & vQueID & "','" & vQueDocDate & "','" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "','" & vSourceID & "','" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "'"
        Dim vService1 As New WebReference.WebServiceCalc
        Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

        vQuery = "exec dbo.USP_NP_SearchReqPickingItemZone1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'," & vTimeID & ""
        Dim vService2 As New WebReference.WebServiceCalc
        Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)

        If ds2.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds2.Tables(0).Rows.Count - 1
                vItemCode = Trim(ds2.Tables(0).Rows(i)("itemcode").ToString)
                vItemName = Trim(ds2.Tables(0).Rows(i)("itemname").ToString)
                vQTY = Trim(ds2.Tables(0).Rows(i)("qty").ToString)
                vUnitCode = Trim(ds2.Tables(0).Rows(i)("unitcode").ToString)
                vWHCode = Trim(ds2.Tables(0).Rows(i)("whcode").ToString)
                vShelfCode = Trim(ds2.Tables(0).Rows(i)("shelfcode").ToString)
                vShelfID = Trim(ds2.Tables(0).Rows(i)("shelfid").ToString)
                vBarCode = Trim(ds2.Tables(0).Rows(i)("barcode").ToString)
                vLineNumber = i

                vQuery = "exec dbo.USP_NP_InsertQuePickCenterDriveInSub1 '" & vQueID & "','" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vQueZone & "','" & vPickZone & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & ""
                Dim vService3 As New WebReference.WebServiceCalc
                Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)
            Next
        End If

        vQuery = "exec dbo.USP_NP_UpdateNewDocNo 31"
        Dim vService4 As New WebReference.WebServiceCalc
        Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

    End Sub

    Private Sub InsertQueZoneB(ByVal vDocNo As String, ByVal vDocDate As String, ByVal vTimeID As Integer, ByVal vType As Integer)
        Dim vQueID As Integer
        Dim vQueDocDate As String
        Dim vARCode As String
        Dim vSaleCode As String
        Dim vRefNo As String
        Dim vMemberID As String
        Dim vSourceID As Integer
        Dim vQueZone As String
        Dim vQueReqTime As String
        Dim vIsConditionSend As Integer
        Dim vPickZone As String

        Dim vItemCode As String
        Dim vItemName As String
        Dim vQTY As Double
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vBarCode As String
        Dim vLineNumber As Integer

        Dim vInstrSale As Integer
        Dim vLenSale As Integer
        Dim vAddTime As String
        Dim vAddTime1 As Date

        Dim i As Integer

        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_NP_SearchNewDocNo 31"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vQueID = Trim(ds.Tables(0).Rows(0)("autonumber").ToString)
        End If

        vPickZone = "02"
        vQueZone = "B"
        vQueDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        vARCode = Me.TBARCode.Text
        vInstrSale = InStr(Me.TBSaleCode.Text, "/")
        vLenSale = Len(Me.TBSaleCode.Text)
        vSaleCode = vb6.Left(Me.TBSaleCode.Text, vInstrSale - 1)
        vRefNo = Me.TBRefNo.Text
        vMemberID = Me.TBMemberID.Text
        vSourceID = vType
        vAddTime1 = vb6.DateAdd(DateInterval.Minute, 15, Now)
        vAddTime = vAddTime1.Hour & ":" & vAddTime1.Minute
        vQueReqTime = vAddTime
        vIsConditionSend = 0

        vQuery = "exec dbo.USP_NP_InsertQuePickCenterMasterDriveIn1 '" & vQueID & "','" & vQueDocDate & "','" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "','" & vSourceID & "','" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "'"
        Dim vService1 As New WebReference.WebServiceCalc
        Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

        vQuery = "exec dbo.USP_NP_SearchReqPickingItemZone1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'," & vTimeID & ""
        Dim vService2 As New WebReference.WebServiceCalc
        Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)

        If ds2.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds2.Tables(0).Rows.Count - 1
                vItemCode = Trim(ds2.Tables(0).Rows(i)("itemcode").ToString)
                vItemName = Trim(ds2.Tables(0).Rows(i)("itemname").ToString)
                vQTY = Trim(ds2.Tables(0).Rows(i)("qty").ToString)
                vUnitCode = Trim(ds2.Tables(0).Rows(i)("unitcode").ToString)
                vWHCode = Trim(ds2.Tables(0).Rows(i)("whcode").ToString)
                vShelfCode = Trim(ds2.Tables(0).Rows(i)("shelfcode").ToString)
                vShelfID = Trim(ds2.Tables(0).Rows(i)("shelfid").ToString)
                vBarCode = Trim(ds2.Tables(0).Rows(i)("barcode").ToString)
                vLineNumber = i

                vQuery = "exec dbo.USP_NP_InsertQuePickCenterDriveInSub1 '" & vQueID & "','" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vQueZone & "','" & vPickZone & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & ""
                Dim vService3 As New WebReference.WebServiceCalc
                Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)
            Next
        End If

        vQuery = "exec dbo.USP_NP_UpdateNewDocNo 31"
        Dim vService4 As New WebReference.WebServiceCalc
        Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub InsertQueZoneC(ByVal vDocNo As String, ByVal vDocdate As String, ByVal vTimeID As Integer, ByVal vType As Integer)
        Dim vQueID As Integer
        Dim vQueDocDate As String
        Dim vARCode As String
        Dim vSaleCode As String
        Dim vRefNo As String
        Dim vMemberID As String
        Dim vSourceID As Integer
        Dim vQueZone As String
        Dim vQueReqTime As String
        Dim vIsConditionSend As Integer
        Dim vPickZone As String

        Dim vItemCode As String
        Dim vItemName As String
        Dim vQTY As Double
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vBarCode As String
        Dim vLineNumber As Integer

        Dim vInstrSale As Integer
        Dim vLenSale As Integer
        Dim vAddTime As String
        Dim vAddTime1 As Date

        Dim i As Integer

        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_NP_SearchNewDocNo 31"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vQueID = Trim(ds.Tables(0).Rows(0)("autonumber").ToString)
        End If

        vPickZone = "03"
        vQueZone = "C"
        vQueDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        vARCode = Me.TBARCode.Text
        vInstrSale = InStr(Me.TBSaleCode.Text, "/")
        vLenSale = Len(Me.TBSaleCode.Text)
        vSaleCode = vb6.Left(Me.TBSaleCode.Text, vInstrSale - 1)
        vRefNo = Me.TBRefNo.Text
        vMemberID = Me.TBMemberID.Text
        vSourceID = vType
        vAddTime1 = vb6.DateAdd(DateInterval.Minute, 15, Now)
        vAddTime = vAddTime1.Hour & ":" & vAddTime1.Minute
        vQueReqTime = vAddTime
        vIsConditionSend = 0

        vQuery = "exec dbo.USP_NP_InsertQuePickCenterMasterDriveIn1 '" & vQueID & "','" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "','" & vSourceID & "','" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "'"
        Dim vService1 As New WebReference.WebServiceCalc
        Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

        vQuery = "exec dbo.USP_NP_SearchReqPickingItemZone1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'," & vTimeID & ""
        Dim vService2 As New WebReference.WebServiceCalc
        Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)

        If ds2.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds2.Tables(0).Rows.Count - 1
                vItemCode = Trim(ds2.Tables(0).Rows(i)("itemcode").ToString)
                vItemName = Trim(ds2.Tables(0).Rows(i)("itemname").ToString)
                vQTY = Trim(ds2.Tables(0).Rows(i)("qty").ToString)
                vUnitCode = Trim(ds2.Tables(0).Rows(i)("unitcode").ToString)
                vWHCode = Trim(ds2.Tables(0).Rows(i)("whcode").ToString)
                vShelfCode = Trim(ds2.Tables(0).Rows(i)("shelfcode").ToString)
                vShelfID = Trim(ds2.Tables(0).Rows(i)("shelfid").ToString)
                vBarCode = Trim(ds2.Tables(0).Rows(i)("barcode").ToString)
                vLineNumber = i

                vQuery = "exec dbo.USP_NP_InsertQuePickCenterDriveInSub1 '" & vQueID & "','" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vQueZone & "','" & vPickZone & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & ""
                Dim vService3 As New WebReference.WebServiceCalc
                Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)
            Next
        End If

        vQuery = "exec dbo.USP_NP_UpdateNewDocNo 31"
        Dim vService4 As New WebReference.WebServiceCalc
        Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)


ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub InsertQueZoneD(ByVal vDocNo As String, ByVal vDocdate As String, ByVal vTimeID As Integer, ByVal vType As Integer)
        Dim vQueID As Integer
        Dim vQueDocDate As String
        Dim vARCode As String
        Dim vSaleCode As String
        Dim vRefNo As String
        Dim vMemberID As String
        Dim vSourceID As Integer
        Dim vQueZone As String
        Dim vQueReqTime As String
        Dim vIsConditionSend As Integer
        Dim vPickZone As String

        Dim vItemCode As String
        Dim vItemName As String
        Dim vQTY As Double
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vBarCode As String
        Dim vLineNumber As Integer

        Dim vInstrSale As Integer
        Dim vLenSale As Integer
        Dim vAddTime As String
        Dim vAddTime1 As Date

        Dim i As Integer

        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_NP_SearchNewDocNo 31"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vQueID = Trim(ds.Tables(0).Rows(0)("autonumber").ToString)
        End If

        vPickZone = "04"
        vQueZone = "D"
        vQueDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        vARCode = Me.TBARCode.Text
        vInstrSale = InStr(Me.TBSaleCode.Text, "/")
        vLenSale = Len(Me.TBSaleCode.Text)
        vSaleCode = vb6.Left(Me.TBSaleCode.Text, vInstrSale - 1)
        vRefNo = Me.TBRefNo.Text
        vMemberID = Me.TBMemberID.Text
        vSourceID = vType
        vAddTime1 = vb6.DateAdd(DateInterval.Minute, 15, Now)

        vAddTime = vAddTime1.Hour & ":" & vAddTime1.Minute
        vQueReqTime = vAddTime
        vIsConditionSend = 0

        vQuery = "exec dbo.USP_NP_InsertQuePickCenterMasterDriveIn1 '" & vQueID & "','" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "','" & vSourceID & "','" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "'"
        Dim vService1 As New WebReference.WebServiceCalc
        Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

        vQuery = "exec dbo.USP_NP_SearchReqPickingItemZone1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'," & vTimeID & ""
        Dim vService2 As New WebReference.WebServiceCalc
        Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)

        If ds2.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds2.Tables(0).Rows.Count - 1
                vItemCode = Trim(ds2.Tables(0).Rows(i)("itemcode").ToString)
                vItemName = Trim(ds2.Tables(0).Rows(i)("itemname").ToString)
                vQTY = Trim(ds2.Tables(0).Rows(i)("qty").ToString)
                vUnitCode = Trim(ds2.Tables(0).Rows(i)("unitcode").ToString)
                vWHCode = Trim(ds2.Tables(0).Rows(i)("whcode").ToString)
                vShelfCode = Trim(ds2.Tables(0).Rows(i)("shelfcode").ToString)
                vShelfID = Trim(ds2.Tables(0).Rows(i)("shelfid").ToString)
                vBarCode = Trim(ds2.Tables(0).Rows(i)("barcode").ToString)
                vLineNumber = i

                vQuery = "exec dbo.USP_NP_InsertQuePickCenterDriveInSub1 '" & vQueID & "','" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vQueZone & "','" & vPickZone & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & ""
                Dim vService3 As New WebReference.WebServiceCalc
                Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)
            Next
        End If

        vQuery = "exec dbo.USP_NP_UpdateNewDocNo 31"
        Dim vService4 As New WebReference.WebServiceCalc
        Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub SavePickUp()
        Dim vCountItem As Integer
        Dim vHeader As String
        Dim vNumber As Integer
        Dim vDocNumber As String

        Dim vDocNo As String
        Dim vDocDate As String
        Dim vARCode As String
        Dim vSaleCode As String
        Dim vMemberID As String
        Dim vRefNo As String
        Dim vTotalNetAmount As Double
        Dim vBeforeTaxAmount As Double
        Dim vTaxAmount As Double

        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vDiscountWord As String
        Dim vDiscountAmount As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vUserID As String
        Dim i As Integer
        Dim vBarCode As String
        Dim vShelfID As String
        Dim vZoneID As String
        Dim vLinePickZone As String
        Dim vLineNumber As Integer

        Dim a As Integer
        Dim b As Integer
        Dim vCheckItemCode As String
        Dim vCheckUnitCode As String
        Dim vCheckBarCode As String
        Dim vCheckWHCode As String
        Dim vCheckShelfCode As String
        Dim vCheckZoneID As String
        Dim vCheckPickZone As String

        Dim vOldItem As String
        Dim vOldUnit As String
        Dim vOldBar As String
        Dim vOldWH As String
        Dim vOldShelf As String
        Dim vOldZone As String
        Dim vOldPick As String
        Dim vOld As Integer

        Dim vCountItemPickZone As Integer
        Dim vItemPickZone As String
        Dim vCount As Integer
        Dim vQueZone As String

        Dim vCheckIsConfirm As Integer
        Dim vCheckHoldBillNo As String

        On Error GoTo ErrDescription

        If Me.ListViewItem.Items.Count > 0 And Me.TBItemAmount.Text <> "" Then
            vUserID = Me.TBUserID.Text

            If Me.TBRefNo.Text = "" Then
                MsgBox("��ͧ��͡�Ţ�����ҧ�ԧ��ǡ�͹�ѹ�֡�����ŷء����", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBRefNo.Focus()
                Me.TBRefNo.SelectAll()
                Exit Sub
            End If

            If Me.TBSaleCode.Text = "" Then
                MsgBox("��سҡ�͡ ���ʾ�ѡ�ҹ��� ��͹�ѹ�֡������", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBSaleCode.Focus()
                Me.TBSaleCode.SelectAll()
                Exit Sub
            End If


            If Me.TBDocNo.Text = "" Then
                vQuery = "exec dbo.usp_np_searchnewdocno 29"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count > 0 Then
                    vHeader = Trim(ds.Tables(0).Rows(0)("header").ToString)
                    vNumber = ds.Tables(0).Rows(0)("autonumber").ToString
                    vDocNumber = Trim(ds.Tables(0).Rows(0)("docnumber").ToString)
                End If

                vDocNo = Trim(vDocNumber & vHeader & "-" & Format(vNumber, "0000"))
            Else
                vDocNo = Me.TBDocNo.Text
            End If

            If vDocNo <> "" Then
                vDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year

                vQuery = "exec dbo.USP_NP_CheckDocDate"
                Dim vService7 As New WebReference.WebServiceCalc
                Dim ds7 As DataSet = vService7.vGetQueryAnlyzer(vQuery)
                If ds7.Tables(0).Rows.Count > 0 Then
                    vDocDate = ds7.Tables(0).Rows(0)("vdocdate").ToString
                End If

                vRefNo = Me.TBRefNo.Text

                If Me.RDZone1.Checked = True Then
                    vConnectZone = "01"
                    vQueZone = "A"
                ElseIf Me.RDZone2.Checked = True Then
                    vConnectZone = "02"
                    vQueZone = "B"
                ElseIf Me.RDZone3.Checked = True Then
                    vConnectZone = "03"
                    vQueZone = "C"
                ElseIf Me.RDZone4.Checked = True Then
                    vConnectZone = "04"
                    vQueZone = "D"
                End If

                For vCount = 0 To Me.ListViewItem.Items.Count - 1
                    vItemPickZone = Me.ListViewItem.Items(vCount).SubItems(12).Text
                    If vConnectZone = vItemPickZone Then
                        vCountItemPickZone = vCountItemPickZone + 1
                    End If
                Next

                If vCountItemPickZone = 0 Then
                    If vCountItemZoneOld = 0 Then
                        Call ClearSaveData()
                        Exit Sub
                    End If
                End If

                Dim vInstrSale As Integer
                Dim vLenSale As Integer

                If Me.TBARCode.Text = "1" Then
                    vARCode = "99999"
                Else
                    vARCode = Me.TBARCode.Text
                End If

                vInstrSale = InStr(Me.TBSaleCode.Text, "/")
                If vInstrSale = 0 Then
                    MsgBox("��سҡ�͡���ʾ�ѡ�ҹ���١��ͧ������������к���� �������ʾ�ѡ�ҹ���ǡ� Enter �ա����", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBSaleCode.Focus()
                    Me.TBSaleCode.SelectAll()
                    Exit Sub
                End If
                vLenSale = Len(Me.TBSaleCode.Text)
                vSaleCode = vb6.Left(Me.TBSaleCode.Text, vInstrSale - 1)

                vMemberID = Me.TBMemberID.Text
                vTotalNetAmount = Me.TBItemAmount.Text
                vBeforeTaxAmount = (vTotalNetAmount * 100) / 107
                vTaxAmount = vTotalNetAmount - vBeforeTaxAmount

                If vIsOpen = 0 Then

                    Call BeforeSaveData()
                    vQuery = "exec dbo.usp_np_insertdriveinslip1 '" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vMemberID & "','" & vSaleCode & "','" & vRefNo & "'," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalNetAmount & ",'" & vUserID & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

                    For i = 0 To Me.ListViewItem.Items.Count - 1
                        vItemCode = Me.ListViewItem.Items(i).SubItems(4).Text
                        vItemName = Me.ListViewItem.Items(i).SubItems(1).Text
                        vWHCode = Me.ListViewItem.Items(i).SubItems(7).Text
                        vShelfCode = Me.ListViewItem.Items(i).SubItems(8).Text
                        vQTY = Me.ListViewItem.Items(i).SubItems(2).Text
                        vPrice = Me.ListViewItem.Items(i).SubItems(5).Text
                        vAmount = Me.ListViewItem.Items(i).SubItems(6).Text
                        vUnitCode = Me.ListViewItem.Items(i).SubItems(3).Text
                        vBarCode = Me.ListViewItem.Items(i).SubItems(9).Text
                        vShelfID = Me.ListViewItem.Items(i).SubItems(10).Text
                        vZoneID = Me.ListViewItem.Items(i).SubItems(11).Text
                        vLinePickZone = Me.ListViewItem.Items(i).SubItems(12).Text
                        vDiscountWord = ""
                        vDiscountAmount = 0
                        vLineNumber = i

                        vQuery = "exec dbo.usp_np_insertdriveinslipsub1 '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "','" & vLinePickZone & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vAmount & ",'" & vBarCode & "','" & vSaleCode & "'," & vLineNumber & " "
                        Dim vService2 As New WebReference.WebServiceCalc
                        Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                    Next

                    vQuery = "exec dbo.usp_np_updatenewdocno 29"
                    Dim vService3 As New WebReference.WebServiceCalc
                    Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

                    MsgBox("�ѹ�֡���������Ţ����͡��� " & vDocNo & " ", MsgBoxStyle.Information, "Send Information Message")

                    Dim vAnswer As Integer

                    vAnswer = MsgBox("�س��ͧ����� �͡���价ӡ�� Check Out �������", MsgBoxStyle.YesNo, "Send Information Message")
                    If vAnswer = 6 Then
                        Call SendCheckQue(vDocNo, vDocDate, vConnectZone)
                        Call ClearSaveData()

                    Else
                        Call ClearSaveData()
                    End If

                End If

                If vIsOpen = 1 Then
                    Call BeforeSaveData()
                    vQuery = "exec dbo.USP_NP_CheckQueHoldBill1 '" & vDocNo & "','" & vARCode & "'"
                    Dim vService As New WebReference.WebServiceCalc
                    Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                    If ds.Tables(0).Rows.Count > 0 Then
                        vCheckIsConfirm = ds.Tables(0).Rows(0)("isconfirm").ToString()
                        vCheckHoldBillNo = ds.Tables(0).Rows(0)("holdbillno").ToString()
                    End If

                    If vCheckIsConfirm = 1 And vCheckHoldBillNo <> "" Then
                        MsgBox("�Ţ����͡�����ӡ�� ��Ǩ�ͺ�Թ�����з��͡��þѡ������º�������� �������ö�觤������͡����� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
                        Call ClearSaveData()
                        Call AfterSaveData()
                        Me.TBRefNo.Focus()
                        Me.TBRefNo.SelectAll()
                        Exit Sub
                    End If

                    vQuery = "exec dbo.usp_np_insertdriveinslip1 '" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vMemberID & "','" & vSaleCode & "','" & vRefNo & "'," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalNetAmount & ",'" & vUserID & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

                    vCountItem = Me.ListViewItem.Items.Count

                    For a = 0 To vCountItemOld
                        vOldItem = vMemItemCodeOld(a)
                        vOldUnit = vMemUnitCodeOld(a)
                        vOldBar = vMemBarCodeOld(a)
                        vOldWH = vMemWHCodeOld(a)
                        vOldShelf = vMemShelfCodeOld(a)
                        vOldZone = vMemZoneIDOld(a)
                        vOldPick = vMemPickZoneOld(a)

                        For b = 0 To Me.ListViewItem.Items.Count - 1
                            vCheckItemCode = Me.ListViewItem.Items(b).SubItems(4).Text
                            vCheckUnitCode = Me.ListViewItem.Items(b).SubItems(3).Text
                            vCheckBarCode = Me.ListViewItem.Items(b).SubItems(9).Text
                            vCheckWHCode = Me.ListViewItem.Items(b).SubItems(7).Text
                            vCheckShelfCode = Me.ListViewItem.Items(b).SubItems(8).Text
                            vCheckZoneID = Me.ListViewItem.Items(b).SubItems(11).Text
                            vCheckPickZone = Me.ListViewItem.Items(b).SubItems(12).Text

                            If vCheckItemCode = vOldItem And vCheckUnitCode = vOldUnit And vCheckBarCode = vOldBar And vCheckWHCode = vOldWH And vCheckShelfCode = vOldShelf And vCheckZoneID = vOldZone And vCheckPickZone = vOldPick Then
                                vOld = 1
                                GoTo Line1
                            Else
                                vOld = 0
                            End If
                        Next
Line1:

                        If vOld = 0 Then
                            vItemCode = vOldItem
                            vWHCode = vOldWH
                            vShelfCode = vOldShelf
                            vUnitCode = vOldUnit
                            vBarCode = vOldBar
                            vZoneID = vOldZone
                            vLinePickZone = vOldPick

                            vQuery = "exec dbo.usp_np_deletedriveinslipsub1 '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vZoneID & "','" & vLinePickZone & "','" & vUnitCode & "','" & vBarCode & "'," & vTotalNetAmount & " "
                            Dim vService2 As New WebReference.WebServiceCalc
                            Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                        End If
                    Next

                    For i = 0 To Me.ListViewItem.Items.Count - 1
                        vItemCode = Me.ListViewItem.Items(i).SubItems(4).Text
                        vItemName = Me.ListViewItem.Items(i).SubItems(1).Text
                        vWHCode = Me.ListViewItem.Items(i).SubItems(7).Text
                        vShelfCode = Me.ListViewItem.Items(i).SubItems(8).Text
                        vQTY = Me.ListViewItem.Items(i).SubItems(2).Text
                        vPrice = Me.ListViewItem.Items(i).SubItems(5).Text
                        vAmount = Me.ListViewItem.Items(i).SubItems(6).Text
                        vUnitCode = Me.ListViewItem.Items(i).SubItems(3).Text
                        vBarCode = Me.ListViewItem.Items(i).SubItems(9).Text
                        vShelfID = Me.ListViewItem.Items(i).SubItems(10).Text
                        vZoneID = Me.ListViewItem.Items(i).SubItems(11).Text
                        vLinePickZone = Me.ListViewItem.Items(i).SubItems(12).Text
                        vDiscountWord = ""
                        vDiscountAmount = 0
                        vLineNumber = i

                        If vConnectZone = vLinePickZone Then
                            vQuery = "exec dbo.usp_np_insertdriveinslipsub1 '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "','" & vLinePickZone & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vAmount & ",'" & vBarCode & "','" & vSaleCode & "'," & vLineNumber & " "
                            Dim vService3 As New WebReference.WebServiceCalc
                            Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)
                        End If
                    Next
                    MsgBox("����Ţ����͡��� " & vDocNo & " ���º��������", MsgBoxStyle.Information, "Send Information Message")

                    Me.TBRefNo.Enabled = True
                    Me.TBARCode.Enabled = True
                    Me.TBSaleCode.Enabled = True

                    Dim vAnswer As Integer

                    vAnswer = MsgBox("�س��ͧ����� �͡���价ӡ�� Check Out �������", MsgBoxStyle.YesNo, "Send Information Message")
                    If vAnswer = 6 Then

                        Dim m As Integer
                        Dim vQueItemCode As String
                        Dim vQueItemName As String
                        Dim vQueUnit As String
                        Dim vQueQty As Double
                        Dim vQueID As Integer
                        Dim vQueArName As String
                        Dim vQueSaleName As String
                        Dim vQueZoneID As String
                        Dim vQueRefNo As String
                        Dim vIndex As Integer
                        Dim vQueDocNo As String
                        Dim vQueWHCode As String
                        Dim vQueShelfCode As String
                        Dim vQueShelfID As String
                        Dim vQueBarCode As String
                        Dim vQuePickZone As String


                        vQuery = "exec dbo.USP_NP_CheckQueDriveIn1 '" & vDocNo & "','" & vDocDate & "','" & vQueZone & "' "
                        Dim vService4 As New WebReference.WebServiceCalc
                        Dim ds4 As DataSet = vService4.vGetQueryAnlyzer(vQuery)

                        If ds4.Tables(0).Rows.Count > 0 Then
                            Me.PNLastQueSend.Visible = True
                            Me.TBCarLicense.Text = Trim(ds4.Tables(0).Rows(0)("refno").ToString)
                            Me.TBQueAR.Text = Trim(ds4.Tables(0).Rows(0)("arcode").ToString) & "/" & Trim(ds4.Tables(0).Rows(0)("arname").ToString)

                            Me.ListViewItemLastSend.Items.Clear()
                            For m = 0 To ds4.Tables(0).Rows.Count - 1
                                vIndex = vIndex + 1
                                vQueItemCode = Trim(ds4.Tables(0).Rows(m)("itemcode").ToString)
                                vQueItemName = Trim(ds4.Tables(0).Rows(m)("itemname").ToString)
                                vQueUnit = Trim(ds4.Tables(0).Rows(m)("unitcode").ToString)
                                vQueQty = Trim(ds4.Tables(0).Rows(m)("qty").ToString)
                                vQueID = Trim(ds4.Tables(0).Rows(m)("queid").ToString)
                                vQueArName = Trim(ds4.Tables(0).Rows(m)("arname").ToString)
                                vQueSaleName = Trim(ds4.Tables(0).Rows(m)("salename").ToString)
                                vQueZoneID = Trim(ds4.Tables(0).Rows(m)("quezone").ToString)
                                vQueRefNo = Trim(ds4.Tables(0).Rows(m)("refno").ToString)
                                vQueDocNo = Trim(ds4.Tables(0).Rows(m)("docno").ToString)
                                vQueWHCode = Trim(ds4.Tables(0).Rows(m)("whcode").ToString)
                                vQueShelfCode = Trim(ds4.Tables(0).Rows(m)("shelfcode").ToString)
                                vQueShelfID = Trim(ds4.Tables(0).Rows(m)("shelfid").ToString)
                                vQueBarCode = Trim(ds4.Tables(0).Rows(m)("barcode").ToString)
                                vQuePickZone = Trim(ds4.Tables(0).Rows(m)("pickzone").ToString)

                                Dim listItem As New ListViewItem(vIndex)
                                listItem.SubItems.Add(vQueItemName)
                                listItem.SubItems.Add(Format(vQueQty, "##,##0.00"))
                                listItem.SubItems.Add(vQueUnit)
                                listItem.SubItems.Add(vQueID)
                                listItem.SubItems.Add(vQueZoneID)
                                listItem.SubItems.Add(vQueDocNo)
                                listItem.SubItems.Add(vQueItemCode)
                                listItem.SubItems.Add(vQueWHCode)
                                listItem.SubItems.Add(vQueShelfCode)
                                listItem.SubItems.Add(vQueBarCode)
                                listItem.SubItems.Add(vQuePickZone)
                                listItem.SubItems.Add(vQueShelfID)
                                Me.ListViewItemLastSend.Items.Add(listItem)
                            Next

                            Me.ListViewItemLastSend.Focus()
                            Me.ListViewItemLastSend.Items(0).Selected = True
                            Me.ListViewItemLastSend.Items(0).Focused = True
                        Else
                            Call SendCheckQue(vDocNo, vDocDate, vConnectZone)
                            Call ClearSaveData()
                        End If
                    Else
                        Call ClearSaveData()
                    End If

                End If

            End If
        Else
            MsgBox("�������¡���Թ������ѹ�֡������ DriveThru ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
        End If

        Call AfterSaveData()


ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub
    Private Sub frmProgram1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        Me.PNDriveIn.Visible = False

        Me.PNLogIn.Visible = True
        Me.PNLogIn.BringToFront()
        Me.RDZone1.Focus()
    End Sub

    Private Sub CallIDNumber()
        Me.TBARCode.Text = "99999"
    End Sub

    Private Sub TBRefNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBRefNo.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Dim vCarLicense As String
            Dim vCountDocNo As Integer
            Dim vDocno As String

            If Me.TBRefNo.Text <> "" Then
                vCarLicense = Me.TBRefNo.Text

                vQuery = "exec dbo.USP_NP_SearchCarLicenseDriveIn1 '" & vCarLicense & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                vCountDocNo = ds.Tables(0).Rows.Count

                If ds.Tables(0).Rows.Count = 1 Then
                    vDocno = ds.Tables(0).Rows(0)("docno").ToString()
                End If

                If vCountDocNo = 1 Then
                    Dim i As Integer
                    Dim vNetItemAmount As Double

                    Dim vItemCode As String
                    Dim vItemName As String
                    Dim vWHCode As String
                    Dim vShelfCode As String
                    Dim vQTY As Double
                    Dim vPrice As Double
                    Dim vAmount As Double
                    Dim vUnitCode As String
                    Dim vPickZone As String
                    Dim vBarCode As String
                    Dim vShelfID As String
                    Dim vZoneID As String
                    Dim vIndex As Integer
                    Dim vPointZone As String

                    If Me.RDZone1.Checked = True Then
                        vPointZone = "01"
                    End If

                    If Me.RDZone2.Checked = True Then
                        vPointZone = "02"
                    End If

                    If Me.RDZone3.Checked = True Then
                        vPointZone = "03"
                    End If

                    If Me.RDZone4.Checked = True Then
                        vPointZone = "04"
                    End If

                    vQuery = "exec dbo.usp_np_SearchDriveInDetails1 '" & vDocno & "','" & vPointZone & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)

                    Me.ListViewItem.Items.Clear()
                    If ds1.Tables(0).Rows.Count > 0 Then
                        vCountItemZoneOld = 0
                        vIsOpen = 1
                        vIsCancal = ds1.Tables(0).Rows(i)("iscancel").ToString
                        vIsconfirm = ds1.Tables(0).Rows(i)("isconfirm").ToString
                        vIsSendQue = ds1.Tables(0).Rows(i)("issendque").ToString

                        Me.TBARCode.Text = ds1.Tables(0).Rows(i)("arcode").ToString
                        Me.TBARName.Text = ds1.Tables(0).Rows(i)("arname").ToString
                        Me.TBRefNo.Text = ds1.Tables(0).Rows(i)("refno").ToString
                        vNetItemAmount = ds1.Tables(0).Rows(i)("totalnetamount").ToString
                        Me.TBItemAmount.Text = Format(vNetItemAmount, "##,##0.00")
                        Me.TBDocNo.Text = ds1.Tables(0).Rows(i)("docno").ToString
                        If ds1.Tables(0).Rows(i)("salecode").ToString <> "" Then
                            Me.TBSaleCode.Text = ds1.Tables(0).Rows(i)("salecode").ToString
                        Else
                            Me.TBSaleCode.Text = vMemSaleName
                        End If
                        vIndex = 0
                        vCountItemOld = ds1.Tables(0).Rows.Count - 1

                        ReDim vMemItemCodeOld(vCountItemOld)
                        ReDim vMemUnitCodeOld(vCountItemOld)
                        ReDim vMemWHCodeOld(vCountItemOld)
                        ReDim vMemShelfCodeOld(vCountItemOld)
                        ReDim vMemZoneIDOld(vCountItemOld)
                        ReDim vMemBarCodeOld(vCountItemOld)
                        ReDim vMemPickZoneOld(vCountItemOld)

                        For i = 0 To ds1.Tables(0).Rows.Count - 1
                            vMemItemCodeOld(i) = ds1.Tables(0).Rows(i)("itemcode").ToString
                            vMemUnitCodeOld(i) = ds1.Tables(0).Rows(i)("unitcode").ToString
                            vMemWHCodeOld(i) = ds1.Tables(0).Rows(i)("whcode").ToString
                            vMemShelfCodeOld(i) = ds1.Tables(0).Rows(i)("shelfcode").ToString
                            vMemZoneIDOld(i) = ds1.Tables(0).Rows(i)("zoneid").ToString
                            vMemBarCodeOld(i) = ds1.Tables(0).Rows(i)("barcode").ToString
                            vMemPickZoneOld(i) = ds1.Tables(0).Rows(i)("pickzone").ToString

                            If vPointZone = vMemPickZoneOld(i) Then
                                vCountItemZoneOld = vCountItemZoneOld + 1
                            End If

                            vPickZone = ds1.Tables(0).Rows(i)("pickzone").ToString
                            vItemCode = ds1.Tables(0).Rows(i)("itemcode").ToString
                            vItemName = ds1.Tables(0).Rows(i)("itemname").ToString
                            vWHCode = ds1.Tables(0).Rows(i)("whcode").ToString
                            vShelfCode = ds1.Tables(0).Rows(i)("shelfcode").ToString
                            vQTY = ds1.Tables(0).Rows(i)("qty").ToString
                            vUnitCode = ds1.Tables(0).Rows(i)("unitcode").ToString
                            vPrice = ds1.Tables(0).Rows(i)("price").ToString
                            vAmount = ds1.Tables(0).Rows(i)("amount").ToString
                            vBarCode = ds1.Tables(0).Rows(i)("barcode").ToString
                            vShelfID = ds1.Tables(0).Rows(i)("shelfid").ToString
                            vZoneID = ds1.Tables(0).Rows(i)("zoneid").ToString

                            vIndex = vIndex + 1
                            Dim listItem As New ListViewItem(vIndex)
                            listItem.SubItems.Add(vItemName)
                            listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                            listItem.SubItems.Add(vUnitCode)
                            listItem.SubItems.Add(vItemCode)
                            listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                            listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                            listItem.SubItems.Add(vWHCode)
                            listItem.SubItems.Add(vShelfCode)
                            listItem.SubItems.Add(vBarCode)
                            listItem.SubItems.Add(vShelfID)
                            listItem.SubItems.Add(vZoneID)
                            listItem.SubItems.Add(vPickZone)
                            Me.ListViewItem.Items.Add(listItem)

                            If vPickZone = vPointZone Then
                                Me.ListViewItem.Items.Item(i).BackColor = Color.White
                            End If
                        Next

                        Me.TBRefNo.Enabled = False

                    End If

                    Me.PNDriveIn.Visible = True
                    Me.TBRefNo.Enabled = False
                    Me.PNDriveIn.BringToFront()
                    Me.TBBarCode.Focus()
                    Me.TBBarCode.SelectAll()

                ElseIf vCountDocNo > 1 Then
                    Dim vSearch As String
                    Dim i As Integer
                    Dim vDocDate As String
                    Dim vRefID As String
                    Dim vAmount As Double
                    Dim vIndex As Integer

                    Me.PNLogIn.Visible = False
                    Me.PNDriveIn.Visible = False
                    Me.PNSearchPickUp.Visible = True
                    Me.PNSearchPickUp.BringToFront()
                    Me.TBSearchPickup.Text = ""

                    vSearch = Me.TBRefNo.Text
                    vQuery = "exec dbo.usp_np_SearchDriveInMaster1 '" & vSearch & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)

                    Me.ListViewSearhPickup.Items.Clear()
                    vIndex = 0
                    If ds1.Tables(0).Rows.Count > 0 Then
                        For i = 0 To ds1.Tables(0).Rows.Count - 1
                            vDocno = ds1.Tables(0).Rows(i)("docno").ToString
                            vDocDate = ds1.Tables(0).Rows(i)("docdate").ToString
                            vRefID = ds1.Tables(0).Rows(i)("refid").ToString
                            vAmount = ds1.Tables(0).Rows(i)("totalnetamount").ToString

                            vIndex = vIndex + 1
                            Dim listItem As New ListViewItem(vIndex)
                            listItem.SubItems.Add(vRefID)
                            listItem.SubItems.Add(vDocno)
                            listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                            Me.ListViewSearhPickup.Items.Add(listItem)

                        Next

                        Dim a As Integer

                        For a = 0 To Me.ListViewItem.Items.Count - 1
                            If a Mod 2 <> 0 Then
                                Me.ListViewItem.Items(a).BackColor = Color.Silver
                            End If
                        Next

                        Me.ListViewSearhPickup.Focus()
                        Me.ListViewSearhPickup.Items(0).Focused = True
                        Me.ListViewSearhPickup.Items(0).Selected = True

                    End If
                Else
                    Me.TBARCode.Focus()
                    Me.TBARCode.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBARCode.Focus()
            Me.TBARCode.SelectAll()
        End If

        If e.KeyCode = Keys.Right Then
            Me.TBARCode.Focus()
            Me.TBARCode.SelectAll()
        End If


ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBQTY_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBQTY.KeyDown
        Dim vBarCode As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vIndex As Integer
        Dim vItemLine As Integer
        Dim vCheckExist As Integer

        Dim vCheckShelf As String
        Dim vCheckUnit As String

        Dim v As Integer
        Dim vShelfQTY As Double
        Dim vShelfUnit As String
        Dim vListShelf As String
        Dim vListUnit As String
        Dim vRate As Integer
        Dim vTotalQTY As Double
        Dim vShelfID As String
        Dim vZoneID As String
        Dim vPickZone As String

        Dim vAnswer As Integer
        Dim vCheckPrice As Double

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Me.TBItem.Text = ""
            Me.TBItemName.Text = ""
            Me.TBPrice.Text = ""
            Me.TBUnit.Text = ""
            Me.TBWHCode.Text = ""
            Me.TBShelfCode.Text = ""
            Me.TBQTY.Text = ""
            Me.TBRate.Text = ""
            Me.TBReserve.Text = ""
            Me.TBMemBarCode.Text = ""
            Me.TBShelfID.Text = ""
            Me.TBZoneID.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNItemDetails.Visible = False
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBBarCode.Focus()
        End If

        If Me.TBRefNo.Text = "" Then
            MsgBox("��ͧ��͡����¹ö�����Ţ�����ҧ�ԧ��͹ �֧�����͡��Ժ�Թ�����", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBBarCode.Text = ""
            Me.TBRefNo.Focus()
            Exit Sub
        End If

        If e.KeyCode = Keys.Enter Then

            If Me.TBPrice.Text <> "" Then
                vCheckPrice = Me.TBPrice.Text
            End If
            If vCheckPrice = 0 Then
                MsgBox("�Թ��Ҵѧ������ѧ������˹��ҤҢ�� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
                Exit Sub
            End If

            If Me.RDZone1.Checked = True Then
                vPickZone = "01"
            End If

            If Me.RDZone2.Checked = True Then
                vPickZone = "02"
            End If

            If Me.RDZone3.Checked = True Then
                vPickZone = "03"
            End If

            If Me.RDZone4.Checked = True Then
                vPickZone = "04"
            End If

            If Me.ListViewStock.Items.Count > 0 And Me.TBItem.Text <> "" Then
                vCheckShelf = Me.TBShelfCode.Text
                vCheckUnit = Me.TBUnit.Text

                If Me.ListViewStock.Items.Count > 0 Then
                    For v = 0 To Me.ListViewStock.Items.Count - 1
                        vListShelf = Me.ListViewStock.Items(v).SubItems(1).Text 'Me.ListViewStock.Items(v).Text
                        vListUnit = Me.ListViewStock.Items(v).SubItems(3).Text
                        If vCheckShelf = vListShelf And vCheckUnit = vListUnit Then
                            vShelfQTY = Me.ListViewStock.Items(v).SubItems(2).Text
                            vShelfUnit = Me.ListViewStock.Items(v).SubItems(3).Text
                            GoTo Line1
                        End If
                    Next
                End If

Line1:
                vCheckExist = 0
                vBarCode = Me.TBMemBarCode.Text
                vItemCode = Me.TBItem.Text
                vItemName = Me.TBItemName.Text
                vWHCode = Me.TBWHCode.Text
                vShelfCode = Me.TBShelfCode.Text
                vUnitCode = Me.TBUnit.Text
                vRate = Me.TBRate.Text
                vShelfID = Me.TBShelfID.Text
                vZoneID = Me.TBZoneID.Text



                If Me.TBQTY.Text <> "" Then
                    vQTY = Me.TBQTY.Text
                End If

                If vShelfUnit <> vUnitCode Then
                    vTotalQTY = vShelfQTY / vRate
                    If vQTY > vTotalQTY Then
                        vAnswer = MsgBox("�Թ������� " & vItemCode & " STOCK ���͢�� ��ͧ��â���Թ��ҹ�� ���������", MsgBoxStyle.YesNo, "Send Question Message ")
                        If vAnswer = 7 Then
                            Me.TBQTY.SelectAll()
                            Exit Sub
                        End If
                    End If
                End If

                If vQTY > vShelfQTY And vShelfUnit = vUnitCode Then
                    vAnswer = MsgBox("�Թ������� " & vItemCode & " STOCK ���͢�� ��ͧ��â���Թ��ҹ�� ���������", MsgBoxStyle.YesNo, "Send Question Message ")
                    If vAnswer = 7 Then
                        Me.TBQTY.SelectAll()
                        Exit Sub
                    End If
                End If

                If Me.TBPrice.Text <> "" Then
                    vPrice = Me.TBPrice.Text
                End If
                vAmount = vQTY * vPrice

                vIndex = Me.ListViewItem.Items.Count + 1
                vItemLine = Me.ListViewItem.Items.Count

                If vQTY = 0 Then
                    MsgBox("������кبӹǹ�ͧ�Թ��ҷ���ͧ��� ���͵�ͧ�кبӹǹ�Թ��ҷ���ͧ����ҡ���� 0", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBQTY.Focus()
                    Me.TBQTY.SelectAll()
                    Exit Sub
                End If

                Dim n As Integer
                Dim vCheckItemCode As String
                Dim vCheckWHCode As String
                Dim vCheckPickZone As String
                Dim vCheckZoneID As String
                Dim vCheckUnitCode As String
                Dim vCheckShelfCode As String

                Dim vEditQTY As Double
                Dim vEditPrice As Double
                Dim vItemAmount As Double


                If Me.ListViewItem.Items.Count > 0 Then
                    For n = 0 To Me.ListViewItem.Items.Count - 1
                        vCheckItemCode = Me.ListViewItem.Items(n).SubItems(4).Text
                        vCheckUnitCode = Me.ListViewItem.Items(n).SubItems(3).Text
                        vCheckShelfCode = Me.ListViewItem.Items(n).SubItems(8).Text
                        vCheckPickZone = Me.ListViewItem.Items(n).SubItems(12).Text
                        vCheckWHCode = Me.ListViewItem.Items(n).SubItems(7).Text
                        vCheckZoneID = Me.ListViewItem.Items(n).SubItems(11).Text

                        If vItemCode = vCheckItemCode And vUnitCode = vCheckUnitCode And vWHCode = vCheckWHCode And vShelfCode = vCheckShelfCode And vZoneID = vCheckZoneID And vPickZone = vCheckPickZone Then
                            vEditPrice = Me.TBPrice.Text
                            vEditQTY = Me.TBQTY.Text
                            vItemAmount = vEditQTY * vEditPrice
                            Me.ListViewItem.Items(n).SubItems(2).Text = Format(vEditQTY, "##,##0.00")
                            Me.ListViewItem.Items(n).SubItems(6).Text = Format(vItemAmount, "##,##0.00")
                            vCheckExist = 1
                            GoTo line2
                        End If
                    Next
                End If

line2:

                If vCheckExist = 0 Then
                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(vItemName)
                    listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                    listItem.SubItems.Add(vUnitCode)
                    listItem.SubItems.Add(vItemCode)
                    listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                    listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                    listItem.SubItems.Add(vWHCode)
                    listItem.SubItems.Add(vShelfCode)
                    listItem.SubItems.Add(vBarCode)
                    listItem.SubItems.Add(vShelfID)
                    listItem.SubItems.Add(vZoneID)
                    listItem.SubItems.Add(vPickZone)
                    Me.ListViewItem.Items.Add(listItem)

                    Me.ListViewItem.Items.Item(vItemLine).BackColor = Color.White

                End If

                Call CalcItemAmount()

                If vQTY >= 10000 Then
                    MsgBox("�ӹǹ�Թ��ҷ�����͡� �Թ 10,000 ��سҵ�Ǩ�ͺ�������ա��", MsgBoxStyle.Information, "Send Error Message")
                End If

                Me.TBItem.Text = ""
                Me.TBMemBarCode.Text = ""
                Me.TBItemName.Text = ""
                Me.TBPrice.Text = ""
                Me.TBUnit.Text = ""
                Me.TBWHCode.Text = ""
                Me.TBShelfCode.Text = ""
                Me.TBShelfID.Text = ""
                Me.TBZoneID.Text = ""
                Me.TBQTY.Text = ""
                Me.TBReserve.Text = ""
                Me.ListViewStock.Items.Clear()
                Me.PNItemDetails.Visible = False
                Me.BTNSave.Visible = True
                Me.TBBarCode.Text = ""
                Me.TBBarCode.Focus()
            Else
                MsgBox("�������¡���Թ����������ö���� ��¡���Թ���ŧ�С�����", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBBarCode.Focus()
                Me.TBBarCode.SelectAll()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNLogIN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vUserCode As String
        Dim vPassWord As String
        Dim vCheckTypeLogIn As String

        On Error GoTo ErrDescription

        vUserCode = Me.TBUserCode.Text
        vPassWord = Me.TBPassword.Text

        If vCheckLogIn <> "" Then
            Me.PNLogIn.Visible = False
            Me.PNDriveIn.Visible = False
            Me.TBUserID.Text = vCheckLogIn
            Call CallIDNumber()

            If Me.RDZone1.Checked = True Then
                vConnectZone = "01"
                vCheckTypeLogIn = "�ش���·��1"
            ElseIf Me.RDZone2.Checked = True Then
                vConnectZone = "02"
                vCheckTypeLogIn = "�ش���·��2"
            ElseIf Me.RDZone3.Checked = True Then
                vConnectZone = "03"
                vCheckTypeLogIn = "�ش���·��3"
            ElseIf Me.RDZone4.Checked = True Then
                vConnectZone = "04"
                vCheckTypeLogIn = "�ش���·��4"
            End If

            If vCheckTypeLogIn <> "05-Checker" Then
                Me.PNLogIn.Visible = False
                Me.PNDriveIn.Visible = True
                Me.PNDriveIn.BringToFront()
                Me.TBRefNo.Focus()
            Else
                Me.PNLogIn.Visible = False
                Me.PNDriveIn.Visible = False
            End If

        Else
            MsgBox("�������ö�����ҹ������� ��سҵ�Ǩ�ͺ����������ʼ�ҹ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBPassword.Text = ""
            Me.TBPassword.Focus()
            Me.TBPassword.SelectAll()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBUserCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBUserCode.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Me.TBPassword.Focus()
            Me.TBPassword.SelectAll()
        End If
        If e.KeyCode = Keys.Up Then
            Me.RDZone3.Focus()
        End If
        If e.KeyCode = Keys.Down Then
            Me.TBPassword.Focus()
            Me.TBPassword.SelectAll()
        End If
        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("�س��ͧ����͡�ҡ��������������", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBPassword.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Up Then
            Me.TBUserCode.Focus()
            Me.TBUserCode.SelectAll()
        End If
        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("�س��ͧ����͡�ҡ��������������", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub MenuDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuDelete.Click
        Dim i As Integer
        Dim vLinePickZone As String
        Dim vPickZone As String

        On Error GoTo ErrDescription

        i = Me.ListViewItem.FocusedItem.Index

        vLinePickZone = Me.ListViewItem.Items(i).SubItems(12).Text

        If Me.RDZone1.Checked = True Then
            vPickZone = "01"
        End If

        If Me.RDZone2.Checked = True Then
            vPickZone = "02"
        End If

        If Me.RDZone3.Checked = True Then
            vPickZone = "03"
        End If

        If Me.RDZone4.Checked = True Then
            vPickZone = "04"
        End If

        If vPickZone <> vLinePickZone Then
            MsgBox("�Թ������褹��⫹ Drive Thru �������öź�������� Drive Thru � �ش�˹����öź � �ش�����ҹ��", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBBarCode.Focus()
            Exit Sub
        End If

        Me.ListViewItem.Items.RemoveAt(i)
        Call GenIDNumber()
        Call CalcItemAmount()
        Me.TBBarCode.Focus()

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub
    Private Sub GenIDNumber()
        Dim i As Integer
        Dim j As Integer

        On Error Resume Next

        If Me.ListViewItem.Items.Count > 0 Then
            j = 0
            For i = 0 To Me.ListViewItem.Items.Count - 1
                j = j + 1
                Me.ListViewItem.Items(i).SubItems(0).Text = j
            Next
        End If
    End Sub

    Private Sub CMBZone_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next

        Me.TBUserCode.Focus()
    End Sub


    Private Sub BTNCloseLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next

        If vCheckLogIn = "" Then
            Application.Exit()
        Else
            Me.PNLogIn.Visible = False
        End If
    End Sub

    Private Sub MenuSearchPickUp()
        On Error Resume Next

        Me.PNLogIn.Visible = False
        Me.PNDriveIn.Visible = False

        Me.PNSearchPickUp.Visible = True
        Me.PNSearchPickUp.BringToFront()
        Me.TBSearchPickup.Focus()
    End Sub

    Private Sub BTNClosePickup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClosePickup.Click
        Dim vAnswer As Integer

        On Error Resume Next

        MsgBox("���駵��� ��顴���� < ���͡�Ѻ��ѧ˹�����͡�ش DriveThru ��ѧ�ҡ�����顴 ���� ESC �����͡�����", MsgBoxStyle.Information, "Send Information Message")


        vAnswer = MsgBox("�س��ͧ����͡�������������� ?", MsgBoxStyle.YesNo, "Send Question Information")
        If vAnswer = 6 Then
            Application.Exit()
        Else
            Exit Sub
        End If

    End Sub

    Private Sub TBSearchPickup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchPickup.KeyDown
        Dim vSearch As String
        Dim i As Integer
        Dim vDocno As String
        Dim vDocDate As String
        Dim vRefID As String
        Dim vAmount As Double
        Dim vIndex As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            vSearch = Me.TBSearchPickup.Text

            vQuery = "exec dbo.usp_np_SearchDriveInMaster1 '" & vSearch & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            Me.ListViewSearhPickup.Items.Clear()
            vIndex = 0
            If ds.Tables(0).Rows.Count > 0 Then
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    vDocno = ds.Tables(0).Rows(i)("docno").ToString
                    vDocDate = ds.Tables(0).Rows(i)("docdate").ToString
                    vRefID = ds.Tables(0).Rows(i)("refid").ToString
                    vAmount = ds.Tables(0).Rows(i)("totalnetamount").ToString

                    vIndex = vIndex + 1
                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(vRefID)
                    listItem.SubItems.Add(vDocno)
                    listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                    Me.ListViewSearhPickup.Items.Add(listItem)

                Next

                Dim a As Integer

                For a = 0 To Me.ListViewItem.Items.Count - 1
                    If a Mod 2 <> 0 Then
                        Me.ListViewItem.Items(a).BackColor = Color.Silver
                    End If
                Next

                Me.ListViewSearhPickup.Focus()

                Me.ListViewSearhPickup.Items(0).Focused = True
                Me.ListViewSearhPickup.Items(0).Selected = True
            Else
                Me.TBSearchPickup.Focus()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.ListViewSearhPickup.Items.Clear()
            Me.TBSearchPickup.Text = ""
            Me.PNSearchPickUp.Visible = False
            Me.PNDriveIn.Visible = True
            Me.PNDriveIn.BringToFront()
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Down Then
            If Me.ListViewSearhPickup.Items.Count > 0 Then
                Me.ListViewSearhPickup.Focus()

                Me.ListViewSearhPickup.Items(0).Focused = True
                Me.ListViewSearhPickup.Items(0).Selected = True
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSearchPickup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vSearch As String
        Dim i As Integer
        Dim vDocno As String
        Dim vDocDate As String
        Dim vRefID As String
        Dim vPickZone As String
        Dim vAmount As Double
        Dim vIndex As Integer

        On Error GoTo ErrDescription

        vSearch = Me.TBSearchPickup.Text

        vQuery = "exec dbo.usp_np_SearchDriveInMaster1 '" & vSearch & "'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        Me.ListViewSearhPickup.Items.Clear()
        vIndex = 0
        If ds.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds.Tables(0).Rows.Count - 1
                vDocno = ds.Tables(0).Rows(i)("docno").ToString
                vDocDate = ds.Tables(0).Rows(i)("docdate").ToString
                vRefID = ds.Tables(0).Rows(i)("refid").ToString
                vPickZone = ds.Tables(0).Rows(i)("pickzone").ToString
                vAmount = ds.Tables(0).Rows(i)("totalnetamount").ToString

                If vPickZone = vConnectZone Then
                    vIndex = vIndex + 1
                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(vRefID)
                    listItem.SubItems.Add(vDocno)
                    listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                    Me.ListViewSearhPickup.Items.Add(listItem)
                End If

            Next
            Me.ListViewSearhPickup.Focus()
        Else
            Me.TBSearchPickup.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub LBLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vUserCode As String
        Dim vPassWord As String
        Dim vCheckTypeLogIn As String

        On Error GoTo ErrDescription

        vUserCode = Me.TBUserCode.Text
        vPassWord = Me.TBPassword.Text

        If vCheckLogIn <> "" Then
            Me.PNLogIn.Visible = False
            Me.PNDriveIn.Visible = False

            Me.TBUserID.Text = vCheckLogIn
            Call CallIDNumber()

            If Me.RDZone1.Checked = True Then
                vConnectZone = "01"
                vCheckTypeLogIn = "�ش���·��1"
            ElseIf Me.RDZone2.Checked = True Then
                vConnectZone = "02"
                vCheckTypeLogIn = "�ش���·��2"
            ElseIf Me.RDZone3.Checked = True Then
                vConnectZone = "03"
                vCheckTypeLogIn = "�ش���·��3"
            ElseIf Me.RDZone4.Checked = True Then
                vConnectZone = "04"
                vCheckTypeLogIn = "�ش���·��4"
            End If

            If vCheckTypeLogIn <> "05-Checker" Then
                Me.PNLogIn.Visible = False
                Me.PNDriveIn.Visible = True
                Me.PNDriveIn.BringToFront()
                Me.TBRefNo.Focus()
            Else
                Me.PNLogIn.Visible = False
                Me.PNDriveIn.Visible = False
            End If

        Else
            MsgBox("�������ö�����ҹ������� ��سҵ�Ǩ�ͺ����������ʼ�ҹ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBPassword.Text = ""
            Me.TBPassword.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub LBCloseLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vAnswer As Integer

        On Error Resume Next

        If vCheckLogIn = "" Then
            vAnswer = MsgBox("�س��ͧ����͡�������������� ?", MsgBoxStyle.YesNo, "Send Question Information")
            If vAnswer = 6 Then
                Application.Exit()
            Else
                Exit Sub
            End If
        Else
            Me.PNLogIn.Visible = False
        End If
    End Sub

    Private Sub LBAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vIndex As Integer
        Dim vCheckExist As Integer

        On Error GoTo ErrDescription

        If Me.ListViewStock.Items.Count > 0 And Me.TBItem.Text <> "" Then
            vCheckExist = 0
            vItemCode = Me.TBItem.Text
            vItemName = Me.TBItemName.Text
            vWHCode = Me.TBWHCode.Text
            vShelfCode = Me.TBShelfCode.Text
            If Me.TBQTY.Text <> "" Then
                vQTY = Me.TBQTY.Text
            End If
            If Me.TBPrice.Text <> "" Then
                vPrice = Me.TBPrice.Text
            End If
            vAmount = vQTY * vPrice
            vUnitCode = Me.TBUnit.Text
            vIndex = Me.ListViewItem.Items.Count + 1

            If vQTY = 0 Then
                MsgBox("������кبӹǹ�ͧ�Թ��ҷ���ͧ��� ���͵�ͧ�кبӹǹ�Թ��ҷ���ͧ����ҡ���� 0", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBQTY.Focus()
                Exit Sub
            End If


            Dim n As Integer
            Dim vCheckItemCode As String
            Dim vEditQTY As Double


            If Me.ListViewItem.Items.Count > 0 Then
                For n = 0 To Me.ListViewItem.Items.Count - 1
                    vCheckItemCode = Me.ListViewItem.Items(n).SubItems(4).Text
                    vEditQTY = Me.TBQTY.Text
                    If vItemCode = vCheckItemCode Then
                        Me.ListViewItem.Items(n).SubItems(2).Text = Format(vEditQTY, "##,##0.00")
                        Call CalcItemAmount()
                        vCheckExist = 1
                        GoTo line1
                    End If
                Next
            End If

line1:

            If vCheckExist = 0 Then
                Dim listItem As New ListViewItem(vIndex)
                listItem.SubItems.Add(vItemName)
                listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                listItem.SubItems.Add(vUnitCode)
                listItem.SubItems.Add(vItemCode)
                listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                listItem.SubItems.Add(vWHCode)
                listItem.SubItems.Add(vShelfCode)
                'listItem.SubItems.Add(vBarCode)
                Me.ListViewItem.Items.Add(listItem)
            End If

            Call CalcItemAmount()

            Me.TBItem.Text = ""
            Me.TBItemName.Text = ""
            Me.TBPrice.Text = ""
            Me.TBUnit.Text = ""
            Me.TBWHCode.Text = ""
            Me.TBShelfCode.Text = ""
            Me.TBQTY.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNItemDetails.Visible = False
            Me.BTNSave.Visible = True
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        Else
            MsgBox("�������¡���Թ����������ö���� ��¡���Թ���ŧ�С�����", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBQTY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBQTY.KeyPress
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

    Private Sub TBQTY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBQTY.TextChanged
        Dim vPrice As Double
        Dim vItemcode As String
        Dim vUnitCode As String
        Dim vQty As Double

        On Error GoTo ErrDescription

        vItemcode = Me.TBItem.Text
        vUnitCode = Me.TBUnit.Text
        If Me.TBQTY.Text <> "" Then
            vQty = Me.TBQTY.Text
        End If

        If vQty > 0 Then
            vQuery = "exec dbo.USP_NP_SearchItemPriceQty1 '" & vItemcode & "'," & vQty & ",'" & vUnitCode & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
            If ds.Tables(0).Rows.Count > 0 Then
                vPrice = ds.Tables(0).Rows(0)("saleprice1").ToString
            End If

            Me.TBPrice.Text = Format(vPrice, "##,##0.00")
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub MenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEdit.Click
        Dim vBarCode As String
        Dim vRate As Integer
        Dim vDefShelfCode As String
        Dim vStockUnit As String
        Dim i As Integer
        Dim vStore As String
        Dim vStkQTY As Double

        On Error GoTo ErrDescription

        vSelectLineEdit = Me.ListViewItem.FocusedItem.Index
        vBarCode = Me.ListViewItem.Items(vSelectLineEdit).SubItems(9).Text
        vDefShelfCode = Me.ListViewItem.Items(vSelectLineEdit).SubItems(8).Text

        Me.ListViewStock.Items.Clear()

        vQuery = "exec dbo.USP_MB_SearchBarCode '" & vBarCode & "'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vRate = ds.Tables(0).Rows(0)("rate").ToString

            For i = 0 To ds.Tables(0).Rows.Count - 1
                vStore = ds.Tables(0).Rows(i)("shelfcode").ToString
                vStkQTY = ds.Tables(0).Rows(i)("stock").ToString
                vStockUnit = ds.Tables(0).Rows(i)("stkunitcode").ToString

                If vDefShelfCode = vStore Then
                    Me.TBEditStock.Text = Format(vStkQTY, "##,##0.00")
                    Me.TBEditStockUnit.Text = vStockUnit
                End If
            Next
        End If

        Me.PNItemEdit.Visible = True
        Me.TBEditCode.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(4).Text
        Me.TBEditName.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(1).Text
        Me.TBEditUnit.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(3).Text
        Me.TBEditPrice.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(5).Text
        Me.TBEditQty.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(2).Text
        Me.TBEditRate.Text = Format(vRate, "##,##0.00")
        Me.TBDefSaleUnitCode.Text = vDefShelfCode
        Me.TBEditIndex.Text = vSelectLineEdit
        Me.TBEditQty.Focus()
        Me.TBEditQty.SelectAll()

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub LBItemEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double

        On Error GoTo ErrDescription

        If Me.TBEditQty.Text <> "" Then
            vQTY = Me.TBEditQty.Text
        End If
        If Me.TBEditPrice.Text <> "" Then
            vPrice = Me.TBEditPrice.Text
        End If
        vAmount = vQTY * vPrice

        Me.ListViewItem.Items(vSelectLineEdit).SubItems(2).Text = Format(vQTY, "##,##0.00")
        Me.ListViewItem.Items(vSelectLineEdit).SubItems(6).Text = Format(vAmount, "##,##0.00")
        Call CalcItemAmount()
        Me.TBEditQty.Text = ""
        Me.PNItemEdit.Visible = False
        Me.TBBarCode.Focus()

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub LBCloseEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next

        Me.PNItemEdit.Visible = False
        Me.TBBarCode.Focus()
    End Sub

    Private Sub MenuSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuSelect.Click
        Dim i As Integer
        Dim vDocno As String
        Dim n As Integer
        Dim vNetItemAmount As Double

        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vPickZone As String
        Dim vBarCode As String
        Dim vIndex As Integer
        Dim vShelfID As String
        Dim vZoneID As String
        Dim vPointZone As String

        On Error GoTo ErrDescription

        If Me.ListViewSearhPickup.Items.Count > 0 Then
            n = Me.ListViewSearhPickup.FocusedItem.Index
            vDocno = Me.ListViewSearhPickup.Items(n).SubItems(2).Text

            If Me.RDZone1.Checked = True Then
                vPointZone = "01"
            End If

            If Me.RDZone2.Checked = True Then
                vPointZone = "02"
            End If

            If Me.RDZone3.Checked = True Then
                vPointZone = "03"
            End If

            If Me.RDZone4.Checked = True Then
                vPointZone = "04"
            End If

            vQuery = "exec dbo.usp_np_SearchDriveInDetails1 '" & vDocno & "','" & vPointZone & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            Me.ListViewItem.Items.Clear()
            If ds.Tables(0).Rows.Count > 0 Then
                vIsOpen = 1
                vIsCancal = ds.Tables(0).Rows(i)("iscancel").ToString
                vIsconfirm = ds.Tables(0).Rows(i)("isconfirm").ToString
                vIsSendQue = ds.Tables(0).Rows(i)("issendque").ToString

                Me.TBARCode.Text = ds.Tables(0).Rows(i)("arcode").ToString
                Me.TBARName.Text = ds.Tables(0).Rows(i)("arname").ToString
                Me.TBRefNo.Text = ds.Tables(0).Rows(i)("refno").ToString
                vNetItemAmount = ds.Tables(0).Rows(i)("totalnetamount").ToString
                Me.TBItemAmount.Text = Format(vNetItemAmount, "##,##0.00")
                Me.TBDocNo.Text = ds.Tables(0).Rows(i)("docno").ToString
                Me.TBSaleCode.Text = ds.Tables(0).Rows(i)("salecode").ToString

                vIndex = 0
                vCountItemOld = ds.Tables(0).Rows.Count - 1

                ReDim vMemItemCodeOld(vCountItemOld)
                ReDim vMemUnitCodeOld(vCountItemOld)
                ReDim vMemWHCodeOld(vCountItemOld)
                ReDim vMemShelfCodeOld(vCountItemOld)
                ReDim vMemZoneIDOld(vCountItemOld)
                ReDim vMemBarCodeOld(vCountItemOld)
                ReDim vMemPickZoneOld(vCountItemOld)

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    vMemItemCodeOld(i) = ds.Tables(0).Rows(i)("itemcode").ToString
                    vMemUnitCodeOld(i) = ds.Tables(0).Rows(i)("unitcode").ToString
                    vMemWHCodeOld(i) = ds.Tables(0).Rows(i)("whcode").ToString
                    vMemShelfCodeOld(i) = ds.Tables(0).Rows(i)("shelfcode").ToString
                    vMemZoneIDOld(i) = ds.Tables(0).Rows(i)("zoneid").ToString
                    vMemBarCodeOld(i) = ds.Tables(0).Rows(i)("barcode").ToString
                    vMemPickZoneOld(i) = ds.Tables(0).Rows(i)("pickzone").ToString

                    vPickZone = ds.Tables(0).Rows(i)("pickzone").ToString
                    vItemCode = ds.Tables(0).Rows(i)("itemcode").ToString
                    vItemName = ds.Tables(0).Rows(i)("itemname").ToString
                    vWHCode = ds.Tables(0).Rows(i)("whcode").ToString
                    vShelfCode = ds.Tables(0).Rows(i)("shelfcode").ToString
                    vQTY = ds.Tables(0).Rows(i)("qty").ToString
                    vUnitCode = ds.Tables(0).Rows(i)("unitcode").ToString
                    vPrice = ds.Tables(0).Rows(i)("price").ToString
                    vAmount = ds.Tables(0).Rows(i)("amount").ToString
                    vBarCode = ds.Tables(0).Rows(i)("barcode").ToString
                    vShelfID = ds.Tables(0).Rows(i)("shelfid").ToString
                    vZoneID = ds.Tables(0).Rows(i)("zoneid").ToString

                    vIndex = vIndex + 1
                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(vItemName)
                    listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                    listItem.SubItems.Add(vUnitCode)
                    listItem.SubItems.Add(vItemCode)
                    listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                    listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                    listItem.SubItems.Add(vWHCode)
                    listItem.SubItems.Add(vShelfCode)
                    listItem.SubItems.Add(vBarCode)
                    listItem.SubItems.Add(vShelfID)
                    listItem.SubItems.Add(vZoneID)
                    listItem.SubItems.Add(vPickZone)
                    Me.ListViewItem.Items.Add(listItem)

                    If vPickZone = vPointZone Then
                        Me.ListViewItem.Items.Item(i).BackColor = Color.White
                    End If
                Next
            End If
            Me.ListViewSearhPickup.Items.Clear()
            Me.TBSearchPickup.Text = ""
            Me.PNSearchPickUp.Visible = False
            Me.PNDriveIn.Visible = True
            Me.TBRefNo.Enabled = False
            Me.PNDriveIn.BringToFront()
            Me.TBBarCode.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNCloseSelectPickup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next

        Me.ListViewSearhPickup.Items.Clear()
        Me.TBSearchPickup.Text = ""
        Me.PNSearchPickUp.Visible = False
    End Sub

    Private Sub TBEditQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBEditQty.KeyDown
        Dim vItemCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String

        Dim vShelfUnit As String
        Dim vShelfQTY As Double
        Dim vTotalQTY As Double
        Dim vRate As Integer

        Dim vAnswer As Integer
        Dim vEditIndex As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            vEditIndex = Me.TBEditIndex.Text
            Me.PNItemEdit.Visible = False
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
        End If

        If Me.TBRefNo.Text = "" Then
            MsgBox("��ͧ��͡����¹ö�����Ţ�����ҧ�ԧ��͹ �֧�����͡��Ժ�Թ�����", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBBarCode.Text = ""
            Me.TBRefNo.Focus()
            Exit Sub
        End If

        If e.KeyCode = Keys.Enter Then

            Dim vLinePickZone As String
            Dim vPickZone As String

            vLinePickZone = Me.TBPickZone.Text

            If Me.RDZone1.Checked = True Then
                vPickZone = "01"
            End If

            If Me.RDZone2.Checked = True Then
                vPickZone = "02"
            End If

            If Me.RDZone3.Checked = True Then
                vPickZone = "03"
            End If

            If Me.RDZone4.Checked = True Then
                vPickZone = "04"
            End If

            If vPickZone <> vLinePickZone Then
                MsgBox("�Թ������褹��⫹ Drive Thru �������ö��䢢������� Drive Thru � �ش�˹����ö���� � �ش�����ҹ��", MsgBoxStyle.Critical, "Send Error Message")
                Me.PNItemEdit.Visible = False
                Me.TBBarCode.Focus()
                Exit Sub
            End If

            If Me.TBEditQty.Text <> "" Then
                vQTY = Me.TBEditQty.Text
            End If
            If vQTY <= 0 Then
                MsgBox("��ͧ��˹��ӹǹ����ͧ����ҡ���� 0 ��ҹ��", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBEditQty.Focus()
                Me.TBEditQty.SelectAll()
                Exit Sub
            End If
            vEditIndex = Me.TBEditIndex.Text
            vItemCode = Me.TBEditCode.Text
            vUnitCode = Me.TBEditUnit.Text
            vShelfUnit = Me.TBEditStockUnit.Text
            If Me.TBEditRate.Text <> "" Then
                vRate = Me.TBEditRate.Text
            End If
            If Me.TBEditStock.Text <> "" Then
                vShelfQTY = Me.TBEditStock.Text
            End If

            If vShelfUnit <> vUnitCode Then
                vTotalQTY = vShelfQTY / vRate
                If vQTY > vTotalQTY Then
                    vAnswer = MsgBox("�Թ������� " & vItemCode & " STOCK ���͢�� ��ͧ��â���Թ��ҹ�� ���������", MsgBoxStyle.YesNo, "Send Question Message ")
                    If vAnswer = 7 Then
                        Me.TBEditQty.SelectAll()
                        Exit Sub
                    End If
                End If
            End If

            If vQTY > vShelfQTY And vShelfUnit = vUnitCode Then
                vAnswer = MsgBox("�Թ������� " & vItemCode & " STOCK ���͢�� ��ͧ��â���Թ��ҹ�� ���������", MsgBoxStyle.YesNo, "Send Question Message ")
                If vAnswer = 7 Then
                    Me.TBEditQty.SelectAll()
                    Exit Sub
                End If
            End If


            If Me.TBEditPrice.Text <> "" Then
                vPrice = Me.TBEditPrice.Text
            End If
            vAmount = vQTY * vPrice

            Me.ListViewItem.Items(vSelectLineEdit).SubItems(2).Text = Format(vQTY, "##,##0.00")
            Me.ListViewItem.Items(vSelectLineEdit).SubItems(6).Text = Format(vAmount, "##,##0.00")
            Call CalcItemAmount()
            Me.TBEditQty.Text = ""
            Me.PNItemEdit.Visible = False
            If Me.ListViewItem.Items.Count = 1 Then
                Me.TBBarCode.Focus()
            ElseIf vEditIndex = Me.ListViewItem.Items.Count - 1 And Me.ListViewItem.Items.Count > 1 Then
                Me.ListViewItem.Focus()
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            ElseIf vEditIndex < Me.ListViewItem.Items.Count - 1 And Me.ListViewItem.Items.Count > 1 Then
                Me.ListViewItem.Focus()
                Me.ListViewItem.Items(vEditIndex + 1).Selected = True
                Me.ListViewItem.Items(vEditIndex + 1).Focused = True
            Else
                Me.ListViewItem.Focus()
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If

            If vQTY >= 10000 Then
                MsgBox("�ӹǹ�Թ��ҷ�����͡� �Թ 10,000 ��سҵ�Ǩ�ͺ�������ա��", MsgBoxStyle.Information, "Send Error Message")
            End If


            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNClearPickUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearPickUp.Click
        On Error Resume Next

        MsgBox("���駵��� ��顴���� �����+�����Ţ 7 ���ͷӡ����ҧ˹�Ҩ�", MsgBoxStyle.Information, "Send Information Message")

        Me.TBDocNo.Text = ""
        Me.TBRefNo.Text = ""
        Me.TBRefNo.Enabled = True
        Me.TBARCode.Enabled = True
        Me.TBSaleCode.Enabled = True

        Me.TBBarCode.Text = ""
        Me.TBItemAmount.Text = ""
        Me.TBSaleCode.Text = vMemSaleName
        Me.ListViewItem.Items.Clear()
        vIsOpen = 0
        vIsCancal = 0
        vIsconfirm = 0
        vIsSendQue = 0
        Me.TBRefNo.Focus()
    End Sub

    Private Sub ClearScreen()
        On Error Resume Next

        Me.TBDocNo.Text = ""
        Me.TBRefNo.Text = ""
        Me.TBRefNo.Enabled = True
        Me.TBARCode.Enabled = True
        Me.TBSaleCode.Enabled = True

        Me.TBBarCode.Text = ""
        Me.TBItemAmount.Text = ""
        Me.TBSaleCode.Text = vMemSaleName
        Me.ListViewItem.Items.Clear()
        vIsOpen = 0
        vIsCancal = 0
        vIsconfirm = 0
        vIsSendQue = 0
        Me.TBRefNo.Focus()
    End Sub

    Private Sub RDZone1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDZone1.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D1 Then
            Me.RDZone1.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D2 Then
            Me.RDZone2.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D3 Then
            Me.RDZone3.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D4 Then
            Me.RDZone4.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("�س��ͧ����͡�ҡ��������������", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub RDZone2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDZone2.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D1 Then
            Me.RDZone1.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D2 Then
            Me.RDZone2.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D3 Then
            Me.RDZone3.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D4 Then
            Me.RDZone4.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("�س��ͧ����͡�ҡ��������������", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub RDZone3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDZone3.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D1 Then
            Me.RDZone1.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D2 Then
            Me.RDZone2.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D3 Then
            Me.RDZone3.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D4 Then
            Me.RDZone4.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("�س��ͧ����͡�ҡ��������������", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub RDZone4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDZone4.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D1 Then
            Me.RDZone1.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D2 Then
            Me.RDZone2.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D3 Then
            Me.RDZone3.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.D4 Then
            Me.RDZone4.Checked = True
            Me.TBUserCode.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Dim vAnswer As Integer
            vAnswer = MsgBox("�س��ͧ����͡�ҡ��������������", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPassword.TextChanged
        Dim vLenPassword As Integer
        Dim vCheckTypeLogIn As String

        On Error GoTo ErrDescription

        vLenPassword = Len(Me.TBPassword.Text)
        If vLenPassword = 4 And Me.TBUserCode.Text <> "" Then

            Me.TBPassword.Visible = False
            vUserCode = Me.TBUserCode.Text
            vPassWord = Me.TBPassword.Text

            Dim vService1 As New WebReference.WebServiceCalc
            Dim ds1 As DataSet = vService1.vLogIn(vUserCode, vPassWord)

            'Dim ds2 As DataSet = vService1.v

            If ds1.Tables(0).Rows.Count > 0 Then
                vCheckLogIn = ds1.Tables(0).Rows(0)("username").ToString
                vUserName = ds1.Tables(0).Rows(0)("username").ToString
                vDuty = ds1.Tables(0).Rows(0)("duty").ToString
                vLevelID = ds1.Tables(0).Rows(0)("levelid").ToString
                vMemSaleName = ds1.Tables(0).Rows(0)("salename").ToString
            Else
                vCheckLogIn = ""
                vUserName = ""
                vDuty = ""
                vLevelID = 0
                vMemSaleName = ""
            End If

            If vCheckLogIn <> "" Then

                Me.PNLogIn.Visible = False
                Me.TBUserID.Text = vCheckLogIn
                Call CallIDNumber()

                If Me.RDZone1.Checked = True Then
                    vConnectZone = "01"
                    vCheckTypeLogIn = "�ش���·��1"
                    Me.LBLZoneID.Text = "01"
                ElseIf Me.RDZone2.Checked = True Then
                    vConnectZone = "02"
                    vCheckTypeLogIn = "�ش���·��2"
                    Me.LBLZoneID.Text = "02"
                ElseIf Me.RDZone3.Checked = True Then
                    vConnectZone = "03"
                    vCheckTypeLogIn = "�ش���·��3"
                    Me.LBLZoneID.Text = "03"
                ElseIf Me.RDZone4.Checked = True Then
                    vConnectZone = "04"
                    vCheckTypeLogIn = "�ش���·��4"
                    Me.LBLZoneID.Text = "04"
                End If
                Me.TBSaleCode.Text = vMemSaleName
                Me.PNDriveIn.Visible = True
                Me.PNDriveIn.BringToFront()
                Me.TBRefNo.Focus()
            Else
                MsgBox("�������ö�����ҹ������� ��سҵ�Ǩ�ͺ����������ʼ�ҹ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBPassword.Visible = True
                Me.TBPassword.Text = ""
                Me.TBSaleCode.Text = ""
                Me.TBPassword.Focus()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub ListViewItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItem.KeyDown
        Dim vItemCode As String
        Dim vIndex As Integer
        Dim vAnswerDelete As Integer
        Dim vLinePickZone As String
        Dim vPickZone As String

        On Error GoTo ErrDescription

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If


        If e.KeyCode = Keys.Back Then
            If Me.ListViewItem.Items.Count > 0 Then
                vIndex = Me.ListViewItem.FocusedItem.Index

                vLinePickZone = Me.ListViewItem.Items(vIndex).SubItems(12).Text

                If Me.RDZone1.Checked = True Then
                    vPickZone = "01"
                End If

                If Me.RDZone2.Checked = True Then
                    vPickZone = "02"
                End If

                If Me.RDZone3.Checked = True Then
                    vPickZone = "03"
                End If

                If Me.RDZone4.Checked = True Then
                    vPickZone = "04"
                End If

                If vPickZone <> vLinePickZone Then
                    MsgBox("�Թ������褹��⫹ Drive Thru �������öź�������� Drive Thru � �ش�˹����öź � �ش�����ҹ��", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBBarCode.Focus()
                    Exit Sub
                End If


                vItemCode = Me.ListViewItem.Items(vIndex).SubItems(1).Text
                vAnswerDelete = MsgBox("�س��ͧ���ź��¡�� ���� " & vItemCode & " ������������", MsgBoxStyle.YesNo, "Send Question Message")
                If vAnswerDelete = 6 Then
                    Me.ListViewItem.Items.RemoveAt(vIndex)
                    Call GenIDNumber()
                    Call CalcItemAmount()
                    Me.TBBarCode.Focus()
                End If
            End If
        End If

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewItem.Items.Count > 0 Then
                Dim vBarCode As String
                Dim vRate As Integer
                Dim vDefShelfCode As String
                Dim vStockUnit As String
                Dim i As Integer
                Dim vStore As String
                Dim vStkQTY As Double

                On Error Resume Next

                vSelectLineEdit = Me.ListViewItem.FocusedItem.Index
                vBarCode = Me.ListViewItem.Items(vSelectLineEdit).SubItems(9).Text
                vDefShelfCode = Me.ListViewItem.Items(vSelectLineEdit).SubItems(8).Text

                vQuery = "exec dbo.USP_MB_SearchBarCode '" & vBarCode & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                Me.ListViewStock.Items.Clear()

                If ds.Tables(0).Rows.Count > 0 Then
                    vRate = ds.Tables(0).Rows(0)("rate").ToString

                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        vStore = ds.Tables(0).Rows(i)("shelfcode").ToString
                        vStkQTY = ds.Tables(0).Rows(i)("stock").ToString
                        vStockUnit = ds.Tables(0).Rows(i)("stkunitcode").ToString

                        If vDefShelfCode = vStore Then
                            Me.TBEditStock.Text = Format(vStkQTY, "##,##0.00")
                            Me.TBEditStockUnit.Text = vStockUnit
                        End If
                    Next
                End If

                Me.TBRefNo.Enabled = False
                Me.TBARCode.Enabled = False
                Me.TBSaleCode.Enabled = False

                Me.PNItemEdit.Visible = True
                Me.TBEditCode.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(4).Text
                Me.TBEditName.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(1).Text
                Me.TBEditUnit.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(3).Text
                Me.TBEditPrice.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(5).Text
                Me.TBEditQty.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(2).Text
                Me.TBPickZone.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(12).Text
                Me.TBEditRate.Text = Format(vRate, "##,##0.00")
                Me.TBDefSaleUnitCode.Text = vDefShelfCode
                Me.TBEditIndex.Text = vSelectLineEdit
                Me.TBEditQty.Focus()
                Me.TBEditQty.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.Up Then
            Dim vCount As Integer
            Dim vSelectID As Integer
            Dim i As Integer

            If Me.ListViewItem.Items.Count > 0 Then
                vCount = Me.ListViewItem.Items.Count
                For i = 0 To Me.ListViewItem.Items.Count - 1
                    If Me.ListViewItem.Items(i).Selected = True Then
                        vSelectID = i + 1
                        GoTo Line2
                    Else
                        vSelectID = 0
                    End If
                Next

            End If
Line2:
            If vSelectID = 0 Then
                Me.TBBarCode.Focus()
            ElseIf vSelectID = 1 Then
                Me.TBBarCode.Focus()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNBack.Click
        On Error GoTo ErrDescription

        MsgBox("���駵��� ��顴���� < ���͡�Ѻ��ѧ˹�����͡�ش DriveThru", MsgBoxStyle.Information, "Send Information Message")

        Me.TBRefNo.Text = ""
        Me.TBRefNo.Enabled = True
        Me.TBBarCode.Text = ""
        Me.TBItemAmount.Text = ""
        Me.TBSaleCode.Text = vMemSaleName
        Me.TBARCode.Text = ""
        Me.TBARCode.Text = "99999"
        Me.ListViewItem.Items.Clear()
        Me.PNDriveIn.Visible = False
        Me.PNLogIn.Visible = True
        Me.PNLogIn.BringToFront()
        Me.TBPassword.Text = ""
        Me.TBPassword.Visible = True
        vIsOpen = 0
        vIsCancal = 0
        vIsconfirm = 0
        vIsSendQue = 0
        Me.RDZone2.Focus()

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub PageLogIn()
        On Error Resume Next

        Me.TBRefNo.Text = ""
        Me.TBRefNo.Enabled = True
        Me.TBBarCode.Text = ""
        Me.TBItemAmount.Text = ""
        Me.TBSaleCode.Text = vMemSaleName
        Me.TBARCode.Text = ""
        Me.TBARCode.Text = "99999"
        Me.ListViewItem.Items.Clear()
        Me.PNDriveIn.Visible = False
        Me.PNLogIn.Visible = True
        Me.PNLogIn.BringToFront()
        Me.TBPassword.Text = ""
        Me.TBPassword.Visible = True
        vIsOpen = 0
        vIsCancal = 0
        vIsconfirm = 0
        vIsSendQue = 0
        Me.RDZone2.Focus()
    End Sub

    Private Sub TBEditQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBEditQty.KeyPress
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

    Private Sub TBEditQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBEditQty.TextChanged
        Dim vPrice As Double
        Dim vItemcode As String
        Dim vUnitCode As String
        Dim vQty As Double

        On Error GoTo ErrDescription

        vItemcode = Me.TBEditCode.Text
        vUnitCode = Me.TBEditUnit.Text
        If Me.TBEditQty.Text <> "" Then
            vQty = Me.TBEditQty.Text
        End If

        If vQty > 0 Then
            vQuery = "exec dbo.USP_NP_SearchItemPriceQty1 '" & vItemcode & "'," & vQty & ",'" & vUnitCode & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
            If ds.Tables(0).Rows.Count > 0 Then
                vPrice = ds.Tables(0).Rows(0)("saleprice1").ToString
            End If

            Me.TBEditPrice.Text = Format(vPrice, "##,##0.00")
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        Dim vSearch As String
        Dim i As Integer
        Dim vDocno As String
        Dim vDocDate As String
        Dim vRefID As String
        Dim vAmount As Double
        Dim vIndex As Integer

        On Error GoTo ErrDescription

        Me.PNLogIn.Visible = False
        Me.PNDriveIn.Visible = False
        Me.PNSearchPickUp.Visible = True
        Me.PNSearchPickUp.BringToFront()
        Me.TBSearchPickup.Text = ""


        MsgBox("���駵��� ��顴���� �����+�����Ţ 1 ���ͷӡ�ä��Ң�����", MsgBoxStyle.Information, "Send Information Message")


        vSearch = ""
        vQuery = "exec dbo.usp_np_SearchDriveInMaster1 '" & vSearch & "'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        Me.ListViewSearhPickup.Items.Clear()
        vIndex = 0
        If ds.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds.Tables(0).Rows.Count - 1
                vDocno = ds.Tables(0).Rows(i)("docno").ToString
                vDocDate = ds.Tables(0).Rows(i)("docdate").ToString
                vRefID = ds.Tables(0).Rows(i)("refid").ToString
                vAmount = ds.Tables(0).Rows(i)("totalnetamount").ToString

                vIndex = vIndex + 1
                Dim listItem As New ListViewItem(vIndex)
                listItem.SubItems.Add(vRefID)
                listItem.SubItems.Add(vDocno)
                listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                Me.ListViewSearhPickup.Items.Add(listItem)

            Next

            Dim a As Integer

            For a = 0 To Me.ListViewItem.Items.Count - 1
                If a Mod 2 <> 0 Then
                    Me.ListViewItem.Items(a).BackColor = Color.Silver
                End If
            Next

            Me.ListViewSearhPickup.Focus()

            Me.ListViewSearhPickup.Items(0).Focused = True
            Me.ListViewSearhPickup.Items(0).Selected = True
        Else
            Me.TBSearchPickup.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub SearchDocNo()
        Dim vSearch As String
        Dim i As Integer
        Dim vDocno As String
        Dim vDocDate As String
        Dim vRefID As String
        Dim vAmount As Double
        Dim vIndex As Integer

        On Error GoTo ErrDescription

        Me.PNLogIn.Visible = False
        Me.PNDriveIn.Visible = False
        Me.PNSearchPickUp.Visible = True
        Me.PNSearchPickUp.BringToFront()
        Me.TBSearchPickup.Text = ""

        vSearch = ""
        vQuery = "exec dbo.usp_np_SearchDriveInMaster1 '" & vSearch & "'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        Me.ListViewSearhPickup.Items.Clear()
        vIndex = 0
        If ds.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds.Tables(0).Rows.Count - 1
                vDocno = ds.Tables(0).Rows(i)("docno").ToString
                vDocDate = ds.Tables(0).Rows(i)("docdate").ToString
                vRefID = ds.Tables(0).Rows(i)("refid").ToString
                vAmount = ds.Tables(0).Rows(i)("totalnetamount").ToString

                vIndex = vIndex + 1
                Dim listItem As New ListViewItem(vIndex)
                listItem.SubItems.Add(vRefID)
                listItem.SubItems.Add(vDocno)
                listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                Me.ListViewSearhPickup.Items.Add(listItem)

            Next

            Dim a As Integer

            For a = 0 To Me.ListViewItem.Items.Count - 1
                If a Mod 2 <> 0 Then
                    Me.ListViewItem.Items(a).BackColor = Color.Silver
                End If
            Next

            Me.ListViewSearhPickup.Focus()

            Me.ListViewSearhPickup.Items(0).Focused = True
            Me.ListViewSearhPickup.Items(0).Selected = True
        Else
            Me.TBSearchPickup.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSave.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub ListViewSearhPickup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewSearhPickup.KeyDown
        Dim i As Integer
        Dim vDocno As String
        Dim n As Integer
        Dim vNetItemAmount As Double

        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vPickZone As String
        Dim vBarCode As String
        Dim vShelfID As String
        Dim vZoneID As String
        Dim vIndex As Integer
        Dim vPointZone As String

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Me.ListViewSearhPickup.Items.Clear()
            Me.TBSearchPickup.Text = ""
            Me.PNSearchPickUp.Visible = False
            Me.PNDriveIn.Visible = True
            Me.PNDriveIn.BringToFront()
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            If Me.ListViewSearhPickup.FocusedItem.Index = 0 Then
                Me.TBSearchPickup.Focus()
                Me.TBSearchPickup.SelectAll()
            End If
        End If


        If e.KeyCode = Keys.Enter Then
            'On Error Resume Next
            If Me.ListViewSearhPickup.Items.Count > 0 Then
                n = Me.ListViewSearhPickup.FocusedItem.Index
                vDocno = Me.ListViewSearhPickup.Items(n).SubItems(2).Text

                If Me.RDZone1.Checked = True Then
                    vPointZone = "01"
                End If

                If Me.RDZone2.Checked = True Then
                    vPointZone = "02"
                End If

                If Me.RDZone3.Checked = True Then
                    vPointZone = "03"
                End If

                If Me.RDZone4.Checked = True Then
                    vPointZone = "04"
                End If

                vQuery = "exec dbo.usp_np_SearchDriveInDetails1 '" & vDocno & "','" & vPointZone & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                Me.ListViewItem.Items.Clear()
                If ds.Tables(0).Rows.Count > 0 Then
                    vIsOpen = 1
                    vIsCancal = ds.Tables(0).Rows(i)("iscancel").ToString
                    vIsconfirm = ds.Tables(0).Rows(i)("isconfirm").ToString
                    vIsSendQue = ds.Tables(0).Rows(i)("issendque").ToString

                    Me.TBARCode.Text = ds.Tables(0).Rows(i)("arcode").ToString
                    Me.TBARName.Text = ds.Tables(0).Rows(i)("arname").ToString
                    Me.TBRefNo.Text = ds.Tables(0).Rows(i)("refno").ToString
                    vNetItemAmount = ds.Tables(0).Rows(i)("totalnetamount").ToString
                    Me.TBItemAmount.Text = Format(vNetItemAmount, "##,##0.00")
                    Me.TBDocNo.Text = ds.Tables(0).Rows(i)("docno").ToString
                    Me.TBSaleCode.Text = ds.Tables(0).Rows(i)("salecode").ToString

                    vIndex = 0
                    vCountItemOld = ds.Tables(0).Rows.Count - 1

                    ReDim vMemItemCodeOld(vCountItemOld)
                    ReDim vMemUnitCodeOld(vCountItemOld)
                    ReDim vMemWHCodeOld(vCountItemOld)
                    ReDim vMemShelfCodeOld(vCountItemOld)
                    ReDim vMemZoneIDOld(vCountItemOld)
                    ReDim vMemBarCodeOld(vCountItemOld)
                    ReDim vMemPickZoneOld(vCountItemOld)

                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        vMemItemCodeOld(i) = ds.Tables(0).Rows(i)("itemcode").ToString
                        vMemUnitCodeOld(i) = ds.Tables(0).Rows(i)("unitcode").ToString
                        vMemWHCodeOld(i) = ds.Tables(0).Rows(i)("whcode").ToString
                        vMemShelfCodeOld(i) = ds.Tables(0).Rows(i)("shelfcode").ToString
                        vMemZoneIDOld(i) = ds.Tables(0).Rows(i)("zoneid").ToString
                        vMemBarCodeOld(i) = ds.Tables(0).Rows(i)("barcode").ToString
                        vMemPickZoneOld(i) = ds.Tables(0).Rows(i)("pickzone").ToString

                        vPickZone = ds.Tables(0).Rows(i)("pickzone").ToString
                        vItemCode = ds.Tables(0).Rows(i)("itemcode").ToString
                        vItemName = ds.Tables(0).Rows(i)("itemname").ToString
                        vWHCode = ds.Tables(0).Rows(i)("whcode").ToString
                        vShelfCode = ds.Tables(0).Rows(i)("shelfcode").ToString
                        vQTY = ds.Tables(0).Rows(i)("qty").ToString
                        vUnitCode = ds.Tables(0).Rows(i)("unitcode").ToString
                        vPrice = ds.Tables(0).Rows(i)("price").ToString
                        vAmount = ds.Tables(0).Rows(i)("amount").ToString
                        vBarCode = ds.Tables(0).Rows(i)("barcode").ToString
                        vShelfID = ds.Tables(0).Rows(i)("shelfid").ToString
                        vZoneID = ds.Tables(0).Rows(i)("zoneid").ToString

                        vIndex = vIndex + 1
                        Dim listItem As New ListViewItem(vIndex)
                        listItem.SubItems.Add(vItemName)
                        listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                        listItem.SubItems.Add(vUnitCode)
                        listItem.SubItems.Add(vItemCode)
                        listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                        listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                        listItem.SubItems.Add(vWHCode)
                        listItem.SubItems.Add(vShelfCode)
                        listItem.SubItems.Add(vBarCode)
                        listItem.SubItems.Add(vShelfID)
                        listItem.SubItems.Add(vZoneID)
                        listItem.SubItems.Add(vPickZone)
                        Me.ListViewItem.Items.Add(listItem)

                        If vPickZone = vPointZone Then
                            Me.ListViewItem.Items.Item(i).BackColor = Color.White
                        End If
                    Next
                End If
                Me.ListViewSearhPickup.Items.Clear()
                Me.TBSearchPickup.Text = ""
                Me.PNSearchPickUp.Visible = False
                Me.PNDriveIn.Visible = True
                Me.TBRefNo.Enabled = False
                Me.PNDriveIn.BringToFront()
                Me.TBBarCode.Focus()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBSaleCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSaleCode.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Right Then
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If

        If e.KeyCode = Keys.Enter Then
            Dim vSaleCode As String
            Dim vLen As Integer
            Dim vInstr As Integer
            Dim vSearch As String

            If Me.TBSaleCode.Text <> "" Then
                vSearch = Me.TBSaleCode.Text

                If InStr(vSearch, "/") <> 0 Then
                    vInstr = InStr(vSearch, "/")
                    vLen = Len(vSearch)
                    vSaleCode = vb6.Left(vSearch, vInstr - 1)

                    vQuery = "exec dbo.USP_CRM_EmployeeDetails1  1,'" & vSaleCode & "'"
                    Dim vService As New WebReference.WebServiceCalc
                    Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                    If ds.Tables(0).Rows.Count > 0 Then
                        Me.TBSaleCode.Text = ds.Tables(0).Rows(0)("empcode").ToString & "/" & ds.Tables(0).Rows(0)("empname").ToString
                        Me.TBBarCode.Focus()
                        Me.TBBarCode.SelectAll()
                    Else
                        MsgBox("��辺�����ž�ѡ�ҹ��� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
                        Me.TBSaleCode.Focus()
                    End If

                Else
                    vQuery = "exec dbo.USP_CRM_EmployeeDetails1 1,'" & vSearch & "'"
                    Dim vService As New WebReference.WebServiceCalc
                    Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                    If ds.Tables(0).Rows.Count > 0 Then
                        Me.TBSaleCode.Text = ds.Tables(0).Rows(0)("empcode").ToString & "/" & ds.Tables(0).Rows(0)("empname").ToString
                        Me.TBBarCode.Focus()
                        Me.TBBarCode.SelectAll()
                    Else
                        MsgBox("��辺�����ž�ѡ�ҹ��� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
                        Me.TBSaleCode.Focus()
                    End If

                End If
            Else
                Me.TBSaleCode.Text = ""
                Me.TBSaleCode.Focus()
            End If

        End If
        If e.KeyCode = Keys.Down Then
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
        End If
        If e.KeyCode = Keys.Up Then
            Me.TBARCode.Focus()
            Me.TBARCode.SelectAll()
        End If

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBARCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBARCode.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Right Then
            Me.TBSaleCode.Focus()
            Me.TBSaleCode.SelectAll()
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBSaleCode.Focus()
            Me.TBSaleCode.SelectAll()
        End If

        If e.KeyCode = Keys.Enter Then
            Me.TBSaleCode.Focus()
            Me.TBSaleCode.SelectAll()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBRefNo.Focus()
            Me.TBRefNo.SelectAll()
        End If

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBARCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBARCode.TextChanged
        Dim vQuery As String
        Dim vSearchAR As String

        On Error GoTo ErrDescription

        If Me.TBARCode.Text <> "" Then
            vSearchAR = Me.TBARCode.Text

            vQuery = "exec dbo.usp_ar_searchar1 '" & vSearchAR & "' "
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            If ds.Tables(0).Rows.Count > 0 Then
                Me.TBARName.Text = ds.Tables(0).Rows(0)("arname").ToString()
                Me.TBMemberID.Text = ds.Tables(0).Rows(0)("memberid").ToString
                Me.TBSaleCode.Focus()
            Else
                Me.TBARName.Text = ""
                Me.TBMemBarCode.Text = ""
                Me.TBARCode.Focus()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNExitSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExitSend.Click
        On Error Resume Next

        Me.ListViewItem.Items.Clear()
        Me.TBARCode.Text = ""
        Me.TBRefNo.Text = ""
        Me.TBItemAmount.Text = ""
        Me.TBDocNo.Text = ""
        Me.TBBarCode.Text = ""
        Me.TBARCode.Text = "99999"
        Me.TBSaleCode.Text = vMemSaleName
        Me.TBRefNo.Enabled = True
        Me.TBRefNo.Focus()
        vIsOpen = 0
        vIsCancal = 0
        vIsconfirm = 0
        vIsSendQue = 0

        Me.ListViewItemLastSend.Items.Clear()
        Me.TBQueAR.Text = ""
        Me.TBCarLicense.Text = ""

        Me.PNLastQueSend.Visible = False
        Me.TBBarCode.Focus()
    End Sub

    Private Sub ClearSendAgian()
        On Error Resume Next

        Me.ListViewItem.Items.Clear()
        Me.TBARCode.Text = ""
        Me.TBRefNo.Text = ""
        Me.TBItemAmount.Text = ""
        Me.TBDocNo.Text = ""
        Me.TBBarCode.Text = ""
        Me.TBARCode.Text = "99999"
        Me.TBSaleCode.Text = vMemSaleName
        Me.TBRefNo.Enabled = True
        Me.TBRefNo.Focus()
        vIsOpen = 0
        vIsCancal = 0
        vIsconfirm = 0
        vIsSendQue = 0

        Me.ListViewItemLastSend.Items.Clear()
        Me.TBQueAR.Text = ""
        Me.TBCarLicense.Text = ""

        Me.PNLastQueSend.Visible = False
        Me.TBBarCode.Focus()
    End Sub

    Private Sub BTNSendAgain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSendAgain.Click
        Dim vCheckCarLicense As Integer
        Dim vCheckQueARCode As String
        Dim vInstrAR As Integer
        Dim vLenAR As Integer
        Dim vDocNo As String
        Dim vDocDate As String
        Dim a As Integer
        Dim b As Integer

        Dim vLastQueID As Integer
        Dim vLastQueDocDate As String
        Dim vLastDocNo As String

        Dim vLastItemCode As String
        Dim vLastUnitCode As String
        Dim vLastWHCode As String
        Dim vLastShelfCode As String
        Dim vLastBarCode As String
        Dim vLastPickZone As String
        Dim vLastShelfID As String
        Dim vLastZoneID As String
        Dim vLastQty As Double

        Dim vQueID As Integer
        Dim vQueDocDate As String
        Dim vItemCode As String
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vBarCode As String
        Dim vPickZone As String
        Dim vZoneID As String
        Dim vShelfID As String
        Dim vQty As Double

        Dim vPointZone As String
        Dim vMemItemExist As Integer

        Dim vCheckIsConfirm As Integer
        Dim vCheckHoldBillNo As String

        On Error GoTo ErrDescription

        If vIsOpen = 0 Then
            MsgBox("�͡��� Drive Thru  �ѧ�����ѹ�֡�������������ö�觤�� CheckOut ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBRefNo.Focus()
            Exit Sub
        End If

        MsgBox("���駵��� ��顴���� �����+�����Ţ 9 ���� ������ Enter ���ͷӡ���觤���� CheckOut", MsgBoxStyle.Information, "Send Information Message")

        vDocNo = Me.TBDocNo.Text
        vDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year

        vQuery = "exec dbo.USP_NP_CheckDocDate"
        Dim vService7 As New WebReference.WebServiceCalc
        Dim ds7 As DataSet = vService7.vGetQueryAnlyzer(vQuery)
        If ds7.Tables(0).Rows.Count > 0 Then
            vDocDate = ds7.Tables(0).Rows(0)("vdocdate").ToString
        End If


        vInstrAR = InStr(Me.TBQueAR.Text, "/")
        vLenAR = Len(Me.TBQueAR.Text)
        vCheckQueARCode = vb6.Left(Me.TBQueAR.Text, vInstrAR - 1)

        vQuery = "exec dbo.USP_NP_CheckQueHoldBill1 '" & vDocNo & "','" & vCheckQueARCode & "'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vCheckIsConfirm = ds.Tables(0).Rows(0)("isconfirm").ToString()
            vCheckHoldBillNo = ds.Tables(0).Rows(0)("holdbillno").ToString()
        End If

        If vCheckIsConfirm = 1 And vCheckHoldBillNo <> "" Then
            MsgBox("�Ţ����͡�����ӡ�� ��Ǩ�ͺ�Թ�����з��͡��þѡ������º�������� �������ö�觤������͡����� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
            Me.ListViewItemLastSend.Items.Clear()
            Me.TBCarLicense.Text = ""
            Me.TBQueAR.Text = ""
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
            Exit Sub
        End If


        If Me.RDZone1.Checked = True Then
            vZoneID = "A"
            vPointZone = "01"
        End If

        If Me.RDZone2.Checked = True Then
            vZoneID = "B"
            vPointZone = "02"
        End If

        If Me.RDZone3.Checked = True Then
            vZoneID = "C"
            vPointZone = "03"
        End If

        If Me.RDZone4.Checked = True Then
            vZoneID = "D"
            vPointZone = "04"
        End If

        If Me.ListViewItem.Items.Count > 0 Then

            For a = 0 To Me.ListViewItemLastSend.Items.Count - 1

                vLastQueID = Me.ListViewItemLastSend.Items(a).SubItems(4).Text
                vLastQueDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
                vLastDocNo = Me.ListViewItemLastSend.Items(a).SubItems(6).Text
                vLastItemCode = Me.ListViewItemLastSend.Items(a).SubItems(7).Text
                vLastUnitCode = Me.ListViewItemLastSend.Items(a).SubItems(3).Text
                vLastWHCode = Me.ListViewItemLastSend.Items(a).SubItems(8).Text
                vLastShelfCode = Me.ListViewItemLastSend.Items(a).SubItems(9).Text
                vLastBarCode = Me.ListViewItemLastSend.Items(a).SubItems(10).Text
                vLastPickZone = Me.ListViewItemLastSend.Items(a).SubItems(11).Text
                vLastZoneID = Me.ListViewItemLastSend.Items(a).SubItems(5).Text
                vLastShelfID = Me.ListViewItemLastSend.Items(a).SubItems(12).Text
                vLastQty = Me.ListViewItemLastSend.Items(a).SubItems(2).Text

                For b = 0 To Me.ListViewItem.Items.Count - 1
                    vItemCode = Me.ListViewItem.Items(b).SubItems(4).Text
                    vUnitCode = Me.ListViewItem.Items(b).SubItems(3).Text
                    vWHCode = Me.ListViewItem.Items(b).SubItems(7).Text
                    vShelfCode = Me.ListViewItem.Items(b).SubItems(8).Text
                    vBarCode = Me.ListViewItem.Items(b).SubItems(9).Text
                    vPickZone = Me.ListViewItem.Items(b).SubItems(12).Text
                    vQty = Me.ListViewItem.Items(b).SubItems(2).Text

                    If vLastItemCode = vItemCode And vLastUnitCode = vUnitCode And vLastWHCode = vWHCode And vLastShelfCode = vShelfCode And vLastPickZone = vPickZone And vLastBarCode = vBarCode Then
                        vMemItemExist = 1
                        GoTo Line1
                    Else
                        vMemItemExist = 0
                    End If

                Next
Line1:

                If vMemItemExist = 0 And vLastPickZone = vPointZone Then
                    vQuery = "exec dbo.USP_NP_InsertUpdateCancelQueItem1 1," & vLastQueID & ",'" & vLastQueDocDate & "','" & vLastItemCode & "','" & vLastWHCode & "','" & vLastShelfCode & "','" & vLastShelfID & "','" & vLastZoneID & "','" & vLastPickZone & "','" & vLastDocNo & "','" & vLastBarCode & "'," & vLastQty & ",'" & vLastUnitCode & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)
                End If
            Next


            For a = 0 To Me.ListViewItem.Items.Count - 1
                vQueID = Me.ListViewItemLastSend.Items(0).SubItems(4).Text
                vQueDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
                vDocNo = Me.TBDocNo.Text
                vItemCode = Me.ListViewItem.Items(a).SubItems(4).Text
                vUnitCode = Me.ListViewItem.Items(a).SubItems(3).Text
                vWHCode = Me.ListViewItem.Items(a).SubItems(7).Text
                vShelfCode = Me.ListViewItem.Items(a).SubItems(8).Text
                vBarCode = Me.ListViewItem.Items(a).SubItems(9).Text
                vPickZone = Me.ListViewItem.Items(a).SubItems(12).Text
                vQty = Me.ListViewItem.Items(a).SubItems(2).Text
                vShelfID = Me.ListViewItem.Items(a).SubItems(10).Text

                For b = 0 To Me.ListViewItemLastSend.Items.Count - 1
                    vLastItemCode = Me.ListViewItemLastSend.Items(b).SubItems(7).Text
                    vLastUnitCode = Me.ListViewItemLastSend.Items(b).SubItems(3).Text
                    vLastWHCode = Me.ListViewItemLastSend.Items(b).SubItems(8).Text
                    vLastShelfCode = Me.ListViewItemLastSend.Items(b).SubItems(9).Text
                    vLastBarCode = Me.ListViewItemLastSend.Items(b).SubItems(10).Text
                    vLastPickZone = Me.ListViewItemLastSend.Items(b).SubItems(11).Text
                    vLastZoneID = Me.ListViewItemLastSend.Items(b).SubItems(5).Text


                    If vLastItemCode = vItemCode And vLastUnitCode = vUnitCode And vLastWHCode = vWHCode And vLastShelfCode = vShelfCode And vLastPickZone = vPickZone And vLastBarCode = vBarCode Then
                        vMemItemExist = 1
                        GoTo Line2
                    Else
                        vMemItemExist = 0
                    End If

                Next
Line2:
                If vPickZone = vPointZone Then
                    vQuery = "exec dbo.USP_NP_InsertUpdateCancelQueItem1 2," & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "','" & vPickZone & "','" & vDocNo & "','" & vBarCode & "'," & vQty & ",'" & vUnitCode & "'"
                    Dim vService2 As New WebReference.WebServiceCalc
                    Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)
                End If
            Next

            vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto1 3,'" & vDocNo & "','" & vPointZone & "','" & vUserName & "'"
            Dim vService3 As New WebReference.WebServiceCalc
            Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

            MsgBox("����¡���Թ���价ӡ�� CheckOut ���º��������", MsgBoxStyle.Information, "Send Information Message")

            Me.ListViewItem.Items.Clear()
            Me.TBARCode.Text = ""
            Me.TBRefNo.Text = ""
            Me.TBItemAmount.Text = ""
            Me.TBDocNo.Text = ""
            Me.TBBarCode.Text = ""
            Me.TBARCode.Text = "99999"
            Me.ListViewItemLastSend.Items.Clear()
            Me.TBQueAR.Text = ""
            Me.TBCarLicense.Text = ""
            Me.PNLastQueSend.Visible = False
            Me.TBRefNo.Enabled = True
            Me.TBRefNo.Focus()

            vIsOpen = 0
            vIsCancal = 0
            vIsconfirm = 0
            vIsSendQue = 0
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub SendQueAgain()
        Dim vCheckCarLicense As Integer
        Dim vCheckQueARCode As String
        Dim vInstrAR As Integer
        Dim vLenAR As Integer
        Dim vDocNo As String
        Dim vDocDate As String
        Dim a As Integer
        Dim b As Integer

        Dim vLastQueID As Integer
        Dim vLastQueDocDate As String
        Dim vLastDocNo As String

        Dim vLastItemCode As String
        Dim vLastUnitCode As String
        Dim vLastWHCode As String
        Dim vLastShelfCode As String
        Dim vLastBarCode As String
        Dim vLastPickZone As String
        Dim vLastShelfID As String
        Dim vLastZoneID As String
        Dim vLastQty As Double

        Dim vQueID As Integer
        Dim vQueDocDate As String
        Dim vItemCode As String
        Dim vUnitCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vBarCode As String
        Dim vPickZone As String
        Dim vZoneID As String
        Dim vShelfID As String
        Dim vQty As Double

        Dim vPointZone As String
        Dim vMemItemExist As Integer

        Dim vCheckIsConfirm As Integer
        Dim vCheckHoldBillNo As String

        On Error GoTo ErrDescription

        If vIsOpen = 0 Then
            MsgBox("�͡��� Drive Thru  �ѧ�����ѹ�֡�������������ö�觤�� CheckOut ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBRefNo.Focus()
            Exit Sub
        End If

        vDocNo = Me.TBDocNo.Text
        vDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year

        vQuery = "exec dbo.USP_NP_CheckDocDate"
        Dim vService7 As New WebReference.WebServiceCalc
        Dim ds7 As DataSet = vService7.vGetQueryAnlyzer(vQuery)
        If ds7.Tables(0).Rows.Count > 0 Then
            vDocDate = ds7.Tables(0).Rows(0)("vdocdate").ToString
        End If

        vInstrAR = InStr(Me.TBQueAR.Text, "/")
        vLenAR = Len(Me.TBQueAR.Text)
        vCheckQueARCode = vb6.Left(Me.TBQueAR.Text, vInstrAR - 1)

        vQuery = "exec dbo.USP_NP_CheckQueHoldBill1 '" & vDocNo & "','" & vCheckQueARCode & "'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vCheckIsConfirm = ds.Tables(0).Rows(0)("isconfirm").ToString()
            vCheckHoldBillNo = ds.Tables(0).Rows(0)("holdbillno").ToString()
        End If

        If vCheckIsConfirm = 1 And vCheckHoldBillNo <> "" Then
            MsgBox("�Ţ����͡�����ӡ�� ��Ǩ�ͺ�Թ�����з��͡��þѡ������º�������� �������ö�觤������͡����� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
            Me.ListViewItemLastSend.Items.Clear()
            Me.TBCarLicense.Text = ""
            Me.TBQueAR.Text = ""
            Me.TBBarCode.Focus()
            Me.TBBarCode.SelectAll()
            Exit Sub
        End If


        If Me.RDZone1.Checked = True Then
            vZoneID = "A"
            vPointZone = "01"
        End If

        If Me.RDZone2.Checked = True Then
            vZoneID = "B"
            vPointZone = "02"
        End If

        If Me.RDZone3.Checked = True Then
            vZoneID = "C"
            vPointZone = "03"
        End If

        If Me.RDZone4.Checked = True Then
            vZoneID = "D"
            vPointZone = "04"
        End If


        If Me.ListViewItem.Items.Count > 0 Then

            For a = 0 To Me.ListViewItemLastSend.Items.Count - 1

                vLastQueID = Me.ListViewItemLastSend.Items(a).SubItems(4).Text
                vLastQueDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
                vLastDocNo = Me.ListViewItemLastSend.Items(a).SubItems(6).Text
                vLastItemCode = Me.ListViewItemLastSend.Items(a).SubItems(7).Text
                vLastUnitCode = Me.ListViewItemLastSend.Items(a).SubItems(3).Text
                vLastWHCode = Me.ListViewItemLastSend.Items(a).SubItems(8).Text
                vLastShelfCode = Me.ListViewItemLastSend.Items(a).SubItems(9).Text
                vLastBarCode = Me.ListViewItemLastSend.Items(a).SubItems(10).Text
                vLastPickZone = Me.ListViewItemLastSend.Items(a).SubItems(11).Text
                vLastZoneID = Me.ListViewItemLastSend.Items(a).SubItems(5).Text
                vLastShelfID = Me.ListViewItemLastSend.Items(a).SubItems(12).Text
                vLastQty = Me.ListViewItemLastSend.Items(a).SubItems(2).Text

                For b = 0 To Me.ListViewItem.Items.Count - 1
                    vItemCode = Me.ListViewItem.Items(b).SubItems(4).Text
                    vUnitCode = Me.ListViewItem.Items(b).SubItems(3).Text
                    vWHCode = Me.ListViewItem.Items(b).SubItems(7).Text
                    vShelfCode = Me.ListViewItem.Items(b).SubItems(8).Text
                    vBarCode = Me.ListViewItem.Items(b).SubItems(9).Text
                    vPickZone = Me.ListViewItem.Items(b).SubItems(12).Text
                    vQty = Me.ListViewItem.Items(b).SubItems(2).Text

                    If vLastItemCode = vItemCode And vLastUnitCode = vUnitCode And vLastWHCode = vWHCode And vLastShelfCode = vShelfCode And vLastPickZone = vPickZone And vLastBarCode = vBarCode Then
                        vMemItemExist = 1
                        GoTo Line1
                    Else
                        vMemItemExist = 0
                    End If

                Next
Line1:

                If vMemItemExist = 0 And vLastPickZone = vPointZone Then
                    vQuery = "exec dbo.USP_NP_InsertUpdateCancelQueItem1 1," & vLastQueID & ",'" & vLastQueDocDate & "','" & vLastItemCode & "','" & vLastWHCode & "','" & vLastShelfCode & "','" & vLastShelfID & "','" & vLastZoneID & "','" & vLastPickZone & "','" & vLastDocNo & "','" & vLastBarCode & "'," & vLastQty & ",'" & vLastUnitCode & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)
                End If
            Next


            For a = 0 To Me.ListViewItem.Items.Count - 1
                vQueID = Me.ListViewItemLastSend.Items(0).SubItems(4).Text
                vQueDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year
                vDocNo = Me.TBDocNo.Text
                vItemCode = Me.ListViewItem.Items(a).SubItems(4).Text
                vUnitCode = Me.ListViewItem.Items(a).SubItems(3).Text
                vWHCode = Me.ListViewItem.Items(a).SubItems(7).Text
                vShelfCode = Me.ListViewItem.Items(a).SubItems(8).Text
                vBarCode = Me.ListViewItem.Items(a).SubItems(9).Text
                vPickZone = Me.ListViewItem.Items(a).SubItems(12).Text
                vQty = Me.ListViewItem.Items(a).SubItems(2).Text
                vShelfID = Me.ListViewItem.Items(a).SubItems(10).Text

                For b = 0 To Me.ListViewItemLastSend.Items.Count - 1
                    vLastItemCode = Me.ListViewItemLastSend.Items(b).SubItems(7).Text
                    vLastUnitCode = Me.ListViewItemLastSend.Items(b).SubItems(3).Text
                    vLastWHCode = Me.ListViewItemLastSend.Items(b).SubItems(8).Text
                    vLastShelfCode = Me.ListViewItemLastSend.Items(b).SubItems(9).Text
                    vLastBarCode = Me.ListViewItemLastSend.Items(b).SubItems(10).Text
                    vLastPickZone = Me.ListViewItemLastSend.Items(b).SubItems(11).Text
                    vLastZoneID = Me.ListViewItemLastSend.Items(b).SubItems(5).Text


                    If vLastItemCode = vItemCode And vLastUnitCode = vUnitCode And vLastWHCode = vWHCode And vLastShelfCode = vShelfCode And vLastPickZone = vPickZone And vLastBarCode = vBarCode Then
                        vMemItemExist = 1
                        GoTo Line2
                    Else
                        vMemItemExist = 0
                    End If

                Next
Line2:
                If vPickZone = vPointZone Then
                    vQuery = "exec dbo.USP_NP_InsertUpdateCancelQueItem1 2," & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "','" & vPickZone & "','" & vDocNo & "','" & vBarCode & "'," & vQty & ",'" & vUnitCode & "'"
                    Dim vService2 As New WebReference.WebServiceCalc
                    Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)
                End If
            Next

            vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto1 3,'" & vDocNo & "','" & vPointZone & "','" & vUserName & "'"
            Dim vService3 As New WebReference.WebServiceCalc
            Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

            MsgBox("����¡���Թ���价ӡ�� CheckOut ���º��������", MsgBoxStyle.Information, "Send Information Message")

            Me.ListViewItem.Items.Clear()
            Me.TBARCode.Text = ""
            Me.TBRefNo.Text = ""
            Me.TBItemAmount.Text = ""
            Me.TBDocNo.Text = ""
            Me.TBBarCode.Text = ""
            Me.TBARCode.Text = "99999"
            Me.ListViewItemLastSend.Items.Clear()
            Me.TBQueAR.Text = ""
            Me.TBCarLicense.Text = ""
            Me.PNLastQueSend.Visible = False
            Me.TBRefNo.Enabled = True
            Me.TBRefNo.Focus()

            vIsOpen = 0
            vIsCancal = 0
            vIsconfirm = 0
            vIsSendQue = 0
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBARName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBARName.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBARName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBARName.TextChanged

    End Sub

    Private Sub TBMemberID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBMemberID.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBUserID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBUserID.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNBack_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNBack.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNClearPickUp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClearPickUp.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSearch.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub ListViewItemLastSend_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItemLastSend.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Call SendQueAgain()
        End If

        If e.KeyCode = 34 Then
            Call SendQueAgain()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearSendAgian()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNExitSend_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNExitSend.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Call SendQueAgain()
        End If

        If e.KeyCode = 34 Then
            Call SendQueAgain()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearSendAgian()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSendAgain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSendAgain.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Call SendQueAgain()
        End If

        If e.KeyCode = 34 Then
            Call SendQueAgain()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearSendAgian()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub


    Private Sub ListViewStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewStock.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Me.PNItemDetails.Visible = False
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBBarCode.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBItem.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Me.PNItemDetails.Visible = False
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBBarCode.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub


    Private Sub TBItemName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBItemName.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Me.PNItemDetails.Visible = False
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBBarCode.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBUnit.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Me.PNItemDetails.Visible = False
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBBarCode.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBPrice.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Me.PNItemDetails.Visible = False
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBBarCode.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBReserve_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBReserve.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Me.PNItemDetails.Visible = False
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBBarCode.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBEditCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBEditCode.KeyDown
        Dim vEditIndex As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            vEditIndex = Me.TBEditIndex.Text
            Me.PNItemEdit.Visible = False
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBPickZone_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBPickZone.KeyDown
        Dim vEditIndex As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            vEditIndex = Me.TBEditIndex.Text
            Me.PNItemEdit.Visible = False
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBEditName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBEditName.KeyDown
        Dim vEditIndex As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            vEditIndex = Me.TBEditIndex.Text
            Me.PNItemEdit.Visible = False
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBEditUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBEditUnit.KeyDown
        Dim vEditIndex As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            vEditIndex = Me.TBEditIndex.Text
            Me.PNItemEdit.Visible = False
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBEditPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBEditPrice.KeyDown
        Dim vEditIndex As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            vEditIndex = Me.TBEditIndex.Text
            Me.PNItemEdit.Visible = False
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBEditRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBEditRate.KeyDown
        Dim vEditIndex As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            vEditIndex = Me.TBEditIndex.Text
            Me.PNItemEdit.Visible = False
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBDefSaleUnitCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBDefSaleUnitCode.KeyDown
        Dim vEditIndex As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            vEditIndex = Me.TBEditIndex.Text
            Me.PNItemEdit.Visible = False
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBEditStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBEditStock.KeyDown
        Dim vEditIndex As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            vEditIndex = Me.TBEditIndex.Text
            Me.PNItemEdit.Visible = False
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBEditStockUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBEditStockUnit.KeyDown
        Dim vEditIndex As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            vEditIndex = Me.TBEditIndex.Text
            Me.PNItemEdit.Visible = False
            Me.ListViewItem.Focus()
            If Me.ListViewItem.Items.Count > 0 Then
                Me.ListViewItem.Items(vEditIndex).Selected = True
                Me.ListViewItem.Items(vEditIndex).Focused = True
            End If
            Me.TBRefNo.Enabled = True
            Me.TBARCode.Enabled = True
            Me.TBSaleCode.Enabled = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBCarLicense_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBCarLicense.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Call SendQueAgain()
        End If

        If e.KeyCode = 34 Then
            Call SendQueAgain()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearSendAgian()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBQueAR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBQueAR.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Call SendQueAgain()
        End If

        If e.KeyCode = 34 Then
            Call SendQueAgain()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ClearSendAgian()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCancel.Click
        Dim vDocNo As String
        Dim vCheckIsConfirm As Integer
        Dim vCheckIsCancel As Integer
        Dim vCheckHoldBillNo As String
        Dim vARCode As String
        Dim vAnswer As Integer

        On Error GoTo ErrDescription

        If vIsOpen = 1 Then
            If Me.TBRefNo.Text <> "" And Me.TBDocNo.Text <> "" And Me.ListViewItem.Items.Count > 0 Then
                Call BeforeSaveData()
                vDocNo = Me.TBDocNo.Text
                vARCode = Me.TBARCode.Text

                vQuery = "exec dbo.USP_NP_CheckQueHoldBill1 '" & vDocNo & "','" & vARCode & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count > 0 Then
                    vCheckIsConfirm = ds.Tables(0).Rows(0)("isconfirm").ToString()
                    vCheckIsCancel = ds.Tables(0).Rows(0)("iscancel").ToString()
                    vCheckHoldBillNo = ds.Tables(0).Rows(0)("holdbillno").ToString()
                End If

                If vCheckIsCancel = 1 Then
                    MsgBox("�Ţ����͡�����ӡ��¡��ԡ����º�������� �������ö¡��ԡ�͡����� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBRefNo.Focus()
                    Me.TBRefNo.SelectAll()
                    Exit Sub
                End If

                If vCheckIsConfirm = 1 And vCheckHoldBillNo <> "" Then
                    MsgBox("�Ţ����͡�����ӡ�� ��Ǩ�ͺ�Թ�����з��͡��þѡ������º�������� �������ö¡��ԡ�͡����� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBRefNo.Focus()
                    Me.TBRefNo.SelectAll()
                    Exit Sub
                End If

                vAnswer = MsgBox("�س��ͧ���¡��ԡ�͡��� DriveThru �Ţ��� " & vDocNo & " ������������", MsgBoxStyle.YesNo, "Send Query Message ?")

                If vAnswer = 6 Then
                    vQuery = "exec dbo.USP_NP_CancelDriveInDocNo '" & vDocNo & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)
                Else
                    Me.TBARCode.Enabled = True
                    Me.TBSaleCode.Enabled = True
                    Me.TBBarCode.Enabled = True
                    Me.ListViewItem.Enabled = True
                    Me.BTNBack.Enabled = True
                    Me.BTNClearPickUp.Enabled = True
                    Me.BTNSave.Enabled = True
                    Me.BTNSearch.Enabled = True
                    Me.BTNClosePickup.Enabled = True
                    Me.BTNCancel.Enabled = True

                    Me.TBARCode.Focus()
                    Me.TBARCode.SelectAll()
                    Exit Sub
                End If
                Call AfterSaveData()
                Call ClearScreen()
                MsgBox("¡��ԡ�Ţ����͡��� " & vDocNo & " ���º��������", MsgBoxStyle.Information, "Send Information Message")
                Me.TBRefNo.Focus()
                Me.TBRefNo.SelectAll()
            Else
                MsgBox("�����Ţͧ�͡��������ú ��سҵ�Ǩ�ͺ", MsgBoxStyle.Information, "Send Information Message")
                Me.TBRefNo.Focus()
                Me.TBRefNo.SelectAll()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub CancelDriveIn()
        Dim vDocNo As String
        Dim vCheckIsConfirm As Integer
        Dim vCheckIsCancel As Integer
        Dim vCheckHoldBillNo As String
        Dim vARCode As String
        Dim vAnswer As Integer

        On Error GoTo ErrDescription

        If vIsOpen = 1 Then
            If Me.TBRefNo.Text <> "" And Me.TBDocNo.Text <> "" And Me.ListViewItem.Items.Count > 0 Then
                Call BeforeSaveData()
                vDocNo = Me.TBDocNo.Text
                vARCode = Me.TBARCode.Text

                vQuery = "exec dbo.USP_NP_CheckQueHoldBill1 '" & vDocNo & "','" & vARCode & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count > 0 Then
                    vCheckIsConfirm = ds.Tables(0).Rows(0)("isconfirm").ToString()
                    vCheckIsCancel = ds.Tables(0).Rows(0)("iscancel").ToString()
                    vCheckHoldBillNo = ds.Tables(0).Rows(0)("holdbillno").ToString()
                End If

                If vCheckIsCancel = 1 Then
                    MsgBox("�Ţ����͡�����ӡ��¡��ԡ����º�������� �������ö¡��ԡ�͡����� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBRefNo.Focus()
                    Me.TBRefNo.SelectAll()
                    Exit Sub
                End If

                If vCheckIsConfirm = 1 And vCheckHoldBillNo <> "" Then
                    MsgBox("�Ţ����͡�����ӡ�� ��Ǩ�ͺ�Թ�����з��͡��þѡ������º�������� �������ö¡��ԡ�͡����� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBRefNo.Focus()
                    Me.TBRefNo.SelectAll()
                    Exit Sub
                End If

                vAnswer = MsgBox("�س��ͧ���¡��ԡ�͡��� DriveThru �Ţ��� " & vDocNo & " ������������", MsgBoxStyle.YesNo, "Send Query Message ?")

                If vAnswer = 6 Then
                    vQuery = "exec dbo.USP_NP_CancelDriveInDocNo '" & vDocNo & "'"
                    Dim vService1 As New WebReference.WebServiceCalc
                    Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)
                Else
                    Me.TBARCode.Enabled = True
                    Me.TBSaleCode.Enabled = True
                    Me.TBBarCode.Enabled = True
                    Me.ListViewItem.Enabled = True
                    Me.BTNBack.Enabled = True
                    Me.BTNClearPickUp.Enabled = True
                    Me.BTNSave.Enabled = True
                    Me.BTNSearch.Enabled = True
                    Me.BTNClosePickup.Enabled = True
                    Me.BTNCancel.Enabled = True

                    Me.TBARCode.Focus()
                    Me.TBARCode.SelectAll()
                    Exit Sub
                End If
                Call AfterSaveData()
                Call ClearScreen()
                MsgBox("¡��ԡ�Ţ����͡��� " & vDocNo & " ���º��������", MsgBoxStyle.Information, "Send Information Message")
                Me.TBRefNo.Focus()
                Me.TBRefNo.SelectAll()
            Else
                MsgBox("�����Ţͧ�͡��������ú ��سҵ�Ǩ�ͺ", MsgBoxStyle.Information, "Send Information Message")
                Me.TBRefNo.Focus()
                Me.TBRefNo.SelectAll()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNCancel_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCancel.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 34 Then
            Call SavePickUp()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SearchDocNo()
        End If

        If e.KeyCode = 115 Then
            Call CancelDriveIn()
        End If

        If e.KeyCode = 37 Then
            Call PageLogIn()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub
End Class
