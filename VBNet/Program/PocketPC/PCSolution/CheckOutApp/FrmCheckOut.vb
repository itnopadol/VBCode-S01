Imports System.Data
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms

Public Class FrmCheckOut
    Dim ds As DataSet
    Dim vQuery As String

    Dim vMemSaleName As String
    Dim vUserCode As String
    Dim vPassWord As String

    Dim vIsOpen As Integer

    Private Sub BTNLogIN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vUserCode As String
        Dim vPassWord As String
        Dim vCheckTypeLogIn As String

        On Error GoTo ErrDescription

        vUserCode = Me.TBUserCode.Text
        vPassWord = Me.TBPassword.Text

        If vCheckLogIn <> "" Then
            Me.PNLogIn.Visible = False
            Me.PNChecker.Visible = False

            vConnectZone = "05"
            vCheckTypeLogIn = "�ش������"

            Me.TBUserID.Text = vUserCode
            vUserID = vUserCode
            Me.PNLogIn.Visible = False
            Me.PNChecker.Visible = True
            Me.PNChecker.BringToFront()
            Me.TBSearchCheckOut.Focus()

        Else
            MsgBox("�������ö�����ҹ������� ��سҵ�Ǩ�ͺ����������ʼ�ҹ", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBPassword.Text = ""
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
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBPassword.Focus()
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

    Private Sub GenIDNumberMerge()
        Dim i As Integer
        Dim j As Integer

        On Error GoTo ErrDescription

        If Me.ListViewMerge.Items.Count > 0 Then
            j = 0
            For i = 0 To Me.ListViewMerge.Items.Count - 1
                j = j + 1
                Me.ListViewMerge.Items(i).SubItems(0).Text = j
            Next
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
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

    Private Sub SearchItemCheckOut()
        Dim vRefNo As String
        Dim i As Integer
        Dim vDocno As String
        Dim vDocDate As String
        Dim vQueID As Integer
        Dim vPickZone As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vShelfID As String
        Dim vQTY As Double
        Dim vCarQTY As Double
        Dim vUnitCode As String
        Dim vPrice As Double
        Dim vDisCountAmount As Double
        Dim vPicker As String
        Dim vAmount As Double
        Dim vIndex As Integer
        Dim vLine As Integer
        Dim vBarCode As String

        Dim vMemStatus As Integer
        Dim vMemQty As Double
        Dim vMemPickQty As Double
        Dim vItemDesc As String
        Dim vARCode As String
        Dim vSaleCode As String

        On Error GoTo ErrDescription

        If Me.TBSearchCheckOut.Text <> "" Then
            vRefNo = Me.TBSearchCheckOut.Text

            vQuery = "exec dbo.usp_np_SearchQueCheckOut1 '" & vRefNo & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            vIndex = 0
            If ds.Tables(0).Rows.Count > 0 Then

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    vMemStatus = ds.Tables(0).Rows(i)("questatus").ToString
                    vMemQty = ds.Tables(0).Rows(i)("qty").ToString
                    vMemPickQty = ds.Tables(0).Rows(i)("remainqty").ToString

                    If vMemStatus = 2 And vMemQty = vMemPickQty Then
                        vItemDesc = "�ú"
                    ElseIf vMemStatus = 2 And vMemQty < vMemPickQty Then
                        vItemDesc = "�Թ"
                    ElseIf vMemStatus = 2 And vMemQty > vMemPickQty Then
                        vItemDesc = "���ú"
                    Else
                        vItemDesc = ds.Tables(0).Rows(i)("quedescription").ToString
                    End If

                    vDocno = ds.Tables(0).Rows(i)("docno").ToString
                    vDocDate = ds.Tables(0).Rows(i)("docdate").ToString
                    vQueID = ds.Tables(0).Rows(i)("queid").ToString
                    vPickZone = ds.Tables(0).Rows(i)("pickzone").ToString
                    vItemCode = ds.Tables(0).Rows(i)("itemcode").ToString
                    vItemName = ds.Tables(0).Rows(i)("itemname").ToString
                    vWHCode = ds.Tables(0).Rows(i)("whcode").ToString
                    vShelfCode = ds.Tables(0).Rows(i)("shelfcode").ToString
                    vQTY = ds.Tables(0).Rows(i)("qty").ToString
                    vCarQTY = ds.Tables(0).Rows(i)("remainqty").ToString
                    vUnitCode = ds.Tables(0).Rows(i)("unitcode").ToString
                    vPrice = ds.Tables(0).Rows(i)("price").ToString
                    vAmount = ds.Tables(0).Rows(i)("netamount").ToString
                    vBarCode = ds.Tables(0).Rows(i)("barcode").ToString
                    vPrice = ds.Tables(0).Rows(i)("price").ToString
                    vDisCountAmount = ds.Tables(0).Rows(i)("discountamount").ToString
                    vAmount = ds.Tables(0).Rows(i)("netamount").ToString
                    vShelfID = ds.Tables(0).Rows(i)("shelfid").ToString
                    vPicker = ds.Tables(0).Rows(i)("quepicker").ToString
                    vARCode = ds.Tables(0).Rows(i)("arcode").ToString
                    vSaleCode = ds.Tables(0).Rows(i)("salecode").ToString

                    vIndex = vIndex + 1
                    vLine = vIndex - 1
                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(vQueID)
                    listItem.SubItems.Add(vItemDesc)
                    listItem.SubItems.Add(vItemName)
                    listItem.SubItems.Add(Format(vCarQTY, "##,##0.00"))
                    listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                    listItem.SubItems.Add(vUnitCode)
                    listItem.SubItems.Add(vItemCode)
                    listItem.SubItems.Add(vWHCode)
                    listItem.SubItems.Add(vShelfCode)
                    listItem.SubItems.Add(vPickZone)
                    listItem.SubItems.Add(vDocno)
                    listItem.SubItems.Add(vBarCode)
                    listItem.SubItems.Add(vPickZone)
                    listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                    listItem.SubItems.Add(Format(vDisCountAmount, "##,##0.00"))
                    listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                    listItem.SubItems.Add(vARCode)
                    listItem.SubItems.Add(vSaleCode)
                    listItem.SubItems.Add(vRefNo)
                    listItem.SubItems.Add(vShelfID)
                    listItem.SubItems.Add(vPicker)
                    listItem.SubItems.Add(vMemStatus)
                Next
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub GenHoldingBill()
        Dim i As Integer
        Dim vCount As Integer
        Dim vCheck As Integer
        Dim vCheckZero As Integer
        Dim vIndex As Integer
        Dim vARCode As String
        Dim vQTY As Double

        On Error GoTo ErrDescription

        If Me.ListViewMerge.Items.Count = 0 Then
            MsgBox("�������¡���Թ���㹵��ҧ ���зӡ�þѡ��� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

        For i = 0 To Me.ListViewMerge.Items.Count - 1

            If Me.ListViewMerge.Items(i).SubItems(1).Text <> "" Then
                vCheck = vCheck + 1
            End If
        Next

        For i = 0 To Me.ListViewMerge.Items.Count - 1

            If Me.ListViewMerge.Items(i).SubItems(1).Text = "0" Or Me.ListViewMerge.Items(i).SubItems(1).Text = "0.00" Then
                vCheckZero = vCheckZero + 1
            End If
        Next

        vCount = Me.ListViewMerge.Items.Count

        If vCheckZero = vCount Then
            MsgBox("��¡���Թ�������ըӹǹ����价ӡ�þѡ��� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
            If Me.ListViewMerge.Items.Count > 0 Then
                Me.ListViewMerge.Focus()
                Me.ListViewMerge.Items(0).Selected = True
                Me.ListViewMerge.Items(0).Focused = True
            End If
            Exit Sub
        End If

        If vCount = vCheck Then

            Dim vPrice As Double
            Dim vDiscount As Double
            Dim vAmount As Double

            Me.PNChecker.Enabled = False
            vARCode = Me.ListViewMerge.Items(0).SubItems(14).Text
            Me.TBHoldingAR.Text = vARCode
            Me.LBLHoldingAmount.Text = Me.LBLCheckOutAmount.Text

            For i = 0 To Me.ListViewMerge.Items.Count - 1

                vQTY = Me.ListViewMerge.Items(i).SubItems(1).Text

                If vQTY > 0 Then
                    vIndex = Me.ListViewHolding.Items.Count + 1
                    vPrice = Me.ListViewMerge.Items(i).SubItems(6).Text
                    vDiscount = Me.ListViewMerge.Items(i).SubItems(12).Text
                    vAmount = Me.ListViewMerge.Items(i).SubItems(7).Text

                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(2).Text)
                    listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(4).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(5).Text)
                    listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                    listItem.SubItems.Add(Format(vDiscount, "##,##0.00"))
                    listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(8).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(9).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(10).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(11).Text)
                    If Me.ListViewMerge.Items(i).SubItems(13).Text <> "" Then
                        listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(13).Text)
                    Else
                        listItem.SubItems.Add(vARCode)
                    End If
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(14).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(15).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(16).Text)
                    Me.ListViewHolding.Items.Add(listItem)
                End If
            Next

            Me.PNHolding.Visible = True
            Me.PNHolding.BringToFront()
            Me.BTNGenBill.Visible = True
            Me.BTNPrintHoldBill.Visible = False
            Me.TBHoldARName.Text = vARCode
            Me.Cash03.Checked = True
            Me.Cash03.Focus()
        Else
            MsgBox("��¡���Թ��� �ѧ��Ǩ�ͺ���ú ��سҵ�Ǩ�ͺ��¡���Թ��Ҵ���", MsgBoxStyle.Critical, "Send Error Message")

            If Me.ListViewMerge.Items.Count > 0 Then
                Me.ListViewMerge.Focus()
                Me.ListViewMerge.Items(0).Selected = True
                Me.ListViewMerge.Items(0).Focused = True
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub
    Private Sub TBSearchCheckOut_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchCheckOut.KeyDown
        Dim i As Integer
        Dim vDocno As String

        Dim vHeader As String
        Dim vNumber As Integer
        Dim vDocNumber As String

        Dim vSearch As String
        Dim vCountItem As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.TBSearchCheckOut.Text <> "" Then
                vSearch = Me.TBSearchCheckOut.Text

                vQuery = "exec dbo.usp_np_searchnewdocno 30"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                If ds.Tables(0).Rows.Count > 0 Then
                    vHeader = Trim(ds.Tables(0).Rows(0)("header").ToString)
                    vNumber = ds.Tables(0).Rows(0)("autonumber").ToString
                    vDocNumber = Trim(ds.Tables(0).Rows(0)("docnumber").ToString)
                End If

                vDocno = Trim(vDocNumber & vHeader & "-" & Format(vNumber, "0000"))

                vQuery = "exec dbo.USP_NP_InsertQueDriveInMergeTempCalc '" & vSearch & "','" & vDocno & "'"
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)
                If ds1.Tables(0).Rows.Count > 0 Then
                    vCountItem = Trim(ds1.Tables(0).Rows(0)("vcount").ToString)
                End If

                If vCountItem = 0 Then
                    MsgBox("�������¡���Թ��� ����ͧ��� CheckOut", MsgBoxStyle.Critical, "Send Information Message")
                    Me.TBSearchCheckOut.Focus()
                    Me.TBSearchCheckOut.SelectAll()
                    Exit Sub
                End If

                Me.ListViewMerge.Enabled = True

                vQuery = "exec dbo.usp_np_updatenewdocno 30"
                Dim vService2 As New WebReference.WebServiceCalc
                Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)


                Dim vMergeDocNo As String
                Dim vMergeDocDate1 As Date
                Dim vMergeDocDate As String
                Dim vMergeNetAmount As Double
                Dim m As Integer
                Dim vMergeQTY As Double
                Dim vMergePrice As Double
                Dim vMergeAmount As Double
                Dim vMergeItemCode As String
                Dim vMergeItemName As String
                Dim vMergeItemBar As String
                Dim vMergeUnitCode As String
                Dim vMergeWHCode As String
                Dim vMergeShelfCode As String
                Dim vMergeDriveIn As String
                Dim vMergeDiscount As Double
                Dim vMergeCarLicense As String
                Dim vMergeAR As String
                Dim vMergeSale As String


                vQuery = "exec dbo.USP_NP_CalcDriveInMergeTemp1 '" & vDocno & "'"
                Dim vService3 As New WebReference.WebServiceCalc
                Dim ds3 As DataSet = vService3.vGetQueryAnlyzer(vQuery)
                If ds3.Tables(0).Rows.Count > 0 Then
                    vMergeDocNo = Trim(ds3.Tables(0).Rows(0)("docno").ToString)
                    vMergeDocDate1 = Trim(ds3.Tables(0).Rows(0)("docdate").ToString)
                    vMergeDocDate = vMergeDocDate1.Day & "/" & vMergeDocDate1.Month & "/" & vMergeDocDate1.Year
                    vMergeNetAmount = Trim(ds3.Tables(0).Rows(0)("netamount").ToString)

                    Me.ListViewMerge.Visible = True
                    Me.ListViewMerge.Items.Clear()
                    For i = 0 To ds3.Tables(0).Rows.Count - 1

                        m = i + 1
                        vMergeItemCode = ds3.Tables(0).Rows(i)("itemcode").ToString
                        vMergeItemName = ds3.Tables(0).Rows(i)("itemname").ToString
                        vMergeQTY = ds3.Tables(0).Rows(i)("qty").ToString
                        vMergePrice = ds3.Tables(0).Rows(i)("price").ToString
                        vMergeAmount = ds3.Tables(0).Rows(i)("amount").ToString
                        vMergeUnitCode = ds3.Tables(0).Rows(i)("unitcode").ToString
                        vMergeItemBar = ds3.Tables(0).Rows(i)("barcode").ToString
                        vMergeDriveIn = ds3.Tables(0).Rows(i)("refno").ToString
                        vMergeDiscount = ds3.Tables(0).Rows(i)("discountamount").ToString
                        vMergeWHCode = ds3.Tables(0).Rows(i)("whcode").ToString
                        vMergeShelfCode = ds3.Tables(0).Rows(i)("shelfcode").ToString
                        vMergeCarLicense = ds3.Tables(0).Rows(i)("carlicense").ToString
                        vMergeAR = ds3.Tables(0).Rows(i)("arcode").ToString
                        vMergeSale = ds3.Tables(0).Rows(i)("salecode").ToString

                        Dim listItem As New ListViewItem(m)
                        listItem.SubItems.Add("")
                        listItem.SubItems.Add(vMergeItemName)
                        listItem.SubItems.Add(Format(vMergeQTY, "##,##0.00"))
                        listItem.SubItems.Add(vMergeUnitCode)
                        listItem.SubItems.Add(vMergeItemCode)
                        listItem.SubItems.Add(Format(vMergePrice, "##,##0.00"))
                        listItem.SubItems.Add(Format(vMergeAmount, "##,##0.00"))
                        listItem.SubItems.Add(vMergeWHCode)
                        listItem.SubItems.Add(vMergeShelfCode)
                        listItem.SubItems.Add(vMergeDriveIn)
                        listItem.SubItems.Add(vMergeItemBar)
                        listItem.SubItems.Add(vMergeDiscount)
                        listItem.SubItems.Add(vMergeCarLicense)
                        listItem.SubItems.Add(vMergeAR)
                        listItem.SubItems.Add(vMergeSale)
                        listItem.SubItems.Add(vDocno)
                        Me.ListViewMerge.Items.Add(listItem)
                    Next

                    Me.LBLNetAmount.Text = Format(vMergeNetAmount, "##,##0.00")

                End If

                If Me.ListViewMerge.Items.Count > 0 Then
                    Me.ListViewMerge.Focus()
                    Me.ListViewMerge.Items(0).Selected = True
                    Me.ListViewMerge.Items(0).Focused = True
                Else
                    Me.TBSearchCheckOut.Focus()
                End If

                Me.BTNCheckOut.Enabled = True
                Me.BTNGenCheckOut.Enabled = False
                Me.TBSearchCheckOut.Enabled = False

            End If
        End If


        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldingBill()
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If

        If e.KeyCode = 115 Then
            Call SearchHoldBill()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

    End Sub

    Private Sub ListViewCheckOut_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SelectQueItem()
        End If

        If e.KeyCode = 16 Then
            Call ItemSelectHoldBill()
        End If

        If e.KeyCode = 34 Then
            Call SearchHoldBill()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub vCalcCheckOutAmountLineItem()
        Dim i As Integer
        Dim vAmount As Double
        Dim vKeyQty As Double
        Dim vPrice As Double

        On Error GoTo ErrDescription

        If Me.ListViewMerge.Items.Count > 0 Then
            For i = 0 To Me.ListViewMerge.Items.Count - 1
                vKeyQty = Me.ListViewMerge.Items(i).SubItems(1).Text
                vPrice = Me.ListViewMerge.Items(i).SubItems(6).Text
                vAmount = vKeyQty * vPrice
                Me.ListViewMerge.Items(i).SubItems(6).Text = Format(vAmount, "##,##0.00")
            Next
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub vCalcCheckOutKeyQuanity()
        Dim i As Integer
        Dim vAmount As Double
        Dim vTotalAmount As Double
        Dim vPrice As Double
        Dim vKeyQTY As Double

        On Error GoTo ErrDescription

        If Me.ListViewMerge.Items.Count > 0 Then
            For i = 0 To Me.ListViewMerge.Items.Count - 1
                vPrice = Me.ListViewMerge.Items(i).SubItems(6).Text
                If Me.ListViewMerge.Items(i).SubItems(1).Text <> "" Then
                    vKeyQTY = Me.ListViewMerge.Items(i).SubItems(1).Text
                Else
                    vKeyQTY = 0
                End If
                vAmount = vKeyQTY * vPrice
                vTotalAmount = vTotalAmount + vAmount
            Next
            Me.LBLCheckOutAmount.Text = Format(vTotalAmount, "##,##0.00")
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCheckOut.Click
        Dim i As Integer
        Dim vCount As Integer
        Dim vCheck As Integer
        Dim vIndex As Integer
        Dim vARCode As String
        Dim vQTY As Double
        Dim vCheckZero As Integer

        On Error GoTo ErrDescription

        If Me.ListViewMerge.Items.Count = 0 Then
            MsgBox("�������¡���Թ���㹵��ҧ ���зӡ�þѡ��� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

        MsgBox("���駵��� ��顴���� �����+�����Ţ 9 ���ͷӡ�����͡�ش�ѡ���", MsgBoxStyle.Information, "Send Information Message")

        For i = 0 To Me.ListViewMerge.Items.Count - 1
            If Me.ListViewMerge.Items(i).SubItems(1).Text <> "" And Me.ListViewMerge.Items(i).SubItems(1).Text <> "0" And Me.ListViewMerge.Items(i).SubItems(1).Text <> "0.00" Then
                vCheck = vCheck + 1
            End If
        Next

        For i = 0 To Me.ListViewMerge.Items.Count - 1

            If Me.ListViewMerge.Items(i).SubItems(1).Text = "0" Or Me.ListViewMerge.Items(i).SubItems(1).Text = "0.00" Then
                vCheckZero = vCheckZero + 1
            End If
        Next

        vCount = Me.ListViewMerge.Items.Count

        If vCheckZero = vCount Then
            MsgBox("��¡���Թ�������ըӹǹ����价ӡ�þѡ��� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
            If Me.ListViewMerge.Items.Count > 0 Then
                Me.ListViewMerge.Focus()
                Me.ListViewMerge.Items(0).Selected = True
                Me.ListViewMerge.Items(0).Focused = True
            End If
            Exit Sub
        End If

        If vCount = vCheck Then

            Dim vPrice As Double
            Dim vDiscount As Double
            Dim vAmount As Double

            Me.PNChecker.Enabled = False
            vARCode = Me.ListViewMerge.Items(0).SubItems(14).Text
            Me.TBHoldingAR.Text = vARCode
            Me.LBLHoldingAmount.Text = Me.LBLCheckOutAmount.Text

            For i = 0 To Me.ListViewMerge.Items.Count - 1

                vQTY = Me.ListViewMerge.Items(i).SubItems(1).Text

                If vQTY > 0 Then
                    vIndex = Me.ListViewHolding.Items.Count + 1
                    vPrice = Me.ListViewMerge.Items(i).SubItems(6).Text
                    vDiscount = Me.ListViewMerge.Items(i).SubItems(12).Text
                    vAmount = vPrice * vQTY

                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(2).Text)
                    listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(4).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(5).Text)
                    listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                    listItem.SubItems.Add(Format(vDiscount, "##,##0.00"))
                    listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(8).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(9).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(10).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(11).Text)
                    If Me.ListViewMerge.Items(i).SubItems(13).Text <> "" Then
                        listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(13).Text)
                    Else
                        listItem.SubItems.Add(vARCode)
                    End If
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(14).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(15).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(i).SubItems(16).Text)
                    Me.ListViewHolding.Items.Add(listItem)
                End If
            Next

            Me.PNHolding.Visible = True
            Me.PNHolding.BringToFront()
            Me.BTNGenBill.Visible = True
            Me.BTNPrintHoldBill.Visible = False
            Me.TBHoldARName.Text = vARCode
            Me.Cash03.Checked = True
            Me.Cash03.Focus()
        Else
            MsgBox("��¡���Թ��� �ѧ��Ǩ�ͺ���ú ��سҵ�Ǩ�ͺ��¡���Թ��Ҵ���", MsgBoxStyle.Critical, "Send Error Message")

            If Me.ListViewMerge.Items.Count > 0 Then
                Me.ListViewMerge.Focus()
                Me.ListViewMerge.Items(0).Selected = True
                Me.ListViewMerge.Items(0).Focused = True
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Public Sub ItemSelectHoldBill()
        Dim i As Integer
        Dim vCount As Integer
        Dim vCheck As Integer
        Dim vIndex As Integer
        Dim vARCode As String

        Dim n As Integer
        Dim vItemCode As String
        Dim vQTY As Double
        Dim vUnitCode As String
        Dim vBarCode As String
        Dim vMergeNo As String
        Dim vDocNo As String
        Dim vCashierCode As String

        On Error GoTo ErrDescription

        If Me.ListViewMerge.Items.Count = 0 Then
            MsgBox("�������¡���Թ���㹵��ҧ ���зӡ�þѡ��� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

        For i = 0 To Me.ListViewMerge.Items.Count - 1
            If Me.ListViewMerge.Items(i).SubItems(1).Text <> "" Then
                vCheck = vCheck + 1
            End If
        Next

        vCount = Me.ListViewMerge.Items.Count
        If vCount = vCheck Then

            For n = 0 To Me.ListViewMerge.Items.Count - 1
                vItemCode = Me.ListViewMerge.Items(n).SubItems(5).Text
                vQTY = Me.ListViewMerge.Items(n).SubItems(1).Text
                vUnitCode = Me.ListViewMerge.Items(n).SubItems(4).Text
                vBarCode = Me.ListViewMerge.Items(n).SubItems(11).Text
                vMergeNo = Me.ListViewMerge.Items(n).SubItems(16).Text
                vDocNo = "HoldNo"
                vCashierCode = ""

                vQuery = "exec dbo.USP_NP_UpdateDriveInMergeTempConfirm1 '" & vMergeNo & "','" & vItemCode & "','" & vBarCode & "'," & vQTY & ",'" & vDocNo & "' "
                Dim vService4 As New WebReference.WebServiceCalc
                Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)

                vQuery = "exec dbo.USP_NP_UpdateHoldBillQtyQue1 '" & vMergeNo & "','" & vDocNo & "','" & vCashierCode & "','" & vItemCode & "','" & vBarCode & "','" & vUnitCode & "'," & vQTY & " "
                Dim vService5 As New WebReference.WebServiceCalc
                Dim ds5 As Integer = vService5.vExecuteQuery(vQuery)
            Next n

            vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto1 9,'" & vMergeNo & "','','" & vUserName & "'"
            Dim vService6 As New WebReference.WebServiceCalc
            Dim ds6 As Integer = vService6.vExecuteQuery(vQuery)

            Me.ListViewHolding.Items.Clear()
            Me.TBHoldARName.Text = ""
            Me.LBLHoldingAmount.Text = ""
            Me.PNHolding.Visible = False

            Me.ListViewMerge.Items.Clear()
            Me.TBSearchCheckOut.Text = ""
            Me.LBLNetAmount.Text = ""
            Me.LBLCheckOutAmount.Text = ""
            Me.TBMergeID.Text = ""
            Me.BTNCheckOut.Enabled = False
            Me.BTNGenCheckOut.Enabled = False
            Me.TBSearchCheckOut.Focus()
            Me.PNChecker.Enabled = True
        Else
            MsgBox("��¡���Թ��� �ѧ��Ǩ�ͺ���ú ��سҵ�Ǩ�ͺ��¡���Թ��Ҵ���", MsgBoxStyle.Critical, "Send Error Message")

            If Me.ListViewMerge.Items.Count > 0 Then
                Me.ListViewMerge.Focus()
                Me.ListViewMerge.Items(0).Selected = True
                Me.ListViewMerge.Items(0).Focused = True
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BeforeSave()
        On Error Resume Next

        Me.TBHoldARName.Enabled = False
        Me.Cash01.Enabled = False
        Me.Cash02.Enabled = False
        Me.Cash03.Enabled = False
        Me.ListViewHolding.Enabled = False
        Me.BTNGenBill.Enabled = False
        Me.BTNHoldingClose.Enabled = False
        Me.BTNPrintHoldBill.Enabled = False
    End Sub

    Private Sub AfterSave()
        On Error Resume Next

        Me.TBHoldARName.Enabled = True
        Me.Cash01.Enabled = True
        Me.Cash02.Enabled = True
        Me.Cash03.Enabled = True
        Me.ListViewHolding.Enabled = True
        Me.BTNGenBill.Enabled = True
        Me.BTNHoldingClose.Enabled = True
        Me.BTNPrintHoldBill.Enabled = True
    End Sub

    Private Sub BTNGenBill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNGenBill.Click
        Dim n As Integer
        Dim vDocNo As String
        Dim vDocdate As String
        Dim vExpireCredit As Integer
        Dim vARCode As String
        Dim vCashierCode As String
        Dim vMachineNo As String
        Dim vMachineCode As String
        Dim vSaleCode As String
        Dim vTaxRate As Double
        Dim vSumOfItemAmount As Double
        Dim vAfterDiscount As Double
        Dim vBeforeTaxAmount As Double
        Dim vTaxAmount As Double
        Dim vTotalAmount As Double
        Dim vNetDebtAmount As Double
        Dim vCreatorCode As String
        Dim vSHIFTCODE As String
        Dim vMydescription As String
        Dim vCarlicense As String

        Dim vMaxNo As Integer
        Dim vHeader As String

        Dim vItemCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vDiscountAmount As Double
        Dim vAmount As Double
        Dim vNetAmount As Double
        Dim vUnitCode As String
        Dim vStockType As Integer
        Dim vLineNumber As Integer
        Dim vBarCode As String
        Dim vPosStatus As Integer
        Dim vSORefNo As String
        Dim vMergeNo As String
        Dim vDriveInNo As String

        On Error GoTo ErrDescription

        If Me.ListViewHolding.Items.Count > 0 Then

            MsgBox("���駵��� ��顴���� �����+�����Ţ 9 ���ͷӡ�úѹ�֡㺾ѡ�����зӡ�� CheckOut", MsgBoxStyle.Information, "Send Information Message")

            If vIsOpen = 0 Then
                Call BeforeSave()
                If Me.Cash01.Checked = True Then
                    vMachineNo = "21"
                ElseIf Me.Cash02.Checked = True Then
                    vMachineNo = "22"
                ElseIf Me.Cash03.Checked = True Then
                    vMachineNo = "23"
                End If

                vDocdate = Now.Day & "/" & Now.Month & "/" & Now.Year
                vQuery = "exec dbo.USP_NP_CheckDocDate"
                Dim vService7 As New WebReference.WebServiceCalc
                Dim ds7 As DataSet = vService7.vGetQueryAnlyzer(vQuery)
                If ds7.Tables(0).Rows.Count > 0 Then
                    vDocdate = ds7.Tables(0).Rows(0)("vdocdate").ToString
                End If

                'vQuery = "exec dbo.usp_np_getmaxnoholdingbill1 '" & vMachineNo & "','" & vDocdate & "' "
                'Dim vService As New WebReference.WebServiceCalc
                'Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
                'If ds.Tables(0).Rows.Count > 0 Then
                '    vMaxNo = ds.Tables(0).Rows(0)("maxnumber").ToString
                '    vHeader = ds.Tables(0).Rows(0)("header").ToString
                'End If

                vQuery = "exec dbo.usp_np_getmaxnoholdbilllog '" & vMachineNo & "','" & vDocdate & "' "
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
                If ds.Tables(0).Rows.Count > 0 Then
                    vMaxNo = ds.Tables(0).Rows(0)("maxnumber").ToString
                    vHeader = ds.Tables(0).Rows(0)("header").ToString
                End If


                vDocNo = vHeader + "-" + Format(vMaxNo, "0000")

                vARCode = Me.TBHoldingAR.Text
                If vARCode = "1" Then
                    vARCode = "999999"
                End If
                vExpireCredit = 1

                vQuery = "select top 1 cashiercode,machinecode,shiftcode from dbo.bcarinvoice where docno like '%'+'" & vHeader & "'+'%'  and iscancel = 0 order by createdatetime desc"
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)
                If ds1.Tables(0).Rows.Count > 0 Then
                    vCashierCode = ds1.Tables(0).Rows(0)("cashiercode").ToString
                    vMachineCode = ds1.Tables(0).Rows(0)("machinecode").ToString
                    vSHIFTCODE = ds1.Tables(0).Rows(0)("shiftcode").ToString
                End If


                vSaleCode = Me.ListViewHolding.Items(0).SubItems(14).Text


                vTaxRate = 7
                If Me.LBLHoldingAmount.Text <> "" Then
                    vSumOfItemAmount = Me.LBLHoldingAmount.Text
                Else
                    vSumOfItemAmount = 0
                End If

                vAfterDiscount = vSumOfItemAmount
                vBeforeTaxAmount = ((vSumOfItemAmount * 100) / 107)
                vTaxAmount = vSumOfItemAmount - ((vSumOfItemAmount * 100) / 107)
                vNetDebtAmount = vSumOfItemAmount
                vTotalAmount = vSumOfItemAmount
                vCreatorCode = vUserID
                vMydescription = Me.ListViewHolding.Items(0).SubItems(13).Text
                vCarlicense = Me.ListViewHolding.Items(0).SubItems(12).Text

                'vQuery = "exec dbo.USP_NP_InsertHoldingBillDriveIn1 '" & vDocNo & "','" & vDocdate & "'," & vExpireCredit & ",'" & vARCode & "','" & vCashierCode & "','" & vMachineNo & "','" & vMachineCode & "','" & vSaleCode & "'," & vTaxRate & "," & vSumOfItemAmount & "," & vAfterDiscount & "," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalAmount & "," & vNetDebtAmount & ",'" & vCreatorCode & "','" & vSHIFTCODE & "','" & vMydescription & "','" & vCarlicense & "' "
                'Dim vService2 As New WebReference.WebServiceCalc
                'Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                vQuery = "exec dbo.USP_NP_InsertDriveInHoldBillMaster '" & vDocNo & "','" & vDocdate & "','" & vMachineNo & "'," & vNetDebtAmount & ",'" & vCreatorCode & "'"
                Dim vService8 As New WebReference.WebServiceCalc
                Dim ds8 As Integer = vService8.vExecuteQuery(vQuery)

                For n = 0 To Me.ListViewHolding.Items.Count - 1
                    vItemCode = Me.ListViewHolding.Items(n).SubItems(4).Text
                    vWHCode = Me.ListViewHolding.Items(n).SubItems(8).Text
                    vShelfCode = Me.ListViewHolding.Items(n).SubItems(9).Text
                    If Me.ListViewHolding.Items(n).SubItems(2).Text <> "" Then
                        vQTY = Me.ListViewHolding.Items(n).SubItems(2).Text
                    Else
                        vQTY = 0
                    End If
                    If Me.ListViewHolding.Items(n).SubItems(5).Text <> "" Then
                        vPrice = Me.ListViewHolding.Items(n).SubItems(5).Text
                    Else
                        vPrice = 0
                    End If
                    vDiscountAmount = Me.ListViewHolding.Items(n).SubItems(6).Text
                    vAmount = vQTY * vPrice 'Me.ListViewHolding.Items(n).SubItems(7).Text
                    vNetAmount = vQTY * vPrice 'Me.ListViewHolding.Items(n).SubItems(7).Text
                    vUnitCode = Me.ListViewHolding.Items(n).SubItems(3).Text
                    vStockType = 0
                    vLineNumber = n
                    vDriveInNo = Me.ListViewHolding.Items(n).SubItems(10).Text
                    vBarCode = Me.ListViewHolding.Items(n).SubItems(11).Text
                    vPosStatus = 1
                    vSORefNo = Me.ListViewHolding.Items(n).SubItems(12).Text
                    vMergeNo = Me.ListViewHolding.Items(n).SubItems(15).Text
                    vSaleCode = Me.ListViewHolding.Items(n).SubItems(14).Text

                    'vQuery = "exec dbo.USP_NP_InsertHoldingBillDriveInSub1 '" & vDocNo & "','" & vDocdate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "'," & vQTY & "," & vPrice & "," & vDiscountAmount & "," & vNetAmount & "," & vNetAmount & ",'" & vUnitCode & "'," & vStockType & "," & vLineNumber & ",'" & vBarCode & "','" & vCashierCode & "'," & vPosStatus & ",'" & vSORefNo & "','" & vDriveInNo & "','" & vSaleCode & "'"
                    'Dim vService3 As New WebReference.WebServiceCalc
                    'Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

                    vQuery = "exec dbo.USP_NP_UpdateDriveInMergeTempConfirm1 '" & vMergeNo & "','" & vItemCode & "','" & vBarCode & "'," & vQTY & ",'" & vDocNo & "' "
                    Dim vService4 As New WebReference.WebServiceCalc
                    Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)

                    vQuery = "exec dbo.USP_NP_UpdateHoldBillQtyQue1 '" & vMergeNo & "','" & vDocNo & "','" & vCashierCode & "','" & vItemCode & "','" & vBarCode & "','" & vUnitCode & "'," & vQTY & " "
                    Dim vService5 As New WebReference.WebServiceCalc
                    Dim ds5 As Integer = vService5.vExecuteQuery(vQuery)

                    vQuery = "exec dbo.USP_NP_InsertDriveInHoldBillSub '" & vDocNo & "','" & vDocdate & "','" & vItemCode & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & "," & vNetAmount & "," & vLineNumber & " "
                    Dim vService9 As New WebReference.WebServiceCalc
                    Dim ds9 As Integer = vService9.vExecuteQuery(vQuery)

                Next n

                vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto1 9,'" & vDocNo & "','','" & vUserName & "'"
                Dim vService6 As New WebReference.WebServiceCalc
                Dim ds6 As Integer = vService6.vExecuteQuery(vQuery)

                MsgBox("�ѹ�֡㺾ѡ������º�������� ���Ţ����͡��� " & vDocNo & " ", MsgBoxStyle.Information, "Send Error Message")

                Call AfterSave()
                Me.ListViewMerge.Enabled = False
                Me.TBSearchCheckOut.Enabled = True
                Me.ListViewHolding.Items.Clear()
                Me.TBHoldARName.Text = ""
                Me.LBLHoldingAmount.Text = ""
                Me.PNHolding.Visible = False
                Me.BTNPrintHoldBill.Visible = False
                Me.TBHoldNo.Text = ""

                Me.ListViewMerge.Items.Clear()
                Me.TBSearchCheckOut.Text = ""
                Me.LBLNetAmount.Text = ""
                Me.LBLCheckOutAmount.Text = ""
                Me.TBMergeID.Text = ""
                Me.BTNCheckOut.Enabled = False
                Me.BTNGenCheckOut.Enabled = False
                Me.PNChecker.Enabled = True
                Me.TBSearchCheckOut.Focus()
            Else
                MsgBox("�͡��÷��ѹ�֡������������ö�ѹ�֡��䢢��������ա ", MsgBoxStyle.Critical, "Send Error Message")
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Public Sub GenHoldBill()
        Dim n As Integer
        Dim vDocNo As String
        Dim vDocdate As String
        Dim vExpireCredit As Integer
        Dim vARCode As String
        Dim vCashierCode As String
        Dim vMachineNo As String
        Dim vMachineCode As String
        Dim vSaleCode As String
        Dim vTaxRate As Double
        Dim vSumOfItemAmount As Double
        Dim vAfterDiscount As Double
        Dim vBeforeTaxAmount As Double
        Dim vTaxAmount As Double
        Dim vTotalAmount As Double
        Dim vNetDebtAmount As Double
        Dim vCreatorCode As String
        Dim vSHIFTCODE As String
        Dim vMydescription As String
        Dim vCarLicense As String

        Dim vMaxNo As Integer
        Dim vHeader As String

        Dim vItemCode As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vDiscountAmount As Double
        Dim vAmount As Double
        Dim vNetAmount As Double
        Dim vUnitCode As String
        Dim vStockType As Integer
        Dim vLineNumber As Integer
        Dim vBarCode As String
        Dim vPosStatus As Integer
        Dim vSORefNo As String
        Dim vMergeNo As String
        Dim vDriveInNo As String

        On Error GoTo ErrDescription

        If Me.ListViewHolding.Items.Count > 0 Then
            If vIsOpen = 0 Then
                Call BeforeSave()
                If Me.Cash01.Checked = True Then
                    vMachineNo = "21"
                ElseIf Me.Cash02.Checked = True Then
                    vMachineNo = "22"
                ElseIf Me.Cash03.Checked = True Then
                    vMachineNo = "23"
                End If

                vQuery = "exec dbo.USP_NP_CheckDocDate"
                Dim vService7 As New WebReference.WebServiceCalc
                Dim ds7 As DataSet = vService7.vGetQueryAnlyzer(vQuery)
                If ds7.Tables(0).Rows.Count > 0 Then
                    vDocdate = ds7.Tables(0).Rows(0)("vdocdate").ToString
                End If

                'vDocdate = Now.Day & "/" & Now.Month & "/" & Now.Year

                'vQuery = "exec dbo.usp_np_getmaxnoholdingbill1 '" & vMachineNo & "','" & vDocdate & "' "
                'Dim vService As New WebReference.WebServiceCalc
                'Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
                'If ds.Tables(0).Rows.Count > 0 Then
                '    vMaxNo = ds.Tables(0).Rows(0)("maxnumber").ToString
                '    vHeader = ds.Tables(0).Rows(0)("header").ToString
                'End If

                vQuery = "exec dbo.usp_np_getmaxnoholdbilllog '" & vMachineNo & "','" & vDocdate & "' "
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
                If ds.Tables(0).Rows.Count > 0 Then
                    vMaxNo = ds.Tables(0).Rows(0)("maxnumber").ToString
                    vHeader = ds.Tables(0).Rows(0)("header").ToString
                End If

                vDocNo = vHeader + "-" + Format(vMaxNo, "0000")

                vARCode = Me.TBHoldingAR.Text
                If vARCode = "1" Then
                    vARCode = "999999"
                End If
                vExpireCredit = 1

                vQuery = "select top 1 cashiercode,machinecode,shiftcode from dbo.bcarinvoice where docno like '%'+'" & vHeader & "'+'%'  and iscancel = 0 order by createdatetime desc"
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)
                If ds1.Tables(0).Rows.Count > 0 Then
                    vCashierCode = ds1.Tables(0).Rows(0)("cashiercode").ToString
                    vMachineCode = ds1.Tables(0).Rows(0)("machinecode").ToString
                    vSHIFTCODE = ds1.Tables(0).Rows(0)("shiftcode").ToString
                End If


                vSaleCode = Me.ListViewHolding.Items(0).SubItems(14).Text


                vTaxRate = 7
                If Me.LBLHoldingAmount.Text <> "" Then
                    vSumOfItemAmount = Me.LBLHoldingAmount.Text
                Else
                    vSumOfItemAmount = 0
                End If

                vAfterDiscount = vSumOfItemAmount
                vBeforeTaxAmount = ((vSumOfItemAmount * 100) / 107)
                vTaxAmount = vSumOfItemAmount - ((vSumOfItemAmount * 100) / 107)
                vNetDebtAmount = vSumOfItemAmount
                vTotalAmount = vSumOfItemAmount
                vCreatorCode = vUserID
                vMydescription = Me.ListViewHolding.Items(0).SubItems(13).Text
                vCarlicense = Me.ListViewHolding.Items(0).SubItems(12).Text

                vQuery = "exec dbo.USP_NP_InsertHoldingBillDriveIn1 '" & vDocNo & "','" & vDocdate & "'," & vExpireCredit & ",'" & vARCode & "','" & vCashierCode & "','" & vMachineNo & "','" & vMachineCode & "','" & vSaleCode & "'," & vTaxRate & "," & vSumOfItemAmount & "," & vAfterDiscount & "," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalAmount & "," & vNetDebtAmount & ",'" & vCreatorCode & "','" & vSHIFTCODE & "','" & vMydescription & "','" & vCarLicense & "' "
                Dim vService2 As New WebReference.WebServiceCalc
                Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                vQuery = "exec dbo.USP_NP_InsertDriveInHoldBillMaster '" & vDocNo & "','" & vDocdate & "','" & vMachineNo & "'," & vNetDebtAmount & ",'" & vCreatorCode & "'"
                Dim vService8 As New WebReference.WebServiceCalc
                Dim ds8 As Integer = vService2.vExecuteQuery(vQuery)

                For n = 0 To Me.ListViewHolding.Items.Count - 1
                    vItemCode = Me.ListViewHolding.Items(n).SubItems(4).Text
                    vWHCode = Me.ListViewHolding.Items(n).SubItems(8).Text
                    vShelfCode = Me.ListViewHolding.Items(n).SubItems(9).Text
                    If Me.ListViewHolding.Items(n).SubItems(2).Text <> "" Then
                        vQTY = Me.ListViewHolding.Items(n).SubItems(2).Text
                    Else
                        vQTY = 0
                    End If
                    If Me.ListViewHolding.Items(n).SubItems(5).Text <> "" Then
                        vPrice = Me.ListViewHolding.Items(n).SubItems(5).Text
                    Else
                        vPrice = 0
                    End If
                    vDiscountAmount = Me.ListViewHolding.Items(n).SubItems(6).Text
                    vAmount = vQTY * vPrice 'Me.ListViewHolding.Items(n).SubItems(7).Text
                    vNetAmount = vQTY * vPrice 'Me.ListViewHolding.Items(n).SubItems(7).Text
                    vUnitCode = Me.ListViewHolding.Items(n).SubItems(3).Text
                    vStockType = 0
                    vLineNumber = n
                    vDriveInNo = Me.ListViewHolding.Items(n).SubItems(10).Text
                    vBarCode = Me.ListViewHolding.Items(n).SubItems(11).Text
                    vPosStatus = 1
                    vSORefNo = Me.ListViewHolding.Items(n).SubItems(12).Text
                    vMergeNo = Me.ListViewHolding.Items(n).SubItems(15).Text
                    vSaleCode = Me.ListViewHolding.Items(n).SubItems(14).Text

                    vQuery = "exec dbo.USP_NP_InsertHoldingBillDriveInSub1 '" & vDocNo & "','" & vDocdate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "'," & vQTY & "," & vPrice & "," & vDiscountAmount & "," & vAmount & "," & vNetAmount & ",'" & vUnitCode & "'," & vStockType & "," & vLineNumber & ",'" & vBarCode & "','" & vCashierCode & "'," & vPosStatus & ",'" & vSORefNo & "','" & vDriveInNo & "','" & vSaleCode & "'"
                    Dim vService3 As New WebReference.WebServiceCalc
                    Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

                    vQuery = "exec dbo.USP_NP_UpdateDriveInMergeTempConfirm1 '" & vMergeNo & "','" & vItemCode & "','" & vBarCode & "'," & vQTY & ",'" & vDocNo & "' "
                    Dim vService4 As New WebReference.WebServiceCalc
                    Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)

                    vQuery = "exec dbo.USP_NP_UpdateHoldBillQtyQue1 '" & vMergeNo & "','" & vDocNo & "','" & vCashierCode & "','" & vItemCode & "','" & vBarCode & "','" & vUnitCode & "'," & vQTY & " "
                    Dim vService5 As New WebReference.WebServiceCalc
                    Dim ds5 As Integer = vService5.vExecuteQuery(vQuery)

                    vQuery = "exec dbo.USP_NP_InsertDriveInHoldBillSub '" & vDocNo & "','" & vDocdate & "','" & vItemCode & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & "," & vNetAmount & "," & vLineNumber & " "
                    Dim vService9 As New WebReference.WebServiceCalc
                    Dim ds9 As Integer = vService2.vExecuteQuery(vQuery)

                Next n

                vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto1 9,'" & vDocNo & "','','" & vUserName & "'"
                Dim vService6 As New WebReference.WebServiceCalc
                Dim ds6 As Integer = vService6.vExecuteQuery(vQuery)

                MsgBox("�ѹ�֡㺾ѡ������º�������� ���Ţ����͡��� " & vDocNo & " ", MsgBoxStyle.Information, "Send Error Message")

                Call AfterSave()
                Me.ListViewMerge.Enabled = False
                Me.TBSearchCheckOut.Enabled = True
                Me.ListViewHolding.Items.Clear()
                Me.TBHoldARName.Text = ""
                Me.LBLHoldingAmount.Text = ""
                Me.PNHolding.Visible = False
                Me.BTNPrintHoldBill.Visible = False
                Me.TBHoldNo.Text = ""

                Me.ListViewMerge.Items.Clear()
                Me.TBSearchCheckOut.Text = ""
                Me.LBLNetAmount.Text = ""
                Me.LBLCheckOutAmount.Text = ""
                Me.TBMergeID.Text = ""
                Me.BTNCheckOut.Enabled = False
                Me.BTNGenCheckOut.Enabled = False
                Me.PNChecker.Enabled = True
                Me.TBSearchCheckOut.Focus()
            Else
                MsgBox("�͡��÷��ѹ�֡������������ö�ѹ�֡��䢢��������ա ", MsgBoxStyle.Critical, "Send Error Message")
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub


    Private Sub BTNClearCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearCheckOut.Click

        On Error Resume Next

        MsgBox("���駵��� ��顴���� �����+�����Ţ 7 ���ͷӡ����ҧ˹�Ҩ�", MsgBoxStyle.Information, "Send Information Message")

        Me.ListViewMerge.Items.Clear()
        Me.BTNGenCheckOut.Enabled = False
        Me.BTNCheckOut.Enabled = False
        Me.LBLCheckOutAmount.Text = ""
        Me.TBSearchCheckOut.Enabled = True
        Me.LBLNetAmount.Text = ""
        Me.TBSearchCheckOut.Enabled = True
        Me.ListViewMerge.Enabled = False
        Me.TBSearchCheckOut.Text = ""
        Me.TBSearchCheckOut.Focus()
    End Sub

    Public Sub ClearScreen()
        On Error Resume Next

        Me.ListViewMerge.Items.Clear()
        Me.BTNGenCheckOut.Enabled = False
        Me.BTNCheckOut.Enabled = False
        Me.LBLCheckOutAmount.Text = ""
        Me.LBLNetAmount.Text = ""
        Me.TBSearchCheckOut.Enabled = True
        Me.ListViewMerge.Enabled = False
        Me.TBSearchCheckOut.Text = ""
        Me.TBSearchCheckOut.Focus()
    End Sub

    Private Sub LBCloseAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next

        Me.TBSearchBarCode.Text = ""
        Me.PNAddItem.Visible = False
    End Sub

    Private Sub TBAddQTY_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBAddQTY.KeyDown
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

        Dim vAnswer As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewAddStockQTY.Items.Count > 0 And Me.TBAddItemCode.Text <> "" Then
                vCheckShelf = Me.TBAddDefShelf.Text
                vCheckUnit = Me.TBAddItemUnit.Text
                If Me.ListViewAddStockQTY.Items.Count > 0 Then
                    For v = 0 To Me.ListViewAddStockQTY.Items.Count - 1
                        vListShelf = Me.ListViewAddStockQTY.Items(v).Text
                        vListUnit = Me.ListViewAddStockQTY.Items(v).SubItems(2).Text
                        If vCheckShelf = vListShelf And vCheckUnit = vListUnit Then
                            vShelfQTY = Me.ListViewAddStockQTY.Items(v).SubItems(1).Text
                            vShelfUnit = Me.ListViewAddStockQTY.Items(v).SubItems(2).Text
                            GoTo Line1
                        End If
                    Next
                End If

Line1:
                vCheckExist = 0
                vBarCode = Me.TBAddItemBar.Text
                vItemCode = Me.TBAddItemCode.Text
                vItemName = Me.TBAddItemName.Text
                vWHCode = Me.TBAddDefWHCode.Text
                vShelfCode = Me.TBAddDefShelf.Text
                vUnitCode = Me.TBAddItemUnit.Text
                vRate = Me.TBAddItemRate.Text

                If Me.TBAddQTY.Text <> "" Then
                    vQTY = Me.TBAddQTY.Text
                End If

                If vShelfUnit <> vUnitCode Then
                    vTotalQTY = vShelfQTY / vRate
                    If vQTY > vTotalQTY Then
                        vAnswer = MsgBox("STOCK ���͢�� �ѧ�׹�ѹ��â�������������", MsgBoxStyle.YesNo, "Send Error Message ")
                        If vAnswer = 7 Then
                            Me.TBAddQTY.SelectAll()
                            Exit Sub
                        Else
                            GoTo NextStep1
                        End If
                    End If
                End If

                If vQTY > vShelfQTY And vShelfUnit = vUnitCode Then
                    vAnswer = MsgBox("STOCK ���͢�� �ѧ�׹�ѹ��â�������������", MsgBoxStyle.YesNo, "Send Error Message ")
                    If vAnswer = 7 Then
                        Me.TBAddQTY.SelectAll()
                        Exit Sub
                    Else
                        GoTo NextStep1
                    End If
                End If

NextStep1:
                If Me.TBAddPrice.Text <> "" Then
                    vPrice = Me.TBAddPrice.Text
                End If
                vAmount = vQTY * vPrice

                vIndex = Me.ListViewMerge.Items.Count + 1

                If vQTY = 0 Then
                    MsgBox("������кبӹǹ�ͧ�Թ��ҷ���ͧ��� ���͵�ͧ�кبӹǹ�Թ��ҷ���ͧ����ҡ���� 0", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If


                Dim n As Integer
                Dim vCheckItemCode As String
                Dim vEditQTY As Double
                Dim vEditPrice As Double
                Dim vItemAmount As Double
                Dim vOldQty As Double
                Dim vPickZone As String


                If Me.ListViewMerge.Items.Count > 0 Then
                    For n = 0 To Me.ListViewMerge.Items.Count - 1
                        vCheckItemCode = Me.ListViewMerge.Items(n).SubItems(5).Text

                        If vItemCode = vCheckItemCode Then
                            vEditPrice = Me.TBAddPrice.Text
                            vEditQTY = Me.TBAddQTY.Text
                            vItemAmount = vEditQTY * vEditPrice

                            vOldQty = Me.ListViewMerge.Items(n).SubItems(3).Text
                            vPickZone = Me.ListViewMerge.Items(n).SubItems(10).Text

                            If vEditQTY = vOldQty Then
                                If vPickZone = "01" Then
                                    Me.ListViewMerge.Items(n).ForeColor = Color.DarkBlue
                                ElseIf vPickZone = "02" Then
                                    Me.ListViewMerge.Items(n).ForeColor = Color.DarkGreen
                                ElseIf vPickZone = "03" Then
                                    Me.ListViewMerge.Items(n).ForeColor = Color.DarkOrange
                                ElseIf vPickZone = "04" Then
                                    Me.ListViewMerge.Items(n).ForeColor = Color.DarkMagenta
                                ElseIf vPickZone = "05" Then
                                    Me.ListViewMerge.Items(n).ForeColor = Color.Black
                                End If
                            Else
                                Me.ListViewMerge.Items(n).ForeColor = Color.Red
                            End If

                            Me.ListViewMerge.Items(n).SubItems(1).Text = Format(vEditQTY, "##,##0.00")
                            vCheckExist = 1
                            GoTo line2
                        End If
                    Next
                End If

line2:

                If vCheckExist = 0 Then
                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                    listItem.SubItems.Add(vItemName)
                    listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                    listItem.SubItems.Add(vUnitCode)
                    listItem.SubItems.Add(vItemCode)
                    listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                    listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                    listItem.SubItems.Add(vWHCode)
                    listItem.SubItems.Add(vShelfCode)
                    listItem.SubItems.Add("")
                    listItem.SubItems.Add(vBarCode)
                    listItem.SubItems.Add(0)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(0).SubItems(13).Text)
                    listItem.SubItems.Add(Me.ListViewMerge.Items(0).SubItems(14).Text)
                    listItem.SubItems.Add("CheckerAdd")
                    listItem.SubItems.Add("CheckerNo")

                    Me.ListViewMerge.Items.Add(listItem)
                End If

                Call vCalcCheckOutKeyQuanity()

                If vQTY >= 10000 Then
                    MsgBox("�ӹǹ�Թ��ҷ�����͡� �Թ 10,000 ��سҵ�Ǩ�ͺ�������ա��", MsgBoxStyle.Information, "Send Error Message")
                End If

                Me.TBAddItemCode.Text = ""
                Me.TBAddItemBar.Text = ""
                Me.TBAddItemName.Text = ""
                Me.TBAddPrice.Text = ""
                Me.TBAddReserveQTY.Text = ""
                Me.TBAddItemUnit.Text = ""
                Me.TBAddDefWHCode.Text = ""
                Me.TBAddDefShelf.Text = ""
                Me.TBAddQTY.Text = ""
                Me.TBAddItemRate.Text = ""
                Me.ListViewAddStockQTY.Items.Clear()
                Me.PNAddItem.Visible = False
                Me.TBSearchBarCode.Text = ""
                Me.ListViewMerge.Focus()
                Me.ListViewMerge.Items(0).Selected = True
                Me.ListViewMerge.Items(0).Focused = True
            Else
                MsgBox("�������¡���Թ����������ö���� ��¡���Թ���ŧ�С�����", MsgBoxStyle.Critical, "Send Error Message")
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBSearchBarCode.Text = ""
            Me.PNAddItem.Visible = False
            Me.ListViewMerge.Focus()
            Me.ListViewMerge.Items(0).Selected = True
            Me.ListViewMerge.Items(0).Focused = True
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBSearchBarCode.Text = ""
            Me.TBAddItemCode.Text = ""
            Me.TBAddItemName.Text = ""
            Me.TBAddPrice.Text = ""
            Me.TBAddReserveQTY.Text = ""
            Me.TBAddItemUnit.Text = ""
            Me.TBAddDefWHCode.Text = ""
            Me.TBAddDefShelf.Text = ""
            Me.TBAddQTY.Text = ""
            Me.TBAddItemRate.Text = ""
            Me.TBAddItemBar.Text = ""
            Me.PNAddItem.Visible = False

            Me.ListViewMerge.Focus()
            Me.ListViewMerge.Items(0).Selected = True
            Me.ListViewMerge.Items(0).Focused = True
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBSearchBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchBarCode.KeyDown
        Dim vBarCode As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vPrice As Double
        Dim vRate As Integer
        Dim vUnitCode As String
        Dim i As Integer
        Dim vStore As String
        Dim vStkUnit As String
        Dim vStkQTY As Double
        Dim vReserveQTY As Double
        Dim vDefWHCode As String
        Dim vDefShelfCode As String

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.TBSearchBarCode.Text <> "" Then
                vBarCode = Me.TBSearchBarCode.Text
            Else
                Me.TBSearchBarCode.Focus()
                Me.TBSearchBarCode.SelectAll()
                Exit Sub
            End If

            Me.ListViewAddStockQTY.Items.Clear()

            vQuery = "exec dbo.USP_MB_SearchBarCode '" & vBarCode & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            If ds.Tables(0).Rows.Count > 0 Then
                vItemCode = ds.Tables(0).Rows(0)("itemcode").ToString
                vItemName = ds.Tables(0).Rows(0)("itemname").ToString
                vPrice = ds.Tables(0).Rows(0)("price").ToString
                vRate = ds.Tables(0).Rows(0)("rate").ToString
                vUnitCode = ds.Tables(0).Rows(0)("unitcode").ToString
                vReserveQTY = ds.Tables(0).Rows(0)("reserveqty").ToString
                vDefWHCode = ds.Tables(0).Rows(0)("defsalewhcode").ToString
                vDefShelfCode = ds.Tables(0).Rows(0)("defsaleshelf").ToString

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    vStore = ds.Tables(0).Rows(i)("shelfcode").ToString
                    vStkUnit = ds.Tables(0).Rows(i)("stkunitcode").ToString
                    vStkQTY = ds.Tables(0).Rows(i)("stock").ToString

                    Dim listItem As New ListViewItem(vStore)
                    listItem.SubItems.Add(Format(vStkQTY, "##,##0.00"))
                    listItem.SubItems.Add(vStkUnit)
                    Me.ListViewAddStockQTY.Items.Add(listItem)
                Next

Line1:
                Me.TBAddQTY.Focus()
                Me.TBAddQTY.SelectAll()
            Else
                MsgBox("����� �������Թ��� ��سҵ�Ǩ�ͺ", MsgBoxStyle.Critical, "Send Information Message")
                Me.TBSearchBarCode.Text = ""
                Me.TBSearchBarCode.Focus()
                Me.TBSearchBarCode.SelectAll()
                Exit Sub
            End If

            Me.TBAddItemCode.Text = vItemCode
            Me.TBAddItemName.Text = vItemName
            Me.TBAddPrice.Text = Format(vPrice, "##,##0.00")
            Me.TBAddItemRate.Text = Format(vRate, "##,##0.00")
            Me.TBAddReserveQTY.Text = Format(vReserveQTY, "##,##0.00")
            Me.TBAddItemUnit.Text = vUnitCode
            Me.TBAddDefWHCode.Text = vDefWHCode
            Me.TBAddDefShelf.Text = vDefShelfCode
            Me.TBAddItemBar.Text = vBarCode

        End If

        If e.KeyCode = Keys.Back Then
            Me.TBAddItemCode.Text = ""
            Me.TBAddItemName.Text = ""
            Me.TBAddPrice.Text = ""
            Me.TBAddReserveQTY.Text = ""
            Me.TBAddItemUnit.Text = ""
            Me.TBAddDefWHCode.Text = ""
            Me.TBAddDefShelf.Text = ""
            Me.TBAddQTY.Text = ""
            Me.TBAddItemRate.Text = ""
            Me.TBAddItemBar.Text = ""
            Me.ListViewAddStockQTY.Items.Clear()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBSearchBarCode.Text = ""
            Me.TBAddItemCode.Text = ""
            Me.TBAddItemName.Text = ""
            Me.TBAddPrice.Text = ""
            Me.TBAddReserveQTY.Text = ""
            Me.TBAddItemUnit.Text = ""
            Me.TBAddDefWHCode.Text = ""
            Me.TBAddDefShelf.Text = ""
            Me.TBAddQTY.Text = ""
            Me.TBAddItemRate.Text = ""
            Me.TBAddItemBar.Text = ""
            Me.PNAddItem.Visible = False

            Me.ListViewMerge.Focus()
            Me.ListViewMerge.Items(0).Selected = True
            Me.ListViewMerge.Items(0).Focused = True
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
                Me.PNChecker.Visible = True
                Me.PNChecker.BringToFront()

                vConnectZone = "05"
                vCheckTypeLogIn = "�ش������"
                Me.TBUserID.Text = vCheckLogIn
                vUserID = vCheckLogIn
                vSaleLogIn = vUserCode
                Me.TBSearchCheckOut.Focus()
                Me.TBSearchCheckOut.SelectAll()

            Else
                MsgBox("�������ö�����ҹ������� ��سҵ�Ǩ�ͺ����������ʼ�ҹ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBPassword.Visible = True
                Me.TBPassword.Text = ""
                Me.TBPassword.Focus()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub FrmCheckOut_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        Me.PNLogIn.Visible = True
        Me.PNLogIn.BringToFront()
        Me.TBUserCode.Focus()
    End Sub

    Private Sub ListViewAddStockQTY_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewAddStockQTY.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.TBSearchBarCode.Text = ""
            Me.PNAddItem.Visible = False
            Me.TBSearchCheckOut.Focus()
        End If
    End Sub

    Private Sub BTNSaveHold_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vDocdate As String
        Dim vPosNo As String
        Dim vMachineNo As String
        Dim vHeader As String

        Dim vExpireCredit As Integer
        Dim vArCode As String
        Dim vCashierCode As String
        Dim vMachineCode As String
        Dim vSaleCode As String
        Dim vTaxRate As Double
        Dim vSumOfItemAmount As Double
        Dim vAfterDiscount As Double
        Dim vBeforeTaxAmount As Double
        Dim vTaxAmount As Double
        Dim vTotalAmount As Double
        Dim vNetDebtAmount As Double
        Dim vCreatorCode As String
        Dim vSHIFTCODE As String

        On Error GoTo ErrDescription

        vExpireCredit = 1
        vArCode = "99999"

        vQuery = "select top 1 cashiercode,machinecode from dbo.bcarinvoice where docno like '%'+'" & vHeader & "'+'%'and iscancel = 0 order by createdatetime desc"
        Dim vService2 As New WebReference.WebServiceCalc
        Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)

        If ds2.Tables(0).Rows.Count > 0 Then
            vCashierCode = ds2.Tables(0).Rows(0)("cashiercode").ToString
            vMachineCode = ds2.Tables(0).Rows(0)("machinecode").ToString
        End If


        vSaleCode = ""
        vTaxRate = 7
        If Me.LBLCheckOutAmount.Text <> "" Then
            vSumOfItemAmount = Me.LBLCheckOutAmount.Text
        Else
            vSumOfItemAmount = 0
        End If
        vAfterDiscount = vSumOfItemAmount
        vBeforeTaxAmount = vSumOfItemAmount - ((vSumOfItemAmount * 7) / 100)
        vTaxAmount = ((vSumOfItemAmount * 7) / 100)
        vTotalAmount = vSumOfItemAmount
        vNetDebtAmount = vSumOfItemAmount
        vCreatorCode = ""
        vSHIFTCODE = "��ҧ�ѹ"

        vQuery = "exec dbo.USP_NP_InsertHoldingBillDriveIn1 '" & vPosNo & "','" & vDocdate & "'," & vExpireCredit & ",'" & vArCode & "','" & vCashierCode & "','" & vMachineNo & "','" & vMachineCode & "','" & vSaleCode & "'," & vTaxRate & "," & vTaxRate & "," & vAfterDiscount & "," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalAmount & "," & vNetDebtAmount & ",'" & vCreatorCode & "','" & vSHIFTCODE & "' "
        Dim vService3 As New WebReference.WebServiceCalc
        Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

    End Sub

    Private Sub RBCash1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim vDocdate As String
        Dim vPosNo As String

        Dim vMachineNo As String
        Dim vHeader As String

        Dim vExpireCredit As Integer
        Dim vArCode As String
        Dim vCashierCode As String
        Dim vMachineCode As String
        Dim vSaleCode As String
        Dim vTaxRate As Double
        Dim vSumOfItemAmount As Double
        Dim vAfterDiscount As Double
        Dim vBeforeTaxAmount As Double
        Dim vTaxAmount As Double
        Dim vTotalAmount As Double
        Dim vNetDebtAmount As Double
        Dim vCreatorCode As String
        Dim vSHIFTCODE As String
        Dim vMyDescription As String

        Dim vAnswer As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            vAnswer = MsgBox("�س��ͧ��� �ѡ��ŷ��ᤪ�����ش��� 1 ���������", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then

                vExpireCredit = 1
                vArCode = "99999"

                vQuery = "select top 1 cashiercode,machinecode from dbo.bcarinvoice where docno like '%'+'" & vHeader & "'+'%'and iscancel = 0 order by createdatetime desc"
                Dim vService2 As New WebReference.WebServiceCalc
                Dim ds2 As DataSet = vService2.vGetQueryAnlyzer(vQuery)

                If ds2.Tables(0).Rows.Count > 0 Then
                    vCashierCode = ds2.Tables(0).Rows(0)("cashiercode").ToString
                    vMachineCode = ds2.Tables(0).Rows(0)("machinecode").ToString
                End If

                vSaleCode = vSaleLogIn
                vTaxRate = 7
                If Me.LBLCheckOutAmount.Text <> "" Then
                    vSumOfItemAmount = Me.LBLCheckOutAmount.Text
                Else
                    vSumOfItemAmount = 0
                End If
                vAfterDiscount = vSumOfItemAmount
                vBeforeTaxAmount = vSumOfItemAmount - ((vSumOfItemAmount * 7) / 100)
                vTaxAmount = ((vSumOfItemAmount * 7) / 100)
                vTotalAmount = vSumOfItemAmount
                vNetDebtAmount = vSumOfItemAmount
                vCreatorCode = vUserID
                vSHIFTCODE = "��ҧ�ѹ"
                If Me.TBPosBill.Text = "" And Me.TBPosBill.Visible = False Then
                    vMyDescription = Me.TBSearchCheckOut.Text
                Else
                    vMyDescription = Me.TBRefDocNo.Text
                End If

                vQuery = "exec dbo.USP_NP_InsertHoldingBillDriveIn1 '" & vPosNo & "','" & vDocdate & "'," & vExpireCredit & ",'" & vArCode & "','" & vCashierCode & "','" & vMachineNo & "','" & vMachineCode & "','" & vSaleCode & "'," & vTaxRate & "," & vTaxRate & "," & vAfterDiscount & "," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalAmount & "," & vNetDebtAmount & ",'" & vCreatorCode & "','" & vSHIFTCODE & "','" & vMyDescription & "' "
                Dim vService3 As New WebReference.WebServiceCalc
                Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

                MsgBox("���Ţ���ѡ����Ţ��� " & vPosNo & "", MsgBoxStyle.Information, "Send Information Message")
                Me.TBPosBill.Text = ""
                Me.TBRefDocNo.Text = ""
                Me.LBLCheckOutAmount.Text = ""
                Me.LBLNetAmount.Text = ""
                Me.TBSearchCheckOut.Focus()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub RBCash2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
     
    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        Dim i As Integer
        Dim vDocno As String
        Dim vNetDebtAmount As Double
        Dim vRefNo As String

        On Error GoTo ErrDescription

        MsgBox("���駵��� ��顴���� �����+�����Ţ 2 ���ͷӡ�ä����Ţ���㺾ѡ���", MsgBoxStyle.Information, "Send Information Message")

        vQuery = "exec dbo.usp_np_SearchCheckOutHolding1 ''"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        Me.ListViewHold.Items.Clear()

        If ds.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds.Tables(0).Rows.Count - 1
                vDocno = ds.Tables(0).Rows(i)("docno").ToString
                vRefNo = ds.Tables(0).Rows(i)("mydescription").ToString
                vNetDebtAmount = ds.Tables(0).Rows(i)("netdebtamount").ToString

                Dim listItem As New ListViewItem(vDocno)
                listItem.SubItems.Add(vRefNo)
                listItem.SubItems.Add(Format(vNetDebtAmount, "##,##0.00"))
                Me.ListViewHold.Items.Add(listItem)
            Next

            Dim a As Integer

            For a = 0 To Me.ListViewHold.Items.Count - 1
                If a Mod 2 <> 0 Then
                    Me.ListViewHold.Items(a).BackColor = Color.Silver
                End If
            Next

            Me.PNSearchHold.Visible = True
            Me.PNSearchHold.BringToFront()
            If Me.ListViewHold.Items.Count > 0 Then
                Me.ListViewHold.Focus()
                Me.ListViewHold.Items(0).Selected = True
                Me.ListViewHold.Items(0).Focused = True
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Public Sub SearchHoldBill()
        Dim i As Integer
        Dim vDocno As String
        Dim vRefNo As String
        Dim vNetDebtAmount As Double

        On Error GoTo ErrDescription

        vQuery = "exec dbo.usp_np_SearchCheckOutHolding1 ''"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        Me.ListViewHold.Items.Clear()

        If ds.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds.Tables(0).Rows.Count - 1
                vDocno = ds.Tables(0).Rows(i)("docno").ToString
                vRefNo = ds.Tables(0).Rows(i)("mydescription").ToString
                vNetDebtAmount = ds.Tables(0).Rows(i)("netdebtamount").ToString

                Dim listItem As New ListViewItem(vDocno)
                listItem.SubItems.Add(vRefNo)
                listItem.SubItems.Add(Format(vNetDebtAmount, "##,##0.00"))
                Me.ListViewHold.Items.Add(listItem)
            Next

            Dim a As Integer

            For a = 0 To Me.ListViewHold.Items.Count - 1
                If a Mod 2 <> 0 Then
                    Me.ListViewHold.Items(a).BackColor = Color.Silver
                End If
            Next

            Me.PNSearchHold.Visible = True
            Me.PNSearchHold.BringToFront()
            If Me.ListViewHold.Items.Count > 0 Then
                Me.ListViewHold.Focus()
                Me.ListViewHold.Items(0).Selected = True
                Me.ListViewHold.Items(0).Focused = True
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub ListViewHold_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewHold.KeyDown
        Dim i As Integer
        Dim vDocNo As String
        Dim vDriveInNo As String
        Dim vMergeNo As String
        Dim n As Integer
        Dim vDocDate As String
        Dim vQueID As Integer
        Dim vPickZone As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPickQTY As Double
        Dim vConfirmQTY As Double
        Dim vUnitCode As String
        Dim vPrice As Double
        Dim vDiscountAmount As Double
        Dim vAmount As Double
        Dim vIndex As Integer
        Dim vLine As Integer
        Dim vBarcode As String
        Dim vLicense As String
        Dim vARCode As String
        Dim vSaleCode As String
        Dim vNetDebtAmount As Double

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewHold.Items.Count > 0 Then
                n = Me.ListViewHold.FocusedItem.Index
                vDocNo = Me.ListViewHold.Items(n).SubItems(0).Text

                vQuery = "exec dbo.usp_np_SearchHoldingDetails1 '" & vDocNo & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                vIndex = 0
                If ds.Tables(0).Rows.Count > 0 Then

                    Me.TBHoldNo.Text = ds.Tables(0).Rows(i)("docno").ToString
                    vNetDebtAmount = ds.Tables(0).Rows(i)("netdebtamount").ToString
                    Me.LBLHoldingAmount.Text = Format(vNetDebtAmount, "##,##0.00")

                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        vItemName = ds.Tables(0).Rows(i)("itemname").ToString
                        vQTY = ds.Tables(0).Rows(i)("qty").ToString
                        vUnitCode = ds.Tables(0).Rows(i)("unitcode").ToString
                        vItemCode = ds.Tables(0).Rows(i)("itemcode").ToString
                        vPrice = ds.Tables(0).Rows(i)("price").ToString
                        vDiscountAmount = ds.Tables(0).Rows(i)("price").ToString
                        vAmount = ds.Tables(0).Rows(i)("amount").ToString
                        vWHCode = ds.Tables(0).Rows(i)("whcode").ToString
                        vShelfCode = ds.Tables(0).Rows(i)("shelfcode").ToString
                        vDriveInNo = ds.Tables(0).Rows(i)("driveinrefno").ToString
                        vBarcode = ds.Tables(0).Rows(i)("barcode").ToString
                        vLicense = ds.Tables(0).Rows(i)("license").ToString
                        vARCode = ds.Tables(0).Rows(i)("arcode").ToString
                        vMergeNo = ds.Tables(0).Rows(i)("mergeno").ToString
                        vSaleCode = ds.Tables(0).Rows(i)("salecode").ToString

                        Me.TBHoldARName.Text = vARCode

                        vIndex = vIndex + 1
                        vLine = vIndex - 1
                        Dim listItem As New ListViewItem(vIndex)
                        listItem.SubItems.Add(vItemName)
                        listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                        listItem.SubItems.Add(vUnitCode)
                        listItem.SubItems.Add(vItemCode)
                        listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                        listItem.SubItems.Add(Format(vDiscountAmount, "##,##0.00"))
                        listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                        listItem.SubItems.Add(vWHCode)
                        listItem.SubItems.Add(vShelfCode)
                        listItem.SubItems.Add(vDriveInNo)
                        listItem.SubItems.Add(vBarcode)
                        listItem.SubItems.Add(vLicense)
                        listItem.SubItems.Add(vARCode)
                        listItem.SubItems.Add(vSaleCode)
                        listItem.SubItems.Add(vMergeNo)
                        Me.ListViewHolding.Items.Add(listItem)
                    Next

                    vIsOpen = 1

                    Call vCalcCheckOutKeyQuanity()

                    Me.PNHolding.Visible = True
                    Me.PNHolding.BringToFront()
                    Me.PNSearchHold.Visible = False
                    Me.BTNPrintHoldBill.Visible = True
                    Me.BTNGenBill.Visible = False

                    If ListViewHolding.Items.Count > 0 Then
                        Me.ListViewHolding.Focus()
                        Me.ListViewHolding.Items(0).Selected = True
                        Me.ListViewHolding.Items(0).Focused = True
                    End If

                End If
            End If
        End If

        If e.KeyCode = Keys.Up Then
            Me.TBHoldSearch.Focus()
            Me.TBHoldSearch.SelectAll()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchHold.Visible = False
            vIsOpen = 0
            If Me.TBSearchCheckOut.Enabled = True Then
                Me.TBSearchCheckOut.Focus()
            ElseIf Me.ListViewMerge.Enabled = True Then
                Me.ListViewMerge.Focus()
                Me.ListViewMerge.Items(0).Selected = True
                Me.ListViewMerge.Items(0).Focused = True
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBKeyQTY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
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

    Private Sub TBAddQTY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TBAddQTY.KeyPress, TBAddQTY.KeyPress, TBCheckQty.KeyPress
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

    Private Sub TBAddQTY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBAddQTY.TextChanged
        Dim vPrice As Double
        Dim vItemcode As String
        Dim vUnitCode As String
        Dim vQty As Double

        On Error GoTo ErrDescription

        vItemcode = Me.TBAddItemCode.Text
        vUnitCode = Me.TBAddItemUnit.Text
        If Me.TBAddQTY.Text <> "" Then
            vQty = Me.TBAddQTY.Text
        End If

        If vQty > 0 Then
            vQuery = "exec dbo.USP_NP_SearchItemPriceQty1 '" & vItemcode & "'," & vQty & ",'" & vUnitCode & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
            If ds.Tables(0).Rows.Count > 0 Then
                vPrice = ds.Tables(0).Rows(0)("saleprice1").ToString
            End If

            Me.TBAddPrice.Text = Format(vPrice, "##,##0.00")
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub


    Private Sub TBHoldSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBHoldSearch.KeyDown
        Dim i As Integer
        Dim vSearch As String
        Dim vDocno As String
        Dim vRefNo As String
        Dim vNetDebtAmount As Double

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.TBHoldSearch.Text = "" Then
                MsgBox("��ͧ��͡�ӷ���ͧ��ä鹴���", MsgBoxStyle.Critical, "Send Error Message")
                Exit Sub
            End If
            vSearch = Me.TBHoldSearch.Text
            vQuery = "exec dbo.usp_np_SearchCheckOutHolding1 '" & vSearch & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            Me.ListViewHold.Items.Clear()

            If ds.Tables(0).Rows.Count > 0 Then
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    vDocno = ds.Tables(0).Rows(i)("docno").ToString
                    vRefNo = ds.Tables(0).Rows(i)("mydescription").ToString
                    vNetDebtAmount = ds.Tables(0).Rows(i)("netdebtamount").ToString

                    Dim listItem As New ListViewItem(vDocno)
                    listItem.SubItems.Add(vRefNo)
                    listItem.SubItems.Add(Format(vNetDebtAmount, "##,##0.00"))
                    Me.ListViewHold.Items.Add(listItem)
                Next

                Dim a As Integer

                For a = 0 To Me.ListViewHold.Items.Count - 1
                    If a Mod 2 <> 0 Then
                        Me.ListViewHold.Items(a).BackColor = Color.Silver
                    End If
                Next

                If Me.ListViewHold.Items.Count > 0 Then
                    Me.ListViewHold.Focus()
                    Me.ListViewHold.Items(0).Selected = True
                    Me.ListViewHold.Items(0).Focused = True
                End If
            Else
                Me.TBHoldSearch.Focus()
                Me.TBHoldSearch.SelectAll()
            End If
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchHold.Visible = False
            vIsOpen = 0
            If Me.TBSearchCheckOut.Enabled = True Then
                Me.TBSearchCheckOut.Focus()
            ElseIf Me.ListViewMerge.Enabled = True Then
                Me.ListViewMerge.Focus()
                Me.ListViewMerge.Items(0).Selected = True
                Me.ListViewMerge.Items(0).Focused = True
            End If
        End If

        If e.KeyCode = Keys.Down Then
            If Me.ListViewHold.Items.Count > 0 Then
                Me.ListViewHold.Focus()
                Me.ListViewHold.Items(0).Selected = True
                Me.ListViewHold.Items(0).Focused = True
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub CallIDNumber()
        Dim vNumber As Integer

        On Error GoTo ErrDescription

        vQuery = "exec dbo.usp_np_searchnewdocno 29"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vNumber = ds.Tables(0).Rows(0)("autonumber").ToString
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNGenCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNGenCheckOut.Click
        Dim i As Integer
        Dim vDocNo As String
        Dim vDocDate As String
        Dim vHeader As String
        Dim vNumber As Integer
        Dim vDocNumber As String

        Dim vItemCode As String
        Dim vBarCode As String
        Dim vQTY As Double
        Dim vUnitCode As String
        Dim vPrice As Double
        Dim vDiscountAmount As Double
        Dim vAmount As Double
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vRefNo As String
        Dim vQueID As Integer
        Dim vSaleCode As String
        Dim vARCode As String
        Dim vCarLicense As String

        On Error GoTo ErrDescription

        If Me.TBMergeID.Text = "" Then
            vQuery = "exec dbo.usp_np_searchnewdocno 30"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            If ds.Tables(0).Rows.Count > 0 Then
                vHeader = Trim(ds.Tables(0).Rows(0)("header").ToString)
                vNumber = ds.Tables(0).Rows(0)("autonumber").ToString
                vDocNumber = Trim(ds.Tables(0).Rows(0)("docnumber").ToString)
            End If

            vDocNo = Trim(vDocNumber & vHeader & "-" & Format(vNumber, "0000"))
        Else
            vDocNo = Me.TBMergeID.Text
        End If

        Me.TBMergeID.Text = vDocNo
        vDocDate = Now.Day & "/" & Now.Month & "/" & Now.Year

        vQuery = "exec dbo.USP_NP_CheckDocDate"
        Dim vService7 As New WebReference.WebServiceCalc
        Dim ds7 As DataSet = vService7.vGetQueryAnlyzer(vQuery)
        If ds7.Tables(0).Rows.Count > 0 Then
            vDocDate = ds7.Tables(0).Rows(0)("vdocdate").ToString
        End If


        vQuery = "exec dbo.USP_NP_InsertDriveInMergeTemp1 '" & vDocNo & "','" & vDocDate & "','" & vARCode & "','" & vSaleCode & "','" & vCarLicense & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & "," & vDiscountAmount & "," & vAmount & ",'" & vBarCode & "','" & vRefNo & "'," & vQueID & "," & i & " "
        Dim vService1 As New WebReference.WebServiceCalc
        Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)
        'Next

        vQuery = "exec dbo.usp_np_updatenewdocno 30"
        Dim vService2 As New WebReference.WebServiceCalc
        Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

        Dim vMergeDocNo As String
        Dim vMergeDocDate1 As Date
        Dim vMergeDocDate As String
        Dim vMergeNetAmount As Double
        Dim m As Integer
        Dim vMergeQTY As Double
        Dim vMergePrice As Double
        Dim vMergeAmount As Double
        Dim vMergeItemCode As String
        Dim vMergeItemName As String
        Dim vMergeItemBar As String
        Dim vMergeUnitCode As String
        Dim vMergeWHCode As String
        Dim vMergeShelfCode As String
        Dim vMergeDriveIn As String
        Dim vMergeDiscount As Double
        Dim vMergeCarLicense As String
        Dim vMergeAR As String
        Dim vMergeSale As String


        vQuery = "exec dbo.USP_NP_CalcDriveInMergeTemp1 '" & vDocNo & "'"
        Dim vService3 As New WebReference.WebServiceCalc
        Dim ds3 As DataSet = vService3.vGetQueryAnlyzer(vQuery)
        If ds3.Tables(0).Rows.Count > 0 Then
            vMergeDocNo = Trim(ds3.Tables(0).Rows(0)("docno").ToString)
            vMergeDocDate1 = Trim(ds3.Tables(0).Rows(0)("docdate").ToString)
            vMergeDocDate = vMergeDocDate1.Day & "/" & vMergeDocDate1.Month & "/" & vMergeDocDate1.Year
            vMergeNetAmount = Trim(ds3.Tables(0).Rows(0)("netamount").ToString)

            Me.ListViewMerge.Visible = True
            Me.ListViewMerge.Items.Clear()
            For i = 0 To ds3.Tables(0).Rows.Count - 1

                m = i + 1
                vMergeItemCode = ds3.Tables(0).Rows(i)("itemcode").ToString
                vMergeItemName = ds3.Tables(0).Rows(i)("itemname").ToString
                vMergeQTY = ds3.Tables(0).Rows(i)("qty").ToString
                vMergePrice = ds3.Tables(0).Rows(i)("price").ToString
                vMergeUnitCode = ds3.Tables(0).Rows(i)("unitcode").ToString
                vMergeItemBar = ds3.Tables(0).Rows(i)("barcode").ToString
                vMergeDriveIn = ds3.Tables(0).Rows(i)("refno").ToString
                vMergeDiscount = ds3.Tables(0).Rows(i)("discountamount").ToString
                vMergeWHCode = ds3.Tables(0).Rows(i)("whcode").ToString
                vMergeShelfCode = ds3.Tables(0).Rows(i)("shelfcode").ToString
                vMergeCarLicense = ds3.Tables(0).Rows(i)("carlicense").ToString
                vMergeAR = ds3.Tables(0).Rows(i)("arcode").ToString
                vMergeSale = ds3.Tables(0).Rows(i)("salecode").ToString

                Dim listItem As New ListViewItem(m)
                listItem.SubItems.Add("")
                listItem.SubItems.Add(vMergeItemName)
                listItem.SubItems.Add(Format(vMergeQTY, "##,##0.00"))
                listItem.SubItems.Add(vMergeUnitCode)
                listItem.SubItems.Add(vMergeItemCode)
                listItem.SubItems.Add(Format(vMergePrice, "##,##0.00"))
                listItem.SubItems.Add(Format(vMergeAmount, "##,##0.00"))
                listItem.SubItems.Add(vMergeWHCode)
                listItem.SubItems.Add(vMergeShelfCode)
                listItem.SubItems.Add(vMergeDriveIn)
                listItem.SubItems.Add(vMergeItemBar)
                listItem.SubItems.Add(vMergeDiscount)
                listItem.SubItems.Add(vMergeCarLicense)
                listItem.SubItems.Add(vMergeAR)
                listItem.SubItems.Add(vMergeSale)
                listItem.SubItems.Add(vDocNo)
                Me.ListViewMerge.Items.Add(listItem)
            Next

        End If

        If Me.ListViewMerge.Items.Count > 0 Then
            Me.ListViewMerge.Focus()
            Me.ListViewMerge.Items(0).Selected = True
            Me.ListViewMerge.Items(0).Focused = True
        Else
            Me.TBSearchCheckOut.Focus()
        End If
        Me.BTNCheckOut.Enabled = True
        Me.BTNGenCheckOut.Enabled = False

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub ListViewMerge_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewMerge.KeyDown

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Dim vIndex As Integer

            If Me.ListViewMerge.Items.Count > 0 Then
                vIndex = Me.ListViewMerge.FocusedItem.Index

                'If vIndex > 0 Then
                Me.TBCheckIndex.Text = vIndex
                Me.TBCheckItemCode.Text = Me.ListViewMerge.Items(vIndex).SubItems(5).Text
                Me.TBCheckItemName.Text = Me.ListViewMerge.Items(vIndex).SubItems(2).Text
                Me.TBCheckUnitCode.Text = Me.ListViewMerge.Items(vIndex).SubItems(4).Text

                If Me.ListViewMerge.Items(vIndex).SubItems(1).Text <> "" Then
                    Me.TBCheckQty.Text = Me.ListViewMerge.Items(vIndex).SubItems(1).Text
                Else
                    Me.TBCheckQty.Text = ""
                End If
                Me.PNCheckQty.Visible = True
                Me.PNCheckQty.BringToFront()
                Me.ListViewMerge.Enabled = False

                Me.TBCheckQty.Focus()
                Me.TBCheckQty.SelectAll()
                'End If
            End If
        End If

        If e.KeyCode = Keys.Up Then
            Dim n As Integer

            If Me.ListViewMerge.Items.Count > 0 Then
                n = Me.ListViewMerge.FocusedItem.Index
                If n = 0 Then
                    Me.TBSearchCheckOut.Focus()
                    Me.TBSearchCheckOut.SelectAll()
                End If
            End If
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldingBill()
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If

        If e.KeyCode = 115 Then
            Call SearchHoldBill()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBCheckQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBCheckQty.KeyDown
        Dim vCountItem As Double
        Dim vIndex As Integer
        Dim vNextIndex As Integer
        Dim vQty As Double
        Dim vItemAmount As Double
        Dim vPrice As Double

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewMerge.Items.Count > 0 Then
                If Me.TBCheckQty.Text <> "" Then
                    vCountItem = Me.TBCheckQty.Text
                    vIndex = Me.TBCheckIndex.Text
                    If Me.ListViewMerge.Items(vIndex).SubItems(3).Text <> "" Then
                        vQty = Me.ListViewMerge.Items(vIndex).SubItems(3).Text
                    End If

                    If Me.ListViewMerge.Items(vIndex).SubItems(6).Text <> "" Then
                        vPrice = Me.ListViewMerge.Items(vIndex).SubItems(6).Text
                    End If

                    vItemAmount = vPrice * vQty

                    Me.ListViewMerge.Items(vIndex).SubItems(1).Text = Format(vCountItem, "##,##0.00")
                    Me.ListViewMerge.Items(vIndex).SubItems(7).Text = Format(vItemAmount, "##,##0.00")
                    vNextIndex = vIndex + 1

                    If vQty <> vCountItem Then
                        Me.ListViewMerge.Items(vIndex).BackColor = Color.Red
                        MsgBox("�ӹǹ����Ǩ�ͺ���ç�Ѻ�ӹǹ�������Թ���", MsgBoxStyle.Critical, "Send Information Message")
                    Else
                        Me.ListViewMerge.Items(vIndex).BackColor = Color.White
                    End If

                    Me.PNCheckQty.Visible = False

                    Call vCalcCheckOutKeyQuanity()

                    If vCountItem >= 10000 Then
                        MsgBox("�ӹǹ�Թ��ҷ�����͡� �Թ 10,000 ��سҵ�Ǩ�ͺ�������ա��", MsgBoxStyle.Information, "Send Error Message")
                    End If

                    Me.TBCheckItemCode.Text = ""
                    Me.TBCheckUnitCode.Text = ""
                    Me.TBCheckItemName.Text = ""
                    Me.ListViewMerge.Enabled = True

                    If vIndex < Me.ListViewMerge.Items.Count - 1 Then
                        Me.ListViewMerge.Items(vIndex).Focused = False
                        Me.ListViewMerge.Items(vIndex).Selected = False

                        Me.ListViewMerge.Items(vNextIndex).Focused = True
                        Me.ListViewMerge.Items(vNextIndex).Selected = True
                        Me.ListViewMerge.Focus()

                    ElseIf vIndex = Me.ListViewMerge.Items.Count - 1 Then
                        Me.ListViewMerge.Items(vIndex).Focused = True
                        Me.ListViewMerge.Items(vIndex).Selected = True
                        Me.ListViewMerge.Focus()
                    End If

                End If
            End If
            Me.ListViewMerge.Enabled = True
        End If

        If e.KeyCode = Keys.Escape Then
            Me.ListViewMerge.Enabled = True
            Me.PNCheckQty.Visible = False
            vIndex = Me.TBCheckIndex.Text
            If Me.ListViewMerge.Items.Count > 0 Then
                Me.ListViewMerge.Focus()
                Me.ListViewMerge.Items(vIndex).Selected = True
                Me.ListViewMerge.Items(vIndex).Focused = True
            Else
                Me.TBSearchCheckOut.Focus()
                Me.TBSearchCheckOut.SelectAll()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub MenuMergeSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vIndex As Integer
        Dim vCountItem As Double

        On Error GoTo ErrDescription

        If Me.ListViewMerge.Items.Count > 0 Then

            vIndex = Me.ListViewMerge.FocusedItem.Index
            Me.TBCheckIndex.Text = vIndex
            Me.TBCheckItemCode.Text = Me.ListViewMerge.Items(vIndex).SubItems(5).Text
            Me.TBCheckItemName.Text = Me.ListViewMerge.Items(vIndex).SubItems(2).Text
            Me.TBCheckUnitCode.Text = Me.ListViewMerge.Items(vIndex).SubItems(4).Text
            If Me.ListViewMerge.Items(vIndex).SubItems(1).Text <> "" Then
                vCountItem = Me.ListViewMerge.Items(vIndex).SubItems(1).Text
            Else
                vCountItem = 0
            End If
            If vCountItem <> 0 Then
                Me.TBCheckQty.Text = Format(vCountItem, "##,##0.00")
                Me.TBCheckQty.Focus()
                Me.TBCheckQty.SelectAll()
            Else
                Me.TBCheckQty.Text = ""
                Me.TBCheckQty.Focus()
                Me.TBCheckQty.SelectAll()
            End If

            Me.PNCheckQty.Visible = True
            Me.PNCheckQty.BringToFront()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Public Sub SelectItemMerge()
        Dim vIndex As Integer
        Dim vCountItem As Double

        On Error GoTo ErrDescription

        If Me.ListViewMerge.Items.Count > 0 Then
            vIndex = Me.ListViewMerge.FocusedItem.Index
            Me.TBCheckIndex.Text = vIndex
            Me.TBCheckItemCode.Text = Me.ListViewMerge.Items(vIndex).SubItems(5).Text
            Me.TBCheckItemName.Text = Me.ListViewMerge.Items(vIndex).SubItems(2).Text
            Me.TBCheckUnitCode.Text = Me.ListViewMerge.Items(vIndex).SubItems(4).Text
            If Me.ListViewMerge.Items(vIndex).SubItems(1).Text <> "" Then
                vCountItem = Me.ListViewMerge.Items(vIndex).SubItems(1).Text
            Else
                vCountItem = 0
            End If
            If vCountItem <> 0 Then
                Me.TBCheckQty.Text = Format(vCountItem, "##,##0.00")
                Me.TBCheckQty.Focus()
                Me.TBCheckQty.SelectAll()
            Else
                Me.TBCheckQty.Text = ""
                Me.TBCheckQty.Focus()
                Me.TBCheckQty.SelectAll()
            End If

            Me.PNCheckQty.Visible = True
            Me.PNCheckQty.BringToFront()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub MenuMergeAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuMergeAdd.Click
        On Error Resume Next

        If Me.ListViewMerge.Items.Count > 0 Then
            Me.PNAddItem.Visible = True
            Me.PNAddItem.BringToFront()
            Me.TBSearchBarCode.Focus()
        Else
            MsgBox("��������Թ��� ��ͧ���Թ�����͡��âͧ�١��� ���ҧ���� 1 ��¡�� �ó������ �����价��͡���������º����", MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub AddItem()
        On Error Resume Next

        If Me.ListViewMerge.Items.Count > 0 Then
            Me.PNAddItem.Visible = True
            Me.PNAddItem.BringToFront()
            Me.TBSearchBarCode.Focus()
        Else
            MsgBox("��������Թ��� ��ͧ���Թ�����͡��âͧ�١��� ���ҧ���� 1 ��¡�� �ó������ �����价��͡���������º����", MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub
    Private Sub MenuMergeCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim i As Integer
        Dim vAnswer As Integer
        Dim vIndex As Integer

        On Error GoTo ErrDescription

        If Me.ListViewMerge.Items.Count > 0 Then
            i = Me.ListViewMerge.FocusedItem.Index
            vIndex = i + 1

            vAnswer = MsgBox("�س��ͧ���ź��¡�÷�� " & vIndex & " ��������� ?", MsgBoxStyle.YesNo, "Send Question Message ?")
            If vAnswer = 6 Then
                Me.ListViewMerge.Items.RemoveAt(i)
                Call GenIDNumberMerge()
                Call vCalcCheckOutKeyQuanity()

                If Me.ListViewMerge.Items.Count > 0 Then
                    Me.ListViewMerge.Focus()
                    Me.ListViewMerge.Items(0).Selected = True
                    Me.ListViewMerge.Items(0).Focused = True
                End If
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub
    Public Sub DeleteItemMerge()
        Dim i As Integer
        Dim vAnswer As Integer
        Dim vIndex As Integer

        On Error GoTo ErrDescription

        If Me.ListViewMerge.Items.Count > 0 Then
            i = Me.ListViewMerge.FocusedItem.Index
            vIndex = i + 1

            vAnswer = MsgBox("�س��ͧ���ź��¡�÷�� " & vIndex & " ��������� ?", MsgBoxStyle.YesNo, "Send Question Message ?")
            If vAnswer = 6 Then
                Me.ListViewMerge.Items.RemoveAt(i)
                Call GenIDNumberMerge()
                Call vCalcCheckOutKeyQuanity()

                If Me.ListViewMerge.Items.Count > 0 Then
                    Me.ListViewMerge.Focus()
                    Me.ListViewMerge.Items(0).Selected = True
                    Me.ListViewMerge.Items(0).Focused = True
                End If
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub


    Private Sub TBAutoKeyQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
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

    Private Sub BTNSelectItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectItem.Click
        On Error Resume Next

        Me.TBMergeID.Text = ""
        Me.ListViewMerge.Items.Clear()
        Me.ListViewMerge.Visible = False
        Call SearchItemCheckOut()
        Me.BTNGenCheckOut.Enabled = False
        Me.BTNCheckOut.Enabled = False
    End Sub

    Public Sub SelectQueItem()
        On Error Resume Next

        Me.TBMergeID.Text = ""
        Me.ListViewMerge.Items.Clear()
        Me.ListViewMerge.Visible = False
        Call SearchItemCheckOut()
        Me.BTNGenCheckOut.Enabled = False
        Me.BTNCheckOut.Enabled = False
    End Sub

    Private Sub BTNHoldingClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNHoldingClose.Click

        On Error GoTo ErrDescription

        MsgBox("���駵��� ��顴���� ESC ���͡�Ѻ�˹�Ҩ����͡����� CheckOut", MsgBoxStyle.Information, "Send Information Message")

        Me.ListViewHolding.Items.Clear()
        Me.TBHoldARName.Text = ""
        Me.TBHoldNo.Text = ""
        Me.ListViewHolding.Items.Clear()
        Me.LBLHoldingAmount.Text = ""
        Me.BTNPrintHoldBill.Visible = False
        Me.PNHolding.Visible = False
        Me.PNChecker.Enabled = True
        vIsOpen = 0

        If Me.ListViewMerge.Items.Count > 0 Then
            Me.ListViewMerge.Focus()
            Me.ListViewMerge.Items(0).Selected = True
            Me.ListViewMerge.Items(0).Focused = True
        Else
            Me.TBSearchCheckOut.SelectAll()
            Me.TBSearchCheckOut.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Public Sub HoldBillClose()
        On Error GoTo ErrDescription

        Me.ListViewHolding.Items.Clear()
        Me.TBHoldARName.Text = ""
        Me.TBHoldNo.Text = ""
        Me.ListViewHolding.Items.Clear()
        Me.LBLHoldingAmount.Text = ""
        Me.BTNPrintHoldBill.Visible = False
        Me.PNHolding.Visible = False
        Me.PNChecker.Enabled = True
        vIsOpen = 0

        If Me.ListViewMerge.Items.Count > 0 Then
            Me.ListViewMerge.Focus()
            Me.ListViewMerge.Items(0).Selected = True
            Me.ListViewMerge.Items(0).Focused = True
        Else
            Me.TBSearchCheckOut.SelectAll()
            Me.TBSearchCheckOut.Focus()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBHoldingAR_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBHoldingAR.TextChanged
        Dim vQuery As String
        Dim vSearchAR As String

        On Error GoTo ErrDescription

        If Me.TBHoldingAR.Text <> "" Then
            vSearchAR = Me.TBHoldingAR.Text

            vQuery = "exec dbo.usp_ar_searchar1 '" & vSearchAR & "' "
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            If ds.Tables(0).Rows.Count > 0 Then
                Me.TBHoldARName.Text = ds.Tables(0).Rows(0)("arname").ToString()
                Me.TBHoldingMemberID.Text = ds.Tables(0).Rows(0)("memberid").ToString
            Else
                Me.TBHoldARName.Text = ""
                Me.TBHoldingMemberID.Text = ""
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSelectItemQue_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.TBSearchCheckOut.Text = ""
            Me.TBSearchCheckOut.Focus()
        End If
    End Sub

    Private Sub Cash01_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Cash01.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Right Then
            Me.Cash02.Checked = True
        End If

        If e.KeyCode = Keys.Escape Then
            Call HoldBillClose()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldBill()
        End If

        If e.KeyCode = 115 Then
            Call PrintHoldBillNo()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub Cash02_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Cash02.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Right Then
            Me.Cash03.Checked = True
        End If

        If e.KeyCode = Keys.Left Then
            Me.Cash01.Checked = True
        End If

        If e.KeyCode = Keys.Escape Then
            Call HoldBillClose()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldBill()
        End If

        If e.KeyCode = 115 Then
            Call PrintHoldBillNo()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub Cash03_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Cash03.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Left Then
            Me.Cash02.Checked = True
        End If

        If e.KeyCode = Keys.Escape Then
            Call HoldBillClose()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldBill()
        End If

        If e.KeyCode = 115 Then
            Call PrintHoldBillNo()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSelectItemQueExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next

        Me.TBSearchCheckOut.Focus()
        Me.TBSearchCheckOut.SelectAll()
    End Sub

    Private Sub BTNSelectItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSelectItem.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SelectQueItem()
        End If

        If e.KeyCode = 34 Then
            Call ItemSelectHoldBill()
        End If

        If e.KeyCode = 115 Then
            Call SearchHoldBill()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNClearCheckOut_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNClearCheckOut.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldingBill()
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If

        If e.KeyCode = 115 Then
            Call SearchHoldBill()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNGenCheckOut_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNGenCheckOut.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Call ClearScreen()
        End If

        If e.KeyCode = 114 Then
            Call SelectQueItem()
        End If

        If e.KeyCode = 16 Then
            Call ItemSelectHoldBill()
        End If

        If e.KeyCode = 34 Then
            Call SearchHoldBill()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNCheckOut_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCheckOut.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldingBill()
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If

        If e.KeyCode = 115 Then
            Call SearchHoldBill()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSearch.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldingBill()
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If

        If e.KeyCode = 115 Then
            Call SearchHoldBill()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub ListViewHolding_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewHolding.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Call HoldBillClose()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldBill()
        End If

        If e.KeyCode = 115 Then
            Call PrintHoldBillNo()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNGenBill_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNGenBill.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Call HoldBillClose()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldBill()
        End If

        If e.KeyCode = 115 Then
            Call PrintHoldBillNo()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNHoldingClose_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNHoldingClose.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Call HoldBillClose()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldBill()
        End If

        If e.KeyCode = 115 Then
            Call PrintHoldBillNo()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBHoldingARName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBHoldARName.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Call HoldBillClose()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldBill()
        End If

        If e.KeyCode = 115 Then
            Call PrintHoldBillNo()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Dim vAnswer As Integer

        On Error GoTo ErrDescription

        MsgBox("���駵��� ��顴���� ESC �����͡�����", MsgBoxStyle.Information, "Send Information Message")

        vAnswer = MsgBox("�س��ͧ����͡�ҡ��������������", MsgBoxStyle.YesNo, "Send Question Message")
        If vAnswer = 6 Then
            Application.Exit()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub ExitProgram()
        Dim vAnswer As Integer

        On Error GoTo ErrDescription

        vAnswer = MsgBox("�س��ͧ����͡�ҡ��������������", MsgBoxStyle.YesNo, "Send Question Message")
        If vAnswer = 6 Then
            Application.Exit()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNExit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNExit.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = 33 Then
            Call ClearScreen()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldingBill()
        End If

        If e.KeyCode = 114 Then
            Call AddItem()
        End If

        If e.KeyCode = 115 Then
            Call SearchHoldBill()
        End If

        If e.KeyCode = Keys.Escape Then
            Call ExitProgram()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

    End Sub

    Private Sub PrintHoldBill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPrintHoldBill.Click
        Dim vDocNo As String

        On Error GoTo ErrDescription

        If Me.TBHoldNo.Text <> "" And Me.ListViewHolding.Items.Count > 0 And Me.TBHoldARName.Text <> "" Then

            MsgBox("���駵��� ��顴���� �����+�����Ţ 2 ���ͷӡ�þ���췴᷹ 㺾ѡ���", MsgBoxStyle.Information, "Send Information Message")

            vDocNo = Me.TBHoldNo.Text
            vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto1 9,'" & vDocNo & "','','" & vUserName & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As Integer = vService.vExecuteQuery(vQuery)

            vIsOpen = 0
            Me.ListViewMerge.Enabled = False
            Me.TBSearchCheckOut.Enabled = True
            Me.ListViewHolding.Items.Clear()
            Me.TBHoldARName.Text = ""
            Me.LBLHoldingAmount.Text = ""
            Me.PNHolding.Visible = False
            Me.BTNPrintHoldBill.Visible = False
            Me.TBHoldNo.Text = ""
            vIsOpen = 0

            Me.ListViewMerge.Items.Clear()
            Me.TBSearchCheckOut.Text = ""
            Me.LBLNetAmount.Text = ""
            Me.LBLCheckOutAmount.Text = ""
            Me.TBMergeID.Text = ""
            Me.BTNCheckOut.Enabled = False
            Me.BTNGenCheckOut.Enabled = False
            Me.PNChecker.Enabled = True
            Me.TBSearchCheckOut.Focus()

            MsgBox("��㺾ѡ��� 仾���췴᷹���º�������� ", MsgBoxStyle.Information, "Send Information Message")
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Public Sub PrintHoldBillNo()
        Dim vDocNo As String

        On Error GoTo ErrDescription

        If Me.TBHoldNo.Text <> "" And Me.ListViewHolding.Items.Count > 0 And Me.TBHoldARName.Text <> "" Then
            vDocNo = Me.TBHoldNo.Text
            vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto1 9,'" & vDocNo & "','','" & vUserName & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As Integer = vService.vExecuteQuery(vQuery)

            Me.ListViewMerge.Enabled = False
            Me.TBSearchCheckOut.Enabled = True
            Me.ListViewHolding.Items.Clear()
            Me.TBHoldARName.Text = ""
            Me.LBLHoldingAmount.Text = ""
            Me.PNHolding.Visible = False
            Me.BTNPrintHoldBill.Visible = False
            Me.TBHoldNo.Text = ""
            vIsOpen = 0

            Me.ListViewMerge.Items.Clear()
            Me.TBSearchCheckOut.Text = ""
            Me.LBLNetAmount.Text = ""
            Me.LBLCheckOutAmount.Text = ""
            Me.TBMergeID.Text = ""
            Me.BTNCheckOut.Enabled = False
            Me.BTNGenCheckOut.Enabled = False
            Me.PNChecker.Enabled = True
            Me.TBSearchCheckOut.Focus()
            MsgBox("��㺾ѡ��� 仾���췴᷹���º�������� ", MsgBoxStyle.Information, "Send Information Message")

        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSelectHoldBill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectHoldBill.Click
        Dim i As Integer
        Dim vDocNo As String
        Dim vDriveInNo As String
        Dim vMergeNo As String
        Dim n As Integer
        Dim vDocDate As String
        Dim vQueID As Integer
        Dim vPickZone As String
        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPickQTY As Double
        Dim vConfirmQTY As Double
        Dim vUnitCode As String
        Dim vPrice As Double
        Dim vDiscountAmount As Double
        Dim vAmount As Double
        Dim vIndex As Integer
        Dim vLine As Integer
        Dim vBarcode As String
        Dim vLicense As String
        Dim vARCode As String
        Dim vSaleCode As String
        Dim vNetDebtAmount As Double

        On Error GoTo ErrDescription

        If Me.ListViewHold.Items.Count > 0 Then
            n = Me.ListViewHold.FocusedItem.Index
            vDocNo = Me.ListViewHold.Items(n).SubItems(0).Text

            vQuery = "exec dbo.usp_np_SearchHoldingDetails1 '" & vDocNo & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            vIndex = 0
            If ds.Tables(0).Rows.Count > 0 Then

                Me.TBHoldNo.Text = ds.Tables(0).Rows(i)("docno").ToString
                vNetDebtAmount = ds.Tables(0).Rows(i)("netdebtamount").ToString
                Me.LBLHoldingAmount.Text = Format(vNetDebtAmount, "##,##0.00")
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    vItemName = ds.Tables(0).Rows(i)("itemname").ToString
                    vQTY = ds.Tables(0).Rows(i)("qty").ToString
                    vUnitCode = ds.Tables(0).Rows(i)("unitcode").ToString
                    vItemCode = ds.Tables(0).Rows(i)("itemcode").ToString
                    vPrice = ds.Tables(0).Rows(i)("price").ToString
                    vDiscountAmount = ds.Tables(0).Rows(i)("price").ToString
                    vAmount = ds.Tables(0).Rows(i)("amount").ToString
                    vWHCode = ds.Tables(0).Rows(i)("whcode").ToString
                    vShelfCode = ds.Tables(0).Rows(i)("shelfcode").ToString
                    vDriveInNo = ds.Tables(0).Rows(i)("driveinrefno").ToString
                    vBarcode = ds.Tables(0).Rows(i)("barcode").ToString
                    vLicense = ds.Tables(0).Rows(i)("license").ToString
                    vARCode = ds.Tables(0).Rows(i)("arcode").ToString
                    vMergeNo = ds.Tables(0).Rows(i)("mergeno").ToString
                    vSaleCode = ds.Tables(0).Rows(i)("salecode").ToString

                    Me.TBHoldARName.Text = vARCode

                    vIndex = vIndex + 1
                    vLine = vIndex - 1
                    Dim listItem As New ListViewItem(vIndex)
                    listItem.SubItems.Add(vItemName)
                    listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                    listItem.SubItems.Add(vUnitCode)
                    listItem.SubItems.Add(vItemCode)
                    listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                    listItem.SubItems.Add(Format(vDiscountAmount, "##,##0.00"))
                    listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                    listItem.SubItems.Add(vWHCode)
                    listItem.SubItems.Add(vShelfCode)
                    listItem.SubItems.Add(vDriveInNo)
                    listItem.SubItems.Add(vBarcode)
                    listItem.SubItems.Add(vLicense)
                    listItem.SubItems.Add(vARCode)
                    listItem.SubItems.Add(vSaleCode)
                    listItem.SubItems.Add(vMergeNo)
                    Me.ListViewHolding.Items.Add(listItem)
                Next

                vIsOpen = 1

                Call vCalcCheckOutKeyQuanity()

                Me.PNHolding.Visible = True
                Me.PNHolding.BringToFront()
                Me.PNSearchHold.Visible = False
                Me.BTNPrintHoldBill.Visible = True
                Me.BTNGenBill.Visible = False

                If ListViewHolding.Items.Count > 0 Then
                    Me.ListViewHolding.Focus()
                    Me.ListViewHolding.Items(0).Selected = True
                    Me.ListViewHolding.Items(0).Focused = True
                End If

            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSelectHoldBill_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSelectHoldBill.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchHold.Visible = False
            vIsOpen = 0
            If Me.TBSearchCheckOut.Enabled = True Then
                Me.TBSearchCheckOut.Focus()
            ElseIf Me.ListViewMerge.Enabled = True Then
                Me.ListViewMerge.Focus()
                Me.ListViewMerge.Items(0).Selected = True
                Me.ListViewMerge.Items(0).Focused = True
            End If
        End If
    End Sub

    Private Sub BTNCloseHoldBill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseHoldBill.Click
        On Error Resume Next

        vIsOpen = 0
        Me.PNSearchHold.Visible = False
        If Me.TBSearchCheckOut.Enabled = True Then
            Me.TBSearchCheckOut.Focus()
        ElseIf Me.ListViewMerge.Enabled = True Then
            Me.ListViewMerge.Focus()
            Me.ListViewMerge.Items(0).Selected = True
            Me.ListViewMerge.Items(0).Focused = True
        End If

    End Sub

    Private Sub BTNCloseHoldBill_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNCloseHoldBill.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Escape Then
            Me.PNSearchHold.Visible = False
            vIsOpen = 0
            If Me.TBSearchCheckOut.Enabled = True Then
                Me.TBSearchCheckOut.Focus()
            ElseIf Me.ListViewMerge.Enabled = True Then
                Me.ListViewMerge.Focus()
                Me.ListViewMerge.Items(0).Selected = True
                Me.ListViewMerge.Items(0).Focused = True
            End If
        End If
    End Sub

    Private Sub BTNPrintHoldBill_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNPrintHoldBill.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Escape Then
            Call HoldBillClose()
        End If

        If e.KeyCode = 34 Then
            Call GenHoldBill()
        End If

        If e.KeyCode = 115 Then
            Call PrintHoldBillNo()
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub
End Class
