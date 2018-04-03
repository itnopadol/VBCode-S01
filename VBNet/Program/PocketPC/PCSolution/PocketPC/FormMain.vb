Imports System.Data
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms

Public Class frmDriveIn
    Dim ds As DataSet
    Dim vQuery As String

    Private Sub TBBarCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBBarCode.KeyDown
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


        If e.KeyCode = 134 Then
            Call SavePickUp()
        End If


        If e.KeyCode = Keys.Enter Then
            If Me.TBBarCode.Text <> "" Then
                vBarCode = Me.TBBarCode.Text
            Else
                Me.TBBarCode.Focus()
            End If

            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetDataBarCode(vBarCode)
            Me.ListViewStock.Items.Clear()

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
                    Me.ListViewStock.Items.Add(listItem)
                Next

                Dim n As Integer
                Dim vCheckItemCode As String
                Dim vCheckQTY As Double

                If Me.ListViewItem.Items.Count > 0 Then
                    For n = 0 To Me.ListViewItem.Items.Count - 1
                        vCheckItemCode = Me.ListViewItem.Items(n).SubItems(4).Text
                        vCheckQTY = Me.ListViewItem.Items(n).SubItems(2).Text
                        If vItemCode = vCheckItemCode Then
                            Me.TBQTY.Text = Format(vCheckQTY, "##,##0.00")
                            GoTo Line1
                        End If
                    Next
                End If

Line1:
                Me.TBQTY.Focus()
                Me.TBQTY.SelectAll()
            Else
                Me.TBBarCode.Focus()
                Me.TBQTY.SelectAll()
            End If

            Me.TBItem.Text = vItemCode
            Me.TBItemName.Text = vItemName
            Me.TBPrice.Text = Format(vPrice, "##,##0.00")
            Me.TBRate.Text = Format(vRate, "##,##0.00")
            Me.TBReserve.Text = Format(vReserveQTY, "##,##0.00")
            Me.TBUnit.Text = vUnitCode
            Me.TBWHCode.Text = vDefWHCode
            Me.TBShelfCode.Text = vDefShelfCode
            Me.TBMemBarCode.Text = vBarCode

        End If

        If e.KeyCode = Keys.Back Then
            Me.TBItem.Text = ""
            Me.TBItemName.Text = ""
            Me.TBPrice.Text = ""
            Me.TBReserve.Text = ""
            Me.TBUnit.Text = ""
            Me.TBWHCode.Text = ""
            Me.TBShelfCode.Text = ""
            Me.TBQTY.Text = ""
            Me.TBRate.Text = ""
            Me.TBMemBarCode.Text = ""
            Me.ListViewStock.Items.Clear()
        End If
    End Sub

    Private Sub TBBarCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBBarCode.TextChanged
        If Me.TBBarCode.Text <> "" Then
            Me.PNItemDetails.Visible = True
            Me.PNItemDetails.BringToFront()
            Me.BTNSave.Visible = False
        Else
            Me.PNItemDetails.Visible = False
            Me.PNDriveIn.Visible = True
            Me.PNDriveIn.BringToFront()
            Me.BTNSave.Visible = True
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
            Me.TBReserve.Text = ""
            Me.TBUnit.Text = ""
            Me.TBWHCode.Text = ""
            Me.TBShelfCode.Text = ""
            Me.TBQTY.Text = ""
            Me.ListViewStock.Items.Clear()
            Me.PNItemDetails.Visible = False
            Me.TBBarCode.Focus()
        End If
    End Sub

    Private Sub CalcItemAmount()
        Dim i As Integer
        Dim vAmount As Double
        Dim vSumAmount As Double

        If Me.ListViewItem.Items.Count > 0 Then
            For i = 0 To Me.ListViewItem.Items.Count - 1
                vAmount = Me.ListViewItem.Items(i).SubItems(6).Text
                vSumAmount = vSumAmount + vAmount
            Next
            Me.TBItemAmount.Text = Format(vSumAmount, "##,##0.00")
        Else
            Me.TBItemAmount.Text = Format(0, "##,##0.00")
        End If
    End Sub

    Private Sub CalcCheckItemQtyAmount()
        Dim i As Integer
        Dim vAmount As Double
        Dim vSumAmount As Double

        If Me.ListViewCheckOut.Items.Count > 0 Then
            For i = 0 To Me.ListViewCheckOut.Items.Count - 1
                vAmount = Me.ListViewCheckOut.Items(i).SubItems(7).Text
                vSumAmount = vSumAmount + vAmount
            Next
            Me.TBItemAmount.Text = Format(vSumAmount, "##,##0.00")
        Else
            Me.TBItemAmount.Text = Format(0, "##,##0.00")
        End If
    End Sub
    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vHeader As String
        Dim vNumber As Integer
        Dim vDocNumber As String

        Dim vDocNo As String
        Dim vDocDate As String
        Dim vID As Integer
        Dim vRefNo As String
        Dim vPickZone As String
        Dim vTotalNetAmount As Double
        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vUserID As String
        Dim i As Integer
        Dim vBarCode As String
        Dim vLineNumber As Integer


        If Me.ListViewItem.Items.Count > 0 And Me.TBItemAmount.Text <> "" Then
            vUserID = Me.TBUserID.Text

            If Me.TBRefNo.Text = "" Then
                MsgBox("ต้องกรอกเลขที่อ้างอิงคิวก่อนบันทึกข้อมูลทุกครั้ง", MsgBoxStyle.Critical, "Send Error Message")
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
                vID = vNumber
                vRefNo = Me.TBRefNo.Text
                If Me.RDZone1.Checked = True Then
                    vPickZone = "01"
                ElseIf Me.RDZone2.Checked = True Then
                    vPickZone = "02"
                ElseIf Me.RDZone3.Checked = True Then
                    vPickZone = "03"
                ElseIf Me.RDZone4.Checked = True Then
                    vPickZone = "04"
                End If

                vConnectZone = vPickZone


                vTotalNetAmount = Me.TBItemAmount.Text

                Call CallIDNumber()

                vQuery = "exec dbo.usp_np_insertdriveinslip '" & vDocNo & "','" & vDocDate & "'," & vID & ",'" & vRefNo & "','" & vPickZone & "'," & vTotalNetAmount & ",'" & vUserID & "'"
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
                    vLineNumber = i

                    vQuery = "exec dbo.usp_np_insertdriveinslipsub '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "'," & vQTY & "," & vQTY & ",'" & vUnitCode & "'," & vPrice & "," & vAmount & ",'" & vBarCode & "'," & vLineNumber & " "
                    Dim vService2 As New WebReference.WebServiceCalc
                    Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                Next

                If Me.TBDocNo.Text = "" Then
                    vQuery = "exec dbo.usp_np_updatenewdocno 29"
                    Dim vService3 As New WebReference.WebServiceCalc
                    Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

                    MsgBox("ได้เลขที่คิว " & vID & " ", MsgBoxStyle.Information, "Send Information Message")
                Else
                    MsgBox("แก้ไขเลขที่เอกสาร " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")
                End If

                Me.ListViewItem.Items.Clear()
                Me.TBID.Text = ""
                Me.TBRefNo.Text = ""
                Me.TBItemAmount.Text = ""
                Me.TBID.Enabled = True
                Me.TBDocNo.Text = ""
                Me.TBBarCode.Text = ""
                Call CallIDNumber()
                Me.TBRefNo.Focus()
            End If
            End If
    End Sub

    Private Sub SavePickUp()
        Dim vHeader As String
        Dim vNumber As Integer
        Dim vDocNumber As String

        Dim vDocNo As String
        Dim vDocDate As String
        Dim vID As Integer
        Dim vRefNo As String
        Dim vPickZone As String
        Dim vTotalNetAmount As Double
        Dim vItemCode As String
        Dim vItemName As String
        Dim vWHCode As String
        Dim vShelfCode As String
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String
        Dim vUserID As String
        Dim i As Integer
        Dim vBarCode As String
        Dim vLineNumber As Integer


        If Me.ListViewItem.Items.Count > 0 And Me.TBItemAmount.Text <> "" Then
            vUserID = Me.TBUserID.Text

            If Me.TBRefNo.Text = "" Then
                MsgBox("ต้องกรอกเลขที่อ้างอิงคิวก่อนบันทึกข้อมูลทุกครั้ง", MsgBoxStyle.Critical, "Send Error Message")
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
                vID = vNumber
                vRefNo = Me.TBRefNo.Text
                If Me.RDZone1.Checked = True Then
                    vPickZone = "01"
                ElseIf Me.RDZone2.Checked = True Then
                    vPickZone = "02"
                ElseIf Me.RDZone3.Checked = True Then
                    vPickZone = "03"
                ElseIf Me.RDZone4.Checked = True Then
                    vPickZone = "04"
                End If

                vConnectZone = vPickZone


                vTotalNetAmount = Me.TBItemAmount.Text
                Call CallIDNumber()

                vQuery = "exec dbo.usp_np_insertdriveinslip '" & vDocNo & "','" & vDocDate & "'," & vID & ",'" & vRefNo & "','" & vPickZone & "'," & vTotalNetAmount & ",'" & vUserID & "'"
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
                    vLineNumber = i

                    vQuery = "exec dbo.usp_np_insertdriveinslipsub '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "'," & vQTY & "," & vQTY & ",'" & vUnitCode & "'," & vPrice & "," & vAmount & ",'" & vBarCode & "'," & vLineNumber & " "
                    Dim vService2 As New WebReference.WebServiceCalc
                    Dim ds2 As Integer = vService2.vExecuteQuery(vQuery)

                Next

                If Me.TBDocNo.Text = "" Then
                    vQuery = "exec dbo.usp_np_updatenewdocno 29"
                    Dim vService3 As New WebReference.WebServiceCalc
                    Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

                    MsgBox("ได้เลขที่คิว" & vID & " ", MsgBoxStyle.Information, "Send Information Message")
                Else
                    MsgBox("แก้ไขเลขที่เอกสาร " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information Message")
                End If

                Me.ListViewItem.Items.Clear()
                Me.TBID.Text = ""
                Me.TBRefNo.Text = ""
                Me.TBItemAmount.Text = ""
                Me.TBID.Enabled = True
                Me.TBDocNo.Text = ""
                Me.TBBarCode.Text = ""
                Call CallIDNumber()
                Me.TBRefNo.Focus()
            End If
        End If
    End Sub
    Private Sub frmProgram1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PNDriveIn.Visible = False
        Me.PNChecker.Visible = True

        Me.MenuProgram.Enabled = False
        Me.MenuMain.Enabled = True

        Me.PNLogIn.Visible = True
        Me.PNLogIn.BringToFront()
        Me.RDZone1.Focus()
        'Me.TBUserCode.Focus()

        'Call CallIDNumber()
    End Sub

    Private Sub CallIDNumber()
        Dim vNumber As Integer

        vQuery = "exec dbo.usp_np_searchnewdocno 29"
        'vQuery = "select qty as autonumber from bcstklocation where itemcode = '2120250'and whcode = 's01' and shelfcode = 'bk3'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        If ds.Tables(0).Rows.Count > 0 Then
            vNumber = ds.Tables(0).Rows(0)("autonumber").ToString
        End If

        Me.TBID.Text = vNumber
    End Sub

    Private Sub TBRefNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBRefNo.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBBarCode.Focus()
        End If

        If e.KeyCode = 134 Then
            Call SavePickUp()
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


        If e.KeyCode = Keys.Escape Then
            Me.PNItemDetails.Visible = False
            Me.TBBarCode.Text = ""
            Me.TBBarCode.Focus()
        End If


        If e.KeyCode = Keys.Enter Then
            If Me.ListViewStock.Items.Count > 0 And Me.TBItem.Text <> "" Then
                vCheckShelf = Me.TBShelfCode.Text
                vCheckUnit = Me.TBUnit.Text
                If Me.ListViewStock.Items.Count > 0 Then
                    For v = 0 To Me.ListViewStock.Items.Count - 1
                        vListShelf = Me.ListViewStock.Items(v).Text
                        vListUnit = Me.ListViewStock.Items(v).SubItems(2).Text
                        If vCheckShelf = vListShelf And vCheckUnit = vListUnit Then
                            vShelfQTY = Me.ListViewStock.Items(v).SubItems(1).Text
                            vShelfUnit = Me.ListViewStock.Items(v).SubItems(2).Text
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

                If Me.TBQTY.Text <> "" Then
                    vQTY = Me.TBQTY.Text
                End If

                If vShelfUnit <> vUnitCode Then
                    vTotalQTY = vShelfQTY / vRate
                    If vQTY > vTotalQTY Then
                        MsgBox("ไม่สามารถขายเกินยอดที่มีใน STOCK ได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message ")
                        Me.TBQTY.SelectAll()
                        Exit Sub
                    End If
                End If

                If vQTY > vShelfQTY And vShelfUnit = vUnitCode Then
                    MsgBox("ไม่สามารถขายเกินยอดที่มีใน STOCK ได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message ")
                    Me.TBQTY.SelectAll()
                    Exit Sub
                End If

                If Me.TBPrice.Text <> "" Then
                    vPrice = Me.TBPrice.Text
                End If
                vAmount = vQTY * vPrice

                vIndex = Me.ListViewItem.Items.Count + 1

                If vQTY = 0 Then
                    MsgBox("ไม่ได้ระบุจำนวนของสินค้าที่ต้องการ หรือต้องระบุจำนวนสินค้าที่ต้องการมากกว่า 0", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                Dim n As Integer
                Dim vCheckItemCode As String
                Dim vEditQTY As Double
                Dim vEditPrice As Double
                Dim vItemAmount As Double


                If Me.ListViewItem.Items.Count > 0 Then
                    For n = 0 To Me.ListViewItem.Items.Count - 1
                        vCheckItemCode = Me.ListViewItem.Items(n).SubItems(4).Text

                        If vItemCode = vCheckItemCode Then
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
                    Me.ListViewItem.Items.Add(listItem)
                End If

                Call CalcItemAmount()

                Me.TBItem.Text = ""
                Me.TBMemBarCode.Text = ""
                Me.TBItemName.Text = ""
                Me.TBPrice.Text = ""
                Me.TBReserve.Text = ""
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
                MsgBox("ไม่มีรายการสินค้าไม่สามารถเพิ่ม รายการสินค้าลงตะกร้าได้", MsgBoxStyle.Critical, "Send Error Message")
            End If
        End If
    End Sub

    Private Sub MenuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuExit.Click
        Dim vAnswer As Integer

        vAnswer = MsgBox("คุณต้องการ เลิกใช้งานโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message ?")
        If vAnswer = 6 Then
            Application.Exit()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub MenuLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuLogIn.Click
        Me.PNDriveIn.Visible = False
        Me.PNChecker.Visible = False
        Me.MenuProgram.Enabled = False

        Me.PNLogIn.Visible = True
        Me.PNLogIn.BringToFront()
        Me.TBUserCode.Focus()
    End Sub

    Private Sub BTNLogIN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vUserCode As String
        Dim vPassWord As String
        Dim vCheckTypeLogIn As String

        vUserCode = Me.TBUserCode.Text
        vPassWord = Me.TBPassword.Text

        Dim vService As New WebReference.WebServiceCalc
        vCheckLogIn = vService.vLogIn(vUserCode, vPassWord)

        If vCheckLogIn <> "" Then
            Me.PNLogIn.Visible = False
            Me.PNDriveIn.Visible = False
            Me.PNChecker.Visible = False

            Me.MenuProgram.Enabled = True

            'vCheckTypeLogIn = Me.CMBZone.Items(Me.CMBZone.SelectedIndex).ToString
            Me.TBUserID.Text = vCheckLogIn
            Call CallIDNumber()

            If Me.RDZone1.Checked = True Then
                vConnectZone = "01"
                vCheckTypeLogIn = "จุดจ่ายที่1"
            ElseIf Me.RDZone2.Checked = True Then
                vConnectZone = "02"
                vCheckTypeLogIn = "จุดจ่ายที่2"
            ElseIf Me.RDZone3.Checked = True Then
                vConnectZone = "03"
                vCheckTypeLogIn = "จุดจ่ายที่3"
            ElseIf Me.RDZone4.Checked = True Then
                vConnectZone = "04"
                vCheckTypeLogIn = "จุดจ่ายที่4"
            End If



            If vCheckTypeLogIn <> "05-Checker" Then
                Me.PNLogIn.Visible = False
                Me.PNDriveIn.Visible = True
                Me.PNChecker.Visible = False

                Me.PNDriveIn.BringToFront()
                Me.TBRefNo.Focus()
            Else
                Me.PNLogIn.Visible = False
                Me.PNDriveIn.Visible = False
                Me.PNChecker.Visible = True

                Me.PNChecker.BringToFront()
                Me.TBSearchCheckOut.Focus()
            End If

        Else
            MsgBox("ไม่สามารถเข้าใช้งานโปรแกรมได้ กรุณาตรวจสอบชื่อและรหัสผ่าน", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBPassword.Text = ""
        End If
    End Sub

    Private Sub TBUserCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBUserCode.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBPassword.Focus()
        End If
        If e.KeyCode = Keys.Up Then
            Me.RDZone4.Focus()
        End If
        If e.KeyCode = Keys.Down Then
            Me.TBPassword.Focus()
        End If
    End Sub

    Private Sub TBPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBPassword.KeyDown
        If e.KeyCode = Keys.Up Then
            Me.TBUserCode.Focus()
        End If
    End Sub

    Private Sub MenuDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuDelete.Click
        Dim i As Integer

        i = Me.ListViewItem.FocusedItem.Index
        Me.ListViewItem.Items.RemoveAt(i)
        Call GenIDNumber()
        Call CalcItemAmount()
        Me.TBBarCode.Focus()
    End Sub
    Private Sub GenIDNumber()
        Dim i As Integer
        Dim j As Integer

        If Me.ListViewItem.Items.Count > 0 Then
            j = 0
            For i = 0 To Me.ListViewItem.Items.Count - 1
                j = j + 1
                Me.ListViewItem.Items(i).SubItems(0).Text = j
            Next
        End If
    End Sub

    Private Sub GenIDNumberCheckOut()
        Dim i As Integer
        Dim j As Integer

        If Me.ListViewCheckOut.Items.Count > 0 Then
            j = 0
            For i = 0 To Me.ListViewCheckOut.Items.Count - 1
                j = j + 1
                Me.ListViewCheckOut.Items(i).SubItems(0).Text = j
            Next
        End If
    End Sub
    Private Sub CMBZone_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.TBUserCode.Focus()
    End Sub

    Private Sub MenuPickup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPickup.Click
        Me.PNLogIn.Visible = False
        Me.PNChecker.Visible = False
        Me.PNDriveIn.Visible = True

        Me.PNDriveIn.BringToFront()
        Me.TBUserID.Text = vCheckLogIn
        Me.TBRefNo.Focus()
        Call CallIDNumber()
    End Sub

    Private Sub MenuCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCheckOut.Click
        Me.PNLogIn.Visible = False
        Me.PNDriveIn.Visible = False
        Me.PNChecker.Visible = True

        Me.PNChecker.BringToFront()
        Me.TBUserID.Text = vCheckLogIn
        Me.TBSearchCheckOut.Focus()
    End Sub

    Private Sub BTNCloseLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If vCheckLogIn = "" Then
            Application.Exit()
        Else
            Me.PNLogIn.Visible = False
            Me.MenuProgram.Enabled = True
        End If
    End Sub

    Private Sub MenuSearchPickUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuSearchPickUp.Click
        Me.PNLogIn.Visible = False
        Me.PNDriveIn.Visible = False
        Me.PNChecker.Visible = False

        Me.PNSearchPickUp.Visible = True
        Me.PNSearchPickUp.BringToFront()
        Me.TBSearchPickup.Focus()
    End Sub

    Private Sub BTNClosePickup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClosePickup.Click
        Me.TBRefNo.Text = ""
        Me.TBDocNo.Text = ""
        Me.ListViewItem.Items.Clear()
        Me.TBItemAmount.Text = ""
        Me.TBID.Enabled = True

        Me.PNDriveIn.Visible = False
    End Sub

    Private Sub TBSearchPickup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchPickup.KeyDown
        Dim vSearch As String
        Dim i As Integer
        Dim vDocno As String
        Dim vDocDate As String
        Dim vRefID As String
        Dim vPickZone As String
        Dim vAmount As Double
        Dim vIndex As Integer

        If e.KeyCode = Keys.Enter Then
            vSearch = Me.TBSearchPickup.Text
            vQuery = "exec dbo.usp_np_SearchDriveInMaster '" & vSearch & "'"
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
        End If
    End Sub

    Private Sub TBSearchPickup_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchPickup.TextChanged

    End Sub

    Private Sub BTNSearchPickup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchPickup.Click
        Dim vSearch As String
        Dim i As Integer
        Dim vDocno As String
        Dim vDocDate As String
        Dim vRefID As String
        Dim vPickZone As String
        Dim vAmount As Double
        Dim vIndex As Integer

        vSearch = Me.TBSearchPickup.Text
        vQuery = "exec dbo.usp_np_SearchDriveInMaster '" & vSearch & "'"
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
    End Sub

    Private Sub TBSearchCheckOut_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBSearchCheckOut.KeyDown
        Dim vRefNo As String
        Dim i As Integer
        Dim vDocno As String
        Dim vDocDate As String
        Dim vID As Integer
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
        Dim vAmount As Double
        Dim vIndex As Integer
        Dim vLine As Integer

        If e.KeyCode = Keys.Enter Then
            If Me.TBSearchCheckOut.Text <> "" Then
                vRefNo = Me.TBSearchCheckOut.Text
                vQuery = "exec dbo.usp_np_SearchItemPickUp '" & vRefNo & "'"
                Dim vService As New WebReference.WebServiceCalc
                Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

                Me.ListViewCheckOut.Items.Clear()

                vIndex = 0
                If ds.Tables(0).Rows.Count > 0 Then
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        vDocno = ds.Tables(0).Rows(i)("docno").ToString
                        vDocDate = ds.Tables(0).Rows(i)("docdate").ToString
                        vID = ds.Tables(0).Rows(i)("id").ToString
                        vPickZone = ds.Tables(0).Rows(i)("pickzone").ToString
                        vItemCode = ds.Tables(0).Rows(i)("itemcode").ToString
                        vItemName = ds.Tables(0).Rows(i)("itemname").ToString
                        vWHCode = ds.Tables(0).Rows(i)("whcode").ToString
                        vShelfCode = ds.Tables(0).Rows(i)("shelfcode").ToString
                        vQTY = ds.Tables(0).Rows(i)("qty").ToString
                        vPickQTY = ds.Tables(0).Rows(i)("pickqty").ToString
                        vConfirmQTY = ds.Tables(0).Rows(i)("confirmqty").ToString
                        vUnitCode = ds.Tables(0).Rows(i)("unitcode").ToString
                        vPrice = ds.Tables(0).Rows(i)("price").ToString
                        vAmount = ds.Tables(0).Rows(i)("amount").ToString

                        vIndex = vIndex + 1
                        vLine = vIndex - 1
                        Dim listItem As New ListViewItem(vIndex)
                        listItem.SubItems.Add(Format(vConfirmQTY, "##,##0.00"))
                        listItem.SubItems.Add(vItemName)
                        listItem.SubItems.Add(Format(vQTY, "##,##0.00"))
                        listItem.SubItems.Add(vUnitCode)
                        listItem.SubItems.Add(vItemCode)
                        listItem.SubItems.Add(Format(vPrice, "##,##0.00"))
                        listItem.SubItems.Add(Format(vAmount, "##,##0.00"))
                        listItem.SubItems.Add(vWHCode)
                        listItem.SubItems.Add(vShelfCode)
                        listItem.SubItems.Add(vPickZone)
                        listItem.SubItems.Add(vDocno)
                        Me.ListViewCheckOut.Items.Add(listItem)

                        If vPickZone = "01" Then
                            Me.ListViewCheckOut.Items(vLine).ForeColor = Color.DarkBlue
                        ElseIf vPickZone = "02" Then
                            Me.ListViewCheckOut.Items(vLine).ForeColor = Color.DarkGreen
                        ElseIf vPickZone = "03" Then
                            Me.ListViewCheckOut.Items(vLine).ForeColor = Color.DarkOrange
                        ElseIf vPickZone = "04" Then
                            Me.ListViewCheckOut.Items(vLine).ForeColor = Color.DarkMagenta
                        End If

                    Next

                    Call vCalcCheckOutAmount()
                    Call vCalcCheckOutKeyQuanity()
                    Me.ListViewCheckOut.Items(0).Selected = True
                    Me.ListViewCheckOut.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub MenuEditCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditCheckOut.Click
        Dim i As Integer

        i = Me.ListViewCheckOut.FocusedItem.Index

    End Sub

    Private Sub LBCloseCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LBCloseCheckOut.Click
        Me.PNChecker.Visible = False
    End Sub

    Private Sub LBLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vUserCode As String
        Dim vPassWord As String
        Dim vCheckTypeLogIn As String

        vUserCode = Me.TBUserCode.Text
        vPassWord = Me.TBPassword.Text

        Dim vService As New WebReference.WebServiceCalc
        vCheckLogIn = vService.vLogIn(vUserCode, vPassWord)

        If vCheckLogIn <> "" Then
            Me.PNLogIn.Visible = False
            Me.PNDriveIn.Visible = False
            Me.PNChecker.Visible = False

            Me.MenuProgram.Enabled = True

            'vCheckTypeLogIn = Me.CMBZone.Items(Me.CMBZone.SelectedIndex).ToString
            Me.TBUserID.Text = vCheckLogIn
            Call CallIDNumber()

            If Me.RDZone1.Checked = True Then
                vConnectZone = "01"
                vCheckTypeLogIn = "จุดจ่ายที่1"
            ElseIf Me.RDZone2.Checked = True Then
                vConnectZone = "02"
                vCheckTypeLogIn = "จุดจ่ายที่2"
            ElseIf Me.RDZone3.Checked = True Then
                vConnectZone = "03"
                vCheckTypeLogIn = "จุดจ่ายที่3"
            ElseIf Me.RDZone4.Checked = True Then
                vConnectZone = "04"
                vCheckTypeLogIn = "จุดจ่ายที่4"
            End If



            If vCheckTypeLogIn <> "05-Checker" Then
                Me.PNLogIn.Visible = False
                Me.PNDriveIn.Visible = True
                Me.PNChecker.Visible = False

                Me.PNDriveIn.BringToFront()
                Me.TBRefNo.Focus()
            Else
                Me.PNLogIn.Visible = False
                Me.PNDriveIn.Visible = False
                Me.PNChecker.Visible = True

                Me.PNChecker.BringToFront()
                Me.TBSearchCheckOut.Focus()
            End If

        Else
            MsgBox("ไม่สามารถเข้าใช้งานโปรแกรมได้ กรุณาตรวจสอบชื่อและรหัสผ่าน", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBPassword.Text = ""
        End If
    End Sub

    Private Sub LBCloseLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vAnswer As Integer

        If vCheckLogIn = "" Then
            vAnswer = MsgBox("คุณต้องการออกโปรแกรมใช่หรือไม่ ?", MsgBoxStyle.YesNo, "Send Question Information")
            If vAnswer = 6 Then
                Application.Exit()
            Else
                Exit Sub
            End If
        Else
            Me.PNLogIn.Visible = False
            Me.MenuProgram.Enabled = True
        End If
    End Sub

    Private Sub LBAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LBAddItem.Click
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
                MsgBox("ไม่ได้ระบุจำนวนของสินค้าที่ต้องการ หรือต้องระบุจำนวนสินค้าที่ต้องการมากกว่า 0", MsgBoxStyle.Critical, "Send Error Message")
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
            Me.TBReserve.Text = ""
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
            MsgBox("ไม่มีรายการสินค้าไม่สามารถเพิ่ม รายการสินค้าลงตะกร้าได้", MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub TBQTY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBQTY.TextChanged

    End Sub

    Private Sub MenuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEdit.Click
        Dim vBarCode As String
        Dim vRate As Integer
        Dim vDefShelfCode As String
        Dim vStockUnit As String
        Dim i As Integer
        Dim vStore As String
        Dim vStkQTY As Double

        vSelectLineEdit = Me.ListViewItem.FocusedItem.Index
        vBarCode = Me.ListViewItem.Items(vSelectLineEdit).SubItems(9).Text
        vDefShelfCode = Me.ListViewItem.Items(vSelectLineEdit).SubItems(8).Text
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetDataBarCode(vBarCode)
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

        Me.PNItemEdit.Visible = True
        Me.TBEditCode.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(4).Text
        Me.TBEditName.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(1).Text
        Me.TBEditUnit.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(3).Text
        Me.TBEditPrice.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(5).Text
        Me.TBEditQty.Text = Me.ListViewItem.Items(vSelectLineEdit).SubItems(2).Text
        Me.TBEditRate.Text = Format(vRate, "##,##0.00")
        Me.TBDefSaleUnitCode.Text = vDefShelfCode
        Me.TBEditQty.Focus()
        Me.TBEditQty.SelectAll()
    End Sub

    Private Sub LBItemEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LBItemEdit.Click
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double

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
    End Sub

    Private Sub LBCloseEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

        n = Me.ListViewSearhPickup.FocusedItem.Index
        vDocno = Me.ListViewSearhPickup.Items(n).SubItems(2).Text

        vQuery = "exec dbo.usp_np_SearchPickUp '" & vDocno & "'"
        Dim vService As New WebReference.WebServiceCalc
        Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

        Me.ListViewItem.Items.Clear()
        If ds.Tables(0).Rows.Count > 0 Then
            Me.TBID.Text = ds.Tables(0).Rows(i)("id").ToString
            Me.TBID.Enabled = False
            Me.TBRefNo.Text = ds.Tables(0).Rows(i)("refid").ToString
            vNetItemAmount = ds.Tables(0).Rows(i)("totalnetamount").ToString
            Me.TBItemAmount.Text = Format(vNetItemAmount, "##,##0.00")
            Me.TBDocNo.Text = ds.Tables(0).Rows(i)("docno").ToString

            vIndex = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1

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
                Me.ListViewItem.Items.Add(listItem)
            Next
        End If
        Me.ListViewSearhPickup.Items.Clear()
        Me.TBSearchPickup.Text = ""
        Me.PNSearchPickUp.Visible = False
        Me.PNDriveIn.Visible = True
    End Sub

    Private Sub BTNCloseSelectPickup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseSelectPickup.Click
        Me.ListViewSearhPickup.Items.Clear()
        Me.TBSearchPickup.Text = ""
        Me.PNSearchPickUp.Visible = False
    End Sub

    Private Sub ListViewCheckOut_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewCheckOut.KeyDown
        Dim i As Integer

        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            If Me.ListViewCheckOut.Items.Count > 0 Then
                i = Me.ListViewCheckOut.FocusedItem.Index
                If Me.ListViewCheckOut.Items(i).BackColor = Color.Red Then
                    Exit Sub
                End If
                Me.PNKeyQTY.Visible = True
                vSelectCheckOutLine = Me.ListViewCheckOut.FocusedItem.Index
                Me.TBIndex.Text = i
                Me.TBItemNameKeyQTY.Text = Me.ListViewCheckOut.Items(i).SubItems(2).Text
                Me.TBKeyPrice.Text = Me.ListViewCheckOut.Items(i).SubItems(6).Text
                Me.TBKeyQTY.Focus()
            End If
        End If
    End Sub

    Private Sub LBCloseKeyQTY_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LBCloseKeyQTY.Click
        Me.PNKeyQTY.Visible = False
    End Sub

    Private Sub LBKeyQTY_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LBKeyQTY.Click
        Me.PNKeyQTY.Visible = False
    End Sub

    Private Sub TBKeyQTY_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBKeyQTY.KeyDown
        Dim i As Integer
        Dim vCheckQTY As Double
        Dim vPayQTY As Double
        Dim vAnswer As Integer
        Dim vPickZone As String

        If Me.TBKeyQTY.Text <> "" Then
            If e.KeyCode = Keys.Enter Then
                i = Me.TBIndex.Text
                vCheckQTY = Me.TBKeyQTY.Text
                vPayQTY = Me.ListViewCheckOut.Items(i).SubItems(3).Text
                vPickZone = Me.ListViewCheckOut.Items(i).SubItems(10).Text

                If vCheckQTY <> vPayQTY Then
                    Me.ListViewCheckOut.Items(i).ForeColor = Color.Red
                    vAnswer = MsgBox("จำนวนที่จ่ายมากับจำนวนที่นับได้ของ Checker ไม่เท่ากัน ต้องการนับใหม่ใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message ?")
                    If vAnswer = 6 Then
                        Me.TBKeyQTY.SelectAll()
                        Me.TBKeyQTY.Focus()
                        Exit Sub
                    End If
                Else
                    If vPickZone = "01" Then
                        Me.ListViewCheckOut.Items(i).ForeColor = Color.DarkBlue
                    ElseIf vPickZone = "02" Then
                        Me.ListViewCheckOut.Items(i).ForeColor = Color.DarkGreen
                    ElseIf vPickZone = "03" Then
                        Me.ListViewCheckOut.Items(i).ForeColor = Color.DarkOrange
                    ElseIf vPickZone = "04" Then
                        Me.ListViewCheckOut.Items(i).ForeColor = Color.DarkMagenta
                    ElseIf vPickZone = "05" Then
                        Me.ListViewCheckOut.Items(i).ForeColor = Color.Black
                    End If
                End If
                Me.ListViewCheckOut.Items(i).SubItems(1).Text = Format(vCheckQTY, "##,##0.00")
                Me.TBKeyQTY.Text = ""
                Me.PNKeyQTY.Visible = False

                Call vCalcCheckOutKeyQuanity()
                Me.ListViewCheckOut.Focus()
            End If
        End If
        If e.KeyCode = Keys.Escape Then
            Me.PNKeyQTY.Visible = False
            Me.ListViewCheckOut.Focus()
        End If
    End Sub
    Private Sub vCalcCheckOutAmount()
        Dim i As Integer
        Dim vAmount As Double
        Dim vTotalAmount As Double

        If Me.ListViewCheckOut.Items.Count > 0 Then
            For i = 0 To Me.ListViewCheckOut.Items.Count - 1
                If Me.ListViewCheckOut.Items(i).BackColor <> Color.Red Then
                    vAmount = Me.ListViewCheckOut.Items(i).SubItems(7).Text
                    vTotalAmount = vTotalAmount + vAmount
                End If
            Next
            Me.LBLNetAmount.Text = Format(vTotalAmount, "##,##0.00")
        End If
    End Sub

    Private Sub vCalcCheckOutKeyQuanity()
        Dim i As Integer
        Dim vAmount As Double
        Dim vTotalAmount As Double
        Dim vPrice As Double
        Dim vKeyQTY As Double

        If Me.ListViewCheckOut.Items.Count > 0 Then
            For i = 0 To Me.ListViewCheckOut.Items.Count - 1
                vPrice = Me.ListViewCheckOut.Items(i).SubItems(6).Text
                vKeyQTY = Me.ListViewCheckOut.Items(i).SubItems(1).Text
                vAmount = vKeyQTY * vPrice
                vTotalAmount = vTotalAmount + vAmount
            Next
            Me.LBLCheckOutAmount.Text = Format(vTotalAmount, "##,##0.00")
        End If
    End Sub

    Private Sub MenuItemCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCancel.Click
        Dim i As Integer
        Dim vAnswer As Integer
        Dim vDocno As String
        Dim vDocdate As String
        Dim vItemcode As String
        Dim vIndex As Integer


        If Me.ListViewCheckOut.Items.Count > 0 Then
            i = Me.ListViewCheckOut.FocusedItem.Index
            vDocno = Me.ListViewCheckOut.Items(i).SubItems(11).Text
            vDocdate = Now.Day & "/" & Now.Month & "/" & Now.Year
            vItemcode = Me.ListViewCheckOut.Items(i).SubItems(5).Text
            vIndex = i + 1

            vAnswer = MsgBox("คุณต้องการยกเลิกรายการที่ " & vIndex & " ใช่หรือไม่ ?", MsgBoxStyle.YesNo, "Send Question Message ?")
            If vAnswer = 6 Then
                Me.ListViewCheckOut.Items.RemoveAt(i)
                Call GenIDNumberCheckOut()

                vQuery = "exec dbo.usp_np_updatedriveincancelcheckout '" & vDocno & "','" & vDocdate & "','" & vItemcode & "'"
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)

                Call vCalcCheckOutAmount()
                Call vCalcCheckOutKeyQuanity()
            End If
        End If
    End Sub

    Private Sub MenuItemSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemSelect.Click
        Dim i As Integer

        i = Me.ListViewCheckOut.FocusedItem.Index
        Me.ListViewCheckOut.Items(i).BackColor = Color.White
        Call vCalcCheckOutAmount()
    End Sub

    Private Sub ListViewCheckOut_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewCheckOut.SelectedIndexChanged

    End Sub

    Private Sub TBSearchCheckOut_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSearchCheckOut.TextChanged

    End Sub

    Private Sub PNSearchPickUp_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PNSearchPickUp.GotFocus

    End Sub

    Private Sub BTNCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCheckOut.Click
        Dim vDocno As String
        Dim vDocdate As String
        Dim vItemcode As String
        Dim vKeyQTY As Double

        Dim i As Integer

        If Me.ListViewCheckOut.Items.Count > 0 Then
            vDocdate = Now.Day & "/" & Now.Month & "/" & Now.Year

            For i = 0 To Me.ListViewCheckOut.Items.Count - 1
                vDocno = Me.ListViewCheckOut.Items(i).SubItems(11).Text
                vItemcode = Me.ListViewCheckOut.Items(i).SubItems(5).Text
                vKeyQTY = Me.ListViewCheckOut.Items(i).SubItems(1).Text

                vQuery = "exec dbo.usp_np_updatedriveincheckeritemcheckout '" & vDocno & "','" & vDocdate & "','" & vItemcode & "'," & vKeyQTY & " "
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)
            Next

            Me.ListViewCheckOut.Items.Clear()
            Me.LBLCheckOutAmount.Text = ""
            Me.LBLNetAmount.Text = ""
            Me.TBSearchCheckOut.Focus()
        End If
    End Sub

    Private Sub BTNGenBill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNGenBill.Click
        Dim vDocno As String
        Dim vDocdate As String
        Dim vItemcode As String
        Dim vInvQTY As Double
        Dim vUserID As String
        Dim vAmount As Double
        Dim vPosNo As String

        Dim vMachineNo As String
        Dim vMaxNo As Integer
        Dim vHeader As String


        Dim i As Integer

        If Me.ListViewCheckOut.Items.Count > 0 Then
            vDocdate = Now.Day & "/" & Now.Month & "/" & Now.Year
            vMachineNo = "01"

            vQuery = "exec dbo.USP_NP_GetMaxNoHoldingBill '" & vMachineNo & "','" & vDocdate & "'"
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)

            If ds.Tables(0).Rows.Count > 0 Then
                vMaxNo = ds.Tables(0).Rows(0)("maxnumber").ToString
                vHeader = ds.Tables(0).Rows(0)("header").ToString
            End If

            vPosNo = vHeader + "-" + Format(vMaxNo, "0000")

            For i = 0 To Me.ListViewCheckOut.Items.Count - 1
                vDocno = Me.ListViewCheckOut.Items(i).SubItems(11).Text
                vItemcode = Me.ListViewCheckOut.Items(i).SubItems(5).Text
                vInvQTY = Me.ListViewCheckOut.Items(i).SubItems(1).Text

                vQuery = "exec dbo.usp_np_driveincheckoutpos '" & vDocno & "','" & vDocdate & "','" & vPosNo & "','" & vCheckLogIn & "','" & vItemcode & "'," & vInvQTY & " "
                Dim vService1 As New WebReference.WebServiceCalc
                Dim ds1 As Integer = vService1.vExecuteQuery(vQuery)
            Next

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

            vExpireCredit = 1
            vArCode = "999999"

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
            vSHIFTCODE = "กลางวัน"

            vQuery = "exec dbo.USP_NP_InsertHoldingBillDriveIn '" & vPosNo & "','" & vDocdate & "'," & vExpireCredit & ",'" & vArCode & "','" & vCashierCode & "','" & vMachineNo & "','" & vMachineCode & "','" & vSaleCode & "'," & vTaxRate & "," & vTaxRate & "," & vAfterDiscount & "," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalAmount & "," & vNetDebtAmount & ",'" & vCreatorCode & "','" & vSHIFTCODE & "' "
            Dim vService3 As New WebReference.WebServiceCalc
            Dim ds3 As Integer = vService3.vExecuteQuery(vQuery)

            Dim n As Integer


            Dim vWHCode As String
            Dim vShelfCode As String
            Dim vQTY As Double
            Dim vPrice As Double
            Dim vItemAmount As Double
            Dim vNetAmount As Double
            Dim vUnitCode As String
            Dim vStockType As Integer
            Dim vLineNumber As Integer
            Dim vBarCode As String
            Dim vPosStatus As Integer

            For n = 0 To Me.ListViewCheckOut.Items.Count - 1
                vItemcode = Me.ListViewCheckOut.Items(n).SubItems(5).Text
                vWHCode = Me.ListViewCheckOut.Items(n).SubItems(8).Text
                vShelfCode = Me.ListViewCheckOut.Items(n).SubItems(9).Text
                vQTY = Me.ListViewCheckOut.Items(n).SubItems(1).Text
                vPrice = Me.ListViewCheckOut.Items(n).SubItems(6).Text
                vItemAmount = Me.ListViewCheckOut.Items(n).SubItems(5).Text
                vNetAmount = Me.ListViewCheckOut.Items(n).SubItems(5).Text
                vUnitCode = Me.ListViewCheckOut.Items(n).SubItems(4).Text
                vStockType = 0
                vLineNumber = n
                vBarCode = Me.ListViewCheckOut.Items(n).SubItems(5).Text
                vPosStatus = 0

                vQuery = "exec dbo.USP_NP_InsertHoldingBillDriveInSub '" & vDocno & "','" & vDocdate & "','" & vItemcode & "','" & vWHCode & "','" & vShelfCode & "'," & vQTY & "," & vPrice & "," & vItemAmount & "," & vNetAmount & ",'" & vUnitCode & "'," & vStockType & "," & vLineNumber & ",'" & vBarCode & "','" & vCashierCode & "'," & vPosStatus & " "
                Dim vService4 As New WebReference.WebServiceCalc
                Dim ds4 As Integer = vService4.vExecuteQuery(vQuery)
            Next

            MsgBox("ได้เลขที่พักบิลเลขที่ " & vPosNo & "", MsgBoxStyle.Information, "Send Information Message")
            Me.ListViewCheckOut.Items.Clear()
            Me.LBLCheckOutAmount.Text = ""
            Me.LBLNetAmount.Text = ""
            Me.TBSearchCheckOut.Focus()
        End If
    End Sub

    Private Sub BTNSearchCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchCheckOut.Click

    End Sub

    Private Sub TBKeyQTY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBKeyQTY.TextChanged

    End Sub

    Private Sub BTNClearCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearCheckOut.Click
        Me.ListViewCheckOut.Items.Clear()
        Me.LBLCheckOutAmount.Text = ""
        Me.LBLNetAmount.Text = ""
        Me.TBSearchCheckOut.Text = ""
        Me.TBSearchCheckOut.Focus()
    End Sub

    Private Sub TBEditQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBEditQty.KeyDown
        Dim vQTY As Double
        Dim vPrice As Double
        Dim vAmount As Double
        Dim vUnitCode As String

        Dim vShelfUnit As String
        Dim vShelfQTY As Double
        Dim vTotalQTY As Double
        Dim vRate As Integer


        If e.KeyCode = Keys.Escape Then
            Me.PNItemEdit.Visible = False
        End If

        If e.KeyCode = Keys.Enter Then
            If Me.TBEditQty.Text <> "" Then
                vQTY = Me.TBEditQty.Text
            End If

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
                    MsgBox("ไม่สามารถขายเกินยอดที่มีใน STOCK ได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                    Me.TBEditQty.SelectAll()
                    Exit Sub
                End If
            End If

            If vQTY > vShelfQTY And vShelfUnit = vUnitCode Then
                MsgBox("ไม่สามารถขายเกินยอดที่มีใน STOCK ได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBEditQty.SelectAll()
                Exit Sub
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
        End If
    End Sub

    Private Sub TBEditQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBEditQty.TextChanged

    End Sub

    Private Sub TBItem_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBItem.TextChanged

    End Sub

    Private Sub MenuAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuAdd.Click
        Me.PNAddItem.Visible = True
        Me.PNAddItem.BringToFront()
        Me.TBSearchBarCode.Focus()
    End Sub

    Private Sub LBCloseAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LBCloseAddItem.Click
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
                        MsgBox("ไม่สามารถขายเกินยอดที่มีใน STOCK ได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message ")
                        Me.TBAddQTY.SelectAll()
                        Exit Sub
                    End If
                End If

                If vQTY > vShelfQTY And vShelfUnit = vUnitCode Then
                    MsgBox("ไม่สามารถขายเกินยอดที่มีใน STOCK ได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message ")
                    Me.TBAddQTY.SelectAll()
                    Exit Sub
                End If

                If Me.TBAddPrice.Text <> "" Then
                    vPrice = Me.TBAddPrice.Text
                End If
                vAmount = vQTY * vPrice

                vIndex = Me.ListViewCheckOut.Items.Count + 1

                If vQTY = 0 Then
                    MsgBox("ไม่ได้ระบุจำนวนของสินค้าที่ต้องการ หรือต้องระบุจำนวนสินค้าที่ต้องการมากกว่า 0", MsgBoxStyle.Critical, "Send Error Message")
                    Exit Sub
                End If

                Dim n As Integer
                Dim vCheckItemCode As String
                Dim vEditQTY As Double
                Dim vEditPrice As Double
                Dim vItemAmount As Double
                Dim vOldQty As Double
                Dim vPickZone As String


                If Me.ListViewCheckOut.Items.Count > 0 Then
                    For n = 0 To Me.ListViewCheckOut.Items.Count - 1
                        vCheckItemCode = Me.ListViewCheckOut.Items(n).SubItems(5).Text

                        If vItemCode = vCheckItemCode Then
                            vEditPrice = Me.TBAddPrice.Text
                            vEditQTY = Me.TBAddQTY.Text
                            vItemAmount = vEditQTY * vEditPrice

                            vOldQty = Me.ListViewCheckOut.Items(n).SubItems(3).Text
                            vPickZone = Me.ListViewCheckOut.Items(n).SubItems(10).Text

                            If vEditQTY = vOldQty Then
                                If vPickZone = "01" Then
                                    Me.ListViewCheckOut.Items(n).ForeColor = Color.DarkBlue
                                ElseIf vPickZone = "02" Then
                                    Me.ListViewCheckOut.Items(n).ForeColor = Color.DarkGreen
                                ElseIf vPickZone = "03" Then
                                    Me.ListViewCheckOut.Items(n).ForeColor = Color.DarkOrange
                                ElseIf vPickZone = "04" Then
                                    Me.ListViewCheckOut.Items(n).ForeColor = Color.DarkMagenta
                                ElseIf vPickZone = "05" Then
                                    Me.ListViewCheckOut.Items(n).ForeColor = Color.Black
                                End If
                            Else
                                Me.ListViewCheckOut.Items(n).ForeColor = Color.Red
                            End If

                            Me.ListViewCheckOut.Items(n).SubItems(1).Text = Format(vEditQTY, "##,##0.00")
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
                    listItem.SubItems.Add("05")
                    listItem.SubItems.Add("CheckerAdd")
                    Me.ListViewCheckOut.Items.Add(listItem)
                End If

                Call vCalcCheckOutAmount()
                Call vCalcCheckOutKeyQuanity()

                Me.TBAddItemCode.Text = ""
                Me.TBAddItemBar.Text = ""
                Me.TBAddItemName.Text = ""
                Me.TBAddPrice.Text = ""
                Me.TBAddReserveQTY.Text = ""
                Me.TBAddItemUnit.Text = ""
                Me.TBAddDefWHCode.Text = ""
                Me.TBAddDefShelf.Text = ""
                Me.TBAddQTY.Text = ""
                Me.ListViewAddStockQTY.Items.Clear()
                Me.PNAddItem.Visible = False
                Me.TBSearchBarCode.Text = ""
                Me.TBSearchBarCode.Focus()
            Else
                MsgBox("ไม่มีรายการสินค้าไม่สามารถเพิ่ม รายการสินค้าลงตะกร้าได้", MsgBoxStyle.Critical, "Send Error Message")
            End If
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

        If e.KeyCode = Keys.Enter Then
            If Me.TBSearchBarCode.Text <> "" Then
                vBarCode = Me.TBSearchBarCode.Text
            Else
                Me.TBSearchBarCode.Focus()
            End If

            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetDataBarCode(vBarCode)
            Me.ListViewAddStockQTY.Items.Clear()

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

                Dim n As Integer
                Dim vCheckItemCode As String
                Dim vCheckQTY As Double

                If Me.ListViewCheckOut.Items.Count > 0 Then
                    For n = 0 To Me.ListViewCheckOut.Items.Count - 1
                        vCheckItemCode = Me.ListViewCheckOut.Items(n).SubItems(5).Text
                        vCheckQTY = Me.ListViewCheckOut.Items(n).SubItems(1).Text
                        If vItemCode = vCheckItemCode Then
                            Me.TBAddQTY.Text = Format(vCheckQTY, "##,##0.00")
                            GoTo Line1
                        End If
                    Next
                End If

Line1:
                Me.TBAddQTY.Focus()
                Me.TBAddQTY.SelectAll()
            Else
                Me.TBSearchBarCode.Focus()
                Me.TBAddQTY.SelectAll()
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
    End Sub

    Private Sub TBAddQTY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBAddQTY.TextChanged

    End Sub

    Private Sub BTNClearPickUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearPickUp.Click
        Me.TBRefNo.Text = ""
        Me.TBBarCode.Text = ""
        Me.TBItemAmount.Text = ""
        Me.ListViewItem.Items.Clear()
    End Sub

    Private Sub RDZone1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RDZone1.CheckedChanged

    End Sub

    Private Sub RDZone1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDZone1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBUserCode.Focus()
        End If
    End Sub

    Private Sub RDZone2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RDZone2.CheckedChanged

    End Sub

    Private Sub RDZone2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDZone2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBUserCode.Focus()
        End If
    End Sub

    Private Sub RDZone3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RDZone3.CheckedChanged

    End Sub

    Private Sub RDZone3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDZone3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBUserCode.Focus()
        End If
    End Sub

    Private Sub RDZone4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RDZone4.CheckedChanged

    End Sub

    Private Sub RDZone4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RDZone4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBUserCode.Focus()
        End If
    End Sub

    Private Sub TBPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPassword.TextChanged
        Dim vLenPassword As Integer
        Dim vUserCode As String
        Dim vPassWord As String
        Dim vCheckTypeLogIn As String


        vLenPassword = Len(Me.TBPassword.Text)
        If vLenPassword = 4 And Me.TBUserCode.Text <> "" Then

            Me.TBPassword.Visible = False
            vUserCode = Me.TBUserCode.Text
            vPassWord = Me.TBPassword.Text

            Dim vService As New WebReference.WebServiceCalc
            vCheckLogIn = vService.vLogIn(vUserCode, vPassWord)

            If vCheckLogIn <> "" Then
                Me.PNLogIn.Visible = False
                Me.PNChecker.Visible = False
                Me.MenuProgram.Enabled = True

                Me.TBUserID.Text = vCheckLogIn
                Call CallIDNumber()

                If Me.RDZone1.Checked = True Then
                    vConnectZone = "01"
                    vCheckTypeLogIn = "จุดจ่ายที่1"
                ElseIf Me.RDZone2.Checked = True Then
                    vConnectZone = "02"
                    vCheckTypeLogIn = "จุดจ่ายที่2"
                ElseIf Me.RDZone3.Checked = True Then
                    vConnectZone = "03"
                    vCheckTypeLogIn = "จุดจ่ายที่3"
                ElseIf Me.RDZone4.Checked = True Then
                    vConnectZone = "04"
                    vCheckTypeLogIn = "จุดจ่ายที่4"
                End If

                Me.PNDriveIn.Visible = True
                Me.PNDriveIn.BringToFront()
                Me.TBRefNo.Focus()
            Else
                MsgBox("ไม่สามารถเข้าใช้งานโปรแกรมได้ กรุณาตรวจสอบชื่อและรหัสผ่าน", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBPassword.Visible = True
                Me.TBPassword.Text = ""
                Me.TBPassword.Focus()
            End If
        End If
    End Sub

    Private Sub TBUserCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBUserCode.TextChanged

    End Sub

    Private Sub TBRefNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBRefNo.TextChanged

    End Sub

    Private Sub ListViewItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListViewItem.KeyDown
        If e.KeyCode = 134 Then
            Call SavePickUp()
        End If
    End Sub

    Private Sub ListViewItem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewItem.SelectedIndexChanged

    End Sub
End Class