Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class FormUpdateItemPrice
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand
    Dim vCheckItemComplete As Integer 'เก็บข้อมูลว่าสินค้ามีระดับราคาครบไหมหรือว่าเกินไหม 0 คือ ครบ 1 คือ มีปัญหา
    Dim vIsOpen As Integer 'ตรวจสอบว่าเป็นเอกสารเก่าหรือเอกสารสร้างใหม่
    Dim vReadQuery As SqlDataReader
    Dim vChecCountDocno As Integer
    'Dim frmReportItemChangePrice As New FormReportItemChangePrice
    Dim vIsConfirm As Integer
    Dim vCheckItemExist As Integer

    Private Sub CMBUnitCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBUnitCode.SelectedIndexChanged
        Dim vItemCode As String
        Dim vUnitCode As String
        Dim i As Integer
        Dim n As Integer
        Dim vSaleType As Integer
        Dim vTranSportType As Integer

        On Error GoTo ErrDescription

        If Me.TextItemCode.Text <> "" Then
            vItemCode = Trim(Me.TextItemCode.Text)
            vUnitCode = Trim(Me.CMBUnitCode.Text)
            vCheckItemComplete = 0
            vQuery = "exec dbo.USP_NP_CheckLevelPrice '" & vItemCode & "','" & vUnitCode & "'" 'ตรวจสอบระดับราคาว่ามีเกินกว่าที่กำหนดไหม
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "CheckLevelPrice")
            dt = ds.Tables("CheckLevelPrice")
            If dt.Rows.Count <= 4 Then ' มาตรฐานของการกำหนดราคาจะมีอยู่ 4 ระดับ กรณีที่มี 4 ระดับจะเข้าส่วนนี้
                For i = 0 To dt.Rows.Count - 1
                    If dt.Rows(i).Item("vCountLine") > 1 Then
                        If i = 0 Then
                            MsgBox("ระดับราคาขายสดรับเอง มีมากกว่า 1 ราคา กรุณาตรวจสอบก่อน เพราะจะไม่สามารถใช้งานโปรแกรมนี้ได้", MsgBoxStyle.Critical, "Send Error")
                            vCheckItemComplete = 1
                        ElseIf i = 1 Then
                            MsgBox("ระดับราคาขายสดส่งให้ มีมากกว่า 1 ราคา กรุณาตรวจสอบก่อน เพราะจะไม่สามารถใช้งานโปรแกรมนี้ได้", MsgBoxStyle.Critical, "Send Error")
                            vCheckItemComplete = 1
                        ElseIf i = 2 Then
                            MsgBox("ระดับราคาขายเชื่อรับเอง มีมากกว่า 1 ราคา กรุณาตรวจสอบก่อน เพราะจะไม่สามารถใช้งานโปรแกรมนี้ได้", MsgBoxStyle.Critical, "Send Error")
                            vCheckItemComplete = 1
                        ElseIf i = 3 Then
                            MsgBox("ระดับราคาขายเชื่อส่งให้ มีมากกว่า 1 ราคา กรุณาตรวจสอบก่อน เพราะจะไม่สามารถใช้งานโปรแกรมนี้ได้", MsgBoxStyle.Critical, "Send Error")
                            vCheckItemComplete = 1
                        End If
                    End If
                Next

                If vCheckItemComplete = 0 Then ' ตรวจสอบก่อนว่าข้อมูลครบไหม ถึงจะมาดึงข้อมูลมาแสดงในช่องต่าง ๆ ให้
                    vQuery = "exec dbo.USP_NP_SearchItemPriceDetails '" & vItemCode & "','" & vUnitCode & "'"
                    da = New SqlDataAdapter(vQuery, vConnection)
                    ds = New DataSet
                    da.Fill(ds, "PriceDetails")
                    dt = ds.Tables("PriceDetails")
                    For n = 0 To dt.Rows.Count - 1
                        vSaleType = dt.Rows(n).Item("saletype")
                        vTranSportType = dt.Rows(n).Item("transporttype")
                        If vSaleType = 0 And vTranSportType = 0 Then
                            Me.LBLOldCash01.Text = Format(Int(dt.Rows(n).Item("saleprice1")), "##,##0.00")
                            Me.LBLOldCash02.Text = Format(Int(dt.Rows(n).Item("saleprice2")), "##,##0.00")
                        End If
                        If vSaleType = 0 And vTranSportType = 1 Then
                            Me.LBLOldCash11.Text = Format(Int(dt.Rows(n).Item("saleprice1")), "##,##0.00")
                            Me.LBLOldCash12.Text = Format(Int(dt.Rows(n).Item("saleprice2")), "##,##0.00")
                        End If
                        If vSaleType = 1 And vTranSportType = 0 Then
                            Me.LBLOldCredit01.Text = Format(Int(dt.Rows(n).Item("saleprice1")), "##,##0.00")
                            Me.LBLOldCredit02.Text = Format(Int(dt.Rows(n).Item("saleprice2")), "##,##0.00")
                        End If
                        If vSaleType = 1 And vTranSportType = 1 Then
                            Me.LBLOldCredit11.Text = Format(Int(dt.Rows(n).Item("saleprice1")), "##,##0.00")
                            Me.LBLOldCredit12.Text = Format(Int(dt.Rows(n).Item("saleprice2")), "##,##0.00")
                        End If
                    Next
                End If
            ElseIf dt.Rows.Count < 4 Then 'มีระดับราคาไม่ครบ
                MsgBox("ระดับราคาสินค้ามีไม่ครบ ตามที่กำหนดให้ไว้ กรุณาตรวจสอบ สินค้ารหัส " & vItemCode & " ", MsgBoxStyle.Critical, "Send Error")
                Exit Sub

            ElseIf dt.Rows.Count > 4 Then 'มีระดับราคาเกิน
                MsgBox("ระดับราคาสินค้ามีเกินกว่าที่กำหนดให้ไว้ กรุณาตรวจสอบ สินค้ารหัส " & vItemCode & " ", MsgBoxStyle.Critical, "Send Error")
                Exit Sub
            End If
        Else
        MsgBox("กรุณากรอกรหัสสินค้าที่จะปรับราคาด้วย", MsgBoxStyle.Critical, "Send Error")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub FormUpdateItemPrice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
        DTPDocDate.Text = Date.Now
        DTPDueDate.Text = DateAdd(DateInterval.Day, 1, Date.Now)
        Me.Pic101.Visible = True
        Me.Pic102.Visible = False
    End Sub

    Private Sub TextItemCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextItemCode.KeyDown
        Dim vItemCode As String
        Dim i As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Call ClearData()
            vItemCode = Trim(TextItemCode.Text)
            vQuery = "select code,name1 from dbo.bcitem where code = '" & vItemCode & "' "
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "KeyItemCode")
            dt = ds.Tables("KeyItemCode")
            If dt.Rows.Count > 0 Then
                Me.LBLItemName.Text = dt.Rows(0).Item("name1")
            Else
                MsgBox("ไม่มีรหัสสินค้า รหัส " & vItemCode & " นี้ในระบบ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error")
                Exit Sub
            End If

            vQuery = "select distinct itemcode,unitcode from dbo.bcpricelist where itemcode = '" & vItemCode & "' order by unitcode"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "UnitCode")
            dt = ds.Tables("UnitCode")
            For i = 0 To dt.Rows.Count - 1
                CMBUnitCode.Items.Add(Trim(dt.Rows(i).Item("unitcode")))
            Next
            Me.CMBUnitCode.Focus()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub ClearData()
        Me.TextCash01.Text = ""
        Me.TextCash02.Text = ""
        Me.TextCash11.Text = ""
        Me.TextCash12.Text = ""
        Me.TextCredit01.Text = ""
        Me.TextCredit02.Text = ""
        Me.TextCredit11.Text = ""
        Me.TextCredit12.Text = ""
        Me.LBLOldCash01.Text = ""
        Me.LBLOldCash02.Text = ""
        Me.LBLOldCash11.Text = ""
        Me.LBLOldCash12.Text = ""
        Me.LBLOldCredit01.Text = ""
        Me.LBLOldCredit02.Text = ""
        Me.LBLOldCredit11.Text = ""
        Me.LBLOldCredit12.Text = ""
        Me.LBLItemName.Text = ""
        Me.CMBUnitCode.Items.Clear()
        Me.CHK101.Checked = False
        Me.CHK102.Checked = False
        Me.CHK103.Checked = False
        Me.CHK104.Checked = False
        Me.TextCash01.Enabled = False
        Me.TextCash02.Enabled = False
        Me.TextCash11.Enabled = False
        Me.TextCash12.Enabled = False
        Me.TextCredit01.Enabled = False
        Me.TextCredit02.Enabled = False
        Me.TextCredit11.Enabled = False
        Me.TextCredit12.Enabled = False

    End Sub

    Private Sub ClearScreen()
        Me.TextCash01.Text = ""
        Me.TextCash02.Text = ""
        Me.TextCash11.Text = ""
        Me.TextCash12.Text = ""
        Me.TextCredit01.Text = ""
        Me.TextCredit02.Text = ""
        Me.TextCredit11.Text = ""
        Me.TextCredit12.Text = ""
        Me.LBLOldCash01.Text = ""
        Me.LBLOldCash02.Text = ""
        Me.LBLOldCash11.Text = ""
        Me.LBLOldCash12.Text = ""
        Me.LBLOldCredit01.Text = ""
        Me.LBLOldCredit02.Text = ""
        Me.LBLOldCredit11.Text = ""
        Me.LBLOldCredit12.Text = ""
        Me.LBLItemName.Text = ""
        Me.CMBUnitCode.Items.Clear()
        Me.CHK101.Checked = False
        Me.CHK102.Checked = False
        Me.CHK103.Checked = False
        Me.CHK104.Checked = False
        Me.TextCash01.Enabled = False
        Me.TextCash02.Enabled = False
        Me.TextCash11.Enabled = False
        Me.TextCash12.Enabled = False
        Me.TextCredit01.Enabled = False
        Me.TextCredit02.Enabled = False
        Me.TextCredit11.Enabled = False
        Me.TextCredit12.Enabled = False
        vIsOpen = 0
        vIsConfirm = 0
        Me.TextDocNo.Text = ""
        Me.DTPDocDate.Text = Date.Now
        Me.DTPDueDate.Text = DateAdd(DateInterval.Day, 1, Date.Now)
        Me.ListView101.Items.Clear()
        Me.Pic101.Visible = True
        Me.Pic102.Visible = False
        Me.BTNGenNumber.Enabled = True
    End Sub

    Private Sub TextCash01_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextCash01.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TextCash02.Focus()
        End If
    End Sub

    Private Sub TextCash11_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextCash11.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TextCash12.Focus()
        End If
    End Sub

    Private Sub TextCredit01_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextCredit01.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TextCredit02.Focus()
        End If
    End Sub

    Private Sub TextCredit11_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextCredit11.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.TextCredit12.Focus()
        End If
    End Sub

    Private Sub TextCash02_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextCash02.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.CHK102.Focus()
        End If
    End Sub

    Private Sub TextCash12_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextCash12.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.CHK103.Focus()
        End If
    End Sub

    Private Sub TextCredit02_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextCredit02.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.CHK104.Focus()
        End If
    End Sub

    Private Sub TextCredit12_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextCredit12.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.Enter Then
            Me.BTNInsertBasket.Focus()
        End If
    End Sub

    Private Sub TextCash01_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextCash01.KeyPress, TextCash02.KeyPress, TextCash11.KeyPress, TextCash12.KeyPress, TextCredit01.KeyPress, TextCredit02.KeyPress, TextCredit11.KeyPress, TextCredit12.KeyPress
        On Error Resume Next

        Select Case Asc(e.KeyChar)
            Case 47 To 58, 8, 44, 46
            Case Else
                e.Handled = True
        End Select
    End Sub

    Private Sub CHK101_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHK101.Click, CHK102.Click, CHK103.Click, CHK104.Click
        On Error GoTo ErrDescription

        If Me.TextItemCode.Text <> "" And Me.LBLItemName.Text <> "" Then
            If CHK101.Checked = False Then
                Me.TextCash01.Enabled = False
                Me.TextCash02.Enabled = False
                Me.TextCash01.Text = ""
                Me.TextCash02.Text = ""
            ElseIf CHK101.Checked = True Then
                Me.TextCash01.Enabled = True
                Me.TextCash02.Enabled = True
                Me.TextCash01.Focus()
            End If

            If CHK102.Checked = False Then
                Me.TextCash11.Enabled = False
                Me.TextCash12.Enabled = False
                Me.TextCash11.Text = ""
                Me.TextCash12.Text = ""
            ElseIf CHK102.Checked = True Then
                Me.TextCash11.Enabled = True
                Me.TextCash12.Enabled = True
                Me.TextCash11.Focus()
            End If

            If CHK103.Checked = False Then
                Me.TextCredit01.Enabled = False
                Me.TextCredit02.Enabled = False
                Me.TextCredit01.Text = ""
                Me.TextCredit02.Text = ""
            ElseIf CHK103.Checked = True Then
                Me.TextCredit01.Enabled = True
                Me.TextCredit02.Enabled = True
                Me.TextCredit01.Focus()
            End If

            If CHK104.Checked = False Then
                Me.TextCredit11.Enabled = False
                Me.TextCredit12.Enabled = False
                Me.TextCredit11.Text = ""
                Me.TextCredit12.Text = ""
            ElseIf CHK104.Checked = True Then
                Me.TextCredit11.Enabled = True
                Me.TextCredit12.Enabled = True
                Me.TextCredit11.Focus()
            End If
        Else
            Me.CHK101.Checked = False
            Me.CHK102.Checked = False
            Me.CHK103.Checked = False
            Me.CHK104.Checked = False
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClear.Click
        Call ClearData()
        Me.TextItemCode.Focus()
    End Sub

    '    Private Sub BTNInsertBasket_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNInsertBasket.Click
    '        Dim vListItemCode As ListViewItem
    '        Dim vItemCode As String
    '        Dim vItemName As String
    '        Dim vUnitCode As String
    '        Dim vOldPrice1 As Object
    '        Dim vNewPrice1 As Object
    '        Dim vOldPrice2 As Object
    '        Dim vNewPrice2 As Object
    '        Dim i As Integer
    '        Dim vPriceLevel As String
    '        Dim vSaleType As String
    '        Dim vCheckItem As String
    '        Dim vCheckType As String
    '        Dim vCheckLevel As String
    '        Dim vCheckUnitCode As String
    '        Dim vScheduleDate As String
    '        Dim vLevel As Integer
    '        Dim vType As Integer
    '        Dim vTransType As Integer
    '        Dim vDocno As String

    '        On Error GoTo ErrDescription

    '        If vIsConfirm = 0 Then
    '            vItemCode = Trim(Me.TextItemCode.Text)
    '            vUnitCode = Trim(Me.CMBUnitCode.Text)
    '            vItemName = Trim(Me.LBLItemName.Text)
    '            vScheduleDate = Me.DTPDueDate.Text

    '            If ListView101.Items.Count > 0 Then
    '                If CHK101.Checked = True Then
    '                    vSaleType = "ขายสดรับเอง"
    '                    vType = 0
    '                    vTransType = 0
    '                    If Me.TextCash01.Text <> "" Then
    '                        vPriceLevel = "ราคาที่1"
    '                        vLevel = 1

    '                        vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "CheckCount")
    '                        dt = ds.Tables("CheckCount")
    '                        If dt.Rows.Count > 0 Then
    '                            vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                        Else
    '                            vCheckItemExist = 0
    '                        End If

    '                        If vCheckItemExist > 0 Then
    '                            vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                            da = New SqlDataAdapter(vQuery, vConnection)
    '                            ds = New DataSet
    '                            da.Fill(ds, "Docno")
    '                            dt = ds.Tables("Docno")
    '                            If dt.Rows.Count > 0 Then
    '                                vDocno = Trim(dt.Rows(0).Item("docno"))
    '                            End If
    '                            vDocno = ""
    '                            MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **สดรับเอง** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                            Exit Sub
    '                        End If

    '                        For i = 0 To ListView101.Items.Count - 1
    '                            vCheckItem = ListView101.Items(i).SubItems(0).Text
    '                            vCheckLevel = ListView101.Items(i).SubItems(2).Text
    '                            vCheckType = ListView101.Items(i).SubItems(3).Text
    '                            vCheckUnitCode = ListView101.Items(i).SubItems(6).Text
    '                            If vItemCode = vCheckItem And vPriceLevel = vCheckLevel And vSaleType = vCheckType And vUnitCode = vCheckUnitCode Then
    '                                MsgBox("มีการปรับราคาสินค้า " & vItemCode & " ในระดับราคาดังกล่าวอยู่แล้ว", MsgBoxStyle.Critical, "Send Error")
    '                                Exit Sub
    '                            End If
    '                        Next
    '                    End If
    '                    If TextCash02.Text <> "" Then
    '                        vPriceLevel = "ราคาที่2"
    '                        vLevel = 2

    '                        vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "CheckCount")
    '                        dt = ds.Tables("CheckCount")
    '                        If dt.Rows.Count > 0 Then
    '                            vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                        Else
    '                            vCheckItemExist = 0
    '                        End If

    '                        If vCheckItemExist > 0 Then
    '                            vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                            da = New SqlDataAdapter(vQuery, vConnection)
    '                            ds = New DataSet
    '                            da.Fill(ds, "Docno")
    '                            dt = ds.Tables("Docno")
    '                            If dt.Rows.Count > 0 Then
    '                                vDocno = Trim(dt.Rows(0).Item("docno"))
    '                            End If
    '                            'MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **สดรับเอง** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                            Exit Sub
    '                        End If

    '                        For i = 0 To ListView101.Items.Count - 1
    '                            vCheckItem = ListView101.Items(i).SubItems(0).Text
    '                            vCheckLevel = ListView101.Items(i).SubItems(2).Text
    '                            vCheckType = ListView101.Items(i).SubItems(3).Text
    '                            vCheckUnitCode = ListView101.Items(i).SubItems(6).Text
    '                            If vItemCode = vCheckItem And vPriceLevel = vCheckLevel And vSaleType = vCheckType And vUnitCode = vCheckUnitCode Then
    '                                MsgBox("มีการปรับราคาสินค้า " & vItemCode & " ในระดับราคาดังกล่าวอยู่แล้ว", MsgBoxStyle.Critical, "Send Error")
    '                                Exit Sub
    '                            End If
    '                        Next
    '                    End If
    '                End If
    '                If CHK102.Checked = True Then
    '                    vSaleType = "ขายสดส่งให้"
    '                    vType = 0
    '                    vTransType = 1
    '                    If Me.TextCash11.Text <> "" Then
    '                        vPriceLevel = "ราคาที่1"
    '                        vLevel = 1

    '                        vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "CheckCount")
    '                        dt = ds.Tables("CheckCount")
    '                        If dt.Rows.Count > 0 Then
    '                            vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                        Else
    '                            vCheckItemExist = 0
    '                        End If

    '                        If vCheckItemExist > 0 Then
    '                            vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                            da = New SqlDataAdapter(vQuery, vConnection)
    '                            ds = New DataSet
    '                            da.Fill(ds, "Docno")
    '                            dt = ds.Tables("Docno")
    '                            If dt.Rows.Count > 0 Then
    '                                vDocno = Trim(dt.Rows(0).Item("docno"))
    '                            End If
    '                            MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **สดส่งให้** นี้  ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                            Exit Sub
    '                        End If

    '                        For i = 0 To ListView101.Items.Count - 1
    '                            vCheckItem = ListView101.Items(i).SubItems(0).Text
    '                            vCheckLevel = ListView101.Items(i).SubItems(2).Text
    '                            vCheckType = ListView101.Items(i).SubItems(3).Text
    '                            vCheckUnitCode = ListView101.Items(i).SubItems(6).Text
    '                            If vItemCode = vCheckItem And vPriceLevel = vCheckLevel And vSaleType = vCheckType And vUnitCode = vCheckUnitCode Then
    '                                MsgBox("มีการปรับราคาสินค้า " & vItemCode & " ในระดับราคาดังกล่าวอยู่แล้ว", MsgBoxStyle.Critical, "Send Error")
    '                                Exit Sub
    '                            End If
    '                        Next
    '                    End If
    '                    If TextCash12.Text <> "" Then
    '                        vPriceLevel = "ราคาที่2"
    '                        vLevel = 2

    '                        vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "CheckCount")
    '                        dt = ds.Tables("CheckCount")
    '                        If dt.Rows.Count > 0 Then
    '                            vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                        Else
    '                            vCheckItemExist = 0
    '                        End If

    '                        If vCheckItemExist > 0 Then
    '                            vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                            da = New SqlDataAdapter(vQuery, vConnection)
    '                            ds = New DataSet
    '                            da.Fill(ds, "Docno")
    '                            dt = ds.Tables("Docno")
    '                            If dt.Rows.Count > 0 Then
    '                                vDocno = Trim(dt.Rows(0).Item("docno"))
    '                            End If
    '                            '------------------------------
    '                            Exit Sub
    '                        End If

    '                        For i = 0 To ListView101.Items.Count - 1
    '                            vCheckItem = ListView101.Items(i).SubItems(0).Text
    '                            vCheckLevel = ListView101.Items(i).SubItems(2).Text
    '                            vCheckType = ListView101.Items(i).SubItems(3).Text
    '                            vCheckUnitCode = ListView101.Items(i).SubItems(6).Text
    '                            If vItemCode = vCheckItem And vPriceLevel = vCheckLevel And vSaleType = vCheckType And vUnitCode = vCheckUnitCode Then
    '                                MsgBox("มีการปรับราคาสินค้า " & vItemCode & " ในระดับราคาดังกล่าวอยู่แล้ว", MsgBoxStyle.Critical, "Send Error")
    '                                Exit Sub
    '                            End If
    '                        Next
    '                    End If
    '                End If
    '                If CHK103.Checked = True Then
    '                    vSaleType = "ขายเชื่อรับเอง"
    '                    vType = 1
    '                    vTransType = 0
    '                    If Me.TextCredit01.Text <> "" Then
    '                        vPriceLevel = "ราคาที่1"
    '                        vLevel = 1

    '                        vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "CheckCount")
    '                        dt = ds.Tables("CheckCount")
    '                        If dt.Rows.Count > 0 Then
    '                            vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                        Else
    '                            vCheckItemExist = 0
    '                        End If

    '                        If vCheckItemExist > 0 Then
    '                            vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                            da = New SqlDataAdapter(vQuery, vConnection)
    '                            ds = New DataSet
    '                            da.Fill(ds, "Docno")
    '                            dt = ds.Tables("Docno")
    '                            If dt.Rows.Count > 0 Then
    '                                vDocno = Trim(dt.Rows(0).Item("docno"))
    '                            End If
    '                            vDocno = ""
    '                            MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **เชื่อรับเอง** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                            Exit Sub
    '                        End If

    '                        For i = 0 To ListView101.Items.Count - 1
    '                            vCheckItem = ListView101.Items(i).SubItems(0).Text
    '                            vCheckLevel = ListView101.Items(i).SubItems(2).Text
    '                            vCheckType = ListView101.Items(i).SubItems(3).Text
    '                            vCheckUnitCode = ListView101.Items(i).SubItems(6).Text
    '                            If vItemCode = vCheckItem And vPriceLevel = vCheckLevel And vSaleType = vCheckType And vUnitCode = vCheckUnitCode Then
    '                                MsgBox("มีการปรับราคาสินค้า " & vItemCode & " ในระดับราคาดังกล่าวอยู่แล้ว", MsgBoxStyle.Critical, "Send Error")
    '                                Exit Sub
    '                            End If
    '                        Next
    '                    End If
    '                    If TextCredit02.Text <> "" Then
    '                        vPriceLevel = "ราคาที่2"
    '                        vLevel = 2

    '                        vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "CheckCount")
    '                        dt = ds.Tables("CheckCount")
    '                        If dt.Rows.Count > 0 Then
    '                            vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                        Else
    '                            vCheckItemExist = 0
    '                        End If

    '                        If vCheckItemExist > 0 Then
    '                            vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                            da = New SqlDataAdapter(vQuery, vConnection)
    '                            ds = New DataSet
    '                            da.Fill(ds, "Docno")
    '                            dt = ds.Tables("Docno")
    '                            If dt.Rows.Count > 0 Then
    '                                vDocno = Trim(dt.Rows(0).Item("docno"))
    '                            End If
    '                            MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **เชื่อส่งให้** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                            Exit Sub
    '                        End If

    '                        For i = 0 To ListView101.Items.Count - 1
    '                            vCheckItem = ListView101.Items(i).SubItems(0).Text
    '                            vCheckLevel = ListView101.Items(i).SubItems(2).Text
    '                            vCheckType = ListView101.Items(i).SubItems(3).Text
    '                            vCheckUnitCode = ListView101.Items(i).SubItems(6).Text
    '                            If vItemCode = vCheckItem And vPriceLevel = vCheckLevel And vSaleType = vCheckType And vUnitCode = vCheckUnitCode Then
    '                                MsgBox("มีการปรับราคาสินค้า " & vItemCode & " ในระดับราคาดังกล่าวอยู่แล้ว", MsgBoxStyle.Critical, "Send Error")
    '                                Exit Sub
    '                            End If
    '                        Next
    '                    End If
    '                End If
    '                If CHK104.Checked = True Then
    '                    vSaleType = "ขายเชื่อส่งให้"
    '                    vType = 1
    '                    vTransType = 1
    '                    If Me.TextCredit11.Text <> "" Then
    '                        vPriceLevel = "ราคาที่1"
    '                        vLevel = 1

    '                        vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "CheckCount")
    '                        dt = ds.Tables("CheckCount")
    '                        If dt.Rows.Count > 0 Then
    '                            vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                        Else
    '                            vCheckItemExist = 0
    '                        End If

    '                        If vCheckItemExist > 0 Then
    '                            vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                            da = New SqlDataAdapter(vQuery, vConnection)
    '                            ds = New DataSet
    '                            da.Fill(ds, "Docno")
    '                            dt = ds.Tables("Docno")
    '                            If dt.Rows.Count > 0 Then
    '                                vDocno = Trim(dt.Rows(0).Item("docno"))
    '                            End If
    '                            MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **เชื่อส่งให้** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                            Exit Sub
    '                        End If

    '                        For i = 0 To ListView101.Items.Count - 1
    '                            vCheckItem = ListView101.Items(i).SubItems(0).Text
    '                            vCheckLevel = ListView101.Items(i).SubItems(2).Text
    '                            vCheckType = ListView101.Items(i).SubItems(3).Text
    '                            vCheckUnitCode = ListView101.Items(i).SubItems(6).Text
    '                            If vItemCode = vCheckItem And vPriceLevel = vCheckLevel And vSaleType = vCheckType And vUnitCode = vCheckUnitCode Then
    '                                MsgBox("มีการปรับราคาสินค้า " & vItemCode & " ในระดับราคาดังกล่าวอยู่แล้ว", MsgBoxStyle.Critical, "Send Error")
    '                                Exit Sub
    '                            End If
    '                        Next
    '                    End If
    '                    If TextCredit12.Text <> "" Then
    '                        vPriceLevel = "ราคาที่2"
    '                        vLevel = 2

    '                        vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "CheckCount")
    '                        dt = ds.Tables("CheckCount")
    '                        If dt.Rows.Count > 0 Then
    '                            vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                        Else
    '                            vCheckItemExist = 0
    '                        End If

    '                        If vCheckItemExist > 0 Then
    '                            vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                            da = New SqlDataAdapter(vQuery, vConnection)
    '                            ds = New DataSet
    '                            da.Fill(ds, "Docno")
    '                            dt = ds.Tables("Docno")
    '                            If dt.Rows.Count > 0 Then
    '                                vDocno = Trim(dt.Rows(0).Item("docno"))
    '                            End If
    '                            MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **เชื่อส่งให้** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                            Exit Sub
    '                        End If

    '                        For i = 0 To ListView101.Items.Count - 1
    '                            vCheckItem = ListView101.Items(i).SubItems(0).Text
    '                            vCheckLevel = ListView101.Items(i).SubItems(2).Text
    '                            vCheckType = ListView101.Items(i).SubItems(3).Text
    '                            vCheckUnitCode = ListView101.Items(i).SubItems(6).Text
    '                            If vItemCode = vCheckItem And vPriceLevel = vCheckLevel And vSaleType = vCheckType And vUnitCode = vCheckUnitCode Then
    '                                MsgBox("มีการปรับราคาสินค้า " & vItemCode & " ในระดับราคาดังกล่าวอยู่แล้ว", MsgBoxStyle.Critical, "Send Error")
    '                                Exit Sub
    '                            End If
    '                        Next
    '                    End If
    '                End If
    '            End If

    '            If CHK101.Checked = True Then
    '                If Me.TextCash01.Text <> "" Then
    '                    vOldPrice1 = Trim(Me.LBLOldCash01.Text)
    '                    vNewPrice1 = Trim(Me.TextCash01.Text)
    '                    vType = 0
    '                    vTransType = 0
    '                    vLevel = 1
    '                    vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                    da = New SqlDataAdapter(vQuery, vConnection)
    '                    ds = New DataSet
    '                    da.Fill(ds, "CheckCount")
    '                    dt = ds.Tables("CheckCount")
    '                    If dt.Rows.Count > 0 Then
    '                        vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                    Else
    '                        vCheckItemExist = 0
    '                    End If

    '                    If vCheckItemExist > 0 Then
    '                        vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "Docno")
    '                        dt = ds.Tables("Docno")
    '                        If dt.Rows.Count > 0 Then
    '                            vDocno = Trim(dt.Rows(0).Item("docno"))
    '                        End If
    '                        MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **สดรับเอง** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                        Exit Sub
    '                    End If

    '                    vListItemCode = ListView101.Items.Add(vItemCode)
    '                    vListItemCode.SubItems.Add(1).Text = Trim(vItemName)
    '                    vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่1")
    '                    vListItemCode.SubItems.Add(3).Text = Trim("ขายสดรับเอง")
    '                    vListItemCode.SubItems.Add(4).Text = Trim(vNewPrice1)
    '                    vListItemCode.SubItems.Add(5).Text = Trim(vOldPrice1)
    '                    vListItemCode.SubItems.Add(6).Text = Trim(vUnitCode)
    '                End If
    '                If Me.TextCash02.Text <> "" Then
    '                    vOldPrice2 = Trim(Me.LBLOldCash02.Text)
    '                    vNewPrice2 = Trim(Me.TextCash02.Text)
    '                    vType = 0
    '                    vTransType = 0
    '                    vLevel = 2
    '                    vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                    da = New SqlDataAdapter(vQuery, vConnection)
    '                    ds = New DataSet
    '                    da.Fill(ds, "CheckCount")
    '                    dt = ds.Tables("CheckCount")
    '                    If dt.Rows.Count > 0 Then
    '                        vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                    Else
    '                        vCheckItemExist = 0
    '                    End If

    '                    If vCheckItemExist > 0 Then
    '                        vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "Docno")
    '                        dt = ds.Tables("Docno")
    '                        If dt.Rows.Count > 0 Then
    '                            vDocno = Trim(dt.Rows(0).Item("docno"))
    '                        End If
    '                        MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **สดรับเอง** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                        Exit Sub
    '                    End If

    '                    vListItemCode = ListView101.Items.Add(vItemCode)
    '                    vListItemCode.SubItems.Add(1).Text = Trim(vItemName)
    '                    vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่2")
    '                    vListItemCode.SubItems.Add(3).Text = Trim("ขายสดรับเอง")
    '                    vListItemCode.SubItems.Add(4).Text = Trim(vNewPrice2)
    '                    vListItemCode.SubItems.Add(5).Text = Trim(vOldPrice2)
    '                    vListItemCode.SubItems.Add(6).Text = Trim(vUnitCode)
    '                End If
    '            End If

    '            If CHK102.Checked = True Then
    '                If Me.TextCash11.Text <> "" Then
    '                    vOldPrice1 = Trim(Me.LBLOldCash11.Text)
    '                    vNewPrice1 = Trim(Me.TextCash11.Text)
    '                    vType = 0
    '                    vTransType = 1
    '                    vLevel = 1
    '                    vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                    da = New SqlDataAdapter(vQuery, vConnection)
    '                    ds = New DataSet
    '                    da.Fill(ds, "CheckCount")
    '                    dt = ds.Tables("CheckCount")
    '                    If dt.Rows.Count > 0 Then
    '                        vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                    Else
    '                        vCheckItemExist = 0
    '                    End If

    '                    If vCheckItemExist > 0 Then
    '                        vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "Docno")
    '                        dt = ds.Tables("Docno")
    '                        If dt.Rows.Count > 0 Then
    '                            vDocno = Trim(dt.Rows(0).Item("docno"))
    '                        End If
    '                        MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **สดส่งให้** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                        Exit Sub
    '                    End If

    '                    vListItemCode = ListView101.Items.Add(vItemCode)
    '                    vListItemCode.SubItems.Add(1).Text = Trim(vItemName)
    '                    vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่1")
    '                    vListItemCode.SubItems.Add(3).Text = Trim("ขายสดส่งให้")
    '                    vListItemCode.SubItems.Add(4).Text = Trim(vNewPrice1)
    '                    vListItemCode.SubItems.Add(5).Text = Trim(vOldPrice1)
    '                    vListItemCode.SubItems.Add(6).Text = Trim(vUnitCode)
    '                End If
    '                If Me.TextCash12.Text <> "" Then
    '                    vOldPrice2 = Trim(Me.LBLOldCash12.Text)
    '                    vNewPrice2 = Trim(Me.TextCash12.Text)
    '                    vType = 0
    '                    vTransType = 1
    '                    vLevel = 2
    '                    vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                    da = New SqlDataAdapter(vQuery, vConnection)
    '                    ds = New DataSet
    '                    da.Fill(ds, "CheckCount")
    '                    dt = ds.Tables("CheckCount")
    '                    If dt.Rows.Count > 0 Then
    '                        vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                    Else
    '                        vCheckItemExist = 0
    '                    End If

    '                    If vCheckItemExist > 0 Then
    '                        vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "Docno")
    '                        dt = ds.Tables("Docno")
    '                        If dt.Rows.Count > 0 Then
    '                            vDocno = Trim(dt.Rows(0).Item("docno"))
    '                        End If
    '                        MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **สดส่งให้** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                        Exit Sub
    '                    End If
    '                    vListItemCode = ListView101.Items.Add(vItemCode)
    '                    vListItemCode.SubItems.Add(1).Text = Trim(vItemName)
    '                    vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่2")
    '                    vListItemCode.SubItems.Add(3).Text = Trim("ขายสดส่งให้")
    '                    vListItemCode.SubItems.Add(4).Text = Trim(vNewPrice2)
    '                    vListItemCode.SubItems.Add(5).Text = Trim(vOldPrice2)
    '                    vListItemCode.SubItems.Add(6).Text = Trim(vUnitCode)
    '                End If
    '            End If

    '            If CHK103.Checked = True Then
    '                If Me.TextCredit01.Text <> "" Then
    '                    vOldPrice1 = Trim(Me.LBLOldCredit01.Text)
    '                    vNewPrice1 = Trim(Me.TextCredit01.Text)
    '                    vType = 1
    '                    vTransType = 0
    '                    vLevel = 1
    '                    vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                    da = New SqlDataAdapter(vQuery, vConnection)
    '                    ds = New DataSet
    '                    da.Fill(ds, "CheckCount")
    '                    dt = ds.Tables("CheckCount")
    '                    If dt.Rows.Count > 0 Then
    '                        vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                    Else
    '                        vCheckItemExist = 0
    '                    End If

    '                    If vCheckItemExist > 0 Then
    '                        vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "Docno")
    '                        dt = ds.Tables("Docno")
    '                        If dt.Rows.Count > 0 Then
    '                            vDocno = Trim(dt.Rows(0).Item("docno"))
    '                        End If
    '                        MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **เชื่อรับเอง** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                        Exit Sub
    '                    End If
    '                    vListItemCode = ListView101.Items.Add(vItemCode)
    '                    vListItemCode.SubItems.Add(1).Text = Trim(vItemName)
    '                    vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่1")
    '                    vListItemCode.SubItems.Add(3).Text = Trim("ขายเชื่อรับเอง")
    '                    vListItemCode.SubItems.Add(4).Text = Trim(vNewPrice1)
    '                    vListItemCode.SubItems.Add(5).Text = Trim(vOldPrice1)
    '                    vListItemCode.SubItems.Add(6).Text = Trim(vUnitCode)
    '                End If
    '                If Me.TextCredit02.Text <> "" Then
    '                    vOldPrice2 = Trim(Me.LBLOldCredit02.Text)
    '                    vNewPrice2 = Trim(Me.TextCredit02.Text)
    '                    vType = 1
    '                    vTransType = 0
    '                    vLevel = 2
    '                    vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                    da = New SqlDataAdapter(vQuery, vConnection)
    '                    ds = New DataSet
    '                    da.Fill(ds, "CheckCount")
    '                    dt = ds.Tables("CheckCount")
    '                    If dt.Rows.Count > 0 Then
    '                        vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                    Else
    '                        vCheckItemExist = 0
    '                    End If

    '                    If vCheckItemExist > 0 Then
    '                        vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "Docno")
    '                        dt = ds.Tables("Docno")
    '                        If dt.Rows.Count > 0 Then
    '                            vDocno = Trim(dt.Rows(0).Item("docno"))
    '                        End If
    '                        MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **เชื่อรับเอง** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                        Exit Sub
    '                    End If
    '                    vListItemCode = ListView101.Items.Add(vItemCode)
    '                    vListItemCode.SubItems.Add(1).Text = Trim(vItemName)
    '                    vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่2")
    '                    vListItemCode.SubItems.Add(3).Text = Trim("ขายเชื่อรับเอง")
    '                    vListItemCode.SubItems.Add(4).Text = Trim(vNewPrice2)
    '                    vListItemCode.SubItems.Add(5).Text = Trim(vOldPrice2)
    '                    vListItemCode.SubItems.Add(6).Text = Trim(vUnitCode)
    '                End If
    '            End If

    '            If CHK104.Checked = True Then
    '                If Me.TextCredit11.Text <> "" Then
    '                    vOldPrice1 = Trim(Me.LBLOldCredit11.Text)
    '                    vNewPrice1 = Trim(Me.TextCredit11.Text)
    '                    vType = 1
    '                    vTransType = 1
    '                    vLevel = 1
    '                    vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                    da = New SqlDataAdapter(vQuery, vConnection)
    '                    ds = New DataSet
    '                    da.Fill(ds, "CheckCount")
    '                    dt = ds.Tables("CheckCount")
    '                    If dt.Rows.Count > 0 Then
    '                        vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                    Else
    '                        vCheckItemExist = 0
    '                    End If

    '                    If vCheckItemExist > 0 Then
    '                        vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "Docno")
    '                        dt = ds.Tables("Docno")
    '                        If dt.Rows.Count > 0 Then
    '                            vDocno = Trim(dt.Rows(0).Item("docno"))
    '                        End If
    '                        MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **เชื่อส่งให้** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                        Exit Sub
    '                    End If
    '                    vListItemCode = ListView101.Items.Add(vItemCode)
    '                    vListItemCode.SubItems.Add(1).Text = Trim(vItemName)
    '                    vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่1")
    '                    vListItemCode.SubItems.Add(3).Text = Trim("ขายเชื่อส่งให้")
    '                    vListItemCode.SubItems.Add(4).Text = Trim(vNewPrice1)
    '                    vListItemCode.SubItems.Add(5).Text = Trim(vOldPrice1)
    '                    vListItemCode.SubItems.Add(6).Text = Trim(vUnitCode)
    '                End If
    '                If Me.TextCredit12.Text <> "" Then
    '                    vOldPrice2 = Trim(Me.LBLOldCredit12.Text)
    '                    vNewPrice2 = Trim(Me.TextCredit12.Text)
    '                    vType = 1
    '                    vTransType = 1
    '                    vLevel = 2
    '                    vQuery = "exec dbo.USP_NP_CheckInsertItemCodeChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                    da = New SqlDataAdapter(vQuery, vConnection)
    '                    ds = New DataSet
    '                    da.Fill(ds, "CheckCount")
    '                    dt = ds.Tables("CheckCount")
    '                    If dt.Rows.Count > 0 Then
    '                        vCheckItemExist = dt.Rows(0).Item("vCountItem")
    '                    Else
    '                        vCheckItemExist = 0
    '                    End If

    '                    If vCheckItemExist > 0 Then
    '                        vQuery = "exec dbo.USP_NP_SelectDocnoChangePrice '" & vScheduleDate & "','" & vItemCode & "'," & vLevel & "," & vType & "," & vTransType & ",'" & vUnitCode & "' "
    '                        da = New SqlDataAdapter(vQuery, vConnection)
    '                        ds = New DataSet
    '                        da.Fill(ds, "Docno")
    '                        dt = ds.Tables("Docno")
    '                        If dt.Rows.Count > 0 Then
    '                            vDocno = Trim(dt.Rows(0).Item("docno"))
    '                        End If
    '                        MsgBox("มีข้อมูลการปรับสินค้า รหัส **" & vItemCode & "** หน่วยนับ **" & vUnitCode & "** ระดับราคาที่ **" & vLevel & "** ประเภทการขาย **เชื่อส่งให้** นี้ ในเอกสารเลขที่ " & vDocno & " ที่จะปรับราคาในวันที่ " & vScheduleDate & " อยู่แล้ว กรุณาตรวจสอบ ", MsgBoxStyle.Critical, "Send Error")
    '                        Exit Sub
    '                    End If
    '                    vListItemCode = ListView101.Items.Add(vItemCode)
    '                    vListItemCode.SubItems.Add(1).Text = Trim(vItemName)
    '                    vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่2")
    '                    vListItemCode.SubItems.Add(3).Text = Trim("ขายเชื่อส่งให้")
    '                    vListItemCode.SubItems.Add(4).Text = Trim(vNewPrice2)
    '                    vListItemCode.SubItems.Add(5).Text = Trim(vOldPrice2)
    '                    vListItemCode.SubItems.Add(6).Text = Trim(vUnitCode)
    '                End If
    '            End If
    '            Call ClearData()
    '            Me.TextItemCode.Text = ""
    '            Me.TextItemCode.Focus()
    '        Else
    '            MsgBox("เอกสารถูกอนุมัติเรียบร้อยแล้ว แก้ไขหรือเพิ่มเติมไม่ได้", MsgBoxStyle.Critical, "Send Error")
    '        End If

    'ErrDescription:
    '        If Err.Description <> "" Then
    '            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
    '            Exit Sub
    '        End If
    '    End Sub

    Private Sub TextCash01_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextCash01.LostFocus, TextCash02.LostFocus, TextCash11.LostFocus, TextCash12.LostFocus, TextCredit01.LostFocus, TextCredit02.LostFocus, TextCredit11.LostFocus, TextCredit12.LostFocus
        On Error GoTo ErrDescription

        If Me.TextCash01.Text <> "" Then
            Me.TextCash01.Text = Format(Int(Me.TextCash01.Text), "##,##0.00")
        End If
        If Me.TextCash02.Text <> "" Then
            Me.TextCash02.Text = Format(Int(Me.TextCash02.Text), "##,##0.00")
        End If
        If Me.TextCash11.Text <> "" Then
            Me.TextCash11.Text = Format(Int(Me.TextCash11.Text), "##,##0.00")
        End If
        If Me.TextCash12.Text <> "" Then
            Me.TextCash12.Text = Format(Int(Me.TextCash12.Text), "##,##0.00")
        End If
        If Me.TextCredit01.Text <> "" Then
            Me.TextCredit01.Text = Format(Int(Me.TextCredit01.Text), "##,##0.00")
        End If
        If Me.TextCredit02.Text <> "" Then
            Me.TextCredit02.Text = Format(Int(Me.TextCredit02.Text), "##,##0.00")
        End If
        If Me.TextCredit11.Text <> "" Then
            Me.TextCredit11.Text = Format(Int(Me.TextCredit11.Text), "##,##0.00")
        End If
        If Me.TextCredit12.Text <> "" Then
            Me.TextCredit12.Text = Format(Int(Me.TextCredit12.Text), "##,##0.00")
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNGenNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNGenNumber.Click
        Dim vGenNumber As String
        Dim vYear As String
        Dim vMonth As String
        Dim vNumber As Integer

        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_NP_GenerateItemChangePriceNumber"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "GenNumber")
        dt = ds.Tables("GenNumber")
        If dt.Rows.Count > 0 Then
            vYear = Trim(dt.Rows(0).Item("year1"))
            vMonth = Trim(dt.Rows(0).Item("month1"))
            vNumber = dt.Rows(0).Item("maxnumber")
        Else
            Exit Sub
        End If
        vYear = Microsoft.VisualBasic.Right(vYear, 2)
        If Len(vMonth) < 2 Then
            vMonth = "0" & vMonth
        End If
        vGenNumber = "CP" & vYear & vMonth & "-" & Format(vNumber, "0000")
        Me.TextDocNo.Text = vGenNumber
        Me.TextItemCode.Focus()

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vType As Integer
        Dim vDocNo As String
        Dim vDocDate As String
        Dim vScheduleDate As String
        Dim vCreatorCode As String
        Dim i As Integer
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vPriceLevel As Integer
        Dim vSaleType As Integer
        Dim vTransSportType As Integer
        Dim vNewPrice As Decimal
        Dim vOldPrice As Decimal
        Dim vLineNumber As Integer

        If vIsConfirm = 0 Then
            If Me.TextDocNo.Text <> "" Then
                vDocNo = Trim(Me.TextDocNo.Text)
                vDocDate = DTPDocDate.Text
                vScheduleDate = DTPDueDate.Text
                vCreatorCode = vUserID
                If vDocDate = vScheduleDate Then
                    MsgBox("ไม่สามารถเลือกวันที่ปรับราคาตรงกับวันที่เอกสารได้ เนื่องจากเอกสารต้องไปปรับปรุงของเช้าอีกวัน")
                    Exit Sub
                End If
                If ListView101.Items.Count > 0 Then
                    If vIsOpen = 0 Then
                        If DateDiff(DateInterval.Day, Date.Now, CDate(vDocDate)) < 0 Then
                            MsgBox("วันที่เอกสารต้องไม่น้อยกว่าวันที่ปัจจุบัน")
                            Exit Sub
                        End If
                        vType = 0
                        vQuery = "select isnull(count(docno),0) as vCount from npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster where docno = '" & vDocNo & "' "
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vReadQuery = vCMD.ExecuteReader()
                        While vReadQuery.Read
                            vChecCountDocno = vReadQuery(0)
                        End While
                        vReadQuery.Close()

                        If vChecCountDocno > 0 Then
                            MsgBox("มีเลขที่เอกสาร " & vDocNo & " นี้อยู่แล้วในระบบ กรุณากดปุ่มสร้างเลขที่เอกสารใหม่อีกครั้ง ", MsgBoxStyle.Critical, "Send Error")
                            Exit Sub
                        End If
                    Else
                        vType = 1
                        vQuery = "select isnull(count(docno),0) as vCount from npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster where docno = '" & vDocNo & "' "
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vReadQuery = vCMD.ExecuteReader()
                        While vReadQuery.Read
                            vChecCountDocno = vReadQuery(0)
                        End While
                        vReadQuery.Close()

                        If vChecCountDocno = 0 Then
                            MsgBox("ยังไม่มีเลขที่เอกสาร " & vDocNo & " นี้อยู่แล้วในระบบ กรุณาตรวจสอบเลขที่เอกสารด้วย ", MsgBoxStyle.Critical, "Send Error")
                            Exit Sub
                        End If
                    End If


                    Try
                        vQuery = "begin tran" ' เปิด Transaction
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vCMD.ExecuteNonQuery()


                        'Insert Master  
                        vQuery = "exec dbo.USP_NP_InsertBasketUpdateItemPrice " & vType & ",'" & vDocNo & "','" & vDocDate & "','" & vScheduleDate & "','" & vCreatorCode & "' "
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vCMD.ExecuteNonQuery()

                        For i = 0 To ListView101.Items.Count - 1
                            vItemCode = ListView101.Items(i).SubItems(0).Text
                            vItemName = ListView101.Items(i).SubItems(1).Text
                            Select Case ListView101.Items(i).SubItems(2).Text
                                Case "ราคาที่1"
                                    vPriceLevel = 1
                                Case "ราคาที่2"
                                    vPriceLevel = 2
                            End Select

                            Select Case ListView101.Items(i).SubItems(3).Text
                                Case "ขายสดรับเอง"
                                    vSaleType = 0
                                    vTransSportType = 0
                                Case "ขายสดส่งให้"
                                    vSaleType = 0
                                    vTransSportType = 1
                                Case "ขายเชื่อรับเอง"
                                    vSaleType = 1
                                    vTransSportType = 0
                                Case "ขายเชื่อส่งให้"
                                    vSaleType = 1
                                    vTransSportType = 1
                            End Select
                            vNewPrice = ListView101.Items(i).SubItems(4).Text
                            vOldPrice = ListView101.Items(i).SubItems(5).Text
                            vUnitCode = ListView101.Items(i).SubItems(6).Text
                            vLineNumber = i
                            vQuery = "exec dbo.USP_NP_InsertBasketUpdateItemPriceDetails '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vPriceLevel & "," & vSaleType & "," & vTransSportType & "," & vNewPrice & "," & vOldPrice & "," & vLineNumber & " "
                            vCMD = New SqlCommand(vQuery, vConnection)
                            vCMD.ExecuteNonQuery()
                        Next

                        vQuery = "commit tran"
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vCMD.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error ")
                        vQuery = "rollback tran"
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vCMD.ExecuteNonQuery()
                    End Try
                    MsgBox("บันทึกข้อมูลเอกสารปรับราคาสินค้า เลขที่ " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Error ")
                    ListView101.Items.Clear()
                    Me.BTNGenNumber.Focus()
                    vIsOpen = 0
                    Call ClearScreen()

                Else
                    MsgBox("ยังไม่มีสินค้าที่จะทำการปรับราคาในตารางข้างล่าง", MsgBoxStyle.Critical, "Send Error")
                End If
            Else
                MsgBox("กรุณา กดปุ่มสร้างเลขที่เอกสารด้วย", MsgBoxStyle.Critical, "Send Error")
            End If
        Else
            MsgBox("เอกสารถูกอนุมัติแล้วไม่สามารถแก้ไขรายละเอียดได้", MsgBoxStyle.Critical, "Send Error")
        End If
    End Sub

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Me.Close()
    End Sub


    Private Sub BTNPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNPrint.Click
        'If frmReportItemChangePrice Is Nothing Then
        '    frmReportItemChangePrice = New FormReportItemChangePrice
        'Else
        '    If frmReportItemChangePrice.IsDisposed Then
        '        frmReportItemChangePrice = New FormReportItemChangePrice
        '    End If
        'End If
        'frmReportItemChangePrice.MdiParent = FormMain
        'frmReportItemChangePrice.Show()
        'frmReportItemChangePrice.BringToFront()
    End Sub

    Private Sub BTNSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearch.Click
        GBSearchDocNo.Visible = True
        Me.TextSearchDocNo.Focus()
        Me.ListView102.Items.Clear()
        Me.TextSearchDocNo.Text = ""
    End Sub

    Private Sub BTNSearchExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchExit.Click
        GBSearchDocNo.Visible = False
    End Sub

    Private Sub BTNSearchClick_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchClick.Click
        Dim vSearch As String
        Dim vListDocNo As ListViewItem
        Dim i As Integer

        On Error GoTo ErrDescription

        If Me.TextSearchDocNo.Text = "" Then
            vQuery = "exec dbo.USP_NP_SearchChangePriceDocNo 0,''"
        Else
            vSearch = Trim(Me.TextSearchDocNo.Text)
            vQuery = "exec dbo.USP_NP_SearchChangePriceDocNo 1,'" & vSearch & "'"
        End If
        ListView102.Items.Clear()
        Me.BTNGenNumber.Enabled = False
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchDocNo")
        dt = ds.Tables("SearchDocNo")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                vListDocNo = ListView102.Items.Add(Trim(dt.Rows(i).Item("docno")))
                vListDocNo.SubItems.Add(0).Text = Trim(dt.Rows(i).Item("docdate"))
                vListDocNo.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("creatorcode"))
                vListDocNo.SubItems.Add(2).Text = Trim(dt.Rows(i).Item("isconfirm"))
                If dt.Rows(i).Item("isconfirm") = 1 Then
                    ListView102.Items(i).SubItems(0).ForeColor = Color.Green
                    ListView102.Items(i).SubItems(1).ForeColor = Color.Green
                    ListView102.Items(i).SubItems(2).ForeColor = Color.Green
                    ListView102.Items(i).SubItems(3).ForeColor = Color.Green
                End If
            Next
            ListView102.Focus()
        Else
            ListView102.Items.Clear()
            MsgBox("ไม่มีข้อมูลของข้อมูลการปรับราคาสินค้า ตามคำค้นหา", MsgBoxStyle.Information, "Send Information")
            Me.TextSearchDocNo.Focus()

        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub ListView102_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView102.DoubleClick
        Dim vDocNo As String
        Dim i As Integer
        Dim vListItemCode As ListViewItem

        On Error GoTo ErrDescription

        ListView101.Items.Clear()
        If ListView102.Items.Count > 0 Then
            vIsOpen = 1
            vDocNo = Trim(ListView102.SelectedItems(0).SubItems(0).Text)
            vIsConfirm = Trim(ListView102.SelectedItems(0).SubItems(3).Text)
            If vIsConfirm = 0 Then
                Me.Pic101.Visible = True
                Me.Pic102.Visible = False
            ElseIf vIsConfirm = 1 Then
                Me.Pic101.Visible = False
                Me.Pic102.Visible = True
            End If
            vQuery = "exec dbo.USP_NP_SearchChangePrice '" & vDocNo & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "DocNo")
            dt = ds.Tables("DocNo")

            If dt.Rows.Count > 0 Then
                vIsOpen = 1
                Me.TextDocNo.Text = Trim(dt.Rows(i).Item("docno"))
                Me.DTPDocDate.Text = Trim(dt.Rows(i).Item("docdate"))
                Me.DTPDueDate.Text = Trim(dt.Rows(i).Item("scheduledate"))


                For i = 0 To dt.Rows.Count - 1
                    vListItemCode = ListView101.Items.Add(Trim(dt.Rows(i).Item("itemcode")))
                    vListItemCode.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("itemname"))
                    If dt.Rows(i).Item("pricelevel") = 1 Then
                        vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่1")
                    ElseIf dt.Rows(i).Item("pricelevel") = 2 Then
                        vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่2")
                    End If
                    Select Case dt.Rows(i).Item("saletype")
                        Case 0
                            Select Case dt.Rows(i).Item("transsporttype")
                                Case 0
                                    vListItemCode.SubItems.Add(3).Text = Trim("ขายสดรับเอง")
                                Case 1
                                    vListItemCode.SubItems.Add(3).Text = Trim("ขายสดส่งให้")
                            End Select
                        Case 1
                            Select Case dt.Rows(i).Item("transsporttype")
                                Case 0
                                    vListItemCode.SubItems.Add(3).Text = Trim("ขายเชื่อรับเอง")
                                Case 1
                                    vListItemCode.SubItems.Add(3).Text = Trim("ขายเชื่อส่งให้")
                            End Select
                    End Select
                    vListItemCode.SubItems.Add(4).Text = Format(dt.Rows(i).Item("newprice"), "##,##0.00")
                    vListItemCode.SubItems.Add(5).Text = Format(dt.Rows(i).Item("oldprice"), "##,##0.00")
                    vListItemCode.SubItems.Add(6).Text = Trim(dt.Rows(i).Item("unitcode"))
                Next
                Me.GBSearchDocNo.Visible = False
            End If
        Else
            Me.TextSearchDocNo.Focus()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub


    Private Sub BTNDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNDelete.Click
        Dim vDocNo As String

        On Error GoTo ErrDescription

        vDocNo = Trim(Me.TextDocNo.Text)
        If MessageBox.Show("คุณต้องการลบเอกสาร ใช่หรือไม่", "ข้อความสอบถาม", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
            If vIsConfirm = 1 Then
                MsgBox("ไม่สามารถลบเลขที่เอกสาร " & vDocNo & " ได้เนื่องจากได้ถูกอ้างไปปรับราคาเรียบร้อยแล้ว กรุณาตรวจสอบเลขที่เอกสารด้วย ", MsgBoxStyle.Critical, "Send Error")
                Exit Sub
            End If
            If Me.TextDocNo.Text <> "" And ListView101.Items.Count > 0 Then
                If vIsOpen = 1 Then
                    vQuery = "select isnull(count(docno),0) as vCount from npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster where docno = '" & vDocNo & "' "
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vReadQuery = vCMD.ExecuteReader()
                    While vReadQuery.Read
                        vChecCountDocno = vReadQuery(0)
                    End While
                    vReadQuery.Close()

                    If vChecCountDocno = 0 Then
                        MsgBox("ยังไม่มีเลขที่เอกสาร " & vDocNo & " นี้อยู่แล้วในระบบ กรุณาตรวจสอบเลขที่เอกสารด้วย ", MsgBoxStyle.Critical, "Send Error")
                        Exit Sub
                    End If


                    vQuery = "exec dbo.USP_NP_DeleteChangePriceDocNo '" & vDocNo & "' "
                    vCMD = New SqlCommand(vQuery, vConnection)
                    vCMD.ExecuteNonQuery()

                    MsgBox("ลบข้อมูลของเลขที่เอกสาร " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information")
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


    Private Sub BTNSearchOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchOK.Click
        Dim vDocNo As String
        Dim i As Integer
        Dim vListItemCode As ListViewItem

        On Error GoTo ErrDescription

        ListView101.Items.Clear()
        If ListView102.Items.Count > 0 Then
            vDocNo = Trim(ListView102.SelectedItems(0).SubItems(0).Text)
            vIsConfirm = Trim(ListView102.SelectedItems(0).SubItems(3).Text)
            vQuery = "exec dbo.USP_NP_SearchChangePrice '" & vDocNo & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "DocNo")
            dt = ds.Tables("DocNo")

            If dt.Rows.Count > 0 Then
                vIsOpen = 1
                Me.TextDocNo.Text = Trim(dt.Rows(i).Item("docno"))
                Me.DTPDocDate.Text = Trim(dt.Rows(i).Item("docdate"))
                Me.DTPDueDate.Text = Trim(dt.Rows(i).Item("scheduledate"))


                For i = 0 To dt.Rows.Count - 1
                    vListItemCode = ListView101.Items.Add(Trim(dt.Rows(i).Item("itemcode")))
                    vListItemCode.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("itemname"))
                    If dt.Rows(i).Item("pricelevel") = 1 Then
                        vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่1")
                    ElseIf dt.Rows(i).Item("pricelevel") = 2 Then
                        vListItemCode.SubItems.Add(2).Text = Trim("ราคาที่2")
                    End If
                    Select Case dt.Rows(i).Item("saletype")
                        Case 0
                            Select Case dt.Rows(i).Item("transsporttype")
                                Case 0
                                    vListItemCode.SubItems.Add(3).Text = Trim("ขายสดรับเอง")
                                Case 1
                                    vListItemCode.SubItems.Add(3).Text = Trim("ขายสดส่งให้")
                            End Select
                        Case 1
                            Select Case dt.Rows(i).Item("transsporttype")
                                Case 0
                                    vListItemCode.SubItems.Add(3).Text = Trim("ขายเชื่อรับเอง")
                                Case 1
                                    vListItemCode.SubItems.Add(3).Text = Trim("ขายเชื่อส่งให้")
                            End Select
                    End Select
                    vListItemCode.SubItems.Add(4).Text = Trim(dt.Rows(i).Item("newprice"))
                    vListItemCode.SubItems.Add(5).Text = Trim(dt.Rows(i).Item("oldprice"))
                    vListItemCode.SubItems.Add(6).Text = Trim(dt.Rows(i).Item("unitcode"))
                Next
                Me.GBSearchDocNo.Visible = False
            End If
        Else
            Me.TextSearchDocNo.Focus()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSearchItemCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSearchItemCode.Click
        Me.GB103.Visible = True
        Me.TextSearchItem.Text = ""
        Me.TextSearchItem.Focus()
        Me.ListView103.Items.Clear()
    End Sub

    Private Sub TextSearchItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextSearchItem.KeyDown
        Dim vSearchItemCode As String
        Dim vListItem As ListViewItem
        Dim i As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.TextSearchItem.Text <> "" Then
                vSearchItemCode = Trim(Me.TextSearchItem.Text)
                vQuery = "select * from (select code,name1 as itemname from dbo.bcitem where code like '%" & vSearchItemCode & "%' and  activestatus = 1 union select code,name1 as itemname from dbo.bcitem where name1 like '%" & vSearchItemCode & "%' and  activestatus = 1 ) as result order by code"
                da = New SqlDataAdapter(vQuery, vConnection)
                ds = New DataSet
                da.Fill(ds, "ItemCode")
                dt = ds.Tables("ItemCode")
                ListView103.Items.Clear()
                If dt.Rows.Count > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        vListItem = ListView103.Items.Add(Trim(dt.Rows(i).Item("code")))
                        vListItem.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("itemname"))
                    Next i
                Else
                    MsgBox("ไม่มีข้อมูลสินค้าที่ต้องการค้นหา", MsgBoxStyle.Critical, "Send Information")
                    Exit Sub
                End If
            Else
                MsgBox("ไม่ได้ใส่ข้อมูลในการค้นหาสินค้า", MsgBoxStyle.Critical, "Send Error")
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub ListView103_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView103.DoubleClick
        On Error GoTo ErrDescription

        If ListView103.Items.Count > 0 Then
            Me.TextItemCode.Text = Trim(ListView103.SelectedItems(0).SubItems(0).Text)
            GB103.Visible = False
            Call ClearData()
            Me.TextItemCode.Focus()
            Call ItemCodeClick()
        Else
            Me.TextSearchItem.Focus()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSelectItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectItem.Click
        On Error GoTo ErrDescription

        If ListView103.Items.Count > 0 Then
            Me.TextItemCode.Text = Trim(ListView103.SelectedItems(0).SubItems(0).Text)
            GB103.Visible = False
            Call ClearData()
            Me.TextItemCode.Focus()
            Call ItemCodeClick()
        Else
            Me.TextSearchItem.Focus()
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub BTNItemExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNItemExit.Click
        Me.GB103.Visible = False
    End Sub

    Private Sub ItemCodeClick()
        Dim vItemCode As String
        Dim i As Integer

        On Error GoTo ErrDescription

        Call ClearData()
        vItemCode = Trim(TextItemCode.Text)
        vQuery = "select code,name1 from dbo.bcitem where code = '" & vItemCode & "' "
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "KeyItemCode")
        dt = ds.Tables("KeyItemCode")
        If dt.Rows.Count > 0 Then
            Me.LBLItemName.Text = dt.Rows(0).Item("name1")
        Else
            MsgBox("ไม่มีรหัสสินค้า รหัส " & vItemCode & " นี้ในระบบ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If

        vQuery = "select distinct itemcode,unitcode from dbo.bcpricelist where itemcode = '" & vItemCode & "' order by unitcode"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "UnitCode")
        dt = ds.Tables("UnitCode")
        For i = 0 To dt.Rows.Count - 1
            CMBUnitCode.Items.Add(Trim(dt.Rows(i).Item("unitcode")))
        Next
        Me.CMBUnitCode.Focus()

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TextItemCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextItemCode.TextChanged
        Me.LBLItemName.Text = ""
    End Sub

    Private Sub BTNClearForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClearForm.Click
        Call ClearScreen()
    End Sub

    Private Sub BTNConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNConfirm.Click
        Dim vDocNo As String

        On Error GoTo ErrDescription

        Call ChekAuthorityAccess()
        If vDepartment = "IT" And vLevelID = 1 Then

            vDocNo = Trim(Me.TextDocNo.Text)
            If MessageBox.Show("คุณต้องการอนุมัติเอกสาร ใช่หรือไม่", "ข้อความสอบถาม", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                If vIsConfirm = 1 Then
                    MsgBox("ไม่สามารถอนุมัติเลขที่เอกสาร " & vDocNo & " ได้เนื่องจากได้ถูกอนุมัติเรียบร้อยแล้ว กรุณาตรวจสอบเลขที่เอกสารด้วย ", MsgBoxStyle.Critical, "Send Error")
                    Exit Sub
                End If
                If Me.TextDocNo.Text <> "" And ListView101.Items.Count > 0 Then
                    If vIsOpen = 1 Then
                        vQuery = "select isnull(count(docno),0) as vCount from npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster where docno = '" & vDocNo & "' "
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vReadQuery = vCMD.ExecuteReader()
                        While vReadQuery.Read
                            vChecCountDocno = vReadQuery(0)
                        End While
                        vReadQuery.Close()

                        If vChecCountDocno = 0 Then
                            MsgBox("ยังไม่มีเลขที่เอกสาร " & vDocNo & " นี้อยู่แล้วในระบบ กรุณาตรวจสอบเลขที่เอกสารด้วย ", MsgBoxStyle.Critical, "Send Error")
                            Exit Sub
                        End If


                        vQuery = "update npmaster.dbo.TB_NP_BasketItemUpdatePriceMaster set isconfirm = 1 where docno = '" & vDocNo & "'"
                        vCMD = New SqlCommand(vQuery, vConnection)
                        vCMD.ExecuteNonQuery()

                        MsgBox("อนุมัติเลขที่เอกสาร " & vDocNo & " เรียบร้อยแล้ว", MsgBoxStyle.Information, "Send Information")
                        Call ClearScreen()
                    End If
                End If
            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TextSearchDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextSearchDocNo.KeyDown
        Dim vSearch As String
        Dim vListDocNo As ListViewItem
        Dim i As Integer

        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            If Me.TextSearchDocNo.Text = "" Then
                vQuery = "exec dbo.USP_NP_SearchChangePriceDocNo 0,''"
            Else
                vSearch = Trim(Me.TextSearchDocNo.Text)
                vQuery = "exec dbo.USP_NP_SearchChangePriceDocNo 1,'" & vSearch & "'"
            End If
            ListView102.Items.Clear()
            Me.BTNGenNumber.Enabled = False
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "SearchDocNo")
            dt = ds.Tables("SearchDocNo")
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    vListDocNo = ListView102.Items.Add(Trim(dt.Rows(i).Item("docno")))
                    vListDocNo.SubItems.Add(0).Text = Trim(dt.Rows(i).Item("docdate"))
                    vListDocNo.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("creatorcode"))
                    vListDocNo.SubItems.Add(2).Text = Trim(dt.Rows(i).Item("isconfirm"))
                    If dt.Rows(i).Item("isconfirm") = 1 Then
                        ListView102.Items(i).SubItems(0).ForeColor = Color.Green
                        ListView102.Items(i).SubItems(1).ForeColor = Color.Green
                        ListView102.Items(i).SubItems(2).ForeColor = Color.Green
                        ListView102.Items(i).SubItems(3).ForeColor = Color.Green
                    End If
                Next
                ListView102.Focus()
            Else
                ListView102.Items.Clear()
                MsgBox("ไม่มีข้อมูลของข้อมูลการปรับราคาสินค้า ตามคำค้นหา", MsgBoxStyle.Information, "Send Information")
                Me.TextSearchDocNo.Focus()

            End If
        End If

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If

    End Sub


    Private Sub ListView101_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView101.KeyDown
        Dim vListIndex As Integer
        Dim vItemCode As String

        If ListView101.Items.Count > 0 Then
            If vIsConfirm = 0 Then
                If e.KeyCode = Keys.Delete Then
                    vListIndex = ListView101.SelectedItems(0).Index
                    vItemCode = ListView101.Items(vListIndex).SubItems(0).Text
                    If MessageBox.Show("ต้องการลบรายการปรับสินค้ารหัส  " & vItemCode & " ใช่หรือไม่", "Send Question ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                        ListView101.Items.RemoveAt(vListIndex)
                    End If
                End If
            End If
        End If
    End Sub

End Class