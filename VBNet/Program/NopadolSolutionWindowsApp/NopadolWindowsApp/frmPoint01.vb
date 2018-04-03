Imports System.Data
Imports System.Data.SqlClient
Public Class frmPoint01
    Dim QryString As String
    Dim da As New SqlDataAdapter
    Dim cmd As New SqlCommand
    Dim ds As New DataSet
    Dim dt As New DataTable
    Dim CPnum As String
    Dim spPNdocno As String
    Dim iPoint As Integer
    Dim i As Integer
    Dim spQry As String
    Dim saveStatus As Integer
    Dim vIsInsertItem As Integer


    Private Sub btnCPnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCPnew.Click
        Dim StrDate As String
        StrDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        QryString = "set dateformat dmy select dbo.FT_VP_Newcampaign ('" & StrDate & "')"
        da = New SqlDataAdapter(QryString, vConnectionString)
        ds = New DataSet
        da.Fill(ds, "cpno")
        dt = ds.Tables("cpno")
        If dt.Rows.Count <> 0 Then
            CPnum = dt.Rows(0).Item(0)
            Me.txtCPno.Text = CPnum
        End If
        Me.btnSaveCP.Enabled = True
        Me.btnSearchCP.Enabled = False
    End Sub
    Private Sub ClrNewCP()
        Me.txtCPno.Text = ""
        Me.txtCPthName.Text = ""
        Me.txtCPenName.Text = ""
        Me.txtCPno.ReadOnly = False
        Me.txtCPthName.ReadOnly = False
        Me.txtCPenName.ReadOnly = False
        Me.dtpCPStartDate.Enabled = True
        Me.dtpCPendDate.Enabled = True
        Me.btnCPnew.Enabled = True
        Me.dtpCPStartDate.Text = Now.Date
        Me.dtpCPendDate.Text = DateAdd(DateInterval.Year, 1, Now.Date)
        Me.btnSaveCP.Enabled = False
    End Sub

    Private Sub btnSaveCP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveCP.Click
        Dim pvCPno As String
        Dim pvCPthName As String
        Dim pvCPenName As String
        Dim pvCPstartDate As String
        Dim pvCPendDate As String
        pvCPno = Me.txtCPno.Text
        pvCPthName = Me.txtCPthName.Text
        pvCPenName = Me.txtCPenName.Text
        pvCPstartDate = Me.dtpCPStartDate.Text
        pvCPendDate = Me.dtpCPendDate.Text
        If pvCPno <> "" And pvCPthName <> "" Then
            QryString = "exec USP_VP_CampaignAssign '" & 1 & "','" & pvCPno & "','" & pvCPthName & "','" & pvCPenName & "','" & pvCPstartDate & "','" & pvCPendDate & "'"
            With cmd
                .CommandType = CommandType.Text
                .CommandText = QryString
                .Connection = vConnection
                .ExecuteNonQuery()
            End With
        Else
            MsgBox("คุณใส่ข้อมูลไม่ครบ", MsgBoxStyle.Critical, "Error")

        End If
        MsgBox("บันทึกรายการ Campaign เรียบร้อยแล้ว", MsgBoxStyle.Information, "Information")
    End Sub

    Private Sub frmPoint01_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()

        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Me.ControlBox = False
        Me.WindowState = FormWindowState.Maximized
        Me.btnselectReward.Enabled = False

        'Me.MdiParent = frmMainMember
        'If v = 0 Then
        '    Me.pnChecKPoint.Hide()
        '    Me.pnCampaignMaster.Hide()
        '    Me.PNsaveNewItem.Hide()
        '    Me.PNsaveSpecialPoint.Hide()
        '    Me.PNwithDraw.Hide()
        'Else
        '    Me.pnChecKPoint.Show()
        '    Me.pnCampaignMaster.Show()
        '    Me.PNsaveNewItem.Show()
        '    Me.PNsaveSpecialPoint.Show()
        '    Me.PNwithDraw.Show()
        'End If

    End Sub

    Private Sub btnClearCP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearCP.Click
        Call ClrNewCP()
    End Sub

    Private Sub btnExitCP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExitCP.Click
        'ปิดใช้งานการเพิ่มข้อมูล Campaign
        Me.PNwithDraw.Visible = False
        Me.pnChecKPoint.Visible = False
        Me.pnCampaignMaster.Visible = False
        Me.PNsaveNewItem.Visible = False
        Me.PNsaveSpecialPoint.Visible = False

    End Sub

    Private Sub btnSearchCP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchCP.Click
        frmSearchCampaig.Show()
        frmSearchCampaig.txtFindCP.Focus()
        'Me.pnCampaignMaster.Enabled = False
    End Sub
    Private Sub cbxCampaignDW()
        Call InitializeDataBase()
        Dim cbxQry As String
        Dim ida As SqlDataAdapter
        Dim ids As New DataSet
        cbxQry = "exec dbo.USP_VP_CampaignList"
        ida = New SqlDataAdapter(cbxQry, vConnection)
        ida.Fill(ids, "vcbx")
        Me.cbxDW.DataSource = ids.Tables("vcbx")
        Me.cbxDW.DisplayMember = ("NameTh")
        Me.cbxDW.ValueMember = ("code")
    End Sub
    ' Lisview ของรายการของรางวัลที่แลกได้
    Private Sub getRewardforChainge()
        Dim LVrwd As ListViewItem
        Dim rwQryStr As String
        Dim nm As Integer
        iPoint = 0
        iPoint = CInt(Me.txtPointAmount.Text)
        rwQryStr = "exec dbo.USP_VP_CheckAVLReward '" & iPoint & "'"
        da = New SqlDataAdapter(rwQryStr, vConnection)
        ds = New DataSet
        da.Fill(ds, "rwd")
        dt = ds.Tables("rwd")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                nm = nm + 1
                LVrwd = Me.LVRewardChange.Items.Add(nm)
                LVrwd.SubItems.Add(0).Text = dt.Rows(i).Item("itemcode")
                LVrwd.SubItems.Add(0).Text = dt.Rows(i).Item("itemname")
                LVrwd.SubItems.Add(0).Text = dt.Rows(i).Item("stockqty")
                LVrwd.SubItems.Add(0).Text = dt.Rows(i).Item("unitcode")
                LVrwd.SubItems.Add(0).Text = dt.Rows(i).Item("Amount")
                LVrwd.SubItems.Add(0).Text = dt.Rows(i).Item("point")
            Next
        End If

    End Sub

    Private Sub btnFindMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindMember.Click
        iType = 1 'สถานะการค้นหา
        '  Me.pnChecKPoint.Hide()
        MemberSearch.Show()
        MemberSearch.txtFind.Focus()
        MemberSearch.btnSLOK.Enabled = False
        ' Me.Close()
    End Sub

    Private Sub btnNewDW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewDW.Click
        Dim StrDate As String
        StrDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        QryString = "set dateformat dmy select dbo.FT_VP_NewWithDrawSet ('" & StrDate & "')"
        da = New SqlDataAdapter(QryString, vConnection)
        ds = New DataSet
        da.Fill(ds, "cpno")
        dt = ds.Tables("cpno")
        If dt.Rows.Count <> 0 Then
            CPnum = dt.Rows(0).Item(0)
            Me.txtDwDoc.Text = CPnum
        End If
        saveStatus = 1
        Me.btnselectReward.Enabled = False
        Me.btnSaveCP.Enabled = True
        Me.btnSearchCP.Enabled = False
        Me.btnFindDwp.Enabled = False
        Me.btnPrintDwp.Enabled = False
        Me.lblisConfirm.Visible = True
        Me.btnDWFMB.Enabled = True
        Me.lblisConfirm.Text = "--N--"
        Me.lblisConfirm.BackColor = Color.Green
    End Sub
    Private Sub genDocnoDrawPoint()
        Dim StrDate As String
        StrDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        QryString = "set dateformat dmy select dbo.FT_VP_NewWithDrawSet ('" & StrDate & "')"
        da = New SqlDataAdapter(QryString, vConnection)
        ds = New DataSet
        da.Fill(ds, "cpno")
        dt = ds.Tables("cpno")
        If dt.Rows.Count <> 0 Then
            CPnum = dt.Rows(0).Item(0)
            Me.txtDwDoc.Text = CPnum
        End If
        saveStatus = 1
        Me.btnselectReward.Enabled = False
        Me.btnSaveCP.Enabled = True
        Me.btnSearchCP.Enabled = False
        Me.btnFindDwp.Enabled = False
        Me.btnPrintDwp.Enabled = False
        Me.lblisConfirm.Visible = True
        Me.btnDWFMB.Enabled = False
        Me.btnNewDW.Enabled = False
        Me.lblisConfirm.Text = "--N--"
        Me.lblisConfirm.BackColor = Color.Green
    End Sub

    Private Sub btnDWFMB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDWFMB.Click
        iType = 2
        MemberSearch.Show()
        MemberSearch.txtFind.Focus()
        'frmMainMember.Enabled = False
        ' Me.Close()
    End Sub

    Private Sub dgvDWreward_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        ' dlgRewardItem.Show()
        'Dim vTex1 As String
        'vTex1 = InputBox("", "")
        'dgvDWreward.Item(1, 0).Value = vTex1 '(columnIndex, Rowindex)
    End Sub

    Private Sub btnselectReward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnselectReward.Click
        PtotalPoint = Me.txtDWtotalPoint.Text
        dlgRewardItem.Show()
        'frmMainMember.Enabled = False
        ' Me.Close()
    End Sub

    Private Sub dgvDWreward_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDWreward.CellEndEdit
        Dim rwAmount As Double
        Dim rwCost As Double
        Dim rwPoint As Double
        Dim rsCost As Double
        Dim rsPoint As Double
        ' Dim n As Integer
        Dim TotalAmount As Double
        Dim TotalPoint As Double
        'Dim netTotalpoint As Double
        If Me.dgvDWreward.Rows.Count > 0 Then
vloop:
            If e.ColumnIndex = 3 Then
                TotalPoint = 0
                rwAmount = Me.dgvDWreward.Item(3, dgvDWreward.CurrentRow.Index).Value 'qty
                rwCost = Me.dgvDWreward.Item(5, dgvDWreward.CurrentRow.Index).Value 'amount
                rwPoint = Me.dgvDWreward.Item(6, dgvDWreward.CurrentRow.Index).Value 'point
                rsCost = (rwAmount * rwCost)
                rsPoint = (rwAmount * rwPoint)
                Me.dgvDWreward.Item(7, dgvDWreward.CurrentRow.Index).Value() = Format(rsCost, "##,##0.00")
                Me.dgvDWreward.Item(8, dgvDWreward.CurrentRow.Index).Value() = Format(rsPoint, "##,##0.00")
                For i = 0 To Me.dgvDWreward.Rows.Count - 1
                    TotalAmount = (TotalAmount + Me.dgvDWreward.Item(7, dgvDWreward.Rows(i).Index).Value)
                    TotalPoint = (TotalPoint + Me.dgvDWreward.Item(8, dgvDWreward.Rows(i).Index).Value)
                Next
                If TotalPoint > CDbl(Me.txtDWtotalPoint.Text) Then
                    MsgBox("แต้มสมาชิกไม่พอกับจำนวนที่จะเบิก", MsgBoxStyle.Critical, "Warning")
                    Me.dgvDWreward.Item(3, dgvDWreward.CurrentRow.Index).Value = 1
                    GoTo vloop
                End If
                Me.txtTotaldwAmount.Text = Format(TotalAmount, "##,##0.00")
                Me.txtTotaldwPoint.Text = Format(TotalPoint, "##,##0.00")
            End If

            'If TotalPoint > CDbl(Me.txtDWtotalPoint.TabIndex) Then
            '    MsgBox("แต้มสมาชิกไม่พอกับจำนวนที่จะเบิก", MsgBoxStyle.Critical, "Warning")
            '    Me.dgvDWreward.Item(3, dgvDWreward.CurrentRow.Index).Value = 1
            '    rwAmount = 1 'Me.dgvDWreward.Item(3, dgvDWreward.CurrentRow.Index).Value 'qty
            '    rwCost = Me.dgvDWreward.Item(5, dgvDWreward.CurrentRow.Index).Value 'amount
            '    rwPoint = Me.dgvDWreward.Item(6, dgvDWreward.CurrentRow.Index).Value 'point
            '    rsCost = (rwAmount * rwCost)
            '    rsPoint = (rwAmount * rwPoint)
            '    Me.dgvDWreward.Item(7, dgvDWreward.CurrentRow.Index).Value() = Format(rsCost, "##,##0.00")
            '    Me.dgvDWreward.Item(8, dgvDWreward.CurrentRow.Index).Value() = Format(rsPoint, "##,##0.00")
        End If


    End Sub

    Private Sub btnChangeRW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChangeRW.Click
        Me.btnselectReward.Enabled = False
        Call txtReadonly()
        Call ChangeReward()
    End Sub
    Private Sub ChangeReward()
        Dim i As Integer
        Dim dt1 As New DataTable("vData")
        Dim dr As DataRow
        Dim inum As Integer
        Dim iAmount As Double
        Dim iPoint As Double
        Dim iTotalAmount As Double
        Dim iTotalPoint As Double

        'For x = 0 To Me.LVRewardChange.Items.Count - 1
        'If Me.LVRewardChange.Items(x).Checked = True Then
        dt1.Columns.Add("ลำดับ", GetType(Integer))
        dt1.Columns.Add("รหัสสินค้า", GetType(String))
        dt1.Columns.Add("ชื่อสินค้า", GetType(String))
        dt1.Columns.Add("จำนวน", GetType(String))
        dt1.Columns.Add("หน่วยนับ", GetType(String))
        dt1.Columns.Add("มูลค่า:หน่วย", GetType(String))
        dt1.Columns.Add("แต้ม:หน่วย", GetType(String))
        dt1.Columns.Add("มูลค่า", GetType(String))
        dt1.Columns.Add("แต้ม", GetType(String))
        If Me.LVRewardChange.Items.Count > 0 Then
            If Me.LVRewardChange.CheckedItems.Count > 0 Then
                Me.btnChangeRW.Enabled = False
                For i = 0 To Me.LVRewardChange.Items.Count - 1
                    If Me.LVRewardChange.Items(i).Checked = True Then
                        inum = inum + 1
                        dr = dt1.NewRow
                        dr("ลำดับ") = inum
                        dr("รหัสสินค้า") = Me.LVRewardChange.Items(i).SubItems(1).Text
                        dr("ชื่อสินค้า") = Me.LVRewardChange.Items(i).SubItems(2).Text
                        dr("จำนวน") = 1
                        dr("หน่วยนับ") = Me.LVRewardChange.Items(i).SubItems(4).Text
                        iAmount = Me.LVRewardChange.Items(i).SubItems(5).Text
                        dr("มูลค่า:หน่วย") = iAmount
                        iTotalAmount = (iTotalAmount + iAmount)
                        iPoint = Me.LVRewardChange.Items(i).SubItems(6).Text
                        dr("แต้ม:หน่วย") = iPoint
                        iTotalPoint = (iTotalPoint + iPoint)
                        dr("มูลค่า") = iAmount
                        dr("แต้ม") = iPoint
                        dt1.Rows.Add(dr)
                    End If
                Next
                Me.dgvDWreward.DataSource = dt1
                Me.txtdwMemberID.Text = Me.txtMemberid.Text
                Me.txtdwARcode.Text = Me.txtARcode.Text
                Me.txtdwArname.Text = Me.txtMemberName.Text
                Me.txtDWtotalPoint.Text = Me.txtPointAmount.Text
                Me.cbxSPCampaign.SelectedValue = Me.lblCPCode.Text 'value ที่ได้มาจาก label CampaingCode
                Me.txtTotaldwAmount.Text = Format(iTotalAmount, "##,##0.00")
                Me.txtTotaldwPoint.Text = Format(iTotalPoint, "##,##0.00")
                Call genDocnoDrawPoint()
                Call dgvReadonly()
                Call cbxCampaignDW()
                Call clrCHECKpoint()
                Me.PNwithDraw.Visible = True
                Me.pnChecKPoint.Visible = False
            Else
                MsgBox("คุณยังไม่ได้เลือกรายการของรางวัล", MsgBoxStyle.Critical, "Warning")
            End If
        End If

        'End If
        'Next
    End Sub
    Private Sub txtReadonly()
        Me.txtDwDoc.ReadOnly = True
        Me.txtdwARcode.ReadOnly = True
        Me.txtdwMemberID.ReadOnly = True
        Me.txtdwArname.ReadOnly = True
        Me.txtDWtotalPoint.ReadOnly = True
        Me.txtTotaldwAmount.ReadOnly = True
        Me.txtTotaldwPoint.ReadOnly = True
        Me.btnDWFMB.Enabled = False
        Me.cbxSPCampaign.Enabled = False
    End Sub

    Private Sub btnSaveDwp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveDwp.Click
        Dim dwDocno As String
        Dim dwDocdate As String
        Dim dwCPCode As String
        Dim dwArCode As String
        Dim dwItemcode As String
        Dim dwItemName As String
        Dim dwUnitcode As String
        Dim dwPnt As Integer
        Dim dwAmount As Double
        Dim dwQty As Integer
        Dim dwTotalPoint As Double
        Dim dwTotalAmount As Double
        Dim dwGrandTotalPoint As Double
        Dim dwGrandTotalAmount As Double
        Dim dwQry As String
        Dim cmd As New SqlCommand

        '------------------------------------------------------------
        dwDocno = Me.txtDwDoc.Text
        dwDocdate = Me.dtpDWDate.Text
        dwCPCode = Me.cbxDW.SelectedValue
        dwArCode = Me.txtdwARcode.Text
        dwGrandTotalPoint = CDbl(Me.txtTotaldwPoint.Text)
        dwGrandTotalAmount = CDbl(Me.txtTotaldwAmount.Text)
        If saveStatus <> 1 Then
            saveStatus = 0
        End If
        If Me.txtDwDoc.Text <> "" Then
            dwQry = "begin tran"
            With cmd
                .CommandType = CommandType.Text
                .Connection = vConnection
                .CommandText = dwQry
                .ExecuteNonQuery()
            End With
            On Error GoTo ErrDesc
            If saveStatus = 1 Then
                dwQry = "exec dbo.USP_VP_WithdrawSet '" & saveStatus & "','" & dwDocno & "','" & dwDocdate & "','" & dwCPCode & "','" & dwArCode & "'"
                With cmd
                    .CommandType = CommandType.Text
                    .Connection = vConnection
                    .CommandText = dwQry
                    .ExecuteNonQuery()
                End With
                For i = 0 To Me.dgvDWreward.Rows.Count - 1
                    dwItemcode = Me.dgvDWreward.Rows(i).Cells(1).Value
                    dwItemName = Me.dgvDWreward.Rows(i).Cells(2).Value
                    dwQty = Me.dgvDWreward.Rows(i).Cells(3).Value
                    dwUnitcode = Me.dgvDWreward.Rows(i).Cells(4).Value
                    dwAmount = Me.dgvDWreward.Rows(i).Cells(5).Value
                    dwPnt = Me.dgvDWreward.Rows(i).Cells(6).Value
                    dwTotalPoint = Me.dgvDWreward.Rows(i).Cells(7).Value
                    dwTotalAmount = Me.dgvDWreward.Rows(i).Cells(8).Value
                    dwQry = "exec dbo.USP_VP_WithdrawSubSet '" & dwDocno & "','" & dwItemcode & "','" & dwItemName & "','" & dwUnitcode & "'," & dwPnt & "," & dwAmount & "," & dwQty & "," & dwTotalAmount & "," & dwTotalPoint & ""
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = vConnection
                        .CommandText = dwQry
                        .ExecuteNonQuery()
                    End With
                Next
                dwQry = "exec dbo.USP_VP_WithdrawFinalSet '" & dwDocno & "'," & dwGrandTotalPoint & "," & dwGrandTotalAmount & ""
                With cmd
                    .CommandType = CommandType.Text
                    .Connection = vConnection
                    .CommandText = dwQry
                    .ExecuteNonQuery()
                End With

                dwQry = "commit tran"
                With cmd
                    .CommandType = CommandType.Text
                    .Connection = vConnection
                    .CommandText = dwQry
                    .ExecuteNonQuery()

                End With
                MsgBox("บันทึกข้อมูลการเบิกแต้มเรียบร้อยแล้ว", MsgBoxStyle.Information, "Information")
                Me.btnPrintDwp.Enabled = True
                Me.btnFindDwp.Enabled = True
            End If
        Else
            MsgBox("คุณยังไม่ได้ใส่เลขที่เอกสาร", MsgBoxStyle.Critical, "Error")
        End If
ErrDesc:
        If Err.Description <> "" Then
            dwQry = " rollback tran"
            With cmd
                .CommandType = CommandType.Text
                .Connection = vConnection
                .CommandText = dwQry
                .ExecuteNonQuery()
            End With
            MsgBox("ไม่สามารถเบิกแต้มแลกของรางวัลได้", MsgBoxStyle.Critical, "Error")
        End If
    End Sub

    Private Sub btnClearDwp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearDwp.Click
        Call ClearPageWithDraw()
        Call txtReadonly()
    End Sub
    'เคลียร์หน้าบันทึกเบิกแต้ม
    Private Sub ClearPageWithDraw()
        Me.txtDwDoc.Text = ""
        Me.txtdwARcode.Text = ""
        Me.txtdwMemberID.Text = ""
        Me.txtdwArname.Text = ""
        Me.txtDWtotalPoint.Text = ""
        Me.txtTotaldwAmount.Text = ""
        Me.txtTotaldwPoint.Text = ""
        Me.btnDWFMB.Enabled = True
        Me.cbxSPCampaign.Enabled = True
        Me.btnSaveDwp.Enabled = False
        Me.btnPrintDwp.Enabled = False
        Me.btnFindDwp.Enabled = True
        Me.btnselectReward.Enabled = False
        Me.btnNewDW.Enabled = True
        Me.dgvDWreward.DataSource = Nothing
    End Sub

    Private Sub btnExitDwp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExitDwp.Click
        ' Call ClearPageWithDraw()
        Call ClearPageWithDraw()
        Call txtReadonly()
        Me.PNwithDraw.Visible = False
        Me.pnChecKPoint.Visible = False
        Me.pnCampaignMaster.Visible = False
        Me.PNsaveNewItem.Visible = False
        Me.PNsaveSpecialPoint.Visible = False
    End Sub
    'ค้นหาใบเบิกแต้มสมาชิก
    Private Sub btnFindDwp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindDwp.Click
        dlgSearchDWPoint.Show()
        dlgSearchDWPoint.txtDwDocFND.Focus()
    End Sub

    Private Sub btnPrintDwp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintDwp.Click
        'พิมพ์ใบเบิกแต้ม
        If Me.txtDwDoc.Text <> "" Then
            PdwDocNo = Me.txtDwDoc.Text
            frmPrintDWpoint.Show()
        Else
            MsgBox("ไม่มีเอกสารที่จะพิมพ์", MsgBoxStyle.Critical, "Error")
        End If

    End Sub
    ' บันทึกแต้มพิเศษ
    Private Sub cbxSpecialPoint()
        Call InitializeDataBase()
        Dim cbxQry As String
        Dim ida As SqlDataAdapter
        Dim ids As New DataSet
        cbxQry = "exec dbo.USP_VP_CampaignList"
        ida = New SqlDataAdapter(cbxQry, vConnection)
        ida.Fill(ids, "vcbx")
        Me.cbxSPCampaign.DataSource = ids.Tables("vcbx")
        Me.cbxSPCampaign.DisplayMember = ("NameTh")
        Me.cbxSPCampaign.ValueMember = ("code")
    End Sub
    ' เลขที่เอกสารการใหม่แต้มพิเศษ
    Private Sub NewSPdoc()
        Dim StrDate As String
        StrDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        QryString = "set dateformat dmy select dbo.FT_VP_NewPointSpecialSet ('" & StrDate & "')"
        da = New SqlDataAdapter(QryString, vConnection)
        ds = New DataSet
        da.Fill(ds, "pntno")
        dt = ds.Tables("pntno")
        If dt.Rows.Count <> 0 Then
            spPNdocno = dt.Rows(0).Item(0)
            Me.txtSPDocno.Text = spPNdocno
        End If
        Me.btnSaveSP.Enabled = True
        Me.btnFindSP.Enabled = False
        Me.btnFNmember.Enabled = True
    End Sub
  
    Private Sub btnFNmember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFNmember.Click
        Me.btnNewSPpoint.Enabled = False
        iType = 3
        MemberSearch.Show()
        MemberSearch.txtFind.Focus()
        'frmMainMember.Enabled = False
    End Sub

    Private Sub btnSaveSP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveSP.Click
        'บันทึกแต้มพิเศษ
        Dim spDocno As String
        Dim sparCode As String
        Dim spMemberid As String
        Dim pvSVsp As Integer
        Dim spCPcode As String
        Dim spDocdate As String
        Dim spPoint As String
        Dim spReason As String
        pvSVsp = PsaveSPpoint
        If Me.txtSPDocno.Text <> "" And Me.txtSPARCode.Text <> "" And Me.txtSPmemberid.Text <> "" Then
            spDocno = Me.txtSPDocno.Text
            spDocdate = Me.dtpSPDocdate.Text
            sparCode = Me.txtSPARCode.Text
            spMemberid = Me.txtSPmemberid.Text
            spCPcode = Me.cbxSPCampaign.SelectedValue
            spPoint = Me.txtSPpoint.Text
            spReason = Me.txtIssue.Text
            spQry = "exec USP_VP_PointSpecialSet '" & pvSVsp & "','" & spCPcode & "','" & spDocno & "','" & spDocdate & "','" & sparCode & "'," & spPoint & ",'" & spReason & "'"
            With cmd
                .CommandType = CommandType.Text
                .CommandText = spQry
                .Connection = vConnection
                .ExecuteNonQuery()
            End With
        End If
        MsgBox("บันทึกข้อมูลเรียบร้อยแล้ว", MsgBoxStyle.Information, "Infomation")
    End Sub

    Private Sub btnFindSP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSP.Click
        'ค้นหารายการแต้มพิเศษ
        dgvSearchSPPoint.Show()
        'frmMainMember.Enabled = False
    End Sub

    Private Sub btnExitSP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExitSP.Click
        Me.PNwithDraw.Visible = False
        Me.pnChecKPoint.Visible = False
        Me.pnCampaignMaster.Visible = False
        Me.PNsaveNewItem.Visible = False
        Me.PNsaveSpecialPoint.Visible = False

    End Sub

    Private Sub btnClearSP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearSP.Click
        'เคลียร์หน้าบันทึกแต้มพิเศษ
        Me.txtSPDocno.Text = ""
        Me.dtpSPDocdate.Value = Now.Date
        Me.txtSPmemberid.Text = ""
        Me.txtSPARCode.Text = ""
        Me.txtSParName.Text = ""
        Me.txtSPpoint.Text = ""
        Me.txtIssue.Text = ""
        Me.lblcancel.Visible = False
        Me.lblConfirm.Visible = True
        Me.lblConfirm.Text = "N"
        Me.btnSaveSP.Enabled = False
        Me.btnFNmember.Enabled = False
        Me.btnNewSPpoint.Enabled = True

    End Sub

    Private Sub btnNewItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewItem.Click
        'รหัสสินค้า
        Me.txtItmPictPath.Text = "M:\ของรางวัล\noimage.jpg"
        dlgItemsearch.Show()
        vIsInsertItem = 1
        'frmMainMember.Enabled = False
    End Sub

    Private Sub btnBrows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrows.Click
        Dim vPath As String
        Me.OpenFdialog.InitialDirectory = "M:\\"
        OpenFdialog.Filter = "JPEG Files|*.jpg|GIF Files|*.gif"
        'frmMainMember.Enabled = False
        If Me.OpenFdialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            vPath = Me.OpenFdialog.FileName
            Me.txtItmPictPath.Text = vPath
            Me.pbxItem.Image = Image.FromFile(vPath)
            Me.pbxItem.Show()
            'frmMainMember.Enabled = True
            Me.btnSaveitm.Enabled = True
        End If
        'frmMainMember.Enabled = True
    End Sub

    Private Sub btnSaveitm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveitm.Click
        Dim rwQry As String
        Dim vItemcode As String
        Dim vItemName As String
        Dim vItemUnit As String
        Dim vItemPoint As Integer
        Dim vItemAmount As Integer
        Dim vItemPath As String
        If Me.txtNewitmcode.Text <> "" And Me.txtItmName.Text <> "" And Me.txtItmUnitcode.Text <> "" And Me.txtItmAmount.Text <> "" And Me.txtItmPoint.Text <> "" And Me.txtItmPictPath.Text <> "" Then
            vItemcode = Me.txtNewitmcode.Text
            vItemName = Me.txtItmName.Text

            vItemUnit = Me.txtItmUnitcode.Text
            vItemPoint = Me.txtItmPoint.Text
            vItemAmount = Me.txtItmAmount.Text
            vItemPath = Me.txtItmPictPath.Text
            rwQry = "exec dbo.USP_VP_RewardAssign'" & vIsInsertItem & "','" & vItemcode & "','" & vItemName & "','" & vItemUnit & "','" & vItemPoint & "','" & vItemPath & "'"
            cmd = New SqlCommand
            With cmd
                .CommandType = CommandType.Text
                .CommandText = rwQry
                .Connection = vConnection
                .ExecuteNonQuery()
            End With
            MsgBox("บันทึกข้อมูลเรียบร้อยแล้ว.", MsgBoxStyle.Information, "Infomation")
            Call clrSaveNewItemform()
        Else
            MsgBox("คุณใส่ข้อมูลไม่ครบกรุณาตรวจสอบใหม่", MsgBoxStyle.Critical, "Warning")
        End If

    End Sub
    Private Sub clrSaveNewItemform()
        Me.txtNewitmcode.Text = ""
        Me.txtItmName.Text = ""
        Me.txtItmUnitcode.Text = ""
        Me.txtItmPoint.Text = ""
        Me.txtItmPictPath.Text = ""
        Me.txtItmAmount.Text = ""
        Me.pbxItem.Image = Nothing
        Me.txtNewitmcode.ReadOnly = True
        Me.txtItmName.ReadOnly = True
        Me.txtItmAmount.ReadOnly = True
        Me.txtItmUnitcode.ReadOnly = True
        Me.btnSaveitm.Enabled = False
        Me.btnFinditm.Enabled = True
        Me.btnNewItem.Enabled = True
        Me.btnBrows.Enabled = True
    End Sub
    
    Private Sub btnItmClr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnItmClr.Click
        Me.txtNewitmcode.Text = ""
        Me.txtItmName.Text = ""
        Me.txtItmUnitcode.Text = ""
        Me.txtItmPoint.Text = ""
        Me.txtItmPictPath.Text = ""
        Me.txtItmAmount.Text = ""
        Me.pbxItem.Image = Nothing
        Me.txtNewitmcode.ReadOnly = True
        Me.txtItmName.ReadOnly = True
        Me.txtItmAmount.ReadOnly = True
        Me.txtItmUnitcode.ReadOnly = True
        Me.btnSaveitm.Enabled = False
        Me.btnFinditm.Enabled = True
        Me.btnNewItem.Enabled = True
        Me.btnBrows.Enabled = True
    End Sub

    Private Sub btnExititm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExititm.Click
        Me.PNwithDraw.Visible = False
        Me.pnChecKPoint.Visible = False
        Me.pnCampaignMaster.Visible = False
        Me.PNsaveNewItem.Visible = False
        Me.PNsaveSpecialPoint.Visible = False

    End Sub

    Private Sub btnFinditm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinditm.Click
        dlgSearchRwItem.Show()
        vIsInsertItem = 0
        'frmMainMember.Enabled = False
    End Sub
    Private Sub CheckRwPoint()
        Dim x As Integer
        Dim NPoint As Double
        Dim vSumPoint As Double
        Dim vPointAmount As Double
        vPointAmount = CDbl(Me.txtPointAmount.Text)
        For x = 0 To Me.LVRewardChange.Items.Count - 1
            If Me.LVRewardChange.Items(x).Checked = True Then
                NPoint = Me.LVRewardChange.Items(x).SubItems(6).Text
                vSumPoint = vSumPoint + NPoint
                If vSumPoint > vPointAmount Then
                    MsgBox("แต้มสมาชิกที่เหลือ ไม่พอเบิกรายการนี้เพิ่มได้", MsgBoxStyle.Critical, "Warning")
                    Me.LVRewardChange.Items(x).Checked = False
                End If
            End If
        Next
    End Sub

    Private Sub LVRewardChange_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles LVRewardChange.ItemCheck
        '  Call CheckRwPoint()
    End Sub

    Private Sub LVRewardChange_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles LVRewardChange.ItemChecked
        Call CheckRwPoint()
    End Sub


    Private Sub LVRewardChange_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LVRewardChange.SelectedIndexChanged
        Dim ix As Integer
        Dim iPictPath As String
        For ix = 0 To Me.LVRewardChange.Items.Count - 1
            If Me.LVRewardChange.Items(ix).Selected = True Then
                iPictPath = Me.LVRewardChange.Items(ix).SubItems(7).Text
                On Error GoTo picterr
                Me.pbxPicRewad.Image = Image.FromFile(iPictPath)
                Me.pbxPicRewad.Show()
            End If
        Next
picterr:
        If Err.Description <> "" Then
            iPictPath = "M:\ของรางวัล\noimage.jpg"
            Me.pbxPicRewad.Image = Image.FromFile(iPictPath)
            Me.pbxPicRewad.Show()
            ' MsgBox("ไม่พบไฟล์รูปภาพที่ต้องการ หรือชื่อไฟล์ไม่ถูกต้อง", MsgBoxStyle.Critical, "Error")
        End If
    End Sub

    Private Sub btnCHKpoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCHKpoint.Click
        Me.txtMemberid.ReadOnly = True
        Me.txtARcode.ReadOnly = True
        Me.txtMemberName.ReadOnly = True
        Me.txtbgDate.ReadOnly = True
        Me.txtExpDate.ReadOnly = True
        Me.txtPointAmount.ReadOnly = True
        Me.txtCPName.ReadOnly = True
        Me.txtCPstDate.ReadOnly = True
        Me.txtCPstpDate.ReadOnly = True
        Me.pnChecKPoint.Show()
        Me.pnChecKPoint.BringToFront()
        Me.pnCampaignMaster.Visible = False
        Me.PNsaveNewItem.Visible = False
        Me.PNsaveSpecialPoint.Visible = False
        Me.PNwithDraw.Visible = False
        Me.btnSearch.Enabled = True
        Me.btnChangeRW.Enabled = False
        Me.btnFindMember.Enabled = True
        Me.Text = ":: ตรวจสอบแต้มสมาชิก"
        v = 1
    End Sub

    Private Sub btnMNdwPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMNdwPoint.Click
        Call cbxCampaignDW()
        Me.txtDwDoc.ReadOnly = True
        Me.txtdwMemberID.ReadOnly = True
        Me.txtdwARcode.ReadOnly = True
        Me.txtdwArname.ReadOnly = True
        Me.txtDWtotalPoint.ReadOnly = True
        Me.txtTotaldwAmount.ReadOnly = True
        Me.txtTotaldwPoint.ReadOnly = True
        Me.pnChecKPoint.Visible = False
        Me.pnCampaignMaster.Visible = False
        Me.PNsaveNewItem.Visible = False
        Me.PNsaveSpecialPoint.Visible = False
        Me.PNwithDraw.Visible = True
        Me.PNwithDraw.BringToFront()
        Me.btnPrintDwp.Enabled = False
        Me.lblisConfirm.Visible = False
        Me.Text = ":: เบิกแต้มและแลกของรางวัล"
    End Sub

    Private Sub btnMNspPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMNspPoint.Click
        Call cbxSpecialPoint()
        Me.lblConfirm.Visible = False
        Me.lblcancel.Visible = False
        Me.pnChecKPoint.Visible = False
        Me.pnCampaignMaster.Visible = False
        Me.PNsaveNewItem.Visible = False
        Me.PNsaveSpecialPoint.Visible = True
        Me.PNwithDraw.Visible = False
        Me.PNsaveSpecialPoint.BringToFront()
        Me.txtSPDocno.ReadOnly = True
        Me.txtSPmemberid.ReadOnly = True
        Me.txtSPARCode.ReadOnly = True
        Me.txtSParName.ReadOnly = True
        Me.btnFNmember.Enabled = False
        Me.btnFindSP.Enabled = True
        Me.btnSaveCP.Enabled = False
        Me.Text = ":: บันทึกการให้แต้มพิเศษ"
    End Sub

    Private Sub btnMNnewRW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMNnewRW.Click
        Me.txtNewitmcode.ReadOnly = True
        Me.txtItmPictPath.ReadOnly = True
        Me.pnChecKPoint.Visible = False
        Me.pnCampaignMaster.Visible = False
        Me.PNsaveNewItem.Visible = True
        Me.PNsaveSpecialPoint.Visible = False
        Me.PNwithDraw.Visible = False
        Me.PNsaveNewItem.BringToFront()
        Me.Text = ":: บันทึกรายการของรางวัล"
    End Sub

    Private Sub btnMNnewCP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMNnewCP.Click
        Me.txtCPno.ReadOnly = True
        Me.pnChecKPoint.Visible = False
        Me.pnCampaignMaster.Visible = True
        Me.PNsaveNewItem.Visible = False
        Me.PNsaveSpecialPoint.Visible = False
        Me.PNwithDraw.Visible = False
        Me.pnCampaignMaster.BringToFront()
        Me.gbxCPmaster.Enabled = True
        Me.btnSaveCP.Enabled = False
        Me.Text = ":: บันทึกรายการส่งเสริมการขาย Campaign Master"
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.pnChecKPoint.Visible = False
        Me.pnCampaignMaster.Visible = False
        Me.PNsaveNewItem.Visible = False
        Me.PNsaveSpecialPoint.Visible = False
        Me.PNwithDraw.Visible = False
        Call clrCHKpoint()
    End Sub

    Private Sub btnMNexit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMNexit.Click
        Me.Close()
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Call clrCHECKpoint()
    End Sub
    Private Sub clrCHECKpoint()
        Me.txtMemberid.Text = ""
        Me.txtARcode.Text = ""
        Me.txtMemberName.Text = ""
        Me.txtbgDate.Text = ""
        Me.txtExpDate.Text = ""
        Me.txtPointAmount.Text = ""
        Me.txtCPName.Text = ""
        Me.txtCPstDate.Text = ""
        Me.txtCPstpDate.Text = ""
        Me.LVRewardChange.Items.Clear()
        Me.LVPointDetailMTR.Items.Clear()
        Me.LVSpecialPoint.Items.Clear()
        Me.LVDrwPoint.Items.Clear()
        Me.txtTotalAmount.Text = ""
        Me.txtTotalPoint.Text = ""
        Me.txtSPecialPoint.Text = ""
        Me.txtDrwPointAmount.Text = ""
        Me.pbxPicRewad.Image = Nothing
        Me.btnChangeRW.Enabled = False
        Me.btnSearch.Enabled = True
        Me.btnFindMember.Enabled = True
    End Sub
    Private Sub clrCHKpoint()
        Me.txtMemberid.Text = ""
        Me.txtARcode.Text = ""
        Me.txtMemberName.Text = ""
        Me.txtbgDate.Text = ""
        Me.txtExpDate.Text = ""
        Me.txtPointAmount.Text = ""
        Me.txtCPName.Text = ""
        Me.txtCPstDate.Text = ""
        Me.txtCPstpDate.Text = ""
        Me.LVRewardChange.Items.Clear()
        Me.LVPointDetailMTR.Items.Clear()
        Me.LVSpecialPoint.Items.Clear()
        Me.LVDrwPoint.Items.Clear()
        Me.txtTotalAmount.Text = ""
        Me.txtTotalPoint.Text = ""
        Me.txtSPecialPoint.Text = ""
        Me.txtDrwPointAmount.Text = ""
        Me.pbxPicRewad.Image = Nothing
        Me.btnChangeRW.Enabled = False
        Me.btnSearch.Enabled = True
        Me.btnFindMember.Enabled = True
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        iType = 1
        MemberSearch.Show()
        MemberSearch.txtFind.Focus()
    End Sub

    Private Sub btnNewSPpoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewSPpoint.Click
        Call NewSPdoc()
        Me.btnFindSP.Enabled = False
    End Sub
    
    Private Sub dgvReadonly()
        Me.dgvDWreward.Columns(0).ReadOnly = True
        Me.dgvDWreward.Columns(1).ReadOnly = True
        Me.dgvDWreward.Columns(2).ReadOnly = True
        Me.dgvDWreward.Columns(3).DefaultCellStyle.BackColor = Color.YellowGreen
        Me.dgvDWreward.Columns(4).ReadOnly = True
        Me.dgvDWreward.Columns(5).ReadOnly = True
        Me.dgvDWreward.Columns(6).ReadOnly = True
        Me.dgvDWreward.Columns(7).ReadOnly = True
        Me.dgvDWreward.Columns(8).ReadOnly = True

    End Sub

   
End Class
