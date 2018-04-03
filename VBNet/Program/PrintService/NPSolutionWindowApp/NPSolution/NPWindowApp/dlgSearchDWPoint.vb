Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic
Public Class dlgSearchDWPoint
    Dim dt As New DataTable
    Dim dataQry As String
    Dim i As Integer
    Dim vTXTsearch As String
    Dim iLvw As ListViewItem
    Dim iDWdocno As String

    Private Sub LoadDataDetail()
        dataQry = "exec dbo.USP_VP_WithDrawSearchSub '" & iDWdocno & "'"
        da = New SqlDataAdapter(dataQry, vConnection)
        ds = New DataSet
        da.Fill(ds, "rsDT")
        dt = ds.Tables("rsDT")
        If dt.Rows.Count > 0 Then
            frmPoint01.txtDwDoc.Text = dt.Rows(0).Item("DocNo")
            frmPoint01.dtpDWDate.Text = dt.Rows(0).Item("Docdate")
            frmPoint01.cbxDW.SelectedValue = dt.Rows(0).Item("CampaignCode")
            frmPoint01.txtdwMemberID.Text = dt.Rows(0).Item("memberid")
            frmPoint01.txtdwARcode.Text = dt.Rows(0).Item("Arcode")
            frmPoint01.txtdwArname.Text = dt.Rows(0).Item("arName")
            cf = dt.Rows(0).Item("isconfirm")
            '-------------------------------
            Dim i As Integer
            Dim dt1 As New DataTable("vData")
            Dim dr As DataRow
            Dim inum As Integer
            Dim iPAmount As Integer
            Dim iPPoint As Integer
            Dim iPTotalAmount As Integer
            Dim iPTotalPoint As Integer

            'For x = 0 To Me.LVRewardChange.Items.Count - 1
            '    If Me.LVRewardChange.Items(x).Checked = True Then
            dt1.Columns.Add("ลำดับ", GetType(Integer))
            dt1.Columns.Add("รหัสสินค้า", GetType(String))
            dt1.Columns.Add("ชื่อสินค้า", GetType(String))
            dt1.Columns.Add("จำนวน", GetType(String))
            dt1.Columns.Add("หน่วยนับ", GetType(String))
            dt1.Columns.Add("มูลค่า:หน่วย", GetType(String))
            dt1.Columns.Add("แต้ม:หน่วย", GetType(String))
            dt1.Columns.Add("มูลค่า", GetType(String))
            dt1.Columns.Add("แต้ม", GetType(String))
            For i = 0 To dt.Rows.Count - 1
                inum = inum + 1
                dr = dt1.NewRow
                dr("ลำดับ") = inum
                dr("รหัสสินค้า") = dt.Rows(i).Item("itemcode")
                dr("ชื่อสินค้า") = dt.Rows(i).Item("itemname")
                dr("จำนวน") = dt.Rows(i).Item("Qty")
                dr("หน่วยนับ") = dt.Rows(i).Item("unitcode")
                'iAmount = Me.LVRewardChange.Items(i).SubItems(5).Text
                dr("มูลค่า:หน่วย") = dt.Rows(i).Item("Amount")
                'iTotalAmount = (iTotalAmount + iAmount)

                dr("แต้ม:หน่วย") = dt.Rows(i).Item("Point")
                iPPoint = dt.Rows(i).Item("detailTotalAmount")
                dr("มูลค่า") = iPPoint
                iPTotalPoint = (iPTotalPoint + iPPoint)
                iPAmount = dt.Rows(i).Item("detailTotalPoint")
                dr("แต้ม") = iPAmount
                dt1.Rows.Add(dr)
                iPTotalAmount = (iPTotalAmount + iPAmount)
            Next
            frmPoint01.dgvDWreward.DataSource = dt1
            frmPoint01.txtTotaldwAmount.Text = Format(iPTotalAmount, "##,##0.00")
            frmPoint01.txtTotaldwPoint.Text = Format(iPTotalPoint, "##,##0.00")
            PsaveDWstatus = 1
            If cf = 0 Then
                frmPoint01.lblisConfirm.Text = "--N--"
                frmPoint01.lblisConfirm.BackColor = Color.Red
                frmPoint01.lblisConfirm.Visible = True
            ElseIf cf = 1 Then
                frmPoint01.lblisConfirm.Text = "--CF--"
                frmPoint01.lblisConfirm.BackColor = Color.Green
                frmPoint01.lblisConfirm.Visible = True
                frmPoint01.btnNewDW.Enabled = False
                frmPoint01.btnDWFMB.Enabled = False
            End If
            frmPoint01.txtDWtotalPoint.Visible = False
            frmPoint01.Label35.Visible = False
            'frmPoint01.btnPrintDwp.Enabled = True
        End If
    End Sub
    Private Sub viewDataDGV()
        'Dim i As Integer
        'Dim dt1 As New DataTable("vData")
        'Dim dr As DataRow
        'Dim inum As Integer
        'Dim iAmount As Integer
        'Dim iPoint As Integer
        'Dim iTotalAmount As Integer
        'Dim iTotalPoint As Integer

        ''For x = 0 To Me.LVRewardChange.Items.Count - 1
        ''    If Me.LVRewardChange.Items(x).Checked = True Then
        'dt1.Columns.Add("ลำดับ", GetType(Integer))
        'dt1.Columns.Add("รหัสสินค้า", GetType(String))
        'dt1.Columns.Add("ชื่อสินค้า", GetType(String))
        'dt1.Columns.Add("จำนวน", GetType(String))
        'dt1.Columns.Add("หน่วยนับ", GetType(String))
        'dt1.Columns.Add("มูลค่า:หน่วย", GetType(String))
        'dt1.Columns.Add("แต้ม:หน่วย", GetType(String))
        'dt1.Columns.Add("มูลค่า", GetType(String))
        'dt1.Columns.Add("แต้ม", GetType(String))
        'If frmPoint01.LVRewardChange.Items.Count > 0 Then
        '    For i = 0 To dt.Rows.Count - 1
        '        If Me.LVRewardChange.Items(i).Checked = True Then
        '            inum = inum + 1
        '            dr = dt.NewRow
        '            dr("ลำดับ") = inum
        '            dr("รหัสสินค้า") = Me.LVRewardChange.Items(i).SubItems(1).Text
        '            dr("ชื่อสินค้า") = Me.LVRewardChange.Items(i).SubItems(2).Text
        '            dr("จำนวน") = 1
        '            dr("หน่วยนับ") = Me.LVRewardChange.Items(i).SubItems(4).Text
        '            iAmount = Me.LVRewardChange.Items(i).SubItems(5).Text
        '            dr("มูลค่า:หน่วย") = iAmount
        '            iTotalAmount = (iTotalAmount + iAmount)
        '            iPoint = Me.LVRewardChange.Items(i).SubItems(6).Text
        '            dr("แต้ม:หน่วย") = iPoint
        '            iTotalPoint = (iTotalPoint + iPoint)
        '            dr("มูลค่า") = iAmount
        '            dr("แต้ม") = iPoint
        '        End If
        '    Next
        '    Me.dgvDWreward.DataSource = dt1
        '    Me.txtdwMemberID.Text = Me.txtMemberid.Text
        '    Me.txtdwARcode.Text = Me.txtARcode.Text
        '    Me.txtdwArname.Text = Me.txtMemberName.Text
        '    Me.txtDWtotalPoint.Text = Me.txtPointAmount.Text
        '    Me.cbxSPCampaign.SelectedValue = Me.lblCPCode.Text 'value ที่ได้มาจาก label CampaingCode
        '    Me.txtTotaldwAmount.Text = Format(iTotalAmount, "##,##0.00")
        '    Me.txtTotaldwPoint.Text = Format(iTotalPoint, "##,##0.00")
        'End If
    End Sub

    Private Sub dlgSearchDWPoint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
        Me.Text = ":: ค้นหาใบเบิกแต้ม"
        Me.btnSelect.Enabled = False
    End Sub
    Private Sub LoadData()
        If Me.txtDwDocFND.Text <> "" Then
            Me.LVdwDoc.Items.Clear()
            On Error GoTo errDes
            vTXTsearch = Me.txtDwDocFND.Text
            dataQry = "exec dbo.USP_VP_WithDrawSearch '" & vTXTsearch & "'"
            da = New SqlDataAdapter(dataQry, vConnection)
            ds = New DataSet
            da.Fill(ds, "vdt")
            dt = ds.Tables("vdt")
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    iLvw = Me.LVdwDoc.Items.Add(dt.Rows(i).Item("DocNo"))
                    iLvw.SubItems.Add(0).Text = dt.Rows(i).Item("DocDate")
                    iLvw.SubItems.Add(0).Text = dt.Rows(i).Item("Arcode")
                    iLvw.SubItems.Add(0).Text = dt.Rows(i).Item("Memberid")
                    iLvw.SubItems.Add(0).Text = dt.Rows(i).Item("arName")
                    iLvw.SubItems.Add(0).Text = dt.Rows(i).Item("FinalTotalPoint")
                    iLvw.SubItems.Add(0).Text = dt.Rows(i).Item("FinaltotalAmount")
                Next
                
            Else
                MsgBox("ไม่พบข้อมูลที่ต้องการ ลองใส่คำค้นหาใหม่", MsgBoxStyle.Information, "Information")
            End If

            'Else
            '    MsgBox("กรุณาใส่คำค้นหาก่อน", MsgBoxStyle.Critical, "Warning")
        End If
errDes:
        If Err.Description <> "" Then
            MsgBox("ไม่พบข้อมูลที่ค้นหา", MsgBoxStyle.Information, "Information")
        End If
    End Sub

    Private Sub btnFindDwdoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindDwdoc.Click
        Call LoadData()
    End Sub

    Private Sub txtDwDoc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDwDocFND.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call LoadData()
        End If
    End Sub

    Private Sub btnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click
        Call SLDwDocno()
    End Sub
    Private Sub SLDwDocno()
        Dim x As Integer
        If Me.LVdwDoc.Items.Count > 0 Then
            For x = 0 To Me.LVdwDoc.Items.Count - 1
                If Me.LVdwDoc.Items(x).Selected Then
                    iDWdocno = Me.LVdwDoc.Items(x).SubItems(0).Text
                    Call LoadDataDetail()
                    Call dgvRWReadonly()
                End If
            Next
            
        Else
            MsgBox("คุณยังไม่ได้เลือกเอกสาร", MsgBoxStyle.Critical, "Warning")
        End If
        frmPoint01.btnPrintDwp.Enabled = True
        Me.Close()
        'frmPoint01.Show()

    End Sub

    Private Sub btnExitdw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExitdw.Click
        Me.Close()
    End Sub


    Private Sub LVdwDoc_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles LVdwDoc.ItemSelectionChanged
        Me.btnSelect.Enabled = True
    End Sub

    
    Private Sub LVdwDoc_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LVdwDoc.MouseDoubleClick
        Call SLDwDocno()
    End Sub
    Private Sub dgvRWReadonly()
        frmPoint01.dgvDWreward.Columns(0).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(1).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(2).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(3).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(4).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(5).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(6).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(7).ReadOnly = True
        frmPoint01.dgvDWreward.Columns(8).ReadOnly = True

    End Sub
End Class
