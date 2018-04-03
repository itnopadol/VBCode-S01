Imports System
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic

Public Class MemberSearch
    Dim fQrystr As String
    Dim da As SqlDataAdapter
    Dim ds As New DataSet
    Dim dt As New DataTable
    Dim i As Integer
    Dim iCPcode As String
    Dim PointAM As Double
    Dim iARCODE As String
    Dim SPPoint As Double
    Dim TotalSPPoint As Double
    Dim DWPoint As Double
    Dim ToTalDwPoint As Double
    Dim vTotalAmount As Double


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ''Dim iArCode As String
        'If iType = 1 Then
        '    Dim vARCode As String

        '    i = Me.LVmeberpointFind.SelectedItems(0).Index
        '    'iArCode = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        '    vARCode = Me.LVmeberpointFind.Items(i).SubItems(0).Text
        '    ' MsgBox(vARCode)

        '    frmPoint01.txtMemberid.Text = Me.LVmeberpointFind.Items(i).SubItems(0).Text
        '    frmPoint01.txtARcode.Text = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        '    frmPoint01.txtMemberName.Text = Me.LVmeberpointFind.Items(i).SubItems(2).Text
        '    vTotalAmount = Me.LVmeberpointFind.Items(i).SubItems(3).Text
        '    frmPoint01.txtTotalAmount.Text = Format(vTotalAmount, "##,##0.00")
        '    PointAM = Me.LVmeberpointFind.Items(i).SubItems(4).Text
        '    frmPoint01.txtPointAmount.Text = PointAM
        '    frmPoint01.txtTotalPoint.Text = PointAM
        '    frmPoint01.txtbgDate.Text = Me.LVmeberpointFind.Items(i).SubItems(5).Text
        '    frmPoint01.txtExpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(6).Text
        '    PcampaignCode = Me.LVmeberpointFind.Items(i).SubItems(7).Text
        '    frmPoint01.txtCPName.Text = Me.LVmeberpointFind.Items(i).SubItems(8).Text
        '    frmPoint01.txtCPstDate.Text = Me.LVmeberpointFind.Items(i).SubItems(9).Text
        '    frmPoint01.txtCPstpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(10).Text
        '    frmPoint01.pnChecKPoint.Show()
        '    frmPoint01.pnChecKPoint.Visible = True
        '    'frmPoint01 = New frmPoint01
        '    'frmPoint01.MdiParent = frmMainMember
        '    ' frmPoint01.BringToFront()
        '    'frmPoint01.Show()
        'End If

        ''MsgBox(i)

        ''Call frmP01Readonly()
        ''Call PointListforChange()
        ''Call PointDetailDesc()
        ''Call PointDetailSpecial()
        ''Call PointDetailwihtDraw()
        ''frmPoint01.btnChangeRW.Enabled = True
        ''frmPoint01.btnSearch.Enabled = False
        ''frmPoint01.btnFindMember.Enabled = False
        ' ''frmPoint01.Show()
        ''Me.Close()
        ' '' เพิ่มข้อมูลในหน้าแลกของรางวัล
        ' ''ElseIf iType = 2 Then
        ''For i = 0 To Me.LVmeberpointFind.Items.Count - 1
        ''    If Me.LVmeberpointFind.Items(i).Selected Then
        ''        'iArCode = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        ''        frmPoint01.txtdwMemberID.Text = Me.LVmeberpointFind.Items(i).SubItems(0).Text
        ''        frmPoint01.txtdwARcode.Text = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        ''        frmPoint01.txtdwArname.Text = Me.LVmeberpointFind.Items(i).SubItems(2).Text
        ''        PointAM = Me.LVmeberpointFind.Items(i).SubItems(4).Text
        ''        frmPoint01.txtDWtotalPoint.Text = PointAM
        ''        PcampaignCode = Me.LVmeberpointFind.Items(i).SubItems(7).Text
        ''        'frmPoint01.txtbgDate.Text = Me.LVmeberpointFind.Items(i).SubItems(5).Text
        ''        'frmPoint01.txtExpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(6).Text
        ''        'PcampaignCode = Me.LVmeberpointFind.Items(i).SubItems(7).Text
        ''        'frmPoint01.txtCPName.Text = Me.LVmeberpointFind.Items(i).SubItems(8).Text
        ''        'frmPoint01.txtCPstDate.Text = Me.LVmeberpointFind.Items(i).SubItems(9).Text
        ''        'frmPoint01.txtCPstpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(10).Text

        ''    End If
        ''Next
        ''frmPoint01.lblCPCode.Text = PcampaignCode
        ' ''Me.Close()

        ' ''ElseIf iType = 3 Then
        ''For i = 0 To Me.LVmeberpointFind.Items.Count - 1
        ''    If Me.LVmeberpointFind.Items(i).Selected Then
        ''        'iArCode = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        ''        frmPoint01.txtSPmemberid.Text = Me.LVmeberpointFind.Items(i).SubItems(0).Text
        ''        frmPoint01.txtSPARCode.Text = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        ''        frmPoint01.txtSParName.Text = Me.LVmeberpointFind.Items(i).SubItems(2).Text
        ''        PointAM = Me.LVmeberpointFind.Items(i).SubItems(4).Text
        ''        'PcampaignCode = Me.LVmeberpointFind.Items(i).SubItems(7).Text
        ''        'frmPoint01.txtbgDate.Text = Me.LVmeberpointFind.Items(i).SubItems(5).Text
        ''        'frmPoint01.txtExpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(6).Text
        ''        'PcampaignCode = Me.LVmeberpointFind.Items(i).SubItems(7).Text
        ''        'frmPoint01.txtCPName.Text = Me.LVmeberpointFind.Items(i).SubItems(8).Text
        ''        'frmPoint01.txtCPstDate.Text = Me.LVmeberpointFind.Items(i).SubItems(9).Text
        ''        'frmPoint01.txtCPstpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(10).Text

        ''    End If
        ''Next
        ' ''End If
        ''PsaveSPpoint = 0
        ' ''frmPoint01.Show()
        'Me.Close()
        ' ''frmPoint01.lblCPCode.Text = PcampaignCode
    End Sub
    'รายการของรางวัลที่แลกได้ รายละเอียด Tab1
    Private Sub PointListforChange()
        Dim xLV As ListViewItem
        Dim Pqry As String
        Dim xnum As Integer
        Dim vStkQty As Double
        Dim vAmount As Double
        Dim vPoint As Double

        Pqry = "exec dbo.USP_VP_CheckAVLReward '" & PointAM & "'"
        da = New SqlDataAdapter(Pqry, vConnection)
        ds = New DataSet
        da.Fill(ds, "pchg")
        dt = ds.Tables("pchg")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                xnum = xnum + 1
                xLV = frmPoint01.LVRewardChange.Items.Add(xnum)
                xLV.SubItems.Add(0).Text = dt.Rows(i).Item("itemcode")
                xLV.SubItems.Add(0).Text = dt.Rows(i).Item("itemname")
                vStkQty = dt.Rows(i).Item("stockqty")
                vAmount = dt.Rows(i).Item("Amount")
                vPoint = dt.Rows(i).Item("Point")
                xLV.SubItems.Add(0).Text = Format(vStkQty, "##,##0.00")
                xLV.SubItems.Add(0).Text = dt.Rows(i).Item("unitcode")
                xLV.SubItems.Add(0).Text = Format(vAmount, "##,##0.00")
                xLV.SubItems.Add(0).Text = Format(vPoint, "##,##0.00")
                xLV.SubItems.Add(0).Text = dt.Rows(i).Item("PicturePath")
            Next
        End If
    End Sub
    ' รายการสะสมแต้ม รายละเอียด Tab2
    Private Sub PointDetailDesc()
        Dim dQry As String
        Dim dLV As ListViewItem

        Dim vRefPrice As Double
        Dim vPrice As Double
        Dim vQty As Double
        Dim vDiscountAmount As Double
        Dim vAmount As Double
        Dim vPointBalance As Double
        Dim vPriceDiff As Double
        Dim vDiscountTotal As Double
        Dim vRefPoint As Double
        Dim vPointFinal As Double

        'Dim dnum As Integer
        dQry = "exec dbo.USP_VP_PointDesc '" & iARCODE & "','" & iCPcode & "'"
        da = New SqlDataAdapter(dQry, vConnection)
        ds = New DataSet
        da.Fill(ds, "pdesc")
        dt = ds.Tables("pdesc")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                dLV = frmPoint01.LVPointDetailMTR.Items.Add(dt.Rows(i).Item("docdate"))
                vRefPrice = dt.Rows(i).Item("Refprice")
                vPrice = dt.Rows(i).Item("Price")
                vQty = dt.Rows(i).Item("Qty")
                vDiscountAmount = dt.Rows(i).Item("DiscountAmount")
                vAmount = dt.Rows(i).Item("Amount")
                vPointBalance = dt.Rows(i).Item("PointBalance")
                vPriceDiff = dt.Rows(i).Item("PriceDiff")
                vDiscountTotal = dt.Rows(i).Item("DiscountTotal")
                vRefPoint = dt.Rows(i).Item("RefPoint")
                vPointFinal = dt.Rows(i).Item("PointFinal")

                dLV.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                dLV.SubItems.Add(0).Text = dt.Rows(i).Item("itemcode")
                dLV.SubItems.Add(0).Text = dt.Rows(i).Item("itemname")
                dLV.SubItems.Add(0).Text = dt.Rows(i).Item("unitcode")
                dLV.SubItems.Add(0).Text = Format(vRefPrice, "##,##0.00")
                dLV.SubItems.Add(0).Text = Format(vPrice, "##,##0.00")
                dLV.SubItems.Add(0).Text = Format(vQty, "##,##0.00")
                dLV.SubItems.Add(0).Text = Format(vDiscountAmount, "##,##0.00")
                dLV.SubItems.Add(0).Text = Format(vAmount, "##,##0.00")
                dLV.SubItems.Add(0).Text = dt.Rows(i).Item("RefLevel")
                dLV.SubItems.Add(0).Text = Format(vPointBalance, "##,##0.00")
                dLV.SubItems.Add(0).Text = Format(vPriceDiff, "##,##0.00")
                dLV.SubItems.Add(0).Text = Format(vDiscountTotal, "##,##0.00")
                dLV.SubItems.Add(0).Text = Format(vRefPoint, "##,##0.00")
                dLV.SubItems.Add(0).Text = Format(vPointFinal, "##,##0.00")
            Next
        End If
    End Sub
    'รายละเอียดแต้มพิเศษ Tab 2
    Private Sub PointDetailSpecial()
        Dim spQry As String
        Dim spLV As ListViewItem
        spQry = "exec dbo.USP_VP_CheckPointSpecial '" & iCPcode & "','" & iARCODE & "'"
        da = New SqlDataAdapter(spQry, vConnection)
        ds = New DataSet
        da.Fill(ds, "psp")
        dt = ds.Tables("psp")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                spLV = frmPoint01.LVSpecialPoint.Items.Add(dt.Rows(i).Item("docdate"))
                spLV.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                SPPoint = dt.Rows(i).Item("Point")
                spLV.SubItems.Add(0).Text = SPPoint
                TotalSPPoint = (TotalSPPoint + SPPoint)
                spLV.SubItems.Add(0).Text = dt.Rows(i).Item("Reason")
            Next
            frmPoint01.txtSPecialPoint.Text = Format(TotalSPPoint, "##,##0.00")
        End If
    End Sub
    ' รายละเอียดแต้มที่เบิก Tab 2
    Private Sub PointDetailwihtDraw()
        Dim dwQry As String
        Dim dwLV As ListViewItem
        Dim vQty As Double
        Dim vTotalAmount As Double
        Dim vTotalPoint As Double

        dwQry = "exec dbo.USP_VP_CheckPointwithDraw '" & iCPcode & "','" & iARCODE & "'"
        da = New SqlDataAdapter(dwQry, vConnection)
        ds = New DataSet
        da.Fill(ds, "pdw")
        dt = ds.Tables("pdw")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                dwLV = frmPoint01.LVDrwPoint.Items.Add(dt.Rows(i).Item("docdate"))
                dwLV.SubItems.Add(0).Text = dt.Rows(i).Item("docno")
                dwLV.SubItems.Add(0).Text = dt.Rows(i).Item("itemcode")
                dwLV.SubItems.Add(0).Text = dt.Rows(i).Item("itemname")
                vQty = dt.Rows(i).Item("Qty")
                vTotalAmount = dt.Rows(i).Item("TotalAmount")
                vTotalPoint = dt.Rows(i).Item("TotalPoint")
                dwLV.SubItems.Add(0).Text = Format(vQty, "##,##0.00")
                dwLV.SubItems.Add(0).Text = dt.Rows(i).Item("unitcode")
                dwLV.SubItems.Add(0).Text = Format(vTotalAmount, "##,##0.00")
                DWPoint = Format(vTotalPoint, "##,##0.00")
                dwLV.SubItems.Add(0).Text = DWPoint
                ToTalDwPoint = (ToTalDwPoint + DWPoint)
            Next
            frmPoint01.txtDrwPointAmount.Text = Format(ToTalDwPoint, "##,##0.00")
        End If
    End Sub
    Private Sub frmP01Readonly()
        frmPoint01.txtMemberid.ReadOnly = True
        frmPoint01.txtARcode.ReadOnly = True
        frmPoint01.txtMemberName.ReadOnly = True
        frmPoint01.txtPointAmount.ReadOnly = True
        frmPoint01.txtbgDate.ReadOnly = True
        frmPoint01.txtExpDate.ReadOnly = True
        frmPoint01.txtCPName.ReadOnly = True
        frmPoint01.txtCPstDate.ReadOnly = True
        frmPoint01.txtCPstpDate.ReadOnly = True
    End Sub
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        If Me.txtFind.Text <> "" Then
            Call FindMember()
        Else
            MsgBox("คุณยังไม่ได้ใส่คำค้นหา", MsgBoxStyle.Critical, "Warning")
        End If

    End Sub
    Private Sub FindMember()
        Call InitializeDataBase()
        Dim Lvmbp As ListViewItem
        Dim mbQry As String
        Dim iFindText As String
        Dim vMemberID As String

        Me.LVmeberpointFind.Items.Clear()
        iCPcode = Me.cbxCP.SelectedValue.ToString()
        iFindText = Me.txtFind.Text
        mbQry = "exec dbo.USP_VP_CheckPointSearch '" & iCPcode & "','" & iFindText & "'"
        da = New SqlDataAdapter(mbQry, vConnection)
        ds = New DataSet
        da.Fill(ds, "mbp")
        dt = ds.Tables("mbp")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                If IsDBNull(dt.Rows(i).Item("memberid")) Then
                    vMemberID = "ไม่ใช่สมาชิก"
                Else
                    vMemberID = dt.Rows(i).Item("memberid")
                End If
                Lvmbp = Me.LVmeberpointFind.Items.Add(vMemberID)
                iARCODE = dt.Rows(i).Item("arcode")
                Lvmbp.SubItems.Add(0).Text = iARCODE
                Lvmbp.SubItems.Add(0).Text = dt.Rows(i).Item("arname")
                If dt.Rows(i).Item("amount") Is DBNull.Value = False Then
                    Lvmbp.SubItems.Add(0).Text = Format(dt.Rows(i).Item("Amount"), "##,##0.00")
                Else
                    Lvmbp.SubItems.Add(0).Text = "0.00"
                End If
                Lvmbp.SubItems.Add(0).Text = Format(dt.Rows(i).Item("Point"), "##,##0.00")
                If IsDBNull(dt.Rows(i).Item("begindate")) Then
                    Lvmbp.SubItems.Add(0).Text = ""
                Else
                    Lvmbp.SubItems.Add(0).Text = dt.Rows(i).Item("begindate")
                End If
                If IsDBNull(dt.Rows(i).Item("ExpireDate")) Then
                    Lvmbp.SubItems.Add(0).Text = ""
                Else
                    Lvmbp.SubItems.Add(0).Text = dt.Rows(i).Item("ExpireDate")
                End If
                Lvmbp.SubItems.Add(0).Text = dt.Rows(i).Item("code")
                Lvmbp.SubItems.Add(0).Text = dt.Rows(i).Item("NameTh")
                If IsDBNull(dt.Rows(i).Item("StartDate")) Then
                    Lvmbp.SubItems.Add(0).Text = ""
                Else
                    Lvmbp.SubItems.Add(0).Text = dt.Rows(i).Item("StartDate")
                End If
                If IsDBNull(dt.Rows(i).Item("StopDate")) Then
                    Lvmbp.SubItems.Add(0).Text = ""
                Else
                    Lvmbp.SubItems.Add(0).Text = dt.Rows(i).Item("StopDate")
                End If
            Next
        Else
            MsgBox("ไม่พบข้อมูลลูกค้าที่ค้นหา", MsgBoxStyle.Information, "Information")
        End If
    End Sub
    Private Sub LoadCBX()
        Call InitializeDataBase()
        Dim cbxQry As String
        Dim ida As SqlDataAdapter
        Dim ids As New DataSet
        cbxQry = "exec dbo.USP_VP_CampaignList"
        ida = New SqlDataAdapter(cbxQry, vConnection)
        ida.Fill(ids, "vcbx")
        Me.cbxCP.DataSource = ids.Tables("vcbx")
        Me.cbxCP.DisplayMember = ("NameTh")
        Me.cbxCP.ValueMember = ("code")
    End Sub

    Private Sub cbxCP_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxCP.SelectedIndexChanged
        'iCPcode = Me.cbxCP.SelectedValue.ToString()
    End Sub

    Private Sub MemberSearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadCBX()
        Me.btnSLOK.Enabled = False
        'Me.txtFind.Focus()
    End Sub

    Private Sub txtFind_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyDown
        If Me.txtFind.Text <> "" Then
            If e.KeyCode = Keys.Enter Then
                Call FindMember()
            End If
        End If
    End Sub
    
    
    Private Sub btnSLOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSLOK.Click
        Call LoadMemberDetail()
        ''Dim iArCode As String
        'v = 1
        ''frmPoint01 = frmPoint01


        'If iType = 1 Then
        '    Dim vARCode As String
        '    i = Me.LVmeberpointFind.SelectedItems(0).Index
        '    'iArCode = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        '    vARCode = Me.LVmeberpointFind.Items(i).SubItems(0).Text
        '    ' MsgBox(vARCode)
        '    frmPoint01.txtMemberid.Text = Me.LVmeberpointFind.Items(i).SubItems(0).Text
        '    frmPoint01.txtARcode.Text = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        '    frmPoint01.txtMemberName.Text = Me.LVmeberpointFind.Items(i).SubItems(2).Text
        '    vTotalAmount = Me.LVmeberpointFind.Items(i).SubItems(3).Text
        '    frmPoint01.txtTotalAmount.Text = Format(vTotalAmount, "##,##0.00")
        '    PointAM = Me.LVmeberpointFind.Items(i).SubItems(4).Text
        '    frmPoint01.txtPointAmount.Text = Format(PointAM, "##,##0.00")
        '    frmPoint01.txtTotalPoint.Text = Format(PointAM, "##,##0.00")
        '    frmPoint01.txtbgDate.Text = Me.LVmeberpointFind.Items(i).SubItems(5).Text
        '    frmPoint01.txtExpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(6).Text
        '    PcampaignCode = Me.LVmeberpointFind.Items(i).SubItems(7).Text
        '    frmPoint01.txtCPName.Text = Me.LVmeberpointFind.Items(i).SubItems(8).Text
        '    frmPoint01.txtCPstDate.Text = Me.LVmeberpointFind.Items(i).SubItems(9).Text
        '    frmPoint01.txtCPstpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(10).Text

        '    '--------
        '    frmPoint01.pnChecKPoint.Show()
        '    frmPoint01.PNsaveNewItem.Hide()
        '    frmPoint01.PNsaveSpecialPoint.Hide()
        '    frmPoint01.PNwithDraw.Hide()
        '    frmPoint01.pnCampaignMaster.Hide()
        '    '----------
        '    Call frmP01Readonly()
        '    Call PointListforChange()
        '    Call PointDetailDesc()
        '    Call PointDetailSpecial()
        '    Call PointDetailwihtDraw()
        '    frmPoint01.btnChangeRW.Enabled = True
        '    frmPoint01.btnSearch.Enabled = False
        '    frmPoint01.btnFindMember.Enabled = False

        '    'Me.Close()
        '    ' เพิ่มข้อมูลในหน้าแลกของรางวัล
        'ElseIf iType = 2 Then
        '    ' For i = 0 To Me.LVmeberpointFind.Items.Count - 1
        '    'If Me.LVmeberpointFind.Items(i).Selected Then
        '    i = Me.LVmeberpointFind.SelectedItems(0).Index
        '    'iArCode = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        '    frmPoint01.txtdwMemberID.Text = Me.LVmeberpointFind.Items(i).SubItems(0).Text
        '    frmPoint01.txtdwARcode.Text = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        '    frmPoint01.txtdwArname.Text = Me.LVmeberpointFind.Items(i).SubItems(2).Text
        '    PointAM = Me.LVmeberpointFind.Items(i).SubItems(4).Text
        '    frmPoint01.txtDWtotalPoint.Text = Format(PointAM, "##,##0.00")
        '    PcampaignCode = Me.LVmeberpointFind.Items(i).SubItems(7).Text
        '    'frmPoint01.txtbgDate.Text = Me.LVmeberpointFind.Items(i).SubItems(5).Text
        '    'frmPoint01.txtExpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(6).Text
        '    'PcampaignCode = Me.LVmeberpointFind.Items(i).SubItems(7).Text
        '    'frmPoint01.txtCPName.Text = Me.LVmeberpointFind.Items(i).SubItems(8).Text
        '    'frmPoint01.txtCPstDate.Text = Me.LVmeberpointFind.Items(i).SubItems(9).Text
        '    'frmPoint01.txtCPstpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(10).Text

        '    ' End If
        '    ' Next
        '    '--------
        '    frmPoint01.btnselectReward.Enabled = True
        '    frmPoint01.pnChecKPoint.Hide()
        '    frmPoint01.PNsaveNewItem.Hide()
        '    frmPoint01.PNsaveSpecialPoint.Hide()
        '    frmPoint01.PNwithDraw.Show()
        '    frmPoint01.pnCampaignMaster.Hide()
        '    frmPoint01.lblCPCode.Text = PcampaignCode
        '    'Me.Close()

        'ElseIf iType = 3 Then
        '    'For i = 0 To Me.LVmeberpointFind.Items.Count - 1
        '    'If Me.LVmeberpointFind.Items(i).Selected Then
        '    'iArCode = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        '    i = Me.LVmeberpointFind.SelectedItems(0).Index
        '    frmPoint01.txtSPmemberid.Text = Me.LVmeberpointFind.Items(i).SubItems(0).Text
        '    frmPoint01.txtSPARCode.Text = Me.LVmeberpointFind.Items(i).SubItems(1).Text
        '    frmPoint01.txtSParName.Text = Me.LVmeberpointFind.Items(i).SubItems(2).Text
        '    PointAM = Me.LVmeberpointFind.Items(i).SubItems(4).Text
        '    '--------
        '    frmPoint01.pnChecKPoint.Hide()
        '    frmPoint01.PNsaveNewItem.Hide()
        '    frmPoint01.PNsaveSpecialPoint.Show()
        '    frmPoint01.PNwithDraw.Hide()
        '    frmPoint01.pnCampaignMaster.Hide()
        '    frmPoint01.lblConfirm.Text = "-N-"
        '    frmPoint01.lblConfirm.Show()
        '    frmPoint01.btnSaveCP.Enabled = True
        'End If
        ''Next
        ''End If
        'PsaveSPpoint = 1
        ' ''frmPoint01.Show()
        'frmPoint01.MdiParent = frmMainMember
        'frmPoint01.btnFindMember.Enabled = False

        ''frmPoint01.Show()
        'Me.Close()
        ' ''frmPoint01.lblCPCode.Text = PcampaignCode
    End Sub

    Private Sub btnCLSL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCLSL.Click
        Me.Close()
    End Sub

    Private Sub LVmeberpointFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LVmeberpointFind.Click
        Me.btnSLOK.Enabled = True
    End Sub
    Private Sub LoadMemberDetail()
        'Dim iArCode As String
        v = 1
        'frmPoint01 = frmPoint01


        If iType = 1 Then
            Dim vARCode As String
            i = Me.LVmeberpointFind.SelectedItems(0).Index
            'iArCode = Me.LVmeberpointFind.Items(i).SubItems(1).Text
            vARCode = Me.LVmeberpointFind.Items(i).SubItems(0).Text
            ' MsgBox(vARCode)
            frmPoint01.txtMemberid.Text = Me.LVmeberpointFind.Items(i).SubItems(0).Text
            frmPoint01.txtARcode.Text = Me.LVmeberpointFind.Items(i).SubItems(1).Text
            frmPoint01.txtMemberName.Text = Me.LVmeberpointFind.Items(i).SubItems(2).Text
            vTotalAmount = Me.LVmeberpointFind.Items(i).SubItems(3).Text
            frmPoint01.txtTotalAmount.Text = Format(vTotalAmount, "##,##0.00")
            PointAM = Me.LVmeberpointFind.Items(i).SubItems(4).Text
            frmPoint01.txtPointAmount.Text = Format(PointAM, "##,##0.00")
            frmPoint01.txtTotalPoint.Text = Format(PointAM, "##,##0.00")
            frmPoint01.txtbgDate.Text = Me.LVmeberpointFind.Items(i).SubItems(5).Text
            frmPoint01.txtExpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(6).Text
            PcampaignCode = Me.LVmeberpointFind.Items(i).SubItems(7).Text
            frmPoint01.txtCPName.Text = Me.LVmeberpointFind.Items(i).SubItems(8).Text
            frmPoint01.txtCPstDate.Text = Me.LVmeberpointFind.Items(i).SubItems(9).Text
            frmPoint01.txtCPstpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(10).Text

            '--------
            frmPoint01.pnChecKPoint.Show()
            frmPoint01.PNsaveNewItem.Hide()
            frmPoint01.PNsaveSpecialPoint.Hide()
            frmPoint01.PNwithDraw.Hide()
            frmPoint01.pnCampaignMaster.Hide()
            '----------
            Call frmP01Readonly()
            Call PointListforChange()
            Call PointDetailDesc()
            Call PointDetailSpecial()
            Call PointDetailwihtDraw()
            frmPoint01.btnChangeRW.Enabled = True
            frmPoint01.btnSearch.Enabled = False
            frmPoint01.btnFindMember.Enabled = False

            'Me.Close()
            ' เพิ่มข้อมูลในหน้าแลกของรางวัล
        ElseIf iType = 2 Then
            ' For i = 0 To Me.LVmeberpointFind.Items.Count - 1
            'If Me.LVmeberpointFind.Items(i).Selected Then
            i = Me.LVmeberpointFind.SelectedItems(0).Index
            'iArCode = Me.LVmeberpointFind.Items(i).SubItems(1).Text
            frmPoint01.txtdwMemberID.Text = Me.LVmeberpointFind.Items(i).SubItems(0).Text
            frmPoint01.txtdwARcode.Text = Me.LVmeberpointFind.Items(i).SubItems(1).Text
            frmPoint01.txtdwArname.Text = Me.LVmeberpointFind.Items(i).SubItems(2).Text
            PointAM = Me.LVmeberpointFind.Items(i).SubItems(4).Text
            frmPoint01.txtDWtotalPoint.Text = Format(PointAM, "##,##0.00")
            PcampaignCode = Me.LVmeberpointFind.Items(i).SubItems(7).Text
            'frmPoint01.txtbgDate.Text = Me.LVmeberpointFind.Items(i).SubItems(5).Text
            'frmPoint01.txtExpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(6).Text
            'PcampaignCode = Me.LVmeberpointFind.Items(i).SubItems(7).Text
            'frmPoint01.txtCPName.Text = Me.LVmeberpointFind.Items(i).SubItems(8).Text
            'frmPoint01.txtCPstDate.Text = Me.LVmeberpointFind.Items(i).SubItems(9).Text
            'frmPoint01.txtCPstpDate.Text = Me.LVmeberpointFind.Items(i).SubItems(10).Text

            ' End If
            ' Next
            '--------
            frmPoint01.btnselectReward.Enabled = True
            frmPoint01.pnChecKPoint.Hide()
            frmPoint01.PNsaveNewItem.Hide()
            frmPoint01.PNsaveSpecialPoint.Hide()
            frmPoint01.PNwithDraw.Show()
            frmPoint01.pnCampaignMaster.Hide()
            frmPoint01.lblCPCode.Text = PcampaignCode
            'Me.Close()

        ElseIf iType = 3 Then
            'For i = 0 To Me.LVmeberpointFind.Items.Count - 1
            'If Me.LVmeberpointFind.Items(i).Selected Then
            'iArCode = Me.LVmeberpointFind.Items(i).SubItems(1).Text
            i = Me.LVmeberpointFind.SelectedItems(0).Index
            frmPoint01.txtSPmemberid.Text = Me.LVmeberpointFind.Items(i).SubItems(0).Text
            frmPoint01.txtSPARCode.Text = Me.LVmeberpointFind.Items(i).SubItems(1).Text
            frmPoint01.txtSParName.Text = Me.LVmeberpointFind.Items(i).SubItems(2).Text
            PointAM = Me.LVmeberpointFind.Items(i).SubItems(4).Text
            '--------
            frmPoint01.pnChecKPoint.Hide()
            frmPoint01.PNsaveNewItem.Hide()
            frmPoint01.PNsaveSpecialPoint.Show()
            frmPoint01.PNwithDraw.Hide()
            frmPoint01.pnCampaignMaster.Hide()
            frmPoint01.lblConfirm.Text = "-N-"
            frmPoint01.lblConfirm.Show()
            frmPoint01.btnSaveCP.Enabled = True
        End If
        'Next
        'End If
        PsaveSPpoint = 1
        ''frmPoint01.Show()
        frmPoint01.MdiParent = FormMain
        frmPoint01.btnFindMember.Enabled = False

        'frmPoint01.Show()
        Me.Close()
        ''frmPoint01.lblCPCode.Text = PcampaignCode
    End Sub


    Private Sub LVmeberpointFind_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles LVmeberpointFind.ItemCheck
        Me.btnSLOK.Enabled = True
    End Sub

    Private Sub LVmeberpointFind_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LVmeberpointFind.MouseDoubleClick
        Call LoadMemberDetail()
    End Sub

    Private Sub LVmeberpointFind_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LVmeberpointFind.SelectedIndexChanged

    End Sub
End Class
