Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.Globalization

Public Class FormMain
    'Dim frmBudgetConfig As New FormBudgetConfig
    ''Dim frmUpdateItemPrice As New FormUpdateItemPrice
    'Dim frmBudgetTargetDepartmentConfig As FormBudgetTargetDepartmentConfig
    ''Dim frmUpdateItemPriceDocNo As FormUpdateItemPriceDocNo
    'Dim frmTeamIncentiveConfig As FormTeamIncentiveConfig
    'Dim frmSaleHistory As FormSaleHistory
    'Dim frmCalIncentive As FormCalIncentive
    'Dim frmPayIncentive As FormPayIncentive
    ''Dim frmPriceStructure As FormPriceStructure
    'Dim FrmImportDataPriceStructureFromExcel As FormImportDataPriceStructureFromExcel
    'Dim FrmCouponRecord As FormCouponRecord
    'Dim FrmCouponRequest As FormCouponRequest

    ''Dim FormApproveVolumeSet As FormApproveVolumeSet
    'Dim dlgVolumeSearch As dlgVolumeSearch
    'Dim dlgPSVdocSearch As dlgPSVdocSearch
    '' Public frmPriceVolumeSet As frmPriceVolumeSet

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim vQuery As String
        'Dim vListItem As New ListViewItem
        'Dim i As Integer

        'vConnectionString = "Persist Security Info = False;User ID=vbuser;Password=132;Data Source = Nebula;Initial Catalog = BCNP"
        'vConnection = New SqlConnection(vConnectionString)
        'vConnection.Open()
        'Call InitializeDataBase()

        'vQuery = "exec dbo.USP_WEB_SearchQueueDocno1"
        'da = New SqlDataAdapter(vQuery, vConnection)
        'ds = New DataSet
        'da.Fill(ds, "Queue")
        'dt = ds.Tables("Queue")
        'ListView1.Items.Clear()
        'For i = 0 To dt.Rows.Count - 1
        'vListItem = ListView1.Items.Add(dt.Rows(i).Item("docno"))
        'vListItem.SubItems.Add(1).Text = Trim(dt.Rows(i).Item("picker"))
        'Next

        Dim vFrm As New FormLogIn
        vFrm.MdiParent = Me
        vFrm.Show()
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

    End Sub

    Private Sub ÕÕ°‚ª√·°√¡ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ÕÕ°‚ª√·°√¡ToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub BudgetConfigToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BudgetConfigToolStripMenuItem.Click
        If frmBudgetConfig Is Nothing Then
            frmBudgetConfig = New FormBudgetConfig
        Else
            If frmBudgetConfig.IsDisposed Then
                frmBudgetConfig = New FormBudgetConfig
            End If
        End If
        frmBudgetConfig.MdiParent = Me
        frmBudgetConfig.Show()
        frmBudgetConfig.BringToFront()
    End Sub

    Private Sub CreateChangeItemPrice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If frmUpdateItemPrice Is Nothing Then
        '    frmUpdateItemPrice = New FormUpdateItemPrice
        'Else
        '    If frmUpdateItemPrice.IsDisposed Then
        '        frmUpdateItemPrice = New FormUpdateItemPrice
        '    End If
        'End If
        'frmUpdateItemPrice.MdiParent = Me
        'frmUpdateItemPrice.Show()
        'frmUpdateItemPrice.BringToFront()
    End Sub

    Private Sub BudgetTargetDepartmentConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BudgetTargetDepartmentConfig.Click
        If frmBudgetTargetDepartmentConfig Is Nothing Then
            frmBudgetTargetDepartmentConfig = New FormBudgetTargetDepartmentConfig
        Else
            If frmBudgetTargetDepartmentConfig.IsDisposed Then
                frmBudgetTargetDepartmentConfig = New FormBudgetTargetDepartmentConfig
            End If
        End If
        frmBudgetTargetDepartmentConfig.MdiParent = Me
        frmBudgetTargetDepartmentConfig.Show()
        frmBudgetTargetDepartmentConfig.BringToFront()
    End Sub

    Private Sub UpdateItemPrice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If frmUpdateItemPriceDocNo Is Nothing Then
        '    frmUpdateItemPriceDocNo = New FormUpdateItemPriceDocNo
        'Else
        '    If frmUpdateItemPriceDocNo.IsDisposed Then
        '        frmUpdateItemPriceDocNo = New FormUpdateItemPriceDocNo
        '    End If
        'End If
        'frmUpdateItemPriceDocNo.MdiParent = Me
        'frmUpdateItemPriceDocNo.Show()
        'frmUpdateItemPriceDocNo.BringToFront()

    End Sub

    Private Sub TeamIncentiveConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TeamIncentiveConfig.Click
        If frmTeamIncentiveConfig Is Nothing Then
            frmTeamIncentiveConfig = New FormTeamIncentiveConfig
        Else
            If frmTeamIncentiveConfig.IsDisposed Then
                frmTeamIncentiveConfig = New FormTeamIncentiveConfig
            End If
        End If
        frmTeamIncentiveConfig.MdiParent = Me
        frmTeamIncentiveConfig.Show()
        frmTeamIncentiveConfig.BringToFront()
    End Sub

    Private Sub SaleHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaleHistory.Click
        If frmSaleHistory Is Nothing Then
            frmSaleHistory = New FormSaleHistory
        Else
            If frmSaleHistory.IsDisposed Then
                frmSaleHistory = New FormSaleHistory
            End If
        End If
        frmSaleHistory.MdiParent = Me
        frmSaleHistory.Show()
        frmSaleHistory.BringToFront()
    End Sub

    Private Sub CalcIncentive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CalcIncentive.Click
        If frmCalIncentive Is Nothing Then
            frmCalIncentive = New FormCalIncentive
        Else
            If frmCalIncentive.IsDisposed Then
                frmCalIncentive = New FormCalIncentive
            End If
        End If
        frmCalIncentive.MdiParent = Me
        frmCalIncentive.Show()
        frmCalIncentive.BringToFront()
    End Sub

    Private Sub PayIncentive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PayIncentive.Click
        If frmPayIncentive Is Nothing Then
            frmPayIncentive = New FormPayIncentive
        Else
            If frmPayIncentive.IsDisposed Then
                frmPayIncentive = New FormPayIncentive
            End If
        End If
        frmPayIncentive.MdiParent = Me
        frmPayIncentive.Show()
        frmPayIncentive.BringToFront()
    End Sub

    Private Sub PriceStructure_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox("¬°‡≈‘°°“√„™Èß“π", MsgBoxStyle.Information, "Send Message Information")
        Exit Sub
        'If Me.frmPriceStructure Is Nothing Then
        '    frmPriceStructure = FormPriceStructure
        'Else
        '    If frmPriceStructure.IsDisposed Then
        '        frmPriceStructure = FormPriceStructure
        '    End If
        'End If
        'frmPriceStructure.MdiParent = Me
        'frmPriceStructure.Show()
        'frmPriceStructure.BringToFront()
    End Sub

    'Private Sub ∑¥ Õ∫_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If Me.frm1 Is Nothing Then
    '        frm1 = Form1
    '    Else
    '        If frm1.IsDisposed Then
    '            frm1 = Form1
    '        End If
    '    End If
    '    frm1.MdiParent = Me
    '    frm1.Show()
    '    frm1.BringToFront()
    'End Sub


    Private Sub ExcelExportData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExcelExportData.Click
        If FrmImportDataPriceStructureFromExcel Is Nothing Then
            FrmImportDataPriceStructureFromExcel = FormImportDataPriceStructureFromExcel
        Else
            If FrmImportDataPriceStructureFromExcel.IsDisposed Then
                FrmImportDataPriceStructureFromExcel = FormImportDataPriceStructureFromExcel
            End If
        End If
        FrmImportDataPriceStructureFromExcel.MdiParent = Me
        FrmImportDataPriceStructureFromExcel.Show()
        FrmImportDataPriceStructureFromExcel.BringToFront()
    End Sub

    Private Sub CouponRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CouponRecord.Click
        If FrmCouponRecord Is Nothing Then
            FrmCouponRecord = FormCouponRecord
        Else
            If FrmCouponRecord.IsDisposed Then
                FrmCouponRecord = FormCouponRecord
            End If
        End If
        FormCouponRecord.MdiParent = Me
        FormCouponRecord.Show()
        FormCouponRecord.BringToFront()
    End Sub

    Private Sub CouponRequest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CouponRequest.Click
        If FrmCouponRequest Is Nothing Then
            FrmCouponRequest = FormCouponRequest
        Else
            If FrmCouponRequest.IsDisposed Then
                FrmCouponRequest = FormCouponRequest
            End If
        End If
        FormCouponRequest.MdiParent = Me
        FormCouponRequest.Show()
        FormCouponRequest.BringToFront()
    End Sub

   
    Private Sub mnPriceVolumeSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnPriceVolumeSet.Click

        If frmPriceVolumeSet Is Nothing Then
            frmPriceVolumeSet = frmPriceVolumeSet
        Else
            If frmPriceVolumeSet.IsDisposed Then
                frmPriceVolumeSet = frmPriceVolumeSet
            End If
        End If
        frmPriceVolumeSet.MdiParent = Me
        frmPriceVolumeSet.Show()
        frmPriceVolumeSet.BringToFront()
       
    End Sub

    Private Sub mnApprovePriceVolumeSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnApprovePriceVolumeSet.Click
        If (vUserID = "yuraporn") Or (vUserID = "sawimon") Or (vUserID = "kittima") Or (vUserID = "panuvich") Or (vUserID = "komkrithc") Or (vUserID = "nueng") Then
            If FormApproveVolumeSet Is Nothing Then
                FormApproveVolumeSet = FormApproveVolumeSet
            Else
                If FormApproveVolumeSet.IsDisposed Then
                    FormApproveVolumeSet = FormApproveVolumeSet
                End If
            End If
            FormApproveVolumeSet.MdiParent = Me
            FormApproveVolumeSet.Show()
            FormApproveVolumeSet.BringToFront()
        Else
            MsgBox("§ÿ≥‰¡Ë¡’ ‘∑∏‘Ï‡¢È“√“¬°“√π’È", MsgBoxStyle.Information, "Information")
        End If
    End Sub

    Private Sub MenuCampaign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCampaign.Click
        If FrmSetCommCampaign Is Nothing Then
            FrmSetCommCampaign = FormSetCommCampaign
        Else
            If FrmSetCommCampaign.IsDisposed Then
                FrmSetCommCampaign = FormSetCommCampaign
            End If
        End If
        FrmSetCommCampaign.MdiParent = Me
        FrmSetCommCampaign.Show()
        FrmSetCommCampaign.BringToFront()
    End Sub

    Private Sub MenuRequestCommission_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuRequestCommission.Click
        If FrmReqCommission Is Nothing Then
            FrmReqCommission = FormReqCommission
        Else
            If FrmReqCommission.IsDisposed Then
                FrmReqCommission = FormReqCommission
            End If
        End If
        FrmReqCommission.MdiParent = Me
        FrmReqCommission.Show()
        FrmReqCommission.BringToFront()
    End Sub

    Private Sub MenuApproveCommission_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuApproveCommission.Click
        If FrmApproveCommission Is Nothing Then
            FrmApproveCommission = FormApproveCommission
        Else
            If FrmApproveCommission.IsDisposed Then
                FrmApproveCommission = FormApproveCommission
            End If
        End If
        FrmApproveCommission.MdiParent = Me
        FrmApproveCommission.Show()
        FrmApproveCommission.BringToFront()
    End Sub

    Private Sub MenuPayCommission_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPayCommission.Click
        If FrmPayCommission Is Nothing Then
            FrmPayCommission = FormPayCommission
        Else
            If FrmPayCommission.IsDisposed Then
                FrmPayCommission = FormPayCommission
            End If
        End If
        FrmPayCommission.MdiParent = Me
        FrmPayCommission.Show()
        FrmPayCommission.BringToFront()
    End Sub

    Private Sub MenuPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If vFrmPoint01 Is Nothing Then
            vFrmPoint01 = frmPoint01
        Else
            If vFrmPoint01.IsDisposed Then
                vFrmPoint01 = frmPoint01
            End If
        End If
        vFrmPoint01.MdiParent = Me
        vFrmPoint01.Show()
        vFrmPoint01.BringToFront()
    End Sub

    Private Sub MenuSmartPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuSmartPoint.Click
        If frmSmartPoint Is Nothing Then
            frmSmartPoint = FormSmartPoint
        Else
            If frmSmartPoint.IsDisposed Then
                frmSmartPoint = FormSmartPoint
            End If
        End If
        frmSmartPoint.MdiParent = Me
        frmSmartPoint.Show()
        frmSmartPoint.BringToFront()
    End Sub

    Private Sub MenuCEData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCEData.Click
        If frmCEProgram Is Nothing Then
            frmCEProgram = FormCEProgram
        Else
            If frmCEProgram.IsDisposed Then
                frmCEProgram = FormCEProgram
            End If
        End If
        frmCEProgram.MdiParent = Me
        frmCEProgram.Show()
        frmCEProgram.BringToFront()
    End Sub

    Private Sub MenuAddPayCoupon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuAddPayCoupon.Click
        If frmCouponExpertFair Is Nothing Then
            frmCouponExpertFair = FormCouponExpertFair
        Else
            If frmCouponExpertFair.IsDisposed Then
                frmCouponExpertFair = FormCouponExpertFair
            End If
        End If
        frmCouponExpertFair.MdiParent = Me
        frmCouponExpertFair.Show()
        frmCouponExpertFair.BringToFront()
    End Sub
End Class
