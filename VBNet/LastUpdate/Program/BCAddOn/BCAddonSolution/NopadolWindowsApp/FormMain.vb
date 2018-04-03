Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.Globalization

Public Class FormMain
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

        Call ChekAuthorityAccess()

        If vDepartment = "MC" Or vDepartment = "IT" Or vDepartment = "AC" Then

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
        End If
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
       

    End Sub

    Private Sub MenuPriceStructureAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuPriceStructureAddItem.Click
        If frmItemSetPriceStructure Is Nothing Then
            frmItemSetPriceStructure = FormItemSetPriceStructure
        Else
            If frmItemSetPriceStructure.IsDisposed Then
                frmItemSetPriceStructure = FormItemSetPriceStructure
            End If
        End If
        frmItemSetPriceStructure.MdiParent = Me
        frmItemSetPriceStructure.Show()
        frmItemSetPriceStructure.BringToFront()
    End Sub
End Class
