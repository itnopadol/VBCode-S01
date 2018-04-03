<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormApproveVolumeSet
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnRefresh = New System.Windows.Forms.Button
        Me.P03 = New System.Windows.Forms.Panel
        Me.chkAll = New System.Windows.Forms.CheckBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnApprove = New System.Windows.Forms.Button
        Me.ListQue = New System.Windows.Forms.ListView
        Me.ColumnHeader13 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader14 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader15 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader16 = New System.Windows.Forms.ColumnHeader
        Me.appRequest = New System.Windows.Forms.ColumnHeader
        Me.P03.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnRefresh
        '
        Me.btnRefresh.Image = Global.NopadolWindowsApp.My.Resources.Resources.icon_16_checkin
        Me.btnRefresh.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnRefresh.Location = New System.Drawing.Point(674, 524)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(80, 54)
        Me.btnRefresh.TabIndex = 9
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'P03
        '
        Me.P03.BackColor = System.Drawing.SystemColors.ControlDark
        Me.P03.Controls.Add(Me.chkAll)
        Me.P03.Controls.Add(Me.btnRefresh)
        Me.P03.Controls.Add(Me.Button1)
        Me.P03.Controls.Add(Me.btnApprove)
        Me.P03.Controls.Add(Me.ListQue)
        Me.P03.Location = New System.Drawing.Point(39, 4)
        Me.P03.Name = "P03"
        Me.P03.Size = New System.Drawing.Size(943, 586)
        Me.P03.TabIndex = 12
        '
        'chkAll
        '
        Me.chkAll.AutoSize = True
        Me.chkAll.ForeColor = System.Drawing.Color.Blue
        Me.chkAll.Location = New System.Drawing.Point(5, 511)
        Me.chkAll.Name = "chkAll"
        Me.chkAll.Size = New System.Drawing.Size(83, 17)
        Me.chkAll.TabIndex = 10
        Me.chkAll.Text = "เลือกทั้งหมด"
        Me.chkAll.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Image = Global.NopadolWindowsApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_73
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button1.Location = New System.Drawing.Point(848, 524)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(82, 53)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "ออก"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnApprove
        '
        Me.btnApprove.Image = Global.NopadolWindowsApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_76
        Me.btnApprove.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnApprove.Location = New System.Drawing.Point(760, 524)
        Me.btnApprove.Name = "btnApprove"
        Me.btnApprove.Size = New System.Drawing.Size(82, 54)
        Me.btnApprove.TabIndex = 7
        Me.btnApprove.Text = "อนุมัติ"
        Me.btnApprove.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnApprove.UseVisualStyleBackColor = True
        '
        'ListQue
        '
        Me.ListQue.CheckBoxes = True
        Me.ListQue.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader13, Me.ColumnHeader14, Me.ColumnHeader15, Me.ColumnHeader16, Me.appRequest})
        Me.ListQue.FullRowSelect = True
        Me.ListQue.GridLines = True
        Me.ListQue.Location = New System.Drawing.Point(6, 6)
        Me.ListQue.Name = "ListQue"
        Me.ListQue.Size = New System.Drawing.Size(934, 502)
        Me.ListQue.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.ListQue.TabIndex = 6
        Me.ListQue.UseCompatibleStateImageBehavior = False
        Me.ListQue.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = "เลขที่เอกสาร"
        Me.ColumnHeader13.Width = 167
        '
        'ColumnHeader14
        '
        Me.ColumnHeader14.Text = "วันที่ปรับราคา"
        Me.ColumnHeader14.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader14.Width = 164
        '
        'ColumnHeader15
        '
        Me.ColumnHeader15.Text = "GPเฉลี่ยทุนตลาด"
        Me.ColumnHeader15.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader15.Width = 151
        '
        'ColumnHeader16
        '
        Me.ColumnHeader16.Text = "GPเฉลี่ยทุนตลาด lot"
        Me.ColumnHeader16.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader16.Width = 170
        '
        'appRequest
        '
        Me.appRequest.Text = "ผู้ขออนุมัติ"
        Me.appRequest.Width = 270
        '
        'FormApproveVolumeSet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Desktop
        Me.ClientSize = New System.Drawing.Size(1004, 601)
        Me.Controls.Add(Me.P03)
        Me.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Name = "FormApproveVolumeSet"
        Me.Text = "FormApproveVolumeSet"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.P03.ResumeLayout(False)
        Me.P03.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnApprove As System.Windows.Forms.Button
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents P03 As System.Windows.Forms.Panel
    Friend WithEvents ListQue As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader13 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader14 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader15 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader16 As System.Windows.Forms.ColumnHeader
    Friend WithEvents chkAll As System.Windows.Forms.CheckBox
    Friend WithEvents appRequest As System.Windows.Forms.ColumnHeader
End Class
