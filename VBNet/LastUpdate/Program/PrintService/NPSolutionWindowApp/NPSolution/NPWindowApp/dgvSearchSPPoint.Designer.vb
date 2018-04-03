<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dgvSearchSPPoint
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtFspDocNo = New System.Windows.Forms.TextBox
        Me.btnFsp = New System.Windows.Forms.Button
        Me.LVfSpPoint = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.btnExitSP = New System.Windows.Forms.Button
        Me.btnSPfind = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(20, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "เลขที่เอกสาร"
        '
        'txtFspDocNo
        '
        Me.txtFspDocNo.Location = New System.Drawing.Point(95, 15)
        Me.txtFspDocNo.Name = "txtFspDocNo"
        Me.txtFspDocNo.Size = New System.Drawing.Size(181, 20)
        Me.txtFspDocNo.TabIndex = 2
        '
        'btnFsp
        '
        Me.btnFsp.BackColor = System.Drawing.SystemColors.Control
        Me.btnFsp.Location = New System.Drawing.Point(271, 14)
        Me.btnFsp.Name = "btnFsp"
        Me.btnFsp.Size = New System.Drawing.Size(75, 23)
        Me.btnFsp.TabIndex = 3
        Me.btnFsp.Text = "ค้นหา"
        Me.btnFsp.UseVisualStyleBackColor = False
        '
        'LVfSpPoint
        '
        Me.LVfSpPoint.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8, Me.ColumnHeader9, Me.ColumnHeader10})
        Me.LVfSpPoint.FullRowSelect = True
        Me.LVfSpPoint.GridLines = True
        Me.LVfSpPoint.Location = New System.Drawing.Point(1, 41)
        Me.LVfSpPoint.Name = "LVfSpPoint"
        Me.LVfSpPoint.Size = New System.Drawing.Size(689, 173)
        Me.LVfSpPoint.TabIndex = 4
        Me.LVfSpPoint.UseCompatibleStateImageBehavior = False
        Me.LVfSpPoint.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "เลขที่เอกสาร"
        Me.ColumnHeader1.Width = 82
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "วันที่"
        Me.ColumnHeader2.Width = 76
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "รหัสลูกค้า"
        Me.ColumnHeader3.Width = 85
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "ชื่อลูกค้า"
        Me.ColumnHeader4.Width = 158
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "รหัสสมาชิก"
        Me.ColumnHeader5.Width = 91
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "จำนวนแต้ม"
        Me.ColumnHeader6.Width = 81
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "CampaignCode"
        Me.ColumnHeader7.Width = 106
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Reason"
        Me.ColumnHeader8.Width = 0
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "confirm"
        Me.ColumnHeader9.Width = 0
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "cancel"
        Me.ColumnHeader10.Width = 2
        '
        'btnExitSP
        '
        Me.btnExitSP.BackColor = System.Drawing.SystemColors.Control
        Me.btnExitSP.Location = New System.Drawing.Point(615, 220)
        Me.btnExitSP.Name = "btnExitSP"
        Me.btnExitSP.Size = New System.Drawing.Size(75, 34)
        Me.btnExitSP.TabIndex = 6
        Me.btnExitSP.Text = "ออก"
        Me.btnExitSP.UseVisualStyleBackColor = False
        '
        'btnSPfind
        '
        Me.btnSPfind.BackColor = System.Drawing.SystemColors.Control
        Me.btnSPfind.Location = New System.Drawing.Point(539, 220)
        Me.btnSPfind.Name = "btnSPfind"
        Me.btnSPfind.Size = New System.Drawing.Size(75, 34)
        Me.btnSPfind.TabIndex = 5
        Me.btnSPfind.Text = "ตกลง"
        Me.btnSPfind.UseVisualStyleBackColor = False
        '
        'dgvSearchSPPoint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Desktop
        Me.ClientSize = New System.Drawing.Size(692, 266)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnExitSP)
        Me.Controls.Add(Me.btnSPfind)
        Me.Controls.Add(Me.LVfSpPoint)
        Me.Controls.Add(Me.btnFsp)
        Me.Controls.Add(Me.txtFspDocNo)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dgvSearchSPPoint"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ค้นหารายการแต้มพิเศษ"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtFspDocNo As System.Windows.Forms.TextBox
    Friend WithEvents btnFsp As System.Windows.Forms.Button
    Friend WithEvents LVfSpPoint As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnSPfind As System.Windows.Forms.Button
    Friend WithEvents btnExitSP As System.Windows.Forms.Button

End Class
