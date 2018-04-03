<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormApproveCommission
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
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.ListViewReqComm = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.CBSelectAll = New System.Windows.Forms.CheckBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.BTNRefresh = New System.Windows.Forms.Button
        Me.BTNExit = New System.Windows.Forms.Button
        Me.BTNApprove = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkRed
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(996, 557)
        Me.Panel1.TabIndex = 0
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.ListViewReqComm)
        Me.Panel3.Controls.Add(Me.CBSelectAll)
        Me.Panel3.Location = New System.Drawing.Point(3, 3)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(990, 551)
        Me.Panel3.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.DarkRed
        Me.Label1.Location = New System.Drawing.Point(3, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(264, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "เลือก อนุมัติเอกสารขอเสนอสินค้าคิดค่าคอมมิชชั่น"
        '
        'ListViewReqComm
        '
        Me.ListViewReqComm.BackColor = System.Drawing.Color.LightGreen
        Me.ListViewReqComm.CheckBoxes = True
        Me.ListViewReqComm.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5})
        Me.ListViewReqComm.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ListViewReqComm.FullRowSelect = True
        Me.ListViewReqComm.GridLines = True
        Me.ListViewReqComm.Location = New System.Drawing.Point(3, 70)
        Me.ListViewReqComm.Name = "ListViewReqComm"
        Me.ListViewReqComm.Size = New System.Drawing.Size(985, 481)
        Me.ListViewReqComm.TabIndex = 1
        Me.ListViewReqComm.UseCompatibleStateImageBehavior = False
        Me.ListViewReqComm.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ลำดับ"
        Me.ColumnHeader1.Width = 45
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "เลขที่เอกสาร"
        Me.ColumnHeader2.Width = 150
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "วันที่เอกสาร"
        Me.ColumnHeader3.Width = 120
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "แคมเปญ"
        Me.ColumnHeader4.Width = 350
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "คำอธิบาย"
        Me.ColumnHeader5.Width = 315
        '
        'CBSelectAll
        '
        Me.CBSelectAll.AutoSize = True
        Me.CBSelectAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.CBSelectAll.Location = New System.Drawing.Point(3, 53)
        Me.CBSelectAll.Name = "CBSelectAll"
        Me.CBSelectAll.Size = New System.Drawing.Size(93, 20)
        Me.CBSelectAll.TabIndex = 0
        Me.CBSelectAll.Text = "เลือกทั้งหมด"
        Me.CBSelectAll.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Black
        Me.Panel2.Controls.Add(Me.Panel4)
        Me.Panel2.Location = New System.Drawing.Point(8, 584)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(996, 100)
        Me.Panel2.TabIndex = 1
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.BTNRefresh)
        Me.Panel4.Controls.Add(Me.BTNExit)
        Me.Panel4.Controls.Add(Me.BTNApprove)
        Me.Panel4.Location = New System.Drawing.Point(3, 3)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(990, 94)
        Me.Panel4.TabIndex = 0
        '
        'BTNRefresh
        '
        Me.BTNRefresh.BackColor = System.Drawing.Color.Gainsboro
        Me.BTNRefresh.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.BTNRefresh.Image = Global.NopadolWindowsApp.My.Resources.Resources.Windows_Explorer
        Me.BTNRefresh.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNRefresh.Location = New System.Drawing.Point(638, 17)
        Me.BTNRefresh.Name = "BTNRefresh"
        Me.BTNRefresh.Size = New System.Drawing.Size(102, 61)
        Me.BTNRefresh.TabIndex = 2
        Me.BTNRefresh.Text = "F1-ดูข้อมูล"
        Me.BTNRefresh.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNRefresh.UseVisualStyleBackColor = False
        '
        'BTNExit
        '
        Me.BTNExit.BackColor = System.Drawing.Color.Gainsboro
        Me.BTNExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.BTNExit.Image = Global.NopadolWindowsApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_73
        Me.BTNExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNExit.Location = New System.Drawing.Point(885, 17)
        Me.BTNExit.Name = "BTNExit"
        Me.BTNExit.Size = New System.Drawing.Size(102, 61)
        Me.BTNExit.TabIndex = 4
        Me.BTNExit.Text = "ESC-ออก"
        Me.BTNExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNExit.UseVisualStyleBackColor = False
        '
        'BTNApprove
        '
        Me.BTNApprove.BackColor = System.Drawing.Color.Gainsboro
        Me.BTNApprove.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.BTNApprove.Image = Global.NopadolWindowsApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_48
        Me.BTNApprove.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNApprove.Location = New System.Drawing.Point(761, 17)
        Me.BTNApprove.Name = "BTNApprove"
        Me.BTNApprove.Size = New System.Drawing.Size(102, 61)
        Me.BTNApprove.TabIndex = 3
        Me.BTNApprove.Text = "F5-อนุมัติ"
        Me.BTNApprove.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNApprove.UseVisualStyleBackColor = False
        '
        'FormApproveCommission
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1016, 734)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "FormApproveCommission"
        Me.Text = "อนุมัติ เอกสารขอเสนอสินค้าคิดค่าคอมฯ"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents ListViewReqComm As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BTNExit As System.Windows.Forms.Button
    Friend WithEvents BTNApprove As System.Windows.Forms.Button
    Friend WithEvents BTNRefresh As System.Windows.Forms.Button
    Friend WithEvents CBSelectAll As System.Windows.Forms.CheckBox
End Class
