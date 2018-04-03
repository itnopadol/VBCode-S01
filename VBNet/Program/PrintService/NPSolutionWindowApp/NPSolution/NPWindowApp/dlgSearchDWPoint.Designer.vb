<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgSearchDWPoint
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
        Me.LVdwDoc = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtDwDocFND = New System.Windows.Forms.TextBox
        Me.btnFindDwdoc = New System.Windows.Forms.Button
        Me.btnSelect = New System.Windows.Forms.Button
        Me.btnExitdw = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'LVdwDoc
        '
        Me.LVdwDoc.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8})
        Me.LVdwDoc.FullRowSelect = True
        Me.LVdwDoc.GridLines = True
        Me.LVdwDoc.Location = New System.Drawing.Point(24, 59)
        Me.LVdwDoc.Name = "LVdwDoc"
        Me.LVdwDoc.Size = New System.Drawing.Size(763, 281)
        Me.LVdwDoc.TabIndex = 1
        Me.LVdwDoc.UseCompatibleStateImageBehavior = False
        Me.LVdwDoc.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "เลขที่เอกสาร"
        Me.ColumnHeader1.Width = 89
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "วันที่เอกสาร"
        Me.ColumnHeader2.Width = 78
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "รหัสลูกค้า"
        Me.ColumnHeader4.Width = 87
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "รหัสสมาชิก"
        Me.ColumnHeader5.Width = 103
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "ชื่อลูกค้า"
        Me.ColumnHeader6.Width = 196
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "จำนวนแต้ม"
        Me.ColumnHeader7.Width = 87
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "มูลค่า(บาท)"
        Me.ColumnHeader8.Width = 111
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.Location = New System.Drawing.Point(58, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "ใส่คำค้นหา"
        '
        'txtDwDocFND
        '
        Me.txtDwDocFND.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtDwDocFND.Location = New System.Drawing.Point(124, 27)
        Me.txtDwDocFND.Name = "txtDwDocFND"
        Me.txtDwDocFND.Size = New System.Drawing.Size(532, 22)
        Me.txtDwDocFND.TabIndex = 3
        '
        'btnFindDwdoc
        '
        Me.btnFindDwdoc.BackColor = System.Drawing.SystemColors.Control
        Me.btnFindDwdoc.ForeColor = System.Drawing.Color.Black
        Me.btnFindDwdoc.Location = New System.Drawing.Point(659, 24)
        Me.btnFindDwdoc.Name = "btnFindDwdoc"
        Me.btnFindDwdoc.Size = New System.Drawing.Size(75, 28)
        Me.btnFindDwdoc.TabIndex = 4
        Me.btnFindDwdoc.Text = "ค้นหา"
        Me.btnFindDwdoc.UseVisualStyleBackColor = False
        '
        'btnSelect
        '
        Me.btnSelect.BackColor = System.Drawing.SystemColors.Control
        Me.btnSelect.ForeColor = System.Drawing.Color.Black
        Me.btnSelect.Location = New System.Drawing.Point(637, 346)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(75, 29)
        Me.btnSelect.TabIndex = 5
        Me.btnSelect.Text = "เลือก"
        Me.btnSelect.UseVisualStyleBackColor = False
        '
        'btnExitdw
        '
        Me.btnExitdw.BackColor = System.Drawing.SystemColors.Control
        Me.btnExitdw.ForeColor = System.Drawing.Color.Black
        Me.btnExitdw.Location = New System.Drawing.Point(712, 346)
        Me.btnExitdw.Name = "btnExitdw"
        Me.btnExitdw.Size = New System.Drawing.Size(75, 29)
        Me.btnExitdw.TabIndex = 6
        Me.btnExitdw.Text = "ออก"
        Me.btnExitdw.UseVisualStyleBackColor = False
        '
        'dlgSearchDWPoint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Desktop
        Me.ClientSize = New System.Drawing.Size(802, 387)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnExitdw)
        Me.Controls.Add(Me.btnSelect)
        Me.Controls.Add(Me.btnFindDwdoc)
        Me.Controls.Add(Me.txtDwDocFND)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LVdwDoc)
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgSearchDWPoint"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "dlgSearchDWPoint"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LVdwDoc As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtDwDocFND As System.Windows.Forms.TextBox
    Friend WithEvents btnFindDwdoc As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents btnExitdw As System.Windows.Forms.Button

End Class
