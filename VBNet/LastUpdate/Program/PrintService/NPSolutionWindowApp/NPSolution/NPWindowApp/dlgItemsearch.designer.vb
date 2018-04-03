<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgItemsearch
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
        Me.lbl1 = New System.Windows.Forms.Label
        Me.txtSCHitm = New System.Windows.Forms.TextBox
        Me.btnSCHitm = New System.Windows.Forms.Button
        Me.LVSCHitm = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Location = New System.Drawing.Point(12, 18)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(52, 13)
        Me.lbl1.TabIndex = 1
        Me.lbl1.Text = "รหัสสินค้า"
        '
        'txtSCHitm
        '
        Me.txtSCHitm.Location = New System.Drawing.Point(70, 15)
        Me.txtSCHitm.Name = "txtSCHitm"
        Me.txtSCHitm.Size = New System.Drawing.Size(292, 20)
        Me.txtSCHitm.TabIndex = 2
        '
        'btnSCHitm
        '
        Me.btnSCHitm.Location = New System.Drawing.Point(368, 13)
        Me.btnSCHitm.Name = "btnSCHitm"
        Me.btnSCHitm.Size = New System.Drawing.Size(75, 23)
        Me.btnSCHitm.TabIndex = 3
        Me.btnSCHitm.Text = "ค้นหา"
        Me.btnSCHitm.UseVisualStyleBackColor = True
        '
        'LVSCHitm
        '
        Me.LVSCHitm.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4})
        Me.LVSCHitm.FullRowSelect = True
        Me.LVSCHitm.GridLines = True
        Me.LVSCHitm.Location = New System.Drawing.Point(12, 41)
        Me.LVSCHitm.Name = "LVSCHitm"
        Me.LVSCHitm.Size = New System.Drawing.Size(555, 395)
        Me.LVSCHitm.TabIndex = 4
        Me.LVSCHitm.UseCompatibleStateImageBehavior = False
        Me.LVSCHitm.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "รหัสสินค้า"
        Me.ColumnHeader1.Width = 117
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "ชื่อสินค้า"
        Me.ColumnHeader2.Width = 223
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "UnitCode"
        Me.ColumnHeader3.Width = 88
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "มูลค่า"
        Me.ColumnHeader4.Width = 115
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(418, 439)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 33)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "ตกลง"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(492, 439)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 33)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "ยกเลิก"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'dlgItemsearch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(582, 475)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.LVSCHitm)
        Me.Controls.Add(Me.btnSCHitm)
        Me.Controls.Add(Me.txtSCHitm)
        Me.Controls.Add(Me.lbl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgItemsearch"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = ":: ค้นหาสินค้า"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents txtSCHitm As System.Windows.Forms.TextBox
    Friend WithEvents btnSCHitm As System.Windows.Forms.Button
    Friend WithEvents LVSCHitm As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button

End Class
