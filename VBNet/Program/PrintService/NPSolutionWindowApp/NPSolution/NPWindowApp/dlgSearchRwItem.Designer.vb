<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgSearchRwItem
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
        Me.LVrwItem = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtFrwitm = New System.Windows.Forms.TextBox
        Me.btnFrw = New System.Windows.Forms.Button
        Me.btnRWFN = New System.Windows.Forms.Button
        Me.btnExitRw = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'LVrwItem
        '
        Me.LVrwItem.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.LVrwItem.FullRowSelect = True
        Me.LVrwItem.GridLines = True
        Me.LVrwItem.Location = New System.Drawing.Point(12, 60)
        Me.LVrwItem.Name = "LVrwItem"
        Me.LVrwItem.Size = New System.Drawing.Size(427, 276)
        Me.LVrwItem.TabIndex = 1
        Me.LVrwItem.UseCompatibleStateImageBehavior = False
        Me.LVrwItem.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "รหัสสินค้า"
        Me.ColumnHeader1.Width = 113
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "ชื่อสินค้า"
        Me.ColumnHeader2.Width = 214
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "หน่วยนับ"
        Me.ColumnHeader3.Width = 94
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(15, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 20)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "คำค้นหา"
        '
        'txtFrwitm
        '
        Me.txtFrwitm.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtFrwitm.Location = New System.Drawing.Point(78, 16)
        Me.txtFrwitm.Name = "txtFrwitm"
        Me.txtFrwitm.Size = New System.Drawing.Size(275, 26)
        Me.txtFrwitm.TabIndex = 3
        '
        'btnFrw
        '
        Me.btnFrw.BackColor = System.Drawing.SystemColors.Control
        Me.btnFrw.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnFrw.Location = New System.Drawing.Point(360, 14)
        Me.btnFrw.Name = "btnFrw"
        Me.btnFrw.Size = New System.Drawing.Size(75, 28)
        Me.btnFrw.TabIndex = 4
        Me.btnFrw.Text = "ค้นหา"
        Me.btnFrw.UseVisualStyleBackColor = False
        '
        'btnRWFN
        '
        Me.btnRWFN.BackColor = System.Drawing.SystemColors.Control
        Me.btnRWFN.Location = New System.Drawing.Point(292, 342)
        Me.btnRWFN.Name = "btnRWFN"
        Me.btnRWFN.Size = New System.Drawing.Size(75, 39)
        Me.btnRWFN.TabIndex = 5
        Me.btnRWFN.Text = "ตกลง"
        Me.btnRWFN.UseVisualStyleBackColor = False
        '
        'btnExitRw
        '
        Me.btnExitRw.BackColor = System.Drawing.SystemColors.Control
        Me.btnExitRw.Location = New System.Drawing.Point(364, 342)
        Me.btnExitRw.Name = "btnExitRw"
        Me.btnExitRw.Size = New System.Drawing.Size(75, 39)
        Me.btnExitRw.TabIndex = 6
        Me.btnExitRw.Text = "ออก"
        Me.btnExitRw.UseVisualStyleBackColor = False
        '
        'dlgSearchRwItem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Desktop
        Me.ClientSize = New System.Drawing.Size(447, 383)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnExitRw)
        Me.Controls.Add(Me.btnRWFN)
        Me.Controls.Add(Me.btnFrw)
        Me.Controls.Add(Me.txtFrwitm)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LVrwItem)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgSearchRwItem"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ค้นหารายการของรางวัล"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LVrwItem As System.Windows.Forms.ListView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtFrwitm As System.Windows.Forms.TextBox
    Friend WithEvents btnFrw As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnRWFN As System.Windows.Forms.Button
    Friend WithEvents btnExitRw As System.Windows.Forms.Button

End Class
