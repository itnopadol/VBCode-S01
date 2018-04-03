<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormUpdateItemPriceDocNo
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
        Me.ListView101 = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.PGB101 = New System.Windows.Forms.ProgressBar
        Me.BTNUpdate = New System.Windows.Forms.Button
        Me.BTNCancel = New System.Windows.Forms.Button
        Me.BTNExit = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.Location = New System.Drawing.Point(63, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(468, 29)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "รายการ สินค้าทั้งหมดที่จะทำการปรับราคาวันนี้"
        '
        'ListView101
        '
        Me.ListView101.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8})
        Me.ListView101.FullRowSelect = True
        Me.ListView101.GridLines = True
        Me.ListView101.Location = New System.Drawing.Point(68, 63)
        Me.ListView101.Name = "ListView101"
        Me.ListView101.Size = New System.Drawing.Size(887, 422)
        Me.ListView101.TabIndex = 2
        Me.ListView101.UseCompatibleStateImageBehavior = False
        Me.ListView101.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "เลขที่เอกสาร"
        Me.ColumnHeader1.Width = 120
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "รหัสสินค้า"
        Me.ColumnHeader2.Width = 120
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "ระดับราคา"
        Me.ColumnHeader3.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "ประเภท"
        Me.ColumnHeader4.Width = 100
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "ราคาเก่า"
        Me.ColumnHeader5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader5.Width = 120
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "ราคาใหม่"
        Me.ColumnHeader6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader6.Width = 120
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "หน่วย"
        Me.ColumnHeader7.Width = 100
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "ชื่อสินค้า"
        Me.ColumnHeader8.Width = 300
        '
        'PGB101
        '
        Me.PGB101.Location = New System.Drawing.Point(68, 502)
        Me.PGB101.Name = "PGB101"
        Me.PGB101.Size = New System.Drawing.Size(887, 18)
        Me.PGB101.TabIndex = 3
        '
        'BTNUpdate
        '
        Me.BTNUpdate.Location = New System.Drawing.Point(68, 544)
        Me.BTNUpdate.Name = "BTNUpdate"
        Me.BTNUpdate.Size = New System.Drawing.Size(89, 34)
        Me.BTNUpdate.TabIndex = 4
        Me.BTNUpdate.Text = "ปรับราคา"
        Me.BTNUpdate.UseVisualStyleBackColor = True
        '
        'BTNCancel
        '
        Me.BTNCancel.Location = New System.Drawing.Point(176, 544)
        Me.BTNCancel.Name = "BTNCancel"
        Me.BTNCancel.Size = New System.Drawing.Size(89, 34)
        Me.BTNCancel.TabIndex = 5
        Me.BTNCancel.Text = "ยกเลิก"
        Me.BTNCancel.UseVisualStyleBackColor = True
        '
        'BTNExit
        '
        Me.BTNExit.Location = New System.Drawing.Point(284, 544)
        Me.BTNExit.Name = "BTNExit"
        Me.BTNExit.Size = New System.Drawing.Size(89, 34)
        Me.BTNExit.TabIndex = 6
        Me.BTNExit.Text = "ออก"
        Me.BTNExit.UseVisualStyleBackColor = True
        '
        'FormUpdateItemPriceDocNo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1016, 734)
        Me.Controls.Add(Me.BTNExit)
        Me.Controls.Add(Me.BTNCancel)
        Me.Controls.Add(Me.BTNUpdate)
        Me.Controls.Add(Me.PGB101)
        Me.Controls.Add(Me.ListView101)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FormUpdateItemPriceDocNo"
        Me.Text = "FormUpdateItemPriceDocNo"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ListView101 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents PGB101 As System.Windows.Forms.ProgressBar
    Friend WithEvents BTNUpdate As System.Windows.Forms.Button
    Friend WithEvents BTNCancel As System.Windows.Forms.Button
    Friend WithEvents BTNExit As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
End Class
