<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormCalIncentive
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CMBPeriod = New System.Windows.Forms.ComboBox
        Me.CMBFiscalYear = New System.Windows.Forms.ComboBox
        Me.CMBSaleType = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ListView101 = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.BTNSave = New System.Windows.Forms.Button
        Me.CBPeriod = New System.Windows.Forms.CheckBox
        Me.BTNClose = New System.Windows.Forms.Button
        Me.PGBar101 = New System.Windows.Forms.ProgressBar
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CMBPeriod)
        Me.GroupBox1.Controls.Add(Me.CMBFiscalYear)
        Me.GroupBox1.Controls.Add(Me.CMBSaleType)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(33, 73)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(946, 75)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "รายละเอียดการดูข้อมูล"
        '
        'CMBPeriod
        '
        Me.CMBPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMBPeriod.FormattingEnabled = True
        Me.CMBPeriod.Location = New System.Drawing.Point(717, 36)
        Me.CMBPeriod.Name = "CMBPeriod"
        Me.CMBPeriod.Size = New System.Drawing.Size(134, 21)
        Me.CMBPeriod.TabIndex = 5
        '
        'CMBFiscalYear
        '
        Me.CMBFiscalYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMBFiscalYear.FormattingEnabled = True
        Me.CMBFiscalYear.Location = New System.Drawing.Point(439, 36)
        Me.CMBFiscalYear.Name = "CMBFiscalYear"
        Me.CMBFiscalYear.Size = New System.Drawing.Size(134, 21)
        Me.CMBFiscalYear.TabIndex = 4
        '
        'CMBSaleType
        '
        Me.CMBSaleType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMBSaleType.FormattingEnabled = True
        Me.CMBSaleType.Items.AddRange(New Object() {"ขายเงินเชื่อ", "ขายเงินสด"})
        Me.CMBSaleType.Location = New System.Drawing.Point(203, 36)
        Me.CMBSaleType.Name = "CMBSaleType"
        Me.CMBSaleType.Size = New System.Drawing.Size(134, 21)
        Me.CMBSaleType.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label4.Location = New System.Drawing.Point(596, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(116, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Period Of 4 Week :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.Location = New System.Drawing.Point(357, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Fiscal Year :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label2.Location = New System.Drawing.Point(99, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "ประเภทการขาย :"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.Location = New System.Drawing.Point(28, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(308, 36)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "คำนวณ Incentive "
        '
        'ListView101
        '
        Me.ListView101.CheckBoxes = True
        Me.ListView101.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4})
        Me.ListView101.FullRowSelect = True
        Me.ListView101.GridLines = True
        Me.ListView101.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.ListView101.Location = New System.Drawing.Point(176, 211)
        Me.ListView101.Name = "ListView101"
        Me.ListView101.Size = New System.Drawing.Size(679, 346)
        Me.ListView101.TabIndex = 5
        Me.ListView101.UseCompatibleStateImageBehavior = False
        Me.ListView101.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Code"
        Me.ColumnHeader1.Width = 120
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Department"
        Me.ColumnHeader2.Width = 300
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "IsProcess"
        Me.ColumnHeader3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader3.Width = 120
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "วันเวลาคำนวณล่าสุด"
        Me.ColumnHeader4.Width = 120
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label6.Location = New System.Drawing.Point(713, 189)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(142, 22)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "สถานะเก็บประวัติ"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label5.Location = New System.Drawing.Point(296, 189)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(302, 22)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Department"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label7.Location = New System.Drawing.Point(176, 189)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(122, 22)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Code"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label8.Location = New System.Drawing.Point(597, 189)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(122, 22)
        Me.Label8.TabIndex = 10
        Me.Label8.Text = "IsProcess"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BTNSave
        '
        Me.BTNSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.BTNSave.Location = New System.Drawing.Point(761, 597)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(94, 29)
        Me.BTNSave.TabIndex = 12
        Me.BTNSave.Text = "คำนวณ"
        Me.BTNSave.UseVisualStyleBackColor = True
        '
        'CBPeriod
        '
        Me.CBPeriod.AutoSize = True
        Me.CBPeriod.Location = New System.Drawing.Point(176, 577)
        Me.CBPeriod.Name = "CBPeriod"
        Me.CBPeriod.Size = New System.Drawing.Size(128, 17)
        Me.CBPeriod.TabIndex = 11
        Me.CBPeriod.Text = "Select Department All"
        Me.CBPeriod.UseVisualStyleBackColor = True
        '
        'BTNClose
        '
        Me.BTNClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.BTNClose.Location = New System.Drawing.Point(871, 597)
        Me.BTNClose.Name = "BTNClose"
        Me.BTNClose.Size = New System.Drawing.Size(94, 29)
        Me.BTNClose.TabIndex = 13
        Me.BTNClose.Text = "ออก"
        Me.BTNClose.UseVisualStyleBackColor = True
        '
        'PGBar101
        '
        Me.PGBar101.Location = New System.Drawing.Point(176, 164)
        Me.PGBar101.Name = "PGBar101"
        Me.PGBar101.Size = New System.Drawing.Size(679, 17)
        Me.PGBar101.TabIndex = 14
        '
        'FormCalIncentive
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDark
        Me.ClientSize = New System.Drawing.Size(1016, 734)
        Me.Controls.Add(Me.PGBar101)
        Me.Controls.Add(Me.BTNClose)
        Me.Controls.Add(Me.BTNSave)
        Me.Controls.Add(Me.CBPeriod)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.ListView101)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "FormCalIncentive"
        Me.Text = "FormCalIncentive"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CMBPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents CMBFiscalYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ListView101 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents CBPeriod As System.Windows.Forms.CheckBox
    Friend WithEvents BTNClose As System.Windows.Forms.Button
    Friend WithEvents PGBar101 As System.Windows.Forms.ProgressBar
    Friend WithEvents CMBSaleType As System.Windows.Forms.ComboBox
End Class
