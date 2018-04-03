<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormBudgetConfig
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
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.BTNSave = New System.Windows.Forms.Button
        Me.CMBSaleType = New System.Windows.Forms.ComboBox
        Me.CMBFiscalYear = New System.Windows.Forms.ComboBox
        Me.CMBPeriod = New System.Windows.Forms.ComboBox
        Me.TextSaleMin = New System.Windows.Forms.TextBox
        Me.TextSaleMax = New System.Windows.Forms.TextBox
        Me.TextGPMin = New System.Windows.Forms.TextBox
        Me.TextGPMax = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.NUDReturnItem = New System.Windows.Forms.NumericUpDown
        Me.BTNExit = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        CType(Me.NUDReturnItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.Location = New System.Drawing.Point(293, 127)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ประเภทการขาย :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label2.Location = New System.Drawing.Point(318, 168)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Fiscal Year :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.Location = New System.Drawing.Point(279, 210)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(116, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Period Of 4 Week :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label4.Location = New System.Drawing.Point(279, 252)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(116, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "จำนวนวันคืนสินค้า :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label5.Location = New System.Drawing.Point(272, 301)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Budget Sale "
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label6.Location = New System.Drawing.Point(283, 340)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Budget GP "
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label7.Location = New System.Drawing.Point(355, 301)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(38, 13)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "MIN :"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label8.Location = New System.Drawing.Point(550, 301)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(41, 13)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "MAX :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label9.Location = New System.Drawing.Point(355, 340)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(38, 13)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "MIN :"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label10.Location = New System.Drawing.Point(550, 340)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(41, 13)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "MAX :"
        '
        'BTNSave
        '
        Me.BTNSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.BTNSave.Location = New System.Drawing.Point(542, 397)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(99, 38)
        Me.BTNSave.TabIndex = 8
        Me.BTNSave.Text = "บันทึกข้อมูล"
        Me.BTNSave.UseVisualStyleBackColor = True
        '
        'CMBSaleType
        '
        Me.CMBSaleType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMBSaleType.FormattingEnabled = True
        Me.CMBSaleType.Items.AddRange(New Object() {"ขายเงินเชื่อ", "ขายเงินสด"})
        Me.CMBSaleType.Location = New System.Drawing.Point(401, 124)
        Me.CMBSaleType.Name = "CMBSaleType"
        Me.CMBSaleType.Size = New System.Drawing.Size(139, 21)
        Me.CMBSaleType.TabIndex = 0
        '
        'CMBFiscalYear
        '
        Me.CMBFiscalYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMBFiscalYear.FormattingEnabled = True
        Me.CMBFiscalYear.Location = New System.Drawing.Point(401, 165)
        Me.CMBFiscalYear.Name = "CMBFiscalYear"
        Me.CMBFiscalYear.Size = New System.Drawing.Size(139, 21)
        Me.CMBFiscalYear.TabIndex = 1
        '
        'CMBPeriod
        '
        Me.CMBPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMBPeriod.FormattingEnabled = True
        Me.CMBPeriod.Location = New System.Drawing.Point(401, 207)
        Me.CMBPeriod.Name = "CMBPeriod"
        Me.CMBPeriod.Size = New System.Drawing.Size(139, 21)
        Me.CMBPeriod.TabIndex = 2
        '
        'TextSaleMin
        '
        Me.TextSaleMin.Location = New System.Drawing.Point(401, 298)
        Me.TextSaleMin.Name = "TextSaleMin"
        Me.TextSaleMin.Size = New System.Drawing.Size(139, 20)
        Me.TextSaleMin.TabIndex = 4
        Me.TextSaleMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextSaleMax
        '
        Me.TextSaleMax.Location = New System.Drawing.Point(599, 298)
        Me.TextSaleMax.Name = "TextSaleMax"
        Me.TextSaleMax.Size = New System.Drawing.Size(142, 20)
        Me.TextSaleMax.TabIndex = 5
        Me.TextSaleMax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextGPMin
        '
        Me.TextGPMin.Location = New System.Drawing.Point(401, 337)
        Me.TextGPMin.Name = "TextGPMin"
        Me.TextGPMin.Size = New System.Drawing.Size(142, 20)
        Me.TextGPMin.TabIndex = 6
        Me.TextGPMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextGPMax
        '
        Me.TextGPMax.Location = New System.Drawing.Point(599, 337)
        Me.TextGPMax.Name = "TextGPMax"
        Me.TextGPMax.Size = New System.Drawing.Size(142, 20)
        Me.TextGPMax.TabIndex = 7
        Me.TextGPMax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, CType(((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic) _
                        Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label11.Location = New System.Drawing.Point(277, 59)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(161, 25)
        Me.Label11.TabIndex = 19
        Me.Label11.Text = "Budget Config"
        '
        'NUDReturnItem
        '
        Me.NUDReturnItem.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.NUDReturnItem.Location = New System.Drawing.Point(401, 250)
        Me.NUDReturnItem.Name = "NUDReturnItem"
        Me.NUDReturnItem.ReadOnly = True
        Me.NUDReturnItem.Size = New System.Drawing.Size(140, 20)
        Me.NUDReturnItem.TabIndex = 3
        Me.NUDReturnItem.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NUDReturnItem.Value = New Decimal(New Integer() {28, 0, 0, 0})
        '
        'BTNExit
        '
        Me.BTNExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BTNExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.BTNExit.Location = New System.Drawing.Point(647, 397)
        Me.BTNExit.Name = "BTNExit"
        Me.BTNExit.Size = New System.Drawing.Size(94, 38)
        Me.BTNExit.TabIndex = 20
        Me.BTNExit.Text = "ออก"
        Me.BTNExit.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.NPWindowApp.My.Resources.Resources.LogoNopadol_144x50
        Me.PictureBox1.Location = New System.Drawing.Point(1, 1)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(145, 50)
        Me.PictureBox1.TabIndex = 21
        Me.PictureBox1.TabStop = False
        '
        'FormBudgetConfig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDark
        Me.ClientSize = New System.Drawing.Size(1156, 734)
        Me.Controls.Add(Me.TextSaleMin)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.BTNExit)
        Me.Controls.Add(Me.NUDReturnItem)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.TextGPMax)
        Me.Controls.Add(Me.TextGPMin)
        Me.Controls.Add(Me.TextSaleMax)
        Me.Controls.Add(Me.CMBPeriod)
        Me.Controls.Add(Me.CMBFiscalYear)
        Me.Controls.Add(Me.CMBSaleType)
        Me.Controls.Add(Me.BTNSave)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Name = "FormBudgetConfig"
        Me.Text = "FormBudgetConfig"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.NUDReturnItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents CMBSaleType As System.Windows.Forms.ComboBox
    Friend WithEvents CMBFiscalYear As System.Windows.Forms.ComboBox
    Friend WithEvents CMBPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents TextSaleMin As System.Windows.Forms.TextBox
    Friend WithEvents TextSaleMax As System.Windows.Forms.TextBox
    Friend WithEvents TextGPMin As System.Windows.Forms.TextBox
    Friend WithEvents TextGPMax As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents NUDReturnItem As System.Windows.Forms.NumericUpDown
    Friend WithEvents BTNExit As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
End Class
