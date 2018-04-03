<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormTeamIncentiveConfig
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CMBPeriod = New System.Windows.Forms.ComboBox
        Me.CMBFiscalYear = New System.Windows.Forms.ComboBox
        Me.CMBSaleType = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.ListView101 = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ListView102 = New System.Windows.Forms.ListView
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader12 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader13 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader14 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader15 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader16 = New System.Windows.Forms.ColumnHeader
        Me.GB101 = New System.Windows.Forms.GroupBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.TextTotalPercent = New System.Windows.Forms.TextBox
        Me.LBLDepartment = New System.Windows.Forms.Label
        Me.BTNCancel = New System.Windows.Forms.Button
        Me.BTNSave = New System.Windows.Forms.Button
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.BTNInsert = New System.Windows.Forms.Button
        Me.NTeamBudget = New System.Windows.Forms.NumericUpDown
        Me.CMBTeam = New System.Windows.Forms.ComboBox
        Me.Label33 = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.Label31 = New System.Windows.Forms.Label
        Me.Label30 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.BTNClose = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GB101.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.NTeamBudget, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.Location = New System.Drawing.Point(36, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(308, 36)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "การกำหนด Incentive ของทีม"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CMBPeriod)
        Me.GroupBox1.Controls.Add(Me.CMBFiscalYear)
        Me.GroupBox1.Controls.Add(Me.CMBSaleType)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(34, 45)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(946, 75)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "รายละเอียดการดูข้อมูล"
        '
        'CMBPeriod
        '
        Me.CMBPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMBPeriod.FormattingEnabled = True
        Me.CMBPeriod.Location = New System.Drawing.Point(657, 36)
        Me.CMBPeriod.Name = "CMBPeriod"
        Me.CMBPeriod.Size = New System.Drawing.Size(134, 21)
        Me.CMBPeriod.TabIndex = 5
        '
        'CMBFiscalYear
        '
        Me.CMBFiscalYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMBFiscalYear.FormattingEnabled = True
        Me.CMBFiscalYear.Location = New System.Drawing.Point(379, 36)
        Me.CMBFiscalYear.Name = "CMBFiscalYear"
        Me.CMBFiscalYear.Size = New System.Drawing.Size(134, 21)
        Me.CMBFiscalYear.TabIndex = 4
        '
        'CMBSaleType
        '
        Me.CMBSaleType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMBSaleType.FormattingEnabled = True
        Me.CMBSaleType.Items.AddRange(New Object() {"ขายเงินเชื่อ", "ขายเงินสด"})
        Me.CMBSaleType.Location = New System.Drawing.Point(143, 36)
        Me.CMBSaleType.Name = "CMBSaleType"
        Me.CMBSaleType.Size = New System.Drawing.Size(134, 21)
        Me.CMBSaleType.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label4.Location = New System.Drawing.Point(536, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(116, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Period Of 4 Week :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.Location = New System.Drawing.Point(297, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Fiscal Year :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label2.Location = New System.Drawing.Point(39, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "ประเภทการขาย :"
        '
        'ListView101
        '
        Me.ListView101.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8, Me.ColumnHeader9})
        Me.ListView101.FullRowSelect = True
        Me.ListView101.GridLines = True
        Me.ListView101.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.ListView101.Location = New System.Drawing.Point(34, 153)
        Me.ListView101.Name = "ListView101"
        Me.ListView101.Size = New System.Drawing.Size(946, 481)
        Me.ListView101.TabIndex = 2
        Me.ListView101.UseCompatibleStateImageBehavior = False
        Me.ListView101.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Code"
        Me.ColumnHeader1.Width = 100
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Department"
        Me.ColumnHeader2.Width = 130
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "เป้าขาย"
        Me.ColumnHeader3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader3.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "เป้ากำไร"
        Me.ColumnHeader4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader4.Width = 100
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Budget Sale Min"
        Me.ColumnHeader5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader5.Width = 100
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Budget Sale Max"
        Me.ColumnHeader6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader6.Width = 100
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Budget GP Min"
        Me.ColumnHeader7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader7.Width = 100
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Budget GP Max"
        Me.ColumnHeader8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader8.Width = 100
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Budget Remain"
        Me.ColumnHeader9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader9.Width = 100
        '
        'ListView102
        '
        Me.ListView102.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader10, Me.ColumnHeader11, Me.ColumnHeader12, Me.ColumnHeader13, Me.ColumnHeader14, Me.ColumnHeader15, Me.ColumnHeader16})
        Me.ListView102.FullRowSelect = True
        Me.ListView102.GridLines = True
        Me.ListView102.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.ListView102.Location = New System.Drawing.Point(27, 239)
        Me.ListView102.Name = "ListView102"
        Me.ListView102.Size = New System.Drawing.Size(891, 245)
        Me.ListView102.TabIndex = 3
        Me.ListView102.UseCompatibleStateImageBehavior = False
        Me.ListView102.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "ทีม"
        Me.ColumnHeader10.Width = 285
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "Budget %"
        Me.ColumnHeader11.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader11.Width = 120
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "Budget Sale Min"
        Me.ColumnHeader12.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader12.Width = 120
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = "Budget Sale Max"
        Me.ColumnHeader13.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader13.Width = 120
        '
        'ColumnHeader14
        '
        Me.ColumnHeader14.Text = "Budget GP Min"
        Me.ColumnHeader14.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader14.Width = 120
        '
        'ColumnHeader15
        '
        Me.ColumnHeader15.Text = "Budget GP Max"
        Me.ColumnHeader15.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader15.Width = 80
        '
        'ColumnHeader16
        '
        Me.ColumnHeader16.Text = "Index"
        Me.ColumnHeader16.Width = 50
        '
        'GB101
        '
        Me.GB101.Controls.Add(Me.Label15)
        Me.GB101.Controls.Add(Me.TextTotalPercent)
        Me.GB101.Controls.Add(Me.LBLDepartment)
        Me.GB101.Controls.Add(Me.BTNCancel)
        Me.GB101.Controls.Add(Me.BTNSave)
        Me.GB101.Controls.Add(Me.Label12)
        Me.GB101.Controls.Add(Me.Label11)
        Me.GB101.Controls.Add(Me.Label10)
        Me.GB101.Controls.Add(Me.Label9)
        Me.GB101.Controls.Add(Me.Label8)
        Me.GB101.Controls.Add(Me.Label7)
        Me.GB101.Controls.Add(Me.GroupBox3)
        Me.GB101.Controls.Add(Me.ListView102)
        Me.GB101.Location = New System.Drawing.Point(34, 45)
        Me.GB101.Name = "GB101"
        Me.GB101.Size = New System.Drawing.Size(947, 642)
        Me.GB101.TabIndex = 5
        Me.GB101.TabStop = False
        Me.GB101.Text = "กำหนด Incentive Team ของแผนก :"
        Me.GB101.Visible = False
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label15.Location = New System.Drawing.Point(238, 496)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(71, 13)
        Me.Label15.TabIndex = 15
        Me.Label15.Text = "Maximum %"
        '
        'TextTotalPercent
        '
        Me.TextTotalPercent.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TextTotalPercent.Location = New System.Drawing.Point(314, 494)
        Me.TextTotalPercent.Name = "TextTotalPercent"
        Me.TextTotalPercent.Size = New System.Drawing.Size(121, 20)
        Me.TextTotalPercent.TabIndex = 14
        Me.TextTotalPercent.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LBLDepartment
        '
        Me.LBLDepartment.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.LBLDepartment.Location = New System.Drawing.Point(27, 27)
        Me.LBLDepartment.Name = "LBLDepartment"
        Me.LBLDepartment.Size = New System.Drawing.Size(891, 27)
        Me.LBLDepartment.TabIndex = 13
        Me.LBLDepartment.Text = "xxx"
        Me.LBLDepartment.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BTNCancel
        '
        Me.BTNCancel.Location = New System.Drawing.Point(833, 504)
        Me.BTNCancel.Name = "BTNCancel"
        Me.BTNCancel.Size = New System.Drawing.Size(85, 35)
        Me.BTNCancel.TabIndex = 12
        Me.BTNCancel.Text = "ยกเลิก"
        Me.BTNCancel.UseVisualStyleBackColor = True
        '
        'BTNSave
        '
        Me.BTNSave.Location = New System.Drawing.Point(726, 504)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(85, 35)
        Me.BTNSave.TabIndex = 11
        Me.BTNSave.Text = "บันทึก"
        Me.BTNSave.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label12.Location = New System.Drawing.Point(794, 239)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(124, 22)
        Me.Label12.TabIndex = 10
        Me.Label12.Text = "Budget GP Max"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label11.Location = New System.Drawing.Point(674, 239)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(120, 22)
        Me.Label11.TabIndex = 9
        Me.Label11.Text = "Budget GP Min"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label10.Location = New System.Drawing.Point(554, 239)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(120, 22)
        Me.Label10.TabIndex = 8
        Me.Label10.Text = "Budget Sale Max"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label9.Location = New System.Drawing.Point(435, 239)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(119, 22)
        Me.Label9.TabIndex = 7
        Me.Label9.Text = "Budget Sale Min"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label8.Location = New System.Drawing.Point(315, 239)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(120, 22)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "Budget %"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label7.Location = New System.Drawing.Point(28, 239)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(287, 22)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "ทีม"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label14)
        Me.GroupBox3.Controls.Add(Me.Label13)
        Me.GroupBox3.Controls.Add(Me.BTNInsert)
        Me.GroupBox3.Controls.Add(Me.NTeamBudget)
        Me.GroupBox3.Controls.Add(Me.CMBTeam)
        Me.GroupBox3.Location = New System.Drawing.Point(27, 63)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(891, 138)
        Me.GroupBox3.TabIndex = 4
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "รายละเอียด ทีม"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(35, 82)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(58, 13)
        Me.Label14.TabIndex = 4
        Me.Label14.Text = "กำหนด % :"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(26, 35)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(67, 13)
        Me.Label13.TabIndex = 3
        Me.Label13.Text = "เลือก Team :"
        '
        'BTNInsert
        '
        Me.BTNInsert.Location = New System.Drawing.Point(205, 80)
        Me.BTNInsert.Name = "BTNInsert"
        Me.BTNInsert.Size = New System.Drawing.Size(82, 34)
        Me.BTNInsert.TabIndex = 2
        Me.BTNInsert.Text = "ลงตาราง"
        Me.BTNInsert.UseVisualStyleBackColor = True
        '
        'NTeamBudget
        '
        Me.NTeamBudget.DecimalPlaces = 2
        Me.NTeamBudget.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.NTeamBudget.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NTeamBudget.Location = New System.Drawing.Point(99, 80)
        Me.NTeamBudget.Minimum = New Decimal(New Integer() {100, 0, 0, -2147483648})
        Me.NTeamBudget.Name = "NTeamBudget"
        Me.NTeamBudget.Size = New System.Drawing.Size(94, 20)
        Me.NTeamBudget.TabIndex = 1
        Me.NTeamBudget.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'CMBTeam
        '
        Me.CMBTeam.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMBTeam.FormattingEnabled = True
        Me.CMBTeam.Location = New System.Drawing.Point(99, 35)
        Me.CMBTeam.Name = "CMBTeam"
        Me.CMBTeam.Size = New System.Drawing.Size(376, 21)
        Me.CMBTeam.TabIndex = 0
        '
        'Label33
        '
        Me.Label33.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.Label33.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label33.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label33.Location = New System.Drawing.Point(767, 130)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(100, 23)
        Me.Label33.TabIndex = 26
        Me.Label33.Text = "Budget GP Max"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label32
        '
        Me.Label32.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.Label32.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label32.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label32.Location = New System.Drawing.Point(667, 130)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(100, 23)
        Me.Label32.TabIndex = 25
        Me.Label32.Text = "Budget GP Min"
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label31
        '
        Me.Label31.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.Label31.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label31.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label31.Location = New System.Drawing.Point(568, 130)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(99, 23)
        Me.Label31.TabIndex = 24
        Me.Label31.Text = "BudgetSaleMax"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label30
        '
        Me.Label30.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.Label30.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label30.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label30.Location = New System.Drawing.Point(467, 130)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(101, 23)
        Me.Label30.TabIndex = 23
        Me.Label30.Text = "BudgetSaleMin"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.Label28.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label28.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label28.Location = New System.Drawing.Point(366, 130)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(101, 23)
        Me.Label28.TabIndex = 21
        Me.Label28.Text = "เป้ากำไร"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label27
        '
        Me.Label27.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.Label27.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label27.Location = New System.Drawing.Point(266, 130)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(100, 23)
        Me.Label27.TabIndex = 20
        Me.Label27.Text = "เป้าขาย"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.Label26.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label26.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label26.Location = New System.Drawing.Point(136, 130)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(130, 23)
        Me.Label26.TabIndex = 19
        Me.Label26.Text = "แผนก"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label5.Location = New System.Drawing.Point(34, 130)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(102, 23)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "รหัสแผนก"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label6.Location = New System.Drawing.Point(867, 130)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(113, 23)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "Budget Remain"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BTNClose
        '
        Me.BTNClose.Location = New System.Drawing.Point(903, 640)
        Me.BTNClose.Name = "BTNClose"
        Me.BTNClose.Size = New System.Drawing.Size(78, 31)
        Me.BTNClose.TabIndex = 28
        Me.BTNClose.Text = "ออก"
        Me.BTNClose.UseVisualStyleBackColor = True
        '
        'FormTeamIncentiveConfig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDark
        Me.ClientSize = New System.Drawing.Size(1016, 734)
        Me.Controls.Add(Me.GB101)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label33)
        Me.Controls.Add(Me.Label32)
        Me.Controls.Add(Me.Label31)
        Me.Controls.Add(Me.Label30)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.ListView101)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BTNClose)
        Me.Name = "FormTeamIncentiveConfig"
        Me.Text = "FormTeamIncentiveConfig"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GB101.ResumeLayout(False)
        Me.GB101.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.NTeamBudget, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ListView101 As System.Windows.Forms.ListView
    Friend WithEvents ListView102 As System.Windows.Forms.ListView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CMBPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents CMBFiscalYear As System.Windows.Forms.ComboBox
    Friend WithEvents CMBSaleType As System.Windows.Forms.ComboBox
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
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader13 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader14 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader15 As System.Windows.Forms.ColumnHeader
    Friend WithEvents GB101 As System.Windows.Forms.GroupBox
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents CMBTeam As System.Windows.Forms.ComboBox
    Friend WithEvents NTeamBudget As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents BTNCancel As System.Windows.Forms.Button
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents BTNInsert As System.Windows.Forms.Button
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents LBLDepartment As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TextTotalPercent As System.Windows.Forms.TextBox
    Friend WithEvents BTNClose As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader16 As System.Windows.Forms.ColumnHeader
End Class
