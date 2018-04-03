<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormCouponRecord
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
        Me.TextCouponNo = New System.Windows.Forms.TextBox
        Me.TextCouponName = New System.Windows.Forms.TextBox
        Me.TextCouponQTY = New System.Windows.Forms.TextBox
        Me.TextCountLenght = New System.Windows.Forms.TextBox
        Me.TextCouponFormat = New System.Windows.Forms.TextBox
        Me.BTNBasket = New System.Windows.Forms.Button
        Me.ListViewCoupon = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader("(none)")
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CMBFormat = New System.Windows.Forms.ComboBox
        Me.CBMerge = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.BTNSave = New System.Windows.Forms.Button
        Me.BTNClearScreen = New System.Windows.Forms.Button
        Me.BTNCanCel = New System.Windows.Forms.Button
        Me.BTNExit = New System.Windows.Forms.Button
        Me.TextCouponValue = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TextCouponCode = New System.Windows.Forms.TextBox
        Me.BTNGenDocNo = New System.Windows.Forms.Button
        Me.DTPStartDate = New System.Windows.Forms.DateTimePicker
        Me.DTPStopDate = New System.Windows.Forms.DateTimePicker
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.TextMyDescription = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PB102 = New System.Windows.Forms.PictureBox
        Me.PB101 = New System.Windows.Forms.PictureBox
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PB102, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PB101, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TextCouponNo
        '
        Me.TextCouponNo.BackColor = System.Drawing.Color.White
        Me.TextCouponNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TextCouponNo.ForeColor = System.Drawing.Color.Black
        Me.TextCouponNo.Location = New System.Drawing.Point(151, 172)
        Me.TextCouponNo.Name = "TextCouponNo"
        Me.TextCouponNo.Size = New System.Drawing.Size(100, 20)
        Me.TextCouponNo.TabIndex = 12
        Me.TextCouponNo.Tag = "vvvv"
        '
        'TextCouponName
        '
        Me.TextCouponName.BackColor = System.Drawing.SystemColors.Info
        Me.TextCouponName.Location = New System.Drawing.Point(151, 92)
        Me.TextCouponName.Name = "TextCouponName"
        Me.TextCouponName.Size = New System.Drawing.Size(768, 20)
        Me.TextCouponName.TabIndex = 2
        '
        'TextCouponQTY
        '
        Me.TextCouponQTY.BackColor = System.Drawing.Color.Black
        Me.TextCouponQTY.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TextCouponQTY.ForeColor = System.Drawing.Color.Gold
        Me.TextCouponQTY.Location = New System.Drawing.Point(333, 172)
        Me.TextCouponQTY.Name = "TextCouponQTY"
        Me.TextCouponQTY.Size = New System.Drawing.Size(100, 20)
        Me.TextCouponQTY.TabIndex = 13
        Me.TextCouponQTY.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextCountLenght
        '
        Me.TextCountLenght.BackColor = System.Drawing.Color.Black
        Me.TextCountLenght.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TextCountLenght.ForeColor = System.Drawing.Color.Gold
        Me.TextCountLenght.Location = New System.Drawing.Point(151, 146)
        Me.TextCountLenght.Name = "TextCountLenght"
        Me.TextCountLenght.Size = New System.Drawing.Size(100, 20)
        Me.TextCountLenght.TabIndex = 5
        '
        'TextCouponFormat
        '
        Me.TextCouponFormat.BackColor = System.Drawing.Color.Black
        Me.TextCouponFormat.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TextCouponFormat.ForeColor = System.Drawing.Color.Gold
        Me.TextCouponFormat.Location = New System.Drawing.Point(333, 146)
        Me.TextCouponFormat.Name = "TextCouponFormat"
        Me.TextCouponFormat.ReadOnly = True
        Me.TextCouponFormat.Size = New System.Drawing.Size(100, 20)
        Me.TextCouponFormat.TabIndex = 6
        Me.TextCouponFormat.Text = "0"
        Me.TextCouponFormat.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'BTNBasket
        '
        Me.BTNBasket.Location = New System.Drawing.Point(177, 250)
        Me.BTNBasket.Name = "BTNBasket"
        Me.BTNBasket.Size = New System.Drawing.Size(75, 33)
        Me.BTNBasket.TabIndex = 15
        Me.BTNBasket.Text = "ลงตาราง"
        Me.BTNBasket.UseVisualStyleBackColor = True
        '
        'ListViewCoupon
        '
        Me.ListViewCoupon.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5})
        Me.ListViewCoupon.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ListViewCoupon.FullRowSelect = True
        Me.ListViewCoupon.GridLines = True
        Me.ListViewCoupon.LabelEdit = True
        Me.ListViewCoupon.Location = New System.Drawing.Point(97, 332)
        Me.ListViewCoupon.Name = "ListViewCoupon"
        Me.ListViewCoupon.Size = New System.Drawing.Size(822, 222)
        Me.ListViewCoupon.TabIndex = 16
        Me.ListViewCoupon.UseCompatibleStateImageBehavior = False
        Me.ListViewCoupon.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "หัวคูปอง"
        Me.ColumnHeader1.Width = 178
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "มูลค่า"
        Me.ColumnHeader2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader2.Width = 160
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "จำนวน"
        Me.ColumnHeader3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader3.Width = 160
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "จำนวนที่อนุมัติแล้ว"
        Me.ColumnHeader4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader4.Width = 160
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "คงเหลือ"
        Me.ColumnHeader5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader5.Width = 160
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CMBFormat)
        Me.GroupBox1.Controls.Add(Me.CBMerge)
        Me.GroupBox1.Location = New System.Drawing.Point(465, 115)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(454, 77)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        '
        'CMBFormat
        '
        Me.CMBFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMBFormat.FormattingEnabled = True
        Me.CMBFormat.Items.AddRange(New Object() {"000", "xxx000", "xxx-000", "yy000", "yy-000", "yymm000", "yymm-000"})
        Me.CMBFormat.Location = New System.Drawing.Point(6, 31)
        Me.CMBFormat.Name = "CMBFormat"
        Me.CMBFormat.Size = New System.Drawing.Size(121, 21)
        Me.CMBFormat.TabIndex = 12
        '
        'CBMerge
        '
        Me.CBMerge.AutoSize = True
        Me.CBMerge.Location = New System.Drawing.Point(148, 33)
        Me.CBMerge.Name = "CBMerge"
        Me.CBMerge.Size = New System.Drawing.Size(108, 17)
        Me.CBMerge.TabIndex = 7
        Me.CBMerge.Text = "รวม เลข Running"
        Me.CBMerge.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(56, 95)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "ชื่อทะเบียนคูปอง :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(81, 123)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "เริ่มใช้วันที่ :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(281, 123)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "ถึงวันที่ :"
        '
        'BTNSave
        '
        Me.BTNSave.Location = New System.Drawing.Point(95, 588)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(100, 43)
        Me.BTNSave.TabIndex = 17
        Me.BTNSave.Text = "บันทึก"
        Me.BTNSave.UseVisualStyleBackColor = True
        '
        'BTNClearScreen
        '
        Me.BTNClearScreen.Location = New System.Drawing.Point(205, 588)
        Me.BTNClearScreen.Name = "BTNClearScreen"
        Me.BTNClearScreen.Size = New System.Drawing.Size(100, 43)
        Me.BTNClearScreen.TabIndex = 18
        Me.BTNClearScreen.Text = "ล้างหน้าจอ"
        Me.BTNClearScreen.UseVisualStyleBackColor = True
        '
        'BTNCanCel
        '
        Me.BTNCanCel.Location = New System.Drawing.Point(316, 588)
        Me.BTNCanCel.Name = "BTNCanCel"
        Me.BTNCanCel.Size = New System.Drawing.Size(100, 43)
        Me.BTNCanCel.TabIndex = 19
        Me.BTNCanCel.Text = "ลบคูปอง"
        Me.BTNCanCel.UseVisualStyleBackColor = True
        '
        'BTNExit
        '
        Me.BTNExit.Location = New System.Drawing.Point(427, 588)
        Me.BTNExit.Name = "BTNExit"
        Me.BTNExit.Size = New System.Drawing.Size(100, 43)
        Me.BTNExit.TabIndex = 20
        Me.BTNExit.Text = "ออก"
        Me.BTNExit.UseVisualStyleBackColor = True
        '
        'TextCouponValue
        '
        Me.TextCouponValue.BackColor = System.Drawing.Color.Black
        Me.TextCouponValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TextCouponValue.ForeColor = System.Drawing.Color.Gold
        Me.TextCouponValue.Location = New System.Drawing.Point(151, 198)
        Me.TextCouponValue.Name = "TextCouponValue"
        Me.TextCouponValue.Size = New System.Drawing.Size(100, 20)
        Me.TextCouponValue.TabIndex = 14
        Me.TextCouponValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(438, 149)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(25, 13)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = ">>>"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label5.Location = New System.Drawing.Point(94, 310)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(117, 13)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "รายการ คูปองต่าง ๆ "
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label9.Location = New System.Drawing.Point(437, 330)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(161, 22)
        Me.Label9.TabIndex = 25
        Me.Label9.Text = "จำนวนคูปอง"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label8.Location = New System.Drawing.Point(277, 330)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(161, 22)
        Me.Label8.TabIndex = 24
        Me.Label8.Text = "มูลค่า"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label7.Location = New System.Drawing.Point(97, 330)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(181, 22)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "หัวคูปอง"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(68, 67)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(77, 13)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "ทะเบียนคูปอง :"
        '
        'TextCouponCode
        '
        Me.TextCouponCode.BackColor = System.Drawing.SystemColors.Info
        Me.TextCouponCode.Location = New System.Drawing.Point(151, 64)
        Me.TextCouponCode.Name = "TextCouponCode"
        Me.TextCouponCode.Size = New System.Drawing.Size(100, 20)
        Me.TextCouponCode.TabIndex = 0
        '
        'BTNGenDocNo
        '
        Me.BTNGenDocNo.Location = New System.Drawing.Point(253, 63)
        Me.BTNGenDocNo.Name = "BTNGenDocNo"
        Me.BTNGenDocNo.Size = New System.Drawing.Size(28, 22)
        Me.BTNGenDocNo.TabIndex = 1
        Me.BTNGenDocNo.UseVisualStyleBackColor = True
        '
        'DTPStartDate
        '
        Me.DTPStartDate.CalendarMonthBackground = System.Drawing.SystemColors.Info
        Me.DTPStartDate.CalendarTitleForeColor = System.Drawing.SystemColors.Info
        Me.DTPStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPStartDate.Location = New System.Drawing.Point(152, 120)
        Me.DTPStartDate.Name = "DTPStartDate"
        Me.DTPStartDate.Size = New System.Drawing.Size(100, 20)
        Me.DTPStartDate.TabIndex = 3
        '
        'DTPStopDate
        '
        Me.DTPStopDate.CalendarMonthBackground = System.Drawing.SystemColors.Info
        Me.DTPStopDate.CalendarTitleForeColor = System.Drawing.SystemColors.Info
        Me.DTPStopDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPStopDate.Location = New System.Drawing.Point(333, 120)
        Me.DTPStopDate.Name = "DTPStopDate"
        Me.DTPStopDate.Size = New System.Drawing.Size(100, 20)
        Me.DTPStopDate.TabIndex = 4
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(94, 175)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(51, 13)
        Me.Label11.TabIndex = 32
        Me.Label11.Text = "หัวคูปอง :"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(92, 149)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(53, 13)
        Me.Label12.TabIndex = 33
        Me.Label12.Text = "ตำแหน่ง :"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(107, 201)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(38, 13)
        Me.Label13.TabIndex = 34
        Me.Label13.Text = "มูลค่า :"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(279, 149)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(48, 13)
        Me.Label14.TabIndex = 35
        Me.Label14.Text = "รูปแบบ :"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(256, 175)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(71, 13)
        Me.Label15.TabIndex = 36
        Me.Label15.Text = "จำนวนคูปอง :"
        '
        'TextMyDescription
        '
        Me.TextMyDescription.Location = New System.Drawing.Point(465, 198)
        Me.TextMyDescription.Multiline = True
        Me.TextMyDescription.Name = "TextMyDescription"
        Me.TextMyDescription.Size = New System.Drawing.Size(454, 50)
        Me.TextMyDescription.TabIndex = 37
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(401, 201)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(58, 13)
        Me.Label10.TabIndex = 38
        Me.Label10.Text = "หมายเหตุ :"
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label16.Location = New System.Drawing.Point(597, 330)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(161, 22)
        Me.Label16.TabIndex = 41
        Me.Label16.Text = "จำนวนอนุมัติแล้ว"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label17.Location = New System.Drawing.Point(757, 330)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(161, 22)
        Me.Label17.TabIndex = 42
        Me.Label17.Text = "จำนวนคงเหลือ"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        'Me.PictureBox1.Image = Global.NPWindowApp.My.Resources.Resources.LogoNopadol_144x50
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(146, 50)
        Me.PictureBox1.TabIndex = 43
        Me.PictureBox1.TabStop = False
        '
        'PB102
        '
        'Me.PB102.Image = Global.NPWindowApp.My.Resources.Resources.Confirm
        Me.PB102.Location = New System.Drawing.Point(196, 14)
        Me.PB102.Name = "PB102"
        Me.PB102.Size = New System.Drawing.Size(38, 20)
        Me.PB102.TabIndex = 40
        Me.PB102.TabStop = False
        Me.PB102.Visible = False
        '
        'PB101
        '
        'Me.PB101.Image = Global.NPWindowApp.My.Resources.Resources._New
        Me.PB101.Location = New System.Drawing.Point(152, 14)
        Me.PB101.Name = "PB101"
        Me.PB101.Size = New System.Drawing.Size(38, 20)
        Me.PB101.TabIndex = 39
        Me.PB101.TabStop = False
        '
        'FormCouponRecord
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDark
        Me.ClientSize = New System.Drawing.Size(1016, 734)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.PB102)
        Me.Controls.Add(Me.PB101)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.TextMyDescription)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.DTPStopDate)
        Me.Controls.Add(Me.DTPStartDate)
        Me.Controls.Add(Me.BTNGenDocNo)
        Me.Controls.Add(Me.TextCouponCode)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextCouponValue)
        Me.Controls.Add(Me.BTNExit)
        Me.Controls.Add(Me.BTNCanCel)
        Me.Controls.Add(Me.BTNClearScreen)
        Me.Controls.Add(Me.BTNSave)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ListViewCoupon)
        Me.Controls.Add(Me.BTNBasket)
        Me.Controls.Add(Me.TextCouponFormat)
        Me.Controls.Add(Me.TextCountLenght)
        Me.Controls.Add(Me.TextCouponQTY)
        Me.Controls.Add(Me.TextCouponName)
        Me.Controls.Add(Me.TextCouponNo)
        Me.Name = "FormCouponRecord"
        Me.Text = "ฟอร์มเพิ่มทะเบียนคูปอง"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PB102, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PB101, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextCouponNo As System.Windows.Forms.TextBox
    Friend WithEvents TextCouponName As System.Windows.Forms.TextBox
    Friend WithEvents TextCouponQTY As System.Windows.Forms.TextBox
    Friend WithEvents TextCountLenght As System.Windows.Forms.TextBox
    Friend WithEvents TextCouponFormat As System.Windows.Forms.TextBox
    Friend WithEvents BTNBasket As System.Windows.Forms.Button
    Friend WithEvents ListViewCoupon As System.Windows.Forms.ListView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CBMerge As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents BTNClearScreen As System.Windows.Forms.Button
    Friend WithEvents BTNCanCel As System.Windows.Forms.Button
    Friend WithEvents BTNExit As System.Windows.Forms.Button
    Friend WithEvents TextCouponValue As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextCouponCode As System.Windows.Forms.TextBox
    Friend WithEvents BTNGenDocNo As System.Windows.Forms.Button
    Friend WithEvents DTPStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents DTPStopDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents CMBFormat As System.Windows.Forms.ComboBox
    Friend WithEvents TextMyDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents PB101 As System.Windows.Forms.PictureBox
    Friend WithEvents PB102 As System.Windows.Forms.PictureBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
End Class
