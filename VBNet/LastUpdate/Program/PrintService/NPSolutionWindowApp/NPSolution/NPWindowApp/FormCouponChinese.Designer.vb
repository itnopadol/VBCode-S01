<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormCouponChinese
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
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TBCouponAmount = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TBDocNo = New System.Windows.Forms.TextBox
        Me.BTNSave = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TBArName = New System.Windows.Forms.TextBox
        Me.ListViewItem = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ListViewInvoice = New System.Windows.Forms.ListView
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.TBCoupon = New System.Windows.Forms.TextBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.BTNAddBill = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TBSumItemAmount = New System.Windows.Forms.TextBox
        Me.TBRemainAmount = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.BTNClearScreen = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.TBDiscountAmount = New System.Windows.Forms.TextBox
        Me.TBDisCountCoupon = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("PSL KandaAD", 20.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label3.Location = New System.Drawing.Point(482, 3)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(301, 34)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "ตรวจสอบ คูปอง และบันทึกการจ่ายคูปอง"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("PSL KandaAD", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label4.Location = New System.Drawing.Point(756, 506)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 27)
        Me.Label4.TabIndex = 26
        Me.Label4.Text = "มูลค่ารวม :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TBCouponAmount
        '
        Me.TBCouponAmount.BackColor = System.Drawing.SystemColors.Info
        Me.TBCouponAmount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBCouponAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBCouponAmount.Location = New System.Drawing.Point(835, 503)
        Me.TBCouponAmount.Name = "TBCouponAmount"
        Me.TBCouponAmount.ReadOnly = True
        Me.TBCouponAmount.Size = New System.Drawing.Size(180, 31)
        Me.TBCouponAmount.TabIndex = 23
        Me.TBCouponAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("PSL KandaAD", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label2.Location = New System.Drawing.Point(483, 49)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(119, 27)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "กรอกเลขที่เอกสาร :"
        '
        'TBDocNo
        '
        Me.TBDocNo.BackColor = System.Drawing.SystemColors.Info
        Me.TBDocNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBDocNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBDocNo.Location = New System.Drawing.Point(608, 49)
        Me.TBDocNo.Name = "TBDocNo"
        Me.TBDocNo.Size = New System.Drawing.Size(250, 31)
        Me.TBDocNo.TabIndex = 22
        '
        'BTNSave
        '
        Me.BTNSave.BackColor = System.Drawing.Color.LightGray
        Me.BTNSave.Font = New System.Drawing.Font("PSL KandaAD", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.BTNSave.Location = New System.Drawing.Point(894, 575)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(121, 57)
        Me.BTNSave.TabIndex = 24
        Me.BTNSave.Text = "บันทึกข้อมูล"
        Me.BTNSave.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("PSL KandaAD", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label1.Location = New System.Drawing.Point(555, 86)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 19)
        Me.Label1.TabIndex = 30
        Me.Label1.Text = "ชื่อลูกค้า :"
        '
        'TBArName
        '
        Me.TBArName.BackColor = System.Drawing.SystemColors.Info
        Me.TBArName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBArName.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBArName.Location = New System.Drawing.Point(608, 86)
        Me.TBArName.Name = "TBArName"
        Me.TBArName.Size = New System.Drawing.Size(407, 24)
        Me.TBArName.TabIndex = 29
        '
        'ListViewItem
        '
        Me.ListViewItem.BackColor = System.Drawing.Color.MistyRose
        Me.ListViewItem.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader8})
        Me.ListViewItem.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ListViewItem.FullRowSelect = True
        Me.ListViewItem.GridLines = True
        Me.ListViewItem.Location = New System.Drawing.Point(478, 140)
        Me.ListViewItem.Name = "ListViewItem"
        Me.ListViewItem.Size = New System.Drawing.Size(537, 161)
        Me.ListViewItem.TabIndex = 9
        Me.ListViewItem.UseCompatibleStateImageBehavior = False
        Me.ListViewItem.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ลำดับ"
        Me.ColumnHeader1.Width = 40
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "ประเภทสินค้า"
        Me.ColumnHeader2.Width = 100
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "ชื่อสินค้า"
        Me.ColumnHeader3.Width = 260
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "มูลค่า"
        Me.ColumnHeader8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader8.Width = 120
        '
        'ListViewInvoice
        '
        Me.ListViewInvoice.BackColor = System.Drawing.Color.AliceBlue
        Me.ListViewInvoice.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader11})
        Me.ListViewInvoice.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ListViewInvoice.FullRowSelect = True
        Me.ListViewInvoice.GridLines = True
        Me.ListViewInvoice.Location = New System.Drawing.Point(478, 388)
        Me.ListViewInvoice.Name = "ListViewInvoice"
        Me.ListViewInvoice.Size = New System.Drawing.Size(537, 110)
        Me.ListViewInvoice.TabIndex = 31
        Me.ListViewInvoice.UseCompatibleStateImageBehavior = False
        Me.ListViewInvoice.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "ลำดับ"
        Me.ColumnHeader4.Width = 40
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "เลขที่เอกสาร"
        Me.ColumnHeader5.Width = 120
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "ลูกค้า"
        Me.ColumnHeader6.Width = 250
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "มูลค่าบิล"
        Me.ColumnHeader11.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader11.Width = 120
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("PSL KandaAD", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label5.Location = New System.Drawing.Point(473, 361)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(157, 27)
        Me.Label5.TabIndex = 32
        Me.Label5.Text = "รายการ เอกสารที่จ่ายคูปอง"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("PSL KandaAD", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label7.Location = New System.Drawing.Point(474, 113)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(127, 24)
        Me.Label7.TabIndex = 33
        Me.Label7.Text = "รายการ สินค้าของเอกสาร"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("PSL KandaAD", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label8.Location = New System.Drawing.Point(741, 538)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(88, 27)
        Me.Label8.TabIndex = 35
        Me.Label8.Text = "จำนวนคูปอง :"
        '
        'TBCoupon
        '
        Me.TBCoupon.BackColor = System.Drawing.SystemColors.Info
        Me.TBCoupon.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBCoupon.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBCoupon.Location = New System.Drawing.Point(835, 538)
        Me.TBCoupon.Name = "TBCoupon"
        Me.TBCoupon.ReadOnly = True
        Me.TBCoupon.Size = New System.Drawing.Size(146, 31)
        Me.TBCoupon.TabIndex = 34
        Me.TBCoupon.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'PictureBox2
        '
        'Me.PictureBox2.Image = Global.NPWindowApp.My.Resources.Resources.chinese1
        Me.PictureBox2.Location = New System.Drawing.Point(140, 301)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(190, 320)
        Me.PictureBox2.TabIndex = 36
        Me.PictureBox2.TabStop = False
        '
        'BTNAddBill
        '
        Me.BTNAddBill.BackColor = System.Drawing.Color.LightGray
        Me.BTNAddBill.Font = New System.Drawing.Font("PSL KandaAD", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        'Me.BTNAddBill.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_76
        Me.BTNAddBill.Location = New System.Drawing.Point(478, 305)
        Me.BTNAddBill.Name = "BTNAddBill"
        Me.BTNAddBill.Size = New System.Drawing.Size(93, 50)
        Me.BTNAddBill.TabIndex = 27
        Me.BTNAddBill.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        'Me.PictureBox1.Image = Global.NPWindowApp.My.Resources.Resources.chinese
        Me.PictureBox1.Location = New System.Drawing.Point(2, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(453, 303)
        Me.PictureBox1.TabIndex = 19
        Me.PictureBox1.TabStop = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("PSL KandaAD", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label9.Location = New System.Drawing.Point(830, 306)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(57, 20)
        Me.Label9.TabIndex = 38
        Me.Label9.Text = "มูลค่าสินค้า :"
        '
        'TBSumItemAmount
        '
        Me.TBSumItemAmount.BackColor = System.Drawing.SystemColors.Info
        Me.TBSumItemAmount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBSumItemAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBSumItemAmount.Location = New System.Drawing.Point(893, 305)
        Me.TBSumItemAmount.Name = "TBSumItemAmount"
        Me.TBSumItemAmount.ReadOnly = True
        Me.TBSumItemAmount.Size = New System.Drawing.Size(122, 24)
        Me.TBSumItemAmount.TabIndex = 37
        Me.TBSumItemAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TBRemainAmount
        '
        Me.TBRemainAmount.BackColor = System.Drawing.SystemColors.Info
        Me.TBRemainAmount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBRemainAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBRemainAmount.Location = New System.Drawing.Point(478, 503)
        Me.TBRemainAmount.Name = "TBRemainAmount"
        Me.TBRemainAmount.ReadOnly = True
        Me.TBRemainAmount.Size = New System.Drawing.Size(152, 31)
        Me.TBRemainAmount.TabIndex = 39
        Me.TBRemainAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("PSL KandaAD", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label10.Location = New System.Drawing.Point(987, 538)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(24, 27)
        Me.Label10.TabIndex = 40
        Me.Label10.Text = "ใบ"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("PSL KandaAD", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label11.Location = New System.Drawing.Point(349, 506)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(123, 27)
        Me.Label11.TabIndex = 41
        Me.Label11.Text = "ซื้อสินค้าทั่วไปเพิ่มอีก"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("PSL KandaAD", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label12.Location = New System.Drawing.Point(349, 538)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(140, 27)
        Me.Label12.TabIndex = 42
        Me.Label12.Text = "จะได้คูปองเพิ่มอีก  1 ใบ"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BTNClearScreen
        '
        Me.BTNClearScreen.BackColor = System.Drawing.Color.LightGray
        Me.BTNClearScreen.Font = New System.Drawing.Font("PSL KandaAD", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.BTNClearScreen.Location = New System.Drawing.Point(478, 575)
        Me.BTNClearScreen.Name = "BTNClearScreen"
        Me.BTNClearScreen.Size = New System.Drawing.Size(121, 57)
        Me.BTNClearScreen.TabIndex = 43
        Me.BTNClearScreen.Text = "ล้างหน้าจอ"
        Me.BTNClearScreen.UseVisualStyleBackColor = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("PSL KandaAD", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label6.Location = New System.Drawing.Point(636, 506)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 27)
        Me.Label6.TabIndex = 44
        Me.Label6.Text = "บาท"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TBDiscountAmount
        '
        Me.TBDiscountAmount.BackColor = System.Drawing.SystemColors.Info
        Me.TBDiscountAmount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBDiscountAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBDiscountAmount.Location = New System.Drawing.Point(893, 333)
        Me.TBDiscountAmount.Name = "TBDiscountAmount"
        Me.TBDiscountAmount.ReadOnly = True
        Me.TBDiscountAmount.Size = New System.Drawing.Size(122, 24)
        Me.TBDiscountAmount.TabIndex = 45
        Me.TBDiscountAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TBDisCountCoupon
        '
        Me.TBDisCountCoupon.BackColor = System.Drawing.SystemColors.Info
        Me.TBDisCountCoupon.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBDisCountCoupon.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBDisCountCoupon.Location = New System.Drawing.Point(893, 361)
        Me.TBDisCountCoupon.Name = "TBDisCountCoupon"
        Me.TBDisCountCoupon.ReadOnly = True
        Me.TBDisCountCoupon.Size = New System.Drawing.Size(122, 24)
        Me.TBDisCountCoupon.TabIndex = 46
        Me.TBDisCountCoupon.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("PSL KandaAD", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label13.Location = New System.Drawing.Point(821, 335)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(66, 20)
        Me.Label13.TabIndex = 47
        Me.Label13.Text = "มูลค่าส่วนลด :"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("PSL KandaAD", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label14.Location = New System.Drawing.Point(827, 365)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(60, 20)
        Me.Label14.TabIndex = 48
        Me.Label14.Text = "มูลค่าคูปอง :"
        '
        'FormCouponChinese
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1212, 726)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.TBDisCountCoupon)
        Me.Controls.Add(Me.TBDiscountAmount)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.BTNClearScreen)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.TBRemainAmount)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.TBSumItemAmount)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.TBCoupon)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.ListViewInvoice)
        Me.Controls.Add(Me.ListViewItem)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TBArName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TBCouponAmount)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TBDocNo)
        Me.Controls.Add(Me.BTNSave)
        Me.Controls.Add(Me.BTNAddBill)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.PictureBox2)
        Me.Name = "FormCouponChinese"
        Me.Text = "FormCouponChinese"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TBCouponAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TBDocNo As System.Windows.Forms.TextBox
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents BTNAddBill As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TBArName As System.Windows.Forms.TextBox
    Friend WithEvents ListViewItem As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ListViewInvoice As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TBCoupon As System.Windows.Forms.TextBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TBSumItemAmount As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents TBRemainAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents BTNClearScreen As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TBDiscountAmount As System.Windows.Forms.TextBox
    Friend WithEvents TBDisCountCoupon As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
End Class
