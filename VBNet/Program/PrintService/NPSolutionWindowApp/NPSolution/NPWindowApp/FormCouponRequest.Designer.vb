<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormCouponRequest
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormCouponRequest))
        Me.TextCPName = New System.Windows.Forms.TextBox
        Me.ListViewReqCP = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.TextDocNo = New System.Windows.Forms.TextBox
        Me.TextCPValue = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.BTNSave = New System.Windows.Forms.Button
        Me.BTNCancel = New System.Windows.Forms.Button
        Me.BTNSearch = New System.Windows.Forms.Button
        Me.TextReqReason = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.ListViewGenCoupon = New System.Windows.Forms.ListView
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.TextCPCode = New System.Windows.Forms.TextBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.DTPDocDate = New System.Windows.Forms.DateTimePicker
        Me.CMBReqUserID = New System.Windows.Forms.ComboBox
        Me.GB101 = New System.Windows.Forms.GroupBox
        Me.TextQTY = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.BTNGenDocNo = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GB101.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextCPName
        '
        Me.TextCPName.Location = New System.Drawing.Point(600, 91)
        Me.TextCPName.Name = "TextCPName"
        Me.TextCPName.Size = New System.Drawing.Size(312, 20)
        Me.TextCPName.TabIndex = 0
        '
        'ListViewReqCP
        '
        Me.ListViewReqCP.CheckBoxes = True
        Me.ListViewReqCP.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.ListViewReqCP.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ListViewReqCP.FullRowSelect = True
        Me.ListViewReqCP.GridLines = True
        Me.ListViewReqCP.Location = New System.Drawing.Point(105, 188)
        Me.ListViewReqCP.Name = "ListViewReqCP"
        Me.ListViewReqCP.Size = New System.Drawing.Size(805, 176)
        Me.ListViewReqCP.TabIndex = 3
        Me.ListViewReqCP.UseCompatibleStateImageBehavior = False
        Me.ListViewReqCP.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Width = 135
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader2.Width = 135
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader3.Width = 135
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader4.Width = 135
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader5.Width = 130
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader6.Width = 130
        '
        'TextDocNo
        '
        Me.TextDocNo.Location = New System.Drawing.Point(191, 60)
        Me.TextDocNo.Name = "TextDocNo"
        Me.TextDocNo.Size = New System.Drawing.Size(100, 20)
        Me.TextDocNo.TabIndex = 4
        '
        'TextCPValue
        '
        Me.TextCPValue.Location = New System.Drawing.Point(191, 156)
        Me.TextCPValue.Name = "TextCPValue"
        Me.TextCPValue.Size = New System.Drawing.Size(100, 20)
        Me.TextCPValue.TabIndex = 5
        Me.TextCPValue.Text = "มูลค่าคูปองทั้งหมด"
        '
        'Label1
        '
        Me.Label1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Location = New System.Drawing.Point(85, 156)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 20)
        Me.Label1.TabIndex = 7
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'BTNSave
        '
        Me.BTNSave.Location = New System.Drawing.Point(107, 580)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(100, 50)
        Me.BTNSave.TabIndex = 9
        Me.BTNSave.Text = "บันทึก"
        Me.BTNSave.UseVisualStyleBackColor = True
        '
        'BTNCancel
        '
        Me.BTNCancel.Location = New System.Drawing.Point(235, 580)
        Me.BTNCancel.Name = "BTNCancel"
        Me.BTNCancel.Size = New System.Drawing.Size(100, 50)
        Me.BTNCancel.TabIndex = 10
        Me.BTNCancel.Text = "ยกเลิกเอกสาร"
        Me.BTNCancel.UseVisualStyleBackColor = True
        '
        'BTNSearch
        '
        Me.BTNSearch.Location = New System.Drawing.Point(367, 580)
        Me.BTNSearch.Name = "BTNSearch"
        Me.BTNSearch.Size = New System.Drawing.Size(100, 50)
        Me.BTNSearch.TabIndex = 11
        Me.BTNSearch.Text = "ค้นหา"
        Me.BTNSearch.UseVisualStyleBackColor = True
        '
        'TextReqReason
        '
        Me.TextReqReason.Location = New System.Drawing.Point(191, 123)
        Me.TextReqReason.Name = "TextReqReason"
        Me.TextReqReason.Size = New System.Drawing.Size(721, 20)
        Me.TextReqReason.TabIndex = 15
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label10.Location = New System.Drawing.Point(511, 186)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(136, 22)
        Me.Label10.TabIndex = 19
        Me.Label10.Text = "อนุมัติแล้ว"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label9.Location = New System.Drawing.Point(377, 186)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(135, 22)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "จำนวน"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label8.Location = New System.Drawing.Point(242, 186)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(136, 22)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "มูลค่า"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label7.Location = New System.Drawing.Point(106, 186)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(137, 22)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "หัวคูปอง"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(191, 25)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(721, 20)
        Me.TextBox2.TabIndex = 20
        Me.TextBox2.Text = "เลือกรายการคูปองในตาราง กรอกจำนวนที่ต้องการ ภายในตารางก็มีจำนวนคงเหลือ และมี Chec" & _
            "k Box ให้เลือกด้วยว่าเอารายไหนด้วย"
        '
        'ListViewGenCoupon
        '
        Me.ListViewGenCoupon.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader7, Me.ColumnHeader8, Me.ColumnHeader9, Me.ColumnHeader10})
        Me.ListViewGenCoupon.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ListViewGenCoupon.FullRowSelect = True
        Me.ListViewGenCoupon.GridLines = True
        Me.ListViewGenCoupon.Location = New System.Drawing.Point(107, 396)
        Me.ListViewGenCoupon.Name = "ListViewGenCoupon"
        Me.ListViewGenCoupon.Size = New System.Drawing.Size(805, 165)
        Me.ListViewGenCoupon.TabIndex = 21
        Me.ListViewGenCoupon.UseCompatibleStateImageBehavior = False
        Me.ListViewGenCoupon.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "เลขที่คูปอง"
        Me.ColumnHeader7.Width = 150
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "จากวันที่"
        Me.ColumnHeader8.Width = 150
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "ถึงวันที่"
        Me.ColumnHeader9.Width = 150
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "มูลค่าคูปอง"
        Me.ColumnHeader10.Width = 150
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(107, 370)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(805, 20)
        Me.TextBox3.TabIndex = 22
        Me.TextBox3.Text = "รายการที่ Gen เป็นคูปองเลย เพื่อเอาไว้ตรวจสอบข้อมูลก็พอ"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(104, 63)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 13)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "เลขที่ใบขอเบิก :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(323, 63)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 13)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "วันที่ขอเบิก :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(134, 91)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(51, 13)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "ผู้ขอเบิก :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(91, 123)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(94, 13)
        Me.Label5.TabIndex = 26
        Me.Label5.Text = "เหตุผลการขอเบิก :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(543, 95)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(51, 13)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "ชื่อคูปอง :"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(311, 95)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(77, 13)
        Me.Label11.TabIndex = 28
        Me.Label11.Text = "ทะเบียนคูปอง :"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextCPCode
        '
        Me.TextCPCode.Location = New System.Drawing.Point(394, 92)
        Me.TextCPCode.Name = "TextCPCode"
        Me.TextCPCode.Size = New System.Drawing.Size(100, 20)
        Me.TextCPCode.TabIndex = 29
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = Global.NPWindowApp.My.Resources.Resources.Cancel
        Me.PictureBox3.Location = New System.Drawing.Point(137, 24)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(39, 21)
        Me.PictureBox3.TabIndex = 14
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.NPWindowApp.My.Resources.Resources.Confirm
        Me.PictureBox2.Location = New System.Drawing.Point(92, 24)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(39, 22)
        Me.PictureBox2.TabIndex = 13
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.ErrorImage = CType(resources.GetObject("PictureBox1.ErrorImage"), System.Drawing.Image)
        Me.PictureBox1.Image = Global.NPWindowApp.My.Resources.Resources._New
        Me.PictureBox1.InitialImage = CType(resources.GetObject("PictureBox1.InitialImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(47, 23)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(39, 22)
        Me.PictureBox1.TabIndex = 12
        Me.PictureBox1.TabStop = False
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label12.Location = New System.Drawing.Point(646, 186)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(132, 22)
        Me.Label12.TabIndex = 30
        Me.Label12.Text = "คงเหลือ"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.DodgerBlue
        Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label13.Location = New System.Drawing.Point(777, 186)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(132, 22)
        Me.Label13.TabIndex = 31
        Me.Label13.Text = "ขอเบิก"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DTPDocDate
        '
        Me.DTPDocDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPDocDate.Location = New System.Drawing.Point(394, 60)
        Me.DTPDocDate.Name = "DTPDocDate"
        Me.DTPDocDate.Size = New System.Drawing.Size(100, 20)
        Me.DTPDocDate.TabIndex = 32
        '
        'CMBReqUserID
        '
        Me.CMBReqUserID.FormattingEnabled = True
        Me.CMBReqUserID.Location = New System.Drawing.Point(191, 88)
        Me.CMBReqUserID.Name = "CMBReqUserID"
        Me.CMBReqUserID.Size = New System.Drawing.Size(100, 21)
        Me.CMBReqUserID.TabIndex = 33
        '
        'GB101
        '
        Me.GB101.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.GB101.Controls.Add(Me.TextQTY)
        Me.GB101.Controls.Add(Me.Label14)
        Me.GB101.Location = New System.Drawing.Point(105, 182)
        Me.GB101.Name = "GB101"
        Me.GB101.Size = New System.Drawing.Size(807, 182)
        Me.GB101.TabIndex = 34
        Me.GB101.TabStop = False
        Me.GB101.Visible = False
        '
        'TextQTY
        '
        Me.TextQTY.BackColor = System.Drawing.Color.Black
        Me.TextQTY.Font = New System.Drawing.Font("Microsoft Sans Serif", 30.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TextQTY.ForeColor = System.Drawing.Color.Gold
        Me.TextQTY.Location = New System.Drawing.Point(389, 73)
        Me.TextQTY.Name = "TextQTY"
        Me.TextQTY.Size = New System.Drawing.Size(168, 53)
        Me.TextQTY.TabIndex = 1
        Me.TextQTY.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label14.Location = New System.Drawing.Point(258, 88)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(125, 16)
        Me.Label14.TabIndex = 0
        Me.Label14.Text = "จำนวนคูปองที่ขอเบิก :"
        '
        'BTNGenDocNo
        '
        Me.BTNGenDocNo.Location = New System.Drawing.Point(291, 60)
        Me.BTNGenDocNo.Name = "BTNGenDocNo"
        Me.BTNGenDocNo.Size = New System.Drawing.Size(20, 19)
        Me.BTNGenDocNo.TabIndex = 35
        Me.BTNGenDocNo.Text = "Button1"
        Me.BTNGenDocNo.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(496, 92)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(20, 19)
        Me.Button1.TabIndex = 36
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FormCouponRequest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDark
        Me.ClientSize = New System.Drawing.Size(1016, 734)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.BTNGenDocNo)
        Me.Controls.Add(Me.GB101)
        Me.Controls.Add(Me.CMBReqUserID)
        Me.Controls.Add(Me.DTPDocDate)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.TextCPCode)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.ListViewGenCoupon)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TextReqReason)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.BTNSearch)
        Me.Controls.Add(Me.BTNCancel)
        Me.Controls.Add(Me.BTNSave)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextCPValue)
        Me.Controls.Add(Me.TextDocNo)
        Me.Controls.Add(Me.ListViewReqCP)
        Me.Controls.Add(Me.TextCPName)
        Me.Name = "FormCouponRequest"
        Me.Text = "ฟอร์ม ทำใบขอเบิกคูปอง"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GB101.ResumeLayout(False)
        Me.GB101.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextCPName As System.Windows.Forms.TextBox
    Friend WithEvents ListViewReqCP As System.Windows.Forms.ListView
    Friend WithEvents TextDocNo As System.Windows.Forms.TextBox
    Friend WithEvents TextCPValue As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents BTNCancel As System.Windows.Forms.Button
    Friend WithEvents BTNSearch As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents TextReqReason As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ListViewGenCoupon As System.Windows.Forms.ListView
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TextCPCode As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents DTPDocDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents CMBReqUserID As System.Windows.Forms.ComboBox
    Friend WithEvents GB101 As System.Windows.Forms.GroupBox
    Friend WithEvents TextQTY As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents BTNGenDocNo As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
