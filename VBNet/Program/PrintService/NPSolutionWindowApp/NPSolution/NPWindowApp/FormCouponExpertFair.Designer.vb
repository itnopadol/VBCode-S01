<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormCouponExpertFair
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
        Me.BTNSave = New System.Windows.Forms.Button
        Me.TBMember = New System.Windows.Forms.TextBox
        Me.MEID = New System.Windows.Forms.MaskedTextBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.TBCouponAmount = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.BTNSearchID = New System.Windows.Forms.Button
        Me.PNSearch = New System.Windows.Forms.Panel
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.BTNClose = New System.Windows.Forms.Button
        Me.BTNSearch = New System.Windows.Forms.Button
        Me.TBSearch = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.ListViewSearch = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.BTNMember = New System.Windows.Forms.Button
        Me.PNMember = New System.Windows.Forms.Panel
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.Panel7 = New System.Windows.Forms.Panel
        Me.BTNMemberClose = New System.Windows.Forms.Button
        Me.BTNSearchMember = New System.Windows.Forms.Button
        Me.TBSearchMember = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.ListViewMember = New System.Windows.Forms.ListView
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Panel8 = New System.Windows.Forms.Panel
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PNSearch.SuspendLayout()
        Me.PNMember.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("PSL KandaAD", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label1.Location = New System.Drawing.Point(582, 82)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(256, 34)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "กรอกเลขที่บัตรประจำตัวประชาชน :"
        '
        'BTNSave
        '
        Me.BTNSave.BackColor = System.Drawing.Color.LightGray
        Me.BTNSave.Font = New System.Drawing.Font("PSL KandaAD", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.BTNSave.Location = New System.Drawing.Point(903, 259)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(121, 57)
        Me.BTNSave.TabIndex = 3
        Me.BTNSave.Text = "บันทึกข้อมูล"
        Me.BTNSave.UseVisualStyleBackColor = False
        '
        'TBMember
        '
        Me.TBMember.BackColor = System.Drawing.SystemColors.Info
        Me.TBMember.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBMember.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBMember.Location = New System.Drawing.Point(844, 136)
        Me.TBMember.Name = "TBMember"
        Me.TBMember.Size = New System.Drawing.Size(250, 38)
        Me.TBMember.TabIndex = 1
        '
        'MEID
        '
        Me.MEID.BackColor = System.Drawing.SystemColors.Info
        Me.MEID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MEID.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.MEID.Location = New System.Drawing.Point(844, 82)
        Me.MEID.Mask = "#-####-#####-##-#"
        Me.MEID.Name = "MEID"
        Me.MEID.Size = New System.Drawing.Size(250, 38)
        Me.MEID.TabIndex = 0
        Me.MEID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.NPWindowApp.My.Resources.Resources.ช่างชิงแชมป์
        Me.PictureBox2.Location = New System.Drawing.Point(37, 293)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(478, 346)
        Me.PictureBox2.TabIndex = 5
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.NPWindowApp.My.Resources.Resources.Expert_Fair
        Me.PictureBox1.Location = New System.Drawing.Point(-320, -25)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(835, 287)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.MediumBlue
        Me.Panel1.Location = New System.Drawing.Point(-2, 245)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(517, 8)
        Me.Panel1.TabIndex = 6
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Orange
        Me.Panel2.Location = New System.Drawing.Point(-2, 254)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(517, 8)
        Me.Panel2.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("PSL KandaAD", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label2.Location = New System.Drawing.Point(691, 136)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(147, 34)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "กรอกเลขที่สมาชิก :"
        '
        'TBCouponAmount
        '
        Me.TBCouponAmount.BackColor = System.Drawing.SystemColors.Info
        Me.TBCouponAmount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBCouponAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBCouponAmount.Location = New System.Drawing.Point(844, 194)
        Me.TBCouponAmount.Name = "TBCouponAmount"
        Me.TBCouponAmount.Size = New System.Drawing.Size(180, 38)
        Me.TBCouponAmount.TabIndex = 2
        Me.TBCouponAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("PSL KandaAD", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label4.Location = New System.Drawing.Point(702, 194)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 34)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "มูลค่ารวมคูปอง :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("PSL KandaAD", 20.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label3.Location = New System.Drawing.Point(521, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(326, 34)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "บันทึกข้อมูลการจ่ายคูปองอาหารและเครื่องดื่ม"
        '
        'BTNSearchID
        '
        Me.BTNSearchID.BackColor = System.Drawing.Color.LightGray
        Me.BTNSearchID.Font = New System.Drawing.Font("PSL KandaAD", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.BTNSearchID.Image = Global.NPWindowApp.My.Resources.Resources._2
        Me.BTNSearchID.Location = New System.Drawing.Point(1100, 82)
        Me.BTNSearchID.Name = "BTNSearchID"
        Me.BTNSearchID.Size = New System.Drawing.Size(41, 38)
        Me.BTNSearchID.TabIndex = 13
        Me.BTNSearchID.UseVisualStyleBackColor = False
        '
        'PNSearch
        '
        Me.PNSearch.BackColor = System.Drawing.Color.Gainsboro
        Me.PNSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PNSearch.Controls.Add(Me.Panel4)
        Me.PNSearch.Controls.Add(Me.Panel3)
        Me.PNSearch.Controls.Add(Me.BTNClose)
        Me.PNSearch.Controls.Add(Me.BTNSearch)
        Me.PNSearch.Controls.Add(Me.TBSearch)
        Me.PNSearch.Controls.Add(Me.Label5)
        Me.PNSearch.Controls.Add(Me.ListViewSearch)
        Me.PNSearch.Location = New System.Drawing.Point(525, 46)
        Me.PNSearch.Name = "PNSearch"
        Me.PNSearch.Size = New System.Drawing.Size(673, 593)
        Me.PNSearch.TabIndex = 14
        Me.PNSearch.Visible = False
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.Orange
        Me.Panel4.Location = New System.Drawing.Point(3, 49)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(665, 14)
        Me.Panel4.TabIndex = 8
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.MediumBlue
        Me.Panel3.Location = New System.Drawing.Point(3, 488)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(665, 14)
        Me.Panel3.TabIndex = 7
        '
        'BTNClose
        '
        Me.BTNClose.BackColor = System.Drawing.Color.Silver
        Me.BTNClose.Font = New System.Drawing.Font("PSL Kanda ExtraAD", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.BTNClose.Location = New System.Drawing.Point(561, 505)
        Me.BTNClose.Name = "BTNClose"
        Me.BTNClose.Size = New System.Drawing.Size(107, 59)
        Me.BTNClose.TabIndex = 5
        Me.BTNClose.Text = "ปิด"
        Me.BTNClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNClose.UseVisualStyleBackColor = False
        '
        'BTNSearch
        '
        Me.BTNSearch.BackColor = System.Drawing.Color.Gainsboro
        Me.BTNSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.BTNSearch.Image = Global.NPWindowApp.My.Resources.Resources._4
        Me.BTNSearch.Location = New System.Drawing.Point(422, 18)
        Me.BTNSearch.Name = "BTNSearch"
        Me.BTNSearch.Size = New System.Drawing.Size(27, 22)
        Me.BTNSearch.TabIndex = 3
        Me.BTNSearch.UseVisualStyleBackColor = False
        '
        'TBSearch
        '
        Me.TBSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBSearch.Location = New System.Drawing.Point(79, 18)
        Me.TBSearch.Name = "TBSearch"
        Me.TBSearch.Size = New System.Drawing.Size(342, 22)
        Me.TBSearch.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("PSL Kanda ExtraAD", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label5.Location = New System.Drawing.Point(3, 13)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 27)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "คำที่ค้นหา :"
        '
        'ListViewSearch
        '
        Me.ListViewSearch.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.ListViewSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ListViewSearch.FullRowSelect = True
        Me.ListViewSearch.GridLines = True
        Me.ListViewSearch.Location = New System.Drawing.Point(3, 64)
        Me.ListViewSearch.Name = "ListViewSearch"
        Me.ListViewSearch.Size = New System.Drawing.Size(665, 417)
        Me.ListViewSearch.TabIndex = 0
        Me.ListViewSearch.UseCompatibleStateImageBehavior = False
        Me.ListViewSearch.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ลำดับ"
        Me.ColumnHeader1.Width = 50
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "เลขที่บัตรประชาชน"
        Me.ColumnHeader2.Width = 120
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "รหัสสมาชิก"
        Me.ColumnHeader3.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "ชื่อสมาชิก"
        Me.ColumnHeader4.Width = 250
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "มูลค่าคูปอง"
        Me.ColumnHeader5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader5.Width = 80
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "ครั้งที่"
        '
        'BTNMember
        '
        Me.BTNMember.BackColor = System.Drawing.Color.LightGray
        Me.BTNMember.Font = New System.Drawing.Font("PSL KandaAD", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.BTNMember.Image = Global.NPWindowApp.My.Resources.Resources._4
        Me.BTNMember.Location = New System.Drawing.Point(1100, 136)
        Me.BTNMember.Name = "BTNMember"
        Me.BTNMember.Size = New System.Drawing.Size(41, 38)
        Me.BTNMember.TabIndex = 15
        Me.BTNMember.UseVisualStyleBackColor = False
        '
        'PNMember
        '
        Me.PNMember.BackColor = System.Drawing.Color.Gainsboro
        Me.PNMember.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PNMember.Controls.Add(Me.Panel6)
        Me.PNMember.Controls.Add(Me.Panel7)
        Me.PNMember.Controls.Add(Me.BTNMemberClose)
        Me.PNMember.Controls.Add(Me.BTNSearchMember)
        Me.PNMember.Controls.Add(Me.TBSearchMember)
        Me.PNMember.Controls.Add(Me.Label6)
        Me.PNMember.Controls.Add(Me.ListViewMember)
        Me.PNMember.Location = New System.Drawing.Point(525, 46)
        Me.PNMember.Name = "PNMember"
        Me.PNMember.Size = New System.Drawing.Size(673, 623)
        Me.PNMember.TabIndex = 16
        Me.PNMember.Visible = False
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.Color.Orange
        Me.Panel6.Location = New System.Drawing.Point(3, 49)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(665, 14)
        Me.Panel6.TabIndex = 8
        '
        'Panel7
        '
        Me.Panel7.BackColor = System.Drawing.Color.MediumBlue
        Me.Panel7.Location = New System.Drawing.Point(3, 488)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(665, 14)
        Me.Panel7.TabIndex = 7
        '
        'BTNMemberClose
        '
        Me.BTNMemberClose.BackColor = System.Drawing.Color.Silver
        Me.BTNMemberClose.Font = New System.Drawing.Font("PSL Kanda ExtraAD", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.BTNMemberClose.Location = New System.Drawing.Point(561, 505)
        Me.BTNMemberClose.Name = "BTNMemberClose"
        Me.BTNMemberClose.Size = New System.Drawing.Size(107, 59)
        Me.BTNMemberClose.TabIndex = 5
        Me.BTNMemberClose.Text = "ปิด"
        Me.BTNMemberClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNMemberClose.UseVisualStyleBackColor = False
        '
        'BTNSearchMember
        '
        Me.BTNSearchMember.BackColor = System.Drawing.Color.Gainsboro
        Me.BTNSearchMember.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.BTNSearchMember.Image = Global.NPWindowApp.My.Resources.Resources._4
        Me.BTNSearchMember.Location = New System.Drawing.Point(422, 18)
        Me.BTNSearchMember.Name = "BTNSearchMember"
        Me.BTNSearchMember.Size = New System.Drawing.Size(27, 22)
        Me.BTNSearchMember.TabIndex = 3
        Me.BTNSearchMember.UseVisualStyleBackColor = False
        '
        'TBSearchMember
        '
        Me.TBSearchMember.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBSearchMember.Location = New System.Drawing.Point(79, 18)
        Me.TBSearchMember.Name = "TBSearchMember"
        Me.TBSearchMember.Size = New System.Drawing.Size(342, 22)
        Me.TBSearchMember.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("PSL Kanda ExtraAD", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Label6.Location = New System.Drawing.Point(3, 13)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(76, 27)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "คำที่ค้นหา :"
        '
        'ListViewMember
        '
        Me.ListViewMember.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader7, Me.ColumnHeader9, Me.ColumnHeader10})
        Me.ListViewMember.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ListViewMember.FullRowSelect = True
        Me.ListViewMember.GridLines = True
        Me.ListViewMember.Location = New System.Drawing.Point(3, 68)
        Me.ListViewMember.Name = "ListViewMember"
        Me.ListViewMember.Size = New System.Drawing.Size(665, 417)
        Me.ListViewMember.TabIndex = 0
        Me.ListViewMember.UseCompatibleStateImageBehavior = False
        Me.ListViewMember.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "ลำดับ"
        Me.ColumnHeader7.Width = 50
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "รหัสสมาชิก"
        Me.ColumnHeader9.Width = 150
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "ชื่อสมาชิก"
        Me.ColumnHeader10.Width = 450
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.MediumBlue
        Me.Panel5.Location = New System.Drawing.Point(515, -1)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(7, 760)
        Me.Panel5.TabIndex = 17
        '
        'Panel8
        '
        Me.Panel8.BackColor = System.Drawing.Color.MediumBlue
        Me.Panel8.Location = New System.Drawing.Point(1201, -1)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(7, 760)
        Me.Panel8.TabIndex = 18
        '
        'FormCouponExpertFair
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1212, 726)
        Me.Controls.Add(Me.PNMember)
        Me.Controls.Add(Me.Panel8)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.PNSearch)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TBCouponAmount)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.MEID)
        Me.Controls.Add(Me.TBMember)
        Me.Controls.Add(Me.BTNSave)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.BTNSearchID)
        Me.Controls.Add(Me.BTNMember)
        Me.Name = "FormCouponExpertFair"
        Me.Text = "FormCouponExpertFair"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PNSearch.ResumeLayout(False)
        Me.PNSearch.PerformLayout()
        Me.PNMember.ResumeLayout(False)
        Me.PNMember.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents TBMember As System.Windows.Forms.TextBox
    Friend WithEvents MEID As System.Windows.Forms.MaskedTextBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TBCouponAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BTNSearchID As System.Windows.Forms.Button
    Friend WithEvents PNSearch As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ListViewSearch As System.Windows.Forms.ListView
    Friend WithEvents TBSearch As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents BTNSearch As System.Windows.Forms.Button
    Friend WithEvents BTNClose As System.Windows.Forms.Button
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents BTNMember As System.Windows.Forms.Button
    Friend WithEvents PNMember As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents BTNMemberClose As System.Windows.Forms.Button
    Friend WithEvents BTNSearchMember As System.Windows.Forms.Button
    Friend WithEvents TBSearchMember As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ListViewMember As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
End Class
