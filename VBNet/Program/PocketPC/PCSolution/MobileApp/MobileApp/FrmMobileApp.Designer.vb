<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class FrmMobileApp
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMobileApp))
        Me.PICLogIn = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.TBUserID = New System.Windows.Forms.TextBox
        Me.TBPassword = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.PNLogIn = New System.Windows.Forms.Panel
        Me.Panel10 = New System.Windows.Forms.Panel
        Me.Panel9 = New System.Windows.Forms.Panel
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.PNSelectJob = New System.Windows.Forms.Panel
        Me.RBJob3 = New System.Windows.Forms.RadioButton
        Me.Panel8 = New System.Windows.Forms.Panel
        Me.Label5 = New System.Windows.Forms.Label
        Me.Panel7 = New System.Windows.Forms.Panel
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.BTNLogIn = New System.Windows.Forms.Button
        Me.BTNSelectJob = New System.Windows.Forms.Button
        Me.RBJob6 = New System.Windows.Forms.RadioButton
        Me.RBJob5 = New System.Windows.Forms.RadioButton
        Me.RBJob4 = New System.Windows.Forms.RadioButton
        Me.RBJob2 = New System.Windows.Forms.RadioButton
        Me.RBJob1 = New System.Windows.Forms.RadioButton
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PNLogIn.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.PNSelectJob.SuspendLayout()
        Me.Panel8.SuspendLayout()
        Me.SuspendLayout()
        '
        'PICLogIn
        '
        Me.PICLogIn.Image = CType(resources.GetObject("PICLogIn.Image"), System.Drawing.Image)
        Me.PICLogIn.Location = New System.Drawing.Point(0, -23)
        Me.PICLogIn.Name = "PICLogIn"
        Me.PICLogIn.Size = New System.Drawing.Size(240, 270)
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(32, 181)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 17)
        Me.Label1.Text = "ชื่อผู้ใช้งาน :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(46, 213)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 17)
        Me.Label2.Text = "รหัสผ่าน :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBUserID
        '
        Me.TBUserID.Location = New System.Drawing.Point(102, 178)
        Me.TBUserID.Name = "TBUserID"
        Me.TBUserID.Size = New System.Drawing.Size(100, 21)
        Me.TBUserID.TabIndex = 0
        '
        'TBPassword
        '
        Me.TBPassword.Location = New System.Drawing.Point(102, 210)
        Me.TBPassword.Name = "TBPassword"
        Me.TBPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TBPassword.Size = New System.Drawing.Size(100, 21)
        Me.TBPassword.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle))
        Me.Label3.Location = New System.Drawing.Point(3, 269)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(234, 20)
        Me.Label3.Text = "โปรแกรมเสริม Nopadol Expert"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Regular)
        Me.Label4.Location = New System.Drawing.Point(31, 292)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(87, 12)
        Me.Label4.Text = "V.17112009"
        '
        'PNLogIn
        '
        Me.PNLogIn.BackColor = System.Drawing.Color.White
        Me.PNLogIn.Controls.Add(Me.Panel10)
        Me.PNLogIn.Controls.Add(Me.Panel9)
        Me.PNLogIn.Controls.Add(Me.Panel5)
        Me.PNLogIn.Controls.Add(Me.Panel4)
        Me.PNLogIn.Controls.Add(Me.Panel1)
        Me.PNLogIn.Controls.Add(Me.Label4)
        Me.PNLogIn.Controls.Add(Me.Label3)
        Me.PNLogIn.Location = New System.Drawing.Point(0, 0)
        Me.PNLogIn.Name = "PNLogIn"
        Me.PNLogIn.Size = New System.Drawing.Size(240, 320)
        '
        'Panel10
        '
        Me.Panel10.BackColor = System.Drawing.Color.Navy
        Me.Panel10.Location = New System.Drawing.Point(0, 311)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(240, 8)
        '
        'Panel9
        '
        Me.Panel9.BackColor = System.Drawing.Color.Navy
        Me.Panel9.Location = New System.Drawing.Point(0, 248)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(240, 8)
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.DarkOrange
        Me.Panel5.Location = New System.Drawing.Point(0, 303)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(240, 8)
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.DarkOrange
        Me.Panel4.Location = New System.Drawing.Point(0, 256)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(240, 8)
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkGreen
        Me.Panel1.Controls.Add(Me.TBPassword)
        Me.Panel1.Controls.Add(Me.TBUserID)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.PICLogIn)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(240, 248)
        '
        'PNSelectJob
        '
        Me.PNSelectJob.Controls.Add(Me.RBJob3)
        Me.PNSelectJob.Controls.Add(Me.Panel8)
        Me.PNSelectJob.Controls.Add(Me.Panel7)
        Me.PNSelectJob.Controls.Add(Me.Panel6)
        Me.PNSelectJob.Controls.Add(Me.BTNLogIn)
        Me.PNSelectJob.Controls.Add(Me.BTNSelectJob)
        Me.PNSelectJob.Controls.Add(Me.RBJob6)
        Me.PNSelectJob.Controls.Add(Me.RBJob5)
        Me.PNSelectJob.Controls.Add(Me.RBJob4)
        Me.PNSelectJob.Controls.Add(Me.RBJob2)
        Me.PNSelectJob.Controls.Add(Me.RBJob1)
        Me.PNSelectJob.Controls.Add(Me.Panel3)
        Me.PNSelectJob.Controls.Add(Me.Panel2)
        Me.PNSelectJob.Controls.Add(Me.PictureBox2)
        Me.PNSelectJob.Location = New System.Drawing.Point(0, -2)
        Me.PNSelectJob.Name = "PNSelectJob"
        Me.PNSelectJob.Size = New System.Drawing.Size(240, 322)
        Me.PNSelectJob.Visible = False
        '
        'RBJob3
        '
        Me.RBJob3.ForeColor = System.Drawing.Color.Red
        Me.RBJob3.Location = New System.Drawing.Point(31, 121)
        Me.RBJob3.Name = "RBJob3"
        Me.RBJob3.Size = New System.Drawing.Size(194, 20)
        Me.RBJob3.TabIndex = 7
        Me.RBJob3.TabStop = False
        Me.RBJob3.Text = "3. นับสต็อกสินค้า ตามคลัง"
        '
        'Panel8
        '
        Me.Panel8.BackColor = System.Drawing.Color.Navy
        Me.Panel8.Controls.Add(Me.Label5)
        Me.Panel8.Location = New System.Drawing.Point(-1, 3)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(241, 48)
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 10.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle))
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(16, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(210, 20)
        Me.Label5.Text = "เลือก โปรแกรมใช้งาน"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Panel7
        '
        Me.Panel7.BackColor = System.Drawing.Color.Navy
        Me.Panel7.Location = New System.Drawing.Point(-13, 246)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(267, 6)
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.Color.DarkOrange
        Me.Panel6.Location = New System.Drawing.Point(-13, 52)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(267, 6)
        '
        'BTNLogIn
        '
        Me.BTNLogIn.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNLogIn.Location = New System.Drawing.Point(182, 220)
        Me.BTNLogIn.Name = "BTNLogIn"
        Me.BTNLogIn.Size = New System.Drawing.Size(43, 19)
        Me.BTNLogIn.TabIndex = 12
        Me.BTNLogIn.Text = "LogIn"
        '
        'BTNSelectJob
        '
        Me.BTNSelectJob.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNSelectJob.Location = New System.Drawing.Point(133, 220)
        Me.BTNSelectJob.Name = "BTNSelectJob"
        Me.BTNSelectJob.Size = New System.Drawing.Size(43, 19)
        Me.BTNSelectJob.TabIndex = 11
        Me.BTNSelectJob.Text = "เลือก"
        '
        'RBJob6
        '
        Me.RBJob6.Enabled = False
        Me.RBJob6.ForeColor = System.Drawing.Color.Red
        Me.RBJob6.Location = New System.Drawing.Point(31, 197)
        Me.RBJob6.Name = "RBJob6"
        Me.RBJob6.Size = New System.Drawing.Size(194, 20)
        Me.RBJob6.TabIndex = 10
        Me.RBJob6.TabStop = False
        Me.RBJob6.Text = "5. เสนอซื้อสินค้า"
        Me.RBJob6.Visible = False
        '
        'RBJob5
        '
        Me.RBJob5.Enabled = False
        Me.RBJob5.ForeColor = System.Drawing.Color.Red
        Me.RBJob5.Location = New System.Drawing.Point(31, 180)
        Me.RBJob5.Name = "RBJob5"
        Me.RBJob5.Size = New System.Drawing.Size(194, 20)
        Me.RBJob5.TabIndex = 9
        Me.RBJob5.TabStop = False
        Me.RBJob5.Text = "4. ขอเบิก/โอน สินค้า"
        Me.RBJob5.Visible = False
        '
        'RBJob4
        '
        Me.RBJob4.ForeColor = System.Drawing.Color.Red
        Me.RBJob4.Location = New System.Drawing.Point(31, 147)
        Me.RBJob4.Name = "RBJob4"
        Me.RBJob4.Size = New System.Drawing.Size(194, 20)
        Me.RBJob4.TabIndex = 8
        Me.RBJob4.TabStop = False
        Me.RBJob4.Text = "4. เลือกสินค้าพิมพ์บาร์โค้ด"
        '
        'RBJob2
        '
        Me.RBJob2.ForeColor = System.Drawing.Color.Red
        Me.RBJob2.Location = New System.Drawing.Point(31, 94)
        Me.RBJob2.Name = "RBJob2"
        Me.RBJob2.Size = New System.Drawing.Size(194, 20)
        Me.RBJob2.TabIndex = 6
        Me.RBJob2.TabStop = False
        Me.RBJob2.Text = "2. บันทึกที่เก็บสินค้า"
        '
        'RBJob1
        '
        Me.RBJob1.Checked = True
        Me.RBJob1.ForeColor = System.Drawing.Color.Red
        Me.RBJob1.Location = New System.Drawing.Point(31, 68)
        Me.RBJob1.Name = "RBJob1"
        Me.RBJob1.Size = New System.Drawing.Size(194, 20)
        Me.RBJob1.TabIndex = 5
        Me.RBJob1.Text = "1. ตรวจสอบข้อมูลสินค้า"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkOrange
        Me.Panel3.Location = New System.Drawing.Point(-10, 252)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(267, 6)
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Navy
        Me.Panel2.Location = New System.Drawing.Point(-10, 58)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(267, 6)
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(-1, 258)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(244, 75)
        '
        'FrmMobileApp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(240, 320)
        Me.Controls.Add(Me.PNSelectJob)
        Me.Controls.Add(Me.PNLogIn)
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "FrmMobileApp"
        Me.Text = "โปรแกรมเสริม"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.PNLogIn.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.PNSelectJob.ResumeLayout(False)
        Me.Panel8.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PICLogIn As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TBUserID As System.Windows.Forms.TextBox
    Friend WithEvents TBPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents PNLogIn As System.Windows.Forms.Panel
    Friend WithEvents PNSelectJob As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents RBJob6 As System.Windows.Forms.RadioButton
    Friend WithEvents RBJob5 As System.Windows.Forms.RadioButton
    Friend WithEvents RBJob4 As System.Windows.Forms.RadioButton
    Friend WithEvents RBJob2 As System.Windows.Forms.RadioButton
    Friend WithEvents RBJob1 As System.Windows.Forms.RadioButton
    Friend WithEvents BTNSelectJob As System.Windows.Forms.Button
    Friend WithEvents BTNLogIn As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
    Friend WithEvents RBJob3 As System.Windows.Forms.RadioButton
    Friend WithEvents Panel10 As System.Windows.Forms.Panel
    Friend WithEvents Panel9 As System.Windows.Forms.Panel

End Class
