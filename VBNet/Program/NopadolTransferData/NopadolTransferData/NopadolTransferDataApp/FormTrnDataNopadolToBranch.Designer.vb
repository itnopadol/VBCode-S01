<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormTrnDataNopadolToBranch
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
        Me.components = New System.ComponentModel.Container
        Me.TMTransfer = New System.Windows.Forms.Timer(Me.components)
        Me.ListViewListTrn = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.Label3 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.TBLink = New System.Windows.Forms.TextBox
        Me.LBLTime = New System.Windows.Forms.Label
        Me.Time = New System.Windows.Forms.Timer(Me.components)
        Me.TMActive = New System.Windows.Forms.Timer(Me.components)
        Me.PBActive = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PBGood = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.PBActive, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PBGood, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TMTransfer
        '
        Me.TMTransfer.Enabled = True
        Me.TMTransfer.Interval = 93333
        '
        'ListViewListTrn
        '
        Me.ListViewListTrn.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.ListViewListTrn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ListViewListTrn.FullRowSelect = True
        Me.ListViewListTrn.GridLines = True
        Me.ListViewListTrn.Location = New System.Drawing.Point(13, 137)
        Me.ListViewListTrn.Name = "ListViewListTrn"
        Me.ListViewListTrn.Size = New System.Drawing.Size(415, 142)
        Me.ListViewListTrn.TabIndex = 9
        Me.ListViewListTrn.UseCompatibleStateImageBehavior = False
        Me.ListViewListTrn.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ลำดับ"
        Me.ColumnHeader1.Width = 50
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "ชื่อตาราง"
        Me.ColumnHeader2.Width = 180
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "ข้อมูลที่โอน"
        Me.ColumnHeader3.Width = 160
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.Location = New System.Drawing.Point(10, 119)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(140, 16)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "รายการเอกสารที่โอนวันนี้"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Black
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Location = New System.Drawing.Point(13, 111)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(415, 5)
        Me.Panel1.TabIndex = 12
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Black
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(13, 280)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(415, 5)
        Me.Panel2.TabIndex = 13
        '
        'TBLink
        '
        Me.TBLink.BackColor = System.Drawing.Color.White
        Me.TBLink.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBLink.Location = New System.Drawing.Point(13, 77)
        Me.TBLink.Name = "TBLink"
        Me.TBLink.ReadOnly = True
        Me.TBLink.Size = New System.Drawing.Size(260, 29)
        Me.TBLink.TabIndex = 14
        Me.TBLink.Text = "ติดต่อสาขาได้"
        Me.TBLink.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'LBLTime
        '
        Me.LBLTime.AutoSize = True
        Me.LBLTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.LBLTime.Location = New System.Drawing.Point(386, 26)
        Me.LBLTime.Name = "LBLTime"
        Me.LBLTime.Size = New System.Drawing.Size(53, 16)
        Me.LBLTime.TabIndex = 15
        Me.LBLTime.Text = "HH:MM"
        '
        'Time
        '
        Me.Time.Enabled = True
        Me.Time.Interval = 1500
        '
        'TMActive
        '
        Me.TMActive.Interval = 1000
        '
        'PBActive
        '
        Me.PBActive.Image = Global.NopadolTransferDataApp.My.Resources.Resources._4359
        Me.PBActive.Location = New System.Drawing.Point(13, 8)
        Me.PBActive.Name = "PBActive"
        Me.PBActive.Size = New System.Drawing.Size(71, 66)
        Me.PBActive.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PBActive.TabIndex = 17
        Me.PBActive.TabStop = False
        Me.PBActive.Visible = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.NopadolTransferDataApp.My.Resources.Resources._4444
        Me.PictureBox2.Location = New System.Drawing.Point(351, 45)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(77, 64)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox2.TabIndex = 7
        Me.PictureBox2.TabStop = False
        '
        'PBGood
        '
        Me.PBGood.Image = Global.NopadolTransferDataApp.My.Resources.Resources._4481
        Me.PBGood.Location = New System.Drawing.Point(312, 8)
        Me.PBGood.Name = "PBGood"
        Me.PBGood.Size = New System.Drawing.Size(71, 66)
        Me.PBGood.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PBGood.TabIndex = 8
        Me.PBGood.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(181, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 13)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "version.20101103"
        '
        'FormTrnDataNopadolToBranch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(442, 297)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PBActive)
        Me.Controls.Add(Me.LBLTime)
        Me.Controls.Add(Me.TBLink)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ListViewListTrn)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PBGood)
        Me.Name = "FormTrnDataNopadolToBranch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "โปรแกรม โอนข้อมูลจาก สนญ.ไปสาขาสันกำแพง ระหว่างวัน "
        CType(Me.PBActive, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PBGood, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TMTransfer As System.Windows.Forms.Timer
    Friend WithEvents PBGood As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents ListViewListTrn As System.Windows.Forms.ListView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TBLink As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents LBLTime As System.Windows.Forms.Label
    Friend WithEvents Time As System.Windows.Forms.Timer
    Friend WithEvents TMActive As System.Windows.Forms.Timer
    Friend WithEvents PBActive As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
