<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNopadolTransDataToBranch
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNopadolTransDataToBranch))
        Me.BTNCloseProgram = New System.Windows.Forms.Button
        Me.BTNMinimize = New System.Windows.Forms.Button
        Me.PBActive = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.PBNotConnect = New System.Windows.Forms.PictureBox
        Me.TConnect = New System.Windows.Forms.Timer(Me.components)
        Me.Label1 = New System.Windows.Forms.Label
        Me.TMNotTransfer = New System.Windows.Forms.Timer(Me.components)
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.LBLTime = New System.Windows.Forms.Label
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.TMActive = New System.Windows.Forms.Timer(Me.components)
        Me.TBLink = New System.Windows.Forms.TextBox
        Me.Time = New System.Windows.Forms.Timer(Me.components)
        Me.TMTransfer = New System.Windows.Forms.Timer(Me.components)
        Me.TCheckConnect = New System.Windows.Forms.Timer(Me.components)
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label3 = New System.Windows.Forms.Label
        Me.ListViewListTrn = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        CType(Me.PBActive, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        CType(Me.PBNotConnect, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BTNCloseProgram
        '
        Me.BTNCloseProgram.BackColor = System.Drawing.Color.Silver
        Me.BTNCloseProgram.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.BTNCloseProgram.Location = New System.Drawing.Point(283, 310)
        Me.BTNCloseProgram.Name = "BTNCloseProgram"
        Me.BTNCloseProgram.Size = New System.Drawing.Size(99, 35)
        Me.BTNCloseProgram.TabIndex = 54
        Me.BTNCloseProgram.Text = "ปิดโปรแกรม"
        Me.BTNCloseProgram.UseVisualStyleBackColor = False
        '
        'BTNMinimize
        '
        Me.BTNMinimize.BackColor = System.Drawing.Color.Silver
        Me.BTNMinimize.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.BTNMinimize.Location = New System.Drawing.Point(177, 310)
        Me.BTNMinimize.Name = "BTNMinimize"
        Me.BTNMinimize.Size = New System.Drawing.Size(99, 35)
        Me.BTNMinimize.TabIndex = 53
        Me.BTNMinimize.Text = "ซ่อนโปรแกรม"
        Me.BTNMinimize.UseVisualStyleBackColor = False
        '
        'PBActive
        '
        Me.PBActive.Location = New System.Drawing.Point(12, 44)
        Me.PBActive.Name = "PBActive"
        Me.PBActive.Size = New System.Drawing.Size(39, 37)
        Me.PBActive.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PBActive.TabIndex = 51
        Me.PBActive.TabStop = False
        Me.PBActive.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(74, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(230, 20)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "โปรแกรม ระบบโอนเอกสารทั่วไป"
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.Black
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Location = New System.Drawing.Point(12, 351)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(370, 5)
        Me.Panel4.TabIndex = 48
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Blue
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Panel3.Location = New System.Drawing.Point(12, 6)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(370, 32)
        Me.Panel3.TabIndex = 52
        '
        'PBNotConnect
        '
        Me.PBNotConnect.Location = New System.Drawing.Point(18, 51)
        Me.PBNotConnect.Name = "PBNotConnect"
        Me.PBNotConnect.Size = New System.Drawing.Size(25, 26)
        Me.PBNotConnect.TabIndex = 55
        Me.PBNotConnect.TabStop = False
        Me.PBNotConnect.Visible = False
        '
        'TConnect
        '
        Me.TConnect.Enabled = True
        Me.TConnect.Interval = 3133
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(284, 375)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(99, 13)
        Me.Label1.TabIndex = 50
        Me.Label1.Text = "Version.2014.06.25"
        '
        'TMNotTransfer
        '
        Me.TMNotTransfer.Enabled = True
        Me.TMNotTransfer.Interval = 51373
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Sea Agent Transfer"
        Me.NotifyIcon1.Visible = True
        '
        'LBLTime
        '
        Me.LBLTime.AutoSize = True
        Me.LBLTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.LBLTime.Location = New System.Drawing.Point(329, 88)
        Me.LBLTime.Name = "LBLTime"
        Me.LBLTime.Size = New System.Drawing.Size(53, 16)
        Me.LBLTime.TabIndex = 49
        Me.LBLTime.Text = "HH:MM"
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "ข้อมูลที่โอน"
        Me.ColumnHeader3.Width = 150
        '
        'TMActive
        '
        Me.TMActive.Enabled = True
        Me.TMActive.Interval = 10000
        '
        'TBLink
        '
        Me.TBLink.BackColor = System.Drawing.Color.Orange
        Me.TBLink.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBLink.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TBLink.Location = New System.Drawing.Point(57, 48)
        Me.TBLink.Name = "TBLink"
        Me.TBLink.ReadOnly = True
        Me.TBLink.Size = New System.Drawing.Size(325, 29)
        Me.TBLink.TabIndex = 47
        Me.TBLink.Text = "ติดต่อสาขาได้"
        Me.TBLink.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Time
        '
        Me.Time.Interval = 1500
        '
        'TMTransfer
        '
        Me.TMTransfer.Enabled = True
        Me.TMTransfer.Interval = 71731
        '
        'TCheckConnect
        '
        Me.TCheckConnect.Enabled = True
        Me.TCheckConnect.Interval = 181373
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Black
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(12, 299)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(370, 5)
        Me.Panel2.TabIndex = 46
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "ชื่อตาราง"
        Me.ColumnHeader2.Width = 150
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Black
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Location = New System.Drawing.Point(12, 115)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(370, 5)
        Me.Panel1.TabIndex = 45
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 123)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(140, 16)
        Me.Label3.TabIndex = 44
        Me.Label3.Text = "รายการเอกสารที่โอนวันนี้"
        '
        'ListViewListTrn
        '
        Me.ListViewListTrn.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListViewListTrn.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.ListViewListTrn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ListViewListTrn.FullRowSelect = True
        Me.ListViewListTrn.GridLines = True
        Me.ListViewListTrn.Location = New System.Drawing.Point(12, 142)
        Me.ListViewListTrn.Name = "ListViewListTrn"
        Me.ListViewListTrn.Size = New System.Drawing.Size(370, 142)
        Me.ListViewListTrn.TabIndex = 43
        Me.ListViewListTrn.UseCompatibleStateImageBehavior = False
        Me.ListViewListTrn.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ลำดับ"
        Me.ColumnHeader1.Width = 50
        '
        'frmNopadolTransDataToBranch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(394, 395)
        Me.Controls.Add(Me.BTNCloseProgram)
        Me.Controls.Add(Me.BTNMinimize)
        Me.Controls.Add(Me.PBActive)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.PBNotConnect)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LBLTime)
        Me.Controls.Add(Me.TBLink)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ListViewListTrn)
        Me.Name = "frmNopadolTransDataToBranch"
        Me.Text = "Form1"
        CType(Me.PBActive, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.PBNotConnect, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BTNCloseProgram As System.Windows.Forms.Button
    Friend WithEvents BTNMinimize As System.Windows.Forms.Button
    Friend WithEvents PBActive As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents PBNotConnect As System.Windows.Forms.PictureBox
    Friend WithEvents TConnect As System.Windows.Forms.Timer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TMNotTransfer As System.Windows.Forms.Timer
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents LBLTime As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents TMActive As System.Windows.Forms.Timer
    Friend WithEvents TBLink As System.Windows.Forms.TextBox
    Friend WithEvents Time As System.Windows.Forms.Timer
    Friend WithEvents TMTransfer As System.Windows.Forms.Timer
    Friend WithEvents TCheckConnect As System.Windows.Forms.Timer
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ListViewListTrn As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader

End Class
