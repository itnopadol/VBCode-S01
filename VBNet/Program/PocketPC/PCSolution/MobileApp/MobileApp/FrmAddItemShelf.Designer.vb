<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class FrmAddItemShelf
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmAddItemShelf))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.TBZone = New System.Windows.Forms.TextBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.TBShelf = New System.Windows.Forms.TextBox
        Me.TBWHCode = New System.Windows.Forms.TextBox
        Me.TBBarCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.ListViewSelectItem = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.BTNMenu = New System.Windows.Forms.Button
        Me.BTNClear = New System.Windows.Forms.Button
        Me.BTNSave = New System.Windows.Forms.Button
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.BTNCheckShelf = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer
        Me.PNCheckShelf = New System.Windows.Forms.Panel
        Me.BTNCHKClear = New System.Windows.Forms.Button
        Me.BTNCHKExit = New System.Windows.Forms.Button
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Label6 = New System.Windows.Forms.Label
        Me.ListViewCHKItem = New System.Windows.Forms.ListView
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.TBCHKShelf = New System.Windows.Forms.TextBox
        Me.TBCHKWHCode = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.PNCheckShelf.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkBlue
        Me.Panel1.Controls.Add(Me.TBZone)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.TBShelf)
        Me.Panel1.Controls.Add(Me.TBWHCode)
        Me.Panel1.Controls.Add(Me.TBBarCode)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(240, 82)
        '
        'TBZone
        '
        Me.TBZone.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBZone.Location = New System.Drawing.Point(161, 32)
        Me.TBZone.Name = "TBZone"
        Me.TBZone.ReadOnly = True
        Me.TBZone.Size = New System.Drawing.Size(45, 19)
        Me.TBZone.TabIndex = 8
        Me.TBZone.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(172, 2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(65, 33)
        '
        'TBShelf
        '
        Me.TBShelf.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBShelf.Location = New System.Drawing.Point(61, 32)
        Me.TBShelf.Name = "TBShelf"
        Me.TBShelf.Size = New System.Drawing.Size(97, 19)
        Me.TBShelf.TabIndex = 1
        '
        'TBWHCode
        '
        Me.TBWHCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBWHCode.Location = New System.Drawing.Point(61, 10)
        Me.TBWHCode.Name = "TBWHCode"
        Me.TBWHCode.ReadOnly = True
        Me.TBWHCode.Size = New System.Drawing.Size(53, 19)
        Me.TBWHCode.TabIndex = 0
        Me.TBWHCode.Text = "S02"
        '
        'TBBarCode
        '
        Me.TBBarCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBBarCode.Location = New System.Drawing.Point(61, 55)
        Me.TBBarCode.Name = "TBBarCode"
        Me.TBBarCode.Size = New System.Drawing.Size(145, 19)
        Me.TBBarCode.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(6, 57)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 20)
        Me.Label3.Text = "บาร์โค้ด :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(6, 33)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 20)
        Me.Label2.Text = "ที่เก็บ :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(9, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 20)
        Me.Label1.Text = "คลัง :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DarkOrange
        Me.Panel2.Controls.Add(Me.ListViewSelectItem)
        Me.Panel2.Location = New System.Drawing.Point(0, 84)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(240, 195)
        '
        'ListViewSelectItem
        '
        Me.ListViewSelectItem.Columns.Add(Me.ColumnHeader1)
        Me.ListViewSelectItem.Columns.Add(Me.ColumnHeader2)
        Me.ListViewSelectItem.Columns.Add(Me.ColumnHeader3)
        Me.ListViewSelectItem.Columns.Add(Me.ColumnHeader4)
        Me.ListViewSelectItem.Columns.Add(Me.ColumnHeader5)
        Me.ListViewSelectItem.Columns.Add(Me.ColumnHeader6)
        Me.ListViewSelectItem.Columns.Add(Me.ColumnHeader7)
        Me.ListViewSelectItem.Columns.Add(Me.ColumnHeader8)
        Me.ListViewSelectItem.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Regular)
        Me.ListViewSelectItem.FullRowSelect = True
        Me.ListViewSelectItem.Location = New System.Drawing.Point(3, 14)
        Me.ListViewSelectItem.Name = "ListViewSelectItem"
        Me.ListViewSelectItem.Size = New System.Drawing.Size(234, 178)
        Me.ListViewSelectItem.TabIndex = 3
        Me.ListViewSelectItem.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ลำดับ"
        Me.ColumnHeader1.Width = 35
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "ชื่อสินค้า"
        Me.ColumnHeader2.Width = 150
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "รหัสสินค้า"
        Me.ColumnHeader3.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "ชั้นเก็บ"
        Me.ColumnHeader4.Width = 50
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "คลัง"
        Me.ColumnHeader5.Width = 50
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "บาร์"
        Me.ColumnHeader6.Width = 100
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "หน่วย"
        Me.ColumnHeader7.Width = 60
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "โซน"
        Me.ColumnHeader8.Width = 60
        '
        'BTNMenu
        '
        Me.BTNMenu.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNMenu.Location = New System.Drawing.Point(5, 6)
        Me.BTNMenu.Name = "BTNMenu"
        Me.BTNMenu.Size = New System.Drawing.Size(68, 20)
        Me.BTNMenu.TabIndex = 6
        Me.BTNMenu.Text = "เมนู [ส้ม+1]"
        '
        'BTNClear
        '
        Me.BTNClear.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNClear.Location = New System.Drawing.Point(84, 6)
        Me.BTNClear.Name = "BTNClear"
        Me.BTNClear.Size = New System.Drawing.Size(45, 20)
        Me.BTNClear.TabIndex = 5
        Me.BTNClear.Text = "เคลียร์"
        '
        'BTNSave
        '
        Me.BTNSave.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNSave.Location = New System.Drawing.Point(139, 6)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(44, 20)
        Me.BTNSave.TabIndex = 4
        Me.BTNSave.Text = "บันทึก"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkBlue
        Me.Panel3.Controls.Add(Me.BTNCheckShelf)
        Me.Panel3.Controls.Add(Me.BTNMenu)
        Me.Panel3.Controls.Add(Me.BTNSave)
        Me.Panel3.Controls.Add(Me.BTNClear)
        Me.Panel3.Location = New System.Drawing.Point(0, 282)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(240, 38)
        '
        'BTNCheckShelf
        '
        Me.BTNCheckShelf.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNCheckShelf.Location = New System.Drawing.Point(193, 6)
        Me.BTNCheckShelf.Name = "BTNCheckShelf"
        Me.BTNCheckShelf.Size = New System.Drawing.Size(44, 20)
        Me.BTNCheckShelf.TabIndex = 7
        Me.BTNCheckShelf.Text = "ตรวจ"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 3000
        '
        'PNCheckShelf
        '
        Me.PNCheckShelf.BackColor = System.Drawing.Color.DarkOrange
        Me.PNCheckShelf.Controls.Add(Me.BTNCHKClear)
        Me.PNCheckShelf.Controls.Add(Me.BTNCHKExit)
        Me.PNCheckShelf.Controls.Add(Me.Panel4)
        Me.PNCheckShelf.Controls.Add(Me.PictureBox2)
        Me.PNCheckShelf.Controls.Add(Me.TBCHKShelf)
        Me.PNCheckShelf.Controls.Add(Me.TBCHKWHCode)
        Me.PNCheckShelf.Controls.Add(Me.Label4)
        Me.PNCheckShelf.Controls.Add(Me.Label5)
        Me.PNCheckShelf.Location = New System.Drawing.Point(0, 0)
        Me.PNCheckShelf.Name = "PNCheckShelf"
        Me.PNCheckShelf.Size = New System.Drawing.Size(240, 320)
        Me.PNCheckShelf.Visible = False
        '
        'BTNCHKClear
        '
        Me.BTNCHKClear.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNCHKClear.Location = New System.Drawing.Point(135, 288)
        Me.BTNCHKClear.Name = "BTNCHKClear"
        Me.BTNCHKClear.Size = New System.Drawing.Size(48, 20)
        Me.BTNCHKClear.TabIndex = 19
        Me.BTNCHKClear.Text = "เคลียร์"
        '
        'BTNCHKExit
        '
        Me.BTNCHKExit.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNCHKExit.Location = New System.Drawing.Point(189, 288)
        Me.BTNCHKExit.Name = "BTNCHKExit"
        Me.BTNCHKExit.Size = New System.Drawing.Size(48, 20)
        Me.BTNCHKExit.TabIndex = 14
        Me.BTNCHKExit.Text = "ออก"
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.DarkBlue
        Me.Panel4.Controls.Add(Me.Label6)
        Me.Panel4.Controls.Add(Me.ListViewCHKItem)
        Me.Panel4.Location = New System.Drawing.Point(3, 58)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(234, 224)
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(3, 5)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(108, 16)
        Me.Label6.Text = "รายการสินค้า"
        '
        'ListViewCHKItem
        '
        Me.ListViewCHKItem.BackColor = System.Drawing.Color.White
        Me.ListViewCHKItem.Columns.Add(Me.ColumnHeader9)
        Me.ListViewCHKItem.Columns.Add(Me.ColumnHeader10)
        Me.ListViewCHKItem.Columns.Add(Me.ColumnHeader11)
        Me.ListViewCHKItem.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Regular)
        Me.ListViewCHKItem.FullRowSelect = True
        Me.ListViewCHKItem.Location = New System.Drawing.Point(3, 24)
        Me.ListViewCHKItem.Name = "ListViewCHKItem"
        Me.ListViewCHKItem.Size = New System.Drawing.Size(228, 194)
        Me.ListViewCHKItem.TabIndex = 12
        Me.ListViewCHKItem.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "ลำดับ"
        Me.ColumnHeader9.Width = 35
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "ชื่อสินค้า"
        Me.ColumnHeader10.Width = 200
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "รหัสสินค้า"
        Me.ColumnHeader11.Width = 80
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(172, 3)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(65, 29)
        '
        'TBCHKShelf
        '
        Me.TBCHKShelf.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBCHKShelf.Location = New System.Drawing.Point(61, 33)
        Me.TBCHKShelf.Name = "TBCHKShelf"
        Me.TBCHKShelf.Size = New System.Drawing.Size(105, 19)
        Me.TBCHKShelf.TabIndex = 11
        '
        'TBCHKWHCode
        '
        Me.TBCHKWHCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBCHKWHCode.Location = New System.Drawing.Point(61, 11)
        Me.TBCHKWHCode.Name = "TBCHKWHCode"
        Me.TBCHKWHCode.ReadOnly = True
        Me.TBCHKWHCode.Size = New System.Drawing.Size(53, 19)
        Me.TBCHKWHCode.TabIndex = 10
        Me.TBCHKWHCode.Text = "S02"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(6, 34)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 20)
        Me.Label4.Text = "ที่เก็บ :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(9, 12)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(50, 20)
        Me.Label5.Text = "คลัง :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'FrmAddItemShelf
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(240, 320)
        Me.Controls.Add(Me.PNCheckShelf)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "FrmAddItemShelf"
        Me.Text = "โปรแกรม บันทึกที่เก็บสินค้า"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.PNCheckShelf.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TBBarCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ListViewSelectItem As System.Windows.Forms.ListView
    Friend WithEvents BTNMenu As System.Windows.Forms.Button
    Friend WithEvents BTNClear As System.Windows.Forms.Button
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents TBWHCode As System.Windows.Forms.TextBox
    Friend WithEvents TBShelf As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents BTNCheckShelf As System.Windows.Forms.Button
    Friend WithEvents TBZone As System.Windows.Forms.TextBox
    Friend WithEvents PNCheckShelf As System.Windows.Forms.Panel
    Friend WithEvents ListViewCHKItem As System.Windows.Forms.ListView
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents TBCHKShelf As System.Windows.Forms.TextBox
    Friend WithEvents TBCHKWHCode As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents BTNCHKExit As System.Windows.Forms.Button
    Friend WithEvents BTNCHKClear As System.Windows.Forms.Button
End Class
