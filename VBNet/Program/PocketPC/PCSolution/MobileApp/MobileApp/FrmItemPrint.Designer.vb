<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class FrmItemPrint
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmItemPrint))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel7 = New System.Windows.Forms.Panel
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.TBBarCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.ListViewItem = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.BTNSave = New System.Windows.Forms.Button
        Me.BTNMenu = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PNShowDetails = New System.Windows.Forms.Panel
        Me.TBMemBarCode = New System.Windows.Forms.TextBox
        Me.PNKeyQty = New System.Windows.Forms.Panel
        Me.TBQty = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.TBPromoExpire = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TBPriceType = New System.Windows.Forms.TextBox
        Me.TBPrice = New System.Windows.Forms.TextBox
        Me.TBUnitCode = New System.Windows.Forms.TextBox
        Me.TBItemName = New System.Windows.Forms.TextBox
        Me.TBItemCode = New System.Windows.Forms.TextBox
        Me.BTNAddItem = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.BTNAddQty = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.PNShowDetails.SuspendLayout()
        Me.PNKeyQty.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.BTNAddQty)
        Me.Panel1.Controls.Add(Me.Panel7)
        Me.Panel1.Controls.Add(Me.Panel6)
        Me.Panel1.Controls.Add(Me.Panel5)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.TBBarCode)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.PictureBox2)
        Me.Panel1.Location = New System.Drawing.Point(0, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(240, 70)
        '
        'Panel7
        '
        Me.Panel7.BackColor = System.Drawing.Color.Orange
        Me.Panel7.Location = New System.Drawing.Point(3, 0)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(234, 7)
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.Color.Orange
        Me.Panel6.Location = New System.Drawing.Point(3, 61)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(234, 7)
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.DarkBlue
        Me.Panel5.Location = New System.Drawing.Point(3, 53)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(234, 7)
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.DarkBlue
        Me.Panel4.Location = New System.Drawing.Point(3, 8)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(234, 7)
        '
        'TBBarCode
        '
        Me.TBBarCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.TBBarCode.Location = New System.Drawing.Point(55, 24)
        Me.TBBarCode.Name = "TBBarCode"
        Me.TBBarCode.Size = New System.Drawing.Size(111, 19)
        Me.TBBarCode.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.Label1.Location = New System.Drawing.Point(3, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 20)
        Me.Label1.Text = "บาร์โค้ด :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(92, -16)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(147, 97)
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DarkBlue
        Me.Panel2.Controls.Add(Me.ListViewItem)
        Me.Panel2.Location = New System.Drawing.Point(0, 75)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(240, 203)
        '
        'ListViewItem
        '
        Me.ListViewItem.Columns.Add(Me.ColumnHeader1)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader2)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader3)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader4)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader5)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader6)
        Me.ListViewItem.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.ListViewItem.FullRowSelect = True
        Me.ListViewItem.Location = New System.Drawing.Point(3, 6)
        Me.ListViewItem.Name = "ListViewItem"
        Me.ListViewItem.Size = New System.Drawing.Size(234, 194)
        Me.ListViewItem.TabIndex = 0
        Me.ListViewItem.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ลำดับ"
        Me.ColumnHeader1.Width = 40
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "รหัสสินค้า"
        Me.ColumnHeader2.Width = 80
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "ชื่อสินค้า"
        Me.ColumnHeader3.Width = 110
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "จำนวน"
        Me.ColumnHeader4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader4.Width = 70
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "บาร์โค้ด"
        Me.ColumnHeader5.Width = 100
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "หน่วย"
        Me.ColumnHeader6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader6.Width = 80
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.White
        Me.Panel3.Controls.Add(Me.BTNSave)
        Me.Panel3.Controls.Add(Me.BTNMenu)
        Me.Panel3.Controls.Add(Me.PictureBox1)
        Me.Panel3.Location = New System.Drawing.Point(0, 279)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(240, 41)
        '
        'BTNSave
        '
        Me.BTNSave.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNSave.Location = New System.Drawing.Point(197, 8)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(40, 20)
        Me.BTNSave.TabIndex = 1
        Me.BTNSave.Text = "บันทึก"
        '
        'BTNMenu
        '
        Me.BTNMenu.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNMenu.Location = New System.Drawing.Point(151, 8)
        Me.BTNMenu.Name = "BTNMenu"
        Me.BTNMenu.Size = New System.Drawing.Size(40, 20)
        Me.BTNMenu.TabIndex = 0
        Me.BTNMenu.Text = "กลับ"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(-6, -9)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(254, 49)
        '
        'PNShowDetails
        '
        Me.PNShowDetails.BackColor = System.Drawing.Color.Black
        Me.PNShowDetails.Controls.Add(Me.TBMemBarCode)
        Me.PNShowDetails.Controls.Add(Me.PNKeyQty)
        Me.PNShowDetails.Controls.Add(Me.Label7)
        Me.PNShowDetails.Controls.Add(Me.TBPromoExpire)
        Me.PNShowDetails.Controls.Add(Me.Label6)
        Me.PNShowDetails.Controls.Add(Me.TBPriceType)
        Me.PNShowDetails.Controls.Add(Me.TBPrice)
        Me.PNShowDetails.Controls.Add(Me.TBUnitCode)
        Me.PNShowDetails.Controls.Add(Me.TBItemName)
        Me.PNShowDetails.Controls.Add(Me.TBItemCode)
        Me.PNShowDetails.Controls.Add(Me.BTNAddItem)
        Me.PNShowDetails.Controls.Add(Me.Label5)
        Me.PNShowDetails.Controls.Add(Me.Label4)
        Me.PNShowDetails.Controls.Add(Me.Label3)
        Me.PNShowDetails.Controls.Add(Me.Label2)
        Me.PNShowDetails.Controls.Add(Me.PictureBox3)
        Me.PNShowDetails.Location = New System.Drawing.Point(0, 70)
        Me.PNShowDetails.Name = "PNShowDetails"
        Me.PNShowDetails.Size = New System.Drawing.Size(240, 249)
        Me.PNShowDetails.Visible = False
        '
        'TBMemBarCode
        '
        Me.TBMemBarCode.BackColor = System.Drawing.Color.Khaki
        Me.TBMemBarCode.Location = New System.Drawing.Point(172, 10)
        Me.TBMemBarCode.Name = "TBMemBarCode"
        Me.TBMemBarCode.ReadOnly = True
        Me.TBMemBarCode.Size = New System.Drawing.Size(65, 21)
        Me.TBMemBarCode.TabIndex = 28
        Me.TBMemBarCode.Visible = False
        '
        'PNKeyQty
        '
        Me.PNKeyQty.Controls.Add(Me.TBQty)
        Me.PNKeyQty.Controls.Add(Me.Label9)
        Me.PNKeyQty.Controls.Add(Me.Label8)
        Me.PNKeyQty.Location = New System.Drawing.Point(7, 145)
        Me.PNKeyQty.Name = "PNKeyQty"
        Me.PNKeyQty.Size = New System.Drawing.Size(227, 90)
        Me.PNKeyQty.Visible = False
        '
        'TBQty
        '
        Me.TBQty.BackColor = System.Drawing.Color.Orange
        Me.TBQty.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Bold)
        Me.TBQty.Location = New System.Drawing.Point(59, 30)
        Me.TBQty.Name = "TBQty"
        Me.TBQty.Size = New System.Drawing.Size(136, 39)
        Me.TBQty.TabIndex = 12
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(4, 41)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(53, 20)
        Me.Label9.Text = "จำนวน :"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle))
        Me.Label8.Location = New System.Drawing.Point(4, 5)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(180, 20)
        Me.Label8.Text = "กรอกจำนวนป้ายที่จะพิมพ์"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(7, 170)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(57, 20)
        Me.Label7.Text = "หมดโปรฯ :"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBPromoExpire
        '
        Me.TBPromoExpire.BackColor = System.Drawing.Color.Khaki
        Me.TBPromoExpire.Location = New System.Drawing.Point(66, 168)
        Me.TBPromoExpire.Name = "TBPromoExpire"
        Me.TBPromoExpire.ReadOnly = True
        Me.TBPromoExpire.Size = New System.Drawing.Size(100, 21)
        Me.TBPromoExpire.TabIndex = 21
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(7, 147)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 20)
        Me.Label6.Text = "ประเภท :"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBPriceType
        '
        Me.TBPriceType.BackColor = System.Drawing.Color.Khaki
        Me.TBPriceType.Location = New System.Drawing.Point(66, 145)
        Me.TBPriceType.Name = "TBPriceType"
        Me.TBPriceType.ReadOnly = True
        Me.TBPriceType.Size = New System.Drawing.Size(136, 21)
        Me.TBPriceType.TabIndex = 16
        '
        'TBPrice
        '
        Me.TBPrice.BackColor = System.Drawing.Color.Orange
        Me.TBPrice.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Bold)
        Me.TBPrice.Location = New System.Drawing.Point(66, 104)
        Me.TBPrice.Name = "TBPrice"
        Me.TBPrice.ReadOnly = True
        Me.TBPrice.Size = New System.Drawing.Size(136, 39)
        Me.TBPrice.TabIndex = 11
        '
        'TBUnitCode
        '
        Me.TBUnitCode.BackColor = System.Drawing.Color.Khaki
        Me.TBUnitCode.Location = New System.Drawing.Point(66, 81)
        Me.TBUnitCode.Name = "TBUnitCode"
        Me.TBUnitCode.ReadOnly = True
        Me.TBUnitCode.Size = New System.Drawing.Size(100, 21)
        Me.TBUnitCode.TabIndex = 10
        '
        'TBItemName
        '
        Me.TBItemName.BackColor = System.Drawing.Color.Khaki
        Me.TBItemName.Location = New System.Drawing.Point(66, 33)
        Me.TBItemName.Multiline = True
        Me.TBItemName.Name = "TBItemName"
        Me.TBItemName.ReadOnly = True
        Me.TBItemName.Size = New System.Drawing.Size(171, 46)
        Me.TBItemName.TabIndex = 9
        '
        'TBItemCode
        '
        Me.TBItemCode.BackColor = System.Drawing.Color.Khaki
        Me.TBItemCode.Location = New System.Drawing.Point(66, 10)
        Me.TBItemCode.Name = "TBItemCode"
        Me.TBItemCode.ReadOnly = True
        Me.TBItemCode.Size = New System.Drawing.Size(100, 21)
        Me.TBItemCode.TabIndex = 8
        '
        'BTNAddItem
        '
        Me.BTNAddItem.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNAddItem.Location = New System.Drawing.Point(66, 195)
        Me.BTNAddItem.Name = "BTNAddItem"
        Me.BTNAddItem.Size = New System.Drawing.Size(72, 20)
        Me.BTNAddItem.TabIndex = 7
        Me.BTNAddItem.Text = "เพิ่มรายการ"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(7, 114)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(57, 20)
        Me.Label5.Text = "ราคาขาย :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(3, 83)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(61, 20)
        Me.Label4.Text = "หน่วยนับ :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(3, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 20)
        Me.Label3.Text = "ชื่อสินค้า :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(3, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 20)
        Me.Label2.Text = "รหัสสินค้า :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(73, -64)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(164, 324)
        '
        'BTNAddQty
        '
        Me.BTNAddQty.Location = New System.Drawing.Point(173, 25)
        Me.BTNAddQty.Name = "BTNAddQty"
        Me.BTNAddQty.Size = New System.Drawing.Size(42, 18)
        Me.BTNAddQty.TabIndex = 6
        Me.BTNAddQty.Text = "เพิ่ม"
        '
        'FrmItemPrint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(240, 320)
        Me.Controls.Add(Me.PNShowDetails)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "FrmItemPrint"
        Me.Text = "FrmItemPrint"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.PNShowDetails.ResumeLayout(False)
        Me.PNKeyQty.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents ListViewItem As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents TBBarCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents BTNMenu As System.Windows.Forms.Button
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PNShowDetails As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TBPrice As System.Windows.Forms.TextBox
    Friend WithEvents TBUnitCode As System.Windows.Forms.TextBox
    Friend WithEvents TBItemName As System.Windows.Forms.TextBox
    Friend WithEvents TBItemCode As System.Windows.Forms.TextBox
    Friend WithEvents BTNAddItem As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TBPriceType As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TBPromoExpire As System.Windows.Forms.TextBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PNKeyQty As System.Windows.Forms.Panel
    Friend WithEvents TBQty As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TBMemBarCode As System.Windows.Forms.TextBox
    Friend WithEvents BTNAddQty As System.Windows.Forms.Button
End Class
