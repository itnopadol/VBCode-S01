<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class FormPrintLabel
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
        Me.BTNSaveData = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.BTNClose = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.VScrollBar1 = New System.Windows.Forms.VScrollBar
        Me.CMBLabelType = New System.Windows.Forms.ComboBox
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ListViewItem = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TBPrice = New System.Windows.Forms.TextBox
        Me.TBItemName = New System.Windows.Forms.TextBox
        Me.TBQty = New System.Windows.Forms.TextBox
        Me.TBBarcode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TBUnitPrice = New System.Windows.Forms.TextBox
        Me.BTNRedDot = New System.Windows.Forms.Button
        Me.Label10 = New System.Windows.Forms.Label
        Me.TBItemCode = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.TBTypePrice = New System.Windows.Forms.TextBox
        Me.BTNClearScreen = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'BTNSaveData
        '
        Me.BTNSaveData.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNSaveData.Location = New System.Drawing.Point(67, 269)
        Me.BTNSaveData.Name = "BTNSaveData"
        Me.BTNSaveData.Size = New System.Drawing.Size(59, 22)
        Me.BTNSaveData.TabIndex = 10
        Me.BTNSaveData.Text = "F5-บันทึก"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(2, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 20)
        Me.Label4.Text = "เลือกป้าย :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'BTNClose
        '
        Me.BTNClose.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNClose.Location = New System.Drawing.Point(132, 269)
        Me.BTNClose.Name = "BTNClose"
        Me.BTNClose.Size = New System.Drawing.Size(59, 22)
        Me.BTNClose.TabIndex = 11
        Me.BTNClose.Text = "ESC-ออก"
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label8.Location = New System.Drawing.Point(24, 73)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(41, 18)
        Me.Label8.Text = "หน่วย :"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label7.Location = New System.Drawing.Point(28, 51)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(37, 18)
        Me.Label7.Text = "ชื่อ :"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Black
        Me.Panel1.Location = New System.Drawing.Point(4, 117)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(317, 5)
        '
        'VScrollBar1
        '
        Me.VScrollBar1.Location = New System.Drawing.Point(6, 164)
        Me.VScrollBar1.Name = "VScrollBar1"
        Me.VScrollBar1.Size = New System.Drawing.Size(13, 79)
        Me.VScrollBar1.TabIndex = 72
        Me.VScrollBar1.Visible = False
        '
        'CMBLabelType
        '
        Me.CMBLabelType.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.CMBLabelType.Location = New System.Drawing.Point(67, 5)
        Me.CMBLabelType.Name = "CMBLabelType"
        Me.CMBLabelType.Size = New System.Drawing.Size(254, 19)
        Me.CMBLabelType.TabIndex = 0
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "บาร์โค้ด"
        Me.ColumnHeader2.Width = 100
        '
        'ListViewItem
        '
        Me.ListViewItem.Columns.Add(Me.ColumnHeader1)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader2)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader3)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader4)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader5)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader6)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader7)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader8)
        Me.ListViewItem.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.ListViewItem.FullRowSelect = True
        Me.ListViewItem.Location = New System.Drawing.Point(4, 140)
        Me.ListViewItem.Name = "ListViewItem"
        Me.ListViewItem.Size = New System.Drawing.Size(317, 125)
        Me.ListViewItem.TabIndex = 8
        Me.ListViewItem.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ลำดับ"
        Me.ColumnHeader1.Width = 45
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "จำนวน"
        Me.ColumnHeader3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader3.Width = 50
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "ประเภท"
        Me.ColumnHeader4.Width = 60
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "วันที่"
        Me.ColumnHeader5.Width = 80
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "IsSave"
        Me.ColumnHeader6.Width = 40
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "รหัส"
        Me.ColumnHeader7.Width = 80
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "หน่วย"
        Me.ColumnHeader8.Width = 70
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle))
        Me.Label3.Location = New System.Drawing.Point(5, 123)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(76, 20)
        Me.Label3.Text = "รายการสินค้า"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label6.Location = New System.Drawing.Point(28, 95)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(37, 18)
        Me.Label6.Text = "ราคา :"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBPrice
        '
        Me.TBPrice.BackColor = System.Drawing.Color.LightPink
        Me.TBPrice.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBPrice.Location = New System.Drawing.Point(67, 93)
        Me.TBPrice.Name = "TBPrice"
        Me.TBPrice.ReadOnly = True
        Me.TBPrice.Size = New System.Drawing.Size(100, 19)
        Me.TBPrice.TabIndex = 6
        '
        'TBItemName
        '
        Me.TBItemName.BackColor = System.Drawing.Color.LightPink
        Me.TBItemName.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBItemName.Location = New System.Drawing.Point(67, 51)
        Me.TBItemName.Name = "TBItemName"
        Me.TBItemName.ReadOnly = True
        Me.TBItemName.Size = New System.Drawing.Size(254, 19)
        Me.TBItemName.TabIndex = 3
        '
        'TBQty
        '
        Me.TBQty.BackColor = System.Drawing.Color.DarkOrange
        Me.TBQty.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBQty.Location = New System.Drawing.Point(218, 93)
        Me.TBQty.Name = "TBQty"
        Me.TBQty.Size = New System.Drawing.Size(103, 19)
        Me.TBQty.TabIndex = 7
        '
        'TBBarcode
        '
        Me.TBBarcode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBBarcode.Location = New System.Drawing.Point(67, 30)
        Me.TBBarcode.Name = "TBBarcode"
        Me.TBBarcode.Size = New System.Drawing.Size(100, 19)
        Me.TBBarcode.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(2, 30)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 18)
        Me.Label2.Text = "บาร์โค้ด :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(171, 95)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 20)
        Me.Label1.Text = "จำนวน :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBUnitPrice
        '
        Me.TBUnitPrice.BackColor = System.Drawing.Color.LightPink
        Me.TBUnitPrice.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBUnitPrice.Location = New System.Drawing.Point(67, 72)
        Me.TBUnitPrice.Name = "TBUnitPrice"
        Me.TBUnitPrice.ReadOnly = True
        Me.TBUnitPrice.Size = New System.Drawing.Size(100, 19)
        Me.TBUnitPrice.TabIndex = 4
        '
        'BTNRedDot
        '
        Me.BTNRedDot.BackColor = System.Drawing.Color.Red
        Me.BTNRedDot.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNRedDot.Location = New System.Drawing.Point(302, 30)
        Me.BTNRedDot.Name = "BTNRedDot"
        Me.BTNRedDot.Size = New System.Drawing.Size(19, 19)
        Me.BTNRedDot.TabIndex = 80
        Me.BTNRedDot.Visible = False
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label10.Location = New System.Drawing.Point(175, 30)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(41, 18)
        Me.Label10.Text = "รหัส :"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBItemCode
        '
        Me.TBItemCode.BackColor = System.Drawing.Color.LightPink
        Me.TBItemCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBItemCode.Location = New System.Drawing.Point(218, 30)
        Me.TBItemCode.Name = "TBItemCode"
        Me.TBItemCode.ReadOnly = True
        Me.TBItemCode.Size = New System.Drawing.Size(103, 19)
        Me.TBItemCode.TabIndex = 2
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label11.Location = New System.Drawing.Point(167, 73)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(50, 18)
        Me.Label11.Text = "ประเภท :"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBTypePrice
        '
        Me.TBTypePrice.BackColor = System.Drawing.Color.LightPink
        Me.TBTypePrice.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBTypePrice.Location = New System.Drawing.Point(218, 72)
        Me.TBTypePrice.Name = "TBTypePrice"
        Me.TBTypePrice.ReadOnly = True
        Me.TBTypePrice.Size = New System.Drawing.Size(103, 19)
        Me.TBTypePrice.TabIndex = 5
        '
        'BTNClearScreen
        '
        Me.BTNClearScreen.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNClearScreen.Location = New System.Drawing.Point(4, 269)
        Me.BTNClearScreen.Name = "BTNClearScreen"
        Me.BTNClearScreen.Size = New System.Drawing.Size(59, 22)
        Me.BTNClearScreen.TabIndex = 9
        Me.BTNClearScreen.Text = "F2-ทำใหม่"
        '
        'FormPrintLabel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.SkyBlue
        Me.ClientSize = New System.Drawing.Size(325, 300)
        Me.ControlBox = False
        Me.Controls.Add(Me.BTNClearScreen)
        Me.Controls.Add(Me.TBTypePrice)
        Me.Controls.Add(Me.BTNRedDot)
        Me.Controls.Add(Me.BTNSaveData)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.BTNClose)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.VScrollBar1)
        Me.Controls.Add(Me.CMBLabelType)
        Me.Controls.Add(Me.ListViewItem)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TBPrice)
        Me.Controls.Add(Me.TBItemName)
        Me.Controls.Add(Me.TBQty)
        Me.Controls.Add(Me.TBBarcode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TBUnitPrice)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.TBItemCode)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FormPrintLabel"
        Me.Text = "FormPrintLabel"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BTNSaveData As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BTNClose As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents VScrollBar1 As System.Windows.Forms.VScrollBar
    Friend WithEvents CMBLabelType As System.Windows.Forms.ComboBox
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ListViewItem As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TBPrice As System.Windows.Forms.TextBox
    Friend WithEvents TBItemName As System.Windows.Forms.TextBox
    Friend WithEvents TBQty As System.Windows.Forms.TextBox
    Friend WithEvents TBBarcode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TBUnitPrice As System.Windows.Forms.TextBox
    Friend WithEvents BTNRedDot As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TBItemCode As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TBTypePrice As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents BTNClearScreen As System.Windows.Forms.Button
End Class
