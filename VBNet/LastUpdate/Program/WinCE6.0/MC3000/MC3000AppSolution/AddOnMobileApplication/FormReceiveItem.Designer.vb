<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class FormReceiveItem
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
        Me.Label3 = New System.Windows.Forms.Label
        Me.TBItemName = New System.Windows.Forms.TextBox
        Me.TBQty = New System.Windows.Forms.TextBox
        Me.TBBarcode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TBUnitPrice = New System.Windows.Forms.TextBox
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.BTNClose = New System.Windows.Forms.Button
        Me.VScrollBar1 = New System.Windows.Forms.VScrollBar
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ListViewItem = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.BTNClearScreen = New System.Windows.Forms.Button
        Me.TBPONo = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.TBAP = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TBItemCode = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'BTNSaveData
        '
        Me.BTNSaveData.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNSaveData.Location = New System.Drawing.Point(68, 271)
        Me.BTNSaveData.Name = "BTNSaveData"
        Me.BTNSaveData.Size = New System.Drawing.Size(59, 22)
        Me.BTNSaveData.TabIndex = 101
        Me.BTNSaveData.Text = "F5-บันทึก"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle))
        Me.Label3.Location = New System.Drawing.Point(6, 130)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(76, 20)
        Me.Label3.Text = "รายการสินค้า"
        '
        'TBItemName
        '
        Me.TBItemName.BackColor = System.Drawing.Color.LightPink
        Me.TBItemName.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBItemName.Location = New System.Drawing.Point(68, 79)
        Me.TBItemName.Name = "TBItemName"
        Me.TBItemName.ReadOnly = True
        Me.TBItemName.Size = New System.Drawing.Size(254, 19)
        Me.TBItemName.TabIndex = 94
        '
        'TBQty
        '
        Me.TBQty.BackColor = System.Drawing.Color.DarkOrange
        Me.TBQty.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBQty.Location = New System.Drawing.Point(219, 102)
        Me.TBQty.Name = "TBQty"
        Me.TBQty.Size = New System.Drawing.Size(103, 19)
        Me.TBQty.TabIndex = 98
        '
        'TBBarcode
        '
        Me.TBBarcode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBBarcode.Location = New System.Drawing.Point(68, 56)
        Me.TBBarcode.Name = "TBBarcode"
        Me.TBBarcode.Size = New System.Drawing.Size(100, 19)
        Me.TBBarcode.TabIndex = 92
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(172, 102)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 20)
        Me.Label1.Text = "จำนวน :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBUnitPrice
        '
        Me.TBUnitPrice.BackColor = System.Drawing.Color.LightPink
        Me.TBUnitPrice.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBUnitPrice.Location = New System.Drawing.Point(68, 102)
        Me.TBUnitPrice.Name = "TBUnitPrice"
        Me.TBUnitPrice.ReadOnly = True
        Me.TBUnitPrice.Size = New System.Drawing.Size(100, 19)
        Me.TBUnitPrice.TabIndex = 95
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "หน่วย"
        Me.ColumnHeader8.Width = 70
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label10.Location = New System.Drawing.Point(25, 56)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(41, 18)
        Me.Label10.Text = "รหัส :"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label8.Location = New System.Drawing.Point(25, 102)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(41, 18)
        Me.Label8.Text = "หน่วย :"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label7.Location = New System.Drawing.Point(29, 80)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(37, 18)
        Me.Label7.Text = "ชื่อ :"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Black
        Me.Panel1.Location = New System.Drawing.Point(5, 124)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(317, 5)
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "รหัส"
        Me.ColumnHeader7.Width = 80
        '
        'BTNClose
        '
        Me.BTNClose.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNClose.Location = New System.Drawing.Point(133, 271)
        Me.BTNClose.Name = "BTNClose"
        Me.BTNClose.Size = New System.Drawing.Size(59, 22)
        Me.BTNClose.TabIndex = 102
        Me.BTNClose.Text = "ESC-ออก"
        '
        'VScrollBar1
        '
        Me.VScrollBar1.Location = New System.Drawing.Point(7, 171)
        Me.VScrollBar1.Name = "VScrollBar1"
        Me.VScrollBar1.Size = New System.Drawing.Size(13, 79)
        Me.VScrollBar1.TabIndex = 103
        Me.VScrollBar1.Visible = False
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "นับได้"
        Me.ColumnHeader2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader2.Width = 70
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "สั่งซื้อ"
        Me.ColumnHeader4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader4.Width = 70
        '
        'ListViewItem
        '
        Me.ListViewItem.Columns.Add(Me.ColumnHeader1)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader2)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader3)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader4)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader7)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader8)
        Me.ListViewItem.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.ListViewItem.FullRowSelect = True
        Me.ListViewItem.Location = New System.Drawing.Point(3, 147)
        Me.ListViewItem.Name = "ListViewItem"
        Me.ListViewItem.Size = New System.Drawing.Size(319, 122)
        Me.ListViewItem.TabIndex = 99
        Me.ListViewItem.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ลำดับ"
        Me.ColumnHeader1.Width = 44
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "ชื่อ"
        Me.ColumnHeader3.Width = 130
        '
        'BTNClearScreen
        '
        Me.BTNClearScreen.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNClearScreen.Location = New System.Drawing.Point(3, 271)
        Me.BTNClearScreen.Name = "BTNClearScreen"
        Me.BTNClearScreen.Size = New System.Drawing.Size(59, 22)
        Me.BTNClearScreen.TabIndex = 100
        Me.BTNClearScreen.Text = "F2-ทำใหม่"
        '
        'TBPONo
        '
        Me.TBPONo.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBPONo.Location = New System.Drawing.Point(68, 4)
        Me.TBPONo.Name = "TBPONo"
        Me.TBPONo.Size = New System.Drawing.Size(100, 19)
        Me.TBPONo.TabIndex = 115
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(3, 4)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 18)
        Me.Label4.Text = "PONO :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Black
        Me.Panel2.Location = New System.Drawing.Point(4, 48)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(317, 5)
        '
        'TBAP
        '
        Me.TBAP.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBAP.Location = New System.Drawing.Point(68, 27)
        Me.TBAP.Name = "TBAP"
        Me.TBAP.Size = New System.Drawing.Size(254, 19)
        Me.TBAP.TabIndex = 118
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(3, 27)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 18)
        Me.Label5.Text = "เจ้าหนี้ :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox1.Location = New System.Drawing.Point(219, 4)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(103, 19)
        Me.TextBox1.TabIndex = 130
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label6.Location = New System.Drawing.Point(154, 4)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(63, 18)
        Me.Label6.Text = "RefNO :"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBItemCode
        '
        Me.TBItemCode.BackColor = System.Drawing.Color.LightPink
        Me.TBItemCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBItemCode.Location = New System.Drawing.Point(219, 56)
        Me.TBItemCode.Name = "TBItemCode"
        Me.TBItemCode.ReadOnly = True
        Me.TBItemCode.Size = New System.Drawing.Size(103, 19)
        Me.TBItemCode.TabIndex = 93
        '
        'FormReceiveItem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.SeaGreen
        Me.ClientSize = New System.Drawing.Size(325, 300)
        Me.ControlBox = False
        Me.Controls.Add(Me.TBAP)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.TBPONo)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.BTNSaveData)
        Me.Controls.Add(Me.TBItemName)
        Me.Controls.Add(Me.TBQty)
        Me.Controls.Add(Me.TBBarcode)
        Me.Controls.Add(Me.TBUnitPrice)
        Me.Controls.Add(Me.TBItemCode)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.BTNClose)
        Me.Controls.Add(Me.VScrollBar1)
        Me.Controls.Add(Me.ListViewItem)
        Me.Controls.Add(Me.BTNClearScreen)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label6)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FormReceiveItem"
        Me.Text = "FormReceiveItem"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BTNSaveData As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TBItemName As System.Windows.Forms.TextBox
    Friend WithEvents TBQty As System.Windows.Forms.TextBox
    Friend WithEvents TBBarcode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TBUnitPrice As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents BTNClose As System.Windows.Forms.Button
    Friend WithEvents VScrollBar1 As System.Windows.Forms.VScrollBar
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ListViewItem As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents BTNClearScreen As System.Windows.Forms.Button
    Friend WithEvents TBPONo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TBAP As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TBItemCode As System.Windows.Forms.TextBox
End Class
