<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class FormShelfAddItem
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
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.CMBZone = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TBShelfID = New System.Windows.Forms.TextBox
        Me.BTNRefresh = New System.Windows.Forms.Button
        Me.VScrollBar1 = New System.Windows.Forms.VScrollBar
        Me.CMBWHCode = New System.Windows.Forms.ComboBox
        Me.ListViewItem = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.BTNSaveData = New System.Windows.Forms.Button
        Me.BTNClose = New System.Windows.Forms.Button
        Me.TBBarcode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label6.Location = New System.Drawing.Point(159, 6)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(55, 20)
        Me.Label6.Text = "โซน :"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(15, 6)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(36, 20)
        Me.Label5.Text = "คลัง :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'CMBZone
        '
        Me.CMBZone.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.CMBZone.Location = New System.Drawing.Point(216, 5)
        Me.CMBZone.Name = "CMBZone"
        Me.CMBZone.Size = New System.Drawing.Size(105, 19)
        Me.CMBZone.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(2, 31)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 20)
        Me.Label4.Text = "ที่เก็บ :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBShelfID
        '
        Me.TBShelfID.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.TBShelfID.Location = New System.Drawing.Point(53, 29)
        Me.TBShelfID.Name = "TBShelfID"
        Me.TBShelfID.Size = New System.Drawing.Size(95, 19)
        Me.TBShelfID.TabIndex = 2
        '
        'BTNRefresh
        '
        Me.BTNRefresh.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNRefresh.Location = New System.Drawing.Point(4, 269)
        Me.BTNRefresh.Name = "BTNRefresh"
        Me.BTNRefresh.Size = New System.Drawing.Size(59, 22)
        Me.BTNRefresh.TabIndex = 5
        Me.BTNRefresh.Text = "F2-เคลียร์"
        '
        'VScrollBar1
        '
        Me.VScrollBar1.Location = New System.Drawing.Point(8, 94)
        Me.VScrollBar1.Name = "VScrollBar1"
        Me.VScrollBar1.Size = New System.Drawing.Size(13, 154)
        Me.VScrollBar1.TabIndex = 49
        Me.VScrollBar1.Visible = False
        '
        'CMBWHCode
        '
        Me.CMBWHCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.CMBWHCode.Location = New System.Drawing.Point(53, 5)
        Me.CMBWHCode.Name = "CMBWHCode"
        Me.CMBWHCode.Size = New System.Drawing.Size(74, 19)
        Me.CMBWHCode.TabIndex = 0
        '
        'ListViewItem
        '
        Me.ListViewItem.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ListViewItem.Columns.Add(Me.ColumnHeader1)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader2)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader3)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader5)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader6)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader7)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader8)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader9)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader4)
        Me.ListViewItem.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.ListViewItem.FullRowSelect = True
        Me.ListViewItem.Location = New System.Drawing.Point(4, 66)
        Me.ListViewItem.Name = "ListViewItem"
        Me.ListViewItem.Size = New System.Drawing.Size(317, 200)
        Me.ListViewItem.TabIndex = 4
        Me.ListViewItem.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ลำดับ"
        Me.ColumnHeader1.Width = 40
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "รหัส"
        Me.ColumnHeader2.Width = 100
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "ชื่อสินค้า"
        Me.ColumnHeader3.Width = 200
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "ที่เก็บ"
        Me.ColumnHeader5.Width = 60
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "คลัง"
        Me.ColumnHeader6.Width = 70
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "โซน"
        Me.ColumnHeader7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader7.Width = 80
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "วันที่"
        Me.ColumnHeader8.Width = 60
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "IsSave"
        Me.ColumnHeader9.Width = 60
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "หน่วย"
        Me.ColumnHeader4.Width = 60
        '
        'BTNSaveData
        '
        Me.BTNSaveData.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNSaveData.Location = New System.Drawing.Point(69, 269)
        Me.BTNSaveData.Name = "BTNSaveData"
        Me.BTNSaveData.Size = New System.Drawing.Size(59, 22)
        Me.BTNSaveData.TabIndex = 6
        Me.BTNSaveData.Text = "F5-บันทึก"
        '
        'BTNClose
        '
        Me.BTNClose.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNClose.Location = New System.Drawing.Point(134, 269)
        Me.BTNClose.Name = "BTNClose"
        Me.BTNClose.Size = New System.Drawing.Size(59, 22)
        Me.BTNClose.TabIndex = 7
        Me.BTNClose.Text = "Esc-ออก"
        '
        'TBBarcode
        '
        Me.TBBarcode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Regular)
        Me.TBBarcode.Location = New System.Drawing.Point(216, 29)
        Me.TBBarcode.Name = "TBBarcode"
        Me.TBBarcode.Size = New System.Drawing.Size(105, 19)
        Me.TBBarcode.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(163, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 20)
        Me.Label2.Text = "บาร์โค้ด :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'FormShelfAddItem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.MediumBlue
        Me.ClientSize = New System.Drawing.Size(325, 300)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.CMBZone)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TBShelfID)
        Me.Controls.Add(Me.BTNRefresh)
        Me.Controls.Add(Me.VScrollBar1)
        Me.Controls.Add(Me.CMBWHCode)
        Me.Controls.Add(Me.ListViewItem)
        Me.Controls.Add(Me.BTNSaveData)
        Me.Controls.Add(Me.BTNClose)
        Me.Controls.Add(Me.TBBarcode)
        Me.Controls.Add(Me.Label2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FormShelfAddItem"
        Me.Text = "FormShelfAddItem"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CMBZone As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TBShelfID As System.Windows.Forms.TextBox
    Friend WithEvents BTNRefresh As System.Windows.Forms.Button
    Friend WithEvents VScrollBar1 As System.Windows.Forms.VScrollBar
    Friend WithEvents CMBWHCode As System.Windows.Forms.ComboBox
    Friend WithEvents ListViewItem As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents BTNSaveData As System.Windows.Forms.Button
    Friend WithEvents BTNClose As System.Windows.Forms.Button
    Friend WithEvents TBBarcode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
End Class
