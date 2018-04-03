<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormImportDataPriceStructureFromExcel
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Label1 = New System.Windows.Forms.Label
        Me.LBLFileName = New System.Windows.Forms.Label
        Me.DGHeader = New System.Windows.Forms.DataGridView
        Me.DGDetails = New System.Windows.Forms.DataGridView
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.BTNGenData = New System.Windows.Forms.Button
        Me.PB101 = New System.Windows.Forms.ProgressBar
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextMyDescription = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextDocNo = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.DocDate = New System.Windows.Forms.DateTimePicker
        Me.GBSearchPriceStructure = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.BTNPriceStructureExit = New System.Windows.Forms.Button
        Me.BTNPriceStructureConfirm = New System.Windows.Forms.Button
        Me.ListViewPriceStructure = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.BTNPriceStructureSearch = New System.Windows.Forms.Button
        Me.TextPriceStructureSearch = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.BTNClose = New System.Windows.Forms.Button
        Me.BTNPrint = New System.Windows.Forms.Button
        Me.BTNClearData = New System.Windows.Forms.Button
        Me.PBConfirm = New System.Windows.Forms.PictureBox
        Me.PBNew = New System.Windows.Forms.PictureBox
        Me.BTNFind = New System.Windows.Forms.Button
        Me.BTNSearchDocno = New System.Windows.Forms.Button
        Me.BTNSave = New System.Windows.Forms.Button
        CType(Me.DGHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GBSearchPriceStructure.SuspendLayout()
        CType(Me.PBConfirm, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PBNew, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(31, 52)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ชื่อที่อยู่ของเอกสาร :"
        '
        'LBLFileName
        '
        Me.LBLFileName.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.LBLFileName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LBLFileName.Location = New System.Drawing.Point(131, 49)
        Me.LBLFileName.Name = "LBLFileName"
        Me.LBLFileName.Size = New System.Drawing.Size(744, 21)
        Me.LBLFileName.TabIndex = 1
        '
        'DGHeader
        '
        Me.DGHeader.AllowUserToDeleteRows = False
        Me.DGHeader.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.SkyBlue
        Me.DGHeader.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DGHeader.GridColor = System.Drawing.SystemColors.MenuText
        Me.DGHeader.Location = New System.Drawing.Point(12, 101)
        Me.DGHeader.Name = "DGHeader"
        Me.DGHeader.ReadOnly = True
        Me.DGHeader.Size = New System.Drawing.Size(992, 136)
        Me.DGHeader.TabIndex = 3
        '
        'DGDetails
        '
        Me.DGDetails.AllowUserToDeleteRows = False
        Me.DGDetails.AllowUserToResizeRows = False
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.SkyBlue
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.White
        Me.DGDetails.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.MenuText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGDetails.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DGDetails.GridColor = System.Drawing.SystemColors.MenuText
        Me.DGDetails.Location = New System.Drawing.Point(12, 262)
        Me.DGDetails.Name = "DGDetails"
        Me.DGDetails.ReadOnly = True
        Me.DGDetails.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Black
        Me.DGDetails.Size = New System.Drawing.Size(992, 355)
        Me.DGDetails.TabIndex = 4
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'BTNGenData
        '
        Me.BTNGenData.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.BTNGenData.Location = New System.Drawing.Point(947, 46)
        Me.BTNGenData.Name = "BTNGenData"
        Me.BTNGenData.Size = New System.Drawing.Size(58, 26)
        Me.BTNGenData.TabIndex = 6
        Me.BTNGenData.Text = "GenData"
        Me.BTNGenData.UseVisualStyleBackColor = False
        '
        'PB101
        '
        Me.PB101.Location = New System.Drawing.Point(12, 78)
        Me.PB101.Name = "PB101"
        Me.PB101.Size = New System.Drawing.Size(992, 14)
        Me.PB101.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(451, 21)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "หมายเหตุ :"
        '
        'TextMyDescription
        '
        Me.TextMyDescription.Location = New System.Drawing.Point(515, 18)
        Me.TextMyDescription.Multiline = True
        Me.TextMyDescription.Name = "TextMyDescription"
        Me.TextMyDescription.Size = New System.Drawing.Size(490, 20)
        Me.TextMyDescription.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(56, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "เลขที่เอกสาร :"
        '
        'TextDocNo
        '
        Me.TextDocNo.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.TextDocNo.Location = New System.Drawing.Point(131, 18)
        Me.TextDocNo.Name = "TextDocNo"
        Me.TextDocNo.ReadOnly = True
        Me.TextDocNo.Size = New System.Drawing.Size(124, 20)
        Me.TextDocNo.TabIndex = 11
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(271, 21)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "วันที่เอกสาร :"
        '
        'DocDate
        '
        Me.DocDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DocDate.Location = New System.Drawing.Point(346, 17)
        Me.DocDate.Name = "DocDate"
        Me.DocDate.Size = New System.Drawing.Size(99, 20)
        Me.DocDate.TabIndex = 13
        '
        'GBSearchPriceStructure
        '
        Me.GBSearchPriceStructure.Controls.Add(Me.Label6)
        Me.GBSearchPriceStructure.Controls.Add(Me.BTNPriceStructureExit)
        Me.GBSearchPriceStructure.Controls.Add(Me.BTNPriceStructureConfirm)
        Me.GBSearchPriceStructure.Controls.Add(Me.ListViewPriceStructure)
        Me.GBSearchPriceStructure.Controls.Add(Me.BTNPriceStructureSearch)
        Me.GBSearchPriceStructure.Controls.Add(Me.TextPriceStructureSearch)
        Me.GBSearchPriceStructure.Controls.Add(Me.Label5)
        Me.GBSearchPriceStructure.Location = New System.Drawing.Point(12, 12)
        Me.GBSearchPriceStructure.Name = "GBSearchPriceStructure"
        Me.GBSearchPriceStructure.Size = New System.Drawing.Size(992, 663)
        Me.GBSearchPriceStructure.TabIndex = 16
        Me.GBSearchPriceStructure.TabStop = False
        Me.GBSearchPriceStructure.Text = "ค้นหา เอกสารเพื่อนำมาอนุมัติการปรับโครงสร้างราคา"
        Me.GBSearchPriceStructure.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(71, 83)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(78, 13)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "รายการเอกสาร"
        '
        'BTNPriceStructureExit
        '
        Me.BTNPriceStructureExit.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_07
        Me.BTNPriceStructureExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNPriceStructureExit.Location = New System.Drawing.Point(805, 582)
        Me.BTNPriceStructureExit.Name = "BTNPriceStructureExit"
        Me.BTNPriceStructureExit.Size = New System.Drawing.Size(122, 52)
        Me.BTNPriceStructureExit.TabIndex = 5
        Me.BTNPriceStructureExit.Text = "ออกหน้าอนุมัติ"
        Me.BTNPriceStructureExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNPriceStructureExit.UseVisualStyleBackColor = True
        '
        'BTNPriceStructureConfirm
        '
        Me.BTNPriceStructureConfirm.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_76
        Me.BTNPriceStructureConfirm.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNPriceStructureConfirm.Location = New System.Drawing.Point(666, 582)
        Me.BTNPriceStructureConfirm.Name = "BTNPriceStructureConfirm"
        Me.BTNPriceStructureConfirm.Size = New System.Drawing.Size(122, 52)
        Me.BTNPriceStructureConfirm.TabIndex = 4
        Me.BTNPriceStructureConfirm.Text = "อนุมัติเอกสาร"
        Me.BTNPriceStructureConfirm.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNPriceStructureConfirm.UseVisualStyleBackColor = True
        '
        'ListViewPriceStructure
        '
        Me.ListViewPriceStructure.BackColor = System.Drawing.SystemColors.Menu
        Me.ListViewPriceStructure.CheckBoxes = True
        Me.ListViewPriceStructure.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.ListViewPriceStructure.FullRowSelect = True
        Me.ListViewPriceStructure.GridLines = True
        Me.ListViewPriceStructure.Location = New System.Drawing.Point(74, 103)
        Me.ListViewPriceStructure.Name = "ListViewPriceStructure"
        Me.ListViewPriceStructure.Size = New System.Drawing.Size(853, 458)
        Me.ListViewPriceStructure.TabIndex = 3
        Me.ListViewPriceStructure.UseCompatibleStateImageBehavior = False
        Me.ListViewPriceStructure.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "เลขที่เอกสาร"
        Me.ColumnHeader1.Width = 120
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "วันที่เอกสาร"
        Me.ColumnHeader2.Width = 120
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "ผู้สร้างเอกสาร"
        Me.ColumnHeader3.Width = 200
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "ชื่อไฟล์เอกสาร"
        Me.ColumnHeader4.Width = 500
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "อนุมัติ"
        Me.ColumnHeader5.Width = 100
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "ที่อยู่ไฟล์เอกสาร"
        Me.ColumnHeader6.Width = 300
        '
        'BTNPriceStructureSearch
        '
        Me.BTNPriceStructureSearch.Image = Global.NPWindowApp.My.Resources.Resources.Windows_Explorer1
        Me.BTNPriceStructureSearch.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNPriceStructureSearch.Location = New System.Drawing.Point(599, 23)
        Me.BTNPriceStructureSearch.Name = "BTNPriceStructureSearch"
        Me.BTNPriceStructureSearch.Size = New System.Drawing.Size(96, 47)
        Me.BTNPriceStructureSearch.TabIndex = 2
        Me.BTNPriceStructureSearch.Text = "ค้นหาเอกสาร"
        Me.BTNPriceStructureSearch.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNPriceStructureSearch.UseVisualStyleBackColor = True
        '
        'TextPriceStructureSearch
        '
        Me.TextPriceStructureSearch.Location = New System.Drawing.Point(216, 37)
        Me.TextPriceStructureSearch.Name = "TextPriceStructureSearch"
        Me.TextPriceStructureSearch.Size = New System.Drawing.Size(380, 20)
        Me.TextPriceStructureSearch.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(71, 40)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(139, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "ค้นหาเอกสารตามข้อความนี้ :"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(9, 242)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(79, 13)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "รายการสินค้า"
        '
        'BTNClose
        '
        Me.BTNClose.BackColor = System.Drawing.Color.WhiteSmoke
        Me.BTNClose.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_73
        Me.BTNClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNClose.Location = New System.Drawing.Point(885, 623)
        Me.BTNClose.Name = "BTNClose"
        Me.BTNClose.Size = New System.Drawing.Size(120, 52)
        Me.BTNClose.TabIndex = 22
        Me.BTNClose.Text = "ออก"
        Me.BTNClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNClose.UseVisualStyleBackColor = False
        '
        'BTNPrint
        '
        Me.BTNPrint.BackColor = System.Drawing.Color.WhiteSmoke
        Me.BTNPrint.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_22
        Me.BTNPrint.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNPrint.Location = New System.Drawing.Point(507, 623)
        Me.BTNPrint.Name = "BTNPrint"
        Me.BTNPrint.Size = New System.Drawing.Size(120, 52)
        Me.BTNPrint.TabIndex = 21
        Me.BTNPrint.Text = "พิมพ์เอกสาร"
        Me.BTNPrint.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNPrint.UseVisualStyleBackColor = False
        '
        'BTNClearData
        '
        Me.BTNClearData.BackColor = System.Drawing.Color.WhiteSmoke
        Me.BTNClearData.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_071
        Me.BTNClearData.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNClearData.Location = New System.Drawing.Point(633, 623)
        Me.BTNClearData.Name = "BTNClearData"
        Me.BTNClearData.Size = New System.Drawing.Size(120, 52)
        Me.BTNClearData.TabIndex = 20
        Me.BTNClearData.Text = "เคลียร์ข้อมูล"
        Me.BTNClearData.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNClearData.UseVisualStyleBackColor = False
        '
        'PBConfirm
        '
        Me.PBConfirm.Image = Global.NPWindowApp.My.Resources.Resources.Confirm
        Me.PBConfirm.Location = New System.Drawing.Point(12, 18)
        Me.PBConfirm.Name = "PBConfirm"
        Me.PBConfirm.Size = New System.Drawing.Size(38, 20)
        Me.PBConfirm.TabIndex = 18
        Me.PBConfirm.TabStop = False
        Me.PBConfirm.Visible = False
        '
        'PBNew
        '
        Me.PBNew.Image = Global.NPWindowApp.My.Resources.Resources._New
        Me.PBNew.Location = New System.Drawing.Point(12, 18)
        Me.PBNew.Name = "PBNew"
        Me.PBNew.Size = New System.Drawing.Size(38, 20)
        Me.PBNew.TabIndex = 17
        Me.PBNew.TabStop = False
        Me.PBNew.Visible = False
        '
        'BTNFind
        '
        Me.BTNFind.BackgroundImage = Global.NPWindowApp.My.Resources.Resources.find
        Me.BTNFind.Location = New System.Drawing.Point(881, 46)
        Me.BTNFind.Name = "BTNFind"
        Me.BTNFind.Size = New System.Drawing.Size(58, 26)
        Me.BTNFind.TabIndex = 15
        Me.BTNFind.UseVisualStyleBackColor = True
        '
        'BTNSearchDocno
        '
        Me.BTNSearchDocno.BackColor = System.Drawing.Color.WhiteSmoke
        Me.BTNSearchDocno.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP__My_eBooks_Folder_
        Me.BTNSearchDocno.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNSearchDocno.Location = New System.Drawing.Point(759, 623)
        Me.BTNSearchDocno.Name = "BTNSearchDocno"
        Me.BTNSearchDocno.Size = New System.Drawing.Size(120, 52)
        Me.BTNSearchDocno.TabIndex = 14
        Me.BTNSearchDocno.Text = "เลือกอนุมัติเอกสาร"
        Me.BTNSearchDocno.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNSearchDocno.UseVisualStyleBackColor = False
        '
        'BTNSave
        '
        Me.BTNSave.BackColor = System.Drawing.Color.WhiteSmoke
        Me.BTNSave.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_48
        Me.BTNSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNSave.Location = New System.Drawing.Point(381, 623)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(120, 52)
        Me.BTNSave.TabIndex = 5
        Me.BTNSave.Text = "บันทึกข้อมูล"
        Me.BTNSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTNSave.UseVisualStyleBackColor = False
        '
        'FormImportDataPriceStructureFromExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        Me.BackColor = System.Drawing.SystemColors.ControlDark
        Me.ClientSize = New System.Drawing.Size(1016, 734)
        Me.Controls.Add(Me.GBSearchPriceStructure)
        Me.Controls.Add(Me.BTNClose)
        Me.Controls.Add(Me.BTNPrint)
        Me.Controls.Add(Me.BTNClearData)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.PBConfirm)
        Me.Controls.Add(Me.PBNew)
        Me.Controls.Add(Me.BTNFind)
        Me.Controls.Add(Me.BTNSearchDocno)
        Me.Controls.Add(Me.DocDate)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextDocNo)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextMyDescription)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.PB101)
        Me.Controls.Add(Me.BTNGenData)
        Me.Controls.Add(Me.BTNSave)
        Me.Controls.Add(Me.DGDetails)
        Me.Controls.Add(Me.DGHeader)
        Me.Controls.Add(Me.LBLFileName)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FormImportDataPriceStructureFromExcel"
        Me.Text = "ดึงข้อมูลโครงสร้างราคาจาก Excel เข้าฐานข้อมูลเพื่อทำการปรับราคา"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.DGHeader, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GBSearchPriceStructure.ResumeLayout(False)
        Me.GBSearchPriceStructure.PerformLayout()
        CType(Me.PBConfirm, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PBNew, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LBLFileName As System.Windows.Forms.Label
    Friend WithEvents DGHeader As System.Windows.Forms.DataGridView
    Friend WithEvents DGDetails As System.Windows.Forms.DataGridView
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents BTNGenData As System.Windows.Forms.Button
    Friend WithEvents PB101 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextMyDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextDocNo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DocDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents BTNSearchDocno As System.Windows.Forms.Button
    Friend WithEvents BTNFind As System.Windows.Forms.Button
    Friend WithEvents GBSearchPriceStructure As System.Windows.Forms.GroupBox
    Friend WithEvents BTNPriceStructureSearch As System.Windows.Forms.Button
    Friend WithEvents TextPriceStructureSearch As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents BTNPriceStructureExit As System.Windows.Forms.Button
    Friend WithEvents BTNPriceStructureConfirm As System.Windows.Forms.Button
    Friend WithEvents ListViewPriceStructure As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents PBNew As System.Windows.Forms.PictureBox
    Friend WithEvents PBConfirm As System.Windows.Forms.PictureBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents BTNClearData As System.Windows.Forms.Button
    Friend WithEvents BTNPrint As System.Windows.Forms.Button
    Friend WithEvents BTNClose As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
End Class
