<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPriceVolumeSet
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
        Me.P01 = New System.Windows.Forms.Panel
        Me.LPcbx = New System.Windows.Forms.CheckBox
        Me.BtnGenVLM = New System.Windows.Forms.Button
        Me.btnSaveAS = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnSearch = New System.Windows.Forms.Button
        Me.PBNew = New System.Windows.Forms.PictureBox
        Me.PBConfirm = New System.Windows.Forms.PictureBox
        Me.gvDetail = New System.Windows.Forms.DataGridView
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnProduct = New System.Windows.Forms.Button
        Me.DTPEndDate = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.DTPStartDate = New System.Windows.Forms.DateTimePicker
        Me.smpLV1 = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.DTPdocDate = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnNewDoc = New System.Windows.Forms.Button
        Me.txtDocno = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtSMTP3 = New System.Windows.Forms.TextBox
        Me.txtSMTP2 = New System.Windows.Forms.TextBox
        Me.TextBox12 = New System.Windows.Forms.TextBox
        Me.txtDC3 = New System.Windows.Forms.TextBox
        Me.txtDC2 = New System.Windows.Forms.TextBox
        Me.TextBox9 = New System.Windows.Forms.TextBox
        Me.txtVM3 = New System.Windows.Forms.TextBox
        Me.txtVM2 = New System.Windows.Forms.TextBox
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.txtLv3 = New System.Windows.Forms.TextBox
        Me.txtLv2 = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.okFind = New System.Windows.Forms.Button
        Me.P02 = New System.Windows.Forms.Panel
        Me.GB03 = New System.Windows.Forms.GroupBox
        Me.RdoManual = New System.Windows.Forms.RadioButton
        Me.RdoPSDoc = New System.Windows.Forms.RadioButton
        Me.pgbItem = New System.Windows.Forms.ProgressBar
        Me.btnexitP2 = New System.Windows.Forms.Button
        Me.ckeckedAll = New System.Windows.Forms.CheckBox
        Me.btnSelect = New System.Windows.Forms.Button
        Me.LvProduct = New System.Windows.Forms.ListView
        Me.hdPSDocno = New System.Windows.Forms.ColumnHeader
        Me.ItemCode = New System.Windows.Forms.ColumnHeader
        Me.ItemName = New System.Windows.Forms.ColumnHeader
        Me.UnitCode = New System.Windows.Forms.ColumnHeader
        Me.Price1 = New System.Windows.Forms.ColumnHeader
        Me.MKbudget = New System.Windows.Forms.ColumnHeader
        Me.BudgetAvgLot = New System.Windows.Forms.ColumnHeader
        Me.GB01 = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cbxCategory = New System.Windows.Forms.ComboBox
        Me.cbxProductType = New System.Windows.Forms.ComboBox
        Me.cbxBrand = New System.Windows.Forms.ComboBox
        Me.cbxDepartment = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.btnGenerate = New System.Windows.Forms.Button
        Me.P01.SuspendLayout()
        CType(Me.PBNew, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PBConfirm, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.smpLV1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.P02.SuspendLayout()
        Me.GB03.SuspendLayout()
        Me.GB01.SuspendLayout()
        Me.SuspendLayout()
        '
        'P01
        '
        Me.P01.Controls.Add(Me.LPcbx)
        Me.P01.Controls.Add(Me.BtnGenVLM)
        Me.P01.Controls.Add(Me.btnSaveAS)
        Me.P01.Controls.Add(Me.btnExit)
        Me.P01.Controls.Add(Me.btnPrint)
        Me.P01.Controls.Add(Me.btnSearch)
        Me.P01.Controls.Add(Me.PBNew)
        Me.P01.Controls.Add(Me.PBConfirm)
        Me.P01.Controls.Add(Me.gvDetail)
        Me.P01.Controls.Add(Me.btnCancel)
        Me.P01.Controls.Add(Me.btnSave)
        Me.P01.Controls.Add(Me.btnProduct)
        Me.P01.Controls.Add(Me.DTPEndDate)
        Me.P01.Controls.Add(Me.Label5)
        Me.P01.Controls.Add(Me.Label4)
        Me.P01.Controls.Add(Me.DTPStartDate)
        Me.P01.Controls.Add(Me.smpLV1)
        Me.P01.Controls.Add(Me.Label3)
        Me.P01.Controls.Add(Me.DTPdocDate)
        Me.P01.Controls.Add(Me.Label2)
        Me.P01.Controls.Add(Me.btnNewDoc)
        Me.P01.Controls.Add(Me.txtDocno)
        Me.P01.Controls.Add(Me.Label1)
        Me.P01.Controls.Add(Me.GroupBox1)
        Me.P01.Controls.Add(Me.okFind)
        Me.P01.Location = New System.Drawing.Point(12, 12)
        Me.P01.Name = "P01"
        Me.P01.Size = New System.Drawing.Size(986, 624)
        Me.P01.TabIndex = 0
        '
        'LPcbx
        '
        Me.LPcbx.AutoSize = True
        Me.LPcbx.ForeColor = System.Drawing.Color.White
        Me.LPcbx.Location = New System.Drawing.Point(75, 133)
        Me.LPcbx.Name = "LPcbx"
        Me.LPcbx.Size = New System.Drawing.Size(96, 17)
        Me.LPcbx.TabIndex = 62
        Me.LPcbx.Text = "ลดจากราคา LP"
        Me.LPcbx.UseVisualStyleBackColor = True
        '
        'BtnGenVLM
        '
        Me.BtnGenVLM.Location = New System.Drawing.Point(682, 130)
        Me.BtnGenVLM.Name = "BtnGenVLM"
        Me.BtnGenVLM.Size = New System.Drawing.Size(112, 26)
        Me.BtnGenVLM.TabIndex = 61
        Me.BtnGenVLM.Text = "Generate"
        Me.BtnGenVLM.UseVisualStyleBackColor = True
        '
        'btnSaveAS
        '
        Me.btnSaveAS.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP__My_eBooks_Folder_
        Me.btnSaveAS.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnSaveAS.Location = New System.Drawing.Point(534, 571)
        Me.btnSaveAS.Name = "btnSaveAS"
        Me.btnSaveAS.Size = New System.Drawing.Size(75, 48)
        Me.btnSaveAS.TabIndex = 60
        Me.btnSaveAS.Text = "บันทึกเป็น"
        Me.btnSaveAS.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnSaveAS.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Image = Global.NPWindowApp.My.Resources.Resources.close1
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnExit.Location = New System.Drawing.Point(900, 571)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(77, 49)
        Me.btnExit.TabIndex = 18
        Me.btnExit.Text = "ออก"
        Me.btnExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_38
        Me.btnPrint.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnPrint.Location = New System.Drawing.Point(607, 571)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 49)
        Me.btnPrint.TabIndex = 14
        Me.btnPrint.Text = "พิมพ์"
        Me.btnPrint.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnSearch
        '
        Me.btnSearch.Image = Global.NPWindowApp.My.Resources.Resources.Windows_Explorer
        Me.btnSearch.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnSearch.Location = New System.Drawing.Point(679, 571)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(75, 49)
        Me.btnSearch.TabIndex = 15
        Me.btnSearch.Text = "ค้นหา"
        Me.btnSearch.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'PBNew
        '
        Me.PBNew.Image = Global.NPWindowApp.My.Resources.Resources._New
        Me.PBNew.Location = New System.Drawing.Point(10, 8)
        Me.PBNew.Name = "PBNew"
        Me.PBNew.Size = New System.Drawing.Size(38, 20)
        Me.PBNew.TabIndex = 55
        Me.PBNew.TabStop = False
        Me.PBNew.Visible = False
        '
        'PBConfirm
        '
        Me.PBConfirm.Image = Global.NPWindowApp.My.Resources.Resources.Confirm
        Me.PBConfirm.Location = New System.Drawing.Point(10, 8)
        Me.PBConfirm.Name = "PBConfirm"
        Me.PBConfirm.Size = New System.Drawing.Size(38, 20)
        Me.PBConfirm.TabIndex = 56
        Me.PBConfirm.TabStop = False
        Me.PBConfirm.Visible = False
        '
        'gvDetail
        '
        Me.gvDetail.AllowUserToAddRows = False
        Me.gvDetail.AllowUserToDeleteRows = False
        Me.gvDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gvDetail.Location = New System.Drawing.Point(10, 162)
        Me.gvDetail.Name = "gvDetail"
        Me.gvDetail.Size = New System.Drawing.Size(967, 406)
        Me.gvDetail.TabIndex = 54
        '
        'btnCancel
        '
        Me.btnCancel.Image = Global.NPWindowApp.My.Resources.Resources._2
        Me.btnCancel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnCancel.Location = New System.Drawing.Point(824, 571)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(81, 49)
        Me.btnCancel.TabIndex = 17
        Me.btnCancel.Text = "New"
        Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_48
        Me.btnSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnSave.Location = New System.Drawing.Point(749, 571)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(79, 49)
        Me.btnSave.TabIndex = 15
        Me.btnSave.Text = "บันทึก"
        Me.btnSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnProduct
        '
        Me.btnProduct.Image = Global.NPWindowApp.My.Resources.Resources.Gen
        Me.btnProduct.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnProduct.Location = New System.Drawing.Point(792, 130)
        Me.btnProduct.Name = "btnProduct"
        Me.btnProduct.Size = New System.Drawing.Size(102, 26)
        Me.btnProduct.TabIndex = 9
        Me.btnProduct.Text = "กำหนดสินค้า"
        Me.btnProduct.UseVisualStyleBackColor = True
        '
        'DTPEndDate
        '
        Me.DTPEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPEndDate.Location = New System.Drawing.Point(499, 130)
        Me.DTPEndDate.Name = "DTPEndDate"
        Me.DTPEndDate.Size = New System.Drawing.Size(100, 20)
        Me.DTPEndDate.TabIndex = 13
        Me.DTPEndDate.Value = New Date(2009, 3, 1, 9, 55, 0, 0)
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(457, 133)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 13)
        Me.Label5.TabIndex = 49
        Me.Label5.Text = "ถึงวันที่"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(256, 133)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(84, 13)
        Me.Label4.TabIndex = 48
        Me.Label4.Text = "ให้ปรับราคาวันที่"
        '
        'DTPStartDate
        '
        Me.DTPStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPStartDate.Location = New System.Drawing.Point(341, 130)
        Me.DTPStartDate.Name = "DTPStartDate"
        Me.DTPStartDate.Size = New System.Drawing.Size(104, 20)
        Me.DTPStartDate.TabIndex = 12
        Me.DTPStartDate.Value = New Date(2009, 3, 1, 9, 55, 0, 0)
        '
        'smpLV1
        '
        Me.smpLV1.DecimalPlaces = 2
        Me.smpLV1.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.smpLV1.Location = New System.Drawing.Point(841, 12)
        Me.smpLV1.Name = "smpLV1"
        Me.smpLV1.Size = New System.Drawing.Size(52, 20)
        Me.smpLV1.TabIndex = 4
        Me.smpLV1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(715, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(123, 13)
        Me.Label3.TabIndex = 45
        Me.Label3.Text = "% SmartPoit ราคาระดับ 1"
        '
        'DTPdocDate
        '
        Me.DTPdocDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPdocDate.Location = New System.Drawing.Point(520, 10)
        Me.DTPdocDate.Name = "DTPdocDate"
        Me.DTPdocDate.Size = New System.Drawing.Size(103, 20)
        Me.DTPdocDate.TabIndex = 3
        Me.DTPdocDate.Value = New Date(2009, 2, 16, 0, 0, 0, 0)
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(450, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 43
        Me.Label2.Text = "วันที่เอกสาร"
        '
        'btnNewDoc
        '
        Me.btnNewDoc.Image = Global.NPWindowApp.My.Resources.Resources._2
        Me.btnNewDoc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnNewDoc.Location = New System.Drawing.Point(277, 8)
        Me.btnNewDoc.Name = "btnNewDoc"
        Me.btnNewDoc.Size = New System.Drawing.Size(59, 26)
        Me.btnNewDoc.TabIndex = 2
        Me.btnNewDoc.Text = "New!"
        Me.btnNewDoc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnNewDoc.UseVisualStyleBackColor = True
        '
        'txtDocno
        '
        Me.txtDocno.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtDocno.Location = New System.Drawing.Point(133, 8)
        Me.txtDocno.Name = "txtDocno"
        Me.txtDocno.Size = New System.Drawing.Size(140, 26)
        Me.txtDocno.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(61, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 40
        Me.Label1.Text = "เลขที่เอกสาร"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtSMTP3)
        Me.GroupBox1.Controls.Add(Me.txtSMTP2)
        Me.GroupBox1.Controls.Add(Me.TextBox12)
        Me.GroupBox1.Controls.Add(Me.txtDC3)
        Me.GroupBox1.Controls.Add(Me.txtDC2)
        Me.GroupBox1.Controls.Add(Me.TextBox9)
        Me.GroupBox1.Controls.Add(Me.txtVM3)
        Me.GroupBox1.Controls.Add(Me.txtVM2)
        Me.GroupBox1.Controls.Add(Me.TextBox6)
        Me.GroupBox1.Controls.Add(Me.txtLv3)
        Me.GroupBox1.Controls.Add(Me.txtLv2)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(75, 34)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(818, 92)
        Me.GroupBox1.TabIndex = 39
        Me.GroupBox1.TabStop = False
        '
        'txtSMTP3
        '
        Me.txtSMTP3.Location = New System.Drawing.Point(607, 57)
        Me.txtSMTP3.Name = "txtSMTP3"
        Me.txtSMTP3.Size = New System.Drawing.Size(199, 20)
        Me.txtSMTP3.TabIndex = 11
        Me.txtSMTP3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtSMTP2
        '
        Me.txtSMTP2.Location = New System.Drawing.Point(607, 38)
        Me.txtSMTP2.Name = "txtSMTP2"
        Me.txtSMTP2.Size = New System.Drawing.Size(199, 20)
        Me.txtSMTP2.TabIndex = 10
        Me.txtSMTP2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox12
        '
        Me.TextBox12.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.TextBox12.ForeColor = System.Drawing.Color.Black
        Me.TextBox12.Location = New System.Drawing.Point(607, 19)
        Me.TextBox12.Name = "TextBox12"
        Me.TextBox12.Size = New System.Drawing.Size(199, 20)
        Me.TextBox12.TabIndex = 58
        Me.TextBox12.Text = "% Smart Point"
        Me.TextBox12.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDC3
        '
        Me.txtDC3.Location = New System.Drawing.Point(399, 57)
        Me.txtDC3.Name = "txtDC3"
        Me.txtDC3.Size = New System.Drawing.Size(209, 20)
        Me.txtDC3.TabIndex = 8
        Me.txtDC3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDC2
        '
        Me.txtDC2.Location = New System.Drawing.Point(399, 38)
        Me.txtDC2.Name = "txtDC2"
        Me.txtDC2.Size = New System.Drawing.Size(209, 20)
        Me.txtDC2.TabIndex = 7
        Me.txtDC2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox9
        '
        Me.TextBox9.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.TextBox9.ForeColor = System.Drawing.Color.Black
        Me.TextBox9.Location = New System.Drawing.Point(399, 19)
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.ReadOnly = True
        Me.TextBox9.Size = New System.Drawing.Size(209, 20)
        Me.TextBox9.TabIndex = 57
        Me.TextBox9.Text = "ส่วนลดจากราคาที่ 1"
        Me.TextBox9.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtVM3
        '
        Me.txtVM3.Location = New System.Drawing.Point(199, 57)
        Me.txtVM3.Name = "txtVM3"
        Me.txtVM3.Size = New System.Drawing.Size(201, 20)
        Me.txtVM3.TabIndex = 6
        Me.txtVM3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtVM2
        '
        Me.txtVM2.Location = New System.Drawing.Point(199, 38)
        Me.txtVM2.Name = "txtVM2"
        Me.txtVM2.Size = New System.Drawing.Size(201, 20)
        Me.txtVM2.TabIndex = 5
        Me.txtVM2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox6
        '
        Me.TextBox6.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.TextBox6.ForeColor = System.Drawing.Color.Black
        Me.TextBox6.Location = New System.Drawing.Point(199, 19)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(201, 20)
        Me.TextBox6.TabIndex = 56
        Me.TextBox6.Text = "Volume"
        Me.TextBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtLv3
        '
        Me.txtLv3.Location = New System.Drawing.Point(21, 57)
        Me.txtLv3.Name = "txtLv3"
        Me.txtLv3.Size = New System.Drawing.Size(179, 20)
        Me.txtLv3.TabIndex = 2
        Me.txtLv3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtLv2
        '
        Me.txtLv2.Location = New System.Drawing.Point(21, 38)
        Me.txtLv2.Name = "txtLv2"
        Me.txtLv2.Size = New System.Drawing.Size(179, 20)
        Me.txtLv2.TabIndex = 1
        Me.txtLv2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.TextBox1.ForeColor = System.Drawing.Color.Black
        Me.TextBox1.Location = New System.Drawing.Point(21, 19)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(179, 20)
        Me.TextBox1.TabIndex = 55
        Me.TextBox1.Text = "ระดับราคา"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'okFind
        '
        Me.okFind.Image = Global.NPWindowApp.My.Resources.Resources.icon_16_checkin
        Me.okFind.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.okFind.Location = New System.Drawing.Point(275, 6)
        Me.okFind.Name = "okFind"
        Me.okFind.Size = New System.Drawing.Size(61, 29)
        Me.okFind.TabIndex = 59
        Me.okFind.Text = "ตกลง"
        Me.okFind.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.okFind.UseVisualStyleBackColor = True
        '
        'P02
        '
        Me.P02.BackColor = System.Drawing.SystemColors.Desktop
        Me.P02.Controls.Add(Me.GB03)
        Me.P02.Controls.Add(Me.pgbItem)
        Me.P02.Controls.Add(Me.btnexitP2)
        Me.P02.Controls.Add(Me.ckeckedAll)
        Me.P02.Controls.Add(Me.btnSelect)
        Me.P02.Controls.Add(Me.LvProduct)
        Me.P02.Controls.Add(Me.GB01)
        Me.P02.Controls.Add(Me.btnGenerate)
        Me.P02.Location = New System.Drawing.Point(12, 12)
        Me.P02.Name = "P02"
        Me.P02.Size = New System.Drawing.Size(986, 624)
        Me.P02.TabIndex = 57
        '
        'GB03
        '
        Me.GB03.Controls.Add(Me.RdoManual)
        Me.GB03.Controls.Add(Me.RdoPSDoc)
        Me.GB03.ForeColor = System.Drawing.Color.White
        Me.GB03.Location = New System.Drawing.Point(15, 11)
        Me.GB03.Name = "GB03"
        Me.GB03.Size = New System.Drawing.Size(216, 62)
        Me.GB03.TabIndex = 16
        Me.GB03.TabStop = False
        Me.GB03.Text = "เลือกรายการ"
        '
        'RdoManual
        '
        Me.RdoManual.AutoSize = True
        Me.RdoManual.Location = New System.Drawing.Point(49, 38)
        Me.RdoManual.Name = "RdoManual"
        Me.RdoManual.Size = New System.Drawing.Size(75, 17)
        Me.RdoManual.TabIndex = 1
        Me.RdoManual.TabStop = True
        Me.RdoManual.Text = "กำหนดเอง"
        Me.RdoManual.UseVisualStyleBackColor = True
        '
        'RdoPSDoc
        '
        Me.RdoPSDoc.AutoSize = True
        Me.RdoPSDoc.Location = New System.Drawing.Point(49, 19)
        Me.RdoPSDoc.Name = "RdoPSDoc"
        Me.RdoPSDoc.Size = New System.Drawing.Size(128, 17)
        Me.RdoPSDoc.TabIndex = 0
        Me.RdoPSDoc.TabStop = True
        Me.RdoPSDoc.Text = "เอกสารโครงสร้างราคา"
        Me.RdoPSDoc.UseVisualStyleBackColor = True
        '
        'pgbItem
        '
        Me.pgbItem.Location = New System.Drawing.Point(104, 79)
        Me.pgbItem.Name = "pgbItem"
        Me.pgbItem.Size = New System.Drawing.Size(747, 35)
        Me.pgbItem.TabIndex = 15
        '
        'btnexitP2
        '
        Me.btnexitP2.Location = New System.Drawing.Point(900, 617)
        Me.btnexitP2.Name = "btnexitP2"
        Me.btnexitP2.Size = New System.Drawing.Size(77, 35)
        Me.btnexitP2.TabIndex = 14
        Me.btnexitP2.Text = "ยกเลิก"
        Me.btnexitP2.UseVisualStyleBackColor = True
        '
        'ckeckedAll
        '
        Me.ckeckedAll.AutoSize = True
        Me.ckeckedAll.ForeColor = System.Drawing.Color.White
        Me.ckeckedAll.Location = New System.Drawing.Point(15, 103)
        Me.ckeckedAll.Name = "ckeckedAll"
        Me.ckeckedAll.Size = New System.Drawing.Size(83, 17)
        Me.ckeckedAll.TabIndex = 13
        Me.ckeckedAll.Text = "เลือกทั้งหมด"
        Me.ckeckedAll.UseVisualStyleBackColor = True
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(820, 616)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(78, 37)
        Me.btnSelect.TabIndex = 10
        Me.btnSelect.Text = "ตกลง"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'LvProduct
        '
        Me.LvProduct.CheckBoxes = True
        Me.LvProduct.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.hdPSDocno, Me.ItemCode, Me.ItemName, Me.UnitCode, Me.Price1, Me.MKbudget, Me.BudgetAvgLot})
        Me.LvProduct.GridLines = True
        Me.LvProduct.Location = New System.Drawing.Point(10, 124)
        Me.LvProduct.Name = "LvProduct"
        Me.LvProduct.Size = New System.Drawing.Size(967, 487)
        Me.LvProduct.TabIndex = 9
        Me.LvProduct.UseCompatibleStateImageBehavior = False
        Me.LvProduct.View = System.Windows.Forms.View.Details
        '
        'hdPSDocno
        '
        Me.hdPSDocno.Text = "เลขที่เอกสารโครงสร้างราคา"
        Me.hdPSDocno.Width = 149
        '
        'ItemCode
        '
        Me.ItemCode.Text = "รหัสสินค้า"
        Me.ItemCode.Width = 122
        '
        'ItemName
        '
        Me.ItemName.Text = "ชื่อสินค้า"
        Me.ItemName.Width = 296
        '
        'UnitCode
        '
        Me.UnitCode.Text = "หน่วยขาย"
        Me.UnitCode.Width = 78
        '
        'Price1
        '
        Me.Price1.Text = "ราคาที่1"
        Me.Price1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.Price1.Width = 88
        '
        'MKbudget
        '
        Me.MKbudget.Text = "ทุนตลาดSaleVat"
        Me.MKbudget.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.MKbudget.Width = 106
        '
        'BudgetAvgLot
        '
        Me.BudgetAvgLot.Text = "ทุนเฉลี่ยตาม LotSaleVat"
        Me.BudgetAvgLot.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.BudgetAvgLot.Width = 117
        '
        'GB01
        '
        Me.GB01.BackColor = System.Drawing.SystemColors.Desktop
        Me.GB01.Controls.Add(Me.Label6)
        Me.GB01.Controls.Add(Me.cbxCategory)
        Me.GB01.Controls.Add(Me.cbxProductType)
        Me.GB01.Controls.Add(Me.cbxBrand)
        Me.GB01.Controls.Add(Me.cbxDepartment)
        Me.GB01.Controls.Add(Me.Label7)
        Me.GB01.Controls.Add(Me.Label8)
        Me.GB01.Controls.Add(Me.Label9)
        Me.GB01.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.GB01.Location = New System.Drawing.Point(237, 12)
        Me.GB01.Name = "GB01"
        Me.GB01.Size = New System.Drawing.Size(740, 61)
        Me.GB01.TabIndex = 7
        Me.GB01.TabStop = False
        Me.GB01.Text = "เงื่อนไขสินค้า"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(563, 29)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 13)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Categories"
        '
        'cbxCategory
        '
        Me.cbxCategory.FormattingEnabled = True
        Me.cbxCategory.Location = New System.Drawing.Point(623, 24)
        Me.cbxCategory.Name = "cbxCategory"
        Me.cbxCategory.Size = New System.Drawing.Size(113, 21)
        Me.cbxCategory.TabIndex = 7
        '
        'cbxProductType
        '
        Me.cbxProductType.FormattingEnabled = True
        Me.cbxProductType.Location = New System.Drawing.Point(416, 24)
        Me.cbxProductType.Name = "cbxProductType"
        Me.cbxProductType.Size = New System.Drawing.Size(141, 21)
        Me.cbxProductType.TabIndex = 6
        Me.cbxProductType.Text = "All"
        '
        'cbxBrand
        '
        Me.cbxBrand.FormattingEnabled = True
        Me.cbxBrand.Location = New System.Drawing.Point(225, 23)
        Me.cbxBrand.Name = "cbxBrand"
        Me.cbxBrand.Size = New System.Drawing.Size(136, 21)
        Me.cbxBrand.TabIndex = 5
        Me.cbxBrand.Text = "All"
        '
        'cbxDepartment
        '
        Me.cbxDepartment.FormattingEnabled = True
        Me.cbxDepartment.Location = New System.Drawing.Point(67, 22)
        Me.cbxDepartment.Name = "cbxDepartment"
        Me.cbxDepartment.Size = New System.Drawing.Size(130, 21)
        Me.cbxDepartment.TabIndex = 4
        Me.cbxDepartment.Text = "All"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(362, 28)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(55, 13)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "ชนิดสินค้า"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(199, 27)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(27, 13)
        Me.Label8.TabIndex = 2
        Me.Label8.Text = "ยี่ห้อ"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 26)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(62, 13)
        Me.Label9.TabIndex = 1
        Me.Label9.Text = "Department"
        '
        'btnGenerate
        '
        Me.btnGenerate.Image = Global.NPWindowApp.My.Resources.Resources.TOON_XP_Icons__V1c__Icon_73
        Me.btnGenerate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnGenerate.Location = New System.Drawing.Point(857, 77)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(120, 41)
        Me.btnGenerate.TabIndex = 8
        Me.btnGenerate.Text = "Generate Data"
        Me.btnGenerate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'frmPriceVolumeSet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        Me.BackColor = System.Drawing.SystemColors.Desktop
        Me.ClientSize = New System.Drawing.Size(1004, 650)
        Me.Controls.Add(Me.P01)
        Me.Controls.Add(Me.P02)
        Me.Name = "frmPriceVolumeSet"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.P01.ResumeLayout(False)
        Me.P01.PerformLayout()
        CType(Me.PBNew, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PBConfirm, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvDetail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.smpLV1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.P02.ResumeLayout(False)
        Me.P02.PerformLayout()
        Me.GB03.ResumeLayout(False)
        Me.GB03.PerformLayout()
        Me.GB01.ResumeLayout(False)
        Me.GB01.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents P01 As System.Windows.Forms.Panel
    Friend WithEvents PBNew As System.Windows.Forms.PictureBox
    Friend WithEvents gvDetail As System.Windows.Forms.DataGridView
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnProduct As System.Windows.Forms.Button
    Friend WithEvents DTPEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DTPStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents smpLV1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DTPdocDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnNewDoc As System.Windows.Forms.Button
    Friend WithEvents txtDocno As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents PBConfirm As System.Windows.Forms.PictureBox
    Friend WithEvents P02 As System.Windows.Forms.Panel
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    Friend WithEvents GB01 As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cbxCategory As System.Windows.Forms.ComboBox
    Friend WithEvents cbxProductType As System.Windows.Forms.ComboBox
    Friend WithEvents cbxBrand As System.Windows.Forms.ComboBox
    Friend WithEvents cbxDepartment As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents LvProduct As System.Windows.Forms.ListView
    Friend WithEvents ItemCode As System.Windows.Forms.ColumnHeader
    Friend WithEvents ItemName As System.Windows.Forms.ColumnHeader
    Friend WithEvents UnitCode As System.Windows.Forms.ColumnHeader
    Friend WithEvents Price1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents MKbudget As System.Windows.Forms.ColumnHeader
    Friend WithEvents BudgetAvgLot As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtSMTP3 As System.Windows.Forms.TextBox
    Friend WithEvents txtSMTP2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox12 As System.Windows.Forms.TextBox
    Friend WithEvents txtDC3 As System.Windows.Forms.TextBox
    Friend WithEvents txtDC2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox
    Friend WithEvents txtVM3 As System.Windows.Forms.TextBox
    Friend WithEvents txtVM2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents txtLv3 As System.Windows.Forms.TextBox
    Friend WithEvents txtLv2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ckeckedAll As System.Windows.Forms.CheckBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents okFind As System.Windows.Forms.Button
    Friend WithEvents btnexitP2 As System.Windows.Forms.Button
    Friend WithEvents pgbItem As System.Windows.Forms.ProgressBar
    Friend WithEvents hdPSDocno As System.Windows.Forms.ColumnHeader
    Friend WithEvents GB03 As System.Windows.Forms.GroupBox
    Friend WithEvents RdoManual As System.Windows.Forms.RadioButton
    Friend WithEvents RdoPSDoc As System.Windows.Forms.RadioButton
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnSaveAS As System.Windows.Forms.Button
    Friend WithEvents BtnGenVLM As System.Windows.Forms.Button
    Friend WithEvents LPcbx As System.Windows.Forms.CheckBox
End Class
