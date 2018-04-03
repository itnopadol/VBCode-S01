<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class FormPOS
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
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.TBDocDate = New System.Windows.Forms.TextBox
        Me.TBDocNo = New System.Windows.Forms.TextBox
        Me.TBPrice = New System.Windows.Forms.TextBox
        Me.VScrollBar1 = New System.Windows.Forms.VScrollBar
        Me.TBBarCode = New System.Windows.Forms.TextBox
        Me.ListViewItem = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader14 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.TBUnit = New System.Windows.Forms.TextBox
        Me.TBQty = New System.Windows.Forms.TextBox
        Me.TBItemName = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.TBItemCode = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.TBRate1 = New System.Windows.Forms.TextBox
        Me.TBRemainQty = New System.Windows.Forms.TextBox
        Me.BTNSearch = New System.Windows.Forms.Button
        Me.BTNCancel = New System.Windows.Forms.Button
        Me.BTNNew = New System.Windows.Forms.Button
        Me.BTNExit = New System.Windows.Forms.Button
        Me.BTNSave = New System.Windows.Forms.Button
        Me.PNReceiveMoney = New System.Windows.Forms.Panel
        Me.ListViewPayDetails = New System.Windows.Forms.ListView
        Me.ColumnHeader12 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader13 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader15 = New System.Windows.Forms.ColumnHeader
        Me.BTNPayCancel = New System.Windows.Forms.Button
        Me.BTNPayOK = New System.Windows.Forms.Button
        Me.BTNCreditDelete = New System.Windows.Forms.Button
        Me.BTNCreditUpdate = New System.Windows.Forms.Button
        Me.TCPayMoney = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Label21 = New System.Windows.Forms.Label
        Me.TBOtherDebt = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.TBOtherExpense = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.TBOverMoney = New System.Windows.Forms.TextBox
        Me.TBOverMoneyInv = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.TBCashAmount = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.TBCreditTotalAmount = New System.Windows.Forms.TextBox
        Me.Label31 = New System.Windows.Forms.Label
        Me.TBChargeAmount = New System.Windows.Forms.TextBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.CMBBranch = New System.Windows.Forms.ComboBox
        Me.CMBCreditType = New System.Windows.Forms.ComboBox
        Me.CMBBank = New System.Windows.Forms.ComboBox
        Me.ListViewCreditCard = New System.Windows.Forms.ListView
        Me.ColumnHeader16 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader17 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader18 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader19 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader20 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader21 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader22 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader23 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader28 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader29 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader30 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader31 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader33 = New System.Windows.Forms.ColumnHeader
        Me.TBCreditAmount = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.TBConfirmNo = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.TBCreditCard = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.PNPOSConfig = New System.Windows.Forms.Panel
        Me.TBCharge = New System.Windows.Forms.TextBox
        Me.Label28 = New System.Windows.Forms.Label
        Me.CMBMemCalcStock = New System.Windows.Forms.ComboBox
        Me.TBMemTaxRate = New System.Windows.Forms.TextBox
        Me.CMBMemDepartment = New System.Windows.Forms.ComboBox
        Me.CMBMemPosID = New System.Windows.Forms.ComboBox
        Me.CMBMemShelfCode = New System.Windows.Forms.ComboBox
        Me.TBMemArName = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.CMBMemWHCode = New System.Windows.Forms.ComboBox
        Me.TBMemArCode = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TBWHCode = New System.Windows.Forms.TextBox
        Me.TBShelfCode = New System.Windows.Forms.TextBox
        Me.TBItemAmount = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.TBStkUnit = New System.Windows.Forms.TextBox
        Me.TBPayAmount = New System.Windows.Forms.TextBox
        Me.TBBalanceAmount = New System.Windows.Forms.TextBox
        Me.TBBillBalance = New System.Windows.Forms.TextBox
        Me.TBBillAmount = New System.Windows.Forms.TextBox
        Me.PNReceiveMoney.SuspendLayout()
        Me.TCPayMoney.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.PNPOSConfig.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Blue
        Me.Panel2.Location = New System.Drawing.Point(5, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(316, 3)
        '
        'TBDocDate
        '
        Me.TBDocDate.BackColor = System.Drawing.Color.PeachPuff
        Me.TBDocDate.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBDocDate.Location = New System.Drawing.Point(226, 8)
        Me.TBDocDate.Name = "TBDocDate"
        Me.TBDocDate.ReadOnly = True
        Me.TBDocDate.Size = New System.Drawing.Size(94, 19)
        Me.TBDocDate.TabIndex = 289
        '
        'TBDocNo
        '
        Me.TBDocNo.BackColor = System.Drawing.Color.PeachPuff
        Me.TBDocNo.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBDocNo.Location = New System.Drawing.Point(59, 8)
        Me.TBDocNo.Name = "TBDocNo"
        Me.TBDocNo.ReadOnly = True
        Me.TBDocNo.Size = New System.Drawing.Size(125, 19)
        Me.TBDocNo.TabIndex = 288
        '
        'TBPrice
        '
        Me.TBPrice.BackColor = System.Drawing.Color.PeachPuff
        Me.TBPrice.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBPrice.Location = New System.Drawing.Point(59, 71)
        Me.TBPrice.Name = "TBPrice"
        Me.TBPrice.ReadOnly = True
        Me.TBPrice.Size = New System.Drawing.Size(105, 19)
        Me.TBPrice.TabIndex = 273
        '
        'VScrollBar1
        '
        Me.VScrollBar1.Location = New System.Drawing.Point(11, 141)
        Me.VScrollBar1.Name = "VScrollBar1"
        Me.VScrollBar1.Size = New System.Drawing.Size(13, 76)
        Me.VScrollBar1.TabIndex = 290
        Me.VScrollBar1.Visible = False
        '
        'TBBarCode
        '
        Me.TBBarCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBBarCode.Location = New System.Drawing.Point(59, 29)
        Me.TBBarCode.Name = "TBBarCode"
        Me.TBBarCode.Size = New System.Drawing.Size(125, 19)
        Me.TBBarCode.TabIndex = 268
        '
        'ListViewItem
        '
        Me.ListViewItem.BackColor = System.Drawing.Color.FloralWhite
        Me.ListViewItem.Columns.Add(Me.ColumnHeader1)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader2)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader3)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader6)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader9)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader7)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader8)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader4)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader14)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader5)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader10)
        Me.ListViewItem.Columns.Add(Me.ColumnHeader11)
        Me.ListViewItem.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.ListViewItem.FullRowSelect = True
        Me.ListViewItem.Location = New System.Drawing.Point(6, 119)
        Me.ListViewItem.Name = "ListViewItem"
        Me.ListViewItem.Size = New System.Drawing.Size(314, 76)
        Me.ListViewItem.TabIndex = 285
        Me.ListViewItem.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ลำดับ"
        Me.ColumnHeader1.Width = 50
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "ชื่อสินค้า"
        Me.ColumnHeader2.Width = 120
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "จำนวน"
        Me.ColumnHeader3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader3.Width = 70
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "ราคา"
        Me.ColumnHeader6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader6.Width = 70
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "มูลค่า"
        Me.ColumnHeader9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader9.Width = 70
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "หน่วย"
        Me.ColumnHeader7.Width = 60
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "รหัส"
        Me.ColumnHeader8.Width = 60
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "บาร์"
        Me.ColumnHeader4.Width = 50
        '
        'ColumnHeader14
        '
        Me.ColumnHeader14.Text = "คลัง"
        Me.ColumnHeader14.Width = 80
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "ชั้นเก็บ"
        Me.ColumnHeader5.Width = 60
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "อัตราส่วน"
        Me.ColumnHeader10.Width = 60
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "Issave"
        Me.ColumnHeader11.Width = 60
        '
        'TBUnit
        '
        Me.TBUnit.BackColor = System.Drawing.Color.PeachPuff
        Me.TBUnit.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBUnit.Location = New System.Drawing.Point(226, 71)
        Me.TBUnit.Name = "TBUnit"
        Me.TBUnit.ReadOnly = True
        Me.TBUnit.Size = New System.Drawing.Size(94, 19)
        Me.TBUnit.TabIndex = 272
        '
        'TBQty
        '
        Me.TBQty.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBQty.Location = New System.Drawing.Point(226, 92)
        Me.TBQty.Name = "TBQty"
        Me.TBQty.Size = New System.Drawing.Size(94, 19)
        Me.TBQty.TabIndex = 284
        '
        'TBItemName
        '
        Me.TBItemName.BackColor = System.Drawing.Color.PeachPuff
        Me.TBItemName.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBItemName.Location = New System.Drawing.Point(59, 50)
        Me.TBItemName.Name = "TBItemName"
        Me.TBItemName.ReadOnly = True
        Me.TBItemName.Size = New System.Drawing.Size(261, 19)
        Me.TBItemName.TabIndex = 271
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label13.Location = New System.Drawing.Point(190, 30)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(36, 20)
        Me.Label13.Text = "รหัส :"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBItemCode
        '
        Me.TBItemCode.BackColor = System.Drawing.Color.PeachPuff
        Me.TBItemCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBItemCode.Location = New System.Drawing.Point(226, 29)
        Me.TBItemCode.Name = "TBItemCode"
        Me.TBItemCode.ReadOnly = True
        Me.TBItemCode.Size = New System.Drawing.Size(94, 19)
        Me.TBItemCode.TabIndex = 270
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label15.Location = New System.Drawing.Point(178, 72)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(47, 20)
        Me.Label15.Text = "หน่วย :"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.Blue
        Me.Panel5.Location = New System.Drawing.Point(6, 113)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(314, 3)
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(15, 30)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 20)
        Me.Label4.Text = "ยิงบาร์ :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label14.Location = New System.Drawing.Point(16, 51)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(44, 20)
        Me.Label14.Text = "ชื่อ :"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(10, 72)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 20)
        Me.Label1.Text = "ราคา :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(10, 93)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 20)
        Me.Label3.Text = "คงเหลือ :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label8.Location = New System.Drawing.Point(164, 93)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(61, 20)
        Me.Label8.Text = "ต้องการ :"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label19
        '
        Me.Label19.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label19.Location = New System.Drawing.Point(3, 9)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(57, 20)
        Me.Label19.Text = "เลขที่ :"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label20
        '
        Me.Label20.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label20.Location = New System.Drawing.Point(169, 9)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(57, 20)
        Me.Label20.Text = "วันที่ :"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBRate1
        '
        Me.TBRate1.BackColor = System.Drawing.Color.Gold
        Me.TBRate1.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBRate1.Location = New System.Drawing.Point(152, 92)
        Me.TBRate1.Name = "TBRate1"
        Me.TBRate1.ReadOnly = True
        Me.TBRate1.Size = New System.Drawing.Size(12, 19)
        Me.TBRate1.TabIndex = 311
        Me.TBRate1.Visible = False
        '
        'TBRemainQty
        '
        Me.TBRemainQty.BackColor = System.Drawing.Color.Gold
        Me.TBRemainQty.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBRemainQty.Location = New System.Drawing.Point(59, 92)
        Me.TBRemainQty.Name = "TBRemainQty"
        Me.TBRemainQty.ReadOnly = True
        Me.TBRemainQty.Size = New System.Drawing.Size(87, 19)
        Me.TBRemainQty.TabIndex = 278
        '
        'BTNSearch
        '
        Me.BTNSearch.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNSearch.Location = New System.Drawing.Point(133, 271)
        Me.BTNSearch.Name = "BTNSearch"
        Me.BTNSearch.Size = New System.Drawing.Size(59, 22)
        Me.BTNSearch.TabIndex = 314
        Me.BTNSearch.Text = "F6-ค้นหา"
        '
        'BTNCancel
        '
        Me.BTNCancel.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNCancel.Location = New System.Drawing.Point(197, 271)
        Me.BTNCancel.Name = "BTNCancel"
        Me.BTNCancel.Size = New System.Drawing.Size(59, 22)
        Me.BTNCancel.TabIndex = 315
        Me.BTNCancel.Text = "F8-ยกเลิก"
        '
        'BTNNew
        '
        Me.BTNNew.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNNew.Location = New System.Drawing.Point(6, 271)
        Me.BTNNew.Name = "BTNNew"
        Me.BTNNew.Size = New System.Drawing.Size(59, 22)
        Me.BTNNew.TabIndex = 312
        Me.BTNNew.Text = "F2-ทำใหม่"
        '
        'BTNExit
        '
        Me.BTNExit.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNExit.Location = New System.Drawing.Point(261, 271)
        Me.BTNExit.Name = "BTNExit"
        Me.BTNExit.Size = New System.Drawing.Size(59, 22)
        Me.BTNExit.TabIndex = 316
        Me.BTNExit.Text = "ESC-ออก"
        '
        'BTNSave
        '
        Me.BTNSave.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNSave.Location = New System.Drawing.Point(69, 271)
        Me.BTNSave.Name = "BTNSave"
        Me.BTNSave.Size = New System.Drawing.Size(59, 22)
        Me.BTNSave.TabIndex = 313
        Me.BTNSave.Text = "F5-บันทึก"
        '
        'PNReceiveMoney
        '
        Me.PNReceiveMoney.BackColor = System.Drawing.Color.SkyBlue
        Me.PNReceiveMoney.Controls.Add(Me.ListViewPayDetails)
        Me.PNReceiveMoney.Controls.Add(Me.BTNPayCancel)
        Me.PNReceiveMoney.Controls.Add(Me.BTNPayOK)
        Me.PNReceiveMoney.Controls.Add(Me.BTNCreditDelete)
        Me.PNReceiveMoney.Controls.Add(Me.BTNCreditUpdate)
        Me.PNReceiveMoney.Controls.Add(Me.TCPayMoney)
        Me.PNReceiveMoney.Location = New System.Drawing.Point(27, 8)
        Me.PNReceiveMoney.Name = "PNReceiveMoney"
        Me.PNReceiveMoney.Size = New System.Drawing.Size(294, 287)
        Me.PNReceiveMoney.Visible = False
        '
        'ListViewPayDetails
        '
        Me.ListViewPayDetails.BackColor = System.Drawing.Color.Orange
        Me.ListViewPayDetails.Columns.Add(Me.ColumnHeader12)
        Me.ListViewPayDetails.Columns.Add(Me.ColumnHeader13)
        Me.ListViewPayDetails.Columns.Add(Me.ColumnHeader15)
        Me.ListViewPayDetails.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.ListViewPayDetails.FullRowSelect = True
        Me.ListViewPayDetails.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.ListViewPayDetails.Location = New System.Drawing.Point(128, 0)
        Me.ListViewPayDetails.Name = "ListViewPayDetails"
        Me.ListViewPayDetails.Size = New System.Drawing.Size(183, 68)
        Me.ListViewPayDetails.TabIndex = 280
        Me.ListViewPayDetails.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "No1"
        Me.ColumnHeader12.Width = 100
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = "No2"
        Me.ColumnHeader13.Width = 100
        '
        'ColumnHeader15
        '
        Me.ColumnHeader15.Text = "No3"
        Me.ColumnHeader15.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader15.Width = 100
        '
        'BTNPayCancel
        '
        Me.BTNPayCancel.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNPayCancel.Location = New System.Drawing.Point(241, 258)
        Me.BTNPayCancel.Name = "BTNPayCancel"
        Me.BTNPayCancel.Size = New System.Drawing.Size(70, 24)
        Me.BTNPayCancel.TabIndex = 93
        Me.BTNPayCancel.Text = "ออก"
        '
        'BTNPayOK
        '
        Me.BTNPayOK.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNPayOK.Location = New System.Drawing.Point(162, 258)
        Me.BTNPayOK.Name = "BTNPayOK"
        Me.BTNPayOK.Size = New System.Drawing.Size(70, 24)
        Me.BTNPayOK.TabIndex = 92
        Me.BTNPayOK.Text = "ตกลง"
        '
        'BTNCreditDelete
        '
        Me.BTNCreditDelete.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNCreditDelete.Location = New System.Drawing.Point(83, 258)
        Me.BTNCreditDelete.Name = "BTNCreditDelete"
        Me.BTNCreditDelete.Size = New System.Drawing.Size(70, 24)
        Me.BTNCreditDelete.TabIndex = 91
        Me.BTNCreditDelete.Text = "ลบ"
        Me.BTNCreditDelete.Visible = False
        '
        'BTNCreditUpdate
        '
        Me.BTNCreditUpdate.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.BTNCreditUpdate.Location = New System.Drawing.Point(4, 258)
        Me.BTNCreditUpdate.Name = "BTNCreditUpdate"
        Me.BTNCreditUpdate.Size = New System.Drawing.Size(70, 24)
        Me.BTNCreditUpdate.TabIndex = 90
        Me.BTNCreditUpdate.Text = "เพิ่ม"
        Me.BTNCreditUpdate.Visible = False
        '
        'TCPayMoney
        '
        Me.TCPayMoney.Controls.Add(Me.TabPage1)
        Me.TCPayMoney.Controls.Add(Me.TabPage2)
        Me.TCPayMoney.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TCPayMoney.Location = New System.Drawing.Point(4, 69)
        Me.TCPayMoney.Name = "TCPayMoney"
        Me.TCPayMoney.SelectedIndex = 0
        Me.TCPayMoney.Size = New System.Drawing.Size(307, 185)
        Me.TCPayMoney.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.LightGreen
        Me.TabPage1.Controls.Add(Me.Label21)
        Me.TabPage1.Controls.Add(Me.TBOtherDebt)
        Me.TabPage1.Controls.Add(Me.Label18)
        Me.TabPage1.Controls.Add(Me.TBOtherExpense)
        Me.TabPage1.Controls.Add(Me.Label17)
        Me.TabPage1.Controls.Add(Me.TBOverMoney)
        Me.TabPage1.Controls.Add(Me.TBOverMoneyInv)
        Me.TabPage1.Controls.Add(Me.Label12)
        Me.TabPage1.Controls.Add(Me.TBCashAmount)
        Me.TabPage1.Controls.Add(Me.Label16)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(299, 165)
        Me.TabPage1.Text = "เงินสด"
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label21.Location = New System.Drawing.Point(12, 103)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(50, 20)
        Me.Label21.Text = "รายได้ :"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBOtherDebt
        '
        Me.TBOtherDebt.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBOtherDebt.Location = New System.Drawing.Point(64, 103)
        Me.TBOtherDebt.Name = "TBOtherDebt"
        Me.TBOtherDebt.Size = New System.Drawing.Size(88, 19)
        Me.TBOtherDebt.TabIndex = 282
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label18.Location = New System.Drawing.Point(0, 78)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(62, 20)
        Me.Label18.Text = "ค่าใช้จ่าย :"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBOtherExpense
        '
        Me.TBOtherExpense.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBOtherExpense.Location = New System.Drawing.Point(64, 78)
        Me.TBOtherExpense.Name = "TBOtherExpense"
        Me.TBOtherExpense.Size = New System.Drawing.Size(88, 19)
        Me.TBOtherExpense.TabIndex = 279
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label17.Location = New System.Drawing.Point(158, 51)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(45, 20)
        Me.Label17.Text = "เงินหัก :"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBOverMoney
        '
        Me.TBOverMoney.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBOverMoney.Location = New System.Drawing.Point(206, 51)
        Me.TBOverMoney.Name = "TBOverMoney"
        Me.TBOverMoney.Size = New System.Drawing.Size(88, 19)
        Me.TBOverMoney.TabIndex = 276
        '
        'TBOverMoneyInv
        '
        Me.TBOverMoneyInv.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBOverMoneyInv.Location = New System.Drawing.Point(206, 26)
        Me.TBOverMoneyInv.Name = "TBOverMoneyInv"
        Me.TBOverMoneyInv.Size = New System.Drawing.Size(88, 19)
        Me.TBOverMoneyInv.TabIndex = 273
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label12.Location = New System.Drawing.Point(16, 26)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(46, 20)
        Me.Label12.Text = "เงินสด :"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBCashAmount
        '
        Me.TBCashAmount.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBCashAmount.Location = New System.Drawing.Point(64, 26)
        Me.TBCashAmount.Name = "TBCashAmount"
        Me.TBCashAmount.Size = New System.Drawing.Size(88, 19)
        Me.TBCashAmount.TabIndex = 270
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label16.Location = New System.Drawing.Point(153, 26)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(50, 20)
        Me.Label16.Text = "เงินเกิน :"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.LightPink
        Me.TabPage2.Controls.Add(Me.TBCreditTotalAmount)
        Me.TabPage2.Controls.Add(Me.Label31)
        Me.TabPage2.Controls.Add(Me.TBChargeAmount)
        Me.TabPage2.Controls.Add(Me.Label29)
        Me.TabPage2.Controls.Add(Me.Label27)
        Me.TabPage2.Controls.Add(Me.Label26)
        Me.TabPage2.Controls.Add(Me.Label25)
        Me.TabPage2.Controls.Add(Me.CMBBranch)
        Me.TabPage2.Controls.Add(Me.CMBCreditType)
        Me.TabPage2.Controls.Add(Me.CMBBank)
        Me.TabPage2.Controls.Add(Me.ListViewCreditCard)
        Me.TabPage2.Controls.Add(Me.TBCreditAmount)
        Me.TabPage2.Controls.Add(Me.Label23)
        Me.TabPage2.Controls.Add(Me.TBConfirmNo)
        Me.TabPage2.Controls.Add(Me.Label22)
        Me.TabPage2.Controls.Add(Me.TBCreditCard)
        Me.TabPage2.Controls.Add(Me.Label24)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(299, 159)
        Me.TabPage2.Text = "บัตรเครดิต"
        '
        'TBCreditTotalAmount
        '
        Me.TBCreditTotalAmount.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBCreditTotalAmount.Location = New System.Drawing.Point(217, 135)
        Me.TBCreditTotalAmount.Name = "TBCreditTotalAmount"
        Me.TBCreditTotalAmount.Size = New System.Drawing.Size(79, 19)
        Me.TBCreditTotalAmount.TabIndex = 299
        '
        'Label31
        '
        Me.Label31.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label31.Location = New System.Drawing.Point(156, 136)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(60, 20)
        Me.Label31.Text = "รวม :"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBChargeAmount
        '
        Me.TBChargeAmount.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBChargeAmount.Location = New System.Drawing.Point(217, 112)
        Me.TBChargeAmount.Name = "TBChargeAmount"
        Me.TBChargeAmount.Size = New System.Drawing.Size(79, 19)
        Me.TBChargeAmount.TabIndex = 294
        '
        'Label29
        '
        Me.Label29.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label29.Location = New System.Drawing.Point(156, 112)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(60, 20)
        Me.Label29.Text = "ยอดชาร์ต :"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label27
        '
        Me.Label27.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label27.Location = New System.Drawing.Point(156, 88)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(60, 20)
        Me.Label27.Text = "เลขอนุมัติ :"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label26
        '
        Me.Label26.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label26.Location = New System.Drawing.Point(4, 135)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(60, 20)
        Me.Label26.Text = "สาขา :"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label25
        '
        Me.Label25.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label25.Location = New System.Drawing.Point(4, 112)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(60, 20)
        Me.Label25.Text = "ธนาคาร :"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'CMBBranch
        '
        Me.CMBBranch.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.CMBBranch.Location = New System.Drawing.Point(65, 135)
        Me.CMBBranch.Name = "CMBBranch"
        Me.CMBBranch.Size = New System.Drawing.Size(88, 19)
        Me.CMBBranch.TabIndex = 282
        '
        'CMBCreditType
        '
        Me.CMBCreditType.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.CMBCreditType.Location = New System.Drawing.Point(217, 88)
        Me.CMBCreditType.Name = "CMBCreditType"
        Me.CMBCreditType.Size = New System.Drawing.Size(79, 19)
        Me.CMBCreditType.TabIndex = 281
        '
        'CMBBank
        '
        Me.CMBBank.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.CMBBank.Location = New System.Drawing.Point(65, 112)
        Me.CMBBank.Name = "CMBBank"
        Me.CMBBank.Size = New System.Drawing.Size(88, 19)
        Me.CMBBank.TabIndex = 280
        '
        'ListViewCreditCard
        '
        Me.ListViewCreditCard.BackColor = System.Drawing.Color.Orange
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader16)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader17)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader18)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader19)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader20)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader21)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader22)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader23)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader28)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader29)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader30)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader31)
        Me.ListViewCreditCard.Columns.Add(Me.ColumnHeader33)
        Me.ListViewCreditCard.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.ListViewCreditCard.FullRowSelect = True
        Me.ListViewCreditCard.Location = New System.Drawing.Point(2, 3)
        Me.ListViewCreditCard.Name = "ListViewCreditCard"
        Me.ListViewCreditCard.Size = New System.Drawing.Size(294, 59)
        Me.ListViewCreditCard.TabIndex = 279
        Me.ListViewCreditCard.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader16
        '
        Me.ColumnHeader16.Text = "No"
        Me.ColumnHeader16.Width = 30
        '
        'ColumnHeader17
        '
        Me.ColumnHeader17.Text = "ชื่อสินค้า"
        Me.ColumnHeader17.Width = 120
        '
        'ColumnHeader18
        '
        Me.ColumnHeader18.Text = "จ่าย"
        Me.ColumnHeader18.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader18.Width = 60
        '
        'ColumnHeader19
        '
        Me.ColumnHeader19.Text = "หน่วย"
        Me.ColumnHeader19.Width = 70
        '
        'ColumnHeader20
        '
        Me.ColumnHeader20.Text = "คิวที่"
        Me.ColumnHeader20.Width = 40
        '
        'ColumnHeader21
        '
        Me.ColumnHeader21.Text = "โซน"
        Me.ColumnHeader21.Width = 40
        '
        'ColumnHeader22
        '
        Me.ColumnHeader22.Text = "เลขที่เอกสาร"
        Me.ColumnHeader22.Width = 100
        '
        'ColumnHeader23
        '
        Me.ColumnHeader23.Text = "รหัส"
        Me.ColumnHeader23.Width = 110
        '
        'ColumnHeader28
        '
        Me.ColumnHeader28.Text = "คลัง"
        Me.ColumnHeader28.Width = 60
        '
        'ColumnHeader29
        '
        Me.ColumnHeader29.Text = "ชั้นเก็บ"
        Me.ColumnHeader29.Width = 60
        '
        'ColumnHeader30
        '
        Me.ColumnHeader30.Text = "บาร์โค้ด"
        Me.ColumnHeader30.Width = 110
        '
        'ColumnHeader31
        '
        Me.ColumnHeader31.Text = "PickZone"
        Me.ColumnHeader31.Width = 60
        '
        'ColumnHeader33
        '
        Me.ColumnHeader33.Text = "ที่เก็บ"
        Me.ColumnHeader33.Width = 60
        '
        'TBCreditAmount
        '
        Me.TBCreditAmount.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBCreditAmount.Location = New System.Drawing.Point(217, 65)
        Me.TBCreditAmount.Name = "TBCreditAmount"
        Me.TBCreditAmount.Size = New System.Drawing.Size(79, 19)
        Me.TBCreditAmount.TabIndex = 277
        '
        'Label23
        '
        Me.Label23.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label23.Location = New System.Drawing.Point(4, 88)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(60, 20)
        Me.Label23.Text = "เลขอนุมัติ :"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBConfirmNo
        '
        Me.TBConfirmNo.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBConfirmNo.Location = New System.Drawing.Point(65, 88)
        Me.TBConfirmNo.Name = "TBConfirmNo"
        Me.TBConfirmNo.Size = New System.Drawing.Size(88, 19)
        Me.TBConfirmNo.TabIndex = 274
        '
        'Label22
        '
        Me.Label22.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label22.Location = New System.Drawing.Point(14, 66)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(50, 20)
        Me.Label22.Text = "เลขที่ :"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBCreditCard
        '
        Me.TBCreditCard.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBCreditCard.Location = New System.Drawing.Point(65, 65)
        Me.TBCreditCard.Name = "TBCreditCard"
        Me.TBCreditCard.Size = New System.Drawing.Size(88, 19)
        Me.TBCreditCard.TabIndex = 272
        '
        'Label24
        '
        Me.Label24.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label24.Location = New System.Drawing.Point(156, 66)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(60, 20)
        Me.Label24.Text = "ยอดบัตร :"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'PNPOSConfig
        '
        Me.PNPOSConfig.BackColor = System.Drawing.Color.Orange
        Me.PNPOSConfig.Controls.Add(Me.TBCharge)
        Me.PNPOSConfig.Controls.Add(Me.Label28)
        Me.PNPOSConfig.Controls.Add(Me.CMBMemCalcStock)
        Me.PNPOSConfig.Controls.Add(Me.TBMemTaxRate)
        Me.PNPOSConfig.Controls.Add(Me.CMBMemDepartment)
        Me.PNPOSConfig.Controls.Add(Me.CMBMemPosID)
        Me.PNPOSConfig.Controls.Add(Me.CMBMemShelfCode)
        Me.PNPOSConfig.Controls.Add(Me.TBMemArName)
        Me.PNPOSConfig.Controls.Add(Me.Label11)
        Me.PNPOSConfig.Controls.Add(Me.Label10)
        Me.PNPOSConfig.Controls.Add(Me.Label9)
        Me.PNPOSConfig.Controls.Add(Me.Label7)
        Me.PNPOSConfig.Controls.Add(Me.Label6)
        Me.PNPOSConfig.Controls.Add(Me.Label5)
        Me.PNPOSConfig.Controls.Add(Me.CMBMemWHCode)
        Me.PNPOSConfig.Controls.Add(Me.TBMemArCode)
        Me.PNPOSConfig.Controls.Add(Me.Label2)
        Me.PNPOSConfig.Location = New System.Drawing.Point(308, 8)
        Me.PNPOSConfig.Name = "PNPOSConfig"
        Me.PNPOSConfig.Size = New System.Drawing.Size(12, 255)
        Me.PNPOSConfig.Visible = False
        '
        'TBCharge
        '
        Me.TBCharge.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBCharge.Location = New System.Drawing.Point(72, 163)
        Me.TBCharge.Name = "TBCharge"
        Me.TBCharge.Size = New System.Drawing.Size(79, 19)
        Me.TBCharge.TabIndex = 299
        '
        'Label28
        '
        Me.Label28.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label28.Location = New System.Drawing.Point(10, 164)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(60, 20)
        Me.Label28.Text = "ชาร์ต % :"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'CMBMemCalcStock
        '
        Me.CMBMemCalcStock.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.CMBMemCalcStock.Location = New System.Drawing.Point(221, 132)
        Me.CMBMemCalcStock.Name = "CMBMemCalcStock"
        Me.CMBMemCalcStock.Size = New System.Drawing.Size(69, 19)
        Me.CMBMemCalcStock.TabIndex = 290
        '
        'TBMemTaxRate
        '
        Me.TBMemTaxRate.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBMemTaxRate.Location = New System.Drawing.Point(72, 132)
        Me.TBMemTaxRate.Name = "TBMemTaxRate"
        Me.TBMemTaxRate.Size = New System.Drawing.Size(69, 19)
        Me.TBMemTaxRate.TabIndex = 289
        '
        'CMBMemDepartment
        '
        Me.CMBMemDepartment.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.CMBMemDepartment.Location = New System.Drawing.Point(72, 106)
        Me.CMBMemDepartment.Name = "CMBMemDepartment"
        Me.CMBMemDepartment.Size = New System.Drawing.Size(240, 19)
        Me.CMBMemDepartment.TabIndex = 288
        '
        'CMBMemPosID
        '
        Me.CMBMemPosID.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.CMBMemPosID.Location = New System.Drawing.Point(72, 31)
        Me.CMBMemPosID.Name = "CMBMemPosID"
        Me.CMBMemPosID.Size = New System.Drawing.Size(69, 19)
        Me.CMBMemPosID.TabIndex = 287
        '
        'CMBMemShelfCode
        '
        Me.CMBMemShelfCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.CMBMemShelfCode.Location = New System.Drawing.Point(221, 81)
        Me.CMBMemShelfCode.Name = "CMBMemShelfCode"
        Me.CMBMemShelfCode.Size = New System.Drawing.Size(69, 19)
        Me.CMBMemShelfCode.TabIndex = 286
        '
        'TBMemArName
        '
        Me.TBMemArName.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBMemArName.Location = New System.Drawing.Point(147, 56)
        Me.TBMemArName.Name = "TBMemArName"
        Me.TBMemArName.Size = New System.Drawing.Size(165, 19)
        Me.TBMemArName.TabIndex = 285
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label11.Location = New System.Drawing.Point(6, 107)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(64, 20)
        Me.Label11.Text = "แผนก :"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label10.Location = New System.Drawing.Point(1, 31)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(69, 20)
        Me.Label10.Text = "เครื่อง POS :"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label9.Location = New System.Drawing.Point(133, 132)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(86, 20)
        Me.Label9.Text = "สถานะติดลบ :"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label7.Location = New System.Drawing.Point(3, 132)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 20)
        Me.Label7.Text = "อัตราภาษี :"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label6.Location = New System.Drawing.Point(3, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(67, 20)
        Me.Label6.Text = "ลูกค้า :"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(152, 81)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 20)
        Me.Label5.Text = "ชั้นเก็บขาย :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'CMBMemWHCode
        '
        Me.CMBMemWHCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.CMBMemWHCode.Location = New System.Drawing.Point(72, 81)
        Me.CMBMemWHCode.Name = "CMBMemWHCode"
        Me.CMBMemWHCode.Size = New System.Drawing.Size(69, 19)
        Me.CMBMemWHCode.TabIndex = 272
        '
        'TBMemArCode
        '
        Me.TBMemArCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBMemArCode.Location = New System.Drawing.Point(72, 56)
        Me.TBMemArCode.Name = "TBMemArCode"
        Me.TBMemArCode.Size = New System.Drawing.Size(69, 19)
        Me.TBMemArCode.TabIndex = 270
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(10, 81)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 20)
        Me.Label2.Text = "คลังขาย :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'TBWHCode
        '
        Me.TBWHCode.BackColor = System.Drawing.Color.Gold
        Me.TBWHCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBWHCode.Location = New System.Drawing.Point(169, 71)
        Me.TBWHCode.Name = "TBWHCode"
        Me.TBWHCode.ReadOnly = True
        Me.TBWHCode.Size = New System.Drawing.Size(12, 19)
        Me.TBWHCode.TabIndex = 329
        Me.TBWHCode.Visible = False
        '
        'TBShelfCode
        '
        Me.TBShelfCode.BackColor = System.Drawing.Color.Gold
        Me.TBShelfCode.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBShelfCode.Location = New System.Drawing.Point(163, 92)
        Me.TBShelfCode.Name = "TBShelfCode"
        Me.TBShelfCode.ReadOnly = True
        Me.TBShelfCode.Size = New System.Drawing.Size(12, 19)
        Me.TBShelfCode.TabIndex = 330
        Me.TBShelfCode.Visible = False
        '
        'TBItemAmount
        '
        Me.TBItemAmount.BackColor = System.Drawing.Color.Gold
        Me.TBItemAmount.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBItemAmount.Location = New System.Drawing.Point(6, 244)
        Me.TBItemAmount.Name = "TBItemAmount"
        Me.TBItemAmount.ReadOnly = True
        Me.TBItemAmount.Size = New System.Drawing.Size(122, 19)
        Me.TBItemAmount.TabIndex = 343
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Blue
        Me.Panel1.Location = New System.Drawing.Point(6, 266)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(314, 3)
        '
        'TBStkUnit
        '
        Me.TBStkUnit.BackColor = System.Drawing.Color.Gold
        Me.TBStkUnit.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBStkUnit.Location = New System.Drawing.Point(178, 71)
        Me.TBStkUnit.Name = "TBStkUnit"
        Me.TBStkUnit.ReadOnly = True
        Me.TBStkUnit.Size = New System.Drawing.Size(12, 19)
        Me.TBStkUnit.TabIndex = 346
        Me.TBStkUnit.Visible = False
        '
        'TBPayAmount
        '
        Me.TBPayAmount.BackColor = System.Drawing.Color.Gold
        Me.TBPayAmount.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBPayAmount.Location = New System.Drawing.Point(170, 244)
        Me.TBPayAmount.Name = "TBPayAmount"
        Me.TBPayAmount.ReadOnly = True
        Me.TBPayAmount.Size = New System.Drawing.Size(122, 19)
        Me.TBPayAmount.TabIndex = 360
        '
        'TBBalanceAmount
        '
        Me.TBBalanceAmount.BackColor = System.Drawing.Color.Gold
        Me.TBBalanceAmount.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBBalanceAmount.Location = New System.Drawing.Point(53, 198)
        Me.TBBalanceAmount.Name = "TBBalanceAmount"
        Me.TBBalanceAmount.ReadOnly = True
        Me.TBBalanceAmount.Size = New System.Drawing.Size(122, 19)
        Me.TBBalanceAmount.TabIndex = 361
        '
        'TBBillBalance
        '
        Me.TBBillBalance.BackColor = System.Drawing.Color.Gold
        Me.TBBillBalance.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBBillBalance.Location = New System.Drawing.Point(6, 220)
        Me.TBBillBalance.Name = "TBBillBalance"
        Me.TBBillBalance.ReadOnly = True
        Me.TBBillBalance.Size = New System.Drawing.Size(122, 19)
        Me.TBBillBalance.TabIndex = 362
        '
        'TBBillAmount
        '
        Me.TBBillAmount.BackColor = System.Drawing.Color.Gold
        Me.TBBillAmount.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.TBBillAmount.Location = New System.Drawing.Point(170, 220)
        Me.TBBillAmount.Name = "TBBillAmount"
        Me.TBBillAmount.ReadOnly = True
        Me.TBBillAmount.Size = New System.Drawing.Size(122, 19)
        Me.TBBillAmount.TabIndex = 377
        '
        'FormPOS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.Khaki
        Me.ClientSize = New System.Drawing.Size(325, 300)
        Me.ControlBox = False
        Me.Controls.Add(Me.PNReceiveMoney)
        Me.Controls.Add(Me.PNPOSConfig)
        Me.Controls.Add(Me.TBBillBalance)
        Me.Controls.Add(Me.TBBalanceAmount)
        Me.Controls.Add(Me.TBStkUnit)
        Me.Controls.Add(Me.TBItemAmount)
        Me.Controls.Add(Me.TBShelfCode)
        Me.Controls.Add(Me.TBWHCode)
        Me.Controls.Add(Me.BTNSearch)
        Me.Controls.Add(Me.BTNCancel)
        Me.Controls.Add(Me.BTNNew)
        Me.Controls.Add(Me.BTNExit)
        Me.Controls.Add(Me.BTNSave)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.TBDocDate)
        Me.Controls.Add(Me.TBDocNo)
        Me.Controls.Add(Me.TBPrice)
        Me.Controls.Add(Me.VScrollBar1)
        Me.Controls.Add(Me.TBBarCode)
        Me.Controls.Add(Me.ListViewItem)
        Me.Controls.Add(Me.TBUnit)
        Me.Controls.Add(Me.TBQty)
        Me.Controls.Add(Me.TBItemName)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.TBItemCode)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.TBRemainQty)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.TBRate1)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TBPayAmount)
        Me.Controls.Add(Me.TBBillAmount)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FormPOS"
        Me.Text = "FormPOS"
        Me.PNReceiveMoney.ResumeLayout(False)
        Me.TCPayMoney.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.PNPOSConfig.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TBDocDate As System.Windows.Forms.TextBox
    Friend WithEvents TBDocNo As System.Windows.Forms.TextBox
    Friend WithEvents TBPrice As System.Windows.Forms.TextBox
    Friend WithEvents VScrollBar1 As System.Windows.Forms.VScrollBar
    Friend WithEvents TBBarCode As System.Windows.Forms.TextBox
    Friend WithEvents ListViewItem As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader14 As System.Windows.Forms.ColumnHeader
    Friend WithEvents TBUnit As System.Windows.Forms.TextBox
    Friend WithEvents TBQty As System.Windows.Forms.TextBox
    Friend WithEvents TBItemName As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TBItemCode As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents TBRate1 As System.Windows.Forms.TextBox
    Friend WithEvents TBRemainQty As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents BTNSearch As System.Windows.Forms.Button
    Friend WithEvents BTNCancel As System.Windows.Forms.Button
    Friend WithEvents BTNNew As System.Windows.Forms.Button
    Friend WithEvents BTNExit As System.Windows.Forms.Button
    Friend WithEvents BTNSave As System.Windows.Forms.Button
    Friend WithEvents PNReceiveMoney As System.Windows.Forms.Panel
    Friend WithEvents TCPayMoney As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents PNPOSConfig As System.Windows.Forms.Panel
    Friend WithEvents TBWHCode As System.Windows.Forms.TextBox
    Friend WithEvents TBShelfCode As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CMBMemWHCode As System.Windows.Forms.ComboBox
    Friend WithEvents TBMemArCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents CMBMemCalcStock As System.Windows.Forms.ComboBox
    Friend WithEvents TBMemTaxRate As System.Windows.Forms.TextBox
    Friend WithEvents CMBMemDepartment As System.Windows.Forms.ComboBox
    Friend WithEvents CMBMemPosID As System.Windows.Forms.ComboBox
    Friend WithEvents CMBMemShelfCode As System.Windows.Forms.ComboBox
    Friend WithEvents TBMemArName As System.Windows.Forms.TextBox
    Friend WithEvents TBItemAmount As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TBStkUnit As System.Windows.Forms.TextBox
    Friend WithEvents TBCashAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TBOverMoneyInv As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TBOtherDebt As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TBOtherExpense As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TBOverMoney As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents TBCreditAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents TBConfirmNo As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TBCreditCard As System.Windows.Forms.TextBox
    Friend WithEvents ListViewCreditCard As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader16 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader17 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader18 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader19 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader20 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader21 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader22 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader23 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader28 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader29 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader30 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader31 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader33 As System.Windows.Forms.ColumnHeader
    Friend WithEvents CMBBank As System.Windows.Forms.ComboBox
    Friend WithEvents CMBBranch As System.Windows.Forms.ComboBox
    Friend WithEvents CMBCreditType As System.Windows.Forms.ComboBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents TBChargeAmount As System.Windows.Forms.TextBox
    Friend WithEvents TBCreditTotalAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents BTNPayCancel As System.Windows.Forms.Button
    Friend WithEvents BTNPayOK As System.Windows.Forms.Button
    Friend WithEvents BTNCreditDelete As System.Windows.Forms.Button
    Friend WithEvents BTNCreditUpdate As System.Windows.Forms.Button
    Friend WithEvents TBPayAmount As System.Windows.Forms.TextBox
    Friend WithEvents ListViewPayDetails As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader13 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader15 As System.Windows.Forms.ColumnHeader
    Friend WithEvents TBBalanceAmount As System.Windows.Forms.TextBox
    Friend WithEvents TBBillBalance As System.Windows.Forms.TextBox
    Friend WithEvents TBBillAmount As System.Windows.Forms.TextBox
    Friend WithEvents TBCharge As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
End Class
