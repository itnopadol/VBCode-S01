VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormCheckOutHoldBill 
   Caption         =   "ตรวจสอบการจ่ายสินค้าและทำเอกสารพักบิล"
   ClientHeight    =   10590
   ClientLeft      =   1860
   ClientTop       =   825
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormCheckOutHoldBill.frx":0000
   ScaleHeight     =   14156.96
   ScaleMode       =   0  'User
   ScaleWidth      =   18637.58
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICSearchHoldBill 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9600
      Left            =   0
      Picture         =   "FormCheckOutHoldBill.frx":9D15
      ScaleHeight     =   9570
      ScaleWidth      =   15195
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   15225
      Begin VB.PictureBox Picture6 
         BackColor       =   &H000080FF&
         Height          =   330
         Left            =   -45
         ScaleHeight     =   270
         ScaleWidth      =   14895
         TabIndex        =   116
         Top             =   1530
         Width           =   14955
      End
      Begin VB.CommandButton CMDSearchHoldBill 
         Height          =   330
         Left            =   7830
         Picture         =   "FormCheckOutHoldBill.frx":13A2A
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   675
         Width           =   330
      End
      Begin VB.CommandButton CMDSearchHoldBillExit 
         BackColor       =   &H00808080&
         Caption         =   "ปิด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12915
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   8100
         Width           =   1815
      End
      Begin VB.CommandButton CMDSearchHoldBillOK 
         BackColor       =   &H00808080&
         Caption         =   "เลือก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10845
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   8100
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListViewHoldBill 
         Height          =   5730
         Left            =   135
         TabIndex        =   62
         Top             =   2250
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   10107
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "เลขที่เอกสาร"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "วันที่เอกสาร"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "พนักงานขาย"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "มูลค่าบิล"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "แคชเชียร์"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "อ้างอิง"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.TextBox TBSearchHoldBill 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3465
         TabIndex        =   60
         Top             =   675
         Width           =   4335
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00C00000&
         Height          =   330
         Left            =   -90
         ScaleHeight     =   270
         ScaleWidth      =   15075
         TabIndex        =   115
         Top             =   9225
         Width           =   15135
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการเอกสารที่ค้นหา "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   109
         Top             =   1890
         Width           =   2805
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "คำที่ค้นหา :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2070
         TabIndex        =   59
         Top             =   675
         Width           =   1320
      End
   End
   Begin VB.PictureBox PICHoldBill 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9465
      Left            =   0
      Picture         =   "FormCheckOutHoldBill.frx":13E7D
      ScaleHeight     =   9435
      ScaleWidth      =   14970
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   15000
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00C00000&
         Height          =   375
         Left            =   -2610
         ScaleHeight     =   315
         ScaleWidth      =   19215
         TabIndex        =   118
         Top             =   9135
         Width           =   19275
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H000080FF&
         Height          =   330
         Left            =   -90
         ScaleHeight     =   270
         ScaleWidth      =   15300
         TabIndex        =   117
         Top             =   1530
         Width           =   15360
      End
      Begin VB.CommandButton CMDPrintHoldBill 
         BackColor       =   &H00808080&
         Caption         =   "พิมพ์เอกสารพักบิล"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   6795
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   7470
         Width           =   2175
      End
      Begin VB.CommandButton CMDDeleteHoldBill 
         BackColor       =   &H00808080&
         Caption         =   "ลบเอกสารพักบิล"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   7470
         Width           =   2130
      End
      Begin VB.CommandButton CMDHoldExit 
         BackColor       =   &H00808080&
         Caption         =   "ออก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   2385
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   7470
         Width           =   2130
      End
      Begin VB.CommandButton CMDHoldSave 
         BackColor       =   &H00808080&
         Caption         =   "บันทึก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   7470
         Width           =   2130
      End
      Begin VB.OptionButton OPCash3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "แคชเชียร์จุดที่ 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   11430
         TabIndex        =   44
         Top             =   1035
         Width           =   1905
      End
      Begin VB.OptionButton OPCash2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "แคชเชียร์จุดที่ 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11430
         TabIndex        =   43
         Top             =   585
         Width           =   1905
      End
      Begin VB.OptionButton OPCash1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "แคชเชียร์จุดที่ 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11430
         TabIndex        =   42
         Top             =   135
         Value           =   -1  'True
         Width           =   1905
      End
      Begin MSComctlLib.ListView ListViewItemHoldBill 
         Height          =   5235
         Left            =   180
         TabIndex        =   41
         Top             =   1980
         Width           =   14550
         _ExtentX        =   25665
         _ExtentY        =   9234
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "จำนวน"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วย"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "ราคา"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "ส่วนลด"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "มูลค่า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "คลัง"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "ชั้นเก็บ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "บาร์โค้ด"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "อ้างอิง1"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "อ้างอิง2"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "Rate1"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "Rate2"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ทะเบียนรถ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5490
         TabIndex        =   122
         Top             =   135
         Width           =   1500
      End
      Begin VB.Label LBLHoldCarLicense 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7065
         TabIndex        =   119
         Top             =   135
         Width           =   2085
      End
      Begin VB.Label LBLHoldBillType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   9360
         TabIndex        =   68
         Top             =   1035
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label LBLHoldBillNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6120
         TabIndex        =   66
         Top             =   1035
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.Label LBLHoldSaleCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   55
         Top             =   1035
         Width           =   3120
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   990
         TabIndex        =   54
         Top             =   1035
         Width           =   1815
      End
      Begin VB.Label LBLHoldNetAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   12645
         TabIndex        =   53
         Top             =   8505
         Width           =   2085
      End
      Begin VB.Label LBLHoldTaxAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   12645
         TabIndex        =   52
         Top             =   7965
         Width           =   2085
      End
      Begin VB.Label LBLHoldItemAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   12645
         TabIndex        =   51
         Top             =   7425
         Width           =   2085
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "มูลค่าภาษี :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10665
         TabIndex        =   50
         Top             =   7965
         Width           =   1860
      End
      Begin VB.Label LBLHoldArName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   49
         Top             =   585
         Width           =   6270
      End
      Begin VB.Label LBLHoldArCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2880
         TabIndex        =   48
         Top             =   135
         Width           =   1860
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   900
         TabIndex        =   47
         Top             =   630
         Width           =   1905
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสลูกค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1125
         TabIndex        =   46
         Top             =   135
         Width           =   1680
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกจุดพักบิล :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   9135
         TabIndex        =   45
         Top             =   135
         Width           =   2130
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "มูลค่าบิล :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11205
         TabIndex        =   40
         Top             =   8505
         Width           =   1320
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "มูลค่าสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11205
         TabIndex        =   39
         Top             =   7425
         Width           =   1320
      End
   End
   Begin VB.PictureBox PICSelectQue 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   9420
      Left            =   0
      ScaleHeight     =   9390
      ScaleWidth      =   15195
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   15225
      Begin VB.CommandButton CMDSelectItemClose 
         BackColor       =   &H00808080&
         Caption         =   "ปิด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12780
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7470
         Width           =   1860
      End
      Begin VB.CommandButton CMDSelectItem 
         BackColor       =   &H00808080&
         Caption         =   "เลือก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10710
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   7470
         Width           =   1860
      End
      Begin MSComctlLib.ListView ListViewSelectQue 
         Height          =   6225
         Left            =   180
         TabIndex        =   19
         Top             =   855
         Width           =   14505
         _ExtentX        =   25585
         _ExtentY        =   10980
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   19
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "อ้างอิง"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "เลขที่คิว"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "สถานะ"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "รหัสสินค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ชื่อสินค้า"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "พนักงานจัด"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "สถานะคิว"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "สั่ง"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "จัดได้"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "หน่วย"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "ราคา"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "ส่วนลด"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "มูลค่า"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "บาร์โค้ด"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Text            =   "คลัง"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   15
            Text            =   "ชั้นเก็บ"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "รหัสลูกค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "พนักงานขาย"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "ทะเบียนรถ"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการ สินค้าจากคิวการจัดสินค้า "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   225
         TabIndex        =   67
         Top             =   180
         Width           =   4290
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080FF&
      Height          =   1320
      Left            =   -180
      ScaleHeight     =   1260
      ScaleWidth      =   15345
      TabIndex        =   114
      Top             =   9405
      Width           =   15405
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000080FF&
      Height          =   285
      Left            =   -90
      ScaleHeight     =   225
      ScaleWidth      =   15030
      TabIndex        =   111
      Top             =   1575
      Width           =   15090
   End
   Begin VB.PictureBox PICKeyCheckOutQTY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5910
      Left            =   135
      ScaleHeight     =   5880
      ScaleWidth      =   14565
      TabIndex        =   70
      Top             =   1980
      Visible         =   0   'False
      Width           =   14595
      Begin VB.CommandButton CMDCheckOutClose 
         BackColor       =   &H00808080&
         Caption         =   "ปิด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   3645
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   3645
         Width           =   2535
      End
      Begin VB.TextBox TBCheckOutItemCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3645
         TabIndex        =   76
         Top             =   1080
         Width           =   3705
      End
      Begin VB.TextBox TBCheckOutItemQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3645
         TabIndex        =   74
         Top             =   2790
         Width           =   2535
      End
      Begin VB.Label LBLCheckOutItemCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   7560
         TabIndex        =   78
         Top             =   1080
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label LBLCheckOutItemName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   3645
         TabIndex        =   77
         Top             =   1935
         Width           =   9150
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "จำนวนที่นับได้ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   810
         TabIndex        =   73
         Top             =   2790
         Width           =   2670
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1575
         TabIndex        =   72
         Top             =   1980
         Width           =   1905
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         TabIndex        =   71
         Top             =   1080
         Width           =   1860
      End
   End
   Begin VB.PictureBox PICKeySearchData 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   13050
      ScaleHeight     =   1560
      ScaleWidth      =   4440
      TabIndex        =   105
      Top             =   0
      Visible         =   0   'False
      Width           =   4470
      Begin VB.CommandButton CMDSearchDataClose 
         Caption         =   "ปิด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6525
         TabIndex        =   108
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TBSearchData 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1485
         TabIndex        =   107
         Top             =   360
         Width           =   4965
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "คำที่ค้นหา :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   106
         Top             =   495
         Width           =   1140
      End
   End
   Begin VB.PictureBox PICCOAddItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5910
      Left            =   135
      ScaleHeight     =   5880
      ScaleWidth      =   14565
      TabIndex        =   81
      Top             =   1980
      Visible         =   0   'False
      Width           =   14595
      Begin VB.CommandButton CMDPICCOAddItemClose 
         BackColor       =   &H00808080&
         Caption         =   "ปิด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   5130
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   4725
         Width           =   1995
      End
      Begin VB.CommandButton CMDPICCOAddItemOK 
         BackColor       =   &H00808080&
         Caption         =   "ตกลง"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   4725
         Width           =   1995
      End
      Begin VB.TextBox TBCOKeyQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3015
         TabIndex        =   99
         Top             =   3645
         Width           =   1995
      End
      Begin VB.TextBox TBCOBarCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3015
         TabIndex        =   95
         Top             =   720
         Width           =   2985
      End
      Begin MSComctlLib.ListView ListViewCOStock 
         Height          =   1680
         Left            =   6120
         TabIndex        =   88
         Top             =   2430
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   2963
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "คลัง"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชั้นเก็บ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "ยอดคงเหลือ"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "หน่วยนับ"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label LBLCORate2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   9495
         TabIndex        =   104
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LBLCORate1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8955
         TabIndex        =   103
         Top             =   720
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label LBLCOItemCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8460
         TabIndex        =   102
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LBLCOPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3015
         TabIndex        =   98
         Top             =   2790
         Width           =   1995
      End
      Begin VB.Label LBLCOUnitCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3015
         TabIndex        =   97
         Top             =   2070
         Width           =   1995
      End
      Begin VB.Label LBLCOItemName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3015
         TabIndex        =   96
         Top             =   1395
         Width           =   9510
      End
      Begin VB.Label LBLCOItemNetAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8010
         TabIndex        =   94
         Top             =   720
         Width           =   330
      End
      Begin VB.Label LBLCOBarCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7650
         TabIndex        =   93
         Top             =   720
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label LBLCOZoneID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7290
         TabIndex        =   92
         Top             =   720
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label LBLCOShelfID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6885
         TabIndex        =   91
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LBLCOShelfCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6480
         TabIndex        =   90
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "ยอดคงเหลือตามคลัง"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6075
         TabIndex        =   89
         Top             =   2025
         Width           =   3210
      End
      Begin VB.Label LBLCOWHCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6075
         TabIndex        =   87
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ต้องการจำนวน :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   990
         TabIndex        =   86
         Top             =   3645
         Width           =   1905
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ราคาขาย :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   990
         TabIndex        =   85
         Top             =   2790
         Width           =   1905
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "หน่วยนับ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   990
         TabIndex        =   84
         Top             =   2070
         Width           =   1905
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         TabIndex        =   83
         Top             =   1395
         Width           =   1905
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         TabIndex        =   82
         Top             =   720
         Width           =   1905
      End
   End
   Begin VB.PictureBox PICKeyCheckOut 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   5910
      Left            =   135
      ScaleHeight     =   5880
      ScaleWidth      =   14565
      TabIndex        =   22
      Top             =   1980
      Visible         =   0   'False
      Width           =   14595
      Begin VB.CommandButton CMDCheckOutExit 
         BackColor       =   &H00808080&
         Caption         =   "ปิด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   4725
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox TBKeyQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   795
         Left            =   1530
         TabIndex        =   35
         Top             =   2340
         Width           =   2940
      End
      Begin VB.CommandButton CMDCheckOut 
         BackColor       =   &H00808080&
         Caption         =   "ตกลง"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   2655
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label LBLIndex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   4725
         TabIndex        =   34
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label LBLDisCount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   8190
         TabIndex        =   33
         Top             =   1620
         Width           =   1770
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ส่วนลด :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7200
         TabIndex        =   32
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "นับได้ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         TabIndex        =   31
         Top             =   2565
         Width           =   870
      End
      Begin VB.Label LBLPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   4725
         TabIndex        =   30
         Top             =   1620
         Width           =   2265
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ราคา :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3825
         TabIndex        =   29
         Top             =   1620
         Width           =   825
      End
      Begin VB.Label LBLUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1530
         TabIndex        =   28
         Top             =   1620
         Width           =   1725
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "หน่วย :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   720
         TabIndex        =   27
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label LBLItemName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1530
         TabIndex        =   26
         Top             =   945
         Width           =   11895
      End
      Begin VB.Label LBLItemCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1530
         TabIndex        =   25
         Top             =   270
         Width           =   2940
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   24
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   23
         Top             =   315
         Width           =   1185
      End
   End
   Begin MSComctlLib.ListView ListViewMerge 
      Height          =   5910
      Left            =   135
      TabIndex        =   2
      Top             =   1980
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   10425
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1620
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   4499
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   8997
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "จ่าย"
         Object.Width           =   2699
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หน่วย"
         Object.Width           =   2159
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "ราคา"
         Object.Width           =   2699
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "ส่วนลด"
         Object.Width           =   2699
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "มูลค่า"
         Object.Width           =   2699
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "คลัง"
         Object.Width           =   2159
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "ชั้นเก็บ"
         Object.Width           =   2159
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "บาร์โค้ด"
         Object.Width           =   4499
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   3599
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "คิวที่"
         Object.Width           =   2159
      EndProperty
   End
   Begin VB.CommandButton CMDHoldBill 
      BackColor       =   &H00808080&
      Caption         =   "สร้างเอกสารพักบิล"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8145
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8370
      Width           =   2175
   End
   Begin VB.CommandButton CMDMerge 
      BackColor       =   &H00808080&
      Caption         =   "รวมรายการ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8370
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListViewItem 
      Height          =   5910
      Left            =   180
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   10425
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1620
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "นับได้"
         Object.Width           =   2699
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   8997
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "รหัสสินค้า"
         Object.Width           =   4499
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หน่วยนับ"
         Object.Width           =   2339
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   3599
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "จ่าย"
         Object.Width           =   2699
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "ราคา"
         Object.Width           =   2591
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "ส่วนลด"
         Object.Width           =   2591
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "มูลค่า"
         Object.Width           =   2591
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "บาร์โค้ด"
         Object.Width           =   4499
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "คลัง"
         Object.Width           =   2159
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "ชั้นเก็บ"
         Object.Width           =   2159
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "อ้างอิง"
         Object.Width           =   4499
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Rate1"
         Object.Width           =   1799
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Rate2"
         Object.Width           =   1799
      EndProperty
   End
   Begin VB.CommandButton CMDSearch 
      BackColor       =   &H00808080&
      Caption         =   "ค้นหา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   10350
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8370
      Width           =   2175
   End
   Begin VB.CommandButton CMDExit 
      BackColor       =   &H00808080&
      Caption         =   "ออก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   12555
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8370
      Width           =   2175
   End
   Begin VB.CommandButton CMDClear 
      BackColor       =   &H00808080&
      Caption         =   "ล้างหน้าจอ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8370
      Width           =   2175
   End
   Begin VB.CommandButton CMDSelectItemQue 
      BackColor       =   &H00808080&
      Caption         =   "เลือกรายการ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3015
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8370
      Width           =   2175
   End
   Begin VB.TextBox TBSearchQueID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   3915
      TabIndex        =   0
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CommandButton CMDSearchQueID 
      Height          =   420
      Left            =   8280
      Picture         =   "FormCheckOutHoldBill.frx":1DB92
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   330
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C00000&
      Height          =   1185
      Left            =   -90
      ScaleHeight     =   1125
      ScaleWidth      =   15300
      TabIndex        =   112
      Top             =   8235
      Width           =   15360
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   -45
      ScaleHeight     =   315
      ScaleWidth      =   14940
      TabIndex        =   113
      Top             =   8010
      Width           =   15000
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ทะบียนรถ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8460
      TabIndex        =   121
      Top             =   630
      Width           =   1365
   End
   Begin VB.Label LBLCarLicense 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   9900
      TabIndex        =   120
      Top             =   585
      Width           =   1995
   End
   Begin VB.Label LBLSaleCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   3915
      TabIndex        =   79
      Top             =   585
      Width           =   4335
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสพนักงานขาย :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2115
      TabIndex        =   80
      Top             =   585
      Width           =   1725
   End
   Begin VB.Label TBArCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   3915
      TabIndex        =   69
      Top             =   90
      Width           =   2490
   End
   Begin VB.Label LBLArName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   6435
      TabIndex        =   9
      Top             =   90
      Width           =   8295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อรหัสลูกค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2565
      TabIndex        =   12
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ค้นหา :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2790
      TabIndex        =   11
      Top             =   1125
      Width           =   1050
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9900
      TabIndex        =   17
      Top             =   3780
      Width           =   780
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9810
      TabIndex        =   16
      Top             =   4410
      Width           =   690
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4455
      TabIndex        =   15
      Top             =   6210
      Width           =   2805
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10440
      TabIndex        =   14
      Top             =   3285
      Width           =   780
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10575
      TabIndex        =   13
      Top             =   5220
      Width           =   1005
   End
End
Attribute VB_Name = "FormCheckOutHoldBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDArCode_Click()
'Me.PICAR.Visible = True
'Me.TXTSearchAR.SetFocus
End Sub

Private Sub CMDCheckOut_Click()
Dim vIndex As Integer
Dim vKeyQTY As Double
Dim vMemQty As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vAmount As Double

Dim vAnswer As Integer


vIndex = Me.LBLIndex.Caption

vMemQty = Me.ListViewItem.ListItems(vIndex).SubItems(6)

If Me.TBKeyQty.Text <> "" Then
  vKeyQTY = Me.TBKeyQty.Text
End If

If vMemQty <> vKeyQTY Then
   vAnswer = MsgBox("จำนวนที่นับได้ ไม่เท่ากับจำนวนที่จ่าย จะทำการตรวจนับอีกรอบหรือไม่", vbYesNo, "Send Question Message")
   
   If vAnswer = 6 Then
      Me.TBKeyQty.SetFocus
      Exit Sub
   Else
      Me.ListViewItem.ListItems(vIndex).SubItems(1) = Format(vKeyQTY, "##,##0.00")
      
      If Me.ListViewItem.ListItems(vIndex).SubItems(7) <> "" Then
         vPrice = Me.ListViewItem.ListItems(vIndex).SubItems(7)
      End If
      
      If Me.ListViewItem.ListItems(vIndex).SubItems(8) <> "" Then
         vDiscountAmount = Me.ListViewItem.ListItems(vIndex).SubItems(8)
      End If
      
      vAmount = (vKeyQTY * (vPrice - vDiscountAmount))
      
      Me.ListViewItem.ListItems(vIndex).SubItems(9) = Format(vAmount, "##,##0.00")
            
      ListViewItem.ListItems(vIndex).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(1).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(2).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(3).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(4).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(5).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(6).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(7).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(8).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(9).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(10).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(11).ForeColor = "&H000000FF"
      ListViewItem.ListItems.Item(vIndex).ListSubItems(12).ForeColor = "&H000000FF"
      
   End If
Else
Me.ListViewItem.ListItems(vIndex).SubItems(1) = Format(vKeyQTY, "##,##0.00")
If Me.ListViewItem.ListItems(vIndex).SubItems(7) <> "" Then
   vPrice = Me.ListViewItem.ListItems(vIndex).SubItems(7)
End If

If Me.ListViewItem.ListItems(vIndex).SubItems(8) <> "" Then
   vDiscountAmount = Me.ListViewItem.ListItems(vIndex).SubItems(8)
End If

vAmount = (vKeyQTY * (vPrice - vDiscountAmount))

Me.ListViewItem.ListItems(vIndex).SubItems(9) = Format(vAmount, "##,##0.00")

ListViewItem.ListItems(vIndex).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(1).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(2).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(3).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(4).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(5).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(6).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(7).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(8).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(9).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(10).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(11).ForeColor = "&H00000000"
ListViewItem.ListItems.Item(vIndex).ListSubItems(12).ForeColor = "&H00000000"
End If

Me.PICKeyCheckOut.Visible = False
Me.CMDHoldBill.Enabled = True
Me.CMDSelectItemQue.Enabled = True
Me.ListViewItem.SetFocus

End Sub

Private Sub CMDCheckOutClose_Click()
Me.PICKeyCheckOutQTY.Visible = False
Me.ListViewItem.SetFocus
End Sub

Private Sub CMDCheckOutClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICKeyCheckOutQTY.Visible = False
Me.ListViewItem.SetFocus
End If
End Sub

Private Sub CMDCheckOutExit_Click()
Me.PICKeyCheckOut.Visible = False
End Sub

Private Sub CMDClear_Click()
'Me.OPTDriveIn.Value = True
Me.TBArCode.Caption = ""
Me.LBLArName.Caption = ""
Me.TBSearchQueID.Text = ""

Me.ListViewItem.ListItems.Clear
Me.ListViewItem.Visible = False
Me.ListViewMerge.Visible = True
Me.ListViewMerge.ListItems.Clear

Me.PICKeyCheckOut.Visible = False
Me.LBLSaleCode.Caption = ""
Me.PICSelectQue.Visible = False
Me.PICSearchHoldBill.Visible = False

Me.CMDHoldBill.Enabled = False
Me.CMDSelectItemQue.Enabled = False

Me.PICHoldBill.Visible = False
Me.CMDSelectItemQue.Enabled = False
Me.CMDMerge.Enabled = False
Me.CMDHoldBill.Enabled = False

Me.LBLHoldArCode.Caption = ""
Me.LBLHoldArName.Caption = ""
Me.LBLHoldItemAmount.Caption = ""
Me.LBLHoldTaxAmount.Caption = ""
Me.LBLHoldNetAmount.Caption = ""
Me.OPCash1.Value = True
Me.ListViewItemHoldBill.ListItems.Clear
vOpenHoldBill = 0
Me.CMDSearchQueID.Enabled = True
Me.TBSearchQueID.SetFocus
End Sub

Private Sub CMDClear_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Call CMDSearch_Click
End If

If KeyCode = 113 Then
   Me.PICKeySearchData.Visible = True
   Me.TBSearchData.SetFocus
End If

If KeyCode = 114 Then
    Call CMDSelectItemQue_Click
End If

If KeyCode = 27 Then
    Call CMDClear_Click
End If

If KeyCode = 116 Then
    Call CMDHoldBill_Click
End If

If KeyCode = 115 Then
    Call CMDMerge_Click
End If

End Sub

Private Sub CMDDeleteHoldBill_Click()
Dim vQuery As String
Dim vAnswer As Integer
Dim vType As Integer

Dim vDocNo As String

If vOpenHoldBill = 1 And Me.ListViewHoldBill.ListItems.Count > 0 Then
   vDocNo = Me.LBLHoldBillNo.Caption
   vType = Me.LBLHoldBillType.Caption
   vAnswer = MsgBox("คุณต้องการยกเอกสารพักบิลเลขที่ " & vDocNo & " ใช่หรือไม่", vbYesNo, "Send Question Message")
   If vAnswer = 6 Then
   
      vQuery = "exec dbo.USP_NP_DeleteHoldingBill " & vType & ",'" & vDocNo & "' "
      gConnection.Execute vQuery
      
      Me.LBLHoldArCode.Caption = ""
      Me.LBLHoldArName.Caption = ""
      Me.LBLHoldItemAmount.Caption = ""
      Me.LBLHoldTaxAmount.Caption = ""
      Me.LBLHoldNetAmount.Caption = ""
      Me.OPCash1.Value = True
      Me.ListViewItemHoldBill.ListItems.Clear
      
      Me.ListViewItem.ListItems.Clear
      
      Me.ListViewItem.Visible = False
      Me.ListViewMerge.Visible = True
      Me.CMDHoldBill.Enabled = False
      Me.CMDSelectItemQue.Enabled = False
      Me.ListViewMerge.ListItems.Clear
      Me.TBArCode.Caption = ""
      Me.LBLArName.Caption = ""
      Me.PICHoldBill.Visible = False
      Me.CMDSelectItemQue.Enabled = False
      Me.CMDMerge.Enabled = False
      Me.CMDHoldBill.Enabled = False
      Me.CMDSearchQueID.Enabled = True
      
      MsgBox "ได้ลบเลขที่พักบิลเลขที่ " & vDocNo & "  ออกจากระบบแล้ว", vbInformation, "Send Information Message"
   
      vOpenHoldBill = 0
      Me.CMDDeleteHoldBill.Enabled = False
      Me.PICHoldBill.Visible = False
      Me.LBLHoldBillNo.Caption = ""
      Me.LBLHoldBillType.Caption = ""
      
     Me.CMDHoldSave.Enabled = True
     Me.OPCash1.Enabled = True
     Me.OPCash2.Enabled = True
     Me.OPCash3.Enabled = True
        End If
End If
End Sub

Private Sub CMDDeleteHoldBill_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDHoldExit_Click
End If
End Sub

Private Sub CMDExit_Click()
Unload FormCheckOutHoldBill
End Sub

Private Sub CMDExit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Call CMDSearch_Click
End If

If KeyCode = 113 Then
   Me.PICKeySearchData.Visible = True
   Me.TBSearchData.SetFocus
End If

If KeyCode = 114 Then
    Call CMDSelectItemQue_Click
End If

If KeyCode = 27 Then
    Call CMDClear_Click
End If

If KeyCode = 116 Then
    Call CMDHoldBill_Click
End If

If KeyCode = 115 Then
    Call CMDMerge_Click
End If
End Sub

Private Sub CMDHoldBill_Click()
Dim vQuery As String
Dim i As Integer
Dim vListItem As ListItem
Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vAmount As Double
Dim vSumOfItemAmount As Double
Dim vTaxAmount As Double
Dim vNetDebtAmount As Double

Dim vCheckCount As Integer

Dim vMergeNo As String
Dim vCheckQty As Double
Dim vItemCode As String

If Me.ListViewItem.ListItems.Count > 0 Then

   For i = 1 To Me.ListViewItem.ListItems.Count
   If Me.ListViewItem.ListItems(i).SubItems(1) <> "" Then
      vCheckCount = vCheckCount + 1
   End If
   Next i
   
   If vCheckCount <> Me.ListViewItem.ListItems.Count Then
      MsgBox "รายการสินค้ายังนับไม่ครบ กรุณาตรวจสอบ กรณีที่จำนวนเป็น 0 ก็ให้ใส่ 0 ห้ามเป็นค่าว่าง", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
   If Me.LBLHoldArName.Caption <> "" Then
      MsgBox "กรุณาตรวจสอบรหัสลูกค้า ", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
   Me.LBLHoldArCode.Caption = Me.TBArCode.Caption
   Me.LBLHoldArName.Caption = Me.LBLArName.Caption
   Me.LBLHoldSaleCode.Caption = Me.LBLSaleCode.Caption
   Me.LBLHoldCarLicense.Caption = Me.LBLCarLicense.Caption
   Me.OPCash1.Value = True

   Me.ListViewItemHoldBill.ListItems.Clear
   For i = 1 To Me.ListViewItem.ListItems.Count
   Set vListItem = Me.ListViewItemHoldBill.ListItems.Add(, , i)
   vListItem.SubItems(1) = Me.ListViewItem.ListItems(i).SubItems(3)
   vListItem.SubItems(2) = Me.ListViewItem.ListItems(i).SubItems(2)
   
   If Me.ListViewItem.ListItems(i).SubItems(1) <> "" Then
      vQTY = Me.ListViewItem.ListItems(i).SubItems(1)
   End If
   vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
   
   vListItem.SubItems(4) = Me.ListViewItem.ListItems(i).SubItems(4)
   
   If Me.ListViewItem.ListItems(i).SubItems(7) <> "" Then
      vPrice = Me.ListViewItem.ListItems(i).SubItems(7)
   End If
   vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
   
   If Me.ListViewItem.ListItems(i).SubItems(8) <> "" Then
      vDiscountAmount = Me.ListViewItem.ListItems(i).SubItems(8)
   End If
   vListItem.SubItems(6) = Format(vDiscountAmount, "##,##0.00")
   
   vAmount = vQTY * (vPrice - vDiscountAmount)
   vListItem.SubItems(7) = Format(vAmount, "##,##0.00")
   
   vListItem.SubItems(8) = Me.ListViewItem.ListItems(i).SubItems(11)
   vListItem.SubItems(9) = Me.ListViewItem.ListItems(i).SubItems(12)
   vListItem.SubItems(10) = Me.ListViewItem.ListItems(i).SubItems(10)
   vListItem.SubItems(11) = Me.ListViewItem.ListItems(i).SubItems(5)
   vListItem.SubItems(12) = Me.ListViewItem.ListItems(i).SubItems(13)
   vListItem.SubItems(13) = Me.ListViewItem.ListItems(i).SubItems(14)
   vListItem.SubItems(14) = Me.ListViewItem.ListItems(i).SubItems(15)
   
   vSumOfItemAmount = vSumOfItemAmount + vAmount
   vTaxAmount = (vSumOfItemAmount - ((vSumOfItemAmount * 100) / 107))
   vNetDebtAmount = vSumOfItemAmount
   
   Me.LBLHoldItemAmount.Caption = Format(vSumOfItemAmount, "##,##0.00")
   Me.LBLHoldTaxAmount.Caption = Format(vTaxAmount, "##,##0.00")
   Me.LBLHoldNetAmount.Caption = Format(vNetDebtAmount, "##,##0.00")
   
   vMergeNo = Me.ListViewItem.ListItems(i).SubItems(5)
   vItemCode = Me.ListViewItem.ListItems(i).SubItems(3)
   vCheckQty = Me.ListViewItem.ListItems(i).SubItems(1)
      
   vQuery = "exec dbo.USP_NP_UpdateCheckQtyQue '" & vMergeNo & "','" & vUserID & "','" & vItemCode & "'," & vCheckQty & " "
   gConnection.Execute vQuery
   
   Next i
Me.PICHoldBill.Visible = True
Else
MsgBox "ไม่มีรายการสินค้าที่จะทำการสร้างเอกสารพักบิล กรุณาตรวจสอบ", vbCritical, "Send Error Message"
End If


End Sub

Private Sub CMDHoldBill_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Call CMDSearch_Click
End If

If KeyCode = 113 Then
   Me.PICKeySearchData.Visible = True
   Me.TBSearchData.SetFocus
End If

If KeyCode = 114 Then
    Call CMDSelectItemQue_Click
End If

If KeyCode = 27 Then
    Call CMDClear_Click
End If

If KeyCode = 116 Then
    Call CMDHoldBill_Click
End If

If KeyCode = 115 Then
    Call CMDMerge_Click
End If
End Sub

Private Sub CMDHoldExit_Click()
Me.ListViewItemHoldBill.ListItems.Clear
Me.LBLHoldItemAmount.Caption = ""
Me.LBLHoldTaxAmount.Caption = ""
Me.LBLHoldNetAmount.Caption = ""
Me.LBLHoldArCode.Caption = ""
Me.LBLHoldArName.Caption = ""
Me.OPCash1.Value = True
vOpenHoldBill = 0
Me.CMDDeleteHoldBill.Enabled = False
Me.LBLHoldBillNo.Caption = ""
Me.LBLHoldBillType.Caption = ""

Me.CMDHoldSave.Enabled = True
Me.OPCash1.Enabled = True
Me.OPCash2.Enabled = True
Me.OPCash3.Enabled = True

Me.PICHoldBill.Visible = False
Me.ListViewItem.SetFocus

End Sub

Private Sub CMDHoldExit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDHoldExit_Click
End If
End Sub

Private Sub CMDHoldSave_Click()
Dim vCheckCount As Integer
Dim i As Integer
Dim n As Integer

Dim vDocNo As String
Dim vDocdate As String
Dim vExpireCredit As Integer
Dim vARCode As String
Dim vCashierCode As String
Dim vMachineNo As String
Dim vMachineCode As String
Dim vSaleCode As String
Dim vTaxRate As Double
Dim vSumOfItemAmount As Double
Dim vAfterDiscount As Double
Dim vBeforeTaxAmount As Double
Dim vTaxAmount As Double
Dim vTotalAmount As Double
Dim vNetDebtAmount As Double
Dim vCreatorCode As String
Dim vSHIFTCODE As String
Dim vMydescription As String

Dim vMaxNo As Integer
Dim vHeader As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Dim vItemCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vAmount As Double
Dim vNetAmount As Double
Dim vUnitCode As String
Dim vStockType As Integer
Dim vLineNumber As Integer
Dim vBarCode As String
Dim vPosStatus As Integer
Dim vSORefNo As String
Dim vMergeNo As String


If Me.ListViewItemHoldBill.ListItems.Count > 0 Then
   
   If Me.OPCash1.Value = True Then
      vMachineNo = "21"
   ElseIf Me.OPCash2.Value = True Then
      vMachineNo = "22"
   ElseIf Me.OPCash3.Value = True Then
      vMachineNo = "23"
   End If
   
   vDocdate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   
   
   vQuery = "exec dbo.usp_np_getmaxnoholdingbill '" & vMachineNo & "','" & vDocdate & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vMaxNo = vRecordset.Fields("maxnumber").Value
      vHeader = vRecordset.Fields("header").Value
   End If
   vRecordset.Close
   
   vDocNo = vHeader + "-" + Format(vMaxNo, "0000")
   
   vARCode = Me.LBLHoldArCode.Caption
   If vARCode = "1" Then
      vARCode = "999999"
   End If
   vExpireCredit = 1
   
   vQuery = "select top 1 cashiercode,machinecode,shiftcode from dbo.bcarinvoice where docno like '%'+'" & vHeader & "'+'%'  and iscancel = 0 order by createdatetime desc"
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCashierCode = vRecordset.Fields("cashiercode").Value
      vMachineCode = vRecordset.Fields("machinecode").Value
      vSHIFTCODE = vRecordset.Fields("shiftcode").Value
   End If
   vRecordset.Close
   
   vCheckSale = InStr(Me.LBLHoldSaleCode.Caption, "/")
   If vCheckSale = 0 Then
      vSaleCode = ""
   Else
      vSaleCode = Left(Me.LBLHoldSaleCode.Caption, vCheckSale - 1)
   End If
   
   vTaxRate = 7
   If Me.LBLHoldItemAmount.Caption <> "" Then
      vSumOfItemAmount = Me.LBLHoldItemAmount.Caption
   Else
      vSumOfItemAmount = 0
   End If
   
   vAfterDiscount = vSumOfItemAmount
   vBeforeTaxAmount = ((vSumOfItemAmount * 100) / 107)
   vTaxAmount = vSumOfItemAmount - ((vSumOfItemAmount * 100) / 107)
   vNetDebtAmount = vSumOfItemAmount
   vTotalAmount = vSumOfItemAmount
   vCreatorCode = vUserID
   vMydescription = Me.LBLHoldCarLicense.Caption
   
   vQuery = "exec dbo.USP_NP_InsertHoldingBillDriveIn '" & vDocNo & "','" & vDocdate & "'," & vExpireCredit & ",'" & vARCode & "','" & vCashierCode & "','" & vMachineNo & "','" & vMachineCode & "','" & vSaleCode & "'," & vTaxRate & "," & vSumOfItemAmount & "," & vAfterDiscount & "," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalAmount & "," & vNetDebtAmount & ",'" & vCreatorCode & "','" & vSHIFTCODE & "','" & vMydescription & "' "
   gConnection.Execute vQuery
   
   For n = 1 To Me.ListViewItemHoldBill.ListItems.Count
      vItemCode = Me.ListViewItemHoldBill.ListItems(n).SubItems(1)
      vWHCode = Me.ListViewItemHoldBill.ListItems(n).SubItems(8)
      vShelfCode = Me.ListViewItemHoldBill.ListItems(n).SubItems(9)
      vQTY = Me.ListViewItemHoldBill.ListItems(n).SubItems(3)
      vPrice = Me.ListViewItemHoldBill.ListItems(n).SubItems(5)
      vDiscountAmount = Me.ListViewItemHoldBill.ListItems(n).SubItems(6)
      vAmount = Me.ListViewItemHoldBill.ListItems(n).SubItems(7)
      vNetAmount = Me.ListViewItemHoldBill.ListItems(n).SubItems(7)
      vUnitCode = Me.ListViewItemHoldBill.ListItems(n).SubItems(4)
      vStockType = 0
      vLineNumber = n - 1
      vBarCode = Me.ListViewItemHoldBill.ListItems(n).SubItems(10)
      vPosStatus = 1
      vSORefNo = Me.ListViewItemHoldBill.ListItems(n).SubItems(12)
      vMergeNo = Me.ListViewItemHoldBill.ListItems(n).SubItems(11)
      
      vQuery = "exec dbo.USP_NP_InsertHoldingBillDriveInSub '" & vDocNo & "','" & vDocdate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "'," & vQTY & "," & vPrice & "," & vDiscountAmount & "," & vAmount & "," & vNetAmount & ",'" & vUnitCode & "'," & vStockType & "," & vLineNumber & ",'" & vBarCode & "','" & vCashierCode & "'," & vPosStatus & ",'" & vSORefNo & "'"
      gConnection.Execute vQuery
      
      vQuery = "exec dbo.USP_NP_UpdateDriveInMergeTempConfirm '" & vMergeNo & "','" & vDocdate & "','" & vItemCode & "','" & vDocNo & "' "
      gConnection.Execute vQuery
      
      vQuery = "exec dbo.USP_NP_UpdateHoldBillQtyQue '" & vMergeNo & "','" & vDocNo & "','" & vCashierCode & "','" & vItemCode & "'," & vQTY & " "
      gConnection.Execute vQuery
   Next n
   
   
   Me.LBLHoldArCode.Caption = ""
   Me.LBLHoldArName.Caption = ""
   Me.LBLHoldItemAmount.Caption = ""
   Me.LBLHoldTaxAmount.Caption = ""
   Me.LBLHoldNetAmount.Caption = ""
   Me.TBArCode.Caption = ""
   Me.LBLArName.Caption = ""
   Me.LBLSaleCode.Caption = ""
   Me.OPCash1.Value = True
   Me.CMDSearchQueID.Enabled = True
   Me.ListViewItemHoldBill.ListItems.Clear
   
   Me.ListViewItem.ListItems.Clear
   
   Me.ListViewItem.Visible = False
   Me.ListViewMerge.Visible = True
   Me.CMDHoldBill.Enabled = False
   Me.CMDSelectItemQue.Enabled = False
   Me.ListViewMerge.ListItems.Clear
   Me.TBArCode.Caption = ""
   Me.LBLArName.Caption = ""
   Me.TBSearchQueID.Text = ""
   Me.PICHoldBill.Visible = False
   Me.CMDSelectItemQue.Enabled = False
   Me.CMDMerge.Enabled = False
   Me.CMDHoldBill.Enabled = False
   
   MsgBox "ได้เลขที่พักบิลเลขที่ " & vDocNo & " ออกที่จุด POS จุดที่ " & vMachineNo & "", vbInformation, "Send Information Message"
   
   Call PrintCheckOutHeader(vDocNo)
   Call PrintCheckOutItem(vDocNo)
   
End If
End Sub


Private Sub CMDHoldSave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDHoldExit_Click
End If
End Sub

Private Sub CMDMerge_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vHeader As String
Dim vRunning As Integer
Dim vDocNumber As String

Dim vDocdate As String
Dim vItemCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vPickQty As Double
Dim vUnitCode As String
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vAmount As Double
Dim vBarCode As String
Dim vRefNo As String
Dim vQueID As Integer
Dim vLineNumber As Integer

Dim i As Integer
Dim vListItem As ListItem

If Me.ListViewMerge.ListItems.Count > 0 Then

Dim n As Integer
Dim m As Integer
Dim vCheckItemCode As String
Dim vCheckPrice As Double
Dim vCheckDiscount As Double

Dim vLineItem As String
Dim vLinePrice As Double
Dim vLineDiscount As Double
Dim vCountExist As Integer

For m = 1 To Me.ListViewMerge.ListItems.Count
 vCheckItemCode = Me.ListViewMerge.ListItems(m).SubItems(1)
 vCheckPrice = Me.ListViewMerge.ListItems(m).SubItems(5)
 vCheckDiscount = Me.ListViewMerge.ListItems(m).SubItems(6)
 
 For n = 1 To Me.ListViewMerge.ListItems.Count
  vLineItem = Me.ListViewMerge.ListItems(n).SubItems(1)
  vLinePrice = Me.ListViewMerge.ListItems(n).SubItems(5)
  vLineDiscount = Me.ListViewMerge.ListItems(n).SubItems(6)
  
  If vCheckItemCode = vLineItem Then
     If vCheckPrice <> vLinePrice Then
        MsgBox "รายการสินค้า รหัส " & vCheckItemCode & " มีจำนวนรายการมากกว่า 1 รายการ และราคาไม่เท่ากัน โปรแกรมไม่สามรถคิดราคาได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
        Exit Sub
     End If
     
     If vCheckDiscount <> vLineDiscount Then
        MsgBox "รายการสินค้า รหัส " & vCheckItemCode & " มีจำนวนรายการมากกว่า 1 รายการ และส่วนลดไม่เท่ากัน โปรแกรมไม่สามารถคิดส่วนลดได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
        Exit Sub
     End If
  End If
  
 Next n
 
Next m

vQuery = "exec dbo.USP_NP_SearchNewDocNo  30 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vHeader = Trim(vRecordset.Fields("header").Value)
    vRunning = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
End If
vRecordset.Close
      
vDocNo = vDocNumber & vHeader & "-" & Format(vRunning, "0000")
vDocdate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)

For i = 1 To Me.ListViewMerge.ListItems.Count
vItemCode = Me.ListViewMerge.ListItems(i).SubItems(1)
vWHCode = Me.ListViewMerge.ListItems(i).SubItems(8)
vShelfCode = Me.ListViewMerge.ListItems(i).SubItems(9)
vPickQty = Me.ListViewMerge.ListItems(i).SubItems(3)
vUnitCode = Me.ListViewMerge.ListItems(i).SubItems(4)
vPrice = Me.ListViewMerge.ListItems(i).SubItems(5)
vDiscountAmount = Me.ListViewMerge.ListItems(i).SubItems(6)
vAmount = Me.ListViewMerge.ListItems(i).SubItems(7)
vRefNo = Me.ListViewMerge.ListItems(i).SubItems(11)
vQueID = Me.ListViewMerge.ListItems(i).SubItems(12)
vBarCode = Me.ListViewMerge.ListItems(i).SubItems(10)
vLineNumber = i - 1


vQuery = "exec dbo.USP_NP_InsertDriveInMergeTemp '" & vDocNo & "','" & vDocdate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "'," & vPickQty & ",'" & vUnitCode & "'," & vPrice & "," & vDiscountAmount & "," & vAmount & ",'" & vBarCode & "','" & vRefNo & "'," & vQueID & "," & vLineNumber & " "
gConnection.Execute vQuery
Next i

vQuery = "exec dbo.usp_np_updatenewdocno 30"
gConnection.Execute vQuery

Dim vMQty As Double
Dim vMPrice As Double
Dim vMDiscountAmount As Double
Dim vMAmount As Double

Me.ListViewItem.ListItems.Clear
vQuery = "exec dbo.USP_NP_CalcDriveInMergeTemp '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    For i = 1 To vRecordset.RecordCount
       Set vListItem = Me.ListViewItem.ListItems.Add(, , i)
       vListItem.SubItems(1) = ""
       vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
       vListItem.SubItems(3) = Trim(vRecordset.Fields("itemcode").Value)
       vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
       vListItem.SubItems(5) = Trim(vRecordset.Fields("docno").Value)
       If Trim(vRecordset.Fields("qty").Value) <> "" Then
           vMQty = Trim(vRecordset.Fields("qty").Value)
       End If
       vListItem.SubItems(6) = Format(vMQty, "##,##0.00")
       
       If Trim(vRecordset.Fields("price").Value) <> "" Then
          vMPrice = Trim(vRecordset.Fields("price").Value)
       End If
       vListItem.SubItems(7) = Format(vMPrice, "##,##0.00")
       
       If Trim(vRecordset.Fields("discountamount").Value) <> "" Then
          vMDiscountAmount = Trim(vRecordset.Fields("discountamount").Value)
       End If
       vListItem.SubItems(8) = Format(vMDiscountAmount, "##,##0.00")
       
       If Trim(vRecordset.Fields("amount").Value) <> "" Then
          vMAmount = Trim(vRecordset.Fields("amount").Value)
       End If
       vListItem.SubItems(9) = Format(vMAmount, "##,##0.00")
       
       vListItem.SubItems(10) = Trim(vRecordset.Fields("barcode").Value)
       vListItem.SubItems(11) = Trim(vRecordset.Fields("whcode").Value)
       vListItem.SubItems(12) = Trim(vRecordset.Fields("shelfcode").Value)
       vListItem.SubItems(13) = Trim(vRecordset.Fields("refno").Value)
       vListItem.SubItems(14) = Trim(vRecordset.Fields("rate1").Value)
       vListItem.SubItems(15) = Trim(vRecordset.Fields("rate2").Value)
    vRecordset.MoveNext
    Next i
    Me.ListViewItem.Visible = True
    Me.ListViewItem.SetFocus
    Me.ListViewMerge.Visible = False
End If
vRecordset.Close

Me.CMDSelectItemQue.Enabled = True
Me.CMDHoldBill.Enabled = True
Me.CMDMerge.Enabled = False
Me.CMDSearchQueID.Enabled = False
End If
End Sub

Private Sub CMDMerge_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Call CMDSearch_Click
End If

If KeyCode = 113 Then
   Me.PICKeySearchData.Visible = True
   Me.TBSearchData.SetFocus
End If

If KeyCode = 114 Then
    Call CMDSelectItemQue_Click
End If

If KeyCode = 27 Then
    Call CMDClear_Click
End If

If KeyCode = 116 Then
    Call CMDHoldBill_Click
End If

If KeyCode = 115 Then
    Call CMDMerge_Click
End If
End Sub

Private Sub CMDPICCOAddItemClose_Click()
Me.TBCOBarCode.Text = ""
Me.PICCOAddItem.Visible = False
End Sub

Private Sub CMDPICCOAddItemOK_Click()
Dim i As Integer
Dim n As Integer
Dim vListItem As ListItem
Dim vItemCode As String
Dim vQTY As Double
Dim vCount As Integer
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vNetAmount As Double
Dim vRate1 As Integer

Dim vCheckItemCode As String
Dim vCheckExist As Integer



If Me.LBLCOItemCode.Caption <> "" And Me.TBCOKeyQty.Text <> "" Then
    vItemCode = Me.LBLCOItemCode.Caption
    vQTY = Me.TBCOKeyQty.Text
    
    For i = 1 To Me.ListViewItem.ListItems.Count
       vCheckItemCode = Me.ListViewItem.ListItems(i).SubItems(3)
       
       If vCheckItemCode = vItemCode Then
          vCheckExist = 1
          GoTo Line1
       End If
    
    Next i
    
Line1:
    If vCheckExist = 1 Then
       MsgBox "มีรหัสสินค้า รหัส " & vItemCode & " นี้อยู่แล้วในรายการขาย ในบรรทัดที่ " & i & " กรุณาตรวจสอบ", vbCritical, "Send Error Message"
       Me.TBCOBarCode.Text = ""
       Me.TBCOBarCode.SetFocus
       Exit Sub
    Else
    
       If Me.ListViewItem.ListItems.Count > 0 Then
          vCount = Me.ListViewMerge.ListItems.Count
          n = vCount
       End If
       
       n = n + 1
       Set vListItem = Me.ListViewItem.ListItems.Add(, , n)
       vListItem.SubItems(1) = Format(vQTY, "##,##0.00")
       vListItem.SubItems(2) = Me.LBLCOItemName.Caption
       vListItem.SubItems(3) = Me.LBLCOItemCode.Caption
       vListItem.SubItems(4) = Me.LBLCOUnitCode.Caption
       vListItem.SubItems(5) = ""
       vListItem.SubItems(6) = Format(vQTY, "##,##0.00")
       
       If Me.LBLCOPrice.Caption <> "" Then
       vPrice = Me.LBLCOPrice.Caption
       End If
       vListItem.SubItems(7) = Format(vPrice, "##,##0.00")
       vListItem.SubItems(8) = Format(0, "##,##0.00")
       
       If Me.LBLCOItemNetAmount.Caption <> "" Then
          vNetAmount = Me.LBLCOItemNetAmount.Caption
       End If
       vListItem.SubItems(9) = Format(vNetAmount, "##,##0.00")
       
       vListItem.SubItems(10) = Me.LBLCOBarCode.Caption
       vListItem.SubItems(11) = Me.LBLCOWHCode.Caption
       vListItem.SubItems(12) = Me.LBLCOShelfCode.Caption
       vListItem.SubItems(13) = ""
       If Me.LBLCORate1.Caption <> "" Then
         vRate1 = Me.LBLCORate1.Caption
       Else
         vRate1 = 1
       End If
       vListItem.SubItems(14) = Format(vRate1, "##,##0.00")
       vListItem.SubItems(15) = Format(1, "##,##0.00")
       Me.TBCOBarCode.Text = ""
       Me.TBCOBarCode.SetFocus
       End If
    End If
End Sub

Private Sub CMDPICCOAddItemOK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBCOBarCode.Text = ""
Me.PICCOAddItem.Visible = False
Me.ListViewItem.SetFocus
End If
End Sub

Private Sub CMDPrintHoldBill_Click()
Dim vDocNo As String

If Me.LBLHoldBillNo.Caption <> "" Then
vDocNo = Me.LBLHoldBillNo.Caption
Call PrintCheckOutHeader(vDocNo)
Call PrintCheckOutItem(vDocNo)
End If
End Sub

Private Sub CMDPrintHoldBill_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDHoldExit_Click
End If
End Sub

Private Sub CMDSearch_Click()
Me.PICSearchHoldBill.Visible = True
Me.TBSearchHoldBill.SetFocus
Me.ListViewHoldBill.ListItems.Clear
End Sub

Private Sub CMDSearchAR_Click()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vSearchAR As String
'Dim vListItem As ListItem
'Dim i As Integer

'If Me.TXTSearchAR.Text <> "" Then
 '  vSearchAR = Me.TXTSearchAR.Text
  ' vQuery = "exec dbo.USP_AR_ARProFileSearch '" & vSearchAR & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '   Me.ListViewAR.ListItems.Clear
     '  vRecordset.MoveFirst
      ' For i = 1 To vRecordset.RecordCount
       '   Set vListItem = Me.ListViewAR.ListItems.Add(, , i)
        '  vListItem.SubItems(1) = Trim(vRecordset.Fields("arcode").Value)
         ' vListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
          'vListItem.SubItems(3) = Trim(vRecordset.Fields("memberid").Value)
          'vRecordset.MoveNext
       'Next i
       
       'Me.ListViewAR.SetFocus
   'End If
   'vRecordset.Close
'End If
End Sub

Private Sub CMDSearchARClose_Click()
'Me.PICAR.Visible = False
End Sub

Private Sub CMDSearchAROK_Click()
'Dim vIndex As Integer
'Dim vARCode As String

'If Me.ListViewAR.ListItems.Count > 0 Then
 '  vIndex = Me.ListViewAR.SelectedItem.Index
  ' vARCode = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(1))
   'Me.TBArCode.Text = vARCode
   'Me.LBLArName.Caption = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(2))
   'Me.TBArCode.SetFocus
'End If
'Me.PICAR.Visible = False
End Sub

Private Sub CMDSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Call CMDSearch_Click
End If

If KeyCode = 113 Then
   Me.PICKeySearchData.Visible = True
   Me.TBSearchData.SetFocus
End If

If KeyCode = 114 Then
    Call CMDSelectItemQue_Click
End If

If KeyCode = 27 Then
    Call CMDClear_Click
End If

If KeyCode = 116 Then
    Call CMDHoldBill_Click
End If

If KeyCode = 115 Then
    Call CMDMerge_Click
End If
End Sub

Private Sub CMDSearchDataClose_Click()
Me.TBSearchQueID.SetFocus
Me.PICKeySearchData.Visible = False
End Sub

Private Sub CMDSearchDataClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBSearchQueID.SetFocus
Me.PICKeySearchData.Visible = False
End If
End Sub

Private Sub CMDSearchHoldBill_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem
Dim vSearch As String

Dim vNetDebtAmount As Double

   vSearch = Me.TBSearchHoldBill.Text
   vQuery = "exec dbo.USP_NP_SearchHodingBill '" & vSearch & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.ListViewHoldBill.ListItems.Clear
      vRecordset.MoveFirst
      For i = 1 To vRecordset.RecordCount
      Set vListItem = Me.ListViewHoldBill.ListItems.Add(, , i)
      vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
      vListItem.SubItems(3) = Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)
      If Trim(vRecordset.Fields("salecode").Value) <> "" Then
      vListItem.SubItems(4) = Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      Else
      vListItem.SubItems(4) = ""
      End If
      vNetDebtAmount = Trim(vRecordset.Fields("netdebtamount").Value)
      vListItem.SubItems(5) = Format(vNetDebtAmount, "##,##0.00")
      vListItem.SubItems(6) = Trim(vRecordset.Fields("cashiername").Value)
      vListItem.SubItems(7) = ""
      vRecordset.MoveNext
      Next i
      
      If Me.ListViewHoldBill.ListItems.Count > 0 Then
      Me.ListViewHoldBill.SetFocus
      Else
      Me.TBSearchHoldBill.SetFocus
      End If
   End If
   vRecordset.Close
End Sub

Private Sub CMDSearchHoldBill_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSearchHoldBill.Visible = False
Me.TBSearchQueID.SetFocus
End If
End Sub

Private Sub CMDSearchHoldBillExit_Click()
Me.ListViewHoldBill.ListItems.Clear
Me.TBSearchHoldBill.Text = ""
PICSearchHoldBill.Visible = False
End Sub

Private Sub CMDSearchHoldBillExit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSearchHoldBill.Visible = False
Me.TBSearchQueID.SetFocus
End If
End Sub

Private Sub CMDSearchHoldBillOK_Click()
Dim vIndex As Integer
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem
Dim vDocNo As String

Dim vSumOfItemAmount As Double
Dim vTaxAmount As Double
Dim vNetDebtAmount As Double
Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vAmount As Double
Dim vRate1 As Integer
Dim vRate2 As Integer

Me.ListViewItemHoldBill.ListItems.Clear
If Me.ListViewHoldBill.ListItems.Count > 0 Then
   vOpenHoldBill = 1
   Me.CMDDeleteHoldBill.Enabled = True
   vIndex = Me.ListViewHoldBill.SelectedItem.Index
   vDocNo = Me.ListViewHoldBill.ListItems(vIndex).SubItems(1)
   vQuery = "exec dbo.USP_NP_SearchHodingBillDetails '" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vRecordset.MoveFirst
      Me.PICHoldBill.Visible = True
      Me.PICSearchHoldBill.Visible = False
      
      Me.LBLHoldBillNo.Caption = vDocNo
      Me.LBLHoldBillType.Caption = vRecordset.Fields("type").Value
      vSumOfItemAmount = vRecordset.Fields("sumofitemamount").Value
      vTaxAmount = vRecordset.Fields("taxamount").Value
      vNetDebtAmount = vRecordset.Fields("totalamount").Value
      
      Me.LBLHoldArCode.Caption = vRecordset.Fields("arcode").Value
      Me.LBLHoldArName.Caption = vRecordset.Fields("arname").Value
      Me.LBLHoldSaleCode.Caption = vRecordset.Fields("salecode").Value & "/ " & vRecordset.Fields("salename").Value
      Me.LBLHoldCarLicense.Caption = vRecordset.Fields("mydescription").Value
      
      If vRecordset.Fields("machineno").Value = "11" Or vRecordset.Fields("machineno").Value = "21" Then
         Me.OPCash1.Value = True
      ElseIf vRecordset.Fields("machineno").Value = "12" Or vRecordset.Fields("machineno").Value = "22" Then
         Me.OPCash2.Value = True
      ElseIf vRecordset.Fields("machineno").Value = "13" Or vRecordset.Fields("machineno").Value = "23" Then
         Me.OPCash3.Value = True
      End If

      Me.LBLHoldItemAmount.Caption = Format(vSumOfItemAmount, "##,##0.00")
      Me.LBLHoldTaxAmount.Caption = Format(vTaxAmount, "##,##0.00")
      Me.LBLHoldNetAmount.Caption = Format(vNetDebtAmount, "##,##0.00")
      
      For i = 1 To vRecordset.RecordCount
      vQTY = vRecordset.Fields("qty").Value
      vPrice = vRecordset.Fields("price").Value
      vDiscountAmount = vRecordset.Fields("discountamount").Value
      vAmount = vRecordset.Fields("amount").Value
      
      Set vListItem = Me.ListViewItemHoldBill.ListItems.Add(, , i)
      vListItem.SubItems(1) = vRecordset.Fields("itemcode").Value
      vListItem.SubItems(2) = vRecordset.Fields("itemname").Value
      vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
      vListItem.SubItems(4) = vRecordset.Fields("unitcode").Value
      vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
      vListItem.SubItems(6) = Format(vDiscountAmount, "##,##0.00")
      vListItem.SubItems(7) = Format(vAmount, "##,##0.00")
      vListItem.SubItems(8) = vRecordset.Fields("whcode").Value
      vListItem.SubItems(9) = vRecordset.Fields("shelfcode").Value
      vListItem.SubItems(10) = vRecordset.Fields("barcode").Value
      vListItem.SubItems(11) = vRecordset.Fields("sorefno").Value
      vListItem.SubItems(12) = vRecordset.Fields("sorefno").Value
      vRate1 = vRecordset.Fields("rate1").Value
      vRate1 = vRecordset.Fields("rate2").Value
      vListItem.SubItems(13) = Format(vRate1, "##,##0.00")
      vListItem.SubItems(14) = Format(vRate2, "##,##0.00")
       vRecordset.MoveNext
      Next i
      
   End If
   vRecordset.Close
   
   Me.PICSearchHoldBill.Visible = False
   Me.PICHoldBill.Visible = True
   Me.CMDHoldSave.Enabled = False
   Me.OPCash1.Enabled = False
   Me.OPCash2.Enabled = False
   Me.OPCash3.Enabled = False
End If
End Sub

Private Sub CMDSearchHoldBillOK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSearchHoldBill.Visible = False
Me.TBSearchQueID.SetFocus
End If
End Sub

Private Sub CMDSearchQueID_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchQue As String
Dim vListItem As ListItem
Dim i As Integer
Dim vType As Integer

Dim vMemStatus As Integer
Dim vMemQty As Double
Dim vMemPickQty As Double

Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vNetAmount As Double
Dim vQTY As Double
Dim vPickQty As Double

 If Me.TBSearchQueID.Text = "" Then
    MsgBox "กรุณากรอกข้อมูลที่ต้องการหา เช่น รหัสลูกค้า  เลขที่เอกสาร ทะเบียนรถ เป็นต้น", vbCritical, "Send Information Message"
    Me.TBSearchQueID.SetFocus
    Exit Sub
 End If
 
 'If Me.OPTPickReq.Value = True Then
  '  vType = 1
 'ElseIf Me.OPTSaleOrder.Value = True Then
  '  vType = 2
 'ElseIf Me.OPTDriveIn.Value = True Then
  '   vType = 3
 'End If
 
vSearchQue = Me.TBSearchQueID.Text
Me.ListViewSelectQue.ListItems.Clear

vQuery = "exec dbo.USP_NP_SearchQueCheckOut '" & vSearchQue & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    For i = 1 To vRecordset.RecordCount
       Set vListItem = Me.ListViewSelectQue.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
       vListItem.SubItems(1) = Trim(vRecordset.Fields("queid").Value)
       
       vMemStatus = Trim(vRecordset.Fields("questatus").Value)
       vMemQty = Trim(vRecordset.Fields("qty").Value)
       vMemPickQty = Trim(vRecordset.Fields("oncarqty").Value)
       If vMemStatus = 2 And vMemQty = vMemPickQty Then
       vListItem.SubItems(2) = "ครบ"
       ElseIf vMemStatus = 2 And vMemQty < vMemPickQty Then
       vListItem.SubItems(2) = "เกิน"
       ElseIf vMemStatus = 2 And vMemQty > vMemPickQty Then
       vListItem.SubItems(2) = "ไม่ครบ"
       Else
       vListItem.SubItems(2) = Trim(vRecordset.Fields("quedescription").Value)
       End If
       vListItem.SubItems(3) = Trim(vRecordset.Fields("itemcode").Value)
       vListItem.SubItems(4) = Trim(vRecordset.Fields("itemname").Value)
       vListItem.SubItems(5) = Trim(vRecordset.Fields("quepicker").Value)
       vListItem.SubItems(6) = Trim(vRecordset.Fields("questatus").Value)
       
       vQTY = Trim(vRecordset.Fields("qty").Value)
       vPickQty = Trim(vRecordset.Fields("oncarqty").Value)
       vPrice = Trim(vRecordset.Fields("price").Value)
       vDiscountAmount = Trim(vRecordset.Fields("discountamount").Value)
       vNetAmount = Trim(vRecordset.Fields("netamount").Value)
       
       vListItem.SubItems(7) = Format(vQTY, "##,##0.00")
       vListItem.SubItems(8) = Format(vPickQty, "##,##0.00")
       vListItem.SubItems(9) = Trim(vRecordset.Fields("unitcode").Value)
       
       vListItem.SubItems(10) = Format(vPrice, "##,##0.00")
       vListItem.SubItems(11) = Format(vDiscountAmount, "##,##0.00")
       vListItem.SubItems(12) = Format(vNetAmount, "##,##0.00")
       vListItem.SubItems(13) = Trim(vRecordset.Fields("barcode").Value)
       vListItem.SubItems(14) = Trim(vRecordset.Fields("whcode").Value)
       vListItem.SubItems(15) = Trim(vRecordset.Fields("shelfcode").Value)
       vListItem.SubItems(16) = Trim(vRecordset.Fields("arcode").Value)
       vListItem.SubItems(17) = Trim(vRecordset.Fields("salecode").Value)
       vListItem.SubItems(18) = Trim(vRecordset.Fields("refno").Value)
       vRecordset.MoveNext
    Next i
End If
vRecordset.Close


Dim n As Integer
Dim vCheckQty As Double
Dim vCheckPickQty As Double
Dim vStatus As Integer

For n = 1 To Me.ListViewSelectQue.ListItems.Count
   vStatus = Me.ListViewSelectQue.ListItems(n).SubItems(6)
   vCheckQty = Me.ListViewSelectQue.ListItems(n).SubItems(7)
   vCheckPickQty = Me.ListViewSelectQue.ListItems(n).SubItems(8)
 
   If vStatus = 2 Or vStatus = 3 Then
      If vCheckQty <> vCheckPickQty Then
         ListViewSelectQue.ListItems(n).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(1).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(2).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(3).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(4).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(5).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(6).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(7).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(8).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(9).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(10).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(11).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(12).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(13).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(14).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(15).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(16).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(17).ForeColor = "&H000000FF"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(18).ForeColor = "&H000000FF"
      Else
         ListViewSelectQue.ListItems(n).Checked = True
         ListViewSelectQue.ListItems(n).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(1).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(2).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(3).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(4).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(5).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(6).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(7).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(8).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(9).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(10).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(11).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(12).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(13).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(14).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(15).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(16).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(17).ForeColor = "&H00004000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(18).ForeColor = "&H00004000"
      End If
   ElseIf vStatus = 1 Then
         ListViewSelectQue.ListItems(n).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(1).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(2).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(3).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(4).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(5).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(6).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(7).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(8).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(9).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(10).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(11).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(12).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(13).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(14).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(15).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(16).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(17).ForeColor = "&H00000000"
         ListViewSelectQue.ListItems.Item(n).ListSubItems(18).ForeColor = "&H00000000"
   End If
      
Next n

Me.PICSelectQue.Visible = True
Me.ListViewSelectQue.SetFocus
End Sub

Private Sub CMDSearchQueID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Call CMDSearch_Click
End If

If KeyCode = 113 Then
   Me.PICKeySearchData.Visible = True
   Me.TBSearchData.SetFocus
End If

If KeyCode = 114 Then
    Call CMDSelectItemQue_Click
End If

If KeyCode = 27 Then
    Call CMDClear_Click
End If

If KeyCode = 116 Then
    Call CMDHoldBill_Click
End If

If KeyCode = 115 Then
    Call CMDMerge_Click
End If
End Sub

Private Sub CMDSelectItem_Click()
Dim i As Integer
Dim n As Integer
Dim m As Integer
Dim vListItem As ListItem
Dim vPickQty As Double
Dim vCount As Integer
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vNetAmount As Double

Dim vCheckDocNo As String
Dim vCheckItemCode As String
Dim vCheckQty As Double
Dim vMemDocNo As String
Dim vMemItemCode As String
Dim vMemQty As Double

If Me.ListViewSelectQue.ListItems.Count > 0 Then
   
   For m = 1 To Me.ListViewSelectQue.ListItems.Count
   If Me.ListViewSelectQue.ListItems(m).Checked = True Then
      vMemDocNo = Me.ListViewSelectQue.ListItems(m).Text
      vMemItemCode = Me.ListViewSelectQue.ListItems(m).SubItems(3)
      vMemQty = Me.ListViewSelectQue.ListItems(m).SubItems(8)
      If Me.TBArCode.Caption = "" Then
      Me.TBArCode.Caption = Me.ListViewSelectQue.ListItems(m).SubItems(16)
      End If
      If Me.LBLSaleCode.Caption = "" Then
      Me.LBLSaleCode.Caption = Me.ListViewSelectQue.ListItems(m).SubItems(17)
      End If
      If Me.LBLCarLicense.Caption = "" Then
      Me.LBLCarLicense.Caption = Me.ListViewSelectQue.ListItems(m).SubItems(18)
      End If
      
      For n = 1 To Me.ListViewMerge.ListItems.Count
         vCheckDocNo = Me.ListViewMerge.ListItems(n).SubItems(11)
         vCheckItemCode = Me.ListViewMerge.ListItems(n).SubItems(1)
         vCheckQty = Me.ListViewMerge.ListItems(n).SubItems(3)
         
         If vCheckDocNo = vMemDocNo And vCheckItemCode = vMemItemCode And vCheckQty = vMemQty Then
         GoTo Line1
         End If
         
      Next n
      
      
      If Me.ListViewMerge.ListItems.Count > 0 Then
         vCount = Me.ListViewMerge.ListItems.Count
         i = vCount
      End If
      
      i = i + 1
      Set vListItem = Me.ListViewMerge.ListItems.Add(, , i)
      vPickQty = Me.ListViewSelectQue.ListItems(m).SubItems(8)
      vPrice = Me.ListViewSelectQue.ListItems(m).SubItems(10)
      vDiscountAmount = Me.ListViewSelectQue.ListItems(m).SubItems(11)
      vNetAmount = Me.ListViewSelectQue.ListItems(m).SubItems(12)
      
      vListItem.SubItems(1) = Me.ListViewSelectQue.ListItems(m).SubItems(3)
      vListItem.SubItems(2) = Me.ListViewSelectQue.ListItems(m).SubItems(4)
      vListItem.SubItems(3) = Format(vPickQty, "##,##0.00")
      vListItem.SubItems(4) = Me.ListViewSelectQue.ListItems(m).SubItems(9)
      vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
      vListItem.SubItems(6) = Format(vDiscountAmount, "##,##0.00")
      vListItem.SubItems(7) = Format(vNetAmount, "##,##0.00")
      vListItem.SubItems(8) = Me.ListViewSelectQue.ListItems(m).SubItems(14)
      vListItem.SubItems(9) = Me.ListViewSelectQue.ListItems(m).SubItems(15)
      vListItem.SubItems(10) = Me.ListViewSelectQue.ListItems(m).SubItems(13)
      vListItem.SubItems(11) = Me.ListViewSelectQue.ListItems(m).Text
      vListItem.SubItems(12) = Me.ListViewSelectQue.ListItems(m).SubItems(1)
   End If
   
Line1:
   Next m
End If


Me.PICSelectQue.Visible = False
Me.CMDMerge.Enabled = True
Me.ListViewSelectQue.ListItems.Clear
Me.CMDMerge.SetFocus
End Sub

Private Sub CMDSelectItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICSelectQue.Visible = False
   Me.TBSearchQueID.SetFocus
End If
End Sub

Private Sub CMDSelectItemClose_Click()
Me.PICSelectQue.Visible = False
Me.TBSearchQueID.SetFocus
End Sub

Private Sub CMDSelectItemClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICSelectQue.Visible = False
   Me.TBSearchQueID.SetFocus
End If
End Sub

Private Sub CMDSelectItemQue_Click()
Me.ListViewMerge.Visible = True
Me.ListViewItem.ListItems.Clear
Me.ListViewItem.Visible = False
Me.CMDHoldBill.Enabled = False
Me.CMDSelectItemQue.Enabled = False
Me.CMDMerge.Enabled = True
Me.CMDSearchQueID.Enabled = True
End Sub


Private Sub ListViewAR_DblClick()
'Dim vIndex As Integer
'Dim vARCode As String
'
'If Me.ListViewAR.ListItems.Count > 0 Then
 '  vIndex = Me.ListViewAR.SelectedItem.Index
  ' vARCode = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(1))
   'Me.TBArCode.Text = vARCode
   ''Me.LBLArName.Caption = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(2))
   'Me.TBArCode.SetFocus
'End If
'Me.PICAR.Visible = False
End Sub

Private Sub ListViewAR_KeyPress(KeyAscii As Integer)
'Dim vIndex As Integer
'Dim vARCode As String

'If KeyAscii = 13 Then
 '  If Me.ListViewAR.ListItems.Count > 0 Then
  '    vIndex = Me.ListViewAR.SelectedItem.Index
   '   vARCode = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(1))
    '  Me.TBArCode.Text = vARCode
     '' Me.LBLArName.Caption = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(2))
      'Me.TBArCode.SetFocus
   'End If
   'Me.PICAR.Visible = False
'End If
End Sub

Private Sub CMDSelectItemQue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Call CMDSearch_Click
End If

If KeyCode = 113 Then
   Me.PICKeySearchData.Visible = True
   Me.TBSearchData.SetFocus
End If

If KeyCode = 114 Then
    Call CMDSelectItemQue_Click
End If

If KeyCode = 27 Then
    Call CMDClear_Click
End If

If KeyCode = 116 Then
    Call CMDHoldBill_Click
End If

If KeyCode = 115 Then
    Call CMDMerge_Click
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   Me.PICKeySearchData.Visible = True
   Me.TBSearchData.SetFocus
End If
End Sub

Private Sub Label23_Click()
'Me.OPTSaleOrder.Value = False
'Me.OPTPickReq.Value = True
'Me.OPTDriveIn.Value = False
End Sub

Private Sub Label27_Click()
'Me.OPTSaleOrder.Value = False
'Me.OPTPickReq.Value = False
'Me.OPTDriveIn.Value = True
End Sub

Private Sub Label8_Click()
'Me.OPTSaleOrder.Value = True
'Me.OPTPickReq.Value = False
'Me.OPTDriveIn.Value = False
End Sub

Private Sub LBLSaleCode_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchSale As String
Dim vSaleCode As String
Dim vLen As Integer
Dim vInstr As Integer

If Me.LBLSaleCode.Caption <> "" Then
   vSearchSale = Me.LBLSaleCode.Caption
   If InStr(vSearchSale, "/") <> 0 Then
      vInstr = InStr(vSearchSale, "/")
      vLen = Len(vSearchSale)
      vSaleCode = Left(vSearchSale, vInstr - 1)
      vQuery = "exec dbo.USP_CRM_EmployeeDetails 1,'" & vSaleCode & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         Me.LBLSaleCode.Caption = Trim(vRecordset.Fields("empcode").Value) & "/" & Trim(vRecordset.Fields("empname").Value)
      Else
         Me.LBLSaleCode.Caption = vSaleCode
      End If
      vRecordset.Close
   Else
      vQuery = "exec dbo.USP_CRM_EmployeeDetails 1,'" & vSearchSale & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         Me.LBLSaleCode.Caption = Trim(vRecordset.Fields("empcode").Value) & "/" & Trim(vRecordset.Fields("empname").Value)
      End If
      vRecordset.Close
   End If
End If
End Sub

Private Sub ListViewHoldBill_DblClick()
If Me.ListViewHoldBill.ListItems.Count > 0 Then
Call CMDSearchHoldBillOK_Click
End If
End Sub

Private Sub ListViewHoldBill_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSearchHoldBill.Visible = False
Me.TBSearchQueID.SetFocus
End If
End Sub

Private Sub ListViewHoldBill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Me.ListViewHoldBill.ListItems.Count > 0 Then
Call CMDSearchHoldBillOK_Click
End If
End If
End Sub

Private Sub ListViewItem_DblClick()
Dim vIndex As Integer
Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double

If Me.ListViewItem.ListItems.Count > 0 Then
   vIndex = Me.ListViewItem.SelectedItem.Index
   Me.PICKeyCheckOut.Visible = True
   Me.LBLIndex.Caption = vIndex
   
   Me.LBLItemCode.Caption = Me.ListViewItem.ListItems(vIndex).SubItems(3)
   Me.LBLItemName.Caption = Me.ListViewItem.ListItems(vIndex).SubItems(2)
   Me.LBLUnit.Caption = Me.ListViewItem.ListItems(vIndex).SubItems(4)
   If Me.ListViewItem.ListItems(vIndex).SubItems(1) <> "" Then
   vQTY = Me.ListViewItem.ListItems(vIndex).SubItems(1)
   End If
   vPrice = Me.ListViewItem.ListItems(vIndex).SubItems(7)
   vDiscountAmount = Me.ListViewItem.ListItems(vIndex).SubItems(8)
   Me.LBLPrice.Caption = Format(vPrice, "##,##0.00")
   Me.LBLDisCount.Caption = Format(vDiscountAmount, "##,##0.00")
   If vQTY > 0 Then
   Me.TBKeyQty.Text = vQTY
   Else
   Me.TBKeyQty.Text = ""
   End If
   
   Me.PICKeyCheckOut.Visible = True
   Me.TBKeyQty.SetFocus
End If
End Sub


Private Sub ListViewItem_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer
Dim i As Integer

If Me.ListViewItem.ListItems.Count > 0 Then
   If KeyCode = 46 Then
      vIndex = Me.ListViewItem.SelectedItem.Index
      Me.ListViewItem.ListItems.Remove (vIndex)
      
      If Me.ListViewItem.ListItems.Count > 0 Then
         For i = 1 To Me.ListViewItem.ListItems.Count
         Me.ListViewItem.ListItems(i).Text = i
         Next i
      End If
   End If
End If


If KeyCode = 118 Then
Me.PICKeyCheckOutQTY.Visible = True
Me.TBCheckOutItemCode.SetFocus
End If

If KeyCode = 113 Then
   Me.PICKeySearchData.Visible = True
   Me.TBSearchData.SetFocus
End If


If KeyCode = 120 Then
  Me.PICCOAddItem.Visible = True
  Me.TBCOBarCode.SetFocus
End If


If KeyCode = 112 Then
    Call CMDSearch_Click
End If


If KeyCode = 114 Then
    Call CMDSelectItemQue_Click
End If

If KeyCode = 115 Then
    Call CMDMerge_Click
End If

If KeyCode = 27 Then
    Call CMDClear_Click
End If

If KeyCode = 116 Then
    Call CMDHoldBill_Click
End If

End Sub

Private Sub ListViewItem_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer
Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double

If KeyAscii = 13 Then
    If Me.ListViewItem.ListItems.Count > 0 Then
       vIndex = Me.ListViewItem.SelectedItem.Index
       Me.PICKeyCheckOut.Visible = True
       Me.LBLIndex.Caption = vIndex
       
       Me.LBLItemCode.Caption = Me.ListViewItem.ListItems(vIndex).SubItems(3)
       Me.LBLItemName.Caption = Me.ListViewItem.ListItems(vIndex).SubItems(2)
       Me.LBLUnit.Caption = Me.ListViewItem.ListItems(vIndex).SubItems(4)
       If Me.ListViewItem.ListItems(vIndex).SubItems(1) <> "" Then
       vQTY = Me.ListViewItem.ListItems(vIndex).SubItems(1)
       End If
       vPrice = Me.ListViewItem.ListItems(vIndex).SubItems(7)
       vDiscountAmount = Me.ListViewItem.ListItems(vIndex).SubItems(8)
       Me.LBLPrice.Caption = Format(vPrice, "##,##0.00")
       Me.LBLDisCount.Caption = Format(vDiscountAmount, "##,##0.00")
       If vQTY > 0 Then
       Me.TBKeyQty.Text = vQTY
       Else
       Me.TBKeyQty.Text = ""
       End If
       
       Me.PICKeyCheckOut.Visible = True
       Me.TBKeyQty.SetFocus
    End If
End If
End Sub

Private Sub ListViewItemHoldBill_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDHoldExit_Click
End If
End Sub

Private Sub ListViewMerge_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer
Dim i As Integer

If Me.ListViewMerge.ListItems.Count > 0 Then
   If KeyCode = 46 Then
      vIndex = Me.ListViewMerge.SelectedItem.Index
      Me.ListViewMerge.ListItems.Remove (vIndex)
      
      If Me.ListViewMerge.ListItems.Count > 0 Then
         For i = 1 To Me.ListViewMerge.ListItems.Count
         Me.ListViewMerge.ListItems(i).Text = i
         Next i
      End If
   End If
End If

If KeyCode = 112 Then
    Call CMDSearch_Click
End If

If KeyCode = 113 Then
   Me.PICKeySearchData.Visible = True
   Me.TBSearchData.SetFocus
End If

If KeyCode = 114 Then
    Call CMDSelectItemQue_Click
End If

If KeyCode = 27 Then
    Call CMDClear_Click
End If

If KeyCode = 116 Then
    Call CMDHoldBill_Click
End If

If KeyCode = 115 Then
    Call CMDMerge_Click
End If
End Sub

Private Sub ListViewSelectQue_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim n As Integer
Dim vQTY As Double
Dim vPickQty As Double
Dim vStatus As Integer


If Me.ListViewSelectQue.ListItems.Count > 0 Then
   n = Item.Index
   vStatus = Me.ListViewSelectQue.ListItems(n).SubItems(6)
   vQTY = Me.ListViewSelectQue.ListItems(n).SubItems(7)
   vPickQty = Me.ListViewSelectQue.ListItems(n).SubItems(8)
   
   If Me.ListViewSelectQue.ListItems(n).Checked = True Then
      If vStatus = 2 Or vStatus = 3 Then
         If vQTY > vPickQty And vPickQty > 0 Then
         MsgBox "สินค้ารายการนี้ จัดได้ไม่ครบ", vbCritical, "Send Information Message"
         End If
         
         If vPickQty = 0 Then
         MsgBox "สินค้ารายการนี้ ยอดสต๊อกไม่มี กรุณาตรวจสอบ", vbCritical, "Send Information Message"
         Me.ListViewSelectQue.ListItems(n).Checked = False
         End If
         
         If vQTY < vPickQty Then
         MsgBox "สินค้ารายการนี้ จัดเกิน กรุณาตรวจสอบ", vbCritical, "Send Information Message"
         End If
      Else
         MsgBox "สินค้ารายการนี้ ยังจัดไม่ครบ", vbCritical, "Send Information Message"
         Me.ListViewSelectQue.ListItems(n).Checked = False
      End If
   End If
End If
End Sub

Private Sub ListViewSelectQue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICSelectQue.Visible = False
   Me.TBSearchQueID.SetFocus
End If
End Sub

Private Sub OPCash1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDHoldExit_Click
End If
End Sub

Private Sub OPCash2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDHoldExit_Click
End If
End Sub

Private Sub OPTDriveIn_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then
 '   Call CMDSearch_Click
'End If

'If KeyCode = 113 Then
 '  Me.PICKeySearchData.Visible = True
  ' Me.TBSearchData.SetFocus
'End If

'If KeyCode = 114 Then
 '   Call CMDSelectItemQue_Click
'End If

'If KeyCode = 27 Then
 '   Call CMDClear_Click
'End If

'If KeyCode = 116 Then
 '   Call CMDHoldBill_Click
'End If

'If KeyCode = 115 Then
 '   Call CMDMerge_Click
'End If
End Sub

Private Sub OPTPickReq_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then
 '   Call CMDSearch_Click
'End If

'If KeyCode = 113 Then
 '  Me.PICKeySearchData.Visible = True
  ' Me.TBSearchData.SetFocus
'End If

'If KeyCode = 114 Then
 '   Call CMDSelectItemQue_Click
'End If

'If KeyCode = 27 Then
 '   Call CMDClear_Click
'End If

'If KeyCode = 116 Then
 '   Call CMDHoldBill_Click
'End If

'If KeyCode = 115 Then
 '   Call CMDMerge_Click
'End If
End Sub

Private Sub OPTSaleOrder_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then
 '   Call CMDSearch_Click
'End If

'If KeyCode = 113 Then
 '  Me.PICKeySearchData.Visible = True
  ' Me.TBSearchData.SetFocus
'End If

'If KeyCode = 114 Then
 '   Call CMDSelectItemQue_Click
'End If

'If KeyCode = 27 Then
 '   Call CMDClear_Click
'End If

'If KeyCode = 116 Then
 '   Call CMDHoldBill_Click
'End If

'If KeyCode = 115 Then
 '   Call CMDMerge_Click
'End If
End Sub

Private Sub PICCOAddItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBCOBarCode.Text = ""
Me.PICCOAddItem.Visible = False
Me.ListViewItem.SetFocus
End If
End Sub

Private Sub PICHoldBill_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Call CMDHoldExit_Click
End If
End Sub

Private Sub PICKeyCheckOutQTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICKeyCheckOutQTY.Visible = False
Me.ListViewItem.SetFocus
End If
End Sub

Private Sub PICKeySearchData_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBSearchQueID.SetFocus
Me.PICKeySearchData.Visible = False
End If
End Sub

Private Sub PICSearchHoldBill_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSearchHoldBill.Visible = False
Me.TBSearchQueID.SetFocus
End If
End Sub

Private Sub PICSelectQue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICSelectQue.Visible = False
   Me.TBSearchQueID.SetFocus
End If
End Sub

Private Sub TBArCode_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String

If Me.TBArCode.Caption <> "" Then
   vSearchAR = Me.TBArCode.Caption
   vQuery = "exec dbo.usp_ar_arprofile '" & vSearchAR & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.LBLArName.Caption = Trim(vRecordset.Fields("arname").Value)
   Else
      Me.LBLArName.Caption = ""
   End If
   vRecordset.Close
Else
   Me.LBLArName.Caption = ""
End If
End Sub

Private Sub TBCheckOutItemCode_Change()
If Me.TBCheckOutItemCode.Text = "" Then
   Me.LBLCheckOutItemCode.Caption = ""
   Me.LBLCheckOutItemName.Caption = ""
   Me.TBCheckOutItemQty.Text = ""
   Me.TBCheckOutItemCode.SetFocus
End If
End Sub

Private Sub TBCheckOutItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICKeyCheckOutQTY.Visible = False
Me.ListViewItem.SetFocus
End If
End Sub

Private Sub TBCheckOutItemCode_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vBarCode As String

If KeyAscii = 13 Then
      vBarCode = Me.TBCheckOutItemCode.Text
      vQuery = "exec dbo.USP_MB_SearchBarcode '" & vBarCode & "'"
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Me.LBLCheckOutItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
          Me.LBLCheckOutItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
          Me.TBCheckOutItemQty.SetFocus
      Else
         Me.TBCheckOutItemCode.SetFocus
      End If
      vRecordset.Close
End If
End Sub

Private Sub TBCheckOutItemQty_Change()
Dim vQtyWord As String
Dim vLenQTY As Integer

If Me.TBCheckOutItemQty.Text <> "" Then
   vQtyWord = Me.TBCheckOutItemQty.Text
   CheckNumber (vQtyWord)
      
   If vCheckValueNumber = False Then
      MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      vLenQTY = Len(Me.TBCheckOutItemQty.Text)
      Me.TBCheckOutItemQty.Text = Left(Me.TBCheckOutItemQty.Text, vLenQTY - 1)
      Me.TBCheckOutItemQty.SetFocus
   End If
End If
End Sub

Private Sub TBCheckOutItemQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICKeyCheckOutQTY.Visible = False
Me.ListViewItem.SetFocus
End If
End Sub

Private Sub TBCheckOutItemQty_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim n As Integer
Dim vCheckItemCode As String
Dim vItemCode As String
Dim vQTY As Double
Dim vCheckNotExist As Integer

If KeyAscii = 13 Then
If Me.TBCheckOutItemQty.Text <> "" Then
   If Me.LBLCheckOutItemCode.Caption <> "" Then
      vItemCode = Me.LBLCheckOutItemCode.Caption
      For i = 1 To Me.ListViewItem.ListItems.Count
         vCheckItemCode = Me.ListViewItem.ListItems(i).SubItems(3)
         If vItemCode = vCheckItemCode Then
            vQTY = Me.TBCheckOutItemQty.Text
            Me.ListViewItem.ListItems(i).SubItems(1) = Format(vQTY, "##,##0.00")
            Me.TBCheckOutItemCode.Text = ""
            Me.LBLCheckOutItemCode.Caption = ""
            Me.LBLCheckOutItemName.Caption = ""
            Me.TBCheckOutItemQty.Text = ""
            Me.TBCheckOutItemCode.SetFocus
            Exit Sub
         End If
         vCheckNotExist = 1
      Next i
      
      If vCheckNotExist = 1 Then
         MsgBox "ไม่มีรหัสสินค้า " & vItemCode & " ในรายการตรวจนับสินค้า กรุณาตรวจสอบ กรณีต้องการเพิ่มสินค้าให้กดปุ่ม ESC แล้วกดปุ่ม F9 ", vbCritical, "Send Error Message"
         Me.TBCheckOutItemCode.Text = ""
      End If
   End If
   End If
End If
End Sub

Private Sub TBCOBarCode_Change()
If Me.TBCOBarCode.Text = "" Then
   Me.LBLCOItemCode.Caption = ""
   Me.LBLCOItemName.Caption = ""
   Me.LBLCOItemNetAmount.Caption = ""
   Me.LBLCOPrice.Caption = ""
   Me.LBLCORate1.Caption = ""
   Me.LBLCORate2.Caption = Format(1, "##,##0.00")
   Me.LBLCOShelfCode.Caption = ""
   Me.LBLCOShelfID.Caption = ""
   Me.LBLCOUnitCode.Caption = ""
   Me.LBLCOWHCode.Caption = ""
   Me.LBLCOZoneID.Caption = ""
   Me.TBCOKeyQty.Text = ""
   Me.TBCOBarCode.SetFocus
End If
End Sub

Private Sub TBCOBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBCOBarCode.Text = ""
Me.PICCOAddItem.Visible = False
Me.ListViewItem.SetFocus
End If
End Sub

Private Sub TBCOBarCode_KeyPress(KeyAscii As Integer)
Dim vBarCode As String
Dim vPrice As Double
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vQTY As Double

If KeyAscii = 13 Then
   If Me.TBCOBarCode.Text <> "" Then
      vBarCode = Me.TBCOBarCode.Text
      vQuery = "exec dbo.USP_MB_SearchBarcode '" & vBarCode & "'"
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Me.LBLCOItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
          Me.LBLCOItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
          Me.LBLCOUnitCode.Caption = Trim(vRecordset.Fields("unitcode").Value)
          Me.LBLCOWHCode.Caption = Trim(vRecordset.Fields("defsalewhcode").Value)
          Me.LBLCOShelfCode.Caption = Trim(vRecordset.Fields("defsaleshelf").Value)
          Me.LBLCOShelfID.Caption = Trim(vRecordset.Fields("shelfid").Value)
          Me.LBLCOZoneID.Caption = Trim(vRecordset.Fields("zoneid").Value)
          Me.LBLCOBarCode.Caption = Trim(vRecordset.Fields("barcode").Value)
          Me.LBLCORate1.Caption = Trim(vRecordset.Fields("rate").Value)
          Me.LBLCORate2.Caption = Format(1, "##,##0.00")
          
          vPrice = Trim(vRecordset.Fields("price").Value)
          If vPrice <= 0 Then
             MsgBox "สินค้ารายการนี้ ยังไม่ได้กำหนดราคาขายของหน่วยนับขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
             Me.TBCOKeyQty.Enabled = False
             Exit Sub
          Else
             Me.TBCOKeyQty.Enabled = True
          End If
          
          Me.LBLCOPrice.Caption = Format(vPrice, "##,##0.00")
          
          Me.ListViewCOStock.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             vQTY = Trim(vRecordset.Fields("stock").Value)
             
            Set vListItem = Me.ListViewCOStock.ListItems.Add(, , Trim(vRecordset.Fields("whcode").Value))
            vListItem.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
            vListItem.SubItems(2) = Format(vQTY, "##,##0.00")
            vListItem.SubItems(3) = Trim(vRecordset.Fields("stkunitcode").Value)
             vRecordset.MoveNext
          Next i
             
          Me.TBCOKeyQty.SetFocus
       Else
          Me.TBCOBarCode.SetFocus
       End If
    vRecordset.Close
   End If
End If
End Sub

Private Sub TBCOKeyQty_Change()
Dim vQtyWord As String
Dim vLenQTY As Integer
Dim vQTY As Double
Dim vPrice As Double
Dim vItemNetAmount As Double

If Me.TBCOKeyQty.Text <> "" Then
   vQtyWord = Me.TBCOKeyQty.Text
   CheckNumber (vQtyWord)
      
   If vCheckValueNumber = False Then
      MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      vLenQTY = Len(Me.TBCheckOutItemQty.Text)
      Me.TBCOKeyQty.Text = Left(Me.TBCOKeyQty.Text, vLenQTY - 1)
      Me.TBCOKeyQty.SetFocus
   End If
   
   If Me.TBCOKeyQty.Text <> "" And Me.TBCOKeyQty.Text <> "." Then
      vQTY = Me.TBCOKeyQty.Text
   End If
   
   If Me.LBLCOPrice.Caption <> "" Or Me.LBLCOPrice.Caption <> "." Then
      vPrice = Me.LBLCOPrice.Caption
   End If
   
   vItemNetAmount = vQTY * vPrice
   Me.LBLCOItemNetAmount.Caption = Format(vItemNetAmount, "##,##0.00")
End If

If Me.TBCOKeyQty.Text = "" Then
   Me.LBLCOItemNetAmount.Caption = Format(0, "##,##0.00")
End If
End Sub

Private Sub TBCOKeyQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBCOBarCode.Text = ""
Me.PICCOAddItem.Visible = False
Me.ListViewItem.SetFocus
End If
End Sub

Private Sub TBCOKeyQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CMDPICCOAddItemOK_Click
End If
End Sub

Private Sub TBKeyQty_Change()
Dim vQtyWord As String
Dim vLenQTY As Integer

If Me.TBKeyQty.Text <> "" Then
   vQtyWord = Me.TBKeyQty.Text
   CheckNumber (vQtyWord)
      
   If vCheckValueNumber = False Then
      MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      vLenQTY = Len(Me.TBKeyQty.Text)
      Me.TBKeyQty.Text = Left(Me.TBKeyQty.Text, vLenQTY - 1)
      Me.TBKeyQty.SetFocus
   End If
End If
End Sub

Private Sub TBKeyQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CMDCheckOut_Click
End If
End Sub

Private Sub TBSearchData_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBSearchQueID.SetFocus
Me.PICKeySearchData.Visible = False
End If
End Sub

Private Sub TBSearchData_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Me.TBSearchData.Text <> "" Then
   Me.TBSearchQueID.Text = Me.TBSearchData.Text
   Me.CMDSearchQueID.SetFocus
   Else
   Me.TBSearchQueID.Text = ""
   Me.TBSearchQueID.SetFocus
   End If
   Me.PICKeySearchData.Visible = False
End If
End Sub

Private Sub TBSearchHoldBill_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSearchHoldBill.Visible = False
Me.TBSearchQueID.SetFocus
End If
End Sub

Private Sub TBSearchHoldBill_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem
Dim vSearch As String

Dim vNetDebtAmount As Double

If KeyAscii = 13 Then
vSearch = Me.TBSearchHoldBill.Text
vQuery = "exec dbo.USP_NP_SearchHodingBill '" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   Me.ListViewHoldBill.ListItems.Clear
   vRecordset.MoveFirst
   For i = 1 To vRecordset.RecordCount
   Set vListItem = Me.ListViewHoldBill.ListItems.Add(, , i)
   vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
   vListItem.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
   vListItem.SubItems(3) = Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)
   If Trim(vRecordset.Fields("salecode").Value) <> "" Then
   vListItem.SubItems(4) = Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
   Else
   vListItem.SubItems(4) = ""
   End If
   vNetDebtAmount = Trim(vRecordset.Fields("netdebtamount").Value)
   vListItem.SubItems(5) = Format(vNetDebtAmount, "##,##0.00")
   vListItem.SubItems(6) = Trim(vRecordset.Fields("cashiername").Value)
   vListItem.SubItems(7) = ""
   vRecordset.MoveNext
   Next i
   
    If Me.ListViewHoldBill.ListItems.Count > 0 Then
    Me.ListViewHoldBill.SetFocus
    Else
    Me.TBSearchHoldBill.SetFocus
    End If
      
End If
vRecordset.Close
End If
End Sub

Private Sub TXTSearchAR_KeyPress(KeyAscii As Integer)
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vSearchAR As String
'Dim vListItem As ListItem
''Dim i As Integer

'If KeyAscii = 13 Then
 '  If Me.TXTSearchAR.Text <> "" Then
  '    vSearchAR = Me.TXTSearchAR.Text
   '   vQuery = "exec dbo.USP_AR_ARProFileSearch '" & vSearchAR & "' "
    ''  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      '    Me.ListViewAR.ListItems.Clear
       '   vRecordset.MoveFirst
        '  For i = 1 To vRecordset.RecordCount
         ''    Set vListItem = Me.ListViewAR.ListItems.Add(, , i)
           '  vListItem.SubItems(1) = Trim(vRecordset.Fields("arcode").Value)
            ' vListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
             'vListItem.SubItems(3) = Trim(vRecordset.Fields("memberid").Value)
             'vRecordset.MoveNext
          'Next i
          
          'Me.ListViewAR.SetFocus
      'End If
      'vRecordset.Close
   'End If
'End If
End Sub

Private Sub TBSearchQueID_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 112 Then
    Call CMDSearch_Click
End If

If KeyCode = 113 Then
   Me.PICKeySearchData.Visible = True
   Me.TBSearchData.SetFocus
End If

If KeyCode = 114 Then
    Call CMDSelectItemQue_Click
End If

If KeyCode = 27 Then
    Call CMDClear_Click
End If

If KeyCode = 116 Then
    Call CMDHoldBill_Click
End If

If KeyCode = 115 Then
    Call CMDMerge_Click
End If
End Sub

Private Sub TBSearchQueID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CMDSearchQueID_Click
End If
End Sub

Public Sub PrintCheckOutHeader(vDocNo As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
   
For Each prnPrinter In Printers
   If prnPrinter.DeviceName = "\\hptc-5100418\SRP370CheckOut" Then
      Set Printer = prnPrinter
      Exit For
   End If
Next
        
vQuery = "exec dbo.USP_NP_SearchHodingBillDetails '" & vDocNo & "'  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 20
Printer.CurrentX = 950
Printer.FontBold = True
Printer.Print Trim("CheckOut Master")

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.Print "*" & Trim(vRecordset.Fields("docno").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1550
Printer.FontBold = True
Printer.Print Trim(vRecordset.Fields("docno").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value) & "          " & Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("รหัส/ชื่อลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("arcode").Value) & "*"

If Trim(vRecordset.Fields("mydescription").Value) <> "" Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("ทะเบียนรถ : ") & Trim(vRecordset.Fields("mydescription").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("mydescription").Value) & "*"
End If

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("จุดแคชเชียร์: ") & Trim(vRecordset.Fields("machineno").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value)
    
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 20
Printer.CurrentX = 1000
Printer.FontBold = True
Printer.Print "มูลค่าสุทธิ :" & "       " & Format(Trim(vRecordset.Fields("totalamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 10
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Trim("วันที่พิมพ์ :") & Now
End If
vRecordset.Close

Printer.EndDoc
End Sub


Public Sub PrintCheckOutItem(vDocNo As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
   
For Each prnPrinter In Printers
   If prnPrinter.DeviceName = "\\hptc-5100418\SRP370CheckOut" Then
      Set Printer = prnPrinter
      Exit For
   End If
Next
        
vQuery = "exec dbo.USP_NP_SearchHodingBillDetails '" & vDocNo & "'  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 20
Printer.CurrentX = 950
Printer.FontBold = True
Printer.Print Trim("CheckOut Details")

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.Print "*" & Trim(vRecordset.Fields("docno").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1550
Printer.FontBold = True
Printer.Print Trim(vRecordset.Fields("docno").Value)


Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value) & "          " & Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("รหัส/ชื่อลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("arcode").Value) & "*"

If Trim(vRecordset.Fields("mydescription").Value) <> "" Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("ทะเบียนรถ : ") & Trim(vRecordset.Fields("mydescription").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("mydescription").Value) & "*"
End If

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("จุดแคชเชียร์: ") & Trim(vRecordset.Fields("machineno").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value)
vRecordset.MoveFirst

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
vRecordset.MoveFirst
n = 1
For i = 1 To vRecordset.RecordCount

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value) & "              " & " ขายชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)

Printer.FontName = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.FontBold = False
Printer.Print "*" & Trim(vRecordset.Fields("itemcode").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ชื่อสินค้า :" & Trim(vRecordset.Fields("itemname").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "จำนวน :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("นับได้ : ") & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
    
If i = vRecordset.RecordCount Then

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าสินค้า :" & "     " & Format(Trim(vRecordset.Fields("sumofitemamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าส่วนลด :" & "  " & Format(Trim(vRecordset.Fields("discountamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าภาษี :" & "       " & Format(Trim(vRecordset.Fields("taxamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าสุทธิ :" & "      " & Format(Trim(vRecordset.Fields("totalamount").Value), "##,##0.00")

End If
vRecordset.MoveNext
n = n + 1
Next i
End If
vRecordset.Close

    
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "               ผู้ตรวจสินค้า                                             ผู้รับสินค้า"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "         _____________                                    ______________"
      
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 10
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Trim("วันที่พิมพ์ :") & Now

Printer.EndDoc
End Sub
