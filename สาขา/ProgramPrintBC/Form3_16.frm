VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form3_16 
   Caption         =   "บันทึก ใบเสนอราคา"
   ClientHeight    =   8700
   ClientLeft      =   3045
   ClientTop       =   1290
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   222
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form3_16.frx":0000
   ScaleHeight     =   8700
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictureItemCode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   -90
      ScaleHeight     =   8625
      ScaleWidth      =   12090
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   12120
      Begin VB.ListBox ListUnitCode 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   5220
         TabIndex        =   78
         Top             =   2655
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Frame Frame1 
         Caption         =   "ประเภทราคา"
         Height          =   735
         Left            =   3285
         TabIndex        =   74
         Top             =   270
         Width           =   7575
         Begin VB.CheckBox Check103 
            Caption         =   "ราคาที่1"
            Height          =   330
            Left            =   5265
            TabIndex        =   77
            Top             =   315
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox Check102 
            Caption         =   "ส่งให้"
            Height          =   285
            Left            =   3060
            TabIndex        =   76
            Top             =   315
            Value           =   1  'Checked
            Width           =   1320
         End
         Begin VB.CheckBox Check101 
            Caption         =   "ขายเงินเชื่อ"
            Height          =   375
            Left            =   585
            TabIndex        =   75
            Top             =   270
            Value           =   1  'Checked
            Width           =   1320
         End
      End
      Begin VB.CommandButton CMDExit 
         Caption         =   "ออก"
         Height          =   465
         Left            =   5400
         TabIndex        =   62
         Top             =   5715
         Width           =   1230
      End
      Begin VB.CommandButton CMDSelectItem 
         Caption         =   "ตกลง"
         Height          =   465
         Left            =   3915
         TabIndex        =   61
         Top             =   5715
         Width           =   1230
      End
      Begin VB.TextBox TextDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3285
         TabIndex        =   60
         Top             =   4140
         Width           =   1860
      End
      Begin VB.TextBox TextQTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3285
         TabIndex        =   59
         Top             =   3645
         Width           =   1860
      End
      Begin VB.CommandButton CMDSearchUnitCode 
         Height          =   330
         Left            =   4455
         TabIndex        =   58
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton CMDSearchItem 
         Height          =   330
         Left            =   5220
         TabIndex        =   57
         Top             =   1170
         Width           =   375
      End
      Begin VB.TextBox TextItemCode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3285
         TabIndex        =   56
         Top             =   1170
         Width           =   1905
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "ส่วนลดรวม :"
         Height          =   330
         Left            =   2070
         TabIndex        =   80
         Top             =   4635
         Width           =   1185
      End
      Begin VB.Label LBLItemDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3285
         TabIndex        =   79
         Top             =   4635
         Width           =   1860
      End
      Begin VB.Label LBLTotalAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3285
         TabIndex        =   67
         Top             =   5175
         Width           =   1860
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "มูลค่ารวม :"
         Height          =   330
         Left            =   1845
         TabIndex        =   72
         Top             =   5175
         Width           =   1410
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "ส่วนลดรายตัว :"
         Height          =   285
         Left            =   1935
         TabIndex        =   71
         Top             =   4140
         Width           =   1320
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "จำนวนที่เสนอราคา :"
         Height          =   375
         Left            =   1575
         TabIndex        =   70
         Top             =   3645
         Width           =   1680
      End
      Begin VB.Label LBLOnHand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3285
         TabIndex        =   66
         Top             =   3150
         Width           =   1860
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "OnHand :"
         Height          =   330
         Left            =   2025
         TabIndex        =   69
         Top             =   3150
         Width           =   1185
      End
      Begin VB.Label LBLPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3285
         TabIndex        =   65
         Top             =   2655
         Width           =   1860
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "ราคาขาย :"
         Height          =   285
         Left            =   2115
         TabIndex        =   68
         Top             =   2655
         Width           =   1140
      End
      Begin VB.Label LBLUnitCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3285
         TabIndex        =   64
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "หน่วยนับ :"
         Height          =   240
         Left            =   2070
         TabIndex        =   55
         Top             =   2205
         Width           =   1185
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "ชื่อสินค้า :"
         Height          =   285
         Left            =   2250
         TabIndex        =   54
         Top             =   1665
         Width           =   1005
      End
      Begin VB.Label LBLItemName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3285
         TabIndex        =   63
         Top             =   1665
         Width           =   7575
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "รหัสสินค้า :"
         Height          =   375
         Left            =   2340
         TabIndex        =   53
         Top             =   1170
         Width           =   915
      End
   End
   Begin VB.CommandButton CMDKeyItem 
      Caption         =   "กรอกข้อมูลขายสินค้า"
      Height          =   330
      Left            =   315
      TabIndex        =   73
      Top             =   2295
      Width           =   1770
   End
   Begin MSComctlLib.ListView ListView104 
      Height          =   1590
      Left            =   1710
      TabIndex        =   40
      Top             =   5940
      Visible         =   0   'False
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   2805
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัส"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อพนักงาน"
         Object.Width           =   3704
      EndProperty
   End
   Begin VB.PictureBox Pic101 
      Height          =   510
      Left            =   10215
      ScaleHeight     =   450
      ScaleWidth      =   180
      TabIndex        =   14
      Top             =   7965
      Visible         =   0   'False
      Width           =   240
      Begin VB.CommandButton CMD202 
         Caption         =   "ออก"
         Height          =   420
         Left            =   8550
         TabIndex        =   37
         Top             =   4725
         Width           =   1140
      End
      Begin VB.CommandButton CMD201 
         Caption         =   "เลือก"
         Height          =   420
         Left            =   7245
         TabIndex        =   36
         Top             =   4725
         Width           =   1140
      End
      Begin VB.TextBox Text201 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1530
         TabIndex        =   24
         Top             =   135
         Width           =   2760
      End
      Begin MSComctlLib.ListView ListView102 
         Height          =   3930
         Left            =   495
         TabIndex        =   25
         Top             =   675
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   6932
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสลูกค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   13229
         EndProperty
      End
      Begin VB.Label Label15 
         Caption         =   "ค้นหาลูกค้า :"
         Height          =   330
         Left            =   540
         TabIndex        =   23
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.PictureBox Pic102 
      Height          =   600
      Left            =   10620
      ScaleHeight     =   540
      ScaleWidth      =   405
      TabIndex        =   41
      Top             =   7965
      Visible         =   0   'False
      Width           =   465
      Begin VB.CommandButton CMD110 
         Height          =   330
         Left            =   4905
         Picture         =   "Form3_16.frx":72FB
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   315
         Width           =   375
      End
      Begin VB.CommandButton CMD113 
         Caption         =   "ออก"
         Height          =   420
         Left            =   8550
         TabIndex        =   48
         Top             =   4410
         Width           =   780
      End
      Begin VB.CommandButton CMD112 
         Caption         =   "เลือก"
         Height          =   420
         Left            =   7605
         TabIndex        =   47
         Top             =   4410
         Width           =   780
      End
      Begin MSComctlLib.ListView ListView105 
         Height          =   3165
         Left            =   2745
         TabIndex        =   46
         Top             =   945
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   5583
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "เลขที่"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "วันที่"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   5115
         EndProperty
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2790
         TabIndex        =   45
         Top             =   315
         Width           =   2040
      End
      Begin VB.Label Label21 
         Caption         =   "คำค้นหา ใบเสนอราคา :"
         Height          =   330
         Left            =   945
         TabIndex        =   44
         Top             =   315
         Width           =   1770
      End
   End
   Begin VB.CommandButton CMD111 
      Height          =   420
      Left            =   315
      Picture         =   "Form3_16.frx":763F
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   7650
      Width           =   1185
   End
   Begin VB.TextBox Text115 
      Appearance      =   0  'Flat
      Height          =   1140
      Left            =   1710
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Top             =   6390
      Width           =   5730
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   90
      Top             =   6255
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.TextBox Text109 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   330
      Left            =   1710
      TabIndex        =   27
      Top             =   5580
      Width           =   2040
   End
   Begin VB.TextBox Text113 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   330
      Left            =   9675
      TabIndex        =   38
      Top             =   6795
      Width           =   1995
   End
   Begin VB.CommandButton CMD108 
      Height          =   420
      Left            =   5130
      Picture         =   "Form3_16.frx":7AAA
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7650
      Width           =   1230
   End
   Begin VB.ComboBox CMB102 
      Height          =   315
      Left            =   1710
      TabIndex        =   34
      Text            =   "รับเอง"
      Top             =   5985
      Width           =   2040
   End
   Begin VB.CommandButton CMD104 
      Height          =   330
      Left            =   3780
      Picture         =   "Form3_16.frx":9FC2
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5580
      Width           =   375
   End
   Begin VB.TextBox Text110 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   330
      Left            =   9675
      TabIndex        =   31
      Text            =   "0"
      Top             =   5580
      Width           =   1995
   End
   Begin MSComCtl2.DTPicker DTPicker102 
      Height          =   315
      Left            =   5535
      TabIndex        =   30
      Top             =   5580
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57737217
      CurrentDate     =   38733
   End
   Begin VB.CommandButton CMD107 
      Height          =   420
      Left            =   3915
      Picture         =   "Form3_16.frx":A38F
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7650
      Width           =   1230
   End
   Begin VB.CommandButton CMD102 
      Height          =   330
      Left            =   8010
      Picture         =   "Form3_16.frx":A81D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1305
      Width           =   375
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   4545
      Picture         =   "Form3_16.frx":ABEA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1305
      Width           =   420
   End
   Begin VB.CommandButton CMD106 
      Height          =   420
      Left            =   2700
      Picture         =   "Form3_16.frx":AF2E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7650
      Width           =   1230
   End
   Begin VB.CommandButton CMD105 
      Height          =   420
      Left            =   1485
      Picture         =   "Form3_16.frx":B355
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7650
      Width           =   1230
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2760
      Left            =   315
      TabIndex        =   7
      Top             =   2700
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   4868
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483634
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสสินค้า"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อสินค้า"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "จำนวน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "หน่วย"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "ราคาต่อหน่วย"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "ส่วนลดรายตัว"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "ส่วนลดรวม"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "ราคารวม"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.TextBox Text103 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   330
      Left            =   2295
      TabIndex        =   6
      Top             =   1755
      Width           =   9375
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   330
      Left            =   6300
      TabIndex        =   4
      Top             =   1305
      Width           =   1680
   End
   Begin VB.TextBox Text111 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Height          =   330
      Left            =   9675
      TabIndex        =   8
      Top             =   5985
      Width           =   1995
   End
   Begin VB.TextBox Text112 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Height          =   330
      Left            =   9675
      TabIndex        =   9
      Top             =   6390
      Width           =   1995
   End
   Begin VB.TextBox Text114 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Height          =   330
      Left            =   9675
      TabIndex        =   10
      Top             =   7200
      Width           =   1995
   End
   Begin VB.ComboBox CMB101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2295
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   855
      Width           =   2220
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   315
      Left            =   6300
      TabIndex        =   2
      Top             =   855
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   57737217
      CurrentDate     =   38733
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   330
      Left            =   2295
      TabIndex        =   0
      Top             =   1305
      Width           =   2220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   12015
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สถานะ :"
      ForeColor       =   &H80000014&
      Height          =   240
      Left            =   3195
      TabIndex        =   51
      Top             =   495
      Width           =   735
   End
   Begin VB.Image Image102 
      Height          =   300
      Left            =   3960
      Picture         =   "Form3_16.frx":B7EA
      Top             =   450
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Image101 
      Height          =   300
      Left            =   3960
      Picture         =   "Form3_16.frx":BD26
      Top             =   450
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ :"
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   540
      TabIndex        =   43
      Top             =   6390
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ส่วนลดท้ายบิล (บาท) :"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   7560
      TabIndex        =   39
      Top             =   6795
      Width           =   2040
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วิธีการจัดส่ง :"
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   495
      TabIndex        =   33
      Top             =   5985
      Width           =   1140
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เครดิต(วัน) :"
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   8460
      TabIndex        =   29
      Top             =   5580
      Width           =   1140
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ยืนราคา :"
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   4500
      TabIndex        =   28
      Top             =   5580
      Width           =   960
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสพนักงาน :"
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   225
      TabIndex        =   26
      Top             =   5580
      Width           =   1410
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่ารวม :"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   8010
      TabIndex        =   22
      Top             =   7200
      Width           =   1590
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่าภาษี :"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   7920
      TabIndex        =   21
      Top             =   6390
      Width           =   1680
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่าสินค้า :"
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   8055
      TabIndex        =   20
      Top             =   5985
      Width           =   1545
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   12015
      Y1              =   5490
      Y2              =   5490
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า :"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   765
      TabIndex        =   19
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกค้า :"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   5040
      TabIndex        =   18
      Top             =   1350
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทเอกสาร :"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   900
      TabIndex        =   17
      Top             =   900
      Width           =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เอกสาร :"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   4770
      TabIndex        =   16
      Top             =   900
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   1035
      TabIndex        =   15
      Top             =   1350
      Width           =   1185
   End
   Begin VB.Menu Q1 
      Caption         =   "Quotation"
   End
   Begin VB.Menu Menu1 
      Caption         =   ""
      Begin VB.Menu MenuEdit 
         Caption         =   "แก้ไขรายการ"
      End
   End
End
Attribute VB_Name = "Form3_16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vIndex As Integer
Dim vEdit As Integer

'Private Sub Check101_Click()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vItemCode As String
'Dim vUnitCode As String
'Dim vSaleType As Integer
'Dim vTranSportType As Integer
'Dim vPriceLevel As Integer
'
'Dim vQTY As Double
'Dim vAmount As Double
'Dim vDiscount As Double
'Dim vPrice As Double
''Dim vDiscount1 As Double

'If Me.Check101.Value = 1 Then
 ' vSaleType = 1 'ขายเงินเชื่อ
'Else
 ' vSaleType = 0 'ขายเงินสด
'End If

'If Me.Check102.Value = 1 Then
 ' vTranSportType = 1 'ส่งให้
'Else
 ' vTranSportType = 0 'รับเอง
'End If

'If Me.Check103.Value = 1 Then
 ' vPriceLevel = 1 'ราคา1
'Else
 ' vPriceLevel = 2 'ราคา2
'End If
'vUnitCode = Trim(Me.LBLUnitCode.Caption)
'Me.LBLUnitCode.Caption = vUnitCode
'vItemCode = Trim(Me.TextItemCode.Text)
'vQuery = "exec dbo.USP_NP_QuotationSelectPriceList '" & vItemCode & "','" & vUnitCode & "'," & vSaleType & "," & vTranSportType & "," & vPriceLevel & " "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   Me.LBLPrice.Caption = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
'End If
'vRecordset.Close
'Me.ListUnitCode.Visible = False

'If Me.TextQTY.Text <> "" And Me.LBLPrice.Caption <> "" Then

 ' vPrice = Me.LBLPrice.Caption
  'vQTY = Me.TextQTY.Text
  
  'If Me.TextDiscount.Text <> "%" Then
   ' If InStr(1, Me.TextDiscount.Text, "%") > 0 Then
    '  vDiscount1 = Left(Me.TextDiscount.Text, InStr(1, Me.TextDiscount.Text, "%") - 1)
     ' vDiscount = vQTY * ((vPrice * vDiscount1) / 100)
    'Else
    'If Me.TextDiscount.Text <> "" Then
     ' vDiscount1 = Me.TextDiscount.Text
      'Else
      'vDiscount1 = 0
      'End If
      'vDiscount = vDiscount1 * vQTY
    'End If
    'Me.LBLItemDiscount.Caption = Format(vDiscount, "##,##0.00")
    'vAmount = (vQTY * vPrice) - vDiscount
    'Me.LBLTotalAmount.Caption = Format(vAmount, "##,##0.00")
    'Else
    'Me.LBLTotalAmount.Caption = Format(vPrice * vQTY, "##,##0.00")
  'End If

'End If

'Me.TextQTY.SetFocus
'End Sub

'Private Sub Check102_Click()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vItemCode As String
'Dim vUnitCode As String
'Dim vSaleType As Integer
'Dim vTranSportType As Integer
'Dim vPriceLevel As Integer
'

'Dim vQTY As Double
'Dim vAmount As Double
'Dim vDiscount As Double
'Dim vPrice As Double
'Dim vDiscount1 As Double
'

'If Me.Check101.Value = 1 Then
 ' vSaleType = 1 'ขายเงินเชื่อ
'Else
 ' vSaleType = 0 'ขายเงินสด
'End If
'
'If Me.Check102.Value = 1 Then
 ' vTranSportType = 1 'ส่งให้
'Else
 ' vTranSportType = 0 'รับเอง
'End If

'If Me.Check103.Value = 1 Then
 ' vPriceLevel = 1 'ราคา1
'Else
 ' vPriceLevel = 2 'ราคา2
'End If
'vUnitCode = Trim(Me.LBLUnitCode.Caption)
'Me.LBLUnitCode.Caption = vUnitCode
'vItemCode = Trim(Me.TextItemCode.Text)
'vQuery = "exec dbo.USP_NP_QuotationSelectPriceList '" & vItemCode & "','" & vUnitCode & "'," & vSaleType & "," & vTranSportType & "," & vPriceLevel & " "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   Me.LBLPrice.Caption = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
'End If
'vRecordset.Close
'Me.ListUnitCode.Visible = False

'If Me.TextQTY.Text <> "" And Me.LBLPrice.Caption <> "" Then

 ' vPrice = Me.LBLPrice.Caption
  'vQTY = Me.TextQTY.Text
  
  'If Me.TextDiscount.Text <> "%" Then
   ' If InStr(1, Me.TextDiscount.Text, "%") > 0 Then
    '  vDiscount1 = Left(Me.TextDiscount.Text, InStr(1, Me.TextDiscount.Text, "%") - 1)
     ' vDiscount = vQTY * ((vPrice * vDiscount1) / 100)
    'Else
    'If Me.TextDiscount.Text <> "" Then
     ' vDiscount1 = Me.TextDiscount.Text
      'Else
      'vDiscount1 = 0
      'End If
      'vDiscount = vDiscount1 * vQTY
    'End If
    'Me.LBLItemDiscount.Caption = Format(vDiscount, "##,##0.00")
    'vAmount = (vQTY * vPrice) - vDiscount
    'Me.LBLTotalAmount.Caption = Format(vAmount, "##,##0.00")
    'Else
    'Me.LBLTotalAmount.Caption = Format(vPrice * vQTY, "##,##0.00")
  'End If

'End If


'Me.TextQTY.SetFocus
'End Sub

Private Sub Check103_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vItemCode As String
Dim vUnitCode As String
Dim vSaleType As Integer
Dim vTranSportType As Integer
Dim vPriceLevel As Integer

Dim vQTY As Double
Dim vAmount As Double
Dim vDiscount As Double
Dim vPrice As Double
Dim vDiscount1 As Double


If Me.Check101.Value = 1 Then
  vSaleType = 1 'ขายเงินเชื่อ
Else
  vSaleType = 0 'ขายเงินสด
End If

If Me.Check102.Value = 1 Then
  vTranSportType = 1 'ส่งให้
Else
  vTranSportType = 0 'รับเอง
End If

If Me.Check103.Value = 1 Then
  vPriceLevel = 1 'ราคา1
Else
  vPriceLevel = 2 'ราคา2
End If
vUnitCode = Trim(Me.LBLUnitCode.Caption)
Me.LBLUnitCode.Caption = vUnitCode
vItemCode = Trim(Me.TextItemCode.Text)
vQuery = "exec dbo.USP_NP_QuotationSelectPriceList '" & vItemCode & "','" & vUnitCode & "'," & vSaleType & "," & vTranSportType & "," & vPriceLevel & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.LBLPrice.Caption = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
End If
vRecordset.Close
Me.ListUnitCode.Visible = False

If Me.TextQTY.Text <> "" And Me.LBLPrice.Caption <> "" Then

  vPrice = Me.LBLPrice.Caption
  vQTY = Me.TextQTY.Text
  
  If Me.TextDiscount.Text <> "%" Then
    If InStr(1, Me.TextDiscount.Text, "%") > 0 Then
      vDiscount1 = Left(Me.TextDiscount.Text, InStr(1, Me.TextDiscount.Text, "%") - 1)
      vDiscount = vQTY * ((vPrice * vDiscount1) / 100)
    Else
    If Me.TextDiscount.Text <> "" Then
      vDiscount1 = Me.TextDiscount.Text
      Else
      vDiscount1 = 0
      End If
      vDiscount = vDiscount1 * vQTY
    End If
    Me.LBLItemDiscount.Caption = Format(vDiscount, "##,##0.00")
    vAmount = (vQTY * vPrice) - vDiscount
    Me.LBLTotalAmount.Caption = Format(vAmount, "##,##0.00")
    Else
    Me.LBLTotalAmount.Caption = Format(vPrice * vQTY, "##,##0.00")
  End If

End If


Me.TextQTY.SetFocus
End Sub

Private Sub CMB101_Click()
If vIsOpenQuotation = 0 Then
Text101.Text = ""
End If
End Sub

Private Sub CMD101_Click()
Dim vRecordset  As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vTypeDocno As Integer

On Error GoTo ErrDescription

vIsOpenQuotation = 0
If vIsOpenQuotation = 0 Then

    Select Case CMB101.ListIndex
    Case 0
        vTypeDocno = 0
    Case 1
        vTypeDocno = 1
    Case -1
        MsgBox "กรุณาเลือกประเภทเอกสาร", vbCritical, "Send Error"
        Exit Sub
    End Select
    vQuery = "exec dbo.USP_NP_QuotationNewDocNo " & vTypeDocno & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vDocNo = Trim(vRecordset.Fields("newdocno").Value)
    End If
    vRecordset.Close
    Text101.Text = vDocNo
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
PIC101.Visible = True
End Sub



Private Sub CMD103_Click()

End Sub

Private Sub CMD104_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListSaleCode As ListItem

On Error GoTo ErrDescription

ListView104.ListItems.Clear
ListView104.Visible = True
vQuery = "select salecode,salename from npmaster.dbo.bcsalegroup where activestatus = 1 order by salename"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set vListSaleCode = ListView104.ListItems.Add(, , Trim(vRecordset.Fields("salecode").Value))
    vListSaleCode.SubItems(1) = Trim(vRecordset.Fields("salename").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMD105_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vDocdate As Date
Dim vBillType As Integer
Dim vARCode As String
Dim vCreditDay As Integer
Dim vValidate As Date
Dim vIsConditionSend As Integer
Dim vSaleCode As String
Dim vTaxRate As Integer
Dim vIsCancel As Integer
Dim vSumOfItemAmount As Currency
Dim vTaxAmount As Currency
Dim vDiscountAmount As Currency
Dim vTotalAmount As Currency
Dim vNetAmount As Currency
Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Integer
Dim vPrice As Currency
Dim vUnitCode As String
Dim vDisCountAmountSub As Currency
Dim vSumDiscountAmount As Currency
Dim vAmount As Currency
Dim vIsCancelSub As Integer
Dim vLineNumber As Integer
Dim i As Integer
Dim vReturnStatus As Integer
Dim vAnswerPrint As Integer
Dim vMydescription As String
Dim vReportName As String
Dim vCheckDocnoExist As Integer
Dim vRepID As Integer
Dim vRepType As String


If Text101.Text <> "" And Text102.Text <> "" And ListView101.ListItems.Count <> 0 And Text110.Text <> "" And Text109.Text <> "" Then
    If CCur(Text113.Text) < CCur(Text111.Text) Then
            vDocNo = Trim(Text101.Text)
            
            vQuery = "select count(docno) as doccount from npmaster.dbo.tb_np_quotation where docno = '" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckDocnoExist = Trim(vRecordset.Fields("doccount").Value)
            End If
            vRecordset.Close
            
            If vCheckDocnoExist > 0 Then
                Call CMD101_Click
                MsgBox "เอกสารนี้มีอยู่แล้ว กดบันทึกข้อมูลอีกครั้ง", vbCritical, "Message Error"
                Exit Sub
            End If
            vDocdate = Trim(DTPicker101.Day) & "/" & Trim(DTPicker101.Month) & "/" & DTPicker101.Year
            Select Case CMB101.ListIndex
            Case 0
                vBillType = 0
            Case 1
                vBillType = 1
            End Select
            vARCode = Trim(Text102.Text)
            vCreditDay = Trim(Text110.Text)
            vValidate = Trim(DTPicker102.Day) & "/" & Trim(DTPicker102.Month) & "/" & DTPicker102.Year
            Select Case CMB102.ListIndex
            Case 0
                vIsConditionSend = 0
            Case 1
                vIsConditionSend = 1
            End Select
            vSaleCode = Trim(Text109.Text)
            vTaxRate = 7
            vIsCancel = 0
            vSumOfItemAmount = Format(CCur(Trim(Text111.Text)), "##,##0.00")
            vTaxAmount = Format(CCur(Trim(Text112.Text)), "##,##0.00")
            vDiscountAmount = Format(CCur(Trim(Text113.Text)), "##,##0.00")
            vTotalAmount = Format(CCur(Trim(Text114.Text)), "##,##0.00")
            vNetAmount = Format(CCur(Trim(Text114.Text)), "##,##0.00")
            vMydescription = Trim(Text115.Text)
            
            On Error GoTo ErrHeader
            
            vQuery = "exec dbo.USP_NP_QuotaionInsertHeader " & vIsOpenQuotation & ",'" & vDocNo & "','" & vDocdate & "', " _
            & " " & vBillType & ",'" & vARCode & "'," & vCreditDay & ",'" & vValidate & "'," & vIsConditionSend & ", " _
            & " '" & vSaleCode & "'," & vTaxRate & "," & vIsCancel & "," & vSumOfItemAmount & ", " & vTaxAmount & ", " _
            & " " & vDiscountAmount & "," & vTotalAmount & "," & vNetAmount & ",'" & vMydescription & "','" & vUserID & "' "
            gConnection.Execute vQuery
            vReturnStatus = 0
ErrHeader:
            If Err.Description <> "" Then
            vReturnStatus = 1
            MsgBox Err.Description
            End If

            For i = 1 To ListView101.ListItems.Count
                vItemCode = Trim(ListView101.ListItems.Item(i).Text)
                vItemName = Trim(ListView101.ListItems.Item(i).SubItems(1))
                vQTY = Trim(ListView101.ListItems.Item(i).SubItems(2))
                vPrice = Trim(ListView101.ListItems.Item(i).SubItems(4))
                vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(3))
                vDisCountAmountSub = Trim(ListView101.ListItems.Item(i).SubItems(5))
                vSumDiscountAmount = CCur(Trim(ListView101.ListItems.Item(i).SubItems(2))) * CCur(Trim(ListView101.ListItems.Item(i).SubItems(5)))
                vAmount = Trim(ListView101.ListItems.Item(i).SubItems(6))
                vIsCancelSub = 0
                vLineNumber = Trim(ListView101.ListItems.Item(i).Text) - 1
                
                vQuery = "exec dbo.USP_NP_QuotaionInsertDetails " & vReturnStatus & ",'" & vDocNo & "','" & vDocdate & "', " _
                & " '" & vItemCode & "','" & vItemName & "'," & vQTY & "," & vPrice & ",'" & vUnitCode & "'," & vDisCountAmountSub & ", " _
                & " " & vSumDiscountAmount & "," & vAmount & "," & vIsCancelSub & "," & vLineNumber & " "
                gConnection.Execute vQuery
            Next i
            If vReturnStatus = 0 Then
            vAnswerPrint = MsgBox("บันทึกใบเสนอราคาเลขที่ " & vDocNo & " ต้องการพิมพ์เอกสารเลยหรือไม่", vbYesNo, "Send Question")
            If vAnswerPrint = 6 Then
            
            vRepID = 305
            vRepType = "QT"
            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                'vQuery = "select reportname from bcnp.dbo.bcreportname where reptype = 'QT' and repid = 305"
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vReportName = Trim(vRecordset.Fields("reportname").Value)
                End If
                vRecordset.Close
                
                With Crystal101
                .ReportFileName = Trim(vReportName & ".rpt")
                .ParameterFields(0) = "@Docno;" & vDocNo & ";true"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .Action = 1
                End With
            End If
            Text101.Text = ""
            Text102.Text = ""
            Text103.Text = ""
            Text109.Text = ""
            Text110.Text = ""
            Text111.Text = ""
            Text112.Text = ""
            Text113.Text = ""
            Text114.Text = ""
            Text115.Text = ""
            DTPicker101.Value = Now
            DTPicker102.Value = Now
            ListView101.ListItems.Clear
            vIsOpenQuotation = 0
            End If
        Else
            MsgBox "ส่วนลดไม่สามารถลดได้มากกว่ามูลค่าสินค้า", vbCritical, "Send Error"
            Exit Sub
        End If
Else
    MsgBox "กรุณากรอกข้อมูลของเอกสารให้เรียบร้อย", vbInformation, "Send Information"
    Exit Sub
    CMD101.SetFocus
End If
End Sub

Private Sub CMD106_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vAnswer As Integer
Dim vCount As Integer

On Error GoTo ErrDescription

If vIsOpenQuotation = 1 Then
    vDocNo = Trim(Text101.Text)
    vAnswer = MsgBox("ต้องการยกเลิกใบเสนอราคาเลขที่ " & vDocNo & " ใช่หรือไม่", vbYesNo, "Send Question")
    If vAnswer = 6 Then
    vQuery = "exec dbo.USP_NP_UpdateIsCancelQuotation '" & vDocNo & "' "
    gConnection.Execute vQuery
    MsgBox "ยกเลิกเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", vbCritical, "Send Error"
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""

    Text109.Text = ""
    Text110.Text = ""
    Text111.Text = ""
    Text112.Text = ""
    Text113.Text = ""
    Text114.Text = ""
    Text115.Text = ""
    DTPicker101.Value = Now
    DTPicker102.Value = Now
    ListView101.ListItems.Clear
    vIsOpenQuotation = 0
    Image101.Visible = True
    Image102.Visible = False
    Else
        Exit Sub
    End If
Else
    MsgBox "เอกสารยังไม่ได้บันทึก ไม่สามารถยกเลิกเอกสารได้", vbCritical, "Send Error"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMD107_Click()
Pic102.Visible = True
End Sub

Private Sub CMD108_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vReportName As String
Dim vRepType As String

On Error GoTo ErrDescription

If vIsOpenQuotation = 1 Then
    vRepID = 305
    vRepType = "QT"
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = " & vRepID & " and reptype = 'QT' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
    
    vDocNo = Trim(Text101.Text)
    With Crystal101
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@Docno;" & vDocNo & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    End With
                
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""

    Text109.Text = ""
    Text110.Text = ""
    Text111.Text = ""
    Text112.Text = ""
    Text113.Text = ""
    Text114.Text = ""
    DTPicker101.Value = Now
    DTPicker102.Value = Now
    ListView101.ListItems.Clear
    vIsOpenQuotation = 0
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub



Private Sub CMD110_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocsearch As String
Dim vListDocSearch As ListItem
Dim i As Integer
Dim vTypeSearch As Integer

On Error GoTo ErrDescription

    If Text1.Text = "" Then
        vTypeSearch = 1
    Else
        vTypeSearch = 0
    End If
    ListView105.ListItems.Clear
    i = 1
    vDocsearch = Trim(Text1.Text)
    vQuery = "exec dbo.USP_NP_SearchQuotation " & vTypeSearch & ",'" & vDocsearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set vListDocSearch = ListView105.ListItems.Add(, , i)
            vListDocSearch.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
            vListDocSearch.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
            vListDocSearch.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
    
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMD111_Click()

On Error GoTo ErrDescription

Text101.Text = ""
Text102.Text = ""
Text103.Text = ""


Text109.Text = ""
Text110.Text = ""
Text111.Text = ""
Text112.Text = ""
Text113.Text = ""
Text114.Text = ""
DTPicker101.Value = Now
DTPicker102.Value = Now
ListView101.ListItems.Clear
vIsOpenQuotation = 0
Image101.Visible = True
Image102.Visible = False
CMB101.Enabled = True
CMD105.Enabled = True
CMD106.Enabled = True
CMD107.Enabled = True
CMD108.Enabled = True
vEdit = 0

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub



Private Sub CMD113_Click()
Pic102.Visible = False
End Sub

Private Sub CMD202_Click()
PIC101.Visible = False
Text102.SetFocus
End Sub

Private Sub CMDExit_Click()
Me.PictureItemCode.Visible = False
End Sub

Private Sub CMDKeyItem_Click()
Me.PictureItemCode.Visible = True
Me.TextItemCode.SetFocus
End Sub

Private Sub CMDSearchUnitCode_Click()
Me.ListUnitCode.Visible = True
End Sub

Private Sub CMDSelectItem_Click()
Dim i As Integer
Dim j As Integer
Dim vListInsert As ListItem
Dim vSumOfItemAmount As Currency
Dim vTaxAmount As Currency
Dim vTotalAmount As Currency
Dim vNetAmount As Currency
Dim vCalAmount As Currency

On Error GoTo ErrDescription

If vEdit = 0 Then
    If Me.TextItemCode.Text <> "" And Me.TextQTY.Text <> "" And Me.LBLTotalAmount.Caption <> "" And Me.LBLOnHand.Caption <> "" Then
        If ListView101.ListItems.Count <> 0 Then
            i = ListView101.ListItems.Count + 1
            If Me.TextDiscount.Text = "" Then
                Me.LBLItemDiscount.Caption = Format(0, "##,##0.00")
            End If
            Text111.Text = Format(CCur(Text111.Text) + (CCur(Me.TextQTY.Text) * CCur(Me.LBLPrice.Caption)) - CCur(Me.LBLItemDiscount.Caption), "##,##0.00")
            Text112.Text = Format(((CCur(Text111.Text) - CCur(Text113.Text)) * 7) / 100, "##,##0.00")
            Text114.Text = Format(CCur(Text111.Text) - CCur(Text113.Text), "##,##0.00")
        Else
            i = 1
            If Me.LBLItemDiscount.Caption = "" Then
                Me.TextDiscount.Text = Format("0", "##,##0.00")
            End If
            Text111.Text = Format((CCur(Me.TextQTY.Text) * CCur(Me.LBLPrice.Caption)) - CCur(Me.LBLItemDiscount.Caption), "##,##0.00")
            Text112.Text = Format((CCur(Text111.Text) * 7) / 100, "##,##0.00")
            Text114.Text = Format(Text111.Text, "##,##0.00")
            
            
            
            
        End If
            Set vListInsert = ListView101.ListItems.Add(, , Trim(Me.TextItemCode.Text))
            vListInsert.SubItems(1) = (Trim(Me.LBLItemName.Caption))
            vListInsert.SubItems(2) = Format(Trim(Me.TextQTY.Text), "##,##0.00")
            vListInsert.SubItems(3) = Trim(Me.LBLUnitCode.Caption)
            vListInsert.SubItems(4) = Format(Trim(Me.LBLPrice.Caption), "##,##0.00")
            vListInsert.SubItems(5) = Trim(Me.TextDiscount.Text)
            vListInsert.SubItems(6) = Format(Trim(Me.LBLItemDiscount.Caption), "##,##0.00")
            vListInsert.SubItems(7) = Format(Me.LBLTotalAmount.Caption, "##,##0.00")
            Me.TextItemCode.Text = ""
            Me.LBLItemName.Caption = ""
            Me.LBLUnitCode.Caption = ""
            Me.LBLPrice.Caption = ""
            Me.TextDiscount.Text = ""
            Me.LBLTotalAmount.Caption = ""
            Me.TextQTY.Text = ""
            Me.LBLOnHand.Caption = ""
            Me.LBLItemDiscount.Caption = ""
            Me.PictureItemCode.Visible = False
    Else
        MsgBox "กรุณากรอกข้อมูลให้ครบ ตามช่องที่มีตัวหนังสือสีแดง", vbCritical, "Send Error"
        Exit Sub
    End If
Else
        If Me.TextDiscount.Text = "" Then
            Me.TextDiscount.Text = Format(CCur(0), "##,#00.00")
        End If
        vListInsert.SubItems(1) = (Trim(Me.LBLItemName.Caption))
        vListInsert.SubItems(2) = Format(Trim(Me.TextQTY.Text), "##,##0.00")
        vListInsert.SubItems(3) = Trim(Me.LBLUnitCode.Caption)
        vListInsert.SubItems(4) = Format(Trim(Me.LBLPrice.Caption), "##,##0.00")
        vListInsert.SubItems(5) = Trim(Me.TextDiscount.Text)
        vListInsert.SubItems(6) = Format(Trim(Me.LBLItemDiscount.Caption), "##,##0.00")
        vListInsert.SubItems(7) = Format(Me.LBLTotalAmount.Caption, "##,##0.00")
        
        Me.TextItemCode.Text = ""
        Me.LBLItemName.Caption = ""
        Me.LBLUnitCode.Caption = ""
        Me.LBLPrice.Caption = ""
        Me.TextDiscount.Text = ""
        Me.LBLTotalAmount.Caption = ""
        Me.TextQTY.Text = ""
        Me.LBLOnHand.Caption = ""
        Me.LBLItemDiscount.Caption = ""
        Me.PictureItemCode.Visible = False
        
        
        Dim vCount As Integer
        Dim vSumAmount As Currency
        For vCount = 1 To ListView101.ListItems.Count
            vSumAmount = CCur(vSumAmount) + CCur(ListView101.ListItems.Item(vCount).SubItems(7))
        Next vCount
        Text111.Text = Format(vSumAmount, "##,##0.00")
        Text112.Text = Format(((vSumAmount - CCur(Text113.Text)) * 7) / 100, "##,##0.00")
        Text114.Text = Format((vSumAmount - CCur(Text113.Text)), "##,##0.00")
        vEdit = 0
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_Load()

On Error Resume Next

DTPicker101 = Now
DTPicker102 = Now
Image101.Visible = True
Image102.Visible = False
vIsOpenQuotation = 0
CMB101.AddItem Trim("ขายสินค้าเงินสด")
CMB101.AddItem Trim("ขายสินค้าเงินเชื่อ")
CMB102.AddItem Trim("รับเอง")
CMB102.AddItem Trim("ส่งให้")
Text111.Text = "0.00"
Text112.Text = "0.00"
Text113.Text = "0.00"
Text114.Text = "0.00"
End Sub





Private Sub LBLUnitCode_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vItemCode As String
Dim vUnitCode As String
Dim vSaleType As Integer
Dim vTranSportType As Integer
Dim vPriceLevel As Integer


If Me.Check101.Value = 1 Then
  vSaleType = 1 'ขายเงินเชื่อ
Else
  vSaleType = 0 'ขายเงินสด
End If

If Me.Check102.Value = 1 Then
  vTranSportType = 1 'ส่งให้
Else
  vTranSportType = 0 'รับเอง
End If

If Me.Check103.Value = 1 Then
  vPriceLevel = 1 'ราคา1
Else
  vPriceLevel = 2 'ราคา2
End If
vUnitCode = Trim(Me.LBLUnitCode.Caption)
Me.LBLUnitCode.Caption = vUnitCode
vItemCode = Trim(Me.TextItemCode.Text)
vQuery = "exec dbo.USP_NP_QuotationSelectPriceList '" & vItemCode & "','" & vUnitCode & "'," & vSaleType & "," & vTranSportType & "," & vPriceLevel & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.LBLPrice.Caption = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
End If
vRecordset.Close
Me.ListUnitCode.Visible = False
Me.TextQTY.SetFocus
End Sub

Private Sub ListUnitCode_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vItemCode As String
Dim vUnitCode As String
Dim vSaleType As Integer
Dim vTranSportType As Integer
Dim vPriceLevel As Integer

Dim vQTY As Double
Dim vAmount As Double
Dim vDiscount As Double
Dim vPrice As Double
Dim vDiscount1 As Double


If Me.Check101.Value = 1 Then
  vSaleType = 1 'ขายเงินเชื่อ
Else
  vSaleType = 0 'ขายเงินสด
End If

If Me.Check102.Value = 1 Then
  vTranSportType = 1 'ส่งให้
Else
  vTranSportType = 0 'รับเอง
End If

If Me.Check103.Value = 1 Then
  vPriceLevel = 1 'ราคา1
Else
  vPriceLevel = 2 'ราคา2
End If
vUnitCode = Trim(Me.ListUnitCode.Text)
Me.LBLUnitCode.Caption = vUnitCode
vItemCode = Trim(Me.TextItemCode.Text)
vQuery = "exec dbo.USP_NP_QuotationSelectPriceList '" & vItemCode & "','" & vUnitCode & "'," & vSaleType & "," & vTranSportType & "," & vPriceLevel & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.LBLPrice.Caption = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
End If
vRecordset.Close
Me.ListUnitCode.Visible = False

If Me.TextQTY.Text <> "" And Me.LBLPrice.Caption <> "" Then

  vPrice = Me.LBLPrice.Caption
  vQTY = Me.TextQTY.Text
  
  If Me.TextDiscount.Text <> "%" Then
    If InStr(1, Me.TextDiscount.Text, "%") > 0 Then
      vDiscount1 = Left(Me.TextDiscount.Text, InStr(1, Me.TextDiscount.Text, "%") - 1)
      vDiscount = vQTY * ((vPrice * vDiscount1) / 100)
    Else
    If Me.TextDiscount.Text <> "" Then
      vDiscount1 = Me.TextDiscount.Text
      Else
      vDiscount1 = 0
      End If
      vDiscount = vDiscount1 * vQTY
    End If
    Me.LBLItemDiscount.Caption = Format(vDiscount, "##,##0.00")
    vAmount = (vQTY * vPrice) - vDiscount
    Me.LBLTotalAmount.Caption = Format(vAmount, "##,##0.00")
    Else
    Me.LBLTotalAmount.Caption = Format(vPrice * vQTY, "##,##0.00")
  End If

End If

Me.TextQTY.SetFocus
End Sub

Private Sub ListView101_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim vSumPriceOfLine As Currency
Dim j As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 And vEdit = 0 Then
    If KeyCode = 46 Then
        i = ListView101.SelectedItem.Index
        vSumPriceOfLine = ListView101.ListItems.Item(i).SubItems(6)
        Text111.Text = Format((CCur(Text111.Text) - CCur(vSumPriceOfLine)), "##,##0.00")
        If Text113.Text = "" Then
            Text113.Text = Format("0", "##,##0.00")
        End If
        Text112.Text = Format((((CCur(Text111.Text) - CCur(Text113.Text)) * 7) / 100), "##,##0.00")
        Text114.Text = Format(CCur(Text111.Text) - CCur(Text113.Text), "##,##0.00")
        
        ListView101.ListItems.Remove (i)
        For j = 1 To ListView101.ListItems.Count
            ListView101.ListItems.Item(j).Text = j
        Next j
    End If
End If
If ListView101.ListItems.Count = 0 Then
    Text111.Text = "0.00"
    Text112.Text = "0.00"
    Text113.Text = "0.00"
    Text114.Text = "0.00"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListView101_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then
        Popup_Menu Menu1
    End If
End Sub

Private Sub ListView102_DblClick()
Dim i As Integer

On Error GoTo ErrDescription

If ListView102.ListItems.Count <> 0 Then
    i = ListView102.SelectedItem.Index
    Text102.Text = Trim(ListView102.ListItems.Item(i).Text)
    Text103.Text = Trim(ListView102.ListItems.Item(i).SubItems(1))
    PIC101.Visible = False
    'Text104.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub





Private Sub ListView104_DblClick()

On Error GoTo ErrDescription

    If ListView104.ListItems.Count <> 0 Then
    Text109.Text = Trim(ListView104.SelectedItem.Text)
    ListView104.Visible = False
    DTPicker102.SetFocus
    End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListView104_LostFocus()
ListView104.Visible = False
End Sub

Private Sub ListView105_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim i As Integer
Dim vListDocno As ListItem

On Error GoTo ErrDescription

If ListView105.ListItems.Count <> 0 Then
    vDocNo = Trim(ListView105.SelectedItem.SubItems(1))
    vIsOpenQuotation = 1
    vQuery = "exec dbo.USP_NP_SearchQuotationDetails '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        Text101.Text = Trim(vRecordset.Fields("docno").Value)
        Text102.Text = Trim(vRecordset.Fields("arcode").Value)
        Text103.Text = Trim(vRecordset.Fields("arname").Value)
        Text109.Text = Trim(vRecordset.Fields("salecode").Value)
        Text110.Text = Trim(vRecordset.Fields("creditday").Value)
        Text111.Text = Format(Trim(vRecordset.Fields("sumofitemamount").Value), "##,##0.00")
        Text112.Text = Format(Trim(vRecordset.Fields("taxamount").Value), "##,##0.00")
        Text113.Text = Format(Trim(vRecordset.Fields("discountamount").Value), "##,##0.00")
        Text114.Text = Format(Trim(vRecordset.Fields("totalamount").Value), "##,##0.00")
        Text115.Text = Trim(vRecordset.Fields("mydescription").Value)
        DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
        DTPicker102.Value = Trim(vRecordset.Fields("validate").Value)
        If Trim(vRecordset.Fields("billtype").Value) = 1 Then
            CMB101.ListIndex = 1
        Else
            CMB101.ListIndex = 0
        End If
        If Trim(vRecordset.Fields("isconditionsend").Value) = 1 Then
            CMB102.ListIndex = 1
        Else
            CMB102.ListIndex = 0
        End If
        If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
            Image101.Visible = False
            Image102.Visible = True
            CMB101.Enabled = False
            CMD105.Enabled = False
            CMD106.Enabled = False
            CMD107.Enabled = False
            CMD108.Enabled = False
            'CMD103.Enabled = False
        Else
            Image101.Visible = True
            Image102.Visible = False
        End If
        
        i = 1
        ListView101.ListItems.Clear
        While Not vRecordset.EOF
        Set vListDocno = ListView101.ListItems.Add(, , i)
        vListDocno.SubItems(1) = Trim(vRecordset.Fields("itemname").Value)
        vListDocno.SubItems(2) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
        vListDocno.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
        vListDocno.SubItems(4) = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
        vListDocno.SubItems(5) = Format(Trim(vRecordset.Fields("discountamountsub").Value), "##,##0.00")
        vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("amount").Value), "##,##0.00")
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
    Pic102.Visible = False
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocsearch As String
Dim vListDocSearch As ListItem
Dim i As Integer
Dim vTypeSearch As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Text1.Text = "" Then
        vTypeSearch = 1
    Else
        vTypeSearch = 0
    End If
    ListView105.ListItems.Clear
    i = 1
    vDocsearch = Trim(Text1.Text)
    vQuery = "exec dbo.USP_NP_SearchQuotation " & vTypeSearch & ",'" & vDocsearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set vListDocSearch = ListView105.ListItems.Add(, , i)
            vListDocSearch.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
            vListDocSearch.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
            vListDocSearch.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Text102_Change()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARName As String
Dim vARCode As String

On Error GoTo ErrDescription

If Text102.Text <> "" Then
    vARCode = Trim(Text102.Text)
    vQuery = "exec dbo.USP_NP_SearchArCode '" & vARCode & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vARName = Trim(vRecordset.Fields("name1").Value)
    End If
    vRecordset.Close
    Text103.Text = vARName
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub



Private Sub Text113_LostFocus()
Dim vTextDiscount As String

On Error GoTo ErrDescription

If Text113.Text = "" Then
    Text113.Text = 0
End If

If CCur(Text113.Text <> CCur("0")) Then
    If CCur(Text111.Text) - CCur(Text114.Text) <> CCur(Text113.Text) Then
        vTextDiscount = Trim(Text113.Text)
        
        Call CheckNumeric(vTextDiscount)
        If vCheckValue = False Then
            MsgBox "กรุณาใส่จำนวนเป็นตัวเลข", vbCritical, "Send Error"
            Text113.Text = "0.00"
            Text113.SetFocus
            Exit Sub
        End If
        Text114.Text = Format(CCur(Text111.Text) - CCur(Text113.Text), "##,##0.00")
        Text112.Text = Format((CCur(Text114.Text) * 7) / 100, "##,##0.00")
        Text113.Text = Format(CCur(Text113.Text), "##,##0.00")
    End If
Else
    If Text113.Text = "" Then
        Text113.Text = Format(CCur("0"), "##,##0.00")
    End If
    If CCur(Text111.Text) - CCur(Text114.Text) <> CCur(Text113.Text) Then
        Text114.Text = Format(CCur(Text111.Text), "##,##0.00")
        Text112.Text = Format((CCur(Text114.Text) * 7) / 100, "##,##0.00")
        Text113.Text = Format(CCur(Text113.Text), "##,##0.00")
    Else
        Text112.Text = Format(CCur(Text112.Text), "##,##0.00")
        Text114.Text = Format(CCur(Text114.Text), "##,##0.00")
        Text113.Text = Format(CCur(Text113.Text), "##,##0.00")
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Text201_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARName As String
Dim vARCode As String
Dim vListAR As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Text201.Text <> "" Then
        ListView102.ListItems.Clear
        vARCode = Trim(Text201.Text)
        vQuery = "exec dbo.USP_NP_SearchArCodeLike '" & vARCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                Set vListAR = ListView102.ListItems.Add(, , Trim(vRecordset.Fields("code").Value))
                vListAR.SubItems(1) = Trim(vRecordset.Fields("name1").Value)
                vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
        Text103.Text = vARName
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Public Sub CheckNumeric(vData As String)
Dim vDocNo As String
Dim vText As String
Dim i As Integer

On Error GoTo ErrDescription

vDocNo = Trim(vData)

For i = 1 To Len(vData)
    If Mid(vDocNo, i, 1) = 0 Or Mid(vDocNo, i, 1) = 1 Or Mid(vDocNo, i, 1) = 2 Or Mid(vDocNo, i, 1) = 3 Or Mid(vDocNo, i, 1) = 4 Or Mid(vDocNo, i, 1) = 5 Or Mid(vDocNo, i, 1) = 6 Or Mid(vDocNo, i, 1) = 7 Or Mid(vDocNo, i, 1) = 8 Or Mid(vDocNo, i, 1) = 9 Or Mid(vDocNo, i, 1) = "." Or Mid(vDocNo, i, 1) = "," Then
        vCheckValue = True
    Else
        vCheckValue = False
        Exit Sub
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub CheckPercent(vData As String)
Dim vDocNo As String
Dim vText As String
Dim i As Integer

On Error GoTo ErrDescription

vDocNo = Trim(vData)

For i = 1 To Len(vData)
    If Mid(vDocNo, i, 1) = 0 Or Mid(vDocNo, i, 1) = 1 Or Mid(vDocNo, i, 1) = 2 Or Mid(vDocNo, i, 1) = 3 Or Mid(vDocNo, i, 1) = 4 Or Mid(vDocNo, i, 1) = 5 Or Mid(vDocNo, i, 1) = 6 Or Mid(vDocNo, i, 1) = 7 Or Mid(vDocNo, i, 1) = 8 Or Mid(vDocNo, i, 1) = 9 Or Mid(vDocNo, i, 1) = "." Or Mid(vDocNo, i, 1) = "," Or Mid(vDocNo, i, 1) = "%" Then
        vCheckPercent = True
    Else
        vCheckPercent = False
        Exit Sub
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Popup_Menu(m As Menu)
    Menu1.Visible = True
    PopupMenu m, 4
    Menu1.Visible = False
End Sub

Private Sub menuedit_click()
If ListView101.ListItems.Count <> 0 Then
vIndex = ListView101.SelectedItem.Index


vEdit = 1
End If
End Sub

Private Sub TextDiscount_Change()
Dim vQTY As Double
Dim vAmount As Double
Dim vDiscount As Double
Dim vPrice As Double
Dim vDiscountAmount  As Double
Dim vDiscountWord As String
Dim vDiscount1 As Double

If Me.LBLPrice.Caption <> "" And Me.TextQTY.Text <> "" And Me.TextDiscount.Text <> "" Then
  vPrice = Me.LBLPrice.Caption
  vQTY = Me.TextQTY.Text
  
    Call CheckPercent(TextDiscount.Text)
    
    If InStr(1, Me.TextDiscount.Text, "%") = 1 Then
    TextDiscount.Text = ""
    LBLItemDiscount.Caption = ""
    vCheckPercent = False
    TextDiscount.SetFocus
    End If
    
    If InStr(1, Me.TextDiscount.Text, ".") = 1 And Len(Me.TextDiscount.Text) = 1 Then
      Me.LBLTotalAmount.Caption = Format(vPrice * vQTY, "##,##0.00")
      Exit Sub
    End If
    
    If InStr(1, Me.TextDiscount.Text, "%") > 0 Then
      If Len(Me.TextDiscount.Text) <> InStr(1, Me.TextDiscount.Text, "%") Then
      TextDiscount.Text = ""
      LBLItemDiscount.Caption = ""
      vCheckPercent = False
      TextDiscount.SetFocus
    End If
  End If

    If vCheckPercent = False Then
        MsgBox "กรุณาใส่จำนวนเป็นตัวเลข", vbCritical, "Send Error"
        TextDiscount.Text = ""
        LBLItemDiscount.Caption = ""
        vCheckPercent = False
        TextDiscount.SetFocus
        
        Me.LBLTotalAmount.Caption = Format(vPrice * vQTY, "##,##0.00")
        Exit Sub
    End If
      
  If Me.TextDiscount.Text <> "%" Then
    If InStr(1, Me.TextDiscount.Text, "%") > 0 Then
      vDiscount1 = Left(Me.TextDiscount.Text, InStr(1, Me.TextDiscount.Text, "%") - 1)
      vDiscount = vQTY * ((vPrice * vDiscount1) / 100)
    Else
      If Me.TextDiscount.Text <> "" Then
        vDiscount1 = Me.TextDiscount.Text
      Else
        vDiscount1 = 0
      End If
      vDiscount = vDiscount1 * vQTY
    End If
    Me.LBLItemDiscount.Caption = Format(vDiscount, "##,##0.00")
    vAmount = (vQTY * vPrice) - vDiscount
    Me.LBLTotalAmount.Caption = Format(vAmount, "##,##0.00")
    Else
    Me.LBLTotalAmount.Caption = Format(vPrice * vQTY, "##,##0.00")
  End If
ElseIf Me.LBLPrice.Caption <> "" And Me.TextQTY.Text <> "" Then
  vPrice = Me.LBLPrice.Caption
  vQTY = Me.TextQTY.Text
  Me.LBLTotalAmount.Caption = Format(vPrice * vQTY, "##,##0.00")
End If
End Sub

Private Sub TextDiscount_LostFocus()
If InStr(1, Me.TextDiscount.Text, ".") = 1 And Len(Me.TextDiscount.Text) = 1 Then
Me.TextDiscount.Text = ""
End If
End Sub

Private Sub TextItemCode_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vItemCode As String
Dim i As Integer

If Me.TextItemCode.Text <> "" Then
      vItemCode = Trim(Me.TextItemCode.Text)
      vQuery = "exec dbo.USP_NP_QuotationSearchItem '" & vItemCode & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vRecordset.MoveFirst
          Me.LBLItemName.Caption = Trim(vRecordset.Fields("name1").Value)
          Me.LBLOnHand.Caption = Format(vRecordset.Fields("stockqty").Value, "##,##0.00")
          Me.LBLUnitCode.Caption = Trim(vRecordset.Fields("defsaleunitcode").Value)
      While Not vRecordset.EOF
        Me.ListUnitCode.AddItem Trim(vRecordset.Fields("unitcode").Value)
      vRecordset.MoveNext
      Wend
      Else
      Me.LBLItemName.Caption = ""
      Me.LBLOnHand.Caption = ""
      Me.LBLUnitCode.Caption = ""
      Me.TextQTY.Text = ""
      Me.TextDiscount.Text = ""
      Me.LBLTotalAmount.Caption = ""
      Me.ListUnitCode.Clear
      Me.LBLPrice.Caption = ""
      End If
      vRecordset.Close
End If

End Sub

Private Sub TextQTY_Change()
Dim vQTY As Double
Dim vAmount As Double
Dim vDiscount As Double
Dim vPrice As Double
Dim vDiscountAmount  As Double
Dim vDiscountWord As String
Dim vDiscount1 As Double

If TextQTY.Text <> "" And Me.LBLPrice.Caption <> "" Then

    Call CheckNumeric(TextQTY.Text)
    If vCheckValue = False Then
        MsgBox "กรุณาใส่จำนวนเป็นตัวเลข", vbCritical, "Send Error"
        TextQTY.Text = ""
        TextQTY.SetFocus
        Exit Sub
    End If
    
    vPrice = Me.LBLPrice.Caption

    vQTY = Me.TextQTY.Text
    
    vAmount = vQTY * vPrice
    If Me.TextDiscount.Text = "" Then
      vDiscount = 0
      vAmount = vQTY * vPrice
    Else
      vDiscountWord = Me.TextDiscount.Text
      If InStr(1, vDiscountWord, "%") > 0 Then
         vDiscount1 = Left(Trim(vDiscountWord), InStr(1, vDiscountWord, "%") - 1)
         vDiscount = vQTY * (vPrice * vDiscount1) / 100
      Else
          vDiscount1 = vDiscountWord
          vDiscount = vQTY * vDiscount1
      End If
      vAmount = vAmount - vDiscount
    End If
    Me.LBLItemDiscount.Caption = Format(vDiscount, "##,##0.00")
    Me.LBLTotalAmount.Caption = Format(vAmount, "##,##0.00")
End If
End Sub
