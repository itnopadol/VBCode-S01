VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form3_19 
   Caption         =   "ขอยกเลิก Back Order"
   ClientHeight    =   9000
   ClientLeft      =   4035
   ClientTop       =   840
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form3_19.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICConfirmCancelBackOrder 
      BackColor       =   &H00808080&
      Height          =   9150
      Left            =   3330
      ScaleHeight     =   9090
      ScaleWidth      =   27360
      TabIndex        =   12
      Top             =   -45
      Visible         =   0   'False
      Width           =   27420
      Begin VB.CheckBox CHKRequestNo 
         BackColor       =   &H00808080&
         Caption         =   "มีเอกสารใบขอยกเลิกใบสั่งขายค้างส่ง (Back Order)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   675
         TabIndex        =   49
         Top             =   3285
         Width           =   5640
      End
      Begin VB.CheckBox CHKAll 
         BackColor       =   &H00808080&
         Caption         =   "ยกเลิกทั้งใบ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   675
         TabIndex        =   44
         Top             =   5850
         Width           =   1275
      End
      Begin MSComctlLib.ListView ListViewItemCancel 
         Height          =   2085
         Left            =   675
         TabIndex        =   48
         Top             =   990
         Width           =   10545
         _ExtentX        =   18600
         _ExtentY        =   3678
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
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CheckBox CHKItem 
         BackColor       =   &H00808080&
         Caption         =   "ยกเลิกบางส่วน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   675
         TabIndex        =   45
         Top             =   6435
         Width           =   1725
      End
      Begin VB.TextBox TXTVendorDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3825
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   5310
         Visible         =   0   'False
         Width           =   7395
      End
      Begin VB.TextBox TXTRequestNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   9495
         TabIndex        =   38
         Top             =   6975
         Width           =   1725
      End
      Begin VB.TextBox TXTReserveNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6030
         TabIndex        =   37
         Top             =   6975
         Width           =   1725
      End
      Begin VB.TextBox TXTBackOrderNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   36
         Top             =   6975
         Width           =   1725
      End
      Begin VB.CheckBox CHKReqOrder 
         BackColor       =   &H00808080&
         Caption         =   "ใบเสนอซื้อสินค้า"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7830
         TabIndex        =   34
         Top             =   6975
         Width           =   1635
      End
      Begin VB.CheckBox CHKSaleOrder 
         BackColor       =   &H00808080&
         Caption         =   "ใบสั่งจอง/สั่งขาย"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   33
         Top             =   6975
         Width           =   1725
      End
      Begin VB.CheckBox CHKBackOrder 
         BackColor       =   &H00808080&
         Caption         =   "ใบ Back Order"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         TabIndex        =   32
         Top             =   6975
         Width           =   1635
      End
      Begin VB.CheckBox CHKDocument 
         BackColor       =   &H00808080&
         Caption         =   "มีเอกสารทดแทน ครบ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         TabIndex        =   31
         Top             =   6435
         Width           =   2040
      End
      Begin VB.CheckBox CHKDeposit 
         BackColor       =   &H00808080&
         Caption         =   "มียอดมัดจำคงเหลือ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   675
         TabIndex        =   30
         Top             =   4230
         Width           =   1860
      End
      Begin VB.CheckBox CHKReceive 
         BackColor       =   &H00808080&
         Caption         =   "รับเข้าสินค้าแล้ว"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   675
         TabIndex        =   29
         Top             =   4770
         Width           =   1635
      End
      Begin VB.CommandButton CMDVendorDescription 
         Height          =   285
         Left            =   3420
         Picture         =   "Form3_19.frx":72FB
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   5310
         Width           =   330
      End
      Begin VB.CommandButton CMDExit 
         Caption         =   "ออก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2520
         TabIndex        =   26
         Top             =   7605
         Width           =   1230
      End
      Begin VB.CommandButton CMDApprove 
         Caption         =   "บันทึกยกเลิก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   675
         TabIndex        =   25
         Top             =   7605
         Width           =   1230
      End
      Begin VB.CheckBox CHKVendorCancel 
         BackColor       =   &H00808080&
         Caption         =   "ยกเลิกกับทาง ผู้แทนจำหน่ายได้"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   675
         TabIndex        =   24
         Top             =   5310
         Width           =   2760
      End
      Begin VB.CheckBox CHKRequestQTY 
         BackColor       =   &H00808080&
         Caption         =   "มียอด ตามจำนวนขอยกเลิก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   675
         TabIndex        =   23
         Top             =   3735
         Width           =   3255
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   0
         X2              =   12015
         Y1              =   8235
         Y2              =   8235
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   0
         X2              =   12015
         Y1              =   7425
         Y2              =   7425
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   0
         X2              =   12015
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line Line2 
         BorderWidth     =   5
         X1              =   0
         X2              =   12015
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   0
         X2              =   12015
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   0
         Picture         =   "Form3_19.frx":76E0
         Top             =   0
         Width           =   2160
      End
      Begin VB.Label LBLDepositBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2790
         TabIndex        =   35
         Top             =   4230
         Width           =   1725
      End
      Begin VB.Image IMG104 
         Height          =   300
         Left            =   2295
         Picture         =   "Form3_19.frx":8B42
         Top             =   450
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Image IMG103 
         Height          =   300
         Left            =   2295
         Picture         =   "Form3_19.frx":9025
         Top             =   450
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Image IMG101 
         Height          =   300
         Left            =   2295
         Picture         =   "Form3_19.frx":9561
         Top             =   450
         Width           =   570
      End
      Begin VB.Image IMG102 
         Height          =   300
         Left            =   2295
         Picture         =   "Form3_19.frx":9993
         Top             =   450
         Visible         =   0   'False
         Width           =   570
      End
   End
   Begin VB.CommandButton CMDConfirmCancelBackOrder 
      Caption         =   "อนุมัติยกเลิก Back Order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4905
      TabIndex        =   11
      Top             =   6840
      Width           =   2265
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   810
      Top             =   7740
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   -225
      ScaleHeight     =   90
      ScaleWidth      =   12195
      TabIndex        =   10
      Top             =   7560
      Width           =   12255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   -270
      ScaleHeight     =   90
      ScaleWidth      =   12375
      TabIndex        =   9
      Top             =   6570
      Width           =   12435
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   360
      Top             =   7740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMDCheckDocNo 
      Caption         =   "ตรวจสอบข้อมูล"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   45
      TabIndex        =   8
      Top             =   6840
      Width           =   2265
   End
   Begin VB.Timer Timer101 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   -90
      Top             =   7740
   End
   Begin MSComctlLib.ListView ListViewItemBackOrder 
      Height          =   3660
      Left            =   45
      TabIndex        =   5
      Top             =   2700
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   6456
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777200
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   26
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "จำนวน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "คงเหลือ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "จำนวนขอยกเลิก"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "หน่วย"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ชนิดสินค้า"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "เอกสารขาย"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ใบเสนอซื้อ"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "อนุมัติซื้อ"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ใบสั่งซื้อ"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "ใบรับเข้า"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "บิล"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "RemainBO"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "ขายจำนวน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Remain ขาย"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "ทำบิลจำนวน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "เสนอซื้อจำนวน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "Remainเสนอซื้อ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "จำนวนอนุมัติซื้อ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Text            =   "Remainอนุมัติซื้อ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Text            =   "สั่งซื้อจำนวน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Text            =   "Remainสังซื้อ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "รับเข้าจำนวน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "SpecialOrder"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.CommandButton CMDReqBackOrderCancel 
      Caption         =   "ใบขอยกเลิก Back Order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2475
      TabIndex        =   4
      Top             =   6840
      Width           =   2265
   End
   Begin VB.TextBox TextDocno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1530
      TabIndex        =   0
      Top             =   1350
      Width           =   1770
   End
   Begin VB.PictureBox PICShowChangeQTY 
      BackColor       =   &H8000000C&
      Height          =   3615
      Left            =   90
      ScaleHeight     =   3555
      ScaleWidth      =   11610
      TabIndex        =   39
      Top             =   2745
      Visible         =   0   'False
      Width           =   11670
      Begin VB.CommandButton CMDChangeQTY 
         Caption         =   "เปลี่ยน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1755
         TabIndex        =   43
         Top             =   1980
         Width           =   1095
      End
      Begin VB.OptionButton OPTRemainQTY 
         BackColor       =   &H8000000C&
         Caption         =   "ตามจำนวนคงเหลือจากการขาย"
         Height          =   375
         Left            =   1710
         TabIndex        =   41
         Top             =   1080
         Width           =   2490
      End
      Begin VB.OptionButton OPTSaleQTY 
         BackColor       =   &H8000000C&
         Caption         =   "ตามจำนวนที่สั่งขาย"
         Height          =   330
         Left            =   1710
         TabIndex        =   40
         Top             =   675
         Value           =   -1  'True
         Width           =   2355
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000C&
         Caption         =   "เลือกจำนวนที่ต้องการยกเลิก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         TabIndex        =   42
         Top             =   135
         Width           =   2625
      End
   End
   Begin VB.Label LBLCancelNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1530
      TabIndex        =   47
      Top             =   990
      Width           =   1770
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ขอยกเลิก :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   135
      TabIndex        =   46
      Top             =   990
      Width           =   1275
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "มัดจำคงเหลือ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9090
      TabIndex        =   22
      Top             =   2070
      Width           =   1230
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดมัดจำ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6750
      TabIndex        =   21
      Top             =   2070
      Width           =   915
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ใบมัดจำ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4365
      TabIndex        =   20
      Top             =   2070
      Width           =   825
   End
   Begin VB.Label LBLDepNetAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7740
      TabIndex        =   19
      Top             =   2070
      Width           =   1365
   End
   Begin VB.Label LBLDepBalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10395
      TabIndex        =   18
      Top             =   2070
      Width           =   1365
   End
   Begin VB.Label LBLDepNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5265
      TabIndex        =   17
      Top             =   2070
      Width           =   1365
   End
   Begin VB.Label LBLDocDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4590
      TabIndex        =   16
      Top             =   1350
      Width           =   1725
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เอกสาร :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3330
      TabIndex        =   15
      Top             =   1350
      Width           =   1140
   End
   Begin VB.Label LBLSaleName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1530
      TabIndex        =   14
      Top             =   2070
      Width           =   2760
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "พนักงานขาย :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   270
      TabIndex        =   13
      Top             =   2070
      Width           =   1140
   End
   Begin VB.Image IMGBillStatus 
      Height          =   300
      Left            =   2835
      Picture         =   "Form3_19.frx":9E3C
      Top             =   360
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image IMGConfirm 
      Height          =   300
      Left            =   2835
      Picture         =   "Form3_19.frx":A31F
      Top             =   360
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image IMGCancel 
      Height          =   300
      Left            =   2835
      Picture         =   "Form3_19.frx":A7C8
      Top             =   360
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image IMGNew 
      Height          =   300
      Left            =   2835
      Picture         =   "Form3_19.frx":AD04
      Top             =   360
      Width           =   570
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการสินค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   45
      TabIndex        =   7
      Top             =   2430
      Width           =   1320
   End
   Begin VB.Label LBLARCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1530
      TabIndex        =   6
      Top             =   1710
      Width           =   1770
   End
   Begin VB.Label LBLARName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3330
      TabIndex        =   3
      Top             =   1710
      Width           =   8430
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   1710
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   315
      TabIndex        =   1
      Top             =   1350
      Width           =   1095
   End
End
Attribute VB_Name = "Form3_19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vIsCancel  As Integer
Dim vIsConfirm As Integer
Dim vBillStatus As Integer
Dim vSOBillStatus As Integer
Dim vMemCheckReqCancel As Integer
Dim vMemPercentProcess As Integer
Dim vSelectIndexChangeQTY As Integer

'Private Sub CHKItem_Click()
'If Me.CHKItem.Value = 1 Then
 '  Me.CHK105.Enabled = True
  ' Me.CHKBackOrder.Enabled = True
   'Me.CHKSaleOrder.Enabled = True
   'Me.CHKReqOrder.Enabled = True
   'Me.TXTBackOrderNo.Enabled = True
   'Me.TXTReserveNo.Enabled = True
   'Me.TXTRequestNo.Enabled = True
'Else
 '  Me.CHK105.Enabled = False
  ' Me.CHKBackOrder.Enabled = False
   ''Me.CHKSaleOrder.Enabled = False
   'Me.CHKReqOrder.Enabled = False
   'Me.TXTBackOrderNo.Enabled = False
   'Me.TXTReserveNo.Enabled = False
   'Me.TXTRequestNo.Enabled = False
'End If
'End Sub

Private Sub CMDChangeQTY_Click()
'Dim vChangeQTY As Integer

'If Me.OPTSaleQTY.Value = True Then
 '  vChangeQTY = Me.ListViewItemBackOrder.ListItems(vSelectIndexChangeQTY).ListSubItems(3)
'End If

'If Me.OPTRemainQTY.Value = True Then
 '  vChangeQTY = Me.ListViewItemBackOrder.ListItems(vSelectIndexChangeQTY).ListSubItems(4)
'End If
'Me.ListViewItemBackOrder.ListItems(vSelectIndexChangeQTY).ListSubItems(5) = Format(vChangeQTY, "##,#0.00")
'Me.OPTSaleQTY.Value = True
'Me.PICShowChangeQTY.Visible = False

End Sub

Private Sub CMDCheckDocNo_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vListItem As ListItem
Dim i As Integer


'If Me.TextDocno.Text <> "" Then
 '  vDocNo = Me.TextDocno.Text
  ' Me.ListViewItemBackOrder.ListItems.Clear
   'vQuery = "exec dbo.USP_BO_SearchBackOrderDetails '" & vDocNo & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '  Me.ListViewItemBackOrder.ListItems.Clear
     ' vIsCancel = vRecordset.Fields("iscancel").Value
      'vIsConfirm = vRecordset.Fields("isconfirm").Value
      'vBillStatus = vRecordset.Fields("billstatus").Value
      'vSOBillStatus = vRecordset.Fields("billstatus").Value
      'Me.LBLARCode.Caption = vRecordset.Fields("arcode").Value
      'Me.LBLARName.Caption = vRecordset.Fields("arname").Value
      'vRecordset.MoveFirst
      'i = 1
      'While Not vRecordset.EOF
      'Set vListItem = Me.ListViewItemBackOrder.ListItems.Add(, , i)
      'vListItem.SubItems(1) = vRecordset.Fields("itemcode").Value
      'vListItem.SubItems(2) = vRecordset.Fields("itemname").Value
      'vListItem.SubItems(3) = Format(vRecordset.Fields("qty").Value, "##,##0.00")
      'vListItem.SubItems(4) = Format(vRecordset.Fields("remainqty").Value, "##,##0.00")
      'vListItem.SubItems(5) = Format(vRecordset.Fields("qty").Value, "##,##0.00")
      'vListItem.SubItems(6) = vRecordset.Fields("unitcode").Value
      'If vRecordset.Fields("itemstatus").Value = "1" Then
       '  vListItem.SubItems(7) = "ปกติ"
      'ElseIf vRecordset.Fields("itemstatus").Value = "3" Then
       '  vListItem.SubItems(7) = "สั่งพิเศษ"
        ' Me.ListViewItemBackOrder.ListItems(i).ListSubItems(7).ForeColor = "&H000000C0"
      'ElseIf vRecordset.Fields("itemstatus").Value = "4" Then
       '  vListItem.SubItems(7) = "ของแถม"
        ' Me.ListViewItemBackOrder.ListItems(i).ListSubItems(7).ForeColor = "&H00800000"
      'Else
       '  vListItem.SubItems(7) = ""
      'End If
       
      'vListItem.SubItems(8) = vRecordset.Fields("saleno").Value
      'vListItem.SubItems(9) = vRecordset.Fields("requestno").Value
      'vListItem.SubItems(10) = vRecordset.Fields("confirmno").Value
      'vListItem.SubItems(11) = vRecordset.Fields("purchaseno").Value
      'vListItem.SubItems(12) = vRecordset.Fields("recno").Value
      'vListItem.SubItems(13) = vRecordset.Fields("invno").Value
      'vListItem.SubItems(14) = Format(vRecordset.Fields("boremainqty").Value, "##,##0.00")
      'vListItem.SubItems(15) = Format(vRecordset.Fields("saleqty").Value, "##,##0.00")
      'vListItem.SubItems(16) = Format(vRecordset.Fields("soremainqty").Value, "##,##0.00")
      'vListItem.SubItems(17) = Format(vRecordset.Fields("invqty").Value, "##,##0.00")
      'vListItem.SubItems(18) = Format(vRecordset.Fields("reqqty").Value, "##,##0.00")
      'vListItem.SubItems(19) = Format(vRecordset.Fields("reqremainqty").Value, "##,##0.00")
      'vListItem.SubItems(20) = Format(vRecordset.Fields("confirmqty").Value, "##,##0.00")
      'vListItem.SubItems(21) = Format(vRecordset.Fields("confremainqty").Value, "##,##0.00")
      'vListItem.SubItems(22) = Format(vRecordset.Fields("purchaseqty").Value, "##,##0.00")
      'vListItem.SubItems(23) = Format(vRecordset.Fields("purremainqty").Value, "##,##0.00")
      'vListItem.SubItems(24) = Format(vRecordset.Fields("recqty").Value, "##,##0.00")
      'vListItem.SubItems(25) = Format(vRecordset.Fields("itemstatus").Value, "##,##0.00")
      
Dim vCheckSORemain As Double
Dim vCheckBORemain As Double
Dim vCheckPRRemain As Double
Dim vCheckAPRemain As Double
Dim vCheckPORemain As Double
Dim vCheckItemStatus As Integer

      'vCheckSORemain = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(16)
      'vCheckBORemain = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(14)
  ''    vCheckPRRemain = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(19)
      'vCheckAPRemain = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(21)
      'vCheckPORemain = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(23)
      'vCheckItemStatus = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(25)
      
      'If (vCheckBORemain > 0 Or vCheckSORemain > 0) And vCheckPORemain > 0 Then
       '  Me.ListViewItemBackOrder.ListItems(i).Checked = True
        ' Me.CMDReqBackOrderCancel.Enabled = True
         'Me.CMDConfirmCancelBackOrder.Enabled = True
         'Me.ListViewItemBackOrder.ListItems(i).ListSubItems(1).ForeColor = "&H000000C0"
      'ElseIf (vCheckBORemain > 0 Or vCheckSORemain > 0) And vCheckPORemain <= 0 And vCheckItemStatus <> 3 Then
       '  Me.ListViewItemBackOrder.ListItems(i).Checked = True
        ' Me.CMDReqBackOrderCancel.Enabled = True
         ''Me.CMDConfirmCancelBackOrder.Enabled = True
         'Me.ListViewItemBackOrder.ListItems(i).ListSubItems(1).ForeColor = "&H000000C0"
      'ElseIf (vCheckBORemain > 0 Or vCheckSORemain > 0) And vCheckPORemain <= 0 And (vCheckPRRemain > 0 Or vCheckAPRemain > 0) Then
       '  Me.ListViewItemBackOrder.ListItems(i).Checked = True
        ' Me.CMDReqBackOrderCancel.Enabled = True
         ''Me.CMDConfirmCancelBackOrder.Enabled = True
         'Me.ListViewItemBackOrder.ListItems(i).ListSubItems(1).ForeColor = "&H000000C0"
      'ElseIf (vCheckBORemain > 0 Or vCheckSORemain > 0) And (vCheckPRRemain = 0 And vCheckAPRemain = 0) And vCheckPORemain <= 0 And vCheckItemStatus = 3 Then
       '  Me.ListViewItemBackOrder.ListItems(i).Checked = False
      'ElseIf vCheckBORemain = 0 And vCheckSORemain = 0 Then
       '  Me.ListViewItemBackOrder.ListItems(i).Checked = False
      'End If
      
      'i = i + 1
      'vRecordset.MoveNext
      'Wend
   'End If
   'vRecordset.Close
   
   Dim vDocDetails As ListItem
   
   'If vIsCancel = 0 And vIsConfirm = 0 And vBillStatus = 0 Then
    '  Me.CMDCheckDocNo.Enabled = False
     ' Me.CMDReqBackOrderCancel.Enabled = False

      'Me.IMGNew.Visible = True
      'Me.IMGCancel.Visible = False
      'Me.IMGConfirm.Visible = False
      'Me.IMGBillStatus.Visible = False

   'ElseIf vIsCancel = 1 And vIsConfirm = 0 And vBillStatus = 0 Then
    '  MsgBox "เอกสารถูกยกเลิกไปแล้ว", vbCritical, "Send Error Message"
     ' Me.CMDCheckDocNo.Enabled = False
      'Me.CMDReqBackOrderCancel.Enabled = False
      'Me.CMDConfirmCancelBackOrder.Enabled = False
      
      'Me.IMGNew.Visible = False
      'Me.IMGCancel.Visible = True
      'Me.IMGConfirm.Visible = False
      'Me.IMGBillStatus.Visible = False
   'ElseIf vIsCancel = 0 And vIsConfirm = 1 And vBillStatus = 0 Then
    '  Me.CMDCheckDocNo.Enabled = False
     ' Me.IMGNew.Visible = False
      ''Me.IMGCancel.Visible = False
      'Me.IMGConfirm.Visible = True
      'Me.IMGBillStatus.Visible = False
   'ElseIf vIsCancel = 0 And vIsConfirm = 1 And vBillStatus = 1 Then
    '  Me.CMDCheckDocNo.Enabled = False
     ' Me.IMGNew.Visible = False
      'Me.IMGCancel.Visible = False
      'Me.IMGConfirm.Visible = False
      'Me.IMGBillStatus.Visible = True
   'End If
'End If

End Sub


Private Sub CMDConfirmCancelBackOrder_Click()
'PICConfirmCancelBackOrder.Visible = True
End Sub

Private Sub CMDConfirmCancelCancel_Click()
'Me.PICConfirmCancelBackOrder.Visible = False
End Sub

Private Sub CMDReqBackOrderCancel_Click()
Dim vDocNo As String
Dim vDocdate As String
Dim vARCode As String
Dim vApproveStatus As Integer
Dim vIsCancel As Integer
Dim vBackOrderRefNo As String
Dim vNewBackOrderNo As String
Dim vRequestUser As String
Dim vDepositNo As String
Dim vDepositAmount As Double
Dim vDepositAmountBalance As Double
Dim vApproveCode As String
Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Double
Dim vReqQTYCancel As Double
Dim vUnitCode As String
Dim vItemStatus As Integer
Dim vLineNumber As Integer
Dim i As Integer

Dim vCheckItemRequest As Integer
Dim n As Integer

Dim vMonth As String
Dim vYear As String
Dim vRecordset As New ADODB.Recordset
Dim vHeader As String
Dim vRunningNo As Integer
Dim vQuery As String
Dim vMonth1 As String
Dim vYear1 As String
Dim vCheckReqCancel As String
Dim vAnswer As Integer

'If Me.ListViewItemBackOrder.ListItems.Count > 0 Then
'vBackOrderRefNo = Me.TextDocno.Text
'vQuery = "select top 1 isnull(docno,'') as docno  from npmaster.dbo.TB_BC_ReqBackOrderCancel where backorderrefno  = '" & vBackOrderRefNo & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '  vCheckReqCancel = vRecordset.Fields("docno").Value
'End If
'vRecordset.Close

'If vCheckReqCancel <> "" Then
 '  vAnswer = MsgBox("เอกสารเลขที่ " & vBackOrderRefNo & " ได้ทำการขอยกเลิกไปแล้ว เลขที่ใบขอยกเลิกคือ  " & vCheckReqCancel & "  ต้องการพิมพ์ใบขอยกเลิก Back Order หรือไม่", vbYesNo, "Send Infromation Message")
  ' If vAnswer = 6 Then
   '   Call PrintRequestBackOrderCancel(vCheckReqCancel)
   'Else
    '  Exit Sub
   'End If
   'Exit Sub
'End If
      
 '  For n = 1 To Me.ListViewItemBackOrder.ListItems.Count
  ' If Me.ListViewItemBackOrder.ListItems(n).Checked = True Then
   '   vCheckItemRequest = 1
    '  GoTo Line1
   'End If
   'Next n

'Line1:
 '  If vCheckItemRequest = 1 Then
  '    vDocdate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   '   vARCode = Me.LBLARCode.Caption
    '  vApproveStatus = 0
     ' vIsCancel = 0
      'vBackOrderRefNo = Me.TextDocno.Text
      'vNewBackOrderNo = ""
      'vRequestUser = vUserID
      'vApproveCode = ""
      
      'If Me.LBLDepNo.Caption <> "" Then
       '  vDepositNo = Me.LBLDepNo.Caption
      'Else
       '  vDepositNo = ""
      'End If
      
      'If Me.LBLDepNetAmount.Caption <> "" Then
       '  vDepositAmount = Me.LBLDepNetAmount.Caption
      'Else
       '  vDepositAmount = 0
      'End If
      
      'If Me.LBLDepBalance.Caption <> "" Then
       '  vDepositAmountBalance = Me.LBLDepBalance.Caption
      'Else
       '  vDepositAmountBalance = 0
      'End If
      
      'vMonth = Month(Now)
      'vYear = Year(Now)
      'vHeader = "BC"
      
      'vQuery = "exec dbo.USP_BC_GenDocnoBackOrderCancel '" & vMonth & "','" & vYear & "' "
      'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      'vRunningNo = vRecordset.Fields("maxvalue").Value
      'End If
      'vRecordset.Close
      
      'If Len(vMonth) < 2 Then
       '  vMonth1 = Trim("0" & vMonth)
      'End If
      
      
      'If Left(vYear, 2) < "51" Then
       '  vYear1 = Right(vYear + 543, 2)
      'End If
      
      'vDocNo = vHeader & vYear1 & vMonth1 & "-" & Format(vRunningNo, "0000")
      
      'vQuery = "begin tran"
      'gConnection.Execute (vQuery)
      
      'vQuery = "set dateformat dmy"
      'gConnection.Execute (vQuery)
      
      'vQuery = "exec dbo.USP_BC_InsertUpdateReqBackOrderCancel 0,'" & vDocNo & "','" & vDocdate & "','" & vARCode & "'," & vApproveStatus & "," & vIsCancel & ",'" & vBackOrderRefNo & "','" & vNewBackOrderNo & "','" & vRequestUser & "','" & vApproveCode & "','" & vDepositNo & "'," & vDepositAmount & "," & vDepositAmountBalance & " "
      'gConnection.Execute (vQuery)
      
      'vLineNumber = -1
      'For i = 1 To Me.ListViewItemBackOrder.ListItems.Count
      'If Me.ListViewItemBackOrder.ListItems(i).Checked = True Then
       '  vItemCode = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(1)
        ' vItemName = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(2)
         'vQTY = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(3)
         'vReqQTYCancel = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(5)
         'vUnitCode = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(6)
         'vItemStatus = Me.ListViewItemBackOrder.ListItems(i).ListSubItems(25)
         'vLineNumber = vLineNumber + 1
         
         'vQuery = "exec dbo.USP_BC_InsertReqBackOrderCancelSub '" & vDocNo & "','" & vDocdate & "','" & vItemCode & "','" & vItemName & "'," & vQTY & "," & vReqQTYCancel & ",'" & vUnitCode & "'," & vItemStatus & "," & vLineNumber & " "
         'gConnection.Execute (vQuery)
      'End If
      'Next i
      
      'vQuery = "commit tran"
      'gConnection.Execute (vQuery)
      
      'MsgBox "ทำใบขอยกเลิกเรียบร้อยแล้วครับ ได้เอกสารเลขที่ " & vDocNo & " ", vbInformation, "Send Information"
      'Call PrintRequestBackOrderCancel(vDocNo)
      'Me.CMDReqBackOrderCancel.Enabled = False
      
   'End If
'End If

End Sub

Public Sub PrintRequestBackOrderCancel(vDocNo As String)
Dim vRecordset As New ADODB.Recordset
Dim vReportName As String
Dim vQuery As String

'vQuery = "exec dbo.USP_NP_SelectReportName 391,'BO' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '  vReportName = vRecordset.Fields("reportname").Value
'End If
'vRecordset.Close

'With Crystal101
'.ReportFileName = vReportName & ".rpt"
'.ParameterFields(0) = "@vDocNO;" & vDocNo & ";true"
'.Destination = crptToWindow
'.WindowState = crptMaximized
'.Action = 1
'End With

End Sub

Private Sub CMDVendorDescription_Click()
'Me.TXTVendorDescription.Visible = True
End Sub

Private Sub ListViewItemBackOrder_DblClick()
Dim n As Integer
Dim vCheckItemRequest As Integer


'If Me.ListViewItemBackOrder.ListItems.Count > 0 Then

 '  For n = 1 To Me.ListViewItemBackOrder.ListItems.Count
  ' If Me.ListViewItemBackOrder.ListItems(n).Checked = True Then
   '   vCheckItemRequest = 1
    '  GoTo Line1
   'End If
   'Next n
   
'Line1:
   
 '  If vCheckItemRequest <> 0 Then
  '    vSelectIndexChangeQTY = Me.ListViewItemBackOrder.SelectedItem.Index
   '   If Me.ListViewItemBackOrder.ListItems(vSelectIndexChangeQTY).Checked = True Then
    '  Me.PICShowChangeQTY.Visible = True
     ' End If
   'End If
'End If

End Sub

Private Sub ListViewItemBackOrder_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vIndex As Integer
Dim vCheckSORemain As Double
Dim vCheckBORemain As Double
Dim vCheckPORemain As Double
Dim vCheckItemStatus As Integer
      
'If Me.ListViewItemBackOrder.ListItems.Count > 0 Then
 '  vIndex = Item
  
  '    vCheckSORemain = Me.ListViewItemBackOrder.ListItems(vIndex).ListSubItems(16)
   '   vCheckBORemain = Me.ListViewItemBackOrder.ListItems(vIndex).ListSubItems(14)
    '  vCheckPORemain = Me.ListViewItemBackOrder.ListItems(vIndex).ListSubItems(23)
     ' vCheckItemStatus = Me.ListViewItemBackOrder.ListItems(vIndex).ListSubItems(25)
      
      'If (vCheckBORemain > 0 Or vCheckSORemain > 0) And vCheckPORemain > 0 Then
       '  Me.ListViewItemBackOrder.ListItems(vIndex).Checked = True
        ' Me.CMDReqBackOrderCancel.Enabled = True
         'Me.ListViewItemBackOrder.ListItems(vIndex).ListSubItems(1).ForeColor = "&H000000C0"
      'ElseIf (vCheckBORemain > 0 Or vCheckSORemain > 0) And vCheckPORemain = 0 And vCheckItemStatus <> 3 Then
       '  Me.ListViewItemBackOrder.ListItems(vIndex).Checked = True
        ' Me.CMDReqBackOrderCancel.Enabled = True
         'Me.ListViewItemBackOrder.ListItems(vIndex).ListSubItems(1).ForeColor = "&H000000C0"
      'ElseIf (vCheckBORemain > 0 Or vCheckSORemain > 0) And vCheckPORemain = 0 And vCheckItemStatus = 3 Then
       '  Me.ListViewItemBackOrder.ListItems(vIndex).Checked = False
      'ElseIf vCheckBORemain = 0 And vCheckSORemain = 0 Then
       '  Me.ListViewItemBackOrder.ListItems(vIndex).Checked = False
      'End If
'End If
End Sub

Private Sub TextDocno_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 8 Or KeyCode = 46 Then
 '  If Me.TextDocno.Text <> "" Then
  '    Me.LBLDocDate.Caption = ""
   '   Me.LBLARCode.Caption = ""
    '  Me.LBLARName.Caption = ""
     ' Me.LBLSaleName.Caption = ""
      'Me.LBLDepNo.Caption = ""
      'Me.LBLDepNetAmount.Caption = ""
      'Me.LBLDepBalance.Caption = ""
      'Me.IMGNew.Visible = True
      'Me.IMGCancel.Visible = False
      'Me.IMGConfirm.Visible = False
      ''Me.IMGBillStatus.Visible = False
      'Me.ListViewItemBackOrder.ListItems.Clear
   'End If
'End If
End Sub

Private Sub TextDocNo_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vDocNo As String
Dim vCheckExist As Integer
Dim i As Integer


'If KeyAscii = 13 Then
 '  If Me.TextDocno.Text <> "" Then
  '    Me.ListViewItemBackOrder.ListItems.Clear
   ''   vDocNo = Me.TextDocno.Text
     
   '   vQuery = "exec dbo.USP_BC_SearchCancelBackOrder '" & vDocNo & "' "
    '   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     '     vCheckExist = vRecordset.RecordCount
      '      Me.LBLCancelNo.Caption = vRecordset.Fields("docno").Value
       '     Me.LBLDocDate.Caption = vRecordset.Fields("backorderdate").Value
        '    Me.LBLARCode.Caption = vRecordset.Fields("arcode").Value
         '   Me.LBLARName.Caption = vRecordset.Fields("arname").Value
          '  Me.LBLSaleName.Caption = vRecordset.Fields("salename").Value
            ''Me.LBLDepNo.Caption = vRecordset.Fields("depositno").Value
           ' Me.LBLDepNetAmount.Caption = Format(vRecordset.Fields("depositamount").Value, "##,##0.00")
            'Me.LBLDepBalance.Caption = Format(vRecordset.Fields("depositamountbalance").Value, "##,##0.00")
            'Me.CMDCheckDocNo.Enabled = True
       'End If
       'vRecordset.Close
       
       
      'If vCheckExist = 0 Then
       '  vQuery = "exec dbo.USP_BO_SearchBackOrderMaster '" & vDocNo & "' "
        ' If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         ''   Me.LBLDocDate.Caption = vRecordset.Fields("docdate").Value
         '   Me.LBLARCode.Caption = vRecordset.Fields("arcode").Value
           '' Me.LBLARName.Caption = vRecordset.Fields("arname").Value
          '  Me.LBLSaleName.Caption = vRecordset.Fields("salename").Value
           ' Me.LBLDepNo.Caption = vRecordset.Fields("depositno").Value
            'Me.LBLDepNetAmount.Caption = Format(vRecordset.Fields("netamount").Value, "##,##0.00")
            'Me.LBLDepBalance.Caption = Format(vRecordset.Fields("billbalance").Value, "##,##0.00")
            'Me.CMDCheckDocNo.Enabled = True
         'Else
          '  MsgBox "ไม่พบข้อมูลของใบ Back Order เลขที่ " & vDocNo & "  กรุณาตรวจสอบ", vbCritical, "Send Error Message"
         'End If
         'vRecordset.Close
      'End If
      
      'Me.TextDocno.Text = UCase(Me.TextDocno.Text)
   'End If
'End If
End Sub

Private Sub Timer1_Timer()
'Me.PB101.Max = 10
'Me.PB101.Min = 0
'vMemPercentProcess = vMemPercentProcess + 1
'If vMemPercentProcess > 10 Then
'vMemPercentProcess = 0
'End If
'Me.PB101.Value = vMemPercentProcess
End Sub

Private Sub Timer101_Timer()
'If Me.CMDStatusReceive.Visible = True Then
 '  Me.CMDStatusReceive.Visible = False
'ElseIf Me.CMDStatusReceive.Visible = False Then
 '  Me.CMDStatusReceive.Visible = True
'End If
End Sub
