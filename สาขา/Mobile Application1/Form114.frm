VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form114 
   Caption         =   "ตรวจนับสินค้า ตามระบบ Cycle-Count"
   ClientHeight    =   8790
   ClientLeft      =   2310
   ClientTop       =   450
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form114.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   14385
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBReport 
      Height          =   8745
      Left            =   0
      Picture         =   "Form114.frx":9673
      ScaleHeight     =   8685
      ScaleWidth      =   14310
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   14370
      Begin Crystal.CrystalReport Crystal104 
         Left            =   630
         Top             =   7380
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
      Begin VB.ComboBox CMBItemBrand 
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
         Height          =   315
         Left            =   7875
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   4725
         Width           =   5100
      End
      Begin VB.OptionButton OPTItemBrand 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ตามยี่ห้อ"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5670
         TabIndex        =   47
         Top             =   4725
         Width           =   1005
      End
      Begin VB.OptionButton OPTItemAll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ทั้งหมด"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5670
         TabIndex        =   46
         Top             =   4365
         Width           =   1005
      End
      Begin VB.CommandButton CMDPrintItemNotCount 
         BackColor       =   &H00C0C0C0&
         Caption         =   "รายงานสินค้าไม่ได้นับสต๊อก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4365
         Width           =   2580
      End
      Begin VB.TextBox TXTItemCode 
         Appearance      =   0  'Flat
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
         Left            =   6795
         TabIndex        =   40
         Top             =   3825
         Width           =   2445
      End
      Begin VB.TextBox TXTShelfID 
         Appearance      =   0  'Flat
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
         Left            =   6795
         TabIndex        =   39
         Top             =   3465
         Width           =   1320
      End
      Begin VB.TextBox TXTRow 
         Appearance      =   0  'Flat
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
         Left            =   6795
         TabIndex        =   38
         Top             =   3105
         Width           =   1320
      End
      Begin VB.CommandButton CMDItemMultiShelf 
         BackColor       =   &H00C0C0C0&
         Caption         =   "รายงานสินค้าหลายที่เก็บ"
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3015
         Width           =   2580
      End
      Begin VB.ComboBox CMBReportType1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6795
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2655
         Width           =   2445
      End
      Begin VB.OptionButton OPTBrand 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ตาม ยี่ห้อสินค้า"
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
         Height          =   330
         Left            =   5670
         TabIndex        =   30
         Top             =   2115
         Width           =   3570
      End
      Begin VB.OptionButton OPTDocNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ตาม เลขที่เอกสาร"
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
         Height          =   330
         Left            =   5670
         TabIndex        =   29
         Top             =   1755
         Value           =   -1  'True
         Width           =   3570
      End
      Begin VB.CommandButton CMDReportClose 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5895
         Width           =   2580
      End
      Begin VB.CommandButton CMDPrintItemNotConfirm 
         BackColor       =   &H00C0C0C0&
         Caption         =   "รายงาน สินค้าไม่รับยอด"
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1755
         Width           =   2580
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกยี่ห้อ :"
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
         Height          =   330
         Left            =   6840
         TabIndex        =   48
         Top             =   4770
         Width           =   1005
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกประเภท :"
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
         Left            =   5670
         TabIndex        =   44
         Top             =   2700
         Width           =   1365
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสที่เก็บ :"
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
         Left            =   5670
         TabIndex        =   43
         Top             =   3510
         Width           =   1230
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสสินค้า :"
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
         Left            =   5670
         TabIndex        =   42
         Top             =   3870
         Width           =   1275
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "รหัส Row :"
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
         Left            =   5670
         TabIndex        =   41
         Top             =   3150
         Width           =   1230
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "รายงาน ต่าง ๆ "
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
         Height          =   465
         Left            =   2880
         TabIndex        =   24
         Top             =   1215
         Width           =   2265
      End
   End
   Begin VB.PictureBox PBTransferNo 
      Height          =   8745
      Left            =   0
      Picture         =   "Form114.frx":12CE6
      ScaleHeight     =   8685
      ScaleWidth      =   14310
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   14370
      Begin VB.ComboBox CMBDocNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2970
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2250
         Width           =   2580
      End
      Begin VB.CommandButton CMDTransferClose 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6660
         Width           =   2580
      End
      Begin VB.CommandButton CMDTransfer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "บันทึกโอนสินค้าจากการรับผล"
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
         Height          =   690
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5715
         Width           =   2580
      End
      Begin VB.CommandButton CMDCheckCountItem 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ตรวจสอบการรับผล"
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
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4770
         Width           =   2580
      End
      Begin VB.CommandButton CMDPrintInspect 
         BackColor       =   &H00C0C0C0&
         Caption         =   "พิมพ์ใบสรุปผลการตรวจนับ"
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
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3825
         Width           =   2580
      End
      Begin VB.CommandButton CMDAcceptDiffQty 
         BackColor       =   &H00C0C0C0&
         Caption         =   "บันทึกผลต่างใบสรุปผล"
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
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2880
         Width           =   2580
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกเลขที่ใบสรุปผลการตรวจนับ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   0
         TabIndex        =   23
         Top             =   2295
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "การรับผล ยอดตรวจนับสินค้า"
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
         Height          =   375
         Left            =   2970
         TabIndex        =   21
         Top             =   1485
         Width           =   4875
      End
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   4725
      Top             =   225
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
   Begin MSComCtl2.DTPicker DTPDocDate 
      Height          =   375
      Left            =   2730
      TabIndex        =   12
      Top             =   6570
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16580609
      CurrentDate     =   41236
   End
   Begin VB.CommandButton CMDSearchRefresh 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ฟื้นฟูข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10665
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1170
      Width           =   1395
   End
   Begin VB.CommandButton CMDClearUserID 
      BackColor       =   &H00C0C0C0&
      Caption         =   "เคลียร์ผู้ใช้งาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12645
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1170
      Width           =   1395
   End
   Begin VB.TextBox TXTDocNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11550
      TabIndex        =   6
      Top             =   6570
      Width           =   2475
   End
   Begin VB.CommandButton CMDPrintDocNo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "พิมพ์ใบตรวจนับ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   12270
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7395
      Width           =   1755
   End
   Begin VB.CommandButton CMDPrintItemSlotTag 
      BackColor       =   &H00C0C0C0&
      Caption         =   "พิมพ์SlotTag"
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
      Height          =   600
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7425
      Width           =   1755
   End
   Begin VB.CommandButton CMDPrintItemLabel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "พิมพ์ป้ายติดสินค้า"
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
      Height          =   600
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7425
      Width           =   1755
   End
   Begin VB.CommandButton CMDAddItemToShelf 
      BackColor       =   &H00C0C0C0&
      Caption         =   "บันทึกสินค้าเข้าที่เก็บ"
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
      Height          =   600
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7425
      Width           =   1755
   End
   Begin VB.CommandButton CMDSaveData 
      BackColor       =   &H00C0C0C0&
      Caption         =   "บันทึกรวมเอกสาร"
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
      Height          =   600
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7425
      Width           =   1755
   End
   Begin MSComctlLib.ListView ListViewDocNo 
      Height          =   4755
      Left            =   300
      TabIndex        =   0
      Top             =   1710
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   8387
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่เอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "คำอธิบาย"
         Object.Width           =   7938
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ผู้สร้างเอกสาร"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "วันที่ทำเอกสาร"
         Object.Width           =   3528
      EndProperty
   End
   Begin Crystal.CrystalReport Crystal102 
      Left            =   855
      Top             =   8190
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
   Begin VB.CommandButton CMDApprove 
      BackColor       =   &H00C0C0C0&
      Caption         =   "รับผลการตรวจนับ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7425
      Width           =   1755
   End
   Begin VB.CommandButton CMDReport 
      BackColor       =   &H00C0C0C0&
      Caption         =   "รายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7425
      Width           =   1755
   End
   Begin VB.OptionButton OPTAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ค้นทั้งหมด"
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
      Height          =   330
      Left            =   4725
      TabIndex        =   31
      Top             =   810
      Value           =   -1  'True
      Width           =   1860
   End
   Begin VB.OptionButton OPTDocDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ระหว่างวันที่เอกสาร"
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
      Height          =   330
      Left            =   4725
      TabIndex        =   32
      Top             =   1260
      Width           =   1860
   End
   Begin MSComCtl2.DTPicker DTPSearchDate2 
      Height          =   330
      Left            =   8865
      TabIndex        =   34
      Top             =   1260
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   582
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
      Format          =   16580609
      CurrentDate     =   41251
   End
   Begin MSComCtl2.DTPicker DTPSearchDate1 
      Height          =   330
      Left            =   6660
      TabIndex        =   33
      Top             =   1260
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   582
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
      Format          =   16580609
      CurrentDate     =   41250
   End
   Begin Crystal.CrystalReport Crystal103 
      Left            =   1710
      Top             =   8190
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ก่อนบันทึกเอกสาร กรุณาตรวจสอบวันที่เทียบยอดคงเหลือด้วย"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   330
      TabIndex        =   13
      Top             =   7065
      Width           =   7215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ให้เทียบยอดคงเหลือ ณ วันที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   300
      TabIndex        =   11
      Top             =   6600
      Width           =   2475
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกเอกสาร ใบตรวจนับ เพื่อทำใบสรุปผลการตรวจนับ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   315
      TabIndex        =   8
      Top             =   1305
      Width           =   4515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "พิมพ์เอกสารเลขที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9870
      TabIndex        =   7
      Top             =   6600
      Width           =   1605
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "ถึง"
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
      Left            =   8370
      TabIndex        =   35
      Top             =   1260
      Width           =   465
   End
End
Attribute VB_Name = "Form114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vQuery As String


Private Sub CMBReportType1_Click()
On Error Resume Next

If Me.CMBReportType1.ListIndex = 0 Then
Me.TXTRow.Enabled = True
Me.TXTShelfID.Enabled = False
Me.TXTItemCode.Enabled = False
Me.TXTRow.SetFocus
ElseIf Me.CMBReportType1.ListIndex = 1 Then
Me.TXTRow.Enabled = False
Me.TXTShelfID.Enabled = True
Me.TXTItemCode.Enabled = False
Me.TXTShelfID.SetFocus
ElseIf Me.CMBReportType1.ListIndex = 2 Then
Me.TXTRow.Enabled = False
Me.TXTShelfID.Enabled = False
Me.TXTItemCode.Enabled = True
Me.TXTItemCode.SetFocus
End If
End Sub

Private Sub CMDAcceptDiffQty_Click()
Dim vDocNo As String
Dim vCheckDocDate As Date
Dim vDateNow As Date
Dim vRecordset As New Recordset


On Error Resume Next

If Me.CMBDocNo.Text <> "" Then
  vDocNo = Me.CMBDocNo.Text
  vDateNow = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
  
  
  vQuery = "select docdate from dbo.bcstkinspect where docno = '" & vDocNo & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckDocDate = vRecordset.Fields("docdate").Value
End If
vRecordset.Close

  If vCheckDocDate <> vDateNow Then
    vQuery = "exec dbo.USP_IC_DiffInspect '" & vDocNo & "' "
    gConnection.Execute vQuery
    
    vQuery = "update dbo.bcstkinspect set printcount = 1 where docno = '" & vDocNo & "' "
    gConnection.Execute vQuery
        
  Else
    vQuery = "exec dbo.USP_IC_DiffInspect_Online '" & vDocNo & "' "
    gConnection.Execute vQuery
    
    vQuery = "update dbo.bcstkinspect set printcount = 1 where docno = '" & vDocNo & "' "
    gConnection.Execute vQuery
  End If
  
  MsgBox "บันทึกผลต่างเรียบร้อย พิมพ์เอกสารตรวจนับได้เลย", vbInformation, "Send Information Message"
  
  Call CMDPrintInspect_Click
  
  Me.CMBDocNo.SetFocus
  
Else
  MsgBox "กรุณาเลือกเอกสารที่จะเทียบยอดผลต่าง", vbInformation, "Send Information Message"
  Me.CMBDocNo.SetFocus
End If
End Sub

Private Sub CMDAddItemToShelf_Click()
MsgBox "ยังไม่เปิดใช้งาน", vbInformation, "Message Information Message"
End Sub

Private Sub CMDApprove_Click()
Dim vRecordset As New ADODB.Recordset
Dim i As Integer

Me.CMBDocNo.Clear
vQuery = "exec dbo.USP_MB_SearchInspectToTransfer"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     vRecordset.MoveFirst
      i = 1
      While Not vRecordset.EOF
      Me.CMBDocNo.AddItem (Trim(vRecordset.Fields("docno").Value))
      vRecordset.MoveNext
      i = i + 1
      Wend
End If
vRecordset.Close


Me.PBTransferNo.Visible = True
Me.CMBDocNo.SetFocus
End Sub

Public Sub SearchInspectToTransfer()
Dim vRecordset As New ADODB.Recordset
Dim i As Integer

Me.CMBDocNo.Clear
vQuery = "exec dbo.USP_MB_SearchInspectToTransfer"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     vRecordset.MoveFirst
      i = 1
      While Not vRecordset.EOF
      Me.CMBDocNo.AddItem (Trim(vRecordset.Fields("docno").Value))
      vRecordset.MoveNext
      i = i + 1
      Wend
End If
vRecordset.Close

Me.CMBDocNo.SetFocus
End Sub

Private Sub CMDCheckCountItem_Click()
Me.CMDTransfer.Enabled = True
End Sub

Private Sub CMDItemMultiShelf_Click()
Dim vRecordset As New ADODB.Recordset
Dim vRepID As Integer
Dim vRepType, vReportName, vDocNo As String
Dim vType As Integer
Dim vSearch As String

On Error GoTo ErrDescription

If Me.CMBReportType1.ListIndex = 0 Then
vRepID = 479
vType = 0
vSearch = Me.TXTRow.Text
End If

If Me.CMBReportType1.ListIndex = 1 Then
vRepID = 479
vType = 1
vSearch = Me.TXTShelfID.Text
End If

If Me.CMBReportType1.ListIndex = 2 Then
vRepID = 480
vType = 2
vSearch = Me.TXTItemCode.Text
End If

vRepType = "IS"

vQuery = "select reportname from bcnp.dbo.bcreportname where  repid = " & vRepID & " and reptype = '" & vRepType & "'  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal103
.ReportFileName = Trim(vReportName & ".rpt")
.ParameterFields(0) = "@vType;" & vType & ";true"
.ParameterFields(1) = "@vSearch;" & vSearch & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDPrintDocNo_Click()
Dim vRecordset As New ADODB.Recordset
Dim vRepID As Integer
Dim vRepType, vReportName, vDocNo As String

On Error GoTo ErrDescription

If Me.TXTDocNo.Text <> "" Then
    vRepID = 213
    vRepType = "IV"
    vDocNo = Trim(Me.TXTDocNo.Text)
    vQuery = "select reportname from bcnp.dbo.bcreportname where  repid = " & vRepID & " and reptype = '" & vRepType & "'  "
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
Else
MsgBox "กรุณากรอก เลขที่ใบตรวจนับที่ต้องการพิมพ์", vbInformation, "Send Information Message"
Me.TXTDocNo.SetFocus
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDPrintInspect_Click()
Dim vRecordset As New ADODB.Recordset
Dim vRepID As Integer
Dim vRepType, vReportName, vDocNo As String

On Error GoTo ErrDescription

If Me.CMBDocNo.Text <> "" Then
    vRepID = 213
    vRepType = "IV"
    vDocNo = Trim(Me.CMBDocNo.Text)
    vQuery = "select reportname from bcnp.dbo.bcreportname where  repid = " & vRepID & " and reptype = '" & vRepType & "'  "
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
Else
MsgBox "กรุณากรอก เลขที่ใบตรวจนับที่ต้องการพิมพ์", vbInformation, "Send Information Message"
Me.CMBDocNo.SetFocus
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDPrintItemLabel_Click()
MsgBox "ยังไม่เปิดใช้งาน", vbInformation, "Message Information Message"
End Sub

Private Sub CMDPrintItemNotConfirm_Click()
Dim vRecordset As New ADODB.Recordset
Dim vRepID As Integer
Dim vRepType, vReportName, vDocNo As String

On Error GoTo ErrDescription

If Me.OPTDocNo.Value = True Then
vRepID = 477
End If

If Me.OPTBrand.Value = True Then
vRepID = 478
End If

vRepType = "IS"

vQuery = "select reportname from bcnp.dbo.bcreportname where  repid = " & vRepID & " and reptype = '" & vRepType & "'  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal102
.ReportFileName = Trim(vReportName & ".rpt")
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDPrintItemNotCount_Click()
Dim vRecordset As New ADODB.Recordset
Dim vRepID As Integer
Dim vRepType, vReportName, vDocNo As String
Dim vType As Integer
Dim vProfit As String
Dim vBrandCode As String

On Error GoTo ErrDescription

vProfit = "S02"
vRepID = 481
vRepType = "IS"

If Me.OPTItemAll.Value = True Then
  vType = 0
ElseIf Me.OPTItemBrand.Value = True Then
  vType = 1

  If Me.CMBItemBrand.Text <> "" Then
    vBrandCode = Left(Me.CMBItemBrand.Text, InStr(Me.CMBItemBrand.Text, "/") - 1)
    Else
    vBrandCode = ""
  End If
Else
  vType = 0
End If


vQuery = "select reportname from bcnp.dbo.bcreportname where  repid = " & vRepID & " and reptype = '" & vRepType & "'  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal104
.ReportFileName = Trim(vReportName & ".rpt")
.ParameterFields(0) = "@vProfit;" & vProfit & ";true"
.ParameterFields(1) = "@vType;" & vType & ";true"
.ParameterFields(2) = "@vBrandCode;" & vBrandCode & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDPrintItemSlotTag_Click()
MsgBox "ยังไม่เปิดใช้งาน", vbInformation, "Message Information Message"
End Sub

Private Sub CMDReport_Click()
Call SearchItemBrand
Me.CMBReportType1.Clear
Me.CMBReportType1.AddItem ("1.ตาม Row")
Me.CMBReportType1.AddItem ("2.ตาม ชั้นเก็บ")
Me.CMBReportType1.AddItem ("3.ตาม รหัสสินค้า")
Me.CMBReportType1.ListIndex = 0

Me.PBReport.Visible = True
End Sub

Private Sub CMDReportClose_Click()
Me.PBReport.Visible = False
Me.ListViewDocNo.SetFocus
End Sub

Private Sub CMDSaveData_Click()
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vDocNo As String
Dim vDocDate As String
Dim i As Integer
Dim n As Integer
Dim vCountSelect As Integer
Dim vRepID As Integer
Dim vRepType, vReportName As String
Dim vCDocDate As String

vCountSelect = 0

For i = 1 To Me.ListViewDocNo.ListItems.Count
If Me.ListViewDocNo.ListItems(i).Checked = True Then
vCountSelect = vCountSelect + 1
End If
Next i

If vCountSelect = 0 Then
MsgBox "กรุณาเลือกเอกสารที่ต้องการสรุปผลรวมยอดตรวจนับ", vbInformation, "Send Information Message"
Me.ListViewDocNo.SetFocus
Me.ListViewDocNo.ListItems(0).Selected = True
Exit Sub
End If


If vCountSelect > 0 Then

For i = 1 To Me.ListViewDocNo.ListItems.Count
If Me.ListViewDocNo.ListItems(i).Checked = True Then

 vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
 vDocDate = Me.ListViewDocNo.ListItems(i).SubItems(2)

vQuery = "exec dbo.USP_NP_InsertSelectInspectTemp '" & vDocNo & "','" & vDocDate & "','" & vUserID & "'  "
gConnection.Execute vQuery

End If
Next i

vCDocDate = Day(Me.DTPDocDate.Value) & "/" & Month(Me.DTPDocDate.Value) & "/" & Year(Me.DTPDocDate.Value)

vQuery = "exec dbo.USP_NP_GenInspectAuto 'S02','" & vUserID & "','" & vCDocDate & "' "
gConnection.Execute vQuery

vQuery = "select top 1 isnull(docno,'') as docno  from dbo.bcstkinspect where docno like 'S02-CKT%' and creatorcode = '" & vUserID & "' order by createdatetime desc"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vDocNo = Trim(vRecordset.Fields("docno").Value)
End If
vRecordset.Close

Me.CMDSaveData.Enabled = False

MsgBox "บันทึกข้อมูลเรียบร้อยแล้ว ได้เลขที่เอกสารใบตรวจนับรวมเลขที่ " & vDocNo & " ", vbInformation, "Send Information Message"

Call CMDSearchRefresh_Click

vRepID = 213
vRepType = "IV"
vQuery = "select reportname from bcnp.dbo.bcreportname where  repid = " & vRepID & " and reptype = '" & vRepType & "'  "
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
End Sub

Private Sub CMDSearchRefresh_Click()
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem

Dim vType As Integer
Dim vDocDate1 As String
Dim vDocDate2 As String


On Error GoTo ErrDescription

Me.ListViewDocNo.ListItems.Clear

If Me.OPTAll.Value = True Then
vType = 0
Else
vType = 1
End If

vDocDate1 = Day(Me.DTPSearchDate1.Value) & "/" & Month(Me.DTPSearchDate1.Value) & "/" & Year(Me.DTPSearchDate1.Value)
vDocDate2 = Day(Me.DTPSearchDate2.Value) & "/" & Month(Me.DTPSearchDate2.Value) & "/" & Year(Me.DTPSearchDate2.Value)

vQuery = "exec dbo.USP_NP_SearchInspectNotAdjust 'S02'," & vType & ",'" & vDocDate1 & "','" & vDocDate2 & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then

     vRecordset.MoveFirst
      i = 1
      While Not vRecordset.EOF
      Set vListItem = Me.ListViewDocNo.ListItems.Add(, , i)
      vListItem.SubItems(1) = vRecordset.Fields("docno").Value
      vListItem.SubItems(2) = vRecordset.Fields("docdate").Value
      vListItem.SubItems(3) = vRecordset.Fields("mydescription").Value
      vListItem.SubItems(4) = vRecordset.Fields("creatorcode").Value
      vListItem.SubItems(5) = vRecordset.Fields("createdatetime").Value
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

Private Sub CMDTransfer_Click()
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vProfit As String
Dim vInSpectNo As String


If Me.CMBDocNo.Text <> "" Then
  vProfit = "S02"
  vDocNo = Me.CMBDocNo.Text
  
  vQuery = "exec dbo.USP_IC_TransferFromInspect '" & vProfit & "','" & vDocNo & "' "
  gConnection.Execute vQuery
  
  vQuery = "exec dbo.USP_IC_TransferFromInspect2 '" & vUserID & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vInSpectNo = Trim(vRecordset.Fields("docno").Value)
  End If
  vRecordset.Close
  
  Call SearchInspectToTransfer
  
  Me.CMDTransfer.Enabled = False
  
  MsgBox "บันทึกโอนสินค้าจากการตรวจนับเรียบร้อยแล้ว ได้เลขที่โอนเลขที่ " & vInSpectNo & " กรุณาตรวจสอบ", vbInformation, "Send Informat Message"
  Me.CMBDocNo.SetFocus
  
Else
  MsgBox "กรุณาเลือกเอกสารที่จะเทียบโอนสินค้า", vbInformation, "Send Information Message"
  Me.CMBDocNo.SetFocus
End If
End Sub

Private Sub CMDTransferClose_Click()
Me.PBTransferNo.Visible = False
Me.ListViewDocNo.SetFocus
End Sub

Private Sub DTPDocDate_Change()
Me.CMDSaveData.Enabled = True
End Sub

Private Sub DTPDocDate_Click()
Me.CMDSaveData.Enabled = True
End Sub

Private Sub DTPDocDate_GotFocus()
Me.CMDSaveData.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next

Me.DTPDocDate.Value = Now
Me.DTPSearchDate1.Value = Now
Me.DTPSearchDate2.Value = Now
Call CMDSearchRefresh_Click
Me.DTPDocDate.SetFocus
End Sub

Private Sub ListViewDocNo_KeyPress(KeyAscii As Integer)
Dim i As Integer

If KeyAscii = 32 Then

For i = 1 To Me.ListViewDocNo.ListItems.Count
If Me.ListViewDocNo.ListItems(i).Selected = True Then
  If Me.ListViewDocNo.ListItems(i).Checked = True Then
    Me.ListViewDocNo.ListItems(i).Checked = False
    Else
    Me.ListViewDocNo.ListItems(i).Checked = True
  End If
End If
Next i
End If
End Sub

Public Sub SearchItemBrand()
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

Me.CMBItemBrand.Clear
vQuery = "exec dbo.USP_PS_BrandList"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
      Me.CMBItemBrand.AddItem (Trim(vRecordset.Fields("brandname").Value))
      vRecordset.MoveNext
    Wend
    End If
vRecordset.Close
End Sub

Private Sub OPTBrand_Click()
Me.CMBItemBrand.Enabled = False
End Sub

Private Sub OPTDocNo_Click()
Me.CMBItemBrand.Enabled = False
End Sub

Private Sub OPTItemAll_Click()
Me.CMBItemBrand.Enabled = False
End Sub

Private Sub OPTItemBrand_Click()
Me.CMBItemBrand.Enabled = True
Me.CMBItemBrand.SetFocus
End Sub
