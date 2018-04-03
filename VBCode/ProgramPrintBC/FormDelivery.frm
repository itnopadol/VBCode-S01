VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FormDelivery 
   Caption         =   "ทำใบขนส่งสินค้า"
   ClientHeight    =   8985
   ClientLeft      =   1680
   ClientTop       =   795
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormDelivery.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   7335
      Top             =   7605
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CMD010 
      Caption         =   "สรุปผลการเก็บเงิน"
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
      Height          =   330
      Left            =   5310
      TabIndex        =   27
      Top             =   7245
      Width           =   1860
   End
   Begin VB.CommandButton CMD109 
      Height          =   330
      Left            =   2340
      Picture         =   "FormDelivery.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "ออกหน้าทำใบขนส่ง"
      Top             =   7245
      Width           =   330
   End
   Begin VB.CheckBox Check101 
      BackColor       =   &H80000009&
      Caption         =   "การขนส่งเสร็จสมบูรณ์"
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
      Left            =   3060
      TabIndex        =   26
      Top             =   7245
      Width           =   2085
   End
   Begin VB.CommandButton CMD105 
      Height          =   330
      Left            =   1935
      Picture         =   "FormDelivery.frx":7665
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "ยกเลิก ใบขนส่ง"
      Top             =   7245
      Width           =   330
   End
   Begin VB.CommandButton CMD104 
      Height          =   330
      Left            =   1530
      Picture         =   "FormDelivery.frx":9803
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "พิมพ์ ใบขนส่ง"
      Top             =   7245
      Width           =   330
   End
   Begin VB.CommandButton CMD103 
      Height          =   330
      Left            =   1125
      Picture         =   "FormDelivery.frx":9B45
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "ค้นหา ใบขนส่ง"
      Top             =   7245
      Width           =   330
   End
   Begin VB.CommandButton CMD102 
      Height          =   330
      Left            =   720
      Picture         =   "FormDelivery.frx":9F12
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "บันทึกข้อมูล เอกสารสร้างใหม่และปรับปรุง"
      Top             =   7245
      Width           =   330
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   315
      Picture         =   "FormDelivery.frx":A239
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "เคลียร์เอกสาร"
      Top             =   7245
      Width           =   330
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2535
      Left            =   315
      TabIndex        =   1
      Top             =   1080
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   4471
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
      BackColor       =   16711382
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่อ้างอิง"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ชื่อสินค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "จำนวน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "จำนวนที่ส่งได้"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "หน่วยนับ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "คำอธิบาย"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "QueueID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "InvoiceNo"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.CommandButton CMD106 
      Caption         =   "เลือกรายการจัดคิว"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2565
      TabIndex        =   0
      Top             =   405
      Width           =   1545
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3390
      Left            =   315
      TabIndex        =   2
      Top             =   3735
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   5980
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "ข้อมูลขนส่ง"
      TabPicture(0)   =   "FormDelivery.frx":A61E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text101"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text102"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text108"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text103"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text106"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text107"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "DTPicker101"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "DTPicker102"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "DTPicker103"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "MaskEdBox101"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "MaskEdBox102"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "CMD107"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "CMD108"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text105"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text104"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "ข้อมูลพนักงาน"
      TabPicture(1)   =   "FormDelivery.frx":A63A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PICEmpWage"
      Tab(1).Control(1)=   "CMD201"
      Tab(1).Control(2)=   "ListView102"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "สรุปผลการขนส่ง"
      TabPicture(2)   =   "FormDelivery.frx":A656
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame102"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox PICEmpWage 
         BackColor       =   &H80000009&
         Height          =   2490
         Left            =   -74640
         ScaleHeight     =   2430
         ScaleWidth      =   9495
         TabIndex        =   48
         Top             =   585
         Visible         =   0   'False
         Width           =   9555
         Begin VB.CommandButton CMDEmpWage 
            Caption         =   "ตกลง"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3420
            TabIndex        =   56
            Top             =   1935
            Width           =   915
         End
         Begin VB.TextBox TXTWage 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2475
            TabIndex        =   55
            Top             =   1485
            Width           =   1860
         End
         Begin VB.Label LBLEmpName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2475
            TabIndex        =   54
            Top             =   1035
            Width           =   6405
         End
         Begin VB.Label LBLEmpID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2475
            TabIndex        =   53
            Top             =   585
            Width           =   1860
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าเที่ยว :"
            Height          =   285
            Left            =   765
            TabIndex        =   52
            Top             =   1485
            Width           =   1590
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ชื่อ-สกุล :"
            Height          =   330
            Left            =   765
            TabIndex        =   51
            Top             =   1035
            Width           =   1590
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "รหัสพนักงาน :"
            Height          =   330
            Left            =   810
            TabIndex        =   50
            Top             =   585
            Width           =   1545
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "คิดค่าเที่ยวพนักงาน"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   49
            Top             =   135
            Width           =   1815
         End
      End
      Begin VB.Frame Frame102 
         Caption         =   "ผลการขนส่ง"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   -74505
         TabIndex        =   41
         Top             =   675
         Width           =   9330
         Begin VB.OptionButton Option201 
            Caption         =   "1.สมบูรณ์"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2025
            TabIndex        =   47
            Top             =   135
            Width           =   2400
         End
         Begin VB.OptionButton Option202 
            Caption         =   "2.หา Site งานไม่เจอ"
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
            Left            =   2025
            TabIndex        =   46
            Top             =   585
            Width           =   2400
         End
         Begin VB.OptionButton Option203 
            Caption         =   "3.ส่งสินค้าผิด"
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
            Left            =   2025
            TabIndex        =   45
            Top             =   1035
            Width           =   1995
         End
         Begin VB.OptionButton Option204 
            Caption         =   "4.สินค้าชำรุด"
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
            Left            =   5850
            TabIndex        =   44
            Top             =   180
            Width           =   1725
         End
         Begin VB.OptionButton Option205 
            Caption         =   "5.ส่งมอบสินค้าบางส่วน"
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
            Left            =   5850
            TabIndex        =   43
            Top             =   585
            Width           =   2355
         End
         Begin VB.OptionButton Option206 
            Caption         =   "6.อื่น ๆ "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5850
            TabIndex        =   42
            Top             =   1035
            Width           =   1815
         End
      End
      Begin VB.TextBox Text104 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   6750
         TabIndex        =   13
         Top             =   855
         Width           =   1905
      End
      Begin VB.TextBox Text105 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   6750
         TabIndex        =   14
         Top             =   1260
         Width           =   2535
      End
      Begin VB.CommandButton CMD108 
         Height          =   330
         Left            =   8685
         Picture         =   "FormDelivery.frx":A672
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "ค้นหา เลขที่รถขนส่ง"
         Top             =   450
         Width           =   330
      End
      Begin VB.CommandButton CMD107 
         Height          =   330
         Left            =   4050
         Picture         =   "FormDelivery.frx":AA3F
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "สร้างเลขที่ใบขนส่ง"
         Top             =   855
         Width           =   330
      End
      Begin MSMask.MaskEdBox MaskEdBox102 
         Height          =   330
         Left            =   1980
         TabIndex        =   10
         Top             =   2880
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox101 
         Height          =   330
         Left            =   1980
         TabIndex        =   8
         Top             =   2070
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker DTPicker103 
         Height          =   330
         Left            =   1980
         TabIndex        =   9
         Top             =   2475
         Width           =   2040
         _ExtentX        =   3598
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
         Format          =   69730305
         CurrentDate     =   38695
      End
      Begin MSComCtl2.DTPicker DTPicker102 
         Height          =   330
         Left            =   1980
         TabIndex        =   7
         Top             =   1665
         Width           =   2040
         _ExtentX        =   3598
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
         Format          =   69730305
         CurrentDate     =   38695
      End
      Begin MSComCtl2.DTPicker DTPicker101 
         Height          =   330
         Left            =   1980
         TabIndex        =   6
         Top             =   1260
         Width           =   2040
         _ExtentX        =   3598
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
         Format          =   69730305
         CurrentDate     =   38695
      End
      Begin VB.TextBox Text107 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   330
         Left            =   6750
         TabIndex        =   16
         Top             =   2070
         Width           =   2535
      End
      Begin VB.TextBox Text106 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   6750
         TabIndex        =   15
         Top             =   1665
         Width           =   2535
      End
      Begin VB.TextBox Text103 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   330
         Left            =   6750
         TabIndex        =   11
         Top             =   450
         Width           =   1905
      End
      Begin VB.TextBox Text108 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   6750
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   2475
         Width           =   3120
      End
      Begin VB.TextBox Text102 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   330
         Left            =   1980
         TabIndex        =   4
         Top             =   855
         Width           =   1995
      End
      Begin VB.CommandButton CMD201 
         Caption         =   "เลือกพนักงาน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74595
         TabIndex        =   18
         Top             =   585
         Width           =   1320
      End
      Begin MSComctlLib.ListView ListView102 
         Height          =   1950
         Left            =   -74595
         TabIndex        =   19
         Top             =   1080
         Width           =   9510
         _ExtentX        =   16775
         _ExtentY        =   3440
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "รหัสพนักงาน"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ชื่อพนักงาน"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ตำแหน่ง"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "ค่าเที่ยว"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.TextBox Text101 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   330
         Left            =   1980
         TabIndex        =   3
         Top             =   450
         Width           =   1635
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "หมายเลขรถ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5130
         TabIndex        =   40
         Top             =   855
         Width           =   1500
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "หมายเหตุ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5355
         TabIndex        =   39
         Top             =   2475
         Width           =   1275
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "กม. สิ้นสุด :"
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
         Left            =   5535
         TabIndex        =   38
         Top             =   2070
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "กม. เริ่มต้น :"
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
         Left            =   5220
         TabIndex        =   37
         Top             =   1665
         Width           =   1410
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "ทะเบียนรถ :"
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
         Left            =   5220
         TabIndex        =   36
         Top             =   1260
         Width           =   1410
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "ID รถขนส่ง :"
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
         Left            =   5355
         TabIndex        =   35
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "เวลากลับ :"
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
         Left            =   900
         TabIndex        =   34
         Top             =   2880
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "วันที่กลับ :"
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
         Left            =   765
         TabIndex        =   33
         Top             =   2475
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "เวลาส่ง :"
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
         TabIndex        =   32
         Top             =   2070
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "วันที่ส่ง :"
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
         Left            =   1125
         TabIndex        =   31
         Top             =   1665
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   720
         TabIndex        =   30
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   675
         TabIndex        =   29
         Top             =   810
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ID :"
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
         Left            =   450
         TabIndex        =   28
         Top             =   495
         Width           =   1410
      End
   End
   Begin VB.Image Image102 
      Height          =   300
      Left            =   10035
      Picture         =   "FormDelivery.frx":AD83
      Top             =   180
      Width           =   570
   End
   Begin VB.Image Image101 
      Height          =   300
      Left            =   10035
      Picture         =   "FormDelivery.frx":B2BF
      Top             =   180
      Width           =   570
   End
End
Attribute VB_Name = "FormDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEmpWageID As Integer
Dim vCheckEmpWages As Boolean

Private Sub Check101_Click()
Dim i As Integer

On Error Resume Next

    If Check101.Value = 1 Then
        For i = 1 To ListView101.ListItems.Count
            ListView101.ListItems.Item(i).SubItems(5) = ListView101.ListItems.Item(i).SubItems(4)
        Next i
        Text107.Enabled = True
        MaskEdBox102.Enabled = True
        CMD010.Enabled = True
        SSTab1.TabEnabled(2) = True
    Else
        For i = 1 To ListView101.ListItems.Count
            ListView101.ListItems.Item(i).SubItems(5) = 0
        Next i
        Text107.Enabled = False
        MaskEdBox102.Enabled = False
        CMD010.Enabled = False
        SSTab1.TabEnabled(2) = False
    End If
End Sub

Private Sub CMD010_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDono As String
Dim vSendResultList As ListItem
Dim i As Integer

On Error GoTo ErrDescription

If Text101.Text <> "" And Check101.Value = 1 And Text107.Text <> "0" Then
FormSendResult.Show
FormSendResult.SetFocus
FormDelivery.Enabled = False
vDono = Trim(FormDelivery.Text102.Text)
FormSendResult.Text101.Text = vDono
vQuery = "exec dbo.USP_DO_UpdateSendResult '" & vDono & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    i = 1
    While Not vRecordset.EOF
        Set vSendResultList = FormSendResult.ListView101.ListItems.Add(, , i)
        vSendResultList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
        vSendResultList.SubItems(2) = Trim(vRecordset.Fields("hoardamount").Value)
        vSendResultList.SubItems(3) = Trim(vRecordset.Fields("hoardnet").Value)
        vSendResultList.SubItems(4) = Trim(vRecordset.Fields("hoardresult").Value)
        'vSendResultList.SubItems(5) = Trim(vRecordset.Fields("hoarddate").Value)
    vRecordset.MoveNext
    i = i + 1
    Wend
End If
vRecordset.Close
Else
    MsgBox "ต้องกรอกข้อมูลการส่งสินค้าขากลับให้เรียบร้อยก่อน ถึงจะเข้าไปสรุปการขนส่งได้", vbInformation, "Send Information"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD101_Click()
On Error GoTo ErrDescription

ListView101.ListItems.Clear
ListView102.ListItems.Clear
Check101.Value = 0
Text101.Text = ""
Text102.Text = ""
DTPicker101 = Now
DTPicker102 = Now
DTPicker103 = Now
MaskEdBox101.Mask = "##:##"
MaskEdBox102.Mask = "##:##"
Text103.Text = ""
Text104.Text = ""
Text105.Text = ""
Text106.Text = ""
Text107.Text = ""
Image101.Visible = True
Image102.Visible = False
vIsOpen2 = 0
Text106.Enabled = True
Text107.Enabled = False
MaskEdBox102.Enabled = False
FormDelivery.DTPicker102.Enabled = True
FormDelivery.MaskEdBox101.Enabled = True
Check101.Enabled = True
CMD106.Enabled = True
CMD201.Enabled = True
Option201.Value = True
Option202.Value = False
Option203.Value = False
Option204.Value = False
Option205.Value = False
Option206.Value = False
            
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As String
Dim vDocNo As String
Dim vDocdate As String
Dim vDateSend As Date
Dim vTimeSend As String
Dim vDateReturn As Date
Dim vTimeReturn As String
Dim vMeasureStart As Currency
Dim vMeasureStop As Currency
Dim vMydescription As String
Dim vIsCancel As String
Dim vVehicalID As Integer
Dim vIsReturn As Integer
Dim vCheckDeliveryID As String
Dim vReturn_Status As Integer
Dim vIsCompleteSave As Integer
Dim vDeliveryID As Integer
Dim vEmpBPlusID As Integer
Dim vEmpWage As Integer
Dim vIsCancelEmp As String
Dim i As Integer
Dim vQueueDetailID As Currency
Dim vReturn_StatusSub As Integer
Dim vIsCompleteSaveSub As Integer
Dim vQTY As Currency
Dim vConfirmQty As Currency
Dim vMydescriptionSub As String
Dim vIsCancelSub As String
Dim vDeliveryIDSub As Currency
Dim vQueueID As Currency
Dim vLineNumber As Integer
Dim vCheckCal As Currency
Dim vLineNumberEmp As Integer
Dim vQueueNo As String
Dim vQueueNoList(50) As String
Dim vCountQueue As Integer
Dim m As Integer
Dim n As Integer
Dim vInvoiceNo As String
Dim vItemCode As String
Dim vUnitCode As String
Dim vSendResult As String
Dim vLineCheckQTY As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 And ListView102.ListItems.Count <> 0 And MaskEdBox101.Text <> "__:__" And Text102.Text <> "" And Text103.Text <> "" And Text106.Text <> "" Then
        If DTPicker102.Value < DTPicker101.Value Then
            MsgBox "กรุณาตรวจสอบวันที่ของเอกสารกับวันที่ส่งของ ต้องกรอกข้อมูลตามความเป็นจริงด้วย", vbCritical, "Send Error"
            Exit Sub
        End If
        If Check101.Value = 1 Then
            If DTPicker103.Value < DTPicker102.Value Then
                MsgBox "กรุณาตรวจสอบวันที่กลับ ต้องมากว่าหรือเท่ากับวันที่ส่งสินค้า", vbCritical, "Send Error"
                DTPicker103.Value = DTPicker102.Value
                Exit Sub
            End If
            
            If DTPicker102.Value = DTPicker103.Value Then
                If MaskEdBox101.Text >= MaskEdBox102.Text Then
                    MsgBox "เวลากลับต้องมากกว่าเวลาส่งของ", vbCritical, "Send Error"
                    Exit Sub
                End If
            End If
            If MaskEdBox102.Text = "__:__" Or MaskEdBox102.Text = "00:00" Then
                MsgBox "ต้องกรอกเวลาในการกลับมาด้วย", vbCritical, "Send Error"
                    Exit Sub
            End If
        End If
    
        Call CheckIsCancel
        Call CheckEmpWage
        If vCheckEmpWages = False Then
         Exit Sub
        End If
        
        If vCheckIsCancel1 = 1 Then
            vIsOpen2 = 0
            ListView101.ListItems.Clear
            ListView102.ListItems.Clear
            DTPicker101.Value = Now
            DTPicker102.Value = Now
            DTPicker103.Value = Now
            Text106.Enabled = False
            Check101.Value = 0
            Text101.Text = ""
            Text102.Text = ""
            Text103.Text = ""
            Text104.Text = ""
            Text105.Text = ""
            Text106.Text = ""
            Text107.Text = ""
            MaskEdBox101.Text = "00:00"
            MaskEdBox102.Text = "00:00"
            Image101.Visible = True
            Image102.Visible = False
            Text106.Enabled = True
            Text107.Enabled = False
            MaskEdBox102.Enabled = False
            FormDelivery.DTPicker102.Enabled = True
            FormDelivery.MaskEdBox101.Enabled = True
            Check101.Enabled = True
            CMD106.Enabled = True
            CMD201.Enabled = True
            Exit Sub
        End If
        If vIsOpen2 = 0 Then
            Call CheckDocNo
            vID = "Null"
            vIsReturn = 0
            vSendResult = "0"
        Else
            If Check101.Value = 1 Then
                If (CDate(DTPicker102.Value) < CDate(DTPicker101.Value)) Or (Trim(MaskEdBox102.Text) = "00:00") Then
                    MsgBox "กรุณา ตรวจสอบข้อมูลวันที่กลับ เวลากลับ และ กม.สิ้นสุด ให้ถูกต้องด้วย", vbCritical, "Send Error"
                    Exit Sub
                End If
            End If
            
            vDocNo = Trim(Text102.Text)
            vQuery = "exec bcnp.dbo.USP_DO_SearchQueueCalQty '" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vRecordset.MoveFirst
                vCountQueue = vRecordset.RecordCount
                m = 1
                While Not vRecordset.EOF
                    vQueueNoList(m) = Trim(vRecordset.Fields("docno").Value)
                    m = m + 1
                vRecordset.MoveNext
                Wend
            End If
            vRecordset.Close
        
            vID = Trim(Text101.Text)
            If Check101.Value = 1 Then
                vIsReturn = 1
            Else
                vIsReturn = 0
            End If
            If Option201.Value = True Then
                vSendResult = 0
            ElseIf Option202.Value = True Then
                vSendResult = 1
            ElseIf Option203.Value = True Then
                vSendResult = 2
            ElseIf Option204.Value = True Then
                vSendResult = 3
            ElseIf Option205.Value = True Then
                vSendResult = 4
            ElseIf Option206.Value = True Then
                vSendResult = 5
            End If
            
        End If
        vDocNo = Trim(Text102.Text)

            
        vDocdate = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
        vDateSend = Trim(DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year)
        If Check101.Value = 1 Then
            If DTPicker103.Value < DTPicker102.Value Then
                MsgBox "กรุณาตรวจสอบวันที่กลับ ต้องมากว่าหรือเท่ากับวันที่ส่งสินค้า", vbCritical, "Send Error"
                Exit Sub
            End If
            vDateReturn = Trim(DTPicker103.Day & "/" & DTPicker103.Month & "/" & DTPicker103.Year)
            vTimeReturn = Trim(MaskEdBox102.Text)
            vMeasureStop = Trim(Text107.Text)
        Else
            vDateReturn = Trim(DTPicker103.Day & "/" & DTPicker103.Month & "/" & DTPicker103.Year)
            vTimeReturn = Trim(MaskEdBox102.Text)
            vMeasureStop = 0 'Trim(Text106.Text)
        End If
        vIsCancel = 0
        vTimeSend = Trim(MaskEdBox101.Text)
        vMeasureStart = Trim(Text106.Text)
        vMydescription = Trim(Text108.Text)
        vVehicalID = Trim(Text103.Text)
        'If Check101.Value = 1 Then
         '   vIsReturn = 1
        'Else
         '   vIsReturn = 0
        'End If
    
        'Insert Header
        
        On Error GoTo RollbackHeader
        
        vQuery = "begin tran"
        gConnection.Execute vQuery
        
        vQuery = "exec bcnp.dbo.USP_DO_DeliveryUpdate_Output " & vID & ",'" & vDocNo & "'," _
                            & " '" & vDocdate & "','" & vDateSend & "','" & vTimeSend & "','" & vDateReturn & "'," _
                            & " '" & vTimeReturn & "'," & vMeasureStart & "," & vMeasureStop & ",'" & vMydescription & "'," _
                            & " '" & vIsCancel & "'," & vVehicalID & ",'" & vIsReturn & "' ,'" & vUserID & "','" & vSendResult & "' "
        gConnection.Execute vQuery
'---------------------------------------------------------------------------------------------------------------------------------------
RollbackHeader:
        If Err.Description <> "" Then
            vReturn_Status = 1
            vIsCompleteSave = 0
            vDeliveryID = 1
            vEmpBPlusID = 1
            vEmpWage = 0
            vIsCancelEmp = 0
            vLineNumberEmp = 0
            vQuery = "exec bcnp.dbo.USP_DO_DeliveryEmpUpdate_Output " & vReturn_Status & " ," & vIsCompleteSave & " ," & vDeliveryID & ", " _
                                & " " & vEmpBPlusID & "," & vEmpWage & ",'" & vIsCancelEmp & "'," & vLineNumberEmp & " "
            gConnection.Execute vQuery
            
             vQuery = "rollback tran"
             gConnection.Execute vQuery
                    
             MsgBox Err.Description, vbCritical, "Send Error"
             Exit Sub
        End If
'----------------------------------------------------------------------------------------------------------------------------------------------
        vQuery = "select * from npmaster.dbo.TB_DO_Delivery  where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDeliveryID = Trim(vRecordset.Fields("id").Value)
        End If
        vRecordset.Close
        
        For i = 1 To ListView102.ListItems.Count
            vReturn_Status = 0
            vIsCompleteSave = 0
            vDeliveryID = vCheckDeliveryID
            vEmpBPlusID = ListView102.ListItems.Item(i).SubItems(1)
            If ListView102.ListItems.Item(i).SubItems(5) = "" Then
             vEmpWage = 0
            Else
             vEmpWage = ListView102.ListItems.Item(i).SubItems(5)
            End If
            vIsCancelEmp = 0
            vLineNumberEmp = Trim(ListView102.ListItems.Item(i).Text) - 1
            
            On Error GoTo RollBackEmp
            
            vQuery = "exec bcnp.dbo.USP_DO_DeliveryEmpUpdate_Output " & vReturn_Status & " ," & vIsCompleteSave & " ," & vDeliveryID & ", " _
                                & " " & vEmpBPlusID & ", " & vEmpWage & ", '" & vIsCancelEmp & "'," & vLineNumberEmp & " "
            gConnection.Execute vQuery
        Next i
'---------------------------------------------------------------------------------------------------------------------------------------------------------
RollBackEmp:
If Err.Description <> "" Then
    vReturn_Status = 1
    vIsCompleteSave = 0
    vDeliveryID = 1
    vEmpBPlusID = 1
    vEmpWage = 0
    vIsCancelEmp = 0
    vLineNumberEmp = 0
    vQuery = "exec bcnp.dbo.USP_DO_DeliveryEmpUpdate_Output " & vReturn_Status & " ," & vIsCompleteSave & " ," & vDeliveryID & ", " _
                        & " " & vEmpBPlusID & ", " & vEmpWage = 0 & ", '" & vIsCancelEmp & "'," & vLineNumberEmp & " "
    gConnection.Execute vQuery
    
    vQuery = "rollback tran"
    gConnection.Execute vQuery
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
'--------------------------------------------------------------------------------------------------------------------------------------------------------

vQuery = "exec dbo.USP_DO_DeliverySubDel " & vID & " "
gConnection.Execute vQuery
        For i = 1 To ListView101.ListItems.Count
            If i <> ListView101.ListItems.Count Then
                vReturn_Status = 0
                vIsCompleteSaveSub = 0
                vQTY = Trim(ListView101.ListItems.Item(i).SubItems(4))
                vConfirmQty = Trim(ListView101.ListItems.Item(i).SubItems(5))
                vMydescription = Trim(ListView101.ListItems.Item(i).SubItems(7))
                vIsCancelSub = 0
                vDeliveryIDSub = vCheckDeliveryID
                vQueueID = Trim(ListView101.ListItems.Item(i).SubItems(8))
                vInvoiceNo = Trim(ListView101.ListItems.Item(i).SubItems(9))
                vItemCode = Trim(ListView101.ListItems.Item(i).SubItems(2))
                vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(6))
                vLineNumber = Trim(ListView101.ListItems.Item(i).Text) - 1
                
                On Error GoTo RollbackDetails
                
                vQuery = "exec bcnp.dbo.USP_DO_DeliverySubUpdate_Output " & vReturn_StatusSub & "," & vIsCompleteSaveSub & " ," _
                                    & "  " & vQTY & ", " & vConfirmQty & ",'" & vMydescriptionSub & "','" & vIsCancelSub & "'," & vDeliveryIDSub & "," & vLineNumber & ", " _
                                    & " " & vQueueID & ",'" & vInvoiceNo & "','" & vItemCode & "','" & vUnitCode & "' "
                gConnection.Execute vQuery
                
                vQueueDetailID = Trim(ListView101.ListItems.Item(i).SubItems(1))
                vQuery = "exec bcnp.dbo.USP_DO_QueueRemainCal " & vQueueDetailID & ",1 "
                gConnection.Execute vQuery
                
                vQuery = "exec bcnp.dbo.USP_DO_QueueRemainCheck " & vQueueDetailID & " "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vCheckCal = Trim(vRecordset.Fields("RemainQty").Value)
                End If
                vRecordset.Close
                
                
                If vCheckCal < 0 Then
                    vReturn_StatusSub = 1
                    vIsCompleteSaveSub = 0
                    vQTY = 0
                    vConfirmQty = 0
                    vMydescription = ""
                    vIsCancelSub = 0
                    vDeliveryIDSub = 1
                    vQueueID = 1
                    vLineNumber = 0
                    vInvoiceNo = "1"
                    vItemCode = "1"
                    vUnitCode = "1"
                    
                    vQuery = "exec bcnp.dbo.USP_DO_DeliverySubUpdate_Output " & vReturn_StatusSub & "," & vIsCompleteSaveSub & " ," _
                                        & "  " & vQTY & ", " & vConfirmQty & ",'" & vMydescriptionSub & "','" & vIsCancelSub & "'," & vDeliveryIDSub & "," & vLineNumber & "," & vQueueID & ",'" & vInvoiceNo & "','" & vItemCode & "','" & vUnitCode & "'  "
                    gConnection.Execute vQuery
                    
                    vQuery = "rollback tran"
                    gConnection.Execute vQuery
                    
                    MsgBox "สินค้ารายการ " & Trim(ListView101.ListItems.Item(i).SubItems(2)) & "  ขนส่งเกินกว่าจำนวนที่ทำใบขอจัดคิวขนส่ง  กรุณาตรวจสอบ", vbCritical, "Send Error"
                    Exit Sub
                End If
                
            Else
                vQTY = Trim(ListView101.ListItems.Item(i).SubItems(4))
                vConfirmQty = Trim(ListView101.ListItems.Item(i).SubItems(5))
                vQueueDetailID = Trim(ListView101.ListItems.Item(i).SubItems(1))
                vQuery = "exec bcnp.dbo.USP_DO_QueueRemainCal " & vQueueDetailID & ",1 "
                gConnection.Execute vQuery
                
                vQuery = "exec bcnp.dbo.USP_DO_QueueRemainCheck " & vQueueDetailID & " "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vCheckCal = Trim(vRecordset.Fields("RemainQty").Value)
                End If
                vRecordset.Close
                
                If Me.Check101.Value = 1 Then
                 If vCheckCal <> vConfirmQty Then
                     vReturn_StatusSub = 1
                     vIsCompleteSaveSub = 0
                     vQTY = 0
                     vConfirmQty = 0
                     vMydescription = ""
                     vIsCancelSub = 0
                     vDeliveryIDSub = 1
                     vQueueID = 1
                     vLineNumber = 0
                     vInvoiceNo = "1"
                     vItemCode = "1"
                     vUnitCode = "1"
                     
                     vQuery = "exec bcnp.dbo.USP_DO_DeliverySubUpdate_Output " & vReturn_StatusSub & "," & vIsCompleteSaveSub & " ," _
                                         & "  " & vQTY & ", " & vConfirmQty & ",'" & vMydescriptionSub & "','" & vIsCancelSub & "'," & vDeliveryIDSub & "," & vLineNumber & " ," & vQueueID & ",'" & vInvoiceNo & "','" & vItemCode & "','" & vUnitCode & "' "
                     gConnection.Execute vQuery
                     
                     
                    vQuery = "rollback tran"
                    gConnection.Execute vQuery
                    
                    MsgBox "สินค้าขนส่งไม่เท่ากับจำนวนที่ส่งจริง หรือ จำนวนขนส่งติดลบ กรุณาตรวจสอบ", vbCritical, "Send Error"
                     Exit Sub
                 End If
                End If

                If vCheckCal < 0 Then
                    vReturn_StatusSub = 1
                    vIsCompleteSaveSub = 0
                    vQTY = 0
                    vConfirmQty = 0
                    vMydescription = ""
                    vIsCancelSub = 0
                    vDeliveryIDSub = 1
                    vQueueID = 1
                    vLineNumber = 0
                    vInvoiceNo = "1"
                    vItemCode = "1"
                    vUnitCode = "1"
                    
                    vQuery = "exec bcnp.dbo.USP_DO_DeliverySubUpdate_Output " & vReturn_StatusSub & "," & vIsCompleteSaveSub & " ," _
                                        & "  " & vQTY & ", " & vConfirmQty & ",'" & vMydescriptionSub & "','" & vIsCancelSub & "'," & vDeliveryIDSub & "," & vLineNumber & " ," & vQueueID & ",'" & vInvoiceNo & "','" & vItemCode & "','" & vUnitCode & "' "
                    gConnection.Execute vQuery
                    
                    vQuery = "rollback tran"
                    gConnection.Execute vQuery
                    MsgBox "สินค้าขนส่งไม่เท่ากับจำนวนที่ส่งจริง หรือ จำนวนขนส่งติดลบ กรุณาตรวจสอบ", vbCritical, "Send Error"
                    Exit Sub
                End If
                
                vReturn_Status = 0
                vIsCompleteSaveSub = 1
                'vQty = Trim(ListView101.ListItems.Item(i).SubItems(4))
                vConfirmQty = Trim(ListView101.ListItems.Item(i).SubItems(5))
                vMydescription = Trim(ListView101.ListItems.Item(i).SubItems(7))
                vIsCancelSub = 0
                vDeliveryIDSub = vCheckDeliveryID
                vQueueID = Trim(ListView101.ListItems.Item(i).SubItems(8))
                vInvoiceNo = Trim(ListView101.ListItems.Item(i).SubItems(9))
                vItemCode = Trim(ListView101.ListItems.Item(i).SubItems(2))
                vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(6))
                vLineNumber = Trim(ListView101.ListItems.Item(i).Text) - 1
                
                On Error GoTo RollbackDetails
                vQuery = "exec bcnp.dbo.USP_DO_DeliverySubUpdate_Output " & vReturn_StatusSub & "," & vIsCompleteSaveSub & " ," _
                                    & "  " & vQTY & ", " & vConfirmQty & ",'" & vMydescriptionSub & "','" & vIsCancelSub & "'," & vDeliveryIDSub & "," & vLineNumber & "," & vQueueID & ",'" & vInvoiceNo & "','" & vItemCode & "','" & vUnitCode & "' "
                gConnection.Execute vQuery
                
                vQueueDetailID = Trim(ListView101.ListItems.Item(i).SubItems(1))
                vQuery = "exec bcnp.dbo.USP_DO_QueueRemainCal " & vQueueDetailID & ",1 "
                gConnection.Execute vQuery
            End If
        Next i
        
RollbackDetails:
        If Err.Description <> "" Then
            vReturn_StatusSub = 1
            vIsCompleteSaveSub = 0
            vQTY = 0
            vConfirmQty = 0
            vMydescription = ""
            vIsCancelSub = 0
            vDeliveryIDSub = 1
            vQueueID = 1
            vLineNumber = 0
            vInvoiceNo = "1"
            vItemCode = "1"
            vUnitCode = "1"
            
            vQuery = "exec bcnp.dbo.USP_DO_DeliverySubUpdate_Output " & vReturn_StatusSub & "," & vIsCompleteSaveSub & " ," _
                                & "  " & vQTY & ", " & vConfirmQty & ",'" & vMydescriptionSub & "','" & vIsCancelSub & "'," & vDeliveryIDSub & "," & vLineNumber & "," & vQueueID & ",'" & vInvoiceNo & "','" & vItemCode & "','" & vUnitCode & "' "
            gConnection.Execute vQuery
            
            vQuery = "rollback tran"
            gConnection.Execute vQuery
            MsgBox Err.Description, vbCritical, "Send Error"
            Exit Sub
        End If

    On Error GoTo ErrDescription
    If vIsOpen2 = 0 Then
        MsgBox "เอกสารได้ทำการบันทึกเรียบร้อยแล้ว ได้เอกสารเลขที่ " & vDocNo & " ", vbInformation, "Send Message"
    Else
        MsgBox "เอกสารเลขที่ " & vDocNo & "  ได้ทำการอัพเดทเรียบร้อยแล้ว ", vbInformation, "Send Message"
        'For n = 1 To vCountQueue
         '       vQueueNo = vQueueNoList(n)
          '      vQuery = "exec bcnp.dbo.USP_DO_QueueRemainCal '" & vQueueNo & "' "
           '     gConnection.Execute vQuery
        'Next n
    End If
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
                    
    vIsOpen2 = 0
    ListView101.ListItems.Clear
    ListView102.ListItems.Clear
    DTPicker101.Value = Now
    DTPicker102.Value = Now
    DTPicker103.Value = Now
    Text106.Enabled = False
    Check101.Value = 0
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text104.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text107.Text = ""
    Text108.Text = ""
    MaskEdBox101.Mask = "##:##"
    MaskEdBox102.Mask = "##:##"
    Image101.Visible = True
    Image102.Visible = False
    Text106.Enabled = True
    Text107.Enabled = False
    MaskEdBox102.Enabled = False
    FormDelivery.DTPicker102.Enabled = True
    FormDelivery.MaskEdBox101.Enabled = True
    Check101.Enabled = True
    CMD106.Enabled = True
    CMD201.Enabled = True
    
'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description, vbCritical, "Send Error"
  '  Exit Sub
'End If
Else
    MsgBox "กรุณา กรอกข้อมูลให้ครบตามช่องที่มีตัวหนังสือสีแดง เพราะถ้าไม่ครบจะไม่สามรถบันทึกข้อมูลได้", vbInformation, "Send Message"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
On Error GoTo ErrDescription

FrmOrder013.Show
FormDelivery.Enabled = False
MDIFrmProgramPrint.DO1.Enabled = False

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD104_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vDocNo As String
Dim vRefType As Integer
Dim vRepType As String
Dim vRepID As Integer

On Error GoTo ErrDescription

If Text102.Text <> "" Then
    If vIsOpen2 = 1 Then
    vDocNo = Trim(Text102.Text)
    vRepType = "DO"
    vRepID = 292
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
    
        With Crystal101
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@DODocno;" & vDocNo & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
        End With
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD105_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vID As Currency
Dim vRefSubID As Double
Dim i As Integer
Dim vAnswer As Integer

On Error GoTo ErrDescription

If Text101.Text <> "" And Text102.Text <> "" And ListView101.ListItems.Count > 0 And vIsOpen2 = 1 Then
    vID = Trim(Text101.Text)
    vDocNo = Trim(Text102.Text)
    vAnswer = MsgBox("คุณต้องการยกเลิกเอกสารเลขที่ " & vDocNo & " ใช่หรือไม่", vbYesNo, "Question Respond")
    If vAnswer = 6 Then
        Call CheckIsCancel
            If vCheckIsCancel2 = 0 Then
            vQuery = "exec bcnp.dbo.usp_DO_CancelDeliveryHeader '" & vDocNo & "'," & vID & ",'" & vUserID & "' "
            gConnection.Execute vQuery
            
             vQuery = "exec bcnp.dbo.usp_DO_CancelDeliveryDetails     " & vID & " "
             gConnection.Execute vQuery
             
            MsgBox "ทำการยกเลิกเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้วครับ", vbInformation, "Send Message "
            vIsOpen2 = 0
            ListView101.ListItems.Clear
            ListView102.ListItems.Clear
            DTPicker101.Value = Now
            DTPicker102.Value = Now
            DTPicker103.Value = Now
            Text106.Enabled = False
            Check101.Value = 0
            Text101.Text = ""
            Text102.Text = ""
            Text103.Text = ""
            Text104.Text = ""
            Text105.Text = ""
            Text106.Text = ""
            Text107.Text = ""
            MaskEdBox101.Mask = "##:##"
            MaskEdBox102.Mask = "##:##"
            Text106.Enabled = True
            Text107.Enabled = False
            MaskEdBox102.Enabled = False
            FormDelivery.DTPicker102.Enabled = True
            FormDelivery.MaskEdBox101.Enabled = True
            Check101.Enabled = True
            CMD106.Enabled = True
            CMD201.Enabled = True
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD106_Click()
On Error GoTo ErrDescription

If Check101.Value = 0 Then
    FrmOrder010.Show
    FormDelivery.Enabled = False
    MDIFrmProgramPrint.DO1.Enabled = False
Else
    MsgBox "ได้เลือกให้ใบขนส่งใบนี้เป็นการขนส่งที่สมบูรณ์แล้วไม่สามารถเพิ่มข้อมูลได้", vbInformation, "Send Information"
End If
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD107_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vGenDocNo As String

On Error GoTo ErrDescription

If ListView101.ListItems.Count > 0 And vIsOpen2 = 0 Then
    vQuery = "exec bcnp.dbo.USP_DO_DONewDocNo"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vGenDocNo = Trim(vRecordset.Fields("QueueNewDocNo").Value)
    End If
    vRecordset.Close
    Text102.Text = vGenDocNo
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD108_Click()
On Error GoTo ErrDescription

MDIFrmProgramPrint.Enabled = False
FrmOrder011.Show
vVehicalModule = 1

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD109_Click()
Unload FormDelivery
End Sub

Private Sub CMD201_Click()
On Error GoTo ErrDescription

MDIFrmProgramPrint.Enabled = False
FrmOrder012.Show
vEmpModule = 1

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMDEmpWage_Click()
Dim vWage As Double

On Error Resume Next

If Me.LBLEmpID.Caption <> "" And Me.LBLEmpName.Caption <> "" And Me.TXTWage.Text <> "" Then
 vWage = Me.TXTWage.Text
 Me.ListView102.ListItems(vEmpWageID).SubItems(5) = Format(vWage, "##,##0.00")
 Me.LBLEmpID.Caption = ""
 Me.LBLEmpName.Caption = ""
 Me.TXTWage.Text = ""
 Me.PICEmpWage.Visible = False
 Me.ListView102.SetFocus
 Else
 MsgBox "ต้องกรอกข้อมูลค่าเที่ยวด้วย กรณีไม่มีก็ให้ใส่ 0", vbCritical, "Send Error Message"
 Me.TXTWage.SetFocus
End If
End Sub

Private Sub DTPicker102_Change()
On Error GoTo ErrDescription

If CDate(DTPicker102.Value) < CDate(DTPicker101.Value) Then
    MsgBox "วันที่ส่งสินค้า ต้องเป็นวันเดียวกันกับวันที่เอกสาร หรือ ต้องมากกว่าวันที่เอกสาร", vbCritical, "Send Error"
    DTPicker102.Value = DTPicker101.Value
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub DTPicker102_Click()
On Error GoTo ErrDescription

If CDate(DTPicker102.Value) < CDate(DTPicker101.Value) Then
    MsgBox "วันที่ส่งสินค้า ต้องเป็นวันเดียวกันกับวันที่เอกสาร หรือ ต้องมากกว่าวันที่เอกสาร", vbCritical, "Send Error"
    DTPicker102.Value = DTPicker101.Value
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub DTPicker103_Change()
On Error GoTo ErrDescription

If DTPicker103.Value < DTPicker102.Value Then
    MsgBox "วันที่กลับ ต้องเป็นวันเดียวกันกับวันที่ส่งสินค้า หรือ ต้องมากกว่าวันที่ส่งสินค้า", vbCritical, "Send Error"
    DTPicker103.Value = DTPicker102.Value
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub DTPicker103_Click()
On Error GoTo ErrDescription

If DTPicker103.Value < DTPicker102.Value Then
    MsgBox "วันที่กลับ ต้องเป็นวันเดียวกันกับวันที่ส่งสินค้า หรือ ต้องมากกว่าวันที่ส่งสินค้า", vbCritical, "Send Error"
    DTPicker103.Value = DTPicker102.Value
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrDescription

Image101.Visible = True
Image102.Visible = False
DTPicker101 = Now
DTPicker102 = Now
DTPicker103 = Now
vIsOpen2 = 0
SSTab1.TabEnabled(2) = False

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Public Sub CheckIsCancel()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocNo As String
Dim vGenDocNo As String

vCheckDocNo = Trim(Text102.Text)
vQuery = "select docno,iscancel from npmaster.dbo.TB_DO_Delivery where docno = '" & vCheckDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckIsCancel2 = Trim(vRecordset.Fields("iscancel").Value)
Else
    vCheckIsCancel2 = 0
End If
vRecordset.Close

End Sub

Public Sub CheckDocNo()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocNo As String
Dim vGenDocNo As String

vCheckDocNo = Trim(Text102.Text)
vQuery = "select docno,iscancel from npmaster.dbo.TB_DO_Delivery where docno = '" & vCheckDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDocnoExist2 = 1
    vCheckIsCancel2 = Trim(vRecordset.Fields("iscancel").Value)
Else
    vCheckDocnoExist2 = 0
    vCheckIsCancel2 = 0
End If
vRecordset.Close

If vCheckDocnoExist2 = 1 Then
    vQuery = "exec bcnp.dbo.USP_DO_DONewDocNo"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vGenDocNo = Trim(vRecordset.Fields("QueueNewDocNo").Value)
    End If
    vRecordset.Close
    Text102.Text = vGenDocNo
End If
End Sub

Private Sub ListView101_DblClick()
Dim vCheckQty As String
Dim i As Integer


vCheckQty = InputBox("จำนวนขนส่ง", "กรอกจำนวนที่ส่งได้จริง")
i = ListView101.SelectedItem.Index
If vCheckQty <> "" Then
    If CCur(vCheckQty) <= CCur(ListView101.ListItems.Item(i).SubItems(4)) Then
        ListView101.ListItems.Item(i).SubItems(5) = vCheckQty
    Else
        MsgBox "จำนวนขนส่งจริงต้องไม่มากกว่าจำนวนขนส่ง", vbCritical, "Send Error"
    End If
End If
End Sub

Private Sub ListView101_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vCheckCountOld As Integer
Dim vNewCount As Integer
Dim i As Integer

On Error GoTo ErrDescription

If KeyCode = 46 Then
    If ListView101.ListItems.Count <> 0 Then
        vCheckCountOld = ListView101.ListItems.Count
        
        If Check101.Value = 0 Then
            ListView101.ListItems.Remove (ListView101.SelectedItem.Index)
        Else
            MsgBox "ไม่สามารถลบรายการได้ เนื่องจากการขนส่งสมบูรณ์แล้ว", vbCritical, "Send Informtion"
            Exit Sub
        End If
        
        vNewCount = ListView101.ListItems.Count
        
        If vCheckCountOld <> vNewCount Then
            For i = 1 To vNewCount
                ListView101.ListItems.Item(i).Text = i
            Next i
        End If
    End If
End If
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub ListView102_DblClick()

If ListView102.ListItems.Count > 0 Then
 Me.PICEmpWage.Visible = True
 vEmpWageID = Me.ListView102.SelectedItem.Index
 Me.LBLEmpID.Caption = ListView102.ListItems(vEmpWageID).SubItems(2)
 Me.LBLEmpName.Caption = ListView102.ListItems(vEmpWageID).SubItems(3)
 Me.TXTWage.Text = Format(ListView102.ListItems(vEmpWageID).SubItems(5), "##,##0.00")
 Me.TXTWage.SetFocus
End If
End Sub

Private Sub ListView102_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vCheckCountOld As Integer
Dim vNewCount As Integer
Dim i As Integer

On Error GoTo ErrDescription

vCheckCountOld = ListView102.ListItems.Count

If KeyCode = 46 Then
    ListView102.ListItems.Remove (ListView102.SelectedItem.Index)
End If

vNewCount = ListView102.ListItems.Count

If vCheckCountOld <> vNewCount Then
    For i = 1 To vNewCount
        ListView102.ListItems.Item(i).Text = i
    Next i
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub ListView102_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If ListView102.ListItems.Count > 0 Then
  Me.PICEmpWage.Visible = True
  vEmpWageID = Me.ListView102.SelectedItem.Index
  Me.LBLEmpID.Caption = ListView102.ListItems(vEmpWageID).SubItems(2)
  Me.LBLEmpName.Caption = ListView102.ListItems(vEmpWageID).SubItems(3)
  Me.TXTWage.Text = Format(ListView102.ListItems(vEmpWageID).SubItems(5), "##,##0.00")
  Me.TXTWage.SetFocus
 End If
End If
End Sub

Private Sub MaskEdBox101_LostFocus()
Dim vTime1 As String
Dim vTime2 As String
Dim vTime3 As String
Dim vTime4 As String
Dim vTime5 As String
Dim vTime6 As String

If Text102.Text <> "" Or ListView101.ListItems.Count <> 0 Then
    vTime1 = Left(MaskEdBox101.Text, 1)
    vTime2 = Right(Left(MaskEdBox101.Text, 2), 1)
    vTime3 = Right(MaskEdBox101.Text, 1)
    vTime4 = Left(Right(MaskEdBox101.Text, 2), 1)
    vTime5 = Left(MaskEdBox101.Text, 2)
    vTime6 = Right(MaskEdBox101.Text, 2)
    
    If vTime1 = "_" Or vTime2 = "_" Or vTime3 = "_" Or vTime4 = "_" Then
        MsgBox "รูปแบบเวลาที่ใช้ในโปรแกรม คือ '00:00' ต้องใส่ให้ครบด้วยครับ", vbInformation, "Send Information"
        MaskEdBox101.SetFocus
        Exit Sub
    End If
    
    If vTime5 > "24" Or vTime6 > "59" Then
        MsgBox "รูปแบบเวลาที่ใช้ในโปรแกรม คือ ชั่วโมงต้องใส่ไม่เกิน 24 และ นาทีต้องไม่เกิน 60", vbInformation, "Send Information"
        MaskEdBox101.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub MaskEdBox102_LostFocus()
Dim vTime1 As String
Dim vTime2 As String
Dim vTime3 As String
Dim vTime4 As String
Dim vTime5 As String
Dim vTime6 As String
Dim vDate1 As Date
Dim vDate2 As Date
Dim vTime11 As String
Dim vTime12 As String
Dim vTime13 As String
Dim vTime14 As String
Dim vDTime1 As Integer
Dim vDTime2 As Integer
Dim vDTime3 As Integer
Dim vDTime4 As Integer

On Error GoTo ErrDescription

If Text101.Text <> "" And Check101.Value = 1 Then
    vTime1 = Left(MaskEdBox102.Text, 1)
    vTime2 = Right(Left(MaskEdBox102.Text, 2), 1)
    vTime3 = Right(MaskEdBox102.Text, 1)
    vTime4 = Left(Right(MaskEdBox102.Text, 2), 1)
    vTime5 = Left(MaskEdBox102.Text, 2)
    vTime6 = Right(MaskEdBox102.Text, 2)
    
    If vTime1 = "_" Or vTime2 = "_" Or vTime3 = "_" Or vTime4 = "_" Then
        MsgBox "รูปแบบเวลาที่ใช้ในโปรแกรม คือ '00:00' ต้องใส่ให้ครบด้วยครับ", vbInformation, "Send Information"
        MaskEdBox102.SetFocus
        Exit Sub
    End If
    
    If vTime5 > 24 Or vTime6 > 59 Then
        MsgBox "รูปแบบเวลาที่ใช้ในโปรแกรม คือ ชั่วโมงต้องใส่ไม่เกิน 24 และ นาทีต้องไม่เกิน 60", vbInformation, "Send Information"
        MaskEdBox102.SetFocus
        Exit Sub
    End If
    
    vDate1 = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
    vDate2 = Trim(DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year)
    If vDate2 = vDate1 Then
        vTime11 = Left(MaskEdBox101.Text, 1)
        vTime12 = Right(Left(MaskEdBox101.Text, 2), 1)
        vTime13 = Right(MaskEdBox101.Text, 1)
        vTime14 = Left(Right(MaskEdBox101.Text, 2), 1)
        vDTime1 = vTime1 & vTime2
        vDTime2 = vTime11 & vTime12
        If vDTime2 > vDTime1 Then
            MsgBox "เวลาชั่วโมงกลับน้อยกว่าเวลาชั่วโมงส่ง กรุณาตรวจสอบ", vbCritical, "Send Error"
            MaskEdBox102.SetFocus
            Exit Sub
        End If
        If vDTime1 = vDTime2 Then
        vDTime3 = vTime4 & vTime3
        vDTime4 = vTime14 & vTime13
            If vDTime4 >= vDTime3 Then
                MsgBox "เวลานาทีกลับน้อยกว่าเวลานาทีส่ง กรุณาตรวจสอบ", vbCritical, "Send Error"
                MaskEdBox102.SetFocus
                Exit Sub
            End If
        End If
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.TabIndex = 2 And Me.ListView102.ListItems.Count > 0 Then
  Me.ListView102.SetFocus
End If
End Sub

Private Sub SSTab1_DblClick()
If SSTab1.TabIndex = 2 And Me.ListView102.ListItems.Count > 0 Then
  Me.ListView102.SetFocus
End If
End Sub

Private Sub Text107_LostFocus()
Dim vCheckMeterBeg As Currency
Dim vCheckMeterStop As Currency

On Error Resume Next

If Check101.Value = 1 Then
    vCheckMeterBeg = Trim(Text106.Text)
    vCheckMeterStop = Trim(Text107.Text)
    If vCheckMeterStop < vCheckMeterBeg Then
        MsgBox "กม.ที่สิ้นสุด น่าจะมากกว่าหรือเท่ากับ  กม.ที่เริ่มต้น", vbCritical, "Send Information"
        Text107.SetFocus
    End If
End If
End Sub

Private Sub CheckEmpWage()
Dim i As Integer
Dim vCheckWage As Double

If Me.ListView102.ListItems.Count > 0 Then
 i = Me.ListView102.ListItems.Count
 For i = 1 To i
 vCheckWage = Me.ListView102.ListItems.Item(i).SubItems(5)
 If vCheckWage = 0 Then
  MsgBox "เอกสารจัดส่ง ยังไม่ได้กำหนดค่าเที่ยวให้กับพนักงาน กรุณาตรวจสอบ"
  vCheckEmpWages = False
  Exit Sub
 Else
  vCheckEmpWages = True
 End If
 Next i
End If
 
End Sub


Private Sub TXTWage_Change()
Dim vString As String
Dim vCheckNumber As Boolean

If Me.TXTWage.Text <> "" Then
 vString = Trim(Me.TXTWage.Text)
 CheckNumber (vString)
 
 If vCheckValueNumber = False Then
    MsgBox "กรุณากรอกเฉพาะมูลค่าของค่าเที่ยวเท่านั้น", vbCritical, "Send Error Message"
    Me.TXTWage.Text = Format(ListView102.ListItems(vEmpWageID).SubItems(5), "##,##0.00")
    Me.TXTWage.SetFocus
 End If
End If

End Sub

Private Sub TXTWage_KeyPress(KeyAscii As Integer)
Dim vWage As Double

If KeyAscii = 13 Then
 If Me.LBLEmpID.Caption <> "" And Me.LBLEmpName.Caption <> "" And Me.TXTWage.Text <> "" Then
  vWage = Me.TXTWage.Text
  Me.ListView102.ListItems(vEmpWageID).SubItems(5) = Format(vWage, "##,##0.00")
  Me.LBLEmpID.Caption = ""
  Me.LBLEmpName.Caption = ""
  Me.TXTWage.Text = ""
  Me.PICEmpWage.Visible = False
  Me.ListView102.SetFocus
 Else
  MsgBox "ต้องกรอกข้อมูลค่าเที่ยวด้วย กรณีไม่มีก็ให้ใส่ 0", vbCritical, "Send Error Message"
  Me.TXTWage.SetFocus
 End If
End If
End Sub

Private Sub TXTWage_LostFocus()
If TXTWage.Text <> "" Then
  Me.TXTWage.Text = Format(Me.TXTWage.Text, "##,##0.00")
End If
End Sub
