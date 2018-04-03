VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmQueue 
   BackColor       =   &H00808080&
   Caption         =   "โปรแกรม จัดควบคุมคิวสินค้า"
   ClientHeight    =   7800
   ClientLeft      =   4935
   ClientTop       =   1710
   ClientWidth     =   12090
   Icon            =   "FrmQueue.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9021.44
   ScaleMode       =   0  'User
   ScaleWidth      =   12090
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer105 
      Enabled         =   0   'False
      Interval        =   17000
      Left            =   4140
      Top             =   10080
   End
   Begin VB.Timer Timer104 
      Interval        =   3000
      Left            =   3690
      Top             =   10080
   End
   Begin VB.Timer Timer103 
      Interval        =   12000
      Left            =   1
      Top             =   10080
   End
   Begin VB.Timer Timer102 
      Enabled         =   0   'False
      Interval        =   9000
      Left            =   2790
      Top             =   10080
   End
   Begin VB.Timer Timer101 
      Interval        =   20000
      Left            =   2340
      Top             =   10080
   End
   Begin VB.PictureBox PicScanBar 
      BackColor       =   &H00C0C0C0&
      Height          =   8835
      Left            =   0
      ScaleHeight     =   8775
      ScaleWidth      =   14130
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   14190
      Begin Crystal.CrystalReport Crystal101 
         Left            =   945
         Top             =   7200
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
      Begin VB.PictureBox PTShowMyDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   4245
         Left            =   7740
         ScaleHeight     =   4215
         ScaleWidth      =   3945
         TabIndex        =   85
         Top             =   2565
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CommandButton CMDUpdateMyDescription 
            Caption         =   "ปรับหมายเหตุ"
            Height          =   510
            Left            =   2475
            TabIndex        =   87
            Top             =   3510
            Width           =   1275
         End
         Begin VB.TextBox TextMyDescription 
            Appearance      =   0  'Flat
            Height          =   3210
            Left            =   180
            MultiLine       =   -1  'True
            TabIndex        =   86
            Top             =   90
            Width           =   3570
         End
      End
      Begin VB.PictureBox PicEditQty 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3300
         Left            =   135
         ScaleHeight     =   3270
         ScaleWidth      =   11550
         TabIndex        =   89
         Top             =   3600
         Visible         =   0   'False
         Width           =   11580
         Begin VB.CommandButton CMDCloseEditQty 
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
            Height          =   420
            Left            =   3420
            TabIndex        =   100
            Top             =   2070
            Width           =   1230
         End
         Begin VB.CommandButton CMBQtyEdit 
            Caption         =   "แก้ไขจำนวน"
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
            Left            =   2025
            TabIndex        =   98
            Top             =   2070
            Width           =   1230
         End
         Begin VB.TextBox TBQtyEdit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   97
            Top             =   1575
            Width           =   1230
         End
         Begin VB.Label LBLUnitCode 
            BackStyle       =   0  'Transparent
            Height          =   330
            Left            =   3375
            TabIndex        =   99
            Top             =   1170
            Width           =   915
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "จัดได้จำนวน :"
            Height          =   240
            Left            =   360
            TabIndex        =   96
            Top             =   1575
            Width           =   1500
         End
         Begin VB.Label LBLQtyEdit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2025
            TabIndex        =   95
            Top             =   1170
            Width           =   1230
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ต้องการจำนวน :"
            Height          =   285
            Left            =   360
            TabIndex        =   94
            Top             =   1170
            Width           =   1500
         End
         Begin VB.Label LBLItemNameEdit 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2025
            TabIndex        =   93
            Top             =   765
            Width           =   8835
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ชื่อสินค้า :"
            Height          =   285
            Left            =   405
            TabIndex        =   92
            Top             =   765
            Width           =   1455
         End
         Begin VB.Label LBLItemCodeEdit 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2025
            TabIndex        =   91
            Top             =   360
            Width           =   2625
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "รหัสสินค้า :"
            Height          =   285
            Left            =   450
            TabIndex        =   90
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CommandButton CMDCustRecCancel 
         Caption         =   "ยกเลิกการรับของ"
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
         Left            =   9135
         TabIndex        =   88
         Top             =   810
         Width           =   2580
      End
      Begin VB.CommandButton CMDDescription 
         Caption         =   "หมายเหตุ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6525
         TabIndex        =   84
         Top             =   2565
         Width           =   1140
      End
      Begin VB.CheckBox CHKDate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "คิวข้ามวัน"
         Height          =   285
         Left            =   1215
         TabIndex        =   74
         Top             =   900
         Width           =   1185
      End
      Begin MSComctlLib.ListView ListView105 
         Height          =   2625
         Left            =   135
         TabIndex        =   57
         Top             =   3600
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   4630
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "จำนวนที่จะขาย"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วยนับ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ชั้นเก็บ"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "ต้องการ"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "จัดได้"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Family"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.CommandButton CMDClose 
         BackColor       =   &H00808080&
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
         Left            =   10665
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   6345
         Width           =   1050
      End
      Begin VB.TextBox TextScan101 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   5175
         TabIndex        =   34
         Top             =   90
         Width           =   2490
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่คิว :"
         Height          =   330
         Left            =   8010
         TabIndex        =   73
         Top             =   1305
         Width           =   1050
      End
      Begin VB.Label LBLDocDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   9135
         TabIndex        =   72
         Top             =   1260
         Width           =   2580
      End
      Begin VB.Image Image2 
         Height          =   750
         Left            =   0
         Picture         =   "FrmQueue.frx":1272
         Top             =   0
         Width           =   2160
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "รายการสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   135
         TabIndex        =   56
         Top             =   3330
         Width           =   1140
      End
      Begin VB.Line Line1 
         BorderWidth     =   10
         X1              =   0
         X2              =   12060
         Y1              =   3105
         Y2              =   3105
      End
      Begin VB.Label LBLARCode102 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   3060
         TabIndex        =   55
         Top             =   1665
         Width           =   4605
      End
      Begin VB.Label LBLDocType101 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   5130
         TabIndex        =   54
         Top             =   2565
         Width           =   420
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ประเภท :"
         Height          =   330
         Left            =   4410
         TabIndex        =   53
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label LBLTimeID101 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   3060
         TabIndex        =   52
         Top             =   2115
         Width           =   870
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ครั้งที่ :"
         Height          =   375
         Left            =   2430
         TabIndex        =   51
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "นาที (H:M:S)"
         Height          =   240
         Left            =   2475
         TabIndex        =   50
         Top             =   2610
         Width           =   1455
      End
      Begin VB.Label LBLARCode101 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1215
         TabIndex        =   48
         Top             =   1665
         Width           =   1770
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ชื่อลูกค้า :"
         Height          =   285
         Left            =   360
         TabIndex        =   47
         Top             =   1710
         Width           =   780
      End
      Begin VB.Label LBLTime101 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1215
         TabIndex        =   46
         Top             =   2565
         Width           =   1095
      End
      Begin VB.Label LBLSale101 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   5130
         TabIndex        =   45
         Top             =   2115
         Width           =   2535
      End
      Begin VB.Label LBLWHCode101 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1215
         TabIndex        =   44
         Top             =   2115
         Width           =   1095
      End
      Begin VB.Label LBLPicker101 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   9135
         TabIndex        =   43
         Top             =   2115
         Width           =   2580
      End
      Begin VB.Label LBLRefDocNo101 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   9135
         TabIndex        =   42
         Top             =   1665
         Width           =   2580
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "เวลาที่ใช้จัด :"
         Height          =   330
         Left            =   225
         TabIndex        =   41
         Top             =   2610
         Width           =   1005
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "พนักงานขาย :"
         Height          =   330
         Left            =   4005
         TabIndex        =   40
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "คลัง :"
         Height          =   375
         Left            =   675
         TabIndex        =   39
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "พนักงานจัดสินค้า :"
         Height          =   285
         Left            =   7695
         TabIndex        =   38
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "เลขที่อ้างอิง :"
         Height          =   285
         Left            =   8145
         TabIndex        =   37
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label LBLStatus101 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1215
         TabIndex        =   36
         Top             =   1215
         Width           =   6450
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "สถานะ :"
         Height          =   330
         Left            =   585
         TabIndex        =   35
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "เลขที่คิว :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   645
         Left            =   3015
         TabIndex        =   33
         Top             =   270
         Width           =   1995
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   7845
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   22440
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   22470
      Begin VB.CommandButton CMDOK 
         Caption         =   "ปิดหน้าจอ"
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
         Left            =   10530
         TabIndex        =   22
         Top             =   6030
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView103 
         Height          =   4470
         Left            =   585
         TabIndex        =   4
         Top             =   1440
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   7885
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "จำนวน"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วย"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ชั้นเก็บ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "สถานะการหยิบ"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "สถานะการจ่าย"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Image Image3 
         Height          =   750
         Left            =   0
         Picture         =   "FrmQueue.frx":26D4
         Top             =   0
         Width           =   2160
      End
      Begin VB.Label LBLDocno2 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxx-xxxx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   6570
         TabIndex        =   11
         Top             =   225
         Width           =   4455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการสินค้า"
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
         Left            =   585
         TabIndex        =   19
         Top             =   1125
         Width           =   1185
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "คลังที่เก็บ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   3045
         TabIndex        =   18
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อคนจัด :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   3120
         TabIndex        =   17
         Top             =   585
         Width           =   1140
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่อ้างอิง :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   5190
         TabIndex        =   16
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่คิว :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   2610
         TabIndex        =   15
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label LBLWHCode 
         BackStyle       =   0  'Transparent
         Caption         =   "xxx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   4050
         TabIndex        =   13
         Top             =   945
         Width           =   600
      End
      Begin VB.Label LBLPicker 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   4050
         TabIndex        =   12
         Top             =   585
         Width           =   5730
      End
      Begin VB.Label LBLDocno1 
         BackStyle       =   0  'Transparent
         Caption         =   "xxx-xxxxxxx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   4050
         TabIndex        =   10
         Top             =   225
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0C0C0&
      Height          =   7845
      Left            =   0
      ScaleHeight     =   7785
      ScaleWidth      =   19485
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   19545
      Begin VB.OptionButton OPT104 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4. ลูกค้าเปลี่ยนสินค้า"
         Height          =   285
         Left            =   810
         TabIndex        =   83
         Top             =   3465
         Width           =   5955
      End
      Begin VB.TextBox TextDescription 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   855
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   82
         Top             =   4455
         Width           =   5775
      End
      Begin VB.CommandButton CMDSendExit 
         BackColor       =   &H00808080&
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
         Left            =   5310
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   5355
         Width           =   1320
      End
      Begin VB.CommandButton CMDSendResult 
         BackColor       =   &H00808080&
         Caption         =   "บันทึกผล"
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
         Left            =   3825
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   5355
         Width           =   1320
      End
      Begin VB.OptionButton OPT103 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3. รอจัดส่งรับของส่งให้ลูกค้าวันรุ่งขึ้น"
         Height          =   375
         Left            =   810
         TabIndex        =   65
         Top             =   3060
         Width           =   4200
      End
      Begin VB.OptionButton OPT102 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2. ลูกค้ายกเลิกการซื้อ-ขาย"
         Height          =   375
         Left            =   810
         TabIndex        =   64
         Top             =   2655
         Width           =   4200
      End
      Begin VB.OptionButton OPT101 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1. ลูกค้ารับของแล้ว"
         Height          =   420
         Left            =   810
         TabIndex        =   63
         Top             =   2250
         Value           =   -1  'True
         Width           =   3705
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   855
         TabIndex        =   81
         Top             =   4185
         Width           =   915
      End
      Begin VB.Label LBLSendRefNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4815
         TabIndex        =   71
         Top             =   945
         Width           =   1815
      End
      Begin VB.Label LBLSendQueue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2295
         TabIndex        =   70
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "อ้างถึง :"
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
         Left            =   3780
         TabIndex        =   69
         Top             =   990
         Width           =   915
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "คิวที่ :"
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
         Left            =   1260
         TabIndex        =   68
         Top             =   990
         Width           =   915
      End
      Begin VB.Line Line5 
         X1              =   810
         X2              =   6570
         Y1              =   4005
         Y2              =   4005
      End
      Begin VB.Line Line4 
         X1              =   810
         X2              =   6570
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Line Line3 
         X1              =   810
         X2              =   6570
         Y1              =   2205
         Y2              =   2205
      End
      Begin VB.Line Line2 
         X1              =   810
         X2              =   6570
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   0
         Picture         =   "FrmQueue.frx":3B36
         Top             =   0
         Width           =   2160
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "ผลการจ่าย-รับสินค้าที่จัดได้"
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
         Left            =   810
         TabIndex        =   62
         Top             =   1800
         Width           =   2580
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   45
      ScaleHeight     =   4215
      ScaleWidth      =   5970
      TabIndex        =   6
      Top             =   45
      Width           =   6000
      Begin VB.CommandButton CMDRefresh 
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
         Height          =   420
         Left            =   4590
         TabIndex        =   31
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton CMDCustItem 
         Caption         =   "พิมพ์ใบจ่าย"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   5670
         TabIndex        =   30
         Top             =   765
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox Text101 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   0
         Top             =   495
         Width           =   1275
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   3075
         Left            =   90
         TabIndex        =   1
         Top             =   945
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5424
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "คิวที่"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "พนักงานขาย"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "เวลาขอรับของ"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "อ้างถึง"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "ประเภท"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "TimeID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "CustZone"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "วันที่คิว"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "TimePick"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "ZoneID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "FamilyGroup"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "คลัง"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "ShelfGroup"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "PickZone"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "คิว :"
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
         Left            =   180
         TabIndex        =   28
         Top             =   540
         Width           =   420
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการ เอกสารที่รอการจัดสินค้า"
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
         Height          =   330
         Left            =   180
         TabIndex        =   21
         Top             =   90
         Width           =   4065
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   6120
      ScaleHeight     =   4215
      ScaleWidth      =   5880
      TabIndex        =   8
      Top             =   45
      Width           =   5910
      Begin VB.TextBox Text102 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   2
         Top             =   495
         Width           =   1320
      End
      Begin MSComctlLib.ListView ListView102 
         Height          =   3075
         Left            =   90
         TabIndex        =   3
         Top             =   945
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   5424
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "คิวที่"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "เริ่มจัด"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ใช้เวลา"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "อ้างถึง"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "ประเภท"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "TimeID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "คนหยิบ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "พนักงานขาย"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "StartDate"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "CustZone"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "วันที่คิว"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "TimePick"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "คิว :"
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
         Left            =   180
         TabIndex        =   29
         Top             =   540
         Width           =   420
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการ เอกสารที่กำลังจัดสินค้า"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   90
         Width           =   3570
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   6120
      ScaleHeight     =   2955
      ScaleWidth      =   5880
      TabIndex        =   9
      Top             =   4365
      Width           =   5910
      Begin MSComctlLib.ListView ListView104 
         Height          =   2355
         Left            =   90
         TabIndex        =   5
         Top             =   450
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   4154
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
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "คิวที่"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SaleOrderNo"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Picker"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "SaleOrder"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "SaleMan"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "TypeCode"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "CustomerZone"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "PickingStatus"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "วันที่คิว"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "TimePick"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการ เอกสารที่จัดสินค้าเสร็จแล้ว"
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
         Height          =   285
         Left            =   180
         TabIndex        =   20
         Top             =   90
         Width           =   2940
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   45
      ScaleHeight     =   2955
      ScaleWidth      =   5970
      TabIndex        =   23
      Top             =   4365
      Width           =   6000
      Begin VB.PictureBox PICReserve 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   0
         Picture         =   "FrmQueue.frx":4F98
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   102
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox PicPay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   0
         Picture         =   "FrmQueue.frx":74E5
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   101
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox Pic102 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   5400
         Picture         =   "FrmQueue.frx":9AB3
         ScaleHeight     =   480
         ScaleWidth      =   435
         TabIndex        =   60
         Top             =   90
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox Pic101 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   5400
         Picture         =   "FrmQueue.frx":C248
         ScaleHeight     =   480
         ScaleWidth      =   435
         TabIndex        =   59
         Top             =   90
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label LBLQueueDate 
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1485
         TabIndex        =   80
         Top             =   45
         Width           =   2715
      End
      Begin VB.Line Line12 
         X1              =   0
         X2              =   5310
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line11 
         X1              =   0
         X2              =   5310
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   5310
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   5310
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Line Line8 
         X1              =   5310
         X2              =   5310
         Y1              =   0
         Y2              =   2970
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   5310
         Y1              =   2475
         Y2              =   2475
      End
      Begin VB.Line Line6 
         X1              =   1395
         X2              =   1395
         Y1              =   0
         Y2              =   2475
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่คิว :"
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
         Left            =   180
         TabIndex        =   79
         Top             =   90
         Width           =   1140
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย :"
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
         Left            =   180
         TabIndex        =   78
         Top             =   2115
         Width           =   1140
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้จัดสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   77
         Top             =   1710
         Width           =   1140
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "อ้างถึง :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   76
         Top             =   1305
         Width           =   1140
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า :"
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
         Left            =   180
         TabIndex        =   75
         Top             =   540
         Width           =   1140
      End
      Begin VB.Label LBLCustomerZone 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   58
         Top             =   2565
         Width           =   5235
      End
      Begin VB.Label LBLUserPick 
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1485
         TabIndex        =   26
         Top             =   1710
         Width           =   3795
      End
      Begin VB.Label LBLRefNo 
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1485
         TabIndex        =   25
         Top             =   1305
         Width           =   3795
      End
      Begin VB.Label LBLARName 
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
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
         Left            =   1485
         TabIndex        =   24
         Top             =   540
         Width           =   3795
      End
      Begin VB.Label LBLSale 
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1485
         TabIndex        =   27
         Top             =   2115
         Width           =   3750
      End
   End
   Begin VB.Menu MQueue 
      Caption         =   "Queue"
   End
   Begin VB.Menu Menu1 
      Caption         =   ""
      Begin VB.Menu MReserve 
         Caption         =   "พิมพ์ทดแทน"
      End
      Begin VB.Menu MDescription 
         Caption         =   "กรอกหมายเหตุคิวจัดสินค้า"
      End
   End
End
Attribute VB_Name = "FrmQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String
Dim vStatus101 As Integer
Dim vIsReceived101 As Integer
Dim vARName101 As String
Dim vARName102 As String
Dim vSaleName101 As String
Dim vPicker101 As String
Dim vWHCode101 As String
Dim vRefNo101 As String
Dim vDiffDateTime101 As Currency
Dim vDocType101 As Integer
Dim vTimeID101 As Integer
Dim vDocDate101 As String

Dim vIndex As Integer

Private Sub CHKDate_Click()
Me.TextScan101.SetFocus
End Sub

Private Sub CMBQtyEdit_Click()
Dim vRecordset As New Recordset
Dim vAnswer As Integer
Dim vEditQty As Integer
Dim vItemCode As String
Dim vUnitCode As String
Dim vDocno As String
Dim vDocDate As String
Dim vTimeID As Integer
Dim vIsReceived As Integer
Dim vStopDateTime As String
Dim vCheckQty As Double

On Error GoTo ErrDescription

vAnswer = MsgBox("ต้องการเปลี่ยนจำนวนจัดได้ใช่หรือไม่", vbYesNo, "Send Question Message ?")
If vAnswer = 6 Then
   vDocno = Me.TextScan101.Text
   vDocDate = Me.LBLDocDate.Caption
   vTimeID = Me.LBLTimeID101.Caption
   vItemCode = Me.LBLItemCodeEdit.Caption
   vUnitCode = Me.LBLUnitCode.Caption
   vEditQty = Me.TBQtyEdit.Text
   vCheckQty = Me.LBLQtyEdit.Caption
   
   If vEditQty > vCheckQty Then
      MsgBox "ไม่สามารถกรอกจำนวนจัดได้เกินจำนวนที่ต้องการได้", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
   vQuery = "exec dbo.USP_QM_CheckARReceive '" & vDocno & "','" & vDocDate & "'," & vTimeID & " "
   If OpenDataBase(sConnection, vRecordset, vQuery) <> 0 Then
      vIsReceived = Trim(vRecordset.Fields("isreceived").Value)
      vStopDateTime = Trim(vRecordset.Fields("stopdatetime").Value)
   End If
   vRecordset.Close
   
   If vIsReceived = 1 Then
      MsgBox "เลขที่ใบจัดคิวใบนี้ ลูกค้าได้รับสินค้าไปแล้วไม่สามารถแก้ไขได้", vbCritical, "Send Error Message"
      Exit Sub
   End If

   vQuery = "exec dbo.USP_QM_UpdatePickQty '" & vDocno & "','" & vDocDate & "'," & vTimeID & ",'" & vItemCode & "','" & vUnitCode & "'," & vEditQty & "  "
   vConnection.Execute (vQuery)
   
   Me.ListView105.ListItems(vIndex).SubItems(7) = Format(vEditQty, "##,##0.00")

   MsgBox "แก้ไขจำนวนการจัดสินค้าให้เรียบร้อยแล้ว", vbInformation, "Send Error Message"
   Call RefreshQueueFinish

   Me.LBLItemCodeEdit.Caption = ""
   Me.LBLUnitCode.Caption = ""
   Me.TBQtyEdit.Text = ""
   Me.LBLQtyEdit.Caption = ""
   
   Me.PicEditQty.Visible = False
End If
Me.PicEditQty.Visible = False

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDClose_Click()
On Error GoTo ErrDescription

Call StartTime
Call ClearCheckQueue
PicScanBar.Visible = False
Me.Text101.SetFocus

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDClose_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 27 Then
 Call CMDClose_Click
End If

If LBLStatus101.Caption <> "" Then
  If KeyCode = 116 Then
    Call StopTime
    Select Case vStatus101
    Case 0
      Call InsertPicker
    Case 1
      Call ConfirmPickItemCode
    Case 2
      Select Case vIsReceived101
      Case 0
        Call CheckIsReceived
      End Select
    End Select
  Call ClearCheckQueue
  PicScanBar.Visible = False
  Me.Text101.SetFocus
  End If
End If
End Sub

Private Sub CMDCloseEditQty_Click()
Me.PicEditQty.Visible = False
End Sub

Private Sub CMDCustItem_Click()
Load FrmPayItemCust
Unload FrmQueue
End Sub

Private Sub CMDCustRecCancel_Click()
Dim vDocno As String
Dim vDocDate As String
Dim vAnswer As Integer

vAnswer = MsgBox("คุณต้องการยกเลิกการรับสินค้าของลูกค้าใช่หรือไม่", vbYesNo, "Send Question Message")
If vAnswer = 6 Then
   vDocno = Trim(TextScan101.Text)
   If Me.CHKDate.Value = 0 Then
   vDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   Else
   vDocDate = DateAdd("d", Now, -1)
   End If
   
   vQuery = "exec dbo.USP_NP_UpdateQueueCustRec '" & vDocno & "','" & vDocDate & "'"
   vConnection.Execute (vQuery)
End If
End Sub

Private Sub CMDDescription_Click()
Me.PTShowMyDescription.Visible = True
Me.TextMyDescription.SetFocus
End Sub

Private Sub CMDOK_Click()
On Error Resume Next

Call StartTime
Picture2.Visible = False
End Sub

Private Sub CMDOK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  Call CMDOK_Click
End If
End Sub

Private Sub CMDRefresh_Click()
  Call RefreshQueueBegin
  Call RefreshQueuePicking
  Call RefreshQueueFinish
  'FrmPicker.Show
  
  vCheckClickListview = 1
  MsgBox "ฟื้นฟูข้อมูลเรียบร้อยแล้ว", vbInformation, "Send Information"
End Sub

Private Sub CMDRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
  Call StopTime
  PicScanBar.Visible = True
  TextScan101.SetFocus
End If
End Sub

Private Sub CMDSendExit_Click()
  Dim i As Integer
  
  On Error Resume Next
  
  Me.Picture6.Visible = False
  Call StartTime
If Me.ListView104.ListItems.Count > 0 Then
   For i = 1 To Me.ListView104.ListItems.Count
           If Me.ListView104.ListItems(i).Checked = True Then
              Me.ListView104.ListItems(i).Checked = False
           End If
   Next
End If
End Sub

Private Sub CMDSendResult_Click()
    Dim vStatus As Integer
    Dim vPickingNo As String
    Dim vSaleOrderNo As String
    Dim vDescription As String
    
    On Error Resume Next
    
    If OPT101.Value = True Then
       vStatus = 1
    ElseIf OPT102.Value = True Then
       vStatus = 2
    ElseIf OPT103.Value = True Then
       vStatus = 3
    ElseIf OPT104.Value = True Then
       vStatus = 4
    End If
    
    vPickingNo = Me.LBLSendQueue.Caption
    vSaleOrderNo = Me.LBLSendRefNo.Caption
    vDescription = Me.TextDescription.Text
    
    vQuery = "exec dbo.USP_NP_UpdateQueueReceivedStatus1 '" & vPickingNo & "','" & vSaleOrderNo & "'," & vStatus & " ,'" & vDescription & "' "
    vConnection.Execute vQuery
    Me.TextDescription.Text = ""
    Me.Pic101.Visible = True
    Me.Pic102.Visible = False
    Call RefreshQueueFinish
    ListView104.SetFocus
    
    Me.Picture6.Visible = False
    Call StartTime
    Call RefreshData
End Sub

Private Sub CMDUpdateMyDescription_Click()
Dim vDocno As String
Dim vDocDate As String
Dim vMyDescription As String

If Me.TextMyDescription.Text <> "" Then
   vDocno = Trim(TextScan101.Text)
   If Me.CHKDate.Value = 0 Then
   vDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   Else
   vDocDate = DateAdd("d", Now, -1)
   End If
   vMyDescription = Me.TextMyDescription.Text
   vQuery = "exec dbo.USP_NP_UpdateQueueMyDescription '" & vDocno & "','" & vDocDate & "','" & vMyDescription & "' "
   vConnection.Execute (vQuery)
End If
Me.PTShowMyDescription.Visible = False
End Sub

Private Sub Command1_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListPicker As ListItem
Dim vConnectionString As String
Dim conn As New ADODB.Connection



vConnectionString = "Provider = SQLOLEDB.1;Data Source = Nebula;Initial Catalog = BPLUS4;User ID =VBUSER;PassWord = 132"
conn.Open vConnectionString
vQuery = "exec bcnp.dbo.USP_HR_PickerZone " & vSelectZoneID & ""
vRecordset.Open vQuery, conn, adOpenDynamic, adLockOptimistic
    If Not vRecordset.EOF Then
    vRecordset.MoveFirst
        While Not vRecordset.EOF
            MsgBox Trim(vRecordset.Fields("picker").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close

End Sub

Private Sub Form_Load()
On Error Resume Next

  Call RefreshQueueBegin
  Call RefreshQueuePicking
  Call RefreshQueueFinish
  vCheckClickListview = 1
End Sub

Private Sub ListView101_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim n As Integer
Dim vCheckQueue As String
Dim vQueueID As String
Dim vIndex As Integer
Dim vCustName As String
Dim vRefNo As String
Dim vQuery As String
Dim vPickingNo As String
Dim vSaleOrderNo As String
Dim vDocType As String
Dim vTimes As Integer
Dim vListItem As ListItem
Dim vDocDate As String

On Error Resume Next

If ListView101.ListItems.Count > 0 Then
  Call StopTime

  vIndex = ListView101.SelectedItem.Index
  vQueueID = ListView101.ListItems.Item(vIndex).Text

  If vIndex <> 0 Then
    vIndexBegin = vIndex
    FrmPicker.Show
    vPrintDocno = Trim(ListView101.ListItems.Item(vIndex).Text)
    vTimeID = Trim(ListView101.ListItems.Item(vIndex).SubItems(6))
    vCustName = Trim(ListView101.ListItems.Item(vIndex).SubItems(5))
    vRefNo = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
    vDocType = Trim(ListView101.ListItems.Item(vIndex).SubItems(4))
    vDocDate = Trim(ListView101.ListItems.Item(vIndex).SubItems(8))
    
    FrmPicker.LBLDocno.Caption = vPrintDocno
    FrmPicker.LBLID.Caption = vTimeID
    FrmPicker.LBLCustName.Caption = vCustName
    FrmPicker.LBLRefNo.Caption = vRefNo
    FrmPicker.LBLDocDate.Caption = vDocDate
    vCheckClickListview = 1
    
i = 0
'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails2
vQuery = "exec dbo.USP_NP_SearchQueueItemDetails5 '" & vPrintDocno & "','" & vRefNo & "'," & vDocType & "," & vTimeID & ",'" & vDocDate & "' "
If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst

While Not vRecordset.EOF
i = i + 1
Set vListItem = FrmPicker.ListView102.ListItems.Add(, , i)
  vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
  vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
  vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
  vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
  vListItem.SubItems(5) = Trim(vRecordset.Fields("shelfcode").Value)
  vListItem.SubItems(6) = Trim(vRecordset.Fields("familycode").Value)
  vRecordset.MoveNext
Wend
End If
vRecordset.Close

  Else
    MsgBox "กรอกเลขที่คิวไม่ถูกต้อง กรุณาตรวจสอบ", vbCritical, "Send Error"
  Exit Sub
  End If
End If
End Sub

Private Sub ListView101_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'Dim vIndex As Integer

'vIndex = Item.Index
'vIndexBegin = Item.Index
'FrmPicker.Show
'vPrintDocno = Trim(ListView101.ListItems.Item(vIndex).SubItems(1))
'vTimeID = Trim(ListView101.ListItems.Item(vIndex).SubItems(6))
'FrmPicker.LBLDocno.Caption = vPrintDocno
'FrmPicker.LBLID.Caption = vTimeID
'vCheckClickListview = 1
End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim vARName As String
Dim vRefNo As String
Dim vSaleCode As String
Dim vCustomerZone As String
Dim vQueueDate As String
Dim vCheckTypePick As Integer

On Error Resume Next

If Me.ListView101.ListItems.Count > 0 Then
 Select Case Trim(ListView101.SelectedItem.ListSubItems(7))
 Case 0
   vCustomerZone = "ลูกค้ารอรับของตามจุดออกใบหยิบ"
 Case 1
   vCustomerZone = "ลูกค้ารอรับของฝั่ง : สำนักงานใหญ่ "
 Case 2
   vCustomerZone = "ลูกค้ารอรับของฝั่ง : OutLet"
 End Select
 
 vARName = Trim(ListView101.SelectedItem.ListSubItems(5))
 vRefNo = Trim(ListView101.SelectedItem.ListSubItems(3))
 vSaleCode = Trim(ListView101.SelectedItem.ListSubItems(1))
 vQueueDate = Trim(ListView101.SelectedItem.ListSubItems(8))
 vCheckTypePick = Trim(ListView101.SelectedItem.ListSubItems(9))
 
 If vCheckTypePick = 0 Then
    Me.PICReserve.Visible = False
    Me.PicPay.Visible = True
 ElseIf vCheckTypePick = 1 Then
    Me.PICReserve.Visible = True
    Me.PicPay.Visible = False
 End If
 
 
 Pic101.Visible = False
 Pic102.Visible = False
 
 LBLCustomerZone.Caption = vCustomerZone
 LBLQueueDate = vQueueDate
 LBLARName.Caption = vARName
 LBLRefNo.Caption = vRefNo
 LBLUserPick.Caption = Trim("-")
 LBLSale.Caption = vSaleCode
 vCheckClickListview = 1
End If
End Sub

Private Sub ListView101_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim vPickingNo As String
Dim vSaleOrderNo As String
Dim vDocType As Integer
Dim vDocDate As String
Dim i As Integer
Dim vARName As String
Dim vRefNo As String
Dim vPicker As String
Dim vSaleCode As String
Dim vTimes As Integer

On Error Resume Next

If ListView101.ListItems.Count > 0 Then
vCheckClickListview = 1
  If KeyCode = 112 Then
    ListView103.ListItems.Clear
    LBLDocno1 = ""
    LBLDocno2 = ""
    LBLPicker.Caption = ""
    LBLWHCode.Caption = ""
    Picture2.Visible = True
    Call StopTime
    
    vIndex = ListView101.SelectedItem.Index
    LBLDocno1 = Trim(ListView101.ListItems.Item(vIndex).Text)
    LBLDocno2 = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
    vDocType = Trim(ListView101.ListItems.Item(vIndex).SubItems(4))
    vDocDate = Trim(ListView101.ListItems.Item(vIndex).SubItems(8))
    vPickingNo = Trim(LBLDocno1.Caption)
    vSaleOrderNo = Trim(LBLDocno2.Caption)
    vTimes = Trim(ListView101.ListItems.Item(vIndex).SubItems(6))
    i = 0
    'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails2 '" & vPrintDocno & "','" & vRefNo & "'," & vDocType & "," & vTimeID & ",'" & vDocDate & "' "
    'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails3 '" & vPickingNo & "','" & vSaleOrderNo & "'," & vDocType & "," & vTimes & ",'" & vDocDate & "' "
    vQuery = "exec dbo.USP_NP_SearchQueueItemDetails5 '" & vPickingNo & "','" & vSaleOrderNo & "'," & vDocType & "," & vTimeID & ",'" & vDocDate & "' "
    
    If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    LBLPicker.Caption = Trim(vRecordset.Fields("picker").Value)
    LBLWHCode.Caption = Trim(vRecordset.Fields("whcode").Value)
    While Not vRecordset.EOF
      i = i + 1
      Set vListItem = ListView103.ListItems.Add(, , i)
      vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
      vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
      vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
      vListItem.SubItems(5) = Trim(vRecordset.Fields("shelfcode").Value)
      vRecordset.MoveNext
      Wend
    End If
    vRecordset.Close
    ListView103.SetFocus
    ElseIf KeyCode = 116 Then
    Call StopTime
    PicScanBar.Visible = True
    TextScan101.SetFocus
    End If
End If
End Sub

Private Sub ListView101_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

If ListView101.ListItems.Count <> 0 Then
    If Button = 2 Then
        Popup_Menu Menu1
    End If
End If
End Sub

Private Sub ListView102_DblClick()
Dim vIndex As Integer
Dim vListItem As ListItem
Dim vSaleOrderNo As String
Dim vRecordset As New ADODB.Recordset
Dim vARCode As String
Dim vDocType As Integer
Dim vDocDate As String
Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim vCheckQueue As String
Dim vQueueID As String

On Error Resume Next

If ListView102.ListItems.Count > 0 Then
  Call StopTime
  vIndex = ListView102.SelectedItem.Index
  vQueueID = ListView102.ListItems.Item(vIndex).Text
    
  If vIndex <> 0 Then
    vCheckClickListview = 2
    vIndexFinish = vIndex
    vPrintDocno = Trim(ListView102.ListItems.Item(vIndex).Text)
    vTimeID = Trim(ListView102.ListItems.Item(vIndex).SubItems(6))
    vSaleOrderNo = Trim(ListView102.ListItems.Item(vIndex).SubItems(3))
    vARCode = Trim(ListView102.ListItems.Item(vIndex).SubItems(5))
    vDocType = Trim(ListView102.ListItems.Item(vIndex).SubItems(4))
    vDocDate = Trim(ListView102.ListItems.Item(vIndex).SubItems(11))
    FrmCheckQTY.Show
    FrmCheckQTY.LBLDocno1.Caption = vPrintDocno
    FrmCheckQTY.LBLDocno2.Caption = vSaleOrderNo
    FrmCheckQTY.LBLARCode.Caption = vARCode
    FrmCheckQTY.LBLID.Caption = vTimeID
    FrmCheckQTY.LBLDocDate.Caption = vDocDate
    
    If vPrintDocno <> "" And vSaleOrderNo <> "" Then
    i = 0
    'ListView101.ListItems.Clear
    'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails2
    vQuery = "exec dbo.USP_NP_SearchQueueItemDetails5 '" & vPrintDocno & "','" & vSaleOrderNo & "'," & vDocType & "," & vTimeID & ",'" & vDocDate & "' "
    If OpenDataBase2(qConnection, vRecordset, vQuery) <> 0 Then
      vRecordset.MoveFirst
      While Not vRecordset.EOF
        i = i + 1
        Set vListItem = FrmCheckQTY.ListView101.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
          vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
          vListItem.SubItems(4) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
          vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
          vListItem.SubItems(6) = Trim("F1")
          vListItem.SubItems(7) = Trim(vRecordset.Fields("whcode").Value)
      vRecordset.MoveNext
      Wend
    End If
    vRecordset.Close
    End If
  Else
    MsgBox "กรอกเลขที่คิวไม่ถูกต้อง กรุณาตรวจสอบ", vbCritical, "Send Error"
  End If
End If
End Sub

Private Sub ListView103_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  Call CMDOK_Click
End If
End Sub

Private Sub ListView105_DblClick()
Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Double
Dim vPickQTY As Double
Dim vUnitCode As String

On Error GoTo ErrDescription

If vStatus101 = 2 And vIsReceived101 = 0 Then
   If Me.ListView105.ListItems.Count > 0 Then
      vIndex = Me.ListView105.SelectedItem.Index
      Me.PicEditQty.Visible = True
      vItemCode = Me.ListView105.ListItems(vIndex).SubItems(1)
      vItemName = Me.ListView105.ListItems(vIndex).SubItems(2)
      vPickQTY = Me.ListView105.ListItems(vIndex).SubItems(7)
      vQTY = Me.ListView105.ListItems(vIndex).SubItems(6)
      vUnitCode = Me.ListView105.ListItems(vIndex).SubItems(4)
      
      Me.LBLItemCodeEdit.Caption = vItemCode
      Me.LBLItemNameEdit.Caption = vItemName
      Me.LBLQtyEdit.Caption = vQTY
      Me.TBQtyEdit.Text = vPickQTY
      Me.LBLUnitCode.Caption = vUnitCode
      Me.TBQtyEdit.SetFocus
   End If
Else
   MsgBox "คิวที่สามารถแก้ไขจำนวนการจัดได้นั้นต้องทำการบันทึกการจัดมาก่อนและต้องไม่เป็นคิวที่ลูกค้ารับของแล้ว", vbCritical, "Send Error Message"
   Me.ListView105.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub ListView105_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
   Call vQueueID
End If
End Sub

Private Sub MReserve_Click()
Dim i As Integer
Dim vAnswer As Integer
Dim vDocno As String
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSaleOrderNo As String
Dim vWHCode As String
Dim vShelfGroup As String
Dim vZoneID As String
Dim vDocType As Integer
Dim vFamilyCode As String
Dim vTimeID As Integer
Dim vCustomerZone As Integer
Dim vPickZone As String
Dim vJobID As Integer

On Error GoTo ErrDescription

i = ListView101.SelectedItem.Index
vDocno = Trim(ListView101.ListItems.Item(i).Text)
vTimeID = ListView101.ListItems.Item(i).SubItems(6)
vAnswer = MsgBox("ต้องการพิมพ์ทดแทนคิวที่  " & vDocno & " ใช่หรือไม่", vbYesNo, "Send Question ?")
If vAnswer = 6 Then

vSaleOrderNo = ListView101.ListItems.Item(i).SubItems(3)
vWHCode = ListView101.ListItems.Item(i).SubItems(12)
vShelfGroup = ListView101.ListItems.Item(i).SubItems(13)
vZoneID = ListView101.ListItems.Item(i).SubItems(10)
vFamilyCode = ListView101.ListItems.Item(i).SubItems(11)
vDocType = ListView101.ListItems.Item(i).SubItems(4)
vCustomerZone = ListView101.ListItems.Item(i).SubItems(7)
vPickZone = ListView101.ListItems.Item(i).SubItems(14)

 
 If vDocType = 1 Then
    'Call PrintPickingSlipRes(vSaleOrderNo, vWHCode, vShelfGroup, vZoneID, vFamilyCode, vPickZone, vCustomerZone, vTimeID)
    
  vJobID = 1
  vQuery = "exec dbo.USP_NP_InsertPrintTermal " & vJobID & ",'" & vSaleOrderNo & "','" & vDocno & "','" & vWHCode & "','" & vShelfGroup & "','" & vFamilyCode & "','" & vZoneID & "','" & vPickZone & "','" & vUserID & "' "
  vConnection.Execute vQuery
  
  
    MsgBox "พิมพ์ทดแทนเรียบร้อย"
ElseIf vDocType = 2 Then
    'Call PrintDetails(vSaleOrderNo, vZoneID, vFamilyCode, vPickZone)
    
  vJobID = 2
  vQuery = "exec dbo.USP_NP_InsertPrintTermal " & vJobID & ",'" & vSaleOrderNo & "','" & vDocno & "','" & vWHCode & "','" & vShelfGroup & "','" & vFamilyCode & "','" & vZoneID & "','" & vPickZone & "','" & vUserID & "' "
  vConnection.Execute vQuery
  
ElseIf vDocType = 3 Then
    Call PrintRequestPickingSlip(vSaleOrderNo, vZoneID, vFamilyCode, vPickZone)
End If

Else
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Public Sub PrintPickingSlipRes(vSaleOrder As String, vWHCode As String, vShelfGroup As String, vZoneID As String, vFamilyGroup As String, vPickZone As String, vCustomerZone As Integer, vCount As Integer)
Dim vRecordset1 As New Recordset
Dim vQuery As String
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vDocno As String
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer
Dim vItemName As String
Dim vSoStatus As Integer
Dim vSelectPicked As Integer
Dim vGroupDocNo As String
Dim vPrinterID As Integer


If UCase(Left(vSaleOrder, 3)) = "S01" Or UCase(Left(vSaleOrder, 3)) = "S02" Or UCase(Left(vSaleOrder, 3)) = "W01" Or UCase(Left(vSaleOrder, 3)) = "W02" Then
vGroupDocNo = UCase(Left(Right(vSaleOrder, Len(vSaleOrder) - InStr(vSaleOrder, "-")), 3))
Else
vGroupDocNo = UCase(Left(vSaleOrder, 3))
End If

vQuery = "exec dbo.USP_NP_SearchPrinterPrintZone '" & vPickZone & "' "
If OpenDataBase(vConnection, vRecordset1, vQuery) <> 0 Then
vPrinterName = Trim(vRecordset1.Fields("printername").Value)
End If
vRecordset1.Close

For Each printerObj In Printers
If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
Set Printer = printerObj
Set printerObj = Nothing
Exit For
End If
Next


    vQuery = "exec dbo.USP_SO_PickingQueueFreedom2 '" & vSaleOrder & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "','" & vFamilyGroup & "','" & vPickZone & "'," & vCount & " "
    If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    
      vSoStatus = vRecordset.Fields("sostatus").Value
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")


      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 20
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 80
      Printer.Print "(จ่าย)"

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.FontBold = True
      Printer.CurrentX = 1600
      Printer.CurrentY = 0
      Printer.Print Trim(vRecordset.Fields("queueno").Value)
      
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 1200
      Printer.CurrentY = 1000
      Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("_______________________________________________________________________________________")
      
      
     If vSoStatus = 1 And vSelectPicked = 0 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 12
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ทดแทนต้นฉบับใบจัดสินค้า เพื่อกำกับสินค้าจอง")
      ElseIf vSoStatus = 1 And vSelectPicked = 1 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 12
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ทดแทนต้นฉบับใบจัดสินค้า เพื่อจ่ายสินค้า")
      ElseIf vSoStatus = 0 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 12
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ทดแทนต้นฉบับใบจัดสินค้า เพื่อจ่ายสินค้า")
      End If
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 1900
      Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 2000
      Printer.CurrentY = 1900
      Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      

      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 3150
      If vRecordset.Fields("isconditionsend").Value = 0 Then
            Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("รับเอง")
      Else
            Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("ส่งให้")
      End If
                  
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 16
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 3400
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value)

      If Trim(vRecordset.Fields("carlicense").Value) <> "" Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 16
        Printer.CurrentX = 1400
        Printer.CurrentY = 3400
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.Print Trim("ทะเบียนรถขนส่ง : ") & Trim(vRecordset.Fields("carlicense").Value)
      End If
            
      If vSoStatus = 0 Then
         Printer.Font.Name = "AngsanaUPC"
         Printer.Font.Size = 14
         Printer.FontBold = True
         Printer.FontUnderline = False
         Printer.CurrentX = 0
         Printer.CurrentY = 3800
         Printer.Print Trim("เวลารับของ : ") & Trim(vRecordset.Fields("requesttime").Value)
      Else
         Printer.Font.Name = "AngsanaUPC"
         Printer.Font.Size = 14
         Printer.FontBold = True
         Printer.FontUnderline = False
         Printer.CurrentX = 0
         Printer.CurrentY = 3800
         Printer.Print Trim("วันที่ครบกำหนดรับของ : ") & Trim(vRecordset.Fields("duedate").Value)
      End If
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 14
      Printer.CurrentX = 0
      Printer.CurrentY = 4150
      Printer.Print Trim(vRecordset.Fields("customerzone").Value)
      
      vRecordset.MoveFirst
      vLineX = 50
      vLineY = 50
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY - 30
      Printer.FontBold = False
      Printer.FontUnderline = False
      Printer.Print Trim("-----------------------------------------------------------------------------------------------")
      n = 1
      While Not vRecordset.EOF
          Printer.Font.Size = 18
          Printer.FontBold = True
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfcode1").Value)
          
          Printer.Font.Size = 11
          Printer.FontBold = False
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print Trim("OnHand") & "(" & Trim(vRecordset.Fields("shelfcode").Value) & ")" & ": " & Trim(vRecordset.Fields("qtylocation").Value) & "  " & Trim(vRecordset.Fields("stkunitcode").Value) & "     " & "รวมคลัง :" & Trim(vRecordset.Fields("StkWHCode").Value) & "  " & Trim(vRecordset.Fields("stkunitcode").Value)
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value) & "             " & " ขายชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)
          
          vItemName = Trim(vRecordset.Fields("itemname").Value) & Trim(vRecordset.Fields("descriptionline"))
          If Len(vItemName) <= 55 Then
             Printer.CurrentX = 0
             Printer.CurrentY = Printer.CurrentY
             Printer.Print "ชื่อสินค้า :" & vItemName
          Else
             Printer.CurrentX = 0
             Printer.CurrentY = Printer.CurrentY
             Printer.Print "ชื่อสินค้า :" & Left(vItemName, 55)
             
             Printer.CurrentX = 600
             Printer.CurrentY = Printer.CurrentY
             Printer.Print Right(vItemName, Len(vItemName) - 55)
          End If
          
          Printer.Font.Size = 13
          Printer.CurrentX = Printer.CurrentX + 15
          Printer.CurrentY = Printer.CurrentY + 100
          Printer.FontBold = True
          Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY - 50
          Printer.FontBold = False
          Printer.Print Trim("-----------------------------------------------------------------------------------------------")
          
      vRecordset.MoveNext
      n = n + 1
      Wend
    End If
    vRecordset.Close
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 400
    Printer.Print Trim("_______________________________________________________________________________________________")
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + vLineY
    Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 150
    Printer.Print "         _____________                                    ______________"
          
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    
    Printer.Font.Size = 10
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("วันที่พิมพ์ :") & Now & "          " & vPrinterName
    Printer.EndDoc
End Sub


Private Sub Popup_Menu(m As Menu)
    Menu1.Visible = True
    PopupMenu m, 2
    Menu1.Visible = False
End Sub

Private Sub ListView102_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim vARName As String
Dim vRefNo As String
Dim vPicker As String
Dim vSaleCode As String
Dim vCustomerZone As String
Dim vQueueDate As String
Dim vCheckTypePick As Integer

On Error Resume Next

If Me.ListView102.ListItems.Count > 0 Then
 vCheckClickListview = 2
 Select Case Trim(ListView102.SelectedItem.ListSubItems(10))
 Case 0
   vCustomerZone = "ลูกค้ารอรับของตามจุดออกใบหยิบ"
 Case 1
   vCustomerZone = "ลูกค้ารอรับของฝั่ง : สำนักงานใหญ่ "
 Case 2
   vCustomerZone = "ลูกค้ารอรับของฝั่ง : OutLet"
 End Select
 
 Pic101.Visible = False
 Pic102.Visible = False
 
 vARName = Trim(ListView102.SelectedItem.ListSubItems(5))
 vRefNo = Trim(ListView102.SelectedItem.ListSubItems(3))
 vPicker = Trim(ListView102.SelectedItem.ListSubItems(7))
 vSaleCode = Trim(ListView102.SelectedItem.ListSubItems(8))
 vQueueDate = Trim(ListView102.SelectedItem.ListSubItems(11))
 vCheckTypePick = Trim(ListView102.SelectedItem.ListSubItems(12))
 
 If vCheckTypePick = 0 Then
    Me.PICReserve.Visible = False
    Me.PicPay.Visible = True
 ElseIf vCheckTypePick = 1 Then
    Me.PICReserve.Visible = True
    Me.PicPay.Visible = False
 End If
  
LBLQueueDate.Caption = vQueueDate
 LBLARName.Caption = vARName
 LBLRefNo.Caption = vRefNo
 LBLUserPick.Caption = vPicker
 LBLSale.Caption = vSaleCode
 LBLCustomerZone.Caption = vCustomerZone
End If
End Sub

Private Sub ListView102_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim vPickingNo As String
Dim vSaleOrderNo As String
Dim vDocType As Integer
Dim vDocDate As String
Dim i As Integer
Dim vARName As String
Dim vRefNo As String
Dim vPicker As String
Dim vSaleCode As String
Dim vTimes As Integer

On Error Resume Next

If ListView102.ListItems.Count > 0 Then
vCheckClickListview = 2
  If KeyCode = 112 Then
    ListView103.ListItems.Clear
    LBLDocno1 = ""
    LBLDocno2 = ""
    LBLPicker.Caption = ""
    LBLWHCode.Caption = ""
    Picture2.Visible = True
    Call StopTime
    
    vIndex = ListView102.SelectedItem.Index
    LBLDocno1 = Trim(ListView102.ListItems.Item(vIndex).Text)
    LBLDocno2 = Trim(ListView102.ListItems.Item(vIndex).SubItems(3))
    vDocType = Trim(ListView102.ListItems.Item(vIndex).SubItems(4))
    vDocDate = Trim(ListView102.ListItems.Item(vIndex).SubItems(11))
    vPickingNo = Trim(LBLDocno1.Caption)
    vSaleOrderNo = Trim(LBLDocno2.Caption)
    vTimes = Trim(ListView102.ListItems.Item(vIndex).SubItems(6))
    i = 0
    'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails2
    'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails3 '" & vPickingNo & "','" & vSaleOrderNo & "'," & vDocType & " ," & vTimes & ",'" & vDocDate & "' "
    vQuery = "exec dbo.USP_NP_SearchQueueItemDetails5 '" & vPickingNo & "','" & vSaleOrderNo & "'," & vDocType & " ," & vTimes & ",'" & vDocDate & "' "
    If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    LBLPicker.Caption = Trim(vRecordset.Fields("picker").Value)
    LBLWHCode.Caption = Trim(vRecordset.Fields("whcode").Value)
    While Not vRecordset.EOF
      i = i + 1
      Set vListItem = ListView103.ListItems.Add(, , i)
      vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
      vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
      vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
      vListItem.SubItems(5) = Trim(vRecordset.Fields("shelfcode").Value)
      vRecordset.MoveNext
      Wend
    End If
    vRecordset.Close
    ListView103.SetFocus
    ElseIf KeyCode = 116 Then
    Call StopTime
    PicScanBar.Visible = True
    TextScan101.SetFocus
    End If
End If

End Sub

Private Sub ListView104_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAnswer As Integer
Dim vPickingNo As String
Dim vIndex As Integer
Dim vSaleOrderNo As String
Dim vDocType As Integer

On Error Resume Next

If ListView104.ListItems.Count > 0 Then
  vIndexComplete = Item.Index
  vSaleOrderNo = Trim(ListView104.ListItems.Item(vIndexComplete).SubItems(2))
  vPickingNo = ListView104.ListItems.Item(vIndexComplete).Text
  
  Me.LBLSendQueue.Caption = vPickingNo
  Me.LBLSendRefNo.Caption = vSaleOrderNo
  Me.Picture6.Visible = True
  Me.OPT101.SetFocus
  Call StopTime
End If
End Sub

Private Sub ListView104_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim vARName As String
Dim vRefNo As String
Dim vSaleCode As String
Dim vPicker As String
Dim vCustomerZone As String
Dim vQueueDate As String
Dim vCheckTypePick As Integer

On Error Resume Next

Select Case Trim(ListView104.SelectedItem.ListSubItems(7))
Case 0
  vCustomerZone = "ลูกค้ารอรับของตามจุดออกใบหยิบ"
Case 1
  vCustomerZone = "ลูกค้ารอรับของฝั่ง : สำนักงานใหญ่ "
Case 2
  vCustomerZone = "ลูกค้ารอรับของฝั่ง : OutLet"
End Select
 
 If Trim(ListView104.SelectedItem.ListSubItems(8)) = 1 Then
   Pic101.Visible = True
   Pic102.Visible = False
 ElseIf Trim(ListView104.SelectedItem.ListSubItems(8)) = 2 Then
   Pic101.Visible = False
   Pic102.Visible = True
 Else
   Pic101.Visible = False
   Pic102.Visible = False
 End If
 
vARName = Trim(ListView104.SelectedItem.ListSubItems(1))
vRefNo = Trim(ListView104.SelectedItem.ListSubItems(2))
vSaleCode = Trim(ListView104.SelectedItem.ListSubItems(5))
vPicker = Trim(ListView104.SelectedItem.ListSubItems(3))
vQueueDate = Trim(ListView104.SelectedItem.ListSubItems(9))
vCheckTypePick = Trim(ListView104.SelectedItem.ListSubItems(10))
 
 If vCheckTypePick = 0 Then
    Me.PICReserve.Visible = False
    Me.PicPay.Visible = True
 ElseIf vCheckTypePick = 1 Then
    Me.PICReserve.Visible = True
    Me.PicPay.Visible = False
 End If
 

LBLQueueDate.Caption = vQueueDate
LBLARName.Caption = vARName
LBLRefNo.Caption = vRefNo
LBLUserPick.Caption = vPicker
LBLSale.Caption = vSaleCode
LBLCustomerZone.Caption = vCustomerZone
vCheckClickListview = 3
End Sub

Private Sub ListView104_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim vPickingNo As String
Dim vSaleOrderNo As String
Dim vDocType As Integer
Dim vDocDate As String
Dim i As Integer
Dim vARName As String
Dim vRefNo As String
Dim vPicker As String
Dim vSaleCode As String
Dim vTimes As Integer

On Error Resume Next

If ListView104.ListItems.Count > 0 Then
vCheckClickListview = 2
  If KeyCode = 112 Then
    ListView103.ListItems.Clear
    LBLDocno1 = ""
    LBLDocno2 = ""
    LBLPicker.Caption = ""
    LBLWHCode.Caption = ""
    Picture2.Visible = True
    Call StopTime
    
    vIndex = ListView104.SelectedItem.Index
    LBLDocno1 = Trim(ListView104.ListItems.Item(vIndex).Text)
    LBLDocno2 = Trim(ListView104.ListItems.Item(vIndex).SubItems(2))
    vDocType = Trim(ListView104.ListItems.Item(vIndex).SubItems(6))
    vDocDate = Trim(ListView104.ListItems.Item(vIndex).SubItems(9))
    vPickingNo = Trim(LBLDocno1.Caption)
    vSaleOrderNo = Trim(LBLDocno2.Caption)
    vTimes = Trim(ListView104.ListItems.Item(vIndex).SubItems(4))
    i = 0
    'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails2
    'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails3 '" & vPickingNo & "','" & vSaleOrderNo & "'," & vDocType & " ," & vTimes & ",'" & vDocDate & "' "
    vQuery = "exec dbo.USP_NP_SearchQueueItemDetails5 '" & vPickingNo & "','" & vSaleOrderNo & "'," & vDocType & " ," & vTimes & ",'" & vDocDate & "' "
    If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    LBLPicker.Caption = Trim(vRecordset.Fields("picker").Value)
    LBLWHCode.Caption = Trim(vRecordset.Fields("whcode").Value)
    While Not vRecordset.EOF
      i = i + 1
      Set vListItem = ListView103.ListItems.Add(, , i)
      vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
      vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
      vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
      vListItem.SubItems(5) = Trim(vRecordset.Fields("shelfcode").Value)
      vRecordset.MoveNext
      Wend
    End If
    vRecordset.Close
    ListView103.SetFocus
    ElseIf KeyCode = 116 Then
      Call StopTime
      PicScanBar.Visible = True
      TextScan101.SetFocus
    End If
End If

End Sub

Private Sub OPT101_KeyPress(KeyAscii As Integer)
Dim i As Integer
  
On Error Resume Next

If KeyAscii = 27 Then
   Me.Picture6.Visible = False
   Call StartTime
   If Me.ListView104.ListItems.Count > 0 Then
      For i = 1 To Me.ListView104.ListItems.Count
              If Me.ListView104.ListItems(i).Checked = True Then
                 Me.ListView104.ListItems(i).Checked = False
              End If
      Next
   End If
   Me.Text101.SetFocus
End If
End Sub

Private Sub PicScanBar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 27 Then
 Call CMDClose_Click
End If

If LBLStatus101.Caption <> "" Then
  If KeyCode = 116 Then
    Call StopTime
    Select Case vStatus101
    Case 0
      Call InsertPicker
    Case 1
      Call ConfirmPickItemCode
    Case 2
      Select Case vIsReceived101
      Case 0
        Call CheckIsReceived
      End Select
    End Select
  Call ClearCheckQueue
  PicScanBar.Visible = False
  Me.Text101.SetFocus
  End If
End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
  Call StopTime
  PicScanBar.Visible = True
  TextScan101.SetFocus
End If
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  Call CMDOK_Click
End If
End Sub

Private Sub Picture3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
  Call StopTime
  PicScanBar.Visible = True
  TextScan101.SetFocus
End If
End Sub

Private Sub Picture4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
  Call StopTime
  PicScanBar.Visible = True
  TextScan101.SetFocus
End If
End Sub

Private Sub Picture5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
  Call StopTime
  PicScanBar.Visible = True
  TextScan101.SetFocus
End If
End Sub

Private Sub Text101_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
  Call StopTime
  PicScanBar.Visible = True
  TextScan101.SetFocus
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim n As Integer
Dim vCheckQueue As String
Dim vQueueID As String
Dim vIndex As Integer
Dim vCustName As String
Dim vRefNo As String
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vPickingNo As String
Dim vSaleOrderNo As String
Dim vDocType As String
Dim vTimes As Integer
Dim vListItem As ListItem
Dim vDocDate As String

On Error Resume Next

If KeyAscii = 13 Then
  Call StopTime
  n = ListView101.ListItems.Count
  vQueueID = Trim(Text101.Text)
  
  For i = 1 To n
    vCheckQueue = Trim(ListView101.ListItems.Item(i).Text)
    If vQueueID = vCheckQueue Then
      vIndex = i
    End If
  Next i

  If vIndex <> 0 Then
    Call FrmQueue.StopTime
    vIndexBegin = vIndex
    FrmPicker.Show
    vPrintDocno = Trim(ListView101.ListItems.Item(vIndex).Text)
    vTimeID = Trim(ListView101.ListItems.Item(vIndex).SubItems(6))
    vCustName = Trim(ListView101.ListItems.Item(vIndex).SubItems(5))
    vRefNo = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
    vDocType = Trim(ListView101.ListItems.Item(vIndex).SubItems(4))
    vDocDate = Trim(ListView101.ListItems.Item(vIndex).SubItems(8))
    
    FrmPicker.LBLDocno.Caption = vPrintDocno
    FrmPicker.LBLID.Caption = vTimeID
    FrmPicker.LBLCustName.Caption = vCustName
    FrmPicker.LBLRefNo.Caption = vRefNo
    FrmPicker.LBLDocDate.Caption = vDocDate
    vCheckClickListview = 1
    
i = 0
'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails2
vQuery = "exec dbo.USP_NP_SearchQueueItemDetails3 '" & vPrintDocno & "','" & vRefNo & "'," & vDocType & "," & vTimeID & ",'" & vDocDate & "' "
If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst

While Not vRecordset.EOF
i = i + 1
Set vListItem = FrmPicker.ListView102.ListItems.Add(, , i)
  vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
  vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
  vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
  vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
  vListItem.SubItems(5) = Trim(vRecordset.Fields("shelfcode").Value)
  vRecordset.MoveNext
Wend
End If
vRecordset.Close

  Else
    MsgBox "กรอกเลขที่คิวไม่ถูกต้อง กรุณาตรวจสอบ", vbCritical, "Send Error"
  Exit Sub
  End If
End If
End Sub

Private Sub Text102_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
  Call StopTime
  PicScanBar.Visible = True
  TextScan101.SetFocus
End If
End Sub

Private Sub Text102_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer
Dim vListItem As ListItem
Dim vSaleOrderNo As String
Dim vRecordset As New ADODB.Recordset
Dim vARCode As String
Dim vDocType As Integer
Dim vDocDate As String
Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim vCheckQueue As String
Dim vQueueID As String

On Error Resume Next

If KeyAscii = 13 Then
  Call StopTime
  n = ListView102.ListItems.Count
  vQueueID = Trim(Text102.Text)
  For j = 1 To n
    vCheckQueue = Trim(ListView102.ListItems.Item(j).Text)
    If vQueueID = vCheckQueue Then
      vIndex = j
    End If
  Next j

  If vIndex <> 0 Then
    vCheckClickListview = 2
    vIndexFinish = vIndex
    vPrintDocno = Trim(ListView102.ListItems.Item(vIndex).Text)
    vTimeID = Trim(ListView102.ListItems.Item(vIndex).SubItems(6))
    vSaleOrderNo = Trim(ListView102.ListItems.Item(vIndex).SubItems(3))
    vARCode = Trim(ListView102.ListItems.Item(vIndex).SubItems(5))
    vDocType = Trim(ListView102.ListItems.Item(vIndex).SubItems(4))
    vDocDate = Trim(ListView102.ListItems.Item(vIndex).SubItems(11))
    FrmCheckQTY.Show
    FrmCheckQTY.LBLDocno1.Caption = vPrintDocno
    FrmCheckQTY.LBLDocno2.Caption = vSaleOrderNo
    FrmCheckQTY.LBLARCode.Caption = vARCode
    FrmCheckQTY.LBLID.Caption = vTimeID
    FrmCheckQTY.LBLDocDate.Caption = vDocDate
    
    If vPrintDocno <> "" And vSaleOrderNo <> "" Then
    i = 0
    'ListView101.ListItems.Clear
    'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails2
    vQuery = "exec dbo.USP_NP_SearchQueueItemDetails3 '" & vPrintDocno & "','" & vSaleOrderNo & "'," & vDocType & "," & vTimeID & ",'" & vDocDate & "' "
    If OpenDataBase2(qConnection, vRecordset, vQuery) <> 0 Then
      vRecordset.MoveFirst
      While Not vRecordset.EOF
        i = i + 1
        Set vListItem = FrmCheckQTY.ListView101.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
          vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
          vListItem.SubItems(4) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
          vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
          vListItem.SubItems(6) = Trim("F1")
          vListItem.SubItems(7) = Trim(vRecordset.Fields("whcode").Value)
      vRecordset.MoveNext
      Wend
    End If
    vRecordset.Close
    End If
  Else
    MsgBox "กรอกเลขที่คิวไม่ถูกต้อง กรุณาตรวจสอบ", vbCritical, "Send Error"
  End If
End If
End Sub

Private Sub TextScan101_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 27 Then
 Call CMDClose_Click
End If

If LBLStatus101.Caption <> "" Then
  If KeyCode = 116 Then
    Call StopTime
    Select Case vStatus101
                 Case 0
                            Call InsertPicker
                 Case 1
                            Call ConfirmPickItemCode
                 Case 2
                            Select Case vIsReceived101
                            Case 0
                                        Call CheckIsReceived
                            End Select
    End Select
    Call ClearCheckQueue
    PicScanBar.Visible = False
  End If
End If

If KeyCode = 8 Then
   Me.ListView105.ListItems.Clear
   Me.TextScan101.Text = ""
   Me.LBLARCode101.Caption = ""
   Me.LBLARCode102.Caption = ""
   Me.LBLPicker101.Caption = ""
   Me.LBLRefDocNo101.Caption = ""
   Me.LBLDocDate.Caption = ""
   Me.LBLSale101.Caption = ""
   Me.LBLTime101.Caption = ""
   Me.LBLWHCode101.Caption = ""
   Me.LBLStatus101.Caption = ""
   Me.TextScan101.SetFocus
   LBLTimeID101.Caption = ""
   LBLDocType101.Caption = ""
   vStatus101 = 0
   vIsReceived101 = 0
   vARName101 = ""
   vSaleName101 = ""
   vPicker101 = ""
   vWHCode101 = ""
   vRefNo101 = ""
   vDiffDateTime101 = 0
   vDocType101 = 0
   vTimeID101 = 0
End If
End Sub

Public Sub CheckIsReceived()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAnswer As Integer
Dim vPickingNo As String
Dim vIndex As Integer
Dim vSaleOrderNo As String
Dim vDocType As Integer
Dim vDocno As String

On Error Resume Next

vDocno = Trim(TextScan101.Text)

If vDocno <> "" Then
   vSaleOrderNo = Trim(LBLRefDocNo101.Caption)
   vPickingNo = vDocno
   
   Me.LBLSendQueue.Caption = vPickingNo
   Me.LBLSendRefNo.Caption = vSaleOrderNo
   Me.Picture6.Visible = True
   Me.OPT101.SetFocus
   Call StopTime
End If
End Sub
Public Sub InsertPicker()
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vPrintDocno As String
Dim vTimeID As Integer
Dim vCustName As String
Dim vRefNo As String
Dim vDocType As Integer
Dim vDocDate As String
Dim vListItem As ListItem

On Error GoTo ErrDescription

If TextScan101.Text <> "" Then
  FrmPicker.Show
  vPrintDocno = Trim(TextScan101.Text)
  vTimeID = Trim(LBLTimeID101.Caption)
  vCustName = Trim(LBLARCode101.Caption)
  vRefNo = Trim(LBLRefDocNo101.Caption)
  vDocType = Trim(LBLDocType101.Caption)
  vDocDate = Trim(LBLDocDate.Caption)
  
  FrmPicker.LBLDocno.Caption = vPrintDocno
  FrmPicker.LBLID.Caption = vTimeID
  FrmPicker.LBLCustName.Caption = vCustName
  FrmPicker.LBLRefNo.Caption = vRefNo
  FrmPicker.LBLDocDate.Caption = vDocDate
        
  i = 0
  'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails2
  vQuery = "exec dbo.USP_NP_SearchQueueItemDetails5 '" & vPrintDocno & "','" & vRefNo & "'," & vDocType & "," & vTimeID & ",'" & vDocDate & "' "
  If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
  vRecordset.MoveFirst
  
  While Not vRecordset.EOF
  i = i + 1
  Set vListItem = FrmPicker.ListView102.ListItems.Add(, , i)
    vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
    vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
    vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
    vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
    vListItem.SubItems(5) = Trim(vRecordset.Fields("shelfcode").Value)
    vListItem.SubItems(6) = Trim(vRecordset.Fields("familycode").Value)
    vRecordset.MoveNext
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
Public Sub ConfirmPickItemCode()
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vPrintDocno As String
Dim vTimeID As Integer
Dim vARCode As String
Dim vSaleOrderNo As String
Dim vDocType As Integer
Dim vDocDate As String
Dim vListItem As ListItem

On Error GoTo ErrDescription

FrmCheckQTY.Show
vPrintDocno = Trim(TextScan101.Text)
vTimeID = Trim(LBLTimeID101.Caption)
vSaleOrderNo = Trim(LBLRefDocNo101.Caption)
vARCode = Trim(LBLARCode101.Caption)
vDocType = Trim(LBLDocType101.Caption)
vDocDate = Trim(LBLDocDate.Caption)

FrmCheckQTY.LBLDocno1.Caption = vPrintDocno
FrmCheckQTY.LBLDocno2.Caption = vSaleOrderNo
FrmCheckQTY.LBLARCode.Caption = vARCode
FrmCheckQTY.LBLID.Caption = vTimeID
FrmCheckQTY.LBLDocDate.Caption = vDocDate

If vPrintDocno <> "" And vSaleOrderNo <> "" Then
i = 0
ListView101.ListItems.Clear
'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails2
vQuery = "exec dbo.USP_NP_SearchQueueItemDetails5 '" & vPrintDocno & "','" & vSaleOrderNo & "'," & vDocType & "," & vTimeID & ",'" & vDocDate & "' "
If OpenDataBase2(qConnection, vRecordset, vQuery) <> 0 Then
  vRecordset.MoveFirst
  While Not vRecordset.EOF
    i = i + 1
    Set vListItem = FrmCheckQTY.ListView101.ListItems.Add(, , i)
      vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
      vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
      vListItem.SubItems(4) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
      vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
      vListItem.SubItems(6) = Trim("F1")
      vListItem.SubItems(7) = Trim(vRecordset.Fields("whcode").Value)
  vRecordset.MoveNext
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

Public Sub ClearCheckQueue()
On Error Resume Next

    Me.TextScan101.Text = ""
    Me.LBLARCode101.Caption = ""
    Me.LBLARCode102.Caption = ""
    Me.LBLPicker101.Caption = ""
    Me.LBLRefDocNo101.Caption = ""
    Me.LBLDocDate.Caption = ""
    Me.LBLSale101.Caption = ""
    Me.LBLTime101.Caption = ""
    Me.LBLWHCode101.Caption = ""
    Me.LBLDocType101.Caption = ""
    Me.LBLTimeID101.Caption = ""
    Me.LBLStatus101.Caption = ""
    Me.ListView105.ListItems.Clear
End Sub

Private Sub TextScan101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vDocno As String
Dim vDocDate As String
Dim vListviewItem As ListItem
Dim i As Integer
Dim vMyDescription As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
  vDocno = Trim(TextScan101.Text)
  If Me.CHKDate.Value = 0 Then
  vDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
  Else
  vDocDate = DateAdd("d", Now, -1)
  End If
  'vQuery = "exec dbo.USP_NP_SearchDataQueueDetails1 '" & vDocno & "','" & vSelectZoneID & "','" & vDocDate & "' "
  vQuery = "exec dbo.USP_NP_SearchDataQueueDetails2'" & vDocno & "','" & vSelectZoneID & "','" & vDocDate & "' "
  If OpenDataBase(qConnection, vRecordset, vQuery) <> 0 Then
   vARName101 = Trim(vRecordset.Fields("arcode").Value)
   vARName102 = Trim(vRecordset.Fields("arname").Value)
   vSaleName101 = Trim(vRecordset.Fields("SaleName").Value)
   vPicker101 = Trim(vRecordset.Fields("Picker").Value)
   vWHCode101 = Trim(vRecordset.Fields("WHCode").Value)
   vRefNo101 = Trim(vRecordset.Fields("saleorderno").Value)
   vDiffDateTime101 = Trim(vRecordset.Fields("diffpicking").Value)
   vStatus101 = Trim(vRecordset.Fields("status").Value)
   vIsReceived101 = Trim(vRecordset.Fields("isreceived").Value)
   vDocType101 = Trim(vRecordset.Fields("doctype").Value)
   vTimeID101 = Trim(vRecordset.Fields("timeid").Value)
   vTimeID = Trim(vRecordset.Fields("timeid").Value)
   vDocDate101 = Trim(vRecordset.Fields("docdate").Value)
   vMyDescription = Trim(vRecordset.Fields("mydescription").Value)
   
    Me.LBLARCode101.Caption = vARName101
    Me.LBLARCode102.Caption = vARName102
    Me.LBLPicker101.Caption = vPicker101
    Me.LBLRefDocNo101.Caption = vRefNo101
    Me.LBLSale101.Caption = vSaleName101
    Me.LBLDocDate.Caption = vDocDate101
    Me.LBLTime101.Caption = Trim(vRecordset.Fields("PickingTime").Value)
    Me.LBLWHCode101.Caption = vWHCode101
    Me.LBLDocType101.Caption = vDocType101
    Me.LBLTimeID101.Caption = vTimeID101
    Me.TextMyDescription.Text = vMyDescription
    
    Select Case vStatus101
    Case 0
      Me.LBLStatus101.Caption = Trim("รอจัดสินค้า")
    Case 1
      Me.LBLStatus101.Caption = Trim("กำลังจัดสินค้า")
    Case 2
      Select Case vIsReceived101
      Case 0
      Me.LBLStatus101.Caption = Trim("จัดสินค้าเรียบร้อย ลูกค้ายังไม่มารับของ")
      Case 1
      Me.LBLStatus101.Caption = Trim("จัดสินค้าเรียบร้อย ลูกค้ามารับของแล้ว")
      End Select
    End Select
    
    ListView105.ListItems.Clear
    'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails2
    vQuery = "exec dbo.USP_NP_SearchQueueItemDetails5 '" & vDocno & "','" & vRefNo101 & "'," & vDocType101 & "," & vTimeID101 & ",'" & vDocDate101 & "' "
    If OpenDataBase1(sConnection, vRecordset1, vQuery) <> 0 Then
    vRecordset1.MoveFirst
    i = 0
    While Not vRecordset1.EOF
    i = i + 1
    Set vListviewItem = ListView105.ListItems.Add(, , i)
      vListviewItem.SubItems(1) = Trim(vRecordset1.Fields("itemcode").Value)
      vListviewItem.SubItems(2) = Trim(vRecordset1.Fields("itemname").Value)
      vListviewItem.SubItems(3) = Format(Trim(vRecordset1.Fields("qty").Value), "##,##0.00")
      vListviewItem.SubItems(4) = Trim(vRecordset1.Fields("unitcode").Value)
      vListviewItem.SubItems(5) = Trim(vRecordset1.Fields("shelfcode").Value)
      vListviewItem.SubItems(6) = Format(Trim(vRecordset1.Fields("qty").Value), "##,##0.00")
      vListviewItem.SubItems(7) = Format(Trim(vRecordset1.Fields("pickqty").Value), "##,##0.00")
      vListviewItem.SubItems(8) = Trim(vRecordset1.Fields("familycode").Value)
      vRecordset1.MoveNext
    Wend
    End If
    vRecordset1.Close
    
  Else
    MsgBox "ไม่มีเลขที่คิว " & vDocno & " ที่ต้องการดูในโซนที่เลือกทำงานอยู่ กรุณาตรวจสอบ", vbCritical, "Send Error"
    Me.ListView105.ListItems.Clear
    Me.TextScan101.Text = ""
    Me.LBLARCode101.Caption = ""
    Me.LBLARCode102.Caption = ""
    Me.LBLPicker101.Caption = ""
    Me.LBLRefDocNo101.Caption = ""
    Me.LBLDocDate.Caption = ""
    Me.LBLSale101.Caption = ""
    Me.LBLTime101.Caption = ""
    Me.LBLWHCode101.Caption = ""
    Me.LBLStatus101.Caption = ""
    Me.TextScan101.SetFocus
    LBLTimeID101.Caption = ""
    LBLDocType101.Caption = ""
    vStatus101 = 0
    vIsReceived101 = 0
    vARName101 = ""
    vSaleName101 = ""
    vPicker101 = ""
    vWHCode101 = ""
    vRefNo101 = ""
    vDiffDateTime101 = 0
    vDocType101 = 0
    vTimeID101 = 0
  End If
  vRecordset.Close
  
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub vQueueID()
Call StopTime
Select Case vStatus101
   Case 0
      Call InsertPicker
   Case 1
      Call ConfirmPickItemCode
   Case 2
      Select Case vIsReceived101
      Case 0
      Call CheckIsReceived
   End Select
End Select
Call ClearCheckQueue
PicScanBar.Visible = False
Me.Text101.SetFocus
End Sub
Private Sub Timer101_Timer()
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim vCount As Integer
Dim vPrinted As Integer
Dim vARName  As String
Dim vQueueDate As String
Dim vRefNo As String
Dim vPicker As String
Dim vSaleCode As String
Dim vCustomerZone As String

On Error Resume Next

ListView101.ListItems.Clear
vQuery = "exec dbo.USP_QM_SearchQueueZone5     " & vSelectZoneID & " "
If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
  If vCheckClickListview = 1 Then
    vARName = Trim(vRecordset.Fields("arname").Value)
    vRefNo = Trim(vRecordset.Fields("saleorderno").Value)
    vPicker = Trim(vRecordset.Fields("picker").Value)
    vSaleCode = Trim(vRecordset.Fields("salename").Value)
    vQueueDate = Trim(vRecordset.Fields("docdate").Value)
    Select Case Trim(vRecordset.Fields("customerzone").Value)
    Case 0
      vCustomerZone = "ลูกค้ารอรับของตามจุดออกใบหยิบ"
    Case 1
      vCustomerZone = "ลูกค้ารอรับของฝั่ง : สำนักงานใหญ่ "
    Case 2
      vCustomerZone = "ลูกค้ารอรับของฝั่ง : OutLet"
    End Select

    FrmQueue.LBLQueueDate.Caption = vQueueDate
    FrmQueue.LBLARName.Caption = vARName
    FrmQueue.LBLRefNo.Caption = vRefNo
    FrmQueue.LBLUserPick.Caption = Trim("-")
    FrmQueue.LBLSale.Caption = vSaleCode
    FrmQueue.LBLCustomerZone.Caption = vCustomerZone
  End If
    While Not vRecordset.EOF
        Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        vListItem.SubItems(1) = Trim(vRecordset.Fields("salename").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("requesttime").Value)
        vListItem.SubItems(3) = Trim(vRecordset.Fields("saleorderno").Value)
        vListItem.SubItems(4) = Trim(vRecordset.Fields("doctype").Value)
        vListItem.SubItems(5) = Trim(vRecordset.Fields("arname").Value)
        vListItem.SubItems(6) = Trim(vRecordset.Fields("timeid").Value)
        vListItem.SubItems(7) = Trim(vRecordset.Fields("customerzone").Value)
        vListItem.SubItems(8) = Trim(vRecordset.Fields("docdate").Value)
        vListItem.SubItems(9) = Trim(vRecordset.Fields("timepick").Value)
        vListItem.SubItems(10) = Trim(vRecordset.Fields("zoneid").Value)
        vListItem.SubItems(11) = Trim(vRecordset.Fields("familycode").Value)
        vListItem.SubItems(12) = Trim(vRecordset.Fields("whcode").Value)
        vListItem.SubItems(13) = Trim(vRecordset.Fields("shelfgroup").Value)
        vListItem.SubItems(14) = Trim(vRecordset.Fields("pickzone").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub


Private Sub Timer102_Timer()
'Dim vRecordset As New ADODB.Recordset
'Dim vListItem As ListItem
'Dim vCount As Integer
'Dim vPrinted As Integer

'On Error Resume Next

'ListView102.ListItems.Clear
'vQuery = "exec dbo.USP_QM_SearchQueuePickingZone " & vSelectZoneID & " "
'If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
 '   While Not vRecordset.EOF
  '      Set vListItem = ListView102.ListItems.Add(, , "")
   '     vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
    ''    vListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
        'vListItem.SubItems(3) = Trim(vRecordset.Fields("startdatetime").Value)
      ''  vListItem.SubItems(4) = Format(-1 * DateDiff("s", Now, vListItem.SubItems(3)) / 60, "##,##0.00")
        'vListItem.SubItems(5) = Trim(vRecordset.Fields("statusdesc").Value)
        'vListItem.SubItems(6) = Trim(vRecordset.Fields("saleorderno").Value)
        'vListItem.SubItems(7) = Trim(vRecordset.Fields("doctype").Value)
        
        'If CCur(vListItem.SubItems(4)) >= 15 Then
         ' ListView102.ListItems(i).ForeColor = "&H000000FF"
          'ListView102.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF" 'red
          'ListView102.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
          'ListView102.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
          'ListView102.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
          'ListView102.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
          'ListView102.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
          'ListView102.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
        'End If
        'If CCur(vListItem.SubItems(4)) >= 12 And CCur(vListItem.SubItems(4)) < 15 Then 'violet
         ' ListView102.ListItems(i).ForeColor = "&H00FF00FF"
          'ListView102.ListItems.Item(i).ListSubItems(1).ForeColor = "&H00FF00FF"
          'ListView102.ListItems.Item(i).ListSubItems(2).ForeColor = "&H00FF00FF"
          'ListView102.ListItems.Item(i).ListSubItems(3).ForeColor = "&H00FF00FF"
          'ListView102.ListItems.Item(i).ListSubItems(4).ForeColor = "&H00FF00FF"
          'ListView102.ListItems.Item(i).ListSubItems(5).ForeColor = "&H00FF00FF"
          'ListView102.ListItems.Item(i).ListSubItems(6).ForeColor = "&H00FF00FF"
          'ListView102.ListItems.Item(i).ListSubItems(7).ForeColor = "&H00FF00FF"
        'End If
        'If CCur(vListItem.SubItems(4)) >= 9 And CCur(vListItem.SubItems(4)) < 12 Then 'dark blue
         ' ListView102.ListItems(i).ForeColor = "&H00FF0000"
          'ListView102.ListItems.Item(i).ListSubItems(1).ForeColor = "&H00FF0000"
          'ListView102.ListItems.Item(i).ListSubItems(2).ForeColor = "&H00FF0000"
          'ListView102.ListItems.Item(i).ListSubItems(3).ForeColor = "&H00FF0000"
          'ListView102.ListItems.Item(i).ListSubItems(4).ForeColor = "&H00FF0000"
          'ListView102.ListItems.Item(i).ListSubItems(5).ForeColor = "&H00FF0000"
          'ListView102.ListItems.Item(i).ListSubItems(6).ForeColor = "&H00FF0000"
          'ListView102.ListItems.Item(i).ListSubItems(7).ForeColor = "&H00FF0000"
        'End If
        'If CCur(vListItem.SubItems(4)) >= 6 And CCur(vListItem.SubItems(4)) < 9 Then 'light blue
         ' ListView102.ListItems(i).ForeColor = "&H00FFFF00"
          'ListView102.ListItems.Item(i).ListSubItems(1).ForeColor = "&H00FFFF00"
          'ListView102.ListItems.Item(i).ListSubItems(2).ForeColor = "&H00FFFF00"
          'ListView102.ListItems.Item(i).ListSubItems(3).ForeColor = "&H00FFFF00"
          'ListView102.ListItems.Item(i).ListSubItems(4).ForeColor = "&H00FFFF00"
          'ListView102.ListItems.Item(i).ListSubItems(5).ForeColor = "&H00FFFF00"
          'ListView102.ListItems.Item(i).ListSubItems(6).ForeColor = "&H00FFFF00"
          'ListView102.ListItems.Item(i).ListSubItems(7).ForeColor = "&H00FFFF00"
        'End If
        'If CCur(vListItem.SubItems(4)) >= 3 And CCur(vListItem.SubItems(4)) < 6 Then 'green
         ' ListView102.ListItems(i).ForeColor = "&H0000FF00"
          'ListView102.ListItems.Item(i).ListSubItems(1).ForeColor = "&H0000FF00"
          'ListView102.ListItems.Item(i).ListSubItems(2).ForeColor = "&H0000FF00"
          'ListView102.ListItems.Item(i).ListSubItems(3).ForeColor = "&H0000FF00"
          'ListView102.ListItems.Item(i).ListSubItems(4).ForeColor = "&H0000FF00"
          'ListView102.ListItems.Item(i).ListSubItems(5).ForeColor = "&H0000FF00"
          'ListView102.ListItems.Item(i).ListSubItems(6).ForeColor = "&H0000FF00"
          'ListView102.ListItems.Item(i).ListSubItems(7).ForeColor = "&H0000FF00"
        'End If
        'If CCur(vListItem.SubItems(4)) >= 1 And CCur(vListItem.SubItems(4)) < 3 Then 'yellow
         ' ListView102.ListItems(i).ForeColor = "&H0000FFFF"
          'ListView102.ListItems.Item(i).ListSubItems(1).ForeColor = "&H0000FFFF"
          'ListView102.ListItems.Item(i).ListSubItems(2).ForeColor = "&H0000FFFF"
          'ListView102.ListItems.Item(i).ListSubItems(3).ForeColor = "&H0000FFFF"
          'ListView102.ListItems.Item(i).ListSubItems(4).ForeColor = "&H0000FFFF"
          'ListView102.ListItems.Item(i).ListSubItems(5).ForeColor = "&H0000FFFF"
          'ListView102.ListItems.Item(i).ListSubItems(6).ForeColor = "&H0000FFFF"
          'ListView102.ListItems.Item(i).ListSubItems(7).ForeColor = "&H0000FFFF"
        'End If
    'vRecordset.MoveNext
    'Wend
'End If
'vRecordset.Close
End Sub

Private Sub Timer103_Timer()
Call RefreshQueuePicking
End Sub

Private Sub Timer104_Timer()
Dim i As Integer
Dim vStartTime As String
Dim vCount As Integer

On Error Resume Next

If ListView102.ListItems.Count <> 0 Then
  i = ListView102.ListItems.Count
 For vCount = 1 To i
  vStartTime = ListView102.ListItems.Item(vCount).SubItems(9)
  'ListView102.ListItems.Item(vCount).SubItems(2) = Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")
  
        If CCur(Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")) >= 15 And ListView102.ListItems(vCount).ForeColor <> "&H000000FF" Then
          ListView102.ListItems(vCount).ForeColor = "&H000000FF"
          ListView102.ListItems.Item(vCount).ListSubItems(1).ForeColor = "&H000000FF" 'red
          ListView102.ListItems.Item(vCount).ListSubItems(2).ForeColor = "&H000000FF"
          ListView102.ListItems.Item(vCount).ListSubItems(3).ForeColor = "&H000000FF"
          ListView102.ListItems.Item(vCount).ListSubItems(4).ForeColor = "&H000000FF"
          ListView102.ListItems.Item(vCount).ListSubItems(5).ForeColor = "&H000000FF"
          ListView102.ListItems.Item(vCount).ListSubItems(6).ForeColor = "&H000000FF"
          ListView102.ListItems.Item(vCount).ListSubItems(7).ForeColor = "&H000000FF"
          ListView102.ListItems.Item(vCount).ListSubItems(8).ForeColor = "&H000000FF"
          ListView102.ListItems.Item(vCount).ListSubItems(9).ForeColor = "&H000000FF"
          ListView102.ListItems.Item(vCount).ListSubItems(10).ForeColor = "&H000000FF"
          ListView102.ListItems.Item(vCount).ListSubItems(11).ForeColor = "&H000000FF"
          ListView102.ListItems.Item(vCount).ListSubItems(12).ForeColor = "&H000000FF"
        End If

        If CCur(Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")) >= 10 And CCur(Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")) < 15 And ListView102.ListItems(vCount).ForeColor <> "&H00FF0000" Then 'dark blue
          ListView102.ListItems(vCount).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(1).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(2).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(3).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(4).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(5).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(6).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(7).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(8).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(9).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(10).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(11).ForeColor = "&H00FF0000"
          ListView102.ListItems.Item(vCount).ListSubItems(12).ForeColor = "&H00FF0000"
        End If

        If CCur(Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")) >= 5 And CCur(Format(-1 * DateDiff("s", Now, vStartTime) / 60, "##,##0.00")) < 10 And ListView102.ListItems(vCount).ForeColor <> "&H00008000" Then   'green
          ListView102.ListItems(vCount).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(1).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(2).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(3).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(4).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(5).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(6).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(7).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(8).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(9).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(10).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(11).ForeColor = "&H00008000"
          ListView102.ListItems.Item(vCount).ListSubItems(12).ForeColor = "&H00008000"
        End If
        ListView102.Refresh
  Next vCount
End If
End Sub


Public Sub StartTime()
Timer101.Enabled = True
Timer102.Enabled = True
Timer103.Enabled = True
Timer104.Enabled = True
Timer105.Enabled = True
End Sub

Public Sub StopTime()
Timer101.Enabled = False
Timer102.Enabled = False
Timer103.Enabled = False
Timer104.Enabled = False
Timer105.Enabled = False
End Sub

Public Sub RefreshData()
  Call RefreshQueueBegin
  Call RefreshQueuePicking
  Call RefreshQueueFinish
End Sub

Public Sub PrintPickingSlipHeader(vSaleOrder As String, vWHCode As String, vShelfCode As String, vZone As Integer, vTimeID As Integer)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocno As String
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer

If vZone = 0 Then
vPrinterName = Trim("\\galaxy\EPS TM-T88III-NP")
For Each printerObj In Printers
  If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    Set Printer = printerObj
    Set printerObj = Nothing
    Exit For
  End If
Next
End If
   
If vZone = 1 Then
vPrinterName = Trim("\\galaxy\EPS TM-T88III-OutLet")
For Each printerObj In Printers
  If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    Set Printer = printerObj
    Set printerObj = Nothing
    Exit For
  End If
Next
End If

    vQuery = "exec dbo.USP_SO_PickingQueueFreedom '" & vSaleOrder & "','" & vWHCode & "','" & vShelfCode & "'," & vTimeID & " "
    'vQuery = "exec dbo.USP_SO_PickingQueue1 '" & vSaleOrder & "','" & vWHCode & "','" & vShelfCode & "' "
    If OpenDataBase(sConnection, vRecordset, vQuery) <> 0 Then
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.FontBold = True
      Printer.CurrentX = 1800
      Printer.CurrentY = 0
      Printer.Print Trim(vRecordset.Fields("queueno").Value)
      
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 1400
      Printer.CurrentY = 1000
      Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("_______________________________________________________________________________________")
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า")
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 1900
      Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 1900
      Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 3150
      If vRecordset.Fields("isconditionsend").Value = 0 Then
            Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("รับเอง")
      Else
            Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("ส่งให้")
      End If
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 3400
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value)
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    End If
    vRecordset.Close
    vQuery = "select * from dbo.bcreportname where repid = 323"
    If OpenDataBase(sConnection, vRecordset, vQuery) <> 0 Then
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY
      Printer.Print Trim(vRecordset.Fields("reportname").Value)
      Printer.CurrentX = Printer.CurrentX + 2000
      Printer.Print Trim("วันที่พิมพ์ :") & Now
    End If
    vRecordset.Close

           
    Printer.EndDoc

  End Sub
  
Public Sub PrintPickingSlip(vSaleOrder As String, vWHCode As String, vShelfCode As String, vZone As Integer, vCountPick As Integer)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vQuery As String
Dim vDocno As String
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer
Dim vSoStatus As Integer
Dim vItemName As String
Dim vGroupDocNo As String
Dim vPickStatus As Integer


'If vZone = 0 Then
 '  vPrinterName = Trim("\\galaxy\EPS TM-T88III-NP")
  ' For Each printerObj In Printers
   '  If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    '   Set Printer = printerObj
     '  Set printerObj = Nothing
      ' Exit For
     'End If
   'Next
'End If
   
'If vZone = 1 Then
 '  vPrinterName = Trim("\\galaxy\EPS-TM-PickingOutlet")
  ' For Each printerObj In Printers
   '  If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    '   Set Printer = printerObj
     '  Set printerObj = Nothing
      ' Exit For
     'End If
   'Next
'End If

'If vZone = 2 Then
 '  vPrinterName = Trim("\\it-queuemedia\EPS-TM-PickingHMX")
  ' For Each printerObj In Printers
   '  If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    '   Set Printer = printerObj
     '  Set printerObj = Nothing
      ' Exit For
     'End If
   'Next
'End If

vQuery = "exec dbo.USP_NP_SearchCheckPrinter " & vZone & " "
If OpenDataBase(sConnection, vRecordset1, vQuery) <> 0 Then
vPrinterName = Trim(vRecordset1.Fields("printername").Value)
End If
vRecordset1.Close

For Each printerObj In Printers
If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
Set Printer = printerObj
Set printerObj = Nothing
Exit For
End If
Next

vQuery = "exec dbo.USP_SO_PickingQueueFreedom '" & vSaleOrder & "','" & vWHCode & "','" & vShelfCode & "'," & vCountPick & " "
    If OpenDataBase(sConnection, vRecordset, vQuery) <> 0 Then
    
      vSoStatus = vRecordset.Fields("sostatus").Value
      vPickStatus = vRecordset.Fields("pickstatus").Value
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

If vPickStatus = 0 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 20
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 80
      Printer.Print "(จ่าย)"
ElseIf vPickStatus = 1 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 20
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 80
      Printer.Print "(จอง)"
End If

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.FontBold = True
      Printer.CurrentX = 1600
      Printer.CurrentY = 0
      Printer.Print Trim(vRecordset.Fields("queueno").Value)
      
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 1200
      Printer.CurrentY = 1000
      Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("_______________________________________________________________________________________")
      
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 12
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ทดแทนใบจัดสินค้า")
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 1900
      Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 1900
      Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      

      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 3150
      
      If vRecordset.Fields("isconditionsend").Value = 0 Then
            Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("รับเอง")
      Else
            Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("ส่งให้")
      End If
                  
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 16
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 3400
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value)

      If Trim(vRecordset.Fields("carlicense").Value) <> "" Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 16
        Printer.CurrentX = 1400
        Printer.CurrentY = 3400
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.Print Trim("ทะเบียนรถขนส่ง : ") & Trim(vRecordset.Fields("carlicense").Value)
      End If
            
      If vSoStatus = 0 Then
         Printer.Font.Name = "AngsanaUPC"
         Printer.Font.Size = 14
         Printer.FontBold = True
         Printer.FontUnderline = False
         Printer.CurrentX = 0
         Printer.CurrentY = 3800
         Printer.Print Trim("เวลารับของ : ") & Trim(vRecordset.Fields("requesttime").Value)
      Else
         Printer.Font.Name = "AngsanaUPC"
         Printer.Font.Size = 14
         Printer.FontBold = True
         Printer.FontUnderline = False
         Printer.CurrentX = 0
         Printer.CurrentY = 3800
         Printer.Print Trim("วันที่ครบกำหนดรับของ : ") & Trim(vRecordset.Fields("duedate").Value)
      End If
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 14
      Printer.CurrentX = 0
      Printer.CurrentY = 4150
      Printer.Print Trim(vRecordset.Fields("customerzone").Value)
      
      vRecordset.MoveFirst
      vLineX = 50
      vLineY = 50
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY - 30
      Printer.FontBold = False
      Printer.FontUnderline = False
      Printer.Print Trim("-----------------------------------------------------------------------------------------------")
      n = 1
      While Not vRecordset.EOF
          Printer.Font.Size = 11
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfcode1").Value) & "    " & Trim("OnHand") & "(" & Trim(vRecordset.Fields("shelfcode").Value) & ")" & ": " & Trim(vRecordset.Fields("qtylocation").Value) & "  " & Trim(vRecordset.Fields("stkunitcode").Value) & "     " & "รวมคลัง :" & Trim(vRecordset.Fields("StkWHCode").Value) & "  " & Trim(vRecordset.Fields("stkunitcode").Value)
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value) & "             " & " ขายชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)
          
          vItemName = Trim(vRecordset.Fields("itemname").Value) & Trim(vRecordset.Fields("descriptionline"))
          If Len(vItemName) <= 55 Then
             Printer.CurrentX = 0
             Printer.CurrentY = Printer.CurrentY
             Printer.Print "ชื่อสินค้า :" & vItemName
          Else
             Printer.CurrentX = 0
             Printer.CurrentY = Printer.CurrentY
             Printer.Print "ชื่อสินค้า :" & Left(vItemName, 55)
             
             Printer.CurrentX = 600
             Printer.CurrentY = Printer.CurrentY
             Printer.Print Right(vItemName, Len(vItemName) - 55)
          End If
          
          Printer.Font.Size = 13
          Printer.CurrentX = Printer.CurrentX + 15
          Printer.CurrentY = Printer.CurrentY + 100
          Printer.FontBold = True
          Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY - 50
          Printer.FontBold = False
          Printer.Print Trim("-----------------------------------------------------------------------------------------------")
          
      vRecordset.MoveNext
      n = n + 1
      Wend
    End If
    vRecordset.Close
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 400
    Printer.Print Trim("_______________________________________________________________________________________________")
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + vLineY
    Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 150
    Printer.Print "         _____________                                    ______________"
          
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    
    Printer.Font.Size = 10
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("วันที่พิมพ์ :") & Now
    Printer.EndDoc
    
End Sub

Public Sub PrintHeaderPos(vDocno As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocDate As Date
Dim vPrinterName As String
Dim printerObj As Printer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer


vPrinterName = Trim("\\galaxy\EPS TM-T88III-NP")
For Each printerObj In Printers
  If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    Set Printer = printerObj
    Set printerObj = Nothing
    Exit For
  End If
Next


vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vQuery = "exec dbo.USP_SO_PickingQueuePos1 '" & vDocno & "','" & vDocDate & "' "
    If OpenDataBase(sConnection, vRecordset, vQuery) <> 0 Then
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.FontBold = True
      Printer.CurrentX = 1800
      Printer.CurrentY = 0
      Printer.Print Trim(vRecordset.Fields("queueno").Value)
      
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 1400
      Printer.CurrentY = 1000
      Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("_______________________________________________________________________________________")
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า")
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 1900
      Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 1900
      Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim("OutLet")
      

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 3150
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value)
End If
vRecordset.Close

    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 100
    Printer.Print Trim("______________________________________________________________________________________________")
    
    vQuery = "select * from dbo.bcreportname where repid = 324"
    If OpenDataBase(sConnection, vRecordset, vQuery) <> 0 Then
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY
      Printer.Print Trim(vRecordset.Fields("reportname").Value)
      Printer.CurrentX = Printer.CurrentX + 2000
      Printer.Print Trim("วันที่พิมพ์ :") & Now
    End If
    vRecordset.Close
      
    Printer.EndDoc
End Sub

Public Sub PrintDetails(vDocno As String, vZoneGroup As String, vFamilyGroup As String, vPickZoneGroup As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocDate As Date
Dim vPrinterName As String
Dim printerObj As Printer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim vPrinterID As Integer

On Error Resume Next

If vPickZoneGroup = "01" Then
vPrinterID = 0
End If

If vPickZoneGroup = "02" Then
vPrinterID = 1
End If

If vPickZoneGroup = "03" Then
vPrinterID = 2
End If

vQuery = "exec dbo.USP_NP_SearchCheckPrinter " & vPrinterID & " "
If OpenDataBase(sConnection, vRecordset, vQuery) <> 0 Then
vPrinterName = Trim(vRecordset.Fields("printername").Value)
End If
vRecordset.Close



For Each printerObj In Printers
If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
Set Printer = printerObj
Set printerObj = Nothing
Exit For
End If
Next

    vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vQuery = "exec dbo.USP_SO_PickingQueuePos3 '" & vDocno & "','" & vDocDate & "' ,'" & vZoneGroup & "','" & vFamilyGroup & "','" & vPickZoneGroup & "' "
    If OpenDataBase(sConnection, vRecordset, vQuery) <> 0 Then
    If Not vRecordset.EOF Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.FontBold = True
      Printer.CurrentX = 1700
      Printer.CurrentY = 0
      Printer.Print Trim(vRecordset.Fields("queueno").Value)
      
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 1200
      Printer.CurrentY = 1000
      Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("_______________________________________________________________________________________")
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ทดแทนใบจัดสินค้า")
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 1900
      Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 1900
      Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 12
      Printer.CurrentX = 0
      Printer.CurrentY = 3150
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value)
      
      vRecordset.MoveFirst
      vLineX = 50
      vLineY = 50
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY - 80
      Printer.Print Trim("-----------------------------------------------------------------------------------------------")
      n = 1
      While Not vRecordset.EOF
          Printer.Font.Size = 11
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ขายชั้นเก็บ :" & Trim(vRecordset.Fields("MBShelfCode").Value) & "       " & Trim("OnHand: ") & Trim(vRecordset.Fields("qtyonhand").Value) & "       " & Trim("รวมคลัง : ") & "  " & Trim(vRecordset.Fields("stkwhcode").Value) & "    " & Trim(vRecordset.Fields("unitcode").Value)
                                      
          Printer.Font.Size = 18
          Printer.FontBold = True
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)
          
          Printer.Font.Name = "3 of 9 Barcode"
          Printer.Font.Size = 20
          Printer.FontBold = False
          Printer.CurrentX = 200
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "*" & Trim(vRecordset.Fields("itemcode").Value) & "*"
      
          Printer.Font.Name = "AngsanaUPC"
          Printer.Font.Size = 11
          Printer.FontBold = False
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value)
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ชื่อสินค้า :" & Trim(vRecordset.Fields("itemname").Value)
          
          Printer.CurrentX = Printer.CurrentX + 15
          Printer.CurrentY = Printer.CurrentY + 50
          Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY - 80
          Printer.Print Trim("-----------------------------------------------------------------------------------------------")
          
      vRecordset.MoveNext
      n = n + 1
      Wend
    End If
    End If
    vRecordset.Close
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 400
    Printer.Print Trim("_______________________________________________________________________________________________")
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + vLineY
    Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 150
    Printer.Print "         _____________                                    ______________"
          
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    
    Printer.CurrentX = Printer.CurrentX + 2000
    Printer.Print Trim("วันที่พิมพ์ :") & Now & "          " & vPrinterName

    Printer.EndDoc
End Sub

Public Sub PrintDetailsPos(vDocno As String, vShelfGroup As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocDate As Date
Dim vPrinterName As String
Dim printerObj As Printer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer


If UCase(vShelfGroup) = "PKA" Then
   vPrinterName = Trim("\\it-queuemedia\EPS-TM-PickingHMX")
   For Each printerObj In Printers
     If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
       Set Printer = printerObj
       Set printerObj = Nothing
       Exit For
     End If
   Next
ElseIf UCase(vShelfGroup) = "PKB" Then
   vPrinterName = Trim("\\galaxy\EPS-TM-PickingOutlet")
   For Each printerObj In Printers
     If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
       Set Printer = printerObj
       Set printerObj = Nothing
       Exit For
     End If
   Next
End If

vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vQuery = "exec dbo.USP_SO_PickingQueuePos1 '" & vDocno & "','" & vDocDate & "','" & vShelfGroup & "' "
    If OpenDataBase(sConnection, vRecordset, vQuery) <> 0 Then
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.FontBold = True
      Printer.CurrentX = 1800
      Printer.CurrentY = 0
      Printer.Print Trim(vRecordset.Fields("queueno").Value)
      
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 1400
      Printer.CurrentY = 1000
      Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("_______________________________________________________________________________________")
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ทดแทนใบจัดสินค้า")
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 1900
      Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 1900
      Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim("OutLet")
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 3150
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value)
      
      vRecordset.MoveFirst
      vLineX = 50
      vLineY = 50
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY - 80
      Printer.Print Trim("-----------------------------------------------------------------------------------------------")
      n = 1
      While Not vRecordset.EOF
          Printer.Font.Size = 11
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode").Value) & "       " & Trim("OnHand : ") & Trim(vRecordset.Fields("qtyonhand").Value) & "  " & Trim(vRecordset.Fields("unitcode").Value)
          
          Printer.Font.Name = "3 of 9 Barcode"
          Printer.Font.Size = 20
          Printer.FontBold = False
          Printer.CurrentX = 200
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "*" & Trim(vRecordset.Fields("itemcode").Value) & "*"
      
          Printer.Font.Name = "AngsanaUPC"
          Printer.Font.Size = 11
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value)
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ชื่อสินค้า :" & Trim(vRecordset.Fields("itemname").Value)
          
          Printer.CurrentX = Printer.CurrentX + 15
          Printer.CurrentY = Printer.CurrentY + 50
          Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY - 80
          Printer.Print Trim("-----------------------------------------------------------------------------------------------")
          
      vRecordset.MoveNext
      n = n + 1
      Wend
    End If
    vRecordset.Close
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 400
    Printer.Print Trim("_______________________________________________________________________________________________")
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + vLineY
    Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 150
    Printer.Print "         _____________                                    ______________"
          
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    
    Printer.EndDoc

End Sub

Public Sub PrintRequestPickingSlip(vDocno As String, vZoneGroup As String, vFamilyGroup As String, vPickZoneGroup As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocDate As Date
Dim vPrinterName As String
Dim printerObj As Printer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim vPrinterID As Integer

On Error Resume Next

If vPickZoneGroup = "01" Then
vPrinterID = 0
End If

If vPickZoneGroup = "02" Then
vPrinterID = 1
End If

If vPickZoneGroup = "03" Then
vPrinterID = 2
End If

vQuery = "exec dbo.USP_NP_SearchCheckPrinter " & vPrinterID & " "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
vPrinterName = Trim(vRecordset.Fields("printername").Value)
End If
vRecordset.Close


For Each printerObj In Printers
If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
Set Printer = printerObj
Set printerObj = Nothing
Exit For
End If
Next

    vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vQuery = "exec dbo.USP_NP_SearchPickingRequestDetails '" & vDocno & "','" & vDocDate & "' ,'" & vZoneGroup & "','" & vFamilyGroup & "','" & vPickZoneGroup & "' "
    If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    If Not vRecordset.EOF Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.FontBold = True
      Printer.CurrentX = 1700
      Printer.CurrentY = 0
      Printer.Print Trim(vRecordset.Fields("queueno").Value)
      
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 1200
      Printer.CurrentY = 1000
      Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("_______________________________________________________________________________________")
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ทดแทนใบจัดสินค้า")
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 1900
      Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 1900
      Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 12
      Printer.CurrentX = 0
      Printer.CurrentY = 3150
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value)
      
      vRecordset.MoveFirst
      vLineX = 50
      vLineY = 50
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY - 80
      Printer.Print Trim("-----------------------------------------------------------------------------------------------")
      n = 1
      While Not vRecordset.EOF
          Printer.Font.Size = 11
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ขายชั้นเก็บ :" & Trim(vRecordset.Fields("MBShelfCode").Value) & "       " & Trim("OnHand: ") & Trim(vRecordset.Fields("qtyonhand").Value) & "       " & Trim("รวมคลัง : ") & "  " & Trim(vRecordset.Fields("stkwhcode").Value) & "    " & Trim(vRecordset.Fields("unitcode").Value)
                                      
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)
          
          Printer.Font.Name = "3 of 9 Barcode"
          Printer.Font.Size = 20
          Printer.FontBold = False
          Printer.CurrentX = 200
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "*" & Trim(vRecordset.Fields("itemcode").Value) & "*"
      
          Printer.Font.Name = "AngsanaUPC"
          Printer.Font.Size = 11
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value)
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ชื่อสินค้า :" & Trim(vRecordset.Fields("itemname").Value)
          
          Printer.CurrentX = Printer.CurrentX + 15
          Printer.CurrentY = Printer.CurrentY + 50
          Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY - 80
          Printer.Print Trim("-----------------------------------------------------------------------------------------------")
          
      vRecordset.MoveNext
      n = n + 1
      Wend
    End If
    End If
    vRecordset.Close
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 400
    Printer.Print Trim("_______________________________________________________________________________________________")
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + vLineY
    Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 150
    Printer.Print "         _____________                                    ______________"
          
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    
    Printer.CurrentX = Printer.CurrentX + 2000
    Printer.Print Trim("วันที่พิมพ์ :") & Now & "          " & vPrinterName

    Printer.EndDoc
End Sub

