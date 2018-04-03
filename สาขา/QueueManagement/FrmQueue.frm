VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQueue 
   BackColor       =   &H00808080&
   Caption         =   "โปรแกรม จัดควบคุมคิวสินค้า"
   ClientHeight    =   9315
   ClientLeft      =   2295
   ClientTop       =   1575
   ClientWidth     =   14880
   Icon            =   "FrmQueue.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10773.68
   ScaleMode       =   0  'User
   ScaleWidth      =   14880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicScanBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   9240
      Left            =   45
      ScaleHeight     =   9210
      ScaleWidth      =   14790
      TabIndex        =   32
      Top             =   45
      Visible         =   0   'False
      Width           =   14820
      Begin VB.PictureBox PICPoint4 
         Height          =   285
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   360
         TabIndex        =   145
         Top             =   0
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox PicEditQty 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4380
         Left            =   1215
         ScaleHeight     =   4350
         ScaleWidth      =   13035
         TabIndex        =   89
         Top             =   3600
         Visible         =   0   'False
         Width           =   13065
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
            Width           =   10995
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
      Begin VB.PictureBox PTShowMyDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   7740
         ScaleHeight     =   4305
         ScaleWidth      =   6510
         TabIndex        =   85
         Top             =   2565
         Visible         =   0   'False
         Width           =   6540
         Begin VB.TextBox TextMyDescription 
            Appearance      =   0  'Flat
            Height          =   3210
            Left            =   180
            MultiLine       =   -1  'True
            TabIndex        =   86
            Top             =   90
            Width           =   6135
         End
         Begin VB.CommandButton CMDUpdateMyDescription 
            Caption         =   "ปรับหมายเหตุ"
            Height          =   510
            Left            =   5040
            TabIndex        =   87
            Top             =   3465
            Width           =   1275
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
         Left            =   11700
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
         Height          =   4380
         Left            =   1215
         TabIndex        =   57
         Top             =   3600
         Width           =   13065
         _ExtentX        =   23045
         _ExtentY        =   7726
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
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   7937
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
         Left            =   10575
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
         Left            =   11700
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1215
         TabIndex        =   56
         Top             =   3285
         Width           =   1140
      End
      Begin VB.Line Line1 
         BorderWidth     =   10
         X1              =   -4770
         X2              =   16380
         Y1              =   3150
         Y2              =   3150
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
         Width           =   6630
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
         Width           =   4560
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
         Left            =   11700
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
         Left            =   11700
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
         Left            =   10260
         TabIndex        =   38
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "เลขที่อ้างอิง :"
         Height          =   285
         Left            =   10710
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
         Width           =   8475
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
   Begin VB.PictureBox PICPickItemToCust 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   9285
      Left            =   45
      ScaleHeight     =   9255
      ScaleWidth      =   14790
      TabIndex        =   107
      Top             =   45
      Visible         =   0   'False
      Width           =   14820
      Begin VB.PictureBox PICPoint5 
         Height          =   240
         Left            =   0
         ScaleHeight     =   180
         ScaleWidth      =   360
         TabIndex        =   146
         Top             =   0
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox PICKeyItemQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   6090
         Left            =   225
         ScaleHeight     =   6060
         ScaleWidth      =   14295
         TabIndex        =   118
         Top             =   1980
         Visible         =   0   'False
         Width           =   14325
         Begin VB.CommandButton CMDKeyQtyClose 
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
            Height          =   600
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   125
            Top             =   3195
            Width           =   1410
         End
         Begin VB.TextBox TBKeyQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2520
            TabIndex        =   124
            Top             =   2250
            Width           =   2490
         End
         Begin VB.TextBox TBItemCode 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   2520
            TabIndex        =   120
            Top             =   675
            Width           =   4290
         End
         Begin VB.Label LBLKeyQtyItemCodeIndex 
            Height          =   375
            Left            =   2520
            TabIndex        =   140
            Top             =   225
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.Label LBLKeyQtyItemCode 
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
            Height          =   540
            Left            =   6885
            TabIndex        =   126
            Top             =   675
            Width           =   2850
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "จำนวนสินค้า :"
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
            Left            =   720
            TabIndex        =   123
            Top             =   2385
            Width           =   1725
         End
         Begin VB.Label LBLKeyQtyItemName 
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
            Height          =   465
            Left            =   2520
            TabIndex        =   122
            Top             =   1485
            Width           =   10545
         End
         Begin VB.Label Label43 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   855
            TabIndex        =   121
            Top             =   1575
            Width           =   1590
         End
         Begin VB.Label Label41 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   810
            TabIndex        =   119
            Top             =   765
            Width           =   1635
         End
      End
      Begin VB.PictureBox PICItemPayQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   6090
         Left            =   225
         ScaleHeight     =   6060
         ScaleWidth      =   14295
         TabIndex        =   129
         Top             =   1980
         Visible         =   0   'False
         Width           =   14325
         Begin VB.TextBox TBPayItemQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2520
            TabIndex        =   130
            Top             =   2250
            Width           =   2490
         End
         Begin VB.CommandButton CMDPayQtyOK 
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
            Height          =   600
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   132
            Top             =   3195
            Width           =   1410
         End
         Begin VB.Label LBLPayQtyIndex 
            Alignment       =   2  'Center
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
            Height          =   540
            Left            =   6840
            TabIndex        =   137
            Top             =   675
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label LBLPayItemCode 
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
            Height          =   540
            Left            =   2520
            TabIndex        =   136
            Top             =   675
            Width           =   4245
         End
         Begin VB.Label Label49 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   810
            TabIndex        =   135
            Top             =   765
            Width           =   1635
         End
         Begin VB.Label Label48 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   855
            TabIndex        =   134
            Top             =   1575
            Width           =   1590
         End
         Begin VB.Label LBLPayItemName 
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
            Height          =   465
            Left            =   2520
            TabIndex        =   133
            Top             =   1485
            Width           =   10545
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "จำนวนสินค้า :"
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
            Left            =   720
            TabIndex        =   131
            Top             =   2385
            Width           =   1725
         End
      End
      Begin VB.CommandButton CMDKeyItemQty 
         BackColor       =   &H00808080&
         Caption         =   "ยิงบาร์โค้ดตรวจสอบ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   7245
         Width           =   2265
      End
      Begin VB.CommandButton CMDPickItemToCustClose 
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
         Height          =   825
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   7245
         Width           =   2265
      End
      Begin VB.CommandButton CMDPickItemToCust 
         BackColor       =   &H00808080&
         Caption         =   "บันทึกการจ่าย"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   7245
         Width           =   2265
      End
      Begin VB.CommandButton CMDPayItemSearch 
         Height          =   375
         Left            =   6480
         Picture         =   "FrmQueue.frx":26D4
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   1620
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ListView ListViewPayItemToCust 
         Height          =   4965
         Left            =   225
         TabIndex        =   112
         Top             =   1980
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   8758
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
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "จำนวนขึ้นรถ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "รหัสสินค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วย"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "จัดได้"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "คิวที่"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "อ้างถึง"
            Object.Width           =   4057
         EndProperty
      End
      Begin VB.TextBox TBPayItemSearch 
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
         Height          =   375
         Left            =   5625
         TabIndex        =   110
         Top             =   1620
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label LBLPayItemIndex 
         BackColor       =   &H00808080&
         Height          =   375
         Left            =   4320
         TabIndex        =   141
         Top             =   225
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "คิวที่ :"
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
         Left            =   315
         TabIndex        =   139
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label LBLPayItemQueID 
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
         Left            =   1800
         TabIndex        =   138
         Top             =   225
         Width           =   915
      End
      Begin VB.Label LBLDocType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6885
         TabIndex        =   127
         Top             =   1620
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการ สินค้าที่ขนของขึ้นรถลูกค้า"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   117
         Top             =   1620
         Width           =   4110
      End
      Begin VB.Label LBLPayItemArCode 
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
         Left            =   1800
         TabIndex        =   113
         Top             =   720
         Width           =   2445
      End
      Begin VB.Label Label42 
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
         Height          =   375
         Left            =   4365
         TabIndex        =   111
         Top             =   1620
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label LBLPayItemArName 
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
         Left            =   4320
         TabIndex        =   109
         Top             =   720
         Width           =   10230
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รหัส/ชื่อลูกค้า :"
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
         Left            =   315
         TabIndex        =   108
         Top             =   720
         Width           =   1410
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   -2160
      ScaleHeight     =   2190
      ScaleWidth      =   8760
      TabIndex        =   7
      Top             =   9045
      Visible         =   0   'False
      Width           =   8790
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
         Left            =   13275
         TabIndex        =   22
         Top             =   7065
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView103 
         Height          =   5460
         Left            =   585
         TabIndex        =   4
         Top             =   1440
         Width           =   13785
         _ExtentX        =   24315
         _ExtentY        =   9631
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
      Begin Crystal.CrystalReport Crystal101 
         Left            =   4410
         Top             =   7110
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Image Image3 
         Height          =   750
         Left            =   0
         Picture         =   "FrmQueue.frx":2AA1
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
         Left            =   4050
         TabIndex        =   11
         Top             =   720
         Width           =   2175
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
         Left            =   9855
         TabIndex        =   18
         Top             =   720
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
         Left            =   6615
         TabIndex        =   17
         Top             =   720
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
         Left            =   2700
         TabIndex        =   16
         Top             =   720
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
         Left            =   10980
         TabIndex        =   13
         Top             =   720
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
         Height          =   330
         Left            =   7650
         TabIndex        =   12
         Top             =   720
         Width           =   1500
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
         Width           =   1545
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0C0C0&
      Height          =   8475
      Left            =   7290
      ScaleHeight     =   8415
      ScaleWidth      =   16290
      TabIndex        =   61
      Top             =   9000
      Visible         =   0   'False
      Width           =   16350
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
         Width           =   8880
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
         Left            =   8415
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   5400
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
         Left            =   6930
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   5400
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
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่คิว :"
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
         Left            =   3870
         TabIndex        =   104
         Top             =   990
         Width           =   825
      End
      Begin VB.Label LBLQueDate 
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
         TabIndex        =   103
         Top             =   945
         Width           =   1815
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
         Left            =   7830
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
         Left            =   6795
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
         X2              =   9765
         Y1              =   4005
         Y2              =   4005
      End
      Begin VB.Line Line4 
         X1              =   810
         X2              =   9765
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Line Line3 
         X1              =   810
         X2              =   9765
         Y1              =   2205
         Y2              =   2205
      End
      Begin VB.Line Line2 
         X1              =   810
         X2              =   9765
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   0
         Picture         =   "FrmQueue.frx":3F03
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
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   7695
      ScaleHeight     =   4845
      ScaleWidth      =   7140
      TabIndex        =   8
      Top             =   45
      Width           =   7170
      Begin VB.PictureBox PICPoint2 
         Height          =   240
         Left            =   0
         ScaleHeight     =   180
         ScaleWidth      =   270
         TabIndex        =   143
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox Text102 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   2
         Top             =   495
         Width           =   1320
      End
      Begin MSComctlLib.ListView ListView102 
         Height          =   3570
         Left            =   90
         TabIndex        =   3
         Top             =   945
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   6297
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
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "คิวที่"
            Object.Width           =   1235
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
            Object.Width           =   3528
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
            Text            =   "ครั้งที่"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "พนักงานจัด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "พนักงานขาย"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "เริ่มจัด"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "วันที่คิว"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "เวลาที่ต้องการรับของ"
            Object.Width           =   3528
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
            Underline       =   -1  'True
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
      Height          =   3480
      Left            =   7695
      ScaleHeight     =   3450
      ScaleWidth      =   7140
      TabIndex        =   9
      Top             =   5040
      Width           =   7170
      Begin VB.PictureBox PICPoint3 
         Height          =   240
         Left            =   0
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   144
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin MSComctlLib.ListView ListView104 
         Height          =   2850
         Left            =   90
         TabIndex        =   5
         Top             =   450
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   5027
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
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "คิวที่"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "อ้างอิง"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "พนักงานจัด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "พนักงานขาย"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ประเภท"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "สถานะการจัด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "วันที่คิว"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ครั้งที่"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "เวลาที่ต้องการรับของ"
            Object.Width           =   3528
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
            Underline       =   -1  'True
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
      Height          =   3480
      Left            =   45
      ScaleHeight     =   3450
      ScaleWidth      =   7590
      TabIndex        =   23
      Top             =   5040
      Width           =   7620
      Begin VB.PictureBox PICReserve 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   0
         Picture         =   "FrmQueue.frx":5365
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
         Picture         =   "FrmQueue.frx":78B2
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
         Left            =   6975
         Picture         =   "FrmQueue.frx":9E80
         ScaleHeight     =   480
         ScaleWidth      =   435
         TabIndex        =   60
         Top             =   45
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox Pic101 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   6975
         Picture         =   "FrmQueue.frx":C615
         ScaleHeight     =   480
         ScaleWidth      =   435
         TabIndex        =   59
         Top             =   45
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label LBLReqTime 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   106
         Top             =   2520
         Width           =   5280
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ต้องการเวลา :"
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
         Left            =   45
         TabIndex        =   105
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Line Line13 
         X1              =   0
         X2              =   6840
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label LBLQueueDate 
         BackStyle       =   0  'Transparent
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
         Width           =   5280
      End
      Begin VB.Line Line12 
         X1              =   0
         X2              =   6840
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line11 
         X1              =   0
         X2              =   6840
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   6840
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   6840
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Line Line8 
         X1              =   6840
         X2              =   6840
         Y1              =   0
         Y2              =   3465
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   6840
         Y1              =   2475
         Y2              =   2475
      End
      Begin VB.Line Line6 
         X1              =   1395
         X2              =   1395
         Y1              =   0
         Y2              =   2880
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
         Top             =   3060
         Width           =   6765
      End
      Begin VB.Label LBLUserPick 
         BackStyle       =   0  'Transparent
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
         Width           =   5280
      End
      Begin VB.Label LBLRefNo 
         BackStyle       =   0  'Transparent
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
         Width           =   5280
      End
      Begin VB.Label LBLARName 
         BackStyle       =   0  'Transparent
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
         Width           =   5280
      End
      Begin VB.Label LBLSale 
         BackStyle       =   0  'Transparent
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
         Width           =   5280
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   45
      ScaleHeight     =   4845
      ScaleWidth      =   7590
      TabIndex        =   6
      Top             =   45
      Width           =   7620
      Begin VB.PictureBox PICPoint1 
         Height          =   285
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   142
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
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
         Left            =   6210
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
         Left            =   7290
         TabIndex        =   30
         Top             =   180
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
         Height          =   3570
         Left            =   90
         TabIndex        =   1
         Top             =   945
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   6297
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
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "คิวที่"
            Object.Width           =   1235
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
            Object.Width           =   3528
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
            Text            =   "วันที่คิว"
            Object.Width           =   2646
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
            Underline       =   -1  'True
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
Dim vQueID As Integer
Dim vDocDate As String
Dim vAnswer As Integer

If Me.TextScan101.Text <> "" And Me.LBLStatus101.Caption <> "" And Me.LBLDocDate.Caption <> "" Then
    vAnswer = MsgBox("คุณต้องการยกเลิกการรับสินค้าของลูกค้าใช่หรือไม่", vbYesNo, "Send Question Message")
    If vAnswer = 6 Then
       vQueID = Trim(TextScan101.Text)
       If Me.CHKDate.Value = 0 Then
       vDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       Else
       vDocDate = DateAdd("d", Now, -1)
       End If
       
       vQuery = "exec dbo.USP_NP_UpdateQueueCustRec " & vQueID & ",'" & vDocDate & "'"
       vConnection.Execute (vQuery)
       
       MsgBox "ยกเลิกคิวจัดสินค้าที่  " & vQueID & " เรียบร้อยแล้ว", vbInformation, "Send Information Message"
       
       Call RefreshQueueBegin
       Call RefreshQueuePicking
       Call RefreshQueueFinish
  
       Me.PicScanBar.Visible = False
       Me.Text101.SetFocus
    End If
End If
End Sub

Private Sub CMDDescription_Click()
Me.PTShowMyDescription.Visible = True
Me.TextMyDescription.SetFocus
End Sub

Private Sub CMDKeyItemQty_Click()
   Me.PICKeyItemQty.Visible = True
   Me.TBItemCode.Text = ""
   Me.LBLKeyQtyItemCodeIndex.Caption = ""
   Me.TBItemCode.SetFocus
End Sub

Private Sub CMDKeyItemQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICKeyItemQty.Visible = True
   Me.TBItemCode.Text = ""
   Me.TBItemCode.SetFocus
End If

If KeyCode = 27 Then
Call CMDPickItemToCustClose_Click
End If

If KeyCode = 116 Then
Call CMDPickItemToCust_Click
End If
End Sub

Sub CMDKeyQtyClose_Click()
Dim vIndex As Integer
If Me.LBLKeyQtyItemCode.Caption <> "" And Me.LBLKeyQtyItemName.Caption <> "" Then
vIndex = Me.LBLKeyQtyItemCodeIndex.Caption
Me.PICKeyItemQty.Visible = False
Me.ListViewPayItemToCust.SetFocus
Me.ListViewPayItemToCust.ListItems(vIndex).Selected = True
Else
Me.PICKeyItemQty.Visible = False
Me.ListViewPayItemToCust.SetFocus
End If
End Sub

Private Sub CMDKeyQtyClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDKeyQtyClose_Click
End If
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

Private Sub CMDPayItemSearch_Click()
Dim vRecordset As New ADODB.Recordset
Dim vSearch As String
Dim vZoneID As String
Dim i As Integer
Dim vPickQTY As Double
Dim vListItem  As ListItem
Dim vDocType As Integer

If Me.TBPayItemSearch.Text <> "" Then
   vSearch = Me.TBPayItemSearch.Text
   If vSelectZoneID = 1 Then
   vZoneID = "A"
   ElseIf vSelectZoneID = 2 Then
   vZoneID = "B"
   ElseIf vSelectZoneID = 3 Then
   vZoneID = "C"
   End If
   vDocType = Me.LBLDocType.Caption
   Me.ListViewPayItemToCust.ListItems.Clear
      
   vQuery = "exec dbo.USP_NP_SearchQuePayItem " & vDocType & "," & vZoneID & ",'" & vSearch & "' "
   If OpenDataBase2(vConnection, vRecordset, vQuery) <> 0 Then
   Me.LBLPayItemArCode.Caption = Trim(vRecordset.Fields("arcode").Value)
   Me.LBLPayItemArName.Caption = Trim(vRecordset.Fields("arname").Value)
   vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
           Set vListItem = FrmQueue.ListViewPayItemToCust.ListItems.Add(, , i)
           vListItem.SubItems(1) = ""
           vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
           vListItem.SubItems(3) = Trim(vRecordset.Fields("itemcode").Value)
           vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
           vPickQTY = Trim(vRecordset.Fields("pickqty").Value)
           vListItem.SubItems(5) = Format(vPickQTY, "##,##0.00")
           vListItem.SubItems(6) = Trim(vRecordset.Fields("queid").Value)
           vListItem.SubItems(7) = Trim(vRecordset.Fields("docno").Value)
       vRecordset.MoveNext
       Next i
   End If
   vRecordset.Close
   
   Me.ListViewPayItemToCust.SetFocus
End If
End Sub

Private Sub CMDPayItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICKeyItemQty.Visible = True
   Me.TBItemCode.Text = ""
   Me.TBItemCode.SetFocus
End If

If KeyCode = 27 Then
Me.PICPickItemToCust.Visible = False
Me.ListView104.SetFocus
End If

If KeyCode = 116 Then
Call CMDPickItemToCust_Click
End If
End Sub

Public Sub PrintSalePickingSlip(vQueID As Integer, vQueDocDate As String, vZone As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocno As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
Dim vQueNo As String
   

If vZone = "A" Then
 vQuery = "exec dbo.USP_NP_SearchPrinter 2"
 If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
     vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
 End If
 vRecordset.Close
    
If vPrinterName <> "" Then
   For Each printerObj In Printers
   If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
   Set Printer = printerObj
   Set printerObj = Nothing
   Exit For
   End If
   Next
Else
   Exit Sub
End If
End If

If vZone = "B" Then
 vQuery = "exec dbo.USP_NP_SearchPrinter 3"
 If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
     vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
 End If
 vRecordset.Close
    
   If vPrinterName <> "" Then
      For Each printerObj In Printers
      If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
      Set Printer = printerObj
      Set printerObj = Nothing
      Exit For
      End If
      Next
   Else
      Exit Sub
   End If
End If

        
vQuery = "exec dbo.USP_NP_SearchQueCenterDetails " & vQueID & ",'" & vQueDocDate & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then

vQueNo = Trim(vRecordset.Fields("queid").Value)

Printer.FontName = "AngsanaUPC"
Printer.Font.Size = 50
Printer.CurrentX = 1700
Printer.Print Trim(vRecordset.Fields("queid").Value)

'Printer.Font.Name = "Code128"
Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 1500
'Printer.Print "*" & Trim(vRecordset.Fields("queid").Value) & "*"
Printer.Print "*" & vQueNo & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1000
Printer.FontBold = True
Printer.Print Trim("Picking Request Slip Details")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("quedocdate").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value) & "          " & Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)

'Printer.Font.Name = "Code128"
Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("docno").Value) & "*"

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

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("ทะเบียนรถ : ") & Trim(vRecordset.Fields("refno").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("refno").Value) & "*"


Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value) & "     " & "โซนการจัด :" & Trim(vRecordset.Fields("quezone").Value)
vRecordset.MoveFirst

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
vRecordset.MoveFirst
n = 1
While Not vRecordset.EOF

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfid").Value) & "                 " & "  ยอดพอขายตามคลัง :  " & Trim(vRecordset.Fields("remainsale").Value)

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
Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
    
vRecordset.MoveNext
n = n + 1
Wend
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
Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"

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


Private Sub CMDPayQtyOK_Click()
Dim vIndex As Integer

vIndex = Me.LBLPayQtyIndex.Caption
Me.PICItemPayQty.Visible = False
Me.ListViewPayItemToCust.SetFocus
Me.ListViewPayItemToCust.ListItems(vIndex).Selected = True
End Sub

Private Sub CMDPayQtyOK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDPayQtyOK_Click
End If
End Sub

Private Sub CMDPickItemToCust_Click()
Dim vQueID As Integer
Dim vItemCode As String
Dim vQTY As Double
Dim i As Integer
Dim n As Integer
Dim m As Integer
Dim vCheckKeyQty As Integer
Dim vQueUpdate As Integer

If Me.ListViewPayItemToCust.ListItems.Count > 0 Then

For n = 1 To Me.ListViewPayItemToCust.ListItems.Count
   If Me.ListViewPayItemToCust.ListItems(n).SubItems(1) = "" Then
      vCheckKeyQty = 1
   Else
      vCheckKeyQty = 0
   End If
   If vCheckKeyQty = 1 Then
      MsgBox "กรุณากรอกจำนวนที่จ่ายสินค้า ให้ลูกค้าให้ครบตามรายการสินค้า ถึงจะบันทึกการจ่ายได้ ", vbCritical, "Send Error Message"
      Me.ListViewPayItemToCust.SetFocus
      Exit Sub
   End If
Next n

For i = 1 To Me.ListViewPayItemToCust.ListItems.Count
vQueID = Me.ListViewPayItemToCust.ListItems(i).SubItems(6)
vItemCode = Me.ListViewPayItemToCust.ListItems(i).SubItems(3)
vQTY = Me.ListViewPayItemToCust.ListItems(i).SubItems(1)

vQuery = "exec dbo.USP_NP_UpdatePayItemQtyQue " & vQueID & ",'" & vItemCode & "'," & vQTY & " "
vConnection.Execute (vQuery)



Next i

For m = 1 To Me.ListView104.ListItems.Count
   vQueUpdate = Me.ListView104.ListItems(m).Text
   If vQueID = vQueUpdate Then
      Me.ListView104.ListItems.Remove (m)
      GoTo Line1
   End If
Next m

Line1:
Me.ListViewPayItemToCust.ListItems.Clear
Me.LBLPayItemArCode.Caption = ""
Me.LBLPayItemArName.Caption = ""
Me.TBPayItemSearch.Text = ""
Me.ListView104.SetFocus
Me.PICPickItemToCust.Visible = False
Call StartTime
End If
End Sub

Private Sub CMDPickItemToCust_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICKeyItemQty.Visible = True
   Me.TBItemCode.Text = ""
   Me.TBItemCode.SetFocus
End If

If KeyCode = 27 Then
Call CMDPickItemToCustClose_Click
End If

If KeyCode = 116 Then
Call CMDPickItemToCust_Click
End If
End Sub

Private Sub CMDPickItemToCustClose_Click()
Dim vIndex As Integer

Me.PICPickItemToCust.Visible = False
Me.ListView104.SetFocus
vIndex = Me.LBLPayItemIndex.Caption

Me.ListView104.ListItems(vIndex).Selected = True
Me.ListView104.ListItems(vIndex).Checked = False
Call StartTime
End Sub

Private Sub CMDPickItemToCustClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICKeyItemQty.Visible = True
   Me.TBItemCode.Text = ""
   Me.TBItemCode.SetFocus
End If

If KeyCode = 27 Then
Call CMDPickItemToCustClose_Click
End If

If KeyCode = 116 Then
Call CMDPickItemToCust_Click
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
    Dim vQueDocDate As String
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
    
    vQueDocDate = Me.LBLQueDate.Caption
    vPickingNo = Me.LBLSendQueue.Caption
    vSaleOrderNo = Me.LBLSendRefNo.Caption
    vDescription = Me.TextDescription.Text
    
    vQuery = "exec dbo.USP_NP_UpdateQueReceived '" & vPickingNo & "','" & vQueDocDate & "','" & vSaleOrderNo & "'," & vStatus & " ,'" & vDescription & "' "
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



vConnectionString = "Provider = SQLOLEDB.1;Data Source = S02DB;Initial Catalog = BPLUS4;User ID =VBUSER;PassWord = 132"
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
  
  Call SetListViewColor(ListView101, PICPoint1, vbWhite, vbYellowBright)
  Call SetListViewColor(ListView102, PICPoint2, vbWhite, vbLightBlue)
  Call SetListViewColor(ListView104, PICPoint3, vbWhite, vbLightGreen)
  
  Call SetListViewColor(ListView105, PICPoint4, vbWhite, vbYellowBright)
  Call SetListViewColor(ListViewPayItemToCust, PICPoint5, vbWhite, vbLightGreen)
  
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
    vDocDate = Trim(ListView101.ListItems.Item(vIndex).SubItems(7))
    
    FrmPicker.LBLDocno.Caption = vPrintDocno
    FrmPicker.LBLID.Caption = vTimeID
    FrmPicker.LBLCustName.Caption = vCustName
    FrmPicker.LBLRefNo.Caption = vRefNo
    FrmPicker.LBLDocDate.Caption = vDocDate
    vCheckClickListview = 1
    
    i = 0
    vQuery = "exec dbo.USP_NP_SearchQueCenterDetails '" & vPrintDocno & "','" & vDocDate & "' "
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
Dim vReqTime As String
Dim vCheckTypePick As Integer

On Error Resume Next

If Me.ListView101.ListItems.Count > 0 Then

 vARName = Trim(ListView101.SelectedItem.ListSubItems(5))
 vRefNo = Trim(ListView101.SelectedItem.ListSubItems(3))
 vSaleCode = Trim(ListView101.SelectedItem.ListSubItems(1))
 vQueueDate = Trim(ListView101.SelectedItem.ListSubItems(7))
 vReqTime = Trim(ListView101.SelectedItem.ListSubItems(2))
 
 Pic101.Visible = False
 Pic102.Visible = False
 
 LBLQueueDate = vQueueDate
 LBLARName.Caption = vARName
 LBLRefNo.Caption = vRefNo
 LBLUserPick.Caption = Trim("-")
 LBLSale.Caption = vSaleCode
 Me.LBLReqTime.Caption = vReqTime
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
    vQuery = "exec dbo.USP_NP_SearchQueCenterDetails '" & vPickingNo & "','" & vDocDate & "' "
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
      vListItem.SubItems(6) = "รอจัดสินค้า"
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

Private Sub ListView101_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Me.ListView101.ListItems.Count > 0 Then
   Call ListView101_DblClick
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

'On Error Resume Next

If ListView102.ListItems.Count > 0 Then
  Call StopTime
  vIndex = ListView102.SelectedItem.Index
  vQueueID = ListView102.ListItems.Item(vIndex).Text
    
  If vIndex <> 0 Then
    vCheckClickListview = 2
    vIndexFinish = vIndex
    FrmCheckQTY.Show
    vPrintDocno = Trim(ListView102.ListItems.Item(vIndex).Text)
    vTimeID = Trim(ListView102.ListItems.Item(vIndex).SubItems(6))
    vSaleOrderNo = Trim(ListView102.ListItems.Item(vIndex).SubItems(3))
    vARCode = Trim(ListView102.ListItems.Item(vIndex).SubItems(5))
    vDocType = Trim(ListView102.ListItems.Item(vIndex).SubItems(4))
    vDocDate = Trim(ListView102.ListItems.Item(vIndex).SubItems(10))

    FrmCheckQTY.LBLDocno1.Caption = vPrintDocno
    FrmCheckQTY.LBLDocno2.Caption = vSaleOrderNo
    FrmCheckQTY.LBLARCode.Caption = vARCode
    FrmCheckQTY.LBLID.Caption = vTimeID
    FrmCheckQTY.LBLDocDate.Caption = vDocDate
    
    If vPrintDocno <> "" Then
    'i = 0
    vQuery = "exec dbo.USP_NP_SearchQueCenterDetails '" & vPrintDocno & "','" & vDocDate & "' "
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
      vListItem.SubItems(6) = Trim(vRecordset.Fields("whcode").Value)
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

Private Sub ListView102_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Me.ListView102.ListItems.Count > 0 Then
   Call ListView102_DblClick
End If
End Sub

Private Sub ListView103_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  Call CMDOK_Click
End If
End Sub

Private Sub ListView104_DblClick()
'Me.PICPickItemToCust.Visible = True
'Me.TBPayItemSearch.SetFocus
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

Private Sub ListViewPayItemToCust_DblClick()
Dim i As Integer

If Me.ListViewPayItemToCust.ListItems.Count > 0 Then
    Me.PICKeyItemQty.Visible = False
    Me.PICItemPayQty.Visible = True
    i = Me.ListViewPayItemToCust.SelectedItem.Index
    Me.LBLPayItemCode.Caption = Me.ListViewPayItemToCust.ListItems(i).SubItems(3)
    Me.LBLPayItemName.Caption = Me.ListViewPayItemToCust.ListItems(i).SubItems(2)
    Me.LBLPayQtyIndex.Caption = i
    Me.TBPayItemQty.Text = ""
    Me.TBPayItemQty.SetFocus
End If
End Sub

Private Sub ListViewPayItemToCust_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICKeyItemQty.Visible = True
   Me.TBItemCode.Text = ""
   Me.TBItemCode.SetFocus
End If

If KeyCode = 27 Then
Call CMDPickItemToCustClose_Click
End If

If KeyCode = 116 Then
Call CMDPickItemToCust_Click
End If
End Sub

Private Sub ListViewPayItemToCust_KeyPress(KeyAscii As Integer)
Dim i As Integer

If KeyAscii = 13 Then
If Me.ListViewPayItemToCust.ListItems.Count > 0 Then
    Me.PICKeyItemQty.Visible = False
    Me.PICItemPayQty.Visible = True
    i = Me.ListViewPayItemToCust.SelectedItem.Index
    Me.LBLPayItemCode.Caption = Me.ListViewPayItemToCust.ListItems(i).SubItems(3)
    Me.LBLPayItemName.Caption = Me.ListViewPayItemToCust.ListItems(i).SubItems(2)
    Me.LBLPayQtyIndex.Caption = i
    Me.TBPayItemQty.Text = ""
    Me.TBPayItemQty.SetFocus
End If
End If

End Sub

Private Sub MReserve_Click()
Dim i As Integer
Dim vAnswer As Integer
Dim vRecordset As New ADODB.Recordset
Dim vZoneID As String
Dim vQueID As Integer
Dim vDocDate As String


On Error GoTo ErrDescription

i = ListView101.SelectedItem.Index
vQueID = Trim(ListView101.ListItems.Item(i).Text)
vDocDate = ListView101.ListItems.Item(i).SubItems(7)
If vSelectZoneID = 1 Then
   vZoneID = "A"
ElseIf vSelectZoneID = 2 Then
   vZoneID = "B"
ElseIf vSelectZoneID = 3 Then
   vZoneID = "C"
ElseIf vSelectZoneID = 4 Then
   vZoneID = "D"
End If

vAnswer = MsgBox("ต้องการพิมพ์ทดแทนคิวที่  " & vQueID & " ใช่หรือไม่", vbYesNo, "Send Question ?")
If vAnswer = 6 Then
   Call PrintSalePickingSlip(vQueID, vDocDate, vZoneID)
   MsgBox "พิมพ์ทดแทนเรียบร้อย"
Else
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
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
Dim vReqTime As String
Dim vCheckTypePick As Integer

On Error Resume Next

If Me.ListView102.ListItems.Count > 0 Then
 vCheckClickListview = 2
 
 Pic101.Visible = False
 Pic102.Visible = False
 
 vARName = Trim(ListView102.SelectedItem.ListSubItems(5))
 vRefNo = Trim(ListView102.SelectedItem.ListSubItems(3))
 vPicker = Trim(ListView102.SelectedItem.ListSubItems(7))
 vSaleCode = Trim(ListView102.SelectedItem.ListSubItems(8))
 vQueueDate = Trim(ListView102.SelectedItem.ListSubItems(10))
 vReqTime = Trim(ListView102.SelectedItem.ListSubItems(11))
  
 LBLQueueDate.Caption = vQueueDate
 LBLARName.Caption = vARName
 LBLRefNo.Caption = vRefNo
 LBLUserPick.Caption = vPicker
 LBLSale.Caption = vSaleCode
 Me.LBLReqTime.Caption = vReqTime
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
    vDocDate = Trim(ListView102.ListItems.Item(vIndex).SubItems(10))
    vPickingNo = Trim(LBLDocno1.Caption)
    vSaleOrderNo = Trim(LBLDocno2.Caption)
    vTimes = Trim(ListView102.ListItems.Item(vIndex).SubItems(6))
    i = 0
    vQuery = "exec dbo.USP_NP_SearchQueCenterDetails '" & vPickingNo & "','" & vDocDate & "' "
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
      vListItem.SubItems(6) = "จัดยังไม่เสร็จ"
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
Dim vQueDocDate As String
Dim vIndex As Integer
Dim vSaleOrderNo As String
Dim vDocType As Integer

Dim vSearch As String
Dim vZoneID As String
Dim i As Integer
Dim vPickQTY As Double
Dim vListItem  As ListItem

On Error Resume Next

If ListView104.ListItems.Count > 0 Then
vIndexComplete = Item.Index
If Me.ListView104.ListItems(vIndexComplete).Checked = True Then
  vSaleOrderNo = Trim(ListView104.ListItems.Item(vIndexComplete).SubItems(2))
  vQueDocDate = Trim(ListView104.ListItems.Item(vIndexComplete).SubItems(7))
  vPickingNo = ListView104.ListItems.Item(vIndexComplete).Text
  
  Me.LBLPayItemIndex.Caption = vIndexComplete
  
  Me.LBLDocType.Caption = Trim(ListView104.ListItems.Item(vIndexComplete).SubItems(5))
  Me.PICPickItemToCust.Visible = True

   vSearch = vPickingNo
   If vSelectZoneID = 1 Then
   vZoneID = "A"
   ElseIf vSelectZoneID = 2 Then
   vZoneID = "B"
   ElseIf vSelectZoneID = 3 Then
   vZoneID = "C"
   End If
   
   FrmQueue.ListViewPayItemToCust.ListItems.Clear
   vQuery = "exec dbo.USP_NP_SearchQuePayItem 0," & vZoneID & ",'" & vSearch & "' "
   If OpenDataBase2(vConnection, vRecordset, vQuery) <> 0 Then
   Me.LBLPayItemQueID.Caption = Trim(vRecordset.Fields("queid").Value)
   Me.LBLPayItemArCode.Caption = Trim(vRecordset.Fields("arcode").Value)
   Me.LBLPayItemArName.Caption = Trim(vRecordset.Fields("arname").Value)
   vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
           Set vListItem = FrmQueue.ListViewPayItemToCust.ListItems.Add(, , i)
           vListItem.SubItems(1) = ""
           vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
           vListItem.SubItems(3) = Trim(vRecordset.Fields("itemcode").Value)
           vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
           vPickQTY = Trim(vRecordset.Fields("pickqty").Value)
           vListItem.SubItems(5) = Format(vPickQTY, "##,##0.00")
           vListItem.SubItems(6) = Trim(vRecordset.Fields("queid").Value)
           vListItem.SubItems(7) = Trim(vRecordset.Fields("docno").Value)
       vRecordset.MoveNext
       Next i
   End If
   vRecordset.Close
   Me.ListViewPayItemToCust.SetFocus
  
  Call StopTime
  End If
  End If
End Sub

Private Sub ListView104_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim vARName As String
Dim vRefNo As String
Dim vSaleCode As String
Dim vPicker As String
Dim vCustomerZone As String
Dim vQueueDate As String
Dim vReqTime As String
Dim vCheckTypePick As Integer

On Error Resume Next

 If Trim(ListView104.SelectedItem.ListSubItems(6)) = 1 Then
   Pic101.Visible = True
   Pic102.Visible = False
 ElseIf Trim(ListView104.SelectedItem.ListSubItems(6)) = 2 Then
   Pic101.Visible = False
   Pic102.Visible = True
 Else
   Pic101.Visible = False
   Pic102.Visible = False
 End If
 
vARName = Trim(ListView104.SelectedItem.ListSubItems(1))
vRefNo = Trim(ListView104.SelectedItem.ListSubItems(2))
vSaleCode = Trim(ListView104.SelectedItem.ListSubItems(4))
vPicker = Trim(ListView104.SelectedItem.ListSubItems(3))
vQueueDate = Trim(ListView104.SelectedItem.ListSubItems(7))
vReqTime = Trim(ListView104.SelectedItem.ListSubItems(9))


LBLQueueDate.Caption = vQueueDate
LBLARName.Caption = vARName
LBLRefNo.Caption = vRefNo
LBLUserPick.Caption = vPicker
LBLSale.Caption = vSaleCode
 Me.LBLReqTime.Caption = vReqTime
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
Dim vQTY As Double
Dim vPickQTY As Double

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
    vDocDate = Trim(ListView104.ListItems.Item(vIndex).SubItems(7))
    vPickingNo = Trim(LBLDocno1.Caption)
    vSaleOrderNo = Trim(LBLDocno2.Caption)
    vTimes = Trim(ListView104.ListItems.Item(vIndex).SubItems(4))
    i = 0

    vQuery = "exec dbo.USP_NP_SearchQueCenterDetails '" & vPickingNo & "','" & vDocDate & "' "
    If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    LBLPicker.Caption = Trim(vRecordset.Fields("picker").Value)
    LBLWHCode.Caption = Trim(vRecordset.Fields("whcode").Value)
    While Not vRecordset.EOF
      i = i + 1
      Set vListItem = ListView103.ListItems.Add(, , i)
      vQTY = Trim(vRecordset.Fields("qty").Value)
      vPickQTY = Trim(vRecordset.Fields("pickqty").Value)
      vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
      vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
      vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
      vListItem.SubItems(5) = Trim(vRecordset.Fields("shelfcode").Value)
      If vQTY > vPickQTY Then
          vListItem.SubItems(6) = "ไม่ครบ"
      ElseIf vPickQTY = 0 Then
          vListItem.SubItems(6) = "ไม่มีของ"
      ElseIf vQTY = vPickQTY Then
          vListItem.SubItems(6) = "ครบ"
      ElseIf vQTY < vPickQTY Then
          vListItem.SubItems(6) = "จัดเกิน"
      End If
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

Private Sub PICItemPayQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDPayQtyOK_Click
End If
End Sub

Private Sub PICKeyItemQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDKeyQtyClose_Click
End If
End Sub

Private Sub PICPickItemToCust_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICKeyItemQty.Visible = True
   Me.TBItemCode.Text = ""
   Me.TBItemCode.SetFocus
End If

If KeyCode = 27 Then
Call CMDPickItemToCustClose_Click
End If

If KeyCode = 116 Then
Call CMDPickItemToCust_Click
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

Private Sub TBItemCode_Change()
If Me.TBItemCode.Text = "" Then
Me.LBLKeyQtyItemCode.Caption = ""
Me.LBLKeyQtyItemName.Caption = ""
Me.TBKeyQty.Text = ""
Me.TBItemCode.SetFocus
End If
End Sub

Private Sub TBItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim vIndex As Integer

'If KeyCode = 27 Then
'Me.PICKeyItemQty.Visible = False

'vIndex = Me.LBLKeyQtyItemCodeIndex.Caption
'Me.ListViewPayItemToCust.SetFocus
'Me.ListViewPayItemToCust.ListItems(vIndex).Selected = True
'End If

If KeyCode = 27 Then
Call CMDKeyQtyClose_Click
End If

End Sub

Private Sub TBItemCode_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vBarCode As String
Dim vCheckPayQty As Double
Dim vItemCode As String
Dim vCheckItemCode As String
Dim i As Integer

If KeyAscii = 13 Then
      vBarCode = Me.TBItemCode.Text
      vQuery = "exec dbo.USP_MB_SearchBarcode '" & vBarCode & "'"
      If OpenDataBase2(vConnection, vRecordset, vQuery) <> 0 Then
          Me.LBLKeyQtyItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
          Me.LBLKeyQtyItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
          vItemCode = Trim(vRecordset.Fields("itemcode").Value)
          
        For i = 1 To Me.ListViewPayItemToCust.ListItems.Count
           vCheckItemCode = Me.ListViewPayItemToCust.ListItems(i).SubItems(3)
           
           If Me.ListViewPayItemToCust.ListItems(i).SubItems(1) <> "" Then
           vCheckPayQty = Me.ListViewPayItemToCust.ListItems(i).SubItems(1)
           End If
           
           If vItemCode = vCheckItemCode Then
              Me.LBLKeyQtyItemCodeIndex.Caption = i
           End If
           
           If vItemCode = vCheckItemCode And vCheckPayQty > 0 Then
              Me.TBKeyQty.Text = vCheckPayQty
           End If
           
        Next i
            
          Me.TBKeyQty.SetFocus
      Else
         MsgBox "ไม่พบรหัสสินค้า " & vBarCode & " นี้ในระบบ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
         Me.TBItemCode.SetFocus
      End If
      vRecordset.Close
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

Private Sub TBKeyQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDKeyQtyClose_Click
End If
End Sub

Private Sub TBKeyQty_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim n As Integer
Dim vCheckItemCode As String
Dim vItemCode As String
Dim vQTY As Double
Dim vCheckNotExist As Integer
Dim vCheckPayQty As Double

If KeyAscii = 13 Then
If Me.TBKeyQty.Text <> "" Then
   If Me.LBLKeyQtyItemCode.Caption <> "" Then
      vItemCode = Me.LBLKeyQtyItemCode.Caption
      For i = 1 To Me.ListViewPayItemToCust.ListItems.Count
         vCheckItemCode = Me.ListViewPayItemToCust.ListItems(i).SubItems(3)
         vCheckPayQty = Me.ListViewPayItemToCust.ListItems(i).SubItems(5)
         If vItemCode = vCheckItemCode Then
            vQTY = Me.TBKeyQty.Text
            Me.ListViewPayItemToCust.ListItems(i).SubItems(1) = Format(vQTY, "##,##0.00")
            
        If vCheckPayQty > vQTY Then
                 ListViewPayItemToCust.ListItems(i).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
        ElseIf vCheckPayQty < vQTY Then
                 ListViewPayItemToCust.ListItems(i).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
        ElseIf vCheckPayQty = vQTY Then
                 ListViewPayItemToCust.ListItems(i).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(1).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(2).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(3).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(4).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(5).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(6).ForeColor = "&H00004000"
        End If

            Me.TBItemCode.Text = ""
            Me.LBLKeyQtyItemCode.Caption = ""
            Me.LBLKeyQtyItemName.Caption = ""
            Me.TBKeyQty.Text = ""
            Me.TBItemCode.SetFocus
            Exit Sub
         End If
         vCheckNotExist = 1
      Next i
      
      If vCheckNotExist = 1 Then
         MsgBox "ไม่มีรหัสสินค้า " & vItemCode & " ในรายการที่จัดได้สินค้า กรุณาตรวจสอบ", vbCritical, "Send Error Message"
         Me.TBItemCode.Text = ""
      End If
   End If
   End If
End If


   
End Sub

Private Sub TBPayItemQty_Change()
Dim vQtyWord As String
Dim vLenQTY As Integer

If Me.TBPayItemQty.Text <> "" Then
   vQtyWord = Me.TBPayItemQty.Text
   CheckNumber (vQtyWord)
      
   If vCheckValueNumber = False Then
      MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      vLenQTY = Len(Me.TBPayItemQty.Text)
      Me.TBPayItemQty.Text = Left(Me.TBPayItemQty.Text, vLenQTY - 1)
      Me.TBPayItemQty.SetFocus
   End If
End If
End Sub

Private Sub TBPayItemQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Call CMDPayQtyOK_Click
End If
End Sub

Private Sub TBPayItemQty_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim vQTY As Double
Dim vCheckPayQty As Double


If KeyAscii = 13 Then
If Me.TBPayItemQty.Text <> "" Then
   If Me.LBLPayItemCode.Caption <> "" Then
      i = Me.LBLPayQtyIndex.Caption

            vQTY = Me.TBPayItemQty.Text
            vCheckPayQty = Me.ListViewPayItemToCust.ListItems(i).SubItems(5)
            Me.ListViewPayItemToCust.ListItems(i).SubItems(1) = Format(vQTY, "##,##0.00")
            
        If vCheckPayQty > vQTY Then
                 ListViewPayItemToCust.ListItems(i).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
        ElseIf vCheckPayQty < vQTY Then
                 ListViewPayItemToCust.ListItems(i).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
        ElseIf vCheckPayQty = vQTY Then
                 ListViewPayItemToCust.ListItems(i).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(1).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(2).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(3).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(4).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(5).ForeColor = "&H00004000"
                 ListViewPayItemToCust.ListItems.Item(i).ListSubItems(6).ForeColor = "&H00004000"
        End If

    Me.LBLPayItemCode.Caption = ""
    Me.LBLPayItemName.Caption = ""
    Me.LBLPayQtyIndex.Caption = ""
    Me.TBPayItemQty.Text = ""
    Me.PICItemPayQty.Visible = False
    Me.ListViewPayItemToCust.SetFocus
    
    If i < Me.ListViewPayItemToCust.ListItems.Count Then
    Me.ListViewPayItemToCust.ListItems(i + 1).Selected = True
    ElseIf i = Me.ListViewPayItemToCust.ListItems.Count Then
    Me.ListViewPayItemToCust.ListItems(i).Selected = True
    End If

   End If
End If
End If
End Sub

Private Sub TBPayItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICKeyItemQty.Visible = True
   Me.TBItemCode.Text = ""
   Me.TBItemCode.SetFocus
End If

If KeyCode = 27 Then
Me.PICPickItemToCust.Visible = False
Me.ListView104.SetFocus
End If

If KeyCode = 116 Then
Call CMDPickItemToCust_Click
End If
End Sub

Private Sub TBPayItemSearch_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vSearch As String
Dim vZoneID As String
Dim i As Integer
Dim vPickQTY As Double
Dim vListItem  As ListItem
Dim vDocType As Integer

If KeyAscii = 13 Then
   If Me.TBPayItemSearch.Text <> "" Then
      vSearch = Me.TBPayItemSearch.Text
      If vSelectZoneID = 1 Then
      vZoneID = "A"
      ElseIf vSelectZoneID = 2 Then
      vZoneID = "B"
      ElseIf vSelectZoneID = 3 Then
      vZoneID = "C"
      End If
      vDocType = Me.LBLDocType.Caption
      FrmQueue.ListViewPayItemToCust.ListItems.Clear
      vQuery = "exec dbo.USP_NP_SearchQuePayItem " & vDocType & "," & vZoneID & ",'" & vSearch & "' "
      If OpenDataBase2(vConnection, vRecordset, vQuery) <> 0 Then
      vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
              Set vListItem = FrmQueue.ListViewPayItemToCust.ListItems.Add(, , i)
              vListItem.SubItems(1) = ""
              vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
              vListItem.SubItems(3) = Trim(vRecordset.Fields("itemcode").Value)
              vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
              vPickQTY = Trim(vRecordset.Fields("pickqty").Value)
              vListItem.SubItems(5) = Format(vPickQTY, "##,##0.00")
              vListItem.SubItems(6) = Trim(vRecordset.Fields("queid").Value)
              vListItem.SubItems(7) = Trim(vRecordset.Fields("docno").Value)
          vRecordset.MoveNext
          Next i
      End If
      vRecordset.Close
   End If
End If
End Sub

Private Sub Text101_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
  Call StopTime
  PicScanBar.Visible = True
  TextScan101.SetFocus
End If

If KeyCode = 40 And Me.ListView101.ListItems.Count > 0 Then
   Me.ListView101.SetFocus
   Me.ListView101.ListItems(1).Selected = True
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
    vDocDate = Trim(ListView101.ListItems.Item(vIndex).SubItems(7))
    
    FrmPicker.LBLDocno.Caption = vPrintDocno
    FrmPicker.LBLID.Caption = vTimeID
    FrmPicker.LBLCustName.Caption = vCustName
    FrmPicker.LBLRefNo.Caption = vRefNo
    FrmPicker.LBLDocDate.Caption = vDocDate
    vCheckClickListview = 1
    
i = 0
'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails2
'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails3 '" & vPrintDocno & "','" & vRefNo & "'," & vDocType & "," & vTimeID & ",'" & vDocDate & "' "

vQuery = "exec dbo.USP_NP_SearchQueCenterDetails '" & vPrintDocno & "','" & vDocDate & "' "
    
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

If KeyCode = 40 And Me.ListView102.ListItems.Count > 0 Then
   Me.ListView102.SetFocus
   Me.ListView102.ListItems(1).Selected = True
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
    vDocDate = Trim(ListView102.ListItems.Item(vIndex).SubItems(10))
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
    'vQuery = "exec dbo.USP_NP_SearchQueueItemDetails3 '" & vPrintDocno & "','" & vSaleOrderNo & "'," & vDocType & "," & vTimeID & ",'" & vDocDate & "' "
    vQuery = "exec dbo.USP_NP_SearchQueCenterDetails '" & vPrintDocno & "','" & vDocDate & "' "
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
          vListItem.SubItems(6) = Trim(vRecordset.Fields("whcode").Value)
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
                                        Call CheckOut
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

Public Sub CheckOut()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAnswer As Integer
Dim vPickingNo As String
Dim vQueDocDate As String
Dim vIndex As Integer
Dim vSaleOrderNo As String
Dim vDocType As Integer

Dim vSearch As String
Dim vZoneID As String
Dim i As Integer
Dim vPickQTY As Double
Dim vListItem  As ListItem

On Error Resume Next

If ListView104.ListItems.Count > 0 Then
vIndexComplete = Me.ListView104.SelectedItem.Index
If Me.ListView104.ListItems(1).Checked = True Then
  vSaleOrderNo = Trim(ListView104.ListItems.Item(vIndexComplete).SubItems(2))
  vQueDocDate = Trim(ListView104.ListItems.Item(vIndexComplete).SubItems(7))
  vPickingNo = ListView104.ListItems.Item(vIndexComplete).Text
  
  Me.LBLPayItemIndex.Caption = vIndexComplete
  
  Me.LBLDocType.Caption = Trim(ListView104.ListItems.Item(vIndexComplete).SubItems(5))
  Me.PICPickItemToCust.Visible = True

   vSearch = vPickingNo
   If vSelectZoneID = 1 Then
   vZoneID = "A"
   ElseIf vSelectZoneID = 2 Then
   vZoneID = "B"
   ElseIf vSelectZoneID = 3 Then
   vZoneID = "C"
   End If
   
   FrmQueue.ListViewPayItemToCust.ListItems.Clear
   vQuery = "exec dbo.USP_NP_SearchQuePayItem 0," & vZoneID & ",'" & vSearch & "' "
   If OpenDataBase2(vConnection, vRecordset, vQuery) <> 0 Then
   Me.LBLPayItemQueID.Caption = Trim(vRecordset.Fields("queid").Value)
   Me.LBLPayItemArCode.Caption = Trim(vRecordset.Fields("arcode").Value)
   Me.LBLPayItemArName.Caption = Trim(vRecordset.Fields("arname").Value)
   vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
           Set vListItem = FrmQueue.ListViewPayItemToCust.ListItems.Add(, , i)
           vListItem.SubItems(1) = ""
           vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
           vListItem.SubItems(3) = Trim(vRecordset.Fields("itemcode").Value)
           vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
           vPickQTY = Trim(vRecordset.Fields("pickqty").Value)
           vListItem.SubItems(5) = Format(vPickQTY, "##,##0.00")
           vListItem.SubItems(6) = Trim(vRecordset.Fields("queid").Value)
           vListItem.SubItems(7) = Trim(vRecordset.Fields("docno").Value)
       vRecordset.MoveNext
       Next i
   End If
   vRecordset.Close
   Me.ListViewPayItemToCust.SetFocus
  
  Call StopTime
  End If
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
  
  vQuery = "exec dbo.USP_NP_SearchQueCenterDetails '" & vPrintDocno & "','" & vDocDate & "' "
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

vQuery = "exec dbo.USP_NP_SearchQueCenterDetails '" & vPrintDocno & "','" & vDocDate & "' "
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
      vListItem.SubItems(6) = Trim(vRecordset.Fields("whcode").Value)
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
Dim vQueID As Integer
Dim vDocDate As String
Dim vListviewItem As ListItem
Dim i As Integer
Dim vMyDescription As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
  vDocno = Trim(TextScan101.Text)
  vQueID = TextScan101.Text
  If Me.CHKDate.Value = 0 Then
  vDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
  Else
  vDocDate = DateAdd("d", Now, -1)
  End If
  'vQuery = "exec dbo.USP_NP_SearchDataQueueDetails1 '" & vDocno & "','" & vSelectZoneID & "','" & vDocDate & "' "
  vQuery = "exec dbo.USP_NP_SearchDataQueueDetails2 " & vQueID & "," & vSelectZoneID & ",'" & vDocDate & "' "
  If OpenDataBase(qConnection, vRecordset, vQuery) <> 0 Then
   vARName101 = Trim(vRecordset.Fields("arcode").Value)
   vARName102 = Trim(vRecordset.Fields("arname").Value)
   vSaleName101 = Trim(vRecordset.Fields("SaleName").Value)
   vPicker101 = Trim(vRecordset.Fields("Picker").Value)
   vWHCode101 = Trim(vRecordset.Fields("WHCode").Value)
   vRefNo101 = Trim(vRecordset.Fields("docno").Value)
   vDiffDateTime101 = Trim(vRecordset.Fields("diffpicking").Value)
   vStatus101 = Trim(vRecordset.Fields("questatus").Value)
   vIsReceived101 = Trim(vRecordset.Fields("quereceived").Value)
   vDocType101 = Trim(vRecordset.Fields("sourceid").Value)
   vTimeID101 = Trim(vRecordset.Fields("quetime").Value)
   vTimeID = Trim(vRecordset.Fields("quetime").Value)
   vDocDate101 = Trim(vRecordset.Fields("quedocdate").Value)
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

    vQuery = "exec dbo.USP_NP_SearchQueCenterDetails '" & vQueID & "','" & vDocDate & "' "
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
vQuery = "exec dbo.USP_NP_SearchQueCenterBegin " & vSelectZoneID & " "
If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
  If vCheckClickListview = 1 Then
    vARName = Trim(vRecordset.Fields("arname").Value)
    vRefNo = Trim(vRecordset.Fields("saleorderno").Value)
    vPicker = Trim(vRecordset.Fields("picker").Value)
    vSaleCode = Trim(vRecordset.Fields("salename").Value)
    vQueueDate = Trim(vRecordset.Fields("docdate").Value)

    FrmQueue.LBLQueueDate.Caption = vQueueDate
    FrmQueue.LBLARName.Caption = vARName
    FrmQueue.LBLRefNo.Caption = vRefNo
    FrmQueue.LBLUserPick.Caption = Trim("-")
    FrmQueue.LBLSale.Caption = vSaleCode
    FrmQueue.LBLCustomerZone.Caption = vCustomerZone
  End If
    While Not vRecordset.EOF
        Set vListItem = FrmQueue.ListView101.ListItems.Add(, , Trim(vRecordset.Fields("queid").Value))
        vListItem.SubItems(1) = Trim(vRecordset.Fields("salename").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("quereqtime").Value)
        vListItem.SubItems(3) = Trim(vRecordset.Fields("docno").Value)
        vListItem.SubItems(4) = Trim(vRecordset.Fields("sourceid").Value)
        vListItem.SubItems(5) = Trim(vRecordset.Fields("arname").Value)
        vListItem.SubItems(6) = Trim(vRecordset.Fields("quetime").Value)
        vListItem.SubItems(7) = Trim(vRecordset.Fields("quedocdate").Value)
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


vPrinterName = Trim("\\nova\EPS TM-T88III-NP")
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
   vPrinterName = Trim("\\nova\EPS-TM-PickingOutlet")
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


Public Sub SetListViewColor(pCtrlListView As ListView, _
                            pCtrlPictureBox As PictureBox, _
                            Color1 As Long, _
                            Color2 As Long)

On Error GoTo SetListViewColor_Error

    Dim iLineHeight As Long
    Dim iBarHeight  As Long
    Dim lBarWidth   As Long
    Dim lColor1     As Long
    Dim lColor2     As Long
 
    lColor1 = Color1
    lColor2 = Color2
    
    If pCtrlListView.View = lvwReport Then
        pCtrlListView.Picture = LoadPicture("")
        pCtrlListView.Refresh
        pCtrlPictureBox.Cls
        
        pCtrlPictureBox.AutoRedraw = True
        pCtrlPictureBox.BorderStyle = vbBSNone
        pCtrlPictureBox.ScaleMode = vbTwips
        pCtrlPictureBox.Visible = False
        
        pCtrlListView.PictureAlignment = lvwTile
        pCtrlPictureBox.Font = pCtrlListView.Font
        pCtrlPictureBox.Top = pCtrlListView.Top
        pCtrlPictureBox.Font = pCtrlListView.Font
        With pCtrlPictureBox.Font
            .Size = pCtrlListView.Font.Size '+ 2.75
            .Bold = pCtrlListView.Font.Bold
            .Charset = pCtrlListView.Font.Charset
            .Italic = pCtrlListView.Font.Italic
            .Name = pCtrlListView.Font.Name
            .Strikethrough = pCtrlListView.Font.Strikethrough
            .Underline = pCtrlListView.Font.Underline
            .Weight = pCtrlListView.Font.Weight
        End With
        pCtrlPictureBox.Refresh
        iLineHeight = pCtrlPictureBox.TextHeight("W") + Screen.TwipsPerPixelY
    
        iBarHeight = (iLineHeight * 1)
        lBarWidth = pCtrlListView.Width
    
        pCtrlPictureBox.Height = iBarHeight * 2
        pCtrlPictureBox.Width = lBarWidth
    
        pCtrlPictureBox.Line (0, 0)-(lBarWidth, iBarHeight), lColor1, BF
        pCtrlPictureBox.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), lColor2, BF
    
        pCtrlPictureBox.AutoSize = True
        pCtrlListView.Picture = pCtrlPictureBox.Image
    Else
        pCtrlListView.Picture = LoadPicture("")
    End If
    
    pCtrlListView.Refresh
    Exit Sub
SetListViewColor_Error:
    pCtrlListView.Picture = LoadPicture("")
    pCtrlListView.Refresh
End Sub

