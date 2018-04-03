VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form110 
   Caption         =   "หน้าบันทึกข้อมูลการตรวจนับสินค้า ตามชั้นเก็บ"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form110.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICKeyQTY 
      BackColor       =   &H00404040&
      Height          =   8025
      Left            =   0
      ScaleHeight     =   7965
      ScaleWidth      =   11790
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   11850
      Begin MSComctlLib.ListView ListViewQTY 
         Height          =   1545
         Left            =   2880
         TabIndex        =   46
         Top             =   5535
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   2725
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ชั้นเก็บ"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "จำนวน"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "หน่วย"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.CommandButton CMDCancel 
         Caption         =   "ยกเลิก"
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
         Left            =   8010
         TabIndex        =   28
         Top             =   4635
         Width           =   1770
      End
      Begin VB.CommandButton CMDOK 
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
         Height          =   510
         Left            =   6075
         TabIndex        =   27
         Top             =   4635
         Width           =   1770
      End
      Begin VB.TextBox TextCountQTY 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   8010
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   3915
         Width           =   1770
      End
      Begin VB.TextBox TextCheckQTY 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   2880
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   3915
         Width           =   1770
      End
      Begin VB.TextBox TextSHW 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   8010
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   3195
         Width           =   1770
      End
      Begin VB.TextBox TextVND 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   8010
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   2700
         Width           =   1770
      End
      Begin VB.TextBox TextQTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000008&
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
         Height          =   780
         Left            =   2880
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   2700
         Width           =   3660
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ยอดคงเหลือตามคลัง ณ ปัจจุบัน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   2880
         TabIndex        =   47
         Top             =   5310
         Width           =   2760
      End
      Begin VB.Label LBLOnHand 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6525
         TabIndex        =   44
         Top             =   2205
         Width           =   3255
      End
      Begin VB.Label LBLUnitcode 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2880
         TabIndex        =   43
         Top             =   2205
         Width           =   1770
      End
      Begin VB.Label LBLItemname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2880
         TabIndex        =   42
         Top             =   1710
         Width           =   6900
      End
      Begin VB.Label LBLItemCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2880
         TabIndex        =   29
         Top             =   1215
         Width           =   1770
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "อยู่ใน SHW :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   285
         Left            =   6390
         TabIndex        =   41
         Top             =   3195
         Width           =   1500
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "อยู่ใน VND :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   285
         Left            =   6525
         TabIndex        =   40
         Top             =   2700
         Width           =   1365
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   0
         Picture         =   "Form110.frx":72FB
         Top             =   0
         Width           =   2160
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "ยอดตรวจนับ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Left            =   6255
         TabIndex        =   39
         Top             =   3915
         Width           =   1635
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "ยอดตรวจสอบ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   285
         Left            =   585
         TabIndex        =   38
         Top             =   3915
         Width           =   2175
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "ยอดสุทธิ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   285
         Left            =   1305
         TabIndex        =   37
         Top             =   2970
         Width           =   1500
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "ชื่อสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   285
         Left            =   1395
         TabIndex        =   36
         Top             =   1710
         Width           =   1365
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "ยอดคงเหลือในคลัง :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   240
         Left            =   4455
         TabIndex        =   35
         Top             =   2205
         Width           =   1950
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "หน่วยนับ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   1485
         TabIndex        =   34
         Top             =   2205
         Width           =   1275
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
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
         ForeColor       =   &H80000014&
         Height          =   240
         Left            =   1755
         TabIndex        =   33
         Top             =   1215
         Width           =   1005
      End
   End
   Begin VB.PictureBox PICSearchShelf 
      BackColor       =   &H00808080&
      Height          =   8070
      Left            =   0
      ScaleHeight     =   8010
      ScaleWidth      =   21690
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   21750
      Begin VB.CommandButton CMDSearchShelfDetails 
         Height          =   285
         Left            =   4140
         Picture         =   "Form110.frx":9D85
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   855
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
         Left            =   9315
         TabIndex        =   18
         Top             =   6615
         Width           =   1095
      End
      Begin VB.CommandButton CMDSealect 
         Caption         =   "เลือก"
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
         Left            =   7965
         TabIndex        =   17
         Top             =   6615
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListViewShelf 
         Height          =   5100
         Left            =   945
         TabIndex        =   16
         Top             =   1305
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   8996
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสที่เก็บ "
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อที่เก็บ"
            Object.Width           =   12700
         EndProperty
      End
      Begin VB.TextBox TextSearch 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   855
         Width           =   2310
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "คำค้นหา :"
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
         Left            =   990
         TabIndex        =   14
         Top             =   855
         Width           =   825
      End
   End
   Begin VB.CommandButton CMDClear 
      Caption         =   "เคลียร์หน้าจอ"
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
      Left            =   8820
      Picture         =   "Form110.frx":A152
      TabIndex        =   8
      Top             =   6795
      Width           =   1185
   End
   Begin VB.CommandButton CMDAddItemList 
      Caption         =   "ลงตาราง"
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
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.ComboBox CMBShelf 
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
      Left            =   4545
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1260
      Width           =   1725
   End
   Begin MSComCtl2.DTPicker DTPDocDate 
      Height          =   330
      Left            =   7695
      TabIndex        =   4
      Top             =   1260
      Width           =   1905
      _ExtentX        =   3360
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
      Format          =   61407233
      CurrentDate     =   39363
   End
   Begin VB.CommandButton CMDClose 
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
      Left            =   10170
      TabIndex        =   9
      Top             =   6795
      Width           =   1185
   End
   Begin VB.CommandButton CMDSave 
      Caption         =   "บันทึก"
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
      Left            =   7470
      TabIndex        =   7
      Top             =   6795
      Width           =   1185
   End
   Begin MSComctlLib.ListView ListViewItemList 
      Height          =   4110
      Left            =   405
      TabIndex        =   6
      Top             =   2475
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   7250
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "ยอดรวมคลัง"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "จำนวนสุทธิ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "จำนวนตรวจสอบ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "จำนวนที่นับได้"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "หน่วยนับ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "คลัง"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ที่เก็บ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "ชั้นเก็บ VND"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "ชั้นเก็บ SHW"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "โซน"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.TextBox TextItemCode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2250
      TabIndex        =   5
      Top             =   1665
      Width           =   2085
   End
   Begin VB.CommandButton CMDSearchShelf 
      Height          =   330
      Left            =   6300
      Picture         =   "Form110.frx":A4C3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1260
      Width           =   330
   End
   Begin VB.ComboBox CMBWHCode 
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
      Height          =   315
      Left            =   2250
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   855
      Width           =   1230
   End
   Begin VB.ComboBox CMBZone 
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
      Left            =   2250
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1260
      Width           =   1230
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "โซน :"
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
      Left            =   1395
      TabIndex        =   48
      Top             =   1305
      Width           =   780
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการสินค้า "
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
      Height          =   195
      Left            =   405
      TabIndex        =   45
      Top             =   2205
      Width           =   1410
   End
   Begin VB.Label LBLRefNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   7695
      TabIndex        =   31
      Top             =   855
      Width           =   1905
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "อ้างอิง :"
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
      Left            =   6435
      TabIndex        =   30
      Top             =   855
      Width           =   1185
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่ :"
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
      Left            =   6660
      TabIndex        =   20
      Top             =   1260
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   1
      X1              =   -585
      X2              =   11835
      Y1              =   6660
      Y2              =   6660
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   0
      X1              =   -135
      X2              =   11835
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "คีย์รหัสสินค้า :"
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
      Left            =   720
      TabIndex        =   12
      Top             =   1710
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ที่เก็บ :"
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
      Left            =   3735
      TabIndex        =   11
      Top             =   1305
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "คลัง :"
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
      Left            =   720
      TabIndex        =   10
      Top             =   855
      Width           =   1455
   End
End
Attribute VB_Name = "Form110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vItemCode As String
Dim vCheckSameValue As Integer
Dim vSumQTY As Integer


Private Sub CMBShelf_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i  As Integer
Dim vDocNo As String
Dim vDocDate As String
Dim vListItem As ListItem

On Error GoTo ErrDescription

If Me.CMBWHCode.Text <> "" Then
   vDocNo = Trim(Me.CMBWHCode.Text & "-" & Me.CMBShelf.Text)
   vDocDate = Me.DTPDocdate.Day & "/" & Me.DTPDocdate.Month & "/" & Me.DTPDocdate.Year
   vQuery = "exec dbo.USP_MB_SearchShelfStockCount '" & vDocNo & "','" & vDocDate & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   Me.ListViewItemList.ListItems.Clear
   Me.LBLRefNo.Caption = vRecordset.Fields("docno").Value
   Call ClearScreen
   vRecordset.MoveFirst
   i = 1
   While Not vRecordset.EOF
   Set vListItem = Me.ListViewItemList.ListItems.Add(, , i)
   vListItem.SubItems(1) = vRecordset.Fields("itemcode").Value
   vListItem.SubItems(2) = vRecordset.Fields("itemname").Value
   vListItem.SubItems(3) = Format(vRecordset.Fields("onhand").Value, "####0.00")
   vListItem.SubItems(4) = Format(vRecordset.Fields("qty").Value, "####0.00")
   vListItem.SubItems(5) = Format(vRecordset.Fields("checkqty").Value, "####0.00")
   vListItem.SubItems(6) = Format(vRecordset.Fields("countqty").Value, "####0.00")
   vListItem.SubItems(7) = vRecordset.Fields("unitcode").Value
   vListItem.SubItems(8) = vRecordset.Fields("whcode").Value
   vListItem.SubItems(9) = vRecordset.Fields("shelfcode").Value
   vListItem.SubItems(10) = Format(vRecordset.Fields("vnd").Value, "####0.00")
   vListItem.SubItems(11) = Format(vRecordset.Fields("shw").Value, "####0.00")
   i = i + 1
   vRecordset.MoveNext
   Wend
   Else
      Call ClearScreen
      Me.ListViewItemList.ListItems.Clear
      Me.LBLRefNo.Caption = ""
   End If
   vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMBShelf_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If Me.ListViewItemList.ListItems.Count > 0 Then
      If KeyCode = 116 Then
      Call CMDSave_Click
      Me.CMBShelf.SetFocus
   End If
End If
End Sub

Private Sub CMBWHCode_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String
Dim vZoneCode As String

On Error Resume Next

vWHCode = Trim(CMBWHCode.Text)
vZoneCode = Trim(CMBZone.Text)

CMBShelf.Clear
If vWHCode <> "" And vZoneCode <> "" Then
   vQuery = "exec dbo.USP_MB_SearchShelfID '" & vWHCode & "','" & vZoneCode & "',''  "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
           vRecordset.MoveFirst
           While Not vRecordset.EOF
           CMBShelf.AddItem Trim(vRecordset.Fields("shelfid").Value)
           vRecordset.MoveNext
           Wend
   End If
   vRecordset.Close
End If
End Sub

Private Sub CMBWHCode_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If Me.ListViewItemList.ListItems.Count > 0 Then
      If KeyCode = 116 Then
      Call CMDSave_Click
      Me.CMBShelf.SetFocus
   End If
End If
End Sub

Private Sub CMBZone_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String
Dim vZoneCode As String

On Error Resume Next

vWHCode = Trim(CMBWHCode.Text)
vZoneCode = Trim(CMBZone.Text)

CMBShelf.Clear
If vWHCode <> "" And vZoneCode <> "" Then
   vQuery = "exec dbo.USP_MB_SearchShelfID '" & vWHCode & "','" & vZoneCode & "',''  "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
           vRecordset.MoveFirst
           While Not vRecordset.EOF
           CMBShelf.AddItem Trim(vRecordset.Fields("shelfid").Value)
           vRecordset.MoveNext
           Wend
   End If
   vRecordset.Close
End If
End Sub

Private Sub CMBZone_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If Me.ListViewItemList.ListItems.Count > 0 Then
      If KeyCode = 116 Then
      Call CMDSave_Click
      Me.CMBShelf.SetFocus
   End If
End If
End Sub

Private Sub CMDAddItemList_Click()
Dim vListItem As ListItem
Dim vItemCode As String
Dim vItemName As String
Dim vUnitCode As String
Dim vOnHand As Double
Dim vQty As Double
Dim vCheckQty As Double
Dim vCountQty As Double
Dim vWHCode As String
Dim vShelfCode As String
Dim i As Integer
Dim n As Integer
Dim vCheckItemExist As String
Dim vItemExist As Integer

On Error GoTo ErrDescription

If Me.CMBWHCode.Text <> "" And Me.CMBShelf.Text <> "" And Me.LBLItemCode.Caption <> "" And Me.LBLItemName.Caption <> "" And Me.LBLUnitcode.Caption <> "" And Me.TextQTY.Text <> "" Then
   vItemCode = Me.LBLItemCode.Caption
   vItemName = Me.LBLItemName.Caption
   vUnitCode = Me.LBLUnitcode.Caption
   vOnHand = Me.LBLOnHand.Caption
   vQty = Me.TextQTY.Text
   vCheckQty = Me.TextCheckQTY.Text
   vCountQty = Me.TextCountQTY.Text
   vWHCode = Me.CMBWHCode.Text
   vShelfCode = Me.CMBShelf.Text
   
   If Me.ListViewItemList.ListItems.Count > 0 Then

      
      For n = 1 To Me.ListViewItemList.ListItems.Count
      vCheckItemExist = Me.ListViewItemList.ListItems(n).SubItems(1)
      If vItemCode = vCheckItemExist Then
         'Me.ListViewItemList.ListItems(n).SubItems(3) = Format(vOnHand, "####0.00")
         Me.ListViewItemList.ListItems(n).SubItems(4) = Format(vQty, "####0.00")
         Me.ListViewItemList.ListItems(n).SubItems(5) = Format(vCheckQty, "####0.00")
         Me.ListViewItemList.ListItems(n).SubItems(6) = Format(vCountQty, "####0.00")
         Call ClearScreen
         Exit Sub
      End If
      Next n
      If vItemExist = 0 Then
          i = Me.ListViewItemList.ListItems.Count + 1
          Set vListItem = Me.ListViewItemList.ListItems.Add(, , i)
          vListItem.SubItems(1) = vItemCode
          vListItem.SubItems(2) = vItemName
          vListItem.SubItems(3) = Format(vOnHand, "####0.00")
          vListItem.SubItems(4) = Format(vQty, "####0.00")
          vListItem.SubItems(5) = Format(vCheckQty, "####0.00")
          vListItem.SubItems(6) = Format(vCountQty, "####0.00")
          vListItem.SubItems(7) = vUnitCode
          vListItem.SubItems(8) = vWHCode
          vListItem.SubItems(9) = vShelfCode
          Call ClearScreen
      End If
   Else
         i = Me.ListViewItemList.ListItems.Count + 1
        Set vListItem = Me.ListViewItemList.ListItems.Add(, , i)
        vListItem.SubItems(1) = vItemCode
        vListItem.SubItems(2) = vItemName
        vListItem.SubItems(3) = Format(vOnHand, "####0.00")
        vListItem.SubItems(4) = Format(vQty, "####0.00")
        vListItem.SubItems(5) = Format(vCheckQty, "####0.00")
        vListItem.SubItems(6) = Format(vCountQty, "####0.00")
        vListItem.SubItems(7) = vUnitCode
        vListItem.SubItems(8) = vWHCode
        vListItem.SubItems(9) = vShelfCode
        Call ClearScreen
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ClearScreen()
On Error Resume Next

   vCheckSameValue = 0
   Me.LBLItemCode.Caption = ""
   Me.LBLItemName.Caption = ""
   Me.LBLUnitcode.Caption = ""
   Me.LBLOnHand.Caption = ""
   Me.TextQTY.Text = ""
   Me.TextVND.Text = ""
   Me.TextSHW.Text = ""
   Me.TextCheckQTY.Text = ""
   Me.TextCountQTY.Text = ""
   Me.TextItemCode.SetFocus
End Sub

Private Sub CMDCancel_Click()
Me.PICKeyQTY.Visible = False
vCheckSameValue = 0
vSumQTY = 0
End Sub

Private Sub CMDClear_Click()
On Error Resume Next

Call ClearScreen
Me.LBLRefNo.Caption = ""
Me.CMBWHCode.Text = "014"
Me.DTPDocdate.Value = Now
Me.ListViewItemList.ListItems.Clear
End Sub

Private Sub CMDClear_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If Me.ListViewItemList.ListItems.Count > 0 Then
      If KeyCode = 116 Then
      Call CMDSave_Click
      Me.CMBShelf.SetFocus
   End If
End If
End Sub

Private Sub CMDClose_Click()
Unload Form110
End Sub

Private Sub CMDClose_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If Me.ListViewItemList.ListItems.Count > 0 Then
      If KeyCode = 116 Then
      Call CMDSave_Click
      Me.CMBShelf.SetFocus
   End If
End If
End Sub

Private Sub CMDExit_Click()
PICSearchShelf.Visible = False
End Sub

Private Sub CMDOK_Click()
Dim vListItem As ListItem
Dim vItemCode As String
Dim vItemName As String
Dim vUnitCode As String
Dim vOnHand As Double
Dim vQty As Double
Dim vCheckQty As Double
Dim vCountQty As Double
Dim vWHCode As String
Dim vZoneCode As String
Dim vShelfCode As String
Dim vVNDQty As Double
Dim vSHWQty As Double
Dim i As Integer
Dim n As Integer
Dim vCheckItemExist As String
Dim vItemExist As Integer

On Error GoTo ErrDescription

If Me.CMBWHCode.Text <> "" And Me.CMBShelf.Text <> "" And Me.LBLItemCode.Caption <> "" And Me.LBLItemName.Caption <> "" And Me.LBLUnitcode.Caption <> "" And Me.TextQTY.Text <> "" Then
   vItemCode = Me.LBLItemCode.Caption
   vItemName = Me.LBLItemName.Caption
   vUnitCode = Me.LBLUnitcode.Caption
   vOnHand = Me.LBLOnHand.Caption
   vQty = Me.TextQTY.Text
   vCheckQty = Me.TextCheckQTY.Text
   vCountQty = Me.TextCountQTY.Text
   vWHCode = Me.CMBWHCode.Text
   vZoneCode = Me.CMBZone.Text
   vShelfCode = Me.CMBShelf.Text
   vVNDQty = Me.TextVND.Text
   vSHWQty = Me.TextSHW.Text
   
   If Me.ListViewItemList.ListItems.Count > 0 Then

      
      For n = 1 To Me.ListViewItemList.ListItems.Count
      vCheckItemExist = Me.ListViewItemList.ListItems(n).SubItems(1)
      If vItemCode = vCheckItemExist Then
         'Me.ListViewItemList.ListItems(n).SubItems(3) = Format(vOnHand, "####0.00")
         Me.ListViewItemList.ListItems(n).SubItems(4) = Format(vQty, "####0.00")
         Me.ListViewItemList.ListItems(n).SubItems(5) = Format(vCheckQty, "####0.00")
         Me.ListViewItemList.ListItems(n).SubItems(6) = Format(vCountQty, "####0.00")
         Me.ListViewItemList.ListItems(n).SubItems(10) = Format(vVNDQty, "####0.00")
         Me.ListViewItemList.ListItems(n).SubItems(11) = Format(vSHWQty, "####0.00")
         Me.PICKeyQTY.Visible = False
         Call ClearScreen
         Exit Sub
      End If
      Next n
      If vItemExist = 0 Then
          i = Me.ListViewItemList.ListItems.Count + 1
          Set vListItem = Me.ListViewItemList.ListItems.Add(, , i)
          vListItem.SubItems(1) = vItemCode
          vListItem.SubItems(2) = vItemName
          vListItem.SubItems(3) = Format(vOnHand, "####0.00")
          vListItem.SubItems(4) = Format(vQty, "####0.00")
          vListItem.SubItems(5) = Format(vCheckQty, "####0.00")
          vListItem.SubItems(6) = Format(vCountQty, "####0.00")
          vListItem.SubItems(7) = vUnitCode
          vListItem.SubItems(8) = vWHCode
          vListItem.SubItems(9) = vShelfCode
          vListItem.SubItems(10) = Format(vVNDQty, "####0.00")
          vListItem.SubItems(11) = Format(vSHWQty, "####0.00")
          vListItem.SubItems(12) = vZoneCode
          Call ClearScreen
          Me.PICKeyQTY.Visible = False
      End If
   Else
         i = Me.ListViewItemList.ListItems.Count + 1
        Set vListItem = Me.ListViewItemList.ListItems.Add(, , i)
        vListItem.SubItems(1) = vItemCode
        vListItem.SubItems(2) = vItemName
        vListItem.SubItems(3) = Format(vOnHand, "####0.00")
        vListItem.SubItems(4) = Format(vQty, "####0.00")
        vListItem.SubItems(5) = Format(vCheckQty, "####0.00")
        vListItem.SubItems(6) = Format(vCountQty, "####0.00")
        vListItem.SubItems(7) = vUnitCode
        vListItem.SubItems(8) = vWHCode
        vListItem.SubItems(9) = vShelfCode
        vListItem.SubItems(10) = Format(vVNDQty, "####0.00")
        vListItem.SubItems(11) = Format(vSHWQty, "####0.00")
        vListItem.SubItems(12) = vZoneCode
        Call ClearScreen
        Me.PICKeyQTY.Visible = False
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDOK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICKeyQTY.Visible = False
   vCheckSameValue = 0
   vSumQTY = 0
End If
End Sub

Private Sub CMDSave_Click()
Dim vQuery As String
Dim vIsOpen As Integer
Dim vDocNo As String
Dim vDocDate As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vItemCode As String
Dim vItemName As String
Dim vOnHand As Double
Dim vQty As Double
Dim vCheckQty As Double
Dim vCountQty As Double
Dim vVND As Double
Dim vSHW As Double
Dim vUnitCode As String
Dim vLineNumber As Integer
Dim i As Integer


If Me.ListViewItemList.ListItems.Count > 0 Then
   vDocDate = Me.DTPDocdate.Day & "/" & Me.DTPDocdate.Month & "/" & Me.DTPDocdate.Year
   vWHCode = UCase(Me.CMBWHCode.Text)
   vShelfCode = UCase(Me.CMBShelf.Text)
   If Me.LBLRefNo.Caption <> "" Then
     vIsOpen = 1
     vDocNo = UCase(Me.LBLRefNo.Caption)
   Else
      vIsOpen = 0
      vDocNo = Trim(vWHCode & "-" & vShelfCode)
   End If

On Error GoTo ErrRollBackTran

   vQuery = "begin tran"
   gConnection.Execute vQuery
   
   vQuery = "exec dbo.USP_MB_InsertDataShelfStockCountHeader " & vIsOpen & ",'" & vDocNo & "','" & vDocDate & "','" & vWHCode & "','" & vShelfCode & "','" & vUserID & "' "
  gConnection.Execute vQuery
   
   For i = 1 To Me.ListViewItemList.ListItems.Count
   vItemCode = Me.ListViewItemList.ListItems(i).SubItems(1)
   vItemName = Me.ListViewItemList.ListItems(i).SubItems(2)
   vOnHand = Me.ListViewItemList.ListItems(i).SubItems(3)
   vQty = Me.ListViewItemList.ListItems(i).SubItems(4)
   vCheckQty = Me.ListViewItemList.ListItems(i).SubItems(5)
   vCountQty = Me.ListViewItemList.ListItems(i).SubItems(6)
   vVND = Me.ListViewItemList.ListItems(i).SubItems(10)
   vSHW = Me.ListViewItemList.ListItems(i).SubItems(11)
   vUnitCode = Me.ListViewItemList.ListItems(i).SubItems(7)
   vLineNumber = i - 1
   
   vQuery = "exec dbo.USP_MB_InsertDataShelfStockCountDetails '" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "'," & vOnHand & "," & vQty & "," & vVND & "," & vSHW & "," & vCheckQty & "," & vCountQty & ",'" & vUnitCode & "'," & vLineNumber & " "
   gConnection.Execute vQuery
   Next i
   
   vQuery = "commit tran"
   gConnection.Execute vQuery
   
   Me.ListViewItemList.ListItems.Clear
   Me.LBLRefNo.Caption = ""
   Me.CMBWHCode.Text = "014"
   MsgBox "บันทึกข้อมูลการตรวจนับสต๊อก คลัง " & vWHCode & " ที่เก็บ " & vShelfCode & " เรียบร้อยแล้ว", vbInformation, "Send Information Message"
      
ErrRollBackTran:
   If Err.Description <> "" Then
   vQuery = "rollback tran"
   gConnection.Execute vQuery
   MsgBox Err.Description & "   " & "ไม่สามารถบันทึกข้อมูลได้"
   End If

Else
   MsgBox "ไม่มีข้อมูลการตรวจนับในการบันทึก", vbCritical, "Send Error Message"
End If
End Sub

Private Sub CMDSave_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If Me.ListViewItemList.ListItems.Count > 0 Then
      If KeyCode = 116 Then
      Call CMDSave_Click
      Me.CMBShelf.SetFocus
   End If
End If
End Sub

Private Sub CMDSearchShelf_Click()
On Error Resume Next

PICSearchShelf.Visible = True
Me.TextSearch.SetFocus
End Sub

Private Sub CMDSearchShelf_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If Me.ListViewItemList.ListItems.Count > 0 Then
      If KeyCode = 116 Then
      Call CMDSave_Click
      Me.CMBShelf.SetFocus
   End If
End If
End Sub

Private Sub CMDSearchShelfDetails_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String
Dim vZoneCode As String
Dim vSearch As String
Dim vListShelf As ListItem

On Error Resume Next

vWHCode = Trim(CMBWHCode.Text)
vZoneCode = Trim(CMBZone.Text)

If TextSearch.Text <> "" Then
   vSearch = Trim(TextSearch.Text)
Else
   vSearch = ""
End If

ListViewShelf.ListItems.Clear
vQuery = "exec dbo.USP_MB_SearchShelfID'" & vWHCode & "' ,'" & vZoneCode & "','" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vListShelf = ListViewShelf.ListItems.Add(, , Trim(vRecordset.Fields("shelfid").Value))
        vListShelf.SubItems(1) = Trim(vRecordset.Fields("shelfname").Value)
        vRecordset.MoveNext
        Wend
        Me.ListViewShelf.SetFocus
End If
vRecordset.Close
End Sub

Public Sub GetWHCode()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

CMBWHCode.Clear
vQuery = "exec dbo.USP_MB_SearchWhCodeCode "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        CMBWHCode.AddItem Trim(vRecordset.Fields("whcode").Value)
        vRecordset.MoveNext
        Wend
End If
vRecordset.Close
Me.CMBWHCode.Text = "S01"
End Sub

Public Sub GetZoneCode()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

CMBZone.Clear
vQuery = "exec dbo.USP_MB_SearchZoneCode "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        CMBZone.AddItem Trim(vRecordset.Fields("zonecode").Value)
        vRecordset.MoveNext
        Wend
End If
vRecordset.Close
Me.CMBZone.Text = "AVL"
End Sub

Private Sub DTPDocDate_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i  As Integer
Dim vDocNo As String
Dim vDocDate As String
Dim vListItem As ListItem

On Error GoTo ErrDescription

If Me.CMBWHCode.Text <> "" Then
   vDocNo = Trim(Me.CMBWHCode.Text & "-" & Me.CMBShelf.Text)
   vDocDate = Me.DTPDocdate.Day & "/" & Me.DTPDocdate.Month & "/" & Me.DTPDocdate.Year
   vQuery = "exec dbo.USP_MB_SearchShelfStockCount '" & vDocNo & "','" & vDocDate & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   Me.LBLRefNo.Caption = vRecordset.Fields("docno").Value
   Call ClearScreen
   vRecordset.MoveFirst
   i = 1
   While Not vRecordset.EOF
   Set vListItem = Me.ListViewItemList.ListItems.Add(, , i)
   vListItem.SubItems(1) = vRecordset.Fields("itemcode").Value
   vListItem.SubItems(2) = vRecordset.Fields("itemname").Value
   vListItem.SubItems(3) = Format(vRecordset.Fields("onhand").Value, "####0.00")
   vListItem.SubItems(4) = Format(vRecordset.Fields("qty").Value, "####0.00")
   vListItem.SubItems(5) = Format(vRecordset.Fields("checkqty").Value, "####0.00")
   vListItem.SubItems(6) = Format(vRecordset.Fields("countqty").Value, "####0.00")
   vListItem.SubItems(7) = vRecordset.Fields("unitcode").Value
   vListItem.SubItems(8) = vRecordset.Fields("whcode").Value
   vListItem.SubItems(9) = vRecordset.Fields("shelfcode").Value
   vListItem.SubItems(10) = Format(vRecordset.Fields("vnd").Value, "####0.00")
   vListItem.SubItems(11) = Format(vRecordset.Fields("shw").Value, "####0.00")
   i = i + 1
   vRecordset.MoveNext
   Wend
   Else
      Call ClearScreen
      Me.ListViewItemList.ListItems.Clear
      Me.LBLRefNo.Caption = ""
   End If
   vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub DTPDocDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If Me.ListViewItemList.ListItems.Count > 0 Then
      If KeyCode = 116 Then
      Call CMDSave_Click
      Me.CMBShelf.SetFocus
   End If
End If
End Sub

Private Sub Form_Load()
Call GetWHCode
Call GetZoneCode
Me.DTPDocdate.Value = Now
End Sub

Private Sub ListViewItemList_DblClick()
Dim vIndex As Integer

On Error Resume Next

If Me.ListViewItemList.ListItems.Count > 0 Then
   vCheckSameValue = 1
   Me.PICKeyQTY.Visible = True
   
   vIndex = Me.ListViewItemList.SelectedItem.Index
   Me.LBLItemCode.Caption = Me.ListViewItemList.ListItems(vIndex).SubItems(1)
   Me.LBLItemName.Caption = Me.ListViewItemList.ListItems(vIndex).SubItems(2)
   Me.LBLUnitcode.Caption = Me.ListViewItemList.ListItems(vIndex).SubItems(7)
   Me.LBLOnHand.Caption = Me.ListViewItemList.ListItems(vIndex).SubItems(3)
   Me.TextQTY.Text = Format(Me.ListViewItemList.ListItems(vIndex).SubItems(4), "####0.00")
   Me.TextCheckQTY.Text = Format(Me.ListViewItemList.ListItems(vIndex).SubItems(5), "####0.00")
   Me.TextCountQTY.Text = Format(Me.ListViewItemList.ListItems(vIndex).SubItems(6), "####0.00")
   Me.TextVND.Text = Format(Me.ListViewItemList.ListItems(vIndex).SubItems(10), "####0.00")
   Me.TextSHW.Text = Format(Me.ListViewItemList.ListItems(vIndex).SubItems(11), "####0.00")
   Me.TextQTY.SetFocus
   
End If
End Sub

Private Sub ListViewItemList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer
Dim i As Integer

On Error Resume Next

If Me.ListViewItemList.ListItems.Count > 0 Then
   If KeyCode = 46 Then
   vIndex = Me.ListViewItemList.SelectedItem.Index
   Me.ListViewItemList.ListItems.Remove (vIndex)
   
   For i = 1 To Me.ListViewItemList.ListItems.Count
      Me.ListViewItemList.ListItems(i).Text = i
   Next i
   End If

   If KeyCode = 116 Then
      Call CMDSave_Click
      Me.CMBShelf.SetFocus
   End If
End If
End Sub

Private Sub ListViewQTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICKeyQTY.Visible = False
   vCheckSameValue = 0
   vSumQTY = 0
End If
End Sub

Private Sub ListViewShelf_DblClick()
On Error Resume Next

If Me.ListViewShelf.ListItems.Count > 0 Then
   Me.CMBShelf.Text = Me.ListViewShelf.SelectedItem.Text
   Me.PICSearchShelf.Visible = False
End If
End Sub

Private Sub TextCheckQTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICKeyQTY.Visible = False
   vCheckSameValue = 0
   vSumQTY = 0
End If
End Sub

Private Sub TextCheckQTY_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
   Call TextCheckQTY_LostFocus
End If
End Sub

Private Sub TextCheckQTY_LostFocus()
Dim vCheckQty As Double
Dim vCountDot As Integer

On Error Resume Next

If Me.TextCheckQTY.Text <> "" And Me.PICKeyQTY.Visible = True Then
   Call CheckNumber(Trim(Me.TextCheckQTY.Text))
   If vCheckValueNumber = False Then
      MsgBox "กรอกข้อมูลได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      Me.TextCheckQTY.SetFocus
   Else
      vCountDot = CheckDot(Me.TextCheckQTY.Text)
      If vCountDot <= 1 Then
         If Me.TextCheckQTY.Text <> "." Then
            vCheckQty = Me.TextCheckQTY.Text
            Me.TextCheckQTY.Text = Format(vCheckQty, "####0.00")
         Else
            MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
            Me.TextCheckQTY.SetFocus
         End If
      Else
         MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
         Me.TextCheckQTY.SetFocus
      End If
   End If
End If
End Sub

Private Sub TextCountQTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICKeyQTY.Visible = False
   vCheckSameValue = 0
   vSumQTY = 0
End If
End Sub

Private Sub TextCountQTY_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
   Call TextCountQTY_LostFocus
End If
End Sub

Private Sub TextCountQTY_LostFocus()
Dim vCountQty As Double
Dim vCountDot As Integer

On Error Resume Next

If Me.TextCountQTY.Text <> "" And Me.PICKeyQTY.Visible = True Then
   Call CheckNumber(Trim(Me.TextCountQTY.Text))
   If vCheckValueNumber = False Then
      MsgBox "กรอกข้อมูลได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      Me.TextCountQTY.SetFocus
   Else
      vCountDot = CheckDot(Me.TextCountQTY.Text)
      If vCountDot <= 1 Then
         If Me.TextCountQTY.Text <> "." Then
            vCountQty = Me.TextCountQTY.Text
            Me.TextCountQTY.Text = Format(vCountQty, "####0.00")
         Else
            MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
            Me.TextCountQTY.SetFocus
         End If
      Else
         MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
         Me.TextCountQTY.SetFocus
      End If
   End If
End If
End Sub

Private Sub TextItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If Me.ListViewItemList.ListItems.Count > 0 Then
      If KeyCode = 116 Then
      Call CMDSave_Click
      Me.CMBShelf.SetFocus
   End If
End If
End Sub

Private Sub TextItemCode_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vCheckItem As String
Dim vWHCode As String
Dim vListQTY As ListItem
         
On Error Resume Next

If KeyAscii = 13 Then
If Me.CMBWHCode.Text <> "" And Me.CMBZone.Text <> "" And Me.CMBShelf.Text <> "" Then
    vItemCode = Trim(TextItemCode.Text)
    vWHCode = Me.CMBWHCode.Text
    vQuery = "exec dbo.USP_ISP_SearchProduct1 '" & vItemCode & "','" & vWHCode & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       Me.PICKeyQTY.Visible = True
       LBLItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
       LBLItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
       LBLUnitcode.Caption = Trim(vRecordset.Fields("defstkunitcode").Value)
       Me.LBLOnHand.Caption = Format(vRecordset.Fields("onhand").Value, "####0.00")
       Me.TextQTY.Text = ""
       Me.TextCheckQTY.Text = ""
       Me.TextCountQTY.Text = ""
       Me.TextVND.Text = ""
       Me.TextSHW.Text = ""
      
       Me.ListViewQTY.ListItems.Clear
       vRecordset.MoveFirst
       While Not vRecordset.EOF
       Set vListQTY = Me.ListViewQTY.ListItems.Add(, , vRecordset.Fields("shelfcode").Value)
       vListQTY.SubItems(1) = Format(vRecordset.Fields("qty").Value, "####0.00")
       vListQTY.SubItems(2) = vRecordset.Fields("stkunitcode").Value
       vRecordset.MoveNext
       Wend
      
       Me.TextItemCode.Text = ""
       Me.TextQTY.SetFocus
   Else
      MsgBox "ไม่มีรหัสสินค้า " & vItemCode & " นี้ในระบบ ", vbCritical, "Send Error"
      LBLItemName.Caption = ""
      LBLUnitcode.Caption = ""
      Exit Sub
   End If
   vRecordset.Close
          
      If Me.ListViewItemList.ListItems.Count > 0 Then
         For i = 1 To Me.ListViewItemList.ListItems.Count
         vCheckItem = Me.ListViewItemList.ListItems(i).SubItems(1)
         If vItemCode = vCheckItem Then
            Me.TextQTY.Text = Me.ListViewItemList.ListItems(i).SubItems(4)
            Me.TextCheckQTY.Text = Me.ListViewItemList.ListItems(i).SubItems(5)
            Me.TextCountQTY.Text = Me.ListViewItemList.ListItems(i).SubItems(6)
            Me.TextVND.Text = Me.ListViewItemList.ListItems(i).SubItems(10)
            Me.TextSHW.Text = Me.ListViewItemList.ListItems(i).SubItems(11)
            vCheckSameValue = 1
            Exit Sub
         End If
         Next i
         
      Else
         vCheckSameValue = 0
      End If
Else
   MsgBox "กรุณาระบุคลังและที่เก็บให้เรียบร้อยก่อน กรอกรายการสินค้า", vbCritical, "Send Error Message"
   Me.CMBShelf.SetFocus
End If
End If
End Sub


Private Sub TextQTY_Change()
Dim vQty As Double
Dim vCountDot As Integer

On Error Resume Next

If Me.TextQTY.Text <> "" Then
   Call CheckNumber(Trim(Me.TextQTY.Text))
   If vCheckValueNumber = False Then
      MsgBox "กรอกข้อมูลได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      If Me.PICKeyQTY.Visible = True Then
         Me.TextQTY.SetFocus
      End If
      
   Else
      vCountDot = CheckDot(Me.TextQTY.Text)
      If vCountDot <= 1 Then
            If Me.TextQTY.Text <> "." Then
            vQty = Format(Me.TextQTY.Text, "####0.00")
            If vCheckSameValue = 0 Then
              Me.TextCheckQTY.Text = Format(vQty, "####0.00")
              Me.TextCountQTY.Text = Format(vQty, "####0.00")
            End If
            End If
      Else
         MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
         Me.TextQTY.SetFocus
      End If
   End If
Else
   Me.TextCheckQTY.Text = Format(0, "####0.00")
   Me.TextCountQTY.Text = Format(0, "####0.00")
End If
End Sub

Private Sub TextQTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICKeyQTY.Visible = False
   vCheckSameValue = 0
   vSumQTY = 0
End If
End Sub

Private Sub TextQty_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 And Me.TextQTY.Text <> "" Then
   Call TextQty_LostFocus
End If
End Sub

Private Sub TextQty_LostFocus()
Dim vQty As Double
Dim vCountDot As Integer

If Me.TextQTY.Text <> "" And Me.TextQTY.Text <> "0" And Me.PICKeyQTY.Visible = True Then
   Call CheckNumber(Trim(Me.TextQTY.Text))
   If vCheckValueNumber = False Then
      MsgBox "กรอกข้อมูลได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
   Else
      vCountDot = CheckDot(Me.TextQTY.Text)
      If vCountDot <= 1 Then
         If Me.TextQTY.Text <> "." Then
         vQty = Format(Me.TextQTY.Text, "####0.00")
         Me.TextQTY.Text = Format(vQty, "####0.00")
         If Me.TextVND.Text <> "" And Me.TextVND.Text <> "." Then
            Dim vnd As Double
            vnd = Me.TextVND.Text
         End If
         If Me.TextSHW.Text <> "" And Me.TextSHW.Text <> "." Then
            Dim shw As Double
            shw = Me.TextSHW.Text
         End If
         Dim QTY As Double
         QTY = Me.TextQTY.Text
         
         If (QTY - vnd) - shw <= 0 Then
            Me.TextVND.Text = 0
            Me.TextSHW.Text = 0
         End If
         Else
            MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
            Me.TextQTY.SetFocus
         End If
      Else
         MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
         Me.TextQTY.SetFocus
      End If
   End If
ElseIf Me.TextQTY.Text = "0" Or Me.TextQTY.Text = "" Then
   Me.TextVND.Text = Format(0, "####0.00")
   Me.TextSHW.Text = Format(0, "####0.00")
   Me.TextVND.SetFocus
End If

End Sub

Private Sub TextSearch_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String
Dim vZoneCode As String
Dim vSearch As String
Dim vListShelf As ListItem

On Error Resume Next

If KeyAscii = 13 Then
   vWHCode = Trim(CMBWHCode.Text)
   vZoneCode = Trim(CMBZone.Text)
   
   If TextSearch.Text <> "" Then
      vSearch = Trim(TextSearch.Text)
   Else
      vSearch = ""
   End If
   
   ListViewShelf.ListItems.Clear
   vQuery = "exec dbo.USP_MB_SearchShelfID'" & vWHCode & "','" & vZoneCode & "','" & vSearch & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
           vRecordset.MoveFirst
           While Not vRecordset.EOF
           Set vListShelf = ListViewShelf.ListItems.Add(, , Trim(vRecordset.Fields("shelfid").Value))
           vListShelf.SubItems(1) = Trim(vRecordset.Fields("shelfname").Value)
           vRecordset.MoveNext
           Wend
           Me.ListViewShelf.SetFocus
   End If
   vRecordset.Close
End If
End Sub

Private Sub TextSHW_Change()
vSumQTY = 0
End Sub

Private Sub TextSHW_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICKeyQTY.Visible = False
   vCheckSameValue = 0
   vSumQTY = 0
End If
End Sub

Private Sub TextSHW_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
   Call TextSHW_LostFocus
End If
End Sub

Private Sub TextSHW_LostFocus()
Dim vCount As Double
Dim vVND As Double
Dim vSHW As Double
Dim vCountDot As Integer
Dim vCountDot1 As Integer
Dim vSHWQty As Double

On Error Resume Next

If Me.TextQTY.Text <> "" And Me.TextQTY.Text <> "." And Me.TextVND.Text <> "." And Me.TextSHW.Text <> "" And Me.PICKeyQTY.Visible = True And vSumQTY = 0 Then
   Call CheckNumber(Trim(Me.TextSHW.Text))
   If vCheckValueNumber = False Then
      MsgBox "กรอกข้อมูลได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      Me.TextSHW.SetFocus
   Else
      vCountDot = CheckDot(Me.TextSHW.Text)
      vCountDot1 = CheckDot(Me.TextVND.Text)
      If vCountDot <= 1 And vCountDot1 <= 1 Then
         If Me.TextSHW.Text <> "." Then
         vCount = Me.TextQTY.Text
         vSHW = Me.TextSHW.Text
         vVND = Me.TextVND.Text
         If (vCount - vSHW) - vVND >= 0 Then
            vSHWQty = Me.TextSHW.Text
            Me.TextSHW.Text = Format(vSHWQty, "####0.00")
            vSumQTY = 0
         Else
            MsgBox "ยอดในชั้นเก็บ SHW ต้องไม่มากกว่ายอดทั้งหมดที่นับได้ในคลัง", vbCritical, "Send Error Message"
            vSumQTY = 1
            Me.TextSHW.SetFocus
         End If
      Else
         MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
         Me.TextSHW.SetFocus
         End If
      Else
         MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
         Me.TextSHW.SetFocus
      End If
   End If
ElseIf vSumQTY = 1 Then
MsgBox "ยอดในชั้นเก็บ SHW ต้องไม่มากกว่ายอดทั้งหมดที่นับได้ในคลัง", vbCritical, "Send Error Message"
Me.TextSHW.SetFocus
Me.TextSHW.Text = Format(0, "####0.00")
Me.TextVND.Text = Format(0, "####0.00")
End If
End Sub

Private Sub TextVND_Change()
vSumQTY = 0
End Sub

Private Sub TextVND_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICKeyQTY.Visible = False
   vCheckSameValue = 0
   vSumQTY = 0
End If
End Sub

Private Sub TextVND_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
   Call TextVND_LostFocus
End If
End Sub

Private Sub TextVND_LostFocus()
Dim vCount As Double
Dim vVND As Double
Dim vCountDot As Integer

On Error Resume Next

If Me.TextQTY.Text <> "" And Me.TextQTY.Text <> "." And Me.TextVND.Text <> "" And Me.PICKeyQTY.Visible = True And vSumQTY = 0 Then
   Call CheckNumber(Trim(Me.TextVND.Text))
   If vCheckValueNumber = False Then
      MsgBox "กรอกข้อมูลได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      Me.TextVND.SetFocus
   Else
      vCountDot = CheckDot(Me.TextVND.Text)
      If vCountDot <= 1 Then
         If Me.TextVND.Text <> "." Then
         vCount = Me.TextQTY.Text
         vVND = Me.TextVND.Text
         If vCount - vVND >= 0 Then
            Me.TextVND.Text = Format(vVND, "####0.00")
            vSumQTY = 0
         Else
            MsgBox "ยอดในชั้นเก็บ VND ต้องไม่มากกว่ายอดทั้งหมดที่นับได้ในคลัง", vbCritical, "Send Error Message"
            vSumQTY = 1
            Me.TextVND.SetFocus
         End If
         Else
            MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
            Me.TextVND.SetFocus
         End If
      Else
         MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
         Me.TextVND.SetFocus
      End If
   End If
ElseIf vSumQTY = 1 Then
   MsgBox "ยอดในชั้นเก็บ VND ต้องไม่มากกว่ายอดทั้งหมดที่นับได้ในคลัง", vbCritical, "Send Error Message"
   Me.TextVND.Text = Format(0, "####.00")
   Me.TextSHW.Text = Format(0, "####.00")
   Me.TextVND.SetFocus
End If
End Sub

