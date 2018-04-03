VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form107 
   Caption         =   "ตรวจสอบที่เก็บสินค้า"
   ClientHeight    =   8040
   ClientLeft      =   2265
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form107.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal107 
      Left            =   7965
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Crystal106 
      Left            =   7425
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Crystal105 
      Left            =   6885
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Crystal104 
      Left            =   6390
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Crystal103 
      Left            =   5895
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Crystal102 
      Left            =   5400
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   4905
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox PicCountSheet 
      BackColor       =   &H8000000C&
      Height          =   8070
      Left            =   0
      ScaleHeight     =   8010
      ScaleWidth      =   12015
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   12075
      Begin VB.OptionButton OPTLandScape 
         BackColor       =   &H8000000C&
         Caption         =   "กระดาษแนวนอน"
         Height          =   285
         Left            =   4995
         TabIndex        =   38
         Top             =   6390
         Value           =   -1  'True
         Width           =   1680
      End
      Begin VB.OptionButton OPTPortrait 
         BackColor       =   &H8000000C&
         Caption         =   "กระดาษแนวตั้ง"
         Height          =   285
         Left            =   4995
         TabIndex        =   37
         Top             =   6075
         Width           =   1500
      End
      Begin VB.CommandButton CMDPrintStoreQTY 
         Caption         =   "พิมพ์ยอดแยกสโตร์"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3645
         TabIndex        =   36
         Top             =   6120
         Width           =   1275
      End
      Begin MSComctlLib.ListView ListView104 
         Height          =   4245
         Left            =   6885
         TabIndex        =   27
         Top             =   1395
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   7488
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bay"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.CommandButton CMDExit 
         Caption         =   "ปิด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   8730
         TabIndex        =   26
         Top             =   6120
         Width           =   1275
      End
      Begin MSComctlLib.ListView ListView103 
         Height          =   4245
         Left            =   1800
         TabIndex        =   25
         Top             =   1395
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   7488
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ShelfCode"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.CommandButton CMDPrintCountSheet 
         Caption         =   "พิมพ์ CountSheet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1800
         TabIndex        =   24
         Top             =   6120
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "เลือก Bay ที่จะพิมพ์ :"
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
         Left            =   6885
         TabIndex        =   28
         Top             =   1035
         Width           =   2220
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "เลือก Row ที่จะพิมพ์ :"
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
         Left            =   1800
         TabIndex        =   23
         Top             =   1035
         Width           =   2355
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   0
         Picture         =   "Form107.frx":72FB
         Top             =   -45
         Width           =   2160
      End
   End
   Begin VB.PictureBox PICStoreItemSlotTag 
      BackColor       =   &H00808080&
      Height          =   8070
      Left            =   0
      ScaleHeight     =   8010
      ScaleWidth      =   11925
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   11985
      Begin VB.CommandButton CMDExitStoreSlotTag 
         Caption         =   "ออก"
         Height          =   465
         Left            =   10440
         TabIndex        =   44
         Top             =   6885
         Width           =   1095
      End
      Begin VB.CommandButton CMDPrintStoreSlotTag 
         Caption         =   "พิมพ์"
         Height          =   465
         Left            =   9135
         TabIndex        =   43
         Top             =   6885
         Width           =   1095
      End
      Begin VB.CheckBox CHKSelectAll 
         BackColor       =   &H00808080&
         Caption         =   "เลือกทั้งหมด"
         Height          =   285
         Left            =   390
         TabIndex        =   42
         Top             =   330
         Width           =   1680
      End
      Begin MSComctlLib.ListView ListViewStoreItem 
         Height          =   5955
         Left            =   405
         TabIndex        =   41
         Top             =   675
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   10504
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
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "คลัง"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ชั้นเก็บ"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "จำนวน"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "หน่วย"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ชื่ออังกฤษ"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "บาร์โค้ด1"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "บาร์โค้ด2"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin VB.PictureBox Pic101 
      Height          =   8070
      Left            =   0
      ScaleHeight     =   8010
      ScaleWidth      =   11970
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   12030
      Begin VB.CommandButton CMD105 
         Caption         =   "ออก"
         Height          =   420
         Left            =   9360
         TabIndex        =   16
         Top             =   5760
         Width           =   960
      End
      Begin VB.CommandButton CMD104 
         Caption         =   "เลือก"
         Height          =   420
         Left            =   7965
         TabIndex        =   15
         Top             =   5760
         Width           =   960
      End
      Begin VB.TextBox Text104 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2205
         TabIndex        =   14
         Top             =   630
         Width           =   2040
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   4290
         Left            =   1665
         TabIndex        =   12
         Top             =   1215
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   7567
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสที่เก็บ"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "คลัง"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "ค้นหา :"
         Height          =   285
         Left            =   1665
         TabIndex        =   13
         Top             =   630
         Width           =   825
      End
   End
   Begin VB.ListBox List102 
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
      Height          =   1470
      Left            =   3330
      TabIndex        =   49
      Top             =   1410
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.ListBox List101 
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
      Height          =   1470
      Left            =   3330
      TabIndex        =   10
      Top             =   660
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.TextBox TXTZone 
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
      Height          =   285
      Left            =   3330
      TabIndex        =   48
      Text            =   "AVL"
      Top             =   750
      Width           =   855
   End
   Begin VB.CommandButton CMDSearchZone 
      Height          =   285
      Left            =   4230
      Picture         =   "Form107.frx":9D85
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   750
      Width           =   345
   End
   Begin VB.CommandButton CMDPrintLabel 
      Caption         =   "พิมพ์ป้ายราคา"
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
      Left            =   6750
      TabIndex        =   45
      Top             =   7020
      Width           =   1320
   End
   Begin VB.CommandButton CMDStoreSlotTag 
      Caption         =   "SlotTag คลัง"
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
      Left            =   990
      TabIndex        =   39
      Top             =   7020
      Width           =   1320
   End
   Begin VB.CommandButton CMDInsertShelf 
      Caption         =   "เพิ่มที่เก็บสินค้า"
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
      TabIndex        =   30
      Top             =   6525
      Width           =   1320
   End
   Begin VB.CommandButton CMDSlotTag 
      Caption         =   "พิมพ์ StotTag"
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
      Left            =   990
      TabIndex        =   21
      Top             =   6525
      Width           =   1320
   End
   Begin VB.CommandButton CMDDelete 
      Caption         =   "ลบรายการ"
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
      Left            =   2430
      TabIndex        =   20
      Top             =   6525
      Width           =   1320
   End
   Begin VB.CheckBox Check101 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลือกทั้งหมด"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   990
      TabIndex        =   19
      Top             =   2025
      Width           =   1455
   End
   Begin VB.CommandButton CMDLock 
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
      Left            =   3960
      TabIndex        =   18
      Top             =   7560
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton CMDPrint 
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
      TabIndex        =   17
      Top             =   7560
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton CMD103 
      Caption         =   "ดูรายละเอียด"
      Height          =   420
      Left            =   5265
      TabIndex        =   9
      Top             =   1890
      Width           =   1140
   End
   Begin VB.CommandButton CMD102 
      Height          =   285
      Left            =   4995
      Picture         =   "Form107.frx":A152
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1125
      Width           =   330
   End
   Begin VB.CommandButton CMD101 
      Height          =   285
      Left            =   4230
      Picture         =   "Form107.frx":A51F
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   375
      Width           =   330
   End
   Begin VB.TextBox Text103 
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
      Height          =   300
      Left            =   3330
      TabIndex        =   6
      Top             =   1485
      Width           =   3075
   End
   Begin VB.TextBox Text102 
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
      Height          =   300
      Left            =   3330
      TabIndex        =   4
      Top             =   1110
      Width           =   1635
   End
   Begin VB.TextBox Text101 
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
      Height          =   300
      Left            =   3330
      TabIndex        =   1
      Text            =   "S02"
      Top             =   360
      Width           =   870
   End
   Begin MSComctlLib.ListView ListView102 
      Height          =   3660
      Left            =   990
      TabIndex        =   0
      Top             =   2700
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   6456
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
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสที่เก็บ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "OnHand"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หน่วยนับ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ชื่อสินค้า"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "คลัง"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PB101 
      Height          =   240
      Left            =   990
      TabIndex        =   29
      Top             =   2430
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CMDCountSheet 
      Caption         =   "พิมพ์ CountSheet"
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
      Left            =   3870
      TabIndex        =   31
      Top             =   6525
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   330
      Left            =   9675
      TabIndex        =   33
      Top             =   6525
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   39414
   End
   Begin VB.CommandButton CMDShelfPlan 
      Caption         =   "พิมพ์ทะเบียนคุม"
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
      Left            =   9675
      TabIndex        =   32
      Top             =   7020
      Width           =   1320
   End
   Begin VB.CommandButton CMDStoreQTY 
      Caption         =   "ยอดแยกสโตร์"
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
      Left            =   6750
      TabIndex        =   35
      Top             =   6525
      Width           =   1320
   End
   Begin VB.CommandButton CMDPrintInspectItem 
      Caption         =   "พิมพ์ผลการนับ"
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
      Left            =   8130
      TabIndex        =   50
      Top             =   7020
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "โซน :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2490
      TabIndex        =   46
      Top             =   750
      Width           =   765
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เริ่มจากวันที่ :"
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
      Left            =   8415
      TabIndex        =   34
      Top             =   6525
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "คลัง :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   2475
      TabIndex        =   5
      Top             =   375
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสที่เก็บ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   1575
      TabIndex        =   3
      Top             =   1125
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสสินค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   1530
      TabIndex        =   2
      Top             =   1485
      Width           =   1725
   End
End
Attribute VB_Name = "Form107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ImportBartendor As New FileSystemObject
Dim vCheckCountRow As Integer
Dim vCheckCountShelf As Integer
Dim vRowIndex As Integer
Dim vShelfIndex As Integer

Private Sub Check101_Click()
Dim i As Integer

On Error Resume Next

If Check101.Value = 1 Then
  For i = 1 To ListView102.ListItems.Count
    ListView102.ListItems.Item(i).Checked = True
  Next i
Else
  For i = 1 To ListView102.ListItems.Count
    ListView102.ListItems.Item(i).Checked = False
  Next i
  Check101.Value = 0
End If
End Sub

Private Sub CHKSelectAll_Click()
Dim i As Integer

If Me.ListViewStoreItem.ListItems.Count > 0 Then

   If Me.CHKSelectAll.Value = 1 Then
      For i = 1 To Me.ListViewStoreItem.ListItems.Count
         Me.ListViewStoreItem.ListItems(i).Checked = True
      Next i
   End If

   If Me.CHKSelectAll.Value = 0 Then
      For i = 1 To Me.ListViewStoreItem.ListItems.Count
         Me.ListViewStoreItem.ListItems(i).Checked = False
      Next i
   End If

End If
End Sub

Private Sub Cmd101_Click()
List101.Visible = True
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String
Dim vItemList As ListItem
Dim vZoneCode As String

On Error Resume Next

vWHCode = Trim(Text101.Text)
vZoneCode = Trim(Me.TXTZone.Text)

vQuery = "exec dbo.USP_MB_SearchShelfCode '" & vWHCode & "' ,'" & vZoneCode & "'"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        ListView101.ListItems.Clear
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vItemList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("shelfcode").Value))
        vItemList.SubItems(1) = Trim(vRecordset.Fields("whcode").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
    Pic101.Visible = True
    Me.Text104.SetFocus
End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vItemCode As String
Dim vItemList As ListItem
Dim i As Integer
Dim n As Integer

On Error Resume Next

If Me.Text101.Text <> "" Or Me.Text102.Text <> "" Or Me.Text103.Text <> "" Then
   ListView102.ListItems.Clear
   Me.PB101.Value = 0
   If Text103.Text <> "" And Me.Text101.Text = "" And Me.Text102.Text = "" Then
       vItemCode = Trim(Text103.Text)
       Text101.Text = ""
       Text102.Text = ""
       'vQuery = "exec dbo.USP_MB_SearchItemShelfCode '" & vItemCode & "' "
       vQuery = "exec dbo.USP_MB_SearchItemScanShelfCode '" & vItemCode & "' "
   ElseIf Me.Text101.Text <> "" And Me.Text102.Text <> "" And Text103.Text = "" Then
       Text103.Text = ""
       vWHCode = Text101.Text
       vShelfCode = Text102.Text
       'vQuery = "exec dbo.USP_MB_SearchItemRecProduct '" & vWHCode & "','" & vShelfCode & "' "
       vQuery = "exec dbo.USP_MB_SearchItemScanRecProduct '" & vWHCode & "','" & vShelfCode & "' "
   ElseIf Me.Text101.Text <> "" And (Me.Text102.Text <> "" Or Me.Text102.Text = "") And Text103.Text <> "" Then
       vWHCode = Text101.Text
       vShelfCode = Text102.Text
       vItemCode = Trim(Text103.Text)
       vQuery = "exec dbo.USP_MB_SearchItemCodeRecProduct '" & vWHCode & "','" & vShelfCode & "','" & vItemCode & "' "
   Else
       Exit Sub
   End If
       If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
           Me.PB101.Max = vRecordset.RecordCount
           vRecordset.MoveFirst
           i = 1
           n = 1
           While Not vRecordset.EOF
           Set vItemList = ListView102.ListItems.Add(, , i)
           vItemList.SubItems(1) = UCase(Trim(vRecordset.Fields("shelfcode").Value))
           vItemList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
           vItemList.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
           vItemList.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
           vItemList.SubItems(5) = Trim(vRecordset.Fields("itemname").Value)
           vItemList.SubItems(6) = Trim(vRecordset.Fields("whcode").Value)
           i = i + 1
           Me.PB101.Value = n
           n = n + 1
           vRecordset.MoveNext
           Wend
       Else
         MsgBox "ไม่มีรายการสินค้าที่ยิงที่เก็บสินค้า", vbInformation, "Send Error Message"
       End If
       vRecordset.Close
Else
   MsgBox "ต้องกรอกรายละเอียดที่ต้องการดูข้อมูลก่อน", vbCritical, "Send Error Message"
End If
End Sub

Private Sub CMD104_Click()
Dim i As Integer

On Error Resume Next

If ListView101.ListItems.Count <> 0 Then
    i = ListView101.SelectedItem.Index
    Text102.Text = Trim(ListView101.ListItems.Item(i).Text)
    Pic101.Visible = False
End If
End Sub

Private Sub CMD105_Click()
Pic101.Visible = False
End Sub


Private Sub CMDCountSheet_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String
Dim vItemList As ListItem
Dim vZoneCode As String

On Error Resume Next

If Me.Text101.Text <> "" Then
   PicCountSheet.Visible = True
   Me.ListView104.Visible = True
   Me.Label6.Visible = True
   vWHCode = Trim(Text101.Text)
   vZoneCode = Trim(TXTZone.Text)
   
   vQuery = "exec dbo.USP_MB_SearchShelfPrintCountSheet '" & vWHCode & "','" & vZoneCode & "','','' "
       If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
           ListView103.ListItems.Clear
           vRecordset.MoveFirst
           While Not vRecordset.EOF
           Set vItemList = ListView103.ListItems.Add(, , Trim(vRecordset.Fields("rowcode").Value))
           vRecordset.MoveNext
           Wend
       End If
       vRecordset.Close
Else
   MsgBox "กรุณาเลือกคลังที่ต้องการพิมพ์ CountSheet", vbCritical, "Send Error Message"
End If
End Sub

Private Sub CMDExitStoreSlotTag_Click()
PICStoreItemSlotTag.Visible = False
End Sub

Private Sub CMDPrintInspectItem_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHcodePrint As String
Dim vDatePrint As String
Dim vReportName As String

'On Error Resume Next


vQuery = "exec dbo.USP_NP_SelectReportName 476,'IS' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vReportName = vRecordset.Fields("reportname").Value
End If
vRecordset.Close

With Crystal107
.ReportFileName = vReportName & ".rpt"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End Sub

Private Sub CMDPrintLabel_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vWHCode As String
Dim vShelfCode As String

If Me.Text101.Text <> "" And Me.Text102.Text <> "" And Me.ListView102.ListItems.Count > 0 Then
   vQuery = "exec dbo.USP_NP_SelectReportName 416,'MB' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vReportName = vRecordset.Fields("reportname").Value
   End If
   vRecordset.Close
         
vWHCode = Me.Text101.Text
vShelfCode = Me.Text102.Text
         
   With Crystal105
   .ReportFileName = vReportName & ".rpt"
   .ParameterFields(1) = "@vWhcode;" & vWHCode & " ;true"
   .ParameterFields(2) = "@vShelfcode;" & vShelfCode & " ;true"
   .Destination = crptToWindow
   .WindowState = crptMaximized
   .Action = 1
   End With
End If
End Sub

Private Sub CMDPrintStoreQTY_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHcodePrint As String
Dim vRowCode As String
Dim vReportName As String
Dim i As Integer

On Error Resume Next

If Me.Text101.Text <> "" And Me.ListView103.ListItems.Count > 0 Then
  If Me.OPTPortrait.Value = True Then
     vQuery = "exec dbo.USP_NP_SelectReportName 377,'MB' "
  ElseIf Me.OPTLandScape.Value = True Then
     vQuery = "exec dbo.USP_NP_SelectReportName 378,'MB' "
  End If
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vReportName = vRecordset.Fields("reportname").Value
   End If
   vRecordset.Close
   
   For i = 1 To Me.ListView103.ListItems.Count
      If Me.ListView103.ListItems(i).Checked = True Then
         vRowCode = Me.ListView103.ListItems(i).Text
         With Crystal105
         .ReportFileName = vReportName & ".rpt"
         .ParameterFields(1) = "@Row;" & vRowCode & " ;true"
         .Destination = crptToWindow
         .WindowState = crptMaximized
         .Action = 1
         End With
      End If
   Next i
   For i = 1 To Me.ListView103.ListItems.Count
      Me.ListView103.ListItems(i).Checked = False
   Next i
End If
End Sub

Private Sub CMDPrintStoreSlotTag_Click()
Dim i As Integer
Dim vWHCode As String
Dim vShelfCode As String
Dim vItemCode As String
Dim vLine As Integer

On Error Resume Next

If Me.ListViewStoreItem.ListItems.Count > 0 Then
   For i = 1 To Me.ListViewStoreItem.ListItems.Count
   If Me.ListViewStoreItem.ListItems(i).Checked = True Then
      vWHCode = Me.ListViewStoreItem.ListItems(i).SubItems(3)
      vShelfCode = Me.ListViewStoreItem.ListItems(i).SubItems(4)
      vItemCode = Me.ListViewStoreItem.ListItems(i).SubItems(1)
      vLine = i
      Call vPrintCountSheetWHCode(vWHCode, vShelfCode, vItemCode, vLine)
   End If
   Next i
End If
End Sub

Private Sub CMDSearchZone_Click()
Me.List102.Visible = True
Me.List102.SetFocus
End Sub

Private Sub CMDShelfPlan_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHcodePrint As String
Dim vDatePrint As String
Dim vReportName As String

On Error Resume Next


vQuery = "exec dbo.USP_NP_SelectReportName 372,'ST' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vReportName = vRecordset.Fields("reportname").Value
End If
vRecordset.Close

vWHcodePrint = Trim(Text101.Text)
vDatePrint = Me.DTPDate.Day & "/" & Me.DTPDate.Month & "/" & Me.DTPDate.Year
With Crystal104
.ReportFileName = vReportName & ".rpt"
.ParameterFields(1) = "@vWHCode;" & vWHcodePrint & " ;true"
.ParameterFields(2) = "@vDate1;" & vDatePrint & " ;true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End Sub

Private Sub CMDSlotTag_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vListRow As ListItem
Dim vWHCode As String
Dim vShelfCode As String
Dim vItemCode As String
Dim vCountListItem As Integer

On Error Resume Next

If Me.Text101.Text <> "" And Me.Text102.Text <> "" And Me.ListView102.ListItems.Count > 0 Then
   vWHCode = Me.Text101.Text
   vShelfCode = Me.Text102.Text
   vItemCode = Me.Text103.Text
   
   For i = 1 To Me.ListView102.ListItems.Count
      If Me.ListView102.ListItems(i).Checked = True Then
         vCountListItem = vCountListItem + 1
      End If
   Next i
   
   If vCountListItem = 0 Then
      Call vPrintCountSheet(vWHCode, vShelfCode, vItemCode, 0)
   ElseIf vCountListItem = 1 And Me.ListView102.ListItems.Count = 1 Then
      Call vPrintCountSheet(vWHCode, vShelfCode, vItemCode, 0)
   ElseIf vCountListItem = Me.ListView102.ListItems.Count And vCountListItem > 1 Then
      Call vPrintCountSheet(vWHCode, vShelfCode, "", 0)
   Else
         For i = 1 To Me.ListView102.ListItems.Count
            If Me.ListView102.ListItems(i).Checked = True Then
               vItemCode = Me.ListView102.ListItems(i).SubItems(2)
               Call vPrintCountSheet(vWHCode, vShelfCode, vItemCode, i)
            End If
         Next i
   End If
   
   For i = 1 To Me.ListView102.ListItems.Count
      If Me.ListView102.ListItems(i).Checked = True Then
         Me.ListView102.ListItems(i).Checked = False
      End If
   Next i
   Me.Check101.Value = 0
   
   MsgBox "พิมพ์เสร็จเรียบร้อย กรุณาติด Slot Tag ให้ตรงกับสินค้าด้วย"
Else
   MsgBox "ต้องกรอกคลังและที่เก็บเป็นอย่างน้อย และต้องกดดูรายละเอียดข้อมูลก่อนพิมพ์", vbCritical, "Send Error Message"
End If


'Me.PicCountSheet.Visible = True
'vQuery = "exec dbo.USP_IC_RowOfToilet"
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '  Me.ListView103.ListItems.Clear
  ' vRecordset.MoveFirst
   'While Not vRecordset.EOF
   'Set vListRow = Me.ListView103.ListItems.Add(, , Trim(vRecordset.Fields("row").Value))
   'vRecordset.MoveNext
   'Wend
'End If
'vRecordset.Close
'Me.Opt101.Value = True
'Me.Opt102.Value = False



End Sub

Public Sub vPrintCountSheet(WHCode As String, ShelfCode As String, ItemCode As String, LineNumber As Integer)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim i As Integer
Dim prnPrinter As Printer
Dim Driver As String
Dim sMsg As String
Dim HWidth As Double
Dim HHeight As Double


'vPrinterName = Trim("\\diy01\TM-Mobile")
'vPrinterName = "TM-T88IIR"
'vPrinterName = Trim("\\x21\TM-T88II SlotTrack")
vPrinterName = Trim("\\hptc-5100421\SRP370B")
For Each printerObj In Printers
  If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    Set Printer = printerObj
    Set printerObj = Nothing
    Exit For
  End If
Next

If WHCode <> "" And ShelfCode <> "" And ItemCode = "" Then
   vQuery = "exec dbo.USP_MB_SearchItemScanRecProduct '" & WHCode & "','" & ShelfCode & "' "
ElseIf WHCode <> "" And ShelfCode <> "" And ItemCode <> "" Then
  vQuery = "exec dbo.USP_MB_SearchItemCodeRecProduct '" & WHCode & "','" & ShelfCode & "','" & ItemCode & "' "
End If

If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   i = 1
   While Not vRecordset.EOF
   
   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 20
   Printer.FontBold = True
   Printer.CurrentX = 1200
   Printer.CurrentY = 0
   Printer.Print "ต้นฉบับ Slot Tag"
   
   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 10
   Printer.FontBold = False
   Printer.CurrentX = 3500
   Printer.CurrentY = 350
   Printer.Print "ส่วน A "
   
   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 14
   Printer.CurrentX = 0
   Printer.CurrentY = 400
   Printer.Print "คลัง : " & Trim(vRecordset.Fields("whcode").Value) & "    " & "ที่เก็บ : " & Trim(vRecordset.Fields("shelfcode").Value)
    
   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 10
   Printer.CurrentX = 0
   Printer.CurrentY = 600
   Printer.Print "-------------------------------------------------------------------------------------------------"
   
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 14
    Printer.FontBold = True
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 100
    If LineNumber = 0 Then
       Printer.Print "ลำดับที่ " & "  " & i
    Else
       Printer.Print "ลำดับที่ " & "  " & LineNumber
    End If
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.FontBold = False
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value) & "           " & "หน่วยนับ :" & "  " & Trim(vRecordset.Fields("unitcode").Value)
     
If Trim(vRecordset.Fields("barcode1").Value) <> "" Then
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.FontBold = False
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    If Trim(vRecordset.Fields("barcode1").Value) <> Trim(vRecordset.Fields("barcode2").Value) Then
       Printer.Print "บาร์โค้ด :" & Trim(vRecordset.Fields("barcode1").Value) & "   ,  " & Trim(vRecordset.Fields("barcode2").Value)
    Else
       Printer.Print "บาร์โค้ด :" & Trim(vRecordset.Fields("barcode1").Value)
    End If
End If

    Printer.Font.Name = "3 of 9 Barcode"
    Printer.Font.Size = 20
    Printer.FontBold = False
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print "*" & Trim(vRecordset.Fields("itemcode").Value) & "*"

    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim(vRecordset.Fields("itemname1").Value)
    
    If Trim(vRecordset.Fields("itemname2").Value) <> "" Then
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim(vRecordset.Fields("itemname2").Value)
    End If


    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")

    If Trim(vRecordset.Fields("barcode1").Value) = "" And Trim(vRecordset.Fields("itemname2").Value) = "" Then
       Printer.CurrentY = 2100
       Printer.Line (0, 2100)-(0, 3100)
       Printer.Line (2100, 2100)-(2100, 3100)
       Printer.Line (4090, 2100)-(4090, 3100)
    ElseIf Trim(vRecordset.Fields("barcode1").Value) <> "" And Trim(vRecordset.Fields("itemname2").Value) = "" Then
       Printer.CurrentY = 2400
       Printer.Line (0, 2400)-(0, 3400)
       Printer.Line (2100, 2400)-(2100, 3400)
       Printer.Line (4090, 2400)-(4090, 3400)
    ElseIf Trim(vRecordset.Fields("barcode1").Value) = "" And Trim(vRecordset.Fields("itemname2").Value) <> "" Then
       Printer.CurrentY = 2400
       Printer.Line (0, 2400)-(0, 3400)
       Printer.Line (2100, 2400)-(2100, 3400)
       Printer.Line (4090, 2400)-(4090, 3400)
    ElseIf Trim(vRecordset.Fields("barcode1").Value) <> "" And Trim(vRecordset.Fields("itemname2").Value) <> "" Then
       Printer.CurrentY = 2700
       Printer.Line (0, 2700)-(0, 3700)
       Printer.Line (2100, 2700)-(2100, 3700)
       Printer.Line (4090, 2700)-(4090, 3700)
    End If
    
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 800
    Printer.Print "นับได้ :" & "                                          " & "ตรวจสอบ:"
        
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")
        
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 100
    Printer.Print "ผู้ตรวจนับ:" & "                                     " & "ผู้ตรวจสอบ:"
   
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")
    
    Printer.EndDoc

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 20
   Printer.FontBold = True
   Printer.CurrentX = 1200
   Printer.CurrentY = 0
   Printer.Print "ติด CountSheet"
   
   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 10
   Printer.FontBold = False
   Printer.CurrentX = 3500
   Printer.CurrentY = 350
   Printer.Print "ส่วน B "
   
   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 14
   Printer.CurrentX = 0
   Printer.CurrentY = 400
    If LineNumber = 0 Then
       Printer.Print "คลัง : " & Trim(vRecordset.Fields("whcode").Value) & "    " & "ที่เก็บ : " & Trim(vRecordset.Fields("shelfcode").Value) & "         " & "ลำดับที่ " & "  " & i
    Else
       Printer.Print "คลัง : " & Trim(vRecordset.Fields("whcode").Value) & "    " & "ที่เก็บ : " & Trim(vRecordset.Fields("shelfcode").Value) & "         " & "ลำดับที่ " & "  " & LineNumber
    End If

   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 10
   Printer.CurrentX = 0
   Printer.CurrentY = 600
   Printer.Print "-------------------------------------------------------------------------------------------------"
   
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print "รหัสสินค้า :" & "  " & Trim(vRecordset.Fields("itemcode").Value) & "      " & "หน่วยนับ :" & "  " & Trim(vRecordset.Fields("unitcode").Value)
             
    If Trim(vRecordset.Fields("barcode1").Value) <> "" Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 11
        Printer.FontBold = False
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY
        If Trim(vRecordset.Fields("barcode1").Value) <> Trim(vRecordset.Fields("barcode2").Value) Then
           Printer.Print "บาร์โค้ด :" & Trim(vRecordset.Fields("barcode1").Value) & "   ,  " & Trim(vRecordset.Fields("barcode2").Value)
        Else
           Printer.Print "บาร์โค้ด :" & Trim(vRecordset.Fields("barcode1").Value)
        End If
    End If
    
    Printer.Font.Name = "3 of 9 Barcode"
    Printer.Font.Size = 20
    Printer.FontBold = False
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print "*" & Trim(vRecordset.Fields("itemcode").Value) & "*"

    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim(vRecordset.Fields("itemname1").Value)
    
    If Trim(vRecordset.Fields("itemname2").Value) <> "" Then
       Printer.Font.Name = "AngsanaUPC"
       Printer.Font.Size = 11
       Printer.CurrentX = 0
       Printer.CurrentY = Printer.CurrentY
       Printer.Print Trim(vRecordset.Fields("itemname2").Value)
    End If
    
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")

    If Trim(vRecordset.Fields("barcode1").Value) = "" And Trim(vRecordset.Fields("itemname2").Value) = "" Then
       Printer.CurrentY = 1800
       Printer.Line (0, 1800)-(0, 2800)
       Printer.Line (2100, 1800)-(2100, 2800)
       Printer.Line (4090, 1800)-(4090, 2800)
    ElseIf Trim(vRecordset.Fields("barcode1").Value) <> "" And Trim(vRecordset.Fields("itemname2").Value) = "" Then
       Printer.CurrentY = 2100
       Printer.Line (0, 2100)-(0, 3100)
       Printer.Line (2100, 2100)-(2100, 3100)
       Printer.Line (4090, 2100)-(4090, 3100)
    ElseIf Trim(vRecordset.Fields("barcode1").Value) = "" And Trim(vRecordset.Fields("itemname2").Value) <> "" Then
       Printer.CurrentY = 2100
       Printer.Line (0, 2100)-(0, 3100)
       Printer.Line (2100, 2100)-(2100, 3100)
       Printer.Line (4090, 2100)-(4090, 3100)
    ElseIf Trim(vRecordset.Fields("barcode1").Value) <> "" And Trim(vRecordset.Fields("itemname2").Value) <> "" Then
       Printer.CurrentY = 2400
       Printer.Line (0, 2400)-(0, 3400)
       Printer.Line (2100, 2400)-(2100, 3400)
       Printer.Line (4090, 2400)-(4090, 3400)
    End If
    
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 800
    Printer.Print "นับได้ :" & "                                          " & "ตรวจสอบ:"
        
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")
        
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 100
    Printer.Print "ผู้ตรวจนับ:" & "                                     " & "ผู้ตรวจสอบ:"
    
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")
    
    Printer.EndDoc
    
   i = i + 1
   vRecordset.MoveNext
   Wend
End If
vRecordset.Close


End Sub

Private Sub TBPrintWrap(ByVal Text As String, ByVal LtMar As Long, ByVal RtMar As Long)
Dim i As Integer
Dim j As Integer
Dim currWord As String


Printer.CurrentX = LtMar
i = 1

Do Until i > Len(Text)
currWord = ""
Do Until i > Len(Text) Or Mid$(Text, i, 1) <= " "

currWord = currWord & Mid$(Text, i, 1)

i = i + 1
Loop

If (Printer.CurrentX + Printer.TextWidth(currWord)) > (Printer.ScaleWidth - RtMar + Printer.ScaleLeft) Then

Printer.Print
Printer.CurrentX = LtMar

End If

Printer.Print currWord;
Do Until i > Len(Text) Or Mid$(Text, i, 1) > " "

Select Case Mid$(Text, i, 1)

Case " "
Printer.Print " ";

Case Chr$(10) 'LF

Printer.Print

Printer.CurrentX = LtMar

Case Chr$(9) 'Tab

j = (Printer.CurrentX) / Printer.TextWidth("0")

j = j + (10 - (j Mod 10))

Printer.CurrentX = (j * Printer.TextWidth("0"))

Case Else

End Select

i = i + 1

Loop

Loop

End Sub

Private Sub CMDDelete_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vItemCode As String
Dim vWHCode As String
Dim vShelfCode As String

On Error Resume Next

If ListView102.ListItems.Count > 0 Then
  For i = 1 To ListView102.ListItems.Count
    If ListView102.ListItems.Item(i).Checked = True Then
      vItemCode = ListView102.ListItems.Item(i).SubItems(2)
      vWHCode = ListView102.ListItems.Item(i).SubItems(6)
      vShelfCode = ListView102.ListItems.Item(i).SubItems(1)
      
      'vQuery = "exec dbo.USP_DeleteRecProductShelfCode '" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "' "
      vQuery = "exec dbo.USP_MB_DeleteRecProductShelfCode '" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "' "
      gConnection.Execute vQuery
    
    End If
  Next i
  Call CMD103_Click
End If
End Sub

Private Sub CMDExit_Click()
Me.PicCountSheet.Visible = False
End Sub

Private Sub CMDInsertShelf_Click()
Form105.Show
Form105.SetFocus
End Sub

Private Sub CMDLock_Click()
Dim vQuery As String

On Error Resume Next

vQuery = "drop  table dbo.Report_Temp3 "
gConnection.Execute vQuery
CMDPrint.Enabled = True
CMDLock.Visible = False
ListView102.ListItems.Clear
End Sub

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim BartendorForm As String
Dim vItemCode As String
Dim vItemName As String
Dim vOnHand As Integer
Dim vUnitCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vCheckPrint As Integer
Dim n As Integer
Dim vReportName As String


On Error Resume Next

If ListView102.ListItems.Count > 0 Then

BartendorForm = "R5-9"

For n = 1 To ListView102.ListItems.Count
  If ListView102.ListItems.Item(n).Checked = True Then
    vCheckPrint = 1
    GoTo NextStep
  End If
Next n

NextStep:
If vCheckPrint = 1 Then
vQuery = "select * into dbo.Report_Temp3 From NP_LABEL_TEMP where UsedUser = 'Null' "
gConnection.Execute vQuery

For i = 1 To ListView102.ListItems.Count
If ListView102.ListItems.Item(i).Checked = True Then
  vItemCode = ListView102.ListItems.Item(i).SubItems(2)
  vItemName = ListView102.ListItems.Item(i).SubItems(5)
  vOnHand = 1
  vUnitCode = ListView102.ListItems.Item(i).SubItems(4)
  vWHCode = ListView102.ListItems.Item(i).SubItems(6)
  vShelfCode = ListView102.ListItems.Item(i).SubItems(1)
  
  vQuery = "Insert Into dbo.Report_Temp3(ItemCode, NAME1, QTY, UnitCode, UsedUser, WHCode, ShelfCode) " _
  & " values('" & Trim(vItemCode) & "','" & Trim(vItemName) & "', 1,'" & Trim(vUnitCode) & "','" & Trim(vUserID) & "','" & Trim(vWHCode) & "', '" & Trim(vShelfCode) & "') "
  gConnection.Execute vQuery
  
  vQuery = "exec dbo.USP_ISP_InsertPrintSlotTagBarCodeLog '" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vUserID & "' "
  gConnection.Execute vQuery
 End If
 Next i
 
vQuery = "select reportname from bcreportname where repid = 366 and reptype = 'ST' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

'ImportBartendor.CreateTextFile ("\\backup-server\bartendor\" & BartendorForm & ".txt")

vQuery = "drop  table dbo.Report_Temp3 "
gConnection.Execute vQuery


'CMDPrint.Enabled = False
'CMDLock.Visible = True
End If
End If
End Sub

Private Sub CMDPrintCountSheet_Click()
Dim i As Integer
Dim vSelectRow As String
Dim vSelectBay As String
Dim j As Integer
Dim vCheckPrintCount As Integer
Dim vWHCode As String
Dim vRow As String
Dim vShelfCode As String
Dim vZoneCode As String

On Error Resume Next

vWHCode = Me.Text101.Text
vZoneCode = Me.TXTZone.Text
   
If vCheckCountRow = 1 And vCheckCountShelf = 0 Then
   vRow = Me.ListView103.ListItems(vRowIndex).Text
   Call PrintCountSheet(vWHCode, vZoneCode, vRow, "")
ElseIf vCheckCountRow = 1 And vCheckCountShelf > 0 Then
    vRow = Me.ListView103.ListItems(vRowIndex).Text
    For i = 1 To Me.ListView104.ListItems.Count
       If Me.ListView104.ListItems(i).Checked = True Then
          vShelfCode = Me.ListView104.ListItems(i).Text
          Call PrintCountSheet(vWHCode, vZoneCode, vRow, vShelfCode)
       End If
    Next i
ElseIf vCheckCountRow > 1 And vCheckCountShelf = 0 Then
    For i = 1 To Me.ListView103.ListItems.Count
       If Me.ListView103.ListItems(i).Checked = True Then
          vRow = Me.ListView103.ListItems(i).Text
          Call PrintCountSheet(vWHCode, vZoneCode, vRow, "")
       End If
    Next i
End If


End Sub

Public Sub PrintCountSheet(vWHCode As String, vZoneCode As String, vRow As String, vShelfCode As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String

On Error Resume Next

vQuery = "exec dbo.USP_NP_SelectReportName 371,'ST' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vReportName = vRecordset.Fields("reportname").Value
End If
vRecordset.Close


With Crystal102
.ReportFileName = vReportName & ".rpt"
.ParameterFields(1) = "@vWHCode;" & vWHCode & " ;true"
.ParameterFields(2) = "@vZoneCode;" & vZoneCode & " ;true"
.ParameterFields(3) = "@vRow;" & vRow & " ;true"
.ParameterFields(4) = "@vShelfCode;" & vShelfCode & " ;true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

End Sub

Public Sub PrintCountSheet_HMX(vRow As String, vBay As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRowPrint As String
Dim vBayPrint As String
Dim vReportName As String
Dim vReportID As Integer
Dim vReportType As String

On Error Resume Next

vQuery = "exec dbo.USP_NP_SelectReportName 368,'ST' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vReportName = vRecordset.Fields("reportname").Value
End If
vRecordset.Close

vBayPrint = vBay
vRowPrint = vRow
With Crystal102
.ReportFileName = vReportName & ".rpt"
.ParameterFields(1) = "@Row;" & vRowPrint & " ;true"
.ParameterFields(2) = "@Bay;" & vBayPrint & " ;true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

End Sub

Private Sub CMDStoreQTY_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String
Dim vItemList As ListItem
Dim vZoneCode As String

On Error Resume Next

If Me.Text101.Text <> "" Then
   PicCountSheet.Visible = True
   Me.ListView104.Visible = False
   Me.Label6.Visible = False
   vWHCode = Trim(Text101.Text)
   vZoneCode = Trim(TXTZone.Text)
   
   vQuery = "exec dbo.USP_MB_SearchShelfPrintCountSheet '" & vWHCode & "','" & vZoneCode & "','','' "
       If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
           ListView103.ListItems.Clear
           vRecordset.MoveFirst
           While Not vRecordset.EOF
           Set vItemList = ListView103.ListItems.Add(, , Trim(vRecordset.Fields("rowcode").Value))
           vRecordset.MoveNext
           Wend
       End If
       vRecordset.Close
Else
   MsgBox "กรุณาเลือกคลังที่ต้องการพิมพ์ CountSheet", vbCritical, "Send Error Message"
End If
End Sub

Private Sub CMDStoreSlotTag_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vListItem As ListItem

On Error Resume Next

PICStoreItemSlotTag.Visible = True
Me.ListViewStoreItem.ListItems.Clear
i = 1
vQuery = "exec dbo.USP_IV_SearchStoreItem '097' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vListItem = Me.ListViewStoreItem.ListItems.Add(, , i)
        vListItem.SubItems(1) = (vRecordset.Fields("itemcode").Value)
        vListItem.SubItems(2) = (vRecordset.Fields("itemname1").Value)
        vListItem.SubItems(3) = (vRecordset.Fields("whcode").Value)
        vListItem.SubItems(4) = (vRecordset.Fields("shelfcode").Value)
        vListItem.SubItems(5) = (vRecordset.Fields("qty").Value)
        vListItem.SubItems(6) = (vRecordset.Fields("unitcode").Value)
        vListItem.SubItems(7) = (vRecordset.Fields("itemname2").Value)
        vListItem.SubItems(8) = (vRecordset.Fields("barcode1").Value)
        vListItem.SubItems(9) = (vRecordset.Fields("barcode2").Value)
        i = i + 1
        vRecordset.MoveNext
        Wend
End If
vRecordset.Close

Me.CHKSelectAll.Value = 1
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

List101.Clear
vQuery = "exec dbo.USP_MB_SearchWhCodeCode "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        List101.AddItem Trim(vRecordset.Fields("whcode").Value)
        vRecordset.MoveNext
        Wend
End If
vRecordset.Close

List102.Clear
vQuery = "exec dbo.USP_MB_SearchZoneCode "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        List102.AddItem Trim(vRecordset.Fields("zonecode").Value)
        vRecordset.MoveNext
        Wend
End If
vRecordset.Close
End Sub

Private Sub List101_Click()
Dim i As Integer
i = List101.ListIndex
Text101.Text = List101.List(i)
Text102.Text = ""
ListView102.ListItems.Clear
List101.Visible = False
End Sub

Private Sub List102_Click()
Dim i As Integer
i = List102.ListIndex
TXTZone.Text = List102.List(i)
Text102.Text = ""
ListView102.ListItems.Clear
List102.Visible = False
End Sub

Private Sub ListView101_DblClick()
Dim i As Integer

On Error Resume Next

If ListView101.ListItems.Count <> 0 Then
    i = ListView101.SelectedItem.Index
    Text102.Text = Trim(ListView101.ListItems.Item(i).Text)
    ListView102.ListItems.Clear
    Pic101.Visible = False
End If

End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
Dim i As Integer

If KeyAscii = 13 Then
    If ListView101.ListItems.Count <> 0 Then
        i = ListView101.SelectedItem.Index
        Text102.Text = Trim(ListView101.ListItems.Item(i).Text)
        Pic101.Visible = False
    End If
End If
End Sub


Private Sub ListView102_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vItemCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vIndex As Integer

On Error Resume Next

If KeyCode = 46 Then
   If ListView102.ListItems.Count > 0 Then
     vIndex = Me.ListView102.SelectedItem.Index
     vItemCode = ListView102.ListItems.Item(vIndex).SubItems(2)
     vWHCode = ListView102.ListItems.Item(vIndex).SubItems(6)
     vShelfCode = ListView102.ListItems.Item(vIndex).SubItems(1)
     
     'vQuery = "exec dbo.USP_DeleteRecProductShelfCode '" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "' "
     vQuery = "exec dbo.USP_MB_DeleteRecProductShelfCode '" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "' "
     gConnection.Execute vQuery
       
     Call CMD103_Click
   End If
End If
End Sub

Private Sub Opt101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vListRow As ListItem

On Error Resume Next

Me.ListView104.ListItems.Clear
vQuery = "exec dbo.USP_IC_RowOfToilet"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   Me.ListView103.ListItems.Clear
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   Set vListRow = Me.ListView103.ListItems.Add(, , Trim(vRecordset.Fields("row").Value))
   vRecordset.MoveNext
   Wend
End If
vRecordset.Close
End Sub

Private Sub Opt102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vListRow As ListItem

On Error Resume Next

vQuery = "exec dbo.USP_IC_RowOfToilet_HMX"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   Me.ListView103.ListItems.Clear
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   Set vListRow = Me.ListView103.ListItems.Add(, , Trim(vRecordset.Fields("rowcode").Value))
   vRecordset.MoveNext
   Wend
End If
vRecordset.Close
End Sub

Private Sub ListView103_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
Dim vCheckCount As Integer
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListRow As ListItem
Dim vRow As String
Dim vWHCode As String
Dim vZoneCode As String

On Error Resume Next

If Me.ListView103.ListItems.Count > 0 Then
   For i = 1 To Me.ListView103.ListItems.Count
   If Me.ListView103.ListItems.Item(i).Checked = True Then
      vCheckCount = vCheckCount + 1
      vCheckCountRow = vCheckCount
      If vCheckCount = 1 Then
         vRowIndex = i
      End If
   End If
   Next i
   
      If vCheckCount = 1 Then
         vWHCode = Me.Text101.Text
         vZoneCode = Me.TXTZone.Text
         
         vRow = Me.ListView103.ListItems(vRowIndex).Text
         vQuery = "exec dbo.USP_MB_SearchShelfPrintCountSheet '" & vWHCode & "','" & vZoneCode & "','" & vRow & "','' "
         If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            Me.ListView104.ListItems.Clear
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set vListRow = Me.ListView104.ListItems.Add(, , Trim(vRecordset.Fields("code").Value))
            vRecordset.MoveNext
            Wend
         End If
         vRecordset.Close
      Else
         Me.ListView104.ListItems.Clear
         vCheckCountShelf = 0
      End If
         
End If
      


End Sub

         
Private Sub ListView104_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
Dim vCheckCount As Integer

On Error Resume Next

If Me.ListView104.ListItems.Count > 0 Then
   For i = 1 To Me.ListView104.ListItems.Count
   If Me.ListView104.ListItems(i).Checked = True Then
      vCheckCount = vCheckCount + 1
      vCheckCountShelf = vCheckCount
      If vCheckCount = 1 Then
         vShelfIndex = i
         GoTo Line1
      End If
   Else
   vCheckCountShelf = 0
   End If
   Next i
Line1:
End If

End Sub

Private Sub Text101_Change()
Me.ListView102.ListItems.Clear
Me.PB101.Value = 0
End Sub

Private Sub Text102_Change()
Me.ListView102.ListItems.Clear
Me.PB101.Value = 0
End Sub

Private Sub Text102_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Text102.Text <> "" Then
      Call CMD103_Click
    End If
End If
End Sub

Private Sub Text103_Change()
Me.ListView102.ListItems.Clear
Me.PB101.Value = 0
End Sub

Private Sub Text103_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Text103.Text <> "" Then
      Call CMD103_Click
    End If
End If
End Sub

Private Sub Text104_Change()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vItemList As ListItem

On Error Resume Next

vWHCode = Trim(Text101.Text)
vShelfCode = Trim(Text104.Text)
ListView101.ListItems.Clear
vQuery = "exec dbo.USP_MB_SearchShelfCodeFilter '" & vWHCode & "','" & vShelfCode & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vItemList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("shelfcode").Value))
        vItemList.SubItems(1) = Trim(vRecordset.Fields("whcode").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
End Sub

Public Sub vPrintCountSheetWHCode(WHCode As String, ShelfCode As String, ItemCode As String, LineNumber As Integer)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim i As Integer
Dim prnPrinter As Printer
Dim Driver As String
Dim sMsg As String
Dim HWidth As Double
Dim HHeight As Double

Dim vItemName1 As String
Dim vItemName2 As String
Dim vQty As Double
Dim vUnitCode As String
Dim vBarCode1 As String
Dim vBarCode2 As String



vPrinterName = Trim("\\diy01\TM-Mobile")
'vPrinterName = "TM-T88IIR"
'vPrinterName = Trim("\\x21\TM-T88II SlotTrack")
For Each printerObj In Printers
  If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    Set Printer = printerObj
    Set printerObj = Nothing
    Exit For
  End If
Next


 vItemName1 = Me.ListViewStoreItem.ListItems(LineNumber).SubItems(2)
 vItemName2 = Me.ListViewStoreItem.ListItems(LineNumber).SubItems(7)
 vQty = Me.ListViewStoreItem.ListItems(LineNumber).SubItems(5)
 vUnitCode = Me.ListViewStoreItem.ListItems(LineNumber).SubItems(6)
 vBarCode1 = Me.ListViewStoreItem.ListItems(LineNumber).SubItems(8)
vBarCode2 = Me.ListViewStoreItem.ListItems(LineNumber).SubItems(9)


   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 20
   Printer.FontBold = True
   Printer.CurrentX = 1200
   Printer.CurrentY = 0
   Printer.Print "ต้นฉบับ Slot Tag"
   
   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 10
   Printer.FontBold = False
   Printer.CurrentX = 3500
   Printer.CurrentY = 350
   Printer.Print "ส่วน A "
   
   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 14
   Printer.CurrentX = 0
   Printer.CurrentY = 400
   Printer.Print "คลัง : " & WHCode & "    " & "ที่เก็บ : " & ShelfCode
    
   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 10
   Printer.CurrentX = 0
   Printer.CurrentY = 600
   Printer.Print "-------------------------------------------------------------------------------------------------"
   
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 14
    Printer.FontBold = True
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 100
    Printer.Print "OnHand " & "  " & vQty & "     " & vUnitCode

    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.FontBold = False
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print "รหัสสินค้า :" & ItemCode & "           " & "หน่วยนับ :" & "  " & vUnitCode
     
If vBarCode1 <> "" Then
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.FontBold = False
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    If vBarCode1 <> vBarCode2 Then
       Printer.Print "บาร์โค้ด :" & vBarCode1 & "   ,  " & vBarCode2
    Else
       Printer.Print "บาร์โค้ด :" & vBarCode1
    End If
End If

    Printer.Font.Name = "3 of 9 Barcode"
    Printer.Font.Size = 20
    Printer.FontBold = False
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print "*" & ItemCode & "*"

    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print vItemName1
    
    If vItemName2 <> "" Then
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print vItemName2
    End If


    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")

    If vBarCode1 = "" And vItemName2 = "" Then
       Printer.CurrentY = 2100
       Printer.Line (0, 2100)-(0, 3100)
       Printer.Line (2100, 2100)-(2100, 3100)
       Printer.Line (4090, 2100)-(4090, 3100)
    ElseIf vBarCode1 <> "" And vItemName2 = "" Then
       Printer.CurrentY = 2400
       Printer.Line (0, 2400)-(0, 3400)
       Printer.Line (2100, 2400)-(2100, 3400)
       Printer.Line (4090, 2400)-(4090, 3400)
    ElseIf vBarCode1 = "" And vItemName2 <> "" Then
       Printer.CurrentY = 2400
       Printer.Line (0, 2400)-(0, 3400)
       Printer.Line (2100, 2400)-(2100, 3400)
       Printer.Line (4090, 2400)-(4090, 3400)
    ElseIf vBarCode1 <> "" And vItemName2 <> "" Then
       Printer.CurrentY = 2700
       Printer.Line (0, 2700)-(0, 3700)
       Printer.Line (2100, 2700)-(2100, 3700)
       Printer.Line (4090, 2700)-(4090, 3700)
    End If
    
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 800
    Printer.Print "นับได้ :" & "                                          " & "ตรวจสอบ:"
        
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")
        
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 100
    Printer.Print "ผู้ตรวจนับ:" & "                                     " & "ผู้ตรวจสอบ:"
   
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")
    
    Printer.EndDoc

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 20
   Printer.FontBold = True
   Printer.CurrentX = 1200
   Printer.CurrentY = 0
   Printer.Print "ติด CountSheet"
   
   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 10
   Printer.FontBold = False
   Printer.CurrentX = 3500
   Printer.CurrentY = 350
   Printer.Print "ส่วน B "
   
   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 14
   Printer.CurrentX = 0
   Printer.CurrentY = 400
   Printer.Print "คลัง : " & WHCode & "    " & "ที่เก็บ : " & ShelfCode

   Printer.Font.Name = "AngsanaUPC"
   Printer.Font.Size = 10
   Printer.CurrentX = 0
   Printer.CurrentY = 600
   Printer.Print "-------------------------------------------------------------------------------------------------"
   
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print "รหัสสินค้า :" & "  " & ItemCode & "      " & "หน่วยนับ :" & "  " & vUnitCode
             
    If vBarCode1 <> "" Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 11
        Printer.FontBold = False
        Printer.CurrentX = 0
        Printer.CurrentY = Printer.CurrentY
        If vBarCode1 <> vBarCode2 Then
           Printer.Print "บาร์โค้ด :" & vBarCode1 & "   ,  " & vBarCode2
        Else
           Printer.Print "บาร์โค้ด :" & vBarCode1
        End If
    End If
    
    Printer.Font.Name = "3 of 9 Barcode"
    Printer.Font.Size = 20
    Printer.FontBold = False
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print "*" & ItemCode & "*"

    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print vItemName1
    
    If vItemName2 <> "" Then
       Printer.Font.Name = "AngsanaUPC"
       Printer.Font.Size = 11
       Printer.CurrentX = 0
       Printer.CurrentY = Printer.CurrentY
       Printer.Print vItemName2
    End If
    
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")

    If vBarCode1 = "" And vItemName2 = "" Then
       Printer.CurrentY = 1800
       Printer.Line (0, 1800)-(0, 2800)
       Printer.Line (2100, 1800)-(2100, 2800)
       Printer.Line (4090, 1800)-(4090, 2800)
    ElseIf vBarCode1 <> "" And vItemName2 = "" Then
       Printer.CurrentY = 2100
       Printer.Line (0, 2100)-(0, 3100)
       Printer.Line (2100, 2100)-(2100, 3100)
       Printer.Line (4090, 2100)-(4090, 3100)
    ElseIf vBarCode1 = "" And vItemName2 <> "" Then
       Printer.CurrentY = 2100
       Printer.Line (0, 2100)-(0, 3100)
       Printer.Line (2100, 2100)-(2100, 3100)
       Printer.Line (4090, 2100)-(4090, 3100)
    ElseIf vBarCode1 <> "" And vItemName2 <> "" Then
       Printer.CurrentY = 2400
       Printer.Line (0, 2400)-(0, 3400)
       Printer.Line (2100, 2400)-(2100, 3400)
       Printer.Line (4090, 2400)-(4090, 3400)
    End If
    
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 800
    Printer.Print "นับได้ :" & "                                          " & "ตรวจสอบ:"
        
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")
        
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 100
    Printer.Print "ผู้ตรวจนับ:" & "                                     " & "ผู้ตรวจสอบ:"
    
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 200
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")
    
    Printer.EndDoc
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Sub
