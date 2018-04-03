VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form102 
   Caption         =   "เช็คบาร์โค้ดตรวจสอบสต็อก"
   ClientHeight    =   7995
   ClientLeft      =   2100
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form102.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   11850
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   7245
      Top             =   7335
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
   Begin VB.PictureBox Pic101 
      BackColor       =   &H8000000C&
      Height          =   8070
      Left            =   0
      ScaleHeight     =   8010
      ScaleWidth      =   11835
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin VB.CommandButton CMDInsert 
         Caption         =   "เพิ่มรายใหม่"
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
         Left            =   4050
         TabIndex        =   29
         Top             =   4545
         Width           =   1230
      End
      Begin VB.TextBox LBLShelfCode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5265
         TabIndex        =   23
         Top             =   1890
         Width           =   1635
      End
      Begin VB.TextBox LBLWHCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   2430
         TabIndex        =   22
         Top             =   1890
         Width           =   1635
      End
      Begin VB.TextBox TextDescription 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2430
         TabIndex        =   27
         Top             =   3960
         Width           =   6405
      End
      Begin VB.TextBox TextShelfCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   2430
         TabIndex        =   24
         Top             =   2295
         Width           =   1635
      End
      Begin VB.TextBox TextCount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   600
         Left            =   2430
         TabIndex        =   26
         Top             =   3240
         Width           =   2850
      End
      Begin VB.CommandButton CMDCancel 
         Caption         =   "ยกเลิก"
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
         Left            =   5670
         TabIndex        =   30
         Top             =   4545
         Width           =   1230
      End
      Begin VB.CommandButton CMDOK 
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
         Height          =   465
         Left            =   2430
         TabIndex        =   28
         Top             =   4545
         Width           =   1230
      End
      Begin VB.Label LBLIndex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5265
         TabIndex        =   51
         Top             =   1080
         Width           =   510
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   -15
         Picture         =   "Form102.frx":72FB
         Top             =   0
         Width           =   2160
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
         Caption         =   "เหตุผลอื่น ๆ :"
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
         Left            =   990
         TabIndex        =   44
         Top             =   3960
         Width           =   1365
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
         Caption         =   "ที่เก็บสินค้า :"
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
         Left            =   1125
         TabIndex        =   43
         Top             =   2295
         Width           =   1230
      End
      Begin VB.Label LBLUnitCode1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   5355
         TabIndex        =   32
         Top             =   2700
         Width           =   1365
      End
      Begin VB.Label LBLUnitCode2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   5355
         TabIndex        =   31
         Top             =   3240
         Width           =   1365
      End
      Begin VB.Label LBLOnHand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2430
         TabIndex        =   25
         Top             =   2700
         Width           =   2850
      End
      Begin VB.Label LBLItemName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2430
         TabIndex        =   21
         Top             =   1485
         Width           =   6405
      End
      Begin VB.Label LBLItemCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2430
         TabIndex        =   20
         Top             =   1080
         Width           =   2625
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
         Caption         =   "นับได้ :"
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
         Left            =   720
         TabIndex        =   19
         Top             =   3240
         Width           =   1635
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
         Caption         =   "OnHand :"
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
         Left            =   540
         TabIndex        =   18
         Top             =   2700
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
         Caption         =   "ชั้นเก็บ :"
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
         Left            =   4230
         TabIndex        =   17
         Top             =   1890
         Width           =   960
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
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
         Height          =   285
         Left            =   1305
         TabIndex        =   16
         Top             =   1890
         Width           =   1050
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
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
         Height          =   285
         Left            =   1305
         TabIndex        =   15
         Top             =   1485
         Width           =   1050
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
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
         Height          =   330
         Left            =   1260
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.PictureBox PICKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4500
      ScaleHeight     =   2145
      ScaleWidth      =   7905
      TabIndex        =   36
      Top             =   45
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton CMDClosePic 
         Caption         =   "ปิด"
         Height          =   375
         Left            =   7335
         TabIndex        =   41
         Top             =   45
         Width           =   510
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "F1 : ให้ Focus อยู่ที่ตารางรายการที่ตรวจนับ (ตารางที่1)"
         Height          =   285
         Left            =   225
         TabIndex        =   37
         Top             =   135
         Width           =   4200
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 : ให้ Focus อยู่ที่ตารางรายละเอียดสินค้า (ตารางที่3)"
         Height          =   330
         Left            =   225
         TabIndex        =   42
         Top             =   810
         Width           =   4020
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 : คือการบันทึกเอกสารการตรวจนับ"
         Height          =   420
         Left            =   225
         TabIndex        =   40
         Top             =   1485
         Width           =   3435
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 : เป็นการนำสินค้าที่อยู่ในตารางกรอกจำนวนที่นับได้ ลงตะกร้า"
         Height          =   420
         Left            =   225
         TabIndex        =   39
         Top             =   1170
         Width           =   4740
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 : ให้ Focus อยู่ที่ตารางตะกร้าที่เก็บรายการสินค้าที่ตรวจนับแล้ว (ตารางที่2)"
         Height          =   375
         Left            =   225
         TabIndex        =   38
         Top             =   495
         Width           =   5505
      End
   End
   Begin VB.CommandButton CMDHandHeld 
      Caption         =   "ดึงข้อมูลจากเครื่อง HandHeld"
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
      Left            =   4500
      TabIndex        =   52
      Top             =   45
      Width           =   4200
   End
   Begin VB.ComboBox CMBZone 
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
      Left            =   7335
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   540
      Width           =   1365
   End
   Begin VB.OptionButton OPTZone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "เฉพาะโซน :"
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
      Left            =   6030
      TabIndex        =   49
      Top             =   540
      Width           =   1275
   End
   Begin VB.OptionButton OPTAll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ทุกโซน"
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
      Left            =   4500
      TabIndex        =   48
      Top             =   540
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.CommandButton CMDKey 
      Caption         =   "Key Use"
      Height          =   330
      Left            =   10890
      TabIndex        =   35
      Top             =   45
      Width           =   825
   End
   Begin VB.CommandButton CMDBasket 
      Caption         =   "ลงตาราง"
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
      Left            =   540
      TabIndex        =   12
      Top             =   3825
      Width           =   1050
   End
   Begin VB.ComboBox Cmb102 
      Height          =   315
      Left            =   8460
      TabIndex        =   10
      Top             =   6435
      Width           =   2790
   End
   Begin VB.CommandButton Cmd105 
      Caption         =   "บันทึกข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   540
      TabIndex        =   4
      Top             =   6435
      Width           =   1500
   End
   Begin VB.CommandButton Cmd104 
      Caption         =   "ลบข้อมูลใน Grid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5130
      TabIndex        =   3
      Top             =   6435
      Width           =   1500
   End
   Begin MSComctlLib.ListView ListView103 
      Height          =   1920
      Left            =   540
      TabIndex        =   6
      Top             =   1845
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   3387
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "คลัง"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชั้นเก็บ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "OnHand"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "นับได้"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หน่วย"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ชื่อสินค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "บาร์โค้ด"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ที่เก็บ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "รหัสเหตุผล"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "เหตุผล"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "หมายเหตุเพิ่ม"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gen"
      Height          =   390
      Left            =   11385
      TabIndex        =   9
      Top             =   5670
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Cmd103 
      Caption         =   "พิมพ์รายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   9900
      TabIndex        =   5
      Top             =   6840
      Width           =   1365
   End
   Begin VB.CommandButton Cmd102 
      Caption         =   "ลบข้อมูล"
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
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   6435
      Width           =   1500
   End
   Begin VB.CommandButton Cmd101 
      Caption         =   "แก้ไขข้อมูล"
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
      Height          =   615
      Left            =   2070
      TabIndex        =   1
      Top             =   6435
      Width           =   1500
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2070
      Left            =   540
      TabIndex        =   7
      Top             =   4275
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   3651
      SortKey         =   8
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "บาร์โค้ด"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "OnHand"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "จำนวนยิงได้"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "ผลต่างการนับ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "จำนวนหน้าShelf"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ชื่อสินค้า"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "หน่วยนับ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ShelfCode"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "DateScan"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "WHCode"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Shelf"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "เหตุผลการตรวจนับ"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1305
      MaxLength       =   20
      TabIndex        =   0
      Top             =   945
      Width           =   2100
   End
   Begin VB.ComboBox CMBCuaseStock 
      Height          =   315
      Left            =   4500
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   1395
      Width           =   6765
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เหตุผลการตรวจนับสินค้า :"
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
      Left            =   2295
      TabIndex        =   46
      Top             =   1395
      Width           =   2130
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "*** ภายใน 1 เลขที่เอกสาร สามารถตรวจสอบสต๊อกสินค้าได้ 500 รายการเท่านั้น"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   540
      TabIndex        =   45
      Top             =   7470
      Width           =   7350
   End
   Begin VB.Label LBLItemDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4500
      TabIndex        =   34
      Top             =   945
      Width           =   6765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   3510
      TabIndex        =   33
      Top             =   945
      Width           =   915
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบตรวจนับ :"
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
      Height          =   315
      Left            =   6795
      TabIndex        =   11
      Top             =   6480
      Width           =   1530
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "บาร์โค้ด :"
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
      Height          =   315
      Left            =   495
      TabIndex        =   8
      Top             =   945
      Width           =   720
   End
End
Attribute VB_Name = "Form102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vError As String
Dim vCheckBarCode As Integer
Dim vCheckStock As Integer
Dim vCheck As Integer
Dim vNewDocNo As String
Dim vGenDocNo As String
Dim vChkBar As Integer
Dim vCountItemCode As Integer
Dim vItemExist As Integer
Dim vIndex As Integer
Dim vQuery As String

Private Sub CMBCuaseStock_Change()
If Me.ListView103.ListItems.Count > 0 Then
   Me.ListView103.SetFocus
End If
End Sub

Private Sub CMBCuaseStock_Click()
If Me.ListView103.ListItems.Count > 0 Then
   Me.ListView103.SetFocus
End If
End Sub

Private Sub CMBZone_Click()
If Me.CMBZone.Text <> "" Then
    Me.OPTZone.Value = True
    Me.Text101.SetFocus
End If
End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType, vReportName, vDocNo As String

On Error GoTo ErrDescription

vRepID = 213
vRepType = "IV"
vDocNo = Trim(Cmb102.Text)
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


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD104_Click()
Dim vAnswer As Integer

vAnswer = MsgBox("คุณต้องการลบรายการตารางข้างล่างใช่หรือไม่ ?", vbYesNo, "Send Question Message")

If vAnswer = 6 Then
    vQuery = "exec dbo.USP_NP_DeleteInspectNoLog '" & vUserID & "','','','' "
    gConnection.Execute vQuery
    ListView101.ListItems.Clear
    Text101.SetFocus
End If
End Sub

Private Sub CMD105_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vDocDate As Date
Dim vItem, vUnit, vWH, vShelf, vItemName As String
Dim vQty, vDiff, vInspectQTY As Currency
Dim vLineNumber As Integer
Dim vShelfStock As String
Dim vReasonCode As String

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then

        vQuery = "begin tran"
        gConnection.Execute vQuery
   
        Call GetDocNo
        
        For i = 1 To ListView101.ListItems.Count
        vLineNumber = i - 1
        vShelfStock = Trim(ListView101.ListItems(i).SubItems(10))
        vItem = Trim(ListView101.ListItems(i).Text)
        vItemName = Trim(ListView101.ListItems(i).ListSubItems(5))
        vWH = Trim(ListView101.ListItems(i).SubItems(9))
        vShelf = Trim(ListView101.ListItems(i).ListSubItems(7))
        vUnit = Trim(ListView101.ListItems(i).ListSubItems(6))
        vQty = Format(ListView101.ListItems(i).ListSubItems(2), "####0.00")
        vDiff = Format(ListView101.ListItems(i).ListSubItems(3), "####0.00")
        vInspectQTY = Format(ListView101.ListItems(i).ListSubItems(1), "####0.00")
        vReasonCode = ListView101.ListItems(i).ListSubItems(11)

        vQuery = "exec dbo.USP_NP_InsertInspectLog '" & vGenDocNo & "','" & vItem & "','" & vItemName & "','" & vWH & "','" & vShelf & "'," & vQty & ",'" & vUnit & "','" & vUserID & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
        gConnection.Execute vQuery
        
        If Me.ListView101.ListItems.Count > 0 Then
        vQuery = "exec dbo.USP_NP_UpdateInspectNoLog '" & vGenDocNo & "','" & vUserID & "' "
        gConnection.Execute vQuery
        End If
        
        Next i
        
        vQuery = "commit tran"
        gConnection.Execute vQuery
        
        Call GetInspectNo
        ListView101.ListItems.Clear
        ListView103.ListItems.Clear
        Call PrintInspection
        'End If
        
Else
MsgBox "ตรวจสอบ สินค้าในตารางว่ามีไหม ตรวจสอบว่ายิงชั้นเก็บหรือยัง ตรวจสอบคลัง "
Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub CMDBasket_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListItem As ListItem
Dim vBarCode As String
Dim vShelfCode As String
Dim vCheckKeyQTY As Integer
Dim vWHCode As String
Dim vZoneCode As String
Dim vReasonID As String
Dim i As Integer
Dim n As Integer
Dim v As Integer
Dim x As Integer
Dim vQty As Double

Dim vDocDate As Date
Dim vItem, vUnit, vWH, vShelf, vItemName As String
Dim vDiff, vInspectQTY As Double
Dim vLineNumber As Integer
Dim vShelfStock As String
Dim vReasonCode As String
Dim vAnswer As Integer
Dim vAnswer1 As Integer
Dim vAnswer2 As Integer

Dim vGetOnHand As Double
Dim vGetInspectQty As Double
Dim vGetDiff As Double

Dim vOldInspectQty As Double
Dim vAddInspectQty As Double

On Error GoTo ErrDescription

  
If ListView103.ListItems.Count > 0 Then

For n = 1 To ListView103.ListItems.Count

 vGetOnHand = ListView103.ListItems.Item(n).SubItems(2)
 vGetInspectQty = ListView103.ListItems.Item(n).SubItems(3)

  If ListView103.ListItems.Item(n).SubItems(3) <> "" Then
    vCheckKeyQTY = 1
  Else
    vCheckKeyQTY = 0
    MsgBox "โปรด กรอกจำนวนที่นับได้ทุกชั้นเก็บ กรณีไม่มีก็ให้ใส่ 0 ", vbCritical, "Send Error "
    ListView103.SetFocus
    Exit Sub
  End If
Next n

If ListView101.ListItems.Count > 0 Then
For v = 1 To ListView101.ListItems.Count
  For x = 1 To ListView103.ListItems.Count
  If ListView101.ListItems.Item(v).Text = ListView103.ListItems(x).SubItems(5) And ListView101.ListItems.Item(v).SubItems(6) = ListView103.ListItems(x).SubItems(4) And ListView101.ListItems.Item(v).SubItems(9) = ListView103.ListItems(x).Text And ListView101.ListItems.Item(v).SubItems(10) = ListView103.ListItems(x).SubItems(1) Then
    vAnswer = MsgBox("รหัสสินค้า " & ListView101.ListItems.Item(v).Text & " หน่วยนับ  " & ListView101.ListItems.Item(v).SubItems(6) & " คลัง " & ListView103.ListItems(x).Text & " ชั้นเก็บ " & ListView103.ListItems(x).SubItems(1) & " นี้มีอยู่แล้วในรายการ ที่ " & v & " คุณต้องการแก้ไขจำนวนใช่หรือไม่", vbYesNo, "Send Error ")
    
    vOldInspectQty = Me.ListView101.ListItems(v).SubItems(2)
    
    If vAnswer = 6 Then
        vAnswer1 = MsgBox("เปลี่ยนจำนวนที่นับได้เป็นจำนวนที่กรอกล่าสุดใช่หรือไม่ ถ้าไม่จะเป็นการบวกกับจำนวนเดิมที่นับได้ก่อนหน้านี้", vbYesNo, "Send Question Message")
        
        If vAnswer1 = 6 Then
             vGetDiff = vGetInspectQty - vGetOnHand
             
             Me.ListView101.ListItems(v).SubItems(1) = Format(vGetOnHand, "##,##0.00")
             Me.ListView101.ListItems(v).SubItems(2) = Format(vGetInspectQty, "##,##0.00")
             Me.ListView101.ListItems(v).SubItems(3) = Format(vGetDiff, "##,##0.00")
                          
            vLineNumber = v
            vShelfStock = Trim(ListView101.ListItems(v).SubItems(10))
            vItem = Trim(ListView101.ListItems(v).Text)
            vItemName = Trim(ListView101.ListItems(v).ListSubItems(5))
            vWH = Trim(ListView101.ListItems(v).SubItems(9))
            vShelf = Trim(ListView101.ListItems(v).ListSubItems(7))
            vUnit = Trim(ListView101.ListItems(v).ListSubItems(6))
            vReasonCode = ListView101.ListItems(v).ListSubItems(11)
        
            vQuery = "exec dbo.USP_NP_InsertInspectLog '','" & vItem & "','" & vItemName & "','" & vWH & "','" & vShelf & "'," & vGetInspectQty & ",'" & vUnit & "','" & vUserID & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
            gConnection.Execute vQuery
             
            ListView103.ListItems.Clear
            LBLItemDescription.Caption = ""
            Me.CMBCuaseStock.Clear
            Text101.Text = ""
            Text101.SetFocus
             Exit Sub
        Else
            vAddInspectQty = vGetInspectQty + vOldInspectQty
             vGetDiff = vAddInspectQty - vGetOnHand
             
             Me.ListView101.ListItems(v).SubItems(1) = Format(vGetOnHand, "##,##0.00")
             Me.ListView101.ListItems(v).SubItems(2) = Format(vAddInspectQty, "##,##0.00")
             Me.ListView101.ListItems(v).SubItems(3) = Format(vGetDiff, "##,##0.00")
                          
            vLineNumber = v
            vShelfStock = Trim(ListView101.ListItems(v).SubItems(10))
            vItem = Trim(ListView101.ListItems(v).Text)
            vItemName = Trim(ListView101.ListItems(v).ListSubItems(5))
            vWH = Trim(ListView101.ListItems(v).SubItems(9))
            vShelf = Trim(ListView101.ListItems(v).ListSubItems(7))
            vUnit = Trim(ListView101.ListItems(v).ListSubItems(6))
            vReasonCode = ListView101.ListItems(v).ListSubItems(11)
                     
            vQuery = "exec dbo.USP_NP_InsertInspectLog '','" & vItem & "','" & vItemName & "','" & vWH & "','" & vShelf & "'," & vAddInspectQty & ",'" & vUnit & "','" & vUserID & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
            gConnection.Execute vQuery

            ListView103.ListItems.Clear
            LBLItemDescription.Caption = ""
            Me.CMBCuaseStock.Clear
            Text101.Text = ""
            Text101.SetFocus
            Exit Sub
        End If
    Else
    Me.Text101.SetFocus
    Exit Sub
    End If
    
  End If
  Next x
Next v
End If

  For i = 1 To ListView103.ListItems.Count
  vBarCode = Trim(ListView103.ListItems.Item(i).SubItems(5))
  vWHCode = Trim(ListView103.ListItems.Item(i).Text)
  vShelfCode = UCase(Trim(ListView103.ListItems.Item(i).SubItems(1)))
  If Trim(ListView103.ListItems.Item(i).SubItems(11)) <> "" Then
    vReasonID = Trim(ListView103.ListItems.Item(i).SubItems(9)) & "//" & Trim(ListView103.ListItems.Item(i).SubItems(11))
  Else
    vReasonID = Trim(ListView103.ListItems.Item(i).SubItems(9))
  End If
  
    vQuery = "exec dbo.USP_MB_ProgStockChecking  '" & vBarCode & "','" & vWHCode & "' ,'" & vShelfCode & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQty = Trim(vRecordset.Fields("qty").Value)
    Else
      vQty = 0
    End If
    vRecordset.Close

        Set ListItem = ListView101.ListItems.Add(, , Trim(vBarCode))
        ListItem.SubItems(1) = Format(vQty, "##,##0.00")
        ListItem.SubItems(2) = Format(Trim(ListView103.ListItems.Item(i).SubItems(3)), "##,##0.00")
        ListItem.SubItems(3) = Format(Trim(ListView103.ListItems.Item(i).SubItems(3)) - vQty, "##,##0.00")
        ListItem.SubItems(5) = Trim(ListView103.ListItems.Item(i).SubItems(6))
        ListItem.SubItems(6) = Trim(ListView103.ListItems.Item(i).SubItems(4))
        ListItem.SubItems(7) = Trim(ListView103.ListItems.Item(i).SubItems(8))
        ListItem.SubItems(8) = Now
        ListItem.SubItems(9) = vWHCode
        ListItem.SubItems(10) = vShelfCode
        ListItem.SubItems(11) = vReasonID
        
        vLineNumber = Me.ListView101.ListItems.Count - 1
        vShelfStock = Trim(ListView101.ListItems(i).SubItems(10))
        vItem = Trim(ListView101.ListItems(i).Text)
        vItemName = Trim(ListView101.ListItems(i).ListSubItems(5))
        vWH = Trim(ListView101.ListItems(i).SubItems(9))
        vShelf = Trim(ListView101.ListItems(i).ListSubItems(7))
        vUnit = Trim(ListView101.ListItems(i).ListSubItems(6))
        vQty = Format(ListView101.ListItems(i).ListSubItems(2), "####0.00")
        vDiff = Format(ListView101.ListItems(i).ListSubItems(3), "####0.00")
        vInspectQTY = Format(ListView101.ListItems(i).ListSubItems(1), "####0.00")
        vReasonCode = ListView101.ListItems(i).ListSubItems(11)
        
        vQuery = "exec dbo.USP_NP_InsertInspectLog '','" & vItem & "','" & vItemName & "','" & vWH & "','" & vShelf & "'," & vQty & ",'" & vUnit & "','" & vUserID & "','" & vShelfStock & "','" & vReasonCode & "'," & vLineNumber & "  "
        gConnection.Execute vQuery

  
  Next i
ListView103.ListItems.Clear
LBLItemDescription.Caption = ""
Me.CMBCuaseStock.Clear
Text101.Text = ""
Text101.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub CMDCancel_Click()
Pic101.Visible = False
End Sub

Private Sub CMDClosePic_Click()
PICKey.Visible = False
End Sub

Private Sub CMDHandHeld_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vItemCode As String
Dim vItemName As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vReasonCode As String
Dim i As Integer
Dim n As Integer
Dim vDocDate As Date
Dim vInspectQTY As Double
Dim vUnitCode As String
Dim vLineNumber As Integer

Dim vCheckItemCode As String
Dim vCheckWHCode As String
Dim vCheckShelfCode As String
Dim vCheckUnitCode As String
Dim vCheckQty As Double
Dim vNewInspectQty As Double

Dim vAnswer As Integer
Dim vAnswer1 As Integer

On Error GoTo ErrDescription

vQuery = "exec dbo.USP_NP_SearchStockCountHandHeld '" & vUserID & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    vItemCode = Trim(vRecordset.Fields("itemcode").Value)
    vItemName = Trim(vRecordset.Fields("itemname").Value)
    vWHCode = Trim(vRecordset.Fields("whcode").Value)
    vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
    vInspectQTY = Trim(vRecordset.Fields("countqty").Value)
    vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
    vReasonCode = Trim(vRecordset.Fields("reasoncode").Value)
    vLineNumber = i
    
    
    If Me.ListView101.ListItems.Count > 0 Then
    For n = 1 To Me.ListView101.ListItems.Count
    vCheckItemCode = Me.ListView101.ListItems(n).Text
    vCheckWHCode = Me.ListView101.ListItems(n).SubItems(9)
    vCheckShelfCode = Me.ListView101.ListItems(n).SubItems(10)
    vCheckUnitCode = Me.ListView101.ListItems(n).SubItems(6)
    vCheckQty = Me.ListView101.ListItems(n).SubItems(2)
    
    If vItemCode = vCheckItemCode And vWHCode = vCheckWHCode And vShelfCode = vCheckShelfCode And vUnitCode = vCheckUnitCode Then
        vAnswer = MsgBox("มีรายการสินค้า " & vItemCode & " " & vItemName & " นี้อยู่ในรายการที่ " & n & " ของตาราง คุณต้องการแก้ไขจำนวนที่นับได้หรือไม่ ?", vbYesNo, "Send Question Message")
    
        If vAnswer = 6 Then
            vAnswer1 = MsgBox("เลือก Yes เพื่อแก้ไขจำนวนที่นับได้ เลือก No เพื่อรวมจำนวนที่นับได้ใหม่กับจำนวนเก่า", vbYesNo, "Send Question Message")
            If vAnswer1 = 6 Then
                vQuery = "exec dbo.USP_NP_InsertInspectLog '','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "'," & vInspectQTY & ",'" & vUnitCode & "','" & vUserID & "','" & vShelfCode & "','" & vReasonCode & "'," & vLineNumber & "  "
                gConnection.Execute vQuery
                
                vQuery = "exec dbo.USP_NP_UpdateUsedStockCountHandHeld '" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "' "
                gConnection.Execute vQuery
                
                GoTo LineNext
                
            ElseIf vAnswer1 = 7 Then
                vNewInspectQty = vCheckQty + vInspectQTY
                vQuery = "exec dbo.USP_NP_InsertInspectLog '','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "'," & vNewInspectQty & ",'" & vUnitCode & "','" & vUserID & "','" & vShelfCode & "','" & vReasonCode & "'," & vLineNumber & "  "
                gConnection.Execute vQuery
                
                vQuery = "exec dbo.USP_NP_UpdateUsedStockCountHandHeld '" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "' "
                gConnection.Execute vQuery
                
                GoTo LineNext
            End If
        Else
                vQuery = "exec dbo.USP_NP_UpdateUsedStockCountHandHeld '" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "' "
                gConnection.Execute vQuery
                
                GoTo LineNext
        End If
    End If
    
    Next n
    End If
    
    
    vQuery = "exec dbo.USP_NP_InsertInspectLog '','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "'," & vInspectQTY & ",'" & vUnitCode & "','" & vUserID & "','" & vShelfCode & "','" & vReasonCode & "'," & vLineNumber & "  "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateUsedStockCountHandHeld '" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "' "
    gConnection.Execute vQuery
    
    Call SearchStkInspectLogs
    
LineNext:

    i = i + 1
    vRecordset.MoveNext
    Wend
Else
MsgBox "ไม่มีรายการสินค้าที่ตรวจนับจากเครื่อง HandHeld กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Me.Text101.SetFocus
End If
vRecordset.Close

Call SearchStkInspectLogs

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDInsert_Click()
Dim ListItem As ListItem
Dim i As Integer

On Error GoTo ErrDescription

If TextShelfCode.Text = "" Then
  MsgBox "กรุณากรอกชั้นเก็บที่สินค้าอยู่จริงด้วย", vbCritical, "Send Error"
End If

  If TextCount.Text <> "" Then
    For i = 1 To ListView103.ListItems.Count
      If LBLWHCode.Text = ListView103.ListItems(i).Text And LBLShelfCode.Text = ListView103.ListItems(i).SubItems(1) Then
        MsgBox "มีข้อมูลของคลัง " & LBLWHCode.Text & " ชั้นเก็บ " & LBLShelfCode.Text & " นี้อยู่แล้ว กรุณาตรวจสอบ", vbCritical, "Send Error"
        Exit Sub
      End If
    Next i
    

    Set ListItem = ListView103.ListItems.Add(, , Trim(LBLWHCode.Text))
    ListItem.SubItems(1) = Trim(LBLShelfCode.Text)
    ListItem.SubItems(2) = Format(0, "##,##0.00")
    ListItem.SubItems(3) = Format(Trim(TextCount.Text), "##,##0.00")
    ListItem.SubItems(4) = Trim(LBLUnitCode2.Caption)
    ListItem.SubItems(5) = Trim(LBLItemCode.Caption)
    ListItem.SubItems(6) = Trim(LBLItemName.Caption)
    ListItem.SubItems(7) = Trim(LBLItemCode.Caption)
    ListItem.SubItems(8) = Trim(TextShelfCode.Text)
    If CMBCuaseStock.Text <> "" Then
    ListItem.SubItems(9) = Left(Trim(CMBCuaseStock.Text), InStr(Trim(CMBCuaseStock.Text), "//") - 1)
    ListItem.SubItems(10) = Right(Trim(CMBCuaseStock.Text), Len(Trim(CMBCuaseStock.Text)) - InStr(Trim(CMBCuaseStock.Text), "//") - 1)
    End If
    ListItem.SubItems(11) = Trim(TextDescription.Text)
    
    Pic101.Visible = False
  Else
    MsgBox "กรุณากรอกจำนวนที่นับได้ด้วย", vbCritical, "Send Error"
  End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDKey_Click()
PICKey.Visible = True
End Sub

Private Sub CMDOK_Click()

On Error Resume Next

If vIndex <> 0 Then
  If TextCount.Text <> "" Then
      If TextShelfCode.Text = "" Then
         MsgBox "ยังไม่ได้ระบุที่เก็บที่อยู่จริงของสินค้าของสินค้า", vbCritical, "Send Error"
      End If
      
        Pic101.Visible = False
        ListView103.ListItems.Item(vIndex).SubItems(3) = Format(Trim(TextCount.Text), "##,##0.00")
        ListView103.ListItems.Item(vIndex).SubItems(8) = Trim(TextShelfCode.Text)
        If CMBCuaseStock.Text <> "" Then
        ListView103.ListItems.Item(vIndex).SubItems(9) = Left(Trim(CMBCuaseStock.Text), InStr(Trim(CMBCuaseStock.Text), "//") - 1)
        ListView103.ListItems.Item(vIndex).SubItems(10) = Right(Trim(CMBCuaseStock.Text), Len(Trim(CMBCuaseStock.Text)) - InStr(Trim(CMBCuaseStock.Text), "//") - 1)
        End If
        ListView103.ListItems.Item(vIndex).SubItems(11) = Trim(TextDescription.Text)
        
        ListView103.SetFocus
        If vIndex < Me.ListView103.ListItems.Count Then
        Me.ListView103.ListItems.Item(vIndex + 1).Selected = True
        End If

  End If
Else
  MsgBox "ไม่มีข้อมูลในการปรับจำนวนการตรวจนับ ต้องกดเพิ่มรายการเมื่อไม่มีรายการในตาราง", vbCritical, "Send Error"
End If

End Sub

Private Sub Command1_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer, j As Integer
Dim vItemCode As String, vBarCode As String

On Error GoTo ErrDescription

For i = 1 To 1000
vItemCode = "50" & Format(i, "00000")
vBarCode = vItemCode
vQuery = "Insert into BCnp.dbo.Report_Temp3 (itemcode,barcode,type)" _
                & "  select '" & vItemCode & "' as Itemcode,'" & vBarCode & "' as Barcode,1 as Type"
gConnection.Execute vQuery
vQuery = "Insert into BCnp.dbo.Report_Temp3 (itemcode,barcode,type)" _
                & "  select '" & vItemCode & "' as Itemcode,'" & vBarCode & "' as Barcode,1 as Type"
gConnection.Execute vQuery

Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next


Call InitializeDataBase2

vQuery = "select top 10 docno from bcnp.dbo.bcstkinspect order by docdate desc "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Cmb102.AddItem Trim(vRecordset.Fields("docno").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

Call SearchStkInspectLogs

Me.Text101.SetFocus

End Sub

Public Sub GetCauseProductNegative()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

CMBCuaseStock.Clear
vQuery = "exec dbo.USP_MB_SearchCauseProductNegative"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then '
  vRecordset.MoveFirst
  While Not vRecordset.EOF
    CMBCuaseStock.AddItem Trim(vRecordset.Fields("causename").Value)
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

Public Sub SearchStkInspectLogs()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim n As Integer
Dim vShelfStock As String
Dim vItem As String
Dim vItemName As String
Dim vWH As String
Dim vShelf As String
Dim vUnit As String
Dim vQty As Double
Dim vDiff As Double
Dim vInspectQTY As Double
Dim vReasonCode As String
Dim ListItem  As ListItem

On Error GoTo ErrDescription

Me.ListView101.ListItems.Clear
vQuery = "exec dbo.USP_MB_SearchStkInspect_Logs '" & vUserID & "'"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  vRecordset.MoveFirst
  While Not vRecordset.EOF
  
    vShelfStock = vRecordset.Fields("stockshelf").Value
    vItem = vRecordset.Fields("itemcode").Value
    vItemName = vRecordset.Fields("itemname").Value
    vWH = vRecordset.Fields("whcode").Value
    vShelf = vRecordset.Fields("shelfcode").Value
    vUnit = vRecordset.Fields("unitcode").Value
    vQty = vRecordset.Fields("qty").Value
    vDiff = vRecordset.Fields("diff").Value
    vInspectQTY = vRecordset.Fields("countqty").Value
    vReasonCode = vRecordset.Fields("reasoncode").Value

    Set ListItem = ListView101.ListItems.Add(, , Trim(vItem))
    ListItem.SubItems(1) = Format(vQty, "##,##0.00")
    ListItem.SubItems(2) = Format(vInspectQTY, "##,##0.00")
    ListItem.SubItems(3) = Format(vDiff, "##,##0.00")
    ListItem.SubItems(4) = ""
    ListItem.SubItems(5) = vItemName
    ListItem.SubItems(6) = vUnit
    ListItem.SubItems(7) = vShelf
    ListItem.SubItems(8) = Now
    ListItem.SubItems(9) = vWH
    ListItem.SubItems(10) = vShelfStock
    ListItem.SubItems(11) = vReasonCode

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


Public Sub SearchShelfCode()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

CMBZone.Clear
vQuery = "select distinct code from dbo.bcshelf order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then '
  vRecordset.MoveFirst
  While Not vRecordset.EOF
    Me.CMBZone.AddItem Trim(vRecordset.Fields("code").Value)
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

Private Sub Label3_Click()

End Sub

Private Sub ListView101_DblClick()
On Error GoTo ErrDescription


  If ListView101.ListItems.Count > 0 Then
    vItemClick = ListView101.SelectedItem.Index
    Form102_1.Show
    Form102_1.SetFocus
    Form102.Enabled = False
    Form102_1.Label101.Caption = Form102.ListView101.ListItems(vItemClick).Text
    Form102_1.Label102.Caption = Form102.ListView101.ListItems.Item(vItemClick).SubItems(5)
    Form102_1.LBLWHCode.Caption = Form102.ListView101.ListItems.Item(vItemClick).SubItems(9)
    Form102_1.LBLShelfCode.Caption = Form102.ListView101.ListItems.Item(vItemClick).SubItems(10)
    Form102_1.Label103.Caption = Format(Form102.ListView101.ListItems.Item(vItemClick).SubItems(2), "##,##0.00")
    Form102_1.Text101.SetFocus
  End If

ErrDescription:
If Err.Description <> "" Then
  MsgBox Err.Description
  Exit Sub
End If
End Sub


Private Sub ListView101_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
  ListView103.SetFocus
ElseIf KeyCode = 113 Then
  ListView101.SetFocus
End If

End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
On Error GoTo ErrDescription


If KeyAscii = 13 Then
  If ListView101.ListItems.Count > 0 Then
    vItemClick = ListView101.SelectedItem.Index
    Form102_1.Show
    Form102_1.SetFocus
    Form102.Enabled = False
    Form102_1.Label101.Caption = Form102.ListView101.ListItems(vItemClick).Text
    Form102_1.Label102.Caption = Form102.ListView101.ListItems.Item(vItemClick).SubItems(5)
    Form102_1.LBLWHCode.Caption = Form102.ListView101.ListItems.Item(vItemClick).SubItems(9)
    Form102_1.LBLShelfCode.Caption = Form102.ListView101.ListItems.Item(vItemClick).SubItems(10)
    Form102_1.Label103.Caption = Form102.ListView101.ListItems.Item(vItemClick).SubItems(2)
    Form102_1.Text101.SetFocus
  End If
End If

ErrDescription:
If Err.Description <> "" Then
  MsgBox Err.Description
  Exit Sub
End If
End Sub

Private Sub ListView101_KeyUp(KeyCode As Integer, Shift As Integer)
If ListView101.ListItems.Count > 0 Then
  If KeyCode = 119 Then
    Call CMD105_Click
  End If
End If

End Sub

Private Sub ListView103_DblClick()
Dim vItemCode As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset


On Error GoTo ErrDescription

  If Me.CMBCuaseStock.Text <> "" Then
  If ListView103.ListItems.Count > 0 Then
    vIndex = ListView103.SelectedItem.Index
    TextCount.Text = ""
    Pic101.Visible = True
    Me.LBLIndex.Caption = vIndex
    LBLItemCode.Caption = ListView103.ListItems.Item(vIndex).SubItems(5)
    LBLItemName.Caption = ListView103.ListItems.Item(vIndex).SubItems(6)
    LBLWHCode.Text = Trim(ListView103.ListItems.Item(vIndex).Text)
    LBLShelfCode.Text = Trim(ListView103.ListItems.Item(vIndex).SubItems(1))
    LBLOnHand.Caption = Format(ListView103.ListItems.Item(vIndex).SubItems(2), "##,##0.00")
    LBLUnitCode1.Caption = ListView103.ListItems.Item(vIndex).SubItems(4)
    LBLUnitCode2.Caption = ListView103.ListItems.Item(vIndex).SubItems(4)
    TextCount.Text = Format(ListView103.ListItems.Item(vIndex).SubItems(3), "##,##0.00")
    TextShelfCode.Text = ListView103.ListItems.Item(vIndex).SubItems(8)
    If ListView103.ListItems.Item(vIndex).SubItems(9) <> "" Then
      CMBCuaseStock.Text = Trim(ListView103.ListItems.Item(vIndex).SubItems(9)) & "//" & Trim(ListView103.ListItems.Item(vIndex).SubItems(10))
    End If
    If Trim(ListView103.ListItems.Item(vIndex).SubItems(11)) <> "" Then
      TextDescription.Text = Trim(ListView103.ListItems.Item(vIndex).SubItems(11))
    Else
      TextDescription.Text = ""
    End If
    TextShelfCode.SetFocus
  Else
  
    If Me.LBLItemDescription.Caption <> "" Then
      Pic101.Visible = True
      vItemCode = Left(Trim(Me.LBLItemDescription.Caption), InStr(Trim(Me.LBLItemDescription.Caption), "//") - 1)
      vQuery = "exec dbo.USP_IV_CheckItemDescription '" & vItemCode & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        LBLItemCode.Caption = Trim(vRecordset.Fields("code").Value)
        LBLItemName.Caption = Trim(vRecordset.Fields("name1").Value)
        LBLWHCode.Text = Trim(vRecordset.Fields("defsalewhcode").Value)
        LBLShelfCode.Text = Trim(vRecordset.Fields("defsaleshelf").Value)
        LBLOnHand.Caption = 0
        LBLUnitCode1.Caption = Trim(vRecordset.Fields("defsaleunitcode").Value)
        LBLUnitCode2.Caption = Trim(vRecordset.Fields("defsaleunitcode").Value)
        TextCount.Text = ""
        TextShelfCode.Text = ""
      End If
      vRecordset.Close
    End If
  End If
Else
   MsgBox "กรณีไม่ได้ ระบุเหตุผลในการตรวจนับ กรุณาระบุก่อนด้วย", vbCritical, "Send Error Message"
   Me.CMBCuaseStock.SetFocus
End If
   
  
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListView103_KeyDown(KeyCode As Integer, Shift As Integer)
If ListView103.ListItems.Count > 0 Then
  If KeyCode = 116 Then
    Call CMDBasket_Click
  End If
End If
If KeyCode = 112 Then
  ListView103.SetFocus
ElseIf KeyCode = 113 Then
  ListView101.SetFocus
End If
End Sub

Private Sub ListView103_KeyPress(KeyAscii As Integer)
Dim vItemCode As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If KeyAscii = 13 And Me.CMBCuaseStock.Text <> "" Then

  If ListView103.ListItems.Count > 0 Then
    vIndex = ListView103.SelectedItem.Index
    TextCount.Text = ""
    Pic101.Visible = True
    Me.LBLIndex.Caption = vIndex
    LBLItemCode.Caption = ListView103.ListItems.Item(vIndex).SubItems(5)
    LBLItemName.Caption = ListView103.ListItems.Item(vIndex).SubItems(6)
    LBLWHCode.Text = Trim(ListView103.ListItems.Item(vIndex).Text)
    LBLShelfCode.Text = Trim(ListView103.ListItems.Item(vIndex).SubItems(1))
    LBLOnHand.Caption = ListView103.ListItems.Item(vIndex).SubItems(2)
    LBLUnitCode1.Caption = ListView103.ListItems.Item(vIndex).SubItems(4)
    LBLUnitCode2.Caption = ListView103.ListItems.Item(vIndex).SubItems(4)
    TextCount.Text = ListView103.ListItems.Item(vIndex).SubItems(3)
    TextShelfCode.Text = ListView103.ListItems.Item(vIndex).SubItems(8)
    If ListView103.ListItems.Item(vIndex).SubItems(9) <> "" Then
      CMBCuaseStock.Text = Trim(ListView103.ListItems.Item(vIndex).SubItems(9)) & "//" & Trim(ListView103.ListItems.Item(vIndex).SubItems(10))
    End If
    If Trim(ListView103.ListItems.Item(vIndex).SubItems(11)) <> "" Then
      TextDescription.Text = Trim(ListView103.ListItems.Item(vIndex).SubItems(11))
    Else
      TextDescription.Text = ""
    End If
    TextShelfCode.SetFocus
  Else
  
    If Me.LBLItemDescription.Caption <> "" Then
      Pic101.Visible = True
      Call GetCauseProductNegative
      vItemCode = Left(Trim(Me.LBLItemDescription.Caption), InStr(Trim(Me.LBLItemDescription.Caption), "//") - 1)
      vQuery = "exec dbo.USP_IV_CheckItemDescription '" & vItemCode & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        LBLItemCode.Caption = Trim(vRecordset.Fields("code").Value)
        LBLItemName.Caption = Trim(vRecordset.Fields("name1").Value)
        LBLWHCode.Text = Trim(vRecordset.Fields("defsalewhcode").Value)
        LBLShelfCode.Text = Trim(vRecordset.Fields("defsaleshelf").Value)
        LBLOnHand.Caption = 0
        LBLUnitCode1.Caption = Trim(vRecordset.Fields("defsaleunitcode").Value)
        LBLUnitCode2.Caption = Trim(vRecordset.Fields("defsaleunitcode").Value)
        TextCount.Text = ""
        TextShelfCode.Text = ""
      End If
      vRecordset.Close
    End If
  End If

Else
   If Me.CMBCuaseStock.Text = "" Then
   MsgBox "กรณีไม่ได้ ระบุเหตุผลในการตรวจนับ กรุณาระบุก่อนด้วย", vbCritical, "Send Error Message"
   Me.CMBCuaseStock.SetFocus
   End If

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListView103_KeyUp(KeyCode As Integer, Shift As Integer)
If ListView101.ListItems.Count > 0 Then
  If KeyCode = 119 Then
    Call CMD105_Click
  End If
End If
End Sub

Private Sub OPTAll_Click()
Me.CMBZone.Clear
End Sub

Private Sub OPTZone_Click()
Call SearchShelfCode
Me.CMBZone.SetFocus
End Sub

Private Sub Text101_Change()
If Me.Text101.Text = "" Then
    ListView103.ListItems.Clear
    LBLItemDescription.Caption = ""
    Me.CMBCuaseStock.Clear
End If
End Sub

Private Sub Text101_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
  ListView103.SetFocus
ElseIf KeyCode = 113 Then
  ListView101.SetFocus
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)

On Error GoTo ErrDescription

If KeyAscii = 13 Then
      vIndex = 0
      Call CheckBar
      If vChkBar = 1 Then
          If vConnect = 0 Then
                  Call CheckStockLocation
                  Me.CMBCuaseStock.SetFocus
          End If
    Else
          MsgBox "ไม่มีสินค้า บาร์โค้ด " & Text101.Text & " กรุณาตรวจสอบด้วยนะครับ "
          Text101.Text = ""
      End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub CheckItemInsert()
Dim i As Integer
Dim vCheckNewItem As String
Dim vCheckShelfItem As String
Dim vCheckStockShelfItem As String
Dim vTableItem As String
Dim vTableShelf As String
Dim vTableStockShelf As String

On Error GoTo ErrDescription

If ListView101.ListItems.Count > 0 Then
  vCountItemCode = ListView101.ListItems.Count
  vCheckNewItem = Trim(Text101.Text)
  vCheckShelfItem = Trim(TextShelfCode.Text)
  vCheckStockShelfItem = Trim(LBLShelfCode.Text)
  For i = 1 To vCountItemCode
    vTableItem = Trim(ListView101.ListItems.Item(i).Text)
    vTableShelf = Trim(ListView101.ListItems.Item(i).SubItems(7))
    vTableStockShelf = Left(Trim(ListView101.ListItems.Item(i).SubItems(10)), InStr(Trim(ListView101.ListItems.Item(i).SubItems(10)), "//") - 1)
    If vCheckNewItem = vTableItem And vCheckShelfItem = vTableShelf And vCheckStockShelfItem = vTableStockShelf Then
      MsgBox "ในตารางข้างล่างมีสินค้ารหัส " & vCheckNewItem & " ที่อยู่ " & vCheckShelfItem & " และ ชั้นเก็บ " & vCheckStockShelfItem & " อยู่ในรายการที่ " & i & " กรุณาตรวจสอบด้วย", vbCritical, "Send Information"
      vItemExist = 1
      Exit Sub
    Else
      vItemExist = 0
    End If
  Next i
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub InsertToGrid()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim ListItem As ListItem
Dim vBarCode As String
Dim vShelfCode As String

'On Error GoTo ErrDescription

'vBarCode = Trim(Text101.Text)
'vDateCountStock = DTPicker1.Day & "/" & DTPicker1.Month & "/" & DTPicker1.Year
'vShelfCode = Left(Trim(CMB103.Text), InStr(Trim(CMB103.Text), "//") - 1)
'vQuery = "set dateformat dmy "
'gConnection.Execute vQuery

'vQuery = "select * from bcnp.dbo. vw_IV_ProgStockChecking where barcode = '" & vBarCode & "' and whcode = '" & Cmb101.Text & "'  and shelfcode = '" & vShelfCode & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   Set ListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("barcode").Value))
  '  ListItem.SubItems(1) = Trim(vRecordset.Fields("qty").Value)
   ' ListItem.SubItems(2) = Trim(Text102.Text)
    'ListItem.SubItems(3) = Trim(Text102.Text) - Trim(vRecordset.Fields("qty").Value)
    'ListItem.SubItems(5) = Trim(vRecordset.Fields("name1").Value)
    'ListItem.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
    'ListItem.SubItems(7) = UCase(Trim(Text103.Text))
    'ListItem.SubItems(8) = Now
    'ListItem.SubItems(9) = Trim(Cmb101.Text)
    'ListItem.SubItems(10) = Trim(CMB103.Text)
'Else
 '   MsgBox "ไม่มีข้อมูลจำนวนสินค้าของ บาร์โค้ด รหัส " & vBarCode & " ในฐานข้อมูล คลัง " & Cmb101.Text & " และ ชั้นเก็บ " & vShelfCode & " กรุณาตรวจสอบ รหัสบาร์โค้ดด้วยนะครับ"
  '  vQuery = "select * from bcnp.dbo. vw_IV_ProgStockChecking where barcode = '" & vBarCode & "'  "
   ' If OpenDataBase(vConnection, vRecordset1, vQuery) <> 0 Then
    'Set ListItem = ListView101.ListItems.Add(, , vBarCode)
    'ListItem.SubItems(1) = 0
    'ListItem.SubItems(2) = Trim(Text102.Text)
    'ListItem.SubItems(3) = Trim(Text102.Text) - 0
    'ListItem.SubItems(5) = Trim(vRecordset1.Fields("name1").Value)
    'ListItem.SubItems(6) = Trim(vRecordset1.Fields("unitcode").Value)
    'ListItem.SubItems(7) = UCase(Trim(Text103.Text))
    'ListItem.SubItems(8) = Now
    'ListItem.SubItems(9) = Trim(Cmb101.Text)
    'ListItem.SubItems(10) = Trim(CMB103.Text)
    'End If
    'vRecordset1.Close
'End If
'vRecordset.Close
'ListView103.ListItems.Clear
'ListView104.ListItems.Clear
'Text101.Text = ""
'Text102.Text = ""
'Text101.SetFocus

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Public Sub InsertToGrid_UnConnect()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListItem As ListItem
Dim vBarCode

On Error GoTo ErrDescription

vBarCode = Trim(Text101.Text)
    Set ListItem = ListView101.ListItems.Add(, , Trim(Text101.Text))
    'ListItem.SubItems(4) = Trim(Text102.Text)
Text101.Text = ""
'Text102.Text = ""
Text101.SetFocus

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub CheckBarCode()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vBarCode As String

On Error GoTo ErrDescription

vBarCode = Trim(Text101.Text)
vQuery = "select barcode from bcnp.dbo.bcbarcodemaster where barcode = '" & vBarCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckBarCode = 1
    vError = ""
Else
    MsgBox "ไม่มีบาร์โค้ด รหัส " & vBarCode & " ในทะเบียนบาร์โค้ด  "
    vError = "ไม่มีบาร์โค้ด"
    vCheckBarCode = 0
    Text101.Text = ""
    Text101.SetFocus
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub CheckStock()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vBarCode As String, vWHCode As String
Dim vShelfCode As String

On Error GoTo ErrDescription
vBarCode = Trim(Text101.Text)
vWHCode = Trim(LBLWHCode.Text)
vShelfCode = Trim(LBLShelfCode.Text)
vQuery = "set dateformat dmy"
gConnection.Execute vQuery
vQuery = "exec dbo.USP_MB_CheckStockBarcode '" & vBarCode & "' ,'" & vWHCode & "' ,'" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckStock = 1
    vError = ""
Else

    vError = vError & " ไม่มียอดสต็อก"
    vCheckStock = 1
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub CheckBar()
Dim vBarCode As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListItem As ListItem

On Error GoTo ErrDescription

vBarCode = Trim(Text101.Text)
ListView103.ListItems.Clear
vQuery = "exec dbo.USP_MB_CheckBarcodeExist '" & vBarCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vChkBar = 1
    Call GetCauseProductNegative
Else
vChkBar = 0
End If
vRecordset.Close
            
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub CheckStockLocation()
Dim vBarCode As String
Dim vWHCode As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListItem As ListItem
Dim vZone As String

On Error GoTo ErrDescription

vBarCode = Trim(Text101.Text)
ListView103.ListItems.Clear


vQuery = "exec dbo.USP_IV_CheckItemDescription '" & vBarCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    LBLItemDescription.Caption = Trim(vRecordset.Fields("code").Value) & "//" & Trim(vRecordset.Fields("name1").Value)
    Else
    ListView103.ListItems.Clear
    Text101.Text = ""
    Text101.SetFocus
    End If
vRecordset.Close
    
    
If Me.OPTAll.Value = True Then
vZone = "ALL"
ElseIf Me.OPTZone.Value = True And Me.CMBZone.Text <> "" Then
vZone = Me.CMBZone.Text
Else
vZone = "ALL"
End If

vQuery = "exec dbo.USP_MB_StockLocation'" & vBarCode & "','" & vZone & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst

    While Not vRecordset.EOF
    Set ListItem = ListView103.ListItems.Add(, , Trim(vRecordset.Fields("whcode").Value))
    ListItem.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
    ListItem.SubItems(2) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
    ListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
    ListItem.SubItems(5) = Trim(vRecordset.Fields("itemcode").Value)
    ListItem.SubItems(6) = Trim(vRecordset.Fields("itemname").Value)
    ListItem.SubItems(7) = Trim(vRecordset.Fields("barcode").Value)
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

Private Sub Text101_KeyUp(KeyCode As Integer, Shift As Integer)
If ListView101.ListItems.Count > 0 Then
  If KeyCode = 119 Then
    Call CMD105_Click
 End If
End If
End Sub


Public Sub GetInspectNo()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocno As String
Dim vYear, vYear1 As String
Dim vMonth, vMonth1 As String
Dim vItem(500), vUnitCode(500), vShelf, vItemName(500) As String
Dim vQty(500), vDiff(500), vInspectQTY(500) As Currency
Dim vCountItem As Currency
Dim vSumItem(500), i, j As Currency
Dim vItemCode(500) As String
Dim vShelfCode(500) As String
Dim vWHCode(500) As String
Dim vInSpectDesc(500) As String

On Error GoTo ErrDescription

vQuery = "select top 1 docno from bcnp.dbo.bcstkinspect  where docno like 's01%' order by docno desc"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDocno = Trim(vRecordset.Fields("docno").Value)
End If
vRecordset.Close

If Left(vCheckDocno, 2) = "IH" Then
    vYear = Mid(vCheckDocno, 3, 2)
    vMonth = Mid(vCheckDocno, 5, 2)
    vYear1 = Mid(Year(Now), 3, 2)
    vMonth1 = Month(Now)
    If vYear1 < 48 Then
        vYear1 = vYear1 + 43
    End If
    If Len(vMonth1) <> 2 Then
        vMonth1 = "0" & vMonth1
    End If
    If vYear1 = vYear And vMonth1 = vMonth Then
            vQuery = "select * from V_WEB_IV_ItemCheck_NewDocNo"
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
               vNewDocNo = Trim(vRecordset.Fields("newdocno").Value)
            End If
            vRecordset.Close
    Else
        vNewDocNo = "S01" & "-" & Trim("IH" & vYear1 & vMonth1 & "-0001")
    End If
ElseIf Left(vCheckDocno, 3) = "S01" Then

Dim vLen As Integer
Dim vDocNo As String

vLen = Len(vCheckDocno)
vDocNo = Right(vCheckDocno, vLen - 4)

    vYear = Mid(vDocNo, 3, 2)
    vMonth = Mid(vDocNo, 5, 2)
    vYear1 = Mid(Year(Now), 3, 2)
    vMonth1 = Month(Now)
    If vYear1 < 48 Then
        vYear1 = vYear1 + 43
    End If
    If Len(vMonth1) <> 2 Then
        vMonth1 = "0" & vMonth1
    End If
    If vYear1 = vYear And vMonth1 = vMonth Then
            vQuery = "select * from V_WEB_IV_ItemCheck_NewDocNo"
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
               vNewDocNo = Trim(vRecordset.Fields("newdocno").Value)
            End If
            vRecordset.Close
    Else
        vNewDocNo = "S01" & "-" & Trim("IH" & vYear1 & vMonth1 & "-0001")
    End If
Else
    vYear1 = Mid(Year(Now), 3, 2)
    vMonth1 = Month(Now)
    If vYear1 < 48 Then
        vYear1 = vYear1 + 43
    End If
    If Len(vMonth1) <> 2 Then
        vMonth1 = "0" & vMonth1
    End If
    
    vNewDocNo = "S01" & "-" & Trim("IH" & vYear1 & vMonth1 & "-0001")
End If

'vQuery = "select count(itemcode) as countitem from (select distinct itemcode,whcode,stockshelf from npmaster.dbo.TB_IV_StkInspect_Logs where docno = '" & vGenDocNo & "') as a "
vQuery = "select count(itemcode) as countitem from (select distinct itemcode,whcode,stockshelf from npmaster.dbo.TB_IV_StkInspect_Logs where docno = '" & vGenDocNo & "') as a "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCountItem = Trim(vRecordset.Fields("countitem").Value)
End If
vRecordset.Close

vQuery = "exec dbo.USP_NP_InsertBCSTKInspect '" & vNewDocNo & "','" & vUserID & "' "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_NP_SelectItemInspect '" & vGenDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        j = 0
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        j = j + 1
        vItemCode(j) = Trim(vRecordset.Fields("itemcode").Value)
        vWHCode(j) = Trim(vRecordset.Fields("whcode").Value)
        vShelfCode(j) = Trim(vRecordset.Fields("stockshelf").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

For i = 1 To vCountItem
vQuery = "exec dbo.USP_NP_SelectItemDetailsInspect '" & vGenDocNo & "' , '" & vItemCode(i) & "' ,'" & vShelfCode(i) & "' ,'" & vWHCode(i) & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vItemName(i) = Trim(vRecordset.Fields("itemname").Value)
    vUnitCode(i) = Trim(vRecordset.Fields("unitcode").Value)
    vInSpectDesc(i) = Trim(vRecordset.Fields("reasoncode").Value)
End If
vRecordset.Close

vQuery = "exec dbo.USP_NP_SelectSumItemQtyInspect '" & vGenDocNo & "','" & vItemCode(i) & "','" & vWHCode(i) & "','" & vShelfCode(i) & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vSumItem(i) = vRecordset.Fields("qty").Value
Else
    vSumItem(i) = 0
End If
vRecordset.Close

vQuery = "set dateformat dmy"
gConnection.Execute vQuery
vQuery = "exec dbo.USP_NP_SelectItemQtySTKLocation '" & vItemCode(i) & "','" & vWHCode(i) & "','" & vShelfCode(i) & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vInspectQTY(i) = vRecordset.Fields("qty").Value
End If
vRecordset.Close

vDiff(i) = vSumItem(i) - vInspectQTY(i)
Next i
For i = 1 To vCountItem
vQuery = "exec dbo.USP_NP_InsertBCInspectSub '" & vNewDocNo & "','" & vItemCode(i) & "','" & vUnitCode(i) & "','" & vWHCode(i) & "','" & vShelfCode(i) & "'," & vInspectQTY(i) & "," & vSumItem(i) & "," & vDiff(i) & ",'" & vInSpectDesc(i) & "' "
gConnection.Execute vQuery
Next i
vQuery = "Update npmaster.dbo.TB_IV_StkInspect_Logs set Inspectno = '" & vNewDocNo & "' where docno = '" & vGenDocNo & "' "
gConnection.Execute vQuery
MsgBox "ได้เอกสารตรวจนับเลขที่ " & vNewDocNo & " "

ErrDescription:
  If Err.Description <> "" Then
  MsgBox Err.Description
  Exit Sub
  End If
End Sub
Public Sub GetDocNo()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vYear, vYear1 As String
Dim vMonth, vMonth1 As String
Dim vHeader As String
Dim vAutoNumber As Integer

On Error GoTo ErrDescription

vQuery = "exec dbo.USP_NP_SearchNewDocNo 10 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vHeader = Trim(vRecordset.Fields("header").Value)
    vAutoNumber = Trim(vRecordset.Fields("AutoNumber").Value)
End If
vRecordset.Close
vYear1 = Mid(Year(Now), 3, 2)
If vYear1 < 48 Then
    vYear1 = vYear1 + 43
End If
vGenDocNo = Trim(vHeader & vYear1 & "-" & Format(vAutoNumber, "0000"))
vQuery = "exec dbo.USP_NP_UpdateNewDocNo  10 "
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub PrintInspection()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType, vReportName, vDocNo As String


vRepID = 213
vRepType = "IV"
vDocNo = vNewDocNo
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

End Sub

Private Sub TextCount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Me.CMDOK.SetFocus
End If
End Sub

Private Sub TextCount_LostFocus()

Call CheckNumber(Trim(Me.TextCount.Text))
If vCheckValueNumber = False Then
   MsgBox "กรอกข้อมูลได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
   Me.TextCount.SetFocus
   Me.TextCount.Text = 0
Else
   Me.TextCount.Text = Format(Me.TextCount.Text, "##,##0.00")
End If
End Sub

Private Sub TextShelfCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  TextCount.SetFocus
End If
End Sub

Private Sub TextShelfCode_LostFocus()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckShelfExist As Integer
Dim vWHCode As String
Dim vShelfCode As String

On Error GoTo ErrDescription

If Me.TextShelfCode.Text <> "" Then
   vWHCode = Me.LBLWHCode.Text
   vShelfCode = Me.TextShelfCode.Text
   
   vQuery = "select isnull(count(code),0) as vCount from  Npmaster.dbo.TB_RC_Shelf where whcode = '" & vWHCode & "' and code = '" & vShelfCode & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckShelfExist = vRecordset.Fields("vcount").Value
   End If
   vRecordset.Close
   If vCheckShelfExist = 0 And Text101.Text <> "" Then
      MsgBox "ที่เก็บ " & vShelfCode & " ไม่มีในทะเบียนที่เก็บ ", vbCritical, "Send Error Message"
      Me.TextShelfCode.Text = UCase(Me.TextShelfCode.Text)
   Else
      Me.TextShelfCode.Text = UCase(Me.TextShelfCode.Text)
   End If
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
