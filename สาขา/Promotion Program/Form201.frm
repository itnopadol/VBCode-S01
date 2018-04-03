VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form201 
   Caption         =   "สร้างใบเสนอสินค้า"
   ClientHeight    =   9000
   ClientLeft      =   1545
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD110 
      Caption         =   "พิมพ์เอกสาร"
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
      Left            =   8625
      TabIndex        =   36
      Top             =   8100
      Width           =   1440
   End
   Begin VB.CommandButton CMD109 
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
      Height          =   465
      Left            =   10500
      TabIndex        =   35
      Top             =   8100
      Width           =   1440
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6465
      Left            =   75
      TabIndex        =   23
      Top             =   1425
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   11404
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   529
      TabCaption(0)   =   "รายละเอียดสินค้า"
      TabPicture(0)   =   "Form201.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ItemDetail101"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ItemDetail105"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ItemDetail102"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ItemDetail103"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ItemDetail104"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ItemDetail106"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ListView101"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CHK102"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CMD104"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "CMD105"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CHK103"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Check101"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "ItemDetail107"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "CMD111"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "CMD108"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Crystal101"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      Begin Crystal.CrystalReport Crystal101 
         Left            =   750
         Top             =   5850
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
      Begin VB.CommandButton CMD108 
         Caption         =   "ลบรายการ"
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
         Height          =   465
         Left            =   10200
         TabIndex        =   39
         ToolTipText     =   "ปุ่มนี้ ใช้สำหรับลบรายการสินค้าในตารางข้างบน"
         Top             =   5850
         Width           =   1440
      End
      Begin VB.CommandButton CMD111 
         Height          =   315
         Left            =   7425
         Picture         =   "Form201.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "ปุ่มนี้ ใช้สำหรับประเภทโปรโมชั่นของสินค้า"
         Top             =   1650
         Width           =   315
      End
      Begin VB.TextBox ItemDetail107 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4350
         TabIndex        =   12
         Top             =   1650
         Width           =   3015
      End
      Begin VB.CheckBox Check101 
         Caption         =   "IsBrochure"
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
         Left            =   1575
         TabIndex        =   11
         Top             =   1650
         Width           =   1215
      End
      Begin VB.CheckBox CHK103 
         Caption         =   "ลดเป็น %"
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
         Left            =   7725
         TabIndex        =   9
         Top             =   1200
         Width           =   1065
      End
      Begin VB.CommandButton CMD105 
         Caption         =   "เพิ่มสินค้า"
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
         Left            =   225
         Picture         =   "Form201.frx":03E9
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "ปุ่มนี้ ใช้สำหรับเพิ่มสินค้าทั้งสินค้าตัวใหม่หรืออัพเดท"
         Top             =   1875
         Width           =   1140
      End
      Begin VB.CommandButton CMD104 
         Height          =   315
         Left            =   3525
         Picture         =   "Form201.frx":0703
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "ปุ่มนี้ ใช้สำหรับค้นหารหัสสินค้าหรือชื่อสินค้า"
         Top             =   750
         Width           =   315
      End
      Begin VB.CheckBox CHK102 
         Caption         =   "สมาชิก"
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
         Left            =   1575
         TabIndex        =   5
         Top             =   375
         Width           =   1890
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   3315
         Left            =   225
         TabIndex        =   16
         Top             =   2400
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5847
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสสินค้า"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อสินค้า"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ราคาปกติ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ราคาโปรโมชั่น"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ลดราคา (บาท)"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ส่วนลด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "หน่วยนับ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "หมายเหตุ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ประเภทการลด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "สมาชิก"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "IsBrochure"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "IsCancel"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "PromotoinType"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "PromotoinTypeCode"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.TextBox ItemDetail106 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8550
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1650
         Width           =   3090
      End
      Begin VB.TextBox ItemDetail104 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4950
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox ItemDetail103 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1575
         TabIndex        =   7
         Top             =   1200
         Width           =   2265
      End
      Begin VB.TextBox ItemDetail102 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4950
         TabIndex        =   30
         Top             =   750
         Width           =   6690
      End
      Begin VB.TextBox ItemDetail105 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9675
         TabIndex        =   10
         Top             =   1200
         Width           =   1965
      End
      Begin VB.TextBox ItemDetail101 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1575
         TabIndex        =   6
         Top             =   750
         Width           =   1890
      End
      Begin VB.Label Label5 
         Caption         =   "ประเภทโปรโมชั่น"
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
         Left            =   3000
         TabIndex        =   37
         Top             =   1650
         Width           =   1290
      End
      Begin VB.Label Label7 
         Caption         =   "หมายเหตุ"
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
         Left            =   7800
         TabIndex        =   29
         Top             =   1650
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "ส่วนลด"
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
         Left            =   9000
         TabIndex        =   28
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "ราคาปกติ"
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
         Left            =   675
         TabIndex        =   27
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "หน่วยนับ"
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
         Left            =   4200
         TabIndex        =   26
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "รหัสสินค้า"
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
         Left            =   675
         TabIndex        =   25
         Top             =   750
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "ชื่อสินค้า"
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
         Left            =   4200
         TabIndex        =   24
         Top             =   750
         Width           =   690
      End
   End
   Begin VB.CommandButton CMD107 
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
      Height          =   465
      Left            =   6750
      TabIndex        =   22
      ToolTipText     =   "ปุ่มนี้ ใช้สำหรับบันทึกและอัพเดท"
      Top             =   8100
      Width           =   1440
   End
   Begin VB.Frame Frame2 
      Caption         =   "รายละเอียดใบเสนอสินค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   75
      TabIndex        =   17
      Top             =   75
      Width           =   11865
      Begin VB.CommandButton CMD112 
         Height          =   315
         Left            =   4575
         Picture         =   "Form201.frx":0AD0
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "ปุ่มนี้ ใช้สำหรับลบเอกสาร (ปรับ IScancel = 1ทั้งใบ)"
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton CMD106 
         Height          =   315
         Left            =   4200
         Picture         =   "Form201.frx":0F42
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "ปุ่มนี้ ใช้สำหรับค้นหาเอกสารเสนอสินค้าโปรโมชั่น"
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton CMD103 
         Height          =   315
         Left            =   10200
         Picture         =   "Form201.frx":130F
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "ปุ่มนี้ ใช้สำหรับค้นหา Section Manager"
         Top             =   825
         Width           =   315
      End
      Begin VB.CommandButton CMD102 
         Height          =   315
         Left            =   10200
         Picture         =   "Form201.frx":16DC
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "ปุ่มนี้ ใช้สำหรับค้นหาทะเบียนโปรโมชั่น"
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox Text103 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6825
         TabIndex        =   4
         Top             =   825
         Width           =   3315
      End
      Begin VB.TextBox Text102 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6825
         TabIndex        =   3
         Top             =   300
         Width           =   3315
      End
      Begin MSComCtl2.DTPicker DTPicker101 
         Height          =   315
         Left            =   2250
         TabIndex        =   2
         Top             =   825
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62717953
         CurrentDate     =   38504
      End
      Begin VB.CommandButton CMD101 
         Height          =   315
         Left            =   3825
         Picture         =   "Form201.frx":1AA9
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "ปุ่มนี้ ใช้สำหรับสร้างเลขที่เอกสารเลขที่ใหม่"
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox Text101 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   2250
         TabIndex        =   0
         Top             =   300
         Width           =   1515
      End
      Begin VB.Image Image101 
         Height          =   300
         Left            =   150
         Picture         =   "Form201.frx":1E00
         ToolTipText     =   "แสดงสถานะเอกสาร N : New CF : Confirm CC : Cancel"
         Top             =   225
         Width           =   570
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Sec. Manager "
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
         Left            =   5250
         TabIndex        =   21
         Top             =   825
         Width           =   1440
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "โปรโมชั่น"
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
         Left            =   5700
         TabIndex        =   20
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label16 
         Caption         =   "วันที่เอกสาร"
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
         Left            =   1125
         TabIndex        =   19
         Top             =   825
         Width           =   990
      End
      Begin VB.Label Label15 
         Caption         =   "เลขที่เอกสาร"
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
         Left            =   1125
         TabIndex        =   18
         Top             =   300
         Width           =   990
      End
      Begin VB.Image Image102 
         Height          =   300
         Left            =   150
         Picture         =   "Form201.frx":2232
         Top             =   225
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Image Image103 
         Height          =   300
         Left            =   150
         Picture         =   "Form201.frx":26DB
         Top             =   225
         Visible         =   0   'False
         Width           =   570
      End
   End
End
Attribute VB_Name = "Form201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNEwDocno As String
Dim vCheckRequestOnListview As Integer
Dim vIndexItemUpdate As Integer
Dim vSortResult As Integer

Private Sub CHK102_Click()
Dim vDiscount As Currency

On Error Resume Next

If CHK102.Value = 1 Then
    ItemDetail105.Enabled = False
    CHK103.Enabled = False
    vDiscount = (ItemDetail103.Text - (ItemDetail103.Text * vMemberDiscount) / 100)
    ItemDetail105.Text = vDiscount
Else
    ItemDetail105.Enabled = True
    CHK103.Enabled = True
    CHK103.Value = 0
    ItemDetail105.Text = 0
End If
ItemDetail101.SetFocus


End Sub

Private Sub CHK103_Click()
On Error Resume Next
ItemDetail105.SetFocus
End Sub

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

vCheckJob1 = 1
vQuery = "execute USP_PM_RequestNewDocNo"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vNEwDocno = Trim(vRecordset.Fields("newdocno").Value)
End If
vRecordset.Close
Text101.Text = vNEwDocno

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD102_Click()
MDIForm1.Enabled = False
vMemCommand = 1
FormSearchMainPromotion.Show
End Sub

Private Sub CMD103_Click()
MDIForm1.Enabled = False
vMemCommand = 1
FormSearchSecMan.Show
End Sub

Private Sub CMD104_Click()
MDIForm1.Enabled = False
FormSearchItem.Show
FormSearchItem.Text101.Text = Trim(Form201.ItemDetail101.Text)
End Sub

Private Sub CMD105_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItemCode As ListItem
Dim vDiscount As Currency
Dim vTypeDiscount As Integer
Dim vCountText As String
Dim vPromoPrice As Currency
Dim vIsMember As Integer
Dim vMydescription As String
Dim vUnitCode As String
Dim vDocno As String
Dim vIsCancel As Integer
Dim vIsConfirm As Integer
Dim vIsBrochure As Integer
Dim vPromotionType As String
Dim vPromotionTypeCode As String
Dim vCheckDocno As String
Dim vCheckPMCode As String
Dim vCheckItemDuplicate As Integer
Dim vCheckItemInRequest As Integer
Dim vCheckDuplicatePromotion As Integer
Dim vCheckDuplicateDocno As String
Dim vDiscount1 As Currency

On Error GoTo ErrDescription

    If Text101.Text <> "" And Text102.Text <> "" Then
        If ItemDetail107.Text <> "" Then
        vDocno = Trim(Text101.Text)
        vQuery = "select  *  from npmaster.dbo.tb_pm_request  where docno = '" & vDocno & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vIsCancel = Trim(vRecordset.Fields("iscancel").Value)
            vIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
        Else
            vIsCancel = 0
            vIsConfirm = 0
        End If
        vRecordset.Close
        Else
        MsgBox "กรุณาเลือก ประเภทสินค้าโปรโมชั่นด้วยครับ"
        CMD111.SetFocus
        Exit Sub
        End If
        If vIsCancel <> 1 And vIsConfirm <> 2 Then
        If vCheckRequestOnListview = 0 Then
            '----------------------------------------------------------------------------
            vCheckPMCode = Left(Trim(Text102.Text), InStr((Text102.Text), "/") - 1)
            vQuery = "select itemcode from npmaster.dbo.TB_PM_TempCheckItemDuplicateLine where itemcode = '" & Trim(ItemDetail101.Text) & "' and docno = '" & vDocno & "' and pmcode = '" & vCheckPMCode & "' "
            If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckItemDuplicate = 1 '******************************************************
            Else
                vCheckItemDuplicate = 0
            End If
            vRecordset.Close
            
            vQuery = "select itemcode from npmaster.dbo.TB_PM_requestsub  where itemcode = '" & Trim(ItemDetail101.Text) & "' and docno = '" & vDocno & "' "
            If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckItemInRequest = 1 '*******************************************************
            Else
                vCheckItemInRequest = 0
            End If
            vRecordset.Close
            
            vQuery = "exec USP_PM_ItemDuplicate  '" & Left(Trim(Text102.Text), InStr(Trim(Text102.Text), "/") - 1) & "', '" & Trim(ItemDetail101.Text) & "'  "
            If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckDuplicatePromotion = Trim(vRecordset.Fields("isduplicate").Value)
                vCheckDuplicateDocno = Trim(vRecordset.Fields("duplicate").Value)
            End If
            vRecordset.Close
            If vCheckDuplicatePromotion = 0 Then
        '--------------------------------------------------------------------------------
        If vCheckItemDuplicate = 0 And vCheckItemInRequest = 0 Then
        
            If CHK102.Value = 1 Then
            vIsMember = 1
            vTypeDiscount = 2
            vDiscount = (ItemDetail103.Text - ItemDetail105.Text)
            vCountText = Trim("Member")
            vPromoPrice = ItemDetail105.Text
            If Check101.Value = 1 Then
                vIsBrochure = 1
            Else
                vIsBrochure = 0
            End If
            Else
            If CHK103.Value = 1 Then
            vTypeDiscount = 1
            vDiscount = ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
            vCountText = Trim(ItemDetail105.Text & "%")
            vPromoPrice = Trim(ItemDetail103.Text) - ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
            Else
            vTypeDiscount = 0
            If Trim(ItemDetail105.Text) <> "" Then
                vDiscount = Trim(ItemDetail105.Text)
                vPromoPrice = Trim(ItemDetail103.Text) - Trim(ItemDetail105.Text)
            Else
                vDiscount = 0
                vPromoPrice = Trim(ItemDetail103.Text) - 0
            End If
            vCountText = Trim(ItemDetail105.Text)
            If Check101.Value = 1 Then
                vIsBrochure = 1
            Else
                vIsBrochure = 0
            End If
            End If
            vIsMember = 0
            End If
            vMydescription = Trim(ItemDetail106.Text)
            vUnitCode = Trim(ItemDetail104.Text)
            vPromotionType = Trim(ItemDetail107.Text)
            vPromotionTypeCode = Left(Trim(ItemDetail107.Text), 2)
            
            Set vListItemCode = ListView101.ListItems.Add(, , Trim(ItemDetail101.Text))
            vListItemCode.SubItems(1) = Trim(ItemDetail102.Text)
            vListItemCode.SubItems(2) = Trim(ItemDetail103.Text)
            vListItemCode.SubItems(3) = vPromoPrice
            vListItemCode.SubItems(4) = vDiscount 'vCountText
            vListItemCode.SubItems(5) = vCountText 'vTypeDiscount
            vListItemCode.SubItems(6) = vUnitCode 'vCountText 'vDisCount
            vListItemCode.SubItems(7) = vMydescription 'vUnitCode
            vListItemCode.SubItems(8) = vTypeDiscount 'vIsMember
            vListItemCode.SubItems(9) = vIsMember 'vMydescription
            vListItemCode.SubItems(10) = vIsBrochure
            vListItemCode.SubItems(11) = vIsCancel
            vListItemCode.SubItems(12) = vPromotionType
            vListItemCode.SubItems(13) = vPromotionTypeCode
            vListItemCode.Checked = True
'            ---------------------------------------------------------------------------------------------
            vCheckDocno = Trim(Text101.Text)
            vCheckPMCode = Left(Trim(Text102.Text), InStr((Text102.Text), "/") - 1)
            vQuery = "execute USP_PM_InsertCheckDuplicatItemLine '" & Trim(ItemDetail101.Text) & "','" & vCheckDocno & "','" & vCheckPMCode & "','" & vUserID & "' "
            gConnection.Execute vQuery
            Else
                MsgBox "ไม่สามารถเพิ่มสินค้าเดียวกัน ในเอกสารใบนี้ได้"
                ItemDetail101.Text = ""
                ItemDetail102.Text = ""
                ItemDetail103.Text = ""
                ItemDetail104.Text = ""
                ItemDetail105.Text = ""
                ItemDetail106.Text = ""
                ItemDetail107.Text = ""
                CHK102.Value = 0
                CHK103.Value = 0
                Check101.Value = 0
                Exit Sub
            End If
            '----------------------------------------------------------------------------------------------------
            Else
            MsgBox "สินค้า รหัส " & Trim(ItemDetail101.Text) & " ซ้ำกันในโปรโมชั่นนี้ กับเลขที่เสนอสินค้า " & vCheckDuplicateDocno & "  กรุณาตรวจสอบ"
            End If
    Else

        If UCase(vUserID) = UCase(Trim(Text103.Text)) Then
            ItemDetail101.Enabled = False
                If CHK102.Value = 1 Then
                vIsMember = 1
                    vTypeDiscount = 2
                    vDiscount = ItemDetail103.Text - ItemDetail105.Text
                    vCountText = Trim("Member")
                    vPromoPrice = ItemDetail105.Text
                    If Check101.Value = 1 Then
                        vIsBrochure = 1
                    Else
                        vIsBrochure = 0
                    End If
                    vPromotionType = Trim(ItemDetail107.Text)
                    vPromotionTypeCode = Left(Trim(ItemDetail107.Text), 2)
            Else
                If CHK103.Value = 1 Then
                    vTypeDiscount = 1
                    vDiscount = ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
                    vCountText = Trim(ItemDetail105.Text & "%")
                    vPromoPrice = Trim(ItemDetail103.Text) - ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
                Else
                    vTypeDiscount = 0
                    If Trim(ItemDetail105.Text) <> "" Then
                        vDiscount = Trim(ItemDetail105.Text)
                        vPromoPrice = Trim(ItemDetail103.Text) - Trim(ItemDetail105.Text)
                    Else
                        vDiscount = 0
                        vPromoPrice = Trim(ItemDetail103.Text) - 0
                    End If
                    vCountText = Trim(ItemDetail105.Text)
    
                End If
                    If Check101.Value = 1 Then
                        vIsBrochure = 1
                    Else
                        vIsBrochure = 0
                    End If
                vPromotionType = Trim(ItemDetail107.Text)
                vPromotionTypeCode = Left(Trim(ItemDetail107.Text), 2)
                vIsMember = 0
                If ListView101.ListItems.Item(vIndexItemUpdate).SubItems(11) = 1 Then
                    vIsCancel = 1
                Else
                    vIsCancel = 0
                End If
            End If
                vMydescription = Trim(ItemDetail106.Text)
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(3) = vPromoPrice
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(4) = vDiscount 'vCountText
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(5) = vCountText 'vTypeDiscount
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(7) = vMydescription 'vCountText 'vDisCount
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(8) = vTypeDiscount 'vIsMember
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(9) = vIsMember 'vMydescription
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(10) = vIsBrochure
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(11) = vIsCancel
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(12) = vPromotionType
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(13) = vPromotionTypeCode
                ListView101.ListItems.Item(vIndexItemUpdate).Checked = True
            Else
                MsgBox "ไม่สามารถเพิ่มหรือแก้ไขสินค้าได้ เนื่องจาก User กับ Section Manager ไม่ตรงกัน"
            End If
    End If
        On Error Resume Next
        ItemDetail101.Text = ""
        ItemDetail102.Text = ""
        ItemDetail103.Text = ""
        ItemDetail104.Text = ""
        ItemDetail105.Text = ""
        ItemDetail106.Text = ""
        ItemDetail107.Text = ""
        ItemDetail105.Enabled = True
        ItemDetail101.Enabled = True
        Check101.Value = 0
        CHK103.Enabled = True
        ItemDetail101.SetFocus
        CHK103.Value = 0
        CHK102.Value = 0
        vCheckRequestOnListview = 0
    Else
        If vIsCancel = 1 Then
            MsgBox "เอกสารเลขที่ " & vDocno & " ได้ยกเลิกไปแล้ว ไม่สามารถแก้ไขข้อมูลได้"
        ElseIf vIsConfirm = 2 Then
            MsgBox "เอกสารเลขที่ " & vDocno & " ได้อนุมัติไปแล้ว ไม่สามารถแก้ไขข้อมูลได้"
        End If
        On Error Resume Next
        ItemDetail101.Text = ""
        ItemDetail102.Text = ""
        ItemDetail103.Text = ""
        ItemDetail104.Text = ""
        ItemDetail105.Text = ""
        ItemDetail106.Text = ""
        ItemDetail107.Text = ""
        ItemDetail105.Enabled = True
        ItemDetail101.Enabled = True
        Check101.Value = 0
        CHK103.Enabled = True
        ItemDetail101.SetFocus
        CHK103.Value = 0
        CHK102.Value = 0
    End If
Else
    MsgBox "การเพิ่มสินค้าใน ตารางข้างล่าง ต้องใส่เลขที่เอกสารและรหัสโปรโมชั่นก่อนนะครับ"
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD106_Click()
MDIForm1.Enabled = False
ItemDetail101.Enabled = True
FormSearchReqPromo.Show
End Sub

Private Sub CMD107_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vPromoname As String
Dim vStartPromo As Date
Dim vIsCancel As String
Dim vCountItem As Integer
Dim vSecName As String
Dim vPromotionCode As String
Dim i As Integer
Dim vError As Integer
Dim vIsCompleteSave As Integer
Dim vItemCode As String
Dim vItemName As String
Dim vUnitCode As String
Dim vNormalPrice As Currency
Dim vFromQTY As Long, vToQty As Long
Dim vDiscount As Currency
Dim vDiscountWord As String
Dim vDiscountType As String
Dim vPromotionPrice As Currency
Dim vMydescription As String
Dim vLineNumber As Integer
Dim vIsBrochure As String
Dim vIsMember As String
Dim vIsConfirm As Integer
Dim vPromotionType As String
Dim vPromotionTypeCode As String
Dim vItemIsCancel As Integer
Dim vCheckDeleteDocno As String
Dim vCheckDuplicatePromotion As Integer
Dim vCheckDuplicateDocno As String

On Error GoTo ErrDescription

If Trim(Text101.Text) <> "" And ListView101.ListItems.Count <> 0 Then
vCountItem = ListView101.ListItems.Count
If vCheckJob1 = 1 Then
    If vCountItem > 0 Then
        If Text101.Text <> "" Then
            If Text102.Text <> "" Or Text103.Text <> "" Then
                vStartPromo = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
                vSecName = Trim(Text103.Text)
                vPromotionCode = Left(Trim(Text102.Text), InStr(Trim(Text102.Text), "/") - 1)
                If vCheckJob1 <> 0 Then
                    vQuery = "execute USP_PM_RequestNewDocNo"
                    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                        vNEwDocno = Trim(vRecordset.Fields("newdocno").Value)
                    End If
                    vRecordset.Close
                Else
                    vNEwDocno = Trim(Text101.Text)
                End If
                
                vQuery = "execute USP_PM_InsertRequest " & vCheckJob1 & ",'" & vNEwDocno & "','" & vStartPromo & "','" & vSecName & "','" & vPromotionCode & "','" & vUserID & "' "
                gConnection.Execute vQuery
                
                For i = 1 To ListView101.ListItems.Count
                vError = 0
                If i = ListView101.ListItems.Count Then
                    vIsCompleteSave = 1
                Else
                    vIsCompleteSave = 0
                End If
                vItemCode = Trim(ListView101.ListItems.Item(i).Text)
                vIsCompleteSave = 1
                vError = 0
                vItemName = Trim(ListView101.ListItems.Item(i).SubItems(1))
                vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(6))
                vNormalPrice = Trim(ListView101.ListItems.Item(i).SubItems(2))
                vFromQTY = 1
                vToQty = 99999
                If Trim(ListView101.ListItems.Item(i).SubItems(8)) <> 2 Then
                    vDiscount = Trim(ListView101.ListItems.Item(i).SubItems(4))
                Else
                    vDiscount = 0
                End If
                vDiscountType = Trim(ListView101.ListItems.Item(i).SubItems(8))
                vDiscountWord = Trim(ListView101.ListItems.Item(i).SubItems(5))
                vPromotionPrice = Trim(ListView101.ListItems.Item(i).SubItems(3))
                vMydescription = Trim(ListView101.ListItems.Item(i).SubItems(7))
                vLineNumber = i - 1
                vIsBrochure = Trim(ListView101.ListItems.Item(i).SubItems(10))
                If Trim(ListView101.ListItems.Item(i).SubItems(9)) = 0 Then
                    vIsMember = 0
                Else
                    vIsMember = 1
                End If
                vPromotionType = Trim(ListView101.ListItems.Item(i).SubItems(12))
                vPromotionTypeCode = Trim(ListView101.ListItems.Item(i).SubItems(13))
                If ListView101.ListItems.Item(i).Checked = True Then
                    vItemIsCancel = 0
                Else
                    vItemIsCancel = 0
                End If
                
                vQuery = "exec USP_PM_ItemDuplicate  '" & Trim(vPromotionCode) & "', '" & Trim(vItemCode) & "'  "
                If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                    vCheckDuplicatePromotion = Trim(vRecordset.Fields("isduplicate").Value)
                    vCheckDuplicateDocno = Trim(vRecordset.Fields("duplicate").Value)
                End If
                vRecordset.Close
                If vCheckDuplicatePromotion = 0 Then
                    vQuery = "execute USP_PM_InsertRequestSub " & vError & "," & vIsCompleteSave & ",'" & vNEwDocno & "','" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vNormalPrice & "," & vFromQTY & "," & vToQty & "," & vDiscount & ",'" & vDiscountType & "','" & vDiscountWord & "'," & vPromotionPrice & ",'" & vMydescription & "','" & vItemIsCancel & "'," & vLineNumber & ",'" & vIsBrochure & "','" & vIsMember & "' ,'" & vPromotionTypeCode & "' "
                    gConnection.Execute vQuery
                Else
                    vQuery = "execute USP_PM_InsertRequestSub 1," & vIsCompleteSave & ",'" & vNEwDocno & "','" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vNormalPrice & "," & vFromQTY & "," & vToQty & "," & vDiscount & ",'" & vDiscountType & "','" & vDiscountWord & "'," & vPromotionPrice & ",'" & vMydescription & "','" & vItemIsCancel & "'," & vLineNumber & ",'" & vIsBrochure & "','" & vIsMember & "' ,'" & vPromotionTypeCode & "' "
                    gConnection.Execute vQuery
                    MsgBox "รายการสินค้า รหัส " & vItemCode & " ในโปรโมชั่นนี้มีอยู่แล้ว ในเอกสารเลขที่ " & vCheckDuplicateDocno & " กรุณาตรวจสอบ", vbCritical, "Send Error"
                    Exit Sub
                End If
                
                Next i
                MsgBox "ได้เอกสารเลขที่  " & vNEwDocno & " "
                Call InitializeSendEmail
                vQuery = "execute USP_PM_DeliverySendMail '" & vNEwDocno & "' "
                vGetConnect.Execute vQuery
                
                vCheckDeleteDocno = Trim(Text101.Text)
                vQuery = "USP_PM_DeleteCheckDuplicatItemLine '" & vCheckDeleteDocno & "','" & vPromotionCode & "','" & vUserID & "' "
                gConnection.Execute vQuery
                
            Else
                MsgBox "กรุณาเลือก ประเภทโปรโมชั่น และ Section Manager ด้วยครับ"
                Exit Sub
            End If
        Else
            MsgBox "กรุณา กดปุ่มสร้างเลขที่เอกสารด้วยครับ"
            Exit Sub
        End If
    Else
        MsgBox "ไม่มีสินค้าที่จะทำการเสนอทำโปรโมชั่น"
        Exit Sub
    End If
ElseIf vCheckJob1 = 0 Then
    vCheckJob1 = 0
    vNEwDocno = Trim(Text101.Text)
    vStartPromo = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
    vSecName = Trim(Text103.Text)
    vPromotionCode = Left(Trim(Text102.Text), InStr(Trim(Text102.Text), "/") - 1)
    
    vQuery = "select  *  from npmaster.dbo.tb_pm_request where docno = '" & vNEwDocno & "' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vIsCancel = Trim(vRecordset.Fields("iscancel").Value)
        vIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
    End If
    vRecordset.Close
    If vIsCancel = 0 And vIsConfirm = 0 Then
        vQuery = "execute USP_PM_InsertRequest " & vCheckJob1 & ",'" & vNEwDocno & "','" & vStartPromo & "','" & vSecName & "','" & vPromotionCode & "','" & vUserID & "' "
        gConnection.Execute vQuery
        If ListView101.ListItems.Count <> 0 Then
        For i = 1 To ListView101.ListItems.Count
        vError = 0
        If i = ListView101.ListItems.Count Then
            vIsCompleteSave = 1
        Else
            vIsCompleteSave = 0
        End If
        vItemCode = Trim(ListView101.ListItems.Item(i).Text)
        vIsCompleteSave = 1
        vError = 0
        vItemName = Trim(ListView101.ListItems.Item(i).SubItems(1))
        vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(6))
        vNormalPrice = Trim(ListView101.ListItems.Item(i).SubItems(2))
        vFromQTY = 1
        vToQty = 99999
        vDiscount = Trim(ListView101.ListItems.Item(i).SubItems(4))
        vDiscountType = Trim(ListView101.ListItems.Item(i).SubItems(8))
        vDiscountWord = Trim(ListView101.ListItems.Item(i).SubItems(5))
        vPromotionPrice = Trim(ListView101.ListItems.Item(i).SubItems(3))
        vMydescription = Trim(ListView101.ListItems.Item(i).SubItems(7))
        vLineNumber = i - 1
        If Trim(ListView101.ListItems.Item(i).SubItems(10)) = 0 Then
        vIsBrochure = 0
        Else
        vIsBrochure = 1
        End If
        vPromotionType = Trim(ListView101.ListItems.Item(i).SubItems(12))
        vPromotionTypeCode = Trim(ListView101.ListItems.Item(i).SubItems(13))
        If Trim(ListView101.ListItems.Item(i).SubItems(9)) = 0 Then
            vIsMember = 0
        Else
            vIsMember = 1
        End If
        If Trim(ListView101.ListItems.Item(i).SubItems(11)) = 0 Then
            vItemIsCancel = 0
        Else
            vItemIsCancel = 1
        End If
        
        vQuery = "execute USP_PM_InsertRequestSub " & vError & "," & vIsCompleteSave & ",'" & vNEwDocno & "','" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vNormalPrice & "," & vFromQTY & "," & vToQty & "," & vDiscount & ",'" & vDiscountType & "','" & vDiscountWord & "'," & vPromotionPrice & ",'" & vMydescription & "','" & vItemIsCancel & "'," & vLineNumber & ",'" & vIsBrochure & "','" & vIsMember & "','" & vPromotionTypeCode & "' "
        gConnection.Execute vQuery
        Next i
        
        'If vCheckDuplicatePromotion = 0 Then
        MsgBox "ปรับปรุงเอกสารเลขที่  " & vNEwDocno & " เรียบร้อยแล้วครับ"
        'Call InitializeSendEmail
        'vQuery = "execute USP_PM_DeliverySendMail '" & vNewDocno & "' "
        'vGetConnect.Execute vQuery
        'End If
        
        vCheckDeleteDocno = Trim(Text101.Text)
        vQuery = "USP_PM_DeleteCheckDuplicatItemLine '" & vCheckDeleteDocno & "','" & vPromotionCode & "','" & vUserID & "' "
        gConnection.Execute vQuery
                
        End If
    End If
End If
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    ItemDetail101.Text = ""
    ItemDetail102.Text = ""
    ItemDetail103.Text = ""
    ItemDetail104.Text = ""
    ItemDetail105.Text = ""
    ItemDetail106.Text = ""
    ItemDetail107.Text = ""
    CHK102.Value = 0
    CHK103.Value = 0
    Check101.Value = 0
    ListView101.ListItems.Clear
    Image101.Visible = True
    Image102.Visible = False
    Image103.Visible = False
Else
    MsgBox "เอกสารเสนอสินค้า ควรมีรายการสินค้าอย่างน้อย 1 รายการถึงจะบันทึกหรืออัพเดทได้"
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    DTPicker101.Value = Now
End If
ErrDescription:
If Err.Description <> "" Then
    If Err.Number = "-2147217873" Then
        MsgBox Err.Description
        MsgBox "รหัสสินค้าที่เสนอโปรโมชั่นซ้ำกัน กรุณาตรวจสอบ"
    Else
        MsgBox Err.Description
    End If
Exit Sub
End If
End Sub

Private Sub CMD108_Click()
Dim i As Integer
Dim vCheckDelete As Integer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocno As String
Dim vCheckPromo As String
Dim vCheckItemCode As String

On Error Resume Next

If ListView101.ListItems.Count <> 0 Then
    i = ListView101.ListItems.Item(vIndexItemUpdate).Index
    vCheckDelete = MsgBox("ต้องการลบรายการสินค้า  " & ListView101.ListItems.Item(vIndexItemUpdate).Text & "  นี้หรือไม่", vbYesNo, "คำยืนยันในการลบรายการสินค้า")
    If vCheckDelete = 6 Then
        vCheckDocno = Trim(Text101.Text)
        vCheckPromo = Left(Trim(Text102.Text), InStr(Trim(Text102.Text), "/") - 1)
        vCheckItemCode = Trim(ListView101.ListItems.Item(vIndexItemUpdate).Text)
        vQuery = "USP_PM_DeleteCheckDuplicatItemLine '" & vCheckDocno & "','" & vCheckPromo & "','" & vUserID & "' ,'" & vCheckItemCode & "' "
        gConnection.Execute vQuery
        
        ItemDetail101.Enabled = True
        ListView101.ListItems.Remove (i)
        ItemDetail101.Text = ""
        ItemDetail102.Text = ""
        ItemDetail103.Text = ""
        ItemDetail104.Text = ""
        ItemDetail105.Text = ""
        ItemDetail106.Text = ""
        ItemDetail107.Text = ""
        CHK102.Value = 0
        CHK103.Value = 0
        Check101.Value = 0
        ItemDetail101.SetFocus
        CMD108.Enabled = False
    Else
        ItemDetail101.Enabled = True
        ItemDetail101.Text = ""
        ItemDetail102.Text = ""
        ItemDetail103.Text = ""
        ItemDetail104.Text = ""
        ItemDetail105.Text = ""
        ItemDetail106.Text = ""
        ItemDetail107.Text = ""
        CHK102.Value = 0
        CHK103.Value = 0
        Check101.Value = 0
        ItemDetail101.SetFocus
        CMD108.Enabled = False
        Exit Sub
    End If
Else
    MsgBox "ไม่มีสินค้าให้ลบรายการ"
End If

End Sub

Private Sub CMD109_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocno As String
Dim vCheckPromo As String

On Error GoTo ErrDescription

If Trim(Text101.Text) <> "" Then
vCheckDocno = Trim(Text101.Text)
vCheckPromo = Left(Trim(Text102.Text), InStr(Trim(Text102.Text), "/") - 1)
vQuery = "USP_PM_DeleteCheckDuplicatItemLine '" & vCheckDocno & "','" & vCheckPromo & "','" & vUserID & "'  "
gConnection.Execute vQuery
End If
        
Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
DTPicker101.Value = Now
ItemDetail101.Text = ""
ItemDetail102.Text = ""
ItemDetail103.Text = ""
ItemDetail104.Text = ""
ItemDetail105.Text = ""
ItemDetail106.Text = ""
ItemDetail107.Text = ""
ItemDetail101.Enabled = True
ListView101.ListItems.Clear
Image101.Visible = True
Image102.Visible = False
Image103.Visible = False
CMD108.Enabled = False
CHK102.Value = 0
CHK103.Value = 0
Check101.Value = 0
vCheckRequestOnListview = 0

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD110_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vDocno As String

On Error GoTo ErrDescription

If vCheckStatusPrint = 0 Then
    vDocno = Trim(Text101.Text)
    If vDocno <> "" Then
    vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 269 and reptype = 'PM' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
    
    With Crystal101
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@vDocno;" & vDocno & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    End With
    End If
Else
    MsgBox "ไม่สามารถพิมพ์เอกสารที่ ตรวจสอบแล้ว หรือ อนุมัติแล้ว หรือ ยกเลิก ได้ กรณีต้องการจะพิมพ์ติดต่อแผนกคอมพิวเตอร์", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD111_Click()
MDIForm1.Enabled = False
FormSearchType.Show
End Sub

Private Sub CMD112_Click()
Dim vQuery  As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno As String
Dim vAnswer As Integer
Dim vCheckDocno As Integer
Dim vCheckUser As String
Dim vIsConfirm As Integer

On Error GoTo ErrDescription

If Text103.Text <> "" Then
    vCheckUser = UCase(Trim(Text103.Text))
    If UCase(vUserID) = vCheckUser Then
        If Trim(Text101.Text) <> "" Then
            vDocno = Trim(Text101.Text)
            vQuery = "select docno,isconfirm from npmaster.dbo.tb_pm_request where docno = '" & vDocno & "' "
            If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckDocno = 1
                vIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
            Else
                vCheckDocno = 0
            End If
            vRecordset.Close
            If vIsConfirm <> 2 Then
            If vCheckDocno = 1 Then
            vAnswer = MsgBox("คุณต้องการลบเอกสารเลขที่ " & vDocno & " นี้ใช่หรือไม่", vbYesNo, "คำถามยืนยันการลบ")
            
            If vAnswer = 6 Then
                If vDocno <> "" Then
                    vQuery = "exec USP_PM_DeletePMRequest  '" & vDocno & "' "
                    gConnection.Execute vQuery
                    MsgBox "ได้ทำการลบเอกสารเลขที่ " & vDocno & " เรียบร้อยแล้ว "
                End If
                Text101.Text = ""
                Text102.Text = ""
                Text103.Text = ""
                DTPicker101.Value = Now
                ListView101.ListItems.Clear
                Else
                Exit Sub
            End If
            Else
                MsgBox "เอกสารเลขที่ " & vDocno & " ยังไม่มี"
            End If
            Else
                MsgBox "เอกสารเลขที่ " & vDocno & " ได้อนุมัติแล้วไม่สามารถลบเอกสารได้"
            End If
        End If
    Else
        MsgBox "คุณไม่มีสิทธิ์ในการลบเอกสารเลขที่นี้"
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

'Private Sub CMD113_Click()
'Dim i As Integer
'Dim vCheckIsCancel As Integer
'
'On Error Resume Next

'If ListView101.ListItems.Count <> 0 Then
 '   i = ListView101.ListItems.Item(vIndexItemUpdate).Index
  '  vCheckIsCancel = MsgBox("ต้องการยกเลิกรายการสินค้า  " & ListView101.ListItems.Item(vIndexItemUpdate).Text & "  นี้หรือไม่", vbYesNo, "คำยืนยันในการลบรายการสินค้า")
   ' If vCheckIsCancel = 6 Then
    '    ItemDetail101.Enabled = True
     '   ItemDetail101.SetFocus
      '  ListView101.ListItems.Item(i).SubItems(11) = 1
       ' ItemDetail101.Enabled = True
        'ItemDetail101.Text = ""
        'ItemDetail102.Text = ""
        'itemDetail103.Text = ""
        'ItemDetail104.Text = ""
        'ItemDetail105.Text = ""
        'ItemDetail106.Text = ""
        'ItemDetail107.Text = ""
        'CHK102.Value = 0
        'CHK103.Value = 0
        'Check101.Value = 0
        'ItemDetail101.SetFocus
        'CMD108.Enabled = False
    'Else
     '   Exit Sub
    'End If
'Else
 '   MsgBox "ไม่มีสินค้าให้ยกเลิกรายการ"
'End If

'End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vMemberDisc As String

DTPicker101 = Now
vCheckJob1 = 1
vCheckRequestOnListview = 0
CHK102.Caption = "สมาชิก"
vQuery = "select memberdisc  from bpsconfig"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vMemberDisc = Left(Trim(vRecordset.Fields("memberdisc").Value), InStr(Trim(vRecordset.Fields("memberdisc").Value), "%") - 1)
End If
vRecordset.Close
CHK102.Caption = CHK102.Caption & "    " & "( ลด " & vMemberDisc & "% )"
vMemberDiscount = vMemberDisc

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

vQuery = "delete npmaster.dbo. TB_PM_TempCheckItemDuplicateLine where userid = '" & vUserID & "'  "
gConnection.Execute vQuery

End Sub

Private Sub ItemDetail101_GotFocus()
If Text102.Text = "" Or Text103.Text = "" Then
    MsgBox "กรุณาเลือก รายการโปรโมชั่น และ ชื่อ Section Manager ด้วยครับ"
    Exit Sub
End If
End Sub

Private Sub ItemDetail101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String

On Error Resume Next

If KeyAscii = 13 Then
    vSearch = Trim(ItemDetail101.Text)
    vCheckRequestOnListview = 0
    If vSearch <> "" Then
        vQuery = "execute USP_PM_FindItemFix '" & vSearch & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            ItemDetail102.Text = Trim(vRecordset.Fields("itemname").Value)
            ItemDetail104.Text = Trim(vRecordset.Fields("unitcode").Value)
            ItemDetail103.Text = Trim(vRecordset.Fields("saleprice1").Value)
        Else
            MDIForm1.Enabled = False
            FormSearchItem.Text101.Text = vSearch
            FormSearchItem.Show
        End If
        vRecordset.Close
    End If
    If CHK102.Value = 1 Then
        ItemDetail105.Text = (Trim(ItemDetail103.Text) - (Trim(ItemDetail103.Text) * vMemberDiscount) / 100)
        ItemDetail106.SetFocus
    Else
        ItemDetail105.SetFocus
    End If
End If
End Sub

Private Sub ItemDetail101_LostFocus()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String

On Error GoTo ErrDescription

    vSearch = Trim(ItemDetail101.Text)
    If vSearch <> "" Then
        vQuery = "execute USP_PM_FindItemFix '" & vSearch & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            ItemDetail102.Text = Trim(vRecordset.Fields("itemname").Value)
            ItemDetail104.Text = Trim(vRecordset.Fields("unitcode").Value)
            ItemDetail103.Text = Trim(vRecordset.Fields("saleprice1").Value)
        Else
            MDIForm1.Enabled = False
            FormSearchItem.Show
            FormSearchItem.Text101.Text = vSearch
        End If
        vRecordset.Close
    End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ItemDetail105_KeyPress(KeyAscii As Integer)
Dim vItemCount As Integer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDisCountLine As Integer
Dim vCheckItemDisCount As String
Dim vUnitCode As String

On Error Resume Next

If KeyAscii = 13 Then
    If CHK103.Value = 1 Then
        vItemCount = Trim(ItemDetail103.Text) - ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
    ElseIf CHK103.Value = 0 Then
        If Trim(ItemDetail105.Text) = "" Then
            vItemCount = Trim(ItemDetail103.Text)
        Else
            vItemCount = Trim(ItemDetail103.Text) - Trim(ItemDetail105.Text)
        End If
    End If
    
    vCheckItemDisCount = Trim(ItemDetail101.Text)
    vUnitCode = Trim(ItemDetail104.Text)
    vQuery = "exec USP_PM_LowerCost  '" & vCheckItemDisCount & "'," & vItemCount & ",'" & vUnitCode & "' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vDisCountLine = Trim(vRecordset.Fields("lowercost"))
    End If
    vRecordset.Close
    
    If vItemCount < 0 Then
    MsgBox "ลดราคามากกว่า ราคาปกติ"
    ItemDetail105.SetFocus
    Else
    CMD111.SetFocus
    End If
    
    If vDisCountLine = 1 Then
        MsgBox "สินค้า รหัส " & vCheckItemDisCount & " เสนอราคาโปรโมชั่นขาดทุน "
    End If
End If
End Sub

Private Sub ItemDetail105_LostFocus()
Dim vItemCount As Integer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDisCountLine As Integer
Dim vCheckItemDisCount As String
Dim vUnitCode As String

On Error Resume Next

    If CHK103.Value = 1 Then
        vItemCount = Trim(ItemDetail103.Text) - ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
    ElseIf CHK103.Value = 0 Then
        If Trim(ItemDetail105.Text) = "" Then
            vItemCount = Trim(ItemDetail103.Text)
        Else
            vItemCount = Trim(ItemDetail103.Text) - Trim(ItemDetail105.Text)
        End If
    End If
    
    vCheckItemDisCount = Trim(ItemDetail101.Text)
    vUnitCode = Trim(ItemDetail104.Text)
    vQuery = "exec USP_PM_LowerCost  '" & vCheckItemDisCount & "'," & vItemCount & " ,'" & vUnitCode & "' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vDisCountLine = Trim(vRecordset.Fields("lowercost"))
    End If
    vRecordset.Close
    
    If vItemCount < 0 Then
    MsgBox "ลดราคามากกว่า ราคาปกติ"
    ItemDetail105.SetFocus
    Else
    If vDisCountLine = 1 Then
        MsgBox "สินค้า รหัส " & vCheckItemDisCount & " เสนอราคาโปรโมชั่นขาดทุน "
    End If
    
    ItemDetail106.SetFocus
    End If

End Sub

Private Sub ItemDetail106_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    CMD105.SetFocus
End If
End Sub

Private Sub ListView101_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo ErrDescription

ListView101.Sorted = True
ListView101.SortKey = ColumnHeader.Index - 1
If vSortResult = 0 Then
    ListView101.SortOrder = lvwAscending
    vSortResult = 1
Else
    ListView101.SortOrder = lvwDescending
    vSortResult = 0
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListView101_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
i = Item.Index
'"&H80000008&"
    If ListView101.ListItems.Item(Item.Index).Checked = False Then
        ListView101.ListItems(i).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(10).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(11).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(12).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(13).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).SubItems(11) = "1"
    Else
        ListView101.ListItems(i).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(10).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(11).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(12).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(13).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).SubItems(11) = "0"
    End If

End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckIsConfirmDoc  As Integer
Dim vCheckIsCancelDoc  As Integer
Dim vDocnoCheck As String

On Error Resume Next

If ListView101.ListItems.Item(Item.Index).Checked = True Then
If ListView101.ListItems.Count <> 0 And Trim(Text101.Text) <> "" Then
vDocnoCheck = Trim(Text101.Text)
vQuery = "select isconfirm ,iscancel from npmaster.dbo.tb_pm_request where docno = '" & vDocnoCheck & "' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckIsConfirmDoc = Trim(vRecordset.Fields("isconfirm"))
    vCheckIsCancelDoc = Trim(vRecordset.Fields("iscancel"))
End If
vRecordset.Close

If vCheckIsConfirmDoc <> 2 Or vCheckIsCancelDoc = 1 Then
    vCheckRequestOnListview = 1
    CMD108.Enabled = True
    ItemDetail101.Enabled = False
    vIndexItemUpdate = Item.Index
    If Trim(ListView101.ListItems.Item(Item.Index).SubItems(9)) <> 0 Then
        CHK102.Value = 1
    Else
        CHK102.Value = 0
    End If
    
    If Trim(ListView101.ListItems.Item(Item.Index).SubItems(10)) <> 0 Then
        Check101.Value = 1
    Else
        Check101.Value = 0
    End If
    
    ItemDetail101.Text = Trim(ListView101.ListItems.Item(Item.Index).Text)
    ItemDetail102.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(1))
    ItemDetail103.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(2))
    ItemDetail104.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(6))
    If Trim(ListView101.ListItems.Item(Item.Index).SubItems(8)) = 1 Then
        CHK103.Value = 1
        ItemDetail105.Text = Left(Trim(ListView101.ListItems.Item(Item.Index).SubItems(5)), Len(Trim(ListView101.ListItems.Item(Item.Index).SubItems(5))) - 1)
    ElseIf Trim(ListView101.ListItems.Item(Item.Index).SubItems(8)) = 2 Then
        ItemDetail105.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(3))
        ItemDetail105.Enabled = False
    Else
        ItemDetail105.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(5))
        CHK103.Value = 0
    End If
    ItemDetail106.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(7))
    ItemDetail107.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(12))
End If
End If
Else
    ItemDetail101.Text = ""
    ItemDetail102.Text = ""
    ItemDetail103.Text = ""
    ItemDetail104.Text = ""
    ItemDetail105.Text = ""
    ItemDetail106.Text = ""
    ItemDetail107.Text = ""
    CHK102.Value = 0
    CHK103.Value = 0
    Check101.Value = 0
End If
End Sub

