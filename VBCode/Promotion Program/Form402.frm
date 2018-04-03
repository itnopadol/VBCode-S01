VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form102 
   Caption         =   "เพิ่มและพิมพ์คูปอง"
   ClientHeight    =   8880
   ClientLeft      =   7020
   ClientTop       =   1065
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   14385
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "พิมพ์คูปอง"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   720
      TabIndex        =   18
      Top             =   8280
      Visible         =   0   'False
      Width           =   10545
      Begin VB.TextBox Text104 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   7245
         TabIndex        =   27
         Top             =   2205
         Width           =   3120
      End
      Begin VB.CommandButton CMD104 
         Caption         =   "พิมพ์คูปอง"
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
         Left            =   9270
         TabIndex        =   25
         Top             =   2610
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   1815
         Left            =   2385
         TabIndex        =   24
         Top             =   270
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   3201
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสคูปอง"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อฟอร์มคูปอง"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ที่อยู่คูปอง"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "ชื่อคูปองที่พิมพ์ :"
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
         Left            =   5355
         TabIndex        =   26
         Top             =   2205
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "รายการคูปอง :"
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
         TabIndex        =   23
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "เพิ่มทะเบียนคูปอง"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8070
      Left            =   90
      TabIndex        =   17
      Top             =   90
      Width           =   14190
      Begin VB.PictureBox Crystal101 
         Height          =   480
         Left            =   12600
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   45
         Top             =   7200
         Width           =   1200
      End
      Begin VB.CommandButton CMDApprove 
         BackColor       =   &H00808080&
         Caption         =   "อนุมัติ"
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
         Height          =   420
         Left            =   5490
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   7380
         Width           =   1095
      End
      Begin VB.CommandButton CMDPrintCoupon 
         BackColor       =   &H00808080&
         Caption         =   "พิมพ์คูปอง"
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
         Height          =   420
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   7380
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListViewMemCoupon 
         Height          =   6765
         Left            =   6300
         TabIndex        =   40
         Top             =   495
         Visible         =   0   'False
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   11933
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "เลขที่คูปอง"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "มูลค่า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ทะเบียนคูปอง"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.TextBox TBStopNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   6705
         TabIndex        =   8
         Top             =   3240
         Width           =   1140
      End
      Begin VB.TextBox TBStartNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   4365
         TabIndex        =   7
         Top             =   3240
         Width           =   1140
      End
      Begin VB.ComboBox CMBPosition 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3735
         Width           =   825
      End
      Begin VB.CheckBox CBDash 
         Caption         =   "ตัวคั่น"
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
         Left            =   3240
         TabIndex        =   10
         Top             =   3735
         Width           =   960
      End
      Begin VB.CommandButton CMDViewCoupon 
         BackColor       =   &H00808080&
         Caption         =   "ตรวจสอบ"
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
         Left            =   4230
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7380
         Width           =   1095
      End
      Begin VB.CommandButton CMDAddCoupon 
         BackColor       =   &H00808080&
         Caption         =   "เพิ่มรายการ"
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
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4185
         Width           =   1500
      End
      Begin VB.TextBox TBAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   1710
         TabIndex        =   6
         Top             =   3240
         Width           =   1500
      End
      Begin MSComctlLib.ListView ListViewCoupon 
         Height          =   2265
         Left            =   1710
         TabIndex        =   12
         Top             =   4995
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3995
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648384
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "หัวคูปอง"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "วันที่เริ่ม"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หมดอายุ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "มูลค่า"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "จากเลขที่"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ถึงเลขที่"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "จำนวนคูปอง"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ตัวคั่น"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Position"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Insert"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "เลขสุดท้าย"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.CheckBox Check101 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ยกเลิก คูปอง"
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
         Height          =   420
         Left            =   225
         TabIndex        =   16
         Top             =   7380
         Width           =   1365
      End
      Begin VB.CommandButton CMD103 
         BackColor       =   &H00808080&
         Caption         =   "ล้างหน้าจอ"
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
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7380
         Width           =   1095
      End
      Begin VB.TextBox Text103 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   1710
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2160
         Width           =   10095
      End
      Begin VB.CommandButton CMD101 
         Height          =   330
         Left            =   5670
         Picture         =   "Form402.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   495
         Width           =   330
      End
      Begin VB.ListBox List101 
         Appearance      =   0  'Flat
         Height          =   1980
         Left            =   6885
         TabIndex        =   28
         Top             =   135
         Visible         =   0   'False
         Width           =   3930
      End
      Begin MSComCtl2.DTPicker DTPicker102 
         Height          =   330
         Left            =   1710
         TabIndex        =   4
         Top             =   1755
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   38709
      End
      Begin MSComCtl2.DTPicker DTPicker101 
         Height          =   330
         Left            =   1710
         TabIndex        =   3
         Top             =   1305
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   38709
      End
      Begin VB.TextBox Text102 
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
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   2
         Top             =   900
         Width           =   1500
      End
      Begin VB.TextBox Text101 
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
         Left            =   1710
         TabIndex        =   0
         Top             =   495
         Width           =   3930
      End
      Begin VB.CommandButton CMD102 
         BackColor       =   &H00808080&
         Caption         =   "บันทึกข้อมูล"
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
         Height          =   420
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7380
         Width           =   1095
      End
      Begin VB.Label LBLIndex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4095
         TabIndex        =   44
         Top             =   4230
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label LBLID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3330
         TabIndex        =   43
         Top             =   900
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label11 
         Caption         =   "ใบ"
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
         Left            =   10890
         TabIndex        =   39
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "ตัวอย่างคูปอง :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5265
         TabIndex        =   38
         Top             =   3735
         Width           =   1320
      End
      Begin VB.Label LBLCoupon 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6705
         TabIndex        =   37
         Top             =   3735
         Width           =   4110
      End
      Begin VB.Label LBLCountCoupon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   9720
         TabIndex        =   36
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "ถึงเลข :"
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
         Left            =   5580
         TabIndex        =   35
         Top             =   3240
         Width           =   1005
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "เริ่มจาก :"
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
         Left            =   3420
         TabIndex        =   34
         Top             =   3240
         Width           =   825
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "ตำแหน่ง :"
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
         Left            =   720
         TabIndex        =   33
         Top             =   3735
         Width           =   870
      End
      Begin VB.Label Label10 
         Caption         =   "รายการคูปอง"
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
         Left            =   1710
         TabIndex        =   32
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "จำนวนคูปอง :"
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
         Left            =   8460
         TabIndex        =   31
         Top             =   3240
         Width           =   1140
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "มูลค่าคูปอง :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   405
         TabIndex        =   30
         Top             =   3240
         Width           =   1185
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "คำอธิบายคูปอง :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   29
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "วันสิ้นสุดคูปอง :"
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
         Left            =   225
         TabIndex        =   22
         Top             =   1755
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "วันเริ่มใช้คูปอง :"
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
         Left            =   90
         TabIndex        =   21
         Top             =   1305
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "หัวเลขที่คูปอง :"
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
         Left            =   90
         TabIndex        =   20
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ชื่อคูปอง :"
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
         Left            =   585
         TabIndex        =   19
         Top             =   495
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Form102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vMemCouponNo As String
Dim vMemCountUse As Integer
Dim vMemIsConfirm As Integer
Dim vMemIsCancel As Integer

Private Sub CBDash_Click()
If Me.Text102.Text <> "" Then
Call GenCoupon
Else
Me.CBDash.Value = 0
End If

Me.CMDAddCoupon.SetFocus
End Sub

Private Sub CBDash_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBAmount.Text = ""
Me.TBStartNum.Text = ""
Me.TBStopNum.Text = ""
Me.LBLCountCoupon.Caption = ""
Me.TBAmount.SetFocus
End If
End Sub

Private Sub CMBPosition_Click()
Call GenCoupon
End Sub

Private Sub CMBPosition_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBAmount.Text = ""
Me.TBStartNum.Text = ""
Me.TBStopNum.Text = ""
Me.LBLCountCoupon.Caption = ""
Me.TBAmount.SetFocus
End If
End Sub

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer

List101.Clear

vQuery = "select * from npmaster.dbo.tb_pm_couponmaster where iscancel = 0 order by id "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        List101.AddItem (Trim(vRecordset.Fields("couponname").Value))
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close
If Me.List101.ListCount > 0 Then
List101.Visible = True
Else
MsgBox "ไม่มีทะเบียนคูปองในระบบ", vbCritical, "Send Error Message"
Me.Text101.SetFocus
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As Integer
Dim vHeaderNo As String
Dim vCouponName As String
Dim vStartDate As Date
Dim vEndDate As Date
Dim vMydescription As String
Dim vCheckHeader As Integer
Dim vIsCancel As String
Dim vPosition As Integer
Dim vIsDash As Integer

Dim i As Integer
Dim vCouponCode As String

Dim n As Integer
Dim m As Integer
Dim vCouponAmount As Double
Dim vStartNum As Integer
Dim vStopNum As Integer
Dim vCountCoupon As Integer

If Me.Text101.Text = "" Then
MsgBox "กรุณากรอกชื่อคูปอง", vbCritical, "Send Error Message"
Me.Text101.SetFocus
Exit Sub
End If

If Me.Text102.Text = "" Then
MsgBox "กรุณากรอกหัวคูปอง", vbCritical, "Send Error Message"
Me.Text102.SetFocus
Exit Sub
End If

If Me.ListViewCoupon.ListItems.Count = 0 Then
MsgBox "กรุณากรอกมูลค่าคูปองแต่ละมูลค่า", vbCritical, "Send Error Message"
Me.TBAmount.SetFocus
Exit Sub
End If

If Text101.Text <> "" And Text102.Text <> "" Then
    If Me.ListViewMemCoupon.ListItems.Count > 0 Then
    For i = 1 To Me.ListViewMemCoupon.ListItems.Count
    vCouponCode = Me.ListViewMemCoupon.ListItems(i).SubItems(1)
    
    vQuery = "select * from (select code from dbo.bccoupon  where code ='" & vCouponCode & "' union select couponcode as code from dbo.bccouponreceive where couponcode ='" & vCouponCode & "' ) as aa  "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vCouponName = Trim(vRecordset.Fields("name").Value)
        Me.ListViewMemCoupon.ListItems(i).SubItems(3) = Trim(vRecordset.Fields("code").Value)
        MsgBox "คูปองเลขที่ " & vCouponCode & " มีอยู่แล้วในระบบ เป็นคูปอง " & vCouponName & " กรุณาตรวจสอบ", vbCritical, "Send Error Message"
        Exit Sub
    End If
    vRecordset.Close
    Next i
    End If

    vStartDate = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
    vEndDate = Trim(DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year)
    vMydescription = Trim(Text103.Text)
    If vStartDate > vEndDate Then
        MsgBox "วันที่เริ่มใช้คูปองต้องน้อยกว่า วันที่หมดอายุของคูปอง", vbCritical, "Send Error"
        Exit Sub
    End If
    
    If vIsOpen1 = 0 Then
        vQuery = "select isnull(max(id),0)+1 as MaxID from npmaster.dbo.tb_pm_couponmaster"
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vID = Trim(vRecordset.Fields("maxid").Value)
        End If
        vRecordset.Close
        vHeaderNo = UCase(Trim(Text102.Text))
        vCouponName = Trim(Text101.Text)
        
        vQuery = "select headerno from npmaster.dbo.tb_pm_couponmaster  where headerno = '" & vHeaderNo & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckHeader = 1
        Else
            vCheckHeader = 0
        End If
        vRecordset.Close
        
        If Me.CMBPosition.Text <> "" Then
        vPosition = Me.CMBPosition.Text
        End If
        
        If Me.CBDash.Value = 1 Then
        vIsDash = 1
        Else
        vIsDash = 0
        End If
        
        If vCheckHeader = 0 Then
            vQuery = "exec dbo.USP_PM_InsertCouponMaster 0," & vID & ",'" & vHeaderNo & "','" & vCouponName & "','" & vStartDate & "','" & vEndDate & "','0','" & vMydescription & "','" & vUserID & "'," & vPosition & "," & vIsDash & " "
            gConnection.Execute (vQuery)
            

            
            For n = 1 To Me.ListViewCoupon.ListItems.Count
            vCouponAmount = Me.ListViewCoupon.ListItems(n).SubItems(4)
            vStartNum = Me.ListViewCoupon.ListItems(n).SubItems(5)
            vStopNum = Me.ListViewCoupon.ListItems(n).SubItems(6)
            vCountCoupon = Me.ListViewCoupon.ListItems(n).SubItems(7)
            
            vQuery = "exec dbo.USP_PM_InsertCouponSub " & vID & "," & vCouponAmount & "," & vStartNum & "," & vStopNum & "," & vCountCoupon & " "
            gConnection.Execute (vQuery)
            Next n
            
        Else
            MsgBox "กรุณาตรวจสอบ หัวคูปอง เพราะมีอยู่แล้ว ", vbCritical, "Send Error"
            Exit Sub
        End If
        MsgBox "บันทึกข้อมูลเรียบร้อยแล้ว", vbInformation, "Send Message"
    Else
        vCouponName = Trim(Text101.Text)
        vID = Me.LBLID.Caption
        vHeaderNo = UCase(Trim(Text102.Text))
    
        vQuery = "select headerno from npmaster.dbo.tb_pm_couponmaster  where headerno = '" & vHeaderNo & "' and id <> " & vID & " "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckHeader = 1
        Else
            vCheckHeader = 0
        End If
        vRecordset.Close
        
        If Check101.Value = 1 Then
            vIsCancel = 1
        Else
            vIsCancel = 0
        End If
        
        If Me.CMBPosition.Text <> "" Then
        vPosition = Me.CMBPosition.Text
        End If
        
        If Me.CBDash.Value = 1 Then
        vIsDash = 1
        Else
        vIsDash = 0
        End If
        
        If vCheckHeader = 0 Then
            vQuery = "exec dbo.USP_PM_InsertCouponMaster 1," & vID & ",'" & vHeaderNo & "','" & vCouponName & "','" & vStartDate & "','" & vEndDate & "','" & vIsCancel & "','" & vMydescription & "','" & vUserID & "'," & vPosition & "," & vIsDash & "  "
            gConnection.Execute (vQuery)
            
            For n = 1 To Me.ListViewCoupon.ListItems.Count
            vCouponAmount = Me.ListViewCoupon.ListItems(n).SubItems(4)
            vStartNum = Me.ListViewCoupon.ListItems(n).SubItems(5)
            vStopNum = Me.ListViewCoupon.ListItems(n).SubItems(6)
            vCountCoupon = Me.ListViewCoupon.ListItems(n).SubItems(7)
            m = n - 1
            
            vQuery = "exec dbo.USP_PM_InsertCouponSub " & vID & "," & vCouponAmount & "," & vStartNum & "," & vStopNum & "," & vCountCoupon & " ," & m & ""
            gConnection.Execute (vQuery)
            Next n
            
        Else
            MsgBox "กรุณาตรวจสอบ หัวคูปอง เพราะมีอยู่แล้ว ", vbCritical, "Send Error"
            Exit Sub
        End If
        MsgBox "ปรับปรุงข้อมูลเรียบร้อยแล้ว", vbInformation, "Send Message"
    End If
    
Call ClearScreen
Me.Text101.SetFocus


End If
End Sub

Private Sub CMD103_Click()
vIsOpen1 = 0
vMemCouponNo = ""
vMemCountUse = 0
vMemIsConfirm = 0

Me.LBLID.Caption = ""
Me.Text101.Enabled = True
Me.Text102.Enabled = True
Me.DTPicker101.Value = Now
Me.DTPicker102.Value = Now
Me.Text101.Text = ""
Me.Text102.Text = ""
Me.Text103.Text = ""
Me.ListView101.ListItems.Clear
Me.Text104.Text = ""
Me.TBAmount.Text = ""
Me.TBStartNum.Text = ""
Me.TBStopNum.Text = ""
Me.LBLCountCoupon.Caption = ""
Me.CMBPosition.Enabled = True
Me.CMBPosition.ListIndex = 1
Me.CBDash.Enabled = True
Me.CBDash.Value = 0
Me.LBLCoupon.Caption = ""
Me.LBLIndex.Caption = ""
Me.ListViewCoupon.ListItems.Clear
Me.ListViewMemCoupon.ListItems.Clear
Me.Text101.SetFocus
End Sub

Public Sub ClearScreen()
vIsOpen1 = 0
vMemCouponNo = ""
vMemCountUse = 0
Me.LBLID.Caption = ""
Me.Text101.Enabled = True
Me.Text102.Enabled = True
Me.DTPicker101.Value = Now
Me.DTPicker102.Value = Now
Me.Text101.Text = ""
Me.Text102.Text = ""
Me.Text103.Text = ""
Me.ListView101.ListItems.Clear
Me.Text104.Text = ""
Me.TBAmount.Text = ""
Me.TBStartNum.Text = ""
Me.TBStopNum.Text = ""
Me.LBLCountCoupon.Caption = ""
Me.CMBPosition.Enabled = True
Me.CMBPosition.ListIndex = 1
Me.CBDash.Enabled = True
Me.CBDash.Value = 0
Me.LBLCoupon.Caption = ""
Me.LBLIndex.Caption = ""
Me.ListViewCoupon.ListItems.Clear
Me.ListViewMemCoupon.ListItems.Clear
Me.CMD102.Enabled = False
Me.CMDApprove.Enabled = False
Me.CMDPrintCoupon.Enabled = False
Me.Text101.SetFocus
End Sub

Private Sub CMD104_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCouponName As String
Dim vCPID As Integer

If Text104.Text <> "" Then
    vCPID = ListView101.ListItems.Item(ListView101.SelectedItem.Index)
    vQuery = "select pathname from npmaster.dbo.TB_PM_CouponName where cpid = " & vCPID & " "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vCouponName = Trim(vRecordset.Fields("pathname").Value)
    End If
    vRecordset.Close
    
    'With Crystal101
    '.ReportFileName = Trim(vCouponName)
    '.Destination = crptToWindow
    '.WindowState = crptMaximized
    '.Action = 1
   ' End With
    
Else
    MsgBox "กรุณาเลือก ชื่อของคูปองที่อยู่ในตารางด้วย", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Function CheckCoupon(vNumber As Double) As String
Dim vHeader As String
Dim vPosition As Integer
Dim vCheckDash As Integer


If Me.CBDash.Value = 1 Then
vCheckDash = 1
Else
vCheckDash = 0
End If

If Me.Text102.Text <> "" Then
vHeader = UCase(Me.Text102.Text)
End If

If Me.CMBPosition.Text <> "" Then
vPosition = Me.CMBPosition.Text
End If

If vCheckDash = 0 Then

    If vPosition = 1 Then
    vMemCouponNo = vHeader & Format(vNumber, "0")
    End If

    If vPosition = 2 Then
    vMemCouponNo = vHeader & Format(vNumber, "00")
    End If
    
    If vPosition = 3 Then
    vMemCouponNo = vHeader & Format(vNumber, "000")
    End If
    
    If vPosition = 4 Then
    vMemCouponNo = vHeader & Format(vNumber, "0000")
    End If
    
    If vPosition = 5 Then
    vMemCouponNo = vHeader & Format(vNumber, "00000")
    End If
    
    If vPosition = 6 Then
    vMemCouponNo = vHeader & Format(vNumber, "000000")
    End If
    
End If

If vCheckDash = 1 Then

    If vPosition = 1 Then
    vMemCouponNo = vHeader & "-" & Format(vNumber, "0")
    End If

    If vPosition = 2 Then
    vMemCouponNo = vHeader & "-" & Format(vNumber, "00")
    End If
    
    If vPosition = 3 Then
    vMemCouponNo = vHeader & "-" & Format(vNumber, "000")
    End If
    
    If vPosition = 4 Then
    vMemCouponNo = vHeader & "-" & Format(vNumber, "0000")
    End If
    
    If vPosition = 5 Then
    vMemCouponNo = vHeader & "-" & Format(vNumber, "00000")
    End If
    
    If vPosition = 6 Then
    vMemCouponNo = vHeader & "-" & Format(vNumber, "000000")
    End If
    
End If

End Function

Public Sub GenCoupon()
Dim vHeader As String
Dim vPosition As Integer
Dim vCheckDash As Integer
Dim vStartNum As Double
Dim vStopNum As Double
Dim vCouponNo As String
Dim vCountCoupon As Integer

If Me.TBStartNum.Text <> "" Then
vStartNum = Me.TBStartNum.Text
End If

If Me.TBStopNum.Text <> "" Then
vStopNum = Me.TBStopNum.Text
End If

vCountCoupon = vStopNum - vStartNum

vCountCoupon = vCountCoupon + 1

Me.LBLCountCoupon.Caption = vCountCoupon


If Me.CBDash.Value = 1 Then
vCheckDash = 1
Else
vCheckDash = 0
End If

If Me.Text102.Text <> "" Then
vHeader = UCase(Me.Text102.Text)
End If

If Me.CMBPosition.Text <> "" Then
vPosition = Me.CMBPosition.Text
End If

If vCheckDash = 0 Then

    If vPosition = 1 Then
    vCouponNo = vHeader & Format(vStartNum, "0")
    End If

    If vPosition = 2 Then
    vCouponNo = vHeader & Format(vStartNum, "00")
    End If
    
    If vPosition = 3 Then
    vCouponNo = vHeader & Format(vStartNum, "000")
    End If
    
    If vPosition = 4 Then
    vCouponNo = vHeader & Format(vStartNum, "0000")
    End If
    
    If vPosition = 5 Then
    vCouponNo = vHeader & Format(vStartNum, "00000")
    End If
    
    If vPosition = 6 Then
    vCouponNo = vHeader & Format(vStartNum, "000000")
    End If
    
End If

If vCheckDash = 1 Then

    If vPosition = 1 Then
    vCouponNo = vHeader & "-" & Format(vStartNum, "0")
    End If

    If vPosition = 2 Then
    vCouponNo = vHeader & "-" & Format(vStartNum, "00")
    End If
    
    If vPosition = 3 Then
    vCouponNo = vHeader & "-" & Format(vStartNum, "000")
    End If
    
    If vPosition = 4 Then
    vCouponNo = vHeader & "-" & Format(vStartNum, "0000")
    End If
    
    If vPosition = 5 Then
    vCouponNo = vHeader & "-" & Format(vStartNum, "00000")
    End If
    
    If vPosition = 6 Then
    vCouponNo = vHeader & "-" & Format(vStartNum, "000000")
    End If
    
End If

Me.LBLCoupon.Caption = vCouponNo

End Sub

Private Sub CMDAddCoupon_Click()
Dim vNow As Date
Dim vStartDate As Date
Dim vStopDate As Date

Dim vListItemCode As ListItem
Dim i As Integer
Dim vAmount As Double
Dim vQty As Double
Dim vHeader As String
Dim vHeaderNumber As String
Dim vPosition As Integer
Dim vCheckDash As Integer


If Me.Text101.Text = "" Then
MsgBox "ยังไม่ได้กรอกชื่อทะเบียนคูปอง กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Me.Text101.SetFocus
Exit Sub
End If

If Me.Text102.Text = "" Then
MsgBox "ยังไม่ได้กรอก หัวคูปอง กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Me.Text102.SetFocus
Exit Sub
End If

vNow = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
vStartDate = Me.DTPicker101.Day & "/" & Me.DTPicker101.Month & "/" & Me.DTPicker101.Year
vStopDate = Me.DTPicker102.Day & "/" & Me.DTPicker102.Month & "/" & Me.DTPicker102.Year

If vStopDate < vNow Then
MsgBox "ไม่สามารถกำหนดวันหมดอายุคูปองน้อยกว่าวันที่ปัจจุบันได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Me.DTPicker102.SetFocus
Exit Sub
End If

If vStopDate < vStartDate Then
MsgBox "ไม่สามารถกำหนดวันหมดอายุคูปองน้อยกว่าวันเริ่มใช้งานคูปองได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Me.DTPicker102.SetFocus
Exit Sub
End If

If Me.TBAmount.Text = "" Then
MsgBox "ยังไม่ได้กรอก มูลค่าคูปอง กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Me.TBAmount.SetFocus
Exit Sub
Else
vAmount = Me.TBAmount.Text
End If

Dim vStartNum As Double
Dim vStopNum As Double
Dim vCountCoupon As Integer

If Me.LBLCountCoupon.Caption <> "" Then
vCountCoupon = Me.LBLCountCoupon.Caption
End If

If Me.TBStartNum.Text <> "" Then
vStartNum = Me.TBStartNum.Text
End If

If Me.TBStopNum.Text <> "" Then
vStopNum = Me.TBStopNum.Text
End If


If vStopNum < vStartNum Then
MsgBox "เลขที่คูปองสุดท้าย น้อยกว่าเลขที่คูปองเริ่มต้น  กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Me.TBStopNum.SetFocus
Exit Sub
End If

If Me.CMBPosition.Text = "" Then
vPosition = 1
Else
vPosition = Me.CMBPosition.Text
End If

If Me.CBDash.Value = 1 Then
vCheckDash = 1
Else
vCheckDash = 0
End If


vHeader = UCase(Me.Text102.Text)

Dim m As Integer
Dim n As Integer
Dim vCheckCoupon As String
Dim vNumber As Double

For m = vStartNum To vStopNum
vNumber = m
Call CheckCoupon(vNumber)

Dim vCheckAmount As Double

Dim vIndex As Integer
If Me.LBLIndex.Caption <> "" Then
vIndex = Me.LBLIndex.Caption
End If

If Me.ListViewCoupon.ListItems.Count > 0 Then
    For i = 1 To Me.ListViewCoupon.ListItems.Count
        vCheckAmount = Me.ListViewCoupon.ListItems(i).SubItems(4)
        If Me.LBLIndex.Caption <> "" Then
        vIndex = Me.LBLIndex.Caption
        End If
        If vAmount = vCheckAmount Then
            If vIndex = 0 Then
            MsgBox "มูลค่าคูปอง " & vCheckAmount & "  กำหนดไว้ก่อนหน้านี้แล้ว กรุณาตรวจสอบ", vbCritical, "Send Error Message"
            Me.TBAmount.SetFocus
            Exit Sub
            End If
        End If
    Next i
End If

If vIndex <> 0 Then
Me.ListViewCoupon.ListItems(vIndex).SubItems(1) = vHeader
Me.ListViewCoupon.ListItems(vIndex).SubItems(2) = vStartDate
Me.ListViewCoupon.ListItems(vIndex).SubItems(3) = vStartDate
Me.ListViewCoupon.ListItems(vIndex).SubItems(4) = Format(vAmount, "##,##0.00")
Me.ListViewCoupon.ListItems(vIndex).SubItems(5) = Format(vStartNum, "##,##0.00")
Me.ListViewCoupon.ListItems(vIndex).SubItems(6) = Format(vStopNum, "##,##0.00")
Me.ListViewCoupon.ListItems(vIndex).SubItems(7) = Format(vCountCoupon, "##,##0.00")
Me.ListViewCoupon.ListItems(vIndex).SubItems(8) = vCheckDash
Me.ListViewCoupon.ListItems(vIndex).SubItems(9) = vPosition
Me.TBAmount.Text = ""
Me.TBStartNum.Text = ""
Me.TBStopNum.Text = ""
Me.LBLCountCoupon.Caption = ""
Me.LBLIndex.Caption = ""
Me.CMDApprove.Enabled = False
Me.TBAmount.SetFocus
Exit Sub
End If

If Me.ListViewMemCoupon.ListItems.Count > 0 Then
For n = 1 To Me.ListViewMemCoupon.ListItems.Count
vCheckCoupon = Me.ListViewMemCoupon.ListItems(n).SubItems(1)
If vCheckCoupon = vMemCouponNo Then
    MsgBox "เลขที่คูปอง " & vMemCouponNo & " มีการกำหนดเรียบร้อยแล้ว อยู่ในบรรทัดที่ " & n & " ในรายการตรวจสอบคูปอง", vbCritical, "Send Error Message"
    Me.CMDViewCoupon.SetFocus
    Exit Sub
End If
Next n
End If
Next m

i = Me.ListViewCoupon.ListItems.Count
i = i + 1
Set vListItemCode = ListViewCoupon.ListItems.Add(, , i)
vListItemCode.SubItems(1) = vHeader
vListItemCode.SubItems(2) = vStartDate
vListItemCode.SubItems(3) = vStopDate
vListItemCode.SubItems(4) = Format(vAmount, "##,##0.00")
vListItemCode.SubItems(5) = Format(vStartNum, "##,##0.00")
vListItemCode.SubItems(6) = Format(vStopNum, "##,##0.00")
vListItemCode.SubItems(7) = Me.LBLCountCoupon.Caption
vListItemCode.SubItems(8) = vCheckDash
vListItemCode.SubItems(9) = vPosition

Call InsertCoupon

Me.TBAmount.Text = ""
Me.TBStartNum.Text = ""
Me.TBStopNum.Text = ""
Me.LBLCountCoupon.Caption = ""
Me.LBLCoupon.Caption = ""

Me.Text102.Enabled = False
Me.CMBPosition.Enabled = False
Me.CBDash.Enabled = False
Me.CMDApprove.Enabled = False
Me.TBAmount.SetFocus

End Sub


Public Sub InsertCouponToDatabase()
Dim vQuery As String
Dim vListItemCode As ListItem
Dim i As Integer
Dim n As Integer
Dim vAmount As Double
Dim vQty As Double
Dim vHeader As String
Dim vCheckDash As Integer

Dim vPosition As Integer
Dim vStartNum As Double
Dim vStopNum As Double
Dim vCouponNo As String
Dim X As Integer
Dim vStartDate As String
Dim vStopDate As String
Dim vCouponName As String
Dim vBegDate As String
Dim vEndDate As String


If Me.ListViewCoupon.ListItems.Count > 0 Then

If Me.Text101.Text <> "" Then
vCouponName = Me.Text101.Text
End If

For X = 1 To Me.ListViewCoupon.ListItems.Count
vStartNum = Me.ListViewCoupon.ListItems(X).SubItems(5)
vStopNum = Me.ListViewCoupon.ListItems(X).SubItems(6)
vCheckDash = Me.ListViewCoupon.ListItems(X).SubItems(8)
vHeader = UCase(Me.ListViewCoupon.ListItems(X).SubItems(1))
vPosition = Me.ListViewCoupon.ListItems(X).SubItems(9)
vAmount = Me.ListViewCoupon.ListItems(X).SubItems(4)

If vCheckDash = 0 Then

    If vPosition = 1 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "0")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)

    Next i
    End If

    If vPosition = 2 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "00")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    
    Next i
    End If
    
    If vPosition = 3 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "000")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    
    Next i
    End If
    
    If vPosition = 4 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "0000")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    
    Next i
    End If
    
    If vPosition = 5 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "00000")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    
    Next i
    End If
    
    If vPosition = 6 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "000000")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    
    Next i
    End If
    
End If

If vCheckDash = 1 Then

    If vPosition = 1 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "0")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    
    Next i
    End If

    If vPosition = 2 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "00")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    
    Next i
    End If
    
    If vPosition = 3 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "000")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    
    Next i
    End If
    
    If vPosition = 4 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "0000")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    
    Next i
    End If
    
    If vPosition = 5 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "00000")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    
    Next i
    End If
    
    If vPosition = 6 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "000000")
    
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCouponNo & "','" & vCouponName & "','" & vStartDate & "','" & vStopDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCouponNo & "','" & vCouponName & "','" & vBegDate & "','" & vEndDate & "'," & vAmount & " "
    gConnection.Execute (vQuery)
    
    Next i
    End If
End If

Next X
End If

End Sub
Public Sub InsertCouponEdit(vIndex As Integer)
Dim vListItemCode As ListItem
Dim i As Integer
Dim n As Integer
Dim vAmount As Double
Dim vQty As Double
Dim vHeader As String
Dim vCheckDash As Integer

Dim vPosition As Integer
Dim vStartNum As Double
Dim vStopNum As Double
Dim vCouponNo As String
Dim X As Integer


If Me.ListViewCoupon.ListItems.Count > 0 Then
For X = 1 To Me.ListViewCoupon.ListItems.Count
If X <> vIndex Then
vStartNum = Me.ListViewCoupon.ListItems(X).SubItems(5)
vStopNum = Me.ListViewCoupon.ListItems(X).SubItems(6)
vCheckDash = Me.ListViewCoupon.ListItems(X).SubItems(8)
vHeader = UCase(Me.ListViewCoupon.ListItems(X).SubItems(1))
vPosition = Me.ListViewCoupon.ListItems(X).SubItems(9)
vAmount = Me.ListViewCoupon.ListItems(X).SubItems(4)

If vCheckDash = 0 Then

    If vPosition = 1 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "0")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If

    If vPosition = 2 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "00")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 3 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 4 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "0000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 5 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "00000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 6 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "000000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
End If

If vCheckDash = 1 Then

    If vPosition = 1 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "0")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If

    If vPosition = 2 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "00")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 3 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 4 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "0000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 5 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "00000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 6 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "000000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
End If
End If
Next X
End If

End Sub

Public Sub InsertCouponSearch()
Dim vListItemCode As ListItem
Dim i As Integer
Dim n As Integer
Dim vAmount As Double
Dim vQty As Double
Dim vHeader As String
Dim vCheckDash As Integer

Dim vPosition As Integer
Dim vStartNum As Double
Dim vStopNum As Double
Dim vCouponNo As String
Dim X As Integer


If Me.ListViewCoupon.ListItems.Count > 0 Then
For X = 1 To Me.ListViewCoupon.ListItems.Count
vStartNum = Me.ListViewCoupon.ListItems(X).SubItems(5)
vStopNum = Me.ListViewCoupon.ListItems(X).SubItems(6)
vCheckDash = Me.ListViewCoupon.ListItems(X).SubItems(8)
vHeader = UCase(Me.ListViewCoupon.ListItems(X).SubItems(1))
vPosition = Me.ListViewCoupon.ListItems(X).SubItems(9)
vAmount = Me.ListViewCoupon.ListItems(X).SubItems(4)

If vCheckDash = 0 Then

    If vPosition = 1 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "0")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If

    If vPosition = 2 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "00")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 3 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 4 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "0000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 5 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "00000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 6 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "000000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
End If

If vCheckDash = 1 Then

    If vPosition = 1 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "0")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If

    If vPosition = 2 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "00")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 3 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 4 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "0000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 5 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "00000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 6 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "000000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
End If

Next X
End If

End Sub

Public Sub InsertCoupon()
Dim vListItemCode As ListItem
Dim i As Integer
Dim n As Integer
Dim vAmount As Double
Dim vQty As Double
Dim vHeader As String
Dim vCheckDash As Integer

Dim vPosition As Integer
Dim vStartNum As Double
Dim vStopNum As Double
Dim vCouponNo As String


If Me.TBStartNum.Text <> "" Then
vStartNum = Me.TBStartNum.Text
End If

If Me.TBStopNum.Text <> "" Then
vStopNum = Me.TBStopNum.Text
End If

If Me.CBDash.Value = 1 Then
vCheckDash = 1
Else
vCheckDash = 0
End If

If Me.Text102.Text <> "" Then
vHeader = UCase(Me.Text102.Text)
End If

If Me.CMBPosition.Text <> "" Then
vPosition = Me.CMBPosition.Text
End If

If Me.TBAmount.Text <> "" Then
vAmount = Me.TBAmount.Text
End If

If vCheckDash = 0 Then

    If vPosition = 1 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "0")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If

    If vPosition = 2 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "00")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 3 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 4 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "0000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 5 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "00000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 6 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & Format(i, "000000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
End If

If vCheckDash = 1 Then

    If vPosition = 1 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "0")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If

    If vPosition = 2 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "00")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 3 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 4 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "0000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 5 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "00000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
    If vPosition = 6 Then
    For i = vStartNum To vStopNum
    vCouponNo = vHeader & "-" & Format(i, "000000")
    
    n = Me.ListViewMemCoupon.ListItems.Count
    n = n + 1
    Set vListItemCode = ListViewMemCoupon.ListItems.Add(, , n)
    vListItemCode.SubItems(1) = vCouponNo
    vListItemCode.SubItems(2) = vAmount
    Next i
    End If
    
End If

End Sub

Private Sub CMDAddCoupon_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBAmount.Text = ""
Me.TBStartNum.Text = ""
Me.TBStopNum.Text = ""
Me.LBLCountCoupon.Caption = ""
Me.TBAmount.SetFocus
End If
End Sub

Private Sub CMDApprove_Click()
Dim vCouponName As String
Dim vQuery As String

If vAccess = 1 Then
If vIsOpen1 = 1 Then
If vMemIsCancel = 0 Then
If vMemIsConfirm = 0 Then
vQuery = "exec dbo.USP_PM_ConfirmCoupon '" & vCouponName & "' "
gConnection.Execute (vQuery)

Call InsertCouponToDatabase
MsgBox "อนุมัติทะเบียนคูปองและเพิ่มเลขที่คูปองในระบบเรียบร้อย กรุณาตรวจสอบ", vbInformation, "Send Information Message"
Call ClearScreen
Me.Text101.SetFocus

Else
MsgBox "ทะเบียนคูปองถูกอนุมัติไปแล้ว", vbCritical, "Send Error Message"
End If
Else
MsgBox "ทะเบียนคูปองถูกยกเลิกไปแล้ว", vbCritical, "Send Error Message"
End If
Else
MsgBox "คุณยังไม่ได้บันทึกข้อมูลทะเบียนคูปอง", vbCritical, "Send Error Message"
End If
Else
MsgBox "คุณไม่มีสิทธิ์ใช้งาน การอนุมัติทะเบียนคูปอง", vbCritical, "Send Error Message"
End If

End Sub

Private Sub CMDViewCoupon_Click()
If Me.ListViewMemCoupon.Visible = False Then
Me.ListViewMemCoupon.Visible = True

If vIsOpen1 = 0 Then
Call vCheckCouponUse
ElseIf vIsOpen1 = 1 Then
Me.ListViewMemCoupon.ListItems.Clear
Call InsertCouponSearch
Call vCheckCouponUse
End If

Me.ListViewMemCoupon.SetFocus
Else
Me.ListViewMemCoupon.Visible = False
Me.CMDViewCoupon.SetFocus
End If

If vMemCountUse = 0 Then
    Me.CMD102.Enabled = True
Else
    Me.CMD102.Enabled = False
End If


End Sub

Public Sub vCheckCouponUse()
Dim i As Integer
Dim vCouponCode As String
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

If Me.ListViewMemCoupon.ListItems.Count > 0 Then
For i = 1 To Me.ListViewMemCoupon.ListItems.Count
vCouponCode = Me.ListViewMemCoupon.ListItems(i).SubItems(1)

vQuery = "select code from bccoupon where code ='" & vCouponCode & "' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.ListViewMemCoupon.ListItems(i).SubItems(3) = Trim(vRecordset.Fields("code").Value)
    vMemCountUse = vMemCountUse + 1
    Me.ListViewMemCoupon.ListItems(i).ForeColor = &H8080FF
    Me.ListViewMemCoupon.ListItems(i).ListSubItems(1).ForeColor = &H8080FF
    Me.ListViewMemCoupon.ListItems(i).ListSubItems(2).ForeColor = &H8080FF
    Me.ListViewMemCoupon.ListItems(i).ListSubItems(3).ForeColor = &H8080FF
End If
vRecordset.Close
Next i
End If
End Sub

Private Sub DTPicker101_Click()
Me.DTPicker102.SetFocus
End Sub

Private Sub DTPicker102_Change()
Dim vNow As Date
Dim vStartDate As Date
Dim vStopDate As Date

vNow = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
vStartDate = Me.DTPicker101.Day & "/" & Me.DTPicker101.Month & "/" & Me.DTPicker101.Year
vStopDate = Me.DTPicker102.Day & "/" & Me.DTPicker102.Month & "/" & Me.DTPicker102.Year

If vStopDate < vNow Then
MsgBox "ไม่สามารถกำหนดวันหมดอายุคูปองน้อยกว่าวันที่ปัจจุบันได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Me.DTPicker102.SetFocus
Exit Sub
End If

If vStopDate < vStartDate Then
MsgBox "ไม่สามารถกำหนดวันหมดอายุคูปองน้อยกว่าวันเริ่มใช้งานคูปองได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Me.DTPicker102.SetFocus
Exit Sub
End If
End Sub

Private Sub DTPicker102_Click()
Me.Text103.SetFocus
End Sub

Private Sub Form_Load()
DTPicker101.Value = Now
DTPicker102.Value = Now

Me.CMBPosition.AddItem (1)
Me.CMBPosition.AddItem (2)
Me.CMBPosition.AddItem (3)
Me.CMBPosition.AddItem (4)
Me.CMBPosition.AddItem (5)
Me.CMBPosition.AddItem (6)
Me.CMBPosition.Text = 1
End Sub

Private Sub List101_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCouponName As String
Dim vCouponList As ListItem
Dim i As Integer
Dim vListItem As ListItem
Dim vID As Integer
Dim vCouponAmount As Double
Dim vStartNum As Integer
Dim vStopNum As Integer
Dim vCountCoupon As Integer
Dim vPosition As Integer

List101.Visible = False
Text101.Enabled = False
Text101.Text = List101.Text
vCouponName = Trim(Text101.Text)
Me.ListViewCoupon.ListItems.Clear
Me.ListViewMemCoupon.ListItems.Clear

vQuery = "exec dbo.usp_pm_searchcoupondetails '" & vCouponName & "' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    
    Me.LBLID.Caption = Trim(vRecordset.Fields("id").Value)
    Text102.Text = Trim(vRecordset.Fields("headerno").Value)
    DTPicker101.Value = Trim(vRecordset.Fields("startdate").Value)
    DTPicker102.Value = Trim(vRecordset.Fields("enddate").Value)
    vPosition = vRecordset.Fields("position").Value
    Me.CMBPosition.Text = vPosition
    If vRecordset.Fields("isdash").Value = 1 Then
    Me.CBDash.Value = 1
    End If
    vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
    vMemIsCancel = Trim(vRecordset.Fields("iscancel").Value)
    
    If vMemIsConfirm = 1 Then
    Me.CMBPosition.Enabled = False
    Me.CBDash.Enabled = False
    Me.Text101.Enabled = False
    Me.Text102.Enabled = False
    Me.CMDApprove.Enabled = False
    Me.CMDPrintCoupon.Enabled = True
    Else
    Me.CMBPosition.Enabled = False
    Me.CBDash.Enabled = False
    Me.Text101.Enabled = True
    Me.Text102.Enabled = False
    Me.CMDApprove.Enabled = True
    Me.CMDPrintCoupon.Enabled = False
    End If
    
    vID = Trim(vRecordset.Fields("id").Value)
    Text103.Text = Trim(vRecordset.Fields("mydescription").Value)
    vIsOpen1 = 1
    
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        i = i + 1
        Set vListItem = Me.ListViewCoupon.ListItems.Add(, , i)
        vCouponAmount = Trim(vRecordset.Fields("couponamount").Value)
        vStartNum = Trim(vRecordset.Fields("startnum").Value)
        vStopNum = Trim(vRecordset.Fields("stopnum").Value)
        vCountCoupon = Trim(vRecordset.Fields("couponcount").Value)
        
        vListItem.SubItems(1) = Trim(vRecordset.Fields("headerno").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("startdate").Value)
        vListItem.SubItems(3) = Trim(vRecordset.Fields("enddate").Value)
        vListItem.SubItems(4) = Format(vCouponAmount, "##,##0.00")
        vListItem.SubItems(5) = Format(vStartNum, "##,##0.00")
        vListItem.SubItems(6) = Format(vStopNum, "##,##0.00")
        vListItem.SubItems(7) = Format(vCountCoupon, "##,##0.00")
        vListItem.SubItems(8) = Trim(vRecordset.Fields("isdash").Value)
        vListItem.SubItems(9) = Trim(vRecordset.Fields("position").Value)
        vRecordset.MoveNext
    Wend
    
    Me.ListViewMemCoupon.ListItems.Clear
   Call InsertCouponSearch
   Me.TBAmount.SetFocus
    
Else
    MsgBox "ไม่มีข้อมูลคูปองนี้อยู่ในระบบ กรุณาตรวจสอบ", vbInformation, "Send Information"
End If
vRecordset.Close


End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
If ListView101.ListItems.Count <> 0 Then
    Text104.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(1))
End If
End Sub

Private Sub ListViewCoupon_DblClick()
Dim vIndex As Integer

If Me.ListViewCoupon.ListItems.Count > 0 Then
    vIndex = Me.ListViewCoupon.SelectedItem.Index
    Me.ListViewMemCoupon.ListItems.Clear
    Call InsertCouponEdit(vIndex)
    Me.LBLIndex.Caption = vIndex
    If vIsOpen1 = 1 And vMemIsConfirm = 0 Then
        Me.TBAmount.Text = Me.ListViewCoupon.ListItems(vIndex).SubItems(4)
        Me.TBStartNum.Text = Me.ListViewCoupon.ListItems(vIndex).SubItems(5)
        Me.TBStopNum.Text = Me.ListViewCoupon.ListItems(vIndex).SubItems(6)
        Me.LBLCountCoupon.Caption = Me.ListViewCoupon.ListItems(vIndex).SubItems(7)
    ElseIf vIsOpen1 = 1 And vMemIsConfirm = 1 Then
        Me.TBAmount.Text = Me.ListViewCoupon.ListItems(vIndex).SubItems(4)
        Me.TBStartNum.Text = Me.ListViewCoupon.ListItems(vIndex).SubItems(5)
        Me.TBStopNum.Text = Me.ListViewCoupon.ListItems(vIndex).SubItems(6)
        Me.LBLCountCoupon.Caption = Me.ListViewCoupon.ListItems(vIndex).SubItems(7)
        Me.TBStartNum.Enabled = False
        Me.CMBPosition.Enabled = False
        Me.CBDash.Enabled = False
        Me.TBAmount.Enabled = False
        Me.TBStopNum.SetFocus
    End If
End If
End Sub

Private Sub ListViewCoupon_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer

If KeyCode = 46 Then
    If Me.ListViewCoupon.ListItems.Count > 0 Then
        vIndex = Me.ListViewCoupon.SelectedItem.Index
        If vIsOpen1 = 1 And vMemIsConfirm = 0 Then
        Me.ListViewCoupon.ListItems.Remove (vIndex)
        Call vGenIndex
        Call InsertCouponSearch
        
        If Me.ListViewCoupon.ListItems.Count = 0 Then
        Me.Text102.Enabled = True
        Me.CMBPosition.Enabled = True
        Me.CBDash.Enabled = True
        Me.CMDApprove.Enabled = False
        Me.CMDPrintCoupon.Enabled = False
        Me.Text101.SetFocus
        Else
        Me.ListViewCoupon.SetFocus
        End If
        End If
    End If
End If
End Sub

Public Sub vGenIndex()
Dim i As Integer

If Me.ListViewCoupon.ListItems.Count > 0 Then
For i = 1 To Me.ListViewCoupon.ListItems.Count
Me.ListViewCoupon.ListItems(i).Text = i
Next i
End If
End Sub

Private Sub TBAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBAmount.Text = ""
Me.TBStartNum.Text = ""
Me.TBStopNum.Text = ""
Me.LBLCountCoupon.Caption = ""
Me.TBAmount.SetFocus
End If
End Sub

Private Sub TBAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.TBStartNum.SetFocus
End If
End Sub

Private Sub TBAmount_LostFocus()
If Me.TBAmount.Text <> "" Then
 Me.TBAmount.Text = CheckDegit(Me.TBAmount.Text)
 Else
 Me.TBAmount.Text = ""
 End If
End Sub

Private Sub TBStartNum_Change()
Call GenCoupon
End Sub

Private Sub TBStartNum_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBAmount.Text = ""
Me.TBStartNum.Text = ""
Me.TBStopNum.Text = ""
Me.LBLCountCoupon.Caption = ""
Me.TBAmount.SetFocus
End If
End Sub

Private Sub TBStartNum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.TBStopNum.SetFocus
End If
End Sub

Private Sub TBStartNum_LostFocus()
If Me.TBStartNum.Text <> "" Then
 Me.TBStartNum.Text = CheckDegit(Me.TBStartNum.Text)
 Else
 Me.TBStartNum.Text = ""
 End If
End Sub

Private Sub TBStopNum_Change()
Call GenCoupon
End Sub

Private Sub TBStopNum_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.TBAmount.Text = ""
Me.TBStartNum.Text = ""
Me.TBStopNum.Text = ""
Me.LBLCountCoupon.Caption = ""
Me.TBAmount.SetFocus
End If
End Sub

Private Sub TBStopNum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.CMBPosition.Enabled = True Then
        Me.CMBPosition.SetFocus
    Else
    Me.CMDAddCoupon.SetFocus
    End If
End If
End Sub

Private Sub TBStopNum_LostFocus()
If Me.TBStopNum.Text <> "" Then
 Me.TBStopNum.Text = CheckDegit(Me.TBStopNum.Text)
 Else
 Me.TBStopNum.Text = ""
 End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.Text102.SetFocus
End If
End Sub

Private Sub Text102_Change()
Call GenCoupon
End Sub

Private Sub Text102_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.DTPicker101.SetFocus
End If
End Sub

Private Sub Text102_LostFocus()
Me.Text102.Text = UCase(Me.Text102.Text)
End Sub

Private Sub Text103_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.TBAmount.SetFocus
End If
End Sub
