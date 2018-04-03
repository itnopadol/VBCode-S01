VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form111 
   Caption         =   "บันทึกข้อมูล นับสต๊อกตามชั้นเก็บสินค้า"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form111.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICSearchShelf 
      BackColor       =   &H00808080&
      Height          =   8115
      Left            =   0
      ScaleHeight     =   8055
      ScaleWidth      =   11835
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin VB.CommandButton CMDCancel 
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
         Height          =   465
         Left            =   9090
         TabIndex        =   45
         Top             =   6345
         Width           =   1230
      End
      Begin VB.CommandButton CMDSelect 
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
         Left            =   7605
         TabIndex        =   44
         Top             =   6345
         Width           =   1230
      End
      Begin MSComctlLib.ListView ListViewShelf 
         Height          =   4875
         Left            =   1575
         TabIndex        =   43
         Top             =   1350
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   8599
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
            Text            =   "รหัสชั้นเก็บ"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อที่เก็บ"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.CommandButton CMDSearchShelfDetails 
         Height          =   285
         Left            =   4950
         Picture         =   "Form111.frx":72FB
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   540
         Width           =   330
      End
      Begin VB.TextBox TextSearch 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2205
         TabIndex        =   41
         Top             =   540
         Width           =   2715
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการที่ค้นหา"
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
         Left            =   1575
         TabIndex        =   46
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "ค้นหา :"
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
         Left            =   1575
         TabIndex        =   40
         Top             =   540
         Width           =   645
      End
   End
   Begin VB.PictureBox PIC101 
      BackColor       =   &H00404040&
      Height          =   8160
      Left            =   0
      ScaleHeight     =   8100
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
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
         Left            =   8550
         TabIndex        =   23
         Top             =   6705
         Width           =   1410
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
         Left            =   6975
         TabIndex        =   22
         Top             =   6705
         Width           =   1410
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   14
         Left            =   7650
         TabIndex        =   16
         Top             =   6255
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   13
         Left            =   7650
         TabIndex        =   15
         Top             =   5940
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   7650
         TabIndex        =   14
         Top             =   5625
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   11
         Left            =   7650
         TabIndex        =   13
         Top             =   5310
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   10
         Left            =   7650
         TabIndex        =   12
         Top             =   4995
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   9
         Left            =   7650
         TabIndex        =   11
         Top             =   4680
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   7650
         TabIndex        =   10
         Top             =   4365
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   7650
         TabIndex        =   9
         Top             =   4050
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   7650
         TabIndex        =   8
         Top             =   3735
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   7650
         TabIndex        =   7
         Top             =   3420
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   7650
         TabIndex        =   6
         Top             =   3105
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   7650
         TabIndex        =   5
         Top             =   2790
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   7650
         TabIndex        =   4
         Top             =   2475
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   7650
         TabIndex        =   3
         Top             =   2160
         Width           =   2265
      End
      Begin VB.TextBox QTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   7650
         TabIndex        =   2
         Top             =   1845
         Width           =   2265
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   4785
         Left            =   1800
         TabIndex        =   1
         Top             =   1800
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   8440
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับที่"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "ชั้นเก็บ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "คงเหลือ"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "หน่วย"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "นับได้"
            Object.Width           =   4145
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   0
         Picture         =   "Form111.frx":76C8
         Top             =   0
         Width           =   2160
      End
      Begin VB.Label LBLShelf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5265
         TabIndex        =   56
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label Label18 
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
         Height          =   285
         Left            =   4410
         TabIndex        =   55
         Top             =   810
         Width           =   780
      End
      Begin VB.Label LBLWHCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2655
         TabIndex        =   54
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label Label17 
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
         Left            =   1710
         TabIndex        =   53
         Top             =   810
         Width           =   915
      End
      Begin VB.Label LBLItemName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5265
         TabIndex        =   52
         Top             =   1170
         Width           =   4695
      End
      Begin VB.Label LBLItemCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2655
         TabIndex        =   51
         Top             =   1170
         Width           =   1680
      End
      Begin VB.Label Label15 
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
         Left            =   4095
         TabIndex        =   50
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   1530
         TabIndex        =   49
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "นับได้"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   7605
         TabIndex        =   21
         Top             =   1530
         Width           =   2355
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   6120
         TabIndex        =   20
         Top             =   1530
         Width           =   1500
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "คงเหลือ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   4095
         TabIndex        =   19
         Top             =   1530
         Width           =   2040
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ชั้นเก็บ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   2610
         TabIndex        =   18
         Top             =   1530
         Width           =   1500
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ลำดับ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   1530
         Width           =   825
      End
   End
   Begin VB.CommandButton CMDClose 
      Caption         =   "ออก"
      Height          =   465
      Left            =   10215
      TabIndex        =   32
      Top             =   6795
      Width           =   1320
   End
   Begin VB.CommandButton CMDDelete 
      Caption         =   "ลบ"
      Height          =   465
      Left            =   8775
      TabIndex        =   31
      Top             =   6795
      Width           =   1320
   End
   Begin VB.CommandButton CMDSave 
      Caption         =   "บันทึก"
      Height          =   465
      Left            =   7425
      TabIndex        =   30
      Top             =   6795
      Width           =   1230
   End
   Begin MSComctlLib.ListView ListViewItemList 
      Height          =   4740
      Left            =   270
      TabIndex        =   29
      Top             =   1890
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   8361
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ชั้นเก็บ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "คงเหลือ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "นับได้"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "หน่วยนับ"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.TextBox TextItemCode 
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
      Height          =   330
      Left            =   3285
      TabIndex        =   48
      Top             =   1260
      Width           =   1770
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
      Left            =   3285
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   810
      Width           =   1725
   End
   Begin VB.CommandButton CMDSearchShelf 
      Height          =   330
      Left            =   5040
      Picture         =   "Form111.frx":A152
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   810
      Width           =   330
   End
   Begin VB.ComboBox CMBWHCode 
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
      Left            =   3285
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   360
      Width           =   1050
   End
   Begin MSComCtl2.DTPicker DTPDocdate 
      Height          =   285
      Left            =   8145
      TabIndex        =   33
      Top             =   810
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   503
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
      CurrentDate     =   39414
   End
   Begin VB.CommandButton CMD 
      Caption         =   "ปิดการแก้ไข"
      Height          =   465
      Left            =   6075
      TabIndex        =   35
      Top             =   6795
      Width           =   1230
   End
   Begin VB.Label Label13 
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
      Height          =   285
      Left            =   2025
      TabIndex        =   47
      Top             =   1260
      Width           =   1185
   End
   Begin VB.Label LBLDocNo 
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
      Left            =   8145
      TabIndex        =   34
      Top             =   360
      Width           =   1635
   End
   Begin VB.Label Label10 
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
      Height          =   240
      Left            =   270
      TabIndex        =   28
      Top             =   1665
      Width           =   1230
   End
   Begin VB.Label Label9 
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
      Left            =   6930
      TabIndex        =   27
      Top             =   810
      Width           =   1140
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่อ้างอิง :"
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
      Left            =   6930
      TabIndex        =   26
      Top             =   360
      Width           =   1140
   End
   Begin VB.Label Label7 
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
      Height          =   285
      Left            =   2610
      TabIndex        =   25
      Top             =   810
      Width           =   600
   End
   Begin VB.Label Label6 
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
      Height          =   285
      Left            =   2565
      TabIndex        =   24
      Top             =   360
      Width           =   645
   End
End
Attribute VB_Name = "Form111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCheckSameValue As Integer


Private Sub Command1_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vCheckItem As String
Dim vWHCode As String
Dim vListQTY As ListItem
         
'On Error Resume Next



    'vItemCode = "2120250"
    'vWHCode = "014"
    'vQuery = "exec dbo.USP_ISP_SearchProduct1 '" & vItemCode & "','" & vWHCode & "' "
    'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     '  Me.PICKeyQTY.Visible = True
      ' LBLItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
       'LBLItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
       'LBLUnitcode.Caption = Trim(vRecordset.Fields("defstkunitcode").Value)
       'Me.LBLOnHand.Caption = Format(vRecordset.Fields("onhand").Value, "####0.00")
       'Me.TextQTY.Text = ""
       'Me.TextCheckQTY.Text = ""
       'Me.TextCountQTY.Text = ""
       'Me.TextVND.Text = ""
       'Me.TextSHW.Text = ""
      '
       'Me.ListViewQTY.ListItems.Clear
       'vRecordset.MoveFirst
       'While Not vRecordset.EOF
       'Set vListQTY = Me.ListViewQTY.ListItems.Add(, , vRecordset.Fields("shelfcode").Value)
       'vListQTY.SubItems(1) = Format(vRecordset.Fields("qty").Value, "####0.00")
       'vListQTY.SubItems(2) = vRecordset.Fields("stkunitcode").Value
       'vRecordset.MoveNext
       'Wend
      
       'Me.TextItemCode.Text = ""
       'Me.TextQTY.SetFocus
   'Else
    '  MsgBox "ไม่มีรหัสสินค้า " & vItemCode & " นี้ในระบบ ", vbCritical, "Send Error"
     ' LBLItemName.Caption = ""
      'LBLUnitcode.Caption = ""
      'Exit Sub
   'End If
   'vRecordset.Close
          
    '  If Me.ListViewItemList.ListItems.Count > 0 Then
       '  For i = 1 To Me.ListViewItemList.ListItems.Count
     '    vCheckItem = Me.ListViewItemList.ListItems(i).SubItems(1)
      '   If vItemCode = vCheckItem Then
        '    Me.TextQTY.Text = Me.ListViewItemList.ListItems(i).SubItems(4)
         '   Me.TextCheckQTY.Text = Me.ListViewItemList.ListItems(i).SubItems(5)
          '  Me.TextCountQTY.Text = Me.ListViewItemList.ListItems(i).SubItems(6)
           ' Me.TextVND.Text = Me.ListViewItemList.ListItems(i).SubItems(10)
            'Me.TextSHW.Text = Me.ListViewItemList.ListItems(i).SubItems(11)
            'vCheckSameValue = 1
            'Exit Sub
         'End If
         'Next i
         
      'Else
       '  vCheckSameValue = 0
      'End If


End Sub

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
   Me.LBLDocNo.Caption = vRecordset.Fields("docno").Value
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
      Me.LBLDocNo.Caption = ""
   End If
   vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMBWHCode_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String

'On Error Resume Next

vWHCode = Trim(CMBWHCode.Text)
CMBShelf.Clear
vQuery = "exec dbo.USP_MB_SearchShelfID '" & vWHCode & "',''  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        CMBShelf.AddItem Trim(vRecordset.Fields("shelfid").Value)
        vRecordset.MoveNext
        Wend
End If
vRecordset.Close
End Sub

Private Sub CMDCancel_Click()
PICSearchShelf.Visible = False
End Sub

Private Sub CMDExit_Click()
Me.PIC101.Visible = False
End Sub

Private Sub CMDOK_Click()
Dim n As Integer
Dim vListItem As ListItem
Dim i As Integer
Dim vCount As Integer


For n = 0 To 14
   Me.QTY(n).Enabled = False
Next n

If Me.ListView101.ListItems.Count > 0 Then
   If Me.ListViewItemList.ListItems.Count > 0 Then
      vCount = Me.ListViewItemList.ListItems.Count
   Else
      vCount = 0
   End If
   
   
   Dim a As Integer
   Dim b As Integer
   
   
   If vCheckSameValue = 0 Then
   
      For i = 1 To Me.ListView101.ListItems.Count
      vCount = vCount + 1
      Set vListItem = Me.ListViewItemList.ListItems.Add(, , vCount)
      vListItem.SubItems(1) = Me.LBLItemCode.Caption
      vListItem.SubItems(2) = Me.LBLItemName.Caption
      vListItem.SubItems(3) = Me.ListView101.ListItems(i).SubItems(1)
      vListItem.SubItems(4) = Me.ListView101.ListItems(i).SubItems(2)
      vListItem.SubItems(5) = Me.QTY(i - 1).Text
      vListItem.SubItems(6) = Me.ListView101.ListItems(i).SubItems(3)
   
      Next i
   
   Else
      vCheckSameValue = 0
   
   End If
   
   For n = 0 To 14
      Me.QTY(n).Text = ""
   Next n
   
   Me.PIC101.Visible = False
End If

End Sub

Private Sub CMDSearchShelf_Click()
PICSearchShelf.Visible = True
Me.TextSearch.SetFocus
End Sub

Private Sub CMDSearchShelfDetails_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String
Dim vSearch As String
Dim vListShelf As ListItem

'On Error Resume Next

vWHCode = Trim(CMBWHCode.Text)
If TextSearch.Text <> "" Then
   vSearch = Trim(TextSearch.Text)
Else
   vSearch = ""
End If

ListViewShelf.ListItems.Clear
vQuery = "exec dbo.USP_MB_SearchShelfID'" & vWHCode & "' ,'" & vSearch & "' "
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

Private Sub CMDSelect_Click()
If Me.ListViewShelf.ListItems.Count > 0 Then
   Me.CMBShelf.Text = Me.ListViewShelf.SelectedItem.Text
   Me.PICSearchShelf.Visible = False
End If
End Sub

Private Sub Form_Load()
Call GetWHCode
Me.DTPDocdate.Value = Now
End Sub

Public Sub GetWHCode()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

'On Error Resume Next

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
Me.CMBWHCode.Text = "014"
End Sub

Public Sub ClearScreen()
'On Error Resume Next

   vCheckSameValue = 0
   Me.TextItemCode.SetFocus
End Sub

Private Sub ListViewShelf_DblClick()
If Me.ListViewShelf.ListItems.Count > 0 Then
   Me.CMBShelf.Text = Me.ListViewShelf.SelectedItem.Text
   Me.PICSearchShelf.Visible = False
End If
End Sub

Private Sub ListViewShelf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Me.ListViewShelf.ListItems.Count > 0 Then
      Me.CMBShelf.Text = Me.ListViewShelf.SelectedItem.Text
      Me.PICSearchShelf.Visible = False
   End If
End If
End Sub

Private Sub QTY_LostFocus(Index As Integer)
Dim vCheckQty As Double
Dim vCountDot As Integer
Dim vIndex As Integer

'On Error Resume Next

If Me.QTY(Index).Text <> "" Then
   vIndex = Index
   Call CheckNumber(Trim(Me.QTY(Index).Text))
   If vCheckValueNumber = False Then
      MsgBox "กรอกข้อมูลได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      Me.QTY(Index).SetFocus
   Else
      vCountDot = CheckDot(Me.QTY(Index).Text)
      If vCountDot <= 1 Then
         If Me.QTY(Index).Text <> "." Then
            vCheckQty = Me.QTY(Index).Text
            Me.QTY(Index).Text = Format(vCheckQty, "####0.00")
         Else
            MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
            Me.QTY(Index).SetFocus
         End If
      Else
         MsgBox "กรอกอักขระผิด กรุณาแก้ไข", vbCritical, "Send Error Message"
         Me.QTY(Index).SetFocus
      End If
   End If
End If
End Sub

Private Sub TextItemCode_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vCheckItem As String
Dim vWHCode As String
Dim vListItem As ListItem
Dim vItemCode As String
         
'On Error Resume Next

If KeyAscii = 13 Then
If Me.CMBWHCode.Text <> "" And Me.CMBShelf.Text <> "" Then
    vItemCode = Trim(TextItemCode.Text)
    vWHCode = Me.CMBWHCode.Text
    vQuery = "exec dbo.USP_ISP_SearchProduct2 '" & vItemCode & "','" & vWHCode & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       Me.PIC101.Visible = True
       LBLItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
       LBLItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
       LBLWHCode.Caption = Me.CMBWHCode.Text
       LBLShelf.Caption = Me.CMBShelf.Text
       
       Me.ListView101.ListItems.Clear
       vRecordset.MoveFirst
       i = 1
       While Not vRecordset.EOF
       Set vListItem = Me.ListView101.ListItems.Add(, , i)
       If Trim(vRecordset.Fields("shelfcode").Value) <> "" Then
       vListItem.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
       Else
       vListItem.SubItems(1) = "AVL"
       End If
       vListItem.SubItems(2) = Format(vRecordset.Fields("qty").Value, "####0.00")
       vListItem.SubItems(3) = vRecordset.Fields("defstkunitcode").Value
       Me.QTY(i - 1).Enabled = True
       Me.QTY(i - 1).Text = 0
       i = i + 1
       vRecordset.MoveNext
       Wend
      
       Me.TextItemCode.Text = ""
       Me.QTY(0).SetFocus
   Else
      MsgBox "ไม่มีรหัสสินค้า " & vItemCode & " นี้ในระบบ ", vbCritical, "Send Error"
      LBLItemName.Caption = ""
      Exit Sub
   End If
   vRecordset.Close
          
          
Dim j As Integer
Dim n As Integer
Dim m As Integer

Dim vCheckShelf As String
Dim vShelfCode As String
Dim vKeyItemCode As String

      If Me.ListViewItemList.ListItems.Count > 0 Then
         For j = 1 To Me.ListViewItemList.ListItems.Count
         vCheckItem = Me.ListViewItemList.ListItems(j).SubItems(1)
         vCheckShelf = Me.ListViewItemList.ListItems(j).SubItems(3)
         If vItemCode = vCheckItem Then
                For m = 1 To Me.ListView101.ListItems.Count
                   vShelfCode = Me.ListView101.ListItems(m).SubItems(1)
                   
                   If vCheckShelf = vShelfCode Then
                      Me.QTY(m - 1).Text = Me.ListViewItemList.ListItems(j).SubItems(5)
                      vCheckSameValue = 1
                   End If
                Next m

         End If

         Next j

      Else
         vCheckSameValue = 0
      End If
Else
   MsgBox "กรุณาระบุคลังและที่เก็บให้เรียบร้อยก่อน กรอกรายการสินค้า", vbCritical, "Send Error Message"
   Me.CMBShelf.SetFocus
End If
End If
End Sub

Private Sub TextSearch_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String
Dim vSearch As String
Dim vListShelf As ListItem

'On Error Resume Next

If KeyAscii = 13 Then
   vWHCode = Trim(CMBWHCode.Text)
   If TextSearch.Text <> "" Then
      vSearch = Trim(TextSearch.Text)
   Else
      vSearch = ""
   End If
   
   ListViewShelf.ListItems.Clear
   vQuery = "exec dbo.USP_MB_SearchShelfID'" & vWHCode & "' ,'" & vSearch & "' "
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
