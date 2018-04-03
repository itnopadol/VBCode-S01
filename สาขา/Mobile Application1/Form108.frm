VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form108 
   Caption         =   "นับสต็อกประจำปี"
   ClientHeight    =   8175
   ClientLeft      =   1965
   ClientTop       =   630
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form108.frx":0000
   ScaleHeight     =   8175
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   495
      Top             =   6255
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
   Begin VB.PictureBox Pic103 
      Height          =   8160
      Left            =   0
      ScaleHeight     =   8100
      ScaleWidth      =   11295
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   11355
      Begin VB.CheckBox Check101 
         Caption         =   "เลือกทั้งหมด"
         Height          =   285
         Left            =   4410
         TabIndex        =   75
         Top             =   6030
         Width           =   2085
      End
      Begin VB.CommandButton CMD109 
         Caption         =   "ออก"
         Height          =   555
         Left            =   9135
         TabIndex        =   74
         Top             =   6255
         Width           =   1410
      End
      Begin VB.CommandButton CMD106 
         Caption         =   "พิมพ์บันทึกผลการตรวจนับ"
         Height          =   555
         Left            =   7470
         TabIndex        =   72
         Top             =   6255
         Width           =   1410
      End
      Begin VB.CommandButton CMD108 
         Caption         =   "<<ยกเลิก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3420
         TabIndex        =   71
         Top             =   3195
         Width           =   870
      End
      Begin VB.CommandButton CMD107 
         Caption         =   "สรุป >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3420
         TabIndex        =   70
         Top             =   2250
         Width           =   870
      End
      Begin MSComctlLib.ListView ListView103 
         Height          =   3975
         Left            =   4410
         TabIndex        =   67
         Top             =   1935
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   7011
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "จัดซื้อ"
            Object.Width           =   3792
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "เลขที่บันทึกผล"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "เลขที่ปรับปรุง"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView ListView102 
         Height          =   3930
         Left            =   405
         TabIndex        =   66
         Top             =   1935
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   6932
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "จัดซื้อ "
            Object.Width           =   5062
         EndProperty
      End
      Begin VB.CommandButton CMD105 
         Caption         =   "ฟื้นฟูข้อมูล"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   8910
         TabIndex        =   65
         Top             =   360
         Width           =   1635
      End
      Begin VB.ComboBox CMBWareHouse 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   6435
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox CMBID 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   360
         Width           =   1140
      End
      Begin VB.ComboBox CMBAnnually 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label Label25 
         Caption         =   "รายการสรุปผลการตรวจนับ"
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
         Left            =   3960
         TabIndex        =   69
         Top             =   1440
         Width           =   6585
      End
      Begin VB.Label Label24 
         Caption         =   "รายการบันทึกผลการตรวจนัล"
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
         Left            =   405
         TabIndex        =   68
         Top             =   1395
         Width           =   2895
      End
      Begin VB.Label Label23 
         Caption         =   "คลัง :"
         Height          =   330
         Left            =   5850
         TabIndex        =   63
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "ครั้งที่ :"
         Height          =   285
         Left            =   3555
         TabIndex        =   61
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label21 
         Caption         =   "ประจำปี :"
         Height          =   285
         Left            =   450
         TabIndex        =   59
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.Frame Frame101 
      BackColor       =   &H80000009&
      Caption         =   "รายละเอียดทะเบียนการตรวจนับประจำปี"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8250
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   11355
      Begin VB.PictureBox Pic101 
         Height          =   5235
         Left            =   720
         ScaleHeight     =   5175
         ScaleWidth      =   9495
         TabIndex        =   29
         Top             =   855
         Visible         =   0   'False
         Width           =   9555
         Begin VB.CommandButton CMDExit 
            Caption         =   "ออก"
            Height          =   465
            Left            =   7785
            TabIndex        =   33
            Top             =   4365
            Width           =   1050
         End
         Begin MSComctlLib.ListView ListView101 
            Height          =   3165
            Left            =   540
            TabIndex        =   32
            Top             =   945
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   5583
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "รหัส"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ประจำปี"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ครั้งที่"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "รายละเอียด"
               Object.Width           =   8819
            EndProperty
         End
         Begin VB.TextBox TextSearch 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1620
            TabIndex        =   31
            Top             =   270
            Width           =   1905
         End
         Begin VB.Label Label4 
            Caption         =   "ค้นหาทะเบียน :"
            Height          =   375
            Left            =   495
            TabIndex        =   30
            Top             =   315
            Width           =   1320
         End
      End
      Begin VB.CommandButton CMD101 
         Caption         =   "C"
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
         Left            =   3600
         TabIndex        =   28
         Top             =   900
         Width           =   285
      End
      Begin VB.CommandButton CMD104 
         Caption         =   "ออก"
         Height          =   420
         Left            =   4905
         TabIndex        =   27
         Top             =   4455
         Width           =   1005
      End
      Begin VB.CommandButton CMD103 
         Caption         =   "ค้นหา"
         Height          =   420
         Left            =   3690
         TabIndex        =   26
         Top             =   4455
         Width           =   1005
      End
      Begin VB.CommandButton CMD102 
         Caption         =   "บันทึก"
         Height          =   420
         Left            =   2475
         TabIndex        =   25
         Top             =   4455
         Width           =   1005
      End
      Begin VB.TextBox TextDescription 
         Appearance      =   0  'Flat
         Height          =   1995
         Left            =   2475
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   2160
         Width           =   6855
      End
      Begin VB.TextBox TextTimes 
         Alignment       =   2  'Center
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
         Left            =   2475
         TabIndex        =   23
         Text            =   "1"
         Top             =   1530
         Width           =   780
      End
      Begin MSComCtl2.DTPicker DTPYear 
         Height          =   330
         Left            =   2475
         TabIndex        =   22
         Top             =   900
         Width           =   1050
         _ExtentX        =   1852
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
         Format          =   16646147
         UpDown          =   -1  'True
         CurrentDate     =   38944
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   3240
         TabIndex        =   21
         Top             =   1530
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รายละเอียด :"
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
         Left            =   855
         TabIndex        =   20
         Top             =   2160
         Width           =   1410
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ครั้งที่ :"
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
         Left            =   1215
         TabIndex        =   19
         Top             =   1530
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ประจำปี :"
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
         Left            =   720
         TabIndex        =   18
         Top             =   900
         Width           =   1545
      End
   End
   Begin VB.PictureBox Pic102 
      Height          =   8160
      Left            =   0
      ScaleHeight     =   8100
      ScaleWidth      =   11295
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   11355
      Begin VB.CommandButton CMDShowDetail 
         Caption         =   "ดูข้อมูล"
         Height          =   420
         Left            =   405
         TabIndex        =   85
         Top             =   1170
         Width           =   1050
      End
      Begin VB.TextBox TextItemSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   330
         Left            =   8595
         TabIndex        =   84
         Top             =   540
         Width           =   1905
      End
      Begin VB.ComboBox CMBWHCode1 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   6030
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   540
         Width           =   1500
      End
      Begin VB.ComboBox CMBTimes1 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   3915
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   540
         Width           =   1230
      End
      Begin VB.ComboBox CMBYear1 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   540
         Width           =   1635
      End
      Begin VB.CommandButton CMDQuit 
         Caption         =   "ออก"
         Height          =   420
         Left            =   9630
         TabIndex        =   57
         Top             =   6390
         Width           =   870
      End
      Begin VB.CommandButton CMDSelect 
         Caption         =   "เลือก"
         Height          =   420
         Left            =   8640
         TabIndex        =   56
         Top             =   6390
         Width           =   870
      End
      Begin MSComctlLib.ListView ListViewID 
         Height          =   4560
         Left            =   450
         TabIndex        =   55
         Top             =   1665
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   8043
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ประจำปี"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ครั้งที่"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "คลัง"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ชั้นเก็บ"
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
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "นับได้"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "หน่วย"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label29 
         Caption         =   "รหัสสินค้า :"
         Height          =   375
         Left            =   7650
         TabIndex        =   80
         Top             =   540
         Width           =   870
      End
      Begin VB.Label Label28 
         Caption         =   "คลัง :"
         Height          =   330
         Left            =   5445
         TabIndex        =   79
         Top             =   540
         Width           =   465
      End
      Begin VB.Label Label27 
         Caption         =   "ครั้งที่ :"
         Height          =   330
         Left            =   3240
         TabIndex        =   78
         Top             =   540
         Width           =   555
      End
      Begin VB.Label Label26 
         Caption         =   "ประจำปี :"
         Height          =   375
         Left            =   450
         TabIndex        =   77
         Top             =   540
         Width           =   735
      End
   End
   Begin VB.CommandButton CMDClear 
      Caption         =   "ล้างหน้าจอ"
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
      Left            =   7290
      TabIndex        =   76
      Top             =   6300
      Width           =   1230
   End
   Begin VB.CommandButton CMDComplete 
      Caption         =   "สรุปบันทึกผลการตรวจนับ"
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
      Left            =   8550
      TabIndex        =   73
      Top             =   6300
      Width           =   1230
   End
   Begin VB.CommandButton CMDDelete 
      Caption         =   "ลบ"
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
      Left            =   4770
      TabIndex        =   53
      Top             =   6300
      Width           =   1230
   End
   Begin VB.ComboBox CMBTimes 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1620
      Width           =   1005
   End
   Begin VB.ComboBox CMBYear 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2250
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1620
      Width           =   1815
   End
   Begin VB.CommandButton CMDSearch 
      Caption         =   "ค้นหา"
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
      Left            =   6030
      TabIndex        =   52
      Top             =   6300
      Width           =   1230
   End
   Begin VB.CommandButton CMDEdit 
      Caption         =   "แก้ไข"
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
      Left            =   3510
      TabIndex        =   51
      Top             =   6300
      Width           =   1230
   End
   Begin VB.CommandButton CMDSave 
      Caption         =   "บันทึก"
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
      Left            =   2250
      TabIndex        =   50
      Top             =   6300
      Width           =   1230
   End
   Begin VB.TextBox Text111 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2250
      TabIndex        =   15
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox Text110 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   5760
      TabIndex        =   14
      Top             =   4950
      Width           =   4110
   End
   Begin VB.TextBox Text109 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2250
      TabIndex        =   13
      Top             =   4950
      Width           =   1815
   End
   Begin VB.TextBox Text108 
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
      Height          =   285
      Left            =   5760
      TabIndex        =   6
      Top             =   2655
      Width           =   4920
   End
   Begin VB.TextBox Text107 
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
      Height          =   285
      Left            =   2250
      TabIndex        =   5
      Top             =   2655
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPDocDate 
      Height          =   285
      Left            =   2250
      TabIndex        =   7
      Top             =   3105
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   16777215
      Format          =   16646145
      CurrentDate     =   38944
   End
   Begin VB.ComboBox CMBShelfCode 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4500
      Width           =   4110
   End
   Begin VB.TextBox Text106 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2250
      TabIndex        =   11
      Top             =   4500
      Width           =   1815
   End
   Begin VB.TextBox Text105 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   10
      Top             =   4050
      Width           =   4920
   End
   Begin VB.TextBox Text104 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2250
      TabIndex        =   9
      Top             =   4050
      Width           =   1815
   End
   Begin VB.TextBox Text103 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2250
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.ComboBox CMBRecProduct 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   5760
      TabIndex        =   4
      Text            =   "CMBRecProduct"
      Top             =   2070
      Width           =   2265
   End
   Begin VB.ComboBox CMBWHCode 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2250
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2070
      Width           =   1815
   End
   Begin VB.TextBox Text101 
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
      Height          =   285
      Left            =   2250
      TabIndex        =   0
      Top             =   1170
      Width           =   1140
   End
   Begin VB.CommandButton CMDStockRecord 
      Caption         =   "ทะเบียนการตรวจนับประจำปี"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   9585
      TabIndex        =   16
      Top             =   135
      Width           =   1590
   End
   Begin VB.Label Label20 
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
      Left            =   4050
      TabIndex        =   49
      Top             =   4050
      Width           =   1590
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อจัดซื้อ :"
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
      Height          =   375
      Left            =   4230
      TabIndex        =   48
      Top             =   2655
      Width           =   1410
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวนที่นับได้ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   510
      Left            =   180
      TabIndex        =   47
      Top             =   5400
      Width           =   1950
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ผู้ตรวจนับ :"
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
      Left            =   4095
      TabIndex        =   46
      Top             =   4950
      Width           =   1545
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ทีม :"
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
      Height          =   375
      Left            =   540
      TabIndex        =   45
      Top             =   4950
      Width           =   1590
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสจัดซื้อ :"
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
      Left            =   765
      TabIndex        =   44
      Top             =   2655
      Width           =   1365
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ตรวจนับ :"
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
      Height          =   375
      Left            =   585
      TabIndex        =   43
      Top             =   3105
      Width           =   1545
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชั้นเก็บสต๊อก :"
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
      Left            =   4545
      TabIndex        =   42
      Top             =   4500
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   1080
      TabIndex        =   41
      Top             =   4500
      Width           =   1050
   End
   Begin VB.Label Label11 
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
      Height          =   330
      Left            =   1080
      TabIndex        =   40
      Top             =   4050
      Width           =   1050
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BarCode :"
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
      TabIndex        =   39
      Top             =   3600
      Width           =   1140
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ตำแหน่งชั้นเก็บ :"
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
      Height          =   375
      Left            =   3915
      TabIndex        =   38
      Top             =   2115
      Width           =   1725
   End
   Begin VB.Label Label8 
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
      Height          =   330
      Left            =   1080
      TabIndex        =   37
      Top             =   2070
      Width           =   1050
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ครั้งที่ :"
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
      Left            =   4590
      TabIndex        =   36
      Top             =   1620
      Width           =   1050
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประจำปี :"
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
      Left            =   1080
      TabIndex        =   35
      Top             =   1620
      Width           =   1050
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   1080
      TabIndex        =   34
      Top             =   1170
      Width           =   1050
   End
End
Attribute VB_Name = "Form108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vIsOpen As Integer


Private Sub Check101_Click()
Dim i As Integer

On Error Resume Next

If ListView103.ListItems.Count > 0 Then
  If Check101.Value = 1 Then
    For i = 1 To ListView103.ListItems.Count
    ListView103.ListItems.Item(i).Checked = True
    Next i
  End If
  If Check101.Value = 0 Then
    For i = 1 To ListView103.ListItems.Count
    ListView103.ListItems.Item(i).Checked = False
    Next i
  End If
  

End If
End Sub

Private Sub CMBWHCode_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vWHCode As String

On Error Resume Next

vWHCode = Trim(CMBWHCode.Text)
CMBRecProduct.Clear
vQuery = "exec dbo.USP_MB_SearchShelfCode '" & vWHCode & "'  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        CMBRecProduct.AddItem Trim(vRecordset.Fields("shelfcode").Value)
        vRecordset.MoveNext
        Wend
End If
vRecordset.Close
End Sub


Private Sub CMBYear_Click()
On Error Resume Next

Call GetTimesDetail
End Sub

Private Sub CMD102_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As String
Dim vTimes As Integer
Dim vMyDescription As String

On Error GoTo ErrDescription

If TextTimes.Text <> "" Then
  vYear = DTPYear.Year
  vTimes = TextTimes.Text
  vMyDescription = TextDescription.Text
  vQuery = "exec dbo.USP_ISP_UpdateAnnuallyMaster '" & vYear & "'," & vTimes & ",'" & vMyDescription & "' "
  gConnection.Execute vQuery
  MsgBox "บันทึกข้อมูลเรียบร้อยแล้ว", vbInformation, "Send Message"
  Call GetTimes
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As String
Dim vListItem As ListItem

On Error GoTo ErrDescription

vYear = DTPYear.Year
ListView101.ListItems.Clear
vQuery = "exec dbo.USP_ISP_SearchAnnuallyMaster '" & vYear & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst
While Not vRecordset.EOF
  Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("id").Value))
  vListItem.SubItems(1) = Trim(vRecordset.Fields("annually").Value)
  vListItem.SubItems(2) = Trim(vRecordset.Fields("times").Value)
  vListItem.SubItems(3) = Trim(vRecordset.Fields("inspectdescription").Value)
vRecordset.MoveNext
Wend
End If
vRecordset.Close
PIC101.Visible = True
TextDescription.Text = ""

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD104_Click()
Frame101.Visible = False
End Sub

Private Sub CMD105_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItem As ListItem
Dim vAnnually As String
Dim vTimes As Integer
Dim vWHCode As String

On Error GoTo ErrDescription

If CMBAnnually.Text <> "" And CMBID.Text <> "" And CMBWareHouse.Text <> "" Then
  vAnnually = CMBAnnually.Text
  vTimes = CMBID.Text
  vWHCode = CMBWareHouse.Text
  ListView102.ListItems.Clear
  vQuery = "exec dbo.USP_ISP_InspectProcessWaitting '" & vAnnually & "'," & vTimes & ",'" & vWHCode & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
      Set vListItem = ListView102.ListItems.Add(, , Trim(vRecordset.Fields("buyername").Value))
    vRecordset.MoveNext
    Wend
  End If
  vRecordset.Close
  
  ListView103.ListItems.Clear
  vQuery = "exec dbo.USP_ISP_InspectProcessRecCK '" & vAnnually & "'," & vTimes & ",'" & vWHCode & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
      Set vListItem = ListView103.ListItems.Add(, , Trim(vRecordset.Fields("buyername").Value))
      vListItem.SubItems(1) = Trim(vRecordset.Fields("stkinspectcode").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("stkadjustcode").Value)
    vRecordset.MoveNext
    Wend
  End If
  vRecordset.Close
  MsgBox "ฟื้นฟูข้อมูลตรวจนับเรียบร้อยแล้ว", vbInformation, "Send Information"
Else
MsgBox "กรอกข้อมูลประจำปี คลัง และ ครั้งที่ก่อนกดประมวลผลทุกครั้ง", vbInformation, "Send Information"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Cmd106_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType As String
Dim vReportName As String
Dim vDocNo As String
Dim i As Integer

On Error GoTo ErrDescription

vRepID = 213
vRepType = "IV"

If ListView103.ListItems.Count > 0 Then
  For i = 1 To ListView103.ListItems.Count
    If ListView103.ListItems.Item(i).Checked = True Then
      vDocNo = Trim(ListView103.ListItems.Item(i).SubItems(1))
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
      Next i
      ListView103.ListItems.Clear
    End If



ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD107_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItem As ListItem
Dim vAnnually As String
Dim vTimes As Integer
Dim vWHCode As String

On Error GoTo ErrDescription

If CMBAnnually.Text <> "" And CMBID.Text <> "" And CMBWareHouse.Text <> "" Then
  vAnnually = CMBAnnually.Text
  vTimes = CMBID.Text
  vWHCode = CMBWareHouse.Text

  vQuery = "exec dbo.USP_ISP_InspectProcessRec '" & vAnnually & "'," & vTimes & ",'" & vWHCode & "' "
  gConnection.Execute vQuery
  
  Call CMD105_Click
  
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD108_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItem As ListItem
Dim vAnnually As String
Dim vTimes As Integer
Dim vWHCode As String
Dim vAnswer As Integer

On Error GoTo ErrDescription

vAnswer = MsgBox("คุณต้องการยกเลิกเอกสารการตรวจนับใช่หรือไม่", vbYesNo, "Message Question")
If vAnswer = 6 Then
  If CMBAnnually.Text <> "" And CMBID.Text <> "" And CMBWareHouse.Text <> "" Then
    vAnnually = CMBAnnually.Text
    vTimes = CMBID.Text
    vWHCode = CMBWareHouse.Text
  
    vQuery = "exec dbo.USP_ISP_InspectProcessRecCancel '" & vAnnually & "'," & vTimes & ",'" & vWHCode & "' "
    gConnection.Execute vQuery
    
    Call CMD105_Click
  End If
Else
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD109_Click()
ListView102.ListItems.Clear
ListView103.ListItems.Clear
Pic103.Visible = False
End Sub

Private Sub CMDClear_Click()
On Error Resume Next

CMBYear.Text = "2006"
CMBTimes.Text = "1"
CMBWHCode.Text = ""
CMBRecProduct.Text = ""
Text103.Text = ""
Text104.Text = ""
Text105.Text = ""
Text106.Text = ""
CMBShelfCode.Text = ""
DTPDocDate = Now
Text107.Text = ""
Text108.Text = ""
Text109.Text = ""
Text110.Text = ""
Text111.Text = ""
vIsOpen = 0
End Sub

Private Sub CMDComplete_Click()
If vUserID = "vilaiwan" Or vUserID = "opporn" Or vUserID = "surachai" Or vUserID = "nueng" Or vUserID = "somrod" Or vUserID = "sa" Then
  Pic103.Visible = True
Else
  MsgBox "คุณไม่มีสิทธิ์สรุปผลการตรวจนับ", vbCritical, "Send Message"
  Exit Sub
End If
End Sub

Private Sub CMDDelete_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As Integer
Dim vAnswer As Integer
Dim vAnnually As Integer
Dim vTimes As Integer
Dim vWHCode As String

On Error GoTo ErrDescription

If vIsOpen = 1 Then
  vID = Trim(Text101.Text)
  vAnnually = Trim(CMBYear.Text)
  vTimes = Trim(CMBTimes.Text)
  vWHCode = Trim(CMBWHCode.Text)
  
  vQuery = "exec dbo.USP_ISP_CheckPrompInsertInspectRec " & vAnnually & "," & vTimes & ",'" & vWHCode & "'  "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    If vRecordset.Fields("adjustcount").Value > 0 Then
    MsgBox "ไม่สมารถเพิ่มข้อมูลการตรวจนับคลัง " & vWHCode & " นี้ได้อีก เพราะได้ทำการคำนวณผลการตรวจนับคลังนี้ไปแล้ว นอกจากจะไปยกเลิกเอกสารการตรวจนับคลังดังกล่าวก่อน แล้วมาบันทึกข้อมูลการตรวจนับอีกครั้ง", vbCritical, "Send Error"
    Exit Sub
    End If
  End If
  vRecordset.Close
  
  vAnswer = MsgBox("คุณต้องการลบเอกสารเลขที่ " & vID & " นี้ใช่หรือไม่", vbYesNo, "Send Question ?")
  If vAnswer = 6 Then
    vQuery = "exec dbo.USP_ISP_DelectInspectRec " & vID & " "
    gConnection.Execute vQuery
    vIsOpen = 0
    Call ClearScreen
  End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDEdit_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAnnually As String
Dim vTimes As String
Dim vWHCode As String
Dim vRecproduct As String
Dim vBarCode As String
Dim vItemCode As String
Dim vItemName As String
Dim vUnitCode As String
Dim vShelfCode As String
Dim vDocDate As Date
Dim vBuyerID As String
Dim vBuyerName As String
Dim vTeam As String
Dim vChecker As String
Dim vCount As Currency
Dim vID As Integer


If vIsOpen = 1 Then

If CMBYear.Text <> "" And CMBTimes.Text <> "" And CMBWHCode.Text <> "" And CMBRecProduct.Text <> "" And CMBShelfCode.Text <> "" And Text104.Text <> "" And Text111.Text <> "" Then
  vID = Trim(Text101.Text)
  vAnnually = Trim(CMBYear.Text)
  vTimes = Trim(CMBTimes.Text)
  vWHCode = Trim(CMBWHCode.Text)
  vRecproduct = Trim(CMBRecProduct.Text)
  vBarCode = Trim(Text103.Text)
  vItemCode = Trim(Text104.Text)
  vItemName = Trim(Text105.Text)
  vUnitCode = Trim(Text106.Text)
  vShelfCode = Trim(CMBShelfCode.Text)
  vDocDate = CDate(DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year)
  vBuyerID = Trim(Text107.Text)
  vBuyerName = Trim(Text108.Text)
  vTeam = Trim(Text109.Text)
  vChecker = Trim(Text110.Text)
  vCount = Format(Trim(Text111.Text), "##,##0.000")

  vQuery = "exec dbo.USP_ISP_CheckPrompInsertInspectRec " & vAnnually & "," & vTimes & ",'" & vWHCode & "'  "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    If vRecordset.Fields("adjustcount").Value > 0 Then
    MsgBox "ไม่สมารถเพิ่มข้อมูลการตรวจนับคลัง " & vWHCode & " นี้ได้อีก เพราะได้ทำการคำนวณผลการตรวจนับคลังนี้ไปแล้ว นอกจากจะไปยกเลิกเอกสารการตรวจนับคลังดังกล่าวก่อน แล้วมาบันทึกข้อมูลการตรวจนับอีกครั้ง", vbCritical, "Send Error"
    Exit Sub
    End If
  End If
  vRecordset.Close
  
On Error GoTo ErrDescription

  vQuery = "begin tran"
  gConnection.Execute vQuery
  vQuery = "exec dbo.USP_ISP_EditInspectRec " & vID & ",'" & vAnnually & "','" & vTimes & "','" & vWHCode & "','" & vRecproduct & "','" & vItemCode & "','" & vUnitCode & "','" & vShelfCode & "','" & vDocDate & "','" & vTeam & "','" & vChecker & "'," & vCount & " "
  gConnection.Execute vQuery
  vQuery = "commit tran"
  gConnection.Execute vQuery
  
  MsgBox "แก้ไขข้อมูลเรียบร้อยแล้ว", vbInformation, "Send Message"
  Call ClearData
  Text103.SetFocus
  vIsOpen = 0
Else
MsgBox "กรอกข้อมูลให้ครบก่อนกดแก้ไขผลทุกครั้ง", vbInformation, "Send Information"
End If

ErrDescription:
If Err.Description <> "" Then
  MsgBox Err.Description
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  Exit Sub
  End If
  End If
End Sub

Private Sub CMDExit_Click()
PIC101.Visible = False
End Sub

Private Sub CMDQuit_Click()
Pic102.Visible = False
End Sub

Private Sub CMDSave_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAnnually As String
Dim vTimes As String
Dim vWHCode As String
Dim vRecproduct As String
Dim vBarCode As String
Dim vItemCode As String
Dim vItemName As String
Dim vUnitCode As String
Dim vShelfCode As String
Dim vDocDate As Date
Dim vBuyerID As String
Dim vBuyerName As String
Dim vTeam As String
Dim vChecker As String
Dim vCount As Currency

If vIsOpen = 0 Then
If CMBYear.Text <> "" And CMBTimes.Text <> "" And CMBWHCode.Text <> "" And CMBRecProduct.Text <> "" And CMBShelfCode.Text <> "" And Text104.Text <> "" And Text111.Text <> "" Then
  vAnnually = Trim(CMBYear.Text)
  vTimes = Trim(CMBTimes.Text)
  vWHCode = Trim(CMBWHCode.Text)
  vRecproduct = Trim(CMBRecProduct.Text)
  vBarCode = Trim(Text103.Text)
  vItemCode = Trim(Text104.Text)
  vItemName = Trim(Text105.Text)
  vUnitCode = Trim(Text106.Text)
  vShelfCode = Trim(CMBShelfCode.Text)
  vDocDate = CDate(DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year)
  vBuyerID = Trim(Text107.Text)
  vBuyerName = Trim(Text108.Text)
  vTeam = Trim(Text109.Text)
  vChecker = Trim(Text110.Text)
  vCount = Format(Trim(Text111.Text), "##,##0.000")

  vQuery = "exec dbo.USP_ISP_CheckPrompInsertInspectRec " & vAnnually & "," & vTimes & ",'" & vWHCode & "'  "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    If vRecordset.Fields("adjustcount").Value > 0 Then
    MsgBox "ไม่สมารถเพิ่มข้อมูลการตรวจนับคลัง " & vWHCode & " นี้ได้อีก เพราะได้ทำการคำนวณผลการตรวจนับคลังนี้ไปแล้ว นอกจากจะไปยกเลิกเอกสารการตรวจนับคลังดังกล่าวก่อน แล้วมาบันทึกข้อมูลการตรวจนับอีกครั้ง", vbCritical, "Send Error"
    Exit Sub
    End If
  End If
  vRecordset.Close

On Error GoTo ErrDescription

  vQuery = "begin tran"
  gConnection.Execute vQuery
  
  
  vQuery = "exec dbo.USP_ISP_InsertInspectRec '" & vAnnually & "','" & vTimes & "','" & vWHCode & "','" & vRecproduct & "','" & vItemCode & "','" & vUnitCode & "','" & vShelfCode & "','" & vDocDate & "','" & vTeam & "','" & vChecker & "'," & vCount & " "
  gConnection.Execute vQuery
  vQuery = "commit tran"
  gConnection.Execute vQuery
  
    
  Call ClearData
  Text103.SetFocus
Else
MsgBox "กรอกข้อมูลให้ครบก่อนกดบันทึกผลทุกครั้ง", vbInformation, "Send Information"
End If

ErrDescription:
If Err.Description <> "" Then
  MsgBox Err.Description
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  Exit Sub
End If
End If
End Sub

Private Sub CMDSearch_Click()
On Error Resume Next

Pic102.Visible = True
ListViewID.ListItems.Clear
TextItemSearch.Text = ""
CMBYear1.Text = "2007"
CMBTimes1.Text = "1"
CMBWHCode1.Text = ""
End Sub

Private Sub CMDSelect_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vYear As String
Dim vTimes As String
Dim vWHCode As String
Dim vItemCode As String
Dim vListItem As ListItem
Dim vIndex As Integer
Dim vShelfCode As String

On Error GoTo ErrDescription

If ListViewID.ListItems.Count > 0 Then
  vIndex = ListViewID.SelectedItem.Index
  vYear = Trim(ListViewID.ListItems.Item(vIndex).SubItems(1))
  vTimes = Trim(ListViewID.ListItems.Item(vIndex).SubItems(2))
  vWHCode = Trim(ListViewID.ListItems.Item(vIndex).SubItems(3))
  vItemCode = Trim(ListViewID.ListItems.Item(vIndex).SubItems(5))
  vShelfCode = Trim(ListViewID.ListItems.Item(vIndex).SubItems(4))
  
  vQuery = "exec dbo.USP_ISP_SearchInspectRec '" & vYear & "','" & vTimes & "','" & vWHCode & "','" & vItemCode & "' ,'','" & vShelfCode & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    Text101.Text = Trim(vRecordset.Fields("id").Value)
    Text104.Text = Trim(vRecordset.Fields("itemcode").Value)
    Text105.Text = Trim(vRecordset.Fields("itemname").Value)
    Text106.Text = Trim(vRecordset.Fields("defstkunitcode").Value)
    Text107.Text = Trim(vRecordset.Fields("buyercode").Value)
    Text108.Text = Trim(vRecordset.Fields("buyername").Value)
    Text109.Text = Trim(vRecordset.Fields("team").Value)
    Text110.Text = Trim(vRecordset.Fields("inspector").Value)
    Text111.Text = Trim(vRecordset.Fields("inspectqty").Value)
    DTPDocDate = Trim(vRecordset.Fields("inspectdate").Value)
    CMBYear.Text = Trim(vRecordset.Fields("annually").Value)
    CMBTimes.Text = Trim(vRecordset.Fields("times").Value)
    CMBWHCode.Text = Trim(vRecordset.Fields("whcode").Value)
    CMBRecProduct.Text = Trim(vRecordset.Fields("positionshelf").Value)
    CMBShelfCode.Text = Trim(vRecordset.Fields("stockshelf").Value)
  End If
  vRecordset.Close
  vIsOpen = 1
  Pic102.Visible = False
Else
  MsgBox "ไม่มีข้อมูลในการเลือก", vbCritical, "Send Error"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDShowDetail_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vYear As String
Dim vTimes As String
Dim vWHCode As String
Dim vItemCode As String
Dim vListItem As ListItem

On Error GoTo ErrDescription

If CMBYear1.Text <> "" And CMBTimes1.Text <> "" And CMBWHCode1.Text <> "" And TextItemSearch.Text <> "" Then
  vYear = CMBYear1.Text
  vTimes = CMBTimes1.Text
  vWHCode = CMBWHCode1.Text
  vItemCode = TextItemSearch.Text
  
  vQuery = "exec dbo.USP_ISP_SearchInspectRec '" & vYear & "','" & vTimes & "','" & vWHCode & "','" & vItemCode & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    ListViewID.ListItems.Clear
    While Not vRecordset.EOF
      Set vListItem = ListViewID.ListItems.Add(, , Trim(vRecordset.Fields("id").Value))
      vListItem.SubItems(1) = Trim(vRecordset.Fields("annually").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("times").Value)
      vListItem.SubItems(3) = Trim(vRecordset.Fields("whcode").Value)
      vListItem.SubItems(4) = Trim(vRecordset.Fields("positionshelf").Value)
      vListItem.SubItems(5) = Trim(vRecordset.Fields("itemcode").Value)
      vListItem.SubItems(6) = Trim(vRecordset.Fields("itemname").Value)
      vListItem.SubItems(7) = Trim(vRecordset.Fields("inspectqty").Value)
      vListItem.SubItems(8) = Trim(vRecordset.Fields("DefStkUnitCode").Value)
    vRecordset.MoveNext
    Wend
  End If
  vRecordset.Close

Else
  MsgBox "ต้องกรอกข้อมูลปี,ครั้งที่,คลัง และรหัสสินค้าให้ครบด้วยในการค้นหา", vbCritical, "Send Error"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDStockRecord_Click()
Frame101.Visible = True
End Sub

Private Sub Form_Load()
DTPYear.CustomFormat = "yyyy"
DTPDocDate = Now
Call GetTimes
Call GetWHCode
Call ShelfCode
Call GetAnnually
End Sub

Public Sub GetWHCode()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

CMBWHCode.Clear
CMBWareHouse.Clear
vQuery = "exec dbo.USP_MB_SearchWhCodeCode "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        CMBWHCode.AddItem Trim(vRecordset.Fields("whcode").Value)
        CMBWHCode1.AddItem Trim(vRecordset.Fields("whcode").Value)
        CMBWareHouse.AddItem Trim(vRecordset.Fields("whcode").Value)
        vRecordset.MoveNext
        Wend
End If
vRecordset.Close
End Sub
Public Sub GetAnnually()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

CMBYear.Clear
CMBAnnually.Clear
vQuery = "exec dbo.USP_ISP_AnnuallyMasterYearList "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        CMBYear.AddItem Trim(vRecordset.Fields("annually").Value)
        CMBYear1.AddItem Trim(vRecordset.Fields("annually").Value)
        CMBAnnually.AddItem Trim(vRecordset.Fields("annually").Value)
        vRecordset.MoveNext
        Wend
End If
vRecordset.Close
CMBYear.Text = DTPDocDate.Year
CMBAnnually.Text = DTPDocDate.Year
End Sub

Public Sub GetTimesDetail()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vYear As String

On Error Resume Next

vYear = Trim(CMBYear.Text)
CMBTimes.Clear
CMBID.Clear
vQuery = "exec dbo.USP_ISP_AnnuallyMasterTimesList '" & vYear & "'"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        CMBTimes.AddItem Trim(vRecordset.Fields("times").Value)
        CMBTimes1.AddItem Trim(vRecordset.Fields("times").Value)
        CMBID.AddItem Trim(vRecordset.Fields("times").Value)
        vRecordset.MoveNext
        Wend
End If
vRecordset.Close
CMBTimes.Text = "1"
End Sub

Private Sub ListViewID_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vYear As String
Dim vTimes As String
Dim vWHCode As String
Dim vItemCode As String
Dim vListItem As ListItem
Dim vIndex As Integer
Dim vShelfCode As String

On Error GoTo ErrDescription

If ListViewID.ListItems.Count > 0 Then
  vIndex = ListViewID.SelectedItem.Index
  vYear = Trim(ListViewID.ListItems.Item(vIndex).SubItems(1))
  vTimes = Trim(ListViewID.ListItems.Item(vIndex).SubItems(2))
  vWHCode = Trim(ListViewID.ListItems.Item(vIndex).SubItems(3))
  vItemCode = Trim(ListViewID.ListItems.Item(vIndex).SubItems(5))
  vShelfCode = Trim(ListViewID.ListItems.Item(vIndex).SubItems(4))
  
  vQuery = "exec dbo.USP_ISP_SearchInspectRec '" & vYear & "','" & vTimes & "','" & vWHCode & "','" & vItemCode & "' ,'','" & vShelfCode & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    Text101.Text = Trim(vRecordset.Fields("id").Value)
    Text104.Text = Trim(vRecordset.Fields("itemcode").Value)
    Text105.Text = Trim(vRecordset.Fields("itemname").Value)
    Text106.Text = Trim(vRecordset.Fields("defstkunitcode").Value)
    Text107.Text = Trim(vRecordset.Fields("buyercode").Value)
    Text108.Text = Trim(vRecordset.Fields("buyername").Value)
    Text109.Text = Trim(vRecordset.Fields("team").Value)
    Text110.Text = Trim(vRecordset.Fields("inspector").Value)
    Text111.Text = Trim(vRecordset.Fields("inspectqty").Value)
    DTPDocDate = Trim(vRecordset.Fields("inspectdate").Value)
    CMBYear.Text = Trim(vRecordset.Fields("annually").Value)
    CMBTimes.Text = Trim(vRecordset.Fields("times").Value)
    CMBWHCode.Text = Trim(vRecordset.Fields("whcode").Value)
    CMBRecProduct.Text = Trim(vRecordset.Fields("positionshelf").Value)
    CMBShelfCode.Text = Trim(vRecordset.Fields("stockshelf").Value)
  End If
  vRecordset.Close
  vIsOpen = 1
  Pic102.Visible = False
Else
  MsgBox "ไม่มีข้อมูลในการเลือก", vbCritical, "Send Error"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text103_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery  As String
Dim vBarCode As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
  If Text103.Text <> "" Then
    vBarCode = Trim(Text103.Text)
    vQuery = "exec dbo.USP_ISP_SearchProduct '" & vBarCode & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Text104.Text = Trim(vRecordset.Fields("itemcode").Value)
      Text105.Text = Trim(vRecordset.Fields("itemname").Value)
      Text106.Text = Trim(vRecordset.Fields("defstkunitcode").Value)
    Else
      MsgBox "ไม่มีบาร์โค้ด " & vBarCode & " นี้ในระบบ ", vbCritical, "Send Error"
      Text104.Text = ""
      Text105.Text = ""
      Text106.Text = ""
      Exit Sub
    End If
    vRecordset.Close
  End If
  If Text109.Text = "" And Text110.Text = "" Then
    Text109.SetFocus
  End If
  If Text109.Text <> "" And Text110.Text = "" Then
    Text110.SetFocus
  End If
  If Text109.Text <> "" And Text110.Text <> "" Then
    Text111.SetFocus
  End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Private Sub Text104_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery  As String
Dim vBarCode As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
  If Text104.Text <> "" Then
    vBarCode = Trim(Text104.Text)
    vQuery = "exec dbo.USP_ISP_SearchProduct '" & vBarCode & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Text105.Text = Trim(vRecordset.Fields("itemname").Value)
      Text106.Text = Trim(vRecordset.Fields("defstkunitcode").Value)
    Else
      MsgBox "ไม่มีรหัสสินค้า " & vBarCode & " นี้ในระบบ ", vbCritical, "Send Error"
      Text105.Text = ""
      Text106.Text = ""
      Exit Sub
    End If
    vRecordset.Close
  End If
  If Text109.Text = "" And Text110.Text = "" Then
    Text109.SetFocus
  End If
  If Text109.Text <> "" And Text110.Text = "" Then
    Text110.SetFocus
  End If
  If Text109.Text <> "" And Text110.Text <> "" Then
    Text111.SetFocus
  End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Private Sub Text109_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Text110.SetFocus
End If
End Sub

Private Sub Text110_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Text111.SetFocus
End If
End Sub

Private Sub Text111_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call CMDSave_Click
End If
End Sub

Private Sub TextItemSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call CMDShowDetail_Click
End If
End Sub

Private Sub UpDown1_DownClick()
If TextTimes.Text > 1 Then
  TextTimes.Text = TextTimes.Text - 1
Else
  MsgBox "จำนวนครั้งต้องมากกว่า 0", vbCritical, "Send Error"
End If
End Sub

Private Sub UpDown1_UpClick()
TextTimes.Text = TextTimes.Text + 1
End Sub

Public Sub GetTimes()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As String

vYear = DTPYear.Year
vQuery = "select isnull(max(times),1)+1  as maxTimes from npmaster.dbo.TB_ISP_AnnuallyMaster where annually = '" & vYear & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  TextTimes.Text = Trim(vRecordset.Fields("maxtimes").Value)
End If
vRecordset.Close
End Sub

Public Sub ShelfCode()
Dim vRecordset  As New ADODB.Recordset
Dim vQuery As String


'vQuery = "select  distinct (code+'//'+name) as shelfcode from dbo.bcshelf order   by shelfcode"
vQuery = "select  distinct code as shelfcode from dbo.bcshelf order   by shelfcode"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMBShelfCode.AddItem Trim(vRecordset.Fields("shelfcode").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close
CMBShelfCode.Text = Trim("AVL")
End Sub

Public Sub ClearScreen()
Text101.Text = ""
Text103.Text = ""
Text104.Text = ""
Text105.Text = ""
Text106.Text = ""
Text107.Text = ""
Text108.Text = ""
Text109.Text = ""
Text110.Text = ""
Text111.Text = ""
DTPDocDate = Now
CMBYear = DTPDocDate.Year
CMBTimes.Text = "1"

End Sub

Public Sub ClearData()
Text101.Text = ""
Text103.Text = ""
Text104.Text = ""
Text105.Text = ""
Text106.Text = ""
Text111.Text = ""
Text110.Text = ""
End Sub
