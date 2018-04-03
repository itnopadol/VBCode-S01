VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormPickingRequest 
   BackColor       =   &H00808080&
   Caption         =   "Picking Request App"
   ClientHeight    =   9600
   ClientLeft      =   2040
   ClientTop       =   675
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormPickingRequest.frx":0000
   ScaleHeight     =   15812.31
   ScaleMode       =   0  'User
   ScaleWidth      =   24309.41
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICPickReq 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   11535
      Left            =   0
      ScaleHeight     =   11505
      ScaleWidth      =   30045
      TabIndex        =   83
      Top             =   0
      Visible         =   0   'False
      Width           =   30075
      Begin VB.PictureBox PICPRSearchSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14655
         TabIndex        =   131
         Top             =   45
         Visible         =   0   'False
         Width           =   14685
         Begin VB.CommandButton CMDPRPICPRSearchSaleClose 
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
            Height          =   1005
            Left            =   10710
            Style           =   1  'Graphical
            TabIndex        =   173
            Top             =   7515
            Width           =   2400
         End
         Begin MSComctlLib.ListView ListViewPRSearchSale 
            Height          =   5055
            Left            =   1845
            TabIndex        =   169
            Top             =   2295
            Width           =   11265
            _ExtentX        =   19870
            _ExtentY        =   8916
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "รหัสพนักงาน"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ชื่อพนักงาน"
               Object.Width           =   8819
            EndProperty
         End
         Begin VB.CommandButton CMDPRSearchSaleClick 
            Height          =   375
            Left            =   6210
            Picture         =   "FormPickingRequest.frx":9D15
            Style           =   1  'Graphical
            TabIndex        =   168
            Top             =   1350
            Width           =   375
         End
         Begin VB.TextBox TBPRSearchSale 
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
            Left            =   1800
            TabIndex        =   167
            Top             =   1350
            Width           =   4380
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "คำที่ค้นหา :"
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
            Left            =   405
            TabIndex        =   166
            Top             =   1350
            Width           =   1275
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "ค้นหา พนักงานขาย"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   180
            TabIndex        =   165
            Top             =   180
            Width           =   2130
         End
      End
      Begin VB.PictureBox PICLastSendQue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14655
         TabIndex        =   177
         Top             =   45
         Visible         =   0   'False
         Width           =   14685
         Begin VB.CommandButton CMDSendQuePrint 
            BackColor       =   &H00808080&
            Caption         =   "พิมพ์ทดแทน"
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
            Left            =   9540
            Style           =   1  'Graphical
            TabIndex        =   205
            Top             =   7155
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.CommandButton CMDSendQueExit 
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
            Left            =   11565
            Style           =   1  'Graphical
            TabIndex        =   179
            Top             =   7155
            Width           =   1950
         End
         Begin MSComctlLib.ListView ListViewLastSendQue 
            Height          =   5235
            Left            =   1080
            TabIndex        =   178
            Top             =   1755
            Width           =   12435
            _ExtentX        =   21934
            _ExtentY        =   9234
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "คิวที่"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "คำอธิบาย"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "รหัส/ชื่อสินค้า"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "ต้องการ"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "จัดได้"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "หน่วยนับ"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "พนักงานจัด"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "โซน"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "วันที่คิว"
               Object.Width           =   2646
            EndProperty
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "รายการ คิวจัดสินค้าล่าสุดของเอกสารนี้"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1080
            TabIndex        =   192
            Top             =   1170
            Width           =   14685
         End
      End
      Begin VB.PictureBox PICPRSearchAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14655
         TabIndex        =   130
         Top             =   45
         Visible         =   0   'False
         Width           =   14685
         Begin VB.CommandButton CMDPRPICPRSearchARClose 
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
            Height          =   870
            Left            =   11160
            Style           =   1  'Graphical
            TabIndex        =   172
            Top             =   7920
            Width           =   2040
         End
         Begin MSComctlLib.ListView ListViewPRSearchAR 
            Height          =   5640
            Left            =   1710
            TabIndex        =   164
            Top             =   2070
            Width           =   11490
            _ExtentX        =   20267
            _ExtentY        =   9948
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "รหัสลูกค้า"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ชื่อลูกค้า"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "รหัสสมาชิก"
               Object.Width           =   4410
            EndProperty
         End
         Begin VB.CommandButton CMDPRSearchARClick 
            Height          =   375
            Left            =   6705
            Picture         =   "FormPickingRequest.frx":A0E2
            Style           =   1  'Graphical
            TabIndex        =   163
            Top             =   1170
            Width           =   375
         End
         Begin VB.TextBox TBPRSearchAR 
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
            Left            =   2790
            TabIndex        =   162
            Top             =   1170
            Width           =   3885
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "คำที่ค้นหา :"
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
            Left            =   1170
            TabIndex        =   161
            Top             =   1170
            Width           =   1500
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "ค้นหา ลูกค้า"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   405
            TabIndex        =   160
            Top             =   270
            Width           =   1320
         End
      End
      Begin VB.PictureBox PICPRSearchDocNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14655
         TabIndex        =   129
         Top             =   45
         Visible         =   0   'False
         Width           =   14685
         Begin VB.CommandButton CMDPRPICPRSearchDocNoClose 
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
            Height          =   915
            Left            =   10935
            Style           =   1  'Graphical
            TabIndex        =   171
            Top             =   7605
            Width           =   2355
         End
         Begin MSComctlLib.ListView ListViewPRSearchDocNo 
            Height          =   4875
            Left            =   1620
            TabIndex        =   159
            Top             =   2520
            Width           =   11670
            _ExtentX        =   20585
            _ExtentY        =   8599
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777152
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "เลขที่เอกสาร"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "วันที่เอกสาร"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "ลูกค้า"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "พนักงานขาย"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "มูลค่า"
               Object.Width           =   2822
            EndProperty
         End
         Begin VB.CommandButton CMDPRSearchDocNoClick 
            Height          =   375
            Left            =   5715
            Picture         =   "FormPickingRequest.frx":A4AF
            Style           =   1  'Graphical
            TabIndex        =   158
            Top             =   1305
            Width           =   375
         End
         Begin VB.TextBox TBPRSearchDocNo 
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
            Height          =   375
            Left            =   1665
            TabIndex        =   157
            Top             =   1305
            Width           =   3975
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "คำที่ค้นหา :"
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
            Left            =   180
            TabIndex        =   156
            Top             =   1305
            Width           =   1365
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "ค้นหา เลขที่เอกสาร"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   495
            TabIndex        =   155
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.PictureBox PICSendInformation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14655
         TabIndex        =   180
         Top             =   45
         Visible         =   0   'False
         Width           =   14685
         Begin VB.CommandButton CMDInfClose 
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
            Height          =   870
            Left            =   11430
            Style           =   1  'Graphical
            TabIndex        =   182
            Top             =   7380
            Width           =   2130
         End
         Begin MSComctlLib.ListView ListViewInfQue 
            Height          =   4380
            Left            =   1080
            TabIndex        =   181
            Top             =   2790
            Width           =   12480
            _ExtentX        =   22013
            _ExtentY        =   7726
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "คิวที่"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "โซน"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ลูกค้า :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1215
            TabIndex        =   187
            Top             =   2025
            Width           =   1140
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "เลขที่เอกสาร :"
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
            Left            =   495
            TabIndex        =   186
            Top             =   1485
            Width           =   1860
         End
         Begin VB.Label LBLInfARName 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   2475
            TabIndex        =   185
            Top             =   2025
            Width           =   8610
         End
         Begin VB.Label LBLInfDocNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   2475
            TabIndex        =   184
            Top             =   1485
            Width           =   2355
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "รายการ เลขที่คิวจัดสินค้า"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1035
            TabIndex        =   183
            Top             =   585
            Width           =   2940
         End
      End
      Begin VB.PictureBox PICPRSearchItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14655
         TabIndex        =   128
         Top             =   45
         Visible         =   0   'False
         Width           =   14685
         Begin VB.PictureBox PicPoint1 
            Height          =   195
            Left            =   0
            ScaleHeight     =   135
            ScaleWidth      =   405
            TabIndex        =   208
            Top             =   0
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.CommandButton CMDPRPICPRSearchItemClose 
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
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   170
            Top             =   8100
            Width           =   2130
         End
         Begin VB.CommandButton CMDPRSearchItemClick 
            Height          =   375
            Left            =   6300
            Picture         =   "FormPickingRequest.frx":A87C
            Style           =   1  'Graphical
            TabIndex        =   136
            Top             =   1395
            Width           =   375
         End
         Begin MSComctlLib.ListView ListViewPRSearchItem 
            Height          =   5730
            Left            =   900
            TabIndex        =   135
            Top             =   2250
            Width           =   13290
            _ExtentX        =   23442
            _ExtentY        =   10107
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "รหัสสินค้า"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ชื่อสินค้า"
               Object.Width           =   12347
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "คงเหลือ"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "หน่วย"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "ราคาต่อหน่วย"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "โซน"
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.TextBox TBPRSearchItem 
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
            Height          =   375
            Left            =   1980
            TabIndex        =   134
            Top             =   1395
            Width           =   4290
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "คำที่ค้นหา :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   945
            TabIndex        =   133
            Top             =   1440
            Width           =   1860
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "ค้นหา รหัสสินค้า"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   405
            TabIndex        =   132
            Top             =   405
            Width           =   1905
         End
      End
      Begin VB.PictureBox PICPRArKeyData 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   2040
         Left            =   45
         ScaleHeight     =   2010
         ScaleWidth      =   19065
         TabIndex        =   188
         Top             =   0
         Visible         =   0   'False
         Width           =   19095
         Begin VB.PictureBox PicPoint 
            BackColor       =   &H00FFFFFF&
            Height          =   150
            Left            =   0
            ScaleHeight     =   90
            ScaleWidth      =   180
            TabIndex        =   207
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.TextBox TBPRKeyMember 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   2835
            TabIndex        =   190
            Top             =   675
            Width           =   3885
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "รหัสสมาชิก :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   810
            TabIndex        =   189
            Top             =   720
            Width           =   1860
         End
      End
      Begin VB.CommandButton CMDSearchItem 
         Height          =   645
         Left            =   7110
         Picture         =   "FormPickingRequest.frx":AC49
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   2385
         Width           =   645
      End
      Begin VB.TextBox TBBarCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2925
         TabIndex        =   102
         Top             =   2385
         Width           =   4155
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   2040
         Left            =   0
         ScaleHeight     =   2010
         ScaleWidth      =   15285
         TabIndex        =   84
         Top             =   0
         Width           =   15315
         Begin VB.CommandButton CMDSale 
            Height          =   375
            Left            =   6075
            Picture         =   "FormPickingRequest.frx":B513
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   1395
            Width           =   375
         End
         Begin VB.CommandButton CMDARCode 
            Height          =   375
            Left            =   4230
            Picture         =   "FormPickingRequest.frx":B8E0
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   765
            Width           =   375
         End
         Begin VB.TextBox TXTLicense 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9990
            TabIndex        =   96
            Top             =   180
            Width           =   1635
         End
         Begin VB.TextBox TXTSaleCode 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2025
            TabIndex        =   100
            Top             =   1395
            Width           =   4020
         End
         Begin VB.TextBox TXTArCode 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2025
            TabIndex        =   97
            Top             =   765
            Width           =   2175
         End
         Begin VB.CommandButton CMDDocNo 
            Height          =   375
            Left            =   4230
            Picture         =   "FormPickingRequest.frx":BCAD
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   180
            Width           =   375
         End
         Begin VB.TextBox TXTDocNo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2025
            TabIndex        =   94
            Top             =   180
            Width           =   2175
         End
         Begin VB.Label TXTMember 
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
            Left            =   13095
            TabIndex        =   93
            Top             =   180
            Width           =   1590
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "รหัสสมาชิก :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   11610
            TabIndex        =   92
            Top             =   180
            Width           =   1365
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ทะเบียนรถ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   8505
            TabIndex        =   91
            Top             =   180
            Width           =   1410
         End
         Begin VB.Label DTPDocDate1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   6300
            TabIndex        =   90
            Top             =   180
            Width           =   2040
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "วันที่เอกสาร :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   4590
            TabIndex        =   89
            Top             =   180
            Width           =   1635
         End
         Begin VB.Label LBLArName 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   4635
            TabIndex        =   88
            Top             =   765
            Width           =   10050
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "พนักงานขาย :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   495
            TabIndex        =   87
            Top             =   1395
            Width           =   1410
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "รหัสลูกค้า :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   585
            TabIndex        =   86
            Top             =   765
            Width           =   1320
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "เลขที่เอกสาร :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   360
            TabIndex        =   85
            Top             =   180
            Width           =   1590
         End
      End
      Begin VB.PictureBox PICSearchItem 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   5865
         Left            =   90
         ScaleHeight     =   5835
         ScaleWidth      =   14655
         TabIndex        =   137
         Top             =   3375
         Visible         =   0   'False
         Width           =   14685
         Begin MSComctlLib.ListView ListViewStock 
            Height          =   2625
            Left            =   5265
            TabIndex        =   154
            Top             =   2205
            Width           =   6720
            _ExtentX        =   11853
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "คลัง"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ชั้บเก็บ"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "คงเหลือ"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "หน่วยนับ"
               Object.Width           =   2646
            EndProperty
         End
         Begin VB.CommandButton CMDOK 
            BackColor       =   &H00808080&
            Caption         =   "ตกลง"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Left            =   2295
            Style           =   1  'Graphical
            TabIndex        =   121
            Top             =   3780
            Width           =   2445
         End
         Begin VB.TextBox TXTDisCount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2295
            TabIndex        =   120
            Top             =   3015
            Width           =   2445
         End
         Begin VB.TextBox TBQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   2295
            TabIndex        =   119
            Top             =   2205
            Width           =   2445
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ส่วนลด/รายการ :"
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
            Height          =   375
            Left            =   180
            TabIndex        =   153
            Top             =   3105
            Width           =   1905
         End
         Begin VB.Label LBLNetPrice 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5670
            TabIndex        =   152
            Top             =   990
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label LBLDisCountAmount 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5085
            TabIndex        =   151
            Top             =   990
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label LBLShelfID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   8550
            TabIndex        =   150
            Top             =   990
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label LBLBarCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   7335
            TabIndex        =   149
            Top             =   990
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label LBLZoneID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   7965
            TabIndex        =   148
            Top             =   990
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label LBLShelfCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   6795
            TabIndex        =   147
            Top             =   990
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label LBLWHCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   6210
            TabIndex        =   146
            Top             =   990
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label LBLUnitCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Height          =   375
            Left            =   2295
            TabIndex        =   145
            Top             =   990
            Width           =   2445
         End
         Begin VB.Label LBLPrice 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Height          =   375
            Left            =   2295
            TabIndex        =   144
            Top             =   1575
            Width           =   2445
         End
         Begin VB.Label LBLItemName 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Height          =   375
            Left            =   5085
            TabIndex        =   143
            Top             =   405
            Width           =   8970
         End
         Begin VB.Label LBLItemCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Height          =   375
            Left            =   2295
            TabIndex        =   142
            Top             =   405
            Width           =   2445
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "หน่วยนับ :"
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
            TabIndex        =   141
            Top             =   990
            Width           =   1365
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ต้องการจำนวน :"
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
            Left            =   360
            TabIndex        =   140
            Top             =   2250
            Width           =   1725
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ราคา :"
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
            Left            =   1305
            TabIndex        =   139
            Top             =   1575
            Width           =   780
         End
         Begin VB.Label Label6 
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
            Height          =   375
            Left            =   630
            TabIndex        =   138
            Top             =   405
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView ListViewItem 
         Height          =   4515
         Left            =   90
         TabIndex        =   104
         Top             =   3465
         Width           =   14670
         _ExtentX        =   25876
         _ExtentY        =   7964
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "จำนวน"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วย"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "ราคา"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "มูลค่า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "ส่วนลด"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "คลัง"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ชั้นเก็บ"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "โซน"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "ที่เก็บ"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "บาร์โค้ด"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "ส่วนลด"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.CommandButton CMDClearScreen 
         BackColor       =   &H00808080&
         Caption         =   "ล้างหน้าจอ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   1935
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   8235
         Width           =   1815
      End
      Begin VB.CommandButton CMDSave 
         BackColor       =   &H00808080&
         Caption         =   "บันทึก(F5)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   8235
         Width           =   1815
      End
      Begin VB.CommandButton CMDSearch 
         BackColor       =   &H00808080&
         Caption         =   "ค้นหา(F1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   6390
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   8235
         Width           =   1815
      End
      Begin VB.CommandButton CMDCancel 
         BackColor       =   &H00808080&
         Caption         =   "ยกเลิก(F8)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   8235
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   8235
         Width           =   1815
      End
      Begin VB.CommandButton CMDQue 
         BackColor       =   &H00808080&
         Caption         =   "จัดสินค้า(F9)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   8235
         Width           =   1815
      End
      Begin VB.CommandButton CMDPRMain 
         BackColor       =   &H00808080&
         Caption         =   "กลับหน้าหลัก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   8235
         Width           =   1815
      End
      Begin VB.Label LBLTotalAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   645
         Left            =   11025
         TabIndex        =   107
         Top             =   2340
         Width           =   3705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสสินค้า/บาร์โค้ด :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   405
         TabIndex        =   98
         Top             =   2520
         Width           =   2445
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000080FF&
         BorderWidth     =   15
         X1              =   0
         X2              =   15345
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000080FF&
         BorderWidth     =   15
         X1              =   0
         X2              =   15345
         Y1              =   2160
         Y2              =   2160
      End
   End
   Begin VB.PictureBox PICDriveIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   10815
      Left            =   0
      ScaleHeight     =   10615.16
      ScaleMode       =   0  'User
      ScaleWidth      =   14850
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   14880
      Begin VB.PictureBox PICDISearchDI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14610
         TabIndex        =   111
         Top             =   90
         Visible         =   0   'False
         Width           =   14640
         Begin VB.CommandButton CMDDIPICSearchDIClose 
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
            Height          =   960
            Left            =   12060
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   7875
            Width           =   2130
         End
         Begin VB.CommandButton CMDDISearchDIClick 
            Height          =   375
            Left            =   6030
            Picture         =   "FormPickingRequest.frx":C100
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   1080
            Width           =   375
         End
         Begin MSComctlLib.ListView ListViewDISearchDI 
            Height          =   5730
            Left            =   585
            TabIndex        =   51
            Top             =   1980
            Width           =   13605
            _ExtentX        =   23998
            _ExtentY        =   10107
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "เลขที่เอกสาร"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "วันที่เอกสาร"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "ลูกค้า"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "ทะเบียนรถ"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "พนักงานขาย"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "มูลค่า"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "ยกเลิก"
               Object.Width           =   2
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "อนุมัติ"
               Object.Width           =   2
            EndProperty
         End
         Begin VB.TextBox TBDISearchDI 
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
            Height          =   375
            Left            =   1530
            TabIndex        =   49
            Top             =   1080
            Width           =   4470
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "คำที่ค้นหา :"
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
            Left            =   585
            TabIndex        =   115
            Top             =   1125
            Width           =   1230
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "ค้นหาเอกสาร DriveIn"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   225
            TabIndex        =   114
            Top             =   225
            Width           =   2400
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   14880
         TabIndex        =   204
         Top             =   9540
         Width           =   14910
      End
      Begin VB.PictureBox PICDISendInformation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14565
         TabIndex        =   193
         Top             =   90
         Visible         =   0   'False
         Width           =   14595
         Begin VB.PictureBox PICPoint3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   210
            ScaleWidth      =   435
            TabIndex        =   210
            Top             =   0
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.CommandButton CMDDISendInformationClose 
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
            Height          =   870
            Left            =   11340
            Style           =   1  'Graphical
            TabIndex        =   195
            Top             =   7740
            Width           =   2040
         End
         Begin MSComctlLib.ListView ListViewDIInfQue 
            Height          =   5415
            Left            =   1350
            TabIndex        =   194
            Top             =   2160
            Width           =   12030
            _ExtentX        =   21220
            _ExtentY        =   9551
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "คิวจัดสินค้า"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "โซนจัดสินค้า"
               Object.Width           =   8820
            EndProperty
         End
         Begin VB.Label LBLDIInfDocNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   2790
            TabIndex        =   198
            Top             =   1350
            Width           =   3390
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "เลขที่เอกสาร :"
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
            Left            =   1395
            TabIndex        =   197
            Top             =   1350
            Width           =   1410
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "รายการ เลขที่คิวจัดสินค้า"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1395
            TabIndex        =   196
            Top             =   630
            Width           =   2985
         End
      End
      Begin VB.PictureBox PICDISearchSale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14610
         TabIndex        =   116
         Top             =   90
         Visible         =   0   'False
         Width           =   14640
         Begin VB.CommandButton CMDDIPICSearchSaleClose 
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
            Height          =   960
            Left            =   10575
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   7200
            Width           =   2220
         End
         Begin MSComctlLib.ListView ListViewDISearchSale 
            Height          =   4830
            Left            =   2115
            TabIndex        =   41
            Top             =   2160
            Width           =   10680
            _ExtentX        =   18838
            _ExtentY        =   8520
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "รหัสพนักงาน"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ชื่อพนักงาน"
               Object.Width           =   12347
            EndProperty
         End
         Begin VB.CommandButton CMDDISearchSaleClick 
            Height          =   420
            Left            =   6300
            Picture         =   "FormPickingRequest.frx":C4CD
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1170
            Width           =   375
         End
         Begin VB.TextBox TBDISearchSale 
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
            Height          =   420
            Left            =   2115
            TabIndex        =   39
            Top             =   1170
            Width           =   4155
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "คำที่ค้นหา :"
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
            Left            =   1080
            TabIndex        =   124
            Top             =   1215
            Width           =   1320
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "ค้นหา พนักงาน "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1080
            TabIndex        =   123
            Top             =   360
            Width           =   2085
         End
      End
      Begin VB.PictureBox PICDILastSendQue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14610
         TabIndex        =   174
         Top             =   90
         Visible         =   0   'False
         Width           =   14640
         Begin VB.CommandButton CMDDISendQuePrint 
            BackColor       =   &H00808080&
            Caption         =   "แก้ไขคิวจัดสินค้า"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   9495
            Style           =   1  'Graphical
            TabIndex        =   206
            Top             =   7425
            Width           =   2130
         End
         Begin VB.CommandButton CMDDISendQueExit 
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
            Height          =   870
            Left            =   12105
            Style           =   1  'Graphical
            TabIndex        =   176
            Top             =   7425
            Width           =   2130
         End
         Begin MSComctlLib.ListView ListViewDILastSendQue 
            Height          =   6180
            Left            =   450
            TabIndex        =   175
            Top             =   1080
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   10901
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   12648447
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
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ชื่อสินค้า"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "จ่าย"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "หน่วย"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "คิวที่"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "โซน"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "เลขที่เอกสาร"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "รหัสสินค้า"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "คลัง"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "ชั้นเก็บ"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "บาร์โค้ด"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "PickZone"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "ที่เก็บ"
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "รายการ คิวจัดสินค้าล่าสุดของเอกสารนี้"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   450
            TabIndex        =   191
            Top             =   495
            Width           =   4740
         End
      End
      Begin VB.PictureBox PICDISearchItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14610
         TabIndex        =   125
         Top             =   90
         Visible         =   0   'False
         Width           =   14640
         Begin VB.CommandButton CMDDIPICSearchItemClose 
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
            Height          =   915
            Left            =   12465
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   7785
            Width           =   2040
         End
         Begin MSComctlLib.ListView ListViewDISearchItem 
            Height          =   5820
            Left            =   180
            TabIndex        =   31
            Top             =   1890
            Width           =   14325
            _ExtentX        =   25268
            _ExtentY        =   10266
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "รหัสสินค้า"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ชื่อสินค้า"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "คงเหลือ"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "หน่วยนับ"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "ราคาต่อหน่วย"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "โซน"
               Object.Width           =   2293
            EndProperty
         End
         Begin VB.CommandButton CMDISearchItemClick 
            Height          =   330
            Left            =   5715
            Picture         =   "FormPickingRequest.frx":C89A
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1260
            Width           =   375
         End
         Begin VB.TextBox TBDISearchItem 
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
            Height          =   330
            Left            =   1125
            TabIndex        =   29
            Top             =   1260
            Width           =   4515
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "คำที่ค้นหา :"
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
            TabIndex        =   127
            Top             =   1260
            Width           =   1680
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "ค้นหา รหัสสินค้า"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   126
            Top             =   405
            Width           =   1680
         End
      End
      Begin VB.PictureBox PICDISearchAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9195
         Left            =   90
         ScaleHeight     =   9165
         ScaleWidth      =   14610
         TabIndex        =   117
         Top             =   90
         Visible         =   0   'False
         Width           =   14640
         Begin VB.CommandButton CMDDIPICSearchARClose 
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
            Height          =   870
            Left            =   11610
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   7830
            Width           =   1995
         End
         Begin MSComctlLib.ListView ListViewDISearchAR 
            Height          =   5550
            Left            =   1350
            TabIndex        =   61
            Top             =   2115
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   9790
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ลำดับ"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "รหัสลูกค้า"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ชื่อลูกค้า"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "รหัสสมาชิก"
               Object.Width           =   4410
            EndProperty
         End
         Begin VB.CommandButton CMDDISearchARClick 
            Height          =   375
            Left            =   6705
            Picture         =   "FormPickingRequest.frx":CC67
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   1170
            Width           =   375
         End
         Begin VB.TextBox TBDISearchAR 
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
            Height          =   375
            Left            =   2430
            TabIndex        =   59
            Top             =   1170
            Width           =   4245
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   " คำที่ค้นหา :"
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
            Left            =   1305
            TabIndex        =   122
            Top             =   1215
            Width           =   2040
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "ค้นหา ลูกค้า"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   405
            TabIndex        =   118
            Top             =   315
            Width           =   1500
         End
      End
      Begin VB.PictureBox PICArKeyData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1860
         Left            =   135
         ScaleHeight     =   1830
         ScaleWidth      =   14565
         TabIndex        =   67
         Top             =   90
         Visible         =   0   'False
         Width           =   14595
         Begin VB.TextBox TBDIKeyMember 
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
            Height          =   555
            Left            =   2970
            TabIndex        =   69
            Top             =   855
            Width           =   3930
         End
         Begin VB.TextBox TBDIKeyArCode 
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
            Height          =   555
            Left            =   2970
            TabIndex        =   68
            Top             =   225
            Width           =   3930
         End
         Begin VB.Label Label67 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "รหัสสมาชิก :"
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
            Left            =   855
            TabIndex        =   71
            Top             =   990
            Width           =   1950
         End
         Begin VB.Label Label61 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "รหัสลูกค้า :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1575
            TabIndex        =   70
            Top             =   315
            Width           =   1230
         End
      End
      Begin VB.PictureBox PICDIKeyQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   5955
         Left            =   135
         ScaleHeight     =   5925
         ScaleWidth      =   14565
         TabIndex        =   34
         Top             =   3330
         Visible         =   0   'False
         Width           =   14595
         Begin MSComctlLib.ListView ListViewDIItemStock 
            Height          =   3390
            Left            =   6615
            TabIndex        =   65
            Top             =   1710
            Width           =   6810
            _ExtentX        =   12012
            _ExtentY        =   5980
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "คลัง"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
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
               SubItemIndex    =   3
               Text            =   "หน่วยนับ"
               Object.Width           =   2646
            EndProperty
         End
         Begin VB.CommandButton CMDDIKeyQtyOK 
            BackColor       =   &H00808080&
            Caption         =   "ตกลง"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Left            =   2835
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   4095
            Width           =   2265
         End
         Begin VB.TextBox TBDIKeyDiscount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2835
            TabIndex        =   20
            Top             =   3330
            Width           =   2265
         End
         Begin VB.TextBox TBDIKeyQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2835
            TabIndex        =   19
            Top             =   2520
            Width           =   2265
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "รายการ ยอดคงเหลือ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   6615
            TabIndex        =   81
            Top             =   1305
            Width           =   2670
         End
         Begin VB.Label LBLDIItemCost 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   465
            Left            =   13680
            TabIndex        =   66
            Top             =   4455
            Width           =   375
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ชื่อสินค้า :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1350
            TabIndex        =   63
            Top             =   765
            Width           =   1365
         End
         Begin VB.Label LBLDIZoneID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   13095
            TabIndex        =   58
            Top             =   3285
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LBLDIShelfCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   13095
            TabIndex        =   57
            Top             =   2385
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LBLDIItemNetAmount 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   13095
            TabIndex        =   56
            Top             =   4500
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LBLDIPrice 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   2835
            TabIndex        =   55
            Top             =   1845
            Width           =   2265
         End
         Begin VB.Label LBLDIBarCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   13095
            TabIndex        =   54
            Top             =   3690
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LBLDIWHCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   13095
            TabIndex        =   53
            Top             =   1935
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LBLDIShelfID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   13095
            TabIndex        =   48
            Top             =   2835
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label75 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ต้องการ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   585
            TabIndex        =   47
            Top             =   2565
            Width           =   2085
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ส่วนลด/รายการ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   540
            TabIndex        =   46
            Top             =   3330
            Width           =   2130
         End
         Begin VB.Label LBLDIDiscountWord 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   13095
            TabIndex        =   45
            Top             =   4095
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LBLDIUnitCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   2835
            TabIndex        =   44
            Top             =   1305
            Width           =   1545
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ราคา/หน่วย :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   945
            TabIndex        =   43
            Top             =   1845
            Width           =   1680
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "หน่วย :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   630
            TabIndex        =   38
            Top             =   1305
            Width           =   2040
         End
         Begin VB.Label LBLDIItemCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Height          =   420
            Left            =   2835
            TabIndex        =   37
            Top             =   225
            Width           =   3030
         End
         Begin VB.Label LBLDIItemName 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Height          =   420
            Left            =   2835
            TabIndex        =   36
            Top             =   765
            Width           =   10635
         End
         Begin VB.Label Label65 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "รหัสสินค้า :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   855
            TabIndex        =   35
            Top             =   225
            Width           =   1860
         End
      End
      Begin VB.CommandButton CMDDISelectPoint 
         Caption         =   "เลือกจุด DI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   135
         TabIndex        =   10
         Top             =   8325
         Width           =   2175
      End
      Begin VB.CommandButton CMDDICancel 
         Caption         =   "ยกเลิก-F8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   9720
         TabIndex        =   14
         Top             =   8325
         Width           =   2175
      End
      Begin VB.CommandButton CMDDISendQue 
         Caption         =   "CheckOut-F9"
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
         Height          =   915
         Left            =   12555
         TabIndex        =   15
         Top             =   8325
         Width           =   2175
      End
      Begin VB.CommandButton CMDDISave 
         Caption         =   "บันทึก-F5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   5220
         TabIndex        =   12
         Top             =   8325
         Width           =   2175
      End
      Begin VB.CommandButton CMDDIClearScreen 
         Caption         =   "ล้างหน้าจอ-ESC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   2970
         TabIndex        =   11
         Top             =   8325
         Width           =   2175
      End
      Begin VB.TextBox TBDIBarCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1800
         TabIndex        =   7
         ToolTipText     =   "ค้นหาสินค้า กด F4"
         Top             =   2295
         Width           =   4380
      End
      Begin MSComctlLib.ListView ListViewDIItem 
         Height          =   4425
         Left            =   135
         TabIndex        =   9
         Top             =   3645
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   7805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "จำนวน"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วย"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "ราคา"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "มูลค่า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "ส่วนลด"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "คลัง"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ชั้นเก็บ"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "โซน"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "ที่เก็บ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "บาร์โค้ด"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "ส่วนลด"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "จุดจ่าย"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   2040
         Left            =   0
         ScaleHeight     =   2010
         ScaleWidth      =   15285
         TabIndex        =   72
         Top             =   0
         Width           =   15315
         Begin VB.PictureBox PICPoint2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   0
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   209
            Top             =   0
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.CommandButton CMDDISearchSale 
            Height          =   420
            Left            =   5760
            Picture         =   "FormPickingRequest.frx":D034
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "ค้นหาพนักงานขาย กด F3"
            Top             =   1350
            Width           =   375
         End
         Begin VB.CommandButton CMDDISearchAr 
            Height          =   420
            Left            =   4230
            Picture         =   "FormPickingRequest.frx":D401
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "ค้นหาลูกค้า กด F2"
            Top             =   180
            Width           =   375
         End
         Begin VB.TextBox TBDISaleCode 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1800
            TabIndex        =   5
            ToolTipText     =   "ค้นหาพนักงานขาย กด F3"
            Top             =   1350
            Width           =   3930
         End
         Begin VB.TextBox TBDICarLicense 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1800
            TabIndex        =   4
            Top             =   765
            Width           =   2400
         End
         Begin VB.TextBox TBDIArCode 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1800
            TabIndex        =   0
            Text            =   "99999"
            ToolTipText     =   "ค้นหาลูกค้า กด F2"
            Top             =   180
            Width           =   2400
         End
         Begin VB.Image IMCancel 
            Height          =   300
            Left            =   14085
            Picture         =   "FormPickingRequest.frx":D7CE
            Top             =   180
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Image IMConfirm 
            Height          =   300
            Left            =   14085
            Picture         =   "FormPickingRequest.frx":DD0A
            Top             =   180
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Image IMNormal 
            Height          =   300
            Left            =   14085
            Picture         =   "FormPickingRequest.frx":E1ED
            Top             =   180
            Width           =   570
         End
         Begin VB.Label LBLDIDocDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   6300
            TabIndex        =   80
            Top             =   1350
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.Label LBLDIDocNo 
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
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   10485
            TabIndex        =   79
            Top             =   765
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.Label LBLDI 
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
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   8820
            TabIndex        =   78
            Top             =   765
            Width           =   1500
         End
         Begin VB.Label LBLDIMember 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   5940
            TabIndex        =   3
            Top             =   765
            Width           =   2760
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "รหัสสมาชิก :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   4545
            TabIndex        =   76
            Top             =   765
            Width           =   1365
         End
         Begin VB.Label LBLDIArName 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "เงินสด"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   4635
            TabIndex        =   2
            Top             =   180
            Width           =   8295
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "พนักงานขาย :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   315
            TabIndex        =   75
            Top             =   1350
            Width           =   1455
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ทะเบียนรถ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   405
            TabIndex        =   74
            Top             =   810
            Width           =   1365
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "รหัสลูกค้า :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   450
            TabIndex        =   73
            Top             =   180
            Width           =   1320
         End
      End
      Begin VB.CommandButton CMDDISearchDocNo 
         Caption         =   "ค้นหา-F1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   7470
         TabIndex        =   13
         Top             =   8325
         Width           =   2175
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   0
         ScaleHeight     =   1425
         ScaleWidth      =   15285
         TabIndex        =   82
         Top             =   8100
         Width           =   15315
      End
      Begin VB.CommandButton CMDDISearchItem 
         Height          =   690
         Left            =   6255
         Picture         =   "FormPickingRequest.frx":E61F
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "ค้นหาสินค้า กด F4"
         Top             =   2295
         Width           =   735
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "มูลค่าสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8100
         TabIndex        =   211
         Top             =   2430
         Width           =   2310
      End
      Begin VB.Label LBLDINetAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   675
         Left            =   10485
         TabIndex        =   64
         Top             =   2295
         Width           =   4245
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         BorderWidth     =   15
         X1              =   0
         X2              =   15345
         Y1              =   3100.394
         Y2              =   3100.394
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   15
         X1              =   0
         X2              =   15345
         Y1              =   2081.693
         Y2              =   2081.693
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "บาร์โค้ด :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         TabIndex        =   33
         Top             =   2430
         Width           =   1185
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการสินค้า/รายละเอียด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   135
         TabIndex        =   28
         Top             =   3330
         Width           =   2715
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FF0000&
      Height          =   510
      Left            =   360
      ScaleHeight     =   450
      ScaleWidth      =   14085
      TabIndex        =   201
      Top             =   9045
      Width           =   14145
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000080FF&
      Height          =   510
      Left            =   360
      ScaleHeight     =   450
      ScaleWidth      =   14085
      TabIndex        =   200
      Top             =   7605
      Width           =   14145
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   510
      Left            =   360
      ScaleHeight     =   450
      ScaleWidth      =   14130
      TabIndex        =   199
      Top             =   2205
      Width           =   14190
   End
   Begin VB.PictureBox PICSelectDI 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4650
      Left            =   360
      ScaleHeight     =   4620
      ScaleWidth      =   14115
      TabIndex        =   77
      Top             =   2880
      Visible         =   0   'False
      Width           =   14145
      Begin VB.CommandButton CMDDIMain 
         BackColor       =   &H00808080&
         Caption         =   "กลับหน้าหลัก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   2700
         Width           =   2850
      End
      Begin VB.CommandButton CMDDI04 
         BackColor       =   &H00808080&
         Caption         =   "จุดที่ 4 (D)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   495
         Width           =   2850
      End
      Begin VB.CommandButton CMDDI03 
         BackColor       =   &H00808080&
         Caption         =   "จุดที่ 3 (C)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   7155
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "สินค้ากลุ่ม เหล็ก"
         Top             =   495
         Width           =   2805
      End
      Begin VB.CommandButton CMDDI02 
         BackColor       =   &H00808080&
         Caption         =   "จุดที่ 2 (B)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   4185
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "สินค้ากลุ่ม ปูน ท่อ กระเบื้อง สังกะสี ฯลฯ"
         Top             =   495
         Width           =   2850
      End
      Begin VB.CommandButton CMDDI01 
         BackColor       =   &H00808080&
         Caption         =   "จุดที่ 1 (A)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "สินค้าอาคาร A"
         Top             =   495
         Width           =   2850
      End
   End
   Begin VB.PictureBox PICSelectJob 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4605
      Left            =   360
      ScaleHeight     =   4575
      ScaleWidth      =   14115
      TabIndex        =   25
      Top             =   2880
      Width           =   14145
      Begin VB.CommandButton CMDPickReq 
         BackColor       =   &H00808080&
         Caption         =   "Picking Request"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   8685
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1710
         Width           =   3210
      End
      Begin VB.CommandButton CMDDriveIn 
         BackColor       =   &H00808080&
         Caption         =   "DriveIn"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1710
         Width           =   3210
      End
   End
   Begin VB.CommandButton CMDExit 
      BackColor       =   &H00808080&
      Caption         =   "กลับ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8235
      Width           =   1635
   End
   Begin VB.Label LBLSelectDI 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกจุด DriveIn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   4320
      TabIndex        =   203
      Top             =   810
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.Label LBLJob 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกรูปแบบ การเข้าใช้งานโปรแกรม"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   3600
      TabIndex        =   202
      Top             =   900
      Width           =   8070
   End
End
Attribute VB_Name = "FormPickingRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDIIsOpen As Integer
Dim vIsOpen As Integer
Dim vPRIsOpen As Integer
Dim vDISendQue As Integer
Dim vSendQue As Integer
Dim vMemDIIsConfirm As Integer
Dim vMemDIIsCancel As Integer
Dim vMemIsCancel As Integer

Dim vDIIsCancel As Integer
Dim vDIIsConfirm As Integer
Dim vDIIsSendQue As Integer


Dim vCountItemOld As Integer
Dim vDIItemCodeOld() As String
Dim vDIUnitCodeOld() As String
Dim vDIWHCodeOld() As String
Dim vDIShelfCodeOld() As String
Dim vDIZoneIDOld() As String
Dim vDIPickZoneOld() As String
Dim vDIBarCodeOld() As String

Dim vCountItemPickZoneOld As Integer

Public Sub CalcEditItemQty()
Dim i As Integer
Dim vAmount As Double
Dim vNetAmount As Double
Dim vSumOfItemAmount As Double
Dim vTaxAmount As Double
Dim vTotalAmount As Double
Dim vLastDisCountAmount As Double

'If Me.ListViewSaleOrder.ListItems.Count > 0 Then
 '  For i = 1 To Me.ListViewSaleOrder.ListItems.Count
  '     vAmount = Me.ListViewSaleOrder.ListItems(i).SubItems(7)
   '    vNetAmount = vNetAmount + vAmount
   'Next i
   'vSumOfItemAmount = vNetAmount
   'If Me.TBSOLastDisCount.Text <> "" Then
    '  vLastDisCountAmount = Me.TBSOLastDisCount.Text
   'End If
   'vTotalAmount = vNetAmount - vLastDisCountAmount
   'vTaxAmount = vTotalAmount - ((vTotalAmount * 100) / 107)
   
   'Me.LBLOrderSumOfItemAmount.Caption = Format(vSumOfItemAmount, "##,##0.00")
   'Me.LBLOrderTaxAmount.Caption = Format(vTaxAmount, "##,##0.00")
   'Me.LBLOrderNetAmount.Caption = Format(vTotalAmount, "##,##0.00")
'Else
 '  Me.LBLOrderSumOfItemAmount.Caption = Format(0, "##,##0.00")
  ' Me.LBLOrderTaxAmount.Caption = Format(0, "##,##0.00")
   'Me.LBLOrderNetAmount.Caption = Format(0, "##,##0.00")
'End If
End Sub

Private Sub CMDArCode_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String
Dim vListItem As ListItem
Dim i As Integer

On Error Resume Next

vSearchAR = ""
vQuery = "exec dbo.USP_AR_SearchARLine '" & vSearchAR & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.ListViewPRSearchAR.ListItems.Clear
    vRecordset.MoveFirst
    For i = 1 To vRecordset.RecordCount
       Set vListItem = Me.ListViewPRSearchAR.ListItems.Add(, , i)
       vListItem.SubItems(1) = Trim(vRecordset.Fields("arcode").Value)
       vListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
       vListItem.SubItems(3) = Trim(vRecordset.Fields("memberid").Value)
       vRecordset.MoveNext
    Next i
End If
vRecordset.Close

Me.PICPRSearchAR.Visible = True
Me.TBPRSearchAR.SetFocus
End Sub

Private Sub CMDARCode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If
End Sub

Private Sub CMDCancel_Click()
Dim vQuery As String
Dim vDocNo As String
Dim vAnswer As Integer

Dim vDocdate As String
Dim vCheckDate As String


If vPRIsOpen = 0 Then
   MsgBox "ไม่สามารถยกเลิกเอกสารที่ยังไม่ได้บันทึกได้ กรุณาตรวจสอบ", vbCritical, "SendError Message"
   Exit Sub
End If

If Me.ListViewItem.ListItems.Count > 0 And vPRIsOpen = 1 And Me.TXTDocNo.Text <> "" Then
   vDocNo = Me.TXTDocNo.Text
   vDocdate = Me.DTPDocDate1.Caption
   
   Dim vDay As String
   Dim vMonth As String
   
   If Len(Day(Now)) = 1 Then
      vDay = Trim("0" & Day(Now))
   Else
      vDay = Day(Now)
   End If
   
   If Len(Month(Now)) = 1 Then
      vMonth = Trim("0" & Month(Now))
   Else
      vMonth = Month(Now)
   End If
   
   vCheckDate = vDay & "/" & vMonth & "/" & Year(Now)
   
   If vDocdate < vCheckDate Then
      MsgBox "ไม่สามารถยกเลิกเอกสารได้ กรณีวันที่เอกสารน้อยกว่าวันที่ปัจจุบัน", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
   If vMemIsCancel = 1 Then
      MsgBox "เอกสารเลขที่ " & vDocNo & " ถูกยกเลิกไปแล้ว", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
   If vSendQue = 1 Then
      MsgBox "เอกสารเลขที่ " & vDocNo & " ถูกอ้างอิงไปจัดคิวเรียบร้อยแล้ว ไม่สามารถยกเลิกได้", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
   vAnswer = MsgBox("คุณต้องการยกเลิกเอกสารเลขที่ " & vDocNo & " นี้ใช่หรือไม่ ", vbYesNo, "Send Question Message")
   
   If vAnswer = 6 Then
   vQuery = "exec dbo.USP_NP_CancelReqPicking '" & vDocNo & "' "
   gConnection.Execute vQuery
   
   MsgBox "ยกเลิกเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", vbCritical, "Send Information Message"
   
   Call CMDDocNo_Click

   End If

End If
End Sub

Private Sub CMDCancel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If
End Sub

Private Sub CMDClearScreen_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vHeader As String
Dim vRunning As Integer
Dim vDocNumber As String

Me.TXTDocNo.Text = ""
vQuery = "exec dbo.USP_NP_SearchNewDocNo  32 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vHeader = Trim(vRecordset.Fields("header").Value)
    vRunning = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
End If
vRecordset.Close
      
vDocNo = vDocNumber & vHeader & "-" & Format(vRunning, "0000")
Me.TXTDocNo.Text = vDocNo
vPRIsOpen = 0
Me.CMDQue.Enabled = False
End Sub

'Private Sub CMDClose_Click()
'Me.PICSearchItem.Visible = False
'End Sub

Private Sub CMDClearScreen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If
End Sub

Private Sub CMDDI01_Click()
'Me.PICDriveIn.Visible = True
'Me.LBLDI.Caption = "01"
'Me.TBDIArCode.SetFocus
End Sub

Private Sub CMDDI01_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 97 Then
'Me.PICDriveIn.Visible = True
'Me.LBLDI.Caption = "01"
'Me.TBDIArCode.SetFocus
'End If

If KeyCode = 98 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "02"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 99 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "03"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 100 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "04"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 27 Then
Call CMDDIMain_Click
End If
End Sub

Private Sub CMDDI02_Click()
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "02"
Me.TBDIArCode.SetFocus
End Sub

Private Sub CMDDI02_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 97 Then
'Me.PICDriveIn.Visible = True
'Me.LBLDI.Caption = "01"
'Me.TBDIArCode.SetFocus
'End If

If KeyCode = 98 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "02"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 99 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "03"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 100 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "04"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 27 Then
Call CMDDIMain_Click
End If
End Sub

Private Sub CMDDI03_Click()
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "03"
Me.TBDIArCode.SetFocus
End Sub

Private Sub CMDDI03_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 97 Then
'Me.PICDriveIn.Visible = True
'Me.LBLDI.Caption = "01"
'Me.TBDIArCode.SetFocus
'End If

If KeyCode = 98 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "02"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 99 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "03"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 100 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "04"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 27 Then
Call CMDDIMain_Click
End If
End Sub

Private Sub CMDDI04_Click()
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "04"
Me.TBDIArCode.SetFocus
End Sub

Private Sub CMDDI04_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 97 Then
'Me.PICDriveIn.Visible = True
'Me.LBLDI.Caption = "01"
'Me.TBDIArCode.SetFocus
'End If

If KeyCode = 98 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "02"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 99 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "03"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 100 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "04"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 27 Then
Call CMDDIMain_Click
End If
End Sub

Private Sub CMDDICancel_Click()
Dim vQuery As String
Dim vDocNo As String
Dim vAnswer As Integer

Dim vDocdate As String
Dim vCheckDate As String


If vDIIsOpen = 0 Then
   MsgBox "ไม่สามารถยกเลิกเอกสารที่ยังไม่ได้บันทึกได้ กรุณาตรวจสอบ", vbCritical, "SendError Message"
   Me.TBDIBarCode.SetFocus
   Exit Sub
End If

If Me.ListViewDIItem.ListItems.Count > 0 And vDIIsOpen = 1 And Me.LBLDIDocNo.Caption <> "" Then
   vDocNo = Me.LBLDIDocNo.Caption
   vDocdate = Me.LBLDIDocDate.Caption
   
   Dim vDay As String
   Dim vMonth As String
   
   If Len(Day(Now)) = 1 Then
      vDay = Trim("0" & Day(Now))
   Else
      vDay = Day(Now)
   End If
   
   If Len(Month(Now)) = 1 Then
      vMonth = Trim("0" & Month(Now))
   Else
      vMonth = Month(Now)
   End If
   
   vCheckDate = vDay & "/" & vMonth & "/" & Year(Now)
   
   If vDocdate < vCheckDate Then
      MsgBox "ไม่สามารถยกเลิกเอกสารได้ กรณีวันที่เอกสารน้อยกว่าวันที่ปัจจุบัน", vbCritical, "Send Error Message"
      Me.TBDIBarCode.SetFocus
      Exit Sub
   End If
   
   If vMemDIIsCancel = 1 Then
      MsgBox "เอกสารเลขที่ " & vDocNo & " ถูกยกเลิกไปแล้ว", vbCritical, "Send Error Message"
      Me.TBDIBarCode.SetFocus
      Exit Sub
   End If
   
   If vDISendQue = 1 Then
      MsgBox "เอกสารเลขที่ " & vDocNo & " ถูกอ้างอิงไปจัดคิวเรียบร้อยแล้ว ไม่สามารถยกเลิกได้", vbCritical, "Send Error Message"
      Me.TBDIBarCode.SetFocus
      Exit Sub
   End If
   
   vAnswer = MsgBox("คุณต้องการยกเลิกเอกสารเลขที่ " & vDocNo & " นี้ใช่หรือไม่ ", vbYesNo, "Send Question Message")
   
   If vAnswer = 6 Then
   vQuery = "exec dbo.USP_NP_CancelDriveInDocNo '" & vDocNo & "' "
   gConnection.Execute vQuery
   
   MsgBox "ยกเลิกเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", vbCritical, "Send Information Message"
   
   Call CMDDIClearScreen_Click
   
   End If

End If
End Sub

Private Sub CMDDICancel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub CMDDIClearScreen_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vHeader As String
Dim vRunning As Integer
Dim vDocNumber As String

vQuery = "exec dbo.USP_NP_SearchNewDocNo  29 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vHeader = Trim(vRecordset.Fields("header").Value)
    vRunning = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
End If
vRecordset.Close
      
vDocNo = vDocNumber & vHeader & "-" & Format(vRunning, "0000")
Me.LBLDIDocNo.Caption = vDocNo
vDIIsOpen = 0
vMemDIIsCancel = 0
vDISendQue = 0
Me.TBDIArCode.Text = ""
Me.TBDIArCode.Text = "99999"

Me.TBDISaleCode.Text = ""
Me.LBLDINetAmount.Caption = ""
Me.TBDICarLicense.Text = ""
Me.LBLDIMember.Caption = ""
Me.LBLDIDocDate.Caption = ""

Me.ListViewDIItem.ListItems.Clear
   
Call NewDoc
Me.CMDDISendQue.Enabled = False
Me.TBDICarLicense.SetFocus
End Sub

Private Sub CMDDIClearScreen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub CMDDIKeyQtyOK_Click()
Dim vListItem As ListItem
Dim i As Integer
Dim vBeforeTaxAmount As Double
Dim vTaxAmount As Double
Dim vNetAmount As Double
Dim vQTY As Double
Dim vPrice As Double
Dim vNetprice As Double
Dim vDiscountAmount As Double

Dim vWHCode As String
Dim vShelfCode As String
Dim vUnitCode As String

Dim vItemCode As String
Dim vCheckItemCode As String
Dim vCheckWHCode As String
Dim vCheckShelfCode As String
Dim vCheckUnitCode As String
Dim vAnswer As Integer
Dim vCountItem As Integer
Dim vPickZone As String
Dim vLinePickZone As String

If Me.LBLDIItemCode.Caption <> "" And Me.TBDIKeyQty.Text <> "" Then
   vItemCode = Me.LBLDIItemCode.Caption
   vWHCode = Me.LBLDIWHCode.Caption
   vShelfCode = Me.LBLDIShelfCode.Caption
   vUnitCode = Me.LBLDIUnitCode.Caption
   vQTY = Me.TBDIKeyQty.Text
   
   vPickZone = Me.LBLDI.Caption
   
   If vQTY = 0 Then
      MsgBox "ต้องกรอกจำนวนสินค้าที่ต้องการที่ไม่เท่ากับ 0  เท่านั้น", vbCritical, "Send Error Message"
      Me.TBDIKeyQty.SetFocus
      Exit Sub
   End If
         
   If Me.ListViewDIItem.ListItems.Count = 0 Then

      vCountItem = 0
      vCountItem = vCountItem + 1
      vQTY = Me.TBDIKeyQty.Text
      vPrice = Me.LBLDIPrice.Caption
      If Me.LBLDIDiscountWord.Caption <> "" Then
      vDiscountAmount = Me.LBLDIDiscountWord.Caption
      Else
      vDiscountAmount = 0
      End If
      vNetAmount = vQTY * (vPrice - vDiscountAmount)
                 
     Set vListItem = Me.ListViewDIItem.ListItems.Add(, , vCountItem)
     vListItem.SubItems(1) = Me.LBLDIItemCode.Caption
     vListItem.SubItems(2) = Me.LBLDIItemName.Caption
     vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
     vListItem.SubItems(4) = Me.LBLDIUnitCode.Caption
     vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
     vListItem.SubItems(6) = Format(vNetAmount, "##,##0.00")
     vListItem.SubItems(7) = Format(vDiscountAmount, "##,##0.00")
     vListItem.SubItems(8) = Me.LBLDIWHCode.Caption
     vListItem.SubItems(9) = Me.LBLDIShelfCode.Caption
     vListItem.SubItems(10) = Me.LBLDIZoneID.Caption
     vListItem.SubItems(11) = Me.LBLDIShelfID.Caption
     vListItem.SubItems(12) = Me.LBLDIBarCode.Caption
     vListItem.SubItems(13) = Me.TBDIKeyDiscount.Text
     vListItem.SubItems(14) = vPickZone
   Else
      For i = 1 To Me.ListViewDIItem.ListItems.Count
      vCheckItemCode = Me.ListViewDIItem.ListItems(i).SubItems(1)
      vCheckWHCode = Me.ListViewDIItem.ListItems(i).SubItems(8)
      vCheckShelfCode = Me.ListViewDIItem.ListItems(i).SubItems(9)
      vCheckUnitCode = Me.ListViewDIItem.ListItems(i).SubItems(4)
      vLinePickZone = Me.ListViewDIItem.ListItems(i).SubItems(14)
      
      If vItemCode = vCheckItemCode And vWHCode = vCheckWHCode And vShelfCode = vCheckShelfCode And vUnitCode = vCheckUnitCode And vPickZone = vLinePickZone Then
         vAnswer = MsgBox("มีรายการสินค้ารหัส " & vItemCode & " คลัง " & vWHCode & " ชั้นเก็บ " & vShelfCode & "นี้อยู่แล้ว ที่บรรทัดที่ " & i & " จะบันทึกทับข้อมูลเก่าหรือไม่ ", vbYesNo, "Send Information Message")
         If vAnswer = 7 Then
            Me.TBDIBarCode.Text = ""
            Me.TBDIBarCode.SetFocus
            Exit Sub
         Else
         
           vQTY = Me.TBDIKeyQty.Text
           vPrice = Me.LBLDIPrice.Caption
           If Me.LBLDIDiscountWord.Caption <> "" Then
           vDiscountAmount = Me.LBLDIDiscountWord.Caption
           Else
           vDiscountAmount = 0
           End If
           vNetAmount = vQTY * (vPrice - vDiscountAmount)

            Me.ListViewDIItem.ListItems(i).SubItems(3) = Format(vQTY, "##,##0.00")
            Me.ListViewDIItem.ListItems(i).SubItems(5) = Format(vPrice, "##,##0.00")
            Me.ListViewDIItem.ListItems(i).SubItems(6) = Format(vNetAmount, "##,##0.00")
            Me.ListViewDIItem.ListItems(i).SubItems(7) = Format(vDiscountAmount, "##,##0.00")

            Me.ListViewDIItem.ListItems(i).SubItems(13) = Me.TBDIKeyDiscount.Text
            
            Me.TBDIBarCode.Text = ""
            Me.LBLDIItemCode.Caption = ""
            Me.LBLDIItemName.Caption = ""
            Me.LBLDIUnitCode.Caption = ""
            Me.TBDIKeyQty.Text = ""
            Me.TBDIKeyDiscount.Text = ""
            Me.LBLDIPrice.Caption = ""
            Me.LBLDIDiscountWord.Caption = ""
            Me.LBLDIWHCode.Caption = ""
            Me.LBLDIShelfCode.Caption = ""
            Me.LBLDIZoneID.Caption = ""
            Me.LBLDIShelfID.Caption = ""
            Me.LBLDIBarCode.Caption = ""
            Me.LBLDIItemNetAmount.Caption = ""
            Me.ListViewDIItemStock.ListItems.Clear
            Me.PICDIKeyQty.Visible = False
            Me.TBDIBarCode.SetFocus

            Call CalcDITotalAmount
            Exit Sub
         End If
      End If
      
Line1:

      Next i
      
      If Me.ListViewDIItem.ListItems.Count > 0 Then
         vCountItem = Me.ListViewDIItem.ListItems.Count
      Else
         vCountItem = 0
      End If
      
      vCountItem = vCountItem + 1
      vQTY = Me.TBDIKeyQty.Text
      vPrice = Me.LBLDIPrice.Caption
      If Me.LBLDIDiscountWord.Caption <> "" Then
      vDiscountAmount = Me.LBLDIDiscountWord.Caption
      Else
      vDiscountAmount = 0
      End If
      vNetAmount = vQTY * (vPrice - vDiscountAmount)
      
           
     Set vListItem = Me.ListViewDIItem.ListItems.Add(, , vCountItem)
     vListItem.SubItems(1) = Me.LBLDIItemCode.Caption
     vListItem.SubItems(2) = Me.LBLDIItemName.Caption
     vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
     vListItem.SubItems(4) = Me.LBLDIUnitCode.Caption
     vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
     vListItem.SubItems(6) = Format(vNetAmount, "##,##0.00")
     vListItem.SubItems(7) = Format(vDiscountAmount, "##,##0.00")
     vListItem.SubItems(8) = Me.LBLDIWHCode.Caption
     vListItem.SubItems(9) = Me.LBLDIShelfCode.Caption
     vListItem.SubItems(10) = Me.LBLDIZoneID.Caption
     vListItem.SubItems(11) = Me.LBLDIShelfID.Caption
     vListItem.SubItems(12) = Me.LBLDIBarCode.Caption
     vListItem.SubItems(13) = Me.TBDIKeyDiscount.Text
     vListItem.SubItems(14) = vPickZone
     
   End If
Me.TBDIBarCode.Text = ""
Me.LBLDIItemCode.Caption = ""
Me.LBLDIItemName.Caption = ""
Me.LBLDIUnitCode.Caption = ""
Me.TBDIKeyQty.Text = ""
Me.TBDIKeyDiscount.Text = ""
Me.LBLDIPrice.Caption = ""
Me.LBLDIDiscountWord.Caption = ""
Me.LBLDIWHCode.Caption = ""
Me.LBLDIShelfCode.Caption = ""
Me.LBLDIZoneID.Caption = ""
Me.LBLDIShelfID.Caption = ""
Me.LBLDIBarCode.Caption = ""
Me.LBLDIItemNetAmount.Caption = ""
Me.PICDIKeyQty.Visible = False
Me.TBDIBarCode.SetFocus
End If
Call CalcDITotalAmount
End Sub

Private Sub CMDDIKeyQtyOK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICDIKeyQty.Visible = False
   Me.TBDIBarCode.Text = ""
   Me.TBDIBarCode.SetFocus
End If
End Sub

Private Sub CMDDIMain_Click()
Me.PICPickReq.Visible = False
Me.PICDriveIn.Visible = False
Me.PICSelectJob.Visible = True
Me.PICSelectDI.Visible = False
Me.LBLJob.Visible = True
Me.LBLSelectDI.Visible = False
Me.CMDDriveIn.SetFocus
End Sub

Private Sub CMDDIPICSearchAR_Click()
'Me.PICDISearchAR.Visible = False
End Sub

Private Sub CMDDIMain_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 97 Then
'Me.PICDriveIn.Visible = True
'Me.LBLDI.Caption = "01"
'Me.TBDIArCode.SetFocus
'End If

If KeyCode = 98 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "02"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 99 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "03"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 100 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "04"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 27 Then
Call CMDDIMain_Click
End If
End Sub

Private Sub CMDDIPICSearchARClose_Click()
Me.PICDISearchAR.Visible = False
Me.TBDIArCode.SetFocus
End Sub

Private Sub CMDDIPICSearchARClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchAR.Visible = False
Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub CMDDIPICSearchDIClose_Click()
Me.PICDISearchDI.Visible = False
Me.TBDIArCode.SetFocus
End Sub

Private Sub CMDDIPICSearchDIClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchDI.Visible = False
Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub CMDDIPICSearchItemClose_Click()
Me.PICDISearchItem.Visible = False
Me.TBDIBarCode.SetFocus
End Sub

Private Sub CMDDIPICSearchItemClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchItem.Visible = False
Me.TBDIBarCode.SetFocus
End If
End Sub

Private Sub CMDDIPICSearchSaleClose_Click()
Me.PICDISearchSale.Visible = False
Me.TBDISaleCode.SetFocus
End Sub

Private Sub CMDDIPICSearchSaleClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchSale.Visible = False
Me.TBDISaleCode.SetFocus
End If
End Sub

Private Sub CMDDISave_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vHeader As String
Dim vRunning As Integer
Dim vDocNumber As String

Dim i As Integer
Dim vDocNo As String
Dim vDocdate As String
Dim vARCode As String
Dim vSaleCode As String
Dim vMemberID As String
Dim vRefNo As String
Dim vPickZone As String
Dim vBeforeTaxAmount As Double
Dim vTaxAmount As Double
Dim vTotalNetAmount As Double
Dim vNetDebtAmount As Double

Dim vItemCode As String
Dim vItemName As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vShelfID As String
Dim vZoneID As String
Dim vQTY As Double
Dim vUnitCode As String
Dim vPrice As Double
Dim vDiscountWord As String
Dim vDiscountAmount As Double
Dim vAmount As Double
Dim vBarCode As String
Dim vLineNumber As Integer
Dim vLinePickZone As String
Dim vAnswer As Integer

Dim vDay1 As String
Dim vMonth1 As String

Dim vCountItemZone As Integer
Dim vItemPickZone As String
Dim n As Integer

If Me.TBDIArCode.Text = "" Then
   MsgBox "ต้องระบุรหัสลูกค้าก่อนการบันทึกข้อมูล กรุณาตรวจสอบ", vbCritical, "Send Error Message"
   Me.TBDIArCode.SetFocus
   Exit Sub
End If

If Me.TBDISaleCode.Text = "" Then
   MsgBox "ต้องระบุรหัสพนักงานก่อนการบันทึกข้อมูล กรุณาตรวจสอบ", vbCritical, "Send Error Message"
   Me.TBDISaleCode.SetFocus
   Exit Sub
End If

If Me.ListViewDIItem.ListItems.Count = 0 Then
   MsgBox "ไม่สามารถบันทึกเอกสารที่ไม่มีรายการสินค้าได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
   Me.TBDIBarCode.SetFocus
   Exit Sub
End If

vPickZone = Me.LBLDI.Caption
For n = 1 To Me.ListViewDIItem.ListItems.Count
vItemPickZone = Me.ListViewDIItem.ListItems(n).SubItems(14)
If vPickZone = vItemPickZone Then
vCountItemZone = vCountItemZone + 1
End If
Next n

If vCountItemZone = 0 Then
   If vCountItemPickZoneOld = 0 Then
      Call CMDDIClearScreen_Click
      Exit Sub
   End If
End If

vPickZone = Me.LBLDI.Caption

If vDIIsOpen = 0 Then
   vQuery = "exec dbo.USP_NP_SearchNewDocNo  29 "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vHeader = Trim(vRecordset.Fields("header").Value)
       vRunning = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
       vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
   End If
   vRecordset.Close

vDocNo = UCase(Trim(vDocNumber & vHeader & "-" & Format(vRunning, "0000")))

vDocdate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
vARCode = Me.TBDIArCode.Text
If vARCode = "1" Then
   Me.TBDIArCode.Text = "99999"
   vARCode = Me.TBDIArCode.Text
End If
vCheckSale = InStr(Me.TBDISaleCode.Text, "/")
If vCheckSale = 0 Then
   MsgBox "กรุณาตรวจสอบรหัสพนักงานขาย ต้องมีชื่อภาษาไทยปรากฏด้วย ตัวอย่างเช่น  11111/xxxxx เป็นต้น ", vbCritical, "Send Error Message"
   Me.TBDIArCode.SetFocus
   Exit Sub
End If

vSaleCode = Left(Me.TBDISaleCode.Text, vCheckSale - 1)

If Me.LBLDINetAmount.Caption <> "" Then
vNetDebtAmount = Me.LBLDINetAmount.Caption
vBeforeTaxAmount = (vNetDebtAmount * 100) / 107
vTaxAmount = vNetDebtAmount - vBeforeTaxAmount
Else
vNetDebtAmount = 0
vBeforeTaxAmount = 0
vTaxAmount = 0
End If

vTotalNetAmount = vNetDebtAmount
vRefNo = Me.TBDICarLicense.Text
vMemberID = Me.LBLDIMember.Caption
   
vQuery = "exec dbo.USP_NP_InsertDriveInSlip1 '" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vMemberID & "','" & vSaleCode & "','" & vRefNo & "'," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalNetAmount & ",'" & vUserID & "' "
gConnection.Execute vQuery

For i = 1 To Me.ListViewDIItem.ListItems.Count
 vItemCode = Me.ListViewDIItem.ListItems(i).SubItems(1)
 vItemName = Me.ListViewDIItem.ListItems(i).SubItems(2)
 vWHCode = Me.ListViewDIItem.ListItems(i).SubItems(8)
 vShelfCode = Me.ListViewDIItem.ListItems(i).SubItems(9)
 vShelfID = Me.ListViewDIItem.ListItems(i).SubItems(11)
 vZoneID = Me.ListViewDIItem.ListItems(i).SubItems(10)
 vQTY = Me.ListViewDIItem.ListItems(i).SubItems(3)
 vUnitCode = Me.ListViewDIItem.ListItems(i).SubItems(4)
 vPrice = Me.ListViewDIItem.ListItems(i).SubItems(5)
 vDiscountWord = Me.ListViewDIItem.ListItems(i).SubItems(13)
 vDiscountAmount = Me.ListViewDIItem.ListItems(i).SubItems(7)
 vAmount = Me.ListViewDIItem.ListItems(i).SubItems(6)
 vBarCode = Me.ListViewDIItem.ListItems(i).SubItems(12)
 vLinePickZone = Me.ListViewDIItem.ListItems(i).SubItems(14)
 vLineNumber = i - 1

vQuery = "exec dbo.USP_NP_InsertDriveInSlipSub1 '" & vDocNo & "','" & vDocdate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "','" & vLinePickZone & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vAmount & ",'" & vBarCode & "','" & vSaleCode & "'," & vLineNumber & " "
gConnection.Execute vQuery

Next i

MsgBox "บันทึกข้อมูลเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", vbInformation, "Send Information Message"

If vDIIsOpen = 0 Then
vQuery = "exec dbo.USP_NP_UpdateNewDocNo  29"
gConnection.Execute vQuery
End If

vDIIsOpen = 0
Me.ListViewDIItem.ListItems.Clear


If Len(Day(Now)) = 1 Then
   vDay1 = Trim("0" & Day(Now))
Else
   vDay1 = Day(Now)
End If

If Len(Month(Now)) = 1 Then
   vMonth1 = Trim("0" & Month(Now))
Else
   vMonth1 = Month(Now)
End If

Me.DTPDocDate1.Caption = vDay1 & "/" & vMonth1 & "/" & Year(Now)

Me.TBDIArCode.Text = ""
Me.TBDISaleCode.Text = ""
Me.LBLDIDocNo.Caption = ""
Me.LBLDINetAmount.Caption = ""
Me.TBDICarLicense.Text = ""
Me.LBLDIMember.Caption = ""
Me.ListViewDIItem.ListItems.Clear
Me.TBDIBarCode.Text = ""
Me.TBDICarLicense.Text = ""

Call ShowDIDetails(vDocNo)

vAnswer = MsgBox("คุณต้องการส่ง รายการสินค้าไปทำการตรวจนับหรือไม่", vbYesNo, "Send Question Message ?")
If vAnswer = 6 Then
Call CMDDISendQue_Click
Exit Sub
End If
End If

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

If vDIIsOpen = 1 Then

vDocNo = Me.LBLDIDocNo.Caption

vDocdate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
vARCode = Me.TBDIArCode.Text
If vARCode = "1" Then
   Me.TBDIArCode.Text = "99999"
   vARCode = Me.TBDIArCode.Text
End If
vCheckSale = InStr(Me.TBDISaleCode.Text, "/")
If vCheckSale = 0 Then
   MsgBox "กรุณาตรวจสอบรหัสพนักงานขาย ต้องมีชื่อภาษาไทยปรากฏด้วย ตัวอย่างเช่น  11111/xxxxx เป็นต้น ", vbCritical, "Send Error Message"
   Me.TBDIArCode.SetFocus
   Exit Sub
End If

vSaleCode = Left(Me.TBDISaleCode.Text, vCheckSale - 1)

If Me.LBLDINetAmount.Caption <> "" Then
vNetDebtAmount = Me.LBLDINetAmount.Caption
vBeforeTaxAmount = (vNetDebtAmount * 100) / 107
vTaxAmount = vNetDebtAmount - vBeforeTaxAmount
Else
vNetDebtAmount = 0
vBeforeTaxAmount = 0
vTaxAmount = 0
End If

vTotalNetAmount = vNetDebtAmount
vRefNo = Me.TBDICarLicense.Text
vMemberID = Me.LBLDIMember.Caption

vQuery = "exec dbo.USP_NP_CheckQueHoldBill1 '" & vDocNo & "','" & vARCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vMemDIIsConfirm = vRecordset.Fields("isconfirm").Value
vMemDIIsCancel = vRecordset.Fields("iscancel").Value
vDISendQue = vRecordset.Fields("issendque").Value
End If
vRecordset.Close
   
If vMemDIIsCancel = 1 Then
   MsgBox "เอกสารที่ยกเลิกแล้วไม่สามารถแก้ไขข้อมูลได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
   Me.TBDIBarCode.SetFocus
   Exit Sub
End If

If vMemDIIsConfirm = 1 Then
   MsgBox "เอกสารที่ถูกอ้างไปทำใบพักบิลแล้วไม่สามารถแก้ไขข้อมูลได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
   Me.TBDIBarCode.SetFocus
   Exit Sub
End If

vQuery = "exec dbo.USP_NP_InsertDriveInSlip1 '" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vMemberID & "','" & vSaleCode & "','" & vRefNo & "'," & vBeforeTaxAmount & "," & vTaxAmount & "," & vTotalNetAmount & ",'" & vUserID & "' "
gConnection.Execute vQuery

Dim vCountItem As Integer
Dim a As Integer
Dim b As Integer
Dim vOld As Integer

Dim vOldItem As String
Dim vOldUnit As String
Dim vOldBar As String
Dim vOldWH As String
Dim vOldShelf As String
Dim vOldZone As String
Dim vOldPick As String

Dim vCheckItemCode As String
Dim vCheckItemUnit As String
Dim vCheckItemBar As String
Dim vCheckItemWH As String
Dim vCheckItemShelf As String
Dim vCheckItemZone As String
Dim vCheckItemPick As String


vCountItem = Me.ListViewDIItem.ListItems.Count

For a = 1 To vCountItemOld
vOldItem = vDIItemCodeOld(a)
vOldUnit = vDIUnitCodeOld(a)
vOldBar = vDIBarCodeOld(a)
vOldWH = vDIWHCodeOld(a)
vOldShelf = vDIShelfCodeOld(a)
vOldZone = vDIZoneIDOld(a)
vOldPick = vDIPickZoneOld(a)

For b = 1 To Me.ListViewDIItem.ListItems.Count
   vCheckItemCode = Me.ListViewDIItem.ListItems(b).ListSubItems(1).Text
   vCheckItemUnit = Me.ListViewDIItem.ListItems(b).ListSubItems(4).Text
   vCheckItemBar = Me.ListViewDIItem.ListItems(b).ListSubItems(12).Text
   vCheckItemWH = Me.ListViewDIItem.ListItems(b).ListSubItems(8).Text
   vCheckItemShelf = Me.ListViewDIItem.ListItems(b).ListSubItems(9).Text
   vCheckItemZone = Me.ListViewDIItem.ListItems(b).ListSubItems(10).Text
   vCheckItemPick = Me.ListViewDIItem.ListItems(b).ListSubItems(14).Text

   If vCheckItemCode = vOldItem And vCheckItemUnit = vOldUnit And vCheckItemBar = vOldBar And vCheckItemWH = vOldWH And vCheckItemShelf = vOldShelf And vCheckItemZone = vOldZone And vCheckItemPick = vOldPick Then
      vOld = 1
      GoTo Line1
   Else
      vOld = 0
   End If
Next b

Line1:

If vOld = 0 Then
vItemCode = vOldItem
vWHCode = vOldWH
vShelfCode = vOldShelf
vUnitCode = vOldUnit
vBarCode = vOldBar
vZoneID = vOldZone
vLinePickZone = vOldPick

vQuery = "exec dbo.USP_NP_DeleteDriveInSlipSub1 '" & vDocNo & "','" & vDocdate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vZoneID & "','" & vLinePickZone & "','" & vUnitCode & "','" & vBarCode & "'," & vTotalNetAmount & " "
gConnection.Execute vQuery

End If
Next a

For i = 1 To Me.ListViewDIItem.ListItems.Count
 vItemCode = Me.ListViewDIItem.ListItems(i).SubItems(1)
 vItemName = Me.ListViewDIItem.ListItems(i).SubItems(2)
 vWHCode = Me.ListViewDIItem.ListItems(i).SubItems(8)
 vShelfCode = Me.ListViewDIItem.ListItems(i).SubItems(9)
 vShelfID = Me.ListViewDIItem.ListItems(i).SubItems(11)
 vZoneID = Me.ListViewDIItem.ListItems(i).SubItems(10)
 vQTY = Me.ListViewDIItem.ListItems(i).SubItems(3)
 vUnitCode = Me.ListViewDIItem.ListItems(i).SubItems(4)
 vPrice = Me.ListViewDIItem.ListItems(i).SubItems(5)
 vDiscountWord = Me.ListViewDIItem.ListItems(i).SubItems(13)
 vDiscountAmount = Me.ListViewDIItem.ListItems(i).SubItems(7)
 vAmount = Me.ListViewDIItem.ListItems(i).SubItems(6)
 vBarCode = Me.ListViewDIItem.ListItems(i).SubItems(12)
 vLinePickZone = Me.ListViewDIItem.ListItems(i).SubItems(14)
 vLineNumber = i - 1
 
 
If vPickZone = vLinePickZone Then
vQuery = "exec dbo.USP_NP_InsertDriveInSlipSub1 '" & vDocNo & "','" & vDocdate & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "','" & vLinePickZone & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vAmount & ",'" & vBarCode & "','" & vSaleCode & "'," & vLineNumber & " "
gConnection.Execute vQuery
End If

Next i

MsgBox "บันทึกแก้ไขข้อมูลเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", vbInformation, "Send Information Message"

vDIIsOpen = 0
Me.ListViewDIItem.ListItems.Clear

If Len(Day(Now)) = 1 Then
   vDay1 = Trim("0" & Day(Now))
Else
   vDay1 = Day(Now)
End If

If Len(Month(Now)) = 1 Then
   vMonth1 = Trim("0" & Month(Now))
Else
   vMonth1 = Month(Now)
End If

Me.DTPDocDate1.Caption = vDay1 & "/" & vMonth1 & "/" & Year(Now)

Me.TBDIArCode.Text = ""
Me.TBDISaleCode.Text = ""
Me.LBLDIDocNo.Caption = ""
Me.LBLDINetAmount.Caption = ""
Me.TBDICarLicense.Text = ""
Me.LBLDIMember.Caption = ""
Me.ListViewDIItem.ListItems.Clear
Me.TBDIBarCode.Text = ""
Me.TBDICarLicense.Text = ""

Call ShowDIDetails(vDocNo)

vAnswer = MsgBox("คุณต้องการส่ง รายการสินค้าไปทำการตรวจนับหรือไม่", vbYesNo, "Send Question Message ?")
If vAnswer = 6 Then
Call CMDDISendQue_Click
End If
End If

End Sub

Private Sub CMDDISave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub CMDDISearchAr_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String
Dim vListItem As ListItem
Dim i As Integer

vSearchAR = ""
vQuery = "exec dbo.USP_AR_SearchARLine '" & vSearchAR & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.ListViewDISearchAR.ListItems.Clear
    vRecordset.MoveFirst
    For i = 1 To vRecordset.RecordCount
       Set vListItem = Me.ListViewDISearchAR.ListItems.Add(, , i)
       vListItem.SubItems(1) = Trim(vRecordset.Fields("arcode").Value)
       vListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
       vListItem.SubItems(3) = Trim(vRecordset.Fields("memberid").Value)
       vRecordset.MoveNext
    Next i
End If
vRecordset.Close

Me.PICDISearchAR.Visible = True
Me.TBDISearchAR.SetFocus

End Sub

Private Sub CMDDISearchAr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub CMDDISearchARClick_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String
Dim vListItem As ListItem
Dim i As Integer

If Me.TBDISearchAR.Text <> "" Then
   vSearchAR = Me.TBDISearchAR.Text
   vQuery = "exec dbo.USP_AR_SearchARLine '" & vSearchAR & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       Me.ListViewDISearchAR.ListItems.Clear
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
          Set vListItem = Me.ListViewDISearchAR.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("arcode").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
          vListItem.SubItems(3) = Trim(vRecordset.Fields("memberid").Value)
          vRecordset.MoveNext
       Next i
       Me.ListViewDISearchAR.SetFocus
   Else
   Me.TBDISearchAR.SetFocus
   End If
   vRecordset.Close
End If
End Sub

Private Sub CMDDISearchARClick_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchAR.Visible = False
Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub CMDDISearchDIClick_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vSearch As String
Dim vNetDebtAmount As Double

If Me.TBDISearchDI.Text <> "" Then
   vSearch = Me.TBDISearchDI.Text
   Me.ListViewDISearchDI.ListItems.Clear
   vQuery = "exec dbo.usp_np_SearchDriveInMaster1 " & vSearch & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vRecordset.MoveFirst
      For i = 1 To vRecordset.RecordCount
      Set vListItem = Me.ListViewDISearchDI.ListItems.Add(, , i)
      vListItem.SubItems(1) = vRecordset.Fields("docno").Value
      vListItem.SubItems(2) = vRecordset.Fields("docdate").Value
      vListItem.SubItems(3) = vRecordset.Fields("arname").Value
      vListItem.SubItems(4) = vRecordset.Fields("refid").Value
      vListItem.SubItems(5) = vRecordset.Fields("salename").Value
      vNetDebtAmount = vRecordset.Fields("totalnetamount").Value
      vListItem.SubItems(6) = Format(vNetDebtAmount, "##,##0.00")
      vListItem.SubItems(7) = vRecordset.Fields("iscancel").Value
      vListItem.SubItems(8) = vRecordset.Fields("isconfirm").Value
      vRecordset.MoveNext
      Next i
   Me.ListViewDISearchDI.SetFocus
   Else
   Me.TBDISearchDI.SetFocus
   End If
   vRecordset.Close
End If
End Sub

Private Sub CMDDISearchDIClick_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchDI.Visible = False
End If
End Sub

Private Sub CMDDISearchDocNo_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vSearch As String
Dim vNetDebtAmount As Double

vSearch = ""
Me.ListViewDISearchDI.ListItems.Clear
vQuery = "exec dbo.usp_np_SearchDriveInMaster1 '" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   For i = 1 To vRecordset.RecordCount
   Set vListItem = Me.ListViewDISearchDI.ListItems.Add(, , i)
   vListItem.SubItems(1) = vRecordset.Fields("docno").Value
   vListItem.SubItems(2) = vRecordset.Fields("docdate").Value
   vListItem.SubItems(3) = vRecordset.Fields("arname").Value
   vListItem.SubItems(4) = vRecordset.Fields("refid").Value
   vListItem.SubItems(5) = vRecordset.Fields("salename").Value
   vNetDebtAmount = vRecordset.Fields("totalnetamount").Value
   vListItem.SubItems(6) = Format(vNetDebtAmount, "##,##0.00")
   vListItem.SubItems(7) = vRecordset.Fields("iscancel").Value
   vListItem.SubItems(8) = vRecordset.Fields("isconfirm").Value
   vRecordset.MoveNext
   Next i
End If
vRecordset.Close
Me.PICDISearchDI.Visible = True
Me.TBDISearchDI.SetFocus

End Sub

Private Sub CMDDISearchDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub CMDDISearchItem_Click()
Me.PICDISearchItem.Visible = True
Me.TBDISearchItem.SetFocus
End Sub

Private Sub CMDDISearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub CMDDISearchSale_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchSale As String
Dim vListItem As ListItem
Dim i As Integer

Me.PICDISearchSale.Visible = True

vQuery = "exec dbo.USP_CRM_EmployeeDetails 0,'' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.ListViewDISearchSale.ListItems.Clear
    vRecordset.MoveFirst
    For i = 1 To vRecordset.RecordCount
       Set vListItem = Me.ListViewDISearchSale.ListItems.Add(, , i)
       vListItem.SubItems(1) = Trim(vRecordset.Fields("empcode").Value)
       vListItem.SubItems(2) = Trim(vRecordset.Fields("empname").Value)
       vRecordset.MoveNext
    Next i
End If
vRecordset.Close

Me.TBDISearchSale.SetFocus
End Sub

Private Sub CMDDISearchSale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub CMDDISearchSaleClick_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchSale As String
Dim vListItem As ListItem
Dim i As Integer


If Me.TBDISearchSale.Text <> "" Then
vSearchSale = Me.TBDISearchSale.Text
vQuery = "exec dbo.USP_CRM_EmployeeDetails 0,'" & vSearchSale & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.ListViewDISearchSale.ListItems.Clear
    vRecordset.MoveFirst
    For i = 1 To vRecordset.RecordCount
       Set vListItem = Me.ListViewDISearchSale.ListItems.Add(, , i)
       vListItem.SubItems(1) = Trim(vRecordset.Fields("empcode").Value)
       vListItem.SubItems(2) = Trim(vRecordset.Fields("empname").Value)
       vRecordset.MoveNext
    Next i
    Me.ListViewDISearchSale.SetFocus
Else
   Me.TBDISearchSale.SetFocus
End If
vRecordset.Close
End If
End Sub

Private Sub CMDDISearchSaleClick_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchSale.Visible = False
End If
End Sub

Private Sub CMDDISelectPoint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vHeader As String
Dim vRunning As Integer
Dim vDocNumber As String

vQuery = "exec dbo.USP_NP_SearchNewDocNo  29 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vHeader = Trim(vRecordset.Fields("header").Value)
    vRunning = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
End If
vRecordset.Close
      
vDocNo = vDocNumber & vHeader & "-" & Format(vRunning, "0000")
Me.LBLDIDocNo.Caption = vDocNo
vDIIsOpen = 0
vMemDIIsCancel = 0
vDISendQue = 0
Me.TBDIArCode.Text = ""
Me.TBDIArCode.Text = "99999"

Me.TBDISaleCode.Text = ""
Me.LBLDINetAmount.Caption = ""
Me.TBDICarLicense.Text = ""
Me.LBLDIMember.Caption = ""

Me.ListViewDIItem.ListItems.Clear

Call NewDoc
Me.PICDISearchDI.Visible = False
Me.PICDIKeyQty.Visible = False
Me.PICDISearchAR.Visible = False
Me.PICDISearchSale.Visible = False
Me.PICPickReq.Visible = False
Me.PICDriveIn.Visible = False
Me.PICSelectJob.Visible = False
Me.PICSelectDI.Visible = True
Me.LBLJob.Visible = False
Me.LBLSelectDI.Visible = True
Me.CMDDI02.SetFocus
End Sub

Private Sub CMDDISelectPoint_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub CMDDISendInformationClose_Click()
Me.LBLDIInfDocNo.Caption = ""
Me.ListViewDIInfQue.ListItems.Clear
Me.PICDISendInformation.Visible = False
Me.TBDIArCode.SetFocus
End Sub

Private Sub CMDDISendInformationClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.LBLDIInfDocNo.Caption = ""
   Me.ListViewDIInfQue.ListItems.Clear
   Me.PICDISendInformation.Visible = False
   Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub CMDDISendQue_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim n As Integer
Dim vDocNo As String
Dim vDocdate As String
Dim vQueDocDate As String

Dim vCheckTimeID As Integer
Dim vLastTimeID As Integer
Dim vGroupZone(4) As String

Dim vListItem As ListItem
Dim vCheckDate As String

Dim vCheckQueSend As Integer
Dim vQTY As Double
Dim vPickQty As Double
Dim vType As Integer
Dim vPickZone As String
Dim x As Integer
Dim vQueZone As String


If Me.ListViewDIItem.ListItems.Count > 0 And Me.LBLDIDocNo.Caption <> "" And vDIIsOpen = 1 Then
   vDocNo = Me.LBLDIDocNo.Caption
   vDocdate = Me.LBLDIDocDate.Caption
   vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   vPickZone = Me.LBLDI.Caption
   vType = 3
      
   If vPickZone = "01" Then
      vQueZone = "A"
   ElseIf vPickZone = "02" Then
      vQueZone = "B"
   ElseIf vPickZone = "03" Then
      vQueZone = "C"
   ElseIf vPickZone = "04" Then
      vQueZone = "X"
   End If
   
    vQuery = "exec dbo.USP_NP_CheckQuePickCenter1 '" & vDocNo & "','" & vQueDocDate & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vLastTimeID = Trim(vRecordset.Fields("max1").Value)
    End If
    vRecordset.Close
    
   vCheckTimeID = vLastTimeID + 1
   
   vQuery = "exec dbo.USP_NP_CheckQueDriveIn1 '" & vDocNo & "','" & vDocdate & "','" & vQueZone & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   
      Me.ListViewDILastSendQue.ListItems.Clear
      vCheckQueSend = 1
      vRecordset.MoveFirst
      For i = 1 To vRecordset.RecordCount
        Set vListItem = Me.ListViewDILastSendQue.ListItems.Add(, , i)
        vListItem.SubItems(1) = Trim(vRecordset.Fields("itemname").Value)
        vQTY = Trim(vRecordset.Fields("qty").Value)
        vListItem.SubItems(2) = Format(vQTY, "##,##0.00")
        vListItem.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
        vListItem.SubItems(4) = Trim(vRecordset.Fields("queid").Value)
        vListItem.SubItems(5) = Trim(vRecordset.Fields("quezone").Value)
        vListItem.SubItems(6) = Trim(vRecordset.Fields("docno").Value)
        vListItem.SubItems(7) = Trim(vRecordset.Fields("itemcode").Value)
        vListItem.SubItems(8) = Trim(vRecordset.Fields("whcode").Value)
        vListItem.SubItems(9) = Trim(vRecordset.Fields("shelfcode").Value)
        vListItem.SubItems(10) = Trim(vRecordset.Fields("barcode").Value)
        vListItem.SubItems(11) = Trim(vRecordset.Fields("pickzone").Value)
        vListItem.SubItems(12) = Trim(vRecordset.Fields("shelfid").Value)
        vRecordset.MoveNext
      Next i
   
   End If
   vRecordset.Close
      
   If vCheckQueSend = 1 Then
      Me.PICDILastSendQue.Visible = True
      Exit Sub
   End If
   
   
   Dim vDay As String
   Dim vMonth As String
   
   If Len(Day(Now)) = 1 Then
      vDay = Trim("0" & Day(Now))
   Else
      vDay = Day(Now)
   End If
   
   If Len(Month(Now)) = 1 Then
      vMonth = Trim("0" & Month(Now))
   Else
      vMonth = Month(Now)
   End If
   
   vCheckDate = vDay & "/" & vMonth & "/" & Year(Now)
   
   If vDocdate < vCheckDate Then
      MsgBox "ไม่สามารถส่งคิวจัดสินค้าได้ กรณีวันที่เอกสารน้อยกว่าวันที่ปัจจุบัน", vbCritical, "Send Error Message"
      Me.TBDIBarCode.SetFocus
      Exit Sub
   End If
   
   vQuery = "exec dbo.USP_NP_SearchGroupPicking1 " & vType & ",'" & vDocNo & "','" & vPickZone & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      n = vRecordset.RecordCount
      For i = 1 To vRecordset.RecordCount
      vGroupZone(i) = Trim(vRecordset.Fields("zoneid").Value)
      Next i
   End If
   vRecordset.Close
   
   For x = 1 To n
      If vGroupZone(x) = "A" Then
         Call PrintPicking_A(vDocNo, vCheckTimeID, 3)
      ElseIf vGroupZone(x) = "B" Then
         Call PrintPicking_B(vDocNo, vCheckTimeID, 3)
      ElseIf vGroupZone(x) = "C" Then
         Call PrintPicking_C(vDocNo, vCheckTimeID, 3)
      ElseIf vGroupZone(x) = "X" Then
         Call PrintPicking_X(vDocNo, vCheckTimeID, 3)
      End If
   Next x
   
   'Call PrintDriveInDetails(vDocNo)

   vDIIsOpen = 0
   Me.ListViewDIItem.ListItems.Clear
   
   Dim vDay1 As String
   Dim vMonth1 As String
   
   If Len(Day(Now)) = 1 Then
      vDay1 = Trim("0" & Day(Now))
   Else
      vDay1 = Day(Now)
   End If
   
   If Len(Month(Now)) = 1 Then
      vMonth1 = Trim("0" & Month(Now))
   Else
      vMonth1 = Month(Now)
   End If
   
   Me.LBLDIDocDate.Caption = vDay1 & "/" & vMonth1 & "/" & Year(Now)

   Me.TBDIArCode.Text = "99999"
   Me.TBDISaleCode.Text = ""
   Me.LBLDIDocNo.Caption = ""
   Me.TBDICarLicense.Text = ""

   Me.LBLDINetAmount.Caption = ""
   Me.CMDDISendQue.Enabled = False
   
   Me.PICDISendInformation.Visible = True
   vQuery = "exec dbo.usp_np_SearchReqPickingInformation '" & vDocNo & "'," & vCheckTimeID & " "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.ListViewDIInfQue.ListItems.Clear
      Me.LBLDIInfDocNo.Caption = Trim(vRecordset.Fields("docno").Value)
      vRecordset.MoveFirst
      For i = 1 To vRecordset.RecordCount
        Set vListItem = Me.ListViewDIInfQue.ListItems.Add(, , i)
        vListItem.SubItems(1) = Trim(vRecordset.Fields("queid").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("zoneid").Value)
        vRecordset.MoveNext
      Next i
   End If
   vRecordset.Close
    
   vQuery = "exec dbo.USP_NP_UpdateSendQueuePicking 3,'" & vDocNo & "' "
   gConnection.Execute vQuery
    
   Me.TBDIArCode.Text = ""
   Me.TBDIArCode.Text = "99999"
   Me.ListViewDIInfQue.SetFocus
   
Else
   MsgBox "เอกสารที่จะส่งจัดสินค้าได้ ต้องมีเลขที่เอกสาร รายการสินค้า และต้องเป็นเอกสารที่บันทึกข้อมูลเรียบร้อยแล้วเป็นอย่างน้อย กรุณาตรวจสอบ ", vbCritical, "Send Error Message"
   Me.TBDIBarCode.SetFocus
End If
End Sub

Private Sub CMDDISendQue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub CMDDISendQueExit_Click()
Me.PICDILastSendQue.Visible = False
Me.TBDIArCode.SetFocus
End Sub

Private Sub CMDDISendQueExit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDILastSendQue.Visible = False
Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub CMDDISendQuePrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vQueID As Integer
Dim vQueDocDate As String
Dim vUnitCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vBarCode As String
Dim vPickZone As String
Dim vQTY As Double
Dim vShelfID As String
Dim vARCode As String
Dim vItemCode As String
Dim vItemPickZone As String
Dim vMemItemExist  As Integer

Dim vLastQueID As Integer
Dim vLastQueDocDate As String
Dim vLastDocNo As String
Dim vLastItemCode As String
Dim vLastUnitCode As String
Dim vLastWHCode As String
Dim vLastShelfCode As String
Dim vLastBarCode As String
Dim vLastPickZone As String
Dim vLastZoneID As String
Dim vLastShelfID As String
Dim vLastQTY As Double

Dim vCheckIsConfirm  As Integer
Dim vCheckHoldBillNo As String
Dim vQueZone As String

Dim a As Integer
Dim b As Integer
            
vPickZone = Me.LBLDI.Caption
vDocNo = Me.LBLDIDocNo.Caption
vARCode = Me.TBDIArCode.Text

vQuery = "exec dbo.USP_NP_CheckQueHoldBill1'" & vDocNo & "','" & vARCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vCheckIsConfirm = vRecordset.Fields("isconfirm").Value
   vCheckHoldBillNo = vRecordset.Fields("holdbillno").Value
End If
vRecordset.Close

If vCheckIsConfirm = 1 And vCheckHoldBillNo <> "" Then
    MsgBox "", vbCritical, "Send Error Message"
    Exit Sub
End If

If vPickZone = "01" Then
   vQueZone = "A"
ElseIf vPickZone = "02" Then
   vQueZone = "B"
ElseIf vPickZone = "03" Then
   vQueZone = "C"
ElseIf vPickZone = "04" Then
   vQueZone = "X"
End If
   
If Me.ListViewDIItem.ListItems.Count > 0 Then

For a = 1 To Me.ListViewDILastSendQue.ListItems.Count

vLastQueID = Me.ListViewDILastSendQue.ListItems(a).ListSubItems(4).Text
vLastQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
vLastDocNo = Me.ListViewDILastSendQue.ListItems(a).ListSubItems(6).Text
vLastItemCode = Me.ListViewDILastSendQue.ListItems(a).ListSubItems(7).Text
vLastUnitCode = Me.ListViewDILastSendQue.ListItems(a).ListSubItems(3).Text
vLastWHCode = Me.ListViewDILastSendQue.ListItems(a).ListSubItems(8).Text
vLastShelfCode = Me.ListViewDILastSendQue.ListItems(a).ListSubItems(9).Text
vLastBarCode = Me.ListViewDILastSendQue.ListItems(a).ListSubItems(10).Text
vLastPickZone = Me.ListViewDILastSendQue.ListItems(a).ListSubItems(11).Text
vLastZoneID = Me.ListViewDILastSendQue.ListItems(a).ListSubItems(5).Text
vLastShelfID = Me.ListViewDILastSendQue.ListItems(a).ListSubItems(12).Text
vLastQTY = Me.ListViewDILastSendQue.ListItems(a).ListSubItems(2).Text

For b = 1 To Me.ListViewDIItem.ListItems.Count
vItemCode = Me.ListViewDIItem.ListItems(b).ListSubItems(1).Text
vUnitCode = Me.ListViewDIItem.ListItems(b).ListSubItems(4).Text
vWHCode = Me.ListViewDIItem.ListItems(b).ListSubItems(8).Text
vShelfCode = Me.ListViewDIItem.ListItems(b).ListSubItems(9).Text
vBarCode = Me.ListViewDIItem.ListItems(b).ListSubItems(12).Text
vItemPickZone = Me.ListViewDIItem.ListItems(b).ListSubItems(14).Text
vQTY = Me.ListViewDIItem.ListItems(b).ListSubItems(3).Text

If vLastItemCode = vItemCode And vLastUnitCode = vUnitCode And vLastWHCode = vWHCode And vLastShelfCode = vShelfCode And vLastPickZone = vItemPickZone And vLastBarCode = vBarCode Then
vMemItemExist = 1
GoTo Line1
Else
vMemItemExist = 0
End If

Next
Line1:

If vMemItemExist = 0 And vLastPickZone = vPickZone Then
vQuery = "exec dbo.USP_NP_InsertUpdateCancelQueItem1 1," & vLastQueID & ",'" & vLastQueDocDate & "','" & vLastItemCode & "','" & vLastWHCode & "','" & vLastShelfCode & "','" & vLastShelfID & "','" & vLastZoneID & "','" & vLastPickZone & "','" & vLastDocNo & "','" & vLastBarCode & "'," & vLastQTY & ",'" & vLastUnitCode & "'"
gConnection.Execute vQuery
End If
Next

For a = 1 To Me.ListViewDIItem.ListItems.Count
vQueID = Me.ListViewDILastSendQue.ListItems(1).ListSubItems(4).Text
vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
vDocNo = Me.LBLDIDocNo.Caption
vItemCode = Me.ListViewDIItem.ListItems(a).ListSubItems(1).Text
vUnitCode = Me.ListViewDIItem.ListItems(a).ListSubItems(4).Text
vWHCode = Me.ListViewDIItem.ListItems(a).ListSubItems(8).Text
vShelfCode = Me.ListViewDIItem.ListItems(a).ListSubItems(9).Text
vBarCode = Me.ListViewDIItem.ListItems(a).ListSubItems(12).Text
vItemPickZone = Me.ListViewDIItem.ListItems(a).ListSubItems(14).Text
vQTY = Me.ListViewDIItem.ListItems(a).ListSubItems(3).Text
vShelfID = Me.ListViewDIItem.ListItems(a).ListSubItems(11).Text

For b = 1 To Me.ListViewDILastSendQue.ListItems.Count
vLastItemCode = Me.ListViewDILastSendQue.ListItems(b).ListSubItems(7).Text
vLastUnitCode = Me.ListViewDILastSendQue.ListItems(b).ListSubItems(3).Text
vLastWHCode = Me.ListViewDILastSendQue.ListItems(b).ListSubItems(8).Text
vLastShelfCode = Me.ListViewDILastSendQue.ListItems(b).ListSubItems(9).Text
vLastBarCode = Me.ListViewDILastSendQue.ListItems(b).ListSubItems(10).Text
vLastPickZone = Me.ListViewDILastSendQue.ListItems(b).ListSubItems(11).Text
vLastZoneID = Me.ListViewDILastSendQue.ListItems(b).ListSubItems(5).Text


If vLastItemCode = vItemCode And vLastUnitCode = vUnitCode And vLastWHCode = vWHCode And vLastShelfCode = vShelfCode And vLastPickZone = vPickZone And vLastBarCode = vBarCode Then
vMemItemExist = 1
GoTo Line2
Else
vMemItemExist = 0
End If

Next
Line2:
If vItemPickZone = vPickZone Then
vQuery = "exec dbo.USP_NP_InsertUpdateCancelQueItem1 2," & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vQueZone & "','" & vPickZone & "','" & vDocNo & "','" & vBarCode & "'," & vQTY & ",'" & vUnitCode & "'"
gConnection.Execute vQuery
End If
Next
End If

MsgBox "แก้ไข รายการสินค้าเพื่อไปทำการ CheckOut ใหม่เรียบร้อยแล้ว", vbInformation, "Send Information Message"
Me.ListViewDILastSendQue.ListItems.Clear
Me.PICDILastSendQue.Visible = False
Me.TBDIArCode.SetFocus

End Sub

Private Sub CMDDocNo_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vHeader As String
Dim vRunning As Integer
Dim vDocNumber As String

vQuery = "exec dbo.USP_NP_SearchNewDocNo  32 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vHeader = Trim(vRecordset.Fields("header").Value)
    vRunning = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
End If
vRecordset.Close
      
vDocNo = vDocNumber & vHeader & "-" & Format(vRunning, "0000")
Me.TXTDocNo.Text = vDocNo
Me.TXTDocNo.SetFocus
End Sub

Private Sub CMDDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If
End Sub

Private Sub CMDDriveIn_Click()
Me.PICSelectDI.Visible = True
Me.CMDDI02.SetFocus
Me.PICSelectJob.Visible = False
Me.LBLJob.Visible = False
Me.LBLSelectDI.Visible = True
End Sub

Private Sub CMDDriveIn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 97 Then
   Me.PICSelectDI.Visible = True
   Me.PICPickReq.Visible = False
   Me.PICSelectJob.Visible = False
   Me.LBLJob.Visible = False
   Me.LBLSelectDI.Visible = True
   Me.CMDDI02.SetFocus
End If

If KeyCode = 98 Then
   Me.PICSelectDI.Visible = False
   Me.PICPickReq.Visible = True
   Me.PICSelectJob.Visible = False
   Me.TXTDocNo.SetFocus
End If

End Sub

Private Sub CMDEditExit_Click()
   'Me.LBLEditItemCode.Caption = ""
   'Me.LBLEditItemName.Caption = ""
   'Me.LBLEditItemQty.Caption = ""
   'Me.LBLEditUnitCode.Caption = ""
   'Me.LBLEditDiscount.Caption = ""
   'Me.LBLEditIndex.Caption = ""
   'Me.LBLEditItemAmount.Caption = ""
   'Me.LBLEditPrice.Caption = ""
   'Me.LBLEditRemain.Caption = ""
   'Me.LBLEditItemQty.Caption = ""
   ''Me.TBEditQty.Text = ""
   'Me.PICEditOrder.Visible = False
   'Call CalcEditItemQty
End Sub

Private Sub CMDEditOK_Click()
'Dim vIndex As Integer
'Dim vQTY As Double
'Dim vEditQty As Double
'Dim vPrice As Double
'Dim vDiscountAmount As Double
'Dim vAmount As Double
''
'If Me.TBEditQty.Text <> "" Then
  ' vIndex = Me.LBLEditIndex.Caption
   'vEditQty = Me.TBEditQty.Text
   'If vEditQty = 0 Then
    '  MsgBox "ไม่สามารถกรอกจำนวนที่จะสั่งจัด เท่ากับ 0 กรุณาตรวจสอบ กรณีไม่สั่งจัด ก็ให้ลบรายการออกจากเอกสาร", vbCritical, "Send Error Message"
     ' Exit Sub
   'End If
   
   'vQTY = Me.LBLEditRemain.Caption
   
   'If vEditQty > vQTY Then
    '  MsgBox "ไม่สามารถกรอกจำนวนที่จะสั่งจัด มากกว่าจำนวนคงเหลือในใบสั่งขาย/จอง", vbCritical, "Send Error Message"
     ' Exit Sub
   'End If
   
   'vPrice = Me.LBLEditPrice.Caption
   'vDiscountAmount = Me.LBLEditDiscount.Caption
   'vAmount = vEditQty * (vPrice - vDiscountAmount)
   'Me.LBLEditItemAmount.Caption = Format(vAmount, "##,##0.00")
   '
   'Me.ListViewSaleOrder.ListItems(vIndex).SubItems(3) = Format(vEditQty, "##,##0.00")
   'Me.ListViewSaleOrder.ListItems(vIndex).SubItems(7) = Format(vAmount, "##,##0.00")
   '
   'Me.LBLEditItemCode.Caption = ""
   'Me.LBLEditItemName.Caption = ""
   'Me.LBLEditItemQty.Caption = ""
   'Me.LBLEditUnitCode.Caption = ""
   'Me.LBLEditDiscount.Caption = ""
   'Me.LBLEditIndex.Caption = ""
   'Me.LBLEditItemAmount.Caption = ""
   ''Me.LBLEditPrice.Caption = ""
   'Me.LBLEditRemain.Caption = ""
   ''Me.LBLEditItemQty.Caption = ""
   'Me.TBEditQty.Text = ""
   'Me.PICEditOrder.Visible = False
   'Call CalcEditItemQty

'End If
End Sub


Private Sub CMDExit_Click()
Me.PICSelectDI.Visible = False
Me.PICSelectJob.Visible = True
Me.CMDDriveIn.SetFocus
End Sub

Public Sub CalcTotalAmount()
Dim i As Integer
Dim vLineAmount As Double
Dim vTotalAmount As Double

If Me.ListViewItem.ListItems.Count > 0 Then
   For i = 1 To Me.ListViewItem.ListItems.Count
   vLineAmount = Me.ListViewItem.ListItems(i).SubItems(6)
   vTotalAmount = vTotalAmount + vLineAmount
   Next i
   Me.LBLTotalAmount.Caption = Format(vTotalAmount, "##,##0.00")
Else
   Me.LBLTotalAmount.Caption = Format(0, "##,##0.00")
End If
End Sub

Public Sub CalcDITotalAmount()
Dim i As Integer
Dim vLineAmount As Double
Dim vTotalAmount As Double

If Me.ListViewDIItem.ListItems.Count > 0 Then
   For i = 1 To Me.ListViewDIItem.ListItems.Count
   vLineAmount = Me.ListViewDIItem.ListItems(i).SubItems(6)
   vTotalAmount = vTotalAmount + vLineAmount
   Next i
   Me.LBLDINetAmount.Caption = Format(vTotalAmount, "##,##0.00")
Else
   Me.LBLDINetAmount.Caption = Format(0, "##,##0.00")
End If
End Sub

Private Sub CMDInfClose_Click()
Me.LBLInfDocNo.Caption = ""
Me.LBLInfARName.Caption = ""
Me.ListViewInfQue.ListItems.Clear
Me.PICSendInformation.Visible = False
Me.TXTDocNo.SetFocus

Call CMDDocNo_Click
End Sub

Private Sub CMDInfClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSendInformation.Visible = False
Me.TXTDocNo.SetFocus
End If
End Sub

Private Sub CMDISearchItemClick_Click()
Dim vSearch As String
Dim vPrice As Double
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vQTY As Double
Dim vRemainOutQTY As Double

If Me.TBDISearchItem.Text <> "" Then
   vSearch = Me.TBDISearchItem.Text
   vQuery = "exec dbo.USP_NP_SearchBarCode '" & vSearch & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       Me.ListViewDISearchItem.ListItems.Clear
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
          vQTY = Trim(vRecordset.Fields("stockqty").Value)
          vRemainOutQTY = Trim(vRecordset.Fields("remainoutqty").Value)
          vPrice = Trim(vRecordset.Fields("price").Value)
          
          Set vListItem = Me.ListViewDISearchItem.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
          vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
          vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
          vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
          vListItem.SubItems(6) = Trim(vRecordset.Fields("zone").Value)
          vRecordset.MoveNext
       Next i
       
       Me.ListViewDISearchItem.SetFocus
    Else
       Me.ListViewDISearchItem.ListItems.Clear
       Me.TBDISearchItem.SetFocus
    End If
 vRecordset.Close
End If
End Sub

Private Sub CMDISearchItemClick_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchItem.Visible = False
End If
End Sub

Private Sub CMDOK_Click()
Dim vListItem As ListItem
Dim i As Integer
Dim vBeforeTaxAmount As Double
Dim vTaxAmount As Double
Dim vNetAmount As Double
Dim vQTY As Double
Dim vPrice As Double
Dim vNetprice As Double
Dim vDiscountAmount As Double

Dim vWHCode As String
Dim vShelfCode As String
Dim vUnitCode As String

Dim vItemCode As String
Dim vCheckItemCode As String
Dim vCheckWHCode As String
Dim vCheckShelfCode As String
Dim vCheckUnitCode As String
Dim vAnswer As Integer
Dim vCountItem As Integer


If Me.LBLItemCode.Caption <> "" And Me.TBQty.Text <> "" Then
   vItemCode = Me.LBLItemCode.Caption
   vWHCode = Me.LBLWHCode.Caption
   vShelfCode = Me.LBLShelfCode.Caption
   vUnitCode = Me.LBLUnitCode.Caption
   vQTY = Me.TBQty.Text
   
   If vQTY = 0 Then
      MsgBox "ต้องกรอกจำนวนสินค้าที่ต้องการที่ไม่เท่ากับ 0  เท่านั้น", vbCritical, "Send Error Message"
      Me.TBQty.SetFocus
      Exit Sub
   End If
         
   If Me.ListViewItem.ListItems.Count = 0 Then

      vCountItem = 0
      vCountItem = vCountItem + 1
      vQTY = Me.TBQty.Text
      vPrice = Me.LBLPrice.Caption
      If Me.LBLDisCountAmount.Caption <> "" Then
      vDiscountAmount = Me.LBLDisCountAmount.Caption
      Else
      vDiscountAmount = 0
      End If
      vNetAmount = vQTY * (vPrice - vDiscountAmount)
                 
     Set vListItem = Me.ListViewItem.ListItems.Add(, , vCountItem)
     vListItem.SubItems(1) = Me.LBLItemCode.Caption
     vListItem.SubItems(2) = Me.LBLItemName.Caption
     vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
     vListItem.SubItems(4) = Me.LBLUnitCode.Caption
     vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
     vListItem.SubItems(6) = Format(vNetAmount, "##,##0.00")
     vListItem.SubItems(7) = Format(vDiscountAmount, "##,##0.00")
     vListItem.SubItems(8) = Me.LBLWHCode.Caption
     vListItem.SubItems(9) = Me.LBLShelfCode.Caption
     vListItem.SubItems(10) = Me.LBLZoneID.Caption
     vListItem.SubItems(11) = Me.LBLShelfID.Caption
     vListItem.SubItems(12) = Me.LBLBarCode.Caption
     vListItem.SubItems(13) = Me.TXTDisCount.Text
   Else
      For i = 1 To Me.ListViewItem.ListItems.Count
      vCheckItemCode = Me.ListViewItem.ListItems(i).SubItems(1)
      vCheckWHCode = Me.ListViewItem.ListItems(i).SubItems(8)
      vCheckShelfCode = Me.ListViewItem.ListItems(i).SubItems(9)
      vCheckUnitCode = Me.ListViewItem.ListItems(i).SubItems(4)
      
      If vItemCode = vCheckItemCode And vWHCode = vCheckWHCode And vShelfCode = vCheckShelfCode And vUnitCode = vCheckUnitCode Then
         vAnswer = MsgBox("มีรายการสินค้ารหัส " & vItemCode & " คลัง " & vWHCode & " ชั้นเก็บ " & vShelfCode & "นี้อยู่แล้ว ที่บรรทัดที่ " & i & " จะบันทึกทับข้อมูลเก่าหรือไม่ ", vbYesNo, "Send Information Message")
         If vAnswer = 7 Then
            Exit Sub
         Else
         
           vQTY = Me.TBQty.Text
           vPrice = Me.LBLPrice.Caption
           If Me.LBLDisCountAmount.Caption <> "" Then
           vDiscountAmount = Me.LBLDisCountAmount.Caption
           Else
           vDiscountAmount = 0
           End If
           vNetAmount = vQTY * (vPrice - vDiscountAmount)

            Me.ListViewItem.ListItems(i).SubItems(3) = Format(vQTY, "##,##0.00")
            Me.ListViewItem.ListItems(i).SubItems(5) = Format(vPrice, "##,##0.00")
            Me.ListViewItem.ListItems(i).SubItems(6) = Format(vNetAmount, "##,##0.00")
            Me.ListViewItem.ListItems(i).SubItems(7) = Format(vDiscountAmount, "##,##0.00")
            Me.ListViewItem.ListItems(i).SubItems(13) = Me.TXTDisCount.Text
            
            Me.TBBarCode.Text = ""
            Me.LBLItemCode.Caption = ""
            Me.LBLItemName.Caption = ""
            Me.LBLUnitCode.Caption = ""
            Me.TBQty.Text = ""
            Me.TXTDisCount.Text = ""
            Me.LBLPrice.Caption = ""
            Me.LBLDisCountAmount.Caption = ""
            Me.LBLWHCode.Caption = ""
            Me.LBLShelfCode.Caption = ""
            Me.LBLZoneID.Caption = ""
            Me.LBLShelfID.Caption = ""
            Me.LBLBarCode.Caption = ""
            Me.LBLNetPrice.Caption = ""
            Me.ListViewStock.ListItems.Clear
            Me.PICSearchItem.Visible = False
            Me.TBBarCode.SetFocus

            Call CalcTotalAmount
            Exit Sub
         End If
      End If
      
Line1:

      Next i
      
      If Me.ListViewItem.ListItems.Count > 0 Then
         vCountItem = Me.ListViewItem.ListItems.Count
      Else
         vCountItem = 0
      End If
      
      vCountItem = vCountItem + 1
      vQTY = Me.TBQty.Text
      vPrice = Me.LBLPrice.Caption
      If Me.LBLDisCountAmount.Caption <> "" Then
      vDiscountAmount = Me.LBLDisCountAmount.Caption
      Else
      vDiscountAmount = 0
      End If
      vNetAmount = vQTY * (vPrice - vDiscountAmount)
      
           
     Set vListItem = Me.ListViewItem.ListItems.Add(, , vCountItem)
     vListItem.SubItems(1) = Me.LBLItemCode.Caption
     vListItem.SubItems(2) = Me.LBLItemName.Caption
     vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
     vListItem.SubItems(4) = Me.LBLUnitCode.Caption
     vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
     vListItem.SubItems(6) = Format(vNetAmount, "##,##0.00")
     vListItem.SubItems(7) = Format(vDiscountAmount, "##,##0.00")
     vListItem.SubItems(8) = Me.LBLWHCode.Caption
     vListItem.SubItems(9) = Me.LBLShelfCode.Caption
     vListItem.SubItems(10) = Me.LBLZoneID.Caption
     vListItem.SubItems(11) = Me.LBLShelfID.Caption
     vListItem.SubItems(12) = Me.LBLBarCode.Caption
     vListItem.SubItems(13) = Me.TXTDisCount.Text
     
   End If
Me.TBBarCode.Text = ""
Me.LBLItemCode.Caption = ""
Me.LBLItemName.Caption = ""
Me.LBLUnitCode.Caption = ""
Me.TBQty.Text = ""
Me.TXTDisCount.Text = ""
Me.LBLPrice.Caption = ""
Me.LBLDisCountAmount.Caption = ""
Me.LBLWHCode.Caption = ""
Me.LBLShelfCode.Caption = ""
Me.LBLZoneID.Caption = ""
Me.LBLShelfID.Caption = ""
Me.LBLBarCode.Caption = ""
Me.LBLNetPrice.Caption = ""
Me.PICSearchItem.Visible = False
Me.TBBarCode.SetFocus
End If

Call CalcTotalAmount
End Sub

Private Sub CMDOK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSearchItem.Visible = False
Me.TBBarCode.Text = ""
Me.TBBarCode.SetFocus
End If
End Sub

Private Sub CMDPickReq_Click()
Me.PICSelectDI.Visible = False
Me.PICPickReq.Visible = True
Me.PICSelectJob.Visible = False
Me.TXTDocNo.SetFocus
End Sub

Private Sub CMDPickReq_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 97 Then
   Me.PICSelectDI.Visible = True
   Me.PICPickReq.Visible = False
   Me.PICSelectJob.Visible = False
   Me.LBLJob.Visible = False
   Me.LBLSelectDI.Visible = True
   Me.CMDDI02.SetFocus
End If

If KeyCode = 98 Then
   Me.PICSelectDI.Visible = False
   Me.PICPickReq.Visible = True
   Me.PICSelectJob.Visible = False
   Me.TXTDocNo.SetFocus
End If

End Sub

Private Sub CMDPICPRArKeyDataClose_Click()
'Me.TBPRKeyMember.Text = ""
'Me.PICPRArKeyData.Visible = False
'Me.TXTArCode.SetFocus
End Sub

Private Sub CMDPICPRArKeyDataClose_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 27 Then
 '  Me.PICPRArKeyData.Visible = False
  ' Me.TXTArCode.SetFocus
'End If
End Sub

Private Sub CMDPRMain_Click()
Me.PICDriveIn.Visible = False
Me.PICPickReq.Visible = False
Me.PICSelectJob.Visible = True
Me.PICSelectDI.Visible = False
Me.CMDDriveIn.SetFocus
End Sub

Private Sub CMDPRMain_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If
End Sub

Private Sub CMDPRPICPRSearchARClose_Click()
Me.PICPRSearchAR.Visible = False
Me.TXTArCode.SetFocus
End Sub

Private Sub CMDPRPICPRSearchARClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchAR.Visible = False
Me.TXTArCode.SetFocus
End If
End Sub

Private Sub CMDPRPICPRSearchDocNoClose_Click()
Me.PICPRSearchDocNo.Visible = False
Me.TXTDocNo.SetFocus
End Sub

Private Sub CMDPRPICPRSearchDocNoClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchDocNo.Visible = False
Me.TXTDocNo.SetFocus
End If
End Sub

Private Sub CMDPRPICPRSearchItemClose_Click()
Me.PICPRSearchItem.Visible = False
Me.TBBarCode.SetFocus
End Sub

Private Sub CMDPRPICPRSearchItemClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchItem.Visible = False
Me.TBBarCode.SetFocus
End If
End Sub

Private Sub CMDPRPICPRSearchSaleClose_Click()
Me.PICPRSearchSale.Visible = False
Me.TXTSaleCode.SetFocus
End Sub

Private Sub CMDPRPICPRSearchSaleClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchSale.Visible = False
Me.TXTSaleCode.SetFocus
End If
End Sub

Private Sub CMDPRSearchARClick_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String
Dim vListItem As ListItem
Dim i As Integer

If Me.TBPRSearchAR.Text <> "" Then
   vSearchAR = Me.TBPRSearchAR.Text
   vQuery = "exec dbo.USP_AR_ARProFileSearch '" & vSearchAR & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       Me.ListViewPRSearchAR.ListItems.Clear
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
          Set vListItem = Me.ListViewPRSearchAR.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("arcode").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
          vListItem.SubItems(3) = Trim(vRecordset.Fields("memberid").Value)
          vRecordset.MoveNext
       Next i
       Me.ListViewPRSearchAR.SetFocus
   End If
   vRecordset.Close
End If
End Sub

Private Sub CMDPRSearchARClick_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchAR.Visible = False
Me.TXTArCode.SetFocus
End If
End Sub

Private Sub CMDPRSearchDocNoClick_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearch As String
Dim vListItem As ListItem
Dim i As Integer

   If Me.TBPRSearchDocNo.Text <> "" Then
   vSearch = Me.TBPRSearchDocNo.Text
   vQuery = "exec dbo.USP_NP_SearchReqPicking '" & vSearch & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       Me.ListViewPRSearchDocNo.ListItems.Clear
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
          Set vListItem = Me.ListViewPRSearchDocNo.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
          vListItem.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
          vListItem.SubItems(4) = Trim(vRecordset.Fields("salename").Value)
          vListItem.SubItems(5) = Trim(vRecordset.Fields("netdebtamount").Value)
          vRecordset.MoveNext
       Next i
       Me.ListViewPRSearchDocNo.SetFocus
   Else
   Me.TBPRSearchDocNo.SetFocus
   End If
   vRecordset.Close
   End If
End Sub

Private Sub CMDPRSearchDocNoClick_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchDocNo.Visible = False
Me.TXTDocNo.SetFocus
End If
End Sub

Private Sub CMDPRSearchItemClick_Click()
Dim vSearch As String
Dim vPrice As Double
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vQTY As Double
Dim vRemainOutQTY As Double

If Me.TBPRSearchItem.Text <> "" Then
   vSearch = Me.TBPRSearchItem.Text
   vQuery = "exec dbo.USP_NP_SearchBarCode '" & vSearch & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       Me.ListViewPRSearchItem.ListItems.Clear
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
          vQTY = Trim(vRecordset.Fields("stockqty").Value)
          vRemainOutQTY = Trim(vRecordset.Fields("remainoutqty").Value)
          vPrice = Trim(vRecordset.Fields("price").Value)
          
          Set vListItem = Me.ListViewPRSearchItem.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
          vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
          vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
          vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
          vListItem.SubItems(6) = Trim(vRecordset.Fields("zone").Value)
          vRecordset.MoveNext
       Next i
       
       Me.ListViewPRSearchItem.SetFocus
    Else
       Me.ListViewPRSearchItem.ListItems.Clear
       Me.TBPRSearchItem.SetFocus
    End If
 vRecordset.Close
End If
End Sub

Private Sub CMDPRSearchItemClick_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchItem.Visible = False
Me.TBBarCode.SetFocus
End If
End Sub

Private Sub CMDPRSearchSaleClick_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchSale As String
Dim vListItem As ListItem
Dim i As Integer

If Me.TBPRSearchSale.Text <> "" Then
   vSearchSale = Me.TBPRSearchSale.Text
   vQuery = "exec dbo.USP_CRM_EmployeeDetails 0,'" & vSearchSale & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       Me.ListViewPRSearchSale.ListItems.Clear
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
          Set vListItem = Me.ListViewPRSearchSale.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("empcode").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("empname").Value)
          vRecordset.MoveNext
       Next i
    Me.ListViewPRSearchSale.SetFocus
   End If
   vRecordset.Close
End If
End Sub

Private Sub CMDPRSearchSaleClick_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchSale.Visible = False
Me.TXTSaleCode.SetFocus
End If
End Sub

Private Sub CMDQue_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim n As Integer
Dim vDocNo As String
Dim vDocdate As String
Dim vQueDocDate As String

Dim vCheckTimeID As Integer
Dim vLastTimeID As Integer
Dim vGroupZone(5) As String

Dim vListItem As ListItem
Dim vCheckDate As String

Dim vCheckQueSend As Integer
Dim vQTY As Double
Dim vPickQty As Double

If Me.ListViewItem.ListItems.Count > 0 And Me.TXTDocNo.Text <> "" And vPRIsOpen = 1 Then
   vDocNo = Me.TXTDocNo.Text
   vDocdate = Me.DTPDocDate1.Caption
   vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   
    vQuery = "exec dbo.USP_NP_CheckQuePickCenter '" & vDocNo & "','" & vQueDocDate & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vLastTimeID = Trim(vRecordset.Fields("max1").Value)
    End If
    vRecordset.Close
    
   vCheckTimeID = vLastTimeID + 1
   
   vQuery = "exec dbo.usp_np_SearchReqPickingInformationLastSend '" & vDocNo & "','" & vQueDocDate & "'," & vLastTimeID & " "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.ListViewLastSendQue.ListItems.Clear
      vCheckQueSend = 1
      vRecordset.MoveFirst
      For i = 1 To vRecordset.RecordCount
        Set vListItem = Me.ListViewLastSendQue.ListItems.Add(, , i)
        vListItem.SubItems(1) = Trim(vRecordset.Fields("queid").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("quedescription").Value)
        vListItem.SubItems(3) = Trim(vRecordset.Fields("itemcode").Value) & "/" & Trim(vRecordset.Fields("itemname").Value)
        vQTY = Trim(vRecordset.Fields("qty").Value)
        vPickQty = Trim(vRecordset.Fields("pickqty").Value)
        vListItem.SubItems(4) = Format(vQTY, "##,##0.00")
        vListItem.SubItems(5) = Format(vPickQty, "##,##0.00")
        vListItem.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
        vListItem.SubItems(7) = Trim(vRecordset.Fields("quepicker").Value)
        vListItem.SubItems(8) = Trim(vRecordset.Fields("quezone").Value)
        vListItem.SubItems(9) = Trim(vRecordset.Fields("quedate").Value)
        vRecordset.MoveNext
      Next i
   End If
   vRecordset.Close
   
   
   
   If vCheckQueSend = 1 Then
      Me.PICLastSendQue.Visible = True
      Me.ListViewLastSendQue.SetFocus
      Exit Sub
   End If
   
   Dim vDay As String
   Dim vMonth As String
   
   If Len(Day(Now)) = 1 Then
      vDay = Trim("0" & Day(Now))
   Else
      vDay = Day(Now)
   End If
   
   If Len(Month(Now)) = 1 Then
      vMonth = Trim("0" & Month(Now))
   Else
      vMonth = Month(Now)
   End If
   
   vCheckDate = vDay & "/" & vMonth & "/" & Year(Now)
   
   If vDocdate < vCheckDate Then
      MsgBox "ไม่สามารถส่งคิวจัดสินค้าได้ กรณีวันที่เอกสารน้อยกว่าวันที่ปัจจุบัน", vbCritical, "Send Error Message"
      Me.TBBarCode.SetFocus
      Exit Sub
   End If
   
   vQuery = "exec dbo.USP_NP_SearchGroupPicking 1,'" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       n = vRecordset.RecordCount
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
       vGroupZone(i) = Trim(vRecordset.Fields("zoneid").Value)
       vRecordset.MoveNext
       Next i
   End If
   vRecordset.Close
   
   For i = 1 To n
      If vGroupZone(i) = "A" Then
         Call PrintPicking_A(vDocNo, vCheckTimeID, 1)
      ElseIf vGroupZone(i) = "B" Then
         Call PrintPicking_B(vDocNo, vCheckTimeID, 1)
      ElseIf vGroupZone(i) = "C" Then
         Call PrintPicking_C(vDocNo, vCheckTimeID, 1)
      ElseIf vGroupZone(i) = "X" Then
         Call PrintPicking_X(vDocNo, vCheckTimeID, 1)
      End If
   Next i
   
   vPRIsOpen = 0
   Me.ListViewItem.ListItems.Clear
   
   Dim vDay1 As String
   Dim vMonth1 As String
   
   If Len(Day(Now)) = 1 Then
      vDay1 = Trim("0" & Day(Now))
   Else
      vDay1 = Day(Now)
   End If
   
   If Len(Month(Now)) = 1 Then
      vMonth1 = Trim("0" & Month(Now))
   Else
      vMonth1 = Month(Now)
   End If
   
   Me.DTPDocDate1.Caption = vDay1 & "/" & vMonth1 & "/" & Year(Now)
   Me.TXTArCode.Text = ""
   Me.TXTSaleCode.Text = ""
   Me.TXTDocNo.Text = ""
   Me.LBLArName.Caption = ""
   Me.LBLTotalAmount.Caption = ""
   Me.CMDQue.Enabled = False
   
   Me.PICSendInformation.Visible = True
   vQuery = "exec dbo.usp_np_SearchReqPickingInformation '" & vDocNo & "'," & vCheckTimeID & " "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.ListViewInfQue.ListItems.Clear
      Me.LBLInfDocNo.Caption = Trim(vRecordset.Fields("docno").Value)
      Me.LBLInfARName.Caption = Trim(vRecordset.Fields("arname").Value)
      vRecordset.MoveFirst
      For i = 1 To vRecordset.RecordCount
        Set vListItem = Me.ListViewInfQue.ListItems.Add(, , i)
        vListItem.SubItems(1) = Trim(vRecordset.Fields("queid").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("zoneid").Value)
        vRecordset.MoveNext
      Next i
   End If
   vRecordset.Close
   Me.ListViewInfQue.SetFocus

vQuery = "exec dbo.USP_NP_UpdateSendQueuePicking 1,'" & vDocNo & "' "
gConnection.Execute vQuery
   
Call CMDDocNo_Click
Else
   MsgBox "เอกสารที่จะส่งจัดสินค้าได้ ต้องมีเลขที่เอกสาร รายการสินค้า และต้องเป็นเอกสารที่บันทึกข้อมูลเรียบร้อยแล้วเป็นอย่างน้อย กรุณาตรวจสอบ ", vbCritical, "Send Error Message"
   Me.TBBarCode.SetFocus
End If
End Sub

Public Sub PrintPicking_A(vDocNo As String, vTimeID As Integer, vType As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Dim i As Integer
Dim n As Integer
Dim vQueID As Integer
Dim vQueDocDate As String
Dim vARCode As String
Dim vSaleCode As String
Dim vRefNo As String
Dim vMemberID As String
Dim vSourceID As Integer
Dim vQueZone As String
Dim vQueReqTime As String
Dim vAddTime As Date
Dim vRequestTime As String
Dim vIsConditionSend As Integer
Dim vDocdate As String

Dim vCheckTimeID As Integer

Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Double
Dim vUnitCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vShelfID  As String
Dim vZoneID  As String
Dim vBarCode  As String
Dim vLineNumber As Integer

Dim vPickZone As String


'On Error GoTo ErrRollBack


If vDocNo <> "" Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  31"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    If vType = 1 Then
       vQueZone = "A"
       vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       vDocNo = Me.TXTDocNo.Text
       vDocdate = Me.DTPDocDate1.Caption
       vARCode = Me.TXTArCode.Text
   
       If Me.TXTSaleCode.Text <> "" Then
       vSaleCode = Left(Me.TXTSaleCode.Text, InStr(Me.TXTSaleCode.Text, "/") - 1)
       End If
       
       vRefNo = Me.TXTLicense.Text
       vMemberID = Me.TXTMember.Caption
       vSourceID = vType
       vIsConditionSend = 0
       vAddTime = DateAdd("n", 16, Now)
       If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
          vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
       ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
          vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
       ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
          vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
       ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
          vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
       End If
       vQueReqTime = vRequestTime
       
       vQuery = "exec dbo.USP_NP_InsertQuePickCenterMaster " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
       gConnection.Execute vQuery
   
       vQuery = "exec dbo.usp_np_SearchReqPickingItemZone " & vType & ",'" & vDocNo & "','A'," & vTimeID & " "
       If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          n = 0
          vRecordset.MoveFirst
         While Not vRecordset.EOF
             vItemCode = Trim(vRecordset.Fields("itemcode").Value)
             vItemName = Trim(vRecordset.Fields("itemname").Value)
             vQTY = Trim(vRecordset.Fields("qty").Value)
             vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
             vWHCode = Trim(vRecordset.Fields("whcode").Value)
             vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
             vShelfID = Trim(vRecordset.Fields("shelfid").Value)
             vZoneID = Trim(vRecordset.Fields("zoneid").Value)
             vBarCode = Trim(vRecordset.Fields("itemcode").Value)
             vLineNumber = n
             n = n + 1
   
             vQuery = "exec dbo.USP_NP_InsertQuePickCenterSub " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
             gConnection.Execute vQuery
             vRecordset.MoveNext
             Wend
         End If
         vRecordset.Close
End If
     
    'ElseIf vType = 2 Then
       'vQueZone = "A"
       'vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       'vDocNo = Me.TBDocNo.Text
       'vDocdate = Me.LBLOrderDocDate.Caption
       'vARCode = Me.LBLOrderArCode.Caption
   
       'If Me.LBLOrderSaleCode.Caption <> "" Then
       'vSaleCode = Left(Me.LBLOrderSaleCode.Caption, InStr(Me.LBLOrderSaleCode.Caption, "/") - 1)
       'End If
       
       'vRefNo = ""
       'vMemberID = ""
       'vSourceID = vType
       'vIsConditionSend = 0
       'vAddTime = DateAdd("n", 16, Now)
       'If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
        '  vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
       'ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
        '  vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
       'ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
        '  vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
       'ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
        '  vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
       'End If
       'vQueReqTime = vRequestTime
       
       'vQuery = "exec dbo.USP_NP_InsertQuePickCenterMaster " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
       'gConnection.Execute vQuery
   
       'vQuery = "exec dbo.usp_np_SearchReqPickingItemZone " & vType & ",'" & vDocNo & "','A'," & vTimeID & " "
       'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        '  n = 0
         ' vRecordset.MoveFirst
         'While Not vRecordset.EOF
          '   vItemCode = Trim(vRecordset.Fields("itemcode").Value)
           '  vItemName = Trim(vRecordset.Fields("itemname").Value)
            ' vQTY = Trim(vRecordset.Fields("qty").Value)
             'vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
             'vWHCode = Trim(vRecordset.Fields("whcode").Value)
             'vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
             'vShelfID = Trim(vRecordset.Fields("shelfid").Value)
             'vZoneID = Trim(vRecordset.Fields("zoneid").Value)
             'vBarCode = Trim(vRecordset.Fields("itemcode").Value)
             'vLineNumber = n
             'n = n + 1
   
             'vQuery = "exec dbo.USP_NP_InsertQuePickCenterSub " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
             'gConnection.Execute vQuery
             'vRecordset.MoveNext
             'Wend
         'End If
         'vRecordset.Close
       
    'End If
            
   If vType = 3 Then
      vQueZone = "A"
      vPickZone = "01"
      vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
      vDocNo = vDocNo
      vDocdate = Me.LBLDIDocDate.Caption
      vARCode = Me.TBDIArCode.Text
      
      If Me.TBDISaleCode.Text <> "" Then
      vSaleCode = Left(Me.TBDISaleCode.Text, InStr(Me.TBDISaleCode.Text, "/") - 1)
      End If
      
      vRefNo = Me.TBDICarLicense.Text
      vMemberID = Me.LBLDIMember.Caption
      vSourceID = vType
      vIsConditionSend = 0
      
      vAddTime = DateAdd("n", 6, Now)
      If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
      vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
      ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
      vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
      ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
      vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
      ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
      vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
      End If
      vQueReqTime = vRequestTime
      
    vQuery = "exec dbo.USP_NP_InsertQuePickCenterMasterDriveIn1 " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
    gConnection.Execute vQuery

    vQuery = "exec dbo.usp_np_SearchReqPickingItemZone1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'," & vTimeID & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       n = 0
       vRecordset.MoveFirst
      While Not vRecordset.EOF
          vItemCode = Trim(vRecordset.Fields("itemcode").Value)
          vItemName = Trim(vRecordset.Fields("itemname").Value)
          vQTY = Trim(vRecordset.Fields("qty").Value)
          vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
          vWHCode = Trim(vRecordset.Fields("whcode").Value)
          vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
          vShelfID = Trim(vRecordset.Fields("shelfid").Value)
          vZoneID = Trim(vRecordset.Fields("zoneid").Value)
          vBarCode = Trim(vRecordset.Fields("barcode").Value)
          vLineNumber = n
          n = n + 1

          vQuery = "exec dbo.USP_NP_InsertQuePickCenterDriveInSub1 " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vQueZone & "','" & vPickZone & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
          gConnection.Execute vQuery
          vRecordset.MoveNext
          Wend
      End If
      vRecordset.Close
      
   End If
       
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  31 "
    gConnection.Execute vQuery
    
    'vQuery = "commit tran"
    'gConnection.Execute vQuery
    
   If vType <> 3 Then
  Call PrintPickingHeader(vQueID, vQueDocDate, vZoneID)
  Call PrintPickingSlip(vQueID, vQueDocDate, vZoneID)
  End If
  
  
  MsgBox "ได้คิวเลขที่ " & vQueID & " ", vbInformation, "Send Queue Information"
End If

'ErrRollBack:
'If Err.Description <> "" Then
 ' vQuery = "rollback tran"
  'gConnection.Execute vQuery
  'MsgBox Err.Description
  'MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าชั้นเก็บ AVL ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
'End If

'ErrRunQueueID:
'If Err.Number = -2147217873 And vPosition = 2 Then
 '   vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
  '  gConnection.Execute vQuery
   ' MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    'Exit Sub
'End If

End Sub

Public Sub PrintPicking_B(vDocNo As String, vTimeID As Integer, vType As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Dim i As Integer
Dim n As Integer
Dim vQueID As Integer
Dim vQueDocDate As String
Dim vDocdate As String
Dim vARCode As String
Dim vSaleCode As String
Dim vRefNo As String
Dim vMemberID As String
Dim vSourceID As Integer
Dim vQueZone As String
Dim vQueReqTime As String
Dim vAddTime As Date
Dim vRequestTime As String
Dim vIsConditionSend As Integer

Dim vCheckTimeID As Integer

Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Double
Dim vUnitCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vShelfID  As String
Dim vZoneID  As String
Dim vBarCode  As String
Dim vLineNumber As Integer

Dim vPickZone As String


'On Error GoTo ErrRollBack


If vDocNo <> "" Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  31"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    If vType = 1 Then
       vQueZone = "B"
       vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       vDocNo = Me.TXTDocNo.Text
       vDocdate = Me.DTPDocDate1.Caption
       vARCode = Me.TXTArCode.Text
   
       If Me.TXTSaleCode.Text <> "" Then
       vSaleCode = Left(Me.TXTSaleCode.Text, InStr(Me.TXTSaleCode.Text, "/") - 1)
       End If
       
       vRefNo = Me.TXTLicense.Text
       vMemberID = Me.TXTMember.Caption
       vSourceID = vType
       vIsConditionSend = 0
       vAddTime = DateAdd("n", 16, Now)
       If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
          vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
       ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
          vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
       ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
          vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
       ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
          vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
       End If
       vQueReqTime = vRequestTime
       
       vQuery = "exec dbo.USP_NP_InsertQuePickCenterMaster " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
       gConnection.Execute vQuery
   
       vQuery = "exec dbo.usp_np_SearchReqPickingItemZone " & vType & ",'" & vDocNo & "','B'," & vTimeID & " "
       If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          n = 0
          vRecordset.MoveFirst
         While Not vRecordset.EOF
             vItemCode = Trim(vRecordset.Fields("itemcode").Value)
             vItemName = Trim(vRecordset.Fields("itemname").Value)
             vQTY = Trim(vRecordset.Fields("qty").Value)
             vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
             vWHCode = Trim(vRecordset.Fields("whcode").Value)
             vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
             vShelfID = Trim(vRecordset.Fields("shelfid").Value)
             vZoneID = Trim(vRecordset.Fields("zoneid").Value)
             vBarCode = Trim(vRecordset.Fields("itemcode").Value)
             vLineNumber = n
             n = n + 1
   
             vQuery = "exec dbo.USP_NP_InsertQuePickCenterSub " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
             gConnection.Execute vQuery
             vRecordset.MoveNext
             Wend
         End If
         vRecordset.Close
End If
       
    'If vType = 2 Then
       'vQueZone = "B"
       'vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       'vDocNo = Me.TBDocNo.Text
       'vDocdate = Me.LBLOrderDocDate.Caption
       'vARCode = Me.LBLOrderArCode.Caption
   
       'If Me.LBLOrderSaleCode.Caption <> "" Then
       'vSaleCode = Left(Me.LBLOrderSaleCode.Caption, InStr(Me.LBLOrderSaleCode.Caption, "/") - 1)
       'End If
       
       'vRefNo = ""
       'vMemberID = ""
       'vSourceID = vType
       'vIsConditionSend = 0
       'vAddTime = DateAdd("n", 16, Now)
       'If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
        '  vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
       'ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
        '  vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
       'ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
        '  vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
       'ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
        '  vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
       'End If
       'vQueReqTime = vRequestTime
       
       'vQuery = "exec dbo.USP_NP_InsertQuePickCenterMaster " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
       'gConnection.Execute vQuery
   
       'vQuery = "exec dbo.usp_np_SearchReqPickingItemZone " & vType & ",'" & vDocNo & "','B'," & vTimeID & " "
       'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        '  n = 0
         ' vRecordset.MoveFirst
         'While Not vRecordset.EOF
          '   vItemCode = Trim(vRecordset.Fields("itemcode").Value)
           '  vItemName = Trim(vRecordset.Fields("itemname").Value)
            ' vQTY = Trim(vRecordset.Fields("qty").Value)
             'vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
             'vWHCode = Trim(vRecordset.Fields("whcode").Value)
             'vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
             'vShelfID = Trim(vRecordset.Fields("shelfid").Value)
             'vZoneID = Trim(vRecordset.Fields("zoneid").Value)
             'vBarCode = Trim(vRecordset.Fields("itemcode").Value)
             'vLineNumber = n
             'n = n + 1
   
             'vQuery = "exec dbo.USP_NP_InsertQuePickCenterSub " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
             'gConnection.Execute vQuery
             'vRecordset.MoveNext
             'Wend
         'End If
         'vRecordset.Close
       
    'End If
            
   If vType = 3 Then
      vQueZone = "B"
      vPickZone = "02"
      vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
      vDocNo = vDocNo
      vDocdate = Me.LBLDIDocDate.Caption
      vARCode = Me.TBDIArCode.Text
      
      If Me.TBDISaleCode.Text <> "" Then
      vSaleCode = Left(Me.TBDISaleCode.Text, InStr(Me.TBDISaleCode.Text, "/") - 1)
      End If
      
      vRefNo = Me.TBDICarLicense.Text
      vMemberID = Me.LBLDIMember.Caption
      vSourceID = vType
      vIsConditionSend = 0
      
      vAddTime = DateAdd("n", 6, Now)
      If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
      vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
      ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
      vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
      ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
      vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
      ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
      vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
      End If
      vQueReqTime = vRequestTime
      
    vQuery = "exec dbo.USP_NP_InsertQuePickCenterMasterDriveIn1 " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
    gConnection.Execute vQuery

    vQuery = "exec dbo.usp_np_SearchReqPickingItemZone1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'," & vTimeID & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       n = 0
       vRecordset.MoveFirst
      While Not vRecordset.EOF
          vItemCode = Trim(vRecordset.Fields("itemcode").Value)
          vItemName = Trim(vRecordset.Fields("itemname").Value)
          vQTY = Trim(vRecordset.Fields("qty").Value)
          vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
          vWHCode = Trim(vRecordset.Fields("whcode").Value)
          vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
          vShelfID = Trim(vRecordset.Fields("shelfid").Value)
          vZoneID = Trim(vRecordset.Fields("zoneid").Value)
          vBarCode = Trim(vRecordset.Fields("barcode").Value)
          vLineNumber = n
          n = n + 1

          vQuery = "exec dbo.USP_NP_InsertQuePickCenterDriveInSub1 " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vQueZone & "','" & vPickZone & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
          gConnection.Execute vQuery
          vRecordset.MoveNext
          Wend
      End If
      vRecordset.Close
      
   End If
       
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  31 "
    gConnection.Execute vQuery
    
    'vQuery = "commit tran"
    'gConnection.Execute vQuery
    
   If vType <> 3 Then
  Call PrintPickingHeader(vQueID, vQueDocDate, vZoneID)
  Call PrintPickingSlip(vQueID, vQueDocDate, vZoneID)
  End If
  
  
  MsgBox "ได้คิวเลขที่ " & vQueID & " ", vbInformation, "Send Queue Information"
End If

'ErrRollBack:
'If Err.Description <> "" Then
 ' vQuery = "rollback tran"
  'gConnection.Execute vQuery
  'MsgBox Err.Description
  'MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าชั้นเก็บ AVL ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
'End If

'ErrRunQueueID:
'If Err.Number = -2147217873 And vPosition = 2 Then
 '   vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
  '  gConnection.Execute vQuery
   ' MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    'Exit Sub
'End If

End Sub


Public Sub PrintPicking_C(vDocNo As String, vTimeID As Integer, vType As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Dim i As Integer
Dim n As Integer
Dim vQueID As Integer
Dim vQueDocDate As String
Dim vDocdate As String
Dim vARCode As String
Dim vSaleCode As String
Dim vRefNo As String
Dim vMemberID As String
Dim vSourceID As Integer
Dim vQueZone As String
Dim vQueReqTime As String
Dim vAddTime As Date
Dim vRequestTime As String
Dim vIsConditionSend As Integer

Dim vCheckTimeID As Integer

Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Double
Dim vUnitCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vShelfID  As String
Dim vZoneID  As String
Dim vBarCode  As String
Dim vLineNumber As Integer

Dim vPickZone As String


'On Error GoTo ErrRollBack


If vDocNo <> "" Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  31"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    If vType = 1 Then
       vQueZone = "C"
       vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       vDocNo = Me.TXTDocNo.Text
       vDocdate = Me.DTPDocDate1.Caption
       vARCode = Me.TXTArCode.Text
   
       If Me.TXTSaleCode.Text <> "" Then
       vSaleCode = Left(Me.TXTSaleCode.Text, InStr(Me.TXTSaleCode.Text, "/") - 1)
       End If
       
       vRefNo = Me.TXTLicense.Text
       vMemberID = Me.TXTMember.Caption
       vSourceID = vType
       vIsConditionSend = 0
       vAddTime = DateAdd("n", 16, Now)
       If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
          vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
       ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
          vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
       ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
          vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
       ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
          vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
       End If
       vQueReqTime = vRequestTime
       
       vQuery = "exec dbo.USP_NP_InsertQuePickCenterMaster " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
       gConnection.Execute vQuery
   
       vQuery = "exec dbo.usp_np_SearchReqPickingItemZone " & vType & ",'" & vDocNo & "','C'," & vTimeID & " "
       If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          n = 0
          vRecordset.MoveFirst
         While Not vRecordset.EOF
             vItemCode = Trim(vRecordset.Fields("itemcode").Value)
             vItemName = Trim(vRecordset.Fields("itemname").Value)
             vQTY = Trim(vRecordset.Fields("qty").Value)
             vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
             vWHCode = Trim(vRecordset.Fields("whcode").Value)
             vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
             vShelfID = Trim(vRecordset.Fields("shelfid").Value)
             vZoneID = Trim(vRecordset.Fields("zoneid").Value)
             vBarCode = Trim(vRecordset.Fields("itemcode").Value)
             vLineNumber = n
             n = n + 1
   
             vQuery = "exec dbo.USP_NP_InsertQuePickCenterSub " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
             gConnection.Execute vQuery
             vRecordset.MoveNext
             Wend
         End If
         vRecordset.Close
    End If
       
    'If vType = 2 Then
     '  vQueZone = "C"
      ' vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       'vDocNo = Me.TBDocNo.Text
       'vDocdate = Me.LBLOrderDocDate.Caption
       'vARCode = Me.LBLOrderArCode.Caption
   
       'If Me.LBLOrderSaleCode.Caption <> "" Then
       'vSaleCode = Left(Me.LBLOrderSaleCode.Caption, InStr(Me.LBLOrderSaleCode.Caption, "/") - 1)
       'End If
       
       'vRefNo = ""
       'vMemberID = ""
       'vSourceID = vType
       'vIsConditionSend = 0
       'vAddTime = DateAdd("n", 16, Now)
       'If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
        '  vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
       'ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
        '  vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
       'ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
        '  vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
       'ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
        '  vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
       'End If
       'vQueReqTime = vRequestTime
       
       'vQuery = "exec dbo.USP_NP_InsertQuePickCenterMaster " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
       'gConnection.Execute vQuery
   
       'vQuery = "exec dbo.usp_np_SearchReqPickingItemZone " & vType & ",'" & vDocNo & "','C'," & vTimeID & " "
       'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        '  n = 0
         ' vRecordset.MoveFirst
         'While Not vRecordset.EOF
          '   vItemCode = Trim(vRecordset.Fields("itemcode").Value)
           '''  vItemName = Trim(vRecordset.Fields("itemname").Value)
            'vQTY = Trim(vRecordset.Fields("qty").Value)
             'vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
             'vWHCode = Trim(vRecordset.Fields("whcode").Value)
             'vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
             'vShelfID = Trim(vRecordset.Fields("shelfid").Value)
             'vZoneID = Trim(vRecordset.Fields("zoneid").Value)
             ''vBarCode = Trim(vRecordset.Fields("itemcode").Value)
             'vLineNumber = n
             'n = n + 1
   
             'vQuery = "exec dbo.USP_NP_InsertQuePickCenterSub " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
             'gConnection.Execute vQuery
             'vRecordset.MoveNext
             'Wend
         'End If
         'vRecordset.Close
       
    'End If
            
     If vType = 3 Then
      vQueZone = "C"
      vPickZone = "03"
      vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
      vDocNo = vDocNo
      vDocdate = Me.LBLDIDocDate.Caption
      vARCode = Me.TBDIArCode.Text
      
      If Me.TBDISaleCode.Text <> "" Then
      vSaleCode = Left(Me.TBDISaleCode.Text, InStr(Me.TBDISaleCode.Text, "/") - 1)
      End If
      
      vRefNo = Me.TBDICarLicense.Text
      vMemberID = Me.LBLDIMember.Caption
      vSourceID = vType
      vIsConditionSend = 0
      
      vAddTime = DateAdd("n", 6, Now)
      If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
      vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
      ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
      vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
      ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
      vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
      ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
      vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
      End If
      vQueReqTime = vRequestTime
      
    vQuery = "exec dbo.USP_NP_InsertQuePickCenterMasterDriveIn1 " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
    gConnection.Execute vQuery

    vQuery = "exec dbo.usp_np_SearchReqPickingItemZone1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'," & vTimeID & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       n = 0
       vRecordset.MoveFirst
      While Not vRecordset.EOF
          vItemCode = Trim(vRecordset.Fields("itemcode").Value)
          vItemName = Trim(vRecordset.Fields("itemname").Value)
          vQTY = Trim(vRecordset.Fields("qty").Value)
          vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
          vWHCode = Trim(vRecordset.Fields("whcode").Value)
          vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
          vShelfID = Trim(vRecordset.Fields("shelfid").Value)
          vZoneID = Trim(vRecordset.Fields("zoneid").Value)
          vBarCode = Trim(vRecordset.Fields("barcode").Value)
          vLineNumber = n
          n = n + 1

          vQuery = "exec dbo.USP_NP_InsertQuePickCenterDriveInSub1 " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vQueZone & "','" & vPickZone & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
          gConnection.Execute vQuery
          vRecordset.MoveNext
          Wend
      End If
      vRecordset.Close
      
   End If
       
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  31 "
    gConnection.Execute vQuery
    
    'vQuery = "commit tran"
    'gConnection.Execute vQuery
    
   If vType <> 3 Then
  Call PrintPickingHeader(vQueID, vQueDocDate, vZoneID)
  Call PrintPickingSlip(vQueID, vQueDocDate, vZoneID)
  End If
  
  
  MsgBox "ได้คิวเลขที่ " & vQueID & " ", vbInformation, "Send Queue Information"
End If


'ErrRollBack:
'If Err.Description <> "" Then
 ' vQuery = "rollback tran"
  'gConnection.Execute vQuery
  'MsgBox Err.Description
  'MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าชั้นเก็บ AVL ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
'End If

'ErrRunQueueID:
'If Err.Number = -2147217873 And vPosition = 2 Then
 '   vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
  '  gConnection.Execute vQuery
   ' MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    'Exit Sub
'End If

End Sub

Public Sub PrintPicking_X(vDocNo As String, vTimeID As Integer, vType As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Dim i As Integer
Dim n As Integer
Dim vQueID As Integer
Dim vQueDocDate As String
Dim vDocdate As String
Dim vARCode As String
Dim vSaleCode As String
Dim vRefNo As String
Dim vMemberID As String
Dim vSourceID As Integer
Dim vQueZone As String
Dim vQueReqTime As String
Dim vAddTime As Date
Dim vRequestTime As String
Dim vIsConditionSend As Integer

Dim vCheckTimeID As Integer

Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Double
Dim vUnitCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vShelfID  As String
Dim vZoneID  As String
Dim vBarCode  As String
Dim vLineNumber As Integer

Dim vPickZone As String


'On Error GoTo ErrRollBack


If vDocNo <> "" Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  31"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    If vType = 1 Then
       vQueZone = "X"
       vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       vDocNo = Me.TXTDocNo.Text
       vDocdate = Me.DTPDocDate1.Caption
       vARCode = Me.TXTArCode.Text
   
       If Me.TXTSaleCode.Text <> "" Then
       vSaleCode = Left(Me.TXTSaleCode.Text, InStr(Me.TXTSaleCode.Text, "/") - 1)
       End If
       
       vRefNo = Me.TXTLicense.Text
       vMemberID = Me.TXTMember.Caption
       vSourceID = vType
       vIsConditionSend = 0
       vAddTime = DateAdd("n", 16, Now)
       If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
          vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
       ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
          vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
       ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
          vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
       ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
          vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
       End If
       vQueReqTime = vRequestTime
       
       vQuery = "exec dbo.USP_NP_InsertQuePickCenterMaster " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
       gConnection.Execute vQuery
   
       vQuery = "exec dbo.usp_np_SearchReqPickingItemZone " & vType & ",'" & vDocNo & "','X'," & vTimeID & " "
       If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          n = 0
          vRecordset.MoveFirst
         While Not vRecordset.EOF
             vItemCode = Trim(vRecordset.Fields("itemcode").Value)
             vItemName = Trim(vRecordset.Fields("itemname").Value)
             vQTY = Trim(vRecordset.Fields("qty").Value)
             vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
             vWHCode = Trim(vRecordset.Fields("whcode").Value)
             vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
             vShelfID = Trim(vRecordset.Fields("shelfid").Value)
             vZoneID = Trim(vRecordset.Fields("zoneid").Value)
             vBarCode = Trim(vRecordset.Fields("itemcode").Value)
             vLineNumber = n
             n = n + 1
   
             vQuery = "exec dbo.USP_NP_InsertQuePickCenterSub " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
             gConnection.Execute vQuery
             vRecordset.MoveNext
             Wend
         End If
         vRecordset.Close
    End If
       
    'ElseIf vType = 2 Then
     '  vQueZone = "X"
      ' vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       'vDocNo = Me.TBDocNo.Text
       'vDocdate = Me.LBLOrderDocDate.Caption
       'vARCode = Me.LBLOrderArCode.Caption
   '
    '   If Me.LBLOrderSaleCode.Caption <> "" Then
     '  vSaleCode = Left(Me.LBLOrderSaleCode.Caption, InStr(Me.LBLOrderSaleCode.Caption, "/") - 1)
      ' End If
       '
       'vRefNo = ""
       'vMemberID = ""
       'vSourceID = vType
       'vIsConditionSend = 0
       'vAddTime = DateAdd("n", 16, Now)
       'If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
        '  vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
       'ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
        '  vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
       'ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
        '  vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
       'ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
        '  vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
       'End If
       'vQueReqTime = vRequestTime
       
       'vQuery = "exec dbo.USP_NP_InsertQuePickCenterMaster " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
       'gConnection.Execute vQuery
   
       'vQuery = "exec dbo.usp_np_SearchReqPickingItemZone " & vType & ",'" & vDocNo & "','X'," & vTimeID & " "
       'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        '  n = 0
         ' vRecordset.MoveFirst
         'While Not vRecordset.EOF
          '   vItemCode = Trim(vRecordset.Fields("itemcode").Value)
           '  vItemName = Trim(vRecordset.Fields("itemname").Value)
            ' vQTY = Trim(vRecordset.Fields("qty").Value)
             'vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
             'vWHCode = Trim(vRecordset.Fields("whcode").Value)
             ''vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
             'vShelfID = Trim(vRecordset.Fields("shelfid").Value)
             'vZoneID = Trim(vRecordset.Fields("zoneid").Value)
             'vBarCode = Trim(vRecordset.Fields("itemcode").Value)
             'vLineNumber = n
             'n = n + 1
   
             'vQuery = "exec dbo.USP_NP_InsertQuePickCenterSub " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
             'gConnection.Execute vQuery
             'vRecordset.MoveNext
             'Wend
         'End If
         'vRecordset.Close
       
    'End If
            
    If vType = 3 Then
      vQueZone = "X"
      vPickZone = "04"
      vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
      vDocNo = vDocNo
      vDocdate = Me.LBLDIDocDate.Caption
      vARCode = Me.TBDIArCode.Text
      
      If Me.TBDISaleCode.Text <> "" Then
      vSaleCode = Left(Me.TBDISaleCode.Text, InStr(Me.TBDISaleCode.Text, "/") - 1)
      End If
      
      vRefNo = Me.TBDICarLicense.Text
      vMemberID = Me.LBLDIMember.Caption
      vSourceID = vType
      vIsConditionSend = 0
      
      vAddTime = DateAdd("n", 6, Now)
      If Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) = 1 Then
      vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Trim("0" & Minute(vAddTime))
      ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) = 1 Then
      vRequestTime = Hour(vAddTime) & ":" & Trim("0" & Minute(vAddTime))
      ElseIf Len(Hour(vAddTime)) = 1 And Len(Minute(vAddTime)) > 1 Then
      vRequestTime = Trim("0" & Hour(vAddTime)) & ":" & Minute(vAddTime)
      ElseIf Len(Hour(vAddTime)) > 1 And Len(Minute(vAddTime)) > 1 Then
      vRequestTime = Hour(vAddTime) & ":" & Minute(vAddTime)
      End If
      vQueReqTime = vRequestTime
      
    vQuery = "exec dbo.USP_NP_InsertQuePickCenterMasterDriveIn1 " & vQueID & ",'" & vQueDocDate & "','" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vSourceID & ",'" & vQueZone & "'," & vIsConditionSend & ",'" & vQueReqTime & "','" & vTimeID & "' "
    gConnection.Execute vQuery

    vQuery = "exec dbo.usp_np_SearchReqPickingItemZone1 " & vType & ",'" & vDocNo & "','" & vPickZone & "'," & vTimeID & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       n = 0
       vRecordset.MoveFirst
      While Not vRecordset.EOF
          vItemCode = Trim(vRecordset.Fields("itemcode").Value)
          vItemName = Trim(vRecordset.Fields("itemname").Value)
          vQTY = Trim(vRecordset.Fields("qty").Value)
          vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
          vWHCode = Trim(vRecordset.Fields("whcode").Value)
          vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
          vShelfID = Trim(vRecordset.Fields("shelfid").Value)
          vZoneID = Trim(vRecordset.Fields("zoneid").Value)
          vBarCode = Trim(vRecordset.Fields("barcode").Value)
          vLineNumber = n
          n = n + 1

          vQuery = "exec dbo.USP_NP_InsertQuePickCenterDriveInSub1 " & vQueID & ",'" & vQueDocDate & "','" & vItemCode & "','" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vQueZone & "','" & vPickZone & "'," & vQTY & ",0,0,'" & vUnitCode & "','" & vBarCode & "','" & vDocNo & "'," & vTimeID & "," & vLineNumber & " "
          gConnection.Execute vQuery
          vRecordset.MoveNext
          Wend
      End If
      vRecordset.Close
      
   End If
       
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  31 "
    gConnection.Execute vQuery
    
    'vQuery = "commit tran"
    'gConnection.Execute vQuery
    
   If vType <> 3 Then
  Call PrintPickingHeader(vQueID, vQueDocDate, vZoneID)
  Call PrintPickingSlip(vQueID, vQueDocDate, vZoneID)
  End If
  
  
  MsgBox "ได้คิวเลขที่ " & vQueID & " ", vbInformation, "Send Queue Information"
End If


'ErrRollBack:
'If Err.Description <> "" Then
 ' vQuery = "rollback tran"
  'gConnection.Execute vQuery
  'MsgBox Err.Description
  'MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าชั้นเก็บ AVL ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
'End If

'ErrRunQueueID:
'If Err.Number = -2147217873 And vPosition = 2 Then
 '   vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
  '  gConnection.Execute vQuery
   ' MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    'Exit Sub
'End If

End Sub

Private Sub CMDQue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If
End Sub

Private Sub CMDSale_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchSale As String
Dim vListItem As ListItem
Dim i As Integer

Me.PICPRSearchSale.Visible = True

vQuery = "exec dbo.USP_CRM_EmployeeDetails 0,'' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.ListViewPRSearchSale.ListItems.Clear
    vRecordset.MoveFirst
    For i = 1 To vRecordset.RecordCount
       Set vListItem = Me.ListViewPRSearchSale.ListItems.Add(, , i)
       vListItem.SubItems(1) = Trim(vRecordset.Fields("empcode").Value)
       vListItem.SubItems(2) = Trim(vRecordset.Fields("empname").Value)
       vRecordset.MoveNext
    Next i
End If
vRecordset.Close
Me.TBPRSearchSale.SetFocus
End Sub

Private Sub CMDSaleClose_Click()
'Me.PICSaleCode.Visible = False
End Sub

Private Sub CMDSaleOK_Click()
'Dim vIndex As Integer
'Dim vSaleCode As String

'If Me.ListViewSaleCode.ListItems.Count > 0 Then
'vIndex = Me.ListViewSaleCode.SelectedItem.Index
'vSaleCode = Trim(Me.ListViewSaleCode.ListItems(vIndex).SubItems(1) & "/" & Me.ListViewSaleCode.ListItems(vIndex).SubItems(2))
'Me.TXTSaleCode.Text = vSaleCode
'Me.TXTSaleCode.SetFocus
'End If
'Me.PICSaleCode.Visible = False
End Sub

Private Sub CMDSaleOrder_Click()
'Me.PICSaleOrder.Visible = True
'Me.TBDocNo.SetFocus
End Sub

Private Sub CMDSaleOrderExit_Click()
'Me.TBDocNo.Text = ""
'Me.PICSaleOrder.Visible = False
End Sub

Private Sub CMDSale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If
End Sub

Private Sub CMDSaleOrderInformationClose_Click()
'Dim vAnswer As Integer

'Me.PICSaleOrderQueInformation.Visible = False
'Me.TBDocNo.SetFocus
End Sub

Private Sub CMDSaleOrderSendQue_Click()
'Dim vQuery As String
'Dim vRecordset As New Recordset
'Dim i As Integer

'Dim vDocNo As String
'Dim vARCode As String
'Dim vSaleCode As String
'Dim vBillStatus As String
'Dim vSoStatus As Integer
'Dim n As Integer
'Dim vBillType As Integer
'Dim vCarLicense As String
'Dim vDeliveryDate As String
'Dim vPickStatus As Integer
'Dim vSOCountNumber As Integer
'
'Dim vDocdate As String
'Dim vQueDocDate As String
'Dim vPickingDate As String
'Dim vItemCode As String
'Dim vItemName As String
'Dim vReqQTY As Double
'Dim vUnitCode As String
'Dim vWHCode As String
'Dim vShelfCode As String
'Dim vZoneID As String
'Dim vIsCancel As Integer
'Dim vLineNumber As Integer
''Dim j As Integer
'Dim vIsConditionSend As Integer
''Dim vCountNumber As Integer
'Dim vCheckShelfGroup As String
'Dim vDueDate As String
'Dim vSelectItemDateTime As String

'Dim vSumOfItemAmount As Double
'Dim vTaxAmount As Double
'Dim vTotalAmount As Double
'
'Dim vPrice As Double
'Dim vDiscountAmount As Double
'Dim vItemAmount As Double
'
'Dim vListItem As ListItem
'Dim vAnswer As Integer
'Dim vLastTimeID As Integer
'Dim vCheckTimeID As Integer
'Dim vQTY As Double
'Dim vPickQty As Double

'If Me.LBLOrderArCode.Caption <> "" And Me.ListViewSaleOrder.ListItems.Count > 0 Then

 '  vDocNo = Me.TBDocNo.Text
  ' vDocdate = Me.LBLOrderDocDate.Caption
   'vPickingDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   'vARCode = Me.LBLOrderArCode.Caption
   'vSaleCode = Left(Me.LBLOrderSaleCode.Caption, InStr(Me.LBLOrderSaleCode.Caption, "/") - 1)
   
   'vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
      
    'vQuery = "exec dbo.USP_NP_CheckQuePickCenter '" & vDocNo & "','" & vQueDocDate & "' "
    'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     '  vLastTimeID = Trim(vRecordset.Fields("max1").Value)
    'End If
    'vRecordset.Close
    
   'vCheckTimeID = vLastTimeID + 1
   
   'vQuery = "exec dbo.usp_np_SearchReqPickingInformationLastSend '" & vDocNo & "','" & vQueDocDate & "'," & vLastTimeID & " "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '  Me.ListViewSaleOrderLastQue.ListItems.Clear
     ' Me.ListViewSaleOrderQueInformation.ListItems.Clear
      'vRecordset.MoveFirst
      ''For i = 1 To vRecordset.RecordCount
        'Set vListItem = Me.ListViewSaleOrderLastQue.ListItems.Add(, , i)
        'vListItem.SubItems(1) = Trim(vRecordset.Fields("queid").Value)
        'vListItem.SubItems(2) = Trim(vRecordset.Fields("quedescription").Value)
        ''vListItem.SubItems(3) = Trim(vRecordset.Fields("itemcode").Value) & "/" & Trim(vRecordset.Fields("itemname").Value)
        'vQTY = Trim(vRecordset.Fields("qty").Value)
        'vPickQty = Trim(vRecordset.Fields("pickqty").Value)
        ''vListItem.SubItems(4) = Format(vQTY, "##,##0.00")
        'vListItem.SubItems(5) = Format(vPickQty, "##,##0.00")
        'vListItem.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
        'vListItem.SubItems(7) = Trim(vRecordset.Fields("quepicker").Value)
        'vListItem.SubItems(8) = Trim(vRecordset.Fields("quezone").Value)
        ''vListItem.SubItems(9) = Trim(vRecordset.Fields("quedate").Value)
        'vRecordset.MoveNext
      'Next i
   'End If
   'vRecordset.Close
   
   'If Me.ListViewSaleOrderLastQue.ListItems.Count > 0 Then
    '  Me.LBLSaleOrderQueInf.Caption = "รายการ คิวจัดสินค้าล่าสุดของเอกสารนี้"
     ' Me.PICSaleOrderQueInformation.Visible = True
      'Me.ListViewSaleOrderQueInformation.Visible = False
      ''Me.ListViewSaleOrderLastQue.Visible = True
      
      'vAnswer = MsgBox("คุณต้องการ ส่งคิวจัดสินค้าต่อหรือไม่", vbYesNo, "Send Question Message")
      'If vAnswer = 7 Then
       ' Exit Sub
      'End If
   'Else
    '  Me.LBLSaleOrderQueInf.Caption = "รายการ คิวที่สั่งจัดสินค้า"
     ' Me.PICSaleOrderQueInformation.Visible = False
      ''Me.ListViewSaleOrderQueInformation.Visible = False
      'Me.ListViewSaleOrderLastQue.Visible = False
   'End If
   
   
   'If vSaleCode = "" Then
    '  MsgBox "ไม่ได้ระบุ รหัสพนักงานกรุณาตรวจสอบ ", vbCritical, "Send Error Message"
     ' Exit Sub
   'End If
   
   'vIsConditionSend = 0
   'vCarLicense = ""
   'vBillType = 0 'Me.LBLOrderBillType.Caption
   'vSoStatus = 0 'Me.LBLOrderSoStatus.Caption
   ''vDeliveryDate = vPickingDate
   'vDueDate = vPickingDate
   'vPickStatus = 0
   
   'If Me.LBLOrderSumOfItemAmount.Caption <> "" Then
   'vSumOfItemAmount = Me.LBLOrderSumOfItemAmount.Caption
   'End If
   
   'If Me.LBLOrderTaxAmount.Caption <> "" Then
   'vTaxAmount = Me.LBLOrderTaxAmount.Caption
   'End If
   
   'If Me.LBLOrderNetAmount.Caption <> "" Then
   'vTotalAmount = Me.LBLOrderNetAmount.Caption
   'End If
   
   'vQuery = "exec dbo.USP_NP_SearchCheckCountSOPicking '" & vDocNo & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '  vSOCountNumber = vRecordset.Fields("vCount").Value
   'End If
   'vRecordset.Close
   
   'vQuery = "exec dbo.USP_NP_SearchSaleOrderGroupShelf '" & vDocNo & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '  vRecordset.MoveFirst
     ' While Not vRecordset.EOF
      'vCheckShelfGroup = vRecordset.Fields("shelfgroup").Value
      'vQuery = "exec dbo.USP_NP_InsertOrderPickHoldBill '" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vPickingDate & "'," & vBillType & "," & vSoStatus & ",0,'" & vSaleCode & "','" & vCarLicense & "'," & vIsConditionSend & "," & vSOCountNumber & ",'" & vCheckShelfGroup & "','" & vDueDate & "'," & vPickStatus & "," & vSumOfItemAmount & "," & vTaxAmount & "," & vTotalAmount & ",'" & vUserID & "' "
      ''gConnection.Execute vQuery
      'vRecordset.MoveNext
      'Wend
   'End If
   'vRecordset.Close
   
   'For j = 1 To Me.ListViewSaleOrder.ListItems.Count
    '  vItemCode = Me.ListViewSaleOrder.ListItems(j).SubItems(1)
     ' vItemName = Me.ListViewSaleOrder.ListItems(j).SubItems(2)
      'vReqQTY = Me.ListViewSaleOrder.ListItems(j).SubItems(3)
      ''vUnitCode = Me.ListViewSaleOrder.ListItems(j).SubItems(4)
      'vWHCode = Me.ListViewSaleOrder.ListItems(j).SubItems(8)
      'vShelfCode = Me.ListViewSaleOrder.ListItems(j).SubItems(9)
      'vZoneID = Me.ListViewSaleOrder.ListItems(j).SubItems(10)
      ''vPrice = Me.ListViewSaleOrder.ListItems(j).SubItems(5)
      'vDiscountAmount = Me.ListViewSaleOrder.ListItems(j).SubItems(6)
      'vItemAmount = Me.ListViewSaleOrder.ListItems(j).SubItems(7)
      ''vIsCancel = 0
      'vSelectItemDateTime = Now
      ''vLineNumber = j - 1
      ''vQuery = "exec dbo.USP_NP_InsertOrderPickHoldBillSub '" & vDocNo & "','" & vDocdate & "','" & vPickingDate & "','" & vItemCode & "','" & vItemName & "'," & vReqQTY & ",'" & vUnitCode & "','" & vWHCode & "','" & vShelfCode & "','" & vZoneID & "'," & vIsCancel & ",'" & vSelectItemDateTime & "'," & vSOCountNumber & "," & vPrice & "," & vDiscountAmount & "," & vItemAmount & "," & vLineNumber & " "
      'gConnection.Execute vQuery
   'Next j
   
   'Call SendQue(vSOCountNumber)
   
    'Me.PICSaleOrderQueInformation.Visible = True
    'Me.ListViewSaleOrderLastQue.Visible = False
    'Me.ListViewSaleOrderQueInformation.Visible = True
    'Me.ListViewSaleOrderQueInformation.ListItems.Clear
    'vQuery = "exec dbo.usp_np_SearchReqPickingInformation '" & vDocNo & "'," & vSOCountNumber & " "
    'if OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    'Me.ListViewSaleOrderQueInformation.ListItems.Clear
    'vRecordset.MoveFirst
    'For i = 1 To vRecordset.RecordCount
    'Set vListItem = Me.ListViewSaleOrderQueInformation.ListItems.Add(, , i)
    'vListItem.SubItems(1) = Trim(vRecordset.Fields("queid").Value)
    'vListItem.SubItems(2) = Trim(vRecordset.Fields("zoneid").Value)
    'vRecordset.MoveNext
    'Next i
    'End If
    'vRecordset.Close
    'Me.ListViewSaleOrderQueInformation.SetFocus
   
   'Me.TBDocNo.Text = ""
   'Me.TBDocNo.SetFocus
'End If
End Sub

Public Sub SendQue(vTimeID As Integer)
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim i As Integer
''Dim n As Integer
'Dim vDocNo As String
'Dim vDocdate As String
'Dim vQueDocDate As String
'Dim vGroupZone(5) As String

'Dim vListItem As ListItem
'Dim vCheckDate As String

'Dim vCheckQueSend As Integer
'Dim vQTY As Double
'Dim vPickQty As Double

'If Me.ListViewSaleOrder.ListItems.Count > 0 And Me.LBLOrderArCode.Caption <> "" Then
 '  vDocNo = Me.TBDocNo.Text
  ' vDocdate = Me.LBLOrderDocDate.Caption
   'vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   '
   'vQuery = "exec dbo.usp_np_SearchReqPickingInformationLastSend '" & vDocNo & "','" & vQueDocDate & "'," & vTimeID & " "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '  Me.ListViewLastSendQue.ListItems.Clear
     ' vCheckQueSend = 1
      'vRecordset.MoveFirst
      'For i = 1 To vRecordset.RecordCount
       ' Set vListItem = Me.ListViewLastSendQue.ListItems.Add(, , i)
        'vListItem.SubItems(1) = Trim(vRecordset.Fields("queid").Value)
        'vListItem.SubItems(2) = Trim(vRecordset.Fields("quedescription").Value)
        ''vListItem.SubItems(3) = Trim(vRecordset.Fields("itemcode").Value) & "/" & Trim(vRecordset.Fields("itemname").Value)
        'vQTY = Trim(vRecordset.Fields("qty").Value)
        'vPickQty = Trim(vRecordset.Fields("pickqty").Value)
        'vListItem.SubItems(4) = Format(vQTY, "##,##0.00")
        'vListItem.SubItems(5) = Format(vPickQty, "##,##0.00")
        'vListItem.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
        'vListItem.SubItems(7) = Trim(vRecordset.Fields("quepicker").Value)
        'vListItem.SubItems(8) = Trim(vRecordset.Fields("quezone").Value)
        ''vListItem.SubItems(9) = Trim(vRecordset.Fields("quedate").Value)
        'vRecordset.MoveNext
      'Next i
   'End If
   'vRecordset.Close
   
   'If vCheckQueSend = 1 Then
    '  Me.PICLastSendQue.Visible = True
     ' Exit Sub
   'End If
   
   'vQuery = "exec dbo.USP_NP_SearchGroupPicking 2,'" & vDocNo & "' "
   'if OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '   n = vRecordset.RecordCount
     '  vRecordset.MoveFirst
      ' For i = 1 To vRecordset.RecordCount
       'vGroupZone(i) = Trim(vRecordset.Fields("zoneid").Value)
       'vRecordset.MoveNext
       'Next i
   'End If
   'vRecordset.Close
   
   'vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   '
   'For i = 1 To n
    '  If vGroupZone(i) = "A" Then
     '    Call PrintPicking_A(vDocNo, vTimeID, 2)
      'ElseIf vGroupZone(i) = "B" Then
       '  Call PrintPicking_B(vDocNo, vTimeID, 2)
      'ElseIf vGroupZone(i) = "C" Then
       '  Call PrintPicking_C(vDocNo, vTimeID, 2)
      'ElseIf vGroupZone(i) = "X" Then
       '  Call PrintPicking_X(vDocNo, vTimeID, 2)
      'End If
   'Next i
   
   
   'Dim vDay1 As String
   'Dim vMonth1 As String
   
   'If Len(Day(Now)) = 1 Then
    '  vDay1 = Trim("0" & Day(Now))
   'Else
    '  vDay1 = Day(Now)
   'End If
   
   'If Len(Month(Now)) = 1 Then
    '  vMonth1 = Trim("0" & Month(Now))
   'Else
    '  vMonth1 = Month(Now)
   'End If
   
   'Me.DTPDocDate1.Caption = vDay1 & "/" & vMonth1 & "/" & Year(Now)
   'Me.TXTArCode.Text = ""
   'Me.TXTSaleCode.Text = ""
   ''Me.TXTDocNo.Text = ""
   'Me.LBLArName.Caption = ""
   'Me.LBLTotalAmount.Caption = ""
   'Me.CMDQue.Enabled = False
   '
   'Me.PICSendInformation.Visible = True
   'vQuery = "exec dbo.usp_np_SearchReqPickingInformation '" & vDocNo & "'," & vTimeID & " "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '  Me.ListViewInfQue.ListItems.Clear
     ' Me.LBLInfDocNo.Caption = Trim(vRecordset.Fields("docno").Value)
      'Me.LBLInfARName.Caption = Trim(vRecordset.Fields("arname").Value)
      ''vRecordset.MoveFirst
      'For i = 1 To vRecordset.RecordCount
       ' Set vListItem = Me.ListViewInfQue.ListItems.Add(, , i)
       ' vListItem.SubItems(1) = Trim(vRecordset.Fields("queid").Value)
        ''vListItem.SubItems(2) = Trim(vRecordset.Fields("zoneid").Value)
        'vRecordset.MoveNext
      'Next i
   'End If
   'vRecordset.Close

'vQuery = "exec dbo.USP_NP_UpdateSendQueuePicking 1,'" & vDocNo & "' "
'gConnection.Execute vQuery
   
'Else
 '  MsgBox "เอกสารที่จะส่งจัดสินค้าได้ ต้องมีเลขที่เอกสาร รายการสินค้า และต้องเป็นเอกสารที่บันทึกข้อมูลเรียบร้อยแล้วเป็นอย่างน้อย กรุณาตรวจสอบ ", vbCritical, "Send Error Message"
'End If
End Sub


Private Sub CMDSave_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer

Dim vDocNo As String
Dim vDocdate As String
Dim vARCode As String
Dim vSaleCode As String
Dim vRefNo As String
Dim vMemberID As String
Dim vReqTime1 As Date
Dim vReqTime As String
Dim vMydescription As String
Dim vCheckSale As Integer

Dim vHeader As String
Dim vRunning As Integer
Dim vDocNumber As String

Dim vCheckTimeID As Integer

Dim vItemCode As String
Dim vQTY As Double
Dim vUnitCode As String
Dim vPrice As Double
Dim vDiscountWord  As String
Dim vDiscountAmount  As Double
Dim vBeforeTaxAmount As Double
Dim vTaxAmount As Double
Dim vNetDebtAmount As Double
Dim vNetAmount As Double
Dim vIsConditionSend As Integer
Dim vWHCode As String
Dim vShelfCode As String
Dim vShelfID  As String
Dim vZoneID  As String
Dim vBarCode  As String
Dim vLineNumber As Integer

Dim vCheckDate As String
Dim vTime As Date
Dim vNowTime As Date


If Me.TXTDocNo.Text = "" Then
   MsgBox "ต้องระบุเลขที่เอกสารก่อนการบันทึกข้อมูล กรุณาตรวจสอบ", vbCritical, "Send Error Message"
   Me.TBBarCode.SetFocus
   Exit Sub
End If

If Me.TXTArCode.Text = "" Then
   MsgBox "ต้องระบุรหัสลูกค้าก่อนการบันทึกข้อมูล กรุณาตรวจสอบ", vbCritical, "Send Error Message"
   Me.TXTArCode.SetFocus
   Exit Sub
End If

If Me.TXTSaleCode.Text = "" Then
   MsgBox "ต้องระบุรหัสพนักงานก่อนการบันทึกข้อมูล กรุณาตรวจสอบ", vbCritical, "Send Error Message"
   Me.TXTSaleCode.SetFocus
   Exit Sub
End If

If vMemIsCancel = 1 Then
   MsgBox "เอกสารที่ยกเลิกแล้วไม่สามารถแก้ไขข้อมูลได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
   Me.TBBarCode.SetFocus
   Exit Sub
End If

If vSendQue = 1 Then
   MsgBox "เอกสารที่ส่งจัดคิวสินค้าแล้วไม่สามารถแก้ไขข้อมูลได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
   Me.TBBarCode.SetFocus
   Exit Sub
End If

If Me.ListViewItem.ListItems.Count > 0 Then
   If vPRIsOpen = 0 Then
      vQuery = "exec dbo.USP_NP_SearchNewDocNo  32 "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          vHeader = Trim(vRecordset.Fields("header").Value)
          vRunning = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
          vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
      End If
      vRecordset.Close
      
      vDocNo = vDocNumber & vHeader & "-" & Format(vRunning, "0000")
   Else
      vDocNo = Me.TXTDocNo.Text
   End If
         
   vDocdate = Me.DTPDocDate1.Caption
   
   Dim vDay As String
   Dim vMonth As String
   
   If Len(Day(Now)) = 1 Then
      vDay = Trim("0" & Day(Now))
   Else
      vDay = Day(Now)
   End If
   
   If Len(Month(Now)) = 1 Then
      vMonth = Trim("0" & Month(Now))
   Else
      vMonth = Month(Now)
   End If
   
   vCheckDate = vDay & "/" & vMonth & "/" & Year(Now)
   
   If vDocdate < vCheckDate Then
      MsgBox "ไม่สามารถบันทึกเอกสาร ย้อนหลังได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      Me.TBBarCode.SetFocus
      Exit Sub
   End If
   
   vARCode = Me.TXTArCode.Text
   vCheckSale = InStr(Me.TXTSaleCode.Text, "/")
   If vCheckSale = 0 Then
      MsgBox "กรุณาตรวจสอบรหัสพนักงานขาย ต้องมีชื่อภาษาไทยปรากฏด้วย ตัวอย่างเช่น  11111/xxxxx เป็นต้น ", vbCritical, "Send Error Message"
      Me.TXTArCode.SetFocus
      Exit Sub
   End If
   vSaleCode = Left(Me.TXTSaleCode.Text, vCheckSale - 1)
   vMydescription = ""
   If Me.LBLTotalAmount.Caption <> "" Then
      vNetDebtAmount = Me.LBLTotalAmount.Caption
      vBeforeTaxAmount = (vNetDebtAmount * 100) / 107
      vTaxAmount = vNetDebtAmount - vBeforeTaxAmount
   Else
      vNetDebtAmount = 0
      vBeforeTaxAmount = 0
      vTaxAmount = 0
   End If
   vRefNo = Me.TXTLicense.Text
   vMemberID = Me.TXTMember.Caption

    
   vQuery = "exec dbo.USP_NP_InsertPickingRequestMaster '" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRefNo & "','" & vMemberID & "'," & vBeforeTaxAmount & "," & vTaxAmount & "," & vNetDebtAmount & "," & vIsConditionSend & ",'" & vReqTime & "','" & vMydescription & "','" & vUserID & "' "
   gConnection.Execute vQuery

   For i = 1 To Me.ListViewItem.ListItems.Count
    vItemCode = Me.ListViewItem.ListItems(i).SubItems(1)
    vQTY = Me.ListViewItem.ListItems(i).SubItems(3)
    vUnitCode = Me.ListViewItem.ListItems(i).SubItems(4)
    vPrice = Me.ListViewItem.ListItems(i).SubItems(5)
    vDiscountWord = Me.ListViewItem.ListItems(i).SubItems(13)
    vDiscountAmount = Me.ListViewItem.ListItems(i).SubItems(7)
    vNetAmount = Me.ListViewItem.ListItems(i).SubItems(6)
    vWHCode = Me.ListViewItem.ListItems(i).SubItems(8)
    vShelfCode = Me.ListViewItem.ListItems(i).SubItems(9)
    vShelfID = Me.ListViewItem.ListItems(i).SubItems(11)
    vZoneID = Me.ListViewItem.ListItems(i).SubItems(10)
    vBarCode = Me.ListViewItem.ListItems(i).SubItems(12)
    vLineNumber = i - 1
   
   vQuery = "exec dbo.USP_NP_InsertPickingRequestSub '" & vDocNo & "','" & vDocdate & "','" & vItemCode & "'," & vQTY & ",'" & vUnitCode & "'," & vPrice & ",'" & vDiscountWord & "'," & vDiscountAmount & "," & vNetAmount & ",'" & vWHCode & "','" & vShelfCode & "','" & vShelfID & "','" & vZoneID & "','" & vBarCode & "'," & vLineNumber & " "
   gConnection.Execute vQuery
         
   Next i
   
   MsgBox "บันทึกข้อมูลเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว", vbInformation, "Send Information Message"
   
   If vPRIsOpen = 0 Then
   vQuery = "exec dbo.USP_NP_UpdateNewDocNo  32 "
   gConnection.Execute vQuery
   End If
         
   vPRIsOpen = 0
   Me.ListViewItem.ListItems.Clear

   Dim vDay1 As String
   Dim vMonth1 As String
   
   If Len(Day(Now)) = 1 Then
      vDay1 = Trim("0" & Day(Now))
   Else
      vDay1 = Day(Now)
   End If
   
   If Len(Month(Now)) = 1 Then
      vMonth1 = Trim("0" & Month(Now))
   Else
      vMonth1 = Month(Now)
   End If
   
   Me.DTPDocDate1.Caption = vDay1 & "/" & vMonth1 & "/" & Year(Now)

   Me.TXTArCode.Text = ""
   Me.TXTSaleCode.Text = ""
   Me.TXTDocNo.Text = ""
   Me.LBLArName.Caption = ""
   Me.LBLTotalAmount.Caption = ""
   Me.TXTLicense.Text = ""
   Me.TXTMember.Caption = ""
   Me.ListViewItem.ListItems.Clear
   Me.TBBarCode.Text = ""
   
   Call ShowReqDetails(vDocNo)
End If
End Sub

Public Sub ShowReqDetails(vDocNo As String)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Dim vDocdate As String
Dim vARCode As String
Dim vSaleCode As String
Dim vRefNo As String
Dim vMemberID As String
Dim vReqTime1 As Date
Dim vReqTime As String
Dim vMydescription As String

Dim vCheckTimeID As Integer

Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Double
Dim vUnitCode As String
Dim vPrice As Double
Dim vDiscountWord  As String
Dim vDiscountAmount  As Double
Dim vBeforeTaxAmount As Double
Dim vTaxAmount As Double
Dim vNetDebtAmount As Double
Dim vNetAmount As Double
Dim vIsConditionSend As Integer
Dim vWHCode As String
Dim vShelfCode As String
Dim vShelfID  As String
Dim vZoneID  As String
Dim vBarCode  As String
Dim vLineNumber As Integer

Dim vListItem As ListItem
Dim i As Integer
Dim n As Integer


vQuery = "exec dbo.usp_np_SearchReqPickingDetails '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vPRIsOpen = 1
   vMemIsCancel = Trim(vRecordset.Fields("iscancel").Value)
   vSendQue = Trim(vRecordset.Fields("issendque").Value)
   vDocdate = Trim(vRecordset.Fields("docdate").Value)
   vARCode = Trim(vRecordset.Fields("arcode").Value)
   vSaleCode = Trim(vRecordset.Fields("salecode").Value)
   vReqTime = Trim(vRecordset.Fields("reqtime").Value)
   vNetDebtAmount = Trim(vRecordset.Fields("netdebtamount").Value)
   vMydescription = Trim(vRecordset.Fields("mydescription").Value)
   vIsConditionSend = Trim(vRecordset.Fields("isconditionsend").Value)
   vRefNo = Trim(vRecordset.Fields("refno").Value)
   vMemberID = Trim(vRecordset.Fields("memberid").Value)
   
   Me.CMDQue.Enabled = True
   
   Me.TXTDocNo.Text = vDocNo
   Me.DTPDocDate1.Caption = vDocdate
   Me.TXTArCode.Text = vARCode
   Me.TXTSaleCode.Text = vSaleCode
   Me.LBLTotalAmount.Caption = Format(vNetDebtAmount, "##,##0.00")
   Me.TXTLicense.Text = vRefNo
   Me.TXTMember.Caption = vMemberID
   
   Me.ListViewItem.ListItems.Clear
   n = vRecordset.RecordCount
   vRecordset.MoveFirst
   For i = 1 To vRecordset.RecordCount
      vItemCode = Trim(vRecordset.Fields("itemcode").Value)
      vItemName = Trim(vRecordset.Fields("itemname").Value)
      vQTY = Trim(vRecordset.Fields("qty").Value)
      vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
      vPrice = Trim(vRecordset.Fields("price").Value)
      vDiscountWord = Trim(vRecordset.Fields("discountword").Value)
      vDiscountAmount = Trim(vRecordset.Fields("discountamount").Value)
      vNetAmount = Trim(vRecordset.Fields("netamount").Value)
      vWHCode = Trim(vRecordset.Fields("whcode").Value)
      vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
      vShelfID = Trim(vRecordset.Fields("shelfid").Value)
      vZoneID = Trim(vRecordset.Fields("zoneid").Value)
      vBarCode = Trim(vRecordset.Fields("barcode").Value)
      vLineNumber = i
      
     Set vListItem = Me.ListViewItem.ListItems.Add(, , vLineNumber)
     vListItem.SubItems(1) = vItemCode
     vListItem.SubItems(2) = vItemName
     vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
     vListItem.SubItems(4) = vUnitCode
     vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
     vListItem.SubItems(6) = Format(vNetAmount, "##,##0.00")
     vListItem.SubItems(7) = Format(vDiscountAmount, "##,##0.00")
     vListItem.SubItems(8) = vWHCode
     vListItem.SubItems(9) = vShelfCode
     vListItem.SubItems(10) = vZoneID
     vListItem.SubItems(11) = vShelfID
     vListItem.SubItems(12) = vBarCode
     vListItem.SubItems(13) = vDiscountWord
         
      vRecordset.MoveNext
   Next i
Else
   vPRIsOpen = 0
   vMemIsCancel = 0
   vSendQue = 0
   Me.ListViewItem.ListItems.Clear
   'Me.DTPDocDate.Value = Now

   Dim vDay As String
   Dim vMonth As String
   
   If Len(Day(Now)) = 1 Then
      vDay = Trim("0" & Day(Now))
   Else
      vDay = Day(Now)
   End If
   
   If Len(Month(Now)) = 1 Then
      vMonth = Trim("0" & Month(Now))
   Else
      vMonth = Month(Now)
   End If
   
   Me.DTPDocDate1.Caption = vDay & "/" & vMonth & "/" & Year(Now)

   Me.TXTArCode.Text = ""
   Me.TXTSaleCode.Text = ""
   Me.LBLArName.Caption = ""
   Me.LBLTotalAmount.Caption = ""
   Me.TXTLicense.Text = ""
   Me.TXTMember.Caption = ""
   Me.CMDQue.Enabled = False
   Me.ListViewItem.ListItems.Clear
   On Error Resume Next
   Me.TXTDocNo.SetFocus
End If
vRecordset.Close

If vMemIsCancel = 1 Then
   MsgBox "เอกสารเลขที่ " & vDocNo & " ถูกยกเลิกไปแล้ว", vbCritical, "Send Error Message"
   Me.TXTArCode.Text = ""
   Me.TXTSaleCode.Text = ""
   Me.LBLArName.Caption = ""
   Me.LBLTotalAmount.Caption = ""
   Me.TXTLicense.Text = ""
   Me.TXTMember.Caption = ""
   Me.CMDQue.Enabled = False
   Me.ListViewItem.ListItems.Clear
   Call CMDDocNo_Click
   Exit Sub
End If

End Sub

Public Sub ShowDIDetails(vDocNo As String)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Dim vDocdate As String
Dim vARCode As String
Dim vSaleCode As String
Dim vRefNo As String
Dim vMemberID As String
Dim vReqTime1 As Date
Dim vReqTime As String
Dim vMydescription As String

Dim vCheckTimeID As Integer

Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Double
Dim vUnitCode As String
Dim vPrice As Double
Dim vDiscountWord  As String
Dim vDiscountAmount  As Double
Dim vBeforeTaxAmount As Double
Dim vTaxAmount As Double
Dim vNetDebtAmount As Double
Dim vNetAmount As Double
Dim vIsConditionSend As Integer
Dim vWHCode As String
Dim vShelfCode As String
Dim vShelfID  As String
Dim vZoneID  As String
Dim vBarCode  As String
Dim vLineNumber As Integer
Dim vLinePickZone As String

Dim vListItem As ListItem
Dim i As Integer
Dim n As Integer

Dim vPointZone As String
Dim x As Integer
Dim vMemLinePickZone As String

vPointZone = Me.LBLDI.Caption

vQuery = "exec dbo.USP_NP_SearchDriveInDetails1 '" & vDocNo & "','" & vPointZone & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vDIIsOpen = 1
   
   vCountItemPickZoneOld = 0
   vDIIsCancel = vRecordset.Fields("iscancel").Value
   vDIIsConfirm = vRecordset.Fields("isconfirm").Value
   vDIIsSendQue = vRecordset.Fields("issendque").Value
            
   vMemDIIsCancel = Trim(vRecordset.Fields("iscancel").Value)
   vDISendQue = Trim(vRecordset.Fields("issendque").Value)
   vDocdate = Trim(vRecordset.Fields("docdate").Value)
   vARCode = Trim(vRecordset.Fields("arcode").Value)
   vSaleCode = Trim(vRecordset.Fields("salecode").Value)
   vNetDebtAmount = Trim(vRecordset.Fields("totalnetamount").Value)
   vRefNo = Trim(vRecordset.Fields("refno").Value)
   vMemberID = Trim(vRecordset.Fields("memberid").Value)
   
   
    vCountItemOld = vRecordset.RecordCount

    ReDim vDIItemCodeOld(vCountItemOld) As String
    ReDim vDIUnitCodeOld(vCountItemOld) As String
    ReDim vDIWHCodeOld(vCountItemOld) As String
    ReDim vDIShelfCodeOld(vCountItemOld) As String
    ReDim vDIZoneIDOld(vCountItemOld) As String
    ReDim vDIPickZoneOld(vCountItemOld) As String
    ReDim vDIBarCodeOld(vCountItemOld) As String
            
   
   If vDIIsCancel = 1 Then
   Call CancelDoc
   ElseIf vDIIsConfirm = 1 Then
   Call ConfirmDoc
   Else
   Call NewDoc
   End If
   
   Me.LBLDIDocNo.Caption = vDocNo
   Me.LBLDIDocDate.Caption = vDocdate
   Me.TBDIArCode.Text = vARCode
   Me.TBDISaleCode.Text = vSaleCode
   Me.LBLDINetAmount.Caption = Format(vNetDebtAmount, "##,##0.00")
   Me.TBDICarLicense.Text = vRefNo
   Me.LBLDIMember.Caption = vMemberID
  
   Me.ListViewDIItem.ListItems.Clear
   n = vRecordset.RecordCount
   vRecordset.MoveFirst
   For i = 1 To vRecordset.RecordCount
   
      vDIItemCodeOld(i) = vRecordset.Fields("itemcode").Value
      vDIUnitCodeOld(i) = vRecordset.Fields("unitcode").Value
      vDIWHCodeOld(i) = vRecordset.Fields("whcode").Value
      vDIShelfCodeOld(i) = vRecordset.Fields("shelfcode").Value
      vDIZoneIDOld(i) = vRecordset.Fields("zoneid").Value
      vDIBarCodeOld(i) = vRecordset.Fields("barcode").Value
      vDIPickZoneOld(i) = vRecordset.Fields("pickzone").Value
    
      If vPointZone = vDIPickZoneOld(i) Then
         vCountItemPickZoneOld = vCountItemPickZoneOld + 1
      End If
            
      vItemCode = Trim(vRecordset.Fields("itemcode").Value)
      vItemName = Trim(vRecordset.Fields("itemname").Value)
      vQTY = Trim(vRecordset.Fields("qty").Value)
      vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
      vPrice = Trim(vRecordset.Fields("price").Value)
      vDiscountWord = Trim(vRecordset.Fields("discountword").Value)
      vDiscountAmount = Trim(vRecordset.Fields("discountamount").Value)
      vNetAmount = Trim(vRecordset.Fields("amount").Value)
      vWHCode = Trim(vRecordset.Fields("whcode").Value)
      vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
      vShelfID = Trim(vRecordset.Fields("shelfid").Value)
      vZoneID = Trim(vRecordset.Fields("zoneid").Value)
      vBarCode = Trim(vRecordset.Fields("barcode").Value)
      vLinePickZone = Trim(vRecordset.Fields("pickzone").Value)
      vLineNumber = i
      
     Set vListItem = Me.ListViewDIItem.ListItems.Add(, , vLineNumber)
     vListItem.SubItems(1) = vItemCode
     vListItem.SubItems(2) = vItemName
     vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
     vListItem.SubItems(4) = vUnitCode
     vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
     vListItem.SubItems(6) = Format(vNetAmount, "##,##0.00")
     vListItem.SubItems(7) = Format(vDiscountAmount, "##,##0.00")
     vListItem.SubItems(8) = vWHCode
     vListItem.SubItems(9) = vShelfCode
     vListItem.SubItems(10) = vZoneID
     vListItem.SubItems(11) = vShelfID
     vListItem.SubItems(12) = vBarCode
     vListItem.SubItems(13) = vDiscountWord
     vListItem.SubItems(14) = vLinePickZone
     vRecordset.MoveNext
   Next i
   
    For x = 1 To Me.ListViewDIItem.ListItems.Count
    vMemLinePickZone = Me.ListViewDIItem.ListItems(x).SubItems(14)
    If vPointZone <> vMemLinePickZone Then
    Me.ListViewDIItem.ListItems.Item(x).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(1).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(2).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(3).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(4).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(5).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(6).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(7).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(8).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(9).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(10).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(11).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(12).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(13).ForeColor = "&H00008000"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(14).ForeColor = "&H00008000"
    Else
    Me.ListViewDIItem.ListItems.Item(x).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(1).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(2).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(3).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(4).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(5).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(6).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(7).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(8).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(9).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(10).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(11).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(12).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(13).ForeColor = "&H80000008"
    Me.ListViewDIItem.ListItems.Item(x).ListSubItems(14).ForeColor = "&H80000008"
    End If
    
    Next x
            
Else
   vDIIsOpen = 0
   vMemDIIsCancel = 0
   vDISendQue = 0
   Me.ListViewItem.ListItems.Clear

   Dim vDay As String
   Dim vMonth As String
   
   If Len(Day(Now)) = 1 Then
      vDay = Trim("0" & Day(Now))
   Else
      vDay = Day(Now)
   End If
   
   If Len(Month(Now)) = 1 Then
      vMonth = Trim("0" & Month(Now))
   Else
      vMonth = Month(Now)
   End If
   
   Me.LBLDIDocDate.Caption = vDay & "/" & vMonth & "/" & Year(Now)

   Me.TBDIArCode.Text = ""
   Me.TBDISaleCode.Text = ""
   Me.LBLDIArName.Caption = ""
   Me.LBLDINetAmount.Caption = ""
   Me.TBDICarLicense.Text = ""
   Me.LBLDIMember.Caption = ""

   Me.ListViewDIItem.ListItems.Clear
   Me.TBDIBarCode.SetFocus
End If
vRecordset.Close

If vMemDIIsCancel = 1 Then
   MsgBox "เอกสารเลขที่ " & vDocNo & " ถูกยกเลิกไปแล้ว", vbCritical, "Send Error Message"
   
   Me.TBDISaleCode.Text = ""
   Me.LBLDINetAmount.Caption = ""
   Me.TBDICarLicense.Text = ""
   Me.LBLDIMember.Caption = ""
   Me.ListViewDIItem.ListItems.Clear
   Me.CMDDISendQue.Enabled = False
   Call CMDDocNo_Click
   Exit Sub
End If

End Sub

Private Sub CMDSave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If
End Sub

Private Sub CMDSearch_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearch As String
Dim vListItem As ListItem
Dim i As Integer

vSearch = ""
vQuery = "exec dbo.USP_NP_SearchReqPicking '" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.ListViewPRSearchDocNo.ListItems.Clear
    vRecordset.MoveFirst
    For i = 1 To vRecordset.RecordCount
       Set vListItem = Me.ListViewPRSearchDocNo.ListItems.Add(, , i)
       vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
       vListItem.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
       vListItem.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
       vListItem.SubItems(4) = Trim(vRecordset.Fields("salename").Value)
       vListItem.SubItems(5) = Trim(vRecordset.Fields("netdebtamount").Value)
       vRecordset.MoveNext
    Next i
End If
vRecordset.Close

Me.PICPRSearchDocNo.Visible = True
Me.TBPRSearchDocNo.SetFocus
End Sub

'Private Sub CMDSearchAR_Click()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vSearchAR As String
'Dim vListItem As ListItem
'Dim i As Integer

'If Me.TXTSearchAR.Text <> "" Then
 '  vSearchAR = Me.TXTSearchAR.Text
  ' vQuery = "exec dbo.USP_AR_ARProFileSearch '" & vSearchAR & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '   Me.ListViewAR.ListItems.Clear
     '  vRecordset.MoveFirst
      ' For i = 1 To vRecordset.RecordCount
       '   Set vListItem = Me.ListViewAR.ListItems.Add(, , i)
        '  vListItem.SubItems(1) = Trim(vRecordset.Fields("arcode").Value)
         ' vListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
          ''vListItem.SubItems(3) = Trim(vRecordset.Fields("memberid").Value)
          'vRecordset.MoveNext
       'Next i
       'Me.ListViewAR.SetFocus
   'End If
   'vRecordset.Close
'End If
'End Sub

Private Sub CMDSearchARClose_Click()
'Me.PICAR.Visible = False
End Sub

Private Sub CMDSearchAROK_Click()
'Dim vIndex As Integer
'Dim vARCode As String

'If Me.ListViewAR.ListItems.Count > 0 Then
 '  vIndex = Me.ListViewAR.SelectedItem.Index
  ' vARCode = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(1))
   'Me.TXTArCode.Text = vARCode
   'Me.LBLArName.Caption = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(2))
   ''Me.TXTMember.Caption = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(3))
   'Me.TXTArCode.SetFocus
'End If
'Me.PICAR.Visible = False
End Sub

Private Sub CMDSearchClose_Click()
'Me.PICSearch.Visible = False
End Sub

Private Sub CMDSearchDocNo_Click()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vSearch As String
'Dim vListItem As ListItem
'Dim i As Integer

'If Me.TXTSearchDocNo.Text <> "" Then
'vSearch = Me.TXTSearchDocNo.Text
'vQuery = "exec dbo.USP_NP_SearchReqPicking '" & vSearch & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   Me.ListViewDocNo.ListItems.Clear
  '  vRecordset.MoveFirst
   ' For i = 1 To vRecordset.RecordCount
    '   Set vListItem = Me.ListViewDocNo.ListItems.Add(, , i)
     '  vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
      ' vListItem.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
       'vListItem.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
       'vListItem.SubItems(4) = Trim(vRecordset.Fields("salename").Value)
       'vListItem.SubItems(5) = Trim(vRecordset.Fields("netdebtamount").Value)
       'vRecordset.MoveNext
    'Next i
    
    'Me.ListViewDocNo.SetFocus
'End If
'vRecordset.Close
'End If
'Me.TXTSearchDocNo.SetFocus
End Sub

Private Sub CMDSearchDocNoOK_Click()
'Dim vIndex As Integer
'Dim vDocNo As String

'If Me.ListViewDocNo.ListItems.Count > 0 Then
 '  vIndex = Me.ListViewDocNo.SelectedItem.Index
  ' vDocNo = Trim(Me.ListViewDocNo.ListItems(vIndex).SubItems(1))
   'Me.TXTDocNo.Text = vDocNo
   'Me.TBBarCode.SetFocus
'End If
'Me.PICSearchDoc.Visible = False
End Sub

Private Sub CMDSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If
End Sub

Private Sub CMDSearchItem_Click()
Me.PICPRSearchItem.Visible = True
Me.TBPRSearchItem.SetFocus
End Sub

Private Sub CMDSearchItemList_Click()
'Dim vSearch As String
'Dim vPrice As Double
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
''Dim vListItem As ListItem
'Dim i As Integer
'Dim vQTY As Double
'Dim vRemainOutQTY As Double
'
'If Me.TXTSearch.Text <> "" Then
 '  Me.PICSearchItem.Visible = True
  ' vSearch = Me.TXTSearch.Text
   'vQuery = "exec dbo.USP_NP_SearchBarCode '" & vSearch & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '   Me.ListViewSearch.ListItems.Clear
     '  vRecordset.MoveFirst
      ' For i = 1 To vRecordset.RecordCount
       '   vQTY = Trim(vRecordset.Fields("stockqty").Value)
        '  vRemainOutQTY = Trim(vRecordset.Fields("remainoutqty").Value)
         ' vPrice = Trim(vRecordset.Fields("price").Value)
          '
          'Set vListItem = Me.ListViewSearch.ListItems.Add(, , i)
          'vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
          'vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
          ''vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
          'vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
          'vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
          ''vListItem.SubItems(6) = Trim(vRecordset.Fields("zone").Value)
          'vRecordset.MoveNext
       'Next i
       
       'Me.ListViewItem.SetFocus
    'Else
     '  Me.TBPRSearchItem.SetFocus
    'End If
    'vRecordset.Close
'End If
End Sub

Private Sub CMDSearchOK_Click()
'Dim vIndex As Integer
'Dim vBarCode As String

'If Me.ListViewSearch.ListItems.Count > 0 Then
'vIndex = Me.ListViewSearch.SelectedItem.Index
'vBarCode = Me.ListViewSearch.ListItems(vIndex).SubItems(1)
''Me.TBBarCode.Text = vBarCode
'Me.TBBarCode.SetFocus
'End If
'Me.PICSearch.Visible = False
End Sub

Private Sub CMDSearchSale_Click()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vSearchSale As String
'Dim vListItem As ListItem
''Dim i As Integer

'If Me.TXTSearchSale.Text <> "" Then
 '  vSearchSale = Me.TXTSearchSale.Text
  ' vQuery = "exec dbo.USP_CRM_EmployeeDetails 0,'" & vSearchSale & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '   Me.ListViewSaleCode.ListItems.Clear
     '  vRecordset.MoveFirst
      '' For i = 1 To vRecordset.RecordCount
        '  Set vListItem = Me.ListViewSaleCode.ListItems.Add(, , i)
         ' vListItem.SubItems(1) = Trim(vRecordset.Fields("empcode").Value)
          ''vListItem.SubItems(2) = Trim(vRecordset.Fields("empname").Value)
          'vRecordset.MoveNext
       'Next i
       
       'Me.ListViewSaleCode.SetFocus
   'End If
   'vRecordset.Close
'End If
End Sub

Private Sub CMDSearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If
End Sub

Private Sub CMDSearchSaleOrder_Click()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim i As Integer
'Dim vListItem As ListItem

'Dim vDocNo As String
'Dim vSumItemAmount As Double
'Dim vTaxAmount As Double
'Dim vNetAmount As Double

'Dim vQTY As Double
'Dim vPrice As Double
'Dim vDiscountAmount As Double
''Dim vAmount As Double

'If Me.TBDocNo.Text <> "" Then
 '  vDocNo = Me.TBDocNo.Text
'
 '  vQuery = "exec dbo.USP_NP_SearchSaleOrder '" & vDocNo & "'"
  ' If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   '   Me.ListViewSaleOrder.ListItems.Clear
    '  vSumItemAmount = Trim(vRecordset.Fields("sumofitemamount").Value)
     ' vTaxAmount = Trim(vRecordset.Fields("taxamount").Value)
      'vNetAmount = Trim(vRecordset.Fields("netamount").Value)
      
      'Me.LBLOrderBillType.Caption = Trim(vRecordset.Fields("billtype").Value)
      'Me.LBLOrderSoStatus.Caption = Trim(vRecordset.Fields("sostatus").Value)
      'Me.LBLOrderDocDate.Caption = Trim(vRecordset.Fields("docdate").Value)
      'Me.LBLOrderArCode.Caption = Trim(vRecordset.Fields("arcode").Value)
      'Me.LBLOrderArName.Caption = Trim(vRecordset.Fields("arname").Value)
      'If Trim(vRecordset.Fields("salecode").Value) = "" Then
       '   MsgBox "เอกสารไม่ได้กำหนดรหัสพนักงานขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      'End If
      'Me.LBLOrderSaleCode.Caption = Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      'Me.LBLOrderSumOfItemAmount.Caption = Format(vSumItemAmount, "##,##0.00")
      'Me.LBLOrderTaxAmount.Caption = Format(vTaxAmount, "##,##0.00")
      'Me.LBLOrderNetAmount.Caption = Format(vNetAmount, "##,##0.00")
      
      'vRecordset.MoveFirst
      'For i = 1 To vRecordset.RecordCount
       ' Set vListItem = Me.ListViewSaleOrder.ListItems.Add(, , i)
        
       ' vQTY = Trim(vRecordset.Fields("remainqty").Value)
        'vPrice = Trim(vRecordset.Fields("price").Value)
        'vDiscountAmount = Trim(vRecordset.Fields("discountamountsub").Value)
        'vAmount = Trim(vRecordset.Fields("amount").Value)

        'vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
        'vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
        'vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
        ''vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
        ''vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
        'vListItem.SubItems(6) = Format(vDiscountAmount, "##,##0.00")
        'vListItem.SubItems(7) = Format(vAmount, "##,##0.00")
        'vListItem.SubItems(8) = Trim(vRecordset.Fields("whcode").Value)
        'vListItem.SubItems(9) = Trim(vRecordset.Fields("shelfcode").Value)
        'vListItem.SubItems(10) = Trim(vRecordset.Fields("zoneid").Value)
        'vListItem.SubItems(11) = Trim(vRecordset.Fields("shelfid").Value)
        'vListItem.SubItems(12) = Trim(vRecordset.Fields("itemcode").Value)
        'vListItem.SubItems(13) = Trim(vRecordset.Fields("discountwordsub").Value)
        'vListItem.SubItems(14) = Format(vQTY, "##,##0.00")
        'vRecordset.MoveNext
      'Next i
   'Else
     ' Me.LBLOrderDocDate.Caption = ""
      'Me.LBLOrderArCode.Caption = ""
      'Me.LBLOrderArName.Caption = ""
      'Me.LBLOrderSaleCode.Caption = ""
      'Me.LBLOrderSumOfItemAmount.Caption = ""
      'Me.LBLOrderTaxAmount.Caption = ""
      'Me.LBLOrderNetAmount.Caption = ""
      'Me.LBLOrderBillType.Caption = ""
      'Me.LBLOrderSoStatus.Caption = ""
   'End If
   'vRecordset.Close
'Else
 '  MsgBox "กรุณากรอกเลขที่ใบสั่งขาย/จอง ให้ถูกต้อง", vbCritical, "Send Error Message"
'End If
End Sub

Private Sub CMDSendQueAgain_Click()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim i As Integer
'Dim n As Integer
'Dim vDocNo As String
'Dim vDocdate As String
'Dim vQueDocDate As String
'
'Dim vCheckTimeID As Integer
'Dim vGroupZone(5) As String

'Dim vListItem As ListItem
'Dim vCheckDate As String

'If Me.ListViewItem.ListItems.Count > 0 And Me.TXTDocNo.Text <> "" Then
 '  vDocNo = Me.TXTDocNo.Text
  ' vDocdate = Me.DTPDocDate1.Caption
   '
   'Dim vDay As String
   'Dim vMonth As String
   
   'If Len(Day(Now)) = 1 Then
    '  vDay = Trim("0" & Day(Now))
   'Else
    '  vDay = Day(Now)
   'End If
   
   'If Len(Month(Now)) = 1 Then
    '  vMonth = Trim("0" & Month(Now))
   'Else
    '  vMonth = Month(Now)
   'End If
   
   'vCheckDate = vDay & "/" & vMonth & "/" & Year(Now)
   
   'If vDocdate < vCheckDate Then
    '  MsgBox "ไม่สามารถส่งคิวจัดสินค้าได้ กรณีวันที่เอกสารน้อยกว่าวันที่ปัจจุบัน", vbCritical, "Send Error Message"
     ' Exit Sub
   'End If
   
   'vQuery = "exec dbo.USP_NP_SearchGroupPicking 1,'" & vDocNo & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '   n = vRecordset.RecordCount
     '  vRecordset.MoveFirst
      ' For i = 1 To vRecordset.RecordCount
       'vGroupZone(i) = Trim(vRecordset.Fields("zoneid").Value)
       'vRecordset.MoveNext
       'Next i
   'End If
   'vRecordset.Close
   
   'vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   
    'vQuery = "exec dbo.USP_NP_CheckQuePickCenter '" & vDocNo & "','" & vQueDocDate & "' "
    'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     '  vCheckTimeID = Trim(vRecordset.Fields("max1").Value)
    'End If
    'vRecordset.Close
    
   'vCheckTimeID = vCheckTimeID + 1
   
   'For i = 1 To n
    '  If vGroupZone(i) = "A" Then
     '    Call PrintPicking_A(vDocNo, vCheckTimeID)
      'ElseIf vGroupZone(i) = "B" Then
       '  Call PrintPicking_B(vDocNo, vCheckTimeID)
      'ElseIf vGroupZone(i) = "C" Then
       '  Call PrintPicking_C(vDocNo, vCheckTimeID)
      'ElseIf vGroupZone(i) = "X" Then
       '  Call PrintPicking_X(vDocNo, vCheckTimeID)
      'End If
   'Next i
   
   'vPRIsOpen = 0
   'Me.ListViewItem.ListItems.Clear
   
   'Dim vDay1 As String
   'Dim vMonth1 As String
   
   'If Len(Day(Now)) = 1 Then
    '  vDay1 = Trim("0" & Day(Now))
   'Else
    '  vDay1 = Day(Now)
   'End If
   
   'If Len(Month(Now)) = 1 Then
    '  vMonth1 = Trim("0" & Month(Now))
   'Else
    '  vMonth1 = Month(Now)
   'End If
   
   'Me.DTPDocDate1.Caption = vDay1 & "/" & vMonth1 & "/" & Year(Now)

   'Me.DTPReqTime.Value = Now
   'Me.TXTArCode.Text = ""
   'Me.TXTSaleCode.Text = ""
   'Me.TXTDocNo.Text = ""
   'Me.LBLArName.Caption = ""
   'Me.TBMyDescription.Text = ""
   'Me.LBLTotalAmount.Caption = ""
   'Me.CBReqTime.Value = 0
   'Me.CBSend.Value = 0
   
   'Me.PICSendInformation.Visible = True
   'vQuery = "exec dbo.usp_np_SearchReqPickingInformation '" & vDocNo & "'," & vCheckTimeID & " "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '  Me.ListViewInfQue.ListItems.Clear
     ' Me.LBLInfDocNo.Caption = Trim(vRecordset.Fields("docno").Value)
      'Me.LBLInfARName.Caption = Trim(vRecordset.Fields("arname").Value)
      'vRecordset.MoveFirst
      'For i = 1 To vRecordset.RecordCount
       ' Set vListItem = Me.ListViewInfQue.ListItems.Add(, , i)
        'vListItem.SubItems(1) = Trim(vRecordset.Fields("queid").Value)
        'vListItem.SubItems(2) = Trim(vRecordset.Fields("zoneid").Value)
        'vRecordset.MoveNext
      'Next i
   'End If
   'vRecordset.Close

'vQuery = "exec dbo.USP_NP_UpdateSendQueuePicking 1,'" & vDocNo & "' "
'gConnection.Execute vQuery
   
'Me.PICLastSendQue.Visible = False
'End If
End Sub

Private Sub CMDSendQueExit_Click()
Me.PICLastSendQue.Visible = False
Me.TXTDocNo.SetFocus
Call CMDDocNo_Click
End Sub

Private Sub CMDSearchSaleOrder_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 116 Then
'Call CMDSaleOrderSendQue_Click
'End If
End Sub

Private Sub CMDSendQueExit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICLastSendQue.Visible = False
'Me.TBDocNo.SetFocus
End If
End Sub

Private Sub CMDSOMain_Click()
'Me.PICSelectJob.Visible = True
'Me.PICSelectDI.Visible = False
'Me.PICPickReq.Visible = False
'Me.PICDriveIn.Visible = False
'Me.CMDDriveIn.SetFocus
End Sub

Private Sub Command1_Click()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vQueID As Integer
'Dim vQueDocDate As String


'vQueID = 1
'vQueDocDate = "11/02/2009"

'vQuery = "exec dbo.USP_NP_SearchPrintPicking " & vQueID & ",'" & vQueDocDate & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   Text1.Text = Trim(vRecordset.Fields("queid").Value)
  '  Text2.Text = Trim(vRecordset.Fields("quedocdate").Value)
'End If
'vRecordset.Close

'PIC.Visible = True
End Sub


Private Sub CMDSendQuePrint_Click()
'Dim vDocNo As String

'If Me.LBLDIDocNo.Caption <> "" And vDIIsOpen = 1 And vMemDIIsCancel = 0 Then
 '   vDocNo = Me.LBLDIDocNo.Caption
  '  Call PrintDriveInDetails(vDocNo)
   ' MsgBox "พิมพ์ทดแทนเอกสาร Pickup เรียบร้อยแล้ว", vbInformation, "Send Information Message"
'End If
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vHeader As String
Dim vRunning As Integer
Dim vDocNumber As String

vQuery = "exec dbo.USP_NP_SearchNewDocNo  32 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vHeader = Trim(vRecordset.Fields("header").Value)
    vRunning = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
End If
vRecordset.Close
      
vDocNo = vDocNumber & vHeader & "-" & Format(vRunning, "0000")
Me.TXTDocNo.Text = vDocNo
Dim vDay As String
Dim vMonth As String

If Len(Day(Now)) = 1 Then
   vDay = Trim("0" & Day(Now))
Else
   vDay = Day(Now)
End If

If Len(Month(Now)) = 1 Then
   vMonth = Trim("0" & Month(Now))
Else
   vMonth = Month(Now)
End If

Me.DTPDocDate1.Caption = vDay & "/" & vMonth & "/" & Year(Now)
Call NewDoc


'Public Const vbViolet = &HFF8080
'Public Const vbVioletBright = &HFFC0C0
'Public Const vbForestGreen = &H228B22
'Public Const vbGray = &HE0E0E0
'Public Const vbLightBlue = &HFFD3A4
'Public Const vbLightGreen = &HABFCBD
'Public Const vbGreenLemon = &HB3FFBE
'Public Const vbYellowBright = &HC0FFFF
'Public Const vbOrange = &H2CCDFC

        
Call SetListViewColor(ListViewPRSearchSale, PicPoint, vbWhite, vbLightGreen)
Call SetListViewColor(ListViewLastSendQue, PicPoint, vbWhite, vbLightGreen)
Call SetListViewColor(ListViewPRSearchAR, PicPoint, vbWhite, vbLightGreen)
Call SetListViewColor(ListViewPRSearchDocNo, PicPoint, vbWhite, vbLightGreen)
Call SetListViewColor(ListViewInfQue, PicPoint, vbWhite, vbLightGreen)
Call SetListViewColor(ListViewPRSearchItem, PicPoint1, vbWhite, vbLightGreen)
Call SetListViewColor(ListViewItem, PicPoint, vbWhite, vbVeryLightBlue)


Call SetListViewColor(ListViewDIItem, PICPoint2, vbWhite, vbVeryLightBlue)
Call SetListViewColor(ListViewDISearchDI, PICPoint3, vbWhite, vbVeryLightGreen)
Call SetListViewColor(ListViewDISearchAR, PICPoint2, vbWhite, vbVeryLightGreen)
Call SetListViewColor(ListViewDISearchItem, PICPoint2, vbWhite, vbVeryLightGreen)
Call SetListViewColor(ListViewDILastSendQue, PICPoint2, vbWhite, vbVeryLightGreen)
Call SetListViewColor(ListViewDISearchSale, PICPoint2, vbWhite, vbVeryLightGreen)
Call SetListViewColor(ListViewDIInfQue, PICPoint2, vbWhite, vbVeryLightGreen)

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

Public Sub NewDoc()
Me.IMNormal.Visible = True
Me.IMConfirm.Visible = False
Me.IMCancel.Visible = False
End Sub

Public Sub ConfirmDoc()
Me.IMNormal.Visible = False
Me.IMConfirm.Visible = True
Me.IMCancel.Visible = False
End Sub


Public Sub CancelDoc()
Me.IMNormal.Visible = False
Me.IMConfirm.Visible = False
Me.IMCancel.Visible = True
End Sub

Private Sub LBLDIDocNo_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vDocdate As String
Dim vARCode As String
Dim vSaleCode As String
Dim vRefNo As String
Dim vMemberID As String
Dim vReqTime1 As Date
Dim vReqTime As String
Dim vMydescription As String

Dim vCheckTimeID As Integer

Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Double
Dim vUnitCode As String
Dim vPrice As Double
Dim vDiscountWord  As String
Dim vDiscountAmount  As Double
Dim vBeforeTaxAmount As Double
Dim vTaxAmount As Double
Dim vNetDebtAmount As Double
Dim vNetAmount As Double
Dim vIsConditionSend As Integer
Dim vWHCode As String
Dim vShelfCode As String
Dim vShelfID  As String
Dim vZoneID  As String
Dim vBarCode  As String
Dim vLineNumber As Integer

Dim vListItem As ListItem
Dim i As Integer
Dim n As Integer

'vDocNo = Me.LBLDIDocNo.Caption
'vQuery = "exec dbo.USP_NP_SearchDriveInDetails '" & vDocNo & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '  vDIIsOpen = 1
  ' vMemDIIsCancel = Trim(vRecordset.Fields("iscancel").Value)
   'vDISendQue = Trim(vRecordset.Fields("issendque").Value)
   'vDocdate = Trim(vRecordset.Fields("docdate").Value)
   'vARCode = Trim(vRecordset.Fields("arcode").Value)
   'vSaleCode = Trim(vRecordset.Fields("salecode").Value)
   'vNetDebtAmount = Trim(vRecordset.Fields("TotalNetAmount").Value)
   'vRefNo = Trim(vRecordset.Fields("refno").Value)
   'vMemberID = Trim(vRecordset.Fields("memberid").Value)
   
   'Me.CMDDISendQue.Enabled = True
   
   'If vMemDIIsCancel = 1 Then
   'Call CancelDoc
   'ElseIf vDISendQue = 1 Then
   'Call ConfirmDoc
   'Else
   'Call NewDoc
   'End If
   
   'Me.LBLDIDocNo.Caption = vDocNo
   'Me.LBLDIDocDate.Caption = vDocdate
   'Me.TBDIArCode.Text = vARCode
   'Me.TBDISaleCode.Text = vSaleCode
   'Me.LBLDINetAmount.Caption = Format(vNetDebtAmount, "##,##0.00")
   'Me.TBDICarLicense.Text = vRefNo
   'Me.LBLDIMember.Caption = vMemberID
  
   'Me.ListViewDIItem.ListItems.Clear
   'n = vRecordset.RecordCount
   'vRecordset.MoveFirst
   'For i = 1 To vRecordset.RecordCount
    '  vItemCode = Trim(vRecordset.Fields("itemcode").Value)
     ' vItemName = Trim(vRecordset.Fields("itemname").Value)
      'vQTY = Trim(vRecordset.Fields("qty").Value)
      'vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
      'vPrice = Trim(vRecordset.Fields("price").Value)
      'vDiscountWord = Trim(vRecordset.Fields("discountword").Value)
      'vDiscountAmount = Trim(vRecordset.Fields("discountamount").Value)
      'vNetAmount = Trim(vRecordset.Fields("amount").Value)
      'vWHCode = Trim(vRecordset.Fields("whcode").Value)
      'vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
      'vShelfID = Trim(vRecordset.Fields("shelfid").Value)
      'vZoneID = Trim(vRecordset.Fields("zoneid").Value)
      'vBarCode = Trim(vRecordset.Fields("barcode").Value)
      'vLineNumber = i
      
     'Set vListItem = Me.ListViewDIItem.ListItems.Add(, , vLineNumber)
     'vListItem.SubItems(1) = vItemCode
     'vListItem.SubItems(2) = vItemName
     'vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
     'vListItem.SubItems(4) = vUnitCode
     'vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
     'vListItem.SubItems(6) = Format(vNetAmount, "##,##0.00")
     'vListItem.SubItems(7) = Format(vDiscountAmount, "##,##0.00")
     'vListItem.SubItems(8) = vWHCode
     'vListItem.SubItems(9) = vShelfCode
     'vListItem.SubItems(10) = vZoneID
     'vListItem.SubItems(11) = vShelfID
     'vListItem.SubItems(12) = vBarCode
     'vListItem.SubItems(13) = vDiscountWord
         
      'vRecordset.MoveNext
   'Next i
'Else
  ' vDIIsOpen = 0
   'vMemDIIsCancel = 0
   'vDISendQue = 0
   'Me.ListViewItem.ListItems.Clear

   'Dim vDay As String
   'Dim vMonth As String
   
   'If Len(Day(Now)) = 1 Then
    '  vDay = Trim("0" & Day(Now))
   'Else
    '  vDay = Day(Now)
   'End If
   
   'If Len(Month(Now)) = 1 Then
    '  vMonth = Trim("0" & Month(Now))
   'Else
    '  vMonth = Month(Now)
   'End If
   
   'Me.LBLDIDocDate.Caption = vDay & "/" & vMonth & "/" & Year(Now)

   'Me.TBDISaleCode.Text = ""
   'Me.LBLDINetAmount.Caption = ""
   'Me.TBDICarLicense.Text = ""
   'Me.LBLDIMember.Caption = ""

   'Me.ListViewDIItem.ListItems.Clear
   'Me.TBDIArCode.Text = "99999"
   'Me.TBDIArCode.SetFocus
'End If
'vRecordset.Close

'If vMemDIIsCancel = 1 Then
 '  MsgBox "เอกสารเลขที่ " & vDocNo & " ถูกยกเลิกไปแล้ว", vbCritical, "Send Error Message"
  ' Me.TXTArCode.Text = ""
   'Me.TXTSaleCode.Text = ""
   'Me.LBLArName.Caption = ""
   'Me.LBLTotalAmount.Caption = ""
   'Me.TXTLicense.Text = ""
   'Me.TXTMember.Caption = ""
   'Me.CMDDISendQue.Enabled = False
   'Me.ListViewItem.ListItems.Clear
   'Call CMDDocNo_Click
   'Exit Sub
'End If

End Sub

Private Sub ListViewAR_DblClick()
'Dim vIndex As Integer
'Dim vARCode As String

'If Me.ListViewAR.ListItems.Count > 0 Then
 '  vIndex = Me.ListViewAR.SelectedItem.Index
  ' vARCode = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(1))
   'Me.TXTArCode.Text = vARCode
   'Me.LBLArName.Caption = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(2))
   'Me.TXTMember.Caption = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(3))
   'Me.TXTArCode.SetFocus
'End If
'Me.PICAR.Visible = False
End Sub

Private Sub ListViewAR_KeyPress(KeyAscii As Integer)
'Dim vIndex As Integer
'Dim vARCode As String

'If KeyAscii = 13 Then
 '  If Me.ListViewAR.ListItems.Count > 0 Then
  '    vIndex = Me.ListViewAR.SelectedItem.Index
   '   vARCode = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(1))
    '  Me.TXTArCode.Text = vARCode
     ' Me.LBLArName.Caption = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(2))
      'Me.TXTMember.Caption = Trim(Me.ListViewAR.ListItems(vIndex).SubItems(3))
      'Me.TXTArCode.SetFocus
   'End If
'End If

'Me.PICAR.Visible = False
End Sub

Private Sub ListViewDIInfQue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.LBLDIInfDocNo.Caption = ""
   Me.ListViewDIInfQue.ListItems.Clear
   Me.PICDISendInformation.Visible = False
   Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub ListViewDIItem_DblClick()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vBarCode As String
Dim vIndex As Integer
Dim vQTY As Double
Dim vDiscountWord As String
Dim vDiscountAmount As Double
Dim vNetAmount As Double
Dim vPrice As Double
Dim i As Integer
Dim vListItem As ListItem

Dim vOldQty As Double
Dim vUnitCode As String
Dim vItemCode As String
Dim vNewPrice As Double
Dim vPickZone As String
Dim vLinePickZone As String


If Me.ListViewDIItem.ListItems.Count > 0 Then
   If vDIIsConfirm = 0 Then
      vPickZone = Me.LBLDI.Caption
   
      vIndex = Me.ListViewDIItem.SelectedItem.Index
      Me.PICDIKeyQty.Visible = True
      vBarCode = Me.ListViewDIItem.ListItems(vIndex).SubItems(1)
      vLinePickZone = Me.ListViewDIItem.ListItems(vIndex).ListSubItems(14).Text
            
      If vPickZone <> vLinePickZone Then
      MsgBox "ไม่สามารถแก้ไขรายการสินค้าที่ได้สั่งจ่ายมาจากจุดจ่ายอื่นได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      Me.ListViewDIItem.SetFocus
      Exit Sub
      End If
      
      vUnitCode = Me.ListViewDIItem.ListItems(vIndex).SubItems(4)
      vItemCode = Me.ListViewDIItem.ListItems(vIndex).SubItems(1)
      vOldQty = Me.ListViewDIItem.ListItems(vIndex).SubItems(3)

      vQuery = "exec dbo.USP_MB_SearchBarcode '" & vBarCode & "'"
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Me.LBLDIItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
          Me.LBLDIItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
          Me.LBLDIUnitCode.Caption = Trim(vRecordset.Fields("unitcode").Value)
          Me.LBLDIWHCode.Caption = Trim(vRecordset.Fields("defsalewhcode").Value)
          Me.LBLDIShelfCode.Caption = Trim(vRecordset.Fields("defsaleshelf").Value)
          Me.LBLDIShelfID.Caption = Trim(vRecordset.Fields("shelfid").Value)
          Me.LBLDIZoneID.Caption = Trim(vRecordset.Fields("zoneid").Value)
          Me.LBLDIBarCode.Caption = Trim(vRecordset.Fields("barcode").Value)
          
          vPrice = Trim(vRecordset.Fields("price").Value)
          If vPrice <= 0 Then
             MsgBox "สินค้ารายการนี้ ยังไม่ได้กำหนดราคาขายของหน่วยนับขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
             Me.TBDIKeyQty.Enabled = False
             Me.TBDIKeyDiscount.Enabled = False
             Exit Sub
          Else
             Me.TBDIKeyQty.Enabled = True
             Me.TBDIKeyDiscount.Enabled = False
          End If
          
          Me.LBLDIPrice.Caption = Format(vPrice, "##,##0.00")
          
          Me.ListViewDIItemStock.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             vQTY = Trim(vRecordset.Fields("stock").Value)
             
            Set vListItem = Me.ListViewDIItemStock.ListItems.Add(, , Trim(vRecordset.Fields("whcode").Value))
            vListItem.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
            vListItem.SubItems(2) = Format(vQTY, "##,##0.00")
            vListItem.SubItems(3) = Trim(vRecordset.Fields("stkunitcode").Value)
                
             vRecordset.MoveNext
          Next i
             
          vQTY = Me.ListViewDIItem.ListItems(vIndex).SubItems(3)
          vDiscountWord = Me.ListViewDIItem.ListItems(vIndex).SubItems(13)
          vDiscountAmount = Me.ListViewDIItem.ListItems(vIndex).SubItems(7)
          vNetAmount = Me.ListViewDIItem.ListItems(vIndex).SubItems(6)
          
          
          Me.TBDIKeyQty.Text = Format(vQTY, "##,##0.00")
          Me.TBDIKeyDiscount.Text = vDiscountWord
          Me.LBLDIDiscountWord.Caption = Format(vDiscountAmount, "##,##0.00")
          Me.LBLDIItemNetAmount.Caption = Format(vNetAmount, "##,##0.00")

          Me.TBDIKeyQty.SetFocus
       End If
    vRecordset.Close
    
    vQuery = "exec dbo.USP_NP_SearchItemPriceQty1 '" & vItemCode & "'," & vOldQty & ",'" & vUnitCode & "'"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vNewPrice = Trim(vRecordset.Fields("saleprice1").Value)
       Me.LBLDIPrice.Caption = Format(vNewPrice, "##,##0.00")
    End If
    vRecordset.Close
    
   Else
      MsgBox "ไม่สามารถแก้ไขรายการของเอกสารนี้ได้ เนื่องจากเอกสารนี้ถูกอ้างไปจัดคิวเรียบร้อยแล้ว", vbCritical, "Send Error Message"
      Me.TBDICarLicense.SetFocus
   End If
End If
End Sub

Private Sub ListViewDIItem_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer
Dim vDocNo As String
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSendQue As Integer
Dim vPickZone As String
Dim vLinePickZone As String

If Me.ListViewDIItem.ListItems.Count > 0 Then
   If KeyCode = 46 Then
      If vDIIsOpen = 0 Then
      If vDIIsConfirm = 0 Then
         vDocNo = Me.LBLDIDocNo.Caption
         vQuery = "exec dbo.USP_NP_SearchSendQueuePicking '" & vDocNo & "' "
          If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
              vSendQue = Trim(vRecordset.Fields("issendque").Value)
          End If
          vRecordset.Close
          
         If vDIIsConfirm = 0 Then
            vPickZone = Me.LBLDI.Caption
            vIndex = Me.ListViewDIItem.SelectedItem.Index
            vLinePickZone = Me.ListViewDIItem.ListItems(vIndex).ListSubItems(14).Text
            If vPickZone <> vLinePickZone Then
            MsgBox "ไม่สามารถแก้ไข รายการสินค้าที่ถูกจ่ายมาจากคนละ จุดจ่ายได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
            Me.ListViewDIItem.SetFocus
            Exit Sub
            End If
      
            Me.ListViewDIItem.ListItems.Remove (vIndex)
            Call CalcDIItemAmount
            Call CalcDITotalAmount
            Call GenDIItemLine
         Else
            MsgBox "ไม่สามารถลบรายการนี้ได้ เนื่องจากถูกอ้างอิงทำคิวเรียบร้อยแล้ว", vbCritical, "Send Error Message"
            Me.TBDIBarCode.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "เลขที่เอกสารนี้ ได้ถูกอ้างไปทำการจัดคิวเรียบร้อยแล้วไม่สามารถลบได้", vbCritical, "Send Error Message"
         Me.TBDIBarCode.SetFocus
      End If
      End If
      
      
      Dim vPointZone As String
      Dim vItemPickZone As String
      
      If vDIIsOpen = 1 Then
      If vDIIsConfirm = 0 Then
         vPointZone = Me.LBLDI.Caption
         vIndex = Me.ListViewDIItem.SelectedItem.Index
         vItemPickZone = Me.ListViewDIItem.ListItems(vIndex).ListSubItems(14)
         
         If vPointZone <> vItemPickZone Then
            MsgBox "รายการสินค้าดังกล่าว ไม่สามารถลบได้ เนื่องจากสินค้าดังกล่าวได้สั่งจัดจากอีก จุดจ่าย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
            Me.TBDIBarCode.SetFocus
            Exit Sub
         End If
         
         vDocNo = Me.LBLDIDocNo.Caption
         vQuery = "exec dbo.USP_NP_SearchSendQueuePicking '" & vDocNo & "' "
          If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
              vDIIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
          End If
          vRecordset.Close
          
         If vDIIsConfirm = 0 Then
         
         If Me.ListViewDIItem.ListItems.Count <= 1 Then
            MsgBox "เอกสารที่เคยบันทึกไปแล้ว ต้องมีรายการสินค้าอย่างน้อย 1 รายการ ", vbCritical, "Send Error Message"
            Me.TBDIBarCode.SetFocus
            Exit Sub
         End If
         
            Me.ListViewDIItem.ListItems.Remove (vIndex)
            Call CalcDIItemAmount
            Call CalcDITotalAmount
            Call GenDIItemLine
         Else
            MsgBox "ไม่สามารถลบรายการนี้ได้ เนื่องจากถูกอ้างอิงทำคิวเรียบร้อยแล้ว", vbCritical, "Send Error Message"
            Me.TBDIBarCode.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "เลขที่เอกสารนี้ ได้ถูกอ้างไปทำการจัดคิวเรียบร้อยแล้วไม่สามารถลบได้", vbCritical, "Send Error Message"
         Me.TBDIBarCode.SetFocus
      End If
      End If
      
   End If
End If

If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If

If KeyCode = 38 Then
Dim n As Integer
Dim vListIndex As Integer

If Me.ListViewDIItem.ListItems.Count > 0 Then
   vListIndex = Me.ListViewDIItem.SelectedItem.Index
   If vListIndex = 1 Then
       Me.TBDIBarCode.SetFocus
   End If
End If

End If
End Sub

Private Sub ListViewDIItem_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vBarCode As String
Dim vIndex As Integer
Dim vQTY As Double
Dim vDiscountWord As String
Dim vDiscountAmount As Double
Dim vNetAmount As Double
Dim vPrice As Double
Dim i As Integer
Dim vListItem As ListItem
Dim vPickZone As String
Dim vLinePickZone As String

If KeyAscii = 13 Then
If Me.ListViewDIItem.ListItems.Count > 0 Then

vPickZone = Me.LBLDI.Caption
   If vDIIsConfirm = 0 Then
      vIndex = Me.ListViewDIItem.SelectedItem.Index
      vLinePickZone = Me.ListViewDIItem.ListItems(vIndex).ListSubItems(14).Text
      
      If vPickZone <> vLinePickZone Then
      MsgBox "ไม่สามารถแก้ไข รายการสินค้าที่ถูกจ่ายมาจากจุดจ่ายอื่นได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      Me.ListViewDIItem.SetFocus
      Exit Sub
      End If
      
      Me.PICDIKeyQty.Visible = True
      vBarCode = Me.ListViewDIItem.ListItems(vIndex).SubItems(1)
      vQuery = "exec dbo.USP_MB_SearchBarcode '" & vBarCode & "'"
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Me.LBLDIItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
          Me.LBLDIItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
          Me.LBLDIUnitCode.Caption = Trim(vRecordset.Fields("unitcode").Value)
          Me.LBLDIWHCode.Caption = Trim(vRecordset.Fields("defsalewhcode").Value)
          Me.LBLDIShelfCode.Caption = Trim(vRecordset.Fields("defsaleshelf").Value)
          Me.LBLDIShelfID.Caption = Trim(vRecordset.Fields("shelfid").Value)
          Me.LBLDIZoneID.Caption = Trim(vRecordset.Fields("zoneid").Value)
          Me.LBLDIBarCode.Caption = Trim(vRecordset.Fields("barcode").Value)
          
          vPrice = Trim(vRecordset.Fields("price").Value)
          If vPrice <= 0 Then
             MsgBox "สินค้ารายการนี้ ยังไม่ได้กำหนดราคาขายของหน่วยนับขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
             Me.TBDIKeyQty.Enabled = False
             Me.TBDIKeyDiscount.Enabled = False
             Exit Sub
          Else
             Me.TBDIKeyQty.Enabled = True
             Me.TBDIKeyDiscount.Enabled = False
          End If
          
          Me.LBLDIPrice.Caption = Format(vPrice, "##,##0.00")
          
          Me.ListViewDIItemStock.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             vQTY = Trim(vRecordset.Fields("stock").Value)
             
            Set vListItem = Me.ListViewDIItemStock.ListItems.Add(, , Trim(vRecordset.Fields("whcode").Value))
            vListItem.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
            vListItem.SubItems(2) = Format(vQTY, "##,##0.00")
            vListItem.SubItems(3) = Trim(vRecordset.Fields("stkunitcode").Value)
                
             vRecordset.MoveNext
          Next i
             
          vQTY = Me.ListViewDIItem.ListItems(vIndex).SubItems(3)
          vDiscountWord = Me.ListViewDIItem.ListItems(vIndex).SubItems(13)
          vDiscountAmount = Me.ListViewDIItem.ListItems(vIndex).SubItems(7)
          vNetAmount = Me.ListViewDIItem.ListItems(vIndex).SubItems(6)
          
          
          Me.TBDIKeyQty.Text = Format(vQTY, "##,##0.00")
          Me.TBDIKeyDiscount.Text = vDiscountWord
          Me.LBLDIDiscountWord.Caption = Format(vDiscountAmount, "##,##0.00")
          Me.LBLDIItemNetAmount.Caption = Format(vNetAmount, "##,##0.00")

          Me.TBDIKeyQty.SetFocus
       End If
    vRecordset.Close
   Else
      MsgBox "ไม่สามารถแก้ไขรายการของเอกสารนี้ได้ เนื่องจากเอกสารนี้ถูกอ้างไปจัดคิวเรียบร้อยแล้ว", vbCritical, "Send Error Message"
      Me.TBDIBarCode.SetFocus
   End If
End If
End If
End Sub

Private Sub ListViewDILastSendQue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDILastSendQue.Visible = False
Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub ListViewDISearchAR_DblClick()
Dim vIndex As Integer
Dim vARCode As String

If Me.ListViewDISearchAR.ListItems.Count > 0 Then
   vIndex = Me.ListViewDISearchAR.SelectedItem.Index
   vARCode = Me.ListViewDISearchAR.ListItems(vIndex).SubItems(1)
   Me.TBDIArCode.Text = vARCode
   Me.TBDICarLicense.SetFocus
   Me.PICDISearchAR.Visible = False
End If
End Sub

Private Sub ListViewDISearchAR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchAR.Visible = False
Me.TBDIArCode.SetFocus
End If

If KeyCode = 38 Then
Me.TBDISearchAR.SetFocus
End If
End Sub

Private Sub ListViewDISearchAR_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer
Dim vARCode As String

If KeyAscii = 13 Then
   If Me.ListViewDISearchAR.ListItems.Count > 0 Then
      vIndex = Me.ListViewDISearchAR.SelectedItem.Index
      vARCode = Me.ListViewDISearchAR.ListItems(vIndex).SubItems(1)
      Me.TBDIArCode.Text = vARCode
      Me.TBDICarLicense.SetFocus
      Me.PICDISearchAR.Visible = False
   End If
End If
End Sub

Private Sub ListViewDISearchDI_DblClick()
Dim vIndex As Integer
Dim vDocNo As String
Dim vRefNo As String

If Me.ListViewDISearchDI.ListItems.Count > 0 Then
   vIndex = Me.ListViewDISearchDI.SelectedItem.Index
   'vDocNo = Me.ListViewDISearchDI.ListItems(vIndex).SubItems(1)
   vRefNo = Me.ListViewDISearchDI.ListItems(vIndex).SubItems(4)
   Me.TBDICarLicense.Text = vRefNo
   'Me.LBLDIDocNo.Caption = vDocNo
   Me.TBDIBarCode.SetFocus
   Me.PICDISearchDI.Visible = False
End If
End Sub

Private Sub ListViewDISearchDI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchDI.Visible = False
Me.TBDIArCode.SetFocus
End If

If KeyCode = 38 Then
Me.TBDISearchDI.SetFocus
End If
End Sub

Private Sub ListViewDISearchDI_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer
Dim vDocNo As String
Dim vRefNo As String

If KeyAscii = 13 Then
If Me.ListViewDISearchDI.ListItems.Count > 0 Then
   vIndex = Me.ListViewDISearchDI.SelectedItem.Index
   'vDocNo = Me.ListViewDISearchDI.ListItems(vIndex).SubItems(1)
   vRefNo = Me.ListViewDISearchDI.ListItems(vIndex).SubItems(4)
   Me.TBDICarLicense.Text = vRefNo
   'Me.LBLDIDocNo.Caption = vDocNo
   Me.TBDIBarCode.SetFocus
   Me.PICDISearchDI.Visible = False
End If
End If
End Sub

Private Sub ListViewDISearchItem_DblClick()
Dim vIndex As Integer
Dim vItemCode As String

If Me.ListViewDISearchItem.ListItems.Count > 0 Then
   vIndex = Me.ListViewDISearchItem.SelectedItem.Index
   vItemCode = Me.ListViewDISearchItem.ListItems(vIndex).SubItems(1)
   Me.TBDIBarCode.Text = vItemCode
   Me.TBDIBarCode.SetFocus
   Me.PICDISearchItem.Visible = False
End If
End Sub

Private Sub ListViewDISearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchItem.Visible = False
Me.TBDIBarCode.SetFocus
End If

If KeyCode = 38 Then
Me.TBDISearchItem.SetFocus
End If

End Sub

Private Sub ListViewDISearchItem_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer
Dim vItemCode As String

If KeyAscii = 13 Then
   If Me.ListViewDISearchItem.ListItems.Count > 0 Then
      vIndex = Me.ListViewDISearchItem.SelectedItem.Index
      vItemCode = Me.ListViewDISearchItem.ListItems(vIndex).SubItems(1)
      Me.TBDIBarCode.Text = vItemCode
      Me.TBDIBarCode.SetFocus
      Me.PICDISearchItem.Visible = False
   End If
End If
End Sub

Private Sub ListViewDISearchSale_DblClick()
Dim vIndex As Integer
Dim vSaleCode As String

If Me.ListViewDISearchSale.ListItems.Count > 0 Then
   vIndex = Me.ListViewDISearchSale.SelectedItem.Index
   vSaleCode = Me.ListViewDISearchSale.ListItems(vIndex).SubItems(1)
   Me.TBDISaleCode.Text = vSaleCode
   Me.TBDIBarCode.SetFocus
   Me.PICDISearchSale.Visible = False
End If
End Sub

Private Sub ListViewDocNo_DblClick()
'Dim vIndex As Integer
'Dim vDocNo As String

'If Me.ListViewDocNo.ListItems.Count > 0 Then
 '  vIndex = Me.ListViewDocNo.SelectedItem.Index
  ' vDocNo = Trim(Me.ListViewDocNo.ListItems(vIndex).SubItems(1))
   'Me.TXTDocNo.Text = vDocNo
   'Me.TBBarCode.SetFocus
'End If
'Me.PICSearchDoc.Visible = False
End Sub

Private Sub ListViewDocNo_KeyPress(KeyAscii As Integer)
'Dim vIndex As Integer
'Dim vDocNo As String

'If KeyAscii = 13 Then
 '  If Me.ListViewDocNo.ListItems.Count > 0 Then
  '    vIndex = Me.ListViewDocNo.SelectedItem.Index
   '   vDocNo = Trim(Me.ListViewDocNo.ListItems(vIndex).SubItems(1))
    '  Me.TXTDocNo.Text = vDocNo
     ' Me.TBBarCode.SetFocus
   'End If
'End If
'Me.PICSearchDoc.Visible = False
End Sub

Private Sub ListViewDISearchSale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchSale.Visible = False
Me.TBDISaleCode.SetFocus
End If

If KeyCode = 38 Then
   Me.TBDISearchSale.SetFocus
End If
End Sub

Private Sub ListViewDISearchSale_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer
Dim vSaleCode As String

If KeyAscii = 13 Then
   If Me.ListViewDISearchSale.ListItems.Count > 0 Then
      vIndex = Me.ListViewDISearchSale.SelectedItem.Index
      vSaleCode = Me.ListViewDISearchSale.ListItems(vIndex).SubItems(1)
      Me.TBDISaleCode.Text = vSaleCode
      Me.TBDIBarCode.SetFocus
      Me.PICDISearchSale.Visible = False
   End If
End If
End Sub

Private Sub ListViewInfQue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSendInformation.Visible = False
Me.TXTDocNo.SetFocus
End If
End Sub

Private Sub ListViewItem_DblClick()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vBarCode As String
Dim vIndex As Integer
Dim vQTY As Double
Dim vDiscountWord As String
Dim vDiscountAmount As Double
Dim vNetAmount As Double
Dim vPrice As Double
Dim i As Integer
Dim vListItem As ListItem


If Me.ListViewItem.ListItems.Count > 0 Then
   If vSendQue = 0 Then
      vIndex = Me.ListViewItem.SelectedItem.Index
      Me.PICSearchItem.Visible = True
      vBarCode = Me.ListViewItem.ListItems(vIndex).SubItems(1)
      vQuery = "exec dbo.USP_MB_SearchBarcode '" & vBarCode & "'"
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Me.LBLItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
          Me.LBLItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
          Me.LBLUnitCode.Caption = Trim(vRecordset.Fields("unitcode").Value)
          Me.LBLWHCode.Caption = Trim(vRecordset.Fields("defsalewhcode").Value)
          Me.LBLShelfCode.Caption = Trim(vRecordset.Fields("defsaleshelf").Value)
          Me.LBLShelfID.Caption = Trim(vRecordset.Fields("shelfid").Value)
          Me.LBLZoneID.Caption = Trim(vRecordset.Fields("zoneid").Value)
          Me.LBLBarCode.Caption = Trim(vRecordset.Fields("barcode").Value)
          
          vPrice = Trim(vRecordset.Fields("price").Value)
          If vPrice <= 0 Then
             MsgBox "สินค้ารายการนี้ ยังไม่ได้กำหนดราคาขายของหน่วยนับขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
             Me.TBQty.Enabled = False
             Me.TXTDisCount.Enabled = False
             Me.TBBarCode.SetFocus
             Exit Sub
          Else
             Me.TBQty.Enabled = True
             Me.TXTDisCount.Enabled = True
          End If
          
          Me.LBLPrice.Caption = Format(vPrice, "##,##0.00")
          
          Me.ListViewStock.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             vQTY = Trim(vRecordset.Fields("stock").Value)
             
            Set vListItem = Me.ListViewStock.ListItems.Add(, , Trim(vRecordset.Fields("whcode").Value))
            vListItem.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
            vListItem.SubItems(2) = Format(vQTY, "##,##0.00")
            vListItem.SubItems(3) = Trim(vRecordset.Fields("stkunitcode").Value)
                
             vRecordset.MoveNext
          Next i
             
          vQTY = Me.ListViewItem.ListItems(vIndex).SubItems(3)
          vDiscountWord = Me.ListViewItem.ListItems(vIndex).SubItems(13)
          vDiscountAmount = Me.ListViewItem.ListItems(vIndex).SubItems(7)
          vNetAmount = Me.ListViewItem.ListItems(vIndex).SubItems(6)
          
          
          Me.TBQty.Text = Format(vQTY, "##,##0.00")
          Me.TXTDisCount.Text = vDiscountWord
          Me.LBLDisCountAmount.Caption = Format(vDiscountAmount, "##,##0.00")
          Me.LBLNetPrice.Caption = Format(vNetAmount, "##,##0.00")

          Me.TBQty.SetFocus
       End If
    vRecordset.Close
   Else
      MsgBox "ไม่สามารถแก้ไขรายการของเอกสารนี้ได้ เนื่องจากเอกสารนี้ถูกอ้างไปจัดคิวเรียบร้อยแล้ว", vbCritical, "Send Error Message"
      Me.TBBarCode.SetFocus
   End If
End If
End Sub

Private Sub ListViewItem_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer
Dim vDocNo As String
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSendQue As Integer

If Me.ListViewItem.ListItems.Count > 0 Then
   If KeyCode = 46 Then
   
      If vSendQue = 0 Then
         vDocNo = Me.TXTDocNo.Text
         vQuery = "exec dbo.USP_NP_SearchSendQueuePicking '" & vDocNo & "' "
          If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
              vSendQue = Trim(vRecordset.Fields("issendque").Value)
          End If
          vRecordset.Close
          
         If vSendQue = 0 Then
            vIndex = Me.ListViewItem.SelectedItem.Index
            Me.ListViewItem.ListItems.Remove (vIndex)
            Call CalcPRItemAmount
            Call CalcTotalAmount
            Call GenPRItemLine
         Else
            MsgBox "ไม่สามารถลบรายการนี้ได้ เนื่องจากถูกอ้างอิงทำคิวเรียบร้อยแล้ว", vbCritical, "Send Error Message"
            Me.TBBarCode.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "เลขที่เอกสารนี้ ได้ถูกอ้างไปทำการจัดคิวเรียบร้อยแล้วไม่สามารถลบได้", vbCritical, "Send Error Message"
         Me.TBBarCode.SetFocus
      End If
   End If
End If

If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If

If KeyCode = 38 Then
   Me.TBBarCode.SetFocus
End If
End Sub

Private Sub ListViewSaleCode_DblClick()
'Dim vIndex As Integer
'Dim vSaleCode As String

'If Me.ListViewSaleCode.ListItems.Count > 0 Then
'vIndex = Me.ListViewSaleCode.SelectedItem.Index
'vSaleCode = Trim(Me.ListViewSaleCode.ListItems(vIndex).SubItems(1) & "/" & Me.ListViewSaleCode.ListItems(vIndex).SubItems(2))
'Me.TXTSaleCode.Text = vSaleCode
'Me.TXTSaleCode.SetFocus
'End If
'Me.PICSaleCode.Visible = False
End Sub

Private Sub ListViewSaleCode_KeyPress(KeyAscii As Integer)
'Dim vIndex As Integer
'Dim vSaleCode As String

'If KeyAscii = 13 Then
 '  If Me.ListViewSaleCode.ListItems.Count > 0 Then
  ' vIndex = Me.ListViewSaleCode.SelectedItem.Index
   'vSaleCode = Trim(Me.ListViewSaleCode.ListItems(vIndex).SubItems(1) & "/" & Me.ListViewSaleCode.ListItems(vIndex).SubItems(2))
   'Me.TXTSaleCode.Text = vSaleCode
   'Me.TXTSaleCode.SetFocus
   'End If
'End If
'Me.PICSaleCode.Visible = False
End Sub

Private Sub ListViewItem_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vBarCode As String
Dim vIndex As Integer
Dim vQTY As Double
Dim vDiscountWord As String
Dim vDiscountAmount As Double
Dim vNetAmount As Double
Dim vPrice As Double
Dim i As Integer
Dim vListItem As ListItem

If KeyAscii = 13 Then
If Me.ListViewItem.ListItems.Count > 0 Then
   If vSendQue = 0 Then
      vIndex = Me.ListViewItem.SelectedItem.Index
      Me.PICSearchItem.Visible = True
      vBarCode = Me.ListViewItem.ListItems(vIndex).SubItems(1)
      vQuery = "exec dbo.USP_MB_SearchBarcode '" & vBarCode & "'"
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Me.LBLItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
          Me.LBLItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
          Me.LBLUnitCode.Caption = Trim(vRecordset.Fields("unitcode").Value)
          Me.LBLWHCode.Caption = Trim(vRecordset.Fields("defsalewhcode").Value)
          Me.LBLShelfCode.Caption = Trim(vRecordset.Fields("defsaleshelf").Value)
          Me.LBLShelfID.Caption = Trim(vRecordset.Fields("shelfid").Value)
          Me.LBLZoneID.Caption = Trim(vRecordset.Fields("zoneid").Value)
          Me.LBLBarCode.Caption = Trim(vRecordset.Fields("barcode").Value)
          
          vPrice = Trim(vRecordset.Fields("price").Value)
          If vPrice <= 0 Then
             MsgBox "สินค้ารายการนี้ ยังไม่ได้กำหนดราคาขายของหน่วยนับขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
             Me.TBQty.Enabled = False
             Me.TXTDisCount.Enabled = False
             Me.TBBarCode.SetFocus
             Exit Sub
          Else
             Me.TBQty.Enabled = True
             Me.TXTDisCount.Enabled = True
          End If
          
          Me.LBLPrice.Caption = Format(vPrice, "##,##0.00")
          
          Me.ListViewStock.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             vQTY = Trim(vRecordset.Fields("stock").Value)
             
            Set vListItem = Me.ListViewStock.ListItems.Add(, , Trim(vRecordset.Fields("whcode").Value))
            vListItem.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
            vListItem.SubItems(2) = Format(vQTY, "##,##0.00")
            vListItem.SubItems(3) = Trim(vRecordset.Fields("stkunitcode").Value)
                
             vRecordset.MoveNext
          Next i
             
          vQTY = Me.ListViewItem.ListItems(vIndex).SubItems(3)
          vDiscountWord = Me.ListViewItem.ListItems(vIndex).SubItems(13)
          vDiscountAmount = Me.ListViewItem.ListItems(vIndex).SubItems(7)
          vNetAmount = Me.ListViewItem.ListItems(vIndex).SubItems(6)
          
          
          Me.TBQty.Text = Format(vQTY, "##,##0.00")
          Me.TXTDisCount.Text = vDiscountWord
          Me.LBLDisCountAmount.Caption = Format(vDiscountAmount, "##,##0.00")
          Me.LBLNetPrice.Caption = Format(vNetAmount, "##,##0.00")

          Me.TBQty.SetFocus
       End If
    vRecordset.Close
   Else
      MsgBox "ไม่สามารถแก้ไขรายการของเอกสารนี้ได้ เนื่องจากเอกสารนี้ถูกอ้างไปจัดคิวเรียบร้อยแล้ว", vbCritical, "Send Error Message"
      Me.TBBarCode.SetFocus
   End If
End If
End If
End Sub

Private Sub ListViewLastSendQue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICLastSendQue.Visible = False
'Me.TBDocNo.SetFocus
End If
End Sub

Private Sub ListViewPRSearchAR_DblClick()
Dim vIndex As Integer
Dim vARCode As String

If Me.ListViewPRSearchAR.ListItems.Count > 0 Then
   vIndex = Me.ListViewPRSearchAR.SelectedItem.Index
   vARCode = Trim(Me.ListViewPRSearchAR.ListItems(vIndex).SubItems(1))
   Me.TXTArCode.Text = vARCode
   Me.LBLArName.Caption = Trim(Me.ListViewPRSearchAR.ListItems(vIndex).SubItems(2))
   Me.TXTMember.Caption = Trim(Me.ListViewPRSearchAR.ListItems(vIndex).SubItems(3))
   Me.TXTArCode.SetFocus
End If
Me.PICPRSearchAR.Visible = False
End Sub

Private Sub ListViewPRSearchAR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchAR.Visible = False
Me.TXTArCode.SetFocus
End If

If KeyCode = 38 Then
Me.TBPRSearchAR.SetFocus
End If
End Sub

Private Sub ListViewPRSearchAR_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer
Dim vARCode As String

If KeyAscii = 13 Then
If Me.ListViewPRSearchAR.ListItems.Count > 0 Then
   vIndex = Me.ListViewPRSearchAR.SelectedItem.Index
   vARCode = Trim(Me.ListViewPRSearchAR.ListItems(vIndex).SubItems(1))
   Me.TXTArCode.Text = vARCode
   Me.LBLArName.Caption = Trim(Me.ListViewPRSearchAR.ListItems(vIndex).SubItems(2))
   Me.TXTMember.Caption = Trim(Me.ListViewPRSearchAR.ListItems(vIndex).SubItems(3))
   Me.TXTArCode.SetFocus
End If
Me.PICPRSearchAR.Visible = False
End If
End Sub

Private Sub ListViewPRSearchDocNo_DblClick()
Dim vIndex As Integer
Dim vDocNo As String

If Me.ListViewPRSearchDocNo.ListItems.Count > 0 Then
   vIndex = Me.ListViewPRSearchDocNo.SelectedItem.Index
   vDocNo = Trim(Me.ListViewPRSearchDocNo.ListItems(vIndex).SubItems(1))
   Me.TXTDocNo.Text = vDocNo
   Me.TBBarCode.SetFocus
End If
Me.PICPRSearchDocNo.Visible = False
End Sub

Private Sub ListViewPRSearchDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchDocNo.Visible = False
Me.TXTDocNo.SetFocus
End If
If KeyCode = 38 Then
Me.TBPRSearchDocNo.SetFocus
End If
End Sub

Private Sub ListViewPRSearchDocNo_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer
Dim vDocNo As String

If KeyAscii = 13 Then
If Me.ListViewPRSearchDocNo.ListItems.Count > 0 Then
   vIndex = Me.ListViewPRSearchDocNo.SelectedItem.Index
   vDocNo = Trim(Me.ListViewPRSearchDocNo.ListItems(vIndex).SubItems(1))
   Me.TXTDocNo.Text = vDocNo
   Me.TBBarCode.SetFocus
End If
Me.PICPRSearchDocNo.Visible = False
End If
End Sub

Private Sub ListViewPRSearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchItem.Visible = False
Me.TBBarCode.SetFocus
End If

If KeyCode = 38 Then
Me.TBPRSearchItem.SetFocus
End If
End Sub

Private Sub ListViewPRSearchItem_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer
Dim vItemCode As String

If KeyAscii = 13 Then
   If Me.ListViewPRSearchItem.ListItems.Count > 0 Then
      vIndex = Me.ListViewPRSearchItem.SelectedItem.Index
      vItemCode = Me.ListViewPRSearchItem.ListItems(vIndex).SubItems(1)
      Me.TBBarCode.Text = vItemCode
      Me.TBBarCode.SetFocus
      Me.PICPRSearchItem.Visible = False
   End If
End If
End Sub

Private Sub ListViewPRSearchSale_DblClick()
Dim vIndex As Integer
Dim vSaleCode As String

If Me.ListViewPRSearchSale.ListItems.Count > 0 Then
vIndex = Me.ListViewPRSearchSale.SelectedItem.Index
vSaleCode = Trim(Me.ListViewPRSearchSale.ListItems(vIndex).SubItems(1) & "/" & Me.ListViewPRSearchSale.ListItems(vIndex).SubItems(2))
Me.TXTSaleCode.Text = vSaleCode
Me.TXTSaleCode.SetFocus
End If
Me.PICPRSearchSale.Visible = False
End Sub

Private Sub ListViewPRSearchSale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchSale.Visible = False
Me.TXTSaleCode.SetFocus
End If

If KeyCode = 38 Then
Me.TBPRSearchSale.SetFocus
End If
End Sub

Private Sub ListViewPRSearchSale_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer
Dim vSaleCode As String

If KeyAscii = 13 Then
If Me.ListViewPRSearchSale.ListItems.Count > 0 Then
vIndex = Me.ListViewPRSearchSale.SelectedItem.Index
vSaleCode = Trim(Me.ListViewPRSearchSale.ListItems(vIndex).SubItems(1) & "/" & Me.ListViewPRSearchSale.ListItems(vIndex).SubItems(2))
Me.TXTSaleCode.Text = vSaleCode
Me.TXTSaleCode.SetFocus
End If
Me.PICPRSearchSale.Visible = False
End If
End Sub

Private Sub ListViewSaleOrder_DblClick()
'Dim vIndex As Integer
'Dim vItemCode As String
'Dim vItemName As String
'Dim vQTY As Double
'Dim vDiscountAmount As Double
'Dim vNetAmount As Double
'Dim vPrice As Double
'Dim vRemainQty As Double

'If Me.ListViewSaleOrder.ListItems.Count > 0 Then
 '  vIndex = Me.ListViewSaleOrder.SelectedItem.Index
  ' Me.PICEditOrder.Visible = True
   
   'Me.LBLEditIndex.Caption = vIndex
   
   'Me.LBLEditItemCode.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(1)
   'Me.LBLEditItemName.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(2)
   'Me.LBLEditUnitCode.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(4)
   
   'vQTY = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(3)
   'vPrice = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(5)
   'vDiscountAmount = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(6)
   'vNetAmount = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(7)
   'vRemainQty = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(14)
   
   'Me.LBLEditItemQty.Caption = Format(vQTY, "##,##0.00")
   'Me.LBLEditPrice.Caption = Format(vPrice, "##,##0.00")
   'Me.LBLEditDiscount.Caption = Format(vDiscountAmount, "##,##0.00")
   'Me.LBLEditItemAmount.Caption = Format(vNetAmount, "##,##0.00")
   'Me.LBLEditRemain.Caption = Format(vRemainQty, "##,##0.00")
   
   'Me.PICEditOrder.Visible = True
   'Me.TBEditQty.Text = vQTY
   'Me.TBEditQty.SetFocus

'End If
End Sub


Private Sub ListViewSaleOrder_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim vIndex As Integer

'If KeyCode = 46 Then
 '  If Me.ListViewSaleOrder.ListItems.Count > 0 And Me.ListViewSaleOrder.ListItems.Count <> 1 Then
  '    vIndex = Me.ListViewSaleOrder.SelectedItem.Index
   '   Me.ListViewSaleOrder.ListItems.Remove (vIndex)
    '  Call CalcEditItemQty
   'End If
   'If Me.ListViewSaleOrder.ListItems.Count = 1 Then
    '  MsgBox "ไม่สามารถลบรายการสินค้าทั้งหมดได้ กรณีไม่เอา ก็ให้ปิดหน้าจอเท่านั้น", vbCritical, "Send Error Message"
   'End If
'End If

'If KeyCode = 116 Then
'Call CMDSaleOrderSendQue_Click
'End If
End Sub

Private Sub ListViewSearch_DblClick()
'Dim vIndex As Integer
'Dim vBarCode As String

'If Me.ListViewSearch.ListItems.Count > 0 Then
'vIndex = Me.ListViewSearch.SelectedItem.Index
'vBarCode = Me.ListViewSearch.ListItems(vIndex).SubItems(1)
'Me.TBBarCode.Text = vBarCode
'Me.TBBarCode.SetFocus
'End If
'Me.PICSearch.Visible = False
End Sub

Private Sub ListViewSearch_KeyPress(KeyAscii As Integer)
'Dim vIndex As Integer
'Dim vBarCode As String

'If KeyAscii = 13 Then
 '  If Me.ListViewSearch.ListItems.Count > 0 Then
  ' vIndex = Me.ListViewSearch.SelectedItem.Index
   'vBarCode = Me.ListViewSearch.ListItems(vIndex).SubItems(1)
   'Me.TBBarCode.Text = vBarCode
   'Me.TBBarCode.SetFocus
   'End If
'End If
'Me.PICSearch.Visible = False
End Sub

Private Sub SearchDocNoClose_Click()
'Me.PICSearchDoc.Visible = False
End Sub

Private Sub ListViewPRSearchItem_DblClick()
Dim vIndex As Integer
Dim vItemCode As String

If Me.ListViewPRSearchItem.ListItems.Count > 0 Then
   vIndex = Me.ListViewPRSearchItem.SelectedItem.Index
   vItemCode = Me.ListViewPRSearchItem.ListItems(vIndex).SubItems(1)
   Me.TBBarCode.Text = vItemCode
   Me.TBBarCode.SetFocus
   Me.PICPRSearchItem.Visible = False
End If
End Sub

Private Sub ListViewSaleOrder_KeyPress(KeyAscii As Integer)
'Dim vIndex As Integer
'Dim vItemCode As String
'Dim vItemName As String
'Dim vQTY As Double
'Dim vDiscountAmount As Double
'Dim vNetAmount As Double
'Dim vPrice As Double
'Dim vRemainQty As Double

'If KeyAscii = 13 Then
 '   If Me.ListViewSaleOrder.ListItems.Count > 0 Then
  '     vIndex = Me.ListViewSaleOrder.SelectedItem.Index
   '    Me.PICEditOrder.Visible = True
    '
     '  Me.LBLEditIndex.Caption = vIndex
       
      ' Me.LBLEditItemCode.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(1)
       'Me.LBLEditItemName.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(2)
       'Me.LBLEditUnitCode.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(4)
       
       'vQTY = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(3)
       'vPrice = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(5)
       'vDiscountAmount = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(6)
       'vNetAmount = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(7)
       'vRemainQty = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(14)
       
       'Me.LBLEditItemQty.Caption = Format(vQTY, "##,##0.00")
       'Me.LBLEditPrice.Caption = Format(vPrice, "##,##0.00")
       'Me.LBLEditDiscount.Caption = Format(vDiscountAmount, "##,##0.00")
       'Me.LBLEditItemAmount.Caption = Format(vNetAmount, "##,##0.00")
       'Me.LBLEditRemain.Caption = Format(vRemainQty, "##,##0.00")
       
       'Me.PICEditOrder.Visible = True
       'Me.TBEditQty.Text = vQTY
       'Me.TBEditQty.SetFocus
    
    'End If
'End If
End Sub

Private Sub PICArKeyData_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICArKeyData.Visible = False
   Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub PICDIKeyQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICDIKeyQty.Visible = False
   Me.TBDIBarCode.Text = ""
   Me.TBDIBarCode.SetFocus
End If
End Sub

Private Sub PICDILastSendQue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDILastSendQue.Visible = False
Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub PICDISearchAR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchAR.Visible = False
Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub PICDISearchDI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchDI.Visible = False
Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub PICDISearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchItem.Visible = False
Me.TBDIBarCode.SetFocus
End If
End Sub

Private Sub PICDISearchSale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchSale.Visible = False
Me.TBDISaleCode.SetFocus
End If
End Sub

Private Sub PICDISendInformation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.LBLDIInfDocNo.Caption = ""
   Me.ListViewDIInfQue.ListItems.Clear
   Me.PICDISendInformation.Visible = False
   Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub PICDriveIn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub PICLastSendQue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICLastSendQue.Visible = False
'Me.TBDocNo.SetFocus
End If
End Sub

Private Sub PICOrder_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 116 Then
'Call CMDSaleOrderSendQue_Click
'End If
End Sub

Private Sub PICPickReq_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If
End Sub

Private Sub PICPRArKeyData_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICPRArKeyData.Visible = False
   Me.TXTArCode.SetFocus
End If
End Sub

Private Sub PICPRSearchAR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchAR.Visible = False
Me.TXTArCode.SetFocus
End If
End Sub

Private Sub PICPRSearchDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchDocNo.Visible = False
Me.TXTDocNo.SetFocus
End If
End Sub

Private Sub PICPRSearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchItem.Visible = False
Me.TBBarCode.SetFocus
End If
End Sub

Private Sub PICPRSearchSale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchSale.Visible = False
Me.TXTSaleCode.SetFocus
End If
End Sub

Private Sub PICSearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSearchItem.Visible = False
Me.TBBarCode.Text = ""
Me.TBBarCode.SetFocus
End If
End Sub

Private Sub PICSelectDI_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 97 Then
'Me.PICDriveIn.Visible = True
'Me.LBLDI.Caption = "01"
'Me.TBDIArCode.SetFocus
'End If

If KeyCode = 98 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "02"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 99 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "03"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 100 Then
Me.PICDriveIn.Visible = True
Me.LBLDI.Caption = "04"
Me.TBDIArCode.SetFocus
End If

If KeyCode = 27 Then
Call CMDDIMain_Click
End If
End Sub

Private Sub PICSelectJob_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 97 Then
   Me.PICSelectDI.Visible = True
   Me.PICPickReq.Visible = False
   Me.PICSelectJob.Visible = False
   Me.LBLJob.Visible = False
   Me.LBLSelectDI.Visible = True
   Me.CMDDI02.SetFocus
End If

If KeyCode = 98 Then
   Me.PICSelectDI.Visible = False
   Me.PICPickReq.Visible = True
   Me.PICSelectJob.Visible = False
   Me.TXTDocNo.SetFocus
End If

End Sub

Private Sub PICSendInformation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSendInformation.Visible = False
Me.TXTDocNo.SetFocus
End If
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
Call CMDSaleOrderSendQue_Click
End If
End Sub

Private Sub Picture3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub Picture4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If
End Sub

Private Sub Picture5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If
End Sub

Private Sub TBBarCode_Change()
If Me.TBBarCode.Text <> "" Then
   Me.PICSearchItem.Visible = True
   Me.CMDClearScreen.Enabled = False
   Me.CMDSave.Enabled = False
   Me.CMDSearch.Enabled = False
   Me.CMDCancel.Enabled = False
   Me.CMDQue.Enabled = False
   
   Me.LBLItemCode.Caption = ""
   Me.LBLItemName.Caption = ""
   Me.LBLUnitCode.Caption = ""
   Me.TBQty.Text = ""
   Me.TXTDisCount.Text = ""
   Me.LBLPrice.Caption = ""
   Me.LBLDisCountAmount.Caption = ""
   Me.LBLWHCode.Caption = ""
   Me.LBLShelfCode.Caption = ""
   Me.LBLZoneID.Caption = ""
   Me.LBLShelfID.Caption = ""
   Me.LBLBarCode.Caption = ""
   Me.LBLNetPrice.Caption = ""
   Me.ListViewStock.ListItems.Clear
            
Else
   Me.PICSearchItem.Visible = False
   Me.CMDClearScreen.Enabled = True
   Me.CMDSave.Enabled = True
   Me.CMDSearch.Enabled = True
   Me.CMDCancel.Enabled = True
   Me.CMDQue.Enabled = True
End If
End Sub

Private Sub TBBarCode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If

If KeyCode = 40 Then
   Me.ListViewItem.SetFocus
End If

If KeyCode = 38 Then
   Me.TXTSaleCode.SetFocus
End If
End Sub

Private Sub TBBarCode_KeyPress(KeyAscii As Integer)
Dim vBarCode As String
Dim vPrice As Double
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vQTY As Double

Dim n As Integer
Dim vOldQty As Double
Dim vCheckItemCode  As String
Dim vItemCode As String
Dim vOldDiscount As Double

If KeyAscii = 13 Then
   If vSendQue = 1 Then
      MsgBox "เอกสารเลขที่นี้ ได้ถูกส่งไปจัดสินค้าเรียบร้อยแล้วไม่สามารถแก้ไขหรือเพิ่มรายการได้ ", vbCritical, "Send Error Message"
      Me.TBBarCode.Text = ""
      Me.TBBarCode.SetFocus
      Exit Sub
   End If

   If Me.TBBarCode.Text <> "" Then
      Me.PICSearchItem.Visible = True
      vBarCode = Me.TBBarCode.Text
      vQuery = "exec dbo.USP_MB_SearchBarcode '" & vBarCode & "'"
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          vItemCode = Trim(vRecordset.Fields("itemcode").Value)
          Me.LBLItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
          Me.LBLItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
          Me.LBLUnitCode.Caption = Trim(vRecordset.Fields("unitcode").Value)
          Me.LBLWHCode.Caption = Trim(vRecordset.Fields("defsalewhcode").Value)
          Me.LBLShelfCode.Caption = Trim(vRecordset.Fields("defsaleshelf").Value)
          Me.LBLShelfID.Caption = Trim(vRecordset.Fields("shelfid").Value)
          Me.LBLZoneID.Caption = Trim(vRecordset.Fields("zoneid").Value)
          Me.LBLBarCode.Caption = Trim(vRecordset.Fields("barcode").Value)
          
          vPrice = Trim(vRecordset.Fields("price").Value)
          If vPrice <= 0 Then
             MsgBox "สินค้ารายการนี้ ยังไม่ได้กำหนดราคาขายของหน่วยนับขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
             Me.TBQty.Enabled = False
             Me.TXTDisCount.Enabled = False
             Me.TBBarCode.SetFocus
             Exit Sub
          Else
             Me.TBQty.Enabled = True
             Me.TXTDisCount.Enabled = True
          End If
          
          Me.LBLPrice.Caption = Format(vPrice, "##,##0.00")
          
          Me.ListViewStock.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             vQTY = Trim(vRecordset.Fields("stock").Value)
             
            Set vListItem = Me.ListViewStock.ListItems.Add(, , Trim(vRecordset.Fields("whcode").Value))
            vListItem.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
            vListItem.SubItems(2) = Format(vQTY, "##,##0.00")
            vListItem.SubItems(3) = Trim(vRecordset.Fields("stkunitcode").Value)
             vRecordset.MoveNext
          Next i
          
         For n = 1 To Me.ListViewItem.ListItems.Count
            vCheckItemCode = Me.ListViewItem.ListItems(n).SubItems(1)
            If vItemCode = vCheckItemCode Then
               vOldQty = Me.ListViewItem.ListItems(n).SubItems(3)
               If Me.ListViewItem.ListItems(n).SubItems(7) <> "" Then
                  vOldDiscount = Me.ListViewItem.ListItems(n).SubItems(7)
               End If
            End If
         Next n
          If vOldQty = 0 Then
             Me.TBQty.Text = 1
          Else
             Me.TBQty.Text = vOldQty
          End If
          
          If vOldDiscount <> 0 Then
             Me.TXTDisCount.Text = vOldDiscount
          End If
          
          Me.TBQty.SetFocus
       Else
          Me.TBBarCode.SetFocus
       End If
    vRecordset.Close
   End If
End If

End Sub
Public Sub GenDIItemLine()
Dim i As Integer

For i = 1 To Me.ListViewDIItem.ListItems.Count
Me.ListViewDIItem.ListItems(i).Text = i
Next i
End Sub

Public Sub GenPRItemLine()
Dim i As Integer

For i = 1 To Me.ListViewItem.ListItems.Count
Me.ListViewItem.ListItems(i).Text = i
Next i
End Sub

Public Sub CalcDIItemAmount()
Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vNetprice As Double

If Me.TBDIKeyQty.Text <> "" And Me.LBLDIPrice.Caption <> "" And Me.LBLDIDiscountWord.Caption <> "" And Me.LBLDIDiscountWord.Caption <> "." Then
   vQTY = Me.TBDIKeyQty.Text
   vPrice = Me.LBLDIPrice.Caption
   If Me.LBLDIDiscountWord.Caption <> "" Then
      vDiscountAmount = Me.LBLDIDiscountWord.Caption
   Else
      vDiscountAmount = 0
   End If
   
   vNetprice = vQTY * (vPrice - vDiscountAmount)
   Me.LBLDINetAmount.Caption = Format(vNetprice, "##,##0.00")
End If
End Sub

Public Sub CalcPRItemAmount()
Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vNetprice As Double

If Me.TBQty.Text <> "" And Me.LBLPrice.Caption <> "" And Me.TXTDisCount.Text <> "" And Me.TXTDisCount.Text <> "." Then
   vQTY = Me.TBQty.Text
   vPrice = Me.LBLPrice.Caption
   If Me.TXTDisCount.Text <> "" Then
      vDiscountAmount = Me.TXTDisCount.Text
   Else
      vDiscountAmount = 0
   End If
   
   vNetprice = vQTY * (vPrice - vDiscountAmount)
   Me.LBLNetPrice.Caption = Format(vNetprice, "##,##0.00")
End If
End Sub

Private Sub TBDIArCode_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String

If Me.TBDIArCode.Text <> "" Then
   vSearchAR = Me.TBDIArCode.Text
   vQuery = "exec dbo.USP_AR_SearchAR '" & vSearchAR & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.LBLDIArName.Caption = Trim(vRecordset.Fields("arname").Value)
      Me.LBLDIMember.Caption = Trim(vRecordset.Fields("memberid").Value)
      Me.TBDICarLicense.SetFocus
   Else
      Me.LBLDIArName.Caption = ""
      Me.LBLDIMember.Caption = ""
      Me.TBDIArCode.SetFocus
   End If
   vRecordset.Close
Else
      Me.LBLDIArName.Caption = ""
      Me.LBLDIMember.Caption = ""
      Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub TBDIArCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If

If KeyCode = 40 Then
   Me.TBDICarLicense.SetFocus
End If
End Sub

Private Sub TBDIArCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.TBDICarLicense.SetFocus
End If
End Sub

Private Sub TBDIBarCode_Change()
If Me.TBDIBarCode.Text <> "" And vDISendQue = 0 Then
   Me.PICDIKeyQty.Visible = True

   Me.LBLDIItemCode.Caption = ""
   Me.LBLDIItemName.Caption = ""
   Me.LBLDIUnitCode.Caption = ""
   Me.TBDIKeyQty.Text = ""
   Me.TBDIKeyDiscount.Text = ""
   Me.LBLDIPrice.Caption = ""
   Me.LBLDIDiscountWord.Caption = ""
   Me.LBLDIWHCode.Caption = ""
   Me.LBLDIShelfCode.Caption = ""
   Me.LBLDIZoneID.Caption = ""
   Me.LBLDIShelfID.Caption = ""
   Me.LBLDIBarCode.Caption = ""
   Me.LBLDIItemNetAmount.Caption = ""
   Me.ListViewDIItemStock.ListItems.Clear
            
Else
   Me.PICDIKeyQty.Visible = False
End If
End Sub

Private Sub TBDIBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If

If KeyCode = 38 Then
   Me.TBDISaleCode.SetFocus
End If

If KeyCode = 40 Then
   Me.ListViewDIItem.SetFocus
End If
End Sub

Private Sub TBDIBarCode_KeyPress(KeyAscii As Integer)
Dim vBarCode As String
Dim vPrice As Double
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vQTY As Double
Dim n As Integer
Dim vCheckItemCode As String
Dim vItemCode As String
Dim vOldQty As Double
Dim vOldDiscount As Double

If KeyAscii = 13 Then
   If Me.TBDIBarCode.Text <> "" Then
      Me.PICDIKeyQty.Visible = True
      vBarCode = Me.TBDIBarCode.Text
      vQuery = "exec dbo.USP_MB_SearchBarcode '" & vBarCode & "'"
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          vItemCode = Trim(vRecordset.Fields("itemcode").Value)
          Me.LBLDIItemCode.Caption = Trim(vRecordset.Fields("itemcode").Value)
          Me.LBLDIItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
          Me.LBLDIUnitCode.Caption = Trim(vRecordset.Fields("unitcode").Value)
          Me.LBLDIWHCode.Caption = Trim(vRecordset.Fields("defsalewhcode").Value)
          Me.LBLDIShelfCode.Caption = Trim(vRecordset.Fields("defsaleshelf").Value)
          Me.LBLDIShelfID.Caption = Trim(vRecordset.Fields("shelfid").Value)
          Me.LBLDIZoneID.Caption = Trim(vRecordset.Fields("zoneid").Value)
          Me.LBLDIBarCode.Caption = Trim(vRecordset.Fields("barcode").Value)
          
          vPrice = Trim(vRecordset.Fields("price").Value)
          If vPrice <= 0 Then
             MsgBox "สินค้ารายการนี้ ยังไม่ได้กำหนดราคาขายของหน่วยนับขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
             Me.TBDIKeyQty.Enabled = False
             Me.TBDIKeyDiscount.Enabled = False
             Me.TBDIBarCode.SetFocus
             Exit Sub
          Else
             Me.TBDIKeyQty.Enabled = True
             Me.TBDIKeyDiscount.Enabled = False
          End If
          
          Me.LBLDIPrice.Caption = Format(vPrice, "##,##0.00")
          
          Me.ListViewDIItemStock.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             vQTY = Trim(vRecordset.Fields("stock").Value)
            
            Set vListItem = Me.ListViewDIItemStock.ListItems.Add(, , Trim(vRecordset.Fields("whcode").Value))
            vListItem.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
            vListItem.SubItems(2) = Format(vQTY, "##,##0.00")
            vListItem.SubItems(3) = Trim(vRecordset.Fields("stkunitcode").Value)
             vRecordset.MoveNext
          Next i
          
          
         For n = 1 To Me.ListViewDIItem.ListItems.Count
            vCheckItemCode = Me.ListViewDIItem.ListItems(n).SubItems(1)
            If vItemCode = vCheckItemCode Then
               vOldQty = Me.ListViewDIItem.ListItems(n).SubItems(3)
               If Me.ListViewDIItem.ListItems(n).SubItems(7) <> "" Then
               vOldDiscount = Me.ListViewDIItem.ListItems(n).SubItems(7)
               Else
               vOldDiscount = 0
               End If
            End If
         Next n
          If vOldQty = 0 Then
             Me.TBDIKeyQty.Text = 1
          Else
             Me.TBDIKeyQty.Text = vOldQty
          End If
          If vOldDiscount <> 0 Then
             Me.TBDIKeyDiscount.Text = vOldDiscount
          Else
             Me.TBDIKeyDiscount.Text = ""
          End If
          Me.TBDIKeyQty.SetFocus
       Else
          MsgBox "บาร์โค้ด/รหัสสินค้า นี้ " & vBarCode & " ไม่มีข้อมูลในระบบ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
          Me.TBDIBarCode.Text = ""
          Me.TBDIBarCode.SetFocus
       End If
    vRecordset.Close
   End If
End If
End Sub

Private Sub TBDICarLicense_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCarlicense As String
Dim vCountDocNo As Integer
Dim vDocNo As String

Dim i As Integer
Dim vNetItemAmount As Double
Dim vItemCode As String
Dim vItemName As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vQTY As Double
Dim vPrice As Double
Dim vAmount As Double
Dim vUnitCode As String
Dim vPickZone As String
Dim vBarCode As String
Dim vShelfID As String
Dim vZoneID As String
Dim vIndex As Integer
Dim vPointZone As String

Dim vTotalNetAmount As Double
Dim vListItem As ListItem
Dim n As Integer

Dim vMemLinePickZone As String
Dim x As Integer

   If Me.TBDICarLicense.Text <> "" Then
          
      Me.ListViewDIItem.ListItems.Clear
      Me.LBLDINetAmount.Caption = ""
      
      vCarlicense = Me.TBDICarLicense.Text
      
      vQuery = "exec dbo.USP_NP_SearchCarLicenseDriveIn1 '" & vCarlicense & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         vCountDocNo = vRecordset.RecordCount
         If vCountDocNo = 1 Then
            vDocNo = vRecordset.Fields("docno").Value
         End If
      End If
      vRecordset.Close
      
      If vCountDocNo > 1 Then

        Dim vSearch As String
        Dim vNetDebtAmount As Double
        
        vSearch = Me.TBDICarLicense.Text
        Me.ListViewDISearchDI.ListItems.Clear
        vQuery = "exec dbo.usp_np_SearchDriveInMaster1 '" & vSearch & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
           vRecordset.MoveFirst
           For i = 1 To vRecordset.RecordCount
           Set vListItem = Me.ListViewDISearchDI.ListItems.Add(, , i)
           vListItem.SubItems(1) = vRecordset.Fields("docno").Value
           vListItem.SubItems(2) = vRecordset.Fields("docdate").Value
           vListItem.SubItems(3) = vRecordset.Fields("arname").Value
           vListItem.SubItems(4) = vRecordset.Fields("refid").Value
           vListItem.SubItems(5) = vRecordset.Fields("salename").Value
           vNetDebtAmount = vRecordset.Fields("totalnetamount").Value
           vListItem.SubItems(6) = Format(vNetDebtAmount, "##,##0.00")
           vListItem.SubItems(7) = vRecordset.Fields("iscancel").Value
           vListItem.SubItems(8) = vRecordset.Fields("isconfirm").Value
           vRecordset.MoveNext
           Next i
        End If
        vRecordset.Close
        Me.PICDISearchDI.Visible = True
        Me.TBDISearchDI.SetFocus
        Exit Sub
      End If
      
      If vCountDocNo = 1 Then
            
         vPointZone = Me.LBLDI.Caption
        
         vQuery = "exec dbo.USP_NP_SearchDriveInDetails1 '" & vDocNo & "','" & vPointZone & "' "
         If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCountItemPickZoneOld = 0
            vDIIsOpen = 1
            vDIIsCancel = vRecordset.Fields("iscancel").Value
            vDIIsConfirm = vRecordset.Fields("isconfirm").Value
            vDIIsSendQue = vRecordset.Fields("issendque").Value
            
            vCountItemOld = vRecordset.RecordCount
            Me.LBLDIDocNo.Caption = vRecordset.Fields("docno").Value
            Me.LBLDIDocDate.Caption = vRecordset.Fields("docdate").Value
            Me.TBDIArCode.Text = vRecordset.Fields("arcode").Value
            Me.TBDISaleCode.Text = vRecordset.Fields("salecode").Value
            vTotalNetAmount = vRecordset.Fields("totalnetamount").Value
            Me.LBLDINetAmount.Caption = Format(vTotalNetAmount, "##,##0.00")
            
            If vDIIsCancel = 1 Then
            Call CancelDoc
            ElseIf vDIIsConfirm = 1 Then
            Call ConfirmDoc
            Else
            Call NewDoc
            End If
            
            ReDim vDIItemCodeOld(vCountItemOld) As String
            ReDim vDIUnitCodeOld(vCountItemOld) As String
            ReDim vDIWHCodeOld(vCountItemOld) As String
            ReDim vDIShelfCodeOld(vCountItemOld) As String
            ReDim vDIZoneIDOld(vCountItemOld) As String
            ReDim vDIPickZoneOld(vCountItemOld) As String
            ReDim vDIBarCodeOld(vCountItemOld) As String
            
            For i = 1 To vRecordset.RecordCount
            vDIItemCodeOld(i) = vRecordset.Fields("itemcode").Value
            vDIUnitCodeOld(i) = vRecordset.Fields("unitcode").Value
            vDIWHCodeOld(i) = vRecordset.Fields("whcode").Value
            vDIShelfCodeOld(i) = vRecordset.Fields("shelfcode").Value
            vDIZoneIDOld(i) = vRecordset.Fields("zoneid").Value
            vDIBarCodeOld(i) = vRecordset.Fields("barcode").Value
            vDIPickZoneOld(i) = vRecordset.Fields("pickzone").Value
            
            If vPointZone = vDIPickZoneOld(i) Then
               vCountItemPickZoneOld = vCountItemPickZoneOld + 1
            End If
            
            vPickZone = vRecordset.Fields("pickzone").Value
            vItemCode = vRecordset.Fields("itemcode").Value
            vItemName = vRecordset.Fields("itemname").Value
            vWHCode = vRecordset.Fields("whcode").Value
            vShelfCode = vRecordset.Fields("shelfcode").Value
            vQTY = vRecordset.Fields("qty").Value
            vUnitCode = vRecordset.Fields("unitcode").Value
            vPrice = vRecordset.Fields("price").Value
            vAmount = vRecordset.Fields("amount").Value
            vBarCode = vRecordset.Fields("barcode").Value
            vShelfID = vRecordset.Fields("shelfid").Value
            vZoneID = vRecordset.Fields("zoneid").Value
            
            Set vListItem = Me.ListViewDIItem.ListItems.Add(, , i)
            vListItem.SubItems(1) = vItemCode
            vListItem.SubItems(2) = vItemName
            vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
            vListItem.SubItems(4) = vUnitCode
            vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
            vListItem.SubItems(6) = Format(vAmount, "##,##0.00")
            vListItem.SubItems(7) = Format(0, "##,##0.00")
            vListItem.SubItems(8) = vWHCode
            vListItem.SubItems(9) = vShelfCode
            vListItem.SubItems(10) = vZoneID
            vListItem.SubItems(11) = vShelfID
            vListItem.SubItems(12) = vBarCode
            vListItem.SubItems(13) = ""
            vListItem.SubItems(14) = vPickZone
             vRecordset.MoveNext
            Next i
            
            For x = 1 To Me.ListViewDIItem.ListItems.Count
            vMemLinePickZone = Me.ListViewDIItem.ListItems(x).SubItems(14)
            If vPointZone <> vMemLinePickZone Then
            Me.ListViewDIItem.ListItems.Item(x).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(1).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(2).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(3).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(4).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(5).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(6).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(7).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(8).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(9).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(10).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(11).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(12).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(13).ForeColor = "&H00008000"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(14).ForeColor = "&H00008000"
            Else
            Me.ListViewDIItem.ListItems.Item(x).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(1).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(2).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(3).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(4).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(5).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(6).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(7).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(8).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(9).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(10).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(11).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(12).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(13).ForeColor = "&H80000008"
            Me.ListViewDIItem.ListItems.Item(x).ListSubItems(14).ForeColor = "&H80000008"
            End If
            
            Next x

         End If
         Me.TBDIBarCode.SetFocus
      End If
   End If

End Sub

Private Sub TBDICarLicense_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vHeader As String
Dim vRunning As Integer
Dim vDocNumber As String

vQuery = "exec dbo.USP_NP_SearchNewDocNo  29 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vHeader = Trim(vRecordset.Fields("header").Value)
    vRunning = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
End If
vRecordset.Close
      
vDocNo = vDocNumber & vHeader & "-" & Format(vRunning, "0000")
Me.LBLDIDocNo.Caption = vDocNo
vDIIsOpen = 0
vMemDIIsCancel = 0
vDISendQue = 0
Me.TBDIArCode.Text = ""
Me.TBDIArCode.Text = "99999"

Me.TBDISaleCode.Text = ""
Me.LBLDINetAmount.Caption = ""
Me.LBLDIMember.Caption = ""

Me.ListViewDIItem.ListItems.Clear
   
Call NewDoc
Me.CMDDISendQue.Enabled = False
Me.TBDICarLicense.SetFocus
End If

If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If

If KeyCode = 38 Then
  Me.TBDIArCode.SetFocus
End If

If KeyCode = 40 Then
   Me.TBDISaleCode.SetFocus
End If
End Sub

Private Sub TBDICarLicense_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.TBDISaleCode.SetFocus
End If
End Sub

Private Sub TBDIKeyArCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
   Me.TBDIKeyMember.SetFocus
End If

If KeyCode = 27 Then
   Me.PICArKeyData.Visible = False
   Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub TBDIKeyDiscount_Change()
Dim vDiscountWord As String
Dim vIsNumber As Boolean
Dim vLenDiscount As Integer

Dim vPrice As Double
Dim vDiscountAmount As Double

Dim vLenDisCountWord As Integer
Dim vDiscAmount As Double


If Me.TBDIKeyDiscount.Text <> "" Then
   vDiscountWord = Me.TBDIKeyDiscount.Text
   CheckNumber (vDiscountWord)
   
   If vCheckValueNumber = False Then
      MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      vLenDiscount = Len(Me.TBDIKeyDiscount.Text)
      Me.TBDIKeyDiscount.Text = Left(Me.TBDIKeyDiscount.Text, vLenDiscount - 1)
      'Me.TBDIKeyDiscount.SetFocus
      Exit Sub
   End If
   
   Me.LBLDIDiscountWord.Caption = Me.TBDIKeyDiscount.Text
   If Me.LBLDIDiscountWord.Caption <> "" And Me.LBLDIDiscountWord.Caption <> "." Then
      vDiscAmount = Me.LBLDIDiscountWord.Caption
      Me.LBLDIDiscountWord.Caption = Format(vDiscAmount, "##,##0.00")
   End If
   
   vPrice = Me.LBLDIPrice.Caption
   If Me.LBLDIDiscountWord.Caption <> "." Then
   vDiscountAmount = Me.LBLDIDiscountWord.Caption
   End If
   
   If vPrice - vDiscountAmount <= 0 Then
      MsgBox "ไม่สามารถลดราคาเท่ากับหรือน้อยกว่าราคาจริงได้", vbApplicationModal, "Send Error Message"
      vLenDisCountWord = Len(Me.TBDIKeyDiscount.Text)
      Me.TBDIKeyDiscount.Text = Left(Me.TBDIKeyDiscount.Text, vLenDisCountWord - 1)
      'Me.TBDIKeyDiscount.SetFocus
      Exit Sub
   End If
   
   Call CalcDIItemAmount
Else
   Me.LBLDIDiscountWord.Caption = Format(0, "##,##0.00")
   Call CalcDIItemAmount
End If
End Sub

Private Sub TBDIKeyDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   Me.TBDIKeyQty.SetFocus
End If

If KeyCode = 27 Then
   Me.PICDIKeyQty.Visible = False
   Me.TBDIBarCode.Text = ""
   Me.TBDIBarCode.SetFocus
End If
End Sub

Private Sub TBDIKeyDiscount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CMDDIKeyQtyOK_Click
End If
End Sub

Private Sub TBDIKeyMember_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 27 Then
   Me.PICArKeyData.Visible = False
   Me.TBDIArCode.SetFocus
End If
End Sub

Private Sub TBDIKeyMember_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String

If KeyAscii = 13 Then
If Me.TBDIKeyMember.Text <> "" Then
   vSearchAR = Me.TBDIKeyMember.Text
   vQuery = "exec dbo.USP_AR_SearchMemberID '" & vSearchAR & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.TBDIArCode.Text = Trim(vRecordset.Fields("arcode").Value)
      Me.TBDICarLicense.SetFocus
   Else
      Me.TBDIArCode.Text = ""
      Me.TBDIArCode.SetFocus
   End If
   Me.PICArKeyData.Visible = False
   vRecordset.Close
End If
End If
End Sub

Private Sub TBDIKeyQty_Change()
Dim vQtyWord As String
Dim vIsNumber As Boolean
Dim vLenQTY As Integer

Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vNetprice As Double

Dim vItemCode As String
Dim vUnitCode As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vNewPrice As Double


If Me.TBDIKeyQty.Text <> "" Then
   vQtyWord = Me.TBDIKeyQty.Text
   CheckNumber (vQtyWord)
   
   If vCheckValueNumber = False Then
      MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      vLenQTY = Len(Me.TBDIKeyQty.Text)
      Me.TBDIKeyQty.Text = Left(Me.TBDIKeyQty.Text, vLenQTY - 1)
      Me.TBDIKeyQty.SetFocus
      Exit Sub
   End If
   
   
   vItemCode = Me.LBLDIItemCode.Caption
   vUnitCode = Me.LBLDIUnitCode.Caption
   If Me.TBDIKeyQty.Text <> "" Then
   vQTY = Me.TBDIKeyQty.Text
   End If

   
    vQuery = "exec dbo.USP_NP_SearchItemPriceQty1 '" & vItemCode & "'," & vQTY & ",'" & vUnitCode & "'"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vNewPrice = Trim(vRecordset.Fields("saleprice1").Value)
       Me.LBLDIPrice.Caption = Format(vNewPrice, "##,##0.00")
    End If
    vRecordset.Close
    
    
   'Call CalcDIItemAmount
   
End If
End Sub

Private Sub TBDIKeyQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
   'Me.TBDIKeyDiscount.SetFocus
   Me.CMDDIKeyQtyOK.SetFocus
End If

If KeyCode = 27 Then
   Me.PICDIKeyQty.Visible = False
   Me.TBDIBarCode.Text = ""
   Me.TBDIBarCode.SetFocus
End If
End Sub

Private Sub TBDIKeyQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   'Me.TBDIKeyDiscount.SetFocus
   Call CMDDIKeyQtyOK_Click
End If
End Sub

Private Sub TBDISaleCode_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchSale As String
Dim vSaleCode As String
Dim vLen As Integer
Dim vInstr As Integer

If Me.TBDISaleCode.Text <> "" Then
   vSearchSale = Me.TBDISaleCode.Text
   If InStr(vSearchSale, "/") <> 0 Then
      vInstr = InStr(vSearchSale, "/")
      vLen = Len(vSearchSale)
      vSaleCode = Left(vSearchSale, vInstr - 1)
      vQuery = "exec dbo.USP_CRM_EmployeeDetails 1,'" & vSaleCode & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         Me.TBDISaleCode.Text = Trim(vRecordset.Fields("empcode").Value) & "/" & Trim(vRecordset.Fields("empname").Value)
      Else
         Me.TBDISaleCode.Text = vSaleCode
      End If
      vRecordset.Close
      Me.TBDIBarCode.SetFocus
   Else
      vQuery = "exec dbo.USP_CRM_EmployeeDetails 1,'" & vSearchSale & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         Me.TBDISaleCode.Text = Trim(vRecordset.Fields("empcode").Value) & "/" & Trim(vRecordset.Fields("empname").Value)
         Me.TBDIBarCode.SetFocus
      End If
      vRecordset.Close
   End If
End If
End Sub

Private Sub TBDISaleCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Call CMDDIClearScreen_Click
End If

If KeyCode = 118 Then
   Me.PICArKeyData.Visible = True
   Me.TBDIKeyArCode.Text = ""
   Me.TBDIKeyMember.Text = ""
   Me.TBDIKeyArCode.SetFocus
End If

If KeyCode = 116 Then
Call CMDDISave_Click
End If

If KeyCode = 112 Then
Call CMDDISearchDocNo_Click
End If

If KeyCode = 113 Then
Call CMDDISearchAr_Click
End If

If KeyCode = 114 Then
Call CMDDISearchSale_Click
End If

If KeyCode = 115 Then
Call CMDDISearchItem_Click
End If

If KeyCode = 119 Then
Call CMDDICancel_Click
End If

If KeyCode = 120 Then
Call CMDDISendQue_Click
End If

If KeyCode = 38 Then
   Me.TBDICarLicense.SetFocus
End If

If KeyCode = 40 Then
   Me.TBDIBarCode.SetFocus
End If
End Sub

Private Sub TBDISaleCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.TBDIBarCode.SetFocus
End If
End Sub

Private Sub TBDISearchAR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchAR.Visible = False
Me.TBDIArCode.SetFocus
End If

If KeyCode = 40 Then
Me.ListViewDISearchAR.SetFocus
End If
End Sub

Private Sub TBDISearchAR_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String
Dim vListItem As ListItem
Dim i As Integer

If KeyAscii = 13 Then
   If Me.TBDISearchAR.Text <> "" Then
      vSearchAR = Me.TBDISearchAR.Text
      vQuery = "exec dbo.USP_AR_SearchARLine '" & vSearchAR & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Me.ListViewDISearchAR.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             Set vListItem = Me.ListViewDISearchAR.ListItems.Add(, , i)
             vListItem.SubItems(1) = Trim(vRecordset.Fields("arcode").Value)
             vListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
             vListItem.SubItems(3) = Trim(vRecordset.Fields("memberid").Value)
             vRecordset.MoveNext
          Next i
          Me.ListViewDISearchAR.SetFocus
      Else
      Me.TBDISearchAR.SetFocus
      End If
      vRecordset.Close
   End If
End If
End Sub

Private Sub TBDISearchDI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchDI.Visible = False
Me.TBDIArCode.SetFocus
End If

If KeyCode = 40 Then
Me.ListViewDISearchDI.SetFocus
End If
End Sub

Private Sub TBDISearchDI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call CMDDISearchDIClick_Click
End If
End Sub

Private Sub TBDISearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchItem.Visible = False
Me.TBDIBarCode.SetFocus
End If

If KeyCode = 40 Then
Me.ListViewDISearchItem.SetFocus
End If
End Sub

Private Sub TBDISearchItem_KeyPress(KeyAscii As Integer)
Dim vSearch As String
Dim vPrice As Double
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vQTY As Double
Dim vRemainOutQTY As Double

If KeyAscii = 13 Then
   If Me.TBDISearchItem.Text <> "" Then
      vSearch = Me.TBDISearchItem.Text
      vQuery = "exec dbo.USP_NP_SearchBarCode '" & vSearch & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Me.ListViewDISearchItem.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             vQTY = Trim(vRecordset.Fields("stockqty").Value)
             vRemainOutQTY = Trim(vRecordset.Fields("remainoutqty").Value)
             vPrice = Trim(vRecordset.Fields("price").Value)
             
             Set vListItem = Me.ListViewDISearchItem.ListItems.Add(, , i)
             vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
             vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
             vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
             vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
             vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
             vListItem.SubItems(6) = Trim(vRecordset.Fields("zone").Value)
             vRecordset.MoveNext
          Next i
          
          Me.ListViewDISearchItem.SetFocus
       Else
          Me.ListViewDISearchItem.ListItems.Clear
          Me.TBDISearchItem.SetFocus
       End If
    vRecordset.Close
   End If
End If
End Sub

Private Sub TBDISearchSale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICDISearchSale.Visible = False
Me.TBDISaleCode.SetFocus
End If

If KeyCode = 40 Then
   Me.ListViewDISearchSale.SetFocus
End If
End Sub

Private Sub TBDISearchSale_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchSale As String
Dim vListItem As ListItem
Dim i As Integer

If KeyAscii = 13 Then
   If Me.TBDISearchSale.Text <> "" Then
   vSearchSale = Me.TBDISearchSale.Text
   vQuery = "exec dbo.USP_CRM_EmployeeDetails 0,'" & vSearchSale & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       Me.ListViewDISearchSale.ListItems.Clear
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
          Set vListItem = Me.ListViewDISearchSale.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("empcode").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("empname").Value)
          vRecordset.MoveNext
       Next i
       Me.ListViewDISearchSale.SetFocus
   Else
      Me.TBDISearchSale.SetFocus
   End If
   vRecordset.Close
   End If
End If
End Sub

Private Sub TBDocNo_Change()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim i As Integer
'Dim vListItem As ListItem

'Dim vDocNo As String
'Dim vSumItemAmount As Double
'Dim vTaxAmount As Double
'Dim vNetAmount As Double
'Dim vLastDisCountAmount As Double

'Dim vQTY As Double
'Dim vPrice As Double
'Dim vDiscountAmount As Double
'Dim vAmount As Double

'If Me.TBDocNo.Text <> "" Then
 '  vDocNo = Me.TBDocNo.Text
'
 '  vQuery = "exec dbo.USP_NP_SearchSaleOrder '" & vDocNo & "'"
  ' If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   '   Me.ListViewSaleOrder.ListItems.Clear
    '  vSumItemAmount = Trim(vRecordset.Fields("sumofitemamount").Value)
     ' vTaxAmount = Trim(vRecordset.Fields("taxamount").Value)
      'vNetAmount = Trim(vRecordset.Fields("netamount").Value)
      'vLastDisCountAmount = Trim(vRecordset.Fields("discountamount").Value)
      
      'Me.LBLOrderBillType.Caption = Trim(vRecordset.Fields("billtype").Value)
      'Me.LBLOrderSoStatus.Caption = Trim(vRecordset.Fields("sostatus").Value)
      'Me.LBLOrderDocDate.Caption = Trim(vRecordset.Fields("docdate").Value)
      'Me.LBLOrderArCode.Caption = Trim(vRecordset.Fields("arcode").Value)
      'Me.LBLOrderArName.Caption = Trim(vRecordset.Fields("arname").Value)
      'If Trim(vRecordset.Fields("salecode").Value) = "" Then
       '   MsgBox "เอกสารไม่ได้กำหนดรหัสพนักงานขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      'Else
       '     Me.LBLOrderSaleCode.Caption = Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      'End If

      'Me.LBLOrderSumOfItemAmount.Caption = Format(vSumItemAmount, "##,##0.00")
      'Me.LBLOrderTaxAmount.Caption = Format(vTaxAmount, "##,##0.00")
      'Me.LBLOrderNetAmount.Caption = Format(vNetAmount, "##,##0.00")
      'if vLastDisCountAmount <> 0 Then
      'Me.TBSOLastDisCount.Text = Format(vLastDisCountAmount, "##,##0.00")
      'Me.LBLOrderDiscountOld.Caption = Format(vLastDisCountAmount, "##,##0.00")
      'Me.TBSOLastDisCount.Enabled = True
      'Else
      'Me.TBSOLastDisCount.Text = Format(vLastDisCountAmount, "##,##0.00")
      'Me.LBLOrderDiscountOld.Caption = Format(vLastDisCountAmount, "##,##0.00")
      'Me.TBSOLastDisCount.Enabled = False
      
      'End If
      
      'vRecordset.MoveFirst
      'For i = 1 To vRecordset.RecordCount
       ' Set vListItem = Me.ListViewSaleOrder.ListItems.Add(, , i)
        
        'vQTY = Trim(vRecordset.Fields("remainqty").Value)
        'vPrice = Trim(vRecordset.Fields("price").Value)
 '       'vDiscountAmount = Trim(vRecordset.Fields("discountamountsub").Value)
  '      vAmount = Trim(vRecordset.Fields("amount").Value)
''
   '     vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
    '    vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
     '   vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
      '  vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
       ' vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
        'vListItem.SubItems(6) = Format(vDiscountAmount, "##,##0.00")
        'vListItem.SubItems(7) = Format(vAmount, "##,##0.00")
        'vListItem.SubItems(8) = Trim(vRecordset.Fields("whcode").Value)
        'vListItem.SubItems(9) = Trim(vRecordset.Fields("shelfcode").Value)
        'vListItem.SubItems(10) = Trim(vRecordset.Fields("zoneid").Value)
        'vListItem.SubItems(11) = Trim(vRecordset.Fields("shelfid").Value)
        'vListItem.SubItems(12) = Trim(vRecordset.Fields("itemcode").Value)
        'vListItem.SubItems(13) = Trim(vRecordset.Fields("discountwordsub").Value)
        'vListItem.SubItems(14) = Format(vQTY, "##,##0.00")
        'vRecordset.MoveNext
      'Next i
   'Else
    '  Me.ListViewSaleOrder.ListItems.Clear
     ' Me.LBLOrderDocDate.Caption = ""
      'Me.LBLOrderArCode.Caption = ""
      'Me.LBLOrderArName.Caption = ""
      'Me.LBLOrderSaleCode.Caption = ""
      'Me.LBLOrderSumOfItemAmount.Caption = ""
      'Me.LBLOrderTaxAmount.Caption = ""
      'Me.LBLOrderNetAmount.Caption = ""
      'Me.TBSOLastDisCount.Text = ""
      'Me.LBLOrderDiscountOld.Caption = ""
      'Me.LBLOrderBillType.Caption = ""
      'Me.LBLOrderSoStatus.Caption = ""

   'End If
   'vRecordset.Close
   'Call CalcEditItemQty
'Else
      'Me.ListViewSaleOrder.ListItems.Clear
      'Me.LBLOrderDocDate.Caption = ""
      'Me.LBLOrderArCode.Caption = ""
      'Me.LBLOrderArName.Caption = ""
      'Me.LBLOrderSaleCode.Caption = ""
      'Me.LBLOrderSumOfItemAmount.Caption = ""
      'Me.LBLOrderTaxAmount.Caption = ""
      'Me.LBLOrderNetAmount.Caption = ""
      'Me.TBSOLastDisCount.Text = ""
      'Me.LBLOrderDiscountOld.Caption = ""
      'Me.LBLOrderBillType.Caption = ""
      'Me.LBLOrderSoStatus.Caption = ""
'End If
End Sub

Private Sub TBDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 116 Then
'Call CMDSaleOrderSendQue_Click
'End If
End Sub

Private Sub TBEditQty_Change()
'Dim vQtyWord As String
'Dim vLenQTY As Integer

'If Me.TBEditQty.Text <> "" Then
 '  vQtyWord = Me.TBEditQty.Text
  ' CheckNumber (vQtyWord)
      
  ' If vCheckValueNumber = False Then
   '   MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
    '  vLenQTY = Len(Me.TBEditQty.Text)
     ' Me.TBEditQty.Text = Left(Me.TBEditQty.Text, vLenQTY - 1)
      'Me.TBEditQty.SetFocus
   'End If
'End If
End Sub

Private Sub TBKeyArCode_Change()

End Sub

Private Sub TBDIKeyArCode_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String

If KeyAscii = 13 Then
If Me.TBDIKeyArCode.Text <> "" Then
   vSearchAR = Me.TBDIKeyArCode.Text
   vQuery = "exec dbo.usp_ar_arprofile '" & vSearchAR & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.TBDIArCode.Text = Trim(vRecordset.Fields("arcode").Value)
      Me.TBDICarLicense.SetFocus
   Else
      Me.TBDIArCode.Text = ""
      Me.TBDIArCode.SetFocus
   End If
   vRecordset.Close
   Me.PICArKeyData.Visible = False
End If
End If
End Sub

Private Sub TBEditQty_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
 '  Call CMDEditOK_Click
'End If
End Sub

Private Sub TBPRKeyMember_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.PICPRArKeyData.Visible = False
   Me.TXTArCode.SetFocus
End If
End Sub

Private Sub TBPRKeyMember_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String

If KeyAscii = 13 Then
   If Me.TBPRKeyMember.Text <> "" Then
      vSearchAR = Me.TBPRKeyMember.Text
      vQuery = "exec dbo.USP_AR_SearchMemberID '" & vSearchAR & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         Me.TXTArCode.Text = Trim(vRecordset.Fields("arcode").Value)
         Me.TXTSaleCode.SetFocus
      Else
         Me.TXTArCode.Text = ""
         Me.TXTArCode.SetFocus
      End If
      Me.PICPRArKeyData.Visible = False
      vRecordset.Close
   End If
   
End If
End Sub

Private Sub TBPRSearchAR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchAR.Visible = False
Me.TXTArCode.SetFocus
End If

If KeyCode = 40 Then
   Me.ListViewPRSearchAR.SetFocus
End If
End Sub

Private Sub TBPRSearchAR_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String
Dim vListItem As ListItem
Dim i As Integer

If KeyAscii = 13 Then
   If Me.TBPRSearchAR.Text <> "" Then
      vSearchAR = Me.TBPRSearchAR.Text
      vQuery = "exec dbo.USP_AR_ARProFileSearch '" & vSearchAR & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Me.ListViewPRSearchAR.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             Set vListItem = Me.ListViewPRSearchAR.ListItems.Add(, , i)
             vListItem.SubItems(1) = Trim(vRecordset.Fields("arcode").Value)
             vListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
             vListItem.SubItems(3) = Trim(vRecordset.Fields("memberid").Value)
             vRecordset.MoveNext
          Next i
          Me.ListViewPRSearchAR.SetFocus
      End If
      vRecordset.Close
   End If
End If
End Sub

Private Sub TBPRSearchDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchDocNo.Visible = False
Me.TXTDocNo.SetFocus
End If

If KeyCode = 40 Then
Me.ListViewPRSearchDocNo.SetFocus
End If
End Sub

Private Sub TBPRSearchDocNo_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearch As String
Dim vListItem As ListItem
Dim i As Integer

If KeyAscii = 13 Then
   If Me.TBPRSearchDocNo.Text <> "" Then
   vSearch = Me.TBPRSearchDocNo.Text
   vQuery = "exec dbo.USP_NP_SearchReqPicking '" & vSearch & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       Me.ListViewPRSearchDocNo.ListItems.Clear
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
          Set vListItem = Me.ListViewPRSearchDocNo.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
          vListItem.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
          vListItem.SubItems(4) = Trim(vRecordset.Fields("salename").Value)
          vListItem.SubItems(5) = Trim(vRecordset.Fields("netdebtamount").Value)
          vRecordset.MoveNext
       Next i
       Me.ListViewPRSearchDocNo.SetFocus
   End If
   vRecordset.Close
   End If
   Me.TBPRSearchDocNo.SetFocus
End If
End Sub

Private Sub TBPRSearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchItem.Visible = False
Me.TBBarCode.SetFocus
End If

If KeyCode = 40 Then
Me.ListViewPRSearchItem.SetFocus
End If
End Sub

Private Sub TBPRSearchSale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICPRSearchSale.Visible = False
Me.TXTSaleCode.SetFocus
End If

If KeyCode = 40 Then
Me.ListViewPRSearchSale.SetFocus
End If
End Sub

Private Sub TBPRSearchSale_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchSale As String
Dim vListItem As ListItem
Dim i As Integer

If KeyAscii = 13 Then
   If Me.TBPRSearchSale.Text <> "" Then
      vSearchSale = Me.TBPRSearchSale.Text
      vQuery = "exec dbo.USP_CRM_EmployeeDetails 0,'" & vSearchSale & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Me.ListViewPRSearchSale.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             Set vListItem = Me.ListViewPRSearchSale.ListItems.Add(, , i)
             vListItem.SubItems(1) = Trim(vRecordset.Fields("empcode").Value)
             vListItem.SubItems(2) = Trim(vRecordset.Fields("empname").Value)
             vRecordset.MoveNext
          Next i
       Me.ListViewPRSearchSale.SetFocus
      End If
      vRecordset.Close
   End If
End If
End Sub

Private Sub TBQTY_Change()
Dim vQtyWord As String
Dim vIsNumber As Boolean
Dim vLenQTY As Integer

Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vNetprice As Double

Dim vNewPrice As Double
Dim vItemCode As String
Dim vUnitCode As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

If Me.TBQty.Text <> "" Then
   vQtyWord = Me.TBQty.Text
   CheckNumber (vQtyWord)
   
   If vCheckValueNumber = False Then
      MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      vLenQTY = Len(Me.TBQty.Text)
      Me.TBQty.Text = Left(Me.TBQty.Text, vLenQTY - 1)
      Me.TBQty.SetFocus
      Exit Sub
   End If
   
   vItemCode = Me.LBLItemCode.Caption
   vUnitCode = Me.LBLUnitCode.Caption
   If Me.TBQty.Text <> "" Then
   vQTY = Me.TBQty.Text
   End If
   
    vQuery = "exec dbo.USP_NP_SearchItemPriceQty1 '" & vItemCode & "'," & vQTY & ",'" & vUnitCode & "'"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vNewPrice = Trim(vRecordset.Fields("saleprice1").Value)
       Me.LBLPrice.Caption = Format(vNewPrice, "##,##0.00")
    End If
    vRecordset.Close
    
   Call CalcDIItemAmount
   
End If
End Sub

Private Sub TBQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSearchItem.Visible = False
Me.TBBarCode.Text = ""
Me.TBBarCode.SetFocus
End If
End Sub

Private Sub TBQTY_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.TXTDisCount.SetFocus
End If
End Sub


Private Sub TBPRSearchItem_KeyPress(KeyAscii As Integer)
Dim vSearch As String
Dim vPrice As Double
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vQTY As Double
Dim vRemainOutQTY As Double

If KeyAscii = 13 Then
   If Me.TBPRSearchItem.Text <> "" Then
      vSearch = Me.TBPRSearchItem.Text
      vQuery = "exec dbo.USP_NP_SearchBarCode '" & vSearch & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Me.ListViewPRSearchItem.ListItems.Clear
          vRecordset.MoveFirst
          For i = 1 To vRecordset.RecordCount
             vQTY = Trim(vRecordset.Fields("stockqty").Value)
             vRemainOutQTY = Trim(vRecordset.Fields("remainoutqty").Value)
             vPrice = Trim(vRecordset.Fields("price").Value)
             
             Set vListItem = Me.ListViewPRSearchItem.ListItems.Add(, , i)
             vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
             vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
             vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
             vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
             vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
             vListItem.SubItems(6) = Trim(vRecordset.Fields("zone").Value)
             vRecordset.MoveNext
          Next i
          
          Me.ListViewPRSearchItem.SetFocus
       Else
          Me.ListViewPRSearchItem.ListItems.Clear
          Me.TBPRSearchItem.SetFocus
       End If
    vRecordset.Close
   End If
End If
End Sub

Private Sub TBSOLastDisCount_Change()
'Dim vQtyWord As String
'Dim vIsNumber As Boolean
'Dim vLenQTY As Integer

'Dim vQTY As Double
'Dim vPrice As Double
'Dim vDiscountAmount As Double
'Dim vNetprice As Double


'If Me.TBSOLastDisCount.Text <> "" Then
 '  vQtyWord = Me.TBSOLastDisCount.Text
  ' CheckNumber (vQtyWord)
   '
   'If vCheckValueNumber = False Then
    '  MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
     ' vLenQTY = Len(Me.TBSOLastDisCount.Text)
      'Me.TBSOLastDisCount.Text = Left(Me.TBSOLastDisCount.Text, vLenQTY - 1)
    '  Me.TBSOLastDisCount.SetFocus
      ''Exit Sub
   'End If
     
'End If
'Call CalcEditItemQty
End Sub

Private Sub TXTArCode_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchAR As String

If Me.TXTArCode.Text <> "" Then
   vSearchAR = Me.TXTArCode.Text
   vQuery = "exec dbo.usp_ar_searchar '" & vSearchAR & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.LBLArName.Caption = Trim(vRecordset.Fields("arname").Value)
      Me.TXTMember.Caption = Trim(vRecordset.Fields("memberid").Value)
   Else
      Me.LBLArName.Caption = ""
      Me.TXTMember.Caption = ""
   End If
   vRecordset.Close
Else
   Me.LBLArName.Caption = ""
   Me.TXTMember.Caption = ""
End If
End Sub

Private Sub TXTArCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If


If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If

If KeyCode = 38 Then
   Me.TXTDocNo.SetFocus
End If

If KeyCode = 40 Then
   Me.TXTSaleCode.SetFocus
End If

If KeyCode = 37 Then
   Me.TXTLicense.SetFocus
End If
End Sub

Private Sub TXTArCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.TXTSaleCode.SetFocus
End If
End Sub

Private Sub TXTDisCount_Change()
Dim vDiscountWord As String
Dim vIsNumber As Boolean
Dim vLenDiscount As Integer

Dim vPrice As Double
Dim vDiscountAmount As Double

Dim vLenDisCountWord As Integer
Dim vDiscAmount As Double


If Me.TXTDisCount.Text <> "" Then
   vDiscountWord = Me.TXTDisCount.Text
   CheckNumber (vDiscountWord)
   
   If vCheckValueNumber = False Then
      MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      vLenDiscount = Len(Me.TXTDisCount.Text)
      Me.TXTDisCount.Text = Left(Me.TXTDisCount.Text, vLenDiscount - 1)
      Me.TXTDisCount.SetFocus
      Exit Sub
   End If
   
   Me.LBLDisCountAmount.Caption = Me.TXTDisCount.Text
   If Me.LBLDisCountAmount.Caption <> "" And Me.LBLDisCountAmount.Caption <> "." Then
      vDiscAmount = Me.LBLDisCountAmount.Caption
      Me.LBLDisCountAmount.Caption = Format(vDiscAmount, "##,##0.00")
   End If
   
   vPrice = Me.LBLPrice.Caption
   If Me.LBLDisCountAmount.Caption <> "." Then
   vDiscountAmount = Me.LBLDisCountAmount.Caption
   End If
   
   If vPrice - vDiscountAmount <= 0 Then
      MsgBox "ไม่สามารถลดราคาได้เท่ากับหรือน้อยกว่าราคาขายได้ กรุณาตรวจสอบ", vbApplicationModal, "Send Error Message"
      vLenDisCountWord = Len(Me.TXTDisCount.Text)
      Me.TXTDisCount.Text = Left(Me.TXTDisCount.Text, vLenDisCountWord - 1)
      Me.TXTDisCount.SetFocus
      Exit Sub
   End If
   
   Call CalcPRItemAmount
Else
   Me.LBLDisCountAmount.Caption = Format(0, "##,##0.00")
   Call CalcPRItemAmount
End If
End Sub

Private Sub TXTDisCount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Me.PICSearchItem.Visible = False
Me.TBBarCode.Text = ""
Me.TBBarCode.SetFocus
End If
End Sub

Private Sub TXTDisCount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CMDOK_Click
End If
End Sub

Private Sub TXTDocNo_Change()
Dim vDocNo As String

If Me.TXTDocNo.Text <> "" Then
vDocNo = Me.TXTDocNo.Text
Call ShowReqDetails(vDocNo)
End If
End Sub

Private Sub TXTDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If

If KeyCode = 40 Then
   Me.TXTArCode.SetFocus
End If

If KeyCode = 39 Then
   Me.TXTLicense.SetFocus
End If

End Sub

Private Sub TXTDocNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.TXTLicense.SetFocus
End If
End Sub

Private Sub TXTLicense_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If

If KeyCode = 39 Then
   Me.TXTArCode.SetFocus
End If

If KeyCode = 40 Then
   Me.TXTArCode.SetFocus
End If

If KeyCode = 37 Then
   Me.TXTDocNo.SetFocus
End If

If KeyCode = 38 Then
   Me.TXTDocNo.SetFocus
End If
End Sub

Private Sub TXTSaleCode_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearchSale As String
Dim vSaleCode As String
Dim vLen As Integer
Dim vInstr As Integer

If Me.TXTSaleCode.Text <> "" Then
   vSearchSale = Me.TXTSaleCode.Text
   If InStr(vSearchSale, "/") <> 0 Then
      vInstr = InStr(vSearchSale, "/")
      vLen = Len(vSearchSale)
      vSaleCode = Left(vSearchSale, vInstr - 1)
      vQuery = "exec dbo.USP_CRM_EmployeeDetails 1,'" & vSaleCode & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         Me.TXTSaleCode.Text = Trim(vRecordset.Fields("empcode").Value) & "/" & Trim(vRecordset.Fields("empname").Value)
      Else
         Me.TXTSaleCode.Text = vSaleCode
      End If
      vRecordset.Close
   Else
      vQuery = "exec dbo.USP_CRM_EmployeeDetails 1,'" & vSearchSale & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         Me.TXTSaleCode.Text = Trim(vRecordset.Fields("empcode").Value) & "/" & Trim(vRecordset.Fields("empname").Value)
      End If
      vRecordset.Close
   End If
End If
End Sub


Private Sub TXTSearch_KeyPress(KeyAscii As Integer)
'Dim vSearch As String
'Dim vPrice As Double
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vListItem As ListItem
'Dim i As Integer
'Dim vQTY As Double
'Dim vRemainOutQTY As Double
'
'If KeyAscii = 13 Then
 '  If Me.TXTSearch.Text <> "" Then
  '    Me.PICSearchItem.Visible = True
   '   vSearch = Me.TXTSearch.Text
    '  vQuery = "exec dbo.USP_NP_SearchBarCode '" & vSearch & "' "
     ' If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      '    Me.ListViewSearch.ListItems.Clear
       '   vRecordset.MoveFirst
        '  For i = 1 To vRecordset.RecordCount
         '    vQTY = Trim(vRecordset.Fields("stockqty").Value)
          '   vRemainOutQTY = Trim(vRecordset.Fields("remainoutqty").Value)
           '  vPrice = Trim(vRecordset.Fields("price").Value)
            '
             'Set vListItem = Me.ListViewSearch.ListItems.Add(, , i)
             'vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
             ''vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
             'vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
             ''vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
             'vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
             'vListItem.SubItems(6) = Trim(vRecordset.Fields("zone").Value)
             'vRecordset.MoveNext
          'Next i
          '
          'Me.ListViewItem.SetFocus
       'Else
        '  Me.ListViewSearch.ListItems.Clear
         ' Me.TXTSearch.SetFocus
       'End If
    'vRecordset.Close
   'End If
'End If
End Sub

Private Sub TXTSearchAR_KeyPress(KeyAscii As Integer)
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vSearchAR As String
'Dim vListItem As ListItem
'Dim i As Integer

'If KeyAscii = 13 Then
 '  If Me.TXTSearchAR.Text <> "" Then
  '    vSearchAR = Me.TXTSearchAR.Text
   '   vQuery = "exec dbo.USP_AR_ARProFileSearch '" & vSearchAR & "' "
    '  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     '     Me.ListViewAR.ListItems.Clear
      '    vRecordset.MoveFirst
       '   For i = 1 To vRecordset.RecordCount
        ''     Set vListItem = Me.ListViewAR.ListItems.Add(, , i)
          '   vListItem.SubItems(1) = Trim(vRecordset.Fields("arcode").Value)
           '  vListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
            '' vListItem.SubItems(3) = Trim(vRecordset.Fields("memberid").Value)
             'vRecordset.MoveNext
          'Next i
          'Me.ListViewAR.SetFocus
      'End If
      'vRecordset.Close
   'End If
'End If
End Sub

Private Sub TXTSearchDocNo_KeyPress(KeyAscii As Integer)
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vSearch As String
'Dim vListItem As ListItem
''Dim i As Integer

'If KeyAscii = 13 Then
 '  If Me.TXTSearchDocNo.Text <> "" Then
  ' vSearch = Me.TXTSearchDocNo.Text
   'vQuery = "exec dbo.USP_NP_SearchReqPicking '" & vSearch & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '   Me.ListViewDocNo.ListItems.Clear
     '  vRecordset.MoveFirst
      ' For i = 1 To vRecordset.RecordCount
       '   Set vListItem = Me.ListViewDocNo.ListItems.Add(, , i)
        '  vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
         ' vListItem.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
          'vListItem.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
          ''vListItem.SubItems(4) = Trim(vRecordset.Fields("salename").Value)
          'vListItem.SubItems(5) = Trim(vRecordset.Fields("netdebtamount").Value)
          'vRecordset.MoveNext
       'Next i
       'Me.ListViewDocNo.SetFocus
   'End If
   'vRecordset.Close
   'End If
'End If
'Me.TXTSearchDocNo.SetFocus
End Sub

Private Sub TXTSearchSale_KeyPress(KeyAscii As Integer)
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vSearchSale As String
'Dim vListItem As ListItem
'Dim i As Integer

'If KeyAscii = 13 Then
 '  If Me.TXTSearchSale.Text <> "" Then
  '    vSearchSale = Me.TXTSearchSale.Text
   '   vQuery = "exec dbo.USP_CRM_EmployeeDetails 0,'" & vSearchSale & "' "
    '  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     '     Me.ListViewSaleCode.ListItems.Clear
      '    vRecordset.MoveFirst
       '   For i = 1 To vRecordset.RecordCount
        '     Set vListItem = Me.ListViewSaleCode.ListItems.Add(, , i)
         '    vListItem.SubItems(1) = Trim(vRecordset.Fields("empcode").Value)
          '   vListItem.SubItems(2) = Trim(vRecordset.Fields("empname").Value)
           '  vRecordset.MoveNext
          'Next i
       'Me.ListViewSaleCode.SetFocus
      'End If
      'vRecordset.Close
   'End If
'End If
End Sub

Public Sub PrintPickingHeader(vQueID As Integer, vQueDocDate As String, vZone As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
Dim vQueNo As String
   
If vZone = "A" Then

vQuery = "exec dbo.USP_NP_SearchPrinter 2"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
End If
vRecordset.Close

If vPrinterName <> "" Then
   For Each prnPrinter In Printers
      If UCase(prnPrinter.DeviceName) = UCase(vPrinterName) Then
         Set Printer = prnPrinter
         Exit For
      End If
   Next
Else
   Exit Sub
End If

End If

If vZone = "B" Then

vQuery = "exec dbo.USP_NP_SearchPrinter 3"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
End If
vRecordset.Close

If vPrinterName <> "" Then
   For Each prnPrinter In Printers
      If UCase(prnPrinter.DeviceName) = UCase(vPrinterName) Then
        Set Printer = prnPrinter
         Exit For
      End If
   Next
Else
   Exit Sub
End If

End If

'If vZone = "C" Then
'For Each prnPrinter In Printers
 '  If prnPrinter.DeviceName = "\\hptc-5100420\SRP370C" Then
  '    Set Printer = prnPrinter
   '   Exit For
   'End If
'Next
'End If
        
vQuery = "exec dbo.USP_NP_SearchQueCenterDetails " & vQueID & ",'" & vQueDocDate & "'"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then

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
Printer.Print Trim("Picking Request Slip Master")

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
End If
vRecordset.Close
     
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("______________________________________________________________________________________________")

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

Public Sub PrintPickingSlip(vQueID As Integer, vQueDocDate As String, vZone As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
Dim vQueNo As String
   
If vZone = "A" Then

vQuery = "exec dbo.USP_NP_SearchPrinter 2"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
End If
vRecordset.Close

If vPrinterName <> "" Then
   For Each prnPrinter In Printers
      If UCase(prnPrinter.DeviceName) = UCase(vPrinterName) Then
         Set Printer = prnPrinter
         Exit For
      End If
   Next
Else
   Exit Sub
End If
End If

If vZone = "B" Then

vQuery = "exec dbo.USP_NP_SearchPrinter 3"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
End If
vRecordset.Close

If vPrinterName <> "" Then
   For Each prnPrinter In Printers
      If UCase(prnPrinter.DeviceName) = UCase(vPrinterName) Then
         Set Printer = prnPrinter
         Exit For
      End If
   Next
Else
   Exit Sub
End If

End If

'If vZone = "C" Then
'For Each prnPrinter In Printers
 '  If prnPrinter.DeviceName = "\\hptc-5100420\SRP370C" Then
  '    Set Printer = prnPrinter
   '   Exit For
   'End If
'Next
'End If
        
vQuery = "exec dbo.USP_NP_SearchQueCenterDetails " & vQueID & ",'" & vQueDocDate & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then

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

Public Sub PrintDriveInHeader(vDocNo As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
Dim vDIPoint As Integer
   
vDIPoint = Me.LBLDI.Caption
'If vDIPoint = "01" Then
'For Each prnPrinter In Printers
 '  If prnPrinter.DeviceName = "\\hptc-5100418\SRP370A" Then
  '    Set Printer = prnPrinter
   '   Exit For
   'End If
'Next
'End If

'If vDIPoint = "02" Then
'For Each prnPrinter In Printers
 '  If prnPrinter.DeviceName = "\\hptc-5100421\SRP370B" Then
  '    Set Printer = prnPrinter
   '   Exit For
   'End If
'Next
'End If

'If vDIPoint = "03" Then
'For Each prnPrinter In Printers
 '  If prnPrinter.DeviceName = "\\hptc-5100420\SRP370C" Then
  '    Set Printer = prnPrinter
   '   Exit For
   'End If
'Next
'End If

vQuery = "exec dbo.USP_NP_SearchDriveInDetails '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1450
Printer.FontBold = True
Printer.Print Trim("DriveIn Slip Master")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value) & "          " & Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)

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
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value) & "     " & "โซนการจัด :" & Trim(vRecordset.Fields("zoneid").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 20
Printer.CurrentX = 1000
Printer.FontBold = True
Printer.Print "มูลค่าสุทธิ :" & "       " & Format(Trim(vRecordset.Fields("totalnetamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "               ผู้จ่ายสินค้า                                             ผู้รับสินค้า"

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
End If
vRecordset.Close

'Printer.EndDoc
End Sub

Public Sub PrintDriveInDetails(vDocNo As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
Dim vDIPoint As Integer
   
vDIPoint = Me.LBLDI.Caption
'If vDIPoint = "01" Then
'For Each prnPrinter In Printers
 '  If prnPrinter.DeviceName = "\\hptc-5100418\SRP370A" Then
  '    Set Printer = prnPrinter
   '   Exit For
   'End If
'Next
'End If

'If vDIPoint = "02" Then
'For Each prnPrinter In Printers
 '  If prnPrinter.DeviceName = "\\hptc-5100421\SRP370B" Then
  '    Set Printer = prnPrinter
   '   Exit For
   'End If
'Next
'End If

'If vDIPoint = "03" Then
'For Each prnPrinter In Printers
 '  If prnPrinter.DeviceName = "\\hptc-5100420\SRP370C" Then
  '    Set Printer = prnPrinter
   '   Exit For
   'End If
'Next
'End If
        
vQuery = "exec dbo.USP_NP_SearchDriveInDetails '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1450
Printer.FontBold = True
Printer.Print Trim("DriveIn Slip Details")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value) & "          " & Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)

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
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value) & "     " & "โซนการจัด :" & Trim(vRecordset.Fields("zoneid").Value)


Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")
vRecordset.MoveFirst
n = 1
While Not vRecordset.EOF

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfid").Value)

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
Printer.Print "จ่าย:" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value)

If i = vRecordset.RecordCount Then

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 20
Printer.CurrentX = 1000
Printer.FontBold = True
Printer.Print "มูลค่าสุทธิ :" & "       " & Format(Trim(vRecordset.Fields("totalnetamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

End If

vRecordset.MoveNext
n = n + 1
Wend

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "               ผู้จ่ายสินค้า                                             ผู้รับสินค้า"

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
End If
vRecordset.Close

'Printer.EndDoc
End Sub

Private Sub TXTSaleCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then
   Me.PICPRArKeyData.Visible = True
   Me.TBPRKeyMember.Text = ""
   Me.TBPRKeyMember.SetFocus
End If

If KeyCode = 112 Then
Call CMDSearch_Click
End If

If KeyCode = 113 Then
Call CMDArCode_Click
End If

If KeyCode = 114 Then
Call CMDSale_Click
End If

If KeyCode = 115 Then
Call CMDSearchItem_Click
End If

If KeyCode = 116 Then
Call CMDSave_Click
End If

If KeyCode = 119 Then
Call CMDCancel_Click
End If

If KeyCode = 120 Then
Call CMDQue_Click
End If

If KeyCode = 38 Then
   Me.TXTArCode.SetFocus
End If

If KeyCode = 40 Then
   Me.TBBarCode.SetFocus
End If
End Sub

Private Sub TXTSaleCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.TBBarCode.SetFocus
End If
End Sub
