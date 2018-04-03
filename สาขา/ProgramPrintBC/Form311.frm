VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form311 
   Caption         =   "หน้าพิมพ์ใบสั่งขาย"
   ClientHeight    =   10590
   ClientLeft      =   2280
   ClientTop       =   705
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form311.frx":0000
   ScaleHeight     =   10590
   ScaleWidth      =   14880
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CKUnShowPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ไม่โชว์ราคา"
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
      Height          =   240
      Left            =   7425
      TabIndex        =   147
      Top             =   7200
      Width           =   1950
   End
   Begin VB.PictureBox PICPoint1 
      Height          =   240
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   540
      TabIndex        =   146
      Top             =   0
      Width           =   600
   End
   Begin VB.PictureBox PICOrder 
      BackColor       =   &H00808080&
      Height          =   11175
      Left            =   -45
      ScaleHeight     =   11115
      ScaleWidth      =   15210
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   15270
      Begin VB.PictureBox PICSaleOrderQueInformation 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   11670
         Left            =   -45
         ScaleHeight     =   11640
         ScaleWidth      =   26220
         TabIndex        =   120
         Top             =   -135
         Visible         =   0   'False
         Width           =   26250
         Begin MSComctlLib.ListView ListViewSaleOrderQueInformation 
            Height          =   6945
            Left            =   405
            TabIndex        =   134
            Top             =   630
            Width           =   13920
            _ExtentX        =   24553
            _ExtentY        =   12250
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
               Size            =   18
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
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "คิวที่"
               Object.Width           =   6703
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "โซนการจัดสินค้า"
               Object.Width           =   15875
            EndProperty
         End
         Begin VB.CommandButton CMDSaleOrderInformationClose 
            Caption         =   "ปิด"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   405
            TabIndex        =   122
            Top             =   7740
            Width           =   1410
         End
         Begin MSComctlLib.ListView ListViewSaleOrderLastQue 
            Height          =   6945
            Left            =   405
            TabIndex        =   121
            Top             =   630
            Width           =   13920
            _ExtentX        =   24553
            _ExtentY        =   12250
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
            MousePointer    =   1
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
               Text            =   "ลำดับ"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "คิวที่"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "คำอธิบาย"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "รหัสชื่อสินค้า"
               Object.Width           =   8820
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
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "ชื่อผู้จัดสินค้า"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "โซนการจัด"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "วันที่คิว"
               Object.Width           =   3175
            EndProperty
         End
         Begin VB.Label LBLSaleOrderQueInf 
            BackStyle       =   0  'Transparent
            Caption         =   "รายการ สินค้าที่ได้สั่งจัด"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   405
            TabIndex        =   135
            Top             =   270
            Width           =   5100
         End
      End
      Begin MSMask.MaskEdBox MEBReqTime 
         Height          =   330
         Left            =   6390
         TabIndex        =   96
         Top             =   1710
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TBOrderRefNo 
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
         Left            =   1755
         TabIndex        =   95
         Top             =   1710
         Width           =   2220
      End
      Begin VB.PictureBox PICEditOrder 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   4155
         Left            =   450
         ScaleHeight     =   4125
         ScaleWidth      =   11955
         TabIndex        =   104
         Top             =   2655
         Visible         =   0   'False
         Width           =   11985
         Begin VB.CommandButton CMDEditExit 
            BackColor       =   &H00808080&
            Caption         =   "ออก"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   10935
            Style           =   1  'Graphical
            TabIndex        =   119
            Top             =   2520
            Width           =   1635
         End
         Begin VB.CommandButton CMDEditOK 
            BackColor       =   &H00808080&
            Caption         =   "เปลี่ยน"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   9090
            Style           =   1  'Graphical
            TabIndex        =   118
            Top             =   2520
            Width           =   1635
         End
         Begin VB.TextBox TBEditQty 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   2070
            TabIndex        =   117
            Top             =   1755
            Width           =   1815
         End
         Begin VB.Label LBLEditIndex 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   7380
            TabIndex        =   126
            Top             =   1755
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label LBLEditRemain 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   10755
            TabIndex        =   125
            Top             =   1755
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label LBLEditItemAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   10755
            TabIndex        =   124
            Top             =   1170
            Width           =   1815
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "มูลค่าสินค้า :"
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
            Left            =   9720
            TabIndex        =   123
            Top             =   1170
            Width           =   960
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "สั่งจัดจำนวน :"
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
            Left            =   765
            TabIndex        =   116
            Top             =   1800
            Width           =   1230
         End
         Begin VB.Label LBLEditUnitCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4725
            TabIndex        =   115
            Top             =   1755
            Width           =   1500
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "หน่วยนับ :"
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
            Left            =   3915
            TabIndex        =   114
            Top             =   1755
            Width           =   780
         End
         Begin VB.Label LBLEditDiscount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   7380
            TabIndex        =   113
            Top             =   1170
            Width           =   2085
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ส่วนลด :"
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
            Left            =   6210
            TabIndex        =   112
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label LBLEditPrice 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4725
            TabIndex        =   111
            Top             =   1170
            Width           =   1500
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ราคา :"
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
            Left            =   3915
            TabIndex        =   110
            Top             =   1170
            Width           =   735
         End
         Begin VB.Label LBLEditItemQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2070
            TabIndex        =   109
            Top             =   1170
            Width           =   1815
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "จำนวน :"
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
            Left            =   720
            TabIndex        =   108
            Top             =   1215
            Width           =   1275
         End
         Begin VB.Label LBLEditItemName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3960
            TabIndex        =   107
            Top             =   585
            Width           =   8610
         End
         Begin VB.Label LBLEditItemCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2070
            TabIndex        =   106
            Top             =   585
            Width           =   1815
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "รหัส/ชื่อสินค้า :"
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
            Left            =   675
            TabIndex        =   105
            Top             =   585
            Width           =   1320
         End
      End
      Begin VB.CommandButton CMDSaleOrderSendQue 
         BackColor       =   &H00404040&
         Caption         =   "สั่งจัดสินค้า"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   6885
         Width           =   1950
      End
      Begin VB.CommandButton CMDSOMain 
         BackColor       =   &H00404040&
         Caption         =   "ออก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   6885
         Width           =   1950
      End
      Begin MSComctlLib.ListView ListViewSaleOrder 
         Height          =   4155
         Left            =   450
         TabIndex        =   97
         Top             =   2655
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   7329
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
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   3528
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
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วยนับ"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "ราคา/หน่วย"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "ส่วนลด/หน่วย"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "มูลค่าสินค้า"
            Object.Width           =   3175
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
            Text            =   "โซน"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "ที่เก็บ"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "บาร์โค้ด"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "ส่วนลด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "RemainOrder"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.TextBox TBDocNo 
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
         Left            =   1755
         TabIndex        =   89
         Top             =   360
         Width           =   2220
      End
      Begin VB.TextBox TBSOLastDisCount 
         Alignment       =   1  'Right Justify
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
         Left            =   12195
         TabIndex        =   132
         Top             =   8415
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "มูลค่ารวม :"
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
         Left            =   10890
         TabIndex        =   145
         Top             =   8055
         Width           =   1230
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการ สินค้าที่สั่งจัด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   450
         TabIndex        =   143
         Top             =   2295
         Width           =   1815
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ต้องการรับของเวลา :"
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
         Left            =   4905
         TabIndex        =   142
         Top             =   1710
         Width           =   1410
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ทะเบียนรถ :"
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
         Left            =   315
         TabIndex        =   141
         Top             =   1710
         Width           =   1365
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสสมาชิก :"
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
         Left            =   10575
         TabIndex        =   140
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label LBLOrderSoStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7920
         TabIndex        =   139
         Top             =   360
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label LBLOrderBillType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7380
         TabIndex        =   138
         Top             =   360
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label LBLOrderSendQue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8460
         TabIndex        =   137
         Top             =   360
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label LBLOrderMember 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11655
         TabIndex        =   93
         Top             =   810
         Width           =   1905
      End
      Begin VB.Label LBLOrderDiscountOld 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   12195
         TabIndex        =   136
         Top             =   7335
         Width           =   2220
      End
      Begin VB.Label LBLOrderNetAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   12195
         TabIndex        =   133
         Top             =   8055
         Width           =   2220
      End
      Begin VB.Label LBLOrderTaxAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   12195
         TabIndex        =   131
         Top             =   7695
         Width           =   2220
      End
      Begin VB.Label LBLOrderSumOfItemAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   12195
         TabIndex        =   130
         Top             =   6975
         Width           =   2220
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "มูลค่าส่วนลด :"
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
         Left            =   11070
         TabIndex        =   129
         Top             =   7335
         Width           =   1050
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "มูลค่าภาษี :"
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
         Left            =   11070
         TabIndex        =   128
         Top             =   7695
         Width           =   1050
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "มูลค่าสินค้า :"
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
         Left            =   11070
         TabIndex        =   127
         Top             =   6975
         Width           =   1050
      End
      Begin VB.Label LBLOrderSaleCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1755
         TabIndex        =   94
         Top             =   1260
         Width           =   5550
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รหัส/ชื่อพนักงานขาย :"
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
         Left            =   -315
         TabIndex        =   103
         Top             =   1260
         Width           =   1995
      End
      Begin VB.Label LBLOrderArName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4005
         TabIndex        =   92
         Top             =   810
         Width           =   6450
      End
      Begin VB.Label LBLOrderArCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1755
         TabIndex        =   91
         Top             =   810
         Width           =   2220
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "รหัส/ชื่อลูกค้า :"
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
         Left            =   630
         TabIndex        =   102
         Top             =   810
         Width           =   1050
      End
      Begin VB.Label LBLOrderDocDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5490
         TabIndex        =   90
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่เอกสาร :"
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
         Left            =   4320
         TabIndex        =   101
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบสั่งขาย/จอง :"
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
         Left            =   90
         TabIndex        =   100
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.PictureBox PTBPickingQueue 
      Height          =   9015
      Left            =   13995
      Picture         =   "Form311.frx":9673
      ScaleHeight     =   8955
      ScaleWidth      =   11970
      TabIndex        =   22
      Top             =   9225
      Visible         =   0   'False
      Width           =   12030
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2610
         ScaleHeight     =   345
         ScaleWidth      =   7905
         TabIndex        =   83
         ToolTipText     =   "เงื่อนไขส่วนนี้จะใช้เฉพาะเอกสารใบสั่งจองเท่านั้น เอกสารใบสั่งขายไม่มีผลต่อการเลือกเงื่อนไข"
         Top             =   990
         Width           =   7935
         Begin VB.OptionButton OPTPayItem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "สั่งจัดเพื่อ จ่ายสินค้า"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2925
            TabIndex        =   85
            Top             =   45
            Width           =   2805
         End
         Begin VB.OptionButton OPTReserve 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "สั่งจัดเพื่อ จองสินค้า"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   45
            TabIndex        =   84
            Top             =   45
            Value           =   -1  'True
            Width           =   2355
         End
      End
      Begin VB.OptionButton OPTPayItem1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "สั่งจัดเพื่อ จ่ายสินค้า"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2610
         TabIndex        =   80
         Top             =   7110
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.OptionButton OPTReserve1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "สั่งจัดเพื่อ จองสินค้า"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2610
         TabIndex        =   79
         Top             =   6705
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.CommandButton CMDSelectItemBack 
         Caption         =   "แก้ไขจำนวนจัด"
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
         Left            =   7785
         TabIndex        =   66
         Top             =   6750
         Width           =   1320
      End
      Begin VB.CommandButton CMDSendPicking 
         Caption         =   "ส่งคิวจัดสินค้า"
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
         Left            =   6345
         TabIndex        =   65
         Top             =   6750
         Width           =   1320
      End
      Begin VB.OptionButton OPTTomorrow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "เช้าวันพรุ่งนี้"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2610
         TabIndex        =   51
         Top             =   4590
         Width           =   1770
      End
      Begin VB.TextBox TextCarLicense 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6345
         TabIndex        =   49
         Top             =   6300
         Width           =   4200
      End
      Begin VB.CheckBox CHKLicense 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "กรณีกำหนด เลขทะเบียนรถขนส่ง"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2610
         TabIndex        =   48
         Top             =   6300
         Width           =   2760
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1230
         Left            =   2610
         ScaleHeight     =   1200
         ScaleWidth      =   7905
         TabIndex        =   44
         Top             =   4995
         Width           =   7935
         Begin VB.OptionButton OptOutLet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "จ่ายฝั่ง OutLet"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   0
            TabIndex        =   46
            Top             =   495
            Width           =   1860
         End
         Begin VB.OptionButton OptNormal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ออกใบหยิบปกติ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   0
            TabIndex        =   45
            Top             =   90
            Value           =   -1  'True
            Width           =   1725
         End
         Begin VB.OptionButton OptMain 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "จ่ายฝั่ง สำนักงานใหญ่"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   0
            TabIndex        =   47
            Top             =   810
            Width           =   1950
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "กรณีขายสินค้า ชั้นเก็บ SPO ต้องระบุการรับสินค้าของลูกค้า ระหว่างฝั่งสำนักงานกับฝั่ง Outlet เท่านั้น ถึงจะสามารถพิมพ์ใบหยิบได้"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   1095
            Left            =   1980
            TabIndex        =   86
            Top             =   45
            Width           =   5865
         End
      End
      Begin VB.CommandButton CMDExitQueue 
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
         Left            =   9225
         TabIndex        =   40
         Top             =   6750
         Width           =   1320
      End
      Begin VB.CommandButton CMDSaveQueue 
         Caption         =   "ส่งคิวจัดสินค้า"
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
         Left            =   135
         TabIndex        =   39
         Top             =   7695
         Visible         =   0   'False
         Width           =   1320
      End
      Begin MSMask.MaskEdBox MEDTime 
         Height          =   285
         Left            =   5580
         TabIndex        =   38
         Top             =   4185
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton OPTSchedule 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ตามเวลาที่กำหนด"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2610
         TabIndex        =   36
         ToolTipText     =   "การเลือก ประเภทนี้ต้องกำหนดเวลาที่ต้องการสินค้ามากกว่าเวลาปัจจุบันขึ้นไป"
         Top             =   4185
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton OPTNow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ด่วน"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2610
         TabIndex        =   35
         ToolTipText     =   "การเลือก ประเภทนี้ จะเพิ่มเวลาจากเวลาปัจจุบันไปอีก 10 นาที"
         Top             =   3780
         Width           =   1770
      End
      Begin VB.Label LBLReserveDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2610
         TabIndex        =   87
         Top             =   630
         Width           =   7935
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกเฉพาะ ระบบจอง :"
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
         Height          =   240
         Left            =   675
         TabIndex        =   78
         Top             =   1035
         Width           =   1860
      End
      Begin VB.Label LBLDueDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8685
         TabIndex        =   77
         Top             =   4185
         Width           =   1860
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ครบกำหนด :"
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
         Height          =   330
         Left            =   7290
         TabIndex        =   76
         Top             =   4185
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทสั่งจอง :"
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
         Left            =   7335
         TabIndex        =   74
         Top             =   3780
         Width           =   1275
      End
      Begin VB.Label LBLSaleType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8685
         TabIndex        =   73
         Top             =   3780
         Width           =   1860
      End
      Begin VB.Label LBLBillType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8865
         TabIndex        =   72
         Top             =   1485
         Width           =   1680
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทบิล :"
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
         Left            =   7830
         TabIndex        =   71
         Top             =   1485
         Width           =   1050
      End
      Begin VB.Label LBLIsConditionSend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5580
         TabIndex        =   70
         Top             =   3780
         Width           =   1230
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "วิธีจัดส่ง :"
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
         Height          =   240
         Left            =   4500
         TabIndex        =   69
         Top             =   3780
         Width           =   1005
      End
      Begin VB.Label LBLSaleName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5580
         TabIndex        =   68
         Top             =   3285
         Width           =   4965
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ :"
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
         Left            =   5355
         TabIndex        =   50
         Top             =   6300
         Width           =   915
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้ารับสินค้า ณ จุด :"
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
         Left            =   675
         TabIndex        =   43
         Top             =   5085
         Width           =   1815
      End
      Begin VB.Label LBLSumQTY 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10890
         TabIndex        =   42
         Top             =   1935
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label LBLCountItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10890
         TabIndex        =   41
         Top             =   1485
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เวลา :"
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
         Left            =   4905
         TabIndex        =   37
         Top             =   4185
         Width           =   600
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เวลาที่ต้องการสินค้า :"
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
         Left            =   765
         TabIndex        =   34
         Top             =   3780
         Width           =   1770
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย :"
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
         Left            =   1350
         TabIndex        =   33
         Top             =   3285
         Width           =   1185
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ที่อยู่ลูกค้า :"
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
         Left            =   1350
         TabIndex        =   32
         Top             =   2385
         Width           =   1185
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รหัส-ชื่อลูกค้า :"
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
         Left            =   1350
         TabIndex        =   31
         Top             =   1935
         Width           =   1185
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่เอกสาร :"
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
         Left            =   4320
         TabIndex        =   30
         Top             =   1485
         Width           =   1185
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่เอกสาร :"
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
         Left            =   1350
         TabIndex        =   29
         Top             =   1485
         Width           =   1185
      End
      Begin VB.Label LBLSaleCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2610
         TabIndex        =   28
         Top             =   3285
         Width           =   1770
      End
      Begin VB.Label LBLARAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   2610
         TabIndex        =   27
         Top             =   2385
         Width           =   7935
      End
      Begin VB.Label LBLARName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2610
         TabIndex        =   26
         Top             =   1935
         Width           =   7935
      End
      Begin VB.Label LBLDocDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5580
         TabIndex        =   25
         Top             =   1485
         Width           =   1815
      End
      Begin VB.Label LBLDocNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2610
         TabIndex        =   24
         Top             =   1485
         Width           =   1770
      End
   End
   Begin VB.PictureBox PICSelectPrintSlip 
      Height          =   9015
      Left            =   12735
      Picture         =   "Form311.frx":1096E
      ScaleHeight     =   8955
      ScaleWidth      =   11970
      TabIndex        =   52
      Top             =   9225
      Visible         =   0   'False
      Width           =   12030
      Begin VB.CommandButton CMDPickingCancel 
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
         Left            =   9810
         TabIndex        =   63
         Top             =   6570
         Width           =   1365
      End
      Begin VB.CommandButton CMDPickingOK 
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
         Left            =   8145
         TabIndex        =   62
         Top             =   6570
         Width           =   1365
      End
      Begin VB.CheckBox CHKSelectAllItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "เลือกทั้งหมดตามจำนวนในใบสั่งขาย"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   405
         Picture         =   "Form311.frx":17C69
         TabIndex        =   54
         Top             =   2205
         Value           =   1  'Checked
         Width           =   2850
      End
      Begin MSComctlLib.ListView ListViewSelectItemPicking 
         Height          =   3705
         Left            =   405
         TabIndex        =   53
         Top             =   2520
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6535
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
         NumItems        =   8
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
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "จำนวนคงค้าง"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "จำนวนสั่งหยิบ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "หน่วย"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "คลัง"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "สโตร์"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label LBLRefDocDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5040
         TabIndex        =   61
         Top             =   1215
         Width           =   1905
      End
      Begin VB.Label LBLRefARCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1530
         TabIndex        =   60
         Top             =   1620
         Width           =   1725
      End
      Begin VB.Label LBLRefDocNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1530
         TabIndex        =   59
         Top             =   1215
         Width           =   1725
      End
      Begin VB.Label LBLRefARName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5040
         TabIndex        =   58
         Top             =   1620
         Width           =   6540
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสลูกค้า :"
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
         Left            =   405
         TabIndex        =   57
         Top             =   1665
         Width           =   1410
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่เอกสาร :"
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
         Left            =   3915
         TabIndex        =   56
         Top             =   1215
         Width           =   1590
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่เอกสาร :"
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
         Left            =   405
         TabIndex        =   55
         Top             =   1215
         Width           =   1410
      End
   End
   Begin VB.CheckBox CHKReqPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ขอพิมพ์ฟอร์ม A4 กรณีพิมพ์กระดาษครึ่งหน้าไม่ได้"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7425
      TabIndex        =   75
      Top             =   7470
      Width           =   3795
   End
   Begin VB.CommandButton CMDReqPicking 
      Caption         =   "จัดสินค้า"
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
      Left            =   4095
      TabIndex        =   64
      Top             =   7200
      Width           =   1365
   End
   Begin VB.CommandButton BTNPickingQueue 
      Caption         =   "คิวจัดสินค้า"
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
      Left            =   4590
      TabIndex        =   23
      Top             =   9315
      Visible         =   0   'False
      Width           =   195
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   2115
      Top             =   9855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.PictureBox Pic101 
      Height          =   4200
      Left            =   1350
      ScaleHeight     =   4140
      ScaleWidth      =   11205
      TabIndex        =   13
      Top             =   2295
      Visible         =   0   'False
      Width           =   11265
      Begin VB.CommandButton CMD103 
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
         Height          =   330
         Left            =   5490
         TabIndex        =   19
         Top             =   2295
         Width           =   825
      End
      Begin VB.CommandButton CMD102 
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
         Height          =   330
         Left            =   4500
         TabIndex        =   18
         Top             =   2295
         Width           =   825
      End
      Begin VB.TextBox Text101 
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
         IMEMode         =   3  'DISABLE
         Left            =   1305
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   2295
         Width           =   2940
      End
      Begin VB.Label Label6 
         Caption         =   "รหัสผ่าน :"
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
         Left            =   450
         TabIndex        =   16
         Top             =   2340
         Width           =   915
      End
      Begin VB.Label LBL101 
         Caption         =   "xxxxxx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   405
         TabIndex        =   15
         Top             =   495
         Width           =   10455
      End
      Begin VB.Label Label4 
         Caption         =   "คำอธิบาย :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   405
         TabIndex        =   14
         Top             =   225
         Width           =   7575
      End
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "จัดคิวขนส่ง"
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
      Left            =   11295
      TabIndex        =   12
      Top             =   7830
      Width           =   1365
   End
   Begin Crystal.CrystalReport CrysPicking1 
      Left            =   1620
      Top             =   9855
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
   Begin Crystal.CrystalReport CrysPicking 
      Left            =   1125
      Top             =   9855
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
   Begin VB.CommandButton CMDRefresh 
      Height          =   330
      Left            =   4185
      Picture         =   "Form311.frx":1EF64
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "ฟื้นฟูข้อมูล รายการเอกสาร"
      Top             =   1350
      Width           =   350
   End
   Begin VB.ComboBox CMBSale 
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
      Height          =   315
      Left            =   1350
      TabIndex        =   9
      Text            =   "CMBSale"
      ToolTipText     =   "กดลูกศรเลือก รหัสพนักงานที่ต้องการพิมพ์เอกสาร"
      Top             =   1350
      Width           =   2790
   End
   Begin VB.CommandButton CMDClearDocuments 
      Height          =   315
      Left            =   12330
      Picture         =   "Form311.frx":1F2BB
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1980
      Width           =   315
   End
   Begin VB.CheckBox ShowDiscount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ไม่โชว์ส่วนลด"
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
      Height          =   240
      Left            =   9450
      TabIndex        =   6
      ToolTipText     =   "พิมพ์เอกสาร ใบสั่งขายที่ไม่แสดงส่วนลดในเอกสาร"
      Top             =   7200
      Width           =   1770
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9090
      Top             =   9810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin VB.CommandButton CMDPicking_Print 
      Caption         =   "พิมพ์ใบสั่งขาย"
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
      Left            =   11295
      TabIndex        =   2
      Top             =   7200
      Width           =   1365
   End
   Begin VB.ComboBox CMBPicking1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9855
      TabIndex        =   1
      ToolTipText     =   "เลือกฟอร์มที่ต้องการพิมพ์"
      Top             =   6660
      Width           =   2790
   End
   Begin MSComctlLib.ListView ListViewPicking 
      Height          =   4200
      Left            =   1350
      TabIndex        =   0
      ToolTipText     =   "คลิ๊กเลือกเอกสารที่ต้องการพิมพ์"
      Top             =   2295
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   7408
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
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
         Text            =   "เลขที่ใบสั่งขาย"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่ครบกำหนด"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันที่เวลาเอกสาร"
         Object.Width           =   3704
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PBSaleOrder 
      Height          =   195
      Left            =   1350
      TabIndex        =   67
      Top             =   1755
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   240
      Left            =   3645
      TabIndex        =   81
      Top             =   9630
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   240
      Left            =   4005
      TabIndex        =   82
      Top             =   9630
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox PICPoint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   255
      TabIndex        =   144
      Top             =   0
      Width           =   285
   End
   Begin VB.Label TXTPicking1 
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2700
      TabIndex        =   21
      Top             =   6660
      Width           =   2760
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการเอกสารที่ยังไม่ได้พิมพ์"
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
      Height          =   240
      Left            =   1350
      TabIndex        =   20
      Top             =   2025
      Width           =   2805
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ลบเอกสารที่ไม่ต้องการพิมพ์"
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
      Left            =   10035
      TabIndex        =   10
      Top             =   2070
      Width           =   2265
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกพนักงานขาย"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1350
      TabIndex        =   7
      Top             =   1035
      Width           =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์ใบสั่งขาย/ใบหยิบสินค้า/ใบจัดคิวส่งสินค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   2625
      TabIndex        =   5
      Top             =   300
      Width           =   9285
   End
   Begin VB.Label LBLPicking2 
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์"
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
      Height          =   240
      Left            =   8775
      TabIndex        =   4
      Top             =   6705
      Width           =   1065
   End
   Begin VB.Label LBLPicking1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบสั่งขาย"
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
      Height          =   240
      Left            =   1350
      TabIndex        =   3
      Top             =   6705
      Width           =   1290
   End
End
Attribute VB_Name = "Form311"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vUserPrint As String
Dim vCheckValue As Boolean
Dim vCheckValue1 As Boolean
Dim vKeyword As String
Dim vCheckKeyword As String
Dim vCheckPic101 As Integer
Dim vARCode As String
Dim vQueueID As Integer
Dim vSaleCode As String
Dim vCustomerZone As Integer

Dim vCheckSelectItemPickBack As Integer

Dim vSOCountNumber As Integer


Private Sub BTNPickingQueue_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vQueueNo As String
Dim vQueueStatus As Integer
Dim vQueueDate As Date
Dim vCheckAnswerPrint As Integer

On Error GoTo ErrDescription

If Me.TXTPicking1.Caption <> "" Then
   vDocNo = Trim(Me.TXTPicking1.Caption)
  
  'vQuery = "select top 1  docno,queuedatetime,status from npmaster.dbo.TB_NP_QueueManagement_Test where saleorderno = '" & vDocNo & "' order by queuedatetime desc"
  vQuery = "exec dbo.USP_NP_SearchCheckQueue '" & vDocNo & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     vQueueNo = vRecordset.Fields("docno").Value
     vQueueDate = vRecordset.Fields("queuedatetime").Value
     vQueueStatus = vRecordset.Fields("status").Value
  End If
  vRecordset.Close
  
  If vQueueNo <> "" Then
     If vQueueStatus = 0 Then
       MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ รอจัดสินค้า ต้องการพิมพ์ใหม่ ให้พิมพ์ที่หน้าทดแทนเอกสาร", vbInformation, "Send Information"
     ElseIf vQueueStatus = 1 Then
       MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ กำลังจัดสินค้า ต้องการพิมพ์ใหม่ ให้พิมพ์ที่หน้าทดแทนเอกสาร", vbInformation, "Send Information"
     ElseIf vQueueStatus = 2 Then
       MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ จัดสินค้าเรียบร้อยแล้ว ต้องการพิมพ์ใหม่ ให้พิมพ์ที่หน้าทดแทนเอกสาร", vbInformation, "Send Information"
     End If
     TXTPicking1.Caption = ""
     CMBPicking1.Text = ""
     vCheckPic101 = 0
     Exit Sub
  End If
  
  vQuery = "exec dbo.USP_SO_SaleOrderDetails '" & vDocNo & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.LBLDocNo.Caption = Trim(vRecordset.Fields("docno").Value)
    Me.LBLDocDate.Caption = Trim(vRecordset.Fields("docdate").Value)
    Me.LBLARName.Caption = Trim(vRecordset.Fields("arname").Value)
    Me.LBLARAddress.Caption = Trim(vRecordset.Fields("workaddress").Value)
    Me.LBLSaleName.Caption = Trim(vRecordset.Fields("salename").Value)
    Me.LBLCountItem.Caption = vRecordset.Fields("CountItem").Value
    Me.LBLSumQTY.Caption = vRecordset.Fields("SumRemainQTY").Value
  
    If Len(Hour(Now)) = 1 Then
      If Len(Minute(Now)) = 1 Then
        MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Trim("0" & Minute(Now))
      Else
        MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Minute(Now)
      End If
    Else
      If Len(Minute(Now)) = 1 Then
        MEDTime.Text = Hour(Now) & ":" & Trim("0" & Minute(Now))
      Else
        MEDTime.Text = Hour(Now) & ":" & Minute(Now)
      End If
    End If
    Me.OptNow.Value = True
    Me.OptNormal.Value = True
    PTBPickingQueue.Visible = True
    Me.CHKLicense.Value = 0
  Else
    MsgBox "กรุณาตรวจสอบเลขที่ใบสั่งขาย/จอง ที่จะเข้าคิวจัดสินค้า เพราะเอกสารดังกล่าวไม่สมารถทำคิวรอจัดสินค้าได้ ", vbCritical, "Send Error Message"
  End If
  vRecordset.Close

Else
  MsgBox "กรุณาเลือกเลขที่ใบสั่งขาย/จอง ที่จะเข้าคิวจัดสินค้า", vbCritical, "Send Error Message"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub CHKLicense_Click()

On Error Resume Next

If Me.CHKLicense.Value = 1 Then
  Me.TextCarLicense.Enabled = True
  Me.TextCarLicense.SetFocus
Else
  Me.TextCarLicense.Enabled = False
End If
End Sub

Private Sub CHKSelectAllItem_Click()
Dim i As Integer

On Error Resume Next

If Me.ListViewSelectItemPicking.ListItems.Count > 0 Then
   If Me.CHKSelectAllItem.Value = 1 Then
      For i = 1 To Me.ListViewSelectItemPicking.ListItems.Count
              Me.ListViewSelectItemPicking.ListItems(i).Checked = True
      Next
   Else
      For i = 1 To Me.ListViewSelectItemPicking.ListItems.Count
              Me.ListViewSelectItemPicking.ListItems(i).Checked = False
      Next
   End If
End If
End Sub

Private Sub CMBSale_Click()
Call RefreshData
End Sub

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vSend As Integer
Dim vSaleCode As String

On Error GoTo ErrDescription

'vDocNo = Trim(Me.TXTPicking1.Caption)
'vQuery = "exec dbo.USP_SO_SOStatus '" & vDocNo & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '  vSend = Trim(vRecordset.Fields("isconditionsend").Value)
  ' vSaleCode = Trim(vRecordset.Fields("salecode").Value)
'End If
'vRecordset.Close

'If vSaleCode = "" Then
 '  MsgBox "ไม่ได้ระบุรหัสพนักงานขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message "
'End If

'If vSend = 1 Then
   vTempDocno = Trim(Me.TXTPicking1.Caption)
   Form312.Show
'Else
   'MsgBox "เอกสารขาย ที่ประเภทเป็น ลูกค้ารับเอง ไม่สามารถทำใบจัดคิวขนส่งได้", vbCritical, "Send Massage"
'End If
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD102_Click()

On Error Resume Next

vCheckKeyword = Trim(Text101.Text)
vCheckPic101 = 1
Pic101.Visible = False
Call CMDPicking_Print_Click
Text101.Text = ""
End Sub

Private Sub CMD103_Click()
vCheckKeyword = Trim(Text101.Text)
vCheckPic101 = 0
Pic101.Visible = False
Text101.Text = ""
TXTPicking1.Caption = ""
CMBPicking1.Text = ""
End Sub

Private Sub CMDClearDocuments_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String

On Error Resume Next

If TXTPicking1.Caption <> "" Then
  vDocNo = Trim(TXTPicking1.Caption)
  vQuery = "exec dbo.USP_SO_UpdatePrintStatus '" & vDocNo & "' "
  gConnection.Execute vQuery
  TXTPicking1.Caption = ""
  Call RefreshData
Else
  MsgBox "กรุณากดเลือกเอกสารที่ต้องการยกเลิกการพิมพ์", vbCritical, "Send Error"
End If
End Sub

Private Sub CMDEditExit_Click()
   Me.LBLEditItemCode.Caption = ""
   Me.LBLEditItemName.Caption = ""
   Me.LBLEditItemQty.Caption = ""
   Me.LBLEditUnitCode.Caption = ""
   Me.LBLEditDiscount.Caption = ""
   Me.LBLEditIndex.Caption = ""
   Me.LBLEditItemAmount.Caption = ""
   Me.LBLEditPrice.Caption = ""
   Me.LBLEditRemain.Caption = ""
   Me.LBLEditItemQty.Caption = ""
   Me.TBEditQty.Text = ""
   Me.TBDocNo.Enabled = True
   Me.TBOrderRefNo.Enabled = True
   Me.MEBReqTime.Enabled = True
   
   Me.PICEditOrder.Visible = False
   Call CalcEditItemQty
   Me.ListViewSaleOrder.SetFocus
End Sub

Public Sub CalcEditItemQty()
Dim i As Integer
Dim vAmount As Double
Dim vNetAmount As Double
Dim vSumOfItemAmount As Double
Dim vTaxAmount As Double
Dim vTotalAmount As Double
Dim vLastDisCountAmount As Double

If Me.ListViewSaleOrder.ListItems.Count > 0 Then
   For i = 1 To Me.ListViewSaleOrder.ListItems.Count
       vAmount = Me.ListViewSaleOrder.ListItems(i).SubItems(7)
       vNetAmount = vNetAmount + vAmount
   Next i
   vSumOfItemAmount = vNetAmount
   If Me.TBSOLastDisCount.Text <> "" Then
      vLastDisCountAmount = Me.TBSOLastDisCount.Text
   End If
   vTotalAmount = vNetAmount - vLastDisCountAmount
   vTaxAmount = vTotalAmount - ((vTotalAmount * 100) / 107)
   
   Me.LBLOrderSumOfItemAmount.Caption = Format(vSumOfItemAmount, "##,##0.00")
   Me.LBLOrderTaxAmount.Caption = Format(vTaxAmount, "##,##0.00")
   Me.LBLOrderNetAmount.Caption = Format(vTotalAmount, "##,##0.00")
Else
   Me.LBLOrderSumOfItemAmount.Caption = Format(0, "##,##0.00")
   Me.LBLOrderTaxAmount.Caption = Format(0, "##,##0.00")
   Me.LBLOrderNetAmount.Caption = Format(0, "##,##0.00")
End If
End Sub

Private Sub CMDEditOK_Click()
Dim vIndex As Integer
Dim vQTY As Double
Dim vEditQty As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vAmount As Double

If Me.TBEditQty.Text <> "" Then
   vIndex = Me.LBLEditIndex.Caption
   vEditQty = Me.TBEditQty.Text
   If vEditQty = 0 Then
      MsgBox "ไม่สามารถกรอกจำนวนที่จะสั่งจัด เท่ากับ 0 กรุณาตรวจสอบ กรณีไม่สั่งจัด ก็ให้ลบรายการออกจากเอกสาร", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
   vQTY = Me.LBLEditRemain.Caption
   
   If vEditQty > vQTY Then
      MsgBox "ไม่สามารถกรอกจำนวนที่จะสั่งจัด มากกว่าจำนวนคงเหลือในใบสั่งขาย/จอง", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
   vPrice = Me.LBLEditPrice.Caption
   vDiscountAmount = Me.LBLEditDiscount.Caption
   vAmount = vEditQty * (vPrice - vDiscountAmount)
   Me.LBLEditItemAmount.Caption = Format(vAmount, "##,##0.00")
   
   Me.ListViewSaleOrder.ListItems(vIndex).SubItems(3) = Format(vEditQty, "##,##0.00")
   Me.ListViewSaleOrder.ListItems(vIndex).SubItems(7) = Format(vAmount, "##,##0.00")
   
   Me.LBLEditItemCode.Caption = ""
   Me.LBLEditItemName.Caption = ""
   Me.LBLEditItemQty.Caption = ""
   Me.LBLEditUnitCode.Caption = ""
   Me.LBLEditDiscount.Caption = ""
   Me.LBLEditIndex.Caption = ""
   Me.LBLEditItemAmount.Caption = ""
   Me.LBLEditPrice.Caption = ""
   Me.LBLEditRemain.Caption = ""
   Me.LBLEditItemQty.Caption = ""
   Me.TBEditQty.Text = ""
   Me.PICEditOrder.Visible = False
   Me.TBDocNo.Enabled = True
   Me.TBOrderRefNo.Enabled = True
   Me.MEBReqTime.Enabled = True
   
   Call CalcEditItemQty
   
   Me.ListViewSaleOrder.SetFocus
      
   If vIndex < Me.ListViewSaleOrder.ListItems.Count Then
   Me.ListViewSaleOrder.ListItems(vIndex + 1).Selected = True
   Else
   Me.ListViewSaleOrder.ListItems(vIndex).Selected = True
   End If


End If
End Sub

Private Sub CMDEditOK_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
CMDEditExit_Click
End If
End Sub

Private Sub CMDExitQueue_Click()
Me.PTBPickingQueue.Visible = False
Me.OptNow.Value = True
Me.OptNormal.Value = True
End Sub

Public Sub CMDPicking_Print_Click()
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String, vQuery As String
Dim vShelfGroup(10) As String
Dim i As Integer, vPrint As Integer
Dim vBillStatus As String
Dim n As Integer, vBillType As Integer
Dim vSend As Integer
Dim vIsConfirmPrint As Integer
Dim vOverDue As Integer
Dim vCheckAVLShelf As Integer
Dim vCheckAVLRemain As Integer
Dim vCheckBillType As Integer
Dim vCheckIsCompleteSave As Integer

On Error GoTo ErrDescription

vDocNo = Trim(TXTPicking1.Caption)
vTempDocno = vDocNo

vQuery = "exec dbo.USP_BC_CheckIsCompleteSave 'SO','" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vCheckIsCompleteSave = vRecordset.Fields("iscompletesave").Value
End If
vRecordset.Close

If vCheckIsCompleteSave = 0 Then
   MsgBox "กรุณารอให้เอกสารบันทึกข้อมูลให้เรียบร้อย อีกสักครู่", vbInformation, "Send Information Message"
   Exit Sub
End If


vQuery = "exec dbo.USP_SO_SOStatus '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vBillStatus = Trim(vRecordset.Fields("billstatus").Value)
   vBillType = Trim(vRecordset.Fields("billtype").Value)
   vSend = Trim(vRecordset.Fields("isconditionsend").Value)
   vSaleCode = Trim(vRecordset.Fields("salecode").Value)
End If
vRecordset.Close

If vSaleCode = "" Then
   MsgBox "ไม่ได้ระบุรหัสพนักงานขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message "
   Exit Sub
End If


Call GetComputerandUser

vQuery = "exec dbo.USP_SL_00002 '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vPrint = Trim(vRecordset.Fields("printed").Value)
vARCode = Trim(vRecordset.Fields("code").Value)
End If
vRecordset.Close
    

If TXTPicking1.Caption <> "" And CMBPicking1.Text <> "" Then
  If CMBPicking1.Text = "พิมพ์ใบสั่งขาย" Then
  
     Call SaleOrder
    
  ElseIf CMBPicking1.Text = "พิมพ์ใบสั่งขาย+พิมพ์ใบจัดคิวสินค้า" Then
  
    Call SaleOrder
        
    If vSend = 1 Then
    Call SaleOrder_Delivery
    Else
    MsgBox "เอกสารขาย ที่ประเภทเป็น ลูกค้ารับเอง ไม่สามารถพิมพ์ใบจัดคิวส่งสินค้าได้", vbCritical, "Send Massage"
    End If
  ElseIf CMBPicking1.Text = "พิมพ์ใบจัดคิวสินค้า" Then
    vQuery = "exec dbo.USP_SO_SOStatus '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vBillStatus = Trim(vRecordset.Fields("billstatus").Value)
      vBillType = Trim(vRecordset.Fields("billtype").Value)
      vSend = Trim(vRecordset.Fields("isconditionsend").Value)
      vSaleCode = Trim(vRecordset.Fields("salecode").Value)
    End If
    vRecordset.Close
    If vSend = 1 Then
      Call SaleOrder_Delivery
    Else
      MsgBox "เอกสารขาย ที่ประเภทเป็น ลูกค้ารับเอง ไม่สามารถพิมพ์ใบจัดคิวส่งสินค้าได้", vbCritical, "Send Massage"
    End If
    
  ElseIf CMBPicking1.Text = "พิมพ์ใบสั่งจองสินค้า" Then
  
    vQuery = "exec dbo.usp_so_CheckConfirmPrint '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vIsConfirmPrint = Trim(vRecordset.Fields("isconfirmprint").Value)
    End If
    vRecordset.Close
    
    If vIsConfirmPrint = 0 Then
      vQuery = "exec dbo.usp_so_CheckOverdue '" & vDocNo & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vOverDue = Trim(vRecordset.Fields("doccount").Value)
      End If
      vRecordset.Close
    End If
    
    If vOverDue = 0 And vIsConfirmPrint = 0 Then
      vQuery = "exec dbo.usp_so_SearchKeyword '01' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vKeyword = Trim(vRecordset.Fields("keyword").Value)
      End If
      vRecordset.Close
    If vCheckPic101 = 0 Then
      LBL101.Caption = Trim("กรุณาใส่รหัสผ่าน เพราะใบสั่งจองเลขที่ " & vDocNo & " วันที่ครบกำหนดเกิน 15 วัน")
      Pic101.Visible = True
      Text101.SetFocus
      Exit Sub
    End If
    If vKeyword <> vCheckKeyword Then
      If vCheckKeyword <> "" Then
      MsgBox "รหัสผ่านไม่ถูกต้อง", vbCritical, "Send Error"
      End If
      vCheckPic101 = 0
      Exit Sub
    End If
    End If
    
    'vQuery = "select isnull(billtype,0) as billtype from dbo.bcsaleorder where docno = '" & vDocNo & "' "
    'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     '   vCheckBillType = vRecordset.Fields("billtype").Value
    'End If
    'vRecordset.Close
    
    If vCheckBillType = 0 Then
       'vQuery = "exec dbo.USP_SO_CheckQTYReserve '" & vDocNo & "' "
       'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          'vRecordset.MoveFirst
          'While Not vRecordset.EOF
          'MsgBox "สินค้ารหัส " & vRecordset.Fields("itemcode").Value & " มียอดในชั้นเก็บ AVL ไม่พอขาย ต้องทำเอกสาร BackOrder เพื่อสั่งซื้อสินค้าเพิ่ม", vbCritical, "Send Error Message"
          'vRecordset.MoveNext
         ' Wend
        '  vCheckAVLRemain = 1
       'End If
       'vRecordset.Close
       '--------------------------------------------------------------------------------------------
       'If vCheckAVLRemain = 0 Then
          Call SaleOrder_Reserve
       'Else
        '  Exit Sub
       'End If
         '--------------------------------------------------------------------------------------------
    Else
       Call SaleOrder_Reserve
    End If
     
Else
  MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
End If

End If

TXTPicking1.Caption = ""
CMBPicking1.Text = ""
vCheckPic101 = 0

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub CMDPickingCancel_Click()
PICSelectPrintSlip.Visible = False
End Sub

Private Sub CMDPickingOK_Click()
Dim vDocNo As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vCountPickItem As Double
Dim vTotalCount As Double
   
   
   If Me.ListViewSelectItemPicking.ListItems.Count > 0 Then
      vTotalCount = 0
      For i = 1 To Me.ListViewSelectItemPicking.ListItems.Count
         If Me.ListViewSelectItemPicking.ListItems(i).Checked = True Then
             vCountPickItem = Me.ListViewSelectItemPicking.ListItems(i).SubItems(4)
             vTotalCount = vTotalCount + vCountPickItem
         End If
      Next
   End If
   
   If vTotalCount > 0 Then
      Me.PTBPickingQueue.Visible = True
   Else
      MsgBox "ไม่มีรายการสินค้าที่จะให้จัดสินค้า", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
   vDocNo = Me.LBLRefDocNo.Caption
   
   If vCheckSelectItemPickBack = 0 Then
   vQuery = "exec dbo.USP_SO_SaleOrderDetails '" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     Me.LBLDocNo.Caption = Trim(vRecordset.Fields("docno").Value)
     Me.LBLDocDate.Caption = Trim(vRecordset.Fields("docdate").Value)
     Me.LBLARName.Caption = Trim(vRecordset.Fields("arname").Value)
     Me.LBLARAddress.Caption = Trim(vRecordset.Fields("workaddress").Value)
     Me.LBLSaleCode.Caption = Trim(vRecordset.Fields("salecode").Value)
     Me.LBLSaleName.Caption = Trim(vRecordset.Fields("salename").Value)
     Me.LBLCountItem.Caption = vRecordset.Fields("CountItem").Value
     Me.LBLSumQTY.Caption = vRecordset.Fields("SumRemainQTY").Value
     Me.LBLBillType.Caption = vRecordset.Fields("billtype").Value
     Me.LBLSaleType.Caption = vRecordset.Fields("sostatus").Value
     Me.LBLDueDate.Caption = vRecordset.Fields("deliverydate").Value
     Me.LBLIsConditionSend.Caption = vRecordset.Fields("isconditionsend").Value
        
     If Len(Hour(Now)) = 1 Then
       If Len(Minute(Now)) = 1 Then
         MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Trim("0" & Minute(Now))
       Else
         MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Minute(Now)
       End If
     Else
       If Len(Minute(Now)) = 1 Then
         MEDTime.Text = Hour(Now) & ":" & Trim("0" & Minute(Now))
       Else
         MEDTime.Text = Hour(Now) & ":" & Minute(Now)
       End If
     End If
     Me.OptNow.Value = True
     Me.OptNormal.Value = True
     PTBPickingQueue.Visible = True
     Me.CHKLicense.Value = 0
   Else
     MsgBox "กรุณาตรวจสอบเลขที่ใบสั่งขาย/จอง ที่จะเข้าคิวจัดสินค้า เพราะเอกสารดังกล่าวไม่สมารถทำคิวรอจัดสินค้าได้ ", vbCritical, "Send Error Message"
   End If
   vRecordset.Close
   
   Dim vCheckPickStatus As Integer
   Dim vCheckCountPick As Integer
   
   vQuery = "exec dbo.USP_NP_SearchCheckPickStatus '" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckPickStatus = Trim(vRecordset.Fields("pickstatus").Value)
      vCheckCountPick = Trim(vRecordset.Fields("socountnumber").Value)
   Else
      vCheckPickStatus = 0
      vCheckCountPick = 0
   End If
   vRecordset.Close
      
  If Me.LBLSaleType.Caption = 1 Then
     If vCheckCountPick = 0 Then
        Me.OPTReserve.Value = True
        Me.OPTPayItem.Value = False
     End If
     
     If vCheckCountPick > 0 Then
        If vCheckPickStatus = 0 Then
           Me.OPTReserve.Value = False
           Me.OPTPayItem.Value = True
        Else
           Me.OPTReserve.Value = True
           Me.OPTPayItem.Value = False
        End If
     End If
     
  Else
    Me.OPTReserve.Value = False
    Me.OPTPayItem.Value = True
  End If
  
   End If
   
   vCheckSelectItemPickBack = 0
End Sub

Private Sub CMDRefresh_Click()
   Call RefreshData
End Sub

Private Sub CMDSaveQueue1()
'Dim vRecordset As New ADODB.Recordset
'Dim vQuery As String
'Dim vCheckDate As Date
'Dim vHour As Integer
'Dim vMinute As Integer
'Dim vSaleOrderNo As String
'Dim vSaleOrderDate As Date
'Dim vDocdate As Date
''Dim vARCode As String
'Dim vSaleCode As String
'Dim vRequestDate As Date
'Dim vRequestTime As String
'Dim vRequestStatus As Integer
'Dim vRequestCountItem As Double
'Dim vRequestCountQTY As Double
'Dim vRequestAt As Integer

'If Me.LBLDocNo.Caption <> "" Then
 ' If Me.OPTSchedule.Value = True Then
  '  If DateDiff("d", Now, vCheckDate) = 0 Then
   '   vHour = Left(Trim(Me.MEDTime.Text), 2)
    '  vMinute = Right(Trim(Me.MEDTime.Text), 2)
     ' If vHour < Hour(Now) Then
      '  MsgBox "ไม่สามารถกำหนดเวลาต้องการสินค้า ณ เวลาที่ผ่านมาแล้วได้"
       ' Exit Sub
      'ElseIf vHour = Hour(Now) Then
       ' If vMinute < Minute(Now) Then
        '  MsgBox "ไม่สามารถกำหนดเวลาต้องการสินค้า ณ เวลาที่ผ่านมาแล้วได้"
         ' Exit Sub
        'End If
      'End If
    'End If
    'vRequestDate = Day(Me.DTPDate.Value) & "/" & Month(Me.DTPDate.Value) & "/" & Year(Me.DTPDate.Value)
    'vRequestTime = Me.MEDTime.Text
    'vRequestStatus = 1
  'ElseIf Me.OPTNow.Value = True Then
   ' vRequestDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    
    'If Len(Hour(Now)) = 1 And Len(Minute(Now)) = 1 Then
     ' vRequestTime = Trim("0" & Hour(Now)) & ":" & Trim("0" & Minute(Now))
    'ElseIf Len(Hour(Now)) > 1 And Len(Minute(Now)) = 1 Then
     ' vRequestTime = Hour(Now) & ":" & Trim("0" & Minute(Now))
    'ElseIf Len(Hour(Now)) = 1 And Len(Minute(Now)) > 1 Then
     ' vRequestTime = Trim("0" & Hour(Now)) & ":" & Minute(Now)
    'ElseIf Len(Hour(Now)) > 1 And Len(Minute(Now)) > 1 Then
     ' vRequestTime = Hour(Now) & ":" & Minute(Now)
    'End If
    'vRequestStatus = 0
  'End If
   ' vSaleOrderNo = Me.LBLDocNo.Caption
    'vSaleOrderDate = Me.LBLDocDate.Caption
    'vDocdate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    'vARCode = Left(Me.LBLARName.Caption, InStr(Me.LBLARName.Caption, "//") - 1)
    'vSaleCode = Left(Me.LBLSaleName.Caption, InStr(Me.LBLSaleName.Caption, "//") - 1)
    'vRequestCountItem = Me.LBLCountItem.Caption
    'vRequestCountQTY = Me.LBLSumQTY.Caption
    'vQuery = "exec dbo.USP_NP_SearchRequestQueueItem '" & vSaleOrderNo & "' "
    'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     ' vRequestAt = vRecordset.Fields("RequestAt").Value + 1
    'Else
     ' vRequestAt = 1
    'End If
    'vRecordset.Close
    
    'On Error GoTo CheckInsertError
    
    'vQuery = "begin tran"
    'gConnection.Execute vQuery
    
    'vQuery = "exec dbo.USP_NP_InsertRequestQueueItem '" & vSaleOrderNo & "','" & vSaleOrderDate & "','" & vDocdate & "','" & vARCode & "','" & vSaleCode & "','" & vRequestDate & "','" & vRequestTime & "'," & vRequestStatus & "," & vRequestCountItem & "," & vRequestCountQTY & ",'" & vUserID & "'," & vRequestAt & " "
    'gConnection.Execute vQuery
      
    'vQuery = "commit tran"
    'gConnection.Execute vQuery
    
    'MsgBox "ส่งเลขที่ใบสั่งขาย/จองเลขที่ " & vSaleOrderNo & " เข้าคิวจัดสินค้าให้ลูกค้าเรียบร้อยแล้วครับ", vbInformation, "Send Information Message"
    
    'Me.LBLDocNo.Caption = ""
    'Me.LBLDocDate.Caption = ""
    'Me.LBLARName.Caption = ""
    'Me.LBLSaleName.Caption = ""
    'Me.LBLCountItem.Caption = ""
    'Me.LBLSumQTY.Caption = ""
    'Me.OPTNow.Value = True
    'Me.OPTSchedule.Value = False
    'Me.TXTPicking1.Caption = ""
    '
    'Me.PTBPickingQueue.Visible = False
    
'CheckInsertError:
 '   If Err.Description <> "" Then
  '    vQuery = "rollback tran"
   '   gConnection.Execute vQuery
    'End If
'Else
 ' MsgBox "ไม่มีข้อมูลเลขที่ใบสั่งขาย/จอง ที่จะเข้าคิวจัดสินค้า กรุณาตรวจสอบ", vbCritical, "Send Error Message"
'End If
End Sub

Private Sub CMDReqPicking_Click()
Dim vDocNo As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vBillType As Integer
Dim vBillStatus As Integer
Dim vSend As Integer
Dim vMemRemainQTY As Double


If Me.TXTPicking1.Caption <> "" Then
    vDocNo = Me.TXTPicking1.Caption
    vQuery = "exec dbo.USP_SO_SOStatus '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vBillStatus = Trim(vRecordset.Fields("billstatus").Value)
        vBillType = Trim(vRecordset.Fields("billtype").Value)
        vSend = Trim(vRecordset.Fields("isconditionsend").Value)
        vMemRemainQTY = Trim(vRecordset.Fields("qty").Value)
    End If
    vRecordset.Close
    
    If vBillStatus = 1 And vMemRemainQTY = 0 Then
        MsgBox "กรุณาตรวจสอบเลขที่ใบสั่งขาย/จอง ที่จะเข้าคิวจัดสินค้า เพราะเอกสารดังกล่าวไม่สามารถทำคิวรอจัดสินค้าได้  เนื่องจากดึงไปทำบิลเรียบร้อยแล้ว", vbCritical, "Send Error Message"
        Exit Sub
    End If
    Me.TBDocNo.Text = vDocNo
Else
Me.PICOrder.Visible = True
Me.TBDocNo.SetFocus
End If





'Dim vRecordset As New ADODB.Recordset
'Dim vQuery As String
'Dim vDocNo As String
'Dim vQueueNo As String
'Dim vQueueStatus As Integer
'Dim vQueueDate As Date
'Dim vCheckAnswerPrint As Integer
'Dim vListItemPicking As ListItem
'Dim i As Integer
'Dim vBillType As Integer
'Dim vBillStatus As Integer
'Dim vSend As Integer
'Dim x As Integer
'Dim vMemRemainQTY As Double


'On Error GoTo ErrDescription



'If Me.TXTPicking1.Caption <> "" Then
   'vDocNo = Trim(Me.TXTPicking1.Caption)
   'vQuery = "exec dbo.USP_SO_SOStatus '" & vDocNo & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      'vBillStatus = Trim(vRecordset.Fields("billstatus").Value)
      'vBillType = Trim(vRecordset.Fields("billtype").Value)
      'vSend = Trim(vRecordset.Fields("isconditionsend").Value)
      'vMemRemainQTY = Trim(vRecordset.Fields("qty").Value)
   'End If
   'vRecordset.Close
   
   'If vBillStatus = 1 And vMemRemainQTY = 0 Then
    '  MsgBox "กรุณาตรวจสอบเลขที่ใบสั่งขาย/จอง ที่จะเข้าคิวจัดสินค้า เพราะเอกสารดังกล่าวไม่สามารถทำคิวรอจัดสินค้าได้  เนื่องจากดึงไปทำบิลเรียบร้อยแล้ว", vbCritical, "Send Error Message"
     ' Exit Sub
   'End If
               
  'vQuery = "exec dbo.USP_SO_CheckSendPicking '" & vDocNo & "' "
  'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   ' vQueueNo = vRecordset.Fields("docno").Value
    'vQueueDate = vRecordset.Fields("queuedatetime").Value
    'vQueueStatus = vRecordset.Fields("status").Value
  'End If
  'vRecordset.Close
  
  'If vQueueNo <> "" Then
   '  If vQueueStatus = 0 Then
    '   MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ รอจัดสินค้า ต้องการพิมพ์ใหม่ ให้พิมพ์ที่หน้าทดแทนเอกสาร", vbInformation, "Send Information"
     'ElseIf vQueueStatus = 1 Then
      ' MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ กำลังจัดสินค้า ต้องการพิมพ์ใหม่ ให้พิมพ์ที่หน้าทดแทนเอกสาร", vbInformation, "Send Information"
     'ElseIf vQueueStatus = 2 Then
      ' MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ จัดสินค้าเรียบร้อยแล้ว ต้องการพิมพ์ใหม่ ให้พิมพ์ที่หน้าทดแทนเอกสาร", vbInformation, "Send Information"
     'End If
     'TXTPicking1.Caption = ""
     'CMBPicking1.Text = ""
     'vCheckPic101 = 0
     'Exit Sub
  'End If
  
  'vQuery = "exec dbo.USP_SO_SaleOrderPicking '" & vDocNo & "' "
  'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   '  Me.LBLRefDocNo.Caption = Trim(vRecordset.Fields("docno").Value)
    ' Me.LBLRefDocDate.Caption = Trim(vRecordset.Fields("docdate").Value)
     'Me.LBLRefARCode.Caption = Trim(vRecordset.Fields("arcode").Value)
     'Me.LBLRefARName.Caption = Trim(vRecordset.Fields("arname").Value)
     'Me.ListViewSelectItemPicking.ListItems.Clear
     'vRecordset.MoveFirst
     'While Not vRecordset.EOF
     'i = i + 1
     'Set vListItemPicking = Me.ListViewSelectItemPicking.ListItems.Add(, , i)
     'vListItemPicking.SubItems(1) = vRecordset.Fields("itemcode").Value
     'vListItemPicking.SubItems(2) = vRecordset.Fields("itemname").Value
     'vListItemPicking.SubItems(3) = Format(vRecordset.Fields("remainqty").Value, "##,##0.00")
     'vListItemPicking.SubItems(4) = Format(vRecordset.Fields("remainqty").Value, "##,##0.00")
     'vListItemPicking.SubItems(5) = vRecordset.Fields("unitcode").Value
     'vListItemPicking.SubItems(6) = vRecordset.Fields("whcode").Value
     'vListItemPicking.SubItems(7) = vRecordset.Fields("shelfcode").Value
    'vListItemPicking.Checked = True
     'vRecordset.MoveNext
     'Wend
     'Me.PICOrder.Visible = True
     'Me.TBDocNo.Text = vDocNo
     'PICSelectPrintSlip.Visible = True
     'Me.CHKSelectAllItem.Value = 1
  'Else
   '  MsgBox "กรุณาตรวจสอบเลขที่ใบสั่งขาย/จอง ที่จะเข้าคิวจัดสินค้า เพราะเอกสารดังกล่าวไม่สามารถทำคิวรอจัดสินค้าได้ ", vbCritical, "Send Error Message"
  'End If
  'vRecordset.Close

'Else
 ' MsgBox "กรุณาเลือกเลขที่ใบสั่งขาย/จอง ที่จะเข้าคิวจัดสินค้า", vbCritical, "Send Error Message"
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub CMDSaleOrderInformationClose_Click()
Me.ListViewSaleOrderQueInformation.ListItems.Clear
Me.ListViewSaleOrderLastQue.ListItems.Clear
Me.PICSaleOrderQueInformation.Visible = False
Me.TBDocNo.Text = ""
Me.PICOrder.Visible = False
Me.ListViewPicking.SetFocus
End Sub

Private Sub CMDSaleOrderSendQue_Click()
Dim vQuery As String
Dim vRecordset As New Recordset
Dim i As Integer

Dim vDocNo As String
Dim vARCode As String
Dim vSaleCode As String
Dim vBillStatus As String
Dim vSoStatus As Integer
Dim n As Integer
Dim vBillType As Integer
Dim vCarlicense As String
Dim vDeliveryDate As String
Dim vPickStatus As Integer
Dim vSOCountNumber As Integer

Dim vDocdate As String
Dim vQueDocDate As String
Dim vPickingDate As String
Dim vItemCode As String
Dim vItemName As String
Dim vReqQTY As Double
Dim vUnitCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vZoneID As String
Dim vIsCancel As Integer
Dim vLineNumber As Integer
Dim j As Integer
Dim vIsConditionSend As Integer
Dim vCountNumber As Integer
Dim vCheckShelfGroup As String
Dim vDueDate As String
Dim vSelectItemDateTime As String

Dim vSumOfItemAmount As Double
Dim vTaxAmount As Double
Dim vTotalAmount As Double

Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vItemAmount As Double

Dim vListItem As ListItem
Dim vAnswer As Integer
Dim vLastTimeID As Integer
Dim vCheckTimeID As Integer
Dim vQTY As Double
Dim vPickQty As Double

If Me.LBLOrderArCode.Caption <> "" And Me.ListViewSaleOrder.ListItems.Count > 0 Then

   vDocNo = Me.TBDocNo.Text
   vDocdate = Me.LBLOrderDocDate.Caption
   vPickingDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   vARCode = Me.LBLOrderArCode.Caption
   vSaleCode = Left(Me.LBLOrderSaleCode.Caption, InStr(Me.LBLOrderSaleCode.Caption, "/") - 1)
   
   vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
      
    vQuery = "exec dbo.USP_NP_CheckQuePickCenter '" & vDocNo & "','" & vQueDocDate & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vLastTimeID = Trim(vRecordset.Fields("max1").Value)
    End If
    vRecordset.Close
    
   vCheckTimeID = vLastTimeID + 1
   
   vQuery = "exec dbo.usp_np_SearchReqPickingInformationLastSend '" & vDocNo & "','" & vQueDocDate & "'," & vLastTimeID & " "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.ListViewSaleOrderLastQue.ListItems.Clear
      Me.ListViewSaleOrderQueInformation.ListItems.Clear
      vRecordset.MoveFirst
      For i = 1 To vRecordset.RecordCount
        Set vListItem = Me.ListViewSaleOrderLastQue.ListItems.Add(, , i)
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
   
   If Me.ListViewSaleOrderLastQue.ListItems.Count > 0 Then
      Me.LBLSaleOrderQueInf.Caption = "รายการ คิวจัดสินค้าล่าสุดของเอกสารนี้"
      Me.PICSaleOrderQueInformation.Visible = True
      Me.ListViewSaleOrderQueInformation.Visible = False
      Me.ListViewSaleOrderLastQue.Visible = True
      
      vAnswer = MsgBox("คุณต้องการ ส่งคิวจัดสินค้าต่อหรือไม่", vbYesNo, "Send Question Message")
      If vAnswer = 7 Then
        Exit Sub
      End If
   Else
      Me.LBLSaleOrderQueInf.Caption = "รายการ คิวที่สั่งจัดสินค้า"
      Me.PICSaleOrderQueInformation.Visible = False
      Me.ListViewSaleOrderQueInformation.Visible = False
      Me.ListViewSaleOrderLastQue.Visible = False
   End If
   
   
   If vSaleCode = "" Then
      MsgBox "ไม่ได้ระบุ รหัสพนักงานกรุณาตรวจสอบ ", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
   vIsConditionSend = 0
   vCarlicense = ""
   vBillType = Me.LBLOrderBillType.Caption
   vSoStatus = Me.LBLOrderSoStatus.Caption
   vDeliveryDate = vPickingDate
   vDueDate = vPickingDate
   vPickStatus = 0
   
   If Me.LBLOrderSumOfItemAmount.Caption <> "" Then
   vSumOfItemAmount = Me.LBLOrderSumOfItemAmount.Caption
   End If
   
   If Me.LBLOrderTaxAmount.Caption <> "" Then
   vTaxAmount = Me.LBLOrderTaxAmount.Caption
   End If
   
   If Me.LBLOrderNetAmount.Caption <> "" Then
   vTotalAmount = Me.LBLOrderNetAmount.Caption
   End If
   
   vQuery = "exec dbo.USP_NP_SearchCheckCountSOPicking '" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vSOCountNumber = vRecordset.Fields("vCount").Value
   End If
   vRecordset.Close
   
   'vQuery = "exec dbo.USP_NP_SearchSaleOrderGroupShelf '" & vDocNo & "' "
   vQuery = "exec dbo.USP_NP_SearchSaleOrderGroupShelfPickZone '" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vRecordset.MoveFirst
      While Not vRecordset.EOF
      vCheckShelfGroup = vRecordset.Fields("shelfgroup").Value
      vQuery = "exec dbo.USP_NP_InsertOrderPickHoldBill '" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vPickingDate & "'," & vBillType & "," & vSoStatus & ",0,'" & vSaleCode & "','" & vCarlicense & "'," & vIsConditionSend & "," & vSOCountNumber & ",'" & vCheckShelfGroup & "','" & vDueDate & "'," & vPickStatus & "," & vSumOfItemAmount & "," & vTaxAmount & "," & vTotalAmount & ",'" & vUserID & "' "
      gConnection.Execute vQuery
      vRecordset.MoveNext
      Wend
   End If
   vRecordset.Close
   
   For j = 1 To Me.ListViewSaleOrder.ListItems.Count
      vItemCode = Me.ListViewSaleOrder.ListItems(j).SubItems(1)
      vItemName = Me.ListViewSaleOrder.ListItems(j).SubItems(2)
      vReqQTY = Me.ListViewSaleOrder.ListItems(j).SubItems(3)
      vUnitCode = Me.ListViewSaleOrder.ListItems(j).SubItems(4)
      vWHCode = Me.ListViewSaleOrder.ListItems(j).SubItems(8)
      vShelfCode = Me.ListViewSaleOrder.ListItems(j).SubItems(9)
      vZoneID = Me.ListViewSaleOrder.ListItems(j).SubItems(10)
      vPrice = Me.ListViewSaleOrder.ListItems(j).SubItems(5)
      vDiscountAmount = Me.ListViewSaleOrder.ListItems(j).SubItems(6)
      vItemAmount = Me.ListViewSaleOrder.ListItems(j).SubItems(7)
      vIsCancel = 0
      vSelectItemDateTime = Now
      vLineNumber = j - 1
      vQuery = "exec dbo.USP_NP_InsertOrderPickHoldBillSub '" & vDocNo & "','" & vDocdate & "','" & vPickingDate & "','" & vItemCode & "','" & vItemName & "'," & vReqQTY & ",'" & vUnitCode & "','" & vWHCode & "','" & vShelfCode & "','" & vZoneID & "'," & vIsCancel & ",'" & vSelectItemDateTime & "'," & vSOCountNumber & "," & vPrice & "," & vDiscountAmount & "," & vItemAmount & "," & vLineNumber & " "
      gConnection.Execute vQuery
   Next j
   
   Call SendQue(vSOCountNumber)
   
    Me.PICSaleOrderQueInformation.Visible = True
    Me.ListViewSaleOrderLastQue.Visible = False
    Me.ListViewSaleOrderQueInformation.Visible = True
    Me.ListViewSaleOrderQueInformation.ListItems.Clear
    vQuery = "exec dbo.usp_np_SearchReqPickingInformation '" & vDocNo & "'," & vSOCountNumber & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.ListViewSaleOrderQueInformation.ListItems.Clear
    vRecordset.MoveFirst
    For i = 1 To vRecordset.RecordCount
    Set vListItem = Me.ListViewSaleOrderQueInformation.ListItems.Add(, , i)
    vListItem.SubItems(1) = Trim(vRecordset.Fields("queid").Value)
    vListItem.SubItems(2) = Trim(vRecordset.Fields("zoneid").Value)
    vRecordset.MoveNext
    Next i
    End If
    vRecordset.Close
    Me.ListViewSaleOrderQueInformation.SetFocus
   
   Me.TBDocNo.Text = ""
   Me.TBDocNo.SetFocus
End If
End Sub

Public Sub SendQue(vTimeID As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim n As Integer
Dim vDocNo As String
Dim vDocdate As String
Dim vQueDocDate As String
Dim vGroupZone(5) As String

Dim vListItem As ListItem
Dim vCheckDate As String

Dim vCheckQueSend As Integer
Dim vQTY As Double
Dim vPickQty As Double

If Me.ListViewSaleOrder.ListItems.Count > 0 And Me.LBLOrderArCode.Caption <> "" Then
   vDocNo = Me.TBDocNo.Text
   vDocdate = Me.LBLOrderDocDate.Caption
   vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   
   vQuery = "exec dbo.usp_np_SearchReqPickingInformationLastSend '" & vDocNo & "','" & vQueDocDate & "'," & vTimeID & " "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.ListViewSaleOrderQueInformation.ListItems.Clear
      vCheckQueSend = 1
      vRecordset.MoveFirst
      For i = 1 To vRecordset.RecordCount
        Set vListItem = Me.ListViewSaleOrderQueInformation.ListItems.Add(, , i)
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
   
   'If vCheckQueSend = 1 Then
    '  Me.PICLastSendQue.Visible = True
     ' Exit Sub
   'End If
   
   vQuery = "exec dbo.USP_NP_SearchGroupPicking 2,'" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       n = vRecordset.RecordCount
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
       vGroupZone(i) = Trim(vRecordset.Fields("zoneid").Value)
       vRecordset.MoveNext
       Next i
   End If
   vRecordset.Close
   
   vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
   
   For i = 1 To n
      If vGroupZone(i) = "A" Then
         Call PrintSalePicking_A(vDocNo, vTimeID, 2)
      ElseIf vGroupZone(i) = "B" Then
         Call PrintSalePicking_B(vDocNo, vTimeID, 2)
      ElseIf vGroupZone(i) = "C" Then
         Call PrintSalePicking_C(vDocNo, vTimeID, 2)
      ElseIf vGroupZone(i) = "X" Then
         Call PrintSalePicking_X(vDocNo, vTimeID, 2)
      End If
   Next i
   
   
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
   
   Me.LBLOrderDocDate.Caption = vDay1 & "/" & vMonth1 & "/" & Year(Now)
   Me.LBLOrderArCode.Caption = ""
   Me.LBLOrderArName.Caption = ""
   Me.TBDocNo.Text = ""
   Me.LBLOrderSaleCode.Caption = ""
   Me.LBLOrderNetAmount.Caption = ""
   'Me.CMDQue.Enabled = False
   
   'MePICSaleOrderQueInformation.Visible = True
   vQuery = "exec dbo.usp_np_SearchReqPickingInformation '" & vDocNo & "'," & vTimeID & " "
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

vQuery = "exec dbo.USP_NP_UpdateSendQueuePicking 1,'" & vDocNo & "' "
gConnection.Execute vQuery
   
Else
   MsgBox "เอกสารที่จะส่งจัดสินค้าได้ ต้องมีเลขที่เอกสาร รายการสินค้า และต้องเป็นเอกสารที่บันทึกข้อมูลเรียบร้อยแล้วเป็นอย่างน้อย กรุณาตรวจสอบ ", vbCritical, "Send Error Message"
End If
End Sub

Private Sub CMDSaveQueue_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vShelfGroup(10) As String
Dim i As Integer
Dim vPrint As Integer
Dim vBillStatus As String
Dim n As Integer
Dim vBillType As Integer
Dim vSend As Integer
Dim vHour As Integer
Dim vMinute As Integer
Dim vCheckDateTime As Date
Dim vCheckDateDiff As Integer
Dim vRequestTime As Date
Dim vCarlicense As String
Dim vRemainQtyCheckPrint As Double

On Error GoTo ErrDescription

If Me.OptSchedule.Value = True Then
vHour = Left(Trim(Me.MEDTime.Text), 2)
vMinute = Right(Trim(Me.MEDTime.Text), 2)
vCheckDateTime = Day(Now) & "/" & Month(Now) & "/" & Year(Now) & "    " & vHour & ":" & vMinute & ":" & "00"
vCheckDateDiff = DateDiff("n", Now, vCheckDateTime)
If vCheckDateDiff < 15 Then
  MsgBox "ไม่สามารถกำหนดเวลาที่ลูกค้าต้องการรับสินค้าน้อยกว่า 15 นาทีได้", vbCritical, "Send Error Infromation"
  Exit Sub
End If
End If

If Me.LBLDocNo.Caption <> "" Then
 vDocNo = Me.LBLDocNo.Caption
 'USP_SO_SOStatus1
 vQuery = "exec dbo.USP_SO_SOStatus '" & vDocNo & "' "
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     vBillStatus = Trim(vRecordset.Fields("billstatus").Value)
     vBillType = Trim(vRecordset.Fields("billtype").Value)
     vSend = Trim(vRecordset.Fields("isconditionsend").Value)
     vSaleCode = Trim(vRecordset.Fields("salecode").Value)
     vRemainQtyCheckPrint = vRecordset.Fields("qty").Value
 End If
 vRecordset.Close
         
 If vRemainQtyCheckPrint > 0 Then
   'USP_SO_SearchShelfGroupPicking
   vQuery = "exec dbo.USP_SO_SearchShelfGroupPicking '" & vDocNo & "',1 "
   'vQuery = "exec dbo.USP_SO_SearchPickingSlip1 '" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     n = vRecordset.RecordCount
     vRecordset.MoveFirst
     For i = 1 To vRecordset.RecordCount
     vShelfGroup(i) = Trim(vRecordset.Fields("shelfgroup").Value)
     vRecordset.MoveNext
     Next i
   End If
   vRecordset.Close
   
   If Me.CHKLicense.Value = 1 Then
     vCarlicense = Me.TextCarLicense.Text
     vQuery = "exec dbo.USP_BC_UpdateCarLicense '" & vDocNo & "','" & vCarlicense & "'  "
     gConnection.Execute (vQuery)
   End If
   
   For i = 1 To n
     If vShelfGroup(i) = "A" Then
       Call PrintPicking_A
     ElseIf vShelfGroup(i) = "B" Then
       Call PrintPicking_B
     ElseIf vShelfGroup(i) = "M" Then
       Call PrintPicking_M_OutLet
     ElseIf vShelfGroup(i) = "K" Then
       Call PrintPicking_K_BackStock
     ElseIf vShelfGroup(i) = "H" Then
       If vSend = 1 Then
         Call PrintPicking_M_HMX
       ElseIf vBillType = 1 And vSend = 0 Then
         Call PrintPicking_M_HMX_CustReceive
       ElseIf vSend = 0 And vBillType = 0 Then
         MsgBox "ขายสินค้าเงินสดคลัง 014 ที่ลูกค้ามารับเองจะไม่สามารถเข้าคิวหยิบสินค้าได้ แต่สินค้าคลังอื่น ๆ สามารถเข้าคิวหยิบได้", vbInformation, "Send Information"
       End If
     ElseIf vShelfGroup(i) = "D" Then
       Call PrintPicking_D
     ElseIf vShelfGroup(i) = "Y" Then
       Call PrintPicking_Y
     End If
   Next i
   Me.ListViewPicking.ListItems.Remove (ListViewPicking.SelectedItem.Index)
   Me.PTBPickingQueue.Visible = False
 End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDSearchSaleOrder_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem

Dim vDocNo As String
Dim vSumItemAmount As Double
Dim vTaxAmount As Double
Dim vNetAmount As Double

Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vAmount As Double

If Me.TBDocNo.Text <> "" Then
   vDocNo = Me.TBDocNo.Text

   vQuery = "exec dbo.USP_NP_SearchSaleOrder '" & vDocNo & "'"
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.ListViewSaleOrder.ListItems.Clear
      vSumItemAmount = Trim(vRecordset.Fields("sumofitemamount").Value)
      vTaxAmount = Trim(vRecordset.Fields("taxamount").Value)
      vNetAmount = Trim(vRecordset.Fields("netamount").Value)
      
      'Me.LBLOrderBillType.Caption = Trim(vRecordset.Fields("billtype").Value)
      'Me.LBLOrderSoStatus.Caption = Trim(vRecordset.Fields("sostatus").Value)
      Me.LBLOrderDocDate.Caption = Trim(vRecordset.Fields("docdate").Value)
      Me.LBLOrderArCode.Caption = Trim(vRecordset.Fields("arcode").Value)
      Me.LBLOrderArName.Caption = Trim(vRecordset.Fields("arname").Value)
      If Trim(vRecordset.Fields("salecode").Value) = "" Then
          MsgBox "เอกสารไม่ได้กำหนดรหัสพนักงานขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      End If
      Me.LBLOrderSaleCode.Caption = Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      Me.LBLOrderSumOfItemAmount.Caption = Format(vSumItemAmount, "##,##0.00")
      Me.LBLOrderTaxAmount.Caption = Format(vTaxAmount, "##,##0.00")
      Me.LBLOrderNetAmount.Caption = Format(vNetAmount, "##,##0.00")
      
      vRecordset.MoveFirst
      For i = 1 To vRecordset.RecordCount
        Set vListItem = Me.ListViewSaleOrder.ListItems.Add(, , i)
        
        vQTY = Trim(vRecordset.Fields("remainqty").Value)
        vPrice = Trim(vRecordset.Fields("price").Value)
        vDiscountAmount = Trim(vRecordset.Fields("discountamountsub").Value)
        vAmount = Trim(vRecordset.Fields("amount").Value)

        vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
        vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
        vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
        vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
        vListItem.SubItems(6) = Format(vDiscountAmount, "##,##0.00")
        vListItem.SubItems(7) = Format(vAmount, "##,##0.00")
        vListItem.SubItems(8) = Trim(vRecordset.Fields("whcode").Value)
        vListItem.SubItems(9) = Trim(vRecordset.Fields("shelfcode").Value)
        vListItem.SubItems(10) = Trim(vRecordset.Fields("zoneid").Value)
        vListItem.SubItems(11) = Trim(vRecordset.Fields("shelfid").Value)
        vListItem.SubItems(12) = Trim(vRecordset.Fields("itemcode").Value)
        vListItem.SubItems(13) = Trim(vRecordset.Fields("discountwordsub").Value)
        vListItem.SubItems(14) = Format(vQTY, "##,##0.00")
        vRecordset.MoveNext
      Next i
   Else
      Me.LBLOrderDocDate.Caption = ""
      Me.LBLOrderArCode.Caption = ""
      Me.LBLOrderArName.Caption = ""
      Me.LBLOrderSaleCode.Caption = ""
      Me.LBLOrderSumOfItemAmount.Caption = ""
      Me.LBLOrderTaxAmount.Caption = ""
      Me.LBLOrderNetAmount.Caption = ""
      'Me.LBLOrderBillType.Caption = ""
      'Me.LBLOrderSoStatus.Caption = ""
   End If
   vRecordset.Close
Else
   MsgBox "กรุณากรอกเลขที่ใบสั่งขาย/จอง ให้ถูกต้อง", vbCritical, "Send Error Message"
End If
End Sub

Private Sub CMDSelectItemBack_Click()
   Me.PICSelectPrintSlip.Visible = True
   Me.PTBPickingQueue.Visible = False
   vCheckSelectItemPickBack = 1
End Sub

Private Sub CMDSendPicking_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vShelfGroup(10) As String
Dim i As Integer
Dim vPrint As Integer
Dim vBillStatus As String
Dim vSoStatus As Integer
Dim n As Integer
Dim vBillType As Integer
Dim vSend As Integer
Dim vHour As Integer
Dim vMinute As Integer
Dim vCheckDateTime As Date
Dim vCheckDateDiff As Integer
Dim vRequestTime As Date
Dim vCarlicense As String
Dim vRemainQtyCheckPrint As Double
Dim vDeliveryDate As String

Dim vDocdate As String
Dim vPickingDate As String
Dim vItemCode As String
Dim vItemName As String
Dim vReqQTY As Double
Dim vUnitCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vZoneID As String
Dim vIsCancel As Integer
Dim vSelectItemDateTime As String
Dim vLineNumber As Integer
Dim j As Integer
Dim vCarLicence As String
Dim vIsConditionSend As Integer
Dim vCountNumber As Integer
Dim vCheckShelfGroup As String
Dim vDueDate As String
Dim vCheckSPO As Integer
Dim m As Integer

Dim vPickStatus As Integer

vDocNo = Me.LBLDocNo.Caption
vDocdate = Me.LBLDocDate.Caption
vPickingDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
vARCode = Left(Me.LBLARName.Caption, InStr(Me.LBLARName.Caption, "//") - 1)
vSaleCode = Me.LBLSaleCode.Caption

If vSaleCode = "" Then
   MsgBox "ไม่ได้ระบุ รหัสพนักงานกรุณาตรวจสอบ ", vbCritical, "Send Error Message"
   Exit Sub
End If

vCheckSPO = 0
For m = 1 To Me.ListViewSelectItemPicking.ListItems.Count
   If Me.ListViewSelectItemPicking.ListItems(m).Checked = True Then
        If Me.ListViewSelectItemPicking.ListItems(m).SubItems(7) = "SPO" Then
           vCheckSPO = vCheckSPO + 1
        End If
   End If
Next m

If vCheckSPO > 0 Then
   If Me.OptMain.Value = False And Me.OptOutLet.Value = False Then
      MsgBox "กรณีที่มีการสั่งจัดสินค้าชั้นเก็บ SPO ต้องระบุด้วยว่าลูกค้ารับของฝั่งไหนตามที่อยู่สินค้าที่อยู่จริง เพื่อความสะดวกต่อการจัดสินค้า กรุณาระบุด้วย", vbCritical, "Send Error Message"
      Exit Sub
   End If
End If

vIsConditionSend = Me.LBLIsConditionSend.Caption
vCarlicense = Me.TextCarLicense.Text
vBillType = Me.LBLBillType.Caption
vSoStatus = Me.LBLSaleType.Caption
vDeliveryDate = Me.LBLDueDate.Caption
vDueDate = Me.LBLDueDate.Caption

If vSoStatus = 1 Then
   If Me.OPTPayItem.Value = True Then
      vPickStatus = 0
   ElseIf Me.OPTReserve.Value = True Then
      vPickStatus = 1
   End If
End If

vQuery = "exec dbo.USP_NP_SearchCheckCountSOPicking '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vSOCountNumber = vRecordset.Fields("vCount").Value
End If
vRecordset.Close

vQuery = "exec dbo.USP_NP_SearchCheckShelfSaleOrderData '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   vCheckShelfGroup = vRecordset.Fields("shelfgroup").Value
   vQuery = "exec dbo.USP_NP_InsertSelectItemPickingMaster1 '" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vPickingDate & "'," & vBillType & "," & vSoStatus & ",0,'" & vSaleCode & "','" & vCarlicense & "'," & vIsConditionSend & "," & vSOCountNumber & ",'" & vCheckShelfGroup & "','" & vDueDate & "'," & vPickStatus & ",'" & vUserID & "' "
   gConnection.Execute vQuery
   vRecordset.MoveNext
  Wend
End If
vRecordset.Close

'vQuery = "exec dbo.USP_NP_InsertSelectItemPickingMaster '" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vPickingDate & "',0,'" & vSaleCode & "','" & vCarLicense & "'," & vIsConditionSend & "," & vSOCountNumber & ",'" & vUserID & "' "
'gConnection.Execute vQuery

For j = 1 To Me.ListViewSelectItemPicking.ListItems.Count
       If Me.ListViewSelectItemPicking.ListItems(j).Checked = True Then
          vItemCode = Me.ListViewSelectItemPicking.ListItems(j).SubItems(1)
          vItemName = Me.ListViewSelectItemPicking.ListItems(j).SubItems(2)
          vReqQTY = Me.ListViewSelectItemPicking.ListItems(j).SubItems(4)
          vUnitCode = Me.ListViewSelectItemPicking.ListItems(j).SubItems(5)
          vWHCode = Me.ListViewSelectItemPicking.ListItems(j).SubItems(6)
          vShelfCode = Me.ListViewSelectItemPicking.ListItems(j).SubItems(7)
          vZoneID = Me.ListViewSelectItemPicking.ListItems(j).SubItems(7)
          vIsCancel = 0
          vSelectItemDateTime = Now
          vLineNumber = j - 1
          vQuery = "exec dbo.USP_NP_InsertSelectItemPicking '" & vDocNo & "','" & vDocdate & "','" & vPickingDate & "','" & vItemCode & "','" & vItemName & "'," & vReqQTY & ",'" & vUnitCode & "','" & vWHCode & "','" & vShelfCode & "','" & vZoneID & "'," & vIsCancel & ",'" & vSelectItemDateTime & "'," & vSOCountNumber & "," & vLineNumber & " "
          gConnection.Execute vQuery
       End If
Next j

If Me.OptSchedule.Value = True Then
   vHour = Left(Trim(Me.MEDTime.Text), 2)
   vMinute = Right(Trim(Me.MEDTime.Text), 2)
   vCheckDateTime = Day(Now) & "/" & Month(Now) & "/" & Year(Now) & "    " & vHour & ":" & vMinute & ":" & "00"
   vCheckDateDiff = DateDiff("n", Now, vCheckDateTime)
   If vCheckDateDiff < 15 Then
     MsgBox "ไม่สามารถกำหนดเวลาที่ลูกค้าต้องการรับสินค้าน้อยกว่า 15 นาทีได้", vbCritical, "Send Error Infromation"
     Exit Sub
   End If
End If

If Me.LBLDocNo.Caption <> "" Then
   vDocNo = Me.LBLDocNo.Caption
   vQuery = "exec dbo.USP_SO_SOStatus '" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vBillStatus = Trim(vRecordset.Fields("billstatus").Value)
       vBillType = Trim(vRecordset.Fields("billtype").Value)
       vSend = Trim(vRecordset.Fields("isconditionsend").Value)
       vSaleCode = Trim(vRecordset.Fields("salecode").Value)
       vRemainQtyCheckPrint = vRecordset.Fields("qty").Value
   End If
   vRecordset.Close
         
   If vRemainQtyCheckPrint > 0 Then
     vQuery = "exec dbo.USP_SO_SearchShelfGroupPicking '" & vDocNo & "'," & vSOCountNumber & " "
     If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       n = vRecordset.RecordCount
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
       vShelfGroup(i) = Trim(vRecordset.Fields("shelfgroup").Value)
       vRecordset.MoveNext
       Next i
     End If
     vRecordset.Close
   
     If Me.CHKLicense.Value = 1 Then
       vCarlicense = Me.TextCarLicense.Text
       vQuery = "exec dbo.USP_NP_UpdateCarLicense '" & vDocNo & "'," & vSOCountNumber & ",'" & vCarlicense & "'"
       gConnection.Execute (vQuery)
     End If
   
     If DatePart("w", Now) <> 1 Then
        For i = 1 To n
          If vShelfGroup(i) = "AVL" Then
             Call PrintPicking_AVL(vSOCountNumber)
          ElseIf vShelfGroup(i) = "BK1" Then
             Call PrintPicking_BK1(vSOCountNumber)
          ElseIf vShelfGroup(i) = "BK2" Then
             Call PrintPicking_BK2(vSOCountNumber)
          ElseIf vShelfGroup(i) = "BK3" Then
             Call PrintPicking_BK3(vSOCountNumber)
          ElseIf vShelfGroup(i) = "SPO" Then
             Call PrintPicking_SPO(vSOCountNumber)
          End If
        Next i
     Else
        For i = 1 To n
          If vShelfGroup(i) = "AVL" Then
             Call PrintPicking_AVL(vSOCountNumber)
          ElseIf vShelfGroup(i) = "BK1" Then
             Call PrintPicking_BK1_Sunday(vSOCountNumber)
          ElseIf vShelfGroup(i) = "BK2" Then
             Call PrintPicking_BK2_Sunday(vSOCountNumber)
          ElseIf vShelfGroup(i) = "BK3" Then
             Call PrintPicking_BK3(vSOCountNumber)
          ElseIf vShelfGroup(i) = "SPO" Then
             Call PrintPicking_SPO(vSOCountNumber)
          End If
        Next i
     
     End If
     Me.ListViewPicking.ListItems.Remove (ListViewPicking.SelectedItem.Index)
     Me.TXTPicking1.Caption = ""
     Me.PTBPickingQueue.Visible = False
     Me.PICSelectPrintSlip.Visible = False
 End If
End If
End Sub

Private Sub CMDSOMain_Click()
Me.TBDocNo.Text = ""
Me.PICOrder.Visible = False
Me.TXTPicking1.Caption = ""
Me.ListViewPicking.SetFocus
End Sub

Private Sub Command1_Click()
Dim vSaleOrder As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vTimeID As Integer


Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
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
Dim vZone As Integer


vSaleOrder = "ROV5105-0092"
vWHCode = "S01"
vShelfCode = "AVL"
vTimeID = 1
vZone = 0


If Me.OPTReserve.Value = True Then
   vSelectPicked = 0
Else
   vSelectPicked = 1
End If


vGroupDocNo = UCase(Left(Right(vSaleOrder, Len(vSaleOrder) - InStr(vSaleOrder, "-")), 3)) 'UCase(Left(vSaleOrder, 3))
  
If Me.OPTReserve.Value = True Then
   vSelectPicked = 0
Else
   vSelectPicked = 1
End If


If vGroupDocNo = "RWV" Or vGroupDocNo = "RWN" Then
   If vSelectPicked = 0 Then 'Res
      If vZone = 0 Then
         vPrinterName = Trim("TM_Moo")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If

      If vZone = 1 Then 'Res
         vPrinterName = Trim("TM_Moo")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If
   End If
   
   If vSelectPicked = 1 Then
      If vShelfCode = "AVL" Then 'Pay
         vPrinterName = Trim("TM_Moo")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      Else
         If vZone = 0 Then
            vPrinterName = Trim("TM_Moo")
            For Each printerObj In Printers
            If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
            Set Printer = printerObj
            Set printerObj = Nothing
            Exit For
            End If
            Next
         End If
         
         If vZone = 1 Then
            vPrinterName = Trim("TM_Moo")
            For Each printerObj In Printers
            If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
            Set Printer = printerObj
            Set printerObj = Nothing
            Exit For
            End If
            Next
         End If
      End If
   End If
      
ElseIf vGroupDocNo = "ROV" Or vGroupDocNo = "RON" Then

   If vSelectPicked = 0 Then 'Res
      If vZone = 0 Then
         vPrinterName = Trim("TM_Moo")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If

      If vZone = 1 Then 'Res
         vPrinterName = Trim("TM_Moo")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If
   End If
   
Else
   If vZone = 0 Then
      vPrinterName = Trim("TM_Moo")
      For Each printerObj In Printers
      If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
      Set Printer = printerObj
      Set printerObj = Nothing
      Exit For
      End If
      Next
   End If
   
   If vZone = 1 Then
      vPrinterName = Trim("TM_Moo")
      For Each printerObj In Printers
      If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
      Set Printer = printerObj
      Set printerObj = Nothing
      Exit For
      End If
      Next
   End If
End If
    
    vQuery = "exec dbo.USP_SO_PickingQueueFreedom '" & vSaleOrder & "','" & vWHCode & "','" & vShelfCode & "'," & vTimeID & " "
    'vQuery = "exec dbo.USP_SO_PickingQueueFreedom 'ROV5105-0092','S01','AVL',1"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    
      vSoStatus = vRecordset.Fields("sostatus").Value
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

If vSoStatus = 0 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 20
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 80
      Printer.Print "(จ่าย)"
ElseIf vSoStatus = 1 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 20
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 80
      Printer.Print "(จ่าย)"
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
      
      
     If vSoStatus = 1 And vSelectPicked = 0 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 12
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ต้นฉบับใบจัดสินค้า เพื่อจอง")
      ElseIf vSoStatus = 1 And vSelectPicked = 1 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 12
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ต้นฉบับใบจัดสินค้า เพื่อจ่าย")
      ElseIf vSoStatus = 0 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 12
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ต้นฉบับใบจัดสินค้า")
      End If
      
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

Private Sub Command2_Click()
Dim vSaleOrder As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vTimeID As Integer


Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
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
Dim vSelectPicked As Integer
Dim vGroupDocNo As String
Dim vZone As Integer


vSaleOrder = "ROV5105-0092"
vWHCode = "S01"
vShelfCode = "AVL"
vTimeID = 1
vZone = 0


vGroupDocNo = UCase(Left(Right(vSaleOrder, Len(vSaleOrder) - InStr(vSaleOrder, "-")), 3)) 'UCase(Left(vSaleOrder, 3))
  
If Me.OPTReserve.Value = True Then
   vSelectPicked = 0
Else
   vSelectPicked = 1
End If


If vGroupDocNo = "RWV" Or vGroupDocNo = "RWN" Then
   If vSelectPicked = 0 Then 'Res
      If vZone = 0 Then
         vPrinterName = Trim("TM_Moo")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If

      If vZone = 1 Then 'Res
         vPrinterName = Trim("TM_Moo")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If
   End If
   
   If vSelectPicked = 1 Then
      If vShelfCode = "AVL" Then 'Pay
         vPrinterName = Trim("TM_Moo")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      Else
         If vZone = 0 Then
            vPrinterName = Trim("TM_Moo")
            For Each printerObj In Printers
            If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
            Set Printer = printerObj
            Set printerObj = Nothing
            Exit For
            End If
            Next
         End If
         
         If vZone = 1 Then
            vPrinterName = Trim("TM_Moo")
            For Each printerObj In Printers
            If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
            Set Printer = printerObj
            Set printerObj = Nothing
            Exit For
            End If
            Next
         End If
      End If
   End If
      
ElseIf vGroupDocNo = "ROV" Or vGroupDocNo = "RON" Then

   If vSelectPicked = 0 Then 'Res
      If vZone = 0 Then
         vPrinterName = Trim("TM_Moo")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If

      If vZone = 1 Then 'Res
         vPrinterName = Trim("TM_Moo")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If
   End If
   
Else
   If vZone = 0 Then
      vPrinterName = Trim("TM_Moo")
      For Each printerObj In Printers
      If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
      Set Printer = printerObj
      Set printerObj = Nothing
      Exit For
      End If
      Next
   End If
   
   If vZone = 1 Then
      vPrinterName = Trim("TM_Moo")
      For Each printerObj In Printers
      If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
      Set Printer = printerObj
      Set printerObj = Nothing
      Exit For
      End If
      Next
   End If
End If

    vQuery = "exec dbo.USP_SO_PickingQueueFreedom '" & vSaleOrder & "','" & vWHCode & "','" & vShelfCode & "'," & vTimeID & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    
      vSoStatus = vRecordset.Fields("sostatus").Value
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")
      
If vSoStatus = 0 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 20
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 80
      Printer.Print "(จ่าย)"
ElseIf vSoStatus = 1 Then
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
    
     If vSoStatus = 1 And vSelectPicked = 0 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า เพื่อจอง")
      ElseIf vSoStatus = 1 And vSelectPicked = 1 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า เพื่อจ่าย")
      Else
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า เพื่อจ่าย")
      End If
      
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
      Printer.Font.Size = 14
      Printer.CurrentX = 0
      Printer.CurrentY = 3400
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value)
      
      If Trim(vRecordset.Fields("carlicense").Value) <> "" Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 16
        Printer.CurrentX = 1400
        Printer.CurrentY = 3300
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
      
    Printer.CurrentX = 0
    Printer.CurrentY = 4300
    Printer.Print Trim("______________________________________________________________________________________________")

    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 30
    Printer.CurrentX = 800
    Printer.CurrentY = 4500
    Printer.Print Trim("ครบ")
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 30
    Printer.CurrentX = 2800
    Printer.CurrentY = 4500
    Printer.Print Trim("ไม่ครบ")
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 20
    Printer.CurrentX = 0
    Printer.CurrentY = 4600
    Printer.Print Trim("______")
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 20
    Printer.CurrentX = 2100
    Printer.CurrentY = 4600
    Printer.Print Trim("______")

    Printer.CurrentX = 0
    Printer.CurrentY = 4700
    Printer.Print Trim("______________________________________________________________________________________________")
    End If
    vRecordset.Close

    Printer.Font.Size = 10
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("วันที่พิมพ์ :") & Now
           
    Printer.EndDoc
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vPicture As String
Dim ListX As ListItem
Dim x As ListImage
Dim i As Integer
Dim SOPListItems As ListItem
Dim vTypeDoc As String
Dim vCheckDate As Date
Dim vCheckGenDocno As String
Dim vCheckDateNow As Date
Dim vCheckYear As String
Dim vCheckMonth As String
Dim vCheckDay As String
Dim vCheckNumber As Integer
Dim vCountNumber As Integer
Dim j As Integer, m As Integer
Dim vCheckGenDoc(10) As Integer
Dim vUserLogInProgram As String

On Error GoTo ErrDescription

vCheckValue = False
ListViewPicking.ListItems.Clear
CMBSale.Clear
vQuery = "select * from vw_NP_SaleUserID "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMBSale.AddItem Trim(vRecordset.Fields("salename").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
 
 vQuery = "exec dbo.USP_NP_SeaechUserLogIn '" & vUserID & "' "
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  vUserLogInProgram = Trim(vRecordset.Fields("salename").Value)
 End If
 vRecordset.Close
 
 CMBSale.Text = vUserLogInProgram
 
 vTypeDoc = "SO"
ListViewPicking.ListItems.Clear

'vQuery = "Select Docno,name1,lastprintdatetime  from BCNP.dbo.vw_sl_00002  where Printed = 0  " _
 '                   & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
 
vQuery = "exec dbo.USP_SO_SearchDocumentToPrint 0,'" & vUserID & "','" & vTypeDoc & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    If Not vRecordset.EOF Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set SOPListItems = ListViewPicking.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
        If IsNull(vRecordset.Fields("arname").Value) Then
        SOPListItems.SubItems(1) = ""
        Else
        SOPListItems.SubItems(1) = Trim(vRecordset.Fields("arname").Value)
        SOPListItems.SubItems(2) = Trim(vRecordset.Fields("lastprintdatetime").Value)
        End If
        vRecordset.MoveNext
        Wend
    End If
End If
vRecordset.Close
        '--------------------------------------------------------------------------------------------------------------------
CMBPicking1.Clear
vQuery = "select Name from npmaster.dbo.NPForm where ModuleID = '" & vTypeDoc & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMBPicking1.AddItem Trim(vRecordset.Fields("Name").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
        
        '---------------------------------------------------------------------------------------------------------------------
        
'Public Const vbViolet = &HFF8080
'Public Const vbVioletBright = &HFFC0C0
'Public Const vbForestGreen = &H228B22
'Public Const vbGray = &HE0E0E0
'Public Const vbLightBlue = &HFFD3A4
'Public Const vbLightGreen = &HABFCBD
'Public Const vbGreenLemon = &HB3FFBE
'Public Const vbYellowBright = &HC0FFFF
'Public Const vbOrange = &H2CCDFC

        
Call SetListViewColor(ListViewSaleOrder, PicPoint, vbWhite, vbLightBlue)
Call SetListViewColor(ListViewSaleOrderLastQue, PicPoint, vbWhite, vbLightGreen)
Call SetListViewColor(ListViewSaleOrderQueInformation, PicPoint, vbWhite, vbLightGreen)
Call SetListViewColor(ListViewPicking, PicPoint1, vbWhite, vbLightGreen)

'Call SetListViewColor(ListViewPicking, PICPoint, vbWhite, vbGray)
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Function RefreshData()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim aaa As Integer
Dim i As Integer
Dim DocListItem As ListItem
Dim vDocNo, vNewDoc As String
Dim vPrintStatus As Integer
Dim CountRecordset As Integer, CountList As Integer
Dim vDocHeader As String
Dim vUserprint1 As String

On Error Resume Next

Me.PBSaleOrder.Value = 0
vDocHeader = "SO"
vUserprint1 = Left(Trim(CMBSale.Text), InStr(Trim(CMBSale.Text), "-") - 1)
vUserPrint = vUserprint1
ListViewPicking.ListItems.Clear

'vQuery = "Select  *  from BCNP.dbo.vw_sl_00002   where Printed = 0 " _
'& " and salecode = '" & vUserprint1 & "' and DoctypeID = '" & vDocHeader & "' order by LastPrintDateTime "

vQuery = "exec dbo.USP_SO_SearchDocumentToPrint  1,'" & vUserprint1 & "','" & vDocHeader & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  If Not vRecordset.EOF Then
  CountRecordset = vRecordset.RecordCount
  Me.PBSaleOrder.Max = CountRecordset
  vRecordset.MoveFirst
  For i = 1 To CountRecordset
    Set DocListItem = ListViewPicking.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
    DocListItem.SubItems(1) = Trim(vRecordset.Fields("arname").Value)
    DocListItem.SubItems(2) = Trim(vRecordset.Fields("duedate").Value)
    DocListItem.SubItems(3) = Trim(vRecordset.Fields("lastprintdatetime").Value)
    Me.PBSaleOrder.Value = i
    vRecordset.MoveNext
  Next i
  End If
End If
vRecordset.Close

End Function

Function NewListItems()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim INVItemLists As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription
    vTypeDoc = "SO"
    ListViewPicking.ListItems.Clear
    
    'vQuery = "Select Docno,name1  from BCNP.dbo.vw_sl_00002  where Printed = 0  " _
                        '& " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
                        
        vQuery = "exec dbo.USP_SO_SearchDocumentToPrint 0,'" & vUserID & "','" & vTypeDoc & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListViewPicking.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                INVItemLists.SubItems(1) = Trim(vRecordset.Fields("arname").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close

        '----------------------------------------------------------------------------------------------------------------------
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Function

Public Sub vGetCustomerZone()
If Me.OptNormal.Value = True Then
  vCustomerZone = 0
ElseIf Me.OptMain.Value = True Then
  vCustomerZone = 1
ElseIf Me.OptOutLet.Value = True Then
  vCustomerZone = 2
End If
End Sub

Private Sub ListViewPicking_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXTPicking1.Caption = Item
End Sub


Public Sub PrintPicking_AVL(vCount As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vNamePrint As String
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vPosition As Integer
Dim vAddTime As Date
Dim vRequestTime As String
Dim vTimeID As Integer
Dim vSoStatus As Integer
Dim vTimePick As Integer


On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    If vUserPrint <> "" Then
       vNamePrint = Trim(vUserPrint)
    Else
       vNamePrint = Me.LBLSaleCode.Caption
    End If
    vShelfGroup = Trim(UCase("AVL"))
    vZoneID = Trim("03")
    'vZoneID = Trim("02")
    vWHCode = Trim("S01")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    vTimeID = vCount
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
 
    vSoStatus = Me.LBLSaleType.Caption
    If vSoStatus = 1 Then
       If Me.OPTReserve.Value = True Then
          vTimePick = 1
       ElseIf Me.OPTPayItem.Value = True Then
          vTimePick = 0
       End If
    End If
        
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    
    On Error GoTo ErrRollBack
    vPosition = 2
    Call vGetCustomerZone
    
    
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement_New '" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vTimeID & ",0,'" & vRequestTime & "'," & vCustomerZone & "," & vTimePick & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','" & vShelfGroup & "' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 1)
  Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 2)
  
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าชั้นเก็บ AVL ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_BK1(vCount As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vNamePrint As String
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vPosition As Integer
Dim vAddTime As Date
Dim vRequestTime As String
Dim vTimeID As Integer
Dim vSoStatus As Integer
Dim vTimePick As Integer


On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    If vUserPrint <> "" Then
       vNamePrint = Trim(vUserPrint)
    Else
       vNamePrint = Me.LBLSaleCode.Caption
    End If
    vShelfGroup = Trim(UCase("BK1"))
    vZoneID = Trim("01")
    vWHCode = Trim("S01")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    vTimeID = vCount
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vSoStatus = Me.LBLSaleType.Caption
    If vSoStatus = 1 Then
       If Me.OPTReserve.Value = True Then
          vTimePick = 1
       ElseIf Me.OPTPayItem.Value = True Then
          vTimePick = 0
       End If
    End If
    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    On Error GoTo ErrRollBack
    vPosition = 2
    Call vGetCustomerZone
    
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement_New '" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vTimeID & ",0,'" & vRequestTime & "'," & vCustomerZone & "," & vTimePick & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','" & vShelfGroup & "' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
  
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าชั้นเก็บ BK1 ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_BK1_Sunday(vCount As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vNamePrint As String
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vPosition As Integer
Dim vAddTime As Date
Dim vRequestTime As String
Dim vTimeID As Integer
Dim vSoStatus As Integer
Dim vTimePick As Integer


On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    If vUserPrint <> "" Then
       vNamePrint = Trim(vUserPrint)
    Else
       vNamePrint = Me.LBLSaleCode.Caption
    End If
    vShelfGroup = Trim(UCase("BK1"))
    vZoneID = Trim("02")
    vWHCode = Trim("S01")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    vTimeID = vCount
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vSoStatus = Me.LBLSaleType.Caption
    If vSoStatus = 1 Then
       If Me.OPTReserve.Value = True Then
          vTimePick = 1
       ElseIf Me.OPTPayItem.Value = True Then
          vTimePick = 0
       End If
    End If
    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    On Error GoTo ErrRollBack
    vPosition = 2
    Call vGetCustomerZone
    
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement_New '" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vTimeID & ",0,'" & vRequestTime & "'," & vCustomerZone & "," & vTimePick & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','" & vShelfGroup & "' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าชั้นเก็บ BK1 ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_BK2(vCount As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vNamePrint As String
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vPosition As Integer
Dim vAddTime As Date
Dim vRequestTime As String
Dim vTimeID As Integer
Dim vSoStatus As Integer
Dim vTimePick As Integer


On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    If vUserPrint <> "" Then
       vNamePrint = Trim(vUserPrint)
    Else
       vNamePrint = Me.LBLSaleCode.Caption
    End If
    vShelfGroup = Trim(UCase("BK2"))
    vZoneID = Trim("01")
    vWHCode = Trim("S01")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    vTimeID = vCount
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vSoStatus = Me.LBLSaleType.Caption
    If vSoStatus = 1 Then
       If Me.OPTReserve.Value = True Then
          vTimePick = 1
       ElseIf Me.OPTPayItem.Value = True Then
          vTimePick = 0
       End If
    End If
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

    
    On Error GoTo ErrRollBack
    vPosition = 2
    Call vGetCustomerZone
    
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement_New '" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vTimeID & ",0,'" & vRequestTime & "'," & vCustomerZone & "," & vTimePick & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','" & vShelfGroup & "' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
  
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าชั้นเก็บ BK2 ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_BK2_Sunday(vCount As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vNamePrint As String
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vPosition As Integer
Dim vAddTime As Date
Dim vRequestTime As String
Dim vTimeID As Integer
Dim vSoStatus As Integer
Dim vTimePick As Integer


On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    If vUserPrint <> "" Then
       vNamePrint = Trim(vUserPrint)
    Else
       vNamePrint = Me.LBLSaleCode.Caption
    End If
    vShelfGroup = Trim(UCase("BK2"))
    vZoneID = Trim("02")
    vWHCode = Trim("S01")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    vTimeID = vCount
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vSoStatus = Me.LBLSaleType.Caption
    If vSoStatus = 1 Then
       If Me.OPTReserve.Value = True Then
          vTimePick = 1
       ElseIf Me.OPTPayItem.Value = True Then
          vTimePick = 0
       End If
    End If
    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    On Error GoTo ErrRollBack
    vPosition = 2
    Call vGetCustomerZone
    
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement_New '" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vTimeID & ",0,'" & vRequestTime & "'," & vCustomerZone & "," & vTimePick & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','" & vShelfGroup & "' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าชั้นเก็บ BK2 ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub


Public Sub PrintPicking_BK3(vCount As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vNamePrint As String
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vPosition As Integer
Dim vAddTime As Date
Dim vRequestTime As String
Dim vTimeID As Integer
Dim vSoStatus As Integer
Dim vTimePick As Integer

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    If vUserPrint <> "" Then
       vNamePrint = Trim(vUserPrint)
    Else
       vNamePrint = Me.LBLSaleCode.Caption
    End If
    vShelfGroup = Trim(UCase("BK3"))
    vZoneID = Trim("02")
    vWHCode = Trim("S01")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    vTimeID = vCount
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vSoStatus = Me.LBLSaleType.Caption
    If vSoStatus = 1 Then
       If Me.OPTReserve.Value = True Then
          vTimePick = 1
       ElseIf Me.OPTPayItem.Value = True Then
          vTimePick = 0
       End If
    End If
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

    
    On Error GoTo ErrRollBack
    vPosition = 2
    Call vGetCustomerZone
    
    
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement_New '" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vTimeID & ",0,'" & vRequestTime & "'," & vCustomerZone & "," & vTimePick & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','" & vShelfGroup & "' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 1)
  Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าชั้นเก็บ AVL ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_SPO(vCount As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vNamePrint As String
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vPosition As Integer
Dim vAddTime As Date
Dim vRequestTime As String
Dim vTimeID As Integer
Dim vSoStatus As Integer
Dim vTimePick As Integer


On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    If vUserPrint <> "" Then
       vNamePrint = Trim(vUserPrint)
    Else
       vNamePrint = Me.LBLSaleCode.Caption
    End If
    vShelfGroup = Trim(UCase("SPO"))
    
    If DatePart("w", Now) <> 1 Then
       If Me.OptOutLet.Value = True Then
          vZoneID = Trim("02")
       ElseIf Me.OptMain.Value = True Then
          vZoneID = Trim("01")
       End If
    Else
          vZoneID = Trim("02")
    End If
    
    vWHCode = Trim("S01")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    vTimeID = vCount
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vSoStatus = Me.LBLSaleType.Caption
    If vSoStatus = 1 Then
       If Me.OPTReserve.Value = True Then
          vTimePick = 1
       ElseIf Me.OPTPayItem.Value = True Then
          vTimePick = 0
       End If
    End If
    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    On Error GoTo ErrRollBack
    vPosition = 2
    Call vGetCustomerZone
    
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement_New '" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vTimeID & ",0,'" & vRequestTime & "'," & vCustomerZone & "," & vTimePick & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','" & vShelfGroup & "' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  
  If DatePart("w", Now) <> 1 Then
     If Me.OptMain.Value = True Then
        Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
     ElseIf Me.OptOutLet.Value = True Then
        Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
     End If
  Else
        Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  End If
  
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าชั้นเก็บ AVL ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintSalePickingSlip(vQueID As Integer, vQueDocDate As String, vZone As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
Dim vQueNo As String
   

If vZone = "A" Or vZone = "X" Then
 vQuery = "exec dbo.USP_NP_SearchPrinter 2"
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
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
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
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


If vZone = "C" Then
 vQuery = "exec dbo.USP_NP_SearchPrinter 4"
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
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

Public Sub PrintSalePicking_A(vDocNo As String, vTimeID As Integer, vType As Integer)
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


'On Error GoTo ErrRollBack


If vDocNo <> "" Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  31"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
   
    If vType = 2 Then
       vQueZone = "A"
       vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       vDocNo = Me.TBDocNo.Text
       vDocdate = Me.LBLOrderDocDate.Caption
       vARCode = Me.LBLOrderArCode.Caption
   
       If Me.LBLOrderSaleCode.Caption <> "" Then
       vSaleCode = Left(Me.LBLOrderSaleCode.Caption, InStr(Me.LBLOrderSaleCode.Caption, "/") - 1)
       End If
       
       vRefNo = Me.TBOrderRefNo.Text
       vMemberID = Me.LBLOrderMember.Caption
       vSourceID = vType
       vIsConditionSend = 1
       
       If Me.MEBReqTime.Text = "__:__" Then
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
       Else
       vQueReqTime = Me.MEBReqTime.Text
       End If
       
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
            
          
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  31 "
    gConnection.Execute vQuery
    
    'vQuery = "commit tran"
    'gConnection.Execute vQuery
    
  'Call PrintPickingHeader(vQueID, vQueDocDate, vZoneID)
  'Call PrintSalePickingSlip(vQueID, vQueDocDate, vQueZone)
  
  
  vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto " & 2 & ",'" & vDocNo & "','" & vQueDocDate & "','" & vQueZone & "'," & vQueID & " ,'" & vUserID & "'"
  gConnection.Execute vQuery
  
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

Public Sub PrintSalePicking_B(vDocNo As String, vTimeID As Integer, vType As Integer)
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


'On Error GoTo ErrRollBack


If vDocNo <> "" Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  31"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
           
    If vType = 2 Then
       vQueZone = "B"
       vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       vDocNo = Me.TBDocNo.Text
       vDocdate = Me.LBLOrderDocDate.Caption
       vARCode = Me.LBLOrderArCode.Caption
   
       If Me.LBLOrderSaleCode.Caption <> "" Then
       vSaleCode = Left(Me.LBLOrderSaleCode.Caption, InStr(Me.LBLOrderSaleCode.Caption, "/") - 1)
       End If
       
       vRefNo = Me.TBOrderRefNo.Text
       vMemberID = Me.LBLOrderMember.Caption
       vSourceID = vType
       vIsConditionSend = 1
       
       If Me.MEBReqTime.Text = "__:__" Then
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
       Else
       vQueReqTime = Me.MEBReqTime.Text
       End If
       
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
            
       
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  31 "
    gConnection.Execute vQuery
    
    'vQuery = "commit tran"
    'gConnection.Execute vQuery

  'Call PrintPickingHeader(vQueID, vQueDocDate, vZoneID)
  'Call PrintSalePickingSlip(vQueID, vQueDocDate, vQueZone)

  vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto " & 2 & ",'" & vDocNo & "','" & vQueDocDate & "','" & vQueZone & "'," & vQueID & " ,'" & vUserID & "'"
  gConnection.Execute vQuery
  
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


Public Sub PrintSalePicking_C(vDocNo As String, vTimeID As Integer, vType As Integer)
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


'On Error GoTo ErrRollBack


If vDocNo <> "" Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  31"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
       
    If vType = 2 Then
       vQueZone = "C"
      vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       vDocNo = Me.TBDocNo.Text
       vDocdate = Me.LBLOrderDocDate.Caption
       vARCode = Me.LBLOrderArCode.Caption
   
       If Me.LBLOrderSaleCode.Caption <> "" Then
       vSaleCode = Left(Me.LBLOrderSaleCode.Caption, InStr(Me.LBLOrderSaleCode.Caption, "/") - 1)
       End If
       
       vRefNo = Me.TBOrderRefNo.Text
       vMemberID = Me.LBLOrderMember.Caption
       vSourceID = vType
       vIsConditionSend = 1
       
       If Me.MEBReqTime.Text = "__:__" Then
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
       Else
       vQueReqTime = Me.MEBReqTime.Text
       End If
       
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
            

       
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  31 "
    gConnection.Execute vQuery
    
    'vQuery = "commit tran"
    'gConnection.Execute vQuery

  'Call PrintPickingHeader(vQueID, vQueDocDate, vZoneID)
  'Call PrintSalePickingSlip(vQueID, vQueDocDate, vQueZone)

  vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto " & 2 & ",'" & vDocNo & "','" & vQueDocDate & "','" & vQueZone & "'," & vQueID & " ,'" & vUserID & "'"
  gConnection.Execute vQuery
  
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

Public Sub PrintSalePicking_X(vDocNo As String, vTimeID As Integer, vType As Integer)
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


'On Error GoTo ErrRollBack


If vDocNo <> "" Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  31"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
   If vType = 2 Then
       vQueZone = "X"
       vQueDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
       vDocNo = Me.TBDocNo.Text
       vDocdate = Me.LBLOrderDocDate.Caption
       vARCode = Me.LBLOrderArCode.Caption
   
       If Me.LBLOrderSaleCode.Caption <> "" Then
       vSaleCode = Left(Me.LBLOrderSaleCode.Caption, InStr(Me.LBLOrderSaleCode.Caption, "/") - 1)
       End If
       
       vRefNo = Me.TBOrderRefNo.Text
       vMemberID = Me.LBLOrderMember.Caption
       vSourceID = vType
       vIsConditionSend = 1
       
       If Me.MEBReqTime.Text = "__:__" Then
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
       Else
       vQueReqTime = Me.MEBReqTime.Text
       End If
       
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

    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  31 "
    gConnection.Execute vQuery
    
    'vQuery = "commit tran"
    'gConnection.Execute vQuery

  'Call PrintPickingHeader(vQueID, vQueDocDate, vZoneID)
  'Call PrintSalePickingSlip(vQueID, vQueDocDate, vQueZone)

  vQuery = "exec dbo.USP_NP_InsertPrintNopadolSystemAuto " & 2 & ",'" & vDocNo & "','" & vQueDocDate & "','" & vQueZone & "'," & vQueID & " ,'" & vUserID & "'"
  gConnection.Execute vQuery
  
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



Public Sub PrintPicking_A()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vNamePrint As String
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vPosition As Integer
Dim vAddTime As Date
Dim vRequestTime As String


On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    vNamePrint = Trim(vUserPrint)
    vShelfGroup = Trim(UCase("A"))
    vZoneID = Trim("01")
    vWHCode = Trim("010")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

    
    On Error GoTo ErrRunQueueID
    vPosition = 2
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement'" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "',1,0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','A' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
  
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าคลัง 010 โซน A ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_B()
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vQuery As String
Dim vNamePrint As String
Dim vPosition As Integer
Dim vRequestTime As String
Dim vAddTime As Date

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    vNamePrint = Trim(vUserPrint)
    vShelfGroup = Trim(UCase("B"))
    vZoneID = Trim("01")
    vWHCode = Trim("010")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)

    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

    
    On Error GoTo ErrRunQueueID
    vPosition = 2
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement '" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "',1,0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','B' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    
  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
  
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าคลัง 010 โซน B ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_C()
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vQuery As String
Dim vNamePrint As String
Dim vPosition As Integer
Dim vRequestTime As String
Dim vAddTime As Date

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    vNamePrint = Trim(vUserPrint)
    vShelfGroup = Trim(UCase("C"))
    vZoneID = Trim("01")
    vWHCode = Trim("015")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

On Error GoTo ErrRunQueueID
    vPosition = 2
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement'" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "',1,0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','C' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    
  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าคลัง 015 โซน C ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub


Public Sub PrintPicking_D()
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vQuery As String
Dim vNamePrint As String
Dim vPosition As Integer
Dim vRequestTime As String
Dim vAddTime As Date


On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    vNamePrint = Trim(vUserPrint)
    vShelfGroup = Trim(UCase("D"))
    vZoneID = Trim("01")
    vWHCode = Trim("010")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo  27"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

    On Error GoTo ErrRunQueueID
    vPosition = 2
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement'" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "',1,0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','D' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    
  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าคลัง 010 โซน D ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub SaleOrder()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vReportName As String
Dim vDocNo As String, vWHCode As String
Dim vGroupDoc As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vCheckDocNo As Integer
Dim vHeaderType As Integer
Dim vCheckBillType As Integer
Dim vRunNumber As String
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer

On Error GoTo ErrDescription
        
Call GetComputerandUser

vDocNo = UCase(Trim(TXTPicking1.Caption))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vGroupDoc = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
   vGroupDoc = UCase(Left(vDocNo, 3))
End If

If UCase(vGroupDoc) = "SHV" Or UCase(vGroupDoc) = "SHN" Or UCase(vGroupDoc) = "SCV" Or UCase(vGroupDoc) = "SCN" Or UCase(vGroupDoc) = "SAB" Or UCase(vGroupDoc) = "SVE" Then 'Or UCase(vGroupDoc) = "SVD" Or UCase(vGroupDoc) = "SVN" Or UCase(vGroupDoc) = "SVM" Then
        vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 'เคยพิมพ์แล้ว
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
        Else
            vCheckDocNo = 0 'ยังไม่ได้พิมพ์
        End If
        vRecordset.Close
        
        If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
            vQuery = "select docno,billtype from bcnp.dbo.bcsaleorder where docno = '" & vDocNo & "' and sostatus = 0 "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckBillType = Trim(vRecordset.Fields("billtype").Value)
            End If
            vRecordset.Close
            If vCheckBillType = 0 Then
                vHeaderType = 13
            ElseIf vCheckBillType = 1 Then
                vHeaderType = 14
            End If
            
            
            vNamePrint = Trim(vUserPrint)
            vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                vHeader = Trim(vRecordset.Fields("header").Value)
                vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
            End If
            vRecordset.Close
            vDocuments = vDocNumber & vHeader & "-" & vPicking
            
            vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
            gConnection.Execute vQuery
            
            vTypeNumber = 1
            vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vTypeNumber = 1
            vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
            vQuery = "exec bcnp.dbo.USP_NP_InsertLogPrintRunningRes '" & vDocNo & "','" & vUserID & "' ," & vTypeNumber & " "
            gConnection.Execute vQuery
        End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
        
            vRepType = "SO"
                       
            If (UCase(vDepartment) = "CH" Or UCase(vDepartment) = "CR" Or UCase(vDepartment) = "IS" Or UCase(vDepartment) = "PC") And Me.CHKReqPrint.Value = 0 Then
               If vGroupDoc = "SHV" Or vGroupDoc = "shv" Or vGroupDoc = "SCV" Or vGroupDoc = "scv" Or vGroupDoc = "SAB" Or vGroupDoc = "sab" Or vGroupDoc = "SVE" Or vGroupDoc = "sve" Then 'Or vGroupDoc = "SVM" Or vGroupDoc = "svm" Or vGroupDoc = "SVD" Or vGroupDoc = "svd"
                  If ShowDiscount.Value = False Then
                  vRepID = 398
                  Else
                  vRepID = 402
                  End If
               ElseIf vGroupDoc = "SHN" Or vGroupDoc = "shn" Or vGroupDoc = "SCN" Or vGroupDoc = "scn" Then 'Or vGroupDoc = "SVN" Or vGroupDoc = "svn" Then
                  If ShowDiscount.Value = False Then
                  vRepID = 399
                  Else
                  vRepID = 403
                  End If
               Else
                  MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                  Exit Sub
               End If
            Else
               If vGroupDoc = "SHV" Or vGroupDoc = "shv" Or vGroupDoc = "SCV" Or vGroupDoc = "scv" Or vGroupDoc = "SAB" Or vGroupDoc = "sab" Or vGroupDoc = "SVE" Or vGroupDoc = "sve" Then 'Or vGroupDoc = "SVM" Or vGroupDoc = "svm" Or vGroupDoc = "SVD" Or vGroupDoc = "svd"
                   If ShowDiscount.Value = False Then
                   vRepID = 57
                   Else
                   vRepID = 142
                   End If
                   
                  If Me.CKUnShowPrice.Value = False Then
                   vRepID = 57
                   Else
                   vRepID = 488
                   End If
               ElseIf vGroupDoc = "SHN" Or vGroupDoc = "shn" Or vGroupDoc = "SCN" Or vGroupDoc = "scn" Then 'Or vGroupDoc = "SVN" Or vGroupDoc = "svn" Then
                   If ShowDiscount.Value = False Then
                   vRepID = 58
                   Else
                   vRepID = 143
                   End If
                   
                   If CKUnShowPrice.Value = False Then
                   vRepID = 58
                   Else
                   vRepID = 489
                   End If
                   
               Else
               MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
               Exit Sub
               End If
            End If
            
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrysPicking1
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
Else
    MsgBox "ไม่สามารถพิมพ์ใบสั่งขายได้ คุณเลือกพิมพ์เอกสารผิด"
    Exit Sub
End If
                            '------------------------------------------------------------------------------------------------------
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If


End Sub


Public Sub SaleOrder_Delivery()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String, vWHCode As String
        Dim vGroupDoc As String
        Dim vPrint As Integer
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        'Call GetComputerandUser
        vDocNo = UCase(Trim(TXTPicking1.Caption))
        
        vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
                vPrint = Trim(vRecordset.Fields("printed").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------------------
                            vRepType = "SO"
                                vRepID = 76
       vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
       'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrysPicking1
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '------------------------------------------------------------------------------------------------------
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintPicking_M()
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vQuery As String
Dim vNamePrint As String
Dim vPosition As Integer
Dim vRequestTime As String
Dim vAddTime As Date

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    vNamePrint = Trim(vUserPrint)
    vShelfGroup = Trim(UCase("M"))
    vZoneID = Trim("02")
    vWHCode = Trim("014")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

    On Error GoTo ErrRunQueueID
    vPosition = 2
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement'" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "',1,0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','M' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    
   'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 1)
   'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
   MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าคลัง 014 โซน M ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub



Public Sub CheckValue()
Dim i As Integer
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String

vDocNo = Trim(TXTPicking1.Caption)
vQuery = "select typecode from bcnp.dbo.vw_IV_PackingSlip where docno = '" & vDocNo & "' and shelfgroup1 = 'M' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckValue = True
Else
    vCheckValue = False
End If
vRecordset.Close


End Sub

Public Sub PrintPicking_M_OutLet()
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vQuery As String
Dim vNamePrint As String
Dim vPosition As Integer
Dim vRequestTime As String
Dim vAddTime As Date

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    vNamePrint = Trim(vUserPrint)
    vShelfGroup = Trim(UCase("M"))
    vZoneID = Trim("02")
    vWHCode = Trim("014")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

    On Error GoTo ErrRunQueueID
    vPosition = 2
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement'" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "',1,0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','M' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 1)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ออกที่โซน OutLet", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าคลัง 014 โซน M ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_K_BackStock()
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vQuery As String
Dim vNamePrint As String
Dim vPosition As Integer
Dim vRequestTime As String
Dim vAddTime As Date

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    vNamePrint = Trim(vUserPrint)
    vShelfGroup = Trim(UCase("K"))
    vZoneID = Trim("02")
    vWHCode = Trim("014")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

    On Error GoTo ErrRunQueueID
    vPosition = 2
    
    Call vGetCustomerZone
    
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement'" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "',1,0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','K' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 1)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ออกที่โซน OutLet", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าคลัง 014 โซน M ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub CheckValueHMX()
Dim i As Integer
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String

vDocNo = Trim(TXTPicking1.Caption)
vQuery = "select itemcode from bcnp.dbo.bcsaleordersub  where docno = '" & vDocNo & "' and  whcode = '014' and  typecode not in (select itemtype from npmaster.dbo.NP_ItemOutLet) "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckValue1 = True
Else
    vCheckValue1 = False
End If
vRecordset.Close

End Sub

Public Sub PrintPicking_M_HMX()
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vQuery As String
Dim vNamePrint As String
Dim vPosition As Integer
Dim vRequestTime As String
Dim vAddTime As Date

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    vNamePrint = Trim(vUserPrint)
    vShelfGroup = Trim(UCase("H"))
    vZoneID = Trim("02")
    vWHCode = Trim("014")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)

    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

    On Error GoTo ErrRunQueueID
    vPosition = 2
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement'" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "',1,0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery

    
    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','H' "
    gConnection.Execute vQuery
    
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 1)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าคลัง 014 โซน H ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_Y()
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vQuery As String
Dim vNamePrint As String
Dim vPosition As Integer
Dim vRequestTime As String
Dim vAddTime As Date

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    vNamePrint = Trim(vUserPrint)
    vShelfGroup = Trim(UCase("Y"))
    vZoneID = Trim("01")
    vWHCode = Trim("016")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)
    
    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

    On Error GoTo ErrRunQueueID
    vPosition = 2
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement'" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "',1,0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery
    

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    
  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าคลัง 016 โซน Y ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub


Public Sub SaleOrder_Reserve()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vReportName As String
Dim vDocNo As String, vWHCode As String
Dim vGroupDoc As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vCheckDocNo As String
Dim vSoStatus As Integer
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer


On Error GoTo ErrDescription
        
vDocNo = UCase(Trim(TXTPicking1.Caption))
vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vGroupDoc = UCase(Trim(vRecordset.Fields("groupdoc").Value))
        vPrint = Trim(vRecordset.Fields("printed").Value)
End If
vRecordset.Close
        
If UCase(vGroupDoc) = "ROV" Or UCase(vGroupDoc) = "RON" Or UCase(vGroupDoc) = "RWV" Or UCase(vGroupDoc) = "RWN" Then
        vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 'เคยพิมพ์แล้ว
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
        Else
            vCheckDocNo = 0 'ยังไม่ได้พิมพ์
        End If
        vRecordset.Close
        
        If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
            vQuery = "select docno,billtype from bcnp.dbo.bcsaleorder where docno = '" & vDocNo & "' and sostatus =1"
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckBillType = Trim(vRecordset.Fields("billtype").Value)
            End If
            vRecordset.Close
            If vCheckBillType = 0 Then
                vHeaderType = 20
            ElseIf vCheckBillType = 1 Then
                vHeaderType = 21
            End If
            
            
            vNamePrint = Trim(vUserPrint)
            vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                vHeader = Trim(vRecordset.Fields("header").Value)
                vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
            End If
            vRecordset.Close
            vDocuments = vDocNumber & vHeader & "-" & vPicking
            
            vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
            gConnection.Execute vQuery
            
            vTypeNumber = 2
            vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vTypeNumber = 2
            vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
            vQuery = "exec bcnp.dbo.USP_NP_InsertLogPrintRunningRes '" & vDocNo & "','" & vUserID & "' ," & vTypeNumber & " "
            gConnection.Execute vQuery
        End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
        
        
        vQuery = "select sostatus from bcnp.dbo.bcsaleorder where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vSoStatus = vRecordset.Fields("sostatus").Value
        End If
        vRecordset.Close
        
        If vSoStatus = 1 Then
        vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = UCase(Trim(vRecordset.Fields("groupdoc").Value))
                vPrint = Trim(vRecordset.Fields("printed").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------------------
        
If (UCase(vDepartment) = "CH" Or UCase(vDepartment) = "CR" Or UCase(vDepartment) = "IS" Or UCase(vDepartment) = "PC") And Me.CHKReqPrint.Value = 0 Then
        If vGroupDoc = "ROV" Or vGroupDoc = "rov" Or UCase(vGroupDoc) = "RWV" Or UCase(vGroupDoc) = "RWN" Then
            vRepID = 396
        ElseIf vGroupDoc = "RON" Or vGroupDoc = "ron" Or UCase(vGroupDoc) = "RWV" Or UCase(vGroupDoc) = "RWN" Then
            vRepID = 397
        End If
Else
        If vGroupDoc = "ROV" Or vGroupDoc = "rov" Or UCase(vGroupDoc) = "RWV" Or UCase(vGroupDoc) = "RWN" Then
            If Me.CKUnShowPrice.Value = False Then
                vRepID = 224
                Else
                vRepID = 490
            End If
        ElseIf vGroupDoc = "RON" Or vGroupDoc = "ron" Or UCase(vGroupDoc) = "RWV" Or UCase(vGroupDoc) = "RWN" Then
            If Me.CKUnShowPrice.Value = False Then
                vRepID = 225
                Else
                vRepID = 491
            End If
        End If
End If
        
        vRepType = "RO"
        
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrysPicking1
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '------------------------------------------------------------------------------------------------------
    vQuery = "exec dbo.usp_so_UpdateIsConfirmPrint '" & vDocNo & "' "
    gConnection.Execute vQuery
    Else
        MsgBox "เอกสารเลขที่ " & vDocNo & " ไม่ได้เป็นใบสั่งจองสินค้า กรุณาตรวจสอบ"
    End If
Else
    MsgBox "ไม่สามารถพิมพ์ใบสั่งจองได้ กรุณาเลือกพิมพ์เอกสารใหม่"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintPicking_M_HMX_CustReceive()
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vWHCode As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vQuery As String
Dim vNamePrint As String
Dim vPosition As Integer
Dim vRequestTime As String
Dim vAddTime As Date

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    vNamePrint = Trim(vUserPrint)
    vShelfGroup = Trim(UCase("H"))
    vZoneID = Trim("02")
    vWHCode = Trim("014")
    vDocType = 1
    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = Left(Trim(Me.LBLARName.Caption), InStr(Trim(Me.LBLARName.Caption), "//") - 1)

    If Me.OptNow.Value = True Then
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
    ElseIf Me.OptSchedule.Value = True Then
      vRequestTime = Me.MEDTime.Text
    ElseIf Me.OptTomorrow.Value = True Then
      vRequestTime = "Tomorrow"
    End If
    
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    
    vDocuments = ""
    
    vQuery = "begin tran"
    gConnection.Execute vQuery

    On Error GoTo ErrRunQueueID
    vPosition = 2
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement '" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "',1,0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','H' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 1)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าคลัง 014 โซน H ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Private Sub ListViewSaleOrder_DblClick()
Dim vIndex As Integer
Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Double
Dim vDiscountAmount As Double
Dim vNetAmount As Double
Dim vPrice As Double
Dim vRemainQty As Double

If Me.ListViewSaleOrder.ListItems.Count > 0 Then
   vIndex = Me.ListViewSaleOrder.SelectedItem.Index
   Me.PICEditOrder.Visible = True
   
   Me.LBLEditIndex.Caption = vIndex
   
   Me.LBLEditItemCode.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(1)
   Me.LBLEditItemName.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(2)
   Me.LBLEditUnitCode.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(4)
   
   
   vQTY = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(3)
   vPrice = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(5)
   vDiscountAmount = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(6)
   vNetAmount = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(7)
   vRemainQty = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(14)
   
   Me.LBLEditItemQty.Caption = Format(vQTY, "##,##0.00")
   Me.LBLEditPrice.Caption = Format(vPrice, "##,##0.00")
   Me.LBLEditDiscount.Caption = Format(vDiscountAmount, "##,##0.00")
   Me.LBLEditItemAmount.Caption = Format(vNetAmount, "##,##0.00")
   Me.LBLEditRemain.Caption = Format(vRemainQty, "##,##0.00")
   
   Me.TBDocNo.Enabled = False
   Me.TBOrderRefNo.Enabled = False
   Me.MEBReqTime.Enabled = False
   
   Me.PICEditOrder.Visible = True
   Me.TBEditQty.Text = vQTY
   Me.TBEditQty.SetFocus

End If
End Sub

Private Sub ListViewSaleOrder_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer

If KeyCode = 46 Then
   If Me.ListViewSaleOrder.ListItems.Count > 0 And Me.ListViewSaleOrder.ListItems.Count <> 1 Then
      vIndex = Me.ListViewSaleOrder.SelectedItem.Index
      Me.ListViewSaleOrder.ListItems.Remove (vIndex)
      Call GenIDLineItem
      Call CalcEditItemQty
   End If
   If Me.ListViewSaleOrder.ListItems.Count = 1 Then
      MsgBox "ไม่สามารถลบรายการสินค้าทั้งหมดได้ กรณีไม่เอา ก็ให้ปิดหน้าจอเท่านั้น", vbCritical, "Send Error Message"
   End If
End If


If KeyCode = 116 Then
Call CMDSaleOrderSendQue_Click
End If
End Sub

Public Sub GenIDLineItem()
Dim i As Integer

For i = 1 To Me.ListViewSaleOrder.ListItems.Count

Me.ListViewSaleOrder.ListItems(i).Text = i
Next i
End Sub

Private Sub ListViewSaleOrder_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer
Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Double
Dim vDiscountAmount As Double
Dim vNetAmount As Double
Dim vPrice As Double
Dim vRemainQty As Double

If KeyAscii = 13 Then
    If Me.ListViewSaleOrder.ListItems.Count > 0 Then
       vIndex = Me.ListViewSaleOrder.SelectedItem.Index
       Me.PICEditOrder.Visible = True
       
       Me.LBLEditIndex.Caption = vIndex
       
       Me.LBLEditItemCode.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(1)
       Me.LBLEditItemName.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(2)
       Me.LBLEditUnitCode.Caption = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(4)
       
       vQTY = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(3)
       vPrice = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(5)
       vDiscountAmount = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(6)
       vNetAmount = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(7)
       vRemainQty = Me.ListViewSaleOrder.ListItems(vIndex).SubItems(14)
       
       Me.LBLEditItemQty.Caption = Format(vQTY, "##,##0.00")
       Me.LBLEditPrice.Caption = Format(vPrice, "##,##0.00")
       Me.LBLEditDiscount.Caption = Format(vDiscountAmount, "##,##0.00")
       Me.LBLEditItemAmount.Caption = Format(vNetAmount, "##,##0.00")
       Me.LBLEditRemain.Caption = Format(vRemainQty, "##,##0.00")
       
       Me.PICEditOrder.Visible = True
       Me.TBEditQty.Text = vQTY
       Me.TBEditQty.SetFocus
    
    End If
End If
End Sub

Private Sub ListViewSelectItemPicking_DblClick()
Dim i As Integer
Dim vRecQTY As String
Dim vCheckQty As Double
Dim vPickQty As Double
Dim vCheckNumber As Boolean
Dim vGetPickQTY As Double

If Me.ListViewSelectItemPicking.ListItems.Count > 0 Then
   i = Me.ListViewSelectItemPicking.SelectedItem.Index
   vCheckQty = Me.ListViewSelectItemPicking.ListItems(i).SubItems(3)
   vGetPickQTY = Me.ListViewSelectItemPicking.ListItems(i).SubItems(4)
   vRecQTY = InputBox("สั่งจัดจำนวน :", "กรอกจำนวนที่ต้องการจัดสินค้า", vGetPickQTY)
   
   CheckNumber (vRecQTY)
   
   If vRecQTY <> "" Then
      If vCheckValueNumber = True Then
         If vRecQTY = 0 Then
            Me.ListViewSelectItemPicking.ListItems(i).Checked = False
            Me.ListViewSelectItemPicking.ListItems(i).SubItems(4) = Format(vPickQty, "##,##0.00")
            Exit Sub
         ElseIf vRecQTY > 0 And Me.ListViewSelectItemPicking.ListItems(i).Checked = False Then
            Me.ListViewSelectItemPicking.ListItems(i).Checked = True
         End If
         vPickQty = vRecQTY
         If vPickQty <= vCheckQty And (vCheckQty - vPickQty) >= 0 Then
            Me.ListViewSelectItemPicking.ListItems(i).SubItems(4) = Format(vPickQty, "##,##0.00")
         Else
            MsgBox "สั่งจัดสินค้าเกินกว่าที่สั่งขาย", vbCritical, "Send Error Message"
            Me.ListViewSelectItemPicking.ListItems(i).Checked = False
         End If
      Else
        MsgBox "ต้องกรอกข้อมูลที่เกี่ยวกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      End If
   End If
End If

End Sub

Private Sub ListViewSelectItemPicking_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
Dim vRecQTY As String
Dim vCheckQty As Double
Dim vPickQty As Double
Dim vCheckNumber As Boolean
Dim vGetPickQTY As Double

If Me.ListViewSelectItemPicking.ListItems.Count > 0 Then
   i = Item.Index
   If Me.ListViewSelectItemPicking.ListItems(i).Checked = False Then
      Exit Sub
   End If

   vCheckQty = Me.ListViewSelectItemPicking.ListItems(i).SubItems(3)
   vGetPickQTY = Me.ListViewSelectItemPicking.ListItems(i).SubItems(4)
   vRecQTY = InputBox("สั่งจัดจำนวน :", "กรอกจำนวนที่ต้องการจัดสินค้า", vGetPickQTY)
   
   CheckNumber (vRecQTY)
   
   If vRecQTY <> "" Then
      If vCheckValueNumber = True Then
         If vRecQTY = 0 Then
            Me.ListViewSelectItemPicking.ListItems(i).Checked = False
            Me.ListViewSelectItemPicking.ListItems(i).SubItems(4) = Format(vPickQty, "##,##0.00")
            Exit Sub
         ElseIf vRecQTY > 0 And Me.ListViewSelectItemPicking.ListItems(i).Checked = False Then
            Me.ListViewSelectItemPicking.ListItems(i).Checked = True
         End If
         vPickQty = vRecQTY
         If vPickQty <= vCheckQty And (vCheckQty - vPickQty) >= 0 Then
            Me.ListViewSelectItemPicking.ListItems(i).SubItems(4) = Format(vPickQty, "##,##0.00")
         Else
            MsgBox "สั่งจัดสินค้าเกินกว่าที่สั่งขาย", vbCritical, "Send Error Message"
            Me.ListViewSelectItemPicking.ListItems(i).Checked = False
         End If
      Else
        MsgBox "ต้องกรอกข้อมูลที่เกี่ยวกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      End If
   End If
End If
End Sub

Private Sub ListViewSelectItemPicking_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim vRecQTY As String
Dim vCheckQty As Double
Dim vPickQty As Double
Dim vCheckNumber As Boolean
Dim vGetPickQTY As Double


If KeyAscii = 13 Then
   If Me.ListViewSelectItemPicking.ListItems.Count > 0 Then
      i = Me.ListViewSelectItemPicking.SelectedItem.Index
      vCheckQty = Me.ListViewSelectItemPicking.ListItems(i).SubItems(3)
      vGetPickQTY = Me.ListViewSelectItemPicking.ListItems(i).SubItems(4)
      vRecQTY = InputBox("สั่งจัดจำนวน :", "กรอกจำนวนที่ต้องการจัดสินค้า", vGetPickQTY)
      
      CheckNumber (vRecQTY)
      
      If vRecQTY <> "" Then
         If vCheckValueNumber = True Then
            If vRecQTY = 0 Then
               Me.ListViewSelectItemPicking.ListItems(i).Checked = False
               Me.ListViewSelectItemPicking.ListItems(i).SubItems(4) = Format(vPickQty, "##,##0.00")
               Exit Sub
            ElseIf vRecQTY > 0 And Me.ListViewSelectItemPicking.ListItems(i).Checked = False Then
               Me.ListViewSelectItemPicking.ListItems(i).Checked = True
            End If
            vPickQty = vRecQTY
            If vPickQty <= vCheckQty And (vCheckQty - vPickQty) >= 0 Then
               Me.ListViewSelectItemPicking.ListItems(i).SubItems(4) = Format(vPickQty, "##,##0.00")
            Else
               MsgBox "สั่งจัดสินค้าเกินกว่าที่สั่งขาย", vbCritical, "Send Error Message"
               Me.ListViewSelectItemPicking.ListItems(i).Checked = False
            End If
         Else
           MsgBox "ต้องกรอกข้อมูลที่เกี่ยวกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
         End If
      End If
   End If
End If
End Sub

Private Sub MEBReqTime_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.ListViewSaleOrder.SetFocus
End If
End Sub

Private Sub MEBReqTime_LostFocus()
Dim vReqTime As String
Dim vHour As String
Dim vMinute As String
Dim vHourNow As String
Dim vMinuteNow As String
Dim vNHour As Integer
Dim vNMinute As Integer
Dim vNHourNow As Integer
Dim vNMinuteNow As Integer


If InStr(Me.MEBReqTime.Text, "_") > 0 Then
    MsgBox "กรุณากรอก ข้อมูลเวลาให้ครบ เช่น 08:05,16:05  เป็นต้น", vbCritical, "Send Error Message"
    Me.MEBReqTime.SetFocus
    Exit Sub
End If


If Me.MEBReqTime.Text <> "__:__" Then
    vHour = Left(Me.MEBReqTime.Text, 2)
    vMinute = Right(Me.MEBReqTime.Text, 2)
    vNHour = vHour
    vNMinute = vMinute
    
    vHourNow = Hour(Now)
    vMinuteNow = Minute(Now)
    vNHourNow = vHourNow
    vNMinuteNow = vMinuteNow
     
    If vNHour >= 24 Then
    MsgBox "คุณกรอกชั่วโมงผิด", vbCritical, "Send Error Message"
    Me.MEBReqTime.SetFocus
    Exit Sub
    End If
    
    If vNMinute >= 60 Then
    MsgBox "คุณกรอกนาทีผิด", vbCritical, "Send Error Message"
    Me.MEBReqTime.SetFocus
    Exit Sub
    End If
    
    
    If vNHour < vNHourNow Then
        MsgBox "เวลาที่ต้องการรับของต้องมากกว่าเวลาปัจจุบันเท่านั้น", vbCritical, "Send Error Message"
        Me.MEBReqTime.SetFocus
        Exit Sub
    ElseIf vNHour = vNHourNow Then
        If vNMinute < vNMinuteNow Then
            MsgBox "เวลาที่ต้องการรับของต้องมากกว่าเวลาปัจจุบันเท่านั้น", vbCritical, "Send Error Message"
            Me.MEBReqTime.SetFocus
            Exit Sub
        End If
    End If
    
Me.ListViewSaleOrder.SetFocus
End If
End Sub

Private Sub MEDTime_LostFocus()
Dim vTime1 As String
Dim vTime2 As String
Dim vTime3 As String
Dim vTime4 As String
Dim vTime5 As String
Dim vTime6 As String

Dim vHour As Integer
Dim vMinute As Integer
Dim vCheckDateTime As Date
Dim vCheckDateDiff As Integer

If MEDTime.Text <> "" Then
    vTime1 = Left(MEDTime.Text, 1)
    vTime2 = Right(Left(MEDTime.Text, 2), 1)
    vTime3 = Right(MEDTime.Text, 1)
    vTime4 = Left(Right(MEDTime.Text, 2), 1)
    vTime5 = Left(MEDTime.Text, 2)
    vTime6 = Right(MEDTime.Text, 2)
    
    If vTime1 = "_" Or vTime2 = "_" Or vTime3 = "_" Or vTime4 = "_" Then
        MsgBox "รูปแบบเวลาที่ใช้ในโปรแกรม คือ '00:01' ต้องใส่ให้ครบด้วยครับ", vbInformation, "Send Information"
        MEDTime.SetFocus
        Exit Sub
    End If
    
    If vTime5 > "24" Then
        MsgBox "รูปแบบเวลาที่ใช้ในโปรแกรม คือ ชั่วโมงต้องใส่ไม่เกิน 24 และ นาทีต้องไม่เกิน 60 หรือ เริ่มต้นที่เวลา 00:01 ", vbInformation, "Send Information"
        If Len(Hour(Now)) = 1 Then
          If Len(Minute(Now)) = 1 Then
            MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Trim("0" & Minute(Now))
          Else
            MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Minute(Now)
          End If
        Else
          If Len(Minute(Now)) = 1 Then
            MEDTime.Text = Hour(Now) & ":" & Trim("0" & Minute(Now))
          Else
            MEDTime.Text = Hour(Now) & ":" & Minute(Now)
          End If
        End If
        MEDTime.SetFocus
        Exit Sub
    ElseIf vTime6 > "59" Then
        MsgBox "รูปแบบเวลาที่ใช้ในโปรแกรม คือ ชั่วโมงต้องใส่ไม่เกิน 24 และ นาทีต้องไม่เกิน 60 หรือ เริ่มต้นที่เวลา 00:01", vbInformation, "Send Information"
        If Len(Hour(Now)) = 1 Then
          If Len(Minute(Now)) = 1 Then
            MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Trim("0" & Minute(Now))
          Else
            MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Minute(Now)
          End If
        Else
          If Len(Minute(Now)) = 1 Then
            MEDTime.Text = Hour(Now) & ":" & Trim("0" & Minute(Now))
          Else
            MEDTime.Text = Hour(Now) & ":" & Minute(Now)
          End If
        End If
        MEDTime.SetFocus
        Exit Sub
    ElseIf vTime5 = "24" And vTime6 > "00" Then
        MsgBox "รูปแบบเวลาที่ใช้ในโปรแกรม คือ ชั่วโมงต้องใส่ไม่เกิน 24 และ นาทีต้องไม่เกิน 60 หรือ เริ่มต้นที่เวลา 00:01", vbInformation, "Send Information"
        If Len(Hour(Now)) = 1 Then
          If Len(Minute(Now)) = 1 Then
            MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Trim("0" & Minute(Now))
          Else
            MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Minute(Now)
          End If
        Else
          If Len(Minute(Now)) = 1 Then
            MEDTime.Text = Hour(Now) & ":" & Trim("0" & Minute(Now))
          Else
            MEDTime.Text = Hour(Now) & ":" & Minute(Now)
          End If
        End If
        MEDTime.SetFocus
        Exit Sub
    ElseIf vTime5 = "00" And vTime6 = "00" Then
        MsgBox "รูปแบบเวลาที่ใช้ในโปรแกรม คือ ชั่วโมงต้องใส่ไม่เกิน 24 และ นาทีต้องไม่เกิน 60 หรือ เริ่มต้นที่เวลา 00:01", vbInformation, "Send Information"
        If Len(Hour(Now)) = 1 Then
          If Len(Minute(Now)) = 1 Then
            MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Trim("0" & Minute(Now))
          Else
            MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Minute(Now)
          End If
        Else
          If Len(Minute(Now)) = 1 Then
            MEDTime.Text = Hour(Now) & ":" & Trim("0" & Minute(Now))
          Else
            MEDTime.Text = Hour(Now) & ":" & Minute(Now)
          End If
        End If
        MEDTime.SetFocus
        Exit Sub
    End If
    
    vHour = Left(Trim(Me.MEDTime.Text), 2)
    vMinute = Right(Trim(Me.MEDTime.Text), 2)
    
    vCheckDateTime = Day(Now) & "/" & Month(Now) & "/" & Year(Now) & "    " & vHour & ":" & vMinute & ":" & "00"
    
    vCheckDateDiff = DateDiff("n", Now, vCheckDateTime)
    
    If vCheckDateDiff < 15 Then
      MsgBox "ไม่สามารถกำหนดเวลาที่ลูกค้าต้องการรับสินค้าน้อยกว่า 15 นาทีได้", vbCritical, "Send Error Infromation"
      MEDTime.SetFocus
      Exit Sub
    End If

End If
End Sub

Private Sub OPTNow_Click()
Me.MEDTime.Enabled = False
End Sub

Private Sub OptSchedule_Click()
Me.MEDTime.Enabled = True
End Sub

Private Sub Pic101_GotFocus()
LBL101.Caption = ""
End Sub

Public Sub PrintPickingSlipHeader(vSaleOrder As String, vWHCode As String, vShelfCode As String, vZone As Integer)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
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
Dim vSelectPicked As Integer
Dim vGroupDocNo As String


If UCase(Left(vSaleOrder, 3)) = "S01" Or UCase(Left(vSaleOrder, 3)) = "S02" Or UCase(Left(vSaleOrder, 3)) = "W01" Or UCase(Left(vSaleOrder, 3)) = "W02" Then
vGroupDocNo = UCase(Left(Right(vSaleOrder, Len(vSaleOrder) - InStr(vSaleOrder, "-")), 3)) 'UCase(Left(vSaleOrder, 3))
Else
vGroupDocNo = UCase(Left(vSaleOrder, 3))
End If
  
If Me.OPTReserve.Value = True Then
   vSelectPicked = 0
Else
   vSelectPicked = 1
End If


If vGroupDocNo = "RWV" Or vGroupDocNo = "RWN" Then
   If vSelectPicked = 0 Then 'Res
      If vZone = 0 Then
         vPrinterName = Trim("\\nova\EPS TM-T88III-NP")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If
      
      If vZone = 1 Then 'Res
         vPrinterName = Trim("\\nova\EPS-TM-PickingOutlet")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If
   End If
   
   If vSelectPicked = 1 Then
      If vShelfCode = "AVL" Then 'Pay
         vPrinterName = Trim("\\nova\EPS-TM-PickingOutlet")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      Else
         If vZone = 0 Then
            vPrinterName = Trim("\\nova\EPS TM-T88III-NP")
            For Each printerObj In Printers
            If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
            Set Printer = printerObj
            Set printerObj = Nothing
            Exit For
            End If
            Next
         End If
         
         If vZone = 1 Then
            vPrinterName = Trim("\\nova\EPS-TM-PickingOutlet")
            For Each printerObj In Printers
            If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
            Set Printer = printerObj
            Set printerObj = Nothing
            Exit For
            End If
            Next
         End If
      End If
   End If
      
ElseIf vGroupDocNo = "ROV" Or vGroupDocNo = "RON" Then

   If vSelectPicked = 0 Then 'Res
      If vZone = 0 Then
         vPrinterName = Trim("\\nova\EPS TM-T88III-NP")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If

      If vZone = 1 Then 'Res
         vPrinterName = Trim("\\nova\EPS-TM-PickingOutlet")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If
   End If
   
Else
   If vZone = 0 Then
      vPrinterName = Trim("\\nova\EPS TM-T88III-NP")
      For Each printerObj In Printers
      If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
      Set Printer = printerObj
      Set printerObj = Nothing
      Exit For
      End If
      Next
   End If
   
   If vZone = 1 Then
      vPrinterName = Trim("\\nova\EPS-TM-PickingOutlet")
      For Each printerObj In Printers
      If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
      Set Printer = printerObj
      Set printerObj = Nothing
      Exit For
      End If
      Next
   End If
End If

    vQuery = "exec dbo.USP_SO_PickingQueueFreedom '" & vSaleOrder & "','" & vWHCode & "','" & vShelfCode & "'," & vSOCountNumber & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    
      vSoStatus = vRecordset.Fields("sostatus").Value
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")
      
If vSoStatus = 0 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 20
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 80
      Printer.Print "(จ่าย)"
ElseIf vSoStatus = 1 Then
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
    
     If vSoStatus = 1 And vSelectPicked = 0 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า เพื่อจอง")
      ElseIf vSoStatus = 1 And vSelectPicked = 1 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า เพื่อจ่าย")
      Else
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า เพื่อจ่าย")
      End If
      
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
      Printer.Font.Size = 14
      Printer.CurrentX = 0
      Printer.CurrentY = 3400
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value)
      
      If Trim(vRecordset.Fields("carlicense").Value) <> "" Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 16
        Printer.CurrentX = 1400
        Printer.CurrentY = 3300
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
      
    Printer.CurrentX = 0
    Printer.CurrentY = 4300
    Printer.Print Trim("______________________________________________________________________________________________")

    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 30
    Printer.CurrentX = 800
    Printer.CurrentY = 4500
    Printer.Print Trim("ครบ")
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 30
    Printer.CurrentX = 2800
    Printer.CurrentY = 4500
    Printer.Print Trim("ไม่ครบ")
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 20
    Printer.CurrentX = 0
    Printer.CurrentY = 4600
    Printer.Print Trim("______")
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 20
    Printer.CurrentX = 2100
    Printer.CurrentY = 4600
    Printer.Print Trim("______")

    Printer.CurrentX = 0
    Printer.CurrentY = 4700
    Printer.Print Trim("______________________________________________________________________________________________")
    End If
    vRecordset.Close

    Printer.Font.Size = 10
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("วันที่พิมพ์ :") & Now
           
    Printer.EndDoc
End Sub
Public Sub PrintPickingSlip(vSaleOrder As String, vWHCode As String, vShelfCode As String, vZone As Integer)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
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


If UCase(Left(vSaleOrder, 3)) = "S01" Or UCase(Left(vSaleOrder, 3)) = "S02" Or UCase(Left(vSaleOrder, 3)) = "W01" Or UCase(Left(vSaleOrder, 3)) = "W02" Then
vGroupDocNo = UCase(Left(Right(vSaleOrder, Len(vSaleOrder) - InStr(vSaleOrder, "-")), 3)) 'UCase(Left(vSaleOrder, 3))
Else
vGroupDocNo = UCase(Left(vSaleOrder, 3))
End If
  
If Me.OPTReserve.Value = True Then
   vSelectPicked = 0
Else
   vSelectPicked = 1
End If

   If vZone = 0 Then
    vQuery = "exec dbo.USP_NP_SearchPrinter 2"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
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
   
   If vZone = 1 Then
    vQuery = "exec dbo.USP_NP_SearchPrinter 3"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
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
   
   'If vZone = 2 Then
    'vQuery = "exec dbo.USP_NP_SearchPrinter 4"
    'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     '   vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
    'End If
    'vRecordset.Close
       
      'If vPrinterName <> "" Then
       '  For Each printerObj In Printers
        ' If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         'Set Printer = printerObj
         'Set printerObj = Nothing
         'Exit For
         'End If
         'Next
      'Else
       '  Exit Sub
      'End If
   'End If


    vQuery = "exec dbo.USP_SO_PickingQueueFreedom '" & vSaleOrder & "','" & vWHCode & "','" & vShelfCode & "'," & vSOCountNumber & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    
      vSoStatus = vRecordset.Fields("sostatus").Value
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

If vSoStatus = 0 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 20
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 80
      Printer.Print "(จ่าย)"
ElseIf vSoStatus = 1 And vSelectPicked = 0 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 20
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 80
      Printer.Print "(จอง)"
ElseIf vSoStatus = 1 And vSelectPicked = 1 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 20
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 80
      Printer.Print "(จ่าย)"
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
      
      
     If vSoStatus = 1 And vSelectPicked = 0 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 12
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ต้นฉบับใบจัดสินค้า เพื่อกำกับสินค้าจอง")
      ElseIf vSoStatus = 1 And vSelectPicked = 1 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 12
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ต้นฉบับใบจัดสินค้า เพื่อจ่ายสินค้า")
      ElseIf vSoStatus = 0 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 12
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ต้นฉบับใบจัดสินค้า เพื่อจ่ายสินค้า")
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

Private Sub PICEditOrder_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
CMDEditExit_Click
End If
End Sub

Private Sub TBDocNo_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem

Dim vDocNo As String
Dim vSumItemAmount As Double
Dim vTaxAmount As Double
Dim vNetAmount As Double
Dim vLastDisCountAmount As Double

Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vAmount As Double


If Me.TBDocNo.Text <> "" Then
   vDocNo = Me.TBDocNo.Text

  ' vQuery = "exec dbo.USP_NP_SearchSaleOrder '" & vDocNo & "'"
   vQuery = "exec dbo.USP_NP_SearchSaleOrderPickZone '" & vDocNo & "'"
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.PICOrder.Visible = True
      Me.ListViewSaleOrder.ListItems.Clear
      vSumItemAmount = Trim(vRecordset.Fields("sumofitemamount").Value)
      vTaxAmount = Trim(vRecordset.Fields("taxamount").Value)
      vNetAmount = Trim(vRecordset.Fields("netamount").Value)
      vLastDisCountAmount = Trim(vRecordset.Fields("discountamount").Value)
      
      Me.LBLOrderBillType.Caption = Trim(vRecordset.Fields("billtype").Value)
      Me.LBLOrderSoStatus.Caption = Trim(vRecordset.Fields("sostatus").Value)
      Me.LBLOrderSendQue.Caption = Trim(vRecordset.Fields("isconditionsend").Value)
      
      Me.LBLOrderDocDate.Caption = Trim(vRecordset.Fields("docdate").Value)
      Me.LBLOrderArCode.Caption = Trim(vRecordset.Fields("arcode").Value)
      Me.LBLOrderArName.Caption = Trim(vRecordset.Fields("arname").Value)
      Me.LBLOrderMember.Caption = Trim(vRecordset.Fields("memberid").Value)
      
      If Trim(vRecordset.Fields("salecode").Value) = "" Then
          MsgBox "เอกสารไม่ได้กำหนดรหัสพนักงานขาย กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      Else
            Me.LBLOrderSaleCode.Caption = Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      End If

      Me.LBLOrderSumOfItemAmount.Caption = Format(vSumItemAmount, "##,##0.00")
      Me.LBLOrderTaxAmount.Caption = Format(vTaxAmount, "##,##0.00")
      Me.LBLOrderNetAmount.Caption = Format(vNetAmount, "##,##0.00")
      If vLastDisCountAmount <> 0 Then
      Me.TBSOLastDisCount.Text = Format(vLastDisCountAmount, "##,##0.00")
      Me.LBLOrderDiscountOld.Caption = Format(vLastDisCountAmount, "##,##0.00")
      Me.TBSOLastDisCount.Enabled = True
      Else
      Me.TBSOLastDisCount.Text = Format(vLastDisCountAmount, "##,##0.00")
      Me.LBLOrderDiscountOld.Caption = Format(vLastDisCountAmount, "##,##0.00")
      Me.TBSOLastDisCount.Enabled = False
      
      End If
      
      vRecordset.MoveFirst
      For i = 1 To vRecordset.RecordCount
        Set vListItem = Me.ListViewSaleOrder.ListItems.Add(, , i)
        
        vQTY = Trim(vRecordset.Fields("remainqty").Value)
        vPrice = Trim(vRecordset.Fields("price").Value)
        vDiscountAmount = Trim(vRecordset.Fields("discountamountsub").Value)
        vAmount = Trim(vRecordset.Fields("amount").Value)

        vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
        vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
        vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
        vListItem.SubItems(5) = Format(vPrice, "##,##0.00")
        vListItem.SubItems(6) = Format(vDiscountAmount, "##,##0.00")
        vListItem.SubItems(7) = Format(vAmount, "##,##0.00")
        vListItem.SubItems(8) = Trim(vRecordset.Fields("whcode").Value)
        vListItem.SubItems(9) = Trim(vRecordset.Fields("shelfcode").Value)
        vListItem.SubItems(10) = Trim(vRecordset.Fields("zoneid").Value)
        vListItem.SubItems(11) = Trim(vRecordset.Fields("shelfid").Value)
        vListItem.SubItems(12) = Trim(vRecordset.Fields("itemcode").Value)
        vListItem.SubItems(13) = Trim(vRecordset.Fields("discountwordsub").Value)
        vListItem.SubItems(14) = Format(vQTY, "##,##0.00")
        vRecordset.MoveNext
      Next i
      
      Me.TBOrderRefNo.SetFocus
   Else
      Me.ListViewSaleOrder.ListItems.Clear
      Me.LBLOrderDocDate.Caption = ""
      Me.LBLOrderArCode.Caption = ""
      Me.LBLOrderArName.Caption = ""
      Me.LBLOrderSaleCode.Caption = ""
      Me.LBLOrderSumOfItemAmount.Caption = ""
      Me.LBLOrderTaxAmount.Caption = ""
      Me.LBLOrderNetAmount.Caption = ""
      Me.TBSOLastDisCount.Text = ""
      Me.LBLOrderDiscountOld.Caption = ""
      Me.LBLOrderMember.Caption = ""
      Me.LBLOrderBillType.Caption = ""
      Me.LBLOrderSoStatus.Caption = ""
      Me.LBLOrderSendQue.Caption = ""
      Me.MEBReqTime.Text = "__:__"
      Me.TBOrderRefNo.Text = ""
      
      'Me.TBDocNo.Text = ""
      'MsgBox "ไม่มีรายการสินค้าให้จัด กรุณาตรวจสอบ", vbCritical, "Send Error Message "
      'Me.PICOrder.Visible = False
      'Me.ListViewPicking.SetFocus
       'Me.TBDocNo.SetFocus
   End If
   vRecordset.Close
   Call CalcEditItemQty
Else
      Me.ListViewSaleOrder.ListItems.Clear
      Me.LBLOrderDocDate.Caption = ""
      Me.LBLOrderArCode.Caption = ""
      Me.LBLOrderArName.Caption = ""
      Me.LBLOrderSaleCode.Caption = ""
      Me.LBLOrderSumOfItemAmount.Caption = ""
      Me.LBLOrderTaxAmount.Caption = ""
      Me.LBLOrderNetAmount.Caption = ""
      Me.TBSOLastDisCount.Text = ""
      Me.LBLOrderDiscountOld.Caption = ""
      Me.MEBReqTime.Text = "__:__"
      Me.TBOrderRefNo.Text = ""
      Me.LBLOrderMember.Caption = ""
      Me.LBLOrderBillType.Caption = ""
      Me.LBLOrderSoStatus.Caption = ""
      Me.LBLOrderSendQue.Caption = ""
End If
End Sub

Private Sub TBDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
Call CMDSaleOrderSendQue_Click
End If
End Sub

Private Sub TBEditQty_Change()
Dim vQtyWord As String
Dim vLenQTY As Integer

If Me.TBEditQty.Text <> "" Then
   vQtyWord = Me.TBEditQty.Text
   CheckNumber (vQtyWord)
      
   If vCheckValueNumber = False Then
      MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      vLenQTY = Len(Me.TBEditQty.Text)
      Me.TBEditQty.Text = Left(Me.TBEditQty.Text, vLenQTY - 1)
      Me.TBEditQty.SetFocus
   End If
End If
End Sub

Private Sub TBEditQty_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   Call CMDEditOK_Click
End If

If KeyAscii = 27 Then
CMDEditExit_Click
End If
End Sub

Private Sub TBOrderRefNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.MEBReqTime.SetFocus
End If
End Sub

Private Sub TBSOLastDisCount_Change()
Dim vQtyWord As String
Dim vIsNumber As Boolean
Dim vLenQTY As Integer

Dim vQTY As Double
Dim vPrice As Double
Dim vDiscountAmount As Double
Dim vNetprice As Double


If Me.TBSOLastDisCount.Text <> "" Then
   vQtyWord = Me.TBSOLastDisCount.Text
   CheckNumber (vQtyWord)
   
   If vCheckValueNumber = False Then
      MsgBox "กรอกได้เฉพาะอักขระที่เกี่ยวข้องกับตัวเลขเท่านั้น", vbCritical, "Send Error Message"
      vLenQTY = Len(Me.TBSOLastDisCount.Text)
      Me.TBSOLastDisCount.Text = Left(Me.TBSOLastDisCount.Text, vLenQTY - 1)
      Me.TBSOLastDisCount.SetFocus
      Exit Sub
   End If
     
End If
Call CalcEditItemQty
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

