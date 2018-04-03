VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form03 
   Caption         =   "พิมพ์ใบหยิบสินค้าทดแทน"
   ClientHeight    =   11010
   ClientLeft      =   2655
   ClientTop       =   795
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form03.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PTBPickingQueue 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   11445
      Left            =   0
      ScaleHeight     =   11415
      ScaleWidth      =   15330
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   15360
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   2520
         ScaleHeight     =   435
         ScaleWidth      =   8085
         TabIndex        =   60
         Top             =   900
         Width           =   8115
         Begin VB.OptionButton OPTPayItem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "สั่งจัดเพื่อ จ่ายสินค้า"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3285
            TabIndex        =   62
            Top             =   90
            Width           =   1725
         End
         Begin VB.OptionButton OPTReserve 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "สั่งจัดเพื่อ จองสินค้า"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            TabIndex        =   61
            Top             =   90
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
         Left            =   2520
         TabIndex        =   58
         Top             =   7425
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton OPTReserve1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "สั่งจัดเพื่อ จองสินค้า"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2520
         TabIndex        =   57
         Top             =   7785
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton CMDSelectItemBack 
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
         Height          =   510
         Left            =   7560
         TabIndex        =   56
         Top             =   7020
         Width           =   1500
      End
      Begin VB.TextBox TextCarLicense 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5985
         TabIndex        =   32
         Top             =   6660
         Width           =   1500
      End
      Begin VB.CheckBox CHKLicense 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "กรณี กำหนดทะเบียนรถขนส่ง     เลขที่ :"
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
         Left            =   2520
         TabIndex        =   31
         Top             =   6615
         Width           =   3345
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
         Height          =   510
         Left            =   9135
         TabIndex        =   25
         Top             =   7020
         Width           =   1500
      End
      Begin VB.CommandButton CMDQueue 
         Caption         =   "ส่งคิว"
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
         Left            =   5985
         TabIndex        =   24
         Top             =   7020
         Width           =   1500
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   2520
         ScaleHeight     =   1515
         ScaleWidth      =   8085
         TabIndex        =   21
         Top             =   4995
         Width           =   8115
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
            Height          =   240
            Left            =   90
            TabIndex        =   26
            Top             =   900
            Visible         =   0   'False
            Width           =   2040
         End
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
            Height          =   240
            Left            =   90
            TabIndex        =   23
            Top             =   495
            Visible         =   0   'False
            Width           =   2040
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
            Left            =   90
            TabIndex        =   22
            Top             =   90
            Value           =   -1  'True
            Width           =   1770
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "สินค้า SPO ออกโกดังฝั่งสำนักงานใหญ่"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   1410
            Left            =   2205
            TabIndex        =   63
            Top             =   45
            Width           =   5820
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1185
         Left            =   2520
         ScaleHeight     =   1155
         ScaleWidth      =   8085
         TabIndex        =   13
         Top             =   3690
         Width           =   8115
         Begin VB.OptionButton OptTomorrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "เช้าวันพรุ่งนี้"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            TabIndex        =   33
            Top             =   810
            Width           =   2220
         End
         Begin MSMask.MaskEdBox MEDTime 
            Height          =   285
            Left            =   3285
            TabIndex        =   27
            Top             =   450
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.OptionButton OptSchedule 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ตามเวลาที่กำหนด"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            TabIndex        =   15
            Top             =   450
            Width           =   2040
         End
         Begin VB.OptionButton OptNow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ด่วน"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   90
            Value           =   -1  'True
            Width           =   1860
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "เวลา :"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2385
            TabIndex        =   28
            Top             =   495
            Width           =   870
         End
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         Height          =   285
         Left            =   135
         TabIndex        =   59
         Top             =   990
         Width           =   2310
      End
      Begin VB.Label Label18 
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
         Height          =   285
         Left            =   8010
         TabIndex        =   55
         Top             =   3150
         Width           =   1005
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   7560
         TabIndex        =   54
         Top             =   1710
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   7605
         TabIndex        =   53
         Top             =   2700
         Width           =   1410
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   7920
         TabIndex        =   52
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label LBLSaleCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2520
         TabIndex        =   51
         Top             =   3195
         Width           =   1815
      End
      Begin VB.Label LBLIsConditionSend 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9090
         TabIndex        =   50
         Top             =   3150
         Width           =   1545
      End
      Begin VB.Label LBLDueDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9090
         TabIndex        =   49
         Top             =   1710
         Width           =   1545
      End
      Begin VB.Label LBLSaleType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9090
         TabIndex        =   48
         Top             =   2700
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label LBLBillType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9090
         TabIndex        =   47
         Top             =   2205
         Width           =   1545
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   0
         Picture         =   "Form03.frx":9673
         Top             =   0
         Width           =   2160
      End
      Begin VB.Label LBLSumQTY 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   5580
         TabIndex        =   30
         Top             =   6975
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label LBLCountItem 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   4590
         TabIndex        =   29
         Top             =   6975
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label LBLSaleName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4455
         TabIndex        =   20
         Top             =   3195
         Width           =   3030
      End
      Begin VB.Label LBLARAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2520
         TabIndex        =   19
         Top             =   2700
         Width           =   4965
      End
      Begin VB.Label LBLARName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2520
         TabIndex        =   18
         Top             =   2205
         Width           =   4965
      End
      Begin VB.Label LBLDocDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5850
         TabIndex        =   17
         Top             =   1710
         Width           =   1635
      End
      Begin VB.Label LBLDocNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Top             =   1710
         Width           =   1815
      End
      Begin VB.Label Label9 
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
         Height          =   285
         Left            =   540
         TabIndex        =   12
         Top             =   4995
         Width           =   1905
      End
      Begin VB.Label Label8 
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
         Height          =   330
         Left            =   495
         TabIndex        =   11
         Top             =   3690
         Width           =   1950
      End
      Begin VB.Label Label7 
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
         Height          =   285
         Left            =   945
         TabIndex        =   10
         Top             =   3195
         Width           =   1500
      End
      Begin VB.Label Label6 
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
         Height          =   285
         Left            =   1035
         TabIndex        =   9
         Top             =   2700
         Width           =   1410
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อ-รหัสลูกค้า :"
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
         TabIndex        =   8
         Top             =   2205
         Width           =   1635
      End
      Begin VB.Label Label4 
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
         Height          =   285
         Left            =   4590
         TabIndex        =   7
         Top             =   1710
         Width           =   1185
      End
      Begin VB.Label Label3 
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
         Height          =   240
         Left            =   1215
         TabIndex        =   6
         Top             =   1710
         Width           =   1230
      End
   End
   Begin VB.PictureBox PICSelectPrintSlip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   11445
      Left            =   0
      Picture         =   "Form03.frx":AAD5
      ScaleHeight     =   11415
      ScaleWidth      =   15330
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   15360
      Begin VB.PictureBox PicPoint 
         Height          =   240
         Left            =   0
         ScaleHeight     =   180
         ScaleWidth      =   405
         TabIndex        =   64
         Top             =   0
         Width           =   465
      End
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
         Height          =   510
         Left            =   1620
         TabIndex        =   45
         Top             =   7515
         Width           =   1320
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
         Height          =   510
         Left            =   180
         TabIndex        =   44
         Top             =   7515
         Width           =   1230
      End
      Begin VB.CheckBox CHKSelectAllItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "เลือกทั้งหมดตามจำนวนในใบสั่งขาย"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   43
         Top             =   2565
         Width           =   3390
      End
      Begin MSComctlLib.ListView ListViewSelectItemPicking 
         Height          =   4425
         Left            =   180
         TabIndex        =   42
         Top             =   2880
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   8555
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
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "โซน"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "GroupCode"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "PickZone"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label LBLRefDocDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4590
         TabIndex        =   41
         Top             =   1305
         Width           =   1995
      End
      Begin VB.Label Label16 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3465
         TabIndex        =   40
         Top             =   1350
         Width           =   1230
      End
      Begin VB.Label LBLRefARName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4590
         TabIndex        =   39
         Top             =   1800
         Width           =   7125
      End
      Begin VB.Label LBLRefARCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1350
         TabIndex        =   38
         Top             =   1800
         Width           =   1860
      End
      Begin VB.Label LBLRefDocNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1350
         TabIndex        =   37
         Top             =   1305
         Width           =   1860
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   405
         TabIndex        =   36
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label Label10 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   35
         Top             =   1350
         Width           =   1275
      End
   End
   Begin VB.CommandButton CMDReqPicking 
      Caption         =   "เลือกสินค้า"
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
      Left            =   6660
      TabIndex        =   46
      Top             =   2340
      Width           =   1095
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   495
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
   Begin VB.CommandButton CMD101 
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
      Height          =   510
      Left            =   3240
      TabIndex        =   4
      Top             =   3555
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   2475
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4230
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.TextBox Text101 
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
      Height          =   420
      Left            =   4635
      TabIndex        =   2
      Top             =   1665
      Width           =   3120
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "โซนการออกใบหยิบ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   330
      Left            =   765
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000004&
      Height          =   285
      Left            =   3150
      TabIndex        =   0
      Top             =   1755
      Width           =   1365
   End
End
Attribute VB_Name = "Form03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSaleCode As String
Dim vARCode As String
Dim vQueueID  As Integer
Dim vCustomerZone As Integer

Dim vUserPrint As String
Dim vCheckValue As Boolean
Dim vCheckValue1 As Boolean
Dim vKeyword As String
Dim vCheckKeyword As String
Dim vCheckPic101 As Integer
Dim vCheckSelectItemPickBack As Integer
Dim vSOCountNumber As Integer

Private Sub CHKLicense_Click()
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

Private Sub CMD101_Click()
'Dim vRecordset As New ADODB.Recordset
'Dim vQuery As String
'Dim vZone As String
'Dim vQueueNo As String
'Dim vQueueStatus As Integer
'Dim vQueueDate As Date
'Dim vDocNo As String
'Dim vCheckAnswer As Integer

'If Me.Text101.Text <> "" And Me.CMB101.Text <> "" Then
 'vDocNo = Trim(Me.Text101.Text)
 'vQuery = "select top 1  docno,queuedatetime,status from npmaster.dbo.TB_NP_QueueManagement where saleorderno = '" & vDocNo & "' order by queuedatetime desc"
 'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  ' vQueueNo = vRecordset.Fields("docno").Value
   'vQueueDate = vRecordset.Fields("queuedatetime").Value
   'vQueueStatus = vRecordset.Fields("status").Value
 'End If
 'vRecordset.Close
 
 'If vQueueNo <> "" Then
  '  If vQueueStatus = 0 Then
   '   MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ รอจัดสินค้า ", vbInformation, "Send Information"
    'ElseIf vQueueStatus = 1 Then
     ' MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ กำลังจัดสินค้า ", vbInformation, "Send Information"
    'ElseIf vQueueStatus = 2 Then
     ' MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ จัดสินค้าเรียบร้อยแล้ว ", vbInformation, "Send Information"
    'End If
    
    'vCheckAnswer = MsgBox("ต้องการพิมพ์ทดแทนหรือไม่", vbYesNo, "Send Question ?")
    'If vCheckAnswer = 7 Then
     ' Exit Sub
    'End If
'End If
 
 ' vQuery = "exec dbo.USP_SO_SaleOrderDetails '" & vDocNo & "' "
  'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   ' Me.LBLDocNo.Caption = Trim(vRecordset.Fields("docno").Value)
    'Me.LBLDocDate.Caption = Trim(vRecordset.Fields("docdate").Value)
    'Me.LBLARName.Caption = Trim(vRecordset.Fields("arname").Value)
    'Me.LBLARAddress.Caption = Trim(vRecordset.Fields("workaddress").Value)
    'Me.LBLSaleName.Caption = Trim(vRecordset.Fields("salename").Value)
    'Me.LBLCountItem.Caption = vRecordset.Fields("CountItem").Value
    'Me.LBLSumQTY.Caption = vRecordset.Fields("SumRemainQTY").Value
  
    'If Len(Hour(Now)) = 1 Then
     ' If Len(Minute(Now)) = 1 Then
      '  MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Trim("0" & Minute(Now))
      'Else
       ' MEDTime.Text = Trim("0" & Hour(Now)) & ":" & Minute(Now)
      'End If
    'Else
     ' If Len(Minute(Now)) = 1 Then
      '  MEDTime.Text = Hour(Now) & ":" & Trim("0" & Minute(Now))
      'Else
       ' MEDTime.Text = Hour(Now) & ":" & Minute(Now)
      'End If
    'End If
    'PicSetTimeZone.Visible = True
    'Me.OptNow.Value = True
    'Me.OptNormal.Value = True
    'Me.CHKLicense.Value = 0
  'Else
   ' MsgBox "กรุณาตรวจสอบเลขที่ใบสั่งขาย/จอง ที่จะเข้าคิวจัดสินค้า เพราะเอกสารดังกล่าวไม่สมารถทำคิวรอจัดสินค้าได้ ", vbCritical, "Send Error Message"
  'End If
  'vRecordset.Close
  
'vZone = UCase(Trim(CMB101.Text))
'Select Case vZone
'Case UCase("A"):
 '     Call PrintPicking_A
'Case UCase("B"):
 '     Call PrintPicking_B
'Case UCase("M"):
 '     Call PrintPicking_M
'Case UCase("H"):
 '     Call PrintPicking_H
'Case UCase("D"):
 '     Call PrintPicking_D
'Case UCase("Y"):
 '     Call PrintPicking_Y
'End Select
'Text101.SetFocus
'CMB101.Clear
'End If
End Sub

Public Function CheckDocument()
Dim vRecordset  As New ADODB.Recordset
Dim vQuery As String
Dim vDocGroup As String, vDocument As String, vGroupDoc1 As String
Dim vTable As String, vCheckDoc As String, vTypeDoc As String, vDocGroup1 As String
Dim vMemDocNo As String

On Error GoTo ErrDescription

vMemDocNo = Text101.Text
vDocGroup1 = UCase(Left(Right(vMemDocNo, Len(vMemDocNo) - InStr(vMemDocNo, "-")), 3)) 'Left(Trim(Text101.Text), 3)
vDocument = Trim(UCase(Text101.Text))
vDocGroup = UCase(vDocGroup1)

'--------------------------------------------------------------------------------------------

If vDocGroup = "SHV" Or vDocGroup = "SHN" Or vDocGroup = "SCV" Or vDocGroup = "SCN" Or vDocGroup = "SVD" Or _
    vDocGroup = "SVN" Or vDocGroup = "SVM" Or vDocGroup = "SAB" Or vDocGroup = "ROV" Or vDocGroup = "RON" Then
    vTable = "BCNP.DBO.BCSALEORDER"
    vTypeDoc = "SO"
End If


If vTable <> "" Then
vQuery = "select docno from " & vTable & " where docno = '" & vDocument & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDoc = vRecordset.Fields("Docno").Value
End If
vRecordset.Close
'-------------------------------------------------------------------------------------------------
End If

vGroupDoc1 = UCase(Left(Right(vDocGroup, Len(vDocGroup) - InStr(vDocGroup, "-")), 2)) 'Left(vDocGroup, 2)


If vCheckDoc = "" Or IsNull(vCheckDoc) Then
    MsgBox "ไม่มีเอกสารนี้ในระบบ", vbCritical, "ข้อความเตือน"
Else
'--------------------------------------------------------------------------
    If vTypeDoc = "SO" Then
     vQuery = "INSERT INTO NPMaster.dbo.NPPrintLogs(DocNo,DocTypeID,GroupDOC,Code,Printed,LastPrintedUser,LastPrintDateTime)" & _
                     "select docno,'" & vTypeDoc & "' as DocTypeID,'" & vDocGroup & "' as GroupDoc,'' as arcode,1 as Printed,'" & vUserID & "' as UserID ,getdate() From " & vTable & " where   docno = '" & vDocument & "'  "
    Else
    vQuery = "INSERT INTO NPMaster.dbo.NPPrintLogs(DocNo,DocTypeID,GroupDOC,Code,Printed,LastPrintedUser,LastPrintDateTime)" & _
                     "select docno,'" & vTypeDoc & "' as DocTypeID,'" & vDocGroup & "' as GroupDoc,'' as arcode,1 as Printed,'" & vUserID & "' as UserID ,getdate() From " & vTable & " where   docno = '" & vDocument & "'  "
    End If
    gConnection.Execute vQuery
    MsgBox "กรุณากดปุ่ม Enter อีกครั้งนะครับ", vbCritical, "ข้อความแจ้งให้ทราบ"
    '-------------------------------------------------------------------------------------------
End If

vQuery = "Delete npmaster.dbo.npprintserver where docno = '" & vDocument & "' "
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If

End Function


Private Sub CMDExit_Click()
Me.PTBPickingQueue.Visible = False
Me.OptNow.Value = True
Me.OptNormal.Value = True
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
   vQuery = "exec dbo.USP_SO_SaleOrderDetails '" & vDocNo & "' " 'pass
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
   
   
     If vRecordset.Fields("sostatus").Value = 1 Then
        Me.OPTReserve.Value = True
        Me.OPTPayItem.Value = False
     Else
       Me.OPTReserve.Value = False
       Me.OPTPayItem.Value = True
     End If
     
     
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
      
      vQuery = "exec dbo.USP_NP_SearchCheckPickStatus '" & vDocNo & "' " 'pass
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

Private Sub CMDQueue_Click()
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vQuery As String
Dim vQuery1 As String

Dim vDocNo As String
Dim vWHGroup(12) As String
Dim vShelfGroup(12) As String
Dim vZoneGroup(12) As String
Dim vFamilyGroup(12) As String
Dim vPickZone(12) As String
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
Dim vCarLicense As String
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
Dim vCheckFamilyGroup As String
Dim vCheckWHGroup As String
Dim vCheckZoneID As String
Dim vItemFamily As String
Dim vCountItem As Integer
Dim vCountSelectPick As Integer
Dim vItemPickZone As String
Dim vCheckPickZone As String

vDocNo = Me.LBLDocNo.Caption
vDocdate = Me.LBLDocDate.Caption
vPickingDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
vARCode = Left(Me.LBLARName.Caption, InStr(Me.LBLARName.Caption, "//") - 1)
vSaleCode = Me.LBLSaleCode.Caption

vCheckSPO = 0
For m = 1 To Me.ListViewSelectItemPicking.ListItems.Count
   If Me.ListViewSelectItemPicking.ListItems(m).Checked = True Then
        If Me.ListViewSelectItemPicking.ListItems(m).SubItems(7) = "SPO" Then
           vCheckSPO = vCheckSPO + 1
        End If
   End If
Next m

'If vCheckSPO > 0 Then
 '  If Me.OptMain.Value = False And Me.OptOutLet.Value = False Then
  '    MsgBox "กรณีที่มีการสั่งจัดสินค้าชั้นเก็บ SPO ต้องระบุด้วยว่าลูกค้ารับของฝั่งไหนตามที่อยู่สินค้าที่อยู่จริง เพื่อความสะดวกต่อการจัดสินค้า กรุณาระบุด้วย", vbCritical, "Send Error Message"
   '   Exit Sub
   'End If
'End If

If vSaleCode = "" Then
   MsgBox "ไม่ได้ระบุ รหัสพนักงานกรุณาตรวจสอบ ", vbCritical, "Send Error Message"
   Exit Sub
End If

vIsConditionSend = Me.LBLIsConditionSend.Caption
vCarLicense = Me.TextCarLicense.Text
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

'vQuery = "exec dbo.USP_NP_SearchCheckCountSOPicking '" & vDocNo & "'"
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '  vSOCountNumber = vRecordset.Fields("vCount").Value
'End If
'vRecordset.Close

'vQuery = "exec dbo.USP_NP_CheckSaleOrderPickupZone'" & vDocNo & "' "
vQuery = "exec dbo.USP_NP_CheckSaleOrderPickupZoneUnitCode '" & vDocNo & "' " 'แยกหน่วยนับท่อ
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   vCheckWHGroup = vRecordset.Fields("whcode").Value
   vCheckShelfGroup = vRecordset.Fields("shelfgroup").Value
   vCheckZoneID = vRecordset.Fields("zoneid").Value
   vCheckFamilyGroup = vRecordset.Fields("familygroup").Value
   vCheckPickZone = vRecordset.Fields("pickzone").Value
   
   
  'pass
  vQuery1 = "exec dbo.USP_NP_SearchCheckCountSOPickingByZone_NewWH '" & vDocNo & "','" & vCheckWHGroup & "','" & vCheckShelfGroup & "','" & vCheckZoneID & "','" & vCheckPickZone & "','" & vCheckFamilyGroup & "' "
  If OpenDataBase(gConnection, vRecordset1, vQuery1) <> 0 Then
     vSOCountNumber = vRecordset1.Fields("vCount").Value
  End If
  vRecordset1.Close
      
      
   
   vQuery = "exec dbo.USP_NP_InsertSelectItemPickingMaster3_NewWH  '" & vDocNo & "','" & vDocdate & "','" & vARCode & "','" & vPickingDate & "'," & vBillType & "," & vSoStatus & ",0,'" & vSaleCode & "','" & vCarLicense & "'," & vIsConditionSend & "," & vSOCountNumber & ",'" & vCheckWHGroup & "','" & vCheckShelfGroup & "','" & vCheckZoneID & "','" & vCheckFamilyGroup & "','" & vCheckPickZone & "','" & vDueDate & "'," & vPickStatus & ",'" & vUserID & "' "
   gConnection.Execute vQuery
   vRecordset.MoveNext
  Wend
End If
vRecordset.Close


For j = 1 To Me.ListViewSelectItemPicking.ListItems.Count
   If Me.ListViewSelectItemPicking.ListItems(j).Checked = True Then
      vItemCode = Me.ListViewSelectItemPicking.ListItems(j).SubItems(1)
      vItemName = Me.ListViewSelectItemPicking.ListItems(j).SubItems(2)
      vReqQTY = Me.ListViewSelectItemPicking.ListItems(j).SubItems(4)
      vUnitCode = Me.ListViewSelectItemPicking.ListItems(j).SubItems(5)
      vWHCode = Me.ListViewSelectItemPicking.ListItems(j).SubItems(6)
      vShelfCode = Me.ListViewSelectItemPicking.ListItems(j).SubItems(7)
      vZoneID = Me.ListViewSelectItemPicking.ListItems(j).SubItems(8)
      vItemFamily = Me.ListViewSelectItemPicking.ListItems(j).SubItems(9)
      vItemPickZone = Me.ListViewSelectItemPicking.ListItems(j).SubItems(10)
      vIsCancel = 0
      vSelectItemDateTime = Now
      vLineNumber = j - 1
      vQuery = "exec dbo.USP_NP_InsertSelectItemPicking2 '" & vDocNo & "','" & vDocdate & "','" & vPickingDate & "','" & vItemCode & "','" & vItemName & "'," & vReqQTY & ",'" & vUnitCode & "','" & vWHCode & "','" & vShelfCode & "','" & vZoneID & "','" & vItemFamily & "','" & vItemPickZone & "'," & vIsCancel & ",'" & vSelectItemDateTime & "'," & vSOCountNumber & "," & vLineNumber & " "
      gConnection.Execute vQuery
      vCountSelectPick = vCountSelectPick + 1
   End If
Next j

vCountItem = Me.ListViewSelectItemPicking.ListItems.Count

If vCountItem <> vCountSelectPick Then
vQuery = "exec dbo.USP_NP_DeleteQueRequestMaster  '" & vDocNo & "','" & vPickingDate & "'," & vSOCountNumber & " "
gConnection.Execute vQuery
End If

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
   vQuery = "exec dbo.USP_SO_SOStatus '" & vDocNo & "' " 'pass
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vBillStatus = Trim(vRecordset.Fields("billstatus").Value)
       vBillType = Trim(vRecordset.Fields("billtype").Value)
       vSend = Trim(vRecordset.Fields("isconditionsend").Value)
       vSaleCode = Trim(vRecordset.Fields("salecode").Value)
       vRemainQtyCheckPrint = vRecordset.Fields("qty").Value
   End If
   vRecordset.Close
         
   If vRemainQtyCheckPrint > 0 Then
     vQuery = "exec dbo.USP_SO_SearchShelfGroupPicking2 '" & vDocNo & "'," & vSOCountNumber & " " 'pass
     If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       n = vRecordset.RecordCount
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
       vWHGroup(i) = Trim(vRecordset.Fields("whcode").Value)
       vShelfGroup(i) = Trim(vRecordset.Fields("shelfgroup").Value)
       vZoneGroup(i) = Trim(vRecordset.Fields("zoneid").Value)
       vFamilyGroup(i) = Trim(vRecordset.Fields("familygroup").Value)
       vPickZone(i) = Trim(vRecordset.Fields("pickzone").Value)
       vRecordset.MoveNext
       Next i
     End If
     vRecordset.Close
   
     If Me.CHKLicense.Value = 1 Then
       vCarLicense = Me.TextCarLicense.Text
       vQuery = "exec dbo.USP_NP_UpdateCarLicense '" & vDocNo & "'," & vSOCountNumber & ",'" & vCarLicense & "'"
       gConnection.Execute (vQuery)
     End If
   
        For i = 1 To n
            Call GenQueuePicking(vSOCountNumber, vZoneGroup(i), vWHGroup(i), vShelfGroup(i), vFamilyGroup(i), vPickZone(i))
        Next i

     
     Me.Text101.Text = ""
     Me.PTBPickingQueue.Visible = False
     Me.PICSelectPrintSlip.Visible = False
 End If
End If

End Sub

Public Sub GenQueuePicking(vCount As Integer, vZoneGroup As String, vWHCode As String, vShelfGroup As String, vFamilyGroup As String, vPickZoneGroup As String)

Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vNamePrint As String
Dim vDocNo As String
Dim vPicking As String
Dim vHeader As String
Dim vDocuments As String
Dim vDocNumber As String
Dim vDocdate As Date
Dim vDocType As Integer
Dim vPosition As Integer
Dim vAddTime As Date
Dim vRequestTime As String
Dim vTimeID As Integer
Dim vSoStatus As Integer
Dim vTimePick As Integer
Dim vJobID As Integer


On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Me.LBLDocNo.Caption))
If vDocNo <> "" Then
    If vUserPrint <> "" Then
       vNamePrint = Trim(vUserPrint)
    Else
       vNamePrint = Me.LBLSaleCode.Caption
    End If

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
        
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement_New2 '" & vQueueID & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vDocuments & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneGroup & "','" & vFamilyGroup & "','" & vPickZoneGroup & "'," & vTimeID & ",0,'" & vRequestTime & "'," & vCustomerZone & "," & vTimePick & " "
    gConnection.Execute vQuery

    vQuery = "exec dbo.USP_NP_UpdateRequestPickingQueue2_NewWH  '" & vDocNo & "'," & vSOCountNumber & ",'" & vQueueID & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneGroup & "','" & vFamilyGroup & "','" & vPickZoneGroup & "' "
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo  27 "
    gConnection.Execute vQuery
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, vZoneGroup, vFamilyGroup, vPickZoneGroup, vSOCountNumber)
  
  vJobID = 1

  
  vQuery = "exec dbo.USP_NP_InsertPrintTermal " & vJobID & ",'" & vDocNo & "','" & vQueueID & "','" & vWHCode & "','" & vShelfGroup & "','" & vFamilyGroup & "','" & vZoneGroup & "','" & vPickZoneGroup & "','" & vUserID & "' "
  gConnection.Execute vQuery
  

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

Public Sub PrintPicking_01(vCount As Integer)
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
    vShelfGroup = Trim(UCase("01"))
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
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
  
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าออกจุดจ่าย โกดัง ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub


Public Sub PrintPicking_02(vCount As Integer)
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
    vShelfGroup = Trim(UCase("02"))
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
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าออกจุดจ่าย OutLet ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_03(vCount As Integer)
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
    vShelfGroup = Trim(UCase("03"))
    vZoneID = Trim("03")
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
    

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 3)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 2)
  
  MsgBox "ได้คิวเลขที่ " & vQueueID & " ", vbInformation, "Send Queue Information"
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถเข้าคิวใบหยิบสินค้าออกจุดจ่าย HMX ได้ กรุณาเข้าคิวใหม่ที่หน้าทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 And vPosition = 2 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
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
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 2)
  
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
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
  
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
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  
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
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
  
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
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  
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
    

  'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
  
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
    

If DatePart("w", Now) <> 1 Then
    If Me.OptMain.Value = True Then
       'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
    ElseIf Me.OptOutLet.Value = True Then
       'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
    End If
Else
    'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
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

Public Sub PrintPickingSlip(vSaleOrder As String, vWHCode As String, vShelfGroup As String, vZoneID As String, vFamilyGroup As String, vPickZoneGroup As String, vCount As Integer)
Dim vRecordset1 As New Recordset
Dim vQuery As String
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
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
Dim vPrinterID As Integer


If UCase(Left(vSaleOrder, 3)) = "S01" Or UCase(Left(vSaleOrder, 3)) = "S02" Or UCase(Left(vSaleOrder, 3)) = "W01" Or UCase(Left(vSaleOrder, 3)) = "W02" Then
vGroupDocNo = UCase(Left(Right(vSaleOrder, Len(vSaleOrder) - InStr(vSaleOrder, "-")), 3))
Else
vGroupDocNo = UCase(Left(vSaleOrder, 3))
End If

  
If Me.OPTReserve.Value = True Then
   vSelectPicked = 0
Else
   vSelectPicked = 1
End If


vQuery = "exec dbo.USP_NP_SearchPrinterPrintZone '" & vPickZoneGroup & "' "
If OpenDataBase(gConnection, vRecordset1, vQuery) <> 0 Then
vPrinterName = Trim(vRecordset1.Fields("printername").Value)
End If
vRecordset1.Close

For Each printerObj In Printers
If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
Set Printer = printerObj
Set printerObj = Nothing
Exit For
End If
Next


    vQuery = "exec dbo.USP_SO_PickingQueueFreedom2 '" & vSaleOrder & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "','" & vFamilyGroup & "','" & vPickZoneGroup & "'," & vCount & " "
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
        Printer.Print Trim("ทดแทนต้นฉบับใบจัดสินค้า เพื่อกำกับสินค้าจอง")
      ElseIf vSoStatus = 1 And vSelectPicked = 1 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 12
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ทดแทนต้นฉบับใบจัดสินค้า เพื่อจ่ายสินค้า")
      ElseIf vSoStatus = 0 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 12
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ทดแทนต้นฉบับใบจัดสินค้า เพื่อจ่ายสินค้า")
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
          Printer.Font.Size = 18
          Printer.FontBold = True
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfcode1").Value)
          
          Printer.Font.Size = 11
          Printer.FontBold = False
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print Trim("OnHand") & "(" & Trim(vRecordset.Fields("shelfcode").Value) & ")" & ": " & Trim(vRecordset.Fields("qtylocation").Value) & "  " & Trim(vRecordset.Fields("stkunitcode").Value) & "     " & "รวมคลัง :" & Trim(vRecordset.Fields("StkWHCode").Value) & "  " & Trim(vRecordset.Fields("stkunitcode").Value)
                                        
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


vGroupDocNo = UCase(Left(Right(vSaleOrder, Len(vSaleOrder) - InStr(vSaleOrder, "-")), 3)) 'UCase(Left(vSaleOrder, 3))
  
If Me.OPTReserve.Value = True Then
   vSelectPicked = 0
Else
   vSelectPicked = 1
End If


If vGroupDocNo = "RWV" Or vGroupDocNo = "RWN" Then
   If vSelectPicked = 0 Then 'Res
      If vZone = 0 Then
         vPrinterName = Trim("\\galaxy\EPS TM-T88III-NP")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If

      If vZone = 1 Then 'Res
         vPrinterName = Trim("\\galaxy\EPS-TM-PickingOutlet")
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
         vPrinterName = Trim("\\galaxy\EPS-TM-PickingOutlet")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      Else
         If vZone = 0 Then
            vPrinterName = Trim("\\galaxy\EPS TM-T88III-NP")
            For Each printerObj In Printers
            If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
            Set Printer = printerObj
            Set printerObj = Nothing
            Exit For
            End If
            Next
         End If
         
         If vZone = 1 Then
            vPrinterName = Trim("\\galaxy\EPS-TM-PickingOutlet")
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
         vPrinterName = Trim("\\galaxy\EPS TM-T88III-NP")
         For Each printerObj In Printers
         If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
         Set Printer = printerObj
         Set printerObj = Nothing
         Exit For
         End If
         Next
      End If

      If vZone = 1 Then 'Res
         vPrinterName = Trim("\\galaxy\EPS-TM-PickingOutlet")
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
      vPrinterName = Trim("\\galaxy\EPS TM-T88III-NP")
      For Each printerObj In Printers
      If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
      Set Printer = printerObj
      Set printerObj = Nothing
      Exit For
      End If
      Next
   End If
   
   If vZone = 1 Then
      vPrinterName = Trim("\\galaxy\EPS-TM-PickingOutlet")
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
      Printer.Print Trim("ทดแทนต้นฉบับใบจัดสินค้า เพื่อจอง")
      ElseIf vSoStatus = 1 And vSelectPicked = 1 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ทดแทนต้นฉบับใบจัดสินค้า เพื่อจ่าย")
      Else
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ทดแทนต้นฉบับใบจัดสินค้า เพื่อจ่าย")
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

Private Sub CMDReqPicking_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vQueueNo As String
Dim vQueueStatus As Integer
Dim vQueueDate As Date
Dim vCheckAnswerPrint As Integer
Dim vListItemPicking As ListItem
Dim i As Integer
Dim vBillType As Integer
Dim vBillStatus As Integer
Dim vSend As Integer
Dim x As Integer
Dim vMemRemainQTY As Double
Dim vHoldingStatus As Integer

On Error GoTo ErrDescription

If Me.Text101.Text <> "" Then
   vDocNo = Trim(Me.Text101.Text)
   vQuery = "exec dbo.USP_SO_SOStatus '" & vDocNo & "' " 'pass
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vBillStatus = Trim(vRecordset.Fields("billstatus").Value)
      vBillType = Trim(vRecordset.Fields("billtype").Value)
      vSend = Trim(vRecordset.Fields("isconditionsend").Value)
      vMemRemainQTY = Trim(vRecordset.Fields("qty").Value)
      vHoldingStatus = Trim(vRecordset.Fields("holdingstatus").Value)
   End If
   vRecordset.Close
   
   If vBillStatus = 1 And vMemRemainQTY = 0 Then
      MsgBox "กรุณาตรวจสอบเลขที่ใบสั่งขาย/จอง ที่จะเข้าคิวจัดสินค้า เพราะเอกสารดังกล่าวไม่สามารถทำคิวรอจัดสินค้าได้  เนื่องจากดึงไปทำบิลเรียบร้อยแล้ว", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
 If vHoldingStatus = 1 Then
   MsgBox "เอกสารขาย ติดวงเงิน (Holding) ไม่สามารถทำใบจัดคิวขนส่งได้ กรุณาติดต่อ สินเชื่อปลด Hold ", vbCritical, "Send Massage"
   Exit Sub
End If
               
  vQuery = "exec dbo.USP_SO_CheckSendPicking '" & vDocNo & "' " 'pass
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vQueueNo = vRecordset.Fields("docno").Value
    vQueueDate = vRecordset.Fields("queuedatetime").Value
    vQueueStatus = vRecordset.Fields("status").Value
  End If
  vRecordset.Close
  
  If vQueueNo <> "" Then
     If vQueueStatus = 0 Then
       MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ รอจัดสินค้า ", vbInformation, "Send Information"
     ElseIf vQueueStatus = 1 Then
       MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ กำลังจัดสินค้า ", vbInformation, "Send Information"
     ElseIf vQueueStatus = 2 Then
       MsgBox "เอกสารสั่งขายเลขที่ " & vDocNo & " นี้ได้พิมพ์ใบจัดสินค้าไปแล้วเมื่อวันที่เวลา " & vQueueDate & " สถานะตอนนี้ คือ จัดสินค้าเรียบร้อยแล้ว ", vbInformation, "Send Information"
     End If
  End If
  
  
  'vQuery = "exec dbo.USP_NP_SaleOrderPickupZone '" & vDocNo & "' "
  vQuery = "exec dbo.USP_NP_SaleOrderPickupZoneUnitCode '" & vDocNo & "' "  'unitcode แยกหน่วยนับ ท่อ
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     Me.LBLRefDocNo.Caption = Trim(vRecordset.Fields("docno").Value)
     Me.LBLRefDocDate.Caption = Trim(vRecordset.Fields("docdate").Value)
     Me.LBLRefARCode.Caption = Trim(vRecordset.Fields("arcode").Value)
     Me.LBLRefARName.Caption = Trim(vRecordset.Fields("arname").Value)
     Me.ListViewSelectItemPicking.ListItems.Clear
     vRecordset.MoveFirst
     While Not vRecordset.EOF
     i = i + 1
     Set vListItemPicking = Me.ListViewSelectItemPicking.ListItems.Add(, , i)
     vListItemPicking.SubItems(1) = vRecordset.Fields("itemcode").Value
     vListItemPicking.SubItems(2) = vRecordset.Fields("itemname").Value
     vListItemPicking.SubItems(3) = Format(vRecordset.Fields("remainqty").Value, "##,##0.00")
     vListItemPicking.SubItems(4) = Format(vRecordset.Fields("remainqty").Value, "##,##0.00")
     vListItemPicking.SubItems(5) = vRecordset.Fields("unitcode").Value
     vListItemPicking.SubItems(6) = vRecordset.Fields("whcode").Value
     vListItemPicking.SubItems(7) = vRecordset.Fields("shelfcode").Value
     vListItemPicking.SubItems(8) = vRecordset.Fields("zoneid").Value
     vListItemPicking.SubItems(9) = vRecordset.Fields("familygroup").Value
     vListItemPicking.SubItems(10) = vRecordset.Fields("pickzone").Value
     vListItemPicking.Checked = True
     vRecordset.MoveNext
     Wend
     
     PICSelectPrintSlip.Visible = True
     Me.CHKSelectAllItem.Value = 1
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

Private Sub CMDSelectItemBack_Click()
   Me.PICSelectPrintSlip.Visible = True
   Me.PTBPickingQueue.Visible = False
   vCheckSelectItemPickBack = 1
End Sub

Private Sub Form_Load()
Call SetListViewColor(ListViewSelectItemPicking, PicPoint, vbWhite, vbLightGreen)
End Sub

Private Sub ListViewSelectItemPicking_DblClick()
Dim i As Integer
Dim vRecQTY As String
Dim vCheckQty As Double
Dim vPickQTY As Double
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
            Me.ListViewSelectItemPicking.ListItems(i).SubItems(4) = Format(vPickQTY, "##,##0.00")
            Exit Sub
         ElseIf vRecQTY > 0 And Me.ListViewSelectItemPicking.ListItems(i).Checked = False Then
            Me.ListViewSelectItemPicking.ListItems(i).Checked = True
         End If
         vPickQTY = vRecQTY
         If vPickQTY <= vCheckQty And (vCheckQty - vPickQTY) >= 0 Then
            Me.ListViewSelectItemPicking.ListItems(i).SubItems(4) = Format(vPickQTY, "##,##0.00")
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
Dim vPickQTY As Double
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
            Me.ListViewSelectItemPicking.ListItems(i).SubItems(4) = Format(vPickQTY, "##,##0.00")
            Exit Sub
         ElseIf vRecQTY > 0 And Me.ListViewSelectItemPicking.ListItems(i).Checked = False Then
            Me.ListViewSelectItemPicking.ListItems(i).Checked = True
         End If
         vPickQTY = vRecQTY
         If vPickQTY <= vCheckQty And (vCheckQty - vPickQTY) >= 0 Then
            Me.ListViewSelectItemPicking.ListItems(i).SubItems(4) = Format(vPickQTY, "##,##0.00")
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

Private Sub OptSchedule_Click()
Me.MEDTime.Enabled = True
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vGroupDoc As String
Dim vDocTypeID As String
Dim vPrint As Integer
Dim vBillStatus As Integer
Dim vBillType As String
Dim vSend As Integer
Dim vDocNo As String
Dim vSOCountNumber As Integer

If KeyAscii = 13 Then
  vDocNo = Trim(Text101.Text)
  If vDocNo <> "" Then
    vQuery = "select doctypeid, groupdoc,printed from npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        vDocTypeID = Trim(vRecordset.Fields("doctypeid").Value)
        vPrint = Trim(vRecordset.Fields("printed").Value)
    Else
        Call CheckDocument
        Exit Sub
    End If
    vRecordset.Close
    
    If vPrint = 1 Then
     If vDocTypeID = "SO" Then
        vQuery = "exec dbo.USP_SO_SaleOrderDetails '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vBillStatus = Trim(vRecordset.Fields("billstatus").Value)
            vBillType = Trim(vRecordset.Fields("billtype").Value)
            vSend = Trim(vRecordset.Fields("isconditionsend").Value)
            vSaleCode = Trim(vRecordset.Fields("salecode").Value)
            vARCode = Trim(vRecordset.Fields("arcode").Value)
        End If
        vRecordset.Close
     End If
        
    If vBillStatus = 1 Then
      MsgBox "ไม่สามารถพิมพ์ทดแทนได้ เนื่องจาก เอกสารขายดังกล่าวได้ถูกอ้างไปทำบิลเรียบร้อยแล้ว", vbCritical, "Send Error"
    End If
    
    Else
    MsgBox "ไม่สามารถพิมพ์ทดแทนได้ เนื่องจาก ยังไม่ได้พิมพ์ตัวจริงที่หน้าพิมพ์ใบสั่งขาย", vbCritical, "Send Error"
    End If
       
  End If
End If
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
Dim vExistPicking As String
Dim vCheckPicking As Integer
Dim vCountPicking As Integer
Dim vCheckPickingExist As Integer
Dim vPickingNo As String
Dim vRepType As String
Dim vRepID As Integer
Dim vARCode As String
Dim vAddTime As Date
Dim vRequestTime As String

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Text101.Text))
If vDocNo <> "" Then
    vNamePrint = Trim(vSaleCode)
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
    
    
    vQuery = "exec dbo.USP_NP_SearchQueueLogs '" & vDocNo & "','" & vShelfGroup & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckPicking = 1
      vExistPicking = Trim(vRecordset.Fields("docno").Value)
      vCountPicking = Trim(vRecordset.Fields("countpicking").Value) + 1
    Else
      vCheckPicking = 0
      vCountPicking = 1
    End If
    vRecordset.Close
    
    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    vDocuments = vQueueID 'vDocNumber & vHeader & "-" & vPicking
    Else
    vDocuments = vExistPicking
    End If
    
    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    vPickingNo = ""
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    End If
    

    On Error GoTo ErrRunQueueID
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement '" & vDocuments & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vPickingNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vCountPicking & ",0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery
    

    vQuery = "commit tran"
    gConnection.Execute vQuery
            
'Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
MsgBox "ได้ทำการเข้าคิวหยิบสินค้าเรียบร้อยแล้ว ได้เลขที่คิว  " & vDocuments & " ", vbInformation, "Send Information"
    
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถพิมพ์ใบหยิบสินค้าคลัง 010 โซน A ได้ กรุณาพิมพ์ทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_B()
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
Dim vExistPicking As String
Dim vCheckPicking As Integer
Dim vCountPicking As Integer
Dim vCheckPickingExist As Integer
Dim vPickingNo As String
Dim vRepType As String
Dim vRepID As Integer
Dim vARCode As String
Dim vAddTime As Date
Dim vRequestTime As String

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Text101.Text))
If vDocNo <> "" Then
    vNamePrint = Trim(vSaleCode)
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
    
    vQuery = "exec dbo.USP_NP_SearchQueueLogs '" & vDocNo & "','" & vShelfGroup & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckPicking = 1
      vExistPicking = Trim(vRecordset.Fields("docno").Value)
      vCountPicking = Trim(vRecordset.Fields("countpicking").Value) + 1
    Else
      vCheckPicking = 0
      vCountPicking = 1
    End If
    vRecordset.Close

    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    vDocuments = vQueueID 'vDocNumber & vHeader & "-" & vPicking
    Else
    vDocuments = vExistPicking
    End If

    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    vPickingNo = ""
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    End If
    
    
    On Error GoTo ErrRunQueueID
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement '" & vDocuments & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vPickingNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vCountPicking & ",0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery
            
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
MsgBox "ได้ทำการเข้าคิวหยิบสินค้าเรียบร้อยแล้ว ได้เลขที่คิว  " & vDocuments & " ", vbInformation, "Send Information"
    
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถพิมพ์ใบหยิบสินค้าคลัง 010 โซน B ได้ กรุณาพิมพ์ทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub


Public Sub PrintPicking_D()
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
Dim vExistPicking As String
Dim vCheckPicking As Integer
Dim vCountPicking As Integer
Dim vCheckPickingExist As Integer
Dim vPickingNo As String
Dim vRepType As String
Dim vRepID As Integer
Dim vARCode As String
Dim vAddTime As Date
Dim vRequestTime As String

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Text101.Text))
If vDocNo <> "" Then
    vNamePrint = Trim(vSaleCode)
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
    
    vQuery = "exec dbo.USP_NP_SearchQueueLogs '" & vDocNo & "','" & vShelfGroup & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckPicking = 1
      vExistPicking = Trim(vRecordset.Fields("docno").Value)
      vCountPicking = Trim(vRecordset.Fields("countpicking").Value) + 1
    Else
      vCheckPicking = 0
      vCountPicking = 1
    End If
    vRecordset.Close
    
    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    vDocuments = vQueueID 'vDocNumber & vHeader & "-" & vPicking
    Else
    vDocuments = vExistPicking
    End If

    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    vPickingNo = ""
    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    End If
    
        
    On Error GoTo ErrRunQueueID
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement '" & vDocuments & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vPickingNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vCountPicking & ",0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery
              
   
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
MsgBox "ได้ทำการเข้าคิวหยิบสินค้าเรียบร้อยแล้ว ได้เลขที่คิว  " & vDocuments & " ", vbInformation, "Send Information"
    
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถพิมพ์ใบหยิบสินค้าคลัง 010 โซน D ได้ กรุณาพิมพ์ทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_C()
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
Dim vExistPicking As String
Dim vCheckPicking As Integer
Dim vCountPicking As Integer
Dim vCheckPickingExist As Integer
Dim vPickingNo As String
Dim vRepType As String
Dim vRepID As Integer
Dim vARCode As String
Dim vAddTime As Date
Dim vRequestTime As String

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Text101.Text))
If vDocNo <> "" Then
    vNamePrint = Trim(vSaleCode)
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
    
    vQuery = "exec dbo.USP_NP_SearchQueueLogs '" & vDocNo & "','" & vShelfGroup & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckPicking = 1
      vExistPicking = Trim(vRecordset.Fields("docno").Value)
      vCountPicking = Trim(vRecordset.Fields("countpicking").Value) + 1
    Else
      vCheckPicking = 0
      vCountPicking = 1
    End If
    vRecordset.Close
        
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    vDocuments = vQueueID 'vDocNumber & vHeader & "-" & vPicking
    Else
    vDocuments = vExistPicking
    End If

    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    vPickingNo = ""
    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    End If

    
    On Error GoTo ErrRunQueueID
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement '" & vDocuments & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vPickingNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vCountPicking & ",0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery
               
    vQuery = "commit tran"
    gConnection.Execute vQuery
    
Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
MsgBox "ได้ทำการเข้าคิวหยิบสินค้าเรียบร้อยแล้ว ได้เลขที่คิว  " & vDocuments & " ", vbInformation, "Send Information"
    
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถพิมพ์ใบหยิบสินค้าคลัง 015 โซน C ได้ กรุณาพิมพ์ทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_M()
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
Dim vExistPicking As String
Dim vCheckPicking As Integer
Dim vCountPicking As Integer
Dim vCheckPickingExist As Integer
Dim vPickingNo As String
Dim vRepType As String
Dim vRepID As Integer
Dim vARCode As String
Dim vAddTime As Date
Dim vRequestTime As String

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Text101.Text))
If vDocNo <> "" Then
    vNamePrint = Trim(vSaleCode)
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
    
    vQuery = "exec dbo.USP_NP_SearchQueueLogs '" & vDocNo & "','" & vShelfGroup & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckPicking = 1
      vExistPicking = Trim(vRecordset.Fields("docno").Value)
      vCountPicking = Trim(vRecordset.Fields("countpicking").Value) + 1
    Else
      vCheckPicking = 0
      vCountPicking = 1
    End If
    vRecordset.Close

    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    vDocuments = vQueueID 'vDocNumber & vHeader & "-" & vPicking
    Else
    vDocuments = vExistPicking
    End If

    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    vPickingNo = ""
    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    End If
    
    
    On Error GoTo ErrRunQueueID
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement '" & vDocuments & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vPickingNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vCountPicking & ",0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery
             
    vQuery = "commit tran"
    gConnection.Execute vQuery
    
Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 1)
'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
MsgBox "ได้ทำการเข้าคิวหยิบสินค้าเรียบร้อยแล้ว ได้เลขที่คิว  " & vDocuments & " ", vbInformation, "Send Information"
    
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถพิมพ์ใบหยิบสินค้าคลัง 014 โซน M ได้ กรุณาพิมพ์ทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_K()
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
Dim vExistPicking As String
Dim vCheckPicking As Integer
Dim vCountPicking As Integer
Dim vCheckPickingExist As Integer
Dim vPickingNo As String
Dim vRepType As String
Dim vRepID As Integer
Dim vARCode As String
Dim vAddTime As Date
Dim vRequestTime As String

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Text101.Text))
If vDocNo <> "" Then
    vNamePrint = Trim(vSaleCode)
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
    
    vQuery = "exec dbo.USP_NP_SearchQueueLogs '" & vDocNo & "','" & vShelfGroup & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckPicking = 1
      vExistPicking = Trim(vRecordset.Fields("docno").Value)
      vCountPicking = Trim(vRecordset.Fields("countpicking").Value) + 1
    Else
      vCheckPicking = 0
      vCountPicking = 1
    End If
    vRecordset.Close

    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    vDocuments = vQueueID 'vDocNumber & vHeader & "-" & vPicking
    Else
    vDocuments = vExistPicking
    End If

    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    vPickingNo = ""
    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    End If
    
    
    On Error GoTo ErrRunQueueID
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement '" & vDocuments & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vPickingNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vCountPicking & ",0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery
             
    vQuery = "commit tran"
    gConnection.Execute vQuery
    
Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 1)
'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
MsgBox "ได้ทำการเข้าคิวหยิบสินค้าเรียบร้อยแล้ว ได้เลขที่คิว  " & vDocuments & " ", vbInformation, "Send Information"
    
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถพิมพ์ใบหยิบสินค้าคลัง 014 โซน M ได้ กรุณาพิมพ์ทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_Y()
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
Dim vExistPicking As String
Dim vCheckPicking As Integer
Dim vCountPicking As Integer
Dim vCheckPickingExist As Integer
Dim vPickingNo As String
Dim vRepType As String
Dim vRepID As Integer
Dim vARCode As String
Dim vAddTime As Date
Dim vRequestTime As String

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Text101.Text))
If vDocNo <> "" Then
    vNamePrint = Trim(vSaleCode)
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
    
    vQuery = "exec dbo.USP_NP_SearchQueueLogs '" & vDocNo & "','" & vShelfGroup & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckPicking = 1
      vExistPicking = Trim(vRecordset.Fields("docno").Value)
      vCountPicking = Trim(vRecordset.Fields("countpicking").Value) + 1
    Else
      vCheckPicking = 0
      vCountPicking = 1
    End If
    vRecordset.Close

    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    vDocuments = vQueueID 'vDocNumber & vHeader & "-" & vPicking
    Else
    vDocuments = vExistPicking
    End If

    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    vPickingNo = ""
    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    End If

    
    On Error GoTo ErrRunQueueID
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement '" & vDocuments & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vPickingNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vCountPicking & ",0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery
            
   
    vQuery = "commit tran"
    gConnection.Execute vQuery
    
  Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 0)
  'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 0)
  MsgBox "ได้ทำการเข้าคิวหยิบสินค้าเรียบร้อยแล้ว ได้เลขที่คิว  " & vDocuments & " ", vbInformation, "Send Information"
    
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถพิมพ์ใบหยิบสินค้าคลัง 016 โซน Y ได้ กรุณาพิมพ์ทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Public Sub PrintPicking_H()
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
Dim vExistPicking As String
Dim vCheckPicking As Integer
Dim vCountPicking As Integer
Dim vCheckPickingExist As Integer
Dim vPickingNo As String
Dim vRepType As String
Dim vRepID As Integer
Dim vARCode As String
Dim vAddTime As Date
Dim vRequestTime As String

On Error GoTo ErrRollBack

vDocNo = UCase(Trim(Text101.Text))
If vDocNo <> "" Then
    vNamePrint = Trim(vSaleCode)
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
    
    vQuery = "exec dbo.USP_NP_SearchQueueLogs '" & vDocNo & "','" & vShelfGroup & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckPicking = 1
      vExistPicking = Trim(vRecordset.Fields("docno").Value)
      vCountPicking = Trim(vRecordset.Fields("countpicking").Value) + 1
    Else
      vCheckPicking = 0
      vCountPicking = 1
    End If
    vRecordset.Close

    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vQueueID = Trim(vRecordset.Fields("autonumber").Value)
    End If
    vRecordset.Close
    vDocuments = vQueueID 'vDocNumber & vHeader & "-" & vPicking
    Else
    vDocuments = vExistPicking
    End If

    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    vPickingNo = ""
    
    If vCheckPicking = 0 Then
    vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
    gConnection.Execute vQuery
    End If
    
    
    On Error GoTo ErrRunQueueID
    Call vGetCustomerZone
    vQuery = "exec dbo.USP_NP_NewInsertDataQueueManagement '" & vDocuments & "','" & vDocdate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vDocNo & "','" & vPickingNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vCountPicking & ",0,'" & vRequestTime & "'," & vCustomerZone & " "
    gConnection.Execute vQuery
            
 
    vQuery = "commit tran"
    gConnection.Execute vQuery
    

Call PrintPickingSlipHeader(vDocNo, vWHCode, vShelfGroup, 1)
'Call PrintPickingSlip(vDocNo, vWHCode, vShelfGroup, 1)
MsgBox "ได้ทำการเข้าคิวหยิบสินค้าเรียบร้อยแล้ว ได้เลขที่คิว  " & vDocuments & " ", vbInformation, "Send Information"

    
End If

ErrRollBack:
If Err.Description <> "" Then
  vQuery = "rollback tran"
  gConnection.Execute vQuery
  MsgBox Err.Description
  MsgBox "ไม่สามารถพิมพ์ใบหยิบสินค้าคลัง 014 โซน H ได้ กรุณาพิมพ์ทดแทนเพื่อเข้าคิวหยิบสินค้าใหม่"
End If

ErrRunQueueID:
If Err.Number = -2147217873 Then
    vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
    gConnection.Execute vQuery
    MsgBox "กดพิมพ์ใบจัดสินค้า อีกรอบ เลขที่คิวเกิดการชนกัน", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

'Public Sub PrintPickingSlipHeader(vSaleOrder As String, vWHCode As String, vShelfCode As String, vZone As Integer)
'Dim vPrinterName As String
'Dim printerObj As Printer
'Dim vRecordset As New ADODB.Recordset
'Dim vQuery As String
'Dim vDocNo As String
'Dim vLineX As Integer
'Dim vLineY As Integer
'Dim vStartX As Integer
'Dim vStartY As Integer
'Dim i As Integer
'Dim prnPrinter As Printer
'Dim lngRetVal As Long
'Dim Driver As String
'Dim n As Integer

'If vZone = 0 Then
'vPrinterName = Trim("\\nova\EPS TM-T88III-NP")
'For Each printerObj In Printers
 ' If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
  '  Set Printer = printerObj
   ' Set printerObj = Nothing
    'Exit For
  'End If
'Next
'End If
   
'If vZone = 1 Then
'vPrinterName = Trim("\\nova\EPS-TM-PickingOutlet")
'For Each printerObj In Printers
 ' If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
  '  Set Printer = printerObj
   ' Set printerObj = Nothing
    'Exit For
  'End If
'Next
'End If

 '   vQuery = "exec dbo.USP_SO_PickingQueue '" & vSaleOrder & "','" & vWHCode & "','" & vShelfCode & "' "
  '  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   '
    '  Printer.Font.Name = "AngsanaUPC"
     ' Printer.Font.Size = 10
      'Printer.CurrentX = 0
      'Printer.CurrentY = 0
      'Printer.Print Trim("_______________________________________________________________________________________")

      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 50
      'Printer.FontBold = True
      'Printer.CurrentX = 1800
      'Printer.CurrentY = 0
      'Printer.Print Trim(vRecordset.Fields("queueno").Value)
      '
      'Printer.Font.Name = "3 of 9 Barcode"
      'Printer.Font.Size = 40
      'Printer.FontBold = False
      'Printer.CurrentX = 1400
      'Printer.CurrentY = 1000
      'Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 10
      'Printer.CurrentX = 0
      'Printer.CurrentY = 1400
      'Printer.Print Trim("_______________________________________________________________________________________")
    
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 1500
      'Printer.CurrentY = 1650
      'Printer.Print Trim("ทดแทนต้นฉบับใบจัดสินค้า")
      
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 0
      'Printer.CurrentY = 1900
      'Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 2200
      'Printer.CurrentY = 1900
      'Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 2200
      'Printer.CurrentY = 2150
      'Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 0
      'Printer.CurrentY = 2400
      'Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 0
      ''Printer.CurrentY = 2650
      'Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 0
      'Printer.CurrentY = 2900
      'Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 0
      'Printer.CurrentY = 3150
      'if vRecordset.Fields("isconditionsend").Value = 0 Then
        '    Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("รับเอง")
      'Else
       '     Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("ส่งให้")
      'End If
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 1400
      'Printer.CurrentY = 3150
      'Printer.Print Trim("เวลาที่ต้องการรับของ : ") & Trim(vRecordset.Fields("requesttime").Value)
       '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 14
      'Printer.CurrentX = 0
      'Printer.CurrentY = 3400
      'Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value)
      '
      'if Trim(vRecordset.Fields("carlicense").Value) <> "" Then
       ' Printer.Font.Name = "AngsanaUPC"
        'Printer.Font.Size = 16
        'Printer.CurrentX = 1400
        'Printer.CurrentY = 3300
        'Printer.FontBold = True
        'Printer.FontUnderline = True
        'Printer.Print Trim("ทะเบียนรถขนส่ง : ") & Trim(vRecordset.Fields("carlicense").Value)
      'End If
      
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 14
      'Printer.CurrentX = 0
      'Printer.CurrentY = 3800
      'Printer.Print Trim(vRecordset.Fields("customerzone").Value)
    '
    'Printer.CurrentX = 0
    'Printer.CurrentY = 3900
    'Printer.Print Trim("______________________________________________________________________________________________")

    'Printer.Font.Name = "AngsanaUPC"
    'Printer.Font.Size = 30
    'Printer.CurrentX = 800
    'Printer.CurrentY = 4250
    'Printer.Print Trim("ครบ")
    
    'Printer.Font.Name = "AngsanaUPC"
    'Printer.Font.Size = 30
    'Printer.CurrentX = 2800
    'Printer.CurrentY = 4250
    'Printer.Print Trim("ไม่ครบ")
    
    'Printer.Font.Name = "AngsanaUPC"
    'Printer.Font.Size = 20
    'Printer.CurrentX = 0
    'Printer.CurrentY = 4400
    'Printer.Print Trim("______")
    
    'Printer.Font.Name = "AngsanaUPC"
    'Printer.Font.Size = 20
    'Printer.CurrentX = 2100
    'Printer.CurrentY = 4400
    'Printer.Print Trim("______")

    'Printer.CurrentX = 0
    'Printer.CurrentY = 4500
    'Printer.Print Trim("______________________________________________________________________________________________")
    'End If
    'vRecordset.Close

     ' Printer.Font.Size = 10
      'Printer.CurrentX = 0
      'Printer.CurrentY = Printer.CurrentY
      'Printer.Print Trim("V:\Reports\RP_BC_PickingQueueHead.rpt")
      'Printer.CurrentX = Printer.CurrentX + 2000
      'Printer.Print Trim("วันที่พิมพ์ :") & Now
           
    'Printer.EndDoc

      'End Sub
      
'Public Sub PrintPickingSlip(vSaleOrder As String, vWHCode As String, vShelfCode As String, vZone As Integer)
'Dim vPrinterName As String
'Dim printerObj As Printer
'Dim vRecordset As New ADODB.Recordset
'Dim vQuery As String
'Dim vDocNo As String
'Dim vLineX As Integer
'Dim vLineY As Integer
'Dim vStartX As Integer
'Dim vStartY As Integer
'Dim i As Integer
'Dim prnPrinter As Printer
'Dim lngRetVal As Long
'Dim Driver As String
'Dim n As Integer
'Dim vItemName As String

'If vZone = 0 Then
'vPrinterName = Trim("\\nova\EPS TM-T88III-NP")
'For Each printerObj In Printers
 ' If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
  '  Set Printer = printerObj
   ' Set printerObj = Nothing
    'Exit For
  'End If
'Next
'End If
   
'If vZone = 1 Then
'vPrinterName = Trim("\\nova\EPS-TM-PickingOutlet")
'For Each printerObj In Printers
 ' If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
  '  Set Printer = printerObj
   ' Set printerObj = Nothing
    'Exit For
  'End If
'Next
'End If

 '   vQuery = "exec dbo.USP_SO_PickingQueue '" & vSaleOrder & "','" & vWHCode & "','" & vShelfCode & "' "
  '  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   '
    '  Printer.Font.Name = "AngsanaUPC"
     ' Printer.Font.Size = 10
      'Printer.CurrentX = 0
      'Printer.CurrentY = 0
      'Printer.Print Trim("_______________________________________________________________________________________")
'
 '     Printer.Font.Name = "AngsanaUPC"
  '    Printer.Font.Size = 50
   '   Printer.FontBold = True
    '  Printer.CurrentX = 1800
     ' Printer.CurrentY = 0
      'Printer.Print Trim(vRecordset.Fields("queueno").Value)
      '
      ''Printer.Font.Name = "3 of 9 Barcode"
      'Printer.Font.Size = 40
      '''Printer.FontBold = False
      'Printer.CurrentX = 1400
      'Printer.CurrentY = 1000
      'Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 10
      'Printer.CurrentX = 0
      'Printer.CurrentY = 1400
      'Printer.Print Trim("_______________________________________________________________________________________")
    '
     ' Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 1500
      ''Printer.CurrentY = 1650
      'Printer.Print Trim("ทดแทนต้นฉบับใบจัดสินค้า")
      
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 0
      'Printer.CurrentY = 1900
      'Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 2200
      'Printer.CurrentY = 1900
      'Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 2200
      'Printer.CurrentY = 2150
      'Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 0
      'Printer.CurrentY = 2400
      'Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 14
      'Printer.CurrentX = 2200
      'Printer.CurrentY = 2400
      'Printer.Print "อ้างถึง :" & Trim(vRecordset.Fields("transferno").Value)
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 0
      'Printer.CurrentY = 2650
      'Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 0
      'Printer.CurrentY = 2900
      'Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 11
      'Printer.CurrentX = 0
      'Printer.CurrentY = 3150
      'if vRecordset.Fields("isconditionsend").Value = 0 Then
        '    Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("รับเอง")
      'Else
       '     Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("ส่งให้")
      'End If
       '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 16
      'Printer.FontBold = True
      'Printer.CurrentX = 0
      'Printer.CurrentY = 3400
      'Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value)

      'If Trim(vRecordset.Fields("carlicense").Value) <> "" Then
       ' Printer.Font.Name = "AngsanaUPC"
        'Printer.Font.Size = 16
        'Printer.CurrentX = 1400
        'Printer.CurrentY = 3400
        'Printer.FontBold = True
        'Printer.FontUnderline = True
        'Printer.Print Trim("ทะเบียนรถขนส่ง : ") & Trim(vRecordset.Fields("carlicense").Value)
      'End If
            
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 14
      'Printer.FontBold = True
      'Printer.FontUnderline = False
      'Printer.CurrentX = 0
      'Printer.CurrentY = 3800
      'Printer.Print Trim("เวลารับของ : ") & Trim(vRecordset.Fields("requesttime").Value)
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 14
      'Printer.CurrentX = 0
      'Printer.CurrentY = 4150
      ''Printer.Print Trim(vRecordset.Fields("customerzone").Value)
      
      'vRecordset.MoveFirst
      'vLineX = 50
      'vLineY = 50
      '
      'Printer.Font.Name = "AngsanaUPC"
      'Printer.Font.Size = 10
      'Printer.CurrentX = 0
      'Printer.CurrentY = Printer.CurrentY - 30
      'Printer.FontBold = False
      'Printer.FontUnderline = False
      'Printer.Print Trim("-----------------------------------------------------------------------------------------------")
      'n = 1
      'While Not vRecordset.EOF
       '   Printer.Font.Size = 11
          
        '  Printer.CurrentX = 0
         ' Printer.CurrentY = Printer.CurrentY
          'Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfcode1").Value) & "       " & Trim("OnHand") & "(" & Trim(vRecordset.Fields("shelfcode").Value) & ")" & ": " & Trim(vRecordset.Fields("qtylocation").Value) & "  " & Trim(vRecordset.Fields("stkunitcode").Value) & "        " & "ยอดรวมคลัง :" & Trim(vRecordset.Fields("StkWHCode").Value) & "  " & Trim(vRecordset.Fields("stkunitcode").Value)
          
          'Printer.CurrentX = 0
          'Printer.CurrentY = Printer.CurrentY
          'Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value) & "                  " & " ขายชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)
          
          'vItemName = Trim(vRecordset.Fields("itemname").Value) & Trim(vRecordset.Fields("descriptionline"))
          'If Len(vItemName) <= 55 Then
           '  Printer.CurrentX = 0
            ' Printer.CurrentY = Printer.CurrentY
             'Printer.Print "ชื่อสินค้า :" & vItemName
          'Else
           '  Printer.CurrentX = 0
            ' Printer.CurrentY = Printer.CurrentY
             'Printer.Print "ชื่อสินค้า :" & Left(vItemName, 55)
             
             'Printer.CurrentX = 600
             'Printer.CurrentY = Printer.CurrentY
             'Printer.Print Right(vItemName, Len(vItemName) - 55)
          'End If
          
          'Printer.CurrentX = Printer.CurrentX + 15
          'Printer.CurrentY = Printer.CurrentY + 100
          'Printer.FontBold = True
          'Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")
          
          'Printer.CurrentX = 0
          'Printer.CurrentY = Printer.CurrentY - 50
          'Printer.FontBold = False
          'Printer.Print Trim("-----------------------------------------------------------------------------------------------")
          
      'vRecordset.MoveNext
      'n = n + 1
      'Wend
    'End If
    'vRecordset.Close
    'Printer.CurrentX = 0
    'Printer.CurrentY = Printer.CurrentY - 400
    'Printer.Print Trim("_______________________________________________________________________________________________")
    
    'Printer.CurrentX = 0
    'Printer.CurrentY = Printer.CurrentY + vLineY
    'Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"
    '
    'Printer.CurrentX = 0
    'Printer.CurrentY = Printer.CurrentY + 150
    'Printer.Print "         _____________                                    ______________"
     '
    'Printer.CurrentX = 0
    'Printer.CurrentY = Printer.CurrentY
    'Printer.Print Trim("______________________________________________________________________________________________")
    

     ' Printer.Font.Size = 10
      'Printer.CurrentX = 0
      'Printer.CurrentY = Printer.CurrentY
      'Printer.Print Trim("V:\Reports\RP_BC_PickingQueue.rpt")
      'Printer.CurrentX = Printer.CurrentX + 2000
      'Printer.Print Trim("วันที่พิมพ์ :") & Now
       '
    'Printer.EndDoc
'End Sub

Public Sub vGetCustomerZone()
If Me.OptNormal.Value = True Then
  vCustomerZone = 0
ElseIf Me.OptMain.Value = True Then
  vCustomerZone = 1
ElseIf Me.OptOutLet.Value = True Then
  vCustomerZone = 2
End If
End Sub
