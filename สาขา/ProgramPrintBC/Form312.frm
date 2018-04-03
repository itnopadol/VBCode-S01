VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form312 
   Caption         =   "ทำใบจัดคิว"
   ClientHeight    =   9600
   ClientLeft      =   2250
   ClientTop       =   915
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form312.frx":0000
   ScaleHeight     =   9600
   ScaleMode       =   0  'User
   ScaleWidth      =   16560.66
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CHKReqPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ขอพิมพ์ฟอร์ม A4 กรณีพิมพ์กระดาษครึ่งหน้าไม่ได้"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3510
      TabIndex        =   63
      Top             =   8010
      Width           =   5415
   End
   Begin VB.CheckBox Check102 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ไม่ต้องการล้างหน้าจอ"
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
      Left            =   9045
      TabIndex        =   62
      Top             =   8010
      Width           =   2040
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   2925
      Top             =   8505
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
   Begin VB.CommandButton CMD007 
      Height          =   330
      Left            =   2880
      Picture         =   "Form312.frx":9673
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "ออกหน้า ทำใบจัดคิวสินค้า"
      Top             =   8010
      Width           =   330
   End
   Begin VB.CommandButton CMD006 
      Height          =   330
      Left            =   2475
      Picture         =   "Form312.frx":99DD
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "ยกเลิกใบจัดคิวสินค้า"
      Top             =   8010
      Width           =   330
   End
   Begin VB.CommandButton CMD005 
      Height          =   330
      Left            =   2070
      Picture         =   "Form312.frx":BB7B
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "พิมพ์ใบจัดคิวสินค้า"
      Top             =   8010
      Width           =   330
   End
   Begin VB.CommandButton CMD004 
      Height          =   330
      Left            =   1665
      Picture         =   "Form312.frx":BEBD
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "ค้นหา ใบจัดคิวสินค้า"
      Top             =   8010
      Width           =   330
   End
   Begin VB.CommandButton CMD003 
      Height          =   330
      Left            =   1260
      Picture         =   "Form312.frx":C28A
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "บันทึกข้อมูลที่สร้างเอกสารใหม่และอัพเดท"
      Top             =   8010
      Width           =   330
   End
   Begin VB.CommandButton CMD002 
      Height          =   330
      Left            =   855
      Picture         =   "Form312.frx":C5B1
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "เคลียร์เอกสาร"
      Top             =   8010
      Width           =   330
   End
   Begin VB.CommandButton CMD001 
      Height          =   330
      Left            =   5985
      Picture         =   "Form312.frx":C996
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "ค้นหา บิล POS ใบสั่งขาย และสั่งจอง มาทำใบจัดคิว"
      Top             =   1170
      Width           =   330
   End
   Begin VB.CheckBox Check101 
      BackColor       =   &H80000009&
      Caption         =   ": POS"
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
      Left            =   855
      TabIndex        =   0
      Top             =   1170
      Width           =   825
   End
   Begin VB.ComboBox CMB101 
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
      ItemData        =   "Form312.frx":CD63
      Left            =   3600
      List            =   "Form312.frx":CD65
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1170
      Width           =   2310
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3480
      Left            =   855
      TabIndex        =   4
      Top             =   4455
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   6138
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "          จัดคิว"
      TabPicture(0)   =   "Form312.frx":CD67
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(11)=   "Text101"
      Tab(0).Control(12)=   "Text102"
      Tab(0).Control(13)=   "Text103"
      Tab(0).Control(14)=   "Text105"
      Tab(0).Control(15)=   "Text106"
      Tab(0).Control(16)=   "Text107"
      Tab(0).Control(17)=   "Text108"
      Tab(0).Control(18)=   "DTPicker102"
      Tab(0).Control(19)=   "DTPicker101"
      Tab(0).Control(20)=   "MaskEdBox101"
      Tab(0).Control(21)=   "CMD101"
      Tab(0).Control(22)=   "CMD105"
      Tab(0).Control(23)=   "CMB102"
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "          ผู้รับสินค้า"
      TabPicture(1)   =   "Form312.frx":CD83
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(1)=   "Label17"
      Tab(1).Control(2)=   "Label18"
      Tab(1).Control(3)=   "Text201"
      Tab(1).Control(4)=   "Text202"
      Tab(1).Control(5)=   "Text203"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "          สถานที่"
      TabPicture(2)   =   "Form312.frx":CD9F
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label19"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label20"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label21"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label22"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label23"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label24"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label25"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label26"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label27"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Text301"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Text302"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Text303"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Text304"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Text305"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Text306"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Text307"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Text308"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Text309"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "CMD301"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "CMD302"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).ControlCount=   20
      Begin VB.ComboBox CMB102 
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
         Left            =   -73605
         TabIndex        =   12
         Top             =   2970
         Width           =   1995
      End
      Begin VB.CommandButton CMD302 
         Height          =   285
         Left            =   8370
         Picture         =   "Form312.frx":CDBB
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "ค้นหา เส้นทางขนส่ง"
         Top             =   675
         Width           =   330
      End
      Begin VB.CommandButton CMD301 
         Height          =   285
         Left            =   3645
         Picture         =   "Form312.frx":D188
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "ค้นหา สถานที่"
         Top             =   675
         Width           =   330
      End
      Begin VB.TextBox Text309 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1230
         Left            =   6570
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   2025
         Width           =   3210
      End
      Begin VB.TextBox Text308 
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
         Left            =   6570
         TabIndex        =   30
         Top             =   1575
         Width           =   2175
      End
      Begin VB.TextBox Text307 
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
         Left            =   6570
         TabIndex        =   29
         Top             =   1125
         Width           =   2175
      End
      Begin VB.TextBox Text306 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   6570
         TabIndex        =   27
         Top             =   675
         Width           =   1770
      End
      Begin VB.TextBox Text305 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   780
         Left            =   1890
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   2475
         Width           =   3255
      End
      Begin VB.TextBox Text304 
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
         Left            =   1890
         TabIndex        =   25
         Top             =   2025
         Width           =   2175
      End
      Begin VB.TextBox Text303 
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
         Left            =   1890
         TabIndex        =   24
         Top             =   1575
         Width           =   2175
      End
      Begin VB.TextBox Text302 
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
         Left            =   1890
         TabIndex        =   23
         Top             =   1125
         Width           =   2175
      End
      Begin VB.TextBox Text301 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1890
         TabIndex        =   21
         Top             =   675
         Width           =   1725
      End
      Begin VB.CommandButton CMD105 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -66585
         Picture         =   "Form312.frx":D555
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1845
         Width           =   330
      End
      Begin VB.CommandButton CMD101 
         Height          =   285
         Left            =   -71670
         Picture         =   "Form312.frx":D922
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "สร้างเลขที่เอกสารใบจัดคิว"
         Top             =   945
         Width           =   330
      End
      Begin VB.TextBox Text203 
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
         Height          =   285
         Left            =   -71625
         TabIndex        =   20
         Top             =   1890
         Width           =   3660
      End
      Begin VB.TextBox Text202 
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
         Height          =   285
         Left            =   -71625
         TabIndex        =   19
         Top             =   1395
         Width           =   3660
      End
      Begin VB.TextBox Text201 
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
         Height          =   285
         Left            =   -71625
         TabIndex        =   18
         Top             =   900
         Width           =   3660
      End
      Begin MSMask.MaskEdBox MaskEdBox101 
         Height          =   285
         Left            =   -73605
         TabIndex        =   10
         Top             =   2160
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
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
      Begin MSComCtl2.DTPicker DTPicker101 
         Height          =   330
         Left            =   -73605
         TabIndex        =   8
         Top             =   1350
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   57147393
         CurrentDate     =   38695
      End
      Begin MSComCtl2.DTPicker DTPicker102 
         Height          =   330
         Left            =   -73605
         TabIndex        =   9
         Top             =   1755
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   57147393
         CurrentDate     =   38695
      End
      Begin VB.TextBox Text108 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   -68700
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   2250
         Width           =   3390
      End
      Begin VB.TextBox Text107 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
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
         Left            =   -68700
         TabIndex        =   15
         Top             =   1845
         Width           =   2085
      End
      Begin VB.TextBox Text106 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   825
         Left            =   -68700
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   945
         Width           =   3390
      End
      Begin VB.TextBox Text105 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
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
         Left            =   -68700
         TabIndex        =   13
         Top             =   540
         Width           =   2040
      End
      Begin VB.TextBox Text103 
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
         Height          =   285
         Left            =   -73605
         TabIndex        =   11
         Text            =   "0"
         Top             =   2565
         Width           =   1950
      End
      Begin VB.TextBox Text102 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   -73605
         TabIndex        =   6
         Top             =   945
         Width           =   1905
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
         Left            =   -73605
         TabIndex        =   5
         Top             =   540
         Width           =   1230
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "หมายเหตุ :"
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
         Left            =   5220
         TabIndex        =   61
         Top             =   2025
         Width           =   1275
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "ชื่อ2 :"
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
         Left            =   5310
         TabIndex        =   60
         Top             =   1575
         Width           =   1185
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "ชื่อ1 :"
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
         Left            =   5355
         TabIndex        =   59
         Top             =   1125
         Width           =   1140
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "ID เส้นทาง :"
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
         Left            =   5445
         TabIndex        =   58
         Top             =   675
         Width           =   1050
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "หมายเหตุ :"
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
         Left            =   360
         TabIndex        =   57
         Top             =   2475
         Width           =   1455
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "จังหวัด :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   56
         Top             =   2025
         Width           =   1275
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "อำเภอ :"
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
         Left            =   585
         TabIndex        =   55
         Top             =   1575
         Width           =   1230
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "ตำบล :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   630
         TabIndex        =   54
         Top             =   1125
         Width           =   1185
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "ID สถานที่ :"
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
         TabIndex        =   53
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "เบอร์มือถือ :"
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
         Left            =   -73245
         TabIndex        =   52
         Top             =   1890
         Width           =   1545
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "เบอร์บ้าน :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73245
         TabIndex        =   51
         Top             =   1395
         Width           =   1545
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "ชื่อ - สกุล :"
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
         Left            =   -73155
         TabIndex        =   50
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "แผนที่ :"
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
         Left            =   -70410
         TabIndex        =   49
         Top             =   1890
         Width           =   1635
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "หมายเหตุ :"
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
         Left            =   -70185
         TabIndex        =   48
         Top             =   2250
         Width           =   1410
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "จุดสังเกตที่สำคัญ :"
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
         Left            =   -70275
         TabIndex        =   47
         Top             =   945
         Width           =   1500
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "ระยะทางโดยประมาณ(กม.) :"
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
         Left            =   -70905
         TabIndex        =   46
         Top             =   540
         Width           =   2130
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "ความสำคัญ :"
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
         Left            =   -74775
         TabIndex        =   45
         Top             =   2970
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "เงินที่ต้องเก็บ :"
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
         Left            =   -74865
         TabIndex        =   44
         Top             =   2565
         Width           =   1185
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "เวลานัดรับ :"
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
         Left            =   -74730
         TabIndex        =   43
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "วันที่นัดรับ :"
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
         Left            =   -74685
         TabIndex        =   42
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Left            =   -74820
         TabIndex        =   41
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   -74775
         TabIndex        =   40
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   -74595
         TabIndex        =   39
         Top             =   540
         Width           =   915
      End
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2805
      Left            =   855
      TabIndex        =   3
      Top             =   1575
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   4948
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
      BackColor       =   16641742
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่อ้างอิง"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "รหัสสินค้า"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ชื่อสินค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "จำนวน"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ค้างส่ง"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "หน่วยนับ"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "HeaderID"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "DetailID"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "IsCancel"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Image Image103 
      Height          =   300
      Left            =   13365
      Picture         =   "Form312.frx":DC66
      Top             =   1035
      Width           =   570
   End
   Begin VB.Image Image102 
      Height          =   300
      Left            =   13365
      Picture         =   "Form312.frx":E10F
      Top             =   1035
      Width           =   570
   End
   Begin VB.Image Image101 
      Height          =   300
      Left            =   13365
      Picture         =   "Form312.frx":E64B
      Top             =   1035
      Width           =   570
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทเอกสารอ้างอิง"
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
      Left            =   1800
      TabIndex        =   38
      Top             =   1215
      Width           =   1725
   End
End
Attribute VB_Name = "Form312"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check101_Click()
On Error GoTo ErrDescription

If Check101.Value = 1 Then
    CMB101.Text = Trim("บิลขาย")
    CMB101.Enabled = False
Else
    CMB101.Text = Trim("ใบสั่งขาย/จอง")
    CMB101.Enabled = True
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD001_Click()
On Error GoTo ErrDescription

MDIFrmProgramPrint.Enabled = False
FrmOrder004.Show

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD002_Click()
On Error GoTo ErrDescription

vIsOpen1 = 0
Check101.Value = 0
ListView101.ListItems.Clear
DTPicker101.Value = Now
DTPicker102.Value = Now
CMD006.Enabled = False
Text101.Text = ""
Text102.Text = ""
Text103.Text = "0"
CMB102.Text = ""
Text105.Text = ""
Text106.Text = ""
Text107.Text = ""
Text108.Text = ""
MaskEdBox101.Mask = "##:##"
Text201.Text = ""
Text202.Text = ""
Text203.Text = ""
Text301.Text = ""
Text302.Text = ""
Text303.Text = ""
Text304.Text = ""
Text305.Text = ""
Text306.Text = ""
Text307.Text = ""
Text308.Text = ""
Text309.Text = ""
Form312.Image101.Visible = True
Form312.Image102.Visible = False
Form312.Image103.Visible = False

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD003_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As String
Dim vDocNo As String
Dim vDocdate As Date
Dim vDueDate As Date
Dim vDueTime As String
Dim vIsCancel As String
Dim vHoardAmount As Currency
Dim vTransportLocation As String
Dim vMapPart As String
Dim vDistance As Currency
Dim vMydescription As String
Dim vPlaceID As Double
Dim vRouteID As Double
Dim vReceiveID As Double
Dim vRefID As Double
Dim vRefType As String
Dim vPriorityID  As Double
Dim vCheckPriorityID As String
Dim vCheckRefType As String
Dim i As Integer
Dim vErrorHeader As String
Dim vSaveHeader_Status As Integer
Dim vSaveDetail_Status As Integer
Dim vReturn_Status As Integer
Dim vIsCompleteSave As Integer
Dim vMydescriptionSub As String
Dim vIsCancelSub As String
Dim vRefIDSub As Double
Dim vRefSubID As Double
Dim vQueueID As Double
Dim vLineNumber As Double
Dim vReceiveName As String
Dim vReceiveTelHome As String
Dim vReceiveTelMobile As String
Dim vAnswer As Integer
    
If ListView101.ListItems.Count <> 0 Then
    If DTPicker101.Value > DTPicker102.Value Then
        MsgBox "วันที่นัดรับสินค้า ต้องมากกว่าหรือเท่ากับวันที่ของเอกสาร"
        Exit Sub
    End If
    
    'Check Priority
    vCheckPriorityID = Trim(CMB102.Text)
    vQuery = "select ID from npmaster.dbo.TB_DO_Priority where Priority = '" & vCheckPriorityID & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPriorityID = Trim(vRecordset.Fields("id").Value)
    Else
        MsgBox "ระดับความสำคัญ ไม่มีในระบบ กรุณาตรวจสอบ", vbCritical, "Send Error"
        CMB102.SetFocus
        Exit Sub
    End If
    vRecordset.Close
        
    If Text102.Text <> "" And Text103.Text <> "" And CMB102.Text <> "" And Text106.Text <> "" And Text201.Text <> "" And Text301.Text <> "" And Text306.Text <> "" Then
        Dim vDoDay As Date
        Dim vDueDay As Date
        
        vDoDay = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
        vDueDay = Trim(DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year)
        
        If vDoDay > vDueDay Then
            MsgBox "วันที่นัดรับสินค้า ต้องเป็นวันเดียวกันกับวันที่เอกสาร หรือ ต้องมากกว่าวันที่เอกสาร", vbCritical, "Send Error"
            Exit Sub
        End If
        
        Call CheckIsCancel
        If vCheckIsCancel1 = 1 Or vCheckIsConfirm1 = 1 Then
            If vCheckIsCancel1 = 1 Then
                MsgBox "เอกสารถูกยกเลิกแล้วไม่สามารถแก้ไขบันทึกข้อมูลได้", vbCritical, "Send Message"
            ElseIf vCheckIsConfirm1 = 1 Then
                MsgBox "เอกสารถูกอนุมัติแล้วไม่สามารถแก้ไขบันทึกข้อมูลได้", vbCritical, "Send Message"
            End If
            Exit Sub
        End If
        If vIsOpen1 = 0 Then
            Call CheckDocNo
            vID = "Null"
        Else
            vID = Trim(Text101.Text)
        End If
        vDocNo = Trim(Text102.Text)
        vDocdate = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
        vDueDate = Trim(DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year)
        vDueTime = Trim(MaskEdBox101.Text)
        vIsCancel = 0
        vHoardAmount = Trim(Text103.Text)
        vTransportLocation = Trim(Text106.Text)
        vMapPart = Trim(Text107.Text)
        vDistance = 0
        vMydescription = Trim(Text108.Text)
        vPlaceID = Trim(Text301.Text)
        vRouteID = Trim(Text306.Text)
        vReceiveName = Trim(Text201.Text)
        vReceiveTelHome = Trim(Text202.Text)
        vReceiveTelMobile = Trim(Text203.Text)
        vRefID = Trim(ListView101.ListItems.Item(1).SubItems(7))
        vCheckRefType = Trim(ListView101.ListItems.Item(1).SubItems(1))
        'Check DocType-------------------------------------------------------------------------------------------------------------
        If vIsOpen1 = 0 Then
            'If CMB101.Text = Trim("ใบสั่งขาย/จอง") And Check101.Value = 0 Then
                vRefType = 1
            'Else
             '   vRefType = 2
            'End If
        Else
        vQuery = "select reftype from npmaster.dbo.tb_do_queue where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRefType = Trim(vRecordset.Fields("reftype").Value)
        End If
        vRecordset.Close
        End If
        
        'Insert Header
        'On Error GoTo RollbackHeader
        vQuery = "exec bcnp.dbo.USP_DO_QueueUpdate_Output " & vID & ",'" & vDocNo & "'," _
                            & " '" & vDocdate & "','" & vDueDate & "','" & vDueTime & "','" & vIsCancel & "'," _
                            & " " & vHoardAmount & ",'" & vTransportLocation & "','" & vMapPart & "'," & vDistance & "," _
                            & " '" & vMydescription & "'," & vPlaceID & "," & vRouteID & "," _
                            & " " & vRefID & ",'" & vRefType & "','" & vPriorityID & "','" & vUserID & "' ," _
                            & " '" & vReceiveName & "','" & vReceiveTelHome & "','" & vReceiveTelMobile & "' "
        gConnection.Execute vQuery
        
        vQuery = "select * from npmaster.dbo.tb_do_queue where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vQueueID = Trim(vRecordset.Fields("id").Value)
        End If
        vRecordset.Close
        
RollbackHeader:
        If Err.Description <> "" Then
        vRefIDSub = Trim(ListView101.ListItems.Item(1).SubItems(7))
        vRefSubID = Trim(ListView101.ListItems.Item(1).SubItems(8))
        vQuery = "exec bcnp.dbo.USP_DO_QueueSubUpdate_Output 1,0,'',0," & vRefIDSub & "," & vRefSubID & "," & vQueueID & ",0"
        gConnection.Execute vQuery
        MsgBox Err.Description, vbCritical, "Send Error"
        Exit Sub
        End If
'Insert Sub
        For i = 1 To ListView101.ListItems.Count
        vReturn_Status = 0
        If i < ListView101.ListItems.Count Then
            vIsCompleteSave = 0
        Else
            vIsCompleteSave = 1
        End If
        vMydescriptionSub = ""
        vIsCancelSub = Trim(ListView101.ListItems.Item(i).SubItems(9))
        vRefIDSub = Trim(ListView101.ListItems.Item(i).SubItems(7))
        vRefSubID = Trim(ListView101.ListItems.Item(i).SubItems(8))
        vLineNumber = i - 1
        
        On Error GoTo RollbackDetails
        vQuery = "exec bcnp.dbo.USP_DO_QueueSubUpdate_Output " & vReturn_Status & "," & vIsCompleteSave & ",'" & vMydescriptionSub & "','" & vIsCancelSub & "'," & vRefID & "," & vRefSubID & "," & vQueueID & "," & vLineNumber & " "
        gConnection.Execute vQuery
        If i = 1 Then
        vQuery = "exec bcnp.dbo.USP_DO_QueueSubUpdate_Output " & vReturn_Status & "," & vIsCompleteSave & ",'" & vMydescriptionSub & "','" & vIsCancelSub & "'," & vRefID & "," & vRefSubID & "," & vQueueID & "," & vLineNumber & " "
        gConnection.Execute vQuery
        End If
        If vReturn_Status = 1 Then
        Exit Sub
        End If
        
RollbackDetails:
        If Err.Description <> "" Then
        vRefIDSub = Trim(ListView101.ListItems.Item(i).SubItems(7))
        vRefSubID = Trim(ListView101.ListItems.Item(i).SubItems(8))
        vQuery = "exec bcnp.dbo.USP_DO_QueueSubUpdate_Output 1,0,'',0," & vRefIDSub & "," & vRefSubID & "," & vQueueID & ",0"
        gConnection.Execute vQuery
        MsgBox Err.Description, vbCritical, "Send Error"
        Exit Sub
        End If
        
        Next i
    'On Error GoTo ErrDescription
    If vIsOpen1 = 0 Then
        MsgBox "เอกสารได้ทำการบันทึกเรียบร้อยแล้ว ได้เอกสารเลขที่ " & vDocNo & " ", vbInformation, "Send Message"
        vAnswer = MsgBox("ต้องการพิมพ์เอกสารหรือไม่", vbYesNo, "Message Question")
        If vAnswer = 6 Then
            Call CMD005_Click
        End If
    Else
        'Call CalRemainQueue
        MsgBox "เอกสารเลขที่ " & vDocNo & "  ได้ทำการอัพเดทเรียบร้อยแล้ว ", vbInformation, "Send Message"
        vAnswer = MsgBox("ต้องการพิมพ์เอกสารหรือไม่", vbYesNo, "Message Question")
        If vAnswer = 6 Then
            Call CMD005_Click
        End If
    End If
    vIsOpen1 = 0
    Check101.Value = 0
    ListView101.ListItems.Clear
    DTPicker101.Value = Now
    DTPicker102.Value = Now
    CMD006.Enabled = False
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = "0"
    MaskEdBox101.Mask = "##:##"
If Check102.Value = 0 Then
    CMB102.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text107.Text = ""
    Text108.Text = ""
    Text201.Text = ""
    Text202.Text = ""
    Text203.Text = ""
    Text301.Text = ""
    Text302.Text = ""
    Text303.Text = ""
    Text304.Text = ""
    Text305.Text = ""
    Text306.Text = ""
    Text307.Text = ""
    Text308.Text = ""
    Text309.Text = ""
    End If
    Form312.Image101.Visible = True
    Form312.Image102.Visible = False
    Form312.Image103.Visible = False
    Else
    MsgBox "กรุณา กรอกข้อมูลให้ครบตามช่องที่มีตัวหนังสือสีแดง เพราะถ้าไม่ครบจะไม่สามรถบันทึกข้อมูลได้", vbCritical, "Send Information"
    End If
    
'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description, vbCritical, "Send Error"
  '  Exit Sub
'End If
Else
    MsgBox "ไม่สามารถบันทึกข้อมูลได้ เนื่องจากข้อมูลไม่ครบ", vbInformation, "Send Message"
End If

End Sub

Private Sub CMD004_Click()
Dim vAnswer As Integer

'On Error GoTo ErrDescription

If vIsOpen1 = 0 And ListView101.ListItems.Count > 0 Then
    vAnswer = MsgBox("การเปิดเอกสารเก่าขึ้นมาดู จะทำให้เอกสารใหม่ที่กำลังสร้างหายไป คุณต้องการบันทึกก่อนหรือไม่", vbYesNo, "Question Respond")
    If vAnswer = 7 Then
        FrmOrder006.Show
        'MDIFrmProgramPrint.Order101.Enabled = False
        Form312.Enabled = False
        'Form311.Enabled = False
    Else
        Exit Sub
    End If
End If
        
'MDIFrmProgramPrint.Order101.Enabled = False
FrmOrder006.Show
Form312.Enabled = False
'Form311.Enabled = False

'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description, vbCritical, "Send Error"
  '  Exit Sub
'End If
End Sub

Private Sub CMD005_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vDocNo As String
Dim vRefType As Integer
Dim vRepID As Integer
Dim vRepType As String

If Text102.Text <> "" Then
    vDocNo = Trim(Text102.Text)
    If vIsOpen1 = 1 Then
    vQuery = "select reftype from npmaster.dbo.tb_do_queue where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRefType = Trim(vRecordset.Fields("reftype").Value)
    End If
    vRecordset.Close
    End If
    
    If (UCase(vDepartment) = "CH" Or UCase(vDepartment) = "CR" Or UCase(vDepartment) = "IS" Or UCase(vDepartment) = "PC") And Me.CHKReqPrint.Value = 0 Then
       vRepID = 404
    Else
       vRepID = 291
    End If
    
    vRepType = "DO"
    
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
    
        With Crystal101
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@QRDocno;" & vDocNo & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
        End With
        vIsOpen1 = 0
End If
Me.CHKReqPrint.Value = 0
End Sub

Private Sub CMD006_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vID As Integer
Dim vRefSubID As Double
Dim i As Integer
Dim vAnswer As Integer

On Error GoTo ErrDescription

If Text101.Text <> "" And Text102.Text <> "" And ListView101.ListItems.Count > 0 And vIsOpen1 = 1 Then
    vID = Trim(Text101.Text)
    vDocNo = Trim(Text102.Text)
    vAnswer = MsgBox("คุณต้องการยกเลิกเอกสารเลขที่ " & vDocNo & " ใช่หรือไม่", vbYesNo, "Question Respond")
    If vAnswer = 6 Then
        Call CheckIsCancel
            If (vCheckIsCancel1 = 0 And vCheckIsConfirm1 = 0) Then
            vQuery = "exec bcnp.dbo.usp_DO_CancelDocumentHeader '" & vDocNo & "'," & vID & ",'" & vUserID & "' "
            gConnection.Execute vQuery
            
            For i = 1 To ListView101.ListItems.Count
                vRefSubID = Trim(ListView101.ListItems.Item(i).SubItems(8))
                vQuery = "exec bcnp.dbo.usp_DO_CancelDocumentDetails     " & vID & "," & vRefSubID & " "
                gConnection.Execute vQuery
            Next i
            MsgBox "ทำการยกเลิกเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้วครับ", vbInformation, "Send Message "
            vIsOpen1 = 0
            Check101.Value = 0
            ListView101.ListItems.Clear
            DTPicker101.Value = Now
            DTPicker102.Value = Now
            CMD006.Enabled = False
            Text101.Text = ""
            Text102.Text = ""
            Text103.Text = "0"
            CMB102.Text = ""
            Text105.Text = ""
            Text106.Text = ""
            Text107.Text = ""
            Text108.Text = ""
            MaskEdBox101.Mask = "##:##"
            Text201.Text = ""
            Text202.Text = ""
            Text203.Text = ""
            Text301.Text = ""
            Text302.Text = ""
            Text303.Text = ""
            Text304.Text = ""
            Text305.Text = ""
            Text306.Text = ""
            Text307.Text = ""
            Text308.Text = ""
            Text309.Text = ""
            Form312.Image101.Visible = True
            Form312.Image102.Visible = False
            Form312.Image103.Visible = False
        Else
            MsgBox "เอกสารไม่สามารถยกเลิกข้อมูลได้", vbCritical, "Send Message"
            Exit Sub
        End If
    Else
        Exit Sub
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD007_Click()
Unload Form312
End Sub

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vGenDocNo As String

On Error GoTo ErrDescription

If ListView101.ListItems.Count > 0 And vIsOpen1 = 0 Then
    vQuery = "exec bcnp.dbo.USP_DO_QueueNewDocNo"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vGenDocNo = Trim(vRecordset.Fields("QueueNewDocNo").Value)
    End If
    vRecordset.Close
    Text102.Text = vGenDocNo
    DTPicker101.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD201_Click()
On Error GoTo ErrDescription

MDIFrmProgramPrint.Enabled = False
FrmOrder007.Show
vReceiveModule = 1

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD301_Click()
On Error GoTo ErrDescription

MDIFrmProgramPrint.Enabled = False
FrmOrder008.Show
vPlaceModule = 1

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD302_Click()
On Error GoTo ErrDescription

MDIFrmProgramPrint.Enabled = False
vRouteModule = 1
FrmOrder009.Show

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub DTPicker101_CloseUp()
DTPicker102.SetFocus
End Sub

Private Sub DTPicker102_Change()
On Error GoTo ErrDescription

If DTPicker102.Value < DTPicker101.Value Then
    MsgBox "วันที่นัดรับสินค้า ต้องเป็นวันเดียวกันกับวันที่เอกสาร หรือ ต้องมากกว่าวันที่เอกสาร", vbCritical, "Send Error"
    DTPicker102.Value = DTPicker101.Value
    Exit Sub
End If

MaskEdBox101.SetFocus

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub DTPicker102_CloseUp()
MaskEdBox101.SetFocus
End Sub

Private Sub DTPicker102_GotFocus()
Dim vDate1 As Date
Dim vDate2 As Date

    vDate1 = DTPicker101.Value
    vDate2 = DTPicker102.Value
    If vDate2 < vDate1 Then
        MsgBox "วันที่นัดรับสินค้า ต้องมากกว่าหรือเท่ากับวันที่ของเอกสาร"
        DTPicker102.Value = DTPicker101.Value
        Exit Sub
        
    End If
End Sub

Private Sub DTPicker102_KeyPress(KeyAscii As Integer)
On Error GoTo ErrDescription

If KeyAscii = 13 Then
    MaskEdBox101.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearchList As ListItem
Dim i As Integer
Dim vIsPOS As Integer

On Error GoTo ErrDescription

CMB101.AddItem Trim("ใบสั่งขาย/จอง")
'CMB101.AddItem Trim("บิลขาย")
CMB101.Text = Trim("ใบสั่งขาย/จอง")
DTPicker101.Value = Now
DTPicker102.Value = Now
Form312.Image101.Visible = True
Form312.Image102.Visible = False
Form312.Image103.Visible = False
vIsOpen1 = 0
Text103.Text = 0
'MaskEdBox101.Text = "00:00"

vQuery = "select priority from npmaster.dbo.TB_DO_Priority order by id"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMB102.AddItem Trim(vRecordset.Fields("priority").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
CMB102.ListIndex = 0

If vTempDocno <> "" Then
    i = 1
    vIsPOS = 0
    Form312.ListView101.ListItems.Clear
    vQuery = "exec bcnp.dbo.usp_do_searchrefheader " & vIsPOS & ",'" & vTempDocno & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        Form312.Text103.Text = Trim(vRecordset.Fields("hoardamount").Value)
    End If
    vRecordset.Close
    vQuery = "exec bcnp.dbo.USP_DO_SearchRef " & vIsPOS & ",'" & vTempDocno & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set vSearchList = Form312.ListView101.ListItems.Add(, , Trim(i))
            vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
            vSearchList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
            vSearchList.SubItems(3) = Trim(vRecordset.Fields("itemname").Value)
            vSearchList.SubItems(4) = Trim(vRecordset.Fields("qty").Value)
            vSearchList.SubItems(5) = Trim(vRecordset.Fields("doremainqty").Value)
            vSearchList.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
            vSearchList.SubItems(7) = Trim(vRecordset.Fields("headid").Value)
            vSearchList.SubItems(8) = Trim(vRecordset.Fields("detailid").Value)
            vSearchList.SubItems(9) = Trim(0)
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Public Sub CheckDocNo()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocNo As String
Dim vGenDocNo As String

vCheckDocNo = Trim(Text102.Text)
vQuery = "select docno,iscancel from npmaster.dbo.tb_do_queue where docno = '" & vCheckDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDocnoExist1 = 1
    vCheckIsCancel1 = Trim(vRecordset.Fields("iscancel").Value)
Else
    vCheckDocnoExist1 = 0
    vCheckIsCancel1 = 0
End If
vRecordset.Close

If vCheckDocnoExist1 = 1 Then
    vQuery = "exec bcnp.dbo.USP_DO_QueueNewDocNo"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vGenDocNo = Trim(vRecordset.Fields("QueueNewDocNo").Value)
    End If
    vRecordset.Close
    Text102.Text = vGenDocNo
End If
End Sub

Public Sub CheckIsCancel()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocNo As String
Dim vGenDocNo As String

vCheckDocNo = Trim(Text102.Text)
vQuery = "select docno,iscancel,isconfirm  from npmaster.dbo.tb_do_queue where docno = '" & vCheckDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckIsCancel1 = Trim(vRecordset.Fields("iscancel").Value)
    vCheckIsConfirm1 = Trim(vRecordset.Fields("isconfirm").Value)
Else
    vCheckIsCancel1 = 0
    vCheckIsConfirm1 = 0
End If
vRecordset.Close

End Sub

Public Sub CalRemainQueue()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocNo As String

vCheckDocNo = Trim(Text102.Text)
vQuery = "exec bcnp.dbo.USP_DO_QueueRemainCal '" & vCheckDocNo & "' "
gConnection.Execute vQuery
End Sub

Public Sub CheckInteger()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckString As String
Dim vString As String
Dim i As Integer

If vCheckTextBox = 1 Then
    vString = Trim(Text103.Text)
ElseIf vCheckTextBox = 2 Then
    vString = Trim(Text105.Text)
End If
For i = 1 To Len(vString)
    vCheckString = Mid(vString, i, 1)
    vQuery = "select '" & vCheckString & "' as Number where '" & vCheckString & "' in ('','.','0','1','2','3','4','5','6','7','8','9')"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vExistNumber = 1
    Else
        vExistNumber = 0
        Exit Sub
    End If
    vRecordset.Close
Next i
End Sub

Private Sub ListView101_DblClick()
Dim vAnswer As Integer
Dim vItemCode As String
Dim i As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count > 0 Then
Call CheckIsCancel
If vCheckIsCancel1 = 0 And vCheckIsConfirm1 = 0 Then
 i = ListView101.SelectedItem.Index
vItemCode = ListView101.ListItems.Item(i).SubItems(2)
vAnswer = MsgBox("ต้องการยกเลิกรายการสินค้า " & vItemCode & " นี้ใช่หรือไม่", vbYesNo, "Question")
    If vAnswer = 6 Then
        ListView101.ListItems.Item(i).SubItems(9) = 1
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
    ElseIf vAnswer = 7 Then
        ListView101.ListItems.Item(i).SubItems(9) = 0
        ListView101.ListItems(i).ForeColor = "&H00000000"
        ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H00000000"
        ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H00000000"
        ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H00000000"
        ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H00000000"
        ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H00000000"
        ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H00000000"
        ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H00000000"
        ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H00000000"
        ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H00000000"
        Exit Sub
    End If
Else
    MsgBox "เอกสารที่ถูกยกเลิกหรือถูกอ้างอิงไม่สามารถแก้ไขได้", vbCritical, "Send Message"
End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub MaskEdBox101_KeyPress(KeyAscii As Integer)
On Error GoTo ErrDescription

If KeyAscii = 13 Then
    Text103.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub MaskEdBox101_LostFocus()
Dim vTime1 As String
Dim vTime2 As String
Dim vTime3 As String
Dim vTime4 As String
Dim vTime5 As String
Dim vTime6 As String

If Text102.Text <> "" Or ListView101.ListItems.Count <> 0 Then
    vTime1 = Left(MaskEdBox101.Text, 1)
    vTime2 = Right(Left(MaskEdBox101.Text, 2), 1)
    vTime3 = Right(MaskEdBox101.Text, 1)
    vTime4 = Left(Right(MaskEdBox101.Text, 2), 1)
    vTime5 = Left(MaskEdBox101.Text, 2)
    vTime6 = Right(MaskEdBox101.Text, 2)
    
    If vTime1 = "_" Or vTime2 = "_" Or vTime3 = "_" Or vTime4 = "_" Then
        MsgBox "รูปแบบเวลาที่ใช้ในโปรแกรม คือ '00:00' ต้องใส่ให้ครบด้วยครับ", vbInformation, "Send Information"
        MaskEdBox101.SetFocus
        Exit Sub
    End If
    
    If vTime5 > 24 Or vTime6 > 59 Then
        MsgBox "รูปแบบเวลาที่ใช้ในโปรแกรม คือ ชั่วโมงต้องใส่ไม่เกิน 24 และ นาทีต้องไม่เกิน 60", vbInformation, "Send Information"
        MaskEdBox101.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub Text103_KeyPress(KeyAscii As Integer)
On Error GoTo ErrDescription

If KeyAscii = 13 Then
    'CMD104.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Text103_LostFocus()
On Error GoTo ErrDescription

vCheckTextBox = 1
Call CheckInteger
If vExistNumber = 0 And Text103.Text <> "" Then
    MsgBox "ช่องจำนวนเงินนี้กรอกข้อมูลได้เฉพาะตัวเลขเท่านั้นครับ กรุณาตรวจสอบ", vbCritical, "Send Error"
    Text103.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Text105_KeyPress(KeyAscii As Integer)
On Error GoTo ErrDescription

If KeyAscii = 13 Then
    Text106.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Text105_LostFocus()
On Error GoTo ErrDescription

vCheckTextBox = 2
Call CheckInteger
If vExistNumber = 0 And Text105.Text <> "" Then
    MsgBox "ช่องจำนวนเงินนี้กรอกข้อมูลได้เฉพาะตัวเลขเท่านั้นครับ กรุณาตรวจสอบ", vbCritical, "Send Error"
    Text105.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Text201_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text202.SetFocus
End If
End Sub
Private Sub Text202_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text203.SetFocus
End If
End Sub
