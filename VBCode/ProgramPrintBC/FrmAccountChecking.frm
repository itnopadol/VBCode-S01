VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmAccountChecking 
   Caption         =   "ตรวจความถูกต้องเอกสารทางบัญชี"
   ClientHeight    =   11010
   ClientLeft      =   3120
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmAccountChecking.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame FMPayBillReport 
      BackColor       =   &H00E0E0E0&
      Caption         =   "รายงาน ติดตามการวางบิล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10950
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   15225
      Begin VB.ComboBox CMBRoute 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2970
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   6480
         Width           =   7665
      End
      Begin VB.ComboBox CMBPressMen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2970
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   5805
         Width           =   7665
      End
      Begin VB.CommandButton CMDPayBillClose 
         Caption         =   "ออกจากรายงาน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10800
         TabIndex        =   45
         Top             =   7605
         Width           =   2085
      End
      Begin Crystal.CrystalReport Crystal101 
         Left            =   405
         Top             =   7380
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.CommandButton CMDPrintPayBill 
         Caption         =   "พิมพ์รายงาน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8595
         TabIndex        =   44
         Top             =   7605
         Width           =   2085
      End
      Begin VB.ComboBox CMBKeepMoney 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2970
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   5130
         Width           =   7665
      End
      Begin VB.ComboBox CMBPayBillType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2970
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   4455
         Width           =   7665
      End
      Begin VB.ComboBox CMBAr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2970
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   3780
         Width           =   7665
      End
      Begin VB.ComboBox CMBPayBillReportType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmAccountChecking.frx":9673
         Left            =   2970
         List            =   "FrmAccountChecking.frx":9675
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2340
         Width           =   7710
      End
      Begin MSComCtl2.DTPicker DTPDateStop 
         Height          =   420
         Left            =   6210
         TabIndex        =   38
         Top             =   3105
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65798145
         CurrentDate     =   41458
      End
      Begin MSComCtl2.DTPicker DTPDateStart 
         Height          =   375
         Left            =   2970
         TabIndex        =   36
         Top             =   3105
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65798145
         CurrentDate     =   41458
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือก สายการวางบิล :"
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
         Left            =   855
         TabIndex        =   74
         Top             =   6480
         Width           =   1950
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือก พนักงานเร่งรัด :"
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
         Left            =   765
         TabIndex        =   73
         Top             =   5805
         Width           =   2040
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือก พนักงานติดตามหนี้สิน :"
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
         Left            =   225
         TabIndex        =   47
         Top             =   5130
         Width           =   2580
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกประเภทการวางบิล :"
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
         Left            =   765
         TabIndex        =   46
         Top             =   4455
         Width           =   2040
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกลูกหนี้ :"
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
         Left            =   1755
         TabIndex        =   37
         Top             =   3780
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกประเภทรายงาน :"
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
         Left            =   900
         TabIndex        =   35
         Top             =   2340
         Width           =   1905
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่ :"
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
         Left            =   5130
         TabIndex        =   33
         Top             =   3105
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "จากวันที่ :"
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
         Left            =   1620
         TabIndex        =   32
         Top             =   3150
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "เลือก เงื่อนไขการดูรายงาน"
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
         Left            =   2925
         TabIndex        =   31
         Top             =   990
         Width           =   7710
      End
   End
   Begin VB.Frame FMDocReport 
      BackColor       =   &H00E0E0E0&
      Caption         =   "รายงาน ตรวจสอบความครบถ้วนของเอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10995
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   15225
      Begin VB.ComboBox CMBReportType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3645
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   3105
         Width           =   7215
      End
      Begin VB.CheckBox CKSealectAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "พิมพ์ทุกหัวเอกสาร เลือกหัวเอกสาร :"
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
         Left            =   360
         TabIndex        =   57
         Top             =   2475
         Width           =   3255
      End
      Begin VB.CommandButton CMDRPExit 
         Caption         =   "ออกจากรายงาน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   8775
         TabIndex        =   67
         Top             =   7335
         Width           =   2085
      End
      Begin VB.CommandButton CMDRPPrint 
         Caption         =   "พิมพ์รายงาน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   6480
         TabIndex        =   66
         Top             =   7335
         Width           =   2085
      End
      Begin Crystal.CrystalReport Crystal102 
         Left            =   765
         Top             =   7695
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.ComboBox CMBRPPaymentStatus 
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
         Height          =   420
         Left            =   3645
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   5805
         Width           =   7215
      End
      Begin VB.ComboBox CMBRPCheckStatus 
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
         Height          =   420
         Left            =   3645
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   6480
         Width           =   7215
      End
      Begin VB.ComboBox CMBRPReturnStatus 
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
         Height          =   420
         Left            =   3645
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   5130
         Width           =   7215
      End
      Begin VB.ComboBox CMBRPInvoiceStatus 
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
         Height          =   420
         Left            =   3645
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   4455
         Width           =   7215
      End
      Begin MSComCtl2.DTPicker DTPRPDate2 
         Height          =   420
         Left            =   7020
         TabIndex        =   61
         Top             =   3780
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65798145
         CurrentDate     =   41459
      End
      Begin MSComCtl2.DTPicker DTPRPDate1 
         Height          =   420
         Left            =   3645
         TabIndex        =   60
         Top             =   3780
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65798145
         CurrentDate     =   41459
      End
      Begin VB.ComboBox CMBRPHeader 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3645
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2475
         Width           =   7215
      End
      Begin VB.ComboBox CMBRPDocType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3645
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   1845
         Width           =   7215
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกประเภท รายงาน :"
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
         Left            =   720
         TabIndex        =   70
         Top             =   3150
         Width           =   2760
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกเงื่อนไข ในการดูรายงาน"
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
         Left            =   3645
         TabIndex        =   69
         Top             =   990
         Width           =   5910
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่ :"
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
         Left            =   5940
         TabIndex        =   68
         Top             =   3870
         Width           =   915
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกสถานะภาษี หัก ณ ที่จ่าย :"
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
         Left            =   810
         TabIndex        =   55
         Top             =   5895
         Width           =   2670
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกสถานะการตรวจ :"
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
         Left            =   1305
         TabIndex        =   54
         Top             =   6525
         Width           =   2175
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกสถานะใบลดหนี้ :"
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
         Left            =   1395
         TabIndex        =   53
         Top             =   5175
         Width           =   2085
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกสถานะใบเสร็จ :"
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
         Left            =   1395
         TabIndex        =   52
         Top             =   4500
         Width           =   2085
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "จากวันที่ :"
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
         Left            =   2295
         TabIndex        =   51
         Top             =   3870
         Width           =   1185
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกประเภทเอกสาร :"
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
         TabIndex        =   50
         Top             =   1890
         Width           =   2175
      End
   End
   Begin MSComCtl2.DTPicker DTPPlanPayDate 
      Height          =   330
      Left            =   4365
      TabIndex        =   72
      Top             =   2790
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   582
      _Version        =   393216
      Format          =   65798145
      CurrentDate     =   41667
   End
   Begin VB.CheckBox CBPlanPayDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "วางแผนวางบิลวันที่ :"
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
      Left            =   2295
      TabIndex        =   71
      Top             =   2790
      Width           =   1995
   End
   Begin VB.CommandButton CMDReport 
      Caption         =   "ดูรายงาน ตรวจสอบเอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2295
      TabIndex        =   48
      Top             =   3690
      Width           =   2490
   End
   Begin VB.CommandButton CMDReportPayBill 
      Caption         =   "ดูรายงาน การวางบิล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4860
      TabIndex        =   29
      Top             =   3690
      Width           =   2490
   End
   Begin VB.CheckBox CBRecMoney 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "วันที่นัดเก็บเงิน :"
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
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   11430
      TabIndex        =   28
      Top             =   2790
      Width           =   1680
   End
   Begin VB.CheckBox CBPayBillDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "วันที่วางบิล :"
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
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7830
      TabIndex        =   27
      Top             =   2790
      Width           =   1365
   End
   Begin VB.CommandButton CMDEditData 
      Caption         =   "ปรับข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13410
      TabIndex        =   26
      Top             =   3690
      Width           =   1770
   End
   Begin MSComCtl2.DTPicker DTPRecMoney 
      Height          =   330
      Left            =   13185
      TabIndex        =   25
      Top             =   2790
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   65798145
      CurrentDate     =   41445
   End
   Begin MSComCtl2.DTPicker DTPPayDate 
      Height          =   330
      Left            =   9270
      TabIndex        =   24
      Top             =   2790
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   65798145
      CurrentDate     =   41445
   End
   Begin VB.ComboBox CMBPayBill 
      Enabled         =   0   'False
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
      Left            =   2295
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   2295
      Width           =   5235
   End
   Begin VB.ComboBox CMBHeader 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10215
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   495
      Width           =   1680
   End
   Begin VB.ComboBox CMBCheckStatus 
      Enabled         =   0   'False
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
      Left            =   9270
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1800
      Width           =   1770
   End
   Begin MSComCtl2.DTPicker DTPDocDate 
      Height          =   375
      Left            =   13590
      TabIndex        =   17
      Top             =   495
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   65798145
      CurrentDate     =   41439
   End
   Begin VB.TextBox TXTMyDescription 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2295
      TabIndex        =   13
      Top             =   3240
      Width           =   12885
   End
   Begin MSComctlLib.ListView ListViewDocNo 
      Height          =   4785
      Left            =   90
      TabIndex        =   12
      Top             =   4770
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   8440
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
      NumItems        =   20
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "บาร์โค้ด"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันที่เอกสาร"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "อ้างถึง"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "พนักงาน"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "มูลค่า"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ผู้สร้างเอกสาร"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "หมายเหตุของเอกสาร"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "หมายเหตุการตรวจสอบ"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "สถานะใบเสร็จ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "สถานะใบลดหนี้"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "สถานะภาษีหัก ณ ที่จ่าย"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "สถานะตรวจสอบ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "เงื่อนไขวางบิล"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "แผนวันที่วางบิล"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   16
         Text            =   "วันที่วางบิล"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Text            =   "วันที่นัดเก็บเงิน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Iscancel"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.ComboBox CMBPaymentStatus 
      Enabled         =   0   'False
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
      Left            =   13185
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1800
      Width           =   1995
   End
   Begin VB.ComboBox CMBReturnStatus 
      Enabled         =   0   'False
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
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1800
      Width           =   1770
   End
   Begin VB.ComboBox CMBInvoiceStatus 
      Enabled         =   0   'False
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
      Left            =   2295
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1800
      Width           =   1770
   End
   Begin VB.TextBox TXTBarCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   5130
      TabIndex        =   3
      Top             =   945
      Width           =   6765
   End
   Begin VB.ComboBox CMBDocType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5130
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   495
      Width           =   3165
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เงื่อนไขการวางบิล :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   180
      TabIndex        =   22
      Top             =   2340
      Width           =   2040
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกหัวเอกสาร :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8595
      TabIndex        =   20
      Top             =   540
      Width           =   1545
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สถานะตรวจสอบ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7650
      TabIndex        =   18
      Top             =   1845
      Width           =   1545
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกวันที่เอกสาร :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11880
      TabIndex        =   16
      Top             =   540
      Width           =   1635
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Height          =   60
      Left            =   90
      TabIndex        =   15
      Top             =   4275
      Width           =   15090
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการเอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   90
      TabIndex        =   14
      Top             =   4455
      Width           =   2355
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สถานะภาษีหัก ณ ที่จ่าย :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   10890
      TabIndex        =   8
      Top             =   1845
      Width           =   2220
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สถานะใบลดหนี้ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4050
      TabIndex        =   7
      Top             =   1845
      Width           =   1590
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สถานะใบเสร็จ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   540
      TabIndex        =   6
      Top             =   1845
      Width           =   1680
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   90
      TabIndex        =   5
      Top             =   1665
      Width           =   15090
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ/ผลการติดตาม :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   45
      TabIndex        =   4
      Top             =   3285
      Width           =   2130
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3690
      TabIndex        =   2
      Top             =   1125
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกประเภทเอกสาร :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3060
      TabIndex        =   0
      Top             =   540
      Width           =   1995
   End
End
Attribute VB_Name = "FrmAccountChecking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CMBDocType_Click()
'If Me.CMBDocType.ListIndex <> 16 Then
 ' Me.CMBPayBill.Enabled = False
  'Me.CBPayBillDate.Enabled = False
  'Me.DTPPayDate.Enabled = False
  'Me.CBRecMoney.Enabled = False
  'Me.DTPRecMoney.Enabled = False
'Else
 ' Me.CMBPayBill.Enabled = True
  'Me.CBPayBillDate.Enabled = True
  'Me.DTPPayDate.Enabled = True
  'Me.CBRecMoney.Enabled = True
  'Me.DTPRecMoney.Enabled = True
'End If

Call vGetHeader
End Sub

Public Sub GetKeepMen()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

Me.CMBKeepMoney.Clear
vQuery = "select * from dbo.vw_NP_KeepMoneyMenName "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   Me.CMBKeepMoney.AddItem (vRecordset.Fields("keepmencodename").Value)
   vRecordset.MoveNext
   Wend
End If
vRecordset.Close

End Sub

Public Sub GetPressMen()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

Me.CMBPressMen.Clear
vQuery = "select * from dbo.vw_NP_PressMenName "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   Me.CMBPressMen.AddItem (vRecordset.Fields("pressmencodename").Value)
   vRecordset.MoveNext
   Wend
End If
vRecordset.Close

End Sub


Public Sub GetPayBillRoute()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

Me.CMBRoute.Clear
vQuery = "select * from dbo.vw_NP_PaybillRoute "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   Me.CMBRoute.AddItem (vRecordset.Fields("routestepname9").Value)
   vRecordset.MoveNext
   Wend
End If
vRecordset.Close

End Sub

Private Sub CMBHeader_Click()
Call InsertData
Call SaveData
End Sub


Private Sub CMBReportType_Click()
On Error Resume Next

If Me.CMBReportType.ListIndex = 0 Then
    Me.CMBRPInvoiceStatus.Enabled = False
    Me.CMBRPReturnStatus.Enabled = False
    Me.CMBRPPaymentStatus.Enabled = False
    Me.CMBRPCheckStatus.Enabled = False
ElseIf Me.CMBReportType.ListIndex = 1 Then
    Me.CMBRPInvoiceStatus.Enabled = False
    Me.CMBRPReturnStatus.Enabled = False
    Me.CMBRPPaymentStatus.Enabled = False
    Me.CMBRPCheckStatus.Enabled = False
ElseIf Me.CMBReportType.ListIndex = 2 Then
    Me.CMBRPInvoiceStatus.Enabled = True
    Me.CMBRPReturnStatus.Enabled = False
    Me.CMBRPPaymentStatus.Enabled = False
    Me.CMBRPCheckStatus.Enabled = False
ElseIf Me.CMBReportType.ListIndex = 3 Then
    Me.CMBRPInvoiceStatus.Enabled = False
    Me.CMBRPReturnStatus.Enabled = True
    Me.CMBRPPaymentStatus.Enabled = False
    Me.CMBRPCheckStatus.Enabled = False
ElseIf Me.CMBReportType.ListIndex = 4 Then
    Me.CMBRPInvoiceStatus.Enabled = False
    Me.CMBRPReturnStatus.Enabled = False
    Me.CMBRPPaymentStatus.Enabled = True
    Me.CMBRPCheckStatus.Enabled = False
ElseIf Me.CMBReportType.ListIndex = 5 Then
    Me.CMBRPInvoiceStatus.Enabled = False
    Me.CMBRPReturnStatus.Enabled = False
    Me.CMBRPPaymentStatus.Enabled = False
    Me.CMBRPCheckStatus.Enabled = True
Else
    Me.CMBRPInvoiceStatus.Enabled = False
    Me.CMBRPReturnStatus.Enabled = False
    Me.CMBRPPaymentStatus.Enabled = False
    Me.CMBRPCheckStatus.Enabled = False
End If
End Sub

Private Sub CMBRPDocType_Click()
Call vGetReportHeader
End Sub

Private Sub CMDEditData_Click()
Dim vQuery As String
Dim vDocType As Integer
Dim vHeader As String
Dim vDocdate As String
Dim i As Integer
Dim vBarCode As String
Dim vCheckDocNo As String
Dim vCount As Integer
Dim vInvoiceStatus As String
Dim vReturnStatus As String
Dim vCheckStatus As String
Dim vPayStatus As String
Dim vPayBillType As String
Dim vPaybillDate As String
Dim vRecMoney As String
Dim vMydescription As String
Dim vIndex As Integer
Dim vCheckDate As Date
Dim vIsPaybillDate As Integer
Dim vIsRecMoney As Integer
Dim vRecMoneyDate As String
Dim vIsPlanDate As Integer
Dim vPlanPayDate As String

Dim vARCode As String
Dim vPersonCode As String
Dim vNumberNo As String
Dim vRunNumber As String
Dim vDocNo As String
Dim vExistStatus As Integer
Dim vChecker As String
Dim vExist As Integer

Dim vSlipDescription As String
Dim vReturnDescription As String
Dim vTaxDescription As String
    
Dim vDate As String
Dim vMonth As String
Dim vAutoNumber As String
Dim vDay As String

On Error Resume Next

If Me.TXTBarCode.Text <> "" Then
    vCount = 0
    vIndex = 0
    vBarCode = UCase(Me.TXTBarCode.Text)
    
    
    If Me.CMBInvoiceStatus.Text <> "" Then
      vInvoiceStatus = Me.CMBInvoiceStatus.Text
    Else
      vInvoiceStatus = ""
    End If
    
    If Me.CMBReturnStatus.Text <> "" Then
      vReturnStatus = Me.CMBReturnStatus.Text
    Else
      vReturnStatus = ""
    End If
    
    If Me.CMBCheckStatus.Text <> "" Then
      vCheckStatus = Me.CMBCheckStatus.Text
    Else
      vCheckStatus = ""
    End If
    
    If Me.CMBPaymentStatus.Text <> "" Then
      vPayStatus = Me.CMBPaymentStatus.Text
    Else
     vPayStatus = ""
    End If
    
    If Me.CMBPayBill.Text <> "" Then
    vPayBillType = Me.CMBPayBill.Text
    Else
    vPayBillType = ""
    End If
        
    
    If Me.CBPlanPayDate.Value = 1 Then
    vPlanPayDate = Me.DTPPlanPayDate.Value
    vIsPlanDate = 1
    Else
    vIsPlanDate = 0
    vPlanPayDate = ""
    End If
        
    
    If Me.CBPayBillDate.Value = 1 Then
    vPaybillDate = Me.DTPPayDate.Value
    vIsPaybillDate = 1
    Else
    vIsPaybillDate = 0
    vPaybillDate = ""
    End If
    
    If Me.CBRecMoney.Value = 1 Then
    vIsRecMoney = 1
    vRecMoneyDate = Me.DTPRecMoney.Value
    Else
    vIsRecMoney = 0
    vRecMoneyDate = ""
    End If
        
    vMydescription = Me.TXTMyDescription.Text
    
    For i = 1 To Me.ListViewDocNo.ListItems.Count
    vCheckDocNo = ListViewDocNo.ListItems.Item(i).ListSubItems(1).Text
    
    If vBarCode = vCheckDocNo Then
        vCount = vCount + 1
        vIndex = i
        
        If vCount > 0 Then
          GoTo HaveDoc
        End If
        
    End If
    Next i
    
HaveDoc:
    
    If vCount = 0 Then
      MsgBox "ไม่เจอเลขที่เอกสาร " & vBarCode & " นี้ในระบบ  กรุณาตรวจสอบ", vbCritical, "Send Error Message"
    Else
    
    ListViewDocNo.ListItems.Item(vIndex).Checked = True
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(2).Text = vCheckDocNo
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(9).Text = vMydescription
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(10).Text = vInvoiceStatus
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(11).Text = vReturnStatus
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(12).Text = vPayStatus
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(13).Text = vCheckStatus
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(14).Text = vPayBillType
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(15).Text = vPlanPayDate
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(16).Text = vPaybillDate
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(17).Text = vRecMoneyDate
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(18).Text = 1
    
    
    
    vCheckDate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
    If Len(DTPDocDate.Day) = 1 Then
    vDay = Trim(0 & DTPDocDate.Day)
    Else
    vDay = Trim(DTPDocDate.Day)
    End If
    
    vMonth = Trim(DTPDocDate.Month)

    vDate = Right(DTPDocDate.Year, 2) + 43 & vMonth & vDay
    vDocType = CMBDocType.ListIndex
    vHeader = Trim(CMBHeader.Text)
    If CMBDocType.ListIndex <> 9 Then
    vNumberNo = vHeader & vDate
    Else
    vNumberNo = vHeader
    End If
    
    vRunNumber = vNumberNo
    vDocNo = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(1))
    vARCode = Left(Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(4)), InStr(Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(4)), "/") - 1)
    vPersonCode = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(5))
    vExistStatus = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(18))

    
    If Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(9)) <> "" Then
    vMydescription = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(9))
    Else
    vMydescription = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(9))
    End If
    
    vSlipDescription = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(10))
    vReturnDescription = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(11))
    vTaxDescription = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(12))
    

    vQuery = "exec dbo.USP_NP_InsertCheckDataExist '" & vRunNumber & "'," & vDocType & ",'" & vDocNo & "','" & vCheckDate & "'," & vExistStatus & ",'" & vMydescription & "','" & vUserID & "' ,'" & vARCode & "' ,'" & vPersonCode & "','" & vSlipDescription & "','" & vReturnDescription & "' ,'" & vTaxDescription & "', '" & vCheckStatus & "','" & vPayBillType & "'," & vIsPlanDate & ",'" & vPlanPayDate & "'," & vIsPaybillDate & ",'" & vPaybillDate & "'," & vIsRecMoney & ",'" & vRecMoneyDate & "' "
    gConnection.Execute vQuery
    
    
    ListViewDocNo.ListItems.Item(vIndex).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(1).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(2).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(3).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(4).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(5).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(6).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(7).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(8).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(9).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(10).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(11).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(12).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(13).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(14).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(15).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(16).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(17).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(18).ForeColor = "&H00C00000"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(19).ForeColor = "&H00C00000"
    
    End If
    
    Me.TXTBarCode.Text = ""
    Me.CMBInvoiceStatus.ListIndex = 0
    Me.CMBReturnStatus.ListIndex = 0
    Me.CMBPaymentStatus.ListIndex = 0
    Me.CMBCheckStatus.ListIndex = 0
    Me.TXTMyDescription.Text = ""
    Me.CBPayBillDate.Value = 0
    Me.CBRecMoney.Value = 0
    
    
    Me.CMBCheckStatus.Enabled = False
    Me.CMBInvoiceStatus.Enabled = False
    Me.CMBPaymentStatus.Enabled = False
    Me.CMBReturnStatus.Enabled = False
    Me.TXTMyDescription.Enabled = False
    
    Me.CMBPayBill.Enabled = False
    Me.CBPayBillDate.Enabled = False
    Me.DTPPayDate.Enabled = False
    Me.CBRecMoney.Enabled = False
    Me.DTPRecMoney.Enabled = False
    
    
    Me.TXTBarCode.SetFocus
End If
End Sub

Public Sub SaveData()
Dim vQuery As String
Dim vDocType As Integer
Dim vHeader As String
Dim vDocdate As String
Dim i As Integer
Dim vBarCode As String
Dim vCheckDocNo As String
Dim vCount As Integer
Dim vInvoiceStatus As String
Dim vReturnStatus As String
Dim vCheckStatus As String
Dim vPayStatus As String
Dim vPayBillType As String
Dim vPaybillDate As String
Dim vRecMoney As String
Dim vMydescription As String
Dim vIndex As Integer
Dim vCheckDate As Date
Dim vPlanPayDate As String
Dim vIsPaybillDate As Integer
Dim vIsRecMoney As Integer
Dim vRecMoneyDate As String
Dim vIsPlanDate As Integer

Dim vARCode As String
Dim vPersonCode As String
Dim vNumberNo As String
Dim vRunNumber As String
Dim vDocNo As String
Dim vExistStatus As Integer
Dim vChecker As String
Dim vExist As Integer

Dim vSlipDescription As String
Dim vReturnDescription As String
Dim vTaxDescription As String
    
Dim vDate As String
Dim vMonth As String
Dim vAutoNumber As String
Dim vDay As String


On Error Resume Next

vCheckDate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
If Len(DTPDocDate.Day) = 1 Then
vDay = Trim(0 & DTPDocDate.Day)
Else
vDay = Trim(DTPDocDate.Day)
End If

vMonth = Trim(DTPDocDate.Month)

vDate = Right(DTPDocDate.Year, 2) + 43 & vMonth & vDay

vDocType = CMBDocType.ListIndex
vHeader = Trim(CMBHeader.Text)
If CMBDocType.ListIndex <> 9 Then
vNumberNo = vHeader & vDate
Else
vNumberNo = vHeader
End If
    
For i = 1 To Me.ListViewDocNo.ListItems.Count

    vRunNumber = vNumberNo
    vDocNo = Trim(ListViewDocNo.ListItems.Item(i).SubItems(1))
    vARCode = Left(Trim(ListViewDocNo.ListItems.Item(i).SubItems(4)), InStr(Trim(ListViewDocNo.ListItems.Item(i).SubItems(4)), "/") - 1)
    vPersonCode = Trim(ListViewDocNo.ListItems.Item(i).SubItems(5))
    vExistStatus = Trim(ListViewDocNo.ListItems.Item(i).SubItems(18))
    
    
    If Trim(ListViewDocNo.ListItems.Item(i).SubItems(9)) <> "" Then
    vMydescription = Trim(ListViewDocNo.ListItems.Item(i).SubItems(9))
    Else
    vMydescription = Trim(ListViewDocNo.ListItems.Item(i).SubItems(9))
    End If
    
    vSlipDescription = Trim(ListViewDocNo.ListItems.Item(i).SubItems(10))
    vReturnDescription = Trim(ListViewDocNo.ListItems.Item(i).SubItems(11))
    vTaxDescription = Trim(ListViewDocNo.ListItems.Item(i).SubItems(12))
    
    vCheckStatus = ListViewDocNo.ListItems.Item(i).ListSubItems(13).Text
    vPayBillType = ListViewDocNo.ListItems.Item(i).ListSubItems(14).Text
    vPlanPayDate = ListViewDocNo.ListItems.Item(i).ListSubItems(15).Text
    vPaybillDate = ListViewDocNo.ListItems.Item(i).ListSubItems(16).Text
    vRecMoney = ListViewDocNo.ListItems.Item(i).ListSubItems(17).Text
    
    If vPlanPayDate <> "" Then
    vIsPlanDate = 1
    Else
    vIsPlanDate = 0
    End If
    
    If vPaybillDate <> "" Then
    vIsPaybillDate = 1
    Else
    vIsPaybillDate = 0
    End If
    
    If vRecMoney <> "" Then
    vIsRecMoney = 1
    Else
    vIsRecMoney = 0
    End If
    
    
    vQuery = "exec dbo.USP_NP_InsertCheckDataExist '" & vRunNumber & "'," & vDocType & ",'" & vDocNo & "','" & vCheckDate & "'," & vExistStatus & ",'" & vMydescription & "','" & vUserID & "' ,'" & vARCode & "' ,'" & vPersonCode & "','" & vSlipDescription & "','" & vReturnDescription & "' ,'" & vTaxDescription & "', '" & vCheckStatus & "','" & vPayBillType & "'," & vIsPlanDate & ",'" & vPlanPayDate & "'," & vIsPaybillDate & ",'" & vPaybillDate & "'," & vIsRecMoney & ",'" & vRecMoneyDate & "' "
    gConnection.Execute vQuery
      
      
Next i
    
Me.TXTBarCode.SetFocus

End Sub

Private Sub CMDPayBillClose_Click()
Me.FMPayBillReport.Visible = False
End Sub

Private Sub CMDPrintPayBill_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepType As String
Dim vRepID As Integer
Dim vReportName As String

Dim vType As Integer
Dim vBeginDate As String
Dim vEndDate As String
Dim vARCode As String
Dim vPayBillType As String
Dim vKeepMoneyCode As String
Dim vPressMenCode As String
Dim vRoute As String

On Error GoTo ErrDescription

If Me.CMBPayBillReportType.ListIndex = 0 Then
vRepID = 527
ElseIf Me.CMBPayBillReportType.ListIndex = 1 Then
vRepID = 521
ElseIf Me.CMBPayBillReportType.ListIndex = 2 Then
vRepID = 521
ElseIf Me.CMBPayBillReportType.ListIndex = 3 Then
vRepID = 521
ElseIf Me.CMBPayBillReportType.ListIndex = 4 Then
vRepID = 529
ElseIf Me.CMBPayBillReportType.ListIndex = 5 Then
vRepID = 530
ElseIf Me.CMBPayBillReportType.ListIndex = 6 Then
vRepID = 531
ElseIf Me.CMBPayBillReportType.ListIndex = 7 Then
vRepID = 532
ElseIf Me.CMBPayBillReportType.ListIndex = 8 Then
vRepID = 533
ElseIf Me.CMBPayBillReportType.ListIndex = 9 Then
vRepID = 534
ElseIf Me.CMBPayBillReportType.ListIndex = 10 Then
vRepID = 535
ElseIf Me.CMBPayBillReportType.ListIndex = 11 Then
vRepID = 536
End If
vRepType = "CK"

vBeginDate = Me.DTPDateStart.Day & "/" & Me.DTPDateStart.Month & "/" & Me.DTPDateStart.Year
vEndDate = Me.DTPDateStop.Day & "/" & Me.DTPDateStop.Month & "/" & Me.DTPDateStop.Year


vType = Me.CMBPayBillReportType.ListIndex


If Me.CMBAr.Text <> "" Then
  vARCode = Left(Me.CMBAr.Text, InStr(Me.CMBAr.Text, "/") - 1)
Else
  vARCode = ""
End If

If Me.CMBPayBillType.Text <> "" Then
  vPayBillType = Me.CMBPayBillType.Text
  
  If vPayBillType = "/-" Then
  vPayBillType = ""
  End If
Else
  vPayBillType = ""
End If


  If Me.CMBKeepMoney.Text <> "" Then
    vKeepMoneyCode = Left(Me.CMBKeepMoney.Text, InStr(Me.CMBKeepMoney.Text, "/") - 1)
  Else
    vKeepMoneyCode = ""
  End If

  If Me.CMBPressMen.Text <> "" Then
    vPressMenCode = Left(Me.CMBPressMen.Text, InStr(Me.CMBPressMen.Text, "/") - 1)
  Else
    vPressMenCode = ""
  End If

  If Me.CMBRoute.Text <> "" Then
    vRoute = Left(Me.CMBRoute.Text, InStr(Me.CMBRoute.Text, "/") - 1)
  Else
    vRoute = ""
  End If

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"

.ParameterFields(0) = "@vType;" & vType & ";true"
.ParameterFields(1) = "@vBeginDate;" & vBeginDate & ";true"
.ParameterFields(2) = "@vEndDate;" & vEndDate & ";true"
.ParameterFields(3) = "@vARCode;" & vARCode & ";true"
.ParameterFields(4) = "@vPayBillType;" & vPayBillType & ";true"
.ParameterFields(5) = "@vKeepMoneyCode;" & vKeepMoneyCode & ";true"
.ParameterFields(6) = "@vPressMenCode;" & vPressMenCode & ";true"
.ParameterFields(7) = "@vRoute;" & vRoute & ";true"
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

Private Sub CMDReport_Click()
Me.FMPayBillReport.Visible = False
Me.FMDocReport.Visible = True
End Sub

Private Sub CMDReportPayBill_Click()
Me.FMPayBillReport.Visible = True
Me.FMDocReport.Visible = False
End Sub

Private Sub CMDRPExit_Click()
Me.FMDocReport.Visible = False
End Sub

Private Sub CMDRPPrint_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vDocType As Integer
Dim vHeader As String
Dim vReportType As Integer
Dim vBeginDate As String
Dim vEndDate As String
Dim vRepID As Integer
Dim vRepType As String
Dim vInvoiceStatus As String
Dim vReturnStatus As String
Dim vTaxStatus As String
Dim vCheckStatus As String
Dim vSelectAll As Integer
Dim vTypeData As Integer

On Error GoTo ErrDescription

vRepID = 522
vRepType = "CK"

vBeginDate = Me.DTPRPDate1.Day & "/" & Me.DTPRPDate1.Month & "/" & Me.DTPRPDate1.Year
vEndDate = Me.DTPRPDate2.Day & "/" & Me.DTPRPDate2.Month & "/" & Me.DTPRPDate2.Year

If Me.CMBRPDocType.ListIndex = 0 Then
  vTypeData = 1
ElseIf Me.CMBRPDocType.ListIndex = 1 Then
  vTypeData = 2
ElseIf Me.CMBRPDocType.ListIndex = 2 Then
  vTypeData = 2
ElseIf Me.CMBRPDocType.ListIndex = 3 Then
  vTypeData = 0
ElseIf Me.CMBRPDocType.ListIndex = 4 Then
  vTypeData = 0
ElseIf Me.CMBRPDocType.ListIndex = 5 Then
  vTypeData = 1
ElseIf Me.CMBRPDocType.ListIndex = 6 Then
  vTypeData = 1
ElseIf Me.CMBRPDocType.ListIndex = 7 Then
  vTypeData = 1
ElseIf Me.CMBRPDocType.ListIndex = 8 Then
  vTypeData = 1
ElseIf Me.CMBRPDocType.ListIndex = 9 Then
  vTypeData = 0
ElseIf Me.CMBRPDocType.ListIndex = 10 Then
  vTypeData = 2
ElseIf Me.CMBRPDocType.ListIndex = 11 Then
  vTypeData = 2
ElseIf Me.CMBRPDocType.ListIndex = 12 Then
  vTypeData = 2
ElseIf Me.CMBRPDocType.ListIndex = 13 Then
  vTypeData = 2
ElseIf Me.CMBRPDocType.ListIndex = 14 Then
  vTypeData = 2
ElseIf Me.CMBRPDocType.ListIndex = 15 Then
  vTypeData = 1
ElseIf Me.CMBRPDocType.ListIndex = 16 Then
  vTypeData = 2
End If

If Me.CKSealectAll.Value = 1 Then
  vSelectAll = 1
Else
  vSelectAll = 0
End If

If Me.CMBReportType.ListIndex = 0 Then
  vReportType = Me.CMBReportType.ListIndex
  vDocType = Me.CMBRPDocType.ListIndex
  vHeader = Me.CMBRPHeader.Text
  
  vInvoiceStatus = ""
  vReturnStatus = ""
  vTaxStatus = ""
  vCheckStatus = ""

ElseIf Me.CMBReportType.ListIndex = 1 Then
  vReportType = Me.CMBReportType.ListIndex
  vDocType = Me.CMBRPDocType.ListIndex
  vHeader = Me.CMBRPHeader.Text
  
  vInvoiceStatus = ""
  vReturnStatus = ""
  vTaxStatus = ""
  vCheckStatus = ""

ElseIf Me.CMBReportType.ListIndex = 2 Then
  vReportType = Me.CMBReportType.ListIndex
  vDocType = Me.CMBRPDocType.ListIndex
  vHeader = Me.CMBRPHeader.Text
  
  vInvoiceStatus = Me.CMBRPInvoiceStatus.Text
  vReturnStatus = ""
  vTaxStatus = ""
  vCheckStatus = ""
  
ElseIf Me.CMBReportType.ListIndex = 3 Then
  vReportType = Me.CMBReportType.ListIndex
  vDocType = Me.CMBRPDocType.ListIndex
  vHeader = Me.CMBRPHeader.Text
  
  vInvoiceStatus = ""
  vReturnStatus = Me.CMBRPReturnStatus.Text
  vTaxStatus = ""
  vCheckStatus = ""
ElseIf Me.CMBReportType.ListIndex = 4 Then
  vReportType = Me.CMBReportType.ListIndex
  vDocType = Me.CMBRPDocType.ListIndex
  vHeader = Me.CMBRPHeader.Text
  
  vInvoiceStatus = ""
  vReturnStatus = ""
  vTaxStatus = Me.CMBRPPaymentStatus.Text
  vCheckStatus = ""
  
ElseIf Me.CMBReportType.ListIndex = 5 Then
  vReportType = Me.CMBReportType.ListIndex
  vDocType = Me.CMBRPDocType.ListIndex
  vHeader = Me.CMBRPHeader.Text
  
  vInvoiceStatus = ""
  vReturnStatus = ""
  vTaxStatus = ""
  vCheckStatus = Me.CMBRPCheckStatus.Text
End If

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal102
.ReportFileName = vReportName & ".rpt"

.ParameterFields(0) = "@vReportType;" & vReportType & ";true"
.ParameterFields(1) = "@vSelectAll;" & vSelectAll & ";true"
.ParameterFields(2) = "@vTypeData;" & vTypeData & ";true"
.ParameterFields(3) = "@vDocGroup;" & vDocType & ";true"
.ParameterFields(5) = "@vHeader;" & vHeader & ";true"
.ParameterFields(6) = "@vBeginDate;" & vBeginDate & ";true"
.ParameterFields(7) = "@vEndDate;" & vEndDate & ";true"
.ParameterFields(8) = "@vInvoiceStatus;" & vInvoiceStatus & ";true"
.ParameterFields(9) = "@vReturnStatus;" & vReturnStatus & ";true"
.ParameterFields(10) = "@vTaxStatus;" & vTaxStatus & ";true"
.ParameterFields(11) = "@vCheckStatus;" & vCheckStatus & ";true"
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

Private Sub DTPDocDate_Change()
Call InsertData
Call SaveData
End Sub

Private Sub Form_Load()

On Error Resume Next

DTPDocDate.Value = Now
CMBDocType.AddItem Trim("บิลขาย")
CMBDocType.AddItem Trim("ใบสั่งซื้อ")
CMBDocType.AddItem Trim("ใบรับเข้าสินค้า")
CMBDocType.AddItem Trim("ใบโอนย้ายสินค้า")
CMBDocType.AddItem Trim("ใบเบิกจ่ายสินค้า")
CMBDocType.AddItem Trim("ใบเสร็จรับชำระ")
CMBDocType.AddItem Trim("ใบมัดจำ")
CMBDocType.AddItem Trim("ใบลดหนี้")
CMBDocType.AddItem Trim("ใบเพิ่มหนี้")
CMBDocType.AddItem Trim("ใบจ่ายสินค้า")
CMBDocType.AddItem Trim("ใบปรับปรุงสินค้า")
CMBDocType.AddItem Trim("ใบสำคัญจ่ายเงิน")
CMBDocType.AddItem Trim("ใบสำคัญจ่ายเงินอื่นๆ")
CMBDocType.AddItem Trim("ใบจ่ายเงินล่วงหน้า")
CMBDocType.AddItem Trim("ใบส่งคืน/ลดหนี้")
CMBDocType.AddItem Trim("ใบรับวางบิลเจ้าหนี้")
CMBDocType.AddItem Trim("ใบรับวางบิลลูกหนี้")
CMBDocType.AddItem Trim("ใบบันทึกตั้งหนี้จากการซื้อ")

CMBRPDocType.AddItem Trim("บิลขาย")
CMBRPDocType.AddItem Trim("ใบสั่งซื้อ")
CMBRPDocType.AddItem Trim("ใบรับเข้าสินค้า")
CMBRPDocType.AddItem Trim("ใบโอนย้ายสินค้า")
CMBRPDocType.AddItem Trim("ใบเบิกจ่ายสินค้า")
CMBRPDocType.AddItem Trim("ใบเสร็จรับชำระ")
CMBRPDocType.AddItem Trim("ใบมัดจำ")
CMBRPDocType.AddItem Trim("ใบลดหนี้")
CMBRPDocType.AddItem Trim("ใบเพิ่มหนี้")
CMBRPDocType.AddItem Trim("ใบจ่ายสินค้า")
CMBRPDocType.AddItem Trim("ใบปรับปรุงสินค้า")
CMBRPDocType.AddItem Trim("ใบสำคัญจ่ายเงิน")
CMBRPDocType.AddItem Trim("ใบสำคัญจ่ายเงินอื่นๆ")
CMBRPDocType.AddItem Trim("ใบจ่ายเงินล่วงหน้า")
CMBRPDocType.AddItem Trim("ใบส่งคืน/ลดหนี้")
CMBRPDocType.AddItem Trim("ใบรับวางบิลเจ้าหนี้")
CMBRPDocType.AddItem Trim("ใบรับวางบิลลูกหนี้")
CMBRPDocType.AddItem Trim("ใบบันทึกตั้งหนี้จากการซื้อ")

Me.CMBInvoiceStatus.AddItem ("")
Me.CMBInvoiceStatus.AddItem ("รอใบเสร็จ")
Me.CMBInvoiceStatus.AddItem ("ได้ใบเสร็จแล้ว")

Me.CMBRPInvoiceStatus.AddItem ("")
Me.CMBRPInvoiceStatus.AddItem ("รอใบเสร็จ")
Me.CMBRPInvoiceStatus.AddItem ("ได้ใบเสร็จแล้ว")

Me.CMBReturnStatus.AddItem ("")
Me.CMBReturnStatus.AddItem ("รอใบลดหนี้")
Me.CMBReturnStatus.AddItem ("ได้ใบลดหนี้แล้ว")

Me.CMBRPReturnStatus.AddItem ("")
Me.CMBRPReturnStatus.AddItem ("รอใบลดหนี้")
Me.CMBRPReturnStatus.AddItem ("ได้ใบลดหนี้แล้ว")

Me.CMBPaymentStatus.AddItem ("")
Me.CMBPaymentStatus.AddItem ("รอภาษีหัก ณ ที่จ่าย")
Me.CMBPaymentStatus.AddItem ("ได้ภาษีหัก ณ ที่จ่าย แล้ว")

Me.CMBRPPaymentStatus.AddItem ("")
Me.CMBRPPaymentStatus.AddItem ("รอภาษีหัก ณ ที่จ่าย")
Me.CMBRPPaymentStatus.AddItem ("ได้ภาษีหัก ณ ที่จ่าย แล้ว")

Me.CMBCheckStatus.AddItem ("")
Me.CMBCheckStatus.AddItem ("ไม่ได้ตรวจสอบ")
Me.CMBCheckStatus.AddItem ("ตรวจสอบแล้ว")

Me.CMBRPCheckStatus.AddItem ("")
Me.CMBRPCheckStatus.AddItem ("ไม่ได้ตรวจสอบ")
Me.CMBRPCheckStatus.AddItem ("ตรวจสอบแล้ว")

Me.CMBReportType.AddItem ("รายงาน การตรวจสอบเอกสารประจำวัน")
Me.CMBReportType.AddItem ("รายงาน การตรวจสอบเอกสารไม่ครบประจำวัน")
Me.CMBReportType.AddItem ("รายงาน การตรวจสอบเอกสารประจำวัน ตามสถานะใบเสร็จ")
Me.CMBReportType.AddItem ("รายงาน การตรวจสอบเอกสารประจำวัน ตามสถานะใบลดหนี้")
Me.CMBReportType.AddItem ("รายงาน การตรวจสอบเอกสารประจำวัน ตามสถานะภาษีหัก_ณ_ที่จ่าย")
Me.CMBReportType.AddItem ("รายงาน การตรวจสอบเอกสารประจำวัน ตามสถานะการตรวจสอบ")
Me.CMBReportType.AddItem ("รายงาน สรุปการตรวจสอบเอกสารประจำวัน")

Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ตามช่วงวันที่")
Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ตามรหัสลูกหนี้")
Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ตามประเภทการวางบิล")
Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ตามพนักงานติดตามหนี้สิน")

Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ที่ยังไม่ได้ระบุวันที่วางบิล")
Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ตามวันที่วางบิล")
Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ที่ยังไม่ได้ระบุวันที่นัดเก็บเงิน")
Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ตามวันที่นัดเก็บเงิน")

Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ตามวันที่เอกสารและพนักงานเร่งรัด")
Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ตามวันที่เอกสารและสายการวางบิล")
Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ตามวันที่แผนวันวางบิล")
Me.CMBPayBillReportType.AddItem ("ดูรายงาน การวางบิล ตามวันที่เอกสารที่ไม่ได้ระบุแผนวันวางบิล")


Me.DTPDateStart.Value = Now
Me.DTPDateStop.Value = Now
Me.DTPRPDate1.Value = Now
Me.DTPRPDate2.Value = Now

Call GetRoute
Call SearchAR
Call GetKeepMen
Call GetPressMen
Call GetPayBillRoute

End Sub


Public Sub SearchAR()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

  vQuery = "exec dbo.USP_MP_SearchArCode 1,'' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  Me.CMBAr.Clear
  vRecordset.MoveFirst
  While Not vRecordset.EOF
  Me.CMBAr.AddItem (Trim(vRecordset.Fields("code").Value) + "/" + Trim(vRecordset.Fields("arname").Value))
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

Public Sub InsertData()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vHeader As String
Dim vDocdate As Date
Dim vListDocno As ListItem
Dim i As Integer

On Error Resume Next

If Me.CMBDocType.ListIndex = 0 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_ck_arinvoice '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("salecode").Value)
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netdebtamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("Description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMBDocType.ListIndex = 1 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_ck_purchaseorder '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("apcode").Value) & "/" & Trim(vRecordset.Fields("apname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMBDocType.ListIndex = 2 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_ck_apinvoice '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("apcode").Value) & "/" & Trim(vRecordset.Fields("apname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netdebtamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If
                    
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMBDocType.ListIndex = 3 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_ck_stktransfer '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = ""
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = ""
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMBDocType.ListIndex = 4 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_ck_stkissue'" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = ""
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = ""
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMBDocType.ListIndex = 5 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_ck_receipt '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    ElseIf CMBDocType.ListIndex = 6 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_ck_ardeposit'" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    ElseIf CMBDocType.ListIndex = 7 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_ck_arcreditnote '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netdebtamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    ElseIf CMBDocType.ListIndex = 8 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_ck_debitnote '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netdebtamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMBDocType.ListIndex = 9 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_CK_ReceiptSlip '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("salecode").Value)
                    vListDocno.SubItems(6) = ""
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("userprint").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(9) = ""
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                       vListDocno.SubItems(2) = ""
                       vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMBDocType.ListIndex = 10 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_ck_stkadjust '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = ""
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = ""
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                       vListDocno.SubItems(2) = ""
                       vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    
    ElseIf CMBDocType.ListIndex = 11 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_CK_Payment '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("apcode").Value) & "/" & Trim(vRecordset.Fields("apname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    
    ElseIf CMBDocType.ListIndex = 12 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_CK_OTHEREXPENSE '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("apcode").Value) & "/" & Trim(vRecordset.Fields("apname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    
    ElseIf CMBDocType.ListIndex = 13 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_CK_APDepositSpecial '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("apcode").Value) & "/" & Trim(vRecordset.Fields("apname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netdebtamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    
    ElseIf CMBDocType.ListIndex = 14 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_CK_STKRefund '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("apcode").Value) & "/" & Trim(vRecordset.Fields("apname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netdebtamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    
    ElseIf CMBDocType.ListIndex = 15 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_CK_BillStatementTemp '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("apcode").Value) & "/" & Trim(vRecordset.Fields("apname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netdebtamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
        End If
    ElseIf CMBDocType.ListIndex = 16 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_CK_PayBill '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = Trim(vRecordset.Fields("paybilltype").Value)
                    
                    If Trim(vRecordset.Fields("isplandate").Value) = 1 Then
                    vListDocno.SubItems(15) = Trim(vRecordset.Fields("planpaydate").Value)
                    Else
                    vListDocno.SubItems(15) = ""
                    End If
                    
                    If Trim(vRecordset.Fields("ispaybilldate").Value) = 1 Then
                    vListDocno.SubItems(16) = Trim(vRecordset.Fields("paybilldate").Value)
                    Else
                    vListDocno.SubItems(16) = ""
                    End If
                    
                    If Trim(vRecordset.Fields("isrecmoney").Value) = 1 Then
                    vListDocno.SubItems(17) = Trim(vRecordset.Fields("recmoneydate").Value)
                    Else
                    vListDocno.SubItems(17) = ""
                    End If
                    
                    
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    
    
    ElseIf CMBDocType.ListIndex = 17 Then
    If CMBHeader.Text <> "" Then
        i = 1
        ListViewDocNo.ListItems.Clear
        vHeader = Trim(CMBHeader.Text)
        vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
        vQuery = "exec dbo.usp_CK_APInvoice1 '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListViewDocNo.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = ""
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("apcode").Value) & "/" & Trim(vRecordset.Fields("apname").Value)
                    vListDocno.SubItems(5) = ""
                    vListDocno.SubItems(6) = Format(Trim(vRecordset.Fields("netdebtamount").Value), "##,##0.00")
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(9) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(10) = Trim(vRecordset.Fields("slipdescription").Value)
                    vListDocno.SubItems(11) = Trim(vRecordset.Fields("returndescription").Value)
                    vListDocno.SubItems(12) = Trim(vRecordset.Fields("taxdescription").Value)
                    vListDocno.SubItems(13) = Trim(vRecordset.Fields("checkstatus").Value)
                    vListDocno.SubItems(14) = ""
                    vListDocno.SubItems(15) = ""
                    vListDocno.SubItems(16) = ""
                    vListDocno.SubItems(17) = ""
                    vListDocno.SubItems(18) = Trim(vRecordset.Fields("existstatus").Value)
                    vListDocno.SubItems(19) = Trim(vRecordset.Fields("iscancel").Value)
                    
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
                        vListDocno.Checked = True
                    Else
                        vListDocno.SubItems(2) = ""
                        vListDocno.Checked = False
                    End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    
End If

    Dim a As Integer
    
    If Me.ListViewDocNo.ListItems.Count > 0 Then
        For a = 1 To Me.ListViewDocNo.ListItems.Count
        
                    If ListViewDocNo.ListItems.Item(a).SubItems(18) = 1 Then
                    ListViewDocNo.ListItems.Item(a).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(1).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(2).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(3).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(4).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(5).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(6).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(7).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(8).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(9).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(10).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(11).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(12).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(13).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(14).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(15).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(16).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(17).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(18).ForeColor = "&H000000FF"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(19).ForeColor = "&H000000FF"
                End If
        Next a
        
        
        For a = 1 To Me.ListViewDocNo.ListItems.Count
                    If ListViewDocNo.ListItems.Item(a).SubItems(18) = 1 Then
                    ListViewDocNo.ListItems.Item(a).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(1).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(2).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(3).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(4).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(5).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(6).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(7).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(8).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(9).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(10).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(11).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(12).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(13).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(14).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(15).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(16).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(17).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(18).ForeColor = "&H00C00000"
                    ListViewDocNo.ListItems.Item(a).ListSubItems(19).ForeColor = "&H00C00000"
                End If
        Next a
        
    End If
End Sub

Public Sub GetRoute()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocdate As Date

On Error Resume Next

Me.CMBPayBill.Clear
Me.CMBPayBillType.Clear

vQuery = "exec dbo.USP_CK_SelectRoute"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Me.CMBPayBill.AddItem Trim(vRecordset.Fields("routename").Value)
         Me.CMBPayBillType.AddItem Trim(vRecordset.Fields("routename").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub

Public Sub vGetHeader()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocdate As Date

On Error GoTo ErrDescription

vQuery = "set dateformat dmy"
gConnection.Execute (vQuery)

If Me.CMBDocType.ListIndex = 0 Then
    Me.CMBHeader.Clear
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_ck_arinvoice order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 1 Then
    CMBHeader.Clear
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_ck_purchaseorder order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 2 Then
    CMBHeader.Clear
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_ck_apinvoice order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 3 Then
    CMBHeader.Clear
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_ck_stktransfer order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 4 Then
    CMBHeader.Clear
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_ck_stkissue order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 5 Then
    CMBHeader.Clear
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_ck_receipt order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 6 Then
    CMBHeader.Clear
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_ck_ardeposit order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 7 Then
    CMBHeader.Clear
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_ck_creditnote order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 8 Then
    CMBHeader.Clear
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_ck_debitnote order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 9 Then
    CMBHeader.Clear
    vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
    ListViewDocNo.ListItems.Clear
    vQuery = "exec dbo.usp_CK_SearchReceiptSlip '" & vDocdate & "'"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 10 Then
    CMBHeader.Clear
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_ck_stkadjust order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 11 Then
    CMBHeader.Clear
    vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_CK_Payment order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 12 Then
    CMBHeader.Clear
    vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_CK_OTHEREXPENSE order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 13 Then
    CMBHeader.Clear
    vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_CK_APDepositSpecial order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 14 Then
    CMBHeader.Clear
    vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_CK_StkRefund order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 15 Then
    CMBHeader.Clear
    vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_CK_BillStatementTemp order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 16 Then
    CMBHeader.Clear
    vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_CK_PayBill order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBDocType.ListIndex = 17 Then
    CMBHeader.Clear
    vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
    ListViewDocNo.ListItems.Clear
    vQuery = "select * from dbo.vw_CK_APInvoice1 order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub vGetReportHeader()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocdate As Date

On Error GoTo ErrDescription

vQuery = "set dateformat dmy"
gConnection.Execute (vQuery)

If Me.CMBRPDocType.ListIndex = 0 Then
    Me.CMBRPHeader.Clear
    
    vQuery = "select * from dbo.vw_ck_arinvoice order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 1 Then
    CMBRPHeader.Clear
    
    vQuery = "select * from dbo.vw_ck_purchaseorder order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 2 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_ck_apinvoice order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 3 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_ck_stktransfer order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 4 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_ck_stkissue order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 5 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_ck_receipt order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 6 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_ck_ardeposit order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 7 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_ck_creditnote order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 8 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_ck_debitnote order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 9 Then
    CMBRPHeader.Clear
    vDocdate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year

    vQuery = "exec dbo.usp_CK_SearchReceiptSlip '" & vDocdate & "'"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 10 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_ck_stkadjust order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 11 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_CK_Payment order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 12 Then
    CMBRPHeader.Clear
    
    vQuery = "select * from dbo.vw_CK_OTHEREXPENSE order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 13 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_CK_APDepositSpecial order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 14 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_CK_StkRefund order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 15 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_CK_BillStatementTemp order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMBRPDocType.ListIndex = 16 Then
    CMBRPHeader.Clear

    vQuery = "select * from dbo.vw_CK_PayBill order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBRPHeader.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListViewDocNo_DblClick()
Dim vDocNo As String
Dim vIndex As Integer

On Error Resume Next

vIndex = Me.ListViewDocNo.SelectedItem.Index

vDocNo = Me.ListViewDocNo.ListItems(vIndex).SubItems(1)
Me.TXTBarCode.Text = vDocNo


Me.TXTMyDescription.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(9)
If Me.ListViewDocNo.ListItems(vIndex).SubItems(10) <> "" Then
Me.CMBInvoiceStatus.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(10)
Else
Me.CMBInvoiceStatus.ListIndex = 0
End If

If Me.ListViewDocNo.ListItems(vIndex).SubItems(11) <> "" Then
Me.CMBReturnStatus.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(11)
Else
Me.CMBReturnStatus.ListIndex = 0
End If

If Me.ListViewDocNo.ListItems(vIndex).SubItems(12) <> "" Then
Me.CMBPaymentStatus.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(12)
Else
Me.CMBPaymentStatus.ListIndex = 0
End If

If Me.ListViewDocNo.ListItems(vIndex).SubItems(13) <> "" Then
Me.CMBCheckStatus.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(13)
Else
Me.CMBCheckStatus.ListIndex = 0
End If


If Me.ListViewDocNo.ListItems(vIndex).SubItems(14) <> "" Then
  Me.CMBPayBill.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(14)
Else
  If Me.CMBDocType.ListIndex = 16 Then
  Me.CMBPayBill.ListIndex = 8
  End If
End If

If Me.ListViewDocNo.ListItems(vIndex).SubItems(15) <> "" Then
Me.DTPPayDate.Value = Me.ListViewDocNo.ListItems(vIndex).SubItems(15)
Me.CBPayBillDate.Value = 1
Else
Me.CBPayBillDate.Value = 0
End If

If Me.ListViewDocNo.ListItems(vIndex).SubItems(16) <> "" Then
Me.DTPRecMoney.Value = Me.ListViewDocNo.ListItems(vIndex).SubItems(16)
Me.CBRecMoney.Value = 1
Else
Me.CBRecMoney.Value = 0
End If


If Me.CMBDocType.ListIndex = 16 Then

  Me.CMBCheckStatus.Enabled = True
  Me.CMBInvoiceStatus.Enabled = True
  Me.CMBPaymentStatus.Enabled = True
  Me.CMBReturnStatus.Enabled = True
  Me.TXTMyDescription.Enabled = True
  
  Me.CMBPayBill.Enabled = True
  Me.CBPayBillDate.Enabled = True
  Me.DTPPayDate.Enabled = True
  Me.CBRecMoney.Enabled = True
  Me.DTPRecMoney.Enabled = True
  
Else

  Me.CMBCheckStatus.Enabled = True
  Me.CMBInvoiceStatus.Enabled = True
  Me.CMBPaymentStatus.Enabled = True
  Me.CMBReturnStatus.Enabled = True
  Me.TXTMyDescription.Enabled = True
  
  Me.CMBPayBill.Enabled = False
  Me.CBPayBillDate.Enabled = False
  Me.DTPPayDate.Enabled = False
  Me.CBRecMoney.Enabled = False
  Me.DTPRecMoney.Enabled = False

End If


Me.CMBInvoiceStatus.SetFocus

End Sub

Public Sub ItemCheck(vIndex As Integer)
Dim vDocNo As String

On Error Resume Next

vDocNo = Me.ListViewDocNo.ListItems(vIndex).SubItems(1)
Me.TXTBarCode.Text = vDocNo


Me.TXTMyDescription.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(9)
If Me.ListViewDocNo.ListItems(vIndex).SubItems(10) <> "" Then
Me.CMBInvoiceStatus.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(10)
Else
Me.CMBInvoiceStatus.ListIndex = 0
End If

If Me.ListViewDocNo.ListItems(vIndex).SubItems(11) <> "" Then
Me.CMBReturnStatus.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(11)
Else
Me.CMBReturnStatus.ListIndex = 0
End If

If Me.ListViewDocNo.ListItems(vIndex).SubItems(12) <> "" Then
Me.CMBPaymentStatus.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(12)
Else
Me.CMBPaymentStatus.ListIndex = 0
End If

If Me.ListViewDocNo.ListItems(vIndex).SubItems(13) <> "" Then
Me.CMBCheckStatus.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(13)
Else
Me.CMBCheckStatus.ListIndex = 0
End If


If Me.ListViewDocNo.ListItems(vIndex).SubItems(14) <> "" Then
  Me.CMBPayBill.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(14)
Else
  If Me.CMBDocType.ListIndex = 16 Then
  Me.CMBPayBill.ListIndex = 8
  End If
End If

If Me.ListViewDocNo.ListItems(vIndex).SubItems(15) <> "" Then
Me.DTPPayDate.Value = Me.ListViewDocNo.ListItems(vIndex).SubItems(15)
Me.CBPayBillDate.Value = 1
Else
Me.CBPayBillDate.Value = 0
End If

If Me.ListViewDocNo.ListItems(vIndex).SubItems(16) <> "" Then
Me.DTPRecMoney.Value = Me.ListViewDocNo.ListItems(vIndex).SubItems(16)
Me.CBRecMoney.Value = 1
Else
Me.CBRecMoney.Value = 0
End If


If Me.CMBDocType.ListIndex = 16 Then

  Me.CMBCheckStatus.Enabled = True
  Me.CMBInvoiceStatus.Enabled = True
  Me.CMBPaymentStatus.Enabled = True
  Me.CMBReturnStatus.Enabled = True
  Me.TXTMyDescription.Enabled = True
  
  Me.CMBPayBill.Enabled = True
  Me.CBPayBillDate.Enabled = True
  Me.DTPPayDate.Enabled = True
  Me.CBRecMoney.Enabled = True
  Me.DTPRecMoney.Enabled = True
  
Else

  Me.CMBCheckStatus.Enabled = True
  Me.CMBInvoiceStatus.Enabled = True
  Me.CMBPaymentStatus.Enabled = True
  Me.CMBReturnStatus.Enabled = True
  Me.TXTMyDescription.Enabled = True
  
  Me.CMBPayBill.Enabled = False
  Me.CBPayBillDate.Enabled = False
  Me.DTPPayDate.Enabled = False
  Me.CBRecMoney.Enabled = False
  Me.DTPRecMoney.Enabled = False

End If


Me.CMBInvoiceStatus.SetFocus
End Sub

Private Sub ListViewDocNo_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vIndex As Integer
Dim vQuery As String
Dim vDocType As Integer
Dim vHeader As String
Dim vDocdate As String
Dim vCheckDocNo As String
Dim vInvoiceStatus As String
Dim vReturnStatus As String
Dim vCheckStatus As String
Dim vPayStatus As String
Dim vPayBillType As String
Dim vPaybillDate As String
Dim vRecMoney As String
Dim vMydescription As String
Dim vCheckDate As Date
Dim vIsPaybillDate As Integer
Dim vIsRecMoney As Integer
Dim vRecMoneyDate As String
Dim vIsPlanDate As Integer
Dim vPlanPayDate  As String

Dim vARCode As String
Dim vPersonCode As String
Dim vNumberNo As String
Dim vRunNumber As String
Dim vDocNo As String
Dim vExistStatus As Integer
Dim vChecker As String

Dim vSlipDescription As String
Dim vReturnDescription As String
Dim vTaxDescription As String
    
Dim vDate As String
Dim vMonth As String
Dim vAutoNumber As String
Dim vDay As String


If Me.ListViewDocNo.ListItems.Count > 0 Then
    vIndex = Item.Index
    If Me.ListViewDocNo.ListItems(vIndex).Checked = True Then
      Call ItemCheck(vIndex)
    Else
        vIndex = Item.Index
        ListViewDocNo.ListItems.Item(vIndex).ListSubItems(18).Text = 0
        

      vCheckDate = DTPDocDate.Day & "/" & DTPDocDate.Month & "/" & DTPDocDate.Year
      If Len(DTPDocDate.Day) = 1 Then
      vDay = Trim(0 & DTPDocDate.Day)
      Else
      vDay = Trim(DTPDocDate.Day)
      End If
      
      vMonth = Trim(DTPDocDate.Month)
  
      vDate = Right(DTPDocDate.Year, 2) + 43 & vMonth & vDay
      vDocType = CMBDocType.ListIndex
      vHeader = Trim(CMBHeader.Text)
      If CMBDocType.ListIndex <> 9 Then
      vNumberNo = vHeader & vDate
      Else
      vNumberNo = vHeader
      End If
        
    
    vRunNumber = vNumberNo
    vDocNo = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(1))
    vARCode = Left(Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(4)), InStr(Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(4)), "/") - 1)
    vPersonCode = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(5))
    vExistStatus = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(18))

    
    If Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(9)) <> "" Then
    vMydescription = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(9))
    Else
    vMydescription = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(9))
    End If
    
    vSlipDescription = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(10))
    vReturnDescription = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(11))
    vTaxDescription = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(12))
    vCheckStatus = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(13))
    vPayBillType = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(14))
            
    
    If Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(15)) <> "" Then
    vPlanPayDate = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(15))
    vIsPlanDate = 1
    Else
    vIsPlanDate = 0
    vPlanPayDate = ""
    End If
    
    
    If Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(16)) <> "" Then
    vPaybillDate = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(15))
    vIsPaybillDate = 1
    Else
    vIsPaybillDate = 0
    vPaybillDate = ""
    End If
    
    If Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(17)) <> "" Then
    vIsRecMoney = 1
    vRecMoneyDate = Trim(ListViewDocNo.ListItems.Item(vIndex).SubItems(16))
    Else
    vIsRecMoney = 0
    vRecMoneyDate = ""
    End If
    
   
    vQuery = "exec dbo.USP_NP_InsertCheckDataExist '" & vRunNumber & "'," & vDocType & ",'" & vDocNo & "','" & vCheckDate & "'," & vExistStatus & ",'" & vMydescription & "','" & vUserID & "' ,'" & vARCode & "' ,'" & vPersonCode & "','" & vSlipDescription & "','" & vReturnDescription & "' ,'" & vTaxDescription & "', '" & vCheckStatus & "','" & vPayBillType & "'," & vIsPlanDate & ",'" & vPlanPayDate & "'," & vIsPaybillDate & ",'" & vPaybillDate & "'," & vIsRecMoney & ",'" & vRecMoneyDate & "' "
    gConnection.Execute vQuery
    
        
    ListViewDocNo.ListItems.Item(vIndex).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(1).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(2).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(3).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(4).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(5).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(6).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(7).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(8).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(9).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(10).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(11).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(12).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(13).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(14).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(15).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(16).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(17).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(18).ForeColor = "&H80000008"
    ListViewDocNo.ListItems.Item(vIndex).ListSubItems(19).ForeColor = "&H80000008"
    
    End If
    
    
    
      

End If

End Sub

Private Sub TXTBarCode_KeyPress(KeyAscii As Integer)
Dim vDocType As Integer
Dim vHeader As String
Dim vDocdate As String
Dim i As Integer
Dim vBarCode As String
Dim vCheckDocNo As String
Dim vCount As Integer

On Error Resume Next

If KeyAscii = 13 And Me.TXTBarCode.Text <> "" Then
    vCount = 0
    vBarCode = UCase(Me.TXTBarCode.Text)
    
    For i = 1 To Me.ListViewDocNo.ListItems.Count
    vCheckDocNo = ListViewDocNo.ListItems.Item(i).ListSubItems(1).Text
    
      If vBarCode = vCheckDocNo Then
      ListViewDocNo.ListItems.Item(i).ListSubItems(2).Text = vBarCode
      ListViewDocNo.ListItems.Item(i).ListSubItems(18).Text = 1
      ListViewDocNo.ListItems.Item(i).Checked = True
      vCount = vCount + 1
      End If
    Next i
        
    If vCount = 0 Then
      MsgBox "ไม่เจอเลขที่เอกสาร " & vBarCode & " นี้ในระบบ  กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      Me.CMBCheckStatus.Enabled = False
      Me.CMBInvoiceStatus.Enabled = False
      Me.CMBPaymentStatus.Enabled = False
      Me.CMBReturnStatus.Enabled = False
      
      Me.CMBPayBill.Enabled = False
      Me.CBPayBillDate.Enabled = False
      Me.DTPPayDate.Enabled = False
      Me.CBRecMoney.Enabled = False
      Me.DTPRecMoney.Enabled = False
      Me.TXTMyDescription.Enabled = False

      Else
      Me.TXTBarCode.Text = UCase(Me.TXTBarCode.Text)
      Me.CMBCheckStatus.Enabled = True
      Me.CMBInvoiceStatus.Enabled = True
      Me.CMBPaymentStatus.Enabled = True
      Me.CMBReturnStatus.Enabled = True
      Me.TXTMyDescription.Enabled = True
      
      If Me.CMBDocType.ListIndex <> 16 Then
        Me.CMBPayBill.Enabled = False
        Me.CBPayBillDate.Enabled = False
        Me.DTPPayDate.Enabled = False
        Me.CBRecMoney.Enabled = False
        Me.DTPRecMoney.Enabled = False
      Else
        Me.CMBPayBill.Enabled = True
        Me.CBPayBillDate.Enabled = True
        Me.DTPPayDate.Enabled = True
        Me.CBRecMoney.Enabled = True
        Me.DTPRecMoney.Enabled = True
      End If

    End If
End If
End Sub

