VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form97 
   Caption         =   "เปลี่ยนเลขที่เอกสาร/เลขที่ภาษี/วันที่เอกสาร"
   ClientHeight    =   11010
   ClientLeft      =   2595
   ClientTop       =   450
   ClientWidth     =   15375
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MousePointer    =   99  'Custom
   Picture         =   "Form97.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICXPC 
      BackColor       =   &H80000001&
      Height          =   10995
      Left            =   0
      ScaleHeight     =   10935
      ScaleWidth      =   15300
      TabIndex        =   66
      Top             =   -45
      Visible         =   0   'False
      Width           =   15360
      Begin VB.CommandButton CMD202 
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
         Height          =   555
         Left            =   3600
         TabIndex        =   100
         Top             =   7560
         Width           =   1950
      End
      Begin VB.CommandButton CMD201 
         Caption         =   "ปรับปรุงเอกสาร"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1665
         TabIndex        =   99
         Top             =   7560
         Width           =   1740
      End
      Begin MSComCtl2.DTPicker DTP202 
         Height          =   375
         Left            =   11295
         TabIndex        =   98
         Top             =   6255
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
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
         Format          =   20447233
         CurrentDate     =   42271
      End
      Begin VB.TextBox TXT206 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   420
         Left            =   3780
         TabIndex        =   94
         Top             =   6255
         Width           =   2640
      End
      Begin VB.TextBox TXT203 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   420
         Left            =   3780
         TabIndex        =   93
         Top             =   5535
         Width           =   2640
      End
      Begin VB.TextBox TXT202 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   420
         Left            =   3780
         TabIndex        =   92
         Top             =   4860
         Width           =   2640
      End
      Begin VB.TextBox TXT201 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   3780
         TabIndex        =   91
         Top             =   4185
         Width           =   2640
      End
      Begin VB.TextBox TXT205 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   420
         Left            =   11295
         TabIndex        =   90
         Top             =   5535
         Width           =   2640
      End
      Begin MSComCtl2.DTPicker DTP201 
         Height          =   375
         Left            =   11295
         TabIndex        =   89
         Top             =   4860
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
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
         Format          =   20447233
         CurrentDate     =   42271
      End
      Begin VB.TextBox TXT204 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   420
         Left            =   11295
         TabIndex        =   88
         Top             =   4185
         Width           =   2640
      End
      Begin VB.ComboBox CMB201 
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
         Left            =   11295
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   3285
         Width           =   2715
      End
      Begin VB.PictureBox Picture4 
         Height          =   3885
         Left            =   7830
         ScaleHeight     =   3825
         ScaleWidth      =   45
         TabIndex        =   78
         Top             =   3285
         Width           =   105
      End
      Begin VB.PictureBox Picture3 
         Height          =   105
         Index           =   1
         Left            =   720
         ScaleHeight     =   45
         ScaleWidth      =   14310
         TabIndex        =   77
         Top             =   7290
         Width           =   14370
      End
      Begin VB.PictureBox Picture3 
         Height          =   105
         Index           =   0
         Left            =   675
         ScaleHeight     =   45
         ScaleWidth      =   14355
         TabIndex        =   76
         Top             =   3060
         Width           =   14415
      End
      Begin VB.CheckBox CBCreditNote_XPC 
         BackColor       =   &H80000001&
         Caption         =   "เลือกเลขที่ลดหนี้"
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
         Left            =   1665
         TabIndex        =   74
         Top             =   3285
         Width           =   1815
      End
      Begin VB.CheckBox CHK204 
         BackColor       =   &H80000001&
         Caption         =   "เปลี่ยนวันที่ใบกำกับภาษี"
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
         Height          =   420
         Left            =   1665
         TabIndex        =   73
         Top             =   2520
         Width           =   3930
      End
      Begin VB.CheckBox CHK202 
         BackColor       =   &H80000001&
         Caption         =   "เปลี่ยนเลขที่ใบกำกับภาษี"
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
         Height          =   420
         Left            =   1665
         TabIndex        =   72
         Top             =   2025
         Width           =   3435
      End
      Begin VB.CheckBox CHK203 
         BackColor       =   &H80000001&
         Caption         =   "เปลี่ยนวันที่เอกสาร"
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
         Height          =   420
         Left            =   1665
         TabIndex        =   71
         Top             =   1485
         Width           =   2085
      End
      Begin VB.CheckBox CHK201 
         BackColor       =   &H80000001&
         Caption         =   "เปลี่ยนเลขที่เอกสาร"
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
         Left            =   1665
         TabIndex        =   70
         Top             =   990
         Width           =   3120
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XPC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Left            =   45
         TabIndex        =   96
         Top             =   90
         Width           =   1455
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกหัวเอกสาร"
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
         Left            =   8550
         TabIndex        =   95
         Top             =   3285
         Width           =   1365
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "เปลี่ยนวันที่ใบกำกับภาษี"
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
         Left            =   8550
         TabIndex        =   86
         Top             =   6255
         Width           =   2940
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "เปลี่ยนเป็นใบกำกับภาษีเลขที่"
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
         Left            =   8550
         TabIndex        =   85
         Top             =   5535
         Width           =   3165
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "เปลี่ยนเป็นเอกสารวันที่"
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
         Height          =   420
         Left            =   8550
         TabIndex        =   84
         Top             =   4860
         Width           =   1725
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "เปลี่ยนเป็นเอกสารเลขที่"
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
         Left            =   8550
         TabIndex        =   83
         Top             =   4185
         Width           =   2580
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ใบกำกับภาษี"
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
         Left            =   1665
         TabIndex        =   82
         Top             =   6255
         Width           =   1860
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบกำกับภาษี"
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
         Left            =   1665
         TabIndex        =   81
         Top             =   5535
         Width           =   1725
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่เอกสาร"
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
         Left            =   1665
         TabIndex        =   80
         Top             =   4860
         Width           =   1725
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่เอกสาร"
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
         Left            =   1665
         TabIndex        =   79
         Top             =   4185
         Width           =   1590
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000001&
         Caption         =   "เลือก การเปลี่ยนข้อมูล"
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
         Height          =   285
         Left            =   1665
         TabIndex        =   75
         Top             =   585
         Width           =   2805
      End
   End
   Begin VB.PictureBox PICCompany 
      BackColor       =   &H80000001&
      Height          =   10905
      Left            =   0
      ScaleHeight     =   10845
      ScaleWidth      =   15300
      TabIndex        =   65
      Top             =   -45
      Width           =   15360
      Begin VB.CommandButton CMDCompany 
         Caption         =   "เลือก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6300
         TabIndex        =   69
         Top             =   2700
         Width           =   1770
      End
      Begin VB.ComboBox CMBCompany 
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
         ItemData        =   "Form97.frx":9673
         Left            =   4320
         List            =   "Form97.frx":967D
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   1665
         Width           =   3795
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกบริษัท ทำงาน :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1980
         TabIndex        =   68
         Top             =   1710
         Width           =   2310
      End
   End
   Begin VB.CommandButton CMDExit 
      Caption         =   "เลือกบริษัท"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9585
      TabIndex        =   97
      Top             =   6750
      Width           =   1725
   End
   Begin VB.PictureBox PICDocDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   11445
      Left            =   2745
      ScaleHeight     =   11415
      ScaleWidth      =   15330
      TabIndex        =   33
      Top             =   8100
      Visible         =   0   'False
      Width           =   15360
      Begin VB.PictureBox PICEditData 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   10890
         Left            =   45
         ScaleHeight     =   10860
         ScaleWidth      =   15225
         TabIndex        =   37
         Top             =   45
         Visible         =   0   'False
         Width           =   15255
         Begin VB.CommandButton CMDCloseEditData 
            BackColor       =   &H00C0C0C0&
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
            Height          =   600
            Left            =   9855
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   3600
            Width           =   1725
         End
         Begin VB.ComboBox CMBHeader 
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
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   765
            Width           =   1680
         End
         Begin VB.CommandButton CMDChangeData 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ปรับปรุง"
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
            Left            =   7920
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   3600
            Width           =   1725
         End
         Begin VB.TextBox TXTNewTaxNo 
            Appearance      =   0  'Flat
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
            Height          =   330
            Left            =   7920
            TabIndex        =   52
            Top             =   1890
            Width           =   1680
         End
         Begin MSComCtl2.DTPicker DTPNewTaxDate 
            Height          =   330
            Left            =   7920
            TabIndex        =   51
            Top             =   2970
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20447233
            CurrentDate     =   40463
         End
         Begin MSComCtl2.DTPicker DTPNewDocDate 
            Height          =   330
            Left            =   2700
            TabIndex        =   50
            Top             =   2970
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20447233
            CurrentDate     =   40463
         End
         Begin VB.CheckBox CKTaxDate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "วันที่ใบกำกับภาษี"
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
            Left            =   6345
            TabIndex        =   48
            Top             =   2520
            Width           =   1500
         End
         Begin VB.CheckBox CKTaxNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "เลขที่ใบกำกับภาษี"
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
            Left            =   6300
            TabIndex        =   47
            Top             =   1440
            Width           =   1545
         End
         Begin VB.CheckBox CKDocDate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "วันที่เอกสาร"
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
            Left            =   1440
            TabIndex        =   46
            Top             =   2520
            Width           =   1185
         End
         Begin VB.CheckBox CKDocNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "เลขที่เอกสาร"
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
            Left            =   1395
            TabIndex        =   45
            Top             =   1440
            Width           =   1230
         End
         Begin VB.Label LBLIndex 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   7920
            TabIndex        =   57
            Top             =   810
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label Label16 
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
            Height          =   285
            Left            =   1035
            TabIndex        =   55
            Top             =   810
            Width           =   1635
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "ปรับปรุงข้อมูล"
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
            Left            =   2700
            TabIndex        =   53
            Top             =   225
            Width           =   2895
         End
         Begin VB.Label LBLNewDocNo 
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
            Left            =   2700
            TabIndex        =   49
            Top             =   1890
            Width           =   1680
         End
         Begin VB.Label LBLTaxDate 
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
            Left            =   7920
            TabIndex        =   44
            Top             =   2520
            Width           =   1680
         End
         Begin VB.Label LBLTaxNo 
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
            Left            =   7920
            TabIndex        =   43
            Top             =   1440
            Width           =   1680
         End
         Begin VB.Label LBLDocDate 
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
            Left            =   2700
            TabIndex        =   42
            Top             =   2520
            Width           =   1680
         End
         Begin VB.Label LBLDocNo 
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
            Left            =   2700
            TabIndex        =   41
            Top             =   1440
            Width           =   1680
         End
      End
      Begin VB.CommandButton CMDClose 
         BackColor       =   &H00404040&
         Caption         =   "ปิดหน้าจอ"
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
         Left            =   13185
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   8640
         Width           =   1590
      End
      Begin VB.CommandButton CMDSave 
         BackColor       =   &H00404040&
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
         Left            =   11385
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   8640
         Width           =   1590
      End
      Begin MSComctlLib.ListView ListViewDocNo 
         Height          =   6630
         Left            =   360
         TabIndex        =   36
         Top             =   1935
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   11695
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "เลขที่เอกสาร"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "หัวเอกสาร"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "วันที่เอกสาร"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "วันที่เอกสารใหม่"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "เลขที่ใบกำกับภาษี"
            Object.Width           =   3263
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "เลขที่ใบกำกับภาษีใหม่"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "วันที่ใบกำกับภาษี"
            Object.Width           =   3263
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "วันที่ใบกำกับภาษีใหม่"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPSearchDoc 
         Height          =   375
         Left            =   1980
         TabIndex        =   35
         Top             =   855
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   20447233
         CurrentDate     =   40462
      End
      Begin VB.CheckBox CBSelectAll 
         BackColor       =   &H00C0C0C0&
         Caption         =   "เลือกทั้งหมด"
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
         Left            =   360
         TabIndex        =   61
         Top             =   1395
         Width           =   1590
      End
      Begin MSComCtl2.DTPicker DTPTaxDateAll 
         Height          =   330
         Left            =   5490
         TabIndex        =   60
         Top             =   1395
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
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
         Format          =   20447233
         CurrentDate     =   40474
      End
      Begin VB.CommandButton CMDTaxAll 
         BackColor       =   &H00404040&
         Caption         =   "ปรับ"
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
         Left            =   7155
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1395
         Width           =   510
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ปรับวันที่ใบกำกับภาษีเอกสารที่เลือกไว้ให้เป็นวันที่ :"
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
         Left            =   1980
         TabIndex        =   59
         Top             =   1440
         Width           =   3525
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "ปรับปรุงข้อมูลเอกสาร ตามวันที่"
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
         Left            =   360
         TabIndex        =   40
         Top             =   225
         Width           =   2670
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกวันที่ใบกำกับภาษี :"
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
         Left            =   360
         TabIndex        =   34
         Top             =   900
         Width           =   1635
      End
   End
   Begin VB.CheckBox CBCreditNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลือกเลขที่ลดหนี้"
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
      Left            =   2250
      TabIndex        =   64
      Top             =   2700
      Width           =   1725
   End
   Begin VB.CommandButton CMDDocDate 
      Caption         =   "เปลี่ยนตามวันที่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7740
      TabIndex        =   32
      Top             =   6750
      Width           =   1725
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   360
      Top             =   7650
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
   Begin VB.CheckBox CHK104 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3360
      TabIndex        =   30
      Top             =   1995
      Width           =   165
   End
   Begin MSComCtl2.DTPicker DTP102 
      Height          =   390
      Left            =   8250
      TabIndex        =   28
      Top             =   4950
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   688
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   20447233
      CurrentDate     =   38399
   End
   Begin VB.TextBox TXT106 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2250
      TabIndex        =   26
      Top             =   4950
      Width           =   2640
   End
   Begin VB.ComboBox CMB101 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8250
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   2850
      Width           =   2505
   End
   Begin VB.CommandButton CMD103 
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
      Height          =   540
      Left            =   5895
      TabIndex        =   23
      Top             =   6750
      Width           =   1740
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "ล้างข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4050
      TabIndex        =   22
      Top             =   6750
      Width           =   1740
   End
   Begin VB.PictureBox Picture2 
      Height          =   3390
      Left            =   5250
      ScaleHeight     =   3330
      ScaleWidth      =   30
      TabIndex        =   21
      Top             =   2550
      Width           =   90
   End
   Begin VB.PictureBox Picture1 
      Height          =   90
      Index           =   1
      Left            =   75
      ScaleHeight     =   30
      ScaleWidth      =   11835
      TabIndex        =   20
      Top             =   6000
      Width           =   11895
   End
   Begin VB.PictureBox Picture1 
      Height          =   90
      Index           =   0
      Left            =   75
      ScaleHeight     =   30
      ScaleWidth      =   11835
      TabIndex        =   19
      Top             =   2400
      Width           =   11895
   End
   Begin VB.TextBox TXT105 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8250
      TabIndex        =   16
      Top             =   4425
      Width           =   2505
   End
   Begin VB.TextBox TXT102 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2250
      TabIndex        =   15
      Top             =   3900
      Width           =   2640
   End
   Begin VB.CheckBox CHK101 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3360
      TabIndex        =   11
      Top             =   1095
      Width           =   165
   End
   Begin VB.CheckBox CHK103 
      Caption         =   "Check3"
      Height          =   195
      Left            =   3360
      TabIndex        =   10
      Top             =   1395
      Width           =   165
   End
   Begin VB.CheckBox CHK102 
      Caption         =   "Check2"
      Height          =   195
      Left            =   3360
      TabIndex        =   9
      Top             =   1695
      Width           =   165
   End
   Begin MSComCtl2.DTPicker DTP101 
      Height          =   390
      Left            =   8250
      TabIndex        =   2
      Top             =   3900
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   688
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   20447233
      CurrentDate     =   38397
   End
   Begin VB.TextBox TXT103 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2250
      TabIndex        =   1
      Top             =   4425
      Width           =   2640
   End
   Begin VB.TextBox TXT101 
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
      Left            =   2250
      TabIndex        =   0
      Top             =   3375
      Width           =   2640
   End
   Begin VB.TextBox TXT104 
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
      Left            =   8250
      TabIndex        =   4
      Top             =   3375
      Width           =   2505
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "ปรับปรุงเอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2205
      TabIndex        =   3
      Top             =   6750
      Width           =   1740
   End
   Begin VB.CheckBox CKBI 
      BackColor       =   &H80000001&
      Caption         =   "ต้องการเปลี่ยนที่ BI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3390
      TabIndex        =   63
      Top             =   450
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนวันที่ใบกำกับภาษี"
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
      Left            =   3660
      TabIndex        =   31
      Top             =   1995
      Width           =   1890
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนวันที่ใบกำกับภาษี"
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
      Height          =   390
      Left            =   5760
      TabIndex        =   29
      Top             =   4995
      Width           =   2415
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ใบกำกับภาษี"
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
      Left            =   495
      TabIndex        =   27
      Top             =   4995
      Width           =   1665
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกหัวเอกสาร"
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
      Left            =   6435
      TabIndex        =   25
      Top             =   2880
      Width           =   1740
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนเป็นเอกสารวันที่"
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
      Left            =   5985
      TabIndex        =   18
      Top             =   3960
      Width           =   2190
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนเป็นใบกำกับภาษีเลขที่"
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
      Left            =   5625
      TabIndex        =   17
      Top             =   4500
      Width           =   2565
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เอกสาร"
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
      Height          =   255
      Left            =   765
      TabIndex        =   14
      Top             =   3960
      Width           =   1410
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบกำกับภาษี"
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
      Left            =   360
      TabIndex        =   13
      Top             =   4500
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
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
      Left            =   810
      TabIndex        =   12
      Top             =   3420
      Width           =   1365
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนวันที่เอกสาร"
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
      Height          =   270
      Left            =   3660
      TabIndex        =   8
      Top             =   1350
      Width           =   1440
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนเป็นเอกสารเลขที่"
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
      Left            =   5805
      TabIndex        =   7
      Top             =   3420
      Width           =   2370
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนเลขที่เอกสาร"
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
      Left            =   3660
      TabIndex        =   6
      Top             =   1080
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนเลขที่ใบกำกับภาษี"
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
      Left            =   3660
      TabIndex        =   5
      Top             =   1695
      Width           =   1965
   End
End
Attribute VB_Name = "Form97"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim vCountNum As Integer
Dim vChangeDoc1 As String
Dim vNewDocNo As String


Private Sub CBCreditNote_Click()
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vRecordset2 As New ADODB.Recordset
Dim vDocNo As String
Dim vQuery As String
Dim vType As Integer

'On Error Resume Next

TXT101.Text = ""
TXT102.Text = ""
TXT103.Text = ""
TXT104.Text = ""
TXT105.Text = ""
TXT106.Text = ""
CHK101.Value = 0
CHK102.Value = 0
CHK103.Value = 0
CHK104.Value = 0


If Me.TXT101.Text <> "" Then
        vDocNo = TXT101.Text
        Call CheckNumeric
        

            vQuery = " exec dbo.USP_NP_SearchDocnoEditTaxData " & Me.CBCreditNote.Value & ",'" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                TXT102.Text = Trim(vRecordset.Fields("docdate").Value)
                    If Not IsNull(Trim(vRecordset.Fields("taxno").Value)) Then
                        TXT103.Text = Trim(vRecordset.Fields("taxno").Value)
                    Else
                        TXT103.Text = "NoTaxNo"
                    End If
                    If Not IsNull(Trim(vRecordset.Fields("taxdate").Value)) Then
                        TXT106.Text = Trim(vRecordset.Fields("taxdate").Value)
                    Else
                        TXT106.Text = Trim(vRecordset.Fields("docdate").Value)
                    End If
                Else
                MsgBox "เอกสาร เลขที่ " & vDocNo & " ไม่สามารถเปลี่ยนเอกสารได้ โปรดตรวจสอบด้วยนะครับ"
                TXT104.Text = ""
                TXT105.Text = ""
                Exit Sub
            End If
            vRecordset.Close
                
            TXT104.Text = ""
            TXT105.Text = ""
            DTP101.Value = TXT102.Text
            DTP102.Value = TXT106.Text
    End If
    
    
    If Me.CBCreditNote.Value = 1 Then
        vType = 1
    Else
        vType = 0
    End If
    
    
    Me.CMB101.Clear
    Me.CMBHeader.Clear
    
    vQuery = "exec dbo.USP_NP_CheckHeaderChangeTax " & vType & ""
    If OpenDataBase(gConnection, vRecordset2, vQuery) <> 0 Then
        vRecordset2.MoveFirst
        While Not vRecordset2.EOF
        CMB101.AddItem Trim(vRecordset2.Fields("docno").Value)
        Me.CMBHeader.AddItem Trim(vRecordset2.Fields("docno").Value)
        vRecordset2.MoveNext
        Wend
    End If
    vRecordset2.Close


End Sub

Private Sub CBCreditNote_XPC_Click()
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vRecordset2 As New ADODB.Recordset
Dim vDocNo As String
Dim vQuery As String
Dim vType As Integer

'On Error Resume Next

TXT201.Text = ""
TXT202.Text = ""
TXT203.Text = ""
TXT204.Text = ""
TXT205.Text = ""
TXT206.Text = ""
CHK201.Value = 0
CHK202.Value = 0
CHK203.Value = 0
CHK204.Value = 0


If Me.TXT201.Text <> "" Then
        vDocNo = TXT201.Text
        Call CheckNumeric
        

            vQuery = " exec dbo.USP_NP_SearchDocnoEditTaxData " & Me.CBCreditNote_XPC.Value & ",'" & vDocNo & "' "
            If OpenDataBaseXPC(vXPCConnection, vRecordset, vQuery) <> 0 Then
                TXT202.Text = Trim(vRecordset.Fields("docdate").Value)
                    If Not IsNull(Trim(vRecordset.Fields("taxno").Value)) Then
                        TXT203.Text = Trim(vRecordset.Fields("taxno").Value)
                    Else
                        TXT203.Text = "NoTaxNo"
                    End If
                    If Not IsNull(Trim(vRecordset.Fields("taxdate").Value)) Then
                        TXT206.Text = Trim(vRecordset.Fields("taxdate").Value)
                    Else
                        TXT206.Text = Trim(vRecordset.Fields("docdate").Value)
                    End If
                Else
                MsgBox "เอกสาร เลขที่ " & vDocNo & " ไม่สามารถเปลี่ยนเอกสารได้ โปรดตรวจสอบด้วยนะครับ"
                TXT204.Text = ""
                TXT205.Text = ""
                Exit Sub
            End If
            vRecordset.Close
                
            TXT204.Text = ""
            TXT205.Text = ""
            DTP201.Value = TXT202.Text
            DTP202.Value = TXT206.Text
    End If
    
    
    If Me.CBCreditNote_XPC.Value = 1 Then
        vType = 1
    Else
        vType = 0
    End If
    
    
    'Me.CMB201.Clear
    'Me.CMBHeader_XPC.Clear
    
    'vQuery = "exec dbo.USP_NP_CheckHeaderChangeTax " & vType & ""
    'If OpenDataBaseXPC(vXPCConnection, vRecordset2, vQuery) <> 0 Then
     '   vRecordset2.MoveFirst
      '  While Not vRecordset2.EOF
        'CMB201.AddItem Trim(vRecordset2.Fields("docno").Value)
       ' Me.CMBHeader_XPC.AddItem Trim(vRecordset2.Fields("docno").Value)
        'vRecordset2.MoveNext
        'Wend
    'End If
    'vRecordset2.Close
End Sub

Private Sub CBSelectAll_Click()
Dim i As Integer

If Me.ListViewDocNo.ListItems.Count > 0 Then

If Me.CBSelectAll.Value = 1 Then
For i = 1 To Me.ListViewDocNo.ListItems.Count
Me.ListViewDocNo.ListItems(i).Checked = True
Next i
End If

If Me.CBSelectAll.Value = 0 Then
For i = 1 To Me.ListViewDocNo.ListItems.Count
Me.ListViewDocNo.ListItems(i).Checked = False
Me.ListViewDocNo.ListItems(i).SubItems(2) = ""
Me.ListViewDocNo.ListItems(i).SubItems(4) = ""
Me.ListViewDocNo.ListItems(i).SubItems(6) = ""
Me.ListViewDocNo.ListItems(i).SubItems(8) = ""
Next i
End If

End If
End Sub

Private Sub CHK101_Click()


If CMB101.Enabled = False Then
    CMB101.Enabled = True
Else
    CMB101.Enabled = False
    TXT104.Text = ""
End If
TXT101.SetFocus
End Sub

Private Sub CHK102_Click()
If TXT105.Enabled = False Then
    TXT105.Enabled = True
Else
    TXT105.Enabled = False
End If
TXT101.SetFocus
End Sub

Private Sub CHK103_Click()
If DTP101.Enabled = False Then
    DTP101.Enabled = True
Else
    DTP101.Enabled = False
End If

TXT101.SetFocus
End Sub

Private Sub CHK104_Click()
If DTP102.Enabled = False Then
    DTP102.Enabled = True
Else
    DTP102.Enabled = False
End If

TXT101.SetFocus
End Sub


Private Sub CHK201_Click()
If CMB201.Enabled = False Then
    CMB201.Enabled = True
Else
    CMB201.Enabled = False
    TXT204.Text = ""
End If
TXT201.SetFocus
End Sub

Private Sub CHK202_Click()
If TXT205.Enabled = False Then
    TXT205.Enabled = True
Else
    TXT205.Enabled = False
End If
TXT201.SetFocus
End Sub

Private Sub CHK203_Click()
If DTP201.Enabled = False Then
    DTP201.Enabled = True
Else
    DTP201.Enabled = False
End If

TXT201.SetFocus
End Sub

Private Sub CHK204_Click()
If DTP202.Enabled = False Then
    DTP202.Enabled = True
Else
    DTP202.Enabled = False
End If

TXT201.SetFocus
End Sub

Private Sub CKBI_Click()
If Me.CKBI.Value = 1 Then
Call InitializeDataBaseBI
End If
End Sub

Private Sub CKDocDate_Click()
If Me.CKDocDate.Value = 1 Then
Me.DTPNewDocDate.Enabled = True
Else
Me.DTPNewDocDate.Value = Now
Me.DTPNewDocDate.Enabled = False
End If
End Sub

Private Sub CKDocNo_Click()
If Me.CKDocNo.Value = 1 Then
    Me.CMBHeader.Enabled = True
Else
    Me.CMBHeader.Enabled = False
    Me.LBLNewDocNo.Caption = ""
End If
End Sub

Private Sub CKTaxDate_Click()
If Me.CKTaxDate.Value = 1 Then
Me.DTPNewTaxDate.Enabled = True
Else
Me.DTPNewTaxDate.Value = Now
Me.DTPNewTaxDate.Enabled = False
End If
End Sub

Private Sub CKTaxNo_Click()
If Me.CKTaxNo.Value = 1 Then
Me.TXTNewTaxNo.Enabled = True
Else
Me.TXTNewTaxNo.Enabled = False
Me.TXTNewTaxNo.Text = ""
End If
End Sub

Private Sub CMB101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vChangeDoc As String
Dim vLenHeader As Integer
Dim vLenNumber As Integer
Dim vAnswer As Integer
Dim vLenNumber1 As String, vLenNumber2 As String, vLenNumber3 As String
Dim vLenNumber4 As String, vLenNumber5 As String, vLenNumber6 As String
Dim vSelectDoc As String
Dim vDocNo1 As String

On Error GoTo ErrDescription

If TXT101.Text <> "" Then
vDocNo = Trim(TXT101.Text)
vLenNumber = (Len(vDocNo) - InStr(1, vDocNo, "-"))
vLenHeader = InStr(1, vDocNo, "-")
vLenNumber1 = Year(TXT102.Text)
If Year(TXT102.Text) < 2500 Then
vLenNumber1 = Year(TXT102.Text) + 543
Else
vLenNumber1 = Year(TXT102.Text)
End If
vLenNumber2 = Right(vLenNumber1, 2)
vLenNumber3 = Month(TXT102.Text)
vLenNumber4 = Len(vLenNumber3)
If vLenNumber4 < 2 Then
vLenNumber5 = "0" & vLenNumber3
Else
vLenNumber5 = vLenNumber3
End If
vLenNumber6 = vLenNumber2 & vLenNumber5
vSelectDoc = Trim(CMB101.Text)
vDocNo1 = vLenNumber6 'Left(Right(vDocno, Len(vDocno) - vCountNum), vLenNumber)

'vQuery = "select    top 1 right(docno," & vLenNumber & ")+1 as docno" _
 '                   & " from bcnp.dbo.bcapinvoice  " _
  '                  & " where ltrim(left(docno," & vCountNum & ")) =   '" & vSelectDoc & "' and " _
   '                & " left(right(docno,len(docno)-" & vCountNum & ")," & vLenNumber & ") = left(right('" & vDocNo1 & "',len('" & vDocNo & "')-" & vCountNum & ")," & vLenNumber & ")  order by docno desc "
   

vQuery = "exec dbo.USP_NP_GetNewRunningTaxNo " & Me.CBCreditNote.Value & "," & vLenNumber & "," & vCountNum & ",'" & vDocNo1 & "','" & vDocNo & "','" & vSelectDoc & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vChangeDoc = Trim(vRecordset.Fields("docno").Value)
Else
    vChangeDoc = Format(1, "0000")
End If
vRecordset.Close

vChangeDoc1 = UCase(Format(vChangeDoc, "0000"))
vChangeDoc1 = vSelectDoc & vDocNo1 & "-" & vChangeDoc1
TXT104.Text = vChangeDoc1
Else
MsgBox "กรุณาใส่เลขที่เอกสารที่ต้องการปรับปรุงก่อนนะครับ"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMB201_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vChangeDoc As String
Dim vLenHeader As Integer
Dim vLenNumber As Integer
Dim vAnswer As Integer
Dim vLenNumber1 As String, vLenNumber2 As String, vLenNumber3 As String
Dim vLenNumber4 As String, vLenNumber5 As String, vLenNumber6 As String
Dim vSelectDoc As String
Dim vDocNo1 As String

On Error GoTo ErrDescription

If TXT201.Text <> "" Then
vDocNo = Trim(TXT201.Text)
vLenNumber = (Len(vDocNo) - InStr(1, vDocNo, "-"))
vLenHeader = InStr(1, vDocNo, "-")
vLenNumber1 = Year(TXT202.Text)
If Year(TXT202.Text) < 2500 Then
vLenNumber1 = Year(TXT202.Text) + 543
Else
vLenNumber1 = Year(TXT202.Text)
End If
vLenNumber2 = Right(vLenNumber1, 2)
vLenNumber3 = Month(TXT202.Text)
vLenNumber4 = Len(vLenNumber3)
If vLenNumber4 < 2 Then
vLenNumber5 = "0" & vLenNumber3
Else
vLenNumber5 = vLenNumber3
End If
vLenNumber6 = vLenNumber2 & vLenNumber5
vSelectDoc = Trim(CMB101.Text)
vDocNo1 = vLenNumber6

vQuery = "exec dbo.USP_NP_GetNewRunningTaxNo " & Me.CBCreditNote.Value & "," & vLenNumber & "," & vCountNum & ",'" & vDocNo1 & "','" & vDocNo & "','" & vSelectDoc & "' "
If OpenDataBaseXPC(vXPCConnection, vRecordset, vQuery) <> 0 Then
    vChangeDoc = Trim(vRecordset.Fields("docno").Value)
Else
    vChangeDoc = Format(1, "0000")
End If
vRecordset.Close

vChangeDoc1 = UCase(Format(vChangeDoc, "0000"))
vChangeDoc1 = vSelectDoc & vDocNo1 & "-" & vChangeDoc1
TXT204.Text = vChangeDoc1
Else
MsgBox "กรุณาใส่เลขที่เอกสารที่ต้องการปรับปรุงก่อนนะครับ"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMBHeader_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vChangeDoc As String
Dim vLenHeader As Integer
Dim vLenNumber As Integer
Dim vAnswer As Integer
Dim vLenNumber1 As String, vLenNumber2 As String, vLenNumber3 As String
Dim vLenNumber4 As String, vLenNumber5 As String, vLenNumber6 As String
Dim vSelectDoc As String
Dim vDocNo1 As String

On Error GoTo ErrDescription

If Me.LBLDocNo.Caption <> "" Then
vDocNo = Trim(Me.LBLDocNo.Caption)
vLenNumber = (Len(vDocNo) - InStr(1, vDocNo, "-"))
vLenHeader = InStr(1, vDocNo, "-")
vLenNumber1 = Year(Me.LBLDocDate.Caption)
If Year(Me.LBLDocDate.Caption) < 2500 Then
vLenNumber1 = Year(Me.LBLDocDate.Caption) + 543
Else
vLenNumber1 = Year(Me.LBLDocDate.Caption)
End If
vLenNumber2 = Right(vLenNumber1, 2)
vLenNumber3 = Month(Me.LBLDocDate.Caption)
vLenNumber4 = Len(vLenNumber3)
If vLenNumber4 < 2 Then
vLenNumber5 = "0" & vLenNumber3
Else
vLenNumber5 = vLenNumber3
End If
vLenNumber6 = vLenNumber2 & vLenNumber5
vSelectDoc = Trim(Me.CMBHeader.Text)
vDocNo1 = vLenNumber6

vQuery = "select    top 1 right(docno," & vLenNumber & ")+1 as docno" _
                    & " from bcnp.dbo.bcapinvoice  " _
                    & " where ltrim(left(docno," & vCountNum & ")) =   '" & vSelectDoc & "' and " _
                    & " left(right(docno,len(docno)-" & vCountNum & ")," & vLenNumber & ") = left(right('" & vDocNo1 & "',len('" & vDocNo & "')-" & vCountNum & ")," & vLenNumber & ")  order by docno desc "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vChangeDoc = Trim(vRecordset.Fields("docno").Value)
Else
    vChangeDoc = Format(1, "0000")
End If
vRecordset.Close

vChangeDoc1 = UCase(Format(vChangeDoc, "0000"))
vChangeDoc1 = vSelectDoc & vDocNo1 & "-" & vChangeDoc1

Me.LBLNewDocNo.Caption = vSelectDoc 'vChangeDoc1
Else
MsgBox "กรุณาใส่เลขที่เอกสารที่ต้องการปรับปรุงก่อนนะครับ"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD101_Click()

If CHK101.Value = 0 And CHK102.Value = 0 And CHK103.Value = 0 And CHK104.Value = 0 Then
MsgBox "กรุณาเลือกหัวข้อในการเปลี่ยนข้อมูลด้วยนะครับ"
Exit Sub
End If

If CHK101.Value = 1 And TXT104.Text = "" Then
    MsgBox "เลือกหัวเอกสารที่จะเปลี่ยนด้วยนะครับ"
    Exit Sub
End If

If CHK101.Value = 1 Then
        Call ChangeDocno
End If
If CHK102.Value = 1 Then
        If TXT105.Text <> "" Then
        Call ChangeTaxNo
        Call ChangeTaxNoBackOffice
        Else
        MsgBox "กรุณาใส่ข้อมูลในการเปลี่ยนเลขที่ใบกำกับภาษีด้วยครับ"
        End If
End If
If CHK103.Value = 1 Then
        If DTP101.Value <> Trim(TXT102.Text) Then
        Call ChangeDocDate
        Else
        MsgBox "กรุณาใส่ข้อมูลในการเปลี่ยนเอกสารด้วยครับ"
        End If
End If
If CHK104.Value = 1 Then
        If DTP102.Value <> Trim(TXT106.Text) Then
        Call ChangeTaxDate
        Call ChangeTaxDateBackOffice
        Else
        MsgBox "กรุณาใส่ข้อมูลในการเปลี่ยนวันที่ใบกำกับภาษีด้วยครับ"
        End If
End If

TXT101.Text = ""
TXT102.Text = ""
TXT103.Text = ""
TXT104.Text = ""
TXT105.Text = ""
TXT106.Text = ""
TXT101.SetFocus
CHK101.Value = 0
CHK102.Value = 0
CHK103.Value = 0
CHK104.Value = 0


End Sub

Public Sub CheckNumeric()
Dim vDocNo As String
Dim vText As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT101.Text)

For i = 1 To Len(TXT101.Text)
    If Mid(vDocNo, i, 1) = 0 Or Mid(vDocNo, i, 1) = 1 Or Mid(vDocNo, i, 1) = 2 Or Mid(vDocNo, i, 1) = 3 Or Mid(vDocNo, i, 1) = 4 Or Mid(vDocNo, i, 1) = 5 Or Mid(vDocNo, i, 1) = 6 Or Mid(vDocNo, i, 1) = 7 Or Mid(vDocNo, i, 1) = 8 Or Mid(vDocNo, i, 1) = 9 Then
        vCheckValue = True
        vCountNum = i - 1
        Exit Sub
    Else
        vCheckValue = False
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub CountDocnoNumeric(vDocNo As String)
Dim vText As String

On Error GoTo ErrDescription

For i = 1 To Len(vDocNo)
    If Mid(vDocNo, i, 1) = 0 Or Mid(vDocNo, i, 1) = 1 Or Mid(vDocNo, i, 1) = 2 Or Mid(vDocNo, i, 1) = 3 Or Mid(vDocNo, i, 1) = 4 Or Mid(vDocNo, i, 1) = 5 Or Mid(vDocNo, i, 1) = 6 Or Mid(vDocNo, i, 1) = 7 Or Mid(vDocNo, i, 1) = 8 Or Mid(vDocNo, i, 1) = 9 Then
        vCheckValue = True
        vCountNum = i - 1
        Exit Sub
    Else
        vCheckValue = False
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ChangeDocno()
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vChangeDoc As String
Dim vLenHeader As Integer
Dim vLenNumber As Integer
Dim vAnswer As Integer
Dim vTaxNo As String
Dim vDocdate As String
Dim vTaxDate As String
Dim vCheckNewDocNo As Integer

On Error GoTo ErrDescription

vDocNo = Trim(TXT101.Text)
If TXT104.Text <> "" Then
vChangeDoc1 = UCase(TXT104.Text)
Else
MsgBox "เลือกหัวเอกสารด้วยนะครับ"
Exit Sub
End If

Line1:

If Me.CBCreditNote.Value = 0 Then
    vQuery = "select  *  from bcnp.dbo.bcapinvoice where docno = '" & vChangeDoc1 & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckNewDocNo = 1
        MsgBox "เลขที่เอกสาร " & vChangeDoc1 & " มีอยู่แล้ว ต้องเปลี่ยนเป็นเลขที่ใหม่"
    Else
        vCheckNewDocNo = 0
    End If
    vRecordset.Close
Else
    vQuery = "select  *  from bcnp.dbo.bcstkrefund where docno = '" & vChangeDoc1 & "' "
    If OpenDataBase(gConnection, vRecordset1, vQuery) <> 0 Then
        vCheckNewDocNo = 1
        MsgBox "เลขที่เอกสาร " & vChangeDoc1 & " มีอยู่แล้ว ต้องเปลี่ยนเป็นเลขที่ใหม่"
    Else
        vCheckNewDocNo = 0
    End If
    vRecordset1.Close

End If

If vCheckNewDocNo = 1 Then
Call CMB101_Click
vChangeDoc1 = UCase(TXT104.Text)
GoTo Line1
End If

'vAnswer = MsgBox("คุณต้องเปลี่ยนเลขที่เอกสาร จากเลขที่ " & vDocNo & " เป็นเลขที่ " & vChangeDoc1 & " นี้ใช่หรือไม่", vbYesNo, "ข้อความสอบถาม")
'If vAnswer = 6 Then
    vQuery = "Exec bcnp.dbo.USP_AP_ChangeDocnoTaxData " & Me.CBCreditNote.Value & ", '" & vDocNo & "' ,'" & vChangeDoc1 & "' "
    gConnection.Execute vQuery
    MsgBox "ได้มีการเปลี่ยนเลขที่เอกสาร จากเลขที่ " & vDocNo & " เป็นเลขที่ " & vChangeDoc1 & " เรียบร้อยแล้ว"
'Else
 '   Exit Sub
'End If
vDocNo = UCase(vDocNo)
vTaxNo = Trim(TXT103.Text)
vDocdate = Trim(TXT102.Text)
vTaxDate = Trim(TXT106.Text)
vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
                    & " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate & "','" & vChangeDoc1 & "','','','','" & vUserID & "',getdate())"
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub ChangeDocnoXPC()
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vChangeDoc As String
Dim vLenHeader As Integer
Dim vLenNumber As Integer
Dim vAnswer As Integer
Dim vTaxNo As String
Dim vDocdate As String
Dim vTaxDate As String
Dim vCheckNewDocNo As Integer

On Error GoTo ErrDescription

vDocNo = Trim(TXT201.Text)
If TXT204.Text <> "" Then
vChangeDoc1 = UCase(TXT204.Text)
Else
MsgBox "เลือกหัวเอกสารด้วยนะครับ"
Exit Sub
End If

Line1:

If Me.CBCreditNote_XPC.Value = 0 Then
    vQuery = "select  *  from xpc.dbo.bcapinvoice where docno = '" & vChangeDoc1 & "' "
    If OpenDataBaseXPC(vXPCConnection, vRecordset, vQuery) <> 0 Then
        vCheckNewDocNo = 1
        MsgBox "เลขที่เอกสาร " & vChangeDoc1 & " มีอยู่แล้ว ต้องเปลี่ยนเป็นเลขที่ใหม่"
    Else
        vCheckNewDocNo = 0
    End If
    vRecordset.Close
Else
    vQuery = "select  *  from xpc.dbo.bcstkrefund where docno = '" & vChangeDoc1 & "' "
    If OpenDataBaseXPC(vXPCConnection, vRecordset1, vQuery) <> 0 Then
        vCheckNewDocNo = 1
        MsgBox "เลขที่เอกสาร " & vChangeDoc1 & " มีอยู่แล้ว ต้องเปลี่ยนเป็นเลขที่ใหม่"
    Else
        vCheckNewDocNo = 0
    End If
    vRecordset1.Close

End If

If vCheckNewDocNo = 1 Then
    Call CMB201_Click
    vChangeDoc1 = UCase(TXT204.Text)
    GoTo Line1
End If

vQuery = "Exec xpc.dbo.USP_AP_ChangeDocnoTaxData " & Me.CBCreditNote_XPC.Value & ", '" & vDocNo & "' ,'" & vChangeDoc1 & "' "
vXPCConnection.Execute vQuery
MsgBox "ได้มีการเปลี่ยนเลขที่เอกสาร จากเลขที่ " & vDocNo & " เป็นเลขที่ " & vChangeDoc1 & " เรียบร้อยแล้ว"

vDocNo = UCase(vDocNo)
vTaxNo = Trim(TXT203.Text)
vDocdate = Trim(TXT202.Text)
vTaxDate = Trim(TXT206.Text)
vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
                    & " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate & "','" & vChangeDoc1 & "','','','','" & vUserID & "-XPC" & "',getdate())"
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ChangeDocnoBackOffice()

End Sub

Private Sub CMD102_Click()
On Error Resume Next

TXT101.Text = ""
TXT102.Text = ""
TXT103.Text = ""
TXT104.Text = ""
TXT105.Text = ""
TXT106.Text = ""
CMB101.Text = ""
CHK101.Value = 0
CHK102.Value = 0
CHK103.Value = 0
CHK104.Value = 0

End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vRepID = 206
vRepType = "AP"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 206 and reptype = 'AP' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.WindowState = crptMaximized
.Destination = crptToWindow
.Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD201_Click()

If CHK201.Value = 0 And CHK202.Value = 0 And CHK203.Value = 0 And CHK204.Value = 0 Then
    MsgBox "กรุณาเลือกหัวข้อในการเปลี่ยนข้อมูลด้วยนะครับ"
    Exit Sub
End If

If CHK201.Value = 1 And TXT204.Text = "" Then
    MsgBox "เลือกหัวเอกสารที่จะเปลี่ยนด้วยนะครับ"
    Exit Sub
End If

If CHK201.Value = 1 Then
        Call ChangeDocnoXPC
End If
If CHK202.Value = 1 Then
        If TXT205.Text <> "" Then
            Call ChangeTaxNoXPC
            'Call ChangeTaxNoBackOffice
        Else
            MsgBox "กรุณาใส่ข้อมูลในการเปลี่ยนเลขที่ใบกำกับภาษีด้วยครับ"
        End If
End If
If CHK203.Value = 1 Then
        If DTP201.Value <> Trim(TXT202.Text) Then
            Call ChangeDocDateXPC
        Else
            MsgBox "กรุณาใส่ข้อมูลในการเปลี่ยนเอกสารด้วยครับ"
        End If
End If
If CHK204.Value = 1 Then
        If DTP202.Value <> Trim(TXT206.Text) Then
            Call ChangeTaxDateXPC
            'Call ChangeTaxDateBackOffice
        Else
            MsgBox "กรุณาใส่ข้อมูลในการเปลี่ยนวันที่ใบกำกับภาษีด้วยครับ"
        End If
End If

TXT201.Text = ""
TXT202.Text = ""
TXT203.Text = ""
TXT204.Text = ""
TXT205.Text = ""
TXT206.Text = ""
TXT201.SetFocus
CHK201.Value = 0
CHK202.Value = 0
CHK203.Value = 0
CHK204.Value = 0
End Sub

Private Sub CMD202_Click()
On Error Resume Next

TXT201.Text = ""
TXT202.Text = ""
TXT203.Text = ""
TXT204.Text = ""
TXT205.Text = ""
TXT206.Text = ""
CMB201.Text = ""
CHK201.Value = 0
CHK202.Value = 0
CHK203.Value = 0
CHK204.Value = 0
End Sub

Private Sub CMDChangeData_Click()
Dim vIndex As Integer
Dim i As Integer
Dim vDocdate As String
Dim vCheckDate As String
Dim vTaxDate As String
Dim vCheckTaxDate As String

Dim vDocdate1 As Date
Dim vCheckDate1 As Date
Dim vTaxDate1 As Date
Dim vCheckTaxDate1 As Date

On Error GoTo ErrDescription

If Me.LBLIndex.Caption <> "" Then
vIndex = Me.LBLIndex.Caption

If Me.CKDocNo.Value = 0 And Me.CKDocDate.Value = 0 And Me.CKTaxNo.Value = 0 And Me.CKTaxDate.Value = 0 Then
Me.CKDocNo.SetFocus
Exit Sub
End If

vDocdate = Me.LBLDocDate.Caption
vTaxDate = Me.LBLTaxDate.Caption

vCheckDate = Day(Me.DTPNewDocDate.Value) & "/" & Month(Me.DTPNewDocDate.Value) & "/" & Year(Me.DTPNewDocDate.Value)
vCheckTaxDate = Day(Me.DTPNewTaxDate.Value) & "/" & Month(Me.DTPNewTaxDate.Value) & "/" & Year(Me.DTPNewTaxDate.Value)

vDocdate1 = vDocdate
If vTaxDate <> "" Then
vTaxDate1 = vTaxDate
End If

vCheckDate1 = vCheckDate
vCheckTaxDate1 = vCheckTaxDate

If Me.CKDocDate.Value = 1 Then
If vDocdate1 = vCheckDate1 Then
MsgBox "วันที่เอกสารยังไม่ได้เปลี่ยนข้อมูล กรุณาตรวจสอบ"
Me.DTPNewDocDate.SetFocus
Exit Sub
End If
End If

If Me.CKTaxDate.Value = 1 Then
If vTaxDate1 = vCheckTaxDate1 Then
MsgBox "วันที่ภาษียังไม่ได้เปลี่ยนข้อมูล กรุณาตรวจสอบ"
Me.DTPNewTaxDate.SetFocus
Exit Sub
End If
End If

If Me.CKDocNo.Value = 1 Then
Me.ListViewDocNo.ListItems(vIndex).SubItems(2) = Me.LBLNewDocNo.Caption
Else
Me.ListViewDocNo.ListItems(vIndex).SubItems(2) = ""
End If

If Me.CKDocDate.Value = 1 Then
Me.ListViewDocNo.ListItems(vIndex).SubItems(4) = Day(Me.DTPNewDocDate.Value) & "/" & Month(Me.DTPNewDocDate.Value) & "/" & Year(Me.DTPNewDocDate.Value)
Else
Me.ListViewDocNo.ListItems(vIndex).SubItems(4) = ""
End If

If Me.CKTaxNo.Value = 1 Then
Me.ListViewDocNo.ListItems(vIndex).SubItems(6) = Me.TXTNewTaxNo.Text
Else
Me.ListViewDocNo.ListItems(vIndex).SubItems(6) = ""
End If

If Me.CKTaxDate.Value = 1 Then
Me.ListViewDocNo.ListItems(vIndex).SubItems(8) = Day(Me.DTPNewTaxDate.Value) & "/" & Month(Me.DTPNewTaxDate.Value) & "/" & Year(Me.DTPNewTaxDate.Value)
Else
Me.ListViewDocNo.ListItems(vIndex).SubItems(8) = ""
End If


Me.CKDocNo.Value = 0
Me.CKDocDate.Value = 0
Me.CKTaxNo.Value = 0
Me.CKTaxDate.Value = 0

For i = 1 To Me.ListViewDocNo.ListItems.Count
Me.ListViewDocNo.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000C0"
Me.ListViewDocNo.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000C0"
Me.ListViewDocNo.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000C0"
Me.ListViewDocNo.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000C0"
Next i

Me.PICEditData.Visible = False
Me.DTPSearchDoc.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDClose_Click()
Me.PICDocDate.Visible = False
Me.TXT101.SetFocus
End Sub

Private Sub CMDCloseEditData_Click()
Dim vIndex As Integer

On Error Resume Next

If Me.CKDocNo.Value = 0 And Me.CKDocDate.Value = 0 And Me.CKTaxNo.Value = 0 And Me.CKTaxDate.Value = 0 Then
vIndex = Me.LBLIndex.Caption

Me.ListViewDocNo.ListItems(vIndex).Checked = False

Me.ListViewDocNo.ListItems(vIndex).SubItems(2) = ""
Me.ListViewDocNo.ListItems(vIndex).SubItems(4) = ""
Me.ListViewDocNo.ListItems(vIndex).SubItems(6) = ""
Me.ListViewDocNo.ListItems(vIndex).SubItems(8) = ""

End If

Me.PICEditData.Visible = False
Me.ListViewDocNo.SetFocus
End Sub

Private Sub CMDCompany_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription


Me.PICCompany.Visible = False
If Me.CMBCompany.ListIndex = 0 Then
    Me.PICXPC.Visible = False
ElseIf Me.CMBCompany.ListIndex = 1 Then
    Me.PICXPC.Visible = True
    
    Call InitializeDataBaseXPC
    
    Me.CMB201.Clear
    'Me.CMBHeader_xpc.Clear
    
    vQuery = "select distinct left(docno,2) as docno from xpc.dbo.bcapinvoice where grbillstatus in (0,1) and grirbillstatus = 2 order by docno "
    If OpenDataBaseXPC(vXPCConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        CMB201.AddItem Trim(vRecordset.Fields("docno").Value)
        'Me.CMBHeader.AddItem Trim(vRecordset.Fields("docno").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
    
    
    Me.DTP201.Value = Now
    Me.DTP202.Value = Now

End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDDocDate_Click()
Me.PICDocDate.Visible = True
Call DTPSearchDoc_Change
Me.DTPSearchDoc.SetFocus
End Sub

Private Sub CMDExit_Click()
Me.PICCompany.Visible = True
Me.PICXPC.Visible = False
Me.CMBCompany.SetFocus
End Sub

Private Sub CMDSave_Click()
Dim i As Integer
Dim vCountItem As Integer
Dim vDocNo As String
Dim vDocdate As String
Dim vTaxNo As String
Dim vTaxDate As String
Dim vHeader As String
Dim vNewDocDate As String
Dim vNewTaxNo As String
Dim vNewTaxDate As String

On Error GoTo ErrDescription

If Me.CKBI.Value = 0 Then

    For i = 1 To Me.ListViewDocNo.ListItems.Count
    
    vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
    vDocdate = Me.ListViewDocNo.ListItems(i).SubItems(3)
    vTaxNo = Me.ListViewDocNo.ListItems(i).SubItems(5)
    vTaxDate = Me.ListViewDocNo.ListItems(i).SubItems(7)
    
    vHeader = Me.ListViewDocNo.ListItems(i).SubItems(2)
    vNewDocDate = Me.ListViewDocNo.ListItems(i).SubItems(4)
    vNewTaxNo = Me.ListViewDocNo.ListItems(i).SubItems(6)
    vNewTaxDate = Me.ListViewDocNo.ListItems(i).SubItems(8)
    
    If Me.ListViewDocNo.ListItems(i).SubItems(2) <> "" Then
            Call vInsertDoc(vDocNo, vDocdate, vHeader, vTaxNo, vTaxDate)
    End If
    
    If Me.ListViewDocNo.ListItems(i).SubItems(4) <> "" Then
            Call vInsertDate(vDocNo, vDocdate, vNewDocDate, vTaxNo, vTaxDate)
    End If
    
    If Me.ListViewDocNo.ListItems(i).SubItems(6) <> "" Then
            Call vInsertTax(vDocNo, vDocdate, vNewTaxNo, vTaxNo, vTaxDate)
            Call vInsertTaxBackOffice(vDocNo, vDocdate, vNewTaxNo, vTaxNo, vTaxDate)
    End If
    
    If Me.ListViewDocNo.ListItems(i).SubItems(8) <> "" Then
            Call vInsertTaxDate(vDocNo, vDocdate, vTaxNo, vNewTaxDate, vTaxDate)
            Call vInsertTaxDateBackOffice(vDocNo, vDocdate, vTaxNo, vNewTaxDate, vTaxDate)
    End If
    
    Next i
End If

If Me.CKBI.Value = 1 Then

    For i = 1 To Me.ListViewDocNo.ListItems.Count
    
    vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
    vDocdate = Me.ListViewDocNo.ListItems(i).SubItems(3)
    vTaxNo = Me.ListViewDocNo.ListItems(i).SubItems(5)
    vTaxDate = Me.ListViewDocNo.ListItems(i).SubItems(7)
    
    vHeader = Me.ListViewDocNo.ListItems(i).SubItems(2)
    vNewDocDate = Me.ListViewDocNo.ListItems(i).SubItems(4)
    vNewTaxNo = Me.ListViewDocNo.ListItems(i).SubItems(6)
    vNewTaxDate = Me.ListViewDocNo.ListItems(i).SubItems(8)
    
    If Me.ListViewDocNo.ListItems(i).SubItems(2) <> "" Then
            Call vInsertDocBI(vDocNo, vDocdate, vHeader, vTaxNo, vTaxDate)
    End If
    
    If Me.ListViewDocNo.ListItems(i).SubItems(4) <> "" Then
            Call vInsertDateBI(vDocNo, vDocdate, vNewDocDate, vTaxNo, vTaxDate)
    End If
    
    If Me.ListViewDocNo.ListItems(i).SubItems(6) <> "" Then
            Call vInsertTaxBI(vDocNo, vDocdate, vNewTaxNo, vTaxNo, vTaxDate)
    End If
    
    If Me.ListViewDocNo.ListItems(i).SubItems(8) <> "" Then
            Call vInsertTaxDateBI(vDocNo, vDocdate, vTaxNo, vNewTaxDate, vTaxDate)
    End If
    
    Next i
End If

MsgBox "บันทึกเปลี่ยนแปลงข้อมูลเรียบร้อยแล้ว", vbInformation, "Send Infromation Message"

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub vInsertDoc(vDocNo As String, vDocdate As String, vHeader As String, vTaxNo As String, vTaxDate As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vChangeDoc As String
Dim vLenHeader As Integer
Dim vLenNumber As Integer
Dim vCheckNewDocNo As Integer

On Error GoTo ErrDescription

Call vGetNewDocNo(vDocNo, vDocdate, vHeader)

vChangeDoc1 = vNewDocNo

vQuery = "select  *  from bcnp.dbo.bcapinvoice where docno = '" & vChangeDoc1 & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckNewDocNo = 1
Else
    vCheckNewDocNo = 0
End If
vRecordset.Close

If vCheckNewDocNo = 1 Then
Call vGetNewDocNo(vDocNo, vDocdate, vHeader)
End If

vQuery = "Exec bcnp.dbo.USP_AP_ChangeDocno '" & vDocNo & "' ,'" & vChangeDoc1 & "' "
gConnection.Execute vQuery

vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
                    & " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate & "','" & vChangeDoc1 & "','','','','" & vUserID & "',getdate())"
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub


Public Sub vInsertDocBI(vDocNo As String, vDocdate As String, vHeader As String, vTaxNo As String, vTaxDate As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vChangeDoc As String
Dim vLenHeader As Integer
Dim vLenNumber As Integer
Dim vCheckNewDocNo As Integer

On Error GoTo ErrDescription

Call vGetNewDocNoBI(vDocNo, vDocdate, vHeader)

vChangeDoc1 = vNewDocNo

vQuery = "select  *  from dbo.bcapinvoice where docno = '" & vChangeDoc1 & "' "
If OpenDataBaseBI(vBIConnection, vRecordset, vQuery) <> 0 Then
    vCheckNewDocNo = 1
Else
    vCheckNewDocNo = 0
End If
vRecordset.Close

If vCheckNewDocNo = 1 Then
Call vGetNewDocNoBI(vDocNo, vDocdate, vHeader)
End If

vQuery = "Exec dbo.USP_AP_ChangeDocno '" & vDocNo & "' ,'" & vChangeDoc1 & "' "
vBIConnection.Execute vQuery

vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
                    & " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate & "','" & vChangeDoc1 & "','','','','" & vUserID & "',getdate())"
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Public Sub vGetNewDocNo(vDocNo As String, vDocdate As String, vHeader As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vChangeDoc As String
Dim vLenHeader As Integer
Dim vLenNumber As Integer
Dim vAnswer As Integer
Dim vLenNumber1 As String, vLenNumber2 As String, vLenNumber3 As String
Dim vLenNumber4 As String, vLenNumber5 As String, vLenNumber6 As String
Dim vSelectDoc As String
Dim vDocNo1 As String

On Error GoTo ErrDescription

Call CountDocnoNumeric(vDocNo)

vLenNumber = (Len(vDocNo) - InStr(1, vDocNo, "-"))
vLenHeader = InStr(1, vDocNo, "-")
vLenNumber1 = Year(vDocdate)
If Year(vDocdate) < 2500 Then
vLenNumber1 = Year(vDocdate) + 543
Else
vLenNumber1 = Year(vDocdate)
End If
vLenNumber2 = Right(vLenNumber1, 2)
vLenNumber3 = Month(vDocdate)
vLenNumber4 = Len(vLenNumber3)
If vLenNumber4 < 2 Then
vLenNumber5 = "0" & vLenNumber3
Else
vLenNumber5 = vLenNumber3
End If
vLenNumber6 = vLenNumber2 & vLenNumber5
vSelectDoc = Trim(vHeader)
vDocNo1 = vLenNumber6

vQuery = "select    top 1 right(docno," & vLenNumber & ")+1 as docno" _
                    & " from bcnp.dbo.bcapinvoice  " _
                    & " where ltrim(left(docno," & vCountNum & ")) =   '" & vSelectDoc & "' and " _
                    & " left(right(docno,len(docno)-" & vCountNum & ")," & vLenNumber & ") = left(right('" & vDocNo1 & "',len('" & vDocNo & "')-" & vCountNum & ")," & vLenNumber & ")  order by docno desc "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vChangeDoc = Trim(vRecordset.Fields("docno").Value)
Else
    vChangeDoc = Format(1, "0000")
End If
vRecordset.Close

vChangeDoc1 = UCase(Format(vChangeDoc, "0000"))
vChangeDoc1 = vSelectDoc & vDocNo1 & "-" & vChangeDoc1

vNewDocNo = vChangeDoc1


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub vGetNewDocNoBI(vDocNo As String, vDocdate As String, vHeader As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vChangeDoc As String
Dim vLenHeader As Integer
Dim vLenNumber As Integer
Dim vAnswer As Integer
Dim vLenNumber1 As String, vLenNumber2 As String, vLenNumber3 As String
Dim vLenNumber4 As String, vLenNumber5 As String, vLenNumber6 As String
Dim vSelectDoc As String
Dim vDocNo1 As String

On Error GoTo ErrDescription

Call CountDocnoNumeric(vDocNo)

vLenNumber = (Len(vDocNo) - InStr(1, vDocNo, "-"))
vLenHeader = InStr(1, vDocNo, "-")
vLenNumber1 = Year(vDocdate)
If Year(vDocdate) < 2500 Then
vLenNumber1 = Year(vDocdate) + 543
Else
vLenNumber1 = Year(vDocdate)
End If
vLenNumber2 = Right(vLenNumber1, 2)
vLenNumber3 = Month(vDocdate)
vLenNumber4 = Len(vLenNumber3)
If vLenNumber4 < 2 Then
vLenNumber5 = "0" & vLenNumber3
Else
vLenNumber5 = vLenNumber3
End If
vLenNumber6 = vLenNumber2 & vLenNumber5
vSelectDoc = Trim(vHeader)
vDocNo1 = vLenNumber6

vQuery = "select    top 1 right(docno," & vLenNumber & ")+1 as docno" _
                    & " from dbo.bcapinvoice  " _
                    & " where ltrim(left(docno," & vCountNum & ")) =   '" & vSelectDoc & "' and " _
                    & " left(right(docno,len(docno)-" & vCountNum & ")," & vLenNumber & ") = left(right('" & vDocNo1 & "',len('" & vDocNo & "')-" & vCountNum & ")," & vLenNumber & ")  order by docno desc "
If OpenDataBaseBI(vBIConnection, vRecordset, vQuery) <> 0 Then
    vChangeDoc = Trim(vRecordset.Fields("docno").Value)
Else
    vChangeDoc = Format(1, "0000")
End If
vRecordset.Close


vChangeDoc1 = UCase(Format(vChangeDoc, "0000"))
vChangeDoc1 = vSelectDoc & vDocNo1 & "-" & vChangeDoc1

vNewDocNo = vChangeDoc1


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub vInsertDate(vDocNo As String, vDocdate As String, vNewDocDate As String, vTaxNo As String, vTaxDate As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

vQuery = "Exec bcnp.dbo.USP_AP_ChangeDocDate '" & vDocNo & "','" & vNewDocDate & "' "
gConnection.Execute vQuery

vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
& " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate & "','','" & vNewDocDate & "','','','" & vUserID & "',getdate())"
gConnection.Execute vQuery
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub vInsertDateBI(vDocNo As String, vDocdate As String, vNewDocDate As String, vTaxNo As String, vTaxDate As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

vQuery = "Exec dbo.USP_AP_ChangeDocDate '" & vDocNo & "','" & vNewDocDate & "' "
vBIConnection.Execute vQuery

vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
& " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate & "','','" & vNewDocDate & "','','','" & vUserID & "',getdate())"
gConnection.Execute vQuery
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub vInsertTax(vDocNo As String, vDocdate As String, vNewTaxNo As String, vTaxNo As String, vTaxDate As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckTax As Integer

On Error GoTo ErrDescription

vQuery = "select docno from bcnp.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then
    vQuery = "Exec bcnp.dbo.USP_AP_ChangeTaxno '" & vDocNo & "','" & vNewTaxNo & "' "
    gConnection.Execute vQuery
End If

vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
& " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate & "','','','" & vNewTaxNo & "','','" & vUserID & "',getdate())"
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub vInsertTaxBI(vDocNo As String, vDocdate As String, vNewTaxNo As String, vTaxNo As String, vTaxDate As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckTax As Integer

On Error GoTo ErrDescription

vQuery = "select docno from dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBaseBI(vBIConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then
    vQuery = "Exec dbo.USP_AP_ChangeTaxno '" & vDocNo & "','" & vNewTaxNo & "' "
    vBIConnection.Execute vQuery
End If

vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
& " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate & "','','','" & vNewTaxNo & "','','" & vUserID & "',getdate())"
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub vInsertTaxBackOffice(vDocNo As String, vDocdate As String, vNewTaxNo As String, vTaxNo As String, vTaxDate As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckTax As Integer
Dim vVatRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

vQuery = "select docno from bi.bcvat.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then

      vQuery = "set dateformat dmy "
      gConnection.Execute vQuery
          
      vQuery = "Update  bi.bcvat.dbo.bcapinvoice set taxno = '" & vNewTaxNo & "'  where iscancel = 0 and docno =  '" & vDocNo & "'"
      gConnection.Execute vQuery
      
      vQuery = "Update  bi.bcvat.dbo.bcinputtax set taxno = '" & vNewTaxNo & "'  where iscancel = 0 and docno =  '" & vDocNo & "'"
      gConnection.Execute vQuery
      
      vQuery = "Update  bi.bcvat.dbo.bcirsub set taxno = '" & vNewTaxNo & "'  where iscancel = 0 and docno =  '" & vDocNo & "'"
      gConnection.Execute vQuery

End If

vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment   set mydescription = 'Update BCVat'  where OldDocNo = '" & vDocNo & "' and OldTaxNo = '" & vTaxNo & "' and NewTaxNo = '" & vNewTaxNo & "' "
gConnection.Execute vQuery


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub vInsertTaxDate(vDocNo As String, vDocdate As String, vTaxNo As String, vNewTaxDate As String, vTaxDate As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckTax As Integer

On Error GoTo ErrDescription

vQuery = "select docno from bcnp.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then

vQuery = "Exec bcnp.dbo.USP_AP_ChangeTaxDate '" & vDocNo & "','" & vTaxNo & "','" & vNewTaxDate & "' "
gConnection.Execute vQuery

vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
& " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate & "','','','','" & vNewTaxDate & "','" & vUserID & "',getdate())"
gConnection.Execute vQuery

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub vInsertTaxDateBI(vDocNo As String, vDocdate As String, vTaxNo As String, vNewTaxDate As String, vTaxDate As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckTax As Integer

On Error GoTo ErrDescription

vQuery = "select docno from dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBaseBI(vBIConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then

vQuery = "Exec dbo.USP_AP_ChangeTaxDate '" & vDocNo & "','" & vTaxNo & "','" & vNewTaxDate & "' "
vBIConnection.Execute vQuery

vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
& " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate & "','','','','" & vNewTaxDate & "','" & vUserID & "',getdate())"
gConnection.Execute vQuery

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Private Sub CMDTaxAll_Click()
Dim i As Integer
Dim vTaxDate As String

For i = 1 To Me.ListViewDocNo.ListItems.Count
If Me.ListViewDocNo.ListItems(i).Checked = True Then
Me.ListViewDocNo.ListItems(i).SubItems(8) = Day(Me.DTPTaxDateAll.Value) & "/" & Month(Me.DTPTaxDateAll.Value) & "/" & Year(Me.DTPTaxDateAll.Value)
End If
Next i
End Sub



Private Sub DTPSearchDoc_Change()
Dim vQuery As String
Dim vRecordset As New Recordset
Dim vDocdate As String
Dim i As Integer
Dim vListDocno As ListItem
Dim vDocNo As String
Dim n As Integer

On Error Resume Next

vDocdate = Day(Me.DTPSearchDoc.Value) & "/" & Month(Me.DTPSearchDoc.Value) & "/" & Year(Me.DTPSearchDoc.Value)

Me.ListViewDocNo.ListItems.Clear
i = 1

If Me.CKBI.Value = 0 Then
    vQuery = "exec dbo.USP_NP_SearchDocTaxNo '" & vDocdate & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        
        While Not vRecordset.EOF
        Set vListDocno = Me.ListViewDocNo.ListItems.Add(, , i)
        vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
        vListDocno.SubItems(2) = ""
        vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
        vListDocno.SubItems(4) = ""
        vListDocno.SubItems(5) = Trim(vRecordset.Fields("taxno").Value)
        vListDocno.SubItems(6) = ""
        If Trim(vRecordset.Fields("taxdate").Value) <> "01/01/1900" Then
        vListDocno.SubItems(7) = Trim(vRecordset.Fields("taxdate").Value)
        Else
        vListDocno.SubItems(7) = ""
        End If
        vListDocno.SubItems(8) = ""
    
        vRecordset.MoveNext
        i = i + 1
        Wend
    Else
        Me.ListViewDocNo.ListItems.Clear
        Exit Sub
    End If
    vRecordset.Close
End If


If Me.CKBI.Value = 1 Then
    vQuery = "exec dbo.USP_NP_SearchDocTaxNo '" & vDocdate & "' "
    If OpenDataBaseBI(vBIConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        
        While Not vRecordset.EOF
        Set vListDocno = Me.ListViewDocNo.ListItems.Add(, , i)
        vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
        vListDocno.SubItems(2) = ""
        vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
        vListDocno.SubItems(4) = ""
        vListDocno.SubItems(5) = Trim(vRecordset.Fields("taxno").Value)
        vListDocno.SubItems(6) = ""
        If Trim(vRecordset.Fields("taxdate").Value) <> "01/01/1900" Then
        vListDocno.SubItems(7) = Trim(vRecordset.Fields("taxdate").Value)
        Else
        vListDocno.SubItems(7) = ""
        End If
        vListDocno.SubItems(8) = ""
    
        vRecordset.MoveNext
        i = i + 1
        Wend
    Else
        Me.ListViewDocNo.ListItems.Clear
        Exit Sub
    End If
    vRecordset.Close
End If
    
End Sub


Public Sub vGetData()
Dim vQuery As String
Dim vRecordset As New Recordset
Dim vDocdate As String
Dim i As Integer
Dim vListDocno As ListItem
Dim vDocNo As String
Dim n As Integer

On Error Resume Next

vDocdate = Day(Me.DTPSearchDoc.Value) & "/" & Month(Me.DTPSearchDoc.Value) & "/" & Year(Me.DTPSearchDoc.Value)

Me.ListViewDocNo.ListItems.Clear
i = 1
vQuery = "exec dbo.USP_NP_SearchDocApNo '" & vDocdate & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    
    While Not vRecordset.EOF
    Set vListDocno = Me.ListViewDocNo.ListItems.Add(, , i)
    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
    vListDocno.SubItems(2) = ""
    vListDocno.SubItems(3) = Trim(vRecordset.Fields("docdate").Value)
    vListDocno.SubItems(4) = ""
    vListDocno.SubItems(5) = Trim(vRecordset.Fields("taxno").Value)
    vListDocno.SubItems(6) = ""
    If Trim(vRecordset.Fields("taxdate").Value) <> "01/01/1900" Then
    vListDocno.SubItems(7) = Trim(vRecordset.Fields("taxdate").Value)
    Else
    vListDocno.SubItems(7) = ""
    End If
    vListDocno.SubItems(8) = ""

    vRecordset.MoveNext
    i = i + 1
    Wend
Else
    Me.ListViewDocNo.ListItems.Clear
    Exit Sub
End If
vRecordset.Close
End Sub


Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

Me.CMB101.Clear
Me.CMBHeader.Clear

vQuery = "select distinct left(docno,2) as docno from bcnp.dbo.bcapinvoice where grbillstatus in (0,1) and grirbillstatus = 2 order by docno "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMB101.AddItem Trim(vRecordset.Fields("docno").Value)
    Me.CMBHeader.AddItem Trim(vRecordset.Fields("docno").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

Call InitializeDataBaseVat

Me.DTPSearchDoc.Value = Now
Me.DTPNewDocDate.Value = Now
Me.DTPNewTaxDate.Value = Now
Me.DTP101.Value = Now
Me.DTP102.Value = Now
Me.DTPTaxDateAll.Value = Now

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub


Private Sub ListViewDocNo_DblClick()
Dim vIndex As Integer

On Error GoTo ErrDescription

If Me.ListViewDocNo.ListItems.Count > 0 Then
    vIndex = Me.ListViewDocNo.SelectedItem.Index
    
        Me.LBLIndex.Caption = vIndex
    
        Me.LBLDocNo.Caption = Me.ListViewDocNo.ListItems(vIndex).SubItems(1)
        If Me.ListViewDocNo.ListItems(vIndex).SubItems(2) <> "" Then
        Me.LBLNewDocNo.Caption = Me.ListViewDocNo.ListItems(vIndex).SubItems(2)
        Me.CKDocNo.Value = 1
        Else
        Me.LBLNewDocNo.Caption = ""
        Me.CKDocNo.Value = 0
        End If
        
        Me.LBLDocDate.Caption = Me.ListViewDocNo.ListItems(vIndex).SubItems(3)
        
        If Me.ListViewDocNo.ListItems(vIndex).SubItems(4) <> "" Then
        Me.DTPNewDocDate.Enabled = True
        Me.DTPNewDocDate.Value = Me.ListViewDocNo.ListItems(vIndex).SubItems(4)
        Me.CKDocDate.Value = 1
        Else
        Me.DTPNewDocDate.Enabled = False
        Me.DTPNewDocDate.Value = Now
        Me.CKDocDate.Value = 0
        End If
        
        Me.LBLTaxNo.Caption = Me.ListViewDocNo.ListItems(vIndex).SubItems(5)
        
        If Me.ListViewDocNo.ListItems(vIndex).SubItems(6) <> "" Then
        Me.TXTNewTaxNo.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(6)
        Me.CKTaxNo.Value = 1
        Else
        Me.TXTNewTaxNo.Text = ""
        Me.CKTaxNo.Value = 0
        End If
        
        Me.LBLTaxDate.Caption = Me.ListViewDocNo.ListItems(vIndex).SubItems(7)
                
        If Me.ListViewDocNo.ListItems(vIndex).SubItems(8) <> "" Then
        Me.DTPNewTaxDate.Enabled = True
        Me.DTPNewTaxDate.Value = Me.ListViewDocNo.ListItems(vIndex).SubItems(8)
        Me.CKTaxDate.Value = 1
        Else
        Me.DTPNewTaxDate.Enabled = False
        Me.DTPNewTaxDate.Value = Now
        Me.CKTaxDate.Value = 0
        End If
            
        Me.PICEditData.Visible = True

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListViewDocNo_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vIndex As Integer

vIndex = Item.Index
Me.ListViewDocNo.ListItems(vIndex).SubItems(2) = ""
Me.ListViewDocNo.ListItems(vIndex).SubItems(4) = ""
Me.ListViewDocNo.ListItems(vIndex).SubItems(6) = ""
Me.ListViewDocNo.ListItems(vIndex).SubItems(8) = ""

'If Me.ListViewDocNo.ListItems.Count > 0 Then
'
 '   vIndex = Item.Index
  '  If Me.ListViewDocNo.ListItems(vIndex).Checked = True Then
   ' Me.LBLIndex.Caption = vIndex
'
 '   Me.LBLDocNo.Caption = Me.ListViewDocNo.ListItems(vIndex).SubItems(1)
  '  Me.LBLNewDocNo.Caption = Me.ListViewDocNo.ListItems(vIndex).SubItems(2)
   ' Me.LBLDocDate.Caption = Me.ListViewDocNo.ListItems(vIndex).SubItems(3)
    'If Me.ListViewDocNo.ListItems(vIndex).SubItems(4) <> "" Then
    'Me.DTPNewDocDate.Enabled = True
    'Me.DTPNewDocDate.Value = Me.ListViewDocNo.ListItems(vIndex).SubItems(4)
    'End If
    'Me.LBLTaxNo.Caption = Me.ListViewDocNo.ListItems(vIndex).SubItems(5)
    'Me.TXTNewTaxNo.Text = Me.ListViewDocNo.ListItems(vIndex).SubItems(6)
    'Me.LBLTaxDate.Caption = Me.ListViewDocNo.ListItems(vIndex).SubItems(7)
    'If Me.ListViewDocNo.ListItems(vIndex).SubItems(8) <> "" Then
    'Me.DTPNewTaxDate.Enabled = True
    'Me.DTPNewTaxDate.Value = Me.ListViewDocNo.ListItems(vIndex).SubItems(8)
    'End If
        
    'Me.PICEditData.Visible = True
    'Else
    'Me.LBLDocNo.Caption = ""
    'Me.LBLDocDate.Caption = ""
    'Me.LBLTaxNo.Caption = ""
    'Me.LBLTaxDate.Caption = ""
     '
    'Me.ListViewDocNo.ListItems(vIndex).SubItems(2) = ""
    'Me.ListViewDocNo.ListItems(vIndex).SubItems(4) = ""
    'Me.ListViewDocNo.ListItems(vIndex).SubItems(6) = ""
    'Me.ListViewDocNo.ListItems(vIndex).SubItems(8) = ""

'End If
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub TXT101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If TXT101.Text <> "" Then
        vDocNo = TXT101.Text
        Call CheckNumeric
        
        'If Me.CKBI.Value = 0 Then

            vQuery = " exec dbo.USP_NP_SearchDocnoEditTaxData " & Me.CBCreditNote.Value & ",'" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                TXT102.Text = Trim(vRecordset.Fields("docdate").Value)
                    If Not IsNull(Trim(vRecordset.Fields("taxno").Value)) Then
                        TXT103.Text = Trim(vRecordset.Fields("taxno").Value)
                    Else
                        TXT103.Text = "NoTaxNo"
                    End If
                    If Not IsNull(Trim(vRecordset.Fields("taxdate").Value)) Then
                        TXT106.Text = Trim(vRecordset.Fields("taxdate").Value)
                    Else
                        TXT106.Text = Trim(vRecordset.Fields("docdate").Value)
                    End If
                Else
                MsgBox "เอกสาร เลขที่ " & vDocNo & " ไม่สามารถเปลี่ยนเอกสารได้ โปรดตรวจสอบด้วยนะครับ"
                TXT104.Text = ""
                TXT105.Text = ""
                Exit Sub
            End If
            vRecordset.Close
            'End If
    
        'If Me.CKBI.Value = 1 Then
        
         '       Call InitializeDataBaseBI
          '      vQuery = "select a.docno,a.docdate,b.taxno,b.taxdate from dbo.bcapinvoice a " _
           '     & " left join dbo.bcinputtax b on a.docno = b.docno " _
            '    & " where a.docno = '" & vDocNo & "' and a.grbillstatus in (0,1) and a.grirbillstatus in (0,2) and a.iscancel = 0"
             '   If OpenDataBaseBI(vBIConnection, vRecordset, vQuery) <> 0 Then
              '      TXT102.Text = Trim(vRecordset.Fields("docdate").Value)
               '         If Not IsNull(Trim(vRecordset.Fields("taxno").Value)) Then
                '            TXT103.Text = Trim(vRecordset.Fields("taxno").Value)
                 '       Else
                  '          TXT103.Text = "NoTaxNo"
                   '     End If
                    '    If Not IsNull(Trim(vRecordset.Fields("taxdate").Value)) Then
                     '       TXT106.Text = Trim(vRecordset.Fields("taxdate").Value)
                      '  Else
                       '     TXT106.Text = Trim(vRecordset.Fields("docdate").Value)
                        'End If
                    'Else
                    'MsgBox "เอกสาร เลขที่ " & vDocNo & " ไม่สามารถเปลี่ยนเอกสารได้ โปรดตรวจสอบด้วยนะครับ"
                    'TXT104.Text = ""
                    'TXT105.Text = ""
                    'Exit Sub
                'End If
                'vRecordset.Close
                'End If
                
            TXT104.Text = ""
            TXT105.Text = ""
            DTP101.Value = TXT102.Text
            DTP102.Value = TXT106.Text
        End If
        
    End If
                       


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ChangeTaxNo()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vTaxNo As String, vTaxNo1 As String
Dim vCheckTax As Integer
Dim vDocdate As String, vTaxDate As String

On Error GoTo ErrDescription

If CHK101.Value = 0 Then
    vDocNo = Trim(TXT101.Text)
Else
    vDocNo = vChangeDoc1
End If

vQuery = "select docno from bcnp.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then
    vTaxNo = Trim(TXT105.Text)
    vTaxNo1 = Trim(TXT103.Text)
    If vTaxNo <> "" Then
        vQuery = "Exec bcnp.dbo.USP_AP_ChangeTaxnoTaxData " & Me.CBCreditNote.Value & ",'" & vDocNo & "','" & vTaxNo & "' "
        gConnection.Execute vQuery
        MsgBox "เอกสารเลขที่ " & vDocNo & " ได้แก้ไข ใบกำกับภาษีเป็นเลขที่ " & vTaxNo & " เรียบร้อยแล้วครับ"
    Else
        MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนใบกำกับภาษี"
    End If
Else
    MsgBox "ไม่มีข้อมูลเลขที่เอกสาร " & vDocNo & " ในตาราง BCInputtax กรุณาตรวจสอบด้วยครับ"
End If

If CHK101.Value = 1 Then
    vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set OldTaxNo = '" & vTaxNo1 & "',NewTaxNo = '" & vTaxNo & "' where newdocno = '" & vDocNo & "' and oldtaxno = '" & vTaxNo1 & "' "
    gConnection.Execute vQuery
Else
    vDocNo = UCase(vDocNo)
    vDocdate = Trim(TXT102.Text)
    vTaxDate = Trim(TXT106.Text)
    vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
    & " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo1 & "','" & vTaxDate & "','','','" & vTaxNo & "','','" & vUserID & "',getdate())"
    gConnection.Execute vQuery
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ChangeTaxNoXPC()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vTaxNo As String, vTaxNo1 As String
Dim vCheckTax As Integer
Dim vDocdate As String, vTaxDate As String

On Error GoTo ErrDescription

If CHK201.Value = 0 Then
    vDocNo = Trim(TXT201.Text)
Else
    vDocNo = vChangeDoc1
End If

vQuery = "select docno from xpc.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBaseXPC(vXPCConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then
    vTaxNo = Trim(TXT205.Text)
    vTaxNo1 = Trim(TXT203.Text)
    If vTaxNo <> "" Then
        vQuery = "exec  xpc.dbo.USP_AP_ChangeTaxnoTaxData " & Me.CBCreditNote_XPC.Value & ",'" & vDocNo & "','" & vTaxNo & "' "
        vXPCConnection.Execute vQuery
        MsgBox "เอกสารเลขที่ " & vDocNo & " ได้แก้ไข ใบกำกับภาษีเป็นเลขที่ " & vTaxNo & " เรียบร้อยแล้วครับ"
    Else
        MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนใบกำกับภาษี"
    End If
Else
    MsgBox "ไม่มีข้อมูลเลขที่เอกสาร " & vDocNo & " ในตาราง BCInputtax กรุณาตรวจสอบด้วยครับ"
End If

If CHK201.Value = 1 Then
    vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set OldTaxNo = '" & vTaxNo1 & "',NewTaxNo = '" & vTaxNo & "' where newdocno = '" & vDocNo & "' and oldtaxno = '" & vTaxNo1 & "' "
    gConnection.Execute vQuery
Else
    vDocNo = UCase(vDocNo)
    vDocdate = Trim(TXT202.Text)
    vTaxDate = Trim(TXT206.Text)
    vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
    & " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo1 & "','" & vTaxDate & "','','','" & vTaxNo & "','','" & vUserID & "-XPC" & "',getdate())"
    gConnection.Execute vQuery
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub ChangeTaxNoBackOffice()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vTaxNo As String, vTaxNo1 As String
Dim vCheckTax As Integer
Dim vDocdate As String, vTaxDate As String
Dim vVatRecordset As New ADODB.Recordset


On Error GoTo ErrDescription

If CHK101.Value = 0 Then
    vDocNo = Trim(TXT101.Text)
Else
    vDocNo = vChangeDoc1
End If

vQuery = "select docno from bi.bcvat.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then
    vTaxNo = Trim(TXT105.Text)
    vTaxNo1 = Trim(TXT103.Text)
    If vTaxNo <> "" Then
      vQuery = "set dateformat dmy "
      gConnection.Execute vQuery
          
      vQuery = "Update  bi.bcvat.dbo.bcapinvoice set taxno = '" & vTaxNo & "'  where iscancel = 0 and docno =  '" & vDocNo & "'"
      gConnection.Execute vQuery
      
      vQuery = "Update  bi.bcvat.dbo.bcinputtax set taxno = '" & vTaxNo & "'  where iscancel = 0 and docno =  '" & vDocNo & "'"
      'vVatConnection.Execute vQuery
      gConnection.Execute vQuery
      
      vQuery = "Update  bi.bcvat.dbo.bcirsub set taxno = '" & vTaxNo & "'  where iscancel = 0 and docno =  '" & vDocNo & "'"
      'vVatConnection.Execute vQuery
      gConnection.Execute vQuery
    
      MsgBox "เอกสารเลขที่ " & vDocNo & " ได้แก้ไข ใบกำกับภาษีเป็นเลขที่ " & vTaxNo & " เรียบร้อยแล้วครับ ที่ข้อมูลหลังร้าน"
      
    Else
        MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนใบกำกับภาษี"
    End If
Else
    MsgBox "ไม่มีข้อมูลเลขที่เอกสาร " & vDocNo & " ในตาราง BCInputtax กรุณาตรวจสอบด้วยครับ"
End If


vDocNo = UCase(vDocNo)
vDocdate = Trim(TXT102.Text)
vTaxDate = Trim(TXT106.Text)
vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment   set mydescription = 'Update BCVat'  where OldDocNo = '" & vDocNo & "' and OldTaxNo = '" & vTaxNo1 & "' and NewTaxNo = '" & vTaxNo & "' "
gConnection.Execute vQuery


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub ChangeTaxDateBackOffice()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vTaxDate As String
Dim vTaxDate1 As String
Dim vCheckTax As Integer
Dim vTaxNo As String, vDocdate As String, vDocdate1 As String
Dim vVatRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If CHK101.Value = 0 Then
    vDocNo = Trim(TXT101.Text)
Else
    vDocNo = vChangeDoc1
End If

If CHK102.Value = 0 Then
    vTaxNo = Trim(TXT103.Text)
Else
    vTaxNo = Trim(TXT105.Text)
End If

'gConnection.Execute vQuery
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            
vQuery = "select docno from bi.bcvat.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
'If OpenDataBaseBCVat(vVatConnection, vVatRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then
    vTaxDate1 = Trim(TXT106.Text)
    vTaxDate = DTP102.Day & "/" & DTP102.Month & "/" & DTP102.Year
        If vTaxDate <> vTaxDate1 Then
        
            vQuery = "set dateformat dmy"
            'vVatConnection.Execute vQuery
            gConnection.Execute vQuery
        
            vQuery = "update  bi.bcvat.dbo.bcinputtax set  taxdate = '" & vTaxDate & "' where  iscancel = 0 and docno = '" & vDocNo & "' and taxno = '" & vTaxNo & "' "
            'vVatConnection.Execute vQuery
            gConnection.Execute vQuery
            MsgBox "เลขที่เอกสาร " & vDocNo & " ได้แก้ไขวันที่ใบกำกับภาษีเป็นวันที่ " & vTaxDate & " เรียบร้อยแล้ว  ที่ข้อมูลหลังร้าน"
        Else
            MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนวันที่เอกสาร"
            Exit Sub
        End If
    

vDocNo = UCase(vDocNo)
vDocdate = Trim(TXT102.Text)
vTaxNo = Trim(TXT103.Text)
vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment  set mydescription = 'Update BCVat'   where  OldDocNo = '" & vDocNo & "'  and OldTaxNo = '" & vTaxNo & "' and oldTaxDate = '" & vTaxDate1 & "' and NewTaxDate = '" & vTaxDate & "' "
gConnection.Execute vQuery


    Else
    MsgBox "ไม่มีข้อมูลเลขที่เอกสาร " & vDocNo & " ในตาราง BCInputtax กรุณาตรวจสอบด้วยครับ"
    End If
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub ChangeDocDate()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vDocdate As String
Dim vDocdate1 As String
Dim vTaxNo As String
Dim vTaxDate As String

On Error GoTo ErrDescription

If CHK101.Value = 0 Then
    vDocNo = Trim(TXT101.Text)
Else
    vDocNo = vChangeDoc1
End If
vTaxNo = Trim(TXT103.Text)
vTaxDate = Trim(TXT106.Text)
vDocdate1 = Trim(TXT102.Text)
vDocdate = DTP101.Day & "/" & DTP101.Month & "/" & DTP101.Year
    If vDocdate <> vDocdate1 Then
        vQuery = "Exec bcnp.dbo.USP_AP_ChangeDocDateTaxData " & Me.CBCreditNote.Value & ", '" & vDocNo & "','" & vDocdate & "' "
        gConnection.Execute vQuery
        MsgBox "เลขที่เอกสาร " & vDocNo & " ได้แก้ไขวันที่เอกสารเป็นวันที่ " & vDocdate & " เรียบร้อยแล้ว  "
    Else
        MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนวันที่เอกสาร"
        Exit Sub
    End If

    If CHK101.Value = 1 Then
        vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set OldDocdate = '" & vDocdate1 & "',newdocdate = '" & vDocdate & "' where newdocno = '" & vDocNo & "' "
        gConnection.Execute vQuery
    ElseIf CHK102.Value = 1 Then
        vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set OldDocdate = '" & vDocdate1 & "',newdocdate = '" & vDocdate & "' where olddocno = '" & vDocNo & "' and oldtaxno = '" & vTaxNo & "' "
        gConnection.Execute vQuery
    Else
        vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
        & " values('" & vDocNo & "','" & vDocdate1 & "','" & vTaxNo & "','" & vTaxDate & "','','" & vDocdate & "','','','" & vUserID & "',getdate())"
        gConnection.Execute vQuery
    End If
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ChangeDocDateXPC()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vDocdate As String
Dim vDocdate1 As String
Dim vTaxNo As String
Dim vTaxDate As String

On Error GoTo ErrDescription

If CHK201.Value = 0 Then
    vDocNo = Trim(TXT201.Text)
Else
    vDocNo = vChangeDoc1
End If
vTaxNo = Trim(TXT203.Text)
vTaxDate = Trim(TXT206.Text)
vDocdate1 = Trim(TXT202.Text)
vDocdate = DTP201.Day & "/" & DTP201.Month & "/" & DTP201.Year
    If vDocdate <> vDocdate1 Then
        vQuery = "Exec xpc.dbo.USP_AP_ChangeDocDateTaxData " & Me.CBCreditNote_XPC.Value & ", '" & vDocNo & "','" & vDocdate & "' "
        vXPCConnection.Execute vQuery
        MsgBox "เลขที่เอกสาร " & vDocNo & " ได้แก้ไขวันที่เอกสารเป็นวันที่ " & vDocdate & " เรียบร้อยแล้ว  "
    Else
        MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนวันที่เอกสาร"
        Exit Sub
    End If

    If CHK201.Value = 1 Then
        vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set OldDocdate = '" & vDocdate1 & "',newdocdate = '" & vDocdate & "' where newdocno = '" & vDocNo & "' "
        gConnection.Execute vQuery
    ElseIf CHK202.Value = 1 Then
        vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set OldDocdate = '" & vDocdate1 & "',newdocdate = '" & vDocdate & "' where olddocno = '" & vDocNo & "' and oldtaxno = '" & vTaxNo & "' "
        gConnection.Execute vQuery
    Else
        vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
        & " values('" & vDocNo & "','" & vDocdate1 & "','" & vTaxNo & "','" & vTaxDate & "','','" & vDocdate & "','','','" & vUserID & "',getdate())"
        gConnection.Execute vQuery
    End If
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ChangeDocDateBackOffice()

End Sub

Public Sub ChangeTaxDate()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vTaxDate As String
Dim vTaxDate1 As String
Dim vCheckTax As Integer
Dim vTaxNo As String
Dim vDocdate As String
Dim vDocdate1 As String

On Error GoTo ErrDescription

If CHK101.Value = 0 Then
    vDocNo = Trim(TXT101.Text)
Else
    vDocNo = vChangeDoc1
End If

If CHK102.Value = 0 Then
    vTaxNo = Trim(TXT103.Text)
Else
    vTaxNo = Trim(TXT105.Text)
End If

vQuery = "select docno from bcnp.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then
    vTaxDate1 = Trim(TXT106.Text)
    vTaxDate = DTP102.Day & "/" & DTP102.Month & "/" & DTP102.Year
        If vTaxDate <> vTaxDate1 Then
            vQuery = "Exec bcnp.dbo.USP_AP_ChangeTaxDateTaxData  " & Me.CBCreditNote.Value & ",'" & vDocNo & "','" & vTaxNo & "','" & vTaxDate & "' "
            gConnection.Execute vQuery
            MsgBox "เลขที่เอกสาร " & vDocNo & " ได้แก้ไขวันที่ใบกำกับภาษีเป็นวันที่ " & vTaxDate & " เรียบร้อยแล้ว  "
        Else
            MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนวันที่เอกสาร"
            Exit Sub
        End If
    
        If CHK101.Value = 1 And CHK102.Value = 0 Then
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where newdocno = '" & vDocNo & "' and oldtaxno = '" & vTaxNo & "' "
            gConnection.Execute vQuery
        ElseIf CHK101.Value = 1 And CHK102.Value = 1 Then
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where newdocno = '" & vDocNo & "' and newtaxno = '" & vTaxNo & "' "
            gConnection.Execute vQuery
        ElseIf CHK101.Value = 0 And CHK102.Value = 1 Then
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where olddocno = '" & vDocNo & "' and newtaxno = '" & vTaxNo & "' "
            gConnection.Execute vQuery
        ElseIf CHK103.Value = 1 Then
            vDocdate1 = DTP101.Day & "/" & DTP101.Month & "/" & DTP101.Year
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where olddocno = '" & vDocNo & "' and newdocdate = '" & vDocdate1 & "' "
            gConnection.Execute vQuery
        Else
            vDocNo = UCase(vDocNo)
            vDocdate = Trim(TXT102.Text)
            vTaxNo = Trim(TXT103.Text)
            vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
            & " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate1 & "','','','','" & vTaxDate & "','" & vUserID & "',getdate())"
            gConnection.Execute vQuery
        End If
    Else
    MsgBox "ไม่มีข้อมูลเลขที่เอกสาร " & vDocNo & " ในตาราง BCInputtax กรุณาตรวจสอบด้วยครับ"
    End If
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub ChangeTaxDateXPC()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vTaxDate As String
Dim vTaxDate1 As String
Dim vCheckTax As Integer
Dim vTaxNo As String
Dim vDocdate As String
Dim vDocdate1 As String

On Error GoTo ErrDescription

If CHK201.Value = 0 Then
    vDocNo = Trim(TXT201.Text)
Else
    vDocNo = vChangeDoc1
End If

If CHK202.Value = 0 Then
    vTaxNo = Trim(TXT203.Text)
Else
    vTaxNo = Trim(TXT205.Text)
End If

vQuery = "select docno from xpc.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBaseXPC(vXPCConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then
    vTaxDate1 = Trim(TXT206.Text)
    vTaxDate = DTP202.Day & "/" & DTP202.Month & "/" & DTP202.Year
        If vTaxDate <> vTaxDate1 Then
            vQuery = "Exec xpc.dbo.USP_AP_ChangeTaxDateTaxData  " & Me.CBCreditNote_XPC.Value & ",'" & vDocNo & "','" & vTaxNo & "','" & vTaxDate & "' "
            vXPCConnection.Execute vQuery
            MsgBox "เลขที่เอกสาร " & vDocNo & " ได้แก้ไขวันที่ใบกำกับภาษีเป็นวันที่ " & vTaxDate & " เรียบร้อยแล้ว  "
        Else
            MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนวันที่เอกสาร"
            Exit Sub
        End If
    
        If CHK201.Value = 1 And CHK202.Value = 0 Then
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where newdocno = '" & vDocNo & "' and oldtaxno = '" & vTaxNo & "' "
            gConnection.Execute vQuery
        ElseIf CHK201.Value = 1 And CHK202.Value = 1 Then
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where newdocno = '" & vDocNo & "' and newtaxno = '" & vTaxNo & "' "
            gConnection.Execute vQuery
        ElseIf CHK201.Value = 0 And CHK202.Value = 1 Then
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where olddocno = '" & vDocNo & "' and newtaxno = '" & vTaxNo & "' "
            gConnection.Execute vQuery
        ElseIf CHK203.Value = 1 Then
            vDocdate1 = DTP201.Day & "/" & DTP201.Month & "/" & DTP201.Year
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where olddocno = '" & vDocNo & "' and newdocdate = '" & vDocdate1 & "' "
            gConnection.Execute vQuery
        Else
            vDocNo = UCase(vDocNo)
            vDocdate = Trim(TXT202.Text)
            vTaxNo = Trim(TXT203.Text)
            vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
            & " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate1 & "','','','','" & vTaxDate & "','" & vUserID & "',getdate())"
            gConnection.Execute vQuery
        End If
    Else
    MsgBox "ไม่มีข้อมูลเลขที่เอกสาร " & vDocNo & " ในตาราง BCInputtax กรุณาตรวจสอบด้วยครับ"
    End If
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub vInsertTaxDateBackOffice(vDocNo As String, vDocdate As String, vTaxNo As String, vNewTaxDate As String, vTaxDate As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckTax As Integer

On Error GoTo ErrDescription

       
vQuery = "select docno from bi.bcvat.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then

vQuery = "set dateformat dmy"
gConnection.Execute vQuery

vQuery = "update  bi.bcvat.dbo.bcinputtax set  taxdate = '" & vNewTaxDate & "' where  iscancel = 0 and docno = '" & vDocNo & "' and taxno = '" & vTaxNo & "' "
gConnection.Execute vQuery

vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment  set mydescription = 'Update BCVat'   where  OldDocNo = '" & vDocNo & "'  and OldTaxNo = '" & vTaxNo & "' and oldTaxDate = '" & vTaxDate & "' and NewTaxDate = '" & vNewTaxDate & "' "
gConnection.Execute vQuery


End If
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub TXT201_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If TXT201.Text <> "" Then
        vDocNo = TXT201.Text
        Call CheckNumeric
        

            vQuery = " exec dbo.USP_NP_SearchDocnoEditTaxData " & Me.CBCreditNote.Value & ",'" & vDocNo & "' "
            If OpenDataBaseXPC(vXPCConnection, vRecordset, vQuery) <> 0 Then
                TXT202.Text = Trim(vRecordset.Fields("docdate").Value)
                    If Not IsNull(Trim(vRecordset.Fields("taxno").Value)) Then
                        TXT203.Text = Trim(vRecordset.Fields("taxno").Value)
                    Else
                        TXT203.Text = "NoTaxNo"
                    End If
                    If Not IsNull(Trim(vRecordset.Fields("taxdate").Value)) Then
                        TXT206.Text = Trim(vRecordset.Fields("taxdate").Value)
                    Else
                        TXT206.Text = Trim(vRecordset.Fields("docdate").Value)
                    End If
                Else
                MsgBox "เอกสาร เลขที่ " & vDocNo & " ไม่สามารถเปลี่ยนเอกสารได้ โปรดตรวจสอบด้วยนะครับ"
                TXT204.Text = ""
                TXT205.Text = ""
                Exit Sub
            End If
            vRecordset.Close
                
            TXT204.Text = ""
            TXT205.Text = ""
            DTP201.Value = TXT202.Text
            DTP202.Value = TXT206.Text
        End If
        
    End If
                       


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

