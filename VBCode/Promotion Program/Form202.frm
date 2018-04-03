VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Form202 
   Caption         =   "ตรวจสอบสถานะใบเสนอสินค้าโปรโมชั่น"
   ClientHeight    =   9000
   ClientLeft      =   6000
   ClientTop       =   855
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   959
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "สถานะใบเสนอสินค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   150
      TabIndex        =   2
      Top             =   375
      Width           =   14085
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   9225
         TabIndex        =   9
         Text            =   "เลือก Section Manager"
         Top             =   450
         Width           =   2490
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1275
         TabIndex        =   8
         Text            =   "เลือกโปรโมชั่น"
         Top             =   450
         Width           =   2565
      End
      Begin VB.Label Label7 
         Caption         =   "Sec. Manager"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7875
         TabIndex        =   7
         Top             =   450
         Width           =   1290
      End
      Begin VB.Label Label6 
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
         Height          =   315
         Left            =   450
         TabIndex        =   6
         Top             =   450
         Width           =   765
      End
   End
   Begin VB.Frame Frame4 
      Height          =   6540
      Left            =   150
      TabIndex        =   0
      Top             =   1425
      Width           =   14085
      Begin TabDlg.SSTab SSTab1 
         Height          =   6015
         Left            =   90
         TabIndex        =   1
         Top             =   300
         Width           =   13890
         _ExtentX        =   24500
         _ExtentY        =   10610
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "ใบเสนอสินค้ารอตรวจสอบ"
         TabPicture(0)   =   "Form202.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "ListView1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "ใบเสนอสินค้ารออนุมัติ"
         TabPicture(1)   =   "Form202.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ListView2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "ใบเสนอสินค้าอนุมัติแล้ว"
         TabPicture(2)   =   "Form202.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "ListView3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin MSComctlLib.ListView ListView3 
            Height          =   5415
            Left            =   75
            TabIndex        =   5
            Top             =   450
            Width           =   13680
            _ExtentX        =   24130
            _ExtentY        =   9551
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   5490
            Left            =   -74925
            TabIndex        =   4
            Top             =   450
            Width           =   13695
            _ExtentX        =   24156
            _ExtentY        =   9684
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   5415
            Left            =   -74955
            TabIndex        =   3
            Top             =   450
            Width           =   13740
            _ExtentX        =   24236
            _ExtentY        =   9551
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
      End
   End
End
Attribute VB_Name = "Form202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

