VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form frmWizard 
   Caption         =   "Label Wizard V.1.5"
   ClientHeight    =   8595
   ClientLeft      =   2220
   ClientTop       =   1575
   ClientWidth     =   14100
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   8595
   ScaleWidth      =   14100
   Begin Crystal.CrystalReport Crystal102 
      Left            =   4545
      Top             =   8235
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
   Begin Crystal.CrystalReport Crystal101 
      Left            =   3105
      Top             =   8325
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
   Begin VB.PictureBox picHeader 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   14040
      TabIndex        =   45
      Top             =   0
      Width           =   14100
      Begin VB.Label LabelDetail 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   47
         Top             =   450
         Width           =   9495
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   12735
         Picture         =   "frmWizard.frx":1A7A
         Top             =   120
         Width           =   960
      End
      Begin VB.Label LabelHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   46
         Top             =   165
         Width           =   705
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   180
      Top             =   8280
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000005&
      Height          =   5940
      Index           =   1
      Left            =   0
      Picture         =   "frmWizard.frx":403C
      ScaleHeight     =   5880
      ScaleWidth      =   14040
      TabIndex        =   48
      Top             =   1035
      Visible         =   0   'False
      Width           =   14100
      Begin VB.Frame Frame4 
         BackColor       =   &H80000005&
         Height          =   615
         Left            =   3120
         TabIndex        =   68
         Top             =   4440
         Width           =   10410
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "frmWizard.frx":72CF
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   69
            Top             =   170
            Width           =   270
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "เลือกรูปแบบการทำงานที่ต้องการแล้วกดปุ่ม Enter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   855
            TabIndex        =   70
            Top             =   225
            Width           =   4215
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9600
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   43
         ImageHeight     =   43
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizard.frx":76E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizard.frx":9C0F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizard.frx":C1E9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizard.frx":E61D
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizard.frx":10B09
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizard.frx":10C8B
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizard.frx":13287
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizard.frx":13647
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizard.frx":15BDF
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4815
         Left            =   2880
         TabIndex        =   49
         Top             =   495
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   8493
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   0
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000005&
      Height          =   6195
      Index           =   0
      Left            =   0
      Picture         =   "frmWizard.frx":186AA
      ScaleHeight     =   6135
      ScaleWidth      =   14040
      TabIndex        =   4
      Top             =   1035
      Visible         =   0   'False
      Width           =   14100
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Caption         =   "กรุณากรอกชื่อผู้ใช้งานและรหัสผ่าน"
         Height          =   2295
         Index           =   1
         Left            =   3720
         TabIndex        =   5
         Top             =   1080
         Width           =   6135
         Begin VB.CommandButton cmdReset 
            BackColor       =   &H8000000E&
            Caption         =   "&Reset"
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
            Left            =   4680
            TabIndex        =   79
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H8000000E&
            Caption         =   "&Enter"
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
            Left            =   3465
            TabIndex        =   78
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtPassword 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            IMEMode         =   3  'DISABLE
            Left            =   2520
            PasswordChar    =   "*"
            TabIndex        =   51
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtUsername 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   2520
            TabIndex        =   50
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4680
            TabIndex        =   80
            Top             =   680
            Width           =   855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Password :"
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   77
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "&Username :"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   6
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000005&
         Height          =   615
         Left            =   3720
         TabIndex        =   65
         Top             =   3840
         Width           =   6135
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "frmWizard.frx":1B93D
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   66
            Top             =   200
            Width           =   270
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "พิมพ์รหัสผู้ใช้และรหัสผ่าน แล้วกดปุ่ม Enter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   67
            Top             =   240
            Width           =   4695
         End
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000010&
      Height          =   6210
      Index           =   5
      Left            =   0
      ScaleHeight     =   6150
      ScaleWidth      =   14145
      TabIndex        =   32
      Top             =   1035
      Visible         =   0   'False
      Width           =   14210
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   3090
         TabIndex        =   53
         Top             =   1215
         Visible         =   0   'False
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   37846
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ค้นหา"
         Height          =   330
         Left            =   5850
         TabIndex        =   39
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   345
         Left            =   3090
         TabIndex        =   38
         Top             =   1200
         Width           =   2640
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000010&
         Height          =   615
         Index           =   3
         Left            =   240
         TabIndex        =   33
         Top             =   405
         Width           =   6495
         Begin VB.OptionButton Option9 
            BackColor       =   &H80000010&
            Caption         =   "วันที่ออกใบรับสินค้า"
            Height          =   255
            Left            =   4440
            TabIndex        =   36
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H80000010&
            Caption         =   "รหัสเจ้าหนี้"
            Height          =   255
            Left            =   2880
            TabIndex        =   35
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H80000010&
            Caption         =   "เลขที่เอกสาร"
            Height          =   255
            Left            =   1320
            TabIndex        =   34
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000010&
            Caption         =   "ค้นหาจาก :"
            Height          =   255
            Left            =   360
            TabIndex        =   37
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView LV_PO1_1 
         Height          =   3690
         Left            =   225
         TabIndex        =   40
         Top             =   1725
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6509
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
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Code"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "WHCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "BarCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ItemName"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "UnitCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "DocNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "DocDate"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LV_PO1_2 
         Height          =   4935
         Left            =   6840
         TabIndex        =   41
         Top             =   480
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   8705
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ItemNumber"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "WHCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "BarCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ItemName"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "UnitCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "DocNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "DocDate"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbPO2 
         BackColor       =   &H80000010&
         Caption         =   "เลขที่เอกสาร :"
         Height          =   255
         Left            =   2070
         TabIndex        =   42
         Top             =   1260
         Width           =   1125
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000010&
      Height          =   6210
      Index           =   4
      Left            =   0
      ScaleHeight     =   6150
      ScaleWidth      =   14055
      TabIndex        =   21
      Top             =   1035
      Visible         =   0   'False
      Width           =   14120
      Begin MSComCtl2.DTPicker DTPicker2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Left            =   1305
         TabIndex        =   54
         Top             =   1215
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   37846
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000010&
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   405
         Width           =   6495
         Begin VB.OptionButton Option6 
            BackColor       =   &H80000010&
            Caption         =   "เลขที่เอกสาร"
            Height          =   255
            Left            =   1320
            TabIndex        =   28
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H80000010&
            Caption         =   "รหัสเจ้าหนี้"
            Height          =   255
            Left            =   2880
            TabIndex        =   27
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H80000010&
            Caption         =   "วันที่ออกใบสั่งซื้อ"
            Height          =   255
            Left            =   4560
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000010&
            Caption         =   "ค้นหาจาก :"
            Height          =   255
            Left            =   360
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   1305
         TabIndex        =   24
         Top             =   1215
         Width           =   2550
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ค้นหา"
         Height          =   345
         Left            =   3915
         TabIndex        =   23
         Top             =   1215
         Width           =   855
      End
      Begin MSComctlLib.ListView LV_PO_1 
         Height          =   4020
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7091
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
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Number"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Barcode"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ItemDesc"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หน่วยนับ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ราคา"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "จำนวนที่สั่ง"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LV_PO_2 
         Height          =   5340
         Left            =   6840
         TabIndex        =   30
         Top             =   480
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   9419
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ItemNumber"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Barcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หน่วยนับ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ราคา"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "จำนวนที่พิพม์"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbPO 
         BackColor       =   &H80000010&
         Caption         =   "เลขที่เอกสาร :"
         Height          =   255
         Left            =   270
         TabIndex        =   43
         Top             =   1260
         Width           =   1155
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000010&
      Height          =   6210
      Index           =   2
      Left            =   0
      ScaleHeight     =   6150
      ScaleWidth      =   14055
      TabIndex        =   7
      Top             =   1035
      Visible         =   0   'False
      Width           =   14120
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   510
         TabIndex        =   85
         Top             =   990
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H80000010&
         Height          =   495
         Left            =   30
         TabIndex        =   83
         Top             =   285
         Width           =   1740
         Begin VB.OptionButton Option10 
            BackColor       =   &H80000010&
            Caption         =   "ป้ายติดชั้นวาง"
            Height          =   255
            Left            =   75
            TabIndex        =   84
            Top             =   180
            Visible         =   0   'False
            Width           =   1545
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000010&
         Height          =   1095
         Left            =   6075
         TabIndex        =   63
         Top             =   840
         Width           =   7575
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   75
            TabIndex        =   64
            Top             =   120
            Width           =   7425
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000010&
         Height          =   500
         Index           =   0
         Left            =   1800
         TabIndex        =   12
         Top             =   285
         Width           =   11850
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000010&
            Caption         =   "รหัสสินค้า"
            Height          =   255
            Left            =   1665
            TabIndex        =   15
            Top             =   180
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000010&
            Caption         =   "รหัสบาร์โค้ด"
            Height          =   255
            Left            =   3465
            TabIndex        =   14
            Top             =   180
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000010&
            Caption         =   "ชื่อสินค้า"
            Height          =   255
            Left            =   5280
            TabIndex        =   13
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "ค้นหาจาก :"
            Height          =   255
            Left            =   225
            TabIndex        =   16
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3150
         TabIndex        =   11
         Top             =   945
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ค้นหา"
         Height          =   345
         Left            =   5265
         TabIndex        =   10
         Top             =   945
         Width           =   675
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   3150
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV_Label 
         Height          =   3255
         Left            =   480
         TabIndex        =   9
         Top             =   2040
         Width           =   13170
         _ExtentX        =   23230
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Number"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "คลัง"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Barcode"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ItemDesc"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วยนับ"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ราคา"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "จำนวนที่พิมพ์"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ราคาปกติ"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ที่เก็บ"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000010&
         Caption         =   "คลัง :"
         Height          =   255
         Left            =   90
         TabIndex        =   82
         Top             =   1005
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbDetail 
         BackColor       =   &H80000010&
         Caption         =   "รหัสสินค้า :"
         Height          =   255
         Left            =   2025
         TabIndex        =   18
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000010&
         Caption         =   "จำนวนที่พิมพ์ :"
         Height          =   255
         Left            =   2025
         TabIndex        =   17
         Top             =   1485
         Width           =   1095
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000010&
      Height          =   6210
      Index           =   9
      Left            =   0
      ScaleHeight     =   6150
      ScaleWidth      =   14055
      TabIndex        =   100
      Top             =   1035
      Visible         =   0   'False
      Width           =   14120
      Begin MSComctlLib.ListView LV_TRF2 
         Height          =   3540
         Left            =   7425
         TabIndex        =   103
         Top             =   1800
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   6244
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่เอกสาร"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "รหัสบาร์โค้ดสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ชื่อสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วยนับ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ราคาขาย"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "จำนวน"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "คลัง"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LV_TRF1 
         Height          =   3540
         Left            =   360
         TabIndex        =   102
         Top             =   1800
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   6244
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
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่เอกสาร"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "รหัสบาร์โค้ดสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ชื่อสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วยนับ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ราคาขาย"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "จำนวน"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "คลัง"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H80000010&
         Caption         =   "ค้นหาเอกสาร"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   360
         TabIndex        =   101
         Top             =   150
         Width           =   6555
         Begin VB.TextBox TXTTRF1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2835
            TabIndex        =   125
            Top             =   900
            Width           =   2760
         End
         Begin VB.CommandButton CMDTRF1 
            Caption         =   ">>>"
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
            Left            =   5625
            TabIndex        =   124
            Top             =   900
            Width           =   510
         End
         Begin VB.OptionButton OptTRF2 
            BackColor       =   &H80000010&
            Caption         =   "ตามวันที่เอกสาร"
            Height          =   315
            Left            =   225
            TabIndex        =   105
            Top             =   675
            Width           =   1530
         End
         Begin VB.OptionButton OptTRF1 
            BackColor       =   &H80000010&
            Caption         =   "ตามเลขที่เอกสาร"
            Height          =   390
            Left            =   225
            TabIndex        =   104
            Top             =   300
            Value           =   -1  'True
            Width           =   2490
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000010&
            Caption         =   "คำที่ค้นหา :"
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
            Left            =   1665
            TabIndex        =   126
            Top             =   900
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000010&
      Height          =   6210
      Index           =   8
      Left            =   0
      ScaleHeight     =   6150
      ScaleWidth      =   14055
      TabIndex        =   94
      Top             =   1035
      Visible         =   0   'False
      Width           =   14120
      Begin MSComctlLib.ListView LV_Sop2 
         Height          =   3915
         Left            =   7425
         TabIndex        =   97
         Top             =   1890
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6906
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
            Text            =   "เลขที่ใบสั่งขาย"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หน่วยนับ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ราคาขาย"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "จำนวน"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "คลัง"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LV_Sop1 
         Height          =   3960
         Left            =   315
         TabIndex        =   96
         Top             =   1875
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   6985
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
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่ใบสั่งขาย"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หน่วยนับ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ราคาขาย"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "จำนวน"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "คลัง"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000010&
         Caption         =   "เลือกการค้นหาใบสั่งขาย"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   315
         TabIndex        =   95
         Top             =   300
         Width           =   6750
         Begin VB.CommandButton CMDSop1 
            Caption         =   ">>>"
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
            Left            =   5895
            TabIndex        =   128
            Top             =   945
            Width           =   555
         End
         Begin VB.TextBox TXTSop1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3105
            TabIndex        =   127
            Top             =   945
            Width           =   2760
         End
         Begin VB.OptionButton OptSop2 
            BackColor       =   &H80000010&
            Caption         =   "ค้นหาจากรหัสลูกค้า"
            Height          =   315
            Left            =   135
            TabIndex        =   99
            Top             =   675
            Width           =   1755
         End
         Begin VB.OptionButton OptSop1 
            BackColor       =   &H80000010&
            Caption         =   "ค้นหาจากเลขที่ใบสั่งขาย"
            Height          =   315
            Left            =   150
            TabIndex        =   98
            Top             =   375
            Value           =   -1  'True
            Width           =   2040
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000010&
            Caption         =   "คำที่ค้นหา :"
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
            Left            =   2070
            TabIndex        =   129
            Top             =   945
            Width           =   960
         End
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000010&
      Height          =   6210
      Index           =   11
      Left            =   0
      ScaleHeight     =   6150
      ScaleWidth      =   14055
      TabIndex        =   130
      Top             =   1035
      Visible         =   0   'False
      Width           =   14120
      Begin VB.CommandButton CMBSelect 
         BackColor       =   &H00FFFFFF&
         Height          =   825
         Left            =   6930
         Picture         =   "frmWizard.frx":1BD50
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   2700
         Width           =   600
      End
      Begin VB.ComboBox CMBWH 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   132
         Top             =   315
         Width           =   1680
      End
      Begin VB.CheckBox Check101 
         BackColor       =   &H80000010&
         Caption         =   "เลือกทั้งหมด"
         Height          =   330
         Left            =   405
         TabIndex        =   137
         Top             =   5580
         Width           =   1770
      End
      Begin VB.CommandButton CMDShelf 
         Caption         =   "ดูข้อมูล"
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
         Left            =   4005
         TabIndex        =   135
         Top             =   810
         Width           =   1275
      End
      Begin MSComctlLib.ListView ListViewShelf2 
         Height          =   4110
         Left            =   7695
         TabIndex        =   139
         Top             =   1440
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   7250
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสชั้นเก็บ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อชั้นเก็บ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "คลัง"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Section"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewShelf1 
         Height          =   4110
         Left            =   405
         TabIndex        =   136
         Top             =   1440
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   7250
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสชั้นเก็บ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อชั้นเก็บ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "คลัง"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Section"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.TextBox TXTShelf2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2790
         TabIndex        =   134
         Top             =   855
         Width           =   1050
      End
      Begin VB.TextBox TXTShelf1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   133
         Top             =   855
         Width           =   1050
      End
      Begin VB.Label Label38 
         BackColor       =   &H80000010&
         Caption         =   "ถึง"
         Height          =   330
         Left            =   2565
         TabIndex        =   142
         Top             =   855
         Width           =   240
      End
      Begin VB.Label Label37 
         BackColor       =   &H80000010&
         Caption         =   "ลงตะกร้า"
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
         Left            =   7650
         TabIndex        =   141
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label Label36 
         BackColor       =   &H80000010&
         Caption         =   "ระหว่างชั้นเก็บ"
         Height          =   285
         Left            =   360
         TabIndex        =   140
         Top             =   855
         Width           =   1095
      End
      Begin VB.Label Label35 
         BackColor       =   &H80000010&
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
         Height          =   285
         Left            =   360
         TabIndex        =   131
         Top             =   315
         Width           =   1365
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000010&
      Height          =   6210
      Index           =   10
      Left            =   0
      ScaleHeight     =   6150
      ScaleWidth      =   14055
      TabIndex        =   106
      Top             =   1035
      Visible         =   0   'False
      Width           =   14120
      Begin VB.CommandButton CMDSearchItemChangePrice 
         BackColor       =   &H80000010&
         Caption         =   "ค้นหาข้อมูล"
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
         Left            =   2835
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   225
         Width           =   1050
      End
      Begin VB.CheckBox Check103 
         BackColor       =   &H80000010&
         Caption         =   "ลบข้อมูลในตาราง"
         Height          =   285
         Left            =   7470
         TabIndex        =   146
         Top             =   5760
         Width           =   3390
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   330
         Left            =   1170
         TabIndex        =   145
         Top             =   225
         Width           =   1590
         _ExtentX        =   2805
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
         Format          =   16449537
         CurrentDate     =   38938
      End
      Begin VB.CheckBox Check102 
         BackColor       =   &H80000010&
         Caption         =   "เลือกทั้งหมด"
         Height          =   285
         Left            =   3105
         TabIndex        =   143
         Top             =   1125
         Width           =   1275
      End
      Begin MSComctlLib.ProgressBar PrgBar101 
         Height          =   315
         Left            =   360
         TabIndex        =   122
         Top             =   675
         Width           =   13350
         _ExtentX        =   23548
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.ComboBox CMBSection 
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
         Left            =   8865
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   180
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSComctlLib.ListView LV_ChangePrice2 
         Height          =   4245
         Left            =   7470
         TabIndex        =   108
         Top             =   1500
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   7488
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสสินค้า"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "บาร์โค้ด"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หน่วยนับ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ราคา"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "จำนวน"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "วันที่อัพเดท"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "คลังที่เก็บ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ราคาตั้ง"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ชั้นเก็บ"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView LV_ChangePrice1 
         Height          =   4245
         Left            =   375
         TabIndex        =   107
         Top             =   1500
         Width           =   6780
         _ExtentX        =   11959
         _ExtentY        =   7488
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสสินค้า"
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสบาร์โค้ด"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หน่วยนับ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ราคา"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "จำนวน"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "วันที่อัพเดท"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "คลังที่เก็บ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ราคาตั้ง"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ชั้นเก็บ"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label Label39 
         BackColor       =   &H80000010&
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
         Height          =   240
         Left            =   360
         TabIndex        =   144
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000010&
         Caption         =   "รหัสสินค้าที่เปลี่ยนราคา"
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
         Left            =   375
         TabIndex        =   112
         Top             =   1170
         Width           =   2490
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000010&
         Caption         =   "รหัสสินค้าที่จะพิมพ์"
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
         Left            =   7470
         TabIndex        =   111
         Top             =   1170
         Width           =   2070
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         Caption         =   "SecManName :"
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
         Left            =   7470
         TabIndex        =   110
         Top             =   225
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000010&
      Height          =   6210
      Index           =   7
      Left            =   0
      ScaleHeight     =   6150
      ScaleWidth      =   14055
      TabIndex        =   86
      Top             =   1035
      Visible         =   0   'False
      Width           =   14120
      Begin MSComctlLib.ListView LV_Asset2 
         Height          =   4050
         Left            =   7200
         TabIndex        =   90
         Top             =   1800
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7144
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
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "BarCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UnitCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LV_Asset1 
         Height          =   4005
         Left            =   360
         TabIndex        =   89
         Top             =   1800
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   7064
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
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Barcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "unitcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox TXTAsset1 
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
         Left            =   1980
         TabIndex        =   88
         Top             =   885
         Width           =   4155
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H80000010&
         Caption         =   "เลือกค้นหาข้อมูล"
         Height          =   1290
         Left            =   360
         TabIndex        =   87
         Top             =   75
         Width           =   6585
         Begin VB.CommandButton CMDAsset1 
            Caption         =   ">>>"
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
            Left            =   5850
            TabIndex        =   93
            Top             =   810
            Width           =   540
         End
         Begin VB.OptionButton OptAsset2 
            BackColor       =   &H80000010&
            Caption         =   "ค้นหาตามชื่อสินทรัพย์"
            Height          =   240
            Left            =   3420
            TabIndex        =   92
            Top             =   375
            Width           =   2115
         End
         Begin VB.OptionButton OptAsset1 
            BackColor       =   &H80000010&
            Caption         =   "ค้นหาตามรหัสสินทรัพย์"
            Height          =   240
            Left            =   150
            TabIndex        =   91
            Top             =   375
            Value           =   -1  'True
            Width           =   2490
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000010&
            Caption         =   "คำค้นหา :"
            Height          =   240
            Left            =   270
            TabIndex        =   151
            Top             =   810
            Width           =   1230
         End
      End
   End
   Begin VB.PictureBox PicWizardSelect 
      Height          =   6210
      Left            =   0
      ScaleHeight     =   6150
      ScaleWidth      =   14055
      TabIndex        =   55
      Top             =   1035
      Visible         =   0   'False
      Width           =   14120
      Begin VB.Frame Frame6 
         Height          =   495
         Left            =   4560
         TabIndex        =   74
         Top             =   360
         Width           =   9000
         Begin VB.Label Label10 
            Caption         =   "Tip : กดปุ่ม ""Select All"" เมื่อต้องการเลือกรายการทั้งหมดใน List"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   165
            Width           =   5295
         End
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "&Select All"
         Height          =   375
         Left            =   480
         TabIndex        =   57
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeselect 
         Caption         =   "&Deselect All"
         Height          =   375
         Left            =   2160
         TabIndex        =   56
         Top             =   480
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListResult 
         Height          =   4335
         Left            =   480
         TabIndex        =   58
         Top             =   960
         Width           =   13080
         _ExtentX        =   23072
         _ExtentY        =   7646
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
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ItemNumber"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "คลัง"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "รหัส Barcode"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ชื่อสินค้าภาษาไทย"
            Object.Width           =   5466
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วยนับ"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ราคา"
            Object.Width           =   2222
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "จำนวนที่พิพม์"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ราคาปกติ"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "วันที่อัพเดท"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ที่เก็บ"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000010&
      Height          =   6210
      Index           =   15
      Left            =   0
      ScaleHeight     =   6150
      ScaleWidth      =   14070
      TabIndex        =   113
      Top             =   1035
      Visible         =   0   'False
      Width           =   14130
      Begin VB.ComboBox CMB102 
         Height          =   315
         Left            =   8325
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   360
         Width           =   3615
      End
      Begin VB.ComboBox CMB101 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   375
         Width           =   6855
      End
      Begin MSComctlLib.ListView LV_Promo1 
         Height          =   4620
         Left            =   135
         TabIndex        =   115
         Top             =   1125
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8149
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสสินค้า"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "บาร์โค้ด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หน่วยนับ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ราคาโปรโมชั่น"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ราคาปกติ"
            Object.Width           =   2293
         EndProperty
      End
      Begin MSComctlLib.ListView LV_Promo2 
         Height          =   4620
         Left            =   7380
         TabIndex        =   114
         Top             =   1125
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8149
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "บาร์โค้ด"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หน่วยนับ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ราคาโปรโมชั่น"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ราคาปกติ"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Label Label31 
         BackColor       =   &H80000010&
         Caption         =   "SecMan :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7380
         TabIndex        =   121
         Top             =   405
         Width           =   915
      End
      Begin VB.Label Label30 
         BackColor       =   &H80000010&
         Caption         =   "ตะกร้าสินค้า"
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
         Left            =   7380
         TabIndex        =   119
         Top             =   855
         Width           =   1965
      End
      Begin VB.Label Label28 
         BackColor       =   &H80000010&
         Caption         =   "รายการโปรโมชั่น"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   117
         Top             =   135
         Width           =   1290
      End
      Begin VB.Label Label29 
         BackColor       =   &H80000010&
         Caption         =   "รายการสินค้า"
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
         Left            =   135
         TabIndex        =   118
         Top             =   825
         Width           =   1185
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H00808080&
      Height          =   6180
      Index           =   6
      Left            =   0
      ScaleHeight     =   6120
      ScaleWidth      =   14070
      TabIndex        =   44
      Top             =   1035
      Visible         =   0   'False
      Width           =   14130
      Begin VB.CheckBox CHKPromoAll 
         BackColor       =   &H00808080&
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
         Height          =   285
         Left            =   180
         TabIndex        =   157
         Top             =   630
         Width           =   1995
      End
      Begin VB.ComboBox CMBPromotionCode 
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
         Height          =   360
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   155
         Top             =   90
         Width           =   6045
      End
      Begin VB.CommandButton CMDImportData 
         Caption         =   "รับข้อมูล"
         Height          =   465
         Left            =   7695
         TabIndex        =   149
         Top             =   45
         Width           =   1185
      End
      Begin MSComctlLib.ListView ListViewItemChangePrice 
         Height          =   3930
         Left            =   180
         TabIndex        =   148
         Top             =   945
         Width           =   13740
         _ExtentX        =   24236
         _ExtentY        =   6932
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
         Appearance      =   1
         NumItems        =   11
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
            Text            =   "บาร์โค้ด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ชื่อสินค้า"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วยนับ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ราคา"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "จำนวน"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "วันที่หมด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "คลัง"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ราคาตั้ง"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "ชั้นเก็บ"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton btnprocess 
         Caption         =   "เริ่มต้นกระบวนการ"
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
         Height          =   225
         Left            =   13410
         TabIndex        =   61
         Top             =   270
         Visible         =   0   'False
         Width           =   585
      End
      Begin MSComctlLib.ListView ListChkValue 
         Height          =   315
         Left            =   270
         TabIndex        =   62
         Top             =   5145
         Visible         =   0   'False
         Width           =   13635
         _ExtentX        =   24051
         _ExtentY        =   556
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ItemNumber"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสบาร์โค้ด"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ราคาจากฐานข้อมูล"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ราคาที่รับมา"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "จำนวนที่พิมพ์"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสโปรโมชั่น :"
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
         Left            =   180
         TabIndex        =   156
         Top             =   135
         Width           =   1545
      End
   End
   Begin VB.PictureBox PicWizard 
      BackColor       =   &H80000010&
      Height          =   7065
      Index           =   3
      Left            =   0
      ScaleHeight     =   7005
      ScaleWidth      =   14145
      TabIndex        =   19
      Top             =   1035
      Visible         =   0   'False
      Width           =   14210
      Begin VB.CommandButton CMDHandHeldClose 
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
         Height          =   510
         Left            =   10575
         TabIndex        =   154
         Top             =   6075
         Width           =   1500
      End
      Begin VB.CheckBox CBSelectAll 
         BackColor       =   &H80000010&
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
         Height          =   285
         Left            =   495
         TabIndex        =   153
         Top             =   450
         Width           =   1725
      End
      Begin VB.CommandButton CMDPrintHandHeld 
         Caption         =   "พิมพ์ป้าย"
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
         Left            =   12195
         TabIndex        =   152
         Top             =   6075
         Width           =   1500
      End
      Begin VB.CommandButton cmdSync 
         Caption         =   "ดึงข้อมูลจากฐานข้อมูล"
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
         Left            =   11520
         TabIndex        =   52
         Top             =   360
         Width           =   2175
      End
      Begin MSComctlLib.ListView LV_Palm 
         Height          =   4815
         Left            =   480
         TabIndex        =   20
         Top             =   840
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   8493
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
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ItemNumber"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัส Barcode"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้าภาษาไทย"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "หน่วยนับ"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "ราคา"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "จำนวนที่พิพม์"
            Object.Width           =   2222
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ชื่อฟอร์ม"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ที่อยู่ฟอร์ม"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ออกโปรแกรม"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   165
      TabIndex        =   0
      Top             =   7380
      Width           =   1485
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "เคลียร์ข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1710
      TabIndex        =   123
      Top             =   7380
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "พิมพ์ออกเครื่องพิมพ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3240
      TabIndex        =   3
      Top             =   7380
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "< &Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4770
      TabIndex        =   1
      Top             =   7380
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6300
      TabIndex        =   2
      Top             =   7380
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "ตัวอย่างก่อนพิมพ์"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7830
      TabIndex        =   31
      Top             =   7380
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton CMDBartendor 
      Caption         =   "Bartendor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9360
      TabIndex        =   81
      Top             =   7380
      Width           =   1485
   End
   Begin VB.PictureBox PicWizardReport 
      BackColor       =   &H80000005&
      Height          =   6195
      Left            =   -45
      Picture         =   "frmWizard.frx":1C175
      ScaleHeight     =   6135
      ScaleWidth      =   14085
      TabIndex        =   59
      Top             =   1035
      Visible         =   0   'False
      Width           =   14145
      Begin VB.CheckBox Check104 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ปรับป้ายราคา"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3150
         TabIndex        =   147
         Top             =   4545
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000005&
         Height          =   615
         Left            =   3120
         TabIndex        =   71
         Top             =   4920
         Width           =   10350
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "frmWizard.frx":1F408
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   72
            Top             =   220
            Width           =   270
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "เลือกแบบฟอร์มที่ต้องการ แล้วกดปุ่ม ""พิมพ์ออกเครื่องพิมพ์"" หรือ ""ตัวอย่างก่อนพิมพ์"""
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
            Height          =   255
            Left            =   480
            TabIndex        =   73
            Top             =   240
            Width           =   6735
         End
      End
      Begin MSComctlLib.ListView LV_Report 
         Height          =   3960
         Left            =   3120
         TabIndex        =   60
         Top             =   480
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   6985
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
            Text            =   "ชนิดป้ายราคา"
            Object.Width           =   7497
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Path Name"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อผู้ใช้งาน :"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   45
      X2              =   14220
      Y1              =   7290
      Y2              =   7290
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FormToPrinter As Boolean
Dim CountItem As Integer
Dim gbstrSQL1, gbstrSQL2 As String
Dim JobSelect As String
Dim AddTempRCV, AddTempPO, AddTempSOP As Boolean
Dim strHeader(), strDetail(), arrBar(), arrPrice() As String
Dim tmpUpDateTime As Date
Dim tmpNUMBR, tmpBarcode, tmpTHINAME, tmpENGNAME, tmpRRCLV, tmpRRC, tmpUOM, tmpUsedUser As String
Dim tmpCategory_ID, tmpSite, tmpBIN_ID, tmpVENDR_ID, tmpRemark, tmpONHAND, tmpQTYALLOCATE As String
Dim tmpTHIName1, tmpBuyDocDate As String
Dim tmpRemainInQTY, tmpRemainOutQTY, tmpDocNo, tmpDocDate As String
Dim tmpQTY, tmpID, tmpTYPE As Integer
Dim DataHas As Boolean, vCheckReceipt As Boolean, vChangePrice As Boolean, vAsset As Boolean
Dim vShelfJob  As Boolean
Dim ImportBartendor As New FileSystemObject
Dim ListForm As ListItem
Dim vWHCode As String, vShelfCode As String
Dim vCheckRemove As Integer
Dim vRemove() As Integer

Dim vCountCheckPriceErect As Integer


Private Sub btnprocess_Click()
        frmWait.Show            ' Loading Data
        DoEvents
        Dim recCount As Long
        Dim percent As Long
        Dim count As Long
        Dim ListCount, ProgressCount, iCount As Integer
        Dim sqlString1, strSQL As String
        Dim ListX As ListItem
        
        percent = 0
        count = 0
        recCount = 0
        
        ' Connect to Palm Database Table (Job_Type = 3)
        ConnectNPDEV
        
        On Error GoTo ErrDescription
        
        strSQL = "SELECT CODE,PRICE FROM Palm_Temp Where Job_Type = '3'"
        If Rs1.State = adStateOpen Then Rs1.Close
        Rs1.Open strSQL, ConnDEV, 1, 3
        If Not Rs1.EOF Then
                ' Redim Array
                CountItem = Int(Rs1.RecordCount)
                ReDim arrBar(Rs1.RecordCount)
                ReDim arrPrice(Rs1.RecordCount)
                
                ' Recordset MoveFirst
                Rs1.MoveFirst
                For iCount = 1 To Rs1.RecordCount
                        arrBar(iCount) = Rs1!code
                        arrPrice(iCount) = Rs1!price
                        Rs1.MoveNext
                Next
        End If
        Rs1.Close       ' Close Connection
        
        ' Clear ค่าที่อยู่ใน Listview เดิม
        ListChkValue.ListItems.Clear
         
        ' ตรวจสอบว่ามีข้อมูลอยู่ในฐานข้อมูลหรือไม่
        If CountItem < 1 Then
                MsgBox "ไม่มีข้อมูลรหัสสินค้าและราคาในฐานข้อมูล", vbCritical + vbOKOnly, "คำเตือน"
                Exit Sub
        End If
                
        ' เชื่อมต่อฐานข้อมูล
        ConnectSQL
        recCount = CountItem
        For iCount = 1 To CountItem
            ' Progress Bar show
            count = count + 1
            percent = 100 * (count / recCount)
            frmWait.lbPercent.Caption = CStr(percent) & "% Loaded"
            ProgressMoveEx percent, 100
            DoEvents
            
            sqlString1 = "SELECT ITEMNMBR,ENGNAME,barcod,ITEMDESC,PRCLEVEL,UOMPRICE,UOFM,salepromotion  FROM   V_Label_PriceList WHERE barcod = '" & Trim(arrBar(iCount)) & "'"
            If Rs2.State = adStateOpen Then Rs2.Close
            Rs2.Open sqlString1, ConnSQL, adOpenDynamic, adLockOptimistic
            If Not Rs2.EOF Then
                    If Int(Rs2!UOMPRICE) <> Int(arrPrice(iCount)) Then
                            tmpNUMBR = Trim(Rs2!Itemnmbr)
                            tmpBarcode = Trim(Rs2!barcod)
                            tmpTHINAME = Trim(Rs2!ITEMDESC)
                            tmpENGNAME = Trim(Rs2!ENGNAME)
                            tmpQTY = "1"
                            tmpRRC = Trim(Rs2!UOMPRICE)
                            tmpUOM = Trim(Rs2!UOFM)
                            tmpRRCLV = Trim(Rs2!salepromotion)
                            tmpUsedUser = "ChkValue"
                            tmpCategory_ID = ""
                            
                            ' Add ListChkValue
                            Set ListX = ListChkValue.ListItems.Add(, , Trim(Rs2!Itemnmbr))
                            ListX.SubItems(1) = Trim(Rs2!barcod)
                            ListX.SubItems(2) = Trim(Rs2!ITEMDESC)
                            ListX.SubItems(3) = Trim(Rs2!UOMPRICE)
                            ListX.SubItems(4) = Int(arrPrice(iCount))
                            ListX.SubItems(5) = "5"
                        End If
                End If
                Rs2.Close
        Next
        
        ConnSQL.Close
        ' Loading Complete
        Unload frmWait
        
        cmdNext.SetFocus
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CBSelectAll_Click()
Dim i As Integer

If Me.CBSelectAll.Value = 1 Then
For i = 1 To Me.LV_Palm.ListItems.count
Me.LV_Palm.ListItems(i).Checked = True
Next i
Else
For i = 1 To Me.LV_Palm.ListItems.count
Me.LV_Palm.ListItems(i).Checked = False
Next i
End If
End Sub

Private Sub Check101_Click()
Dim i As Integer

On Error GoTo ErrDescription

For i = 1 To ListViewShelf1.ListItems.count
  If Check101.Value = 1 Then
    ListViewShelf1.ListItems.Item(i).Checked = True
  Else
    ListViewShelf1.ListItems.Item(i).Checked = False
  End If
Next i

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Check102_Click()
Dim vCount As Integer
Dim i As Integer
Dim ListX As ListItem
Dim iCount As Integer


On Error GoTo ErrDescription

vCount = LV_ChangePrice1.ListItems.count

If Check102.Value = 1 Then
For i = 1 To vCount
  LV_ChangePrice1.ListItems.Item(i).Checked = True
  If LV_ChangePrice1.ListItems.Item(i).Checked = True Then
          Set ListX = LV_ChangePrice2.ListItems.Add(, , LV_ChangePrice1.ListItems(i).Text)
          ListX.SubItems(1) = Me.LV_ChangePrice1.ListItems(i).SubItems(1)
          ListX.SubItems(2) = Me.LV_ChangePrice1.ListItems(i).SubItems(2)
          ListX.SubItems(3) = Me.LV_ChangePrice1.ListItems(i).SubItems(3)
          ListX.SubItems(4) = Me.LV_ChangePrice1.ListItems(i).SubItems(4)
          ListX.SubItems(5) = Me.LV_ChangePrice1.ListItems(i).SubItems(5)
          ListX.SubItems(6) = Me.LV_ChangePrice1.ListItems(i).SubItems(6)
          ListX.SubItems(7) = Me.LV_ChangePrice1.ListItems(i).SubItems(7)
          ListX.SubItems(8) = Me.LV_ChangePrice1.ListItems(i).SubItems(8)
          ListX.SubItems(9) = Me.LV_ChangePrice1.ListItems(i).SubItems(9)
  End If
        
Next i
For iCount = Me.LV_ChangePrice1.ListItems.count To 1 Step -1
  LV_ChangePrice1.ListItems.Remove (iCount)
Next iCount

Check102.Value = 0
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Check103_Click()
Dim vCount As Integer
Dim i As Integer

On Error GoTo ErrDescription

If Check103.Value = 1 Then
  vCount = LV_ChangePrice2.ListItems.count
  For i = vCount To 1 Step -1
    LV_ChangePrice2.ListItems.Remove (i)
  Next i
Check103.Value = 0
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CHKPromoAll_Click()
Dim i As Integer

If Me.CHKPromoAll.Value = 1 Then
For i = 1 To Me.ListViewItemChangePrice.ListItems.count
Me.ListViewItemChangePrice.ListItems(i).Checked = True
Next i
Else
For i = 1 To Me.ListViewItemChangePrice.ListItems.count
Me.ListViewItemChangePrice.ListItems(i).Checked = False
Next i
End If
End Sub

Private Sub CMB102_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListItemPro As ListItem
Dim vPromoCode As String
Dim StrCount As Integer
Dim vSecman As String

On Error GoTo ErrDescription

LV_Promo1.ListItems.Clear
If CMB101.Text <> "" Then
    If CMB102.Text <> "" Then
        StrCount = InStr(Trim(CMB101.Text), "/")
        vPromoCode = Trim(Left(CMB101.Text, StrCount - 1))
        vSecman = Trim(CMB102.Text)
        vQuery = "select * from vw_IV_ItemPromotion where pmcode = '" & vPromoCode & "' and secman = '" & vSecman & "'  order by promoprice "
        If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set ListItemPro = LV_Promo1.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
            ListItemPro.SubItems(1) = Trim(vRecordset.Fields("barcode").Value)
            ListItemPro.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
            ListItemPro.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
            ListItemPro.SubItems(4) = Trim(vRecordset.Fields("promoprice").Value)
            ListItemPro.SubItems(5) = Trim(vRecordset.Fields("priceerect").Value)
            vRecordset.MoveNext
        Wend
        End If
        vRecordset.Close
    Else
    MsgBox "กรุณาเลือก ชื่อ Section ด้วยนะครับ"
    End If
Else
MsgBox "กรุณาเลือก โปรโมชั่นที่จะพิมพ์ป้ายราคาด้วยนะครับ"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub


Private Sub CMBSection_Click()
Dim vQuery As String
Dim vSectionCode As String
Dim ListChangePrice As ListItem
Dim vCount As Integer
Dim i As Integer
Dim vDocDate As Date

On Error Resume Next

PrgBar101.Visible = True
LV_ChangePrice1.ListItems.Clear
DTPicker3 = Now
vSectionCode = Left(Trim(CMBSection.Text), InStr(Trim(CMBSection.Text), "//") - 1)
vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
i = 0
If Rs1.State = 1 Then
Rs1.Close
End If
vQuery = "exec  dbo.USP_NP_SearchItemChangePrice '" & vSectionCode & "','" & vDocDate & "' "
Rs1.Open vQuery, ConnSQL, adOpenDynamic, adLockOptimistic
If Not Rs1.EOF Then
    PrgBar101.Max = Rs1.RecordCount
    Rs1.MoveFirst
    While Not Rs1.EOF
    Set ListChangePrice = LV_ChangePrice1.ListItems.Add(, , Trim(Rs1.Fields("itemcode").Value))
    
    ListChangePrice.SubItems(1) = Trim(Rs1.Fields("barcode").Value)
    If IsNull(Rs1.Fields("name1").Value) Then
        ListChangePrice.SubItems(2) = ""
    Else
        ListChangePrice.SubItems(2) = Trim(Rs1.Fields("name1").Value)
    End If
    ListChangePrice.SubItems(3) = Trim(Rs1.Fields("unitcode").Value)
    ListChangePrice.SubItems(4) = Trim(Rs1.Fields("SalePrice1").Value)
    ListChangePrice.SubItems(5) = 1
    ListChangePrice.SubItems(6) = Trim(Rs1.Fields("dateupdate").Value)
    ListChangePrice.SubItems(7) = Trim(Rs1.Fields("whcode").Value)
    ListChangePrice.SubItems(8) = Trim(Rs1.Fields("priceerect").Value)
    ListChangePrice.SubItems(9) = Trim(Rs1.Fields("shelfcode").Value)
    Rs1.MoveNext
    i = i + 1
    PrgBar101.Value = i
    Wend
End If
Rs1.Close
PrgBar101.Visible = False
End Sub

Private Sub CMBSelect_Click()
Dim i As Integer
Dim vListShelfBin As ListItem
Dim j As Integer

On Error GoTo ErrDescription

If ListViewShelf1.ListItems.count <> 0 Then
  For i = 1 To ListViewShelf1.ListItems.count
    If ListViewShelf1.ListItems.Item(i).Checked = True Then
        j = ListViewShelf2.ListItems.count
        j = j + 1
        Set vListShelfBin = ListViewShelf2.ListItems.Add(, , j)
        vListShelfBin.SubItems(1) = Trim(ListViewShelf1.ListItems.Item(i).SubItems(1))
        vListShelfBin.SubItems(2) = Trim(ListViewShelf1.ListItems.Item(i).SubItems(2))
        vListShelfBin.SubItems(3) = Trim(ListViewShelf1.ListItems.Item(i).SubItems(3))
        vListShelfBin.SubItems(4) = Trim(ListViewShelf1.ListItems.Item(i).SubItems(4))
    End If
  Next i
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMBWHCode_Click()
Dim vQuery As String
Dim vSectionCode As String
Dim ListChangePrice As ListItem
Dim vCount As Integer
Dim i As Integer
Dim vDocDate As Date

On Error Resume Next

PrgBar101.Visible = True
LV_ChangePrice1.ListItems.Clear
DTPicker3 = Now
vSectionCode = Left(Trim(CMBSection.Text), InStr(Trim(CMBSection.Text), "//") - 1)
vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
i = 0
If Rs1.State = 1 Then
Rs1.Close
End If
vQuery = "exec  dbo.USP_NP_SearchItemChangePrice '" & vSectionCode & "','" & vDocDate & "' "
Rs1.Open vQuery, ConnSQL, adOpenDynamic, adLockOptimistic
If Not Rs1.EOF Then
    PrgBar101.Max = Rs1.RecordCount
    Rs1.MoveFirst
    While Not Rs1.EOF
    Set ListChangePrice = LV_ChangePrice1.ListItems.Add(, , Trim(Rs1.Fields("itemcode").Value))
    
    ListChangePrice.SubItems(1) = Trim(Rs1.Fields("barcode").Value)
    If IsNull(Rs1.Fields("name1").Value) Then
        ListChangePrice.SubItems(2) = ""
    Else
        ListChangePrice.SubItems(2) = Trim(Rs1.Fields("name1").Value)
    End If
    ListChangePrice.SubItems(3) = Trim(Rs1.Fields("unitcode").Value)
    ListChangePrice.SubItems(4) = Trim(Rs1.Fields("SalePrice1").Value)
    ListChangePrice.SubItems(5) = 1
    ListChangePrice.SubItems(6) = Trim(Rs1.Fields("dateupdate").Value)
    ListChangePrice.SubItems(7) = Trim(Rs1.Fields("whcode").Value)
    ListChangePrice.SubItems(8) = Trim(Rs1.Fields("priceerect").Value)
    ListChangePrice.SubItems(9) = Trim(Rs1.Fields("shelfcode").Value)
    Rs1.MoveNext
    i = i + 1
    PrgBar101.Value = i
    Wend
End If
Rs1.Close
PrgBar101.Visible = False
End Sub

Private Sub CMD101_Click()
LV_Label.ListItems.Clear
End Sub

Private Sub cmdAll_Click()
        Dim j As Integer
        
        On Error GoTo ErrDescription
        
        For j = 1 To ListResult.ListItems.count
                ListResult.ListItems(j).Checked = True
        Next
        cmdNext.SetFocus
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDAsset1_Click()
        Dim tmpDate As String
        On Error GoTo Err1:
        
        ConnectSQL
        Dim ListX As ListItem
        
        If OptAsset1.Value = True Then
                gbstrSQL1 = "Select code,name,unitcode,buyprice From bcassetsmaster Where code = '" & Trim(TXTAsset1.Text) & "'"
        End If
        If OptAsset2.Value = True Then
                gbstrSQL1 = "Select code,name,unitcode,buyprice From bcassetsmaster Where name = '" & Trim(TXTAsset1.Text) & "'"
        End If

    
        Rs1.Open gbstrSQL1, ConnSQL, adOpenDynamic, adLockOptimistic
        If Not Rs1.EOF Then
                LV_Asset1.ListItems.Clear
                Rs1.MoveFirst
                While Not Rs1.EOF
                    Set ListX = LV_Asset1.ListItems.Add(, , Trim(Rs1!code))
                    If IsNull(Rs1!code) = True Then
                            MsgBox "ไม่มีบาร์โค้ด & " & Rs1!Itemcode & " ", vbInformation, "ข้อความเตือน"
                            ListX.SubItems(1) = "No Barcode"
                    Else
                            ListX.SubItems(1) = Trim(Rs1!code)
                    End If
                    ListX.SubItems(2) = Trim(Rs1!Name)
                    If IsNull(Rs1!Unitcode) = True Then
                    ListX.SubItems(3) = "No UnitCode"
                    Else
                    ListX.SubItems(3) = Trim(Rs1!Unitcode)
                    End If
                    If IsNull(Rs1!buyprice) = True Then
                    ListX.SubItems(4) = "NoBuyPrice"
                    Else
                    ListX.SubItems(4) = Trim(Rs1!buyprice)
                    End If
                    ListX.SubItems(5) = 2
                    Rs1.MoveNext
                Wend
                Rs1.Close
        Else        ' ไม่พบรายการที่ค้นหา
                MsgBox "ไม่พบทรัพย์สินเลขที่" & Trim(TXTAsset1.Text), vbOKOnly + vbInformation, "คำเตือน"
                If Me.DTPicker1.Visible = False Then
                        TXTAsset1.Text = ""
                        TXTAsset1.SetFocus
                End If
                Exit Sub
        End If
        Exit Sub
        
' Error Found
Err1:
        MsgBox Err.Description, vbOKOnly + vbCritical, "พบข้อผิดพลาดของตัวโปรแกรม"
        Exit Sub
End Sub

Private Sub CMDBartendor_Click()
        Dim tmpPathName, strCreate_Temp, strDel_Temp As String
        Dim strSQL, strSQL2 As String
        Dim iCount As Integer
        Dim BartendorForm As String
        Dim Rs4 As New ADODB.Recordset
        
        On Error GoTo ErrDescription
        
        'strSQL = "Delete From Report_Temp1 Where Useduser = '" & strUsername & "'"
        'ConnSQL.Execute strSQL
        'On Error GoTo ErrDescription
        tmpPathName = Trim(LV_Report.SelectedItem.SubItems(1)) & ".rpt"
        BartendorForm = Mid(tmpPathName, 21, 4)
        ' Dump Data to Report_Temp
        If BartendorForm = "R2-8" Then
        strCreate_Temp = "select * into dbo.Report_Temp1 From NP_LABEL_TEMP where UsedUser = 'Null'"
        ConnSQL.Execute strCreate_Temp
        strSQL = "Select * From NP_Label_Temp Where UsedUser = '" & strUsername & "' "
        Rs3.Open strSQL, ConnSQL, adOpenDynamic, adLockOptimistic
        If Not Rs3.EOF Then
                Rs3.MoveFirst
                While Not Rs3.EOF
                    If Int(Rs3!QTY) > 0 Then
                            strSQL2 = "Insert Into dbo.Report_Temp1(ItemCode, barcode, NAME1, NAME2, QTY, PriceLevel, Price, UnitCode, UsedUser, Category_ID, WHCode, ShelfCode, VENDR_ID, remark, SPrice,RemainOutQTY,RemainInQTY) " _
                                        & " values('" & Trim(Rs3!Itemcode) & "','" & Trim(Rs3!barcode) & "','" & Trim(Rs3!Name1) & "', '" & Trim(Rs3!Name2) & "', 1,'" & Trim(Rs3!PriceLevel) & "','" & Trim(Rs3!price) & "','" & Trim(Rs3!Unitcode) & "'," _
                                        & "'" & Trim(Rs3!UsedUser) & "','" & Trim(Rs3!Category_ID) & "','" & Trim(Rs3!whcode) & "', '" & Trim(Rs3!ShelfCode) & "','" & Trim(Rs3!VENDR_ID) & "', '" & Trim(Rs3!Remark) & "','" & Trim(Rs3!SPrice) & "' ,'" & Rs3!RemainOutQTY & "','" & Rs3!RemainInQTY & "') "
                            For iCount = 1 To Rs3!QTY
                                    ConnSQL.Execute strSQL2
                            Next iCount
                    End If
                    Rs3.MoveNext
                Wend
        End If
        Rs3.Close
'สร้าง Text Files ที่เครื่องแม่ข่ายเพื่อที่จะให้ Commander จับไปพิมพ์ปายราคา
ElseIf BartendorForm = "R1-1" Or BartendorForm = "R1-2" Or BartendorForm = "R1-3" Or BartendorForm = "R1-4" Or BartendorForm = "R1-5" Or BartendorForm = "R1-6" Or BartendorForm = "R5-1" Or BartendorForm = "R5-2" Or BartendorForm = "R5-3" Or BartendorForm = "R5-4" Or BartendorForm = "R5-5" Or BartendorForm = "R5-6" Or BartendorForm = "R5-7" Then
strCreate_Temp = "select * into tempdb.dbo.Report_Temp2 From NP_LABEL_TEMP where UsedUser = 'Null'"
ConnSQL.Execute strCreate_Temp
strSQL = "Select * From NP_Label_Temp Where UsedUser = '" & strUsername & "' "
        Rs3.Open strSQL, ConnSQL, adOpenDynamic, adLockOptimistic
        If Not Rs3.EOF Then
                Rs3.MoveFirst
                While Not Rs3.EOF
                    If Int(Rs3!QTY) > 0 Then
                            strSQL2 = "Insert Into tempdb.dbo.Report_Temp2(ItemCode, barcode, NAME1, NAME2, QTY, PriceLevel, Price, UnitCode, UsedUser, Category_ID, WHCode, ShelfCode, VENDR_ID, remark, SPrice,OnHand,RemainOutQTY,RemainInQTY,SOPNUM,SOPDOC) " _
                                        & " values('" & Trim(Rs3!Itemcode) & "','" & Trim(Rs3!barcode) & "','" & Trim(Rs3!Name1) & "', '" & Trim(Rs3!Name2) & "', 1,'" & Trim(Rs3!PriceLevel) & "','" & Trim(Rs3!price) & "','" & Trim(Rs3!Unitcode) & "'," _
                                        & "'" & Trim(Rs3!UsedUser) & "','" & Trim(Rs3!Category_ID) & "','" & Trim(Rs3!whcode) & "', '" & Trim(Rs3!ShelfCode) & "','" & Trim(Rs3!VENDR_ID) & "', '" & Trim(GenFormatBarCode(Rs3!barcode)) & "','" & Trim(Rs3!SPrice) & "' ,'" & Rs3!onhand & "','" & Rs3!RemainOutQTY & "','" & Rs3!RemainInQTY & "','" & Rs3!SOPNUM & "','" & Rs3!sopdoc & "') "
                            For iCount = 1 To Rs3!QTY
                                    ConnSQL.Execute strSQL2
                            Next iCount
                    End If
                    Rs3.MoveNext
                Wend
        End If
         Rs3.Close
'ElseIf BartendorForm = "R5-2" Then
'strCreate_Temp = "select * into dbo.Report_Temp2 From NP_LABEL_TEMP where UsedUser = 'Null'"
'ConnSQL.Execute strCreate_Temp
'strSQL = "Select * From NP_Label_Temp Where UsedUser = '" & strUsername & "' "
        'Rs3.Open strSQL, ConnSQL, adOpenDynamic, adLockOptimistic
        'If Not Rs3.EOF Then
         '       Rs3.MoveFirst
          '      While Not Rs3.EOF
           '         If Int(Rs3!QTY) > 0 Then
            '                strSQL2 = "Insert Into dbo.Report_Temp2(ItemCode, barcode, NAME1, NAME2, QTY, PriceLevel, Price, UnitCode, UsedUser, Category_ID, WHCode, ShelfCode, VENDR_ID, remark, SPrice,OnHand,RemainOutQTY,RemainInQTY,SOPNUM,SOPDOC) " _
             '                           & " values('" & Trim(Rs3!Itemcode) & "','" & Trim(Rs3!barcode) & "','" & Trim(Rs3!Name1) & "', '" & Trim(Rs3!Name2) & "', 1,'" & Trim(Rs3!PriceLevel) & "','" & Trim(Rs3!price) & "','" & Trim(Rs3!Unitcode) & "'," _
              '                          & "'" & Trim(Rs3!UsedUser) & "','" & Trim(Rs3!Category_ID) & "','" & Trim(Rs3!whcode) & "', '" & Trim(Rs3!ShelfCode) & "','" & Trim(Rs3!VENDR_ID) & "', '" & Trim(Rs3!Remark) & "','" & Trim(Rs3!SPrice) & "' ,'" & Rs3!onhand & "','" & Rs3!RemainOutQTY & "','" & Rs3!RemainInQTY & "','" & Rs3!SOPNUM & "','" & Rs3!sopdoc & "') "
               '             For iCount = 1 To Rs3!QTY
                '                    ConnSQL.Execute strSQL2
                 '           Next iCount
                  '  End If
                   ' Rs3.MoveNext
               ' Wend
        'End If
        'Rs3.Close
'ElseIf BartendorForm = "R5-3" Then
'strCreate_Temp = "select * into dbo.Report_Temp2 From NP_LABEL_TEMP where UsedUser = 'Null'"
'ConnSQL.Execute strCreate_Temp
'strSQL = "Select * From NP_Label_Temp Where UsedUser = '" & strUsername & "' "
 '       Rs3.Open strSQL, ConnSQL, adOpenDynamic, adLockOptimistic
  '      If Not Rs3.EOF Then
   '             Rs3.MoveFirst
    '            While Not Rs3.EOF
     '               If Int(Rs3!QTY) > 0 Then
      '                      strSQL2 = "Insert Into dbo.Report_Temp2(ItemCode, barcode, NAME1, NAME2, QTY, PriceLevel, Price, UnitCode, UsedUser, Category_ID, WHCode, ShelfCode, VENDR_ID, remark, SPrice,OnHand,RemainOutQTY,RemainInQTY,SOPNUM,SOPDOC) " _
       '                                 & " values('" & Trim(Rs3!Itemcode) & "','" & Trim(Rs3!barcode) & "','" & Trim(Rs3!Name1) & "', '" & Trim(Rs3!Name2) & "', 1,'" & Trim(Rs3!PriceLevel) & "','" & Trim(Rs3!price) & "','" & Trim(Rs3!Unitcode) & "'," _
        '                                & "'" & Trim(Rs3!UsedUser) & "','" & Trim(Rs3!Category_ID) & "','" & Trim(Rs3!whcode) & "', '" & Trim(Rs3!ShelfCode) & "','" & Trim(Rs3!VENDR_ID) & "', '" & Trim(Rs3!Remark) & "','" & Trim(Rs3!SPrice) & "' ,'" & Rs3!onhand & "','" & Rs3!RemainOutQTY & "','" & Rs3!RemainInQTY & "','" & Rs3!SOPNUM & "','" & Rs3!sopdoc & "') "
         '                   For iCount = 1 To Rs3!QTY
           '                         ConnSQL.Execute strSQL2
          '                  Next iCount
            '        End If
             '       Rs3.MoveNext
              '  Wend
        'End If
        'Rs3.Close
ElseIf BartendorForm = "R8-1" Then
strCreate_Temp = "select * into dbo.Report_Temp3 From NP_LABEL_TEMP where UsedUser = 'Null'"
        ConnSQL.Execute strCreate_Temp
        strSQL = "Select * From NP_Label_Temp Where UsedUser = '" & strUsername & "' "
        Rs3.Open strSQL, ConnSQL, adOpenDynamic, adLockOptimistic
        If Not Rs3.EOF Then
                Rs3.MoveFirst
                While Not Rs3.EOF
                    If Int(Rs3!QTY) > 0 Then
                            strSQL2 = "Insert Into dbo.Report_Temp3(ItemCode, barcode, NAME1, NAME2, QTY, PriceLevel, Price, UnitCode, UsedUser, Category_ID, WHCode, ShelfCode, VENDR_ID, remark, SPrice,SOPDOC,RemainOutQTY,RemainInQTY) " _
                                        & " values('" & Trim(Rs3!Itemcode) & "','" & Trim(Rs3!barcode) & "','" & Trim(Rs3!Name1) & "', '" & Trim(Rs3!Name2) & "', 1,'" & Trim(Rs3!PriceLevel) & "','" & Trim(Rs3!price) & "','" & Trim(Rs3!Unitcode) & "'," _
                                        & "'" & Trim(Rs3!UsedUser) & "','" & Trim(Rs3!Category_ID) & "','" & Trim(Rs3!whcode) & "', '" & Trim(Rs3!ShelfCode) & "','" & Trim(Rs3!VENDR_ID) & "', '" & Trim(Rs3!Remark) & "','" & Trim(Rs3!SPrice) & "' ,'" & Rs3!sopdoc & "','" & Rs3!RemainOutQTY & "','" & Rs3!RemainInQTY & "') "
                            For iCount = 1 To Rs3!QTY
                                    ConnSQL.Execute strSQL2
                            Next iCount
                    End If
            '     Rs3.MoveNext
                Wend
        End If
        Rs3.Close
End If


'If BartendorForm <> "R1-1" And BartendorForm <> "R1-2" And BartendorForm <> "R1-3" Then
    'ImportBartendor.CreateTextFile ("\\192.168.2.46\bartendor\" & BartendorForm & ".txt")
    ImportBartendor.CreateTextFile ("\\label.nopadol.com\bartendor\" & BartendorForm & ".txt")
'Else
 '   ImportBartendor.CreateTextFile ("\\192.168.2.85\bartendor\" & BartendorForm & ".txt")
'End If

'ImportBartendor.CreateTextFile ("\\S2GR1P\bartendor\" & BartendorForm & ".txt")
'ImportBartendor.CreateTextFile ("C:\bartendor\" & BartendorForm & ".txt")

        MsgBox "อีกประมาณ 1 นาที ไปเอาป้ายราคาได้ที่เครื่องพิมพ์บาร์โค้ดครับ", vbInformation, "ข้อความแจ้งให้ทราบ"
        CMDBartendor.Visible = False
        cmdPreview.Visible = False
strSQL = "delete  NP_Label_Temp Where UsedUser = '" & strUsername & "' "
ConnSQL.Execute strSQL
        
ErrDescription:
    If Err.Description <> "" Then
        If BartendorForm = "R2-8" Then
            strSQL = "Drop Table Report_Temp1 "
            ConnSQL.Execute strSQL
        ElseIf BartendorForm = "R1-1" Or BartendorForm = "R1-2" Or BartendorForm = "R1-3" Or BartendorForm = "R1-4" Or BartendorForm = "R1-5" Or BartendorForm = "R1-6" Or BartendorForm = "R5-1" Or BartendorForm = "R5-2" Or BartendorForm = "R5-3" Or BartendorForm = "R5-4" Or BartendorForm = "R5-5" Or BartendorForm = "R5-7" Or BartendorForm = "R5-8" Then
            strSQL = "Drop Table  tempdb.dbo.Report_Temp2 "
            ConnSQL.Execute strSQL
       ElseIf BartendorForm = "R8-1" Then
            strSQL = "Drop Table  Report_Temp3 "
            ConnSQL.Execute strSQL
        End If
            MsgBox "กดปุ่มพิมพ์ Bartendor อีกครั้งนะครับ"
    End If

End Sub


Public Function GenFormatBarCode(vGetBarCode As String) As String
Dim vLenBarCode As Integer
Dim vBarCode As String

vBarCode = Trim(vGetBarCode)
vLenBarCode = Len(vBarCode)

If vLenBarCode = 6 Then
GenFormatBarCode = vBarCode
ElseIf vLenBarCode = 7 Then
GenFormatBarCode = Left(vBarCode, 1) + " " + Right(vBarCode, 6)
ElseIf vLenBarCode = 8 Then
GenFormatBarCode = Left(vBarCode, 1) + " " + Mid(vBarCode, 2, 6) + " " + Right(vBarCode, 1)
ElseIf vLenBarCode = 9 Then
GenFormatBarCode = Left(vBarCode, 1) + " " + Mid(vBarCode, 2, 6) + " " + Right(vBarCode, 2)
ElseIf vLenBarCode = 10 Then
GenFormatBarCode = Left(vBarCode, 1) + " " + Mid(vBarCode, 2, 6) + " " + Right(vBarCode, 3)
ElseIf vLenBarCode = 11 Then
GenFormatBarCode = Left(vBarCode, 1) + " " + Mid(vBarCode, 2, 6) + " " + Right(vBarCode, 4)
ElseIf vLenBarCode = 12 Then
GenFormatBarCode = Left(vBarCode, 1) + " " + Mid(vBarCode, 2, 6) + " " + Right(vBarCode, 5)
ElseIf vLenBarCode = 13 Then
GenFormatBarCode = Left(vBarCode, 1) + " " + Mid(vBarCode, 2, 6) + " " + Right(vBarCode, 6)
ElseIf vLenBarCode = 14 Then
GenFormatBarCode = Left(vBarCode, 1) + " " + Mid(vBarCode, 2, 6) + " " + Mid(vBarCode, 8, 6) + " " + Right(vBarCode, 1)
ElseIf vLenBarCode = 15 Then
GenFormatBarCode = Left(vBarCode, 1) + " " + Mid(vBarCode, 2, 6) + " " + Mid(vBarCode, 8, 6) + " " + Right(vBarCode, 2)
ElseIf vLenBarCode >= 16 Then
GenFormatBarCode = vBarCode
End If

End Function
Private Sub cmdCancel_Click()
Dim tmpPathName, strCreate_Temp, strDel_Temp As String
Dim strSQL, strSQL2 As String
Dim BartendorForm As String

On Error Resume Next
    tmpPathName = Trim(LV_Report.SelectedItem.SubItems(1)) & ".rpt"
    BartendorForm = Mid(tmpPathName, 21, 4)
        
        strSQL = "Delete From NP_Label_Temp Where Useduser = '" & strUsername & "' "
        ConnSQL.Execute strSQL
        
    If BartendorForm = "R2-8" Then
    strSQL = "Drop Table Report_Temp1"
    ConnSQL.Execute strSQL
    'ConnSQL.Close
    ElseIf BartendorForm = "R5-1" Then
    strSQL = "Drop Table Report_Temp2"
    ConnSQL.Execute strSQL
    ElseIf BartendorForm = "R5-2" Then
    strSQL = "Drop Table Report_Temp2"
    ConnSQL.Execute strSQL
    ElseIf BartendorForm = "R5-3" Then
    strSQL = "Drop Table Report_Temp2"
    ConnSQL.Execute strSQL
    'ConnSQL.Close
    ElseIf BartendorForm = "R5-4" Then
    strSQL = "Drop Table Report_Temp2"
    ConnSQL.Execute strSQL
    ElseIf BartendorForm = "R5-5" Then
    strSQL = "Drop Table Report_Temp2"
    ConnSQL.Execute strSQL
    ElseIf BartendorForm = "R5-7" Then
    strSQL = "Drop Table Report_Temp2"
    ConnSQL.Execute strSQL
    ElseIf BartendorForm = "R8-1" Then
    strSQL = "Drop Table Report_Temp3"
    ConnSQL.Execute strSQL
    End If
        ' ลบ Table Report_Temp
        'strSQL = "delete  Report_Temp Where Useduser = '" & strUsername & "' "
        'ConnSQL.Execute strSQL
        
        'strSQL = "delete  Report_Temp1  Where Useduser = '" & strUsername & "' "
        'ConnSQL.Execute strSQL
        'strSQL = "delete  Report_Temp2  Where Useduser = '" & strUsername & "' "
        'ConnSQL.Execute strSQL
        ConnSQL.Close
        Unload Me

        
'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description
     'Unload Me
'end If
End Sub

Private Sub CMDChangePrice_Click()
Dim vQuery As String
Dim vWHCode As String
Dim ListChangePrice As ListItem
Dim vCount As Integer

On Error Resume Next
'LV_ChangePrice1.ListItems.Clear
'vWHCode = Trim(CMBWHCode.Text)
'vQuery = " SELECT * from bcnp.dbo.vw_PRG_SearchItemChangePrice where secman = '" & vWHCode & "' "
'Rs1.Open vQuery, ConnSQL, adOpenDynamic, adLockOptimistic
'If Not Rs1.EOF Then
 '   Rs1.MoveFirst
  '  While Not Rs1.EOF
   ' Set ListChangePrice = LV_ChangePrice1.ListItems.Add(, , Trim(Rs1.Fields("itemcode").Value))
    '
    'ListChangePrice.SubItems(1) = Trim(Rs1.Fields("barcode").Value)
    'If IsNull(Rs1.Fields("name1").Value) Then
     '   ListChangePrice.SubItems(2) = ""
    'Else
     '   ListChangePrice.SubItems(2) = Trim(Rs1.Fields("name1").Value)
    'End If
    'ListChangePrice.SubItems(3) = Trim(Rs1.Fields("unitcode").Value)
    'ListChangePrice.SubItems(4) = Trim(Rs1.Fields("SalePrice1").Value)
    'ListChangePrice.SubItems(5) = 1
    ''ListChangePrice.SubItems(6) = Trim(Rs1.Fields("dateupdate").Value)
    'Rs1.MoveNext
    'Wend
'End If
'Rs1.Close

End Sub

Private Sub cmdDeselect_Click()
        Dim j As Integer
        
        On Error GoTo ErrDescription
        
        For j = 1 To ListResult.ListItems.count
                ListResult.ListItems(j).Checked = False
        Next
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub cmdFinish_Click()

On Error GoTo ErrDescription

        FormToPrinter = True        ' Printer
        Call ProcessPrint
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDCloseHandHeld_Click()
PicWizard(IntStep).Visible = False
PicWizard(1).Visible = True
ListView1.SetFocus
IntStep = 1
End Sub


Private Sub CMDHandHeldClose_Click()
PicWizard(IntStep).Visible = False
PicWizard(1).Visible = True
ListView1.SetFocus
IntStep = 1
End Sub

Private Sub CMDImportData_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem
Dim vPromotion As String


If Me.CMBPromotionCode.Text <> "" Then

vPromotion = Left(Trim(Me.CMBPromotionCode.Text), InStr(Trim(CMBPromotionCode.Text), "//") - 1)

vQuery = "exec dbo.USP_PM_PromotionExpire '" & vPromotion & "' "
If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
  ListViewItemChangePrice.ListItems.Clear
  vRecordset.MoveFirst
  i = 1
  While Not vRecordset.EOF
  Set vListItem = ListViewItemChangePrice.ListItems.Add(, , i)
  vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
  vListItem.SubItems(2) = Trim(vRecordset.Fields("barcode").Value)
  vListItem.SubItems(3) = Trim(vRecordset.Fields("itemname").Value)
  vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
  vListItem.SubItems(5) = Trim(vRecordset.Fields("price").Value)
  vListItem.SubItems(6) = Trim(1)
  vListItem.SubItems(7) = Trim(vRecordset.Fields("dateend").Value)
  vListItem.SubItems(8) = Trim(vRecordset.Fields("whcode").Value)
  vListItem.SubItems(9) = Trim(vRecordset.Fields("priceerect").Value)
  vListItem.SubItems(10) = Trim(vRecordset.Fields("shelfcode").Value)
  vRecordset.MoveNext
  i = i + 1
  Wend
End If
vRecordset.Close
End If



'vQuery = "exec dbo.USP_NP_SelectItemChangePricePrintLabel"
'If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
 ' ListViewItemChangePrice.ListItems.Clear
  'vRecordset.MoveFirst
  'i = 1
  'While Not vRecordset.EOF
  'Set vListItem = ListViewItemChangePrice.ListItems.Add(, , i)
  'vListItem.SubItems(1) = Trim(vRecordset.Fields("barcode").Value)
  'vListItem.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
  'vListItem.SubItems(3) = Trim(vRecordset.Fields("newprice").Value)
  'vListItem.SubItems(4) = Trim(vRecordset.Fields("oldprice").Value)
  'vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
  'vListItem.SubItems(6) = Trim(vRecordset.Fields("itemname").Value)
  'vListItem.SubItems(7) = Trim(vRecordset.Fields("docno").Value)
  'vRecordset.MoveNext
  'i = i + 1
  'Wend
'End If
'vRecordset.Close
End Sub

Private Sub cmdPreview_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vLabelName As String
Dim vName1 As String
Dim tmpPathName As String
Dim i As Integer
Dim vItemCode As String
Dim vCheckTypePrint As Integer
Dim vRequestNo As String
Dim vCheckType As String
        
On Error GoTo ErrDescription

FormToPrinter = False

tmpPathName = UCase(Left(Right(Trim(LV_Report.SelectedItem.SubItems(1)), 5), 2))

If tmpPathName = "SP" Then
   Call CheckPriceErect
Else
   vCountCheckPriceErect = 0
End If

'====================================================================================================================================
'====================================================================================================================================
'====================================================================================================================================
vCheckType = "พิมพ์ป้ายราคา"

If vCheckType = "พิมพ์ป้ายราคา" Then
    For i = 1 To Me.ListResult.ListItems.count
    If ListResult.ListItems.Item(i).Checked = True Then
    vItemCode = ListResult.ListItems.Item(i).Text
    
    vQuery = "exec dbo.USP_PM_ItemCheckPromotion '" & vItemCode & "' "
    If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
        vRequestNo = Trim(vRecordset.Fields("requestno").Value)
        vCheckTypePrint = 1
        Else
        vCheckTypePrint = 0
        vRequestNo = ""
    End If
    vRecordset.Close
    
    If vCheckTypePrint = 1 Then
        MsgBox "สินค้า รหัส " & vItemCode & " เป็นสินค้าโปรโมชั่น อยู่ในเอกสารเลขที่ " & vRequestNo & " ไม่สามารถพิมพ์ป้ายราคาเกี่ยวกับราคาปกติได้ ต้องลบรายการสินค้าดังกล่าวออกจากรายการพิมพ์", vbCritical, "Send Error Message"
        Exit Sub
    End If
    End If
    Next i
End If

'====================================================================================================================================
'====================================================================================================================================
'====================================================================================================================================

If vCountCheckPriceErect = 0 Then
   Me.Check104.Value = 1
   Call ProcessPrint
   cmdPreview.Visible = False
   CMDBartendor.Visible = False
Else
   vQuery = "exec dbo.USP_NP_DeleteDataPrintLabel '" & strUsername & "' "
   vConnection.Execute vQuery
   
   cmdNext.Enabled = True
   cmdFinish.Visible = False
   cmdPreview.Visible = False
   PicWizardReport.Visible = False
   PicWizardSelect.Visible = True
   CMDBartendor.Visible = False
   Call DelNPTemp
   IntStep = IntStep - 1
   
   Me.Check104.Value = 0
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CheckBarCodeType()

End Sub

Public Sub CheckPriceErect()
Dim i As Integer
Dim vPrice As Double
Dim vCheckPrice As Double
Dim vItemError As String

vCountCheckPriceErect = 0
For i = 1 To Me.ListResult.ListItems.count
   If Me.ListResult.ListItems(i).Checked = True Then
      vItemError = Me.ListResult.ListItems(i).SubItems(2)
      vPrice = Me.ListResult.ListItems(i).SubItems(5)
      vCheckPrice = Me.ListResult.ListItems(i).SubItems(7)
      If vPrice >= vCheckPrice Then
          vCountCheckPriceErect = vCountCheckPriceErect + 1
          MsgBox "ราคาตั้งที่ได้กำหนดไว้ของรหัสสินค้า " & vItemError & "  มีราคาน้อยกว่าหรือเท่ากับ ราคาที่จะขายจริง หรือ ราคาตั้งของหน่วยที่จะขายไม่ได้ถูกกำหนด กรุณาแจ้งจัดซื้อผู้ที่ดูแลสินค้าดังกล่าว", vbCritical, "Send Error Message"
          Me.ListResult.ListItems(i).Checked = False
      End If
   End If
Next i
End Sub

Private Sub ProcessPrint()
Dim tmpPathName, strCreate_Temp, strDel_Temp, Name1, Name2 As String
Dim strSQL, strSQL2 As String
Dim iCount As Integer
Dim i As Integer
Dim vItemCode As String
Dim vUnitCode As String
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vQuery As String

'On Error GoTo ErrDescription
        
        tmpPathName = Trim(LV_Report.SelectedItem.SubItems(1)) & ".rpt"
        Name1 = Mid(LV_Report.SelectedItem.SubItems(1), 21, 10)
        '------------------------------------------------------------------------------------
        If Trim(Name1) = "R2-8" Or Trim(Name1) = "R5-1" Or Trim(Name1) = "R5-2" Or Trim(Name1) = "R5-3" Or Trim(Name1) = "R8-1" Or Trim(Name1) = "R5-4" Or Trim(Name1) = "R5-5" Or Trim(Name1) = "R5-7" Or Trim(Name1) = "R5-8" Then
        MsgBox "ป้ายราคาจะพิมพ์ออกเครื่องพิมพ์ BarTendor ", vbInformation, "แจ้งเตือนการพิมพ์ป้ายราคา"
        CMDBartendor_Click
        Exit Sub
        End If
        
        If Trim(Name1) = "R8-6" Or Trim(Name1) = "R8-7" Then
        Call PrintForm1
        Exit Sub
        End If
        '------------------------------------------------------------------------------------
        'check การปรับราคาสินค้า
        If Check104.Value = 1 Then
            For i = 1 To ListResult.ListItems.count
            vItemCode = Trim(ListResult.ListItems.Item(i).Text)
            vUnitCode = Trim(ListResult.ListItems.Item(i).SubItems(4))
            vQuery = "exec usp_IV_UpdatePrintUpdateChangePrice '" & vItemCode & "','" & vUnitCode & "','" & strUsername & "','" & Name1 & "' "
            vConnection.Execute vQuery
            Next i
        End If
        ' Check ราคาพิเศษสำหรับโชว์รูม Or ราคาพิเศษ 2 ใบ/หน้า
        'If Me.LV_Report.SelectedItem.Index = 6 Or Me.LV_Report.SelectedItem.Index = 14 Then
         '   AddTempPO = False
          '  AddTempRCV = False
           ' If AddTempPO = True Or AddTempRCV = True Then
            '    MsgBox "คุณไม่สามารถเลือกรายการพิมพ์ รายการนี้ได้", vbOKOnly + vbInformation, "คำแนะนำ"
             '   Exit Sub
            'Else
             '   frmWizard.Enabled = False
              '  frmSPrice.Show
               ' Exit Sub
            'End If
        'End If
        ' Dump Data to Report_Temp
        vQuery = "exec dbo.USP_NP_SearchPrintQTY '" & strUsername & "' "
        If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
                    If Int(vRecordset.Fields("qty").Value) > 0 Then
                            vQuery = "exec dbo.USP_NP_InsertPrintLabel '" & Trim(vRecordset.Fields("Itemcode").Value) & "'," _
                            & " '" & Trim(vRecordset.Fields("barcode").Value) & "','" & Trim(vRecordset.Fields("Name1").Value) & "',  " _
                            & " '" & Trim(vRecordset.Fields("Name2").Value) & "', 1," & Trim(vRecordset.Fields("PriceLevel").Value) & ", " _
                            & " " & Trim(vRecordset.Fields("price").Value) & ",'" & Trim(vRecordset.Fields("Unitcode").Value) & "', " _
                            & " '" & Trim(vRecordset.Fields("UsedUser").Value) & "','" & Trim(vRecordset.Fields("Category_ID").Value) & "', " _
                            & " '" & Trim(vRecordset.Fields("whcode").Value) & "', '" & Trim(vRecordset.Fields("ShelfCode").Value) & "', " _
                            & " '" & Trim(vRecordset.Fields("VENDR_ID").Value) & "', '" & Trim(vRecordset.Fields("Remark").Value) & "', " _
                            & " " & Trim(vRecordset.Fields("SPrice").Value) & ",'" & vRecordset.Fields("onhand").Value & "', " _
                            & " '" & vRecordset.Fields("RemainOutQTY").Value & "','" & vRecordset.Fields("RemainInQTY").Value & "', " _
                            & " '" & vRecordset.Fields("SOPNUM").Value & "', '" & vRecordset.Fields("sopdoc").Value & "' "
                            For iCount = 1 To Int(vRecordset.Fields("qty").Value)
                                    vConnection.Execute vQuery
                            Next iCount
                    
                    End If
                    vRecordset.MoveNext
                Wend
        End If
        vRecordset.Close
        
        With Crystal101
        .ReportFileName = tmpPathName
        .ParameterFields(0) = "@vUserID;" & tmpUsedUser & ";true"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
        End With
        vQuery = "exec dbo.USP_NP_DeleteDataPrintLabel '" & strUsername & "' "
        vConnection.Execute vQuery
        Check104.Value = 0

'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description
'End If
End Sub

Private Sub cmdPrevious_Click()
        
        On Error Resume Next
        
        Select Case IntStep
                Case 1:
                        cmdNext.Visible = False
                        cmdPrevious.Visible = False
                        PicWizard(IntStep).Visible = False
                        PicWizard(IntStep - 1).Visible = True
                        IntStep = IntStep - 1
                        txtUsername = ""
                        txtPassword = ""
                        txtUsername.SetFocus
                Case 2:         ' Now เลือกแบบบันทึกเอง --> Back to Select Job
                        PicWizard(IntStep).Visible = False      ' Invisible Now Step
                        PicWizard(1).Visible = True
                        CMD101.Visible = False
                        ListView1.SetFocus
                        IntStep = 1
                Case 3:         ' Now เลือกพิมพ์ป้ายราคาจากเครื่อง Palm  --> Back to Select Job
                        PicWizard(IntStep).Visible = False      ' Invisible Now Step
                        PicWizard(1).Visible = True
                        ListView1.SetFocus
                        IntStep = 1
                Case 4:         ' Now เลือกใบสั่งซื้อ --> Back to Select Job
                        PicWizard(IntStep).Visible = False      ' Invisible Now Step
                        PicWizard(1).Visible = True
                        ListView1.SetFocus
                        IntStep = 1
                        AddTempPO = False
                Case 5:         ' Now เลือกใบรับสินค้า --> Back to Select Job
                        PicWizard(IntStep).Visible = False      ' Invisible Now Step
                        PicWizard(1).Visible = True
                        ListView1.SetFocus
                        IntStep = 1
                        AddTempRCV = False
                        vCheckReceipt = False
                Case 6:         ' Now Palm Check Value --> Back to Select Job
                        PicWizard(IntStep).Visible = False      ' Invisible Now Step
                        PicWizard(1).Visible = True
                        ListView1.SetFocus
                        IntStep = 1
                        AddTempRCV = False
                Case 7:         ' Now ListResult --> Back to JobSelect
                        PicWizard(IntStep).Visible = False      ' Invisible Now Step
                        PicWizard(1).Visible = True
                        ListView1.SetFocus
                        IntStep = 1
                        AddTempRCV = False
                Case 8:         ' Now ListResult --> Back to JobSelect
                        PicWizard(IntStep).Visible = False      ' Invisible Now Step
                        PicWizard(1).Visible = True
                        ListView1.SetFocus
                        IntStep = 1
                        AddTempRCV = False
                Case 9:         ' Now ListResult --> Back to JobSelect
                        PicWizard(IntStep).Visible = False      ' Invisible Now Step
                        PicWizard(1).Visible = True
                        ListView1.SetFocus
                        IntStep = 1
                        AddTempRCV = False
                Case 10:         ' Now ListResult --> Back to JobSelect
                        PicWizard(IntStep).Visible = False      ' Invisible Now Step
                        PicWizard(1).Visible = True
                        ListView1.SetFocus
                        LV_ChangePrice1.ListItems.Clear
                        LV_ChangePrice2.ListItems.Clear
                        AddTempRCV = False
                        vChangePrice = False
                        IntStep = 1
                Case 11:         ' Now ListResult --> Back to JobSelect
                        PicWizard(IntStep).Visible = False      ' Invisible Now Step
                        PicWizard(1).Visible = True
                        ListView1.SetFocus
                        LV_Promo1.ListItems.Clear
                        LV_Promo2.ListItems.Clear
                        CMB101.Clear
                        CMB102.Clear
                        TXTShelf1.Text = ""
                        TXTShelf2.Text = ""
                        Check101.Value = 0
                        AddTempRCV = False
                        vShelfJob = False
                        IntStep = 1
                Case 12:
                        PicWizardSelect.Visible = False
                        cmdNext.Enabled = True
                        If IntStep1 = 2 Then
                        PicWizard(IntStep - 10).Visible = True
                        CMD101.Visible = True
                        ElseIf IntStep1 = 3 Then
                        PicWizard(IntStep - 9).Visible = True
                        ElseIf IntStep1 = 4 Then
                        PicWizard(IntStep - 8).Visible = True
                        ElseIf IntStep1 = 5 Then
                        PicWizard(IntStep - 7).Visible = True
                        ElseIf IntStep1 = 6 Then
                        PicWizard(IntStep - 6).Visible = True
                        ElseIf IntStep1 = 7 Then
                        PicWizard(IntStep - 5).Visible = True
                        ElseIf IntStep1 = 8 Then
                        PicWizard(IntStep - 4).Visible = True
                        ElseIf IntStep1 = 9 Then
                        PicWizard(IntStep - 3).Visible = True
                        ElseIf IntStep1 = 10 Then
                        PicWizard(IntStep - 2).Visible = True
                        ElseIf IntStep1 = 11 Then
                        PicWizard(IntStep - 1).Visible = True
                        'ElseIf IntStep1 = 12 Then
                        'PicWizard(IntStep - 2).Visible = True
                        'LV_ChangePrice1.ListItems.Clear
                        'LV_ChangePrice2.ListItems.Clear
                        'LV_Promo1.ListItems.Clear
                        'LV_Promo2.ListItems.Clear
                        End If
                        ListResult.ListItems.Clear
                        IntStep = IntStep1
                        
                Case 13:       ' Now Form Select Form Print --> Back to ListResult
                        cmdNext.Enabled = True
                        cmdFinish.Visible = False
                        cmdPreview.Visible = False
                        PicWizardReport.Visible = False         ' Invisible Now Step
                        PicWizardSelect.Visible = True
                        CMDBartendor.Visible = False
                        Call DelNPTemp
                        IntStep = IntStep - 1
        End Select
        
        ' Get Detail
        Call GetDetail
        
'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description
  '  Exit Sub
'End If
End Sub
Private Sub DelNPTemp()
Dim tmpPathName As String
Dim BartendorForm As String
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim strSQL As String
        
'On Error Resume Next
tmpPathName = Trim(LV_Report.SelectedItem.SubItems(1)) & ".rpt"
BartendorForm = Mid(tmpPathName, 21, 4)
vQuery = "exec dbo.USP_NP_DeleteDataPrintLabel '" & strUsername & "' "
vConnection.Execute vQuery
If BartendorForm = "R2-8" Then
strSQL = "Drop Table Report_Temp1 "
ConnSQL.Execute strSQL
ElseIf BartendorForm = "R5-1" Then
strSQL = "Drop Table  tempdb.dbo.Report_Temp2 "
ConnSQL.Execute strSQL
ElseIf BartendorForm = "R5-2" Then
strSQL = "Drop Table  tempdb.dbo.Report_Temp2 "
ConnSQL.Execute strSQL
ElseIf BartendorForm = "R5-3" Then
strSQL = "Drop Table  tempdb.dbo.Report_Temp2 "
ConnSQL.Execute strSQL
ElseIf BartendorForm = "R5-4" Then
strSQL = "Drop Table  tempdb.dbo.Report_Temp2 "
ConnSQL.Execute strSQL
ElseIf BartendorForm = "R5-5" Then
strSQL = "Drop Table  tempdb.dbo.Report_Temp2 "
ConnSQL.Execute strSQL
ElseIf BartendorForm = "R5-7" Then
strSQL = "Drop Table  tempdb.dbo.Report_Temp2 "
ConnSQL.Execute strSQL
ElseIf BartendorForm = "R8-1" Then
strSQL = "Drop Table  Report_Temp3 "
ConnSQL.Execute strSQL
End If
End Sub

Private Sub CMDPrintHandHeld_Click()
Dim tmpPathName, strCreate_Temp, strDel_Temp, Name1, Name2 As String
Dim strSQL, strSQL2 As String
Dim iCount As Integer
Dim i As Integer
Dim vItemCode As String
Dim vUnitCode As String
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vRecordset2 As New ADODB.Recordset
Dim vQuery As String
Dim vFormName As String
Dim vCheckFileName As String
Dim vBarCode As String
Dim vWHCode As String
Dim strItemNumber  As String
Dim vType As String
Dim n As Integer
Dim vMemCountCheck As Integer

On Error GoTo ErrDescription

If Me.LV_Palm.ListItems.count = 0 Then
Exit Sub
End If

For n = 1 To Me.LV_Palm.ListItems.count
If Me.LV_Palm.ListItems(n).Checked = True Then
vMemCountCheck = vMemCountCheck + 1
End If
Next n

If vMemCountCheck = 0 Then
MsgBox "กรุณาเลือกรายการสินค้าที่จะพิมพ์ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Exit Sub
End If

vQuery = "exec dbo.USP_NP_CheckPrintLabel"
If OpenDatabase(vConnection, vRecordset1, vQuery) <> 0 Then
vRecordset1.MoveFirst
While Not vRecordset1.EOF
vFormName = Trim(vRecordset1.Fields("pathname").Value)

If Me.LV_Palm.ListItems.count > 0 Then

For i = 1 To Me.LV_Palm.ListItems.count
If Me.LV_Palm.ListItems(i).Checked = True Then
vCheckFileName = Me.LV_Palm.ListItems.Item(i).SubItems(7)

If vFormName = vCheckFileName Then

vUnitCode = Trim(LV_Palm.ListItems.Item(i).SubItems(3))
vItemCode = Trim(LV_Palm.ListItems.Item(i).Text)
vBarCode = Trim(LV_Palm.ListItems.Item(i).SubItems(1))
vWHCode = "S02"

vQuery = "exec dbo.USP_NP_SearchItemDetails_Market '" & vItemCode & "' ,'" & vUnitCode & "','" & vBarCode & "','" & vWHCode & "' "
If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
tmpDocNo = ""
tmpDocDate = ""
strItemNumber = Trim(vRecordset.Fields("Itemcode").Value)
tmpNUMBR = Trim(vRecordset.Fields("Itemcode").Value)
If Me.LV_Palm.ListItems(i).SubItems(1) = Null Then
tmpBarcode = Trim(vRecordset.Fields("barcode").Value)
Else
tmpBarcode = Me.LV_Palm.ListItems(i).SubItems(1)
End If
tmpTHINAME = Trim(vRecordset.Fields("name1").Value)
If Right(tmpTHINAME, 1) = "'" Then
tmpTHINAME = Left(tmpTHINAME, (Len(tmpTHINAME) - 1)) & """"
End If
tmpENGNAME = Trim(vRecordset.Fields("name2").Value)
tmpUOM = Trim(vRecordset.Fields("unitcode").Value)
tmpUsedUser = strUsername
tmpSite = ""

If tmpBIN_ID = "" Then
tmpBIN_ID = Trim(vRecordset.Fields("shelfcode").Value)
End If

tmpONHAND = Trim(vRecordset.Fields("qty").Value)
tmpQTYALLOCATE = Trim(vRecordset.Fields("reserveqty").Value)
tmpCategory_ID = ""
tmpRemainOutQTY = Trim(vRecordset.Fields("RemainOutQTY").Value)
tmpRemainInQTY = Trim(vRecordset.Fields("RemainInQTY").Value)
tmpVENDR_ID = ""
tmpRRCLV = Trim(vRecordset.Fields("salepromotion").Value)
tmpRRC = Trim(vRecordset.Fields("saleprice1").Value)
tmpSPrice = CheckDegit(Trim(vRecordset.Fields("priceerect").Value))
                        
Else
MsgBox "สินค้ารหัส " & vItemCode & " และคลัง " & vWHCode & " มีข้อมูลไม่ครบไม่สามารถพิมพ์เอกสารได้ กรุณาตรวจสอบ บาร์โค้ด ที่เก็บ ราคาตั้ง ของสินค้าดังกล่าว"
tmpDocNo = ""
tmpDocDate = ""
strItemNumber = ""
tmpNUMBR = ""
tmpBarcode = ""
tmpTHINAME = ""
tmpTHINAME = ""
tmpENGNAME = ""
tmpUOM = ""
tmpSite = ""
tmpBIN_ID = ""
tmpONHAND = ""
tmpQTYALLOCATE = ""
tmpCategory_ID = ""
tmpRemainOutQTY = ""
tmpRemainInQTY = ""
tmpVENDR_ID = ""
tmpRRCLV = 0
tmpRRC = 0
tmpSPrice = ""
End If
vRecordset.Close

tmpQTY = Int(LV_Palm.ListItems(i).SubItems(5))
If IsNull(tmpSPrice) Or tmpSPrice = "" Then
tmpSPrice = 0
End If

vQuery = "exec dbo.USP_NP_InsertLabelTemp '" & tmpNUMBR & "','" & tmpBarcode & "','" & tmpTHINAME & "','" & tmpENGNAME & "'," & Int(tmpQTY) & "," & tmpRRCLV & "," & tmpRRC & ",'" & tmpUOM & "','" & tmpUsedUser & "','" & tmpCategory_ID & "','" & tmpSite & "','" & tmpBIN_ID & "','" & tmpVENDR_ID & "','" & tmpRemark & "'," & Int(tmpSPrice) & ",'" & tmpONHAND & "','" & tmpQTYALLOCATE & "'," & tmpTYPE & ",'" & tmpRemainOutQTY & "','" & tmpRemainInQTY & "','" & tmpDocNo & "','" & tmpDocDate & "' "
vConnection.Execute vQuery
End If
End If
Next i


vQuery = "exec dbo.USP_NP_SearchPrintQTY '" & strUsername & "' "
If OpenDatabase(vConnection, vRecordset2, vQuery) <> 0 Then
vRecordset2.MoveFirst
While Not vRecordset2.EOF
If Int(vRecordset2.Fields("qty").Value) > 0 Then
vQuery = "exec dbo.USP_NP_InsertPrintLabel '" & Trim(vRecordset2.Fields("Itemcode").Value) & "'," _
& " '" & Trim(vRecordset2.Fields("barcode").Value) & "','" & Trim(vRecordset2.Fields("Name1").Value) & "',  " _
& " '" & Trim(vRecordset2.Fields("Name2").Value) & "', 1," & Trim(vRecordset2.Fields("PriceLevel").Value) & ", " _
& " " & Trim(vRecordset2.Fields("price").Value) & ",'" & Trim(vRecordset2.Fields("Unitcode").Value) & "', " _
& " '" & Trim(vRecordset2.Fields("UsedUser").Value) & "','" & Trim(vRecordset2.Fields("Category_ID").Value) & "', " _
& " '" & Trim(vRecordset2.Fields("whcode").Value) & "', '" & Trim(vRecordset2.Fields("ShelfCode").Value) & "', " _
& " '" & Trim(vRecordset2.Fields("VENDR_ID").Value) & "', '" & Trim(vRecordset2.Fields("Remark").Value) & "', " _
& " " & Trim(vRecordset2.Fields("SPrice").Value) & ",'" & vRecordset2.Fields("onhand").Value & "', " _
& " '" & vRecordset2.Fields("RemainOutQTY").Value & "','" & vRecordset2.Fields("RemainInQTY").Value & "', " _
& " '" & vRecordset2.Fields("SOPNUM").Value & "', '" & vRecordset2.Fields("sopdoc").Value & "' "
For iCount = 1 To Int(vRecordset2.Fields("qty").Value)
vConnection.Execute vQuery
Next iCount

End If
vRecordset2.MoveNext
Wend
End If
vRecordset2.Close

With Crystal101
.ReportFileName = vFormName & ".rpt"
.ParameterFields(0) = "@vUserID;" & tmpUsedUser & ";true"
.WindowState = crptMaximized
.Destination = crptToWindow
.Action = 1
End With

vQuery = "exec dbo.USP_NP_DeleteDataPrintLabel '" & strUsername & "' "
vConnection.Execute vQuery

End If

vRecordset1.MoveNext
Wend
End If
vRecordset1.Close

'------------------------------------------------------------------------------------
'check การปรับราคาสินค้า

For i = 1 To LV_Palm.ListItems.count
vBarCode = Trim(LV_Palm.ListItems.Item(i).SubItems(1))
vUnitCode = Trim(LV_Palm.ListItems.Item(i).SubItems(3))
Name1 = Trim(LV_Palm.ListItems.Item(i).SubItems(7))
vType = Trim(LV_Palm.ListItems.Item(i).SubItems(6))

If Me.LV_Palm.ListItems(i).Checked = True Then
vQuery = "exec usp_IV_UpdatePrintUpdateChangePrice '" & vBarCode & "','" & vUnitCode & "','" & strUsername & "','" & Name1 & "' "
vConnection.Execute vQuery

vQuery = "exec dbo.USP_NP_UpdatePrintLabelHandHeld '" & vBarCode & "' "
vConnection.Execute vQuery

End If
Next i

Me.LV_Palm.ListItems.Clear
Me.CBSelectAll.Value = 0

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Private Sub cmdReset_Click()
        On Error GoTo ErrDescription
        
        txtUsername.Text = ""
        txtPassword = ""
        txtUsername.SetFocus
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDSearchItemChangePrice_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim ListChangePrice As ListItem
Dim vCount As Integer
Dim i As Integer
Dim vDocDate As Date

On Error GoTo ErrDescription

Me.MousePointer = 1
PrgBar101.Visible = True
PrgBar101.Value = 0
LV_ChangePrice1.ListItems.Clear
vDocDate = CDate(DTPicker3.Day & "/" & DTPicker3.Month & "/" & DTPicker3.Year)
i = 0
vQuery = "exec  dbo.USP_NP_SearchItemChangePrice '" & vDocDate & "' "
If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
If Not vRecordset.EOF Then
    PrgBar101.Max = vRecordset.RecordCount
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set ListChangePrice = LV_ChangePrice1.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
    ListChangePrice.SubItems(1) = Trim(vRecordset.Fields("barcode").Value)
    If IsNull(vRecordset.Fields("name1").Value) Then
        ListChangePrice.SubItems(2) = ""
    Else
        ListChangePrice.SubItems(2) = Trim(vRecordset.Fields("name1").Value)
    End If
    ListChangePrice.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
    ListChangePrice.SubItems(4) = Trim(vRecordset.Fields("SalePrice1").Value)
    ListChangePrice.SubItems(5) = 1
    ListChangePrice.SubItems(6) = Trim(vRecordset.Fields("dateupdate").Value)
    ListChangePrice.SubItems(7) = Trim(vRecordset.Fields("whcode").Value)
    ListChangePrice.SubItems(8) = Trim(vRecordset.Fields("priceerect").Value)
    ListChangePrice.SubItems(9) = Trim(vRecordset.Fields("shelfcode").Value)
    vRecordset.MoveNext
    i = i + 1
    PrgBar101.Value = i
    Wend
End If
MsgBox "ค้นหาข้อมูลเรียบร้อยแล้ว", vbInformation, "Send Information Message"
Else
MsgBox "ไม่มีข้อมูลที่ต้องการค้นหา", vbInformation, "Send Information Message"
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDShelf_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vShelf1 As String
Dim vShelf2 As String
Dim vListShelf As ListItem
Dim i As Integer
Dim vWHCode As String

On Error GoTo ErrDescription

If TXTShelf1.Text <> "" And TXTShelf2.Text <> "" And CMBWH.Text <> "" Then
    vShelf1 = Trim(TXTShelf1.Text)
    vShelf2 = Trim(TXTShelf2.Text)
    vWHCode = Trim(CMBWH.Text)
    ListViewShelf1.ListItems.Clear
    vQuery = "exec dbo.USP_IC_ShelfHMX '" & vShelf1 & "','" & vShelf2 & "','" & vWHCode & "' "
    If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        i = 1
        While Not vRecordset.EOF
        Set vListShelf = ListViewShelf1.ListItems.Add(, , i)
        vListShelf.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
        vListShelf.SubItems(2) = Trim(vRecordset.Fields("name").Value)
        vListShelf.SubItems(3) = Trim(vRecordset.Fields("whcode").Value)
        vListShelf.SubItems(4) = Trim(vRecordset.Fields("secstaff").Value)
        vRecordset.MoveNext
        i = i + 1
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

Private Sub CMDSop1_Click()
Dim tmpDate As String
        On Error GoTo Err1:
        
        ConnectSQL
        Dim ListX As ListItem
        
        If OptSop1.Value = True Then
                gbstrSQL1 = "select * from ( Select docno,itemcode,itemname,unitcode,price,qty,whcode From bcsaleordersub  where iscancel = 0 " _
                & " Union " _
                & " Select docno,itemcode,itemname,unitcode,price,qty ,whcode From bcquotationsub  where iscancel = 0 " _
                & " ) as Item_Reserve Where docno = '" & Trim(TXTSop1.Text) & "' "
                
        End If
        If OptSop2.Value = True Then
                gbstrSQL1 = "select * from ( Select docno,itemcode,itemname,unitcode,price,qty,whcode From bcsaleordersub  where iscancel = 0 " _
                & " Union " _
                & " Select docno,itemcode,itemname,unitcode,price,qty,whcode From bcquotationsub  where iscancel = 0 " _
                & " ) as Item_Reserve Where arcode = '" & Trim(TXTSop1.Text) & "'"
        End If

    
        Rs1.Open gbstrSQL1, ConnSQL, adOpenDynamic, adLockOptimistic
        If Not Rs1.EOF Then
                LV_Sop1.ListItems.Clear
                Rs1.MoveFirst
                While Not Rs1.EOF
                    Set ListX = LV_Sop1.ListItems.Add(, , Trim(Rs1!docno))
                    If IsNull(Rs1!Itemcode) = True Then
                            MsgBox "ไม่มีบาร์โค้ด & " & Rs1!Itemcode & " ", vbInformation, "ข้อความเตือน"
                            ListX.SubItems(1) = "No Barcode"
                    Else
                            ListX.SubItems(1) = Trim(Rs1!Itemcode)
                    End If
                    ListX.SubItems(2) = Trim(Rs1!ITEMname)
                    ListX.SubItems(3) = Trim(Rs1!Unitcode)
                    ListX.SubItems(4) = Trim(Rs1!price)
                    ListX.SubItems(5) = 1
                    ListX.SubItems(6) = Trim(Rs1!whcode)
                    Rs1.MoveNext
                Wend
                Rs1.Close
        Else        ' ไม่พบรายการที่ค้นหา
                MsgBox "ไม่พบใบสั่งขายเลขที่" & Trim(TXTSop1.Text), vbOKOnly + vbInformation, "คำเตือน"
                If Me.DTPicker1.Visible = False Then
                        TXTSop1.Text = ""
                        TXTSop1.SetFocus
                End If
                Exit Sub
        End If
        Exit Sub
        
' Error Found
Err1:
        MsgBox Err.Description, vbOKOnly + vbCritical, "พบข้อผิดพลาดของตัวโปรแกรม"
        Exit Sub
End Sub

Private Sub cmdSync_Click()
Dim vQuery As String
Dim Rs1 As New Recordset
Dim ListItem As ListItem
Dim i As Integer
Dim n As Integer
Dim vSalePrice As Double

'frmWait.Show            ' Loading Data

LV_Palm.ListItems.Clear


vQuery = "exec dbo.USP_NP_SearchItemFromHandHeld"
Rs1.Open vQuery, ConnSQL, adOpenDynamic, adLockOptimistic
If Not Rs1.EOF Then
    Rs1.MoveFirst
    i = 1
    While Not Rs1.EOF
    Set ListItem = LV_Palm.ListItems.Add(, , Trim(Rs1.Fields("itemcode").Value))
    ListItem.SubItems(1) = Trim(Rs1.Fields("barcode").Value)
    ListItem.SubItems(2) = Trim(Rs1.Fields("barcodename").Value)
    ListItem.SubItems(3) = Trim(Rs1.Fields("unitcode").Value)
    vSalePrice = Rs1.Fields("saleprice1").Value
    ListItem.SubItems(4) = Format(vSalePrice, "##,##0.00")
    ListItem.SubItems(5) = Trim(Rs1.Fields("qty").Value)
    ListItem.SubItems(6) = Trim(Rs1.Fields("reportname").Value)
    ListItem.SubItems(7) = Trim(Rs1.Fields("pathname").Value)
    Rs1.MoveNext
    i = i + 1
    Wend
End If
Rs1.Close
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub
Private Sub ProgressMoveEx(ByVal lTemp As Long, lMax As Long)     ', ByVal sText As String)
    Dim IncrX As Single
    Dim picRect As RECT

On Error GoTo ErrDescription

'    Dim allRect As RECT
    IncrX = lTemp * frmWait.Picture1.Width / lMax
    
    ' picture rectangle
'    allRect.Left = 0
'    allRect.Top = 0
'    allRect.Bottom = frmWait.Picture1.Height
    ' percent rectangle
    picRect.Left = 0
    picRect.Top = 0
    picRect.Bottom = frmWait.Picture1.Height
    picRect.Right = IncrX / Screen.TwipsPerPixelX
     
    frmWait.Picture1.Cls
    DrawGradient frmWait.Picture1.hdc, picRect, vbWhite, vbRed, False
    'frmWait.Picture1.Line (0, 0)-(IncrX, frmWait.Picture1.Height), vbRed, BF
    'frmWait.Picture1.Print "Joe"
    
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Private Sub CMDTRF1_Click()
Dim tmpDate As String
        On Error GoTo Err1:
        
        ConnectSQL
        Dim ListX As ListItem
        
        If OptSop1.Value = True Then
                'gbstrSQL1 = "select a.docno,a.docdate,a.itemcode,b.name1 as itemname,a.price,a.qty,UNITCODE,towh from BCStkTransfSub2 a inner join bcitem b on  a.itemcode = b.code Where a.docno = '" & Trim(TXTTRF1.Text) & "'"
                gbstrSQL1 = "select a.docno,a.docdate,a.itemcode,b.name1 as itemname,a.price,a.qty,a.UNITCODE,towh,c.barcode " _
                                        & " from BCStkTransfSub2 a inner join bcitem b on  a.itemcode = b.code " _
                                        & " left join ( " _
                                        & " select itemcode,barcode from bcbarcodemaster " _
                                        & " where  len(barcode) = 13) as c " _
                                        & " on a.itemcode = c.itemcode " _
                                        & " Where a.docno = '" & Trim(TXTTRF1.Text) & "'"
        End If
        If OptSop2.Value = True Then
                MsgBox "ยังไม่ได้ทำให้ครับ", vbInformation, "ข้อความแจ้ง"
                'gbstrSQL1 = "Select * From BCSTKTRANSFSUB Where arcode = '" & Trim(TXTTRF1.Text) & "'"
        End If

    
        Rs1.Open gbstrSQL1, ConnSQL, adOpenDynamic, adLockOptimistic
        If Not Rs1.EOF Then
                LV_TRF1.ListItems.Clear
                Rs1.MoveFirst
                While Not Rs1.EOF
                    Set ListX = LV_TRF1.ListItems.Add(, , Trim(Rs1!docno))
                    ListX.SubItems(1) = Trim(Rs1!Itemcode)
                    If IsNull(Rs1!barcode) = True Then
                            MsgBox "ไม่มีบาร์โค้ด 13 หลัก & " & Rs1!Itemcode & " ", vbInformation, "ข้อความเตือน"
                            ListX.SubItems(2) = Trim(Rs1!Itemcode)
                    Else
                            ListX.SubItems(2) = Trim(Rs1!barcode)
                    End If
                    ListX.SubItems(3) = Trim(Rs1!ITEMname)
                    ListX.SubItems(4) = Trim(Rs1!Unitcode)
                    ListX.SubItems(5) = Trim(Rs1!price)
                    ListX.SubItems(6) = Trim(Rs1!QTY)
                    ListX.SubItems(7) = Trim(Rs1!towh)
                    Rs1.MoveNext
                Wend
                Rs1.Close
        Else        ' ไม่พบรายการที่ค้นหา
                MsgBox "ไม่พบใบโอนสินค้าเลขที่" & Trim(TXTTRF1.Text), vbOKOnly + vbInformation, "คำเตือน"
                If Me.DTPicker1.Visible = False Then
                        TXTTRF1.Text = ""
                        TXTTRF1.SetFocus
                End If
                Exit Sub
        End If
        Exit Sub
        
' Error Found
Err1:
        MsgBox Err.Description, vbOKOnly + vbCritical, "พบข้อผิดพลาดของตัวโปรแกรม"
        Exit Sub
End Sub

Private Sub Command1_Click()
Dim ListX As ListItem
Dim vLenBarCode As Integer
Dim vAnswer As String
Dim vCheckDegitBar As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vRecordset2 As New ADODB.Recordset

Dim vItemCode As String
Dim vCheckTypePrint As Integer
Dim vRequestNo As String
        
On Error GoTo ErrDescription
                
    
vItemCode = Trim(Text1.Text)

vQuery = "exec dbo.USP_PM_ItemCheckPromotion '" & vItemCode & "' "
If OpenDatabase(vConnection, vRecordset2, vQuery) <> 0 Then
    vRequestNo = Trim(vRecordset2.Fields("requestno").Value)
    vCheckTypePrint = 1
    Else
    vCheckTypePrint = 0
    vRequestNo = ""
End If
vRecordset2.Close

'=============================ถ้าพิมพ์ป้ายราคาให้ตรวจสอบ ราคาโปรโมชั่น=================================================
'==============================================================================
'==============================================================================
'==============================================================================

'If vCheckTypePrint = 1 Then
 '   MsgBox "สินค้า รหัส " & vitemCode & " เป็นสินค้าโปรโมชั่น อยู่ในเอกสารเลขที่ " & vRequestNo & " ไม่สามารถพิมพ์ป้ายราคาเกี่ยวกับราคาปกติได้ ต้องลบรายการสินค้าดังกล่าวออกจากรายการพิมพ์", vbCritical, "Send Error Message"
  '  Exit Sub
'End If
        
        
'==============================================================================
'==============================================================================
'==============================================================================
'==============================================================================

        
        If Option1.Value = True Then
            vQuery = "exec dbo.USP_NP_LabelPriceList_ItemCode '" & Trim(Text1.Text) & "' "
        End If
        If Option2.Value = True Then
            vQuery = "exec dbo.USP_NP_LabelPriceList_BarCode '" & Trim(Text1.Text) & "' "
        End If
        If Option3.Value = True Then
            vQuery = "exec dbo.USP_NP_LabelPriceList_ItemName '" & Trim(Text1.Text) & "' "
        End If
        If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set ListX = frmProductDetail.LV_ProductDetail.ListItems.Add(, , Trim(vRecordset.Fields("Itemcode").Value))
            If IsNull(Trim(Trim(vRecordset.Fields("whcode").Value))) Then
            ListX.SubItems(1) = ""
            Else
            ListX.SubItems(1) = Trim(vRecordset.Fields("whcode").Value)
            End If
            ListX.SubItems(2) = Trim(vRecordset.Fields("barcode").Value)
            ListX.SubItems(3) = Trim(vRecordset.Fields("name1").Value)
            ListX.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
            ListX.SubItems(5) = Format(Trim(vRecordset.Fields("saleprice1").Value), "##,##0.00")
            If IsNull(Trim(vRecordset.Fields("priceerect").Value)) Then
            ListX.SubItems(6) = 0
            Else
            ListX.SubItems(6) = Trim(vRecordset.Fields("priceerect").Value)
            End If
            If IsNull(Trim(Trim(vRecordset.Fields("shelfcode").Value))) Then
            ListX.SubItems(7) = ""
            Else
            ListX.SubItems(7) = Trim(vRecordset.Fields("shelfcode").Value)
            End If
            vRecordset.MoveNext
    Wend
    vCheckDegitBar = Trim(Text1.Text)
    vLenBarCode = Len(vCheckDegitBar)
    If vLenBarCode = 13 Then
        vQuery = "select dbo.FT_CK_BarCodeEAN13 ('" & vCheckDegitBar & "') as DegitAnswer"
        If OpenDatabase(vConnection, vRecordset1, vQuery) <> 0 Then
            vAnswer = UCase(Trim(vRecordset1.Fields("DegitAnswer").Value))
        End If
        vRecordset1.Close
        If vAnswer = UCase(Trim("NO")) Then
            MsgBox "บาร์โค้ด " & vCheckDegitBar & " ไม่ถูกตามกฏ EAN 13 กรุณาแจ้ง บริหารสินค้าที่ดูแลสินค้าตัวนี้ เพราะ ถ้าพิมพ์ฟอร์ม 13 หลัก(5-1) จะทำให้บาร์โค้ดผิด ใช้งานไม่ได้ หรือไม่ ก็ต้องพิมพ์ฟอร์มที่มีจุดทศนิยม(5-3)แทน", vbCritical, "Send Error BarCode"
            frmProductDetail.Visible = False
            frmWizard.Enabled = True
            Exit Sub
        Else
            frmProductDetail.Visible = True
            frmWizard.Enabled = False
        End If
    End If
    
    Else
        MsgBox "ไม่มีข้อมูลที่ทำการค้นหา ตรวจสอบความถูกต้องด้วย", vbCritical, "Send Error"
        Exit Sub
    End If
    vRecordset.Close
    
    frmProductDetail.Visible = True
    frmWizard.Enabled = False
    Text2.Enabled = True
    
    If IsNull(tmpSPrice) Or tmpSPrice = "" Then
        tmpSPrice = 0
    End If
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Command2_Click()            ' Search By V_Label_PODetail (ใบสั่งซื้อ)
        On Error GoTo Err1:
        
        ConnectSQL
        Dim tmpDate As String
        Dim ListX As ListItem
        
        If Option6.Value = True Then            ' Search By เลขที่เอกสาร
                gbstrSQL1 = "Select * From V_Label_PODetail Where PONUM= '" & Trim(Text3.Text) & "'"
        End If
        If Option5.Value = True Then            ' Search By รหัสเจ้าหนี้
                gbstrSQL1 = "Select * From V_Label_PODetail Where VENDORID = '" & Trim(Text3.Text) & "'"
        End If
        If Option4.Value = True Then            ' Search By วันที่ออกใบสั่งซื้อ
                tmpDate = Day(DTPicker2.Value) & "/" & Month(DTPicker2.Value) & "/" & Year(DTPicker2.Value)
                'MsgBox tmpDate
                gbstrSQL1 = "Select * From V_Label_PODetail Where DOCDATE = '" & tmpDate & "'"
        End If
        
        ' Connect and Add To LV_PO_1
        ConnectSQL
        If Rs1.State = adStateOpen Then Rs1.Close
        Rs1.Open gbstrSQL1, ConnSQL, adOpenDynamic, adLockOptimistic
        If Not Rs1.EOF Then
                Me.LV_PO_1.ListItems.Clear
                Rs1.MoveFirst
                While Not Rs1.EOF
                    Set ListX = LV_PO_1.ListItems.Add(, , Trim(Rs1!Itemcode))
                    If IsNull(Rs1!barcode) = True Then
                            ListX.SubItems(1) = "No Barcode"
                    Else
                            ListX.SubItems(1) = Trim(Rs1!barcode)
                    End If
                    ListX.SubItems(2) = Trim(Rs1!ITEMname)
                    ListX.SubItems(3) = Trim(Rs1!Unitcode)
                    ListX.SubItems(4) = Trim(Round(Rs1!price))
                    ListX.SubItems(5) = Trim(Rs1!QTY)
                    
                    ' Move Next Record
                    Rs1.MoveNext
                Wend
        Else        ' ไม่พบรายการที่ค้นหา
                MsgBox "ไม่พบรายการสินค้าตาม" & Trim(lbPO.Caption) & Trim(Text3.Text), vbOKOnly + vbInformation, "คำเตือน"
                If Me.DTPicker2.Visible = False Then
                        Text3.Text = ""
                        Text3.SetFocus
                End If
                Exit Sub
        End If
        
        ' Close Connection
        Rs1.Close
        ConnSQL.Close
        Exit Sub
        
' Found Error
Err1:
        MsgBox Err.Description, vbOKOnly + vbCritical, "พบข้อผิดพลาดของตัวโปรแกรม"
        Exit Sub
End Sub

Private Sub Command3_Click()            ' Search ตามเงื่อนไขของใบรับสินค้า
        Dim tmpDate As String
        ' Error Handling
        On Error GoTo Err1:
        
        ConnectSQL
        Dim ListX As ListItem
        
        If Option7.Value = True Then
                gbstrSQL1 = "exec dbo.USP_LB_RecievingDetail  1,'" & Trim(Text4.Text) & "' "
        End If
        If Option8.Value = True Then
                gbstrSQL1 = "exec dbo.USP_LB_RecievingDetail  2,'" & Trim(Text4.Text) & "' "
        End If
        If Option9.Value = True Then
                tmpDate = Day(DTPicker1.Value) & "/" & Month(DTPicker1.Value) & "/" & Year(DTPicker1.Value)
                'MsgBox tmpDate       'ตรวจสอบเวลา
                'MsgBox Me.DTPicker1.Value
                gbstrSQL1 = "exec dbo.USP_LB_RecievingDetail  3,'" & Trim(Text4.Text) & "'  "
        End If
        
        ' Connect and Add To LV_PO1_1
        Rs1.Open gbstrSQL1, ConnSQL, adOpenDynamic, adLockOptimistic
        If Not Rs1.EOF Then
                LV_PO1_1.ListItems.Clear
                Rs1.MoveFirst
                While Not Rs1.EOF
                    Set ListX = LV_PO1_1.ListItems.Add(, , Trim(Rs1!Itemcode))
                    ListX.SubItems(1) = Trim(Rs1!whcode)
                    If IsNull(Rs1!barcode) = True Then
                            MsgBox "ไม่มีบาร์โค้ด & " & Rs1!Itemcode & " ", vbInformation, "ข้อความเตือน"
                            ListX.SubItems(2) = "No Barcode"
                    Else
                            ListX.SubItems(2) = Trim(Rs1!barcode)
                    End If
                    ListX.SubItems(3) = Trim(Rs1!ITEMname)
                    ListX.SubItems(4) = Trim(Rs1!Unitcode)
                    If IsNull(Rs1!saleprice1) = True Then
                            MsgBox "ไม่มีระดับราคาสินค้า รหัส & " & Rs1!Itemcode & " ", vbInformation, "ข้อความเตือน"
                            ListX.SubItems(5) = "No PriceList"
                    Else
                            ListX.SubItems(5) = Trim(Rs1!saleprice1)
                    End If
                    
                    ListX.SubItems(6) = Rs1!CNQTY
                    ListX.SubItems(7) = Trim(Rs1.Fields("Docno").Value)
                    ListX.SubItems(8) = Rs1.Fields("docdate").Value
                    
                    ' Move Next Record
                    Rs1.MoveNext
                Wend

                Rs1.Close
                LV_PO1_1.SetFocus
        Else        ' ไม่พบรายการที่ค้นหา
                MsgBox "ไม่พบรายการสินค้าตาม" & Trim(lbPO2.Caption) & Trim(Text4.Text), vbOKOnly + vbInformation, "คำเตือน"
                If Me.DTPicker1.Visible = False Then
                        Text4.Text = ""
                        Text4.SetFocus
                End If
                Exit Sub
        End If
        Exit Sub
        
' Error Found
Err1:
        MsgBox Err.Description, vbOKOnly + vbCritical, "พบข้อผิดพลาดของตัวโปรแกรม"
        Exit Sub
End Sub

Private Sub Command4_Click()
        Call ChkPWD
End Sub

Private Sub Command5_Click()

End Sub

Public Sub GeItemChangePrice()
Dim vQuery As String
Dim ListChangePrice As ListItem
Dim vCount As Integer
Dim i As Integer
Dim vDocDate As Date

On Error GoTo ErrDescription

PrgBar101.Visible = True
LV_ChangePrice1.ListItems.Clear
vDocDate = CDate(DTPicker3.Day & "/" & DTPicker3.Month & "/" & DTPicker3.Year)
i = 0
vQuery = "exec  dbo.USP_NP_SearchItemChangePrice '" & vDocDate & "' "
Rs1.Open vQuery, ConnSQL, adOpenDynamic, adLockOptimistic
If Not Rs1.EOF Then
    PrgBar101.Max = Rs1.RecordCount
    Rs1.MoveFirst
    While Not Rs1.EOF
    Set ListChangePrice = LV_ChangePrice1.ListItems.Add(, , Trim(Rs1.Fields("itemcode").Value))
    
    ListChangePrice.SubItems(1) = Trim(Rs1.Fields("barcode").Value)
    If IsNull(Rs1.Fields("name1").Value) Then
        ListChangePrice.SubItems(2) = ""
    Else
        ListChangePrice.SubItems(2) = Trim(Rs1.Fields("name1").Value)
    End If
    ListChangePrice.SubItems(3) = Trim(Rs1.Fields("unitcode").Value)
    ListChangePrice.SubItems(4) = Trim(Rs1.Fields("SalePrice1").Value)
    ListChangePrice.SubItems(5) = 1
    ListChangePrice.SubItems(6) = Trim(Rs1.Fields("dateupdate").Value)
    ListChangePrice.SubItems(7) = Trim(Rs1.Fields("whcode").Value)
    ListChangePrice.SubItems(8) = Trim(Rs1.Fields("priceerect").Value)
    ListChangePrice.SubItems(9) = Trim(Rs1.Fields("shelfcode").Value)
    Rs1.MoveNext
    i = i + 1
    PrgBar101.Value = i
    Wend
End If
Rs1.Close
PrgBar101.Visible = False


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If

End Sub
Private Sub Form_Load()
        Dim strTip, strSQL, strSQL2 As String
        Dim iCount As Integer
        
        On Error GoTo ErrDescription
             
        ' Set เริ่มต้นว่าไม่ได้มาจาก Recieving Form or Puchase Order Form
        CMDBartendor.Visible = False
        AddTempRCV = False
        AddTempPO = False
        DTPicker1 = Now
        DTPicker3 = Now
        
        ' Start Step
        IntStep = 0
        PicWizard(IntStep).Visible = True
                
        Dim ListX As ListItem
        Set ListX = ListView1.ListItems.Add(, , "Manual", 5)
        Set ListX = ListView1.ListItems.Add(, , "Palm Application", 5)
        Set ListX = ListView1.ListItems.Add(, , "จากใบสั่งซื้อ", 5)
        Set ListX = ListView1.ListItems.Add(, , "จากใบรับสินค้า", 5)
        Set ListX = ListView1.ListItems.Add(, , "โปรโมชั่นหมดอายุ", 5)
        Set ListX = ListView1.ListItems.Add(, , "ทะเบียนสินทรัพย์", 5)
        Set ListX = ListView1.ListItems.Add(, , "จากใบสั่งขาย", 5)
        Set ListX = ListView1.ListItems.Add(, , "จากใบขอโอนสินค้า", 5)
        Set ListX = ListView1.ListItems.Add(, , "ปรับราคาสินค้า", 5)
        Set ListX = ListView1.ListItems.Add(, , "พิมพ์ชั้นเก็บ", 5)
        
        ' Start Check Box Option
        Option1.Value = True        ' หน้าเลือกเอง
        Option6.Value = True        ' หน้าใบสั่งซื้อ
        Option7.Value = True        ' หน้าใบรับสินค้า
        
        ' Connect to Tool Tips
        ConnTipDB
        strTip = "Select * From Tips"
        Rs1.Open strTip, ConnAccess, 1, 3
        If Not Rs1.EOF Then
                ' กำหนดค่า Array
                ReDim strHeader(Rs1.RecordCount)
                ReDim strDetail(Rs1.RecordCount)
                
                Rs1.MoveFirst
                For iCount = 0 To Rs1.RecordCount - 1
                        strHeader(iCount) = Trim(Rs1!Header)
                        strDetail(iCount) = Trim(Rs1!Detail)
                        Rs1.MoveNext
                Next
        End If
        Rs1.Close       ' Close Connection
        
        ' Get Detail
        Call GetDetail
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    ConnectSQL
    ' Declaration
    Dim Answer As VbMsgBoxResult
    Dim strSQL As String
    Dim tmpPathName As String
  
    Dim BartendorForm As String
        
    tmpPathName = Trim(LV_Report.SelectedItem.SubItems(1)) & ".rpt"
    BartendorForm = Mid(tmpPathName, 21, 4)
    Answer = MsgBox("คุณต้องการออกจากโปรแกรม Label Wizard?", vbQuestion + vbYesNo + vbDefaultButton1, "Label Wizard Warning")
    If Answer = vbNo Then Cancel = True: Exit Sub
    ' Before End Delete Temp
    strSQL = "Delete From NP_Label_Temp Where Useduser = '" & strUsername & "'"
    ConnSQL.Execute strSQL
    'ConnSQL.Close
    If BartendorForm = "R2-8" Then
    strSQL = "Drop Table Report_Temp1"
    ConnSQL.Execute strSQL
    'ConnSQL.Close
    ElseIf BartendorForm = "R5-1" Or BartendorForm = "R5-2" Or BartendorForm = "R5-3" Or BartendorForm = "R5-4" Or BartendorForm = "R5-5" Or BartendorForm = "R5-7" Then
    strSQL = "Drop Table Report_Temp2"
    ConnSQL.Execute strSQL
    ElseIf BartendorForm = "R8-1" Then
    strSQL = "Drop Table Report_Temp3"
    ConnSQL.Execute strSQL
    End If
        Unload Me
        If ConnSQL.State = 1 Then
            ConnSQL.Close
        End If
        If vConnection.State = 1 Then
            vConnection.Close
        End If
    End 'Just incase it's stuck in the importing function
End Sub


Public Sub GetNamePromotionExpire()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListPromo As ListItem

On Error GoTo ErrDescription

vQuery = " execute  dbo.USP_PM_PromotionMasterExpire "
If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMBPromotionCode.AddItem Trim(vRecordset.Fields("promoname").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub


Private Sub cmdNext_Click()
        'On Error GoTo ErrDescription/////////////////////////////////////////////////////////////////////////////////////////////
 On Error Resume Next
        ' Go to Next Wizard
         
        
        Select Case IntStep
                Case 0:             ' Now Login Next To PicWizard(1)
                        cmdPrevious.Visible = True
                        cmdNext.Visible = True
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizard(IntStep + 1).Visible = True
                        IntStep = IntStep + 1
                        strUsername = Trim(txtUsername)
                        ListView1.SetFocus
                Case 1:             ' Select Source Wizard
                        JobSelect = ListView1.SelectedItem.Index            ' Ex. Index เลือกเอง = 1 แต่เป็น intStep 2 เพื่อเก็บตำแหน่งย้อนกลับ
                        'MsgBox ListView1.SelectedItem.Index
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizard(Int(JobSelect) + 1).Visible = True
                        'PicWizard(2).Visible = True
                        IntStep = Int(JobSelect) + 1
                        Select Case IntStep
                                Case 2:     ' เลือกเอง
                                        Text1.SetFocus      ' Setfocus Text1
                                        LV_Label.ListItems.Clear
                                        Option1.Value = True
                                        CMD101.Visible = True
                                Case 3:     ' พิมพ์ป้ายจากเครื่องปาล์ม
                                        LV_Palm.ListItems.Clear
                                        cmdSync.SetFocus
                                Case 4:     ' PO Order
                                        Option6.Value = True
                                        Text3.Visible = True
                                        DTPicker2.Visible = False
                                        Text3.Text = ""
                                        Text3.SetFocus
                                        LV_PO_1.ListItems.Clear
                                        LV_PO_2.ListItems.Clear
                                Case 5:     ' RCV Order
                                        vCheckReceipt = True
                                        Option7.Value = True
                                        Text4.Visible = True
                                        DTPicker1.Visible = False
                                        Text4.Text = ""
                                        Text4.SetFocus
                                        AddTempRCV = True
                                        LV_PO1_1.ListItems.Clear
                                        LV_PO1_2.ListItems.Clear
                                Case 6:     ' Palm ChkValue
                                        ListChkValue.ListItems.Clear
                                        Call GetNamePromotionExpire
                                        btnprocess.SetFocus
                                Case 7:  'Assets
                                        LV_Asset1.ListItems.Clear
                                        LV_Asset2.ListItems.Clear
                                        TXTAsset1.SetFocus
                                        vAsset = True
                                Case 8:  'SaleOrder
                                        LV_Sop1.ListItems.Clear
                                        LV_Sop2.ListItems.Clear
                                        TXTSop1.SetFocus
                                Case 9:  'Transfer
                                        LV_TRF1.ListItems.Clear
                                        LV_TRF2.ListItems.Clear
                                        TXTTRF1.SetFocus
                                Case 10:  'Change Price Of Item
                                        LV_ChangePrice1.ListItems.Clear
                                        LV_ChangePrice2.ListItems.Clear
                                        Me.DTPicker3.Value = Now
                                        CMBSection.SetFocus
                                        vChangePrice = True
                                        Call GeItemChangePrice
                                        'Call GetWHCode
                                Case 11: 'Promotion
                                        ListViewShelf1.ListItems.Clear
                                        ListViewShelf2.ListItems.Clear
                                        vShelfJob = True
                                        TXTShelf1.SetFocus
                                        Call AddWHCode
                                        'Call GetNamePromotion
                                        'Call GetSectionName
                        End Select
                Case 2:         ' เลือกเอง
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizardSelect.Visible = True
                        ' PicWizard(2).Visible = True
                        Call addListResult(IntStep)
                        CMD101.Visible = False
                        IntStep1 = IntStep
                        IntStep = 12
                Case 3:         ' Palm Application (พิมพ์ป้ายราคา)
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizardSelect.Visible = True
                        Call addListResult(IntStep)
                        IntStep1 = IntStep
                        IntStep = 12
                Case 4:         ' ใบสั่งซื้อ
                        AddTempPO = True            ' เป็นการบอกว่าได้รับค่ามาจากใบสั่งซื้อ เพื่อนำค่าไปเพิ่มใน NP_Label_Temp
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizardSelect.Visible = True
                        Call addListResult(IntStep)
                        IntStep1 = IntStep
                        IntStep = 12
                Case 5:         ' ใบรับสินค้า
                        'AddTempRCV = True            ' เป็นการบอกว่าได้รับค่ามาจากใบรับสินค้า เพื่อนำค่าไปเพิ่มใน NP_Label_Temp
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizardSelect.Visible = True
                        Call addListResult(IntStep)
                        IntStep1 = IntStep
                        IntStep = 12
                Case 6:         ' Palm Application (เปรียบเทียบราคาสินค้า)
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizardSelect.Visible = True
                        Call addListResult(IntStep)
                        IntStep1 = IntStep
                        IntStep = 12
                Case 7:         ' ทรัพย์สิน
                        'AddTempRCV = True            ' เป็นการบอกว่าได้รับค่ามาจากใบรับสินค้า เพื่อนำค่าไปเพิ่มใน NP_Label_Temp
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizardSelect.Visible = True
                        Call addListResult(IntStep)
                        IntStep1 = IntStep
                        IntStep = 12
                Case 8:         ' ใบสั่งขาย
                        AddTempRCV = True            ' เป็นการบอกว่าได้รับค่ามาจากใบรับสินค้า เพื่อนำค่าไปเพิ่มใน NP_Label_Temp
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizardSelect.Visible = True
                        Call addListResult(IntStep)
                        IntStep1 = IntStep
                        IntStep = 12
                Case 9:         ' ใบขอโอนย้ายสินค้าระหว่างคลัง
                        'AddTempRCV = True            ' เป็นการบอกว่าได้รับค่ามาจากใบรับสินค้า เพื่อนำค่าไปเพิ่มใน NP_Label_Temp
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizardSelect.Visible = True
                        Call addListResult(IntStep)
                        IntStep1 = IntStep
                        IntStep = 12
                Case 10:         ' รหัสสินค้าที่ทำการเปลี่ยนแปลงราคาใหม่
                        'AddTempRCV = True            ' เป็นการบอกว่าได้รับค่ามาจากใบรับสินค้า เพื่อนำค่าไปเพิ่มใน NP_Label_Temp
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizardSelect.Visible = True
                        Call addListResult(IntStep)
                        IntStep1 = IntStep
                        IntStep = 12
                Case 11:         ' รหัสสินค้าที่ทำการเปลี่ยนแปลงราคาใหม่
                        'AddTempRCV = True            ' เป็นการบอกว่าได้รับค่ามาจากใบรับสินค้า เพื่อนำค่าไปเพิ่มใน NP_Label_Temp
                        PicWizard(IntStep).Visible = False          ' Invisible Last Step
                        PicWizardSelect.Visible = True
                        Call addListResult(IntStep)
                        IntStep1 = IntStep
                        IntStep = 12
                Case 12:         ' Add Data From ListResult To NP_Label_Temp And Show Form Type Prompt Print
                        ' Check ว่ามาจากใบรับสินค้า หรือเปล่า
                        'If AddTempRCV = True Then
                         '       Call AddTempByRecieving                 ' Add Temp By Recieving
                                If AddTempPO = True Then
                                        Call AddTempByPO                        ' Add Temp By Purchase Order
                                'ElseIf AddTempRCV = True Then
                                        'Call AddTempByRecieving
                                ElseIf vChangePrice = True Then
                                        Call AddTempByChangePrice
                                ElseIf vAsset = True Then
                                        Call AddTempAsset
                                ElseIf vShelfJob = True Then
                                        Call AddTempPromotion
                                Else
                                        Call AddToTemp                              ' Other Job Add temp
                                        Me.Check104.Value = 1
                                        CMDBartendor.Visible = True
                                End If
                        
                        If DataHas = True Then
                                PicWizardSelect.Visible = False             ' Invisible Last Step
                                PicWizardReport.Visible = True
                                Me.LV_Report.SetFocus
                                cmdNext.Enabled = False
                                'cmdFinish.Visible = True
                                cmdPreview.Visible = True
                                CMDBartendor.Visible = True
                                IntStep = IntStep + 1
                        Else
                                MsgBox "คุณไม่ได้เลือกรายการสินค้าที่ต้องการพิมพ์", vbOKOnly + vbInformation, "คำแนะนำ"
                                ' PicWizardSelect.Visible = True
                        End If
        End Select
                
        ' Get Detail
        Call GetDetail
        
'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description
  '  Exit Sub
'End If
End Sub

Private Sub AddTempByPO()
        Dim iCount As Integer
        Dim strSQL As String
        ' Dim ListX As ListItem
        
        ConnectSQL
        On Error GoTo ErrDescription
        
        ' Check Data in Listresult
        If ListResult.ListItems.count = 0 Then
                MsgBox "ไม่พบรายการสินค้าที่ต้องการพิมพ์", vbOKOnly + vbInformation, "คำแนะนำ"
                DataHas = False
                Exit Sub
        End If
        
        For iCount = 1 To ListResult.ListItems.count
            If Me.ListResult.ListItems(iCount).Checked = True Then
                DataHas = True
                tmpNUMBR = Trim(Me.ListResult.ListItems(iCount).Text)
                tmpBarcode = Me.ListResult.ListItems(iCount).SubItems(1)
                tmpTHINAME = Me.ListResult.ListItems(iCount).SubItems(2)
                tmpUOM = Me.ListResult.ListItems(iCount).SubItems(3)
                tmpRRC = Me.ListResult.ListItems(iCount).SubItems(4)
                tmpQTY = Me.ListResult.ListItems(iCount).SubItems(5)
                
                ' ยังไม่มีค่าใส่
                tmpUsedUser = strUsername
                tmpCategory_ID = ""
                tmpSite = ""
                tmpBIN_ID = ""
                tmpVENDR_ID = ""
                tmpRemark = ""
                tmpSPrice = 0
                tmpONHAND = ""
                tmpQTYALLOCATE = ""
                tmpTYPE = 0
                
                ' Add To NP_Label_Temp
                strSQL = "Insert into NP_LABEL_TEMP(ItemCode, barcode, NAME1, QTY, Price, UnitCode,UsedUser,WHCode,ShelfCode,SPrice,Type) " _
                & "values('" & tmpNUMBR & "','" & tmpBarcode & "','" & tmpTHINAME & "'," & Int(tmpQTY) & "," & tmpRRC & ",'" & tmpUOM & "','" & tmpUsedUser & "','" & tmpSite & "','" & tmpBIN_ID & "'," & Int(tmpSPrice) & "," & tmpTYPE & ")"
                ConnSQL.Execute strSQL
            End If
        Next iCount
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Private Sub AddTempByRecieving()
        Dim iCount As Integer
        Dim strSQL As String
        ' Dim ListX As ListItem
        
        ConnectSQL
        On Error GoTo ErrDescription
        
        ' Check Data in Listresult
        If ListResult.ListItems.count = 0 Then
                MsgBox "ไม่พบรายการสินค้าที่ต้องการพิมพ์", vbOKOnly + vbInformation, "คำแนะนำ"
                DataHas = False
                Exit Sub
        End If
        
        For iCount = 1 To ListResult.ListItems.count
            If Me.ListResult.ListItems(iCount).Checked = True Then
               DataHas = True
                'tmpNUMBR = LV_Sop2.ListItems(iCount).Text
                tmpNUMBR = Trim(Me.ListResult.ListItems(iCount).Text)
                tmpBarcode = Trim(Me.ListResult.ListItems(iCount).Text) 'Me.ListResult.ListItems(iCount).SubItems(1)
                tmpTHINAME = Me.ListResult.ListItems(iCount).SubItems(3)
                tmpUOM = Me.ListResult.ListItems(iCount).SubItems(4)
                tmpRRC = Me.ListResult.ListItems(iCount).SubItems(5)
                tmpQTY = Me.ListResult.ListItems(iCount).SubItems(6)
                
                strSQL = "select * from vw_IV_ItemPromotion where itemcode = '" & tmpBarcode & "' and unitcode = '" & tmpUOM & "' "
                Rs1.Open strSQL, ConnSQL, adOpenDynamic, adLockOptimistic
                If Not Rs1.EOF Then
                        tmpRRCLV = Trim(Rs1!promoprice)  'Trim(Rs1!PRCLEVEL) ดึงข้อมูลมาอีก View หนึ่ง
                        tmpRRC = Trim(Rs1!saleprice1) 'Trim(Rs1!UOMPRICE)
                        tmpSPrice = Trim(Rs1!priceerect)
                End If
                Rs1.Close
                
                ' ยังไม่มีค่าใส่
                tmpUsedUser = strUsername
                tmpCategory_ID = ""
                tmpSite = Me.ListResult.ListItems(iCount).SubItems(1)
                tmpBIN_ID = ""
                tmpVENDR_ID = ""
                tmpRemark = ""
                'tmpSPrice = 0
                tmpONHAND = ""
                tmpQTYALLOCATE = ""
                tmpTYPE = 0
                'tmpRRCLV = 0
                ' Add To NP_Label_Temp
                strSQL = "Insert into NP_LABEL_TEMP(Itemcode, barcode, name1, NAME2, QTY, PriceLevel, Price, Unitcode,UsedUser,Category_ID,WHCode,ShelfCode,VENDR_ID,remark,SPrice,ONHAND,QTYALLOCATE,Type)  " _
                & "values('" & tmpNUMBR & "','" & tmpBarcode & "','" & tmpTHINAME & "','" & tmpENGNAME & "'," & Int(tmpQTY) & "," & tmpRRCLV & "," & tmpRRC & ",'" & tmpUOM & "','" & tmpUsedUser & "','" & tmpCategory_ID & "','" & tmpSite & "','" & tmpBIN_ID & "','" & tmpVENDR_ID & "','" & tmpRemark & "'," & Int(tmpSPrice) & ",'" & tmpONHAND & "','" & tmpQTYALLOCATE & "'," & tmpTYPE & ")"
                ConnSQL.Execute strSQL
                
            End If
        Next iCount
 '----------------------------------------------------------------------------------------------------
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub
Private Sub GetDetail()
        Select Case IntStep
        Case 0:
                Me.Caption = "Label Wizard - Login"
        Case 1:
                Me.Caption = "Label Wizard - Wizard I"
        Case 2, 3, 4, 5, 7, 8, 9, 10, 11:
                Me.Caption = "Label Wizard - Wizard II"
        Case 6:
                Me.Caption = "Label Wizard - Wizard III"
        Case 12:
                Me.Caption = "Label Wizard - Wizard IV"
        End Select
        
        LabelHeader.Caption = strHeader(IntStep)
        Me.LabelDetail = strDetail(IntStep)
End Sub
Private Sub BeginConnect()
        ' Set focus User Name
        'Dim ListForm As ListItem
        Dim strTip, strSQL, strSQL2 As String
        Dim iCount As Integer
        
        ' Connect to SQL and Add Report To LV_Report
        ConnectSQL
        On Error GoTo ErrDescription
        
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
       strSQL = "Select  ReportName, PathName From NP_Label_Setup Where StdIndex = '1'  and typelabel in (1,2) order by parentid,reportname" ' คลัง
       'strSQL = "Select  ReportName, PathName From NP_Label_Setup Where StdIndex = '1'  and typelabel =0 order by parentid,reportname" 'การตลาด
       
       
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
        Rs1.Open strSQL, ConnSQL, adOpenDynamic, adLockOptimistic
        If Not Rs1.EOF Then
                Rs1.MoveFirst
                LV_Report.ListItems.Clear
                While Not Rs1.EOF
                        Set ListForm = Me.LV_Report.ListItems.Add(, , Trim(Rs1!ReportName))
                        ListForm.SubItems(1) = Trim(Rs1!PathName)
                        Rs1.MoveNext
                Wend
        End If
        Rs1.Close       ' Close Connection

        ' Get Detail
        Call GetDetail
        
        ' Set Data in ListViewResult
        DataHas = False
        
        Call cmdNext_Click
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub


Private Sub ListView1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                Call cmdNext_Click
        End If
End Sub

Private Sub ListViewShelf2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim j As Integer

On Error GoTo ErrDescription

If KeyCode = 46 Then
    If ListViewShelf2.ListItems.count <> 0 Then
        i = ListViewShelf2.SelectedItem.Index
        ListViewShelf2.ListItems.Remove (i)
          For j = 1 To ListViewShelf2.ListItems.count
          ListViewShelf2.ListItems.Item(j).Text = j
          Next j
    End If
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub LV_Asset1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
 Dim ListX As ListItem
        Dim iCount As Integer
        
        On Error GoTo ErrDescription
        
        If Item.Checked = True Then
                ' Check Box Option
                'MsgBox Item.Index
                Set ListX = LV_Asset2.ListItems.Add(, , LV_Asset1.ListItems(Item.Index).Text)
                ListX.SubItems(1) = Me.LV_Asset1.ListItems(Item.Index).SubItems(1)
                ListX.SubItems(2) = Me.LV_Asset1.ListItems(Item.Index).SubItems(2)
                ListX.SubItems(3) = Me.LV_Asset1.ListItems(Item.Index).SubItems(3)
                ListX.SubItems(4) = Me.LV_Asset1.ListItems(Item.Index).SubItems(4)
                ListX.SubItems(5) = Me.LV_Asset1.ListItems(Item.Index).SubItems(5)
        Else
                ' Uncheck Box Option
                For iCount = Me.LV_Asset2.ListItems.count To 1 Step -1
                        If LV_Asset1.ListItems(Item.Index).SubItems(1) = LV_Asset2.ListItems.Item(iCount).SubItems(1) Then
                                LV_Asset1.ListItems.Remove (iCount)
                        End If
                Next iCount
        End If
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub LV_ChangePrice1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim ListX As ListItem
        Dim iCount As Integer
        
        On Error GoTo ErrDescription
        
        If Item.Checked = True Then
                ' Check Box Option
                'MsgBox Item.Index
                Set ListX = LV_ChangePrice2.ListItems.Add(, , LV_ChangePrice1.ListItems(Item.Index).Text)
                ListX.SubItems(1) = Me.LV_ChangePrice1.ListItems(Item.Index).SubItems(1)
                ListX.SubItems(2) = Me.LV_ChangePrice1.ListItems(Item.Index).SubItems(2)
                ListX.SubItems(3) = Me.LV_ChangePrice1.ListItems(Item.Index).SubItems(3)
                ListX.SubItems(4) = Me.LV_ChangePrice1.ListItems(Item.Index).SubItems(4)
                ListX.SubItems(5) = Me.LV_ChangePrice1.ListItems(Item.Index).SubItems(5)
                ListX.SubItems(6) = Me.LV_ChangePrice1.ListItems(Item.Index).SubItems(6)
                ListX.SubItems(7) = Me.LV_ChangePrice1.ListItems(Item.Index).SubItems(7)
                ListX.SubItems(8) = Me.LV_ChangePrice1.ListItems(Item.Index).SubItems(8)
                ListX.SubItems(9) = Me.LV_ChangePrice1.ListItems(Item.Index).SubItems(9)
        Else
                ' Uncheck Box Option
                For iCount = Me.LV_ChangePrice2.ListItems.count To 1 Step -1
                        If LV_ChangePrice1.ListItems(Item.Index).SubItems(1) = LV_ChangePrice2.ListItems.Item(iCount).SubItems(1) Then
                                LV_ChangePrice1.ListItems.Remove (iCount)
                        End If
                Next iCount
        End If
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub


Private Sub LV_ChangePrice2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer

If KeyCode = 46 And LV_ChangePrice2.ListItems.count > 0 Then
  vIndex = LV_ChangePrice2.SelectedItem.Index
  LV_ChangePrice2.ListItems.Remove (vIndex)
End If
End Sub

Private Sub LV_Label_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer

If KeyCode = 46 And LV_Label.ListItems.count > 0 Then
  vIndex = LV_Label.SelectedItem.Index
  LV_Label.ListItems.Remove (vIndex)
End If
End Sub

Private Sub LV_PO_1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
        Dim ListX As ListItem
        Dim iCount As Integer
        
        On Error GoTo ErrDescription
        
        If Item.Checked = True Then
                ' Check Box Option
                Set ListX = LV_PO_2.ListItems.Add(, , LV_PO_1.ListItems(Item.Index).Text)
                ListX.SubItems(1) = Me.LV_PO_1.ListItems(Item.Index).SubItems(1)
                ListX.SubItems(2) = Me.LV_PO_1.ListItems(Item.Index).SubItems(2)
                ListX.SubItems(3) = Me.LV_PO_1.ListItems(Item.Index).SubItems(3)
                ListX.SubItems(4) = Me.LV_PO_1.ListItems(Item.Index).SubItems(4)
                ListX.SubItems(5) = Me.LV_PO_1.ListItems(Item.Index).SubItems(5)
        Else
                ' Uncheck Box Option
                For iCount = Me.LV_PO_2.ListItems.count To 1 Step -1
                        If LV_PO_1.ListItems(Item.Index).SubItems(1) = LV_PO_2.ListItems.Item(iCount).SubItems(1) Then
                                LV_PO_2.ListItems.Remove (iCount)
                        End If
                Next iCount
        End If
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub LV_PO1_1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
        Dim ListX As ListItem
        Dim iCount As Integer
        Dim vQuantity As String
        Dim vCountPrint As Integer
        
        On Error GoTo ErrDescription
        
        If Item.Checked = True Then
                ' Check Box Option
                'MsgBox Item.Index
                Set ListX = LV_PO1_2.ListItems.Add(, , LV_PO1_1.ListItems(Item.Index).Text)
                ListX.SubItems(1) = Me.LV_PO1_1.ListItems(Item.Index).SubItems(1)
                ListX.SubItems(2) = Me.LV_PO1_1.ListItems(Item.Index).SubItems(2)
                ListX.SubItems(3) = Me.LV_PO1_1.ListItems(Item.Index).SubItems(3)
                ListX.SubItems(4) = Me.LV_PO1_1.ListItems(Item.Index).SubItems(4)
                ListX.SubItems(5) = Me.LV_PO1_1.ListItems(Item.Index).SubItems(5)
                vQuantity = InputBox("จำนวนที่จะพิมพ์", "กรอกจำนวนพิมพ์ป้าย", Me.LV_PO1_1.ListItems(Item.Index).SubItems(6))
                If vQuantity = "" Then
                vCountPrint = 0
                Else
                vCountPrint = CCur(vQuantity)
                End If
                ListX.SubItems(6) = vCountPrint
                ListX.SubItems(7) = Me.LV_PO1_1.ListItems(Item.Index).SubItems(7)
                ListX.SubItems(8) = Me.LV_PO1_1.ListItems(Item.Index).SubItems(8)
                
        Else
                ' Uncheck Box Option
                Dim i As Integer
                Dim vIndex As Integer
                
                vIndex = LV_PO1_2.ListItems.count
                ReDim vRemove(vIndex) As Integer

                If LV_PO1_2.ListItems.count <> 1 Then
                    For i = 1 To LV_PO1_2.ListItems.count
                        Call CheckRemove(LV_PO1_2.ListItems.Item(i).Text)
                        If vCheckRemove = 2 Then
                            vRemove(i) = 1
                        Else
                            vRemove(i) = 2
                        End If
                    Next i
                    
                    For i = 1 To LV_PO1_2.ListItems.count
                        If vRemove(i) = 1 Then
                         LV_PO1_2.ListItems.Remove (i)
                        End If
                    Next i
                Else
                    LV_PO1_2.ListItems.Remove (1)
                End If
        End If
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub
Public Function CheckRemove(Item As String) As Integer
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To LV_PO1_1.ListItems.count
        If LV_PO1_1.ListItems.Item(i).Checked = True Then
                If LV_PO1_1.ListItems.Item(i).Text = Item Then
                    vCheckRemove = 1
                    Exit Function
                Else
                    vCheckRemove = 2
                End If
        End If
    Next i
End Function

Private Sub LV_Promo1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim ListX As ListItem
Dim iCount As Integer
        
        On Error GoTo ErrDescription
        
        If Item.Checked = True Then
                'FrmSaleOrder.Show
                Set ListX = LV_Promo2.ListItems.Add(, , LV_Promo1.ListItems(Item.Index).Text)
                ListX.SubItems(1) = Me.LV_Promo1.ListItems(Item.Index).SubItems(1)
                ListX.SubItems(2) = Me.LV_Promo1.ListItems(Item.Index).SubItems(2)
                ListX.SubItems(3) = Me.LV_Promo1.ListItems(Item.Index).SubItems(3)
                ListX.SubItems(4) = Me.LV_Promo1.ListItems(Item.Index).SubItems(4)
                ListX.SubItems(5) = Me.LV_Promo1.ListItems(Item.Index).SubItems(5)
        Else
                ' Uncheck Box Option
                'For iCount = Me.LV_Promo2.ListItems.count To 1 Step -1
                 '       If LV_Promo1.ListItems(Item.Index).SubItems(1) = LV_Promo2.ListItems.Item(iCount).SubItems(1) Then
                  '              LV_Promo1.ListItems.Remove (iCount)
                   '     End If
                'Next iCount
        End If
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub LV_Sop1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim ListX As ListItem
        Dim iCount As Integer
        
        On Error GoTo ErrDescription
        
        If Item.Checked = True Then
                'FrmSaleOrder.Show
                Set ListX = LV_Sop2.ListItems.Add(, , LV_Sop1.ListItems(Item.Index).Text)
                ListX.SubItems(1) = Me.LV_Sop1.ListItems(Item.Index).SubItems(1)
                ListX.SubItems(2) = Me.LV_Sop1.ListItems(Item.Index).SubItems(2)
                ListX.SubItems(3) = Me.LV_Sop1.ListItems(Item.Index).SubItems(3)
                ListX.SubItems(4) = Me.LV_Sop1.ListItems(Item.Index).SubItems(4)
                ListX.SubItems(5) = Me.LV_Sop1.ListItems(Item.Index).SubItems(5)
                ListX.SubItems(6) = Me.LV_Sop1.ListItems(Item.Index).SubItems(6)
        Else
                ' Uncheck Box Option
                For iCount = Me.LV_Sop2.ListItems.count To 1 Step -1
                        If LV_Sop1.ListItems(Item.Index).SubItems(1) = LV_Sop2.ListItems.Item(iCount).SubItems(1) Then
                                LV_Sop1.ListItems.Remove (iCount)
                        End If
                Next iCount
        End If
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub LV_TRF1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
        Dim ListX As ListItem
        Dim iCount As Integer
        
        On Error GoTo ErrDescription
        
        If Item.Checked = True Then
                ' Check Box Option
                'MsgBox Item.Index
                Set ListX = LV_TRF2.ListItems.Add(, , LV_TRF1.ListItems(Item.Index).Text)
                ListX.SubItems(1) = Me.LV_TRF1.ListItems(Item.Index).SubItems(1)
                ListX.SubItems(2) = Me.LV_TRF1.ListItems(Item.Index).SubItems(2)
                ListX.SubItems(3) = Me.LV_TRF1.ListItems(Item.Index).SubItems(3)
                ListX.SubItems(4) = Me.LV_TRF1.ListItems(Item.Index).SubItems(4)
                ListX.SubItems(5) = Me.LV_TRF1.ListItems(Item.Index).SubItems(5)
                ListX.SubItems(6) = Me.LV_TRF1.ListItems(Item.Index).SubItems(6)
                ListX.SubItems(7) = Me.LV_TRF1.ListItems(Item.Index).SubItems(7)
        Else
                ' Uncheck Box Option
                If Me.LV_TRF2.ListItems.count > 0 Then
                    For iCount = Me.LV_TRF2.ListItems.count To 1 Step -1
                            If LV_TRF1.ListItems(Item.Index).SubItems(1) = LV_TRF2.ListItems.Item(iCount).SubItems(1) Then
                                    LV_TRF2.ListItems.Remove (iCount)
                            End If
                    Next iCount
                End If
        End If
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If

End Sub

Private Sub Option1_Click()
        lbDetail.Caption = "รหัสสินค้า :"
        'Text1.SetFocus
End Sub

Private Sub Option10_Click()
Combo1.Visible = True
Label12.Visible = True
Label12.Visible = True
End Sub

Private Sub Option11_Click()

End Sub

Private Sub Option2_Click()
        lbDetail.Caption = "รหัสบาร์โค้ด :"
        Text1.SetFocus
End Sub

Private Sub Option3_Click()
        lbDetail.Caption = "ชื่อสินค้า :"
        Text1.SetFocus
End Sub

Private Sub Option4_Click()
        lbPO.Caption = "วันที่ออกใบสั่งซื้อ :"
        Text3.Visible = False
        DTPicker2.Visible = True
End Sub

Private Sub Option5_Click()
        lbPO.Caption = "รหัสเจ้าหนี้ :"
        Text3.Visible = True
        DTPicker2.Visible = False
End Sub

Private Sub Option6_Click()
        lbPO.Caption = "เลขที่เอกสาร :"
        Text3.Visible = True
        DTPicker2.Visible = False
End Sub

Private Sub Option7_Click()
On Error Resume Next

        lbPO2.Caption = "เลขที่เอกสาร :"
        Text4.Visible = True
        DTPicker1.Visible = False
        Text4.SetFocus
End Sub

Private Sub Option8_Click()
        Text4.Visible = True
        DTPicker1.Visible = False
        lbPO2.Caption = "รหัสเจ้าหนี้ :"
        Text4.SetFocus
End Sub

Private Sub Option9_Click()
        lbPO2.Caption = "วันที่ออกใบรับสินค้า :"
        Text4.Visible = False
        DTPicker1.Visible = True
        DTPicker1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrDescription
        
        If KeyAscii = 13 Then
            Call Command1_Click
            Call ItemLocation
        End If
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
        Dim ListX As ListItem
        Dim vPrice As Double
        Dim vSPrice As Double
        
        On Error GoTo ErrDescription
        
        If KeyAscii = 13 And vCheckRecProduct > 0 Then
                ' Check ว่าเป็นตัวเลขหรือไม่
                If IsNumeric(Text2.Text) = False Then
                        MsgBox "กรุณากรอกข้อมูลที่เป็นตัวเลข"
                        Text2.Text = ""
                        Text2.SetFocus
                        Exit Sub
                End If
                    
                Set ListX = Me.LV_Label.ListItems.Add(, , Trim(tmpItemNumber))
                ListX.SubItems(1) = Trim(tmpWHCode)
                ListX.SubItems(2) = Trim(tmpBarcod)
                ListX.SubItems(3) = Trim(tmpItemDesc)
                ListX.SubItems(4) = Trim(tmpUOFM)
                ListX.SubItems(5) = Trim(tmpPrice)
                ListX.SubItems(6) = Int(Text2.Text)
                ListX.SubItems(7) = Trim(tmpSPrice)
                ListX.SubItems(8) = Trim(tmpShelfCode)
                If Option10.Value = True Then
                    'ListX.SubItems(6) = Trim(Text5.Text)
                End If
        
            ' Clear Data and set Focus
            Text1.Text = ""
            Text2.Text = ""
            Text2.Enabled = False
            Label6.Caption = ""
            Text1.SetFocus
        End If
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub
Private Sub addListResult(StepNumber As Integer)
       ' Clear ListResult
       ListResult.ListItems.Clear
        
        Dim ListX As ListItem
        Dim iCount As Integer
        'MsgBox StepNumber
        On Error GoTo ErrDescription
        
        Select Case StepNumber
                Case 2:
                        For iCount = 1 To Me.LV_Label.ListItems.count
                                Set ListX = Me.ListResult.ListItems.Add(, , LV_Label.ListItems(iCount).Text)
                                ListX.SubItems(1) = Me.LV_Label.ListItems(iCount).SubItems(1)
                                ListX.SubItems(2) = Me.LV_Label.ListItems(iCount).SubItems(2)
                                ListX.SubItems(3) = Me.LV_Label.ListItems(iCount).SubItems(3)
                                ListX.SubItems(4) = Me.LV_Label.ListItems(iCount).SubItems(4)
                                ListX.SubItems(5) = Me.LV_Label.ListItems(iCount).SubItems(5)
                                ListX.SubItems(6) = Me.LV_Label.ListItems(iCount).SubItems(6)
                                ListX.SubItems(7) = Me.LV_Label.ListItems(iCount).SubItems(7)
                                ListX.SubItems(9) = Me.LV_Label.ListItems(iCount).SubItems(8)
                        Next
                Case 3:
                        For iCount = 1 To Me.LV_Palm.ListItems.count
                                Set ListX = Me.ListResult.ListItems.Add(, , LV_Palm.ListItems(iCount).Text)
                                ListX.SubItems(1) = Me.LV_Palm.ListItems(iCount).SubItems(1)
                                ListX.SubItems(2) = Me.LV_Palm.ListItems(iCount).SubItems(2)
                                ListX.SubItems(3) = Me.LV_Palm.ListItems(iCount).SubItems(3)
                                ListX.SubItems(4) = Me.LV_Palm.ListItems(iCount).SubItems(4)
                                ListX.SubItems(5) = Me.LV_Palm.ListItems(iCount).SubItems(5)
                        Next
                Case 4:
                        For iCount = 1 To Me.LV_PO_2.ListItems.count
                                Set ListX = Me.ListResult.ListItems.Add(, , LV_PO_2.ListItems(iCount).Text)
                                ListX.SubItems(1) = Me.LV_PO_2.ListItems(iCount).SubItems(1)
                                ListX.SubItems(2) = Me.LV_PO_2.ListItems(iCount).SubItems(2)
                                ListX.SubItems(3) = Me.LV_PO_2.ListItems(iCount).SubItems(3)
                                ListX.SubItems(4) = Me.LV_PO_2.ListItems(iCount).SubItems(4)
                                ListX.SubItems(5) = Me.LV_PO_2.ListItems(iCount).SubItems(5)
                        Next
                Case 5:
                        For iCount = 1 To Me.LV_PO1_2.ListItems.count
                                Set ListX = Me.ListResult.ListItems.Add(, , LV_PO1_2.ListItems(iCount).Text)
                                ListX.SubItems(1) = Me.LV_PO1_2.ListItems(iCount).SubItems(1)
                                ListX.SubItems(2) = Me.LV_PO1_2.ListItems(iCount).SubItems(2)
                                ListX.SubItems(3) = Me.LV_PO1_2.ListItems(iCount).SubItems(3)
                                ListX.SubItems(4) = Me.LV_PO1_2.ListItems(iCount).SubItems(4)
                                ListX.SubItems(5) = Me.LV_PO1_2.ListItems(iCount).SubItems(5)
                                ListX.SubItems(6) = Me.LV_PO1_2.ListItems(iCount).SubItems(6)
                                ListX.SubItems(7) = Me.LV_PO1_2.ListItems(iCount).SubItems(7)
                                ListX.SubItems(8) = Me.LV_PO1_2.ListItems(iCount).SubItems(8)
                        Next
                'Case 6:
                        'For iCount = 1 To Me.ListChkValue.ListItems.count
                         '       Set ListX = Me.ListResult.ListItems.Add(, , Me.ListChkValue.ListItems(iCount).Text)
                          '      ListX.SubItems(1) = Me.ListChkValue.ListItems(iCount).SubItems(1)
                           '     ListX.SubItems(2) = Me.ListChkValue.ListItems(iCount).SubItems(2)
                            '    ListX.SubItems(3) = Me.ListChkValue.ListItems(iCount).SubItems(3)
                             '   ListX.SubItems(4) = Me.ListChkValue.ListItems(iCount).SubItems(4)
                              '  ListX.SubItems(5) = Me.ListChkValue.ListItems(iCount).SubItems(5)
                        'Next
                        
                Case 6:
                
                        For iCount = 1 To Me.ListViewItemChangePrice.ListItems.count
                                If Me.ListViewItemChangePrice.ListItems.Item(iCount).Checked = True Then
                                    Set ListX = Me.ListResult.ListItems.Add(, , Me.ListViewItemChangePrice.ListItems(iCount).SubItems(1))
                                    ListX.SubItems(1) = Me.ListViewItemChangePrice.ListItems(iCount).SubItems(8)
                                    ListX.SubItems(2) = Me.ListViewItemChangePrice.ListItems(iCount).SubItems(2)
                                    ListX.SubItems(3) = Me.ListViewItemChangePrice.ListItems(iCount).SubItems(3)
                                    ListX.SubItems(4) = Me.ListViewItemChangePrice.ListItems(iCount).SubItems(4)
                                    ListX.SubItems(5) = Me.ListViewItemChangePrice.ListItems(iCount).SubItems(5)
                                    ListX.SubItems(6) = "1"
                                    ListX.SubItems(7) = Me.ListViewItemChangePrice.ListItems(iCount).SubItems(9)
                                    ListX.SubItems(8) = Me.ListViewItemChangePrice.ListItems(iCount).SubItems(7)
                                    ListX.SubItems(9) = Me.ListViewItemChangePrice.ListItems(iCount).SubItems(10)
                                End If
                        Next
                Case 7:
                        For iCount = 1 To Me.LV_Asset2.ListItems.count
                                Set ListX = Me.ListResult.ListItems.Add(, , LV_Asset2.ListItems(iCount).Text)
                                ListX.SubItems(1) = Me.LV_Asset2.ListItems(iCount).SubItems(1)
                                ListX.SubItems(2) = Me.LV_Asset2.ListItems(iCount).SubItems(2)
                                ListX.SubItems(3) = Me.LV_Asset2.ListItems(iCount).SubItems(3)
                                ListX.SubItems(4) = Me.LV_Asset2.ListItems(iCount).SubItems(4)
                                ListX.SubItems(5) = Me.LV_Asset2.ListItems(iCount).SubItems(5)
                        Next
                Case 8:
                        For iCount = 1 To Me.LV_Sop2.ListItems.count
                                Set ListX = Me.ListResult.ListItems.Add(, , LV_Sop2.ListItems(iCount).SubItems(1))
                                ListX.SubItems(1) = Me.LV_Sop2.ListItems(iCount).SubItems(6)
                                ListX.SubItems(2) = Me.LV_Sop2.ListItems(iCount).SubItems(1)
                                ListX.SubItems(3) = Me.LV_Sop2.ListItems(iCount).SubItems(2)
                                ListX.SubItems(4) = Me.LV_Sop2.ListItems(iCount).SubItems(3)
                                ListX.SubItems(5) = Me.LV_Sop2.ListItems(iCount).SubItems(4)
                                ListX.SubItems(6) = Me.LV_Sop2.ListItems(iCount).SubItems(5)
                        Next
                Case 9:
                        For iCount = 1 To Me.LV_TRF2.ListItems.count
                                Set ListX = Me.ListResult.ListItems.Add(, , LV_TRF2.ListItems(iCount).SubItems(1))
                                ListX.SubItems(1) = Me.LV_TRF2.ListItems(iCount).SubItems(7)
                                ListX.SubItems(2) = Me.LV_TRF2.ListItems(iCount).SubItems(2)
                                ListX.SubItems(3) = Me.LV_TRF2.ListItems(iCount).SubItems(3)
                                ListX.SubItems(4) = Me.LV_TRF2.ListItems(iCount).SubItems(4)
                                ListX.SubItems(5) = Me.LV_TRF2.ListItems(iCount).SubItems(5)
                                ListX.SubItems(6) = Me.LV_TRF2.ListItems(iCount).SubItems(6)
                        Next
                Case 10:
                        For iCount = 1 To Me.LV_ChangePrice2.ListItems.count
                                Set ListX = Me.ListResult.ListItems.Add(, , LV_ChangePrice2.ListItems(iCount).Text)
                                ListX.SubItems(1) = Me.LV_ChangePrice2.ListItems(iCount).SubItems(7)
                                ListX.SubItems(2) = Me.LV_ChangePrice2.ListItems(iCount).SubItems(1)
                                ListX.SubItems(3) = Me.LV_ChangePrice2.ListItems(iCount).SubItems(2)
                                ListX.SubItems(4) = Me.LV_ChangePrice2.ListItems(iCount).SubItems(3)
                                ListX.SubItems(5) = Me.LV_ChangePrice2.ListItems(iCount).SubItems(4)
                                ListX.SubItems(6) = "1"
                                ListX.SubItems(7) = Me.LV_ChangePrice2.ListItems(iCount).SubItems(8)
                                ListX.SubItems(8) = Me.LV_ChangePrice2.ListItems(iCount).SubItems(6)
                        Next
                Case 11:
                        For iCount = 1 To Me.ListViewShelf2.ListItems.count
                                Set ListX = Me.ListResult.ListItems.Add(, , ListViewShelf2.ListItems(iCount).SubItems(1))
                                ListX.SubItems(1) = Me.ListViewShelf2.ListItems(iCount).SubItems(3)
                                ListX.SubItems(2) = Me.ListViewShelf2.ListItems(iCount).SubItems(4)
                                ListX.SubItems(3) = Me.ListViewShelf2.ListItems(iCount).SubItems(2)
                        Next
        End Select
        
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub
Private Sub AddToTemp()
Dim iCount As Integer
Dim strItemNumber As String
Dim strSQL As String, vUnitCode As String
Dim vCheckPromotion As Integer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vItemCode As String
Dim vWHCode As String
Dim vBarCode As String
        
        If ListResult.ListItems.count = 0 Then
                MsgBox "ไม่พบรายการสินค้าที่ต้องการพิมพ์", vbOKOnly + vbInformation, "คำแนะนำ"
                DataHas = False
                Exit Sub
        End If
                
       ReDim SPrice(ListResult.ListItems.count)
        For iCount = 1 To ListResult.ListItems.count
            If Me.ListResult.ListItems(iCount).Checked = True Then
                DataHas = True
                vUnitCode = Trim(ListResult.ListItems.Item(iCount).SubItems(4))
                vItemCode = Trim(ListResult.ListItems.Item(iCount).Text)
                vBarCode = Trim(ListResult.ListItems.Item(iCount).SubItems(2))
                vWHCode = Trim(ListResult.ListItems.Item(iCount).SubItems(1))
                tmpBIN_ID = Trim(ListResult.ListItems.Item(iCount).SubItems(9))
                
                'ป้ายติดโกดัง
               vQuery = "exec dbo.USP_NP_SearchItemDetails '" & vItemCode & "' ,'" & vUnitCode & "','" & vBarCode & "','" & vWHCode & "' "
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                
                'Market
               ' vQuery = "exec dbo.USP_NP_SearchItemDetails_Market '" & vItemCode & "' ,'" & vUnitCode & "','" & vBarCode & "','" & vWHCode & "' "
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
                        If vCheckReceipt = True Then
                            tmpDocNo = ListResult.ListItems(iCount).SubItems(7)
                            tmpDocDate = ListResult.ListItems(iCount).SubItems(8)
                        End If
                        strItemNumber = Trim(vRecordset.Fields("Itemcode").Value)
                        tmpNUMBR = Trim(vRecordset.Fields("Itemcode").Value)
                        If Me.ListResult.ListItems(iCount).SubItems(2) = Null Then
                        tmpBarcode = Trim(vRecordset.Fields("barcode").Value)
                        Else
                        tmpBarcode = Me.ListResult.ListItems(iCount).SubItems(2)
                        End If
                        tmpTHINAME = Trim(vRecordset.Fields("name1").Value)
                        If Right(tmpTHINAME, 1) = "'" Then
                                tmpTHINAME = Left(tmpTHINAME, (Len(tmpTHINAME) - 1)) & """"
                        End If
                        tmpENGNAME = Trim(vRecordset.Fields("name2").Value)
                        tmpUOM = Trim(vRecordset.Fields("unitcode").Value)
                        tmpUsedUser = strUsername
                        tmpSite = Trim(ListResult.ListItems(iCount).SubItems(1))
                        
                        If tmpBIN_ID = "" Then
                           tmpBIN_ID = Trim(vRecordset.Fields("shelfcode").Value)
                        End If
                        
                        tmpONHAND = Trim(vRecordset.Fields("qty").Value)
                        tmpQTYALLOCATE = Trim(vRecordset.Fields("reserveqty").Value)
                        tmpCategory_ID = ""
                        tmpRemainOutQTY = Trim(vRecordset.Fields("RemainOutQTY").Value)
                        tmpRemainInQTY = Trim(vRecordset.Fields("RemainInQTY").Value)
                        tmpVENDR_ID = "" 'Trim(vRecordset.Fields("VENDERCODE").Value)
                        tmpRRCLV = Trim(vRecordset.Fields("salepromotion").Value)
                        tmpRRC = Trim(vRecordset.Fields("saleprice1").Value)
                        tmpSPrice = CheckDegit(Trim(vRecordset.Fields("priceerect").Value))
                        'If IsNull(Trim(vRecordset.Fields("priceerect").Value)) Or Trim(vRecordset.Fields("priceerect").Value) = "" Then
                         ' MsgBox "รหัสสินค้า " & strItemNumber & " ไม่มีราคาตั้ง กรุณากำหนดราคาตั้งและหน่วยราคาตั้งให้ตรงกับหน่วยขายจริงด้วย", vbCritical, "Send Error "
                          'Exit Sub
                        'End If
                        
                Else
                       MsgBox "สินค้ารหัส " & vItemCode & " และคลัง " & vWHCode & " มีข้อมูลไม่ครบไม่สามารถพิมพ์เอกสารได้ กรุณาตรวจสอบ บาร์โค้ด ที่เก็บ ราคาตั้ง ของสินค้าดังกล่าว"
                        tmpDocNo = ""
                        tmpDocDate = ""
                        strItemNumber = ""
                        tmpNUMBR = ""
                        tmpBarcode = ""
                        tmpTHINAME = ""
                        tmpTHINAME = ""
                        tmpENGNAME = ""
                        tmpUOM = ""
                        tmpSite = ""
                        tmpBIN_ID = ""
                        tmpONHAND = ""
                        tmpQTYALLOCATE = ""
                        tmpCategory_ID = ""
                        tmpRemainOutQTY = ""
                        tmpRemainInQTY = ""
                        tmpVENDR_ID = ""
                        tmpRRCLV = 0
                        tmpRRC = 0
                        tmpSPrice = ""
                End If
                vRecordset.Close
                tmpQTY = Int(ListResult.ListItems(iCount).SubItems(6))
                If IsNull(tmpSPrice) Or tmpSPrice = "" Then
                    tmpSPrice = 0
                End If
                If vCheckReceipt = True And strItemNumber <> "" Then
                    vQuery = "exec dbo.USP_NP_InsertLabelTemp '" & tmpNUMBR & "','" & tmpBarcode & "','" & tmpTHINAME & "','" & tmpENGNAME & "'," & Int(tmpQTY) & "," & tmpRRCLV & "," & tmpRRC & ",'" & tmpUOM & "','" & tmpUsedUser & "','" & tmpCategory_ID & "','" & tmpSite & "','" & tmpBIN_ID & "','" & tmpVENDR_ID & "','" & tmpRemark & "'," & Int(tmpSPrice) & ",'" & tmpONHAND & "','" & tmpQTYALLOCATE & "'," & tmpTYPE & ",'" & tmpRemainOutQTY & "','" & tmpRemainInQTY & "','" & tmpDocNo & "','" & tmpDocDate & "' "
                ElseIf vCheckReceipt = False And strItemNumber <> "" Then
                    vQuery = "exec dbo.USP_NP_InsertLabelTemp '" & tmpNUMBR & "','" & tmpBarcode & "','" & tmpTHINAME & "','" & tmpENGNAME & "'," & Int(tmpQTY) & "," & tmpRRCLV & "," & tmpRRC & ",'" & tmpUOM & "','" & tmpUsedUser & "','" & tmpCategory_ID & "','" & tmpSite & "','" & tmpBIN_ID & "','" & tmpVENDR_ID & "','" & tmpRemark & "'," & Int(tmpSPrice) & ",'" & tmpONHAND & "','" & tmpQTYALLOCATE & "'," & tmpTYPE & ",'" & tmpRemainOutQTY & "','" & tmpRemainInQTY & "','" & tmpDocNo & "','" & tmpDocDate & "' "
                End If
                vConnection.Execute vQuery
            End If
        Next
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                Call Command2_Click
        End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                Call Command3_Click
        End If
End Sub

Private Sub Text6_Change()

End Sub

Private Sub TXTAsset1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                Call CMDAsset1_Click
        End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                Call ChkPWD
        End If
End Sub

Private Sub TXTShelf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TXTShelf2.SetFocus
End If
End Sub

Private Sub TXTShelf2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CMDShelf_Click
End If
End Sub

Private Sub TXTSop1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                CMDSop1_Click
        End If
End Sub

Private Sub TXTTRF1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                CMDTRF1_Click
        End If
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                txtPassword.SetFocus
        End If
End Sub

Private Sub ChkPWD()
    ' Declaration
    Dim strError, tmpErrorNumber, tmpErrorDesc As String
    Dim ConnChkPWD As New ADODB.Connection

    'On Error GoTo Err1:
    ' Process Check Password NPDL Table
        On Error GoTo ErrDescription:
    
    With ConnChkPWD
        If .State = adStateOpen Then .Close
                .Provider = "SQLOLEDB"
                .Properties("Persist Security Info").Value = False
                .Properties("User ID").Value = txtUsername.Text
                .Properties("Password").Value = txtPassword.Text
                .Properties("Data Source").Value = "S02DB" '"GALAXY"
                .Properties("Initial Catalog").Value = "BCNP" 'ทำการเปลี่ยนฐานข้อมูลเป็น BCNP
                .CursorLocation = adUseClient
                .Open
    End With
    ConnChkPWD.Close
    
    'On Error GoTo err2:
    'With ConnChkPWD
     '   If .State = adStateOpen Then .Close
      '          .Provider = "SQLOLEDB"
       '         .Properties("Persist Security Info").Value = False
        '        .Properties("User ID").Value = txtUsername.Text
         '       .Properties("Password").Value = txtPassword.Text
          '      .Properties("Data Source").Value = "S02DB" '"GALAXY"
           '     .Properties("Initial Catalog").Value = "NPDEV"
            '    .CursorLocation = adUseClient
             '   .Open
    'End With
    'ConnChkPWD.Close

    ' Add Username & Password
    strUsername = Trim(txtUsername.Text)
    strPassword = Trim(txtPassword.Text)
    Call InitializeDatabase
    Call BeginConnect
    Exit Sub
    
    
' Error Connect NPDL Table
'Err1:
 '       tmpErrorNumber = Err.Number
  '      tmpErrorDesc = Err.Description
        ' MsgBox tmpErrorNumber
        
        ' Connect to Error in Thai Language

   '     ConnTipDB
    '    strError = "Select ErrNum,ErrDesc From ErrorDesc Where ErrNum = '" & Abs(tmpErrorNumber) & "'"
     '   Rs1.Open strError, ConnAccess, 1, 3
      '  If Not Rs1.EOF Then
       '         MsgBox Rs1!ErrDesc, vbOKOnly + vbCritical, "คำแนะนำ (NPDL)"
        'Else
         '       MsgBox Abs(tmpErrorNumber) & " - " & tmpErrorDesc, vbOKOnly + vbCritical, "คำแนะนำ (NPDL)"
          '      Rs1.AddNew
           '     Rs1!ErrNum = Abs(tmpErrorNumber)
            '    Rs1!ErrDesc = Abs(tmpErrorDesc)
             '   Rs1.Update
        'End If
        'Rs1.Close       ' Close Connection
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    txtPassword.Text = ""
    Exit Sub
End If
        
        'txtUsername.Text = ""
        'txtPassword.Text = ""
        'txtUsername.SetFocus
        'Exit Sub

' Error Connect NPDEV Table
'err2:
 '       tmpErrorNumber = Err.Number
  '      tmpErrorDesc = Err.Description
   '     ' MsgBox tmpErrorNumber
        
        ' Connect to Error in Thai Language
    '    ConnTipDB
     '   strError = "Select ErrNum,ErrDesc From ErrorDesc Where ErrNum = '" & Abs(tmpErrorNumber) & "'"
      '  Rs1.Open strError, ConnAccess, 1, 3
       ' If Not Rs1.EOF Then
        '        MsgBox Rs1!ErrDesc, vbOKOnly + vbCritical, "คำแนะนำ (NPDEV)"
        'Else
         '       MsgBox Abs(tmpErrorNumber) & " - " & tmpErrorDesc, vbOKOnly + vbCritical, "คำแนะนำ (NPDEV)"
          '      Rs1.AddNew
           '     Rs1!ErrNum = Abs(tmpErrorNumber)
            '    Rs1!ErrDesc = Abs(tmpErrorDesc)
             '   Rs1.Update
        'End If
        'Rs1.Close       ' Close Connection
        
        'txtUsername.Text = ""
        'txtPassword.Text = ""
        'txtUsername.SetFocus
        'Exit Sub
End Sub
Public Sub ItemLocation()
Dim vQuery As String

On Error GoTo ErrDescription
ConnectSQL
vQuery = "select WHCODE,shelfcode from  bcrecproduct where productcode = '" & Trim(Text1.Text) & "' "
Rs1.Open vQuery, ConnSQL, adOpenDynamic, adLockOptimistic
If Not Rs1.EOF Then
    Rs1.MoveFirst
        Do Until Rs1.EOF
            Combo1.AddItem Trim(Rs1.Fields("WHCODE").Value)
             Rs1.MoveNext
        Loop
   
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Private Sub GetWHCode()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription
CMBSection.Clear
vQuery = "select  distinct secman, secman+'//'+ secmanname as secmanname   from dbo.vw_PRG_SearchSecMan order by secman"
If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBSection.AddItem Trim(vRecordset.Fields("secmanname").Value)
             vRecordset.MoveNext
        Wend
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Private Sub AddWHCode()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

CMBWH.Clear
vQuery = "select code from bcwarehouse order by code"
If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBWH.AddItem Trim(vRecordset.Fields("code").Value)
             vRecordset.MoveNext
        Wend
End If
vRecordset.Close
CMBWH.Text = Trim("014")

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Private Sub AddTempByChangePrice()
        Dim iCount As Integer
        Dim strSQL As String
        Dim vQuery As String
        ' Dim ListX As ListItem
        
        'ConnectSQL
        On Error GoTo ErrDescription
        
        ' Check Data in Listresult
        If ListResult.ListItems.count = 0 Then
                MsgBox "ไม่พบรายการสินค้าที่ต้องการพิมพ์", vbOKOnly + vbInformation, "คำแนะนำ"
                DataHas = False
                Exit Sub
        End If
        
        For iCount = 1 To ListResult.ListItems.count
            If Me.ListResult.ListItems(iCount).Checked = True Then
                DataHas = True
                tmpNUMBR = Me.ListResult.ListItems(iCount).Text
                tmpBarcode = Me.ListResult.ListItems(iCount).SubItems(2)
                tmpTHINAME = Me.ListResult.ListItems(iCount).SubItems(3)
                tmpUOM = Me.ListResult.ListItems(iCount).SubItems(4)
                tmpRRC = Me.ListResult.ListItems(iCount).SubItems(5)
                tmpQTY = 1
                tmpUpDateTime = Me.ListResult.ListItems(iCount).SubItems(8)
                ' ยังไม่มีค่าใส่
                tmpUsedUser = strUsername
                tmpCategory_ID = ""
                tmpSite = Me.ListResult.ListItems(iCount).SubItems(1)
                tmpBIN_ID = Me.LV_ChangePrice2.ListItems(iCount).SubItems(9)
                tmpVENDR_ID = ""
                tmpRemark = ""
                tmpSPrice = Me.ListResult.ListItems(iCount).SubItems(7)
                tmpONHAND = ""
                tmpQTYALLOCATE = ""
                tmpTYPE = 0
                tmpRRCLV = 0
                
                ' Add To NP_Label_Temp
                vQuery = "Insert into NP_LABEL_TEMP(Itemcode, barcode, name1, NAME2, QTY, PriceLevel, Price, Unitcode,UsedUser,Category_ID,WHCode,ShelfCode,VENDR_ID,remark,SPrice,ONHAND,QTYALLOCATE,Type)  " _
                & "values('" & tmpNUMBR & "','" & tmpBarcode & "','" & tmpTHINAME & "','" & tmpENGNAME & "'," & Int(tmpQTY) & "," & tmpRRCLV & "," & tmpRRC & ",'" & tmpUOM & "','" & tmpUsedUser & "','" & tmpCategory_ID & "','" & tmpSite & "','" & tmpBIN_ID & "','" & tmpVENDR_ID & "','" & tmpRemark & "'," & Int(tmpSPrice) & ",'" & tmpONHAND & "','" & tmpQTYALLOCATE & "'," & tmpTYPE & ")"
                vConnection.Execute vQuery
            
            End If
        Next iCount
 '----------------------------------------------------------------------------------------------------
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Private Sub AddTempAsset()
        Dim iCount As Integer
        Dim strSQL As String
        ' Dim ListX As ListItem
        
        ConnectSQL
        On Error GoTo ErrDescription
        
        ' Check Data in Listresult
        If ListResult.ListItems.count = 0 Then
                MsgBox "ไม่พบรายการสินค้าที่ต้องการพิมพ์", vbOKOnly + vbInformation, "คำแนะนำ"
                DataHas = False
                Exit Sub
        End If
        
        For iCount = 1 To ListResult.ListItems.count
            If Me.ListResult.ListItems(iCount).Checked = True Then
               DataHas = True
                tmpNUMBR = Me.ListResult.ListItems(iCount).Text
                tmpBarcode = Me.ListResult.ListItems(iCount).Text
                tmpTHINAME = Me.ListResult.ListItems(iCount).SubItems(2)
                tmpUOM = Me.ListResult.ListItems(iCount).SubItems(3)
                tmpRRC = Me.ListResult.ListItems(iCount).SubItems(4)
                tmpQTY = 2
                tmpUsedUser = strUsername
                tmpSPrice = 0
                tmpTYPE = 0
                strSQL = "select a.buydate,b.name from bcassetsmaster a     " _
                                    & "left join bcdepartment b on a.departcode = b.code   " _
                                    & "where   a.code = '" & tmpNUMBR & "'  "
                Rs1.Open strSQL, ConnSQL, adOpenDynamic, adLockOptimistic
                If Not Rs1.EOF Then
                    tmpTHIName1 = Rs1.Fields("name").Value
                    tmpBuyDocDate = Rs1.Fields("buydate").Value
                End If
                Rs1.Close
                ' Add To NP_Label_Temp
                strSQL = "Insert into NP_LABEL_TEMP(Itemcode, barcode, name1, name2,QTY, Price, Unitcode,UsedUser,SPrice,SOPDoc,Type)  " _
                & "values('" & tmpNUMBR & "','" & tmpBarcode & "','" & tmpTHINAME & "','" & tmpTHIName1 & "'," & Int(tmpQTY) & "," & tmpRRC & ",'" & tmpUOM & "','" & tmpUsedUser & "'," & tmpSPrice & ",'" & tmpBuyDocDate & "'," & tmpTYPE & ")"
                ConnSQL.Execute strSQL
            End If
        Next iCount
 '----------------------------------------------------------------------------------------------------
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub


Public Sub PrintForm1()
Dim tmpPathName As String

            tmpPathName = Trim(LV_Report.SelectedItem.SubItems(1)) & ".rpt"
            With Crystal102
                    .ReportFileName = tmpPathName
                    .ParameterFields(0) = "@vUserID;" & tmpUsedUser & ";true"
                    .WindowState = crptMaximized
                    .Connect = "uid=VBUser;pwd=132"
                        If FormToPrinter = False Then
                                .Destination = crptToWindow
                        End If
                        If FormToPrinter = True Then
'                                    .Destination = crptToPrinter
                        End If
                         .Action = 1
        End With

End Sub

Public Sub GetNamePromotion()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListPromo As ListItem

On Error GoTo ErrDescription

vQuery = " execute  npmaster.dbo.USP_PM_SelectPromotion "
If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB101.AddItem Trim(vRecordset.Fields("pmcode").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Private Sub AddTempPromotion()
        Dim iCount As Integer
        Dim strSQL As String
        ' Dim ListX As ListItem
        
        ConnectSQL
        On Error GoTo ErrDescription
        
        ' Check Data in Listresult
        If ListResult.ListItems.count = 0 Then
                MsgBox "ไม่พบรายการสินค้าที่ต้องการพิมพ์", vbOKOnly + vbInformation, "คำแนะนำ"
                DataHas = False
                Exit Sub
        End If
        
        For iCount = 1 To ListResult.ListItems.count
            If Me.ListResult.ListItems(iCount).Checked = True Then
                DataHas = True
                tmpNUMBR = Me.ListResult.ListItems(iCount).Text
                tmpBarcode = Me.ListResult.ListItems(iCount).SubItems(2)
                tmpTHINAME = Me.ListResult.ListItems(iCount).SubItems(3)
                tmpUsedUser = strUsername
                tmpSite = Me.ListResult.ListItems(iCount).SubItems(1)

                strSQL = "USP_LB_InsertPrintShelf '" & tmpNUMBR & "','" & tmpTHINAME & "','" & tmpUsedUser & "','" & tmpSite & "' "
                ConnSQL.Execute strSQL
                
            End If
        Next iCount
 '----------------------------------------------------------------------------------------------------
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub GetSectionName()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListPromo As ListItem

On Error GoTo ErrDescription

vQuery = "select distinct secman  from NPMaster.dbo.TB_PM_PromotionItem order by secman"
If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB102.AddItem Trim(vRecordset.Fields("secman").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub
