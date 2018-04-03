VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPrintReceiveSlip 
   Caption         =   "พิมพ์ใบจ่ายสินค้า"
   ClientHeight    =   8100
   ClientLeft      =   5325
   ClientTop       =   1800
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   12045
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8115
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   14314
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   706
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "พิมพ์ใบจ่ายสินค้า"
      TabPicture(0)   =   "FrmPrintReceiveSlip.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListViewItemList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Picture1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "บันทึกข้อมูลสินค้าที่มีปัญหาในการจ่าย"
      TabPicture(1)   =   "FrmPrintReceiveSlip.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "LBLWHCode"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label11"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Picture2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "ListViewItemCheckQTY"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "TextSlipNo"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "PICCheckQTY"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Command3"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Command4"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.CommandButton Command4 
         Caption         =   "ใบจ่ายถูกยกเลิก"
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
         Left            =   8775
         TabIndex        =   34
         Top             =   6705
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "บิลถูกยกเลิก"
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
         Left            =   7245
         TabIndex        =   33
         Top             =   6705
         Width           =   1455
      End
      Begin VB.PictureBox PICCheckQTY 
         BackColor       =   &H80000009&
         Height          =   3120
         Left            =   270
         ScaleHeight     =   3060
         ScaleWidth      =   11430
         TabIndex        =   29
         Top             =   2745
         Visible         =   0   'False
         Width           =   11490
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check3"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   855
            TabIndex        =   32
            Top             =   1710
            Width           =   3705
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check2"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   855
            TabIndex        =   31
            Top             =   1260
            Width           =   3705
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   855
            TabIndex        =   30
            Top             =   675
            Width           =   3750
         End
      End
      Begin VB.CommandButton Command2 
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
         Left            =   10305
         TabIndex        =   24
         Top             =   6705
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "บันทึกการจ่าย"
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
         Left            =   5715
         TabIndex        =   23
         Top             =   6705
         Width           =   1455
      End
      Begin VB.TextBox TextSlipNo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1620
         TabIndex        =   22
         Top             =   1080
         Width           =   2580
      End
      Begin MSComctlLib.ListView ListViewItemCheckQTY 
         Height          =   4020
         Left            =   270
         TabIndex        =   19
         Top             =   2385
         Width           =   11490
         _ExtentX        =   20267
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
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   9172
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "จำนวน"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วย"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "คลัง"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewItemList 
         Height          =   5910
         Left            =   -70095
         TabIndex        =   17
         Top             =   1530
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   10425
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   5292
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
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6405
         Left            =   -75000
         ScaleHeight     =   6375
         ScaleWidth      =   4665
         TabIndex        =   6
         Top             =   1035
         Width           =   4695
         Begin VB.ListBox ListChecker 
            Height          =   1425
            Left            =   315
            TabIndex        =   18
            Top             =   4185
            Visible         =   0   'False
            Width           =   4155
         End
         Begin VB.CommandButton CMDPrintReceiveSlip 
            Caption         =   "พิมพ์ใบจ่าย"
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
            Left            =   1935
            TabIndex        =   2
            Top             =   3195
            Width           =   1680
         End
         Begin VB.CommandButton CMDSelectChecker 
            Height          =   339
            Left            =   3645
            TabIndex        =   1
            Top             =   2150
            Width           =   420
         End
         Begin VB.TextBox TextInvoiceNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   330
            Left            =   1440
            TabIndex        =   0
            Top             =   270
            Width           =   1905
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ชื่อลูกค้า :"
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
            Left            =   45
            TabIndex        =   15
            Top             =   1260
            Width           =   1365
         End
         Begin VB.Label LBLARName 
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
            Height          =   735
            Left            =   1440
            TabIndex        =   14
            Top             =   1260
            Width           =   3030
         End
         Begin VB.Label LBLSaleCode 
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
            Left            =   1440
            TabIndex        =   13
            Top             =   2655
            Width           =   2175
         End
         Begin VB.Label LBLChecker 
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
            Left            =   1440
            TabIndex        =   12
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label LBLARCode 
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
            Left            =   1440
            TabIndex        =   11
            Top             =   765
            Width           =   1905
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ชื่อพนักงานขาย :"
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
            Left            =   0
            TabIndex        =   10
            Top             =   2655
            Width           =   1410
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ชื่อเช็คเกอร์ :"
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
            Height          =   240
            Left            =   180
            TabIndex        =   9
            Top             =   2160
            Width           =   1230
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   270
            TabIndex        =   8
            Top             =   765
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "เลขที่บิลขาย :"
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
            Left            =   315
            TabIndex        =   7
            Top             =   270
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   0
         Picture         =   "FrmPrintReceiveSlip.frx":0038
         ScaleHeight     =   390
         ScaleWidth      =   12045
         TabIndex        =   5
         Top             =   495
         Width           =   12075
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   -75000
         Picture         =   "FrmPrintReceiveSlip.frx":01C0
         ScaleHeight     =   390
         ScaleWidth      =   12045
         TabIndex        =   4
         Top             =   495
         Width           =   12075
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4185
         TabIndex        =   28
         Top             =   1530
         Width           =   7575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "เช็คเกอร์ :"
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
         Left            =   3105
         TabIndex        =   27
         Top             =   1575
         Width           =   1050
      End
      Begin VB.Label LBLWHCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1620
         TabIndex        =   26
         Top             =   1530
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Left            =   270
         TabIndex        =   25
         Top             =   1575
         Width           =   1320
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "เลขที่ใขจ่ายสินค้า :"
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
         Left            =   135
         TabIndex        =   21
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Height          =   285
         Left            =   270
         TabIndex        =   20
         Top             =   2070
         Width           =   1275
      End
      Begin VB.Label Label10 
         Caption         =   "รายการสินค้าจุดจ่าย"
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
         Left            =   -70095
         TabIndex        =   16
         Top             =   1170
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmPrintReceiveSlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String

Private Sub Form_Load()
Call InitializeConnectDataBase1
End Sub

Private Sub GetItemDetails(SlipNo As String)
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem

vQuery = "exec dbo.USP_INV_ItemLocateReceiptSlip " & vSelectZoneID & ",'" & SlipNo & "' "
If OpenDataBase1(vConnection, vRecordset, vQuery) <> 0 Then
  i = 1
  While Not vRecordset.EOF
  Set vListItem = ListViewItemCheckQTY.ListItems.Add(, , i)
    vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
    vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
    vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
    vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
    vListItem.SubItems(5) = Trim(vRecordset.Fields("CheckZoneLocation").Value)
    vRecordset.MoveNext
    i = i + 1
  Wend
End If
vRecordset.Close

End Sub

Private Sub TextInvoiceNo_KeyPress(KeyAscii As Integer)
Dim vInvoiceNo As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem


If KeyAscii = 13 Then
  vInvoiceNo = UCase(Me.TextSlipNo.Text)
  vQuery = "exec dbo.USP_INV_ItemLocateReceiptSlip " & vSelectZoneID & ",'" & vInvoiceNo & "' "
  If OpenDataBase1(vConnection, vRecordset, vQuery) <> 0 Then
    Me.LBLARCode.Caption = vRecordset.Fields("arcode").Value
    Me.LBLARName.Caption = vRecordset.Fields("arname").Value
    Me.LBLSaleCode.Caption = vRecordset.Fields("salename").Value
    i = 1
    While Not vRecordset.EOF
    Set vListItem = ListViewItemList.ListItems.Add(, , i)
      vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
      vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
      vListItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
      If Trim(vRecordset.Fields("CheckZoneLocation").Value) = "010" Then
        ListViewItemList.ListItems(i).ForeColor = "&H00400000"
        ListViewItemList.ListItems.Item(i).ListSubItems(1).ForeColor = "&H00400000"
        ListViewItemList.ListItems.Item(i).ListSubItems(2).ForeColor = "&H00400000"
        ListViewItemList.ListItems.Item(i).ListSubItems(3).ForeColor = "&H00400000"
        ListViewItemList.ListItems.Item(i).ListSubItems(4).ForeColor = "&H00400000"
      ElseIf Trim(vRecordset.Fields("CheckZoneLocation").Value) = "010C" Then
        ListViewItemList.ListItems(i).ForeColor = "&H00000000"
        ListViewItemList.ListItems.Item(i).ListSubItems(1).ForeColor = "&H00000000"
        ListViewItemList.ListItems.Item(i).ListSubItems(2).ForeColor = "&H00000000"
        ListViewItemList.ListItems.Item(i).ListSubItems(3).ForeColor = "&H00000000"
        ListViewItemList.ListItems.Item(i).ListSubItems(4).ForeColor = "&H00000000"
      End If
    
      vRecordset.MoveNext
      i = i + 1
    Wend
  End If
  vRecordset.Close
  
  Call GetItemDetails(vInvoiceNo)
End If
End Sub
