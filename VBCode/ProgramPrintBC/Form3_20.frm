VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form3_20 
   Caption         =   "พิมพ์ใบกำกับภาษีเอกสาร POS"
   ClientHeight    =   9600
   ClientLeft      =   2370
   ClientTop       =   1155
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9600
   ScaleWidth      =   14880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICAr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9600
      Left            =   0
      Picture         =   "Form3_20.frx":0000
      ScaleHeight     =   9570
      ScaleWidth      =   39090
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   39120
      Begin VB.CommandButton CMDARExit 
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
         Height          =   555
         Left            =   12735
         TabIndex        =   34
         Top             =   7515
         Width           =   1365
      End
      Begin VB.CommandButton CMDAROK 
         Caption         =   "เลือก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   11250
         TabIndex        =   33
         Top             =   7515
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListViewSearchAr 
         Height          =   4515
         Left            =   720
         TabIndex        =   22
         Top             =   2835
         Width           =   13380
         _ExtentX        =   23601
         _ExtentY        =   7964
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสลูกค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ที่อยู่ลูกค้า"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "เบอร์โทร"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "แฟกส์"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.CommandButton CMDSearchArOK 
         Height          =   420
         Left            =   6750
         Picture         =   "Form3_20.frx":9673
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1845
         Width           =   330
      End
      Begin VB.TextBox TextSearchARCode 
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
         Left            =   1710
         TabIndex        =   19
         Top             =   1845
         Width           =   4920
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ค้นหาลูกค้า เพื่อทำใบกำกับภาษีอย่างเต็ม"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   720
         TabIndex        =   32
         Top             =   1260
         Width           =   6045
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการลูกค้า"
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
         Left            =   720
         TabIndex        =   21
         Top             =   2520
         Width           =   1635
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   18
         Top             =   1890
         Width           =   1545
      End
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   3330
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
   Begin VB.CommandButton CMDExit 
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
      Left            =   13230
      TabIndex        =   5
      Top             =   8145
      Width           =   1455
   End
   Begin VB.CommandButton CMDPrint 
      Caption         =   "พิมพ์"
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
      Left            =   11430
      TabIndex        =   4
      Top             =   8145
      Width           =   1545
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C00000&
      Height          =   1995
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   14805
      TabIndex        =   8
      Top             =   7560
      Width           =   14865
      Begin VB.CommandButton CMDPrintDriveThru 
         Caption         =   "พิมพ์ ไดรฟ์ทรู"
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
         Left            =   9495
         TabIndex        =   35
         Top             =   550
         Width           =   1545
      End
      Begin Crystal.CrystalReport Crystal102 
         Left            =   2790
         Top             =   630
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
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000080FF&
      Height          =   4425
      Left            =   0
      ScaleHeight     =   4365
      ScaleWidth      =   14805
      TabIndex        =   7
      Top             =   3015
      Width           =   14865
      Begin MSComctlLib.ListView ListViewItem 
         Height          =   3885
         Left            =   270
         TabIndex        =   3
         Top             =   270
         Width           =   14235
         _ExtentX        =   25109
         _ExtentY        =   6853
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
            Text            =   "รหัส/ชื่อสินค้า"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "จำนวน"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หน่วย"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "ราคา"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "ส่วนลด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "มูลค่าสินค้า"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C00000&
      Height          =   2535
      Left            =   -45
      ScaleHeight     =   2475
      ScaleWidth      =   14850
      TabIndex        =   6
      Top             =   0
      Width           =   14910
      Begin VB.CommandButton CMDGenTaxNo 
         Height          =   375
         Left            =   5625
         Picture         =   "Form3_20.frx":9A40
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   675
         Width           =   375
      End
      Begin VB.CommandButton CMDSearchAr 
         Height          =   375
         Left            =   4140
         Picture         =   "Form3_20.frx":9E93
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1170
         Width           =   375
      End
      Begin VB.TextBox TXTARCode 
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
         Left            =   2385
         TabIndex        =   15
         Top             =   1170
         Width           =   1725
      End
      Begin VB.CommandButton CMDSearchDocNo 
         Caption         =   "ตรวจสอบข้อมูล/โอน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   13815
         Picture         =   "Form3_20.frx":A260
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.TextBox TBDocNo 
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
         Height          =   390
         Left            =   2385
         TabIndex        =   0
         Top             =   180
         Width           =   3165
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "แฟกส์ :"
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
         Left            =   8145
         TabIndex        =   31
         Top             =   2115
         Width           =   780
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "โทร :"
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
         Left            =   1530
         TabIndex        =   30
         Top             =   2115
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ที่อยู่ :"
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
         Left            =   1575
         TabIndex        =   29
         Top             =   1710
         Width           =   690
      End
      Begin VB.Label LBLFax 
         BackColor       =   &H00C0C0FF&
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
         Left            =   9045
         TabIndex        =   28
         Top             =   2070
         Width           =   5010
      End
      Begin VB.Label LBLTel 
         BackColor       =   &H00C0C0FF&
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
         Left            =   2385
         TabIndex        =   27
         Top             =   2070
         Width           =   5190
      End
      Begin VB.Label LBLAddress 
         BackColor       =   &H00C0C0FF&
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
         Left            =   2385
         TabIndex        =   26
         Top             =   1665
         Width           =   11670
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบกำกับภาษีอย่างเต็ม :"
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
         Left            =   -90
         TabIndex        =   25
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label LBLTaxNo 
         BackColor       =   &H00C0C0FF&
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
         Left            =   2385
         TabIndex        =   23
         Top             =   675
         Width           =   3165
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   1260
         TabIndex        =   14
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Label LBLDocType 
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
         Height          =   420
         Left            =   13680
         TabIndex        =   13
         Top             =   405
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label LBLDocNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   420
         Left            =   13905
         TabIndex        =   11
         Top             =   405
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label LBLArCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   5625
         TabIndex        =   2
         Top             =   1170
         Width           =   8430
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า :"
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
         Left            =   4590
         TabIndex        =   10
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบกำกับภาษีอย่างย่อ :"
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
         Left            =   45
         TabIndex        =   9
         Top             =   225
         Width           =   2220
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการ สินค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   540
      TabIndex        =   12
      Top             =   2610
      Width           =   2355
   End
End
Attribute VB_Name = "Form3_20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vMemIsOpen As Integer

Private Sub CMDARExit_Click()
Me.PICAr.Visible = False
Me.TXTARCode.SetFocus
End Sub

Private Sub CMDAROK_Click()
Dim vIndex As Integer

On Error GoTo ErrDescription

If Me.ListViewSearchAr.ListItems.Count > 0 Then
  vIndex = Me.ListViewSearchAr.SelectedItem.Index
  Me.TXTARCode.Text = Me.ListViewSearchAr.ListItems(vIndex).SubItems(1)
  Me.LBLArCode.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(2)
  Me.LBLAddress.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(3)
  Me.LBLTel.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(4)
  Me.LBLFax.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(5)
  Me.PICAr.Visible = False
Else
  Me.TXTARCode.Text = ""
  Me.LBLArCode.Caption = ""
  Me.LBLAddress.Caption = ""
  Me.LBLTel.Caption = ""
  Me.LBLFax.Caption = ""
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDExit_Click()
Unload Form3_20
End Sub

Private Sub CMDGenTaxNo_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vTaxNo As String

On Error GoTo ErrDescription

If Me.TBDocNo.Text <> "" And Me.ListViewItem.ListItems.Count > 0 And Me.LBLTaxNo.Caption = "" Then
    vQuery = "exec dbo.USP_NP_GetMaxTaxNo 'S01' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vTaxNo = Trim(vRecordset.Fields("taxno").Value)
    End If
    vRecordset.Close
    
    Me.LBLTaxNo.Caption = vTaxNo
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vCheck As Integer
Dim vDueDate As String
Dim vDocType As Integer

Dim vTaxNo As String
Dim vARCode As String

Dim vCheckInputTax As Integer

On Error GoTo ErrDescription


If Me.LBLTaxNo.Caption = "" Then
    MsgBox "กรุณา กดปุ่มกำหนดเลขที่ใบกำกับภาษีอย่างเต็ม", vbCritical, "Send Error Message"
    Me.CMDGenTaxNo.SetFocus
    Exit Sub
End If

If Me.LBLArCode.Caption = "" Then
    MsgBox "กรุณา เลือกลูกค้า", vbCritical, "Send Error Message"
    Me.TXTARCode.SetFocus
    Exit Sub
End If

If Me.LBLArCode.Caption = "เงินสด" Then
    MsgBox "กรุณา เลือกลูกค้าที่ไม่ใช่เงินสด", vbCritical, "Send Error Message"
    Me.TXTARCode.SetFocus
    Exit Sub
End If


If Me.TBDocNo.Text <> "" And Me.LBLTaxNo.Caption <> "" And Me.LBLArCode.Caption <> "" And Me.LBLArCode.Caption <> "เงินสด" Then
vDocNo = Trim(Me.LBLDocNo.Caption)
vDocType = Me.LBLDocType.Caption
vTaxNo = Me.LBLTaxNo.Caption
vARCode = Me.TXTARCode.Text

If vMemIsOpen = 0 Then
    
    vQuery = "exec dbo.USP_NP_CheckInputTax '" & vTaxNo & "'"
    If OpenDataBase(gConnection, vRecordset1, vQuery) <> 0 Then
        vCheckInputTax = Trim(vRecordset1.Fields("vCount").Value)
    End If
    vRecordset1.Close
    
    If vCheckInputTax > 0 Then
        MsgBox "เอกสารใบกำกับภาษีนี้ มีอยู่แล้วกรุณากดปุ่มกำหนดเลขที่ใบกำกับภาษีอย่างเต็มใหม่", vbCritical, "Send Error Message"
        Me.CMDGenTaxNo.SetFocus
        Exit Sub
    End If
    
    vQuery = "exec dbo.USP_NP_InsertOutPutTax '" & vDocNo & "','" & vTaxNo & "','" & vARCode & "' "
    gConnection.Execute vQuery
End If

vMemIsOpen = 0

If vDocType = 1 Then
vRepType = "INV"
vRepID = 480

 vDueDate = ""
 vCheck = 1
 vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     With Me.Crystal101
         .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
         .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
         .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
         .Formulas(0) = "CreditCondition='" & vDueDate & "' "
         .Destination = crptToWindow
         .WindowState = crptMaximized
         .Action = 1
     End With
 End If
 vRecordset.Close
 End If
 
'If vDocType = 2 Then
'vRepType = "RT"
'vRepID = 64

 'vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  '   With Me.Crystal102
   '      .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
    '     .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
     '    .Destination = crptToWindow
      '   .WindowState = crptMaximized
       '  .Action = 1
     'End With
 'End If
 'vRecordset.Close
 'End If
 
 
 Me.TBDocNo.Text = ""
 Me.LBLDocNo.Caption = ""
 Me.LBLDocType.Caption = ""
 Me.LBLArCode.Caption = ""
 Me.ListViewItem.ListItems.Clear
 Me.TBDocNo.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDPrintDriveThru_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vCheck As Integer
Dim vDueDate As String
Dim vDocType As Integer

Dim vTaxNo As String
Dim vARCode As String

Dim vCheckInputTax As Integer

On Error GoTo ErrDescription


If Me.LBLTaxNo.Caption = "" Then
    MsgBox "กรุณา กดปุ่มกำหนดเลขที่ใบกำกับภาษีอย่างเต็ม", vbCritical, "Send Error Message"
    Me.CMDGenTaxNo.SetFocus
    Exit Sub
End If

If Me.LBLArCode.Caption = "" Then
    MsgBox "กรุณา เลือกลูกค้า", vbCritical, "Send Error Message"
    Me.TXTARCode.SetFocus
    Exit Sub
End If

If Me.LBLArCode.Caption = "เงินสด" Then
    MsgBox "กรุณา เลือกลูกค้าที่ไม่ใช่เงินสด", vbCritical, "Send Error Message"
    Me.TXTARCode.SetFocus
    Exit Sub
End If


If Me.TBDocNo.Text <> "" And Me.LBLTaxNo.Caption <> "" And Me.LBLArCode.Caption <> "" And Me.LBLArCode.Caption <> "เงินสด" Then
vDocNo = Trim(Me.LBLDocNo.Caption)
vDocType = Me.LBLDocType.Caption
vTaxNo = Me.LBLTaxNo.Caption
vARCode = Me.TXTARCode.Text

If vMemIsOpen = 0 Then
    
    vQuery = "exec dbo.USP_NP_CheckInputTax '" & vTaxNo & "'"
    If OpenDataBase(gConnection, vRecordset1, vQuery) <> 0 Then
        vCheckInputTax = Trim(vRecordset1.Fields("vCount").Value)
    End If
    vRecordset1.Close
    
    If vCheckInputTax > 0 Then
        MsgBox "เอกสารใบกำกับภาษีนี้ มีอยู่แล้วกรุณากดปุ่มกำหนดเลขที่ใบกำกับภาษีอย่างเต็มใหม่", vbCritical, "Send Error Message"
        Me.CMDGenTaxNo.SetFocus
        Exit Sub
    End If
    
    vQuery = "exec dbo.USP_NP_InsertOutPutTax '" & vDocNo & "','" & vTaxNo & "','" & vARCode & "' "
    gConnection.Execute vQuery
End If

vMemIsOpen = 0

If vDocType = 1 Then
vRepType = "INV"
vRepID = 575

 vDueDate = ""
 vCheck = 1
 vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     With Me.Crystal101
         .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
         .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
         .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
         .Formulas(0) = "CreditCondition='" & vDueDate & "' "
         .Destination = crptToWindow
         .WindowState = crptMaximized
         .Action = 1
     End With
 End If
 vRecordset.Close
 End If
 
 Me.TBDocNo.Text = ""
 Me.LBLDocNo.Caption = ""
 Me.LBLDocType.Caption = ""
 Me.LBLArCode.Caption = ""
 Me.ListViewItem.ListItems.Clear
 Me.TBDocNo.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDSearchAR_Click()
If Me.TBDocNo.Text <> "" And Me.ListViewItem.ListItems.Count > 0 Then
    Me.PICAr.Visible = True
    Me.TextSearchARCode.SetFocus
Else
    MsgBox "กรุณา กรอกเลขที่เอกสารอย่างย่อ", vbCritical, "Send Error Message"
    Me.TBDocNo.SetFocus
End If
End Sub

Private Sub CMDSearchArOK_Click()
Dim vQuery As String
Dim vSearch As String
Dim vListAR As ListItem
Dim vRecordset As New ADODB.Recordset
Dim n As Integer

On Error GoTo ErrDescription

If TextSearchARCode.Text <> "" Then
  vSearch = TextSearchARCode.Text
  vQuery = "exec dbo.USP_MP_SearchArCode 1,'" & vSearch & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  Me.ListViewSearchAr.ListItems.Clear
  vRecordset.MoveFirst
  While Not vRecordset.EOF
  n = n + 1
  Set vListAR = Me.ListViewSearchAr.ListItems.Add(, , n)
  vListAR.SubItems(1) = Trim(vRecordset.Fields("code").Value)
  vListAR.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
  vListAR.SubItems(3) = Trim(vRecordset.Fields("billaddress").Value)
  vListAR.SubItems(4) = Trim(vRecordset.Fields("telephone").Value)
  vListAR.SubItems(5) = Trim(vRecordset.Fields("fax").Value)
  vRecordset.MoveNext
  Wend
  Me.ListViewSearchAr.SetFocus
  Else
  Me.ListViewSearchAr.ListItems.Clear
  Me.TextSearchARCode.SetFocus
  End If
  vRecordset.Close
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDSearchDocNo_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vDocNo As String
Dim i As Integer
Dim vListItem As ListItem

On Error GoTo ErrDescription

If Me.TBDocNo.Text <> "" Then
   vDocNo = Me.TBDocNo.Text
   vQuery = "exec dbo.usp_np_SearchTaxNoPrintInvoice '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.LBLDocNo.Caption = vRecordset.Fields("docno").Value
      Me.LBLDocType.Caption = vRecordset.Fields("doctype").Value
      Me.TXTARCode.Text = vRecordset.Fields("arcode").Value
      Me.LBLArCode.Caption = vRecordset.Fields("arname").Value
      vRecordset.MoveFirst
      i = 1
      Me.ListViewItem.ListItems.Clear
      While Not vRecordset.EOF
      Set vListItem = Me.ListViewItem.ListItems.Add(, , i)
      vListItem.SubItems(1) = vRecordset.Fields("itemcode").Value & "/" & vRecordset.Fields("itemname").Value
      vListItem.SubItems(2) = Format(vRecordset.Fields("qty").Value, "##,##0.00")
      vListItem.SubItems(3) = vRecordset.Fields("unitcode").Value
      vListItem.SubItems(4) = Format(vRecordset.Fields("price").Value, "##,##0.00")
      vListItem.SubItems(5) = Format(vRecordset.Fields("discountamount").Value, "##,##0.00")
      vListItem.SubItems(6) = Format(vRecordset.Fields("amount").Value, "##,##0.00")
      i = i + 1
      vRecordset.MoveNext
      Wend
      vRecordset.Close
Else
        vQuery = "exec dbo.usp_tf_invoicefrompos_S01_New '" & vDocNo & "'"
        If OpenDataBase(gConnection, vRecordset1, vQuery) <> 0 Then
                MsgBox (vRecordset1.Fields("errordesc").Value)
        End If
        vRecordset1.Close
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListViewSearchAR_DblClick()
Dim vIndex As Integer

On Error GoTo ErrDescription

If Me.ListViewSearchAr.ListItems.Count > 0 Then
  vIndex = Me.ListViewSearchAr.SelectedItem.Index
  Me.TXTARCode.Text = Me.ListViewSearchAr.ListItems(vIndex).SubItems(1)
  Me.LBLArCode.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(2)
  Me.LBLAddress.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(3)
  Me.LBLTel.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(4)
  Me.LBLFax.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(5)
  Me.PICAr.Visible = False
Else
  Me.TXTARCode.Text = ""
  Me.LBLArCode.Caption = ""
  Me.LBLAddress.Caption = ""
  Me.LBLTel.Caption = ""
  Me.LBLFax.Caption = ""
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListViewSearchAR_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Me.ListViewSearchAr.ListItems.Count > 0 Then
      vIndex = Me.ListViewSearchAr.SelectedItem.Index
      Me.TXTARCode.Text = Me.ListViewSearchAr.ListItems(vIndex).SubItems(1)
      Me.LBLArCode.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(2)
      Me.LBLAddress.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(3)
      Me.LBLTel.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(4)
      Me.LBLFax.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(5)
      Me.PICAr.Visible = False
    Else
      Me.TXTARCode.Text = ""
      Me.LBLArCode.Caption = ""
      Me.LBLAddress.Caption = ""
      Me.LBLTel.Caption = ""
      Me.LBLFax.Caption = ""
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub TBDocNo_Change()
Dim vErrDesc As String
Dim vIsError As Integer
Dim vDocNo As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vRecordset4 As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem

On Error GoTo ErrDescription
     
If Me.TBDocNo.Text = "" Then
    Me.TBDocNo.Text = ""
    Me.LBLDocNo.Caption = ""
    Me.LBLDocType.Caption = ""
    Me.LBLArCode.Caption = ""
    Me.TXTARCode.Text = ""
    Me.LBLTaxNo.Caption = ""
    Me.CMDGenTaxNo.Enabled = True
    Me.ListViewItem.ListItems.Clear
    Me.TBDocNo.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub TBDocNo_KeyPress(KeyAscii As Integer)
Dim vErrDesc As String
Dim vIsError As Integer
Dim vDocNo As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vRecordset4 As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
If Me.TBDocNo.Text <> "" Then

     vDocNo = Me.TBDocNo.Text
    
    vQuery = "exec dbo.usp_np_SearchTaxNoPrintInvoice '" & vDocNo & "' "
     If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       Me.LBLDocNo.Caption = vRecordset.Fields("docno").Value
       Me.LBLDocType.Caption = vRecordset.Fields("doctype").Value
       Me.TXTARCode.Text = vRecordset.Fields("arcode").Value
       Me.LBLArCode.Caption = vRecordset.Fields("arname").Value
       
       If vRecordset.Fields("taxno").Value <> "" Then
            vMemIsOpen = 1
            Me.LBLTaxNo.Caption = vRecordset.Fields("taxno").Value
       Else
            vMemIsOpen = 0
       End If
       
       If Me.LBLTaxNo.Caption <> "" Then
         Me.CMDGenTaxNo.Enabled = False
       Else
         Me.CMDGenTaxNo.Enabled = True
       End If
       
       vRecordset.MoveFirst
       i = 1
       Me.ListViewItem.ListItems.Clear
       While Not vRecordset.EOF
       Set vListItem = Me.ListViewItem.ListItems.Add(, , i)
       vListItem.SubItems(1) = vRecordset.Fields("itemcode").Value & "/" & vRecordset.Fields("itemname").Value
       vListItem.SubItems(2) = Format(vRecordset.Fields("qty").Value, "##,##0.00")
       vListItem.SubItems(3) = vRecordset.Fields("unitcode").Value
       vListItem.SubItems(4) = Format(vRecordset.Fields("price").Value, "##,##0.00")
       vListItem.SubItems(5) = Format(vRecordset.Fields("discountamount").Value, "##,##0.00")
       vListItem.SubItems(6) = Format(vRecordset.Fields("amount").Value, "##,##0.00")
       i = i + 1
       vRecordset.MoveNext
       Wend
       vRecordset.Close
     Else
     
        Me.LBLDocNo.Caption = ""
        Me.LBLDocType.Caption = ""
        Me.TXTARCode.Text = ""
        Me.LBLArCode.Caption = ""
        Me.LBLTaxNo.Caption = ""
        Me.CMDGenTaxNo.Enabled = True
        Me.ListViewItem.ListItems.Clear

     
         vQuery = "exec dbo.usp_tf_checkinvoice '" & vDocNo & "'"
         If OpenDataBase(gConnection, vRecordset4, vQuery) <> 0 Then
                 vErrDesc = vRecordset4.Fields("errordesc").Value
                 vIsError = vRecordset4.Fields("iserror").Value
         End If
         vRecordset4.Close
         
         If vIsError = 0 Then
             vQuery = "exec dbo.usp_tf_invoicefrompos_S01_ForTaxPrint '" & vDocNo & "'"
             gConnection.Execute (vQuery)
             
             MsgBox "เอกสารโอนมาจาก POS เรียบร้อยแล้ว กรุณากดดูข้อมูลอีกรอบ", vbCritical, "Send Information Message"
             
         Else
             MsgBox vErrDesc
         End If
    End If
End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub TextSearchARCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CMDSearchArOK_Click
End If
End Sub

Private Sub TXTArCode_Change()
Dim vQuery As String
Dim vSearch As String
Dim vRecordset2 As New ADODB.Recordset


On Error GoTo ErrDescription

If TXTARCode.Text <> "" Then
  vSearch = TXTARCode.Text
  vQuery = "exec dbo.USP_MP_SearchArCode 2,'" & vSearch & "' "
  If OpenDataBase(gConnection, vRecordset2, vQuery) <> 0 Then
        vRecordset2.MoveFirst
        While Not vRecordset2.EOF
        Me.LBLArCode.Caption = Trim(vRecordset2.Fields("arname").Value)
        Me.LBLAddress.Caption = Trim(vRecordset2.Fields("billaddress").Value)
        Me.LBLTel.Caption = Trim(vRecordset2.Fields("telephone").Value)
        Me.LBLFax.Caption = Trim(vRecordset2.Fields("fax").Value)
        vRecordset2.MoveNext
        Wend
  Else
        Me.LBLArCode.Caption = ""
        Me.LBLAddress.Caption = ""
        Me.LBLTel.Caption = ""
        Me.LBLFax.Caption = ""
  End If
  vRecordset2.Close
Else
    Me.LBLArCode.Caption = ""
    Me.LBLAddress.Caption = ""
    Me.LBLTel.Caption = ""
    Me.LBLFax.Caption = ""
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub
