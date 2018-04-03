VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
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
      Height          =   9510
      Left            =   0
      Picture         =   "Form3_20.frx":0000
      ScaleHeight     =   9480
      ScaleWidth      =   14835
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   14865
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
         Height          =   645
         Left            =   11385
         TabIndex        =   4
         Top             =   8100
         Width           =   1770
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
         Height          =   645
         Left            =   9540
         TabIndex        =   3
         Top             =   8100
         Width           =   1770
      End
      Begin MSComctlLib.ListView ListViewSearchAr 
         Height          =   4335
         Left            =   900
         TabIndex        =   2
         Top             =   3645
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   7646
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
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ที่อยู่"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "เบอร์โทร"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "แฟกส์"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.CommandButton CMDSearchArOK 
         Height          =   420
         Left            =   7605
         Picture         =   "Form3_20.frx":9673
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1935
         Width           =   375
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
         Height          =   435
         Left            =   1665
         TabIndex        =   0
         Top             =   1935
         Width           =   5910
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการลูกค้า"
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
         Height          =   465
         Left            =   945
         TabIndex        =   29
         Top             =   3285
         Width           =   2130
      End
      Begin VB.Label Label5 
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
         Height          =   465
         Left            =   630
         TabIndex        =   28
         Top             =   2025
         Width           =   1140
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ค้นหาลูกค้า เพื่อทำใบกำกับภาษีอย่างเต็ม"
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
         Height          =   375
         Left            =   495
         TabIndex        =   27
         Top             =   1305
         Width           =   4875
      End
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   540
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
      TabIndex        =   16
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
      Left            =   11295
      TabIndex        =   15
      Top             =   8145
      Width           =   1545
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   -3420
      ScaleHeight     =   1965
      ScaleWidth      =   18255
      TabIndex        =   20
      Top             =   7650
      Width           =   18285
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
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   4380
      Left            =   0
      ScaleHeight     =   4350
      ScaleWidth      =   14835
      TabIndex        =   19
      Top             =   3060
      Width           =   14865
      Begin MSComctlLib.ListView ListViewItem 
         Height          =   3840
         Left            =   585
         TabIndex        =   14
         Top             =   270
         Width           =   13740
         _ExtentX        =   24236
         _ExtentY        =   6773
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   2670
      Left            =   0
      ScaleHeight     =   2640
      ScaleWidth      =   14835
      TabIndex        =   18
      Top             =   0
      Width           =   14865
      Begin VB.CommandButton CMDSearchAr 
         Height          =   375
         Left            =   5220
         Picture         =   "Form3_20.frx":9A40
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1125
         Width           =   420
      End
      Begin VB.CommandButton CMDGenTaxNo 
         Height          =   330
         Left            =   5220
         Picture         =   "Form3_20.frx":9E0D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   675
         Width           =   420
      End
      Begin VB.TextBox TXTARCode 
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
         Left            =   2565
         TabIndex        =   8
         Top             =   1125
         Width           =   2580
      End
      Begin VB.CommandButton CMDSearchDocNo 
         Height          =   120
         Left            =   13005
         Picture         =   "Form3_20.frx":A260
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   105
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
         Left            =   2565
         TabIndex        =   5
         Top             =   180
         Width           =   2580
      End
      Begin VB.Label Label16 
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
         Height          =   330
         Left            =   8370
         TabIndex        =   34
         Top             =   2115
         Width           =   1050
      End
      Begin VB.Label Label15 
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
         Height          =   285
         Left            =   1575
         TabIndex        =   33
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label14 
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
         Height          =   240
         Left            =   1305
         TabIndex        =   32
         Top             =   2070
         Width           =   1185
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
         Left            =   2565
         TabIndex        =   12
         Top             =   2070
         Width           =   6000
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
         Left            =   9540
         TabIndex        =   13
         Top             =   2070
         Width           =   4605
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
         Left            =   2565
         TabIndex        =   11
         Top             =   1620
         Width           =   11580
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสลูกค้า :"
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
         Left            =   1440
         TabIndex        =   31
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label LBLTaxNo 
         BackColor       =   &H00C0C0FF&
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
         Left            =   2565
         TabIndex        =   6
         Top             =   675
         Width           =   2580
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบกำกับภาษีอย่างเต็ม :"
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
         Left            =   90
         TabIndex        =   30
         Top             =   675
         Width           =   2400
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
         TabIndex        =   25
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
         TabIndex        =   23
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
         Left            =   7155
         TabIndex        =   10
         Top             =   1125
         Width           =   6990
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
         Left            =   6120
         TabIndex        =   22
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label Label1 
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
         Height          =   285
         Left            =   1035
         TabIndex        =   21
         Top             =   225
         Width           =   1455
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
      Left            =   585
      TabIndex        =   24
      Top             =   2700
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
    vQuery = "exec dbo.USP_NP_GetMaxTaxNo 'S02' "
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
vRepID = 465

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
 
If vDocType = 2 Then
vRepType = "RT"
vRepID = 64

 vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     With Me.Crystal102
         .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
         .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
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

Private Sub CMDSearchAr_Click()
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
      Me.LBLArCode.Caption = vRecordset.Fields("arname").Value & "/" & vRecordset.Fields("arname").Value
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
    MsgBox "ไม่พบเลขที่เอกสารที่ต้องการพิมพ์ใบกำกับภาษีอย่างเต็ม กรุณาตรวจสอบ", vbCritical, "Send Error Message"
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListViewSearchAr_DblClick()
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

Private Sub ListViewSearchAr_KeyPress(KeyAscii As Integer)
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
             vQuery = "exec dbo.usp_tf_invoicefrompos_S02_ForTaxPrint '" & vDocNo & "'"
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

Private Sub TXTARCode_Change()
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
