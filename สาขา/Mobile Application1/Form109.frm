VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form109 
   Caption         =   "ตรวจสอบสินค้าใกล้เคียง"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form109.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CHKOnHand 
      Caption         =   "ไม่แสดง OnHand"
      Height          =   285
      Left            =   4005
      TabIndex        =   23
      Top             =   7380
      Width           =   1545
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   540
      Top             =   8055
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMDClearScreen 
      Caption         =   "ลบรายการในตารางที่เลือก"
      Height          =   675
      Left            =   7695
      TabIndex        =   21
      Top             =   7200
      Width           =   1755
   End
   Begin VB.CheckBox CBSelectAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลือกทั้งหมด"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3240
      TabIndex        =   20
      Top             =   1950
      Width           =   1485
   End
   Begin VB.OptionButton OPTNoCond 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ไม่ระบุเงื่อนไข"
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
      Height          =   315
      Left            =   1320
      TabIndex        =   17
      Top             =   1050
      Value           =   -1  'True
      Width           =   3435
   End
   Begin VB.OptionButton OPTDepart 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลือก ตามDepart"
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
      Height          =   315
      Left            =   5190
      TabIndex        =   16
      Top             =   1050
      Width           =   1935
   End
   Begin VB.OptionButton OPTCat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลือก ตามCategory"
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
      Height          =   315
      Left            =   5190
      TabIndex        =   15
      Top             =   1500
      Width           =   1935
   End
   Begin VB.OptionButton OPTType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลือก ตามชนิด"
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
      Height          =   315
      Left            =   5190
      TabIndex        =   14
      Top             =   600
      Width           =   1935
   End
   Begin VB.OptionButton OPTBrand 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลือก ตามยี่ห้อ"
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
      Height          =   315
      Left            =   5190
      TabIndex        =   13
      Top             =   150
      Width           =   1935
   End
   Begin VB.ComboBox CMBCat 
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
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1500
      Width           =   4305
   End
   Begin VB.ComboBox CMBDepart 
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
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1050
      Width           =   4305
   End
   Begin VB.ComboBox CMBType 
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
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   600
      Width           =   4305
   End
   Begin VB.ComboBox CMBBrand 
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
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   150
      Width           =   4305
   End
   Begin VB.TextBox Text101 
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
      Height          =   330
      Left            =   1590
      TabIndex        =   7
      Top             =   7350
      Width           =   2280
   End
   Begin VB.CommandButton CMDSearchItem 
      Height          =   330
      Left            =   4770
      Picture         =   "Form109.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1500
      Width           =   375
   End
   Begin VB.CommandButton CMDBasket 
      Caption         =   "เลือกสินค้าลงตาราง"
      Height          =   675
      Left            =   9750
      Picture         =   "Form109.frx":76C8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   1755
   End
   Begin VB.TextBox TextSearchItem 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1305
      TabIndex        =   0
      Top             =   1500
      Width           =   3435
   End
   Begin VB.CommandButton CMDPrint 
      Caption         =   "บันทึกและพิมพ์เอกสาร"
      Height          =   675
      Left            =   5850
      Picture         =   "Form109.frx":79E2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   1770
   End
   Begin MSComctlLib.ListView ListView102 
      Height          =   4695
      Left            =   420
      TabIndex        =   4
      Top             =   2430
      Visible         =   0   'False
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648384
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "OnHand"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หน่วยนับ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "ราคาขาย"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "หน่วยขาย"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   4695
      Left            =   405
      TabIndex        =   2
      Top             =   2430
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   8281
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "OnHand"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หน่วยนับ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "ราคาขาย"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "หน่วยขาย"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "*** เฉพาะยอดคงเหลือ คลัง S02  ได้หักยอดขายระหว่างวันของจุดขาย POS เรียบร้อย"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5190
      TabIndex        =   22
      Top             =   2040
      Width           =   6315
   End
   Begin VB.Label LBLSearch 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการ สินค้าที่ค้นหา"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   420
      TabIndex        =   19
      Top             =   2070
      Width           =   3285
   End
   Begin VB.Label LBLItem 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการ สินค้าที่จะทำการนับสต๊อก"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   420
      TabIndex        =   18
      Top             =   2070
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label2 
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
      Height          =   285
      Left            =   420
      TabIndex        =   8
      Top             =   7380
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาสินค้า :"
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   210
      TabIndex        =   6
      Top             =   1530
      Width           =   1050
   End
End
Attribute VB_Name = "Form109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String
Dim vGenDocNo As String

Private Sub CBSelectAll_Click()
Dim i  As Integer

If Me.CBSelectAll.Value = 1 Then
    If Me.ListView101.ListItems.Count > 0 Then
    For i = 1 To Me.ListView101.ListItems.Count
    Me.ListView101.ListItems(i).Checked = True
    Next
    End If
End If

If Me.CBSelectAll.Value = 0 Then
    If Me.ListView101.ListItems.Count > 0 Then
    For i = 1 To Me.ListView101.ListItems.Count
    Me.ListView101.ListItems(i).Checked = False
    Next
    End If
End If
End Sub

Private Sub CMDBasket_Click()
Dim i As Integer
Dim n As Integer
Dim vCheck As Integer
Dim vItemCode As String
Dim vListItem As ListItem

On Error GoTo ErrDescription

For i = 1 To ListView101.ListItems.Count
  If ListView101.ListItems.Item(i).Checked = True Then
    vItemCode = Trim(ListView101.ListItems.Item(i).SubItems(1))
    If ListView102.ListItems.Count > 0 Then
    vCheck = CheckItemExist(vItemCode)
    End If
  
    If vCheck = 1 Then
      MsgBox "รหัสสินค้า " & vItemCode & " ได้เลือกไว้แล้วในตารางข้างล่าง", vbCritical, "Send Information"
    End If
    
    If vCheck = 0 Then
      If ListView102.ListItems.Count > 0 Then
        n = ListView102.ListItems.Count
      Else
        n = 0
      End If
      n = n + 1
      Set vListItem = ListView102.ListItems.Add(, , Trim(n))
      vListItem.SubItems(1) = ListView101.ListItems.Item(i).SubItems(1)
      vListItem.SubItems(2) = ListView101.ListItems.Item(i).SubItems(2)
      vListItem.SubItems(3) = ListView101.ListItems.Item(i).SubItems(3)
      vListItem.SubItems(4) = ListView101.ListItems.Item(i).SubItems(4)
      vListItem.SubItems(5) = ListView101.ListItems.Item(i).SubItems(5)
      vListItem.SubItems(6) = ListView101.ListItems.Item(i).SubItems(6)
    End If
  
  End If
Next i

Me.ListView101.Visible = False
Me.ListView102.Visible = True
Me.LBLSearch.Visible = False
Me.LBLItem.Visible = True
Me.TextSearchItem.SetFocus

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDClearScreen_Click()
Dim vAnswer As Integer

vAnswer = MsgBox("คุณต้องการ ลบรายการในตารางที่เลือกไว้ใช่หรือไม่ ?", vbYesNo, "Send Question Message")

If vAnswer = 6 Then
Me.ListView102.ListItems.Clear
Me.TextSearchItem.SetFocus
Me.ListView101.Visible = True
Me.ListView102.Visible = False
End If
End Sub

Private Sub CMDPrint_Click()
Dim vRecordset As New ADODB.Recordset
Dim vYear As String
Dim vMonth As String
Dim vMaxNumber As Integer
Dim vDocDate As Date
Dim i As Integer
Dim vItemCode As String
Dim vUnitCode As String
Dim vRepID As Integer
Dim vRepType As String
Dim vReportName As String
Dim vDocNo As String

On Error GoTo ErrDescription

If ListView102.ListItems.Count > 0 And Text101.Text = "" Then
  vQuery = "exec dbo.USP_MB_GeneraterintItemDataCountStock "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vYear = Right(Trim(vRecordset.Fields("year1").Value), 2)
    vMonth = Trim(vRecordset.Fields("month1").Value)
    vMaxNumber = Trim(vRecordset.Fields("maxnumber").Value)
  End If
  vRecordset.Close
  
  If Len(vMonth) = 1 Then
    vMonth = "0" & vMonth
  End If
  
  vGenDocNo = UCase("STK" & vYear & vMonth & "-" & Format(vMaxNumber, "000"))
  
  vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
  
  
  For i = 1 To ListView102.ListItems.Count
    vItemCode = Trim(ListView102.ListItems.Item(i).SubItems(1))
    vUnitCode = Trim(ListView102.ListItems.Item(i).SubItems(4))
    vQuery = "exec dbo.USP_MB_InsertPrintItemDataCountStock '" & vGenDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vUnitCode & "','" & vUserID & "' "
    gConnection.Execute vQuery
  Next i
  
  Call PrintDataItemStock
  ListView102.ListItems.Clear
ElseIf Text101.Text <> "" Then
    If Me.CHKOnHand.Value = 0 Then
      vRepID = 329
    Else
      vRepID = 492
    End If
  vRepType = "MB"
  vDocNo = UCase(Trim(Text101.Text))
  vQuery = "select reportname from bcnp.dbo.bcreportname where  repid = " & vRepID & " and reptype = '" & vRepType & "'  "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  vReportName = Trim(vRecordset.Fields("reportname").Value)
  End If
  vRecordset.Close
  
  With Crystal101
  .ReportFileName = Trim(vReportName & ".rpt")
  .ParameterFields(0) = "@vDocno;" & vDocNo & ";true"
  .Destination = crptToWindow
  .WindowState = crptMaximized
  .Action = 1
  End With
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub SearchItemBrand()
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

Me.CMBBrand.Clear
vQuery = "exec dbo.USP_PS_BrandList"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
      Me.CMBBrand.AddItem (Trim(vRecordset.Fields("brandname").Value))
      vRecordset.MoveNext
    Wend
    End If
vRecordset.Close
End Sub

Public Sub SearchItemType()
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

Me.CMBType.Clear
vQuery = "exec dbo.USP_PS_ItemTypeList"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
      Me.CMBType.AddItem (Trim(vRecordset.Fields("itemtypename").Value))
      vRecordset.MoveNext
    Wend
    End If
vRecordset.Close
End Sub

Public Sub SearchItemCat()
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

Me.CMBCat.Clear
vQuery = "exec dbo.USP_PS_SearchItemSubCat"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
      Me.CMBCat.AddItem (Trim(vRecordset.Fields("subcatname").Value))
      vRecordset.MoveNext
    Wend
    End If
vRecordset.Close
End Sub

Public Sub SearchItemDepart()
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

Me.CMBDepart.Clear
vQuery = "exec dbo.USP_PS_SearchItemDepartment"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
      Me.CMBDepart.AddItem (Trim(vRecordset.Fields("departname").Value))
      vRecordset.MoveNext
    Wend
    End If
vRecordset.Close
End Sub

Private Sub CMDSearchItem_Click()
Dim vRecordset As New ADODB.Recordset
Dim vSearch As String
Dim i As Integer
Dim vListItem As ListItem

Dim vType As Integer
Dim vCode As String

On Error GoTo ErrDescription

ListView101.ListItems.Clear

If Me.OPTNoCond.Value = True And Me.TextSearchItem.Text <> "" Then
    vType = 0
    vCode = ""
 End If
 
 If Me.OPTBrand.Value = True And Me.CMBBrand.Text <> "" Then
    vType = 1
    vCode = Left(Me.CMBBrand.Text, InStr(Me.CMBBrand.Text, "/") - 1)
 ElseIf Me.OPTBrand.Value = True And Me.CMBBrand.Text = "" Then
    MsgBox "กรุณาเลือกเงื่อนไขที่ต้องการดู", vbCritical, "Send Error Message"
    Me.CMBBrand.SetFocus
    Exit Sub
 End If
 
  If Me.OPTType.Value = True And Me.CMBType.Text <> "" Then
    vType = 2
    vCode = Left(Me.CMBType.Text, InStr(Me.CMBType.Text, "/") - 1)
  ElseIf Me.OPTType.Value = True And Me.CMBType.Text = "" Then
    MsgBox "กรุณาเลือกเงื่อนไขที่ต้องการดู", vbCritical, "Send Error Message"
    Me.CMBType.SetFocus
    Exit Sub
 End If
 
  If Me.OPTDepart.Value = True And Me.CMBDepart.Text <> "" Then
    vType = 3
    vCode = Left(Me.CMBDepart.Text, InStr(Me.CMBDepart.Text, "/") - 1)
   ElseIf Me.OPTDepart.Value = True And Me.CMBDepart.Text = "" Then
    MsgBox "กรุณาเลือกเงื่อนไขที่ต้องการดู", vbCritical, "Send Error Message"
    Me.CMBDepart.SetFocus
    Exit Sub
 End If
 
  If Me.OPTCat.Value = True And Me.CMBCat.Text <> "" Then
    vType = 4
    vCode = Left(Me.CMBCat.Text, InStr(Me.CMBCat.Text, "/") - 1)
   ElseIf Me.OPTCat.Value = True And Me.CMBCat.Text = "" Then
    MsgBox "กรุณาเลือกเงื่อนไขที่ต้องการดู", vbCritical, "Send Error Message"
    Me.CMBDepart.SetFocus
    Exit Sub
 End If
 
 If vType = 0 And Me.TextSearchItem.Text = "" Then
 MsgBox "กรุณากรอกคำที่ค้นหา", vbCritical, "Send Error Message"
 Me.TextSearchItem.SetFocus
 Exit Sub
 End If

vSearch = Trim(TextSearchItem.Text)
  
ListView101.ListItems.Clear
i = 1
vQuery = "exec dbo.USP_MB_SearchItemCountSTK " & vType & ",'" & vCode & "','" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst
While Not vRecordset.EOF
Set vListItem = ListView101.ListItems.Add(, , Trim(i))
vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("stockqty").Value), "##,##0.00")
vListItem.SubItems(4) = Trim(vRecordset.Fields("defstkunitcode").Value)
vListItem.SubItems(5) = Format(Trim(vRecordset.Fields("saleprice1").Value), "##,##0.00")
vListItem.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
i = i + 1
vRecordset.MoveNext
Wend
End If
vRecordset.Close


Me.ListView101.Visible = True
Me.ListView102.Visible = False
Me.LBLSearch.Visible = True
Me.LBLItem.Visible = False

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub PrintDataItemStock()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType As String
Dim vReportName As String
Dim vDocNo As String

If Me.CHKOnHand.Value = 0 Then
vRepID = 329
Else
vRepID = 492
End If

vRepType = "MB"
vDocNo = vGenDocNo
vQuery = "select reportname from bcnp.dbo.bcreportname where  repid = " & vRepID & " and reptype = '" & vRepType & "'  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = Trim(vReportName & ".rpt")
.ParameterFields(0) = "@vDocno;" & vDocNo & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

End Sub

Private Sub Form_Load()
Call SearchItemBrand
Call SearchItemType
Call SearchItemDepart
Call SearchItemCat
End Sub

Private Sub OPTBrand_Click()
On Error Resume Next

If Me.OPTBrand.Value = True Then
    Me.ListView101.ListItems.Clear
    Me.CMBBrand.Enabled = True
    Me.CMBCat.Enabled = False
    Me.CMBType.Enabled = False
    Me.CMBDepart.Enabled = False
Else
    Me.CMBBrand.Enabled = False
End If
End Sub

Private Sub OPTCat_Click()
On Error Resume Next

If Me.OPTCat.Value = True Then
    Me.ListView101.ListItems.Clear
    Me.CMBCat.Enabled = True
    Me.CMBBrand.Enabled = False
    Me.CMBType.Enabled = False
    Me.CMBDepart.Enabled = False
Else
    Me.CMBCat.Enabled = False
End If
End Sub

Private Sub OPTDepart_Click()
On Error Resume Next

If Me.OPTDepart.Value = True Then
    Me.ListView101.ListItems.Clear
    Me.CMBDepart.Enabled = True
    Me.CMBBrand.Enabled = False
    Me.CMBCat.Enabled = False
    Me.CMBType.Enabled = False
Else
    Me.CMBDepart.Enabled = False
End If
End Sub

Private Sub OPTType_Click()
On Error Resume Next

If Me.OPTType.Value = True Then
    Me.ListView101.ListItems.Clear
    Me.CMBType.Enabled = True
    Me.CMBBrand.Enabled = False
    Me.CMBCat.Enabled = False
    Me.CMBDepart.Enabled = False
Else
    Me.CMBType.Enabled = False
End If
End Sub

Private Sub TextSearchItem_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vSearch As String
Dim i As Integer
Dim vListItem As ListItem


On Error GoTo ErrDescription

If KeyAscii = 13 Then

Call CMDSearchItem_Click

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Function CheckItemExist(vItem As String) As Integer
Dim i As Integer

For i = 1 To ListView102.ListItems.Count
  If vItem = ListView102.ListItems.Item(i).SubItems(1) Then
    CheckItemExist = 1
  Exit Function
  Else
    CheckItemExist = 0
  End If
Next i

End Function
