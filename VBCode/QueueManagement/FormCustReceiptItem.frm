VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormCustReceiptItem 
   Caption         =   "ตรวจสอบการจ่ายสินค้าของ Checker"
   ClientHeight    =   7575
   ClientLeft      =   5130
   ClientTop       =   1800
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic101 
      Height          =   1995
      Left            =   1890
      ScaleHeight     =   1935
      ScaleWidth      =   8145
      TabIndex        =   13
      Top             =   4410
      Visible         =   0   'False
      Width           =   8205
      Begin VB.CommandButton CMDExit 
         Caption         =   "ออก"
         Height          =   420
         Left            =   7065
         TabIndex        =   23
         Top             =   1305
         Width           =   825
      End
      Begin VB.CommandButton CMDOK 
         Caption         =   "ตกลง"
         Height          =   420
         Left            =   6120
         TabIndex        =   22
         Top             =   1305
         Width           =   825
      End
      Begin VB.TextBox TextQTY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3645
         TabIndex        =   19
         Top             =   675
         Width           =   1770
      End
      Begin VB.Label LBLQTY 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   945
         TabIndex        =   25
         Top             =   675
         Width           =   1905
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "ขาย :"
         Height          =   285
         Left            =   135
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LBLUnitCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6435
         TabIndex        =   21
         Top             =   675
         Width           =   1500
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "หน่วยนับ :"
         Height          =   375
         Left            =   5580
         TabIndex        =   20
         Top             =   675
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "จ่ายไป :"
         Height          =   285
         Left            =   2835
         TabIndex        =   18
         Top             =   675
         Width           =   735
      End
      Begin VB.Label LBLItemName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4410
         TabIndex        =   17
         Top             =   90
         Width           =   3525
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "ชื่อสืนค้า :"
         Height          =   330
         Left            =   3600
         TabIndex        =   16
         Top             =   90
         Width           =   735
      End
      Begin VB.Label LBLItemCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   945
         TabIndex        =   15
         Top             =   90
         Width           =   1905
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "รหัสสินค้า :"
         Height          =   330
         Left            =   45
         TabIndex        =   14
         Top             =   90
         Width           =   825
      End
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "ออก"
      Height          =   420
      Left            =   9315
      TabIndex        =   3
      Top             =   4410
      Width           =   780
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "ตกลง"
      Height          =   420
      Left            =   8370
      TabIndex        =   2
      Top             =   4410
      Width           =   780
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2760
      Left            =   1890
      TabIndex        =   1
      Top             =   1485
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4868
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
      NumItems        =   7
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
         Text            =   "ชื่อสินค้า"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "จำนวน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "จำนวนจ่าย"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "หน่วย"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ชั้นเก็บ"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.ComboBox CMBChecker 
      Height          =   315
      Left            =   2835
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   855
      Width           =   4515
   End
   Begin VB.Label LBLCount 
      Height          =   285
      Left            =   8505
      TabIndex        =   12
      Top             =   855
      Width           =   1635
   End
   Begin VB.Label Label5 
      Caption         =   "ครั้งที่ :"
      Height          =   285
      Left            =   7605
      TabIndex        =   11
      Top             =   855
      Width           =   825
   End
   Begin VB.Label LBLWHCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8505
      TabIndex        =   10
      Top             =   270
      Width           =   1590
   End
   Begin VB.Label LBLInvoice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5985
      TabIndex        =   9
      Top             =   270
      Width           =   1680
   End
   Begin VB.Label LBLRefNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3105
      TabIndex        =   8
      Top             =   270
      Width           =   1725
   End
   Begin VB.Label Label4 
      Caption         =   "Checker :"
      Height          =   240
      Left            =   1890
      TabIndex        =   7
      Top             =   855
      Width           =   870
   End
   Begin VB.Label Label3 
      Caption         =   "คลัง :"
      Height          =   285
      Left            =   7875
      TabIndex        =   6
      Top             =   270
      Width           =   510
   End
   Begin VB.Label Label2 
      Caption         =   "เลขที่บิล :"
      Height          =   285
      Left            =   5085
      TabIndex        =   5
      Top             =   270
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "เลขที่ใบจ่าย :"
      Height          =   285
      Left            =   1935
      TabIndex        =   4
      Top             =   270
      Width           =   1050
   End
End
Attribute VB_Name = "FormCustReceiptItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String
Dim vIndex As Integer

Private Sub CMD101_Click()
Dim vDocno As String
Dim vDocDate As Date
Dim vItemCode As String
Dim vItemName As String
Dim vQTY As Currency
Dim vReceiptQTY As Currency
Dim vUnitCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vPrintCount As Integer
Dim vLineNumber As Integer
Dim i As Integer
Dim vRecordset As New ADODB.Recordset
Dim vSaleOrderNo As String
Dim vShelfGroup As String
Dim vInvoiceNo As String
Dim vChecker As String

If ListView101.ListItems.Count > 0 And CMBChecker.Text <> "" Then
Call FrmPayItemCust.StopTime
vDocno = Trim(LBLRefNo.Caption)
vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
vWHCode = Trim(LBLWHCode.Caption)
vPrintCount = LBLCount.Caption
vInvoiceNo = Trim(LBLInvoice.Caption)
vChecker = Trim(CMBChecker.Text)
For i = 1 To ListView101.ListItems.Count
  vItemCode = ListView101.ListItems.Item(i).SubItems(1)
  vItemName = ListView101.ListItems.Item(i).SubItems(2)
  vQTY = ListView101.ListItems.Item(i).SubItems(3)
  vReceiptQTY = ListView101.ListItems.Item(i).SubItems(4)
  vUnitCode = ListView101.ListItems.Item(i).SubItems(5)
  vShelfCode = ListView101.ListItems.Item(i).SubItems(6)
  vLineNumber = i - 1
  vQuery = "exec dbo.USP_QUE_InsertLineItemReceipt '" & vDocno & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "'," & vQTY & "," & vReceiptQTY & ",'" & vUnitCode & "','" & vWHCode & "','" & vShelfCode & "'," & vPrintCount & "," & vLineNumber & " "
  vConnection.Execute vQuery
Next i
LBLRefNo.Caption = ""
LBLInvoice.Caption = ""
LBLWHCode.Caption = ""
LBLCount.Caption = ""
ListView101.ListItems.Clear
Unload FormCustReceiptItem
vQuery = "exec dbo.USP_QUE_CheckShelfGroup '" & vInvoiceNo & "','" & vWHCode & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  vSaleOrderNo = Trim(vRecordset.Fields("sorefno").Value)
  vShelfGroup = Trim(vRecordset.Fields("shelfgroup").Value)
End If
vRecordset.Close

vQuery = "exec dbo.USP_QUE_UpdateIsReceivedItem '" & vSaleOrderNo & "','" & vWHCode & "','" & vShelfGroup & "' "
vConnection.Execute vQuery
vQuery = "exec dbo.USP_NP_UpdateStatusCustItemReceipt 1,'" & vDocno & "'," & vPrintCount & ",'" & vChecker & "' "
vConnection.Execute vQuery
Call FrmPayItemCust.SearchCustItemReceiptChecking

Else
  MsgBox "ต้องมีรายการสินค้าในตารางและเลือกใส่ชื่อ Checker ให้เรียบร้อย", vbCritical, "Send Error"
End If
FrmPayItemCust.TextChecker.Text = ""
Call FrmPayItemCust.StartTime
End Sub

Private Sub CMD102_Click()
Call FrmPayItemCust.StartTime
Unload FormCustReceiptItem
End Sub

Private Sub CMDOK_Click()
Dim vPickQTY As String

On Error Resume Next

If LBLItemName.Caption <> "" And TextQTY.Text <> "" Then
  vPickQTY = CCur(TextQTY.Text)
  ListView101.ListItems.Item(vIndex).SubItems(4) = Format(vPickQTY, "##,##0.00")
  LBLItemCode.Caption = ""
  LBLItemName.Caption = ""
  LBLQTY.Caption = ""
  LBLUnitCode.Caption = ""
  TextQTY.Text = ""
  ListView101.SetFocus
  Pic101.Visible = False
Else
  MsgBox "กรุณากรอกจำนวนที่หยิบได้ด้วย", vbCritical, "Send Error"
End If

End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

Call FrmPayItemCust.StopTime
vQuery = "select * from  dbo.vw_np_CheckerName "
If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst
CMBChecker.Clear
While Not vRecordset.EOF
  CMBChecker.AddItem Trim(vRecordset.Fields("personname").Value)
vRecordset.MoveNext
Wend
End If
vRecordset.Close
End Sub

Private Sub ListView101_DblClick()

If ListView101.ListItems.Count > 0 Then
  Pic101.Visible = True
  vIndex = ListView101.SelectedItem.Index
  LBLItemCode.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(1))
  LBLItemName.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(2))
  LBLQTY.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
  LBLUnitCode.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(5))
  TextQTY.SetFocus
End If

End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If ListView101.ListItems.Count > 0 Then
  Pic101.Visible = True
  vIndex = ListView101.SelectedItem.Index
  LBLItemCode.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(1))
  LBLItemName.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(2))
  LBLQTY.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
  LBLUnitCode.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(5))
  TextQTY.SetFocus
End If
End If
End Sub
