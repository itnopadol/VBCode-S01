VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormOpenItemPOS 
   Caption         =   "เปิดขายติดลบสินค้า POS"
   ClientHeight    =   11010
   ClientLeft      =   2250
   ClientTop       =   450
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormOpenItemPOS.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICSelectType 
      Height          =   11580
      Left            =   11745
      Picture         =   "FormOpenItemPOS.frx":9673
      ScaleHeight     =   11520
      ScaleWidth      =   15345
      TabIndex        =   23
      Top             =   9090
      Visible         =   0   'False
      Width           =   15405
      Begin VB.CommandButton CMDSelect 
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
         Height          =   600
         Left            =   2970
         TabIndex        =   26
         Top             =   5580
         Width           =   1455
      End
      Begin VB.OptionButton OPTBCAccount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "2.BCAccount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   2970
         TabIndex        =   25
         Top             =   3960
         Width           =   4695
      End
      Begin VB.OptionButton OPTPointOfSale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "1.BC Point Of Sale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   2970
         TabIndex        =   24
         Top             =   2790
         Value           =   -1  'True
         Width           =   4695
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกประเภท ของการเปิดติดลบ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   510
         Left            =   2970
         TabIndex        =   27
         Top             =   1710
         Width           =   6045
      End
   End
   Begin VB.Frame Frame101 
      Caption         =   "สถานะการเปิดติดลบ"
      Height          =   7440
      Left            =   1710
      TabIndex        =   15
      Top             =   1305
      Visible         =   0   'False
      Width           =   12030
      Begin VB.CommandButton CMD101 
         Caption         =   "เปิดติดลบ"
         Height          =   420
         Left            =   5895
         TabIndex        =   7
         Top             =   2160
         Width           =   1680
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "เปิดติดลบ ขายสินค้า POS :"
         Height          =   240
         Left            =   3690
         TabIndex        =   18
         Top             =   2205
         Width           =   1995
      End
      Begin VB.Label LabelStatus 
         Alignment       =   2  'Center
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
         Left            =   5895
         TabIndex        =   17
         Top             =   1620
         Width           =   1680
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "ขณะนี้โปรแกรม กำลัง :"
         Height          =   285
         Left            =   3555
         TabIndex        =   16
         Top             =   1665
         Width           =   2175
      End
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "ปิดติดลบ"
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
      Left            =   4680
      TabIndex        =   21
      Top             =   6345
      Width           =   1230
   End
   Begin VB.CommandButton CMD104 
      Caption         =   "ประมวลข้อมูลเปิดติดลบ"
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
      Left            =   2295
      TabIndex        =   6
      Top             =   6345
      Width           =   2265
   End
   Begin VB.TextBox Text104 
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
      Height          =   555
      Left            =   3510
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2925
      Width           =   9510
   End
   Begin VB.TextBox Text103 
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
      Height          =   330
      Left            =   3510
      TabIndex        =   0
      Top             =   1305
      Width           =   1860
   End
   Begin VB.CommandButton CMD103 
      Caption         =   "ลงตารางเก็บประวัติ"
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
      Left            =   2295
      TabIndex        =   4
      Top             =   3600
      Width           =   1230
   End
   Begin VB.TextBox Text102 
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
      Height          =   330
      Left            =   3510
      TabIndex        =   2
      Top             =   2520
      Width           =   1860
   End
   Begin VB.TextBox Text101 
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
      Height          =   330
      Left            =   3510
      TabIndex        =   1
      Top             =   1710
      Width           =   1860
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   1905
      Left            =   2295
      TabIndex        =   5
      Top             =   4275
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   3360
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
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อสินค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "จำนวนที่จะขาย"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "หน่วย"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เหตุผลการเปิด"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "OnHand"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   495
      Top             =   2925
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "การเปิด-ปิด ติดลบจะไม่มีผล จนกว่าจะมีการเปิดปิดหน้าขายสินค้าทั้ง POS และ BCAccount ใหม่อีกครั้ง"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1005
      Left            =   3510
      TabIndex        =   22
      Top             =   225
      Width           =   11805
   End
   Begin VB.Label LabelStatus1 
      BackColor       =   &H00C0FFFF&
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
      Height          =   375
      Left            =   7785
      TabIndex        =   20
      Top             =   1305
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สถานะการเปิดติดลบ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   1305
      Width           =   1590
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อผู้ขอเปิดติดลบ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   1350
      TabIndex        =   14
      Top             =   1305
      Width           =   2085
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เหตุผลการเปิดติดลบ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1845
      TabIndex        =   13
      Top             =   2925
      Width           =   1590
   End
   Begin VB.Label LBL102 
      BackColor       =   &H00C0FFFF&
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
      Left            =   5445
      TabIndex        =   12
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label LBL101 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3510
      TabIndex        =   11
      Top             =   2115
      Width           =   6135
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวนที่จะขาย :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   2295
      TabIndex        =   10
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อสินค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   2340
      TabIndex        =   9
      Top             =   2115
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสสินค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   2340
      TabIndex        =   8
      Top             =   1710
      Width           =   1095
   End
End
Attribute VB_Name = "FormOpenItemPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vSelectType As Integer

Private Sub CMD101_Click()
Dim vStatus As Integer
Dim vRecordset As New Recordset
Dim vQuery As String
Dim vItemCode As String
Dim vItemName As String
Dim vUnitCode As String
Dim vQTY As Currency
Dim vOnHandQTY As Currency
Dim vDescription As String
Dim i As Integer
Dim vDocdate As Date
Dim vUserRequest As String

vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))

If ListView101.ListItems.Count > 0 Then
For i = 1 To ListView101.ListItems.Count
 vItemCode = Trim(ListView101.ListItems.Item(i).Text)
 vItemName = Trim(ListView101.ListItems.Item(i).SubItems(1))
 vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(3))
 vQTY = Trim(ListView101.ListItems.Item(i).SubItems(2))
 vOnHandQTY = Trim(ListView101.ListItems.Item(i).SubItems(5))
 vDescription = Trim(ListView101.ListItems.Item(i).SubItems(4))
 vUserRequest = Trim(Text103.Text)


vQuery = "exec dbo.USP_NP_InsertOpenItemMinuteLogs '" & vDocdate & "','" & vItemCode & "','" & vItemName & "'," & vQTY & "," & vOnHandQTY & ",'" & vUnitCode & "','" & vUserRequest & "','" & vUserID & "','" & vDescription & "' "
gConnection.Execute vQuery
Next i

vStatus = 1
vQuery = "exec dbo.USP_NP_OpenMinuteItem " & vStatus & " "
gConnection.Execute vQuery

Frame101.Visible = False

MsgBox "กรุณา แจ้ง Cashier ปิดเปิดหน้าขาย POS อีกครั้งหลังเปิดติดลบ  และต้องรอปิดขายติดลบ โปรแกรม POS ทุกครั้ง"
End If


End Sub

Private Sub CMD102_Click()
Dim vRecordset As New Recordset
Dim vQuery As String
Dim vStatus As Integer
Dim i As Integer
Dim vItemCode As String
Dim vUnitCode As String
Dim vDocdate As Date

vStatus = 2
vQuery = "exec dbo.USP_NP_OpenMinuteItem " & vStatus & " "
gConnection.Execute vQuery

vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))

For i = 1 To ListView101.ListItems.Count
 vItemCode = Trim(ListView101.ListItems.Item(i).Text)
 vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(3))
vQuery = "exec dbo.USP_NP_CloseOpenItemMinuteLogs '" & vDocdate & "','" & vItemCode & "','" & vUnitCode & "','" & vUserID & "'"
gConnection.Execute vQuery
Next i
Call ClearScreen
MsgBox "กรุณา แจ้ง Cashier ปิดเปิดหน้าขาย POS อีกครั้งหลังปิดติดลบ ถึงจบกระบวนการ เปิดขายติดลบ โปรแกรม POS ไม่เช่นนั้นจะเป็นการเปิดขายติดลบไปตลอด"
End Sub

Private Sub CMD103_Click()
Dim vRecordset As New Recordset
Dim vQuery As String
Dim vItemName As String
Dim vUnitCode As String
Dim vItemCode As String
Dim vQTY As Currency
Dim vOnHandQTY As Currency
Dim vDescription As String
Dim vListItem As ListItem
Dim vStatusRequest  As Integer
Dim i As Integer
Dim vCheckItem As String


If Text102.Text <> "" And Text104.Text <> "" Then
vItemCode = Trim(Text101.Text)
vQuery = "exec dbo.USP_NP_CheckStockItemCode '" & vItemCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  vOnHandQTY = Int(Trim(vRecordset.Fields("qtyonhand").Value))
End If
vRecordset.Close
vQTY = Int(Text102.Text)

If vOnHandQTY - vQTY < 0 Then
vStatusRequest = 1
Else
vStatusRequest = 0
End If

If vStatusRequest = 0 Then
  MsgBox "สินค้ารหัส " & vItemCode & " นี้มียอดในระบบ " & vOnHandQTY & " จะขาย " & vQTY & " ซึ่งสามารถขายได้ ไม่จำเป็นต้องขอเปิดติดลบ โปรแกรม POS", vbCritical, "Send Error "
  Exit Sub
Else

  For i = 1 To ListView101.ListItems.Count
  vCheckItem = ListView101.ListItems.Item(i).Text
  If vItemCode = vCheckItem Then
    MsgBox "มีรายการสินค้ารหัส  " & vItemCode & " ขอเปิดติดลบอยู่ในตารางข้างล่าง รายที่ " & i & " อยู่แล้ว", vbCritical, "Send Error "
    Call ClearItem
    Exit Sub
  End If
  Next i

  vItemName = Trim(LBL101.Caption)
  vUnitCode = Trim(LBL102.Caption)
  vDescription = Trim(Text104.Text)
  Set vListItem = ListView101.ListItems.Add(, , Trim(vItemCode))
  vListItem.SubItems(1) = Trim(vItemName)
  vListItem.SubItems(2) = Format(Trim(vQTY), "##,##0.00")
  vListItem.SubItems(3) = Trim(vUnitCode)
  vListItem.SubItems(4) = Trim(vDescription)
  vListItem.SubItems(5) = Trim(vOnHandQTY)
  Call ClearItem
End If

Else
  MsgBox "กรุณากรอกข้อมูลจำนวนสินค้าที่จะขายและเหตุผลของการเปิดติดลบด้วย", vbCritical, "Send Error"
  If Text102.Text = "" Then
    Text102.SetFocus
  ElseIf Text104.Text = "" Then
    Text104.SetFocus
  End If
End If

End Sub


Private Sub CMD104_Click()

If ListView101.ListItems.Count > 0 Then
  Frame101.Visible = True
End If

End Sub

Private Sub CMDSelect_Click()
If Me.OPTPointOfSale.Value = 1 Then
vSelectType = 1
End If

If Me.OPTBCAccount.Value = True Then
vSelectType = 2
End If

Me.PICSelectType.Visible = False
Me.Text103.SetFocus
End Sub

Private Sub ListView101_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
i = ListView101.SelectedItem.Index
ListView101.ListItems.Remove (i)
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vItemCode As String
Dim vRecordset As New Recordset
Dim vQuery As String
Dim vItemName As String
Dim vUnitCode As String

If KeyAscii = 13 Then
If Text103.Text = "" Then
  MsgBox "ต้องกรอกชื่อผู้ขอเปิดติดลบก่อนเสมอ กรณีไม่กรอกไม่สามารถกรอกรหัสสินค้าเปิดติดลบได้", vbCritical, "Send Error "
  Exit Sub
Else
  vItemCode = Trim(Text101.Text)
  vQuery = "exec dbo.USP_NP_CheckItemCode '" & vItemCode & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vItemName = Trim(vRecordset.Fields("name1").Value)
    vUnitCode = Trim(vRecordset.Fields("defsaleunitcode").Value)
  Else
    MsgBox "ไม่พบข้อมูลรหัสสินค้าที่ต้องการจะเปิดติดลบ", vbCritical, "Send Error "
    Text101.Text = ""
    Exit Sub
  End If
  vRecordset.Close
  
  LBL101.Caption = vItemName
  LBL102.Caption = vUnitCode
  Text102.SetFocus
End If
End If
End Sub
Public Sub ClearItem()
Text101.Text = ""
Text102.Text = ""
Text104.Text = ""
LBL101.Caption = ""
LBL102.Caption = ""
Text101.SetFocus
End Sub
Public Sub ClearScreen()
Text101.Text = ""
Text102.Text = ""
Text104.Text = ""
Text103.Text = ""
ListView101.ListItems.Clear
LBL101.Caption = ""
LBL102.Caption = ""
Text101.SetFocus
End Sub

Private Sub Text102_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Text104.SetFocus
End If
End Sub

Private Sub Text102_LostFocus()
On Error Resume Next
Call CheckNumber(Text102.Text)
If Text102.Text <> "" Then
  If vCheckValueNumber = True Then
  Text102.Text = Format(Int(Text102.Text), "##,##0.00")
  Else
    MsgBox "กรอกข้อมูลที่เป็นตัวเลขเท่านั้น", vbCritical, "Send Error"
    Text102.SetFocus
    Exit Sub
  End If
End If
End Sub


Private Sub Text103_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Text101.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vStatus As Integer

vQuery = "exec dbo.USP_NP_SearchStatusMinuteItem"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vStatus = Trim(vRecordset.Fields("checkstock").Value)
End If
vRecordset.Close

If vStatus = 0 Then
    LabelStatus.Caption = "ปิดติดลบ"
    LabelStatus1.Caption = "ปิดติดลบ"
ElseIf vStatus = 1 Then
    LabelStatus.Caption = "เปิดติดลบ"
    LabelStatus1.Caption = "เปิดติดลบ"
End If
End Sub
