VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FormStockCardBC 
   Caption         =   "รายงาน สต๊อกการ์ด BCAccount"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormStockCardBC.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal103 
      Left            =   2340
      Top             =   7830
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
   Begin Crystal.CrystalReport Crystal102 
      Left            =   1440
      Top             =   7830
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
   Begin VB.PictureBox PICSearchItem 
      Height          =   9060
      Left            =   -45
      Picture         =   "FormStockCardBC.frx":72FB
      ScaleHeight     =   9000
      ScaleWidth      =   12015
      TabIndex        =   19
      Top             =   -45
      Visible         =   0   'False
      Width           =   12075
      Begin VB.CommandButton CMDClickSearch 
         Height          =   285
         Left            =   5805
         Picture         =   "FormStockCardBC.frx":E5F6
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2070
         Width           =   375
      End
      Begin VB.CommandButton CMDExit 
         Caption         =   "ปิด"
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
         Left            =   9990
         TabIndex        =   25
         Top             =   7110
         Width           =   960
      End
      Begin VB.CommandButton CMDSelect 
         Caption         =   "เลือก"
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
         Left            =   8820
         TabIndex        =   24
         Top             =   7110
         Width           =   960
      End
      Begin MSComctlLib.ListView ListViewItemList 
         Height          =   4200
         Left            =   990
         TabIndex        =   23
         Top             =   2610
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   7408
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   11818
         EndProperty
      End
      Begin VB.TextBox TextSearchItemCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Top             =   2070
         Width           =   3975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "คำที่ค้นหา :"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   990
         TabIndex        =   21
         Top             =   2070
         Width           =   915
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "ค้นหารหัสสินค้า "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   990
         TabIndex        =   20
         Top             =   1260
         Width           =   1455
      End
   End
   Begin VB.CommandButton CMDSearchWHCode2 
      Height          =   285
      Left            =   5355
      Picture         =   "FormStockCardBC.frx":E9C3
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3060
      Width           =   375
   End
   Begin VB.CommandButton CMDSearchItemCode 
      Height          =   285
      Left            =   5940
      Picture         =   "FormStockCardBC.frx":ED90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1890
      Width           =   330
   End
   Begin VB.ListBox ListWHCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   3915
      TabIndex        =   18
      Top             =   3375
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton CMDSearchWHCode1 
      Height          =   285
      Left            =   5355
      Picture         =   "FormStockCardBC.frx":F15D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2700
      Width           =   375
   End
   Begin VB.ComboBox CMBCompany 
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
      Left            =   3915
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1350
      Width           =   1410
   End
   Begin VB.CommandButton CMDPrint 
      Caption         =   "พิมพ์ รายงาน"
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
      Left            =   5355
      TabIndex        =   9
      Top             =   4365
      Width           =   1320
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   585
      Top             =   7830
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
   Begin VB.TextBox TextStopWHCode 
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
      Left            =   3915
      TabIndex        =   5
      Text            =   "010"
      Top             =   3060
      Width           =   1410
   End
   Begin VB.TextBox TextStartWHCode 
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
      Left            =   3915
      TabIndex        =   3
      Text            =   "010"
      Top             =   2700
      Width           =   1410
   End
   Begin MSComCtl2.DTPicker DTPStopDate 
      Height          =   285
      Left            =   3915
      TabIndex        =   8
      Top             =   3870
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   58589185
      CurrentDate     =   39219
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   285
      Left            =   3915
      TabIndex        =   7
      Top             =   3510
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   58589185
      CurrentDate     =   39219
   End
   Begin VB.TextBox TextItemCode 
      Height          =   285
      Left            =   3915
      TabIndex        =   1
      Top             =   1890
      Width           =   1995
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สต๊อคการ์ด :"
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
      Left            =   2115
      TabIndex        =   17
      Top             =   1350
      Width           =   1725
   End
   Begin VB.Label LBLItemName 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3915
      TabIndex        =   16
      Top             =   2250
      Width           =   6090
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่ :"
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
      Height          =   240
      Left            =   1575
      TabIndex        =   15
      Top             =   3870
      Width           =   2265
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่ :"
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
      Height          =   240
      Left            =   2745
      TabIndex        =   14
      Top             =   3510
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงคลัง :"
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
      Height          =   240
      Left            =   2790
      TabIndex        =   13
      Top             =   3015
      Width           =   1050
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากคลัง :"
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
      Height          =   240
      Left            =   2790
      TabIndex        =   12
      Top             =   2655
      Width           =   1050
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อสินค้า :"
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
      Height          =   240
      Left            =   2835
      TabIndex        =   11
      Top             =   2250
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสสินค้า :"
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
      Height          =   240
      Left            =   2880
      TabIndex        =   10
      Top             =   1890
      Width           =   960
   End
End
Attribute VB_Name = "FormStockCardBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String
Dim vClick As Integer

Private Sub CMDClickSearch_Click()
Dim vSearch As String
Dim vRecordset As New ADODB.Recordset
Dim i As Double
Dim vListItem As ListItem

On Error GoTo ErrDescription

 If Me.TextSearchItemCode.Text <> "" Then
  vSearch = Me.TextSearchItemCode.Text
  vQuery = "exec dbo.USP_MB_SearchItem '" & vSearch & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   Me.ListViewItemList.ListItems.Clear
   i = 1
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   Set vListItem = Me.ListViewItemList.ListItems.Add(, , i)
   vListItem.SubItems(1) = vRecordset.Fields("itemcode").Value
   vListItem.SubItems(2) = vRecordset.Fields("itemname").Value
   i = i + 1
   vRecordset.MoveNext
   Wend
   Me.ListViewItemList.SetFocus
  End If
 End If
 
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMDExit_Click()
Me.PICSearchItem.Visible = False
End Sub

Private Sub CMDPrint_Click()
Dim vRecordset As New ADODB.Recordset
Dim vItemCode As String
Dim vStartWHCode As String
Dim vStopWHCode As String
Dim vStartDate As Date
Dim vStopDate As Date
Dim rServer As String
Dim rDataBase As String
Dim vReportName As String
Dim vCompanyName As String

On Error GoTo ErrDescription

If Me.LBLItemName.Caption <> "" And Me.TextStartWHCode.Text <> "" And Me.TextStopWHCode.Text <> "" Then
vCompanyName = UCase(Me.CMBCompany.Text)
Select Case UCase(vCompanyName)
Case "S02"
 rServer = "S02DB"
 rDataBase = "BCNP"
Case "NP48"
 rServer = "NEBULA"
 rDataBase = "BCNP2005_BFClose"
Case "NP47"
 rServer = "NEBULA"
 rDataBase = "BCNP2004a"
Case Else
rServer = ""
rDataBase = ""
Exit Sub
End Select

vItemCode = Me.TextItemCode.Text
vStartWHCode = Me.TextStartWHCode.Text
vStopWHCode = Me.TextStopWHCode.Text
vStartDate = Me.DTPStartDate.Value
vStopDate = Me.DTPStopDate.Value

Dim vRepID As Integer
Dim vRepType As String

vRepID = 349
vRepType = "IV"

'vQuery = "select reportname from bcreportname where repid = 349  and reptype = 'IV' "
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
End If
vRecordset.Close

With Crystal101
.ReportFileName = Trim(vReportName)
.Connect = "dsn =" & rServer & ";dsq = " & rDataBase & ";uid=" & vUserID & ";pwd =" & vPassword & " "
Call ReportSetLocation(Crystal101, rDataBase)
.ParameterFields(0) = "@Itemcode;" & vItemCode & ";true"
.ParameterFields(1) = "@FromWH;" & vStartWHCode & ";true"
.ParameterFields(2) = "@ToWH;" & vStopWHCode & ";true"
.ParameterFields(3) = "@AtDate;" & vStartDate & ";true"
.ParameterFields(4) = "@EndDate;" & vStopDate & ";true"
.Formulas(0) = "Company='" & vCompanyName & "' "
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
Else
 MsgBox "กรุณาตรวจสอบ รหัสสินค้า ว่ามีข้อมูลอยู่ในระบบหรือไม่", vbCritical, "Send Error Message"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMDSearchItemCode_Click()
Me.PICSearchItem.Visible = True
Me.TextSearchItemCode.SetFocus

End Sub

Private Sub CMDSearchWHCode1_Click()
Me.ListWHCode.Visible = True
vClick = 1
End Sub

Private Sub CMDSearchWHCode2_Click()
Me.ListWHCode.Visible = True
vClick = 2
End Sub

Private Sub CMDSelect_Click()
Dim vIndex As Integer

On Error GoTo ErrDescription

If Me.ListViewItemList.ListItems.Count > 0 Then
 vIndex = Me.ListViewItemList.SelectedItem.Index
 Me.TextItemCode.Text = Me.ListViewItemList.ListItems(vIndex).SubItems(1)
 Me.LBLItemName.Caption = Me.ListViewItemList.ListItems(vIndex).SubItems(2)
 Me.PICSearchItem.Visible = False
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Me.DTPStartDate.Value = Now
Me.DTPStopDate.Value = Now
Call GetConnection
Call GetWareHouse
End Sub

Public Sub GetConnection()
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

vQuery = "select  * from npmaster.dbo.TB_NP_BCDataConnection where  servername = 'S02DB' order by companyname"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 vRecordset.MoveFirst
 While Not vRecordset.EOF
 Me.CMBCompany.AddItem vRecordset.Fields("companyname").Value
 vRecordset.MoveNext
 Wend
 Me.CMBCompany.Text = Me.CMBCompany.List(1)
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Public Sub GetWareHouse()
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

vQuery = "select  distinct code from dbo.BCWarehouse order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 vRecordset.MoveFirst
 While Not vRecordset.EOF
 Me.ListWHCode.AddItem vRecordset.Fields("code").Value
 vRecordset.MoveNext
 Wend
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub


Private Sub ListViewItemList_DblClick()
Dim vIndex As Integer

On Error GoTo ErrDescription

If Me.ListViewItemList.ListItems.Count > 0 Then
 vIndex = Me.ListViewItemList.SelectedItem.Index
 Me.TextItemCode.Text = Me.ListViewItemList.ListItems(vIndex).SubItems(1)
 Me.LBLItemName.Caption = Me.ListViewItemList.ListItems(vIndex).SubItems(2)
 Me.PICSearchItem.Visible = False
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub ListViewItemList_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
 If Me.ListViewItemList.ListItems.Count > 0 Then
  vIndex = Me.ListViewItemList.SelectedItem.Index
  Me.TextItemCode.Text = Me.ListViewItemList.ListItems(vIndex).SubItems(1)
  Me.LBLItemName.Caption = Me.ListViewItemList.ListItems(vIndex).SubItems(2)
  Me.PICSearchItem.Visible = False
 End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub ListWHCode_Click()
Dim vIndex As Integer

On Error GoTo ErrDescription

If Me.ListWHCode.ListCount > 0 Then
  If vClick = 1 Then
   Me.TextStartWHCode.Text = Me.ListWHCode.Text
  ElseIf vClick = 2 Then
   Me.TextStopWHCode.Text = Me.ListWHCode.Text
  End If
  Me.ListWHCode.Visible = False
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub TextItemCode_Change()
Dim vItemCode As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If Me.TextItemCode.Text <> "" Then
 vItemCode = Me.TextItemCode.Text
 vQuery = "exec dbo.USP_NP_CheckItemCode '" & vItemCode & "' "
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   Me.TextItemCode.Text = vRecordset.Fields("code").Value
   Me.LBLItemName.Caption = vRecordset.Fields("name1").Value
 Else
  Me.LBLItemName.Caption = ""
 End If
 vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub TextSearchItemCode_KeyPress(KeyAscii As Integer)
Dim vSearch As String
Dim vRecordset As New ADODB.Recordset
Dim i As Double
Dim vListItem As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
 If Me.TextSearchItemCode.Text <> "" Then
  vSearch = Me.TextSearchItemCode.Text
  vQuery = "exec dbo.USP_MB_SearchItem '" & vSearch & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   Me.ListViewItemList.ListItems.Clear
   i = 1
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   Set vListItem = Me.ListViewItemList.ListItems.Add(, , i)
   vListItem.SubItems(1) = vRecordset.Fields("itemcode").Value
   vListItem.SubItems(2) = vRecordset.Fields("itemname").Value
   i = i + 1
   vRecordset.MoveNext
   Wend
   Me.ListViewItemList.SetFocus
  End If
 End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

