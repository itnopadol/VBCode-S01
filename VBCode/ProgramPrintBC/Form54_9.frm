VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form54_9 
   Caption         =   "พิมพ์ Label จ่าหน้าซอง"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "Form54_9.frx":0000
   ScaleHeight     =   8325
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListViewSearch 
      Height          =   3000
      Left            =   2000
      TabIndex        =   8
      Top             =   2000
      Visible         =   0   'False
      Width           =   10000
      _ExtentX        =   17648
      _ExtentY        =   5292
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัส"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อ"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ที่อยู่"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เบอร์โทร"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton CMDAdd 
      Caption         =   "เพิ่มรายการ"
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
      Left            =   8730
      TabIndex        =   7
      Top             =   1400
      Width           =   1050
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
      Height          =   690
      Left            =   9500
      TabIndex        =   6
      Top             =   6200
      Width           =   1500
   End
   Begin MSComctlLib.ListView ListviewPrint 
      Height          =   4000
      Left            =   2000
      TabIndex        =   5
      Top             =   2000
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   7064
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   2646
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
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เบอร์โทร"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.TextBox TXTSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3300
      TabIndex        =   4
      Top             =   1410
      Width           =   5000
   End
   Begin VB.CommandButton CMDSendInformation 
      Caption         =   "แจ้งข่าวสาร"
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
      Left            =   150
      TabIndex        =   2
      Top             =   2925
      Width           =   1590
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "พิมพ์ ทั่วไป"
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
      Left            =   150
      TabIndex        =   1
      Top             =   2250
      Width           =   1590
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1680
      Top             =   5760
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
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์ สมาชิก"
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
      Left            =   150
      TabIndex        =   0
      Top             =   1530
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาสินค้า :"
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
      Height          =   375
      Left            =   2000
      TabIndex        =   3
      Top             =   1485
      Width           =   2175
   End
End
Attribute VB_Name = "Form54_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vRepID = 256
vRepType = "MB"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vRepID = 294
vRepType = "ML"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDAdd_Click()
Dim i As Integer
Dim n As Integer
Dim vListItem As ListItem

On Error GoTo ErrDescription

If Me.ListViewSearch.ListItems.Count > 0 Then
For i = 1 To Me.ListViewSearch.ListItems.Count
    
    If Me.ListViewSearch.ListItems(i).Checked = True Then
    
        n = Me.ListviewPrint.ListItems.Count + 1
    
        Set vListItem = Me.ListviewPrint.ListItems.Add(, , n)
        vListItem.SubItems(1) = Me.ListViewSearch.ListItems(i).SubItems(1)
        vListItem.SubItems(2) = Me.ListViewSearch.ListItems(i).SubItems(2)
        vListItem.SubItems(3) = Me.ListViewSearch.ListItems(i).SubItems(3)
        vListItem.SubItems(4) = Me.ListViewSearch.ListItems(i).SubItems(4)
    
    End If
Next

Me.ListViewSearch.Visible = False
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If

End Sub

Private Sub CMDPrint_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

Dim i As Integer
Dim vARCode As String
Dim vARName As String
Dim vBillAddress As String
Dim vTelephone As String

On Error GoTo ErrDescription

For i = 1 To Me.ListviewPrint.ListItems.Count
     vARCode = Me.ListviewPrint.ListItems(i).SubItems(1)
     vARName = Me.ListviewPrint.ListItems(i).SubItems(2)
     vBillAddress = Me.ListviewPrint.ListItems(i).SubItems(3)
     vTelephone = Me.ListviewPrint.ListItems(i).SubItems(4)
     
    vQuery = "exec dbo.USP_NP_InsertPrintArLetterTemp  '" & vARCode & "','" & vARName & "','" & vBillAddress & "','" & vTelephone & "'"
    gConnection.Execute vQuery
Next i



vRepID = 570
vRepType = "ML"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With


vQuery = "delete npmaster.dbo.TB_NP_PrintArLetterTemp "
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDSendInformation_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vRepID = 467
vRepType = "MB"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With


vRepID = 468
vRepType = "MB"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub ListviewPrint_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer

On Error Resume Next

If KeyCode = 46 Then

    If Me.ListviewPrint.ListItems.Count > 0 Then
        vIndex = Me.ListviewPrint.SelectedItem.Index
        Me.ListviewPrint.ListItems.Remove (vIndex)
        
        Call GenLineNumber
    End If
End If
End Sub


Public Sub GenLineNumber()
Dim i As Integer
Dim n As Integer

On Error Resume Next

n = 1
If Me.ListviewPrint.ListItems.Count > 0 Then
    For i = 1 To Me.ListviewPrint.ListItems.Count
        Me.ListviewPrint.ListItems(i).Text = n
        n = n + 1
    Next i
End If
End Sub

Private Sub TXTSearch_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItem As ListItem
Dim n As Double
Dim vSearch As String

On Error Resume Next

If Me.TXTSearch.Text <> "" And KeyAscii = 13 Then

    vSearch = Me.TXTSearch.Text
    
    Me.ListViewSearch.ListItems.Clear
    
    vQuery = "exec dbo.USP_NP_SearchCustomer '" & vSearch & "'"
    If OpenDataBase(gConnection, vRecordset, vQuery) > 0 Then
       vRecordset.MoveFirst
       While Not vRecordset.EOF
       n = n + 1
       Set vListItem = Me.ListViewSearch.ListItems.Add(, , n)
       vListItem.SubItems(1) = vRecordset.Fields("code").Value
       vListItem.SubItems(2) = vRecordset.Fields("name1").Value
       vListItem.SubItems(3) = vRecordset.Fields("billaddress").Value
       vListItem.SubItems(4) = vRecordset.Fields("telephone").Value
       vRecordset.MoveNext
       Wend
    Else
    Me.ListViewSearch.Visible = False
    Me.TXTSearch.SetFocus
    End If
    vRecordset.Close
    
    Me.ListViewSearch.Visible = True
    Me.ListViewSearch.SetFocus
End If
End Sub
