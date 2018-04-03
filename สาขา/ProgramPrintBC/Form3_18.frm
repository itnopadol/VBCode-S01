VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form3_18 
   Caption         =   "พิมพ์ใบขออนุมัติขาย"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form3_18.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1035
      Top             =   7560
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
   Begin VB.CommandButton CMDPrint 
      Caption         =   "พิมพ์ใบขออนุมัติ"
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
      Left            =   1575
      TabIndex        =   7
      Top             =   6435
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListViewSearchAR 
      Height          =   3390
      Left            =   675
      TabIndex        =   3
      Top             =   1755
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   5980
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสลูกค้า"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   12347
      EndProperty
   End
   Begin VB.CommandButton CMDSearchAR 
      Height          =   330
      Left            =   7830
      Picture         =   "Form3_18.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1170
      Width           =   330
   End
   Begin VB.TextBox TextSearchAR 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1755
      TabIndex        =   1
      ToolTipText     =   "กรอกข้อความค้นหาแล้วกด Enter หรือ กดปุ่มแว่นขยาย"
      Top             =   1170
      Width           =   6045
   End
   Begin VB.Label LBLARCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1575
      TabIndex        =   6
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label LBLARName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1575
      TabIndex        =   5
      Top             =   5850
      Width           =   8430
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   675
      TabIndex        =   4
      Top             =   5400
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาลูกค้า :"
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
      Height          =   330
      Left            =   675
      TabIndex        =   0
      Top             =   1170
      Width           =   1005
   End
End
Attribute VB_Name = "Form3_18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String

Private Sub CMDPrint_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAutoNumber As Integer
Dim vGenNumber As String
Dim vYear As String
Dim vMaxNumber As Integer
Dim vDocdate As Date
Dim vARCode As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If Me.LBLARCode.Caption <> "" And Me.LBLARName.Caption <> "" Then
vQuery = "exec dbo.USP_CD_GenerateConfirmNumber "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  vYear = Trim(vRecordset.Fields("year1").Value)
  vMaxNumber = Trim(vRecordset.Fields("maxnumber").Value)
End If
vRecordset.Close
vGenNumber = Format(vMaxNumber, "0000") & "/" & vYear

vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
vARCode = Trim(LBLARCode.Caption)
vQuery = "exec dbo.USP_CD_InsertConfirmSaleOrderRequestLogs '" & vGenNumber & "','" & vARCode & "','" & vDocdate & "','" & vUserID & "' "
gConnection.Execute vQuery

vRepID = 328
vRepType = "CD"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from dbo.bcreportname where repid = 328 and reptype = 'CD' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal101
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vArCode;" & vARCode & ";true"
        .Formulas(0) = "vDocNo ='" & vGenNumber & "' "
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close

Me.LBLARCode.Caption = ""
Me.LBLARName.Caption = ""
Me.ListViewSearchAR.ListItems.Clear
Me.TextSearchAR.Text = ""
Me.TextSearchAR.SetFocus

Else
  MsgBox "กรุณา กรอกรหัสสินค้าที่ต้องการจะพิมพ์เอกสารด้วย", vbCritical, "Send Error"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDSearchAR_Click()
Dim vSearch As String
Dim vListAR As ListItem
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If TextSearchAR.Text <> "" Then
  vSearch = TextSearchAR.Text
  vQuery = "exec dbo.USP_MP_SearchArCode 1,'" & vSearch & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  Me.ListViewSearchAR.ListItems.Clear
  vRecordset.MoveFirst
  While Not vRecordset.EOF
  Set vListAR = Me.ListViewSearchAR.ListItems.Add(, , vRecordset.Fields("code").Value)
  vListAR.SubItems(1) = Trim(vRecordset.Fields("arname").Value)
  vRecordset.MoveNext
  Wend
  Me.ListViewSearchAR.SetFocus
  Else
  Me.ListViewSearchAR.ListItems.Clear
  Me.TextSearchAR.SetFocus
  End If
  vRecordset.Close
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

If Me.ListViewSearchAR.ListItems.Count > 0 Then
  vIndex = Me.ListViewSearchAR.SelectedItem.Index
  Me.LBLARCode.Caption = Me.ListViewSearchAR.ListItems(vIndex).Text
  Me.LBLARName.Caption = Me.ListViewSearchAR.ListItems(vIndex).SubItems(1)
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListViewSearchAR_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim vIndex As Integer

On Error GoTo ErrDescription

If Me.ListViewSearchAR.ListItems.Count > 0 Then
  vIndex = Me.ListViewSearchAR.SelectedItem.Index
  Me.LBLARCode.Caption = Me.ListViewSearchAR.ListItems(vIndex).Text
  Me.LBLARName.Caption = Me.ListViewSearchAR.ListItems(vIndex).SubItems(1)
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
  If Me.ListViewSearchAR.ListItems.Count > 0 Then
    vIndex = Me.ListViewSearchAR.SelectedItem.Index
    Me.LBLARCode.Caption = Me.ListViewSearchAR.ListItems(vIndex).Text
    Me.LBLARName.Caption = Me.ListViewSearchAR.ListItems(vIndex).SubItems(1)
  End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub TextSearchAR_KeyPress(KeyAscii As Integer)
Dim vSearch As String
Dim vListAR As ListItem
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If KeyAscii = 13 Then
  If TextSearchAR.Text <> "" Then
    vSearch = TextSearchAR.Text
    vQuery = "exec dbo.USP_MP_SearchArCode 1,'" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.ListViewSearchAR.ListItems.Clear
    Me.LBLARCode.Caption = ""
    Me.LBLARName.Caption = ""
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set vListAR = Me.ListViewSearchAR.ListItems.Add(, , vRecordset.Fields("code").Value)
    vListAR.SubItems(1) = Trim(vRecordset.Fields("arname").Value)
    vRecordset.MoveNext
    Wend
    Me.ListViewSearchAR.SetFocus
    Else
    Me.ListViewSearchAR.ListItems.Clear
    Me.LBLARCode.Caption = ""
    Me.LBLARName.Caption = ""
    Me.TextSearchAR.SetFocus
    End If
    vRecordset.Close
  End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub
