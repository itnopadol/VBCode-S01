VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Form43 
   Caption         =   "รายงาน เจ้าหนี้ค้างชำระ"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   222
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form43.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXTAPCode 
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
      Height          =   360
      Left            =   2115
      TabIndex        =   9
      Top             =   5760
      Width           =   3300
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   630
      Top             =   7605
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
      Caption         =   "พิมพ์"
      Height          =   690
      Left            =   2115
      TabIndex        =   7
      Top             =   6750
      Width           =   1455
   End
   Begin VB.CommandButton CMDSearchAP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7245
      Picture         =   "Form43.frx":9673
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1260
      Width           =   330
   End
   Begin VB.TextBox TXTSearchAP 
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
      Left            =   2115
      TabIndex        =   2
      Top             =   1260
      Width           =   5100
   End
   Begin MSComctlLib.ListView ListViewAp 
      Height          =   3570
      Left            =   2115
      TabIndex        =   0
      Top             =   2115
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   6297
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
         Text            =   "รหัสเจ้าหนี้"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อเจ้าหนี้"
         Object.Width           =   12347
      EndProperty
   End
   Begin VB.Label LBLAPName 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2115
      TabIndex        =   8
      Top             =   6255
      Width           =   9825
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อเจ้าหนี้ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   855
      TabIndex        =   6
      Top             =   6255
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสเจ้าหนี้ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   855
      TabIndex        =   5
      Top             =   5805
      Width           =   1185
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ค้นหา :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   990
      TabIndex        =   3
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "รายชื่อ เจ้าหนี้"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   2115
      TabIndex        =   1
      Top             =   1845
      Width           =   1545
   End
End
Attribute VB_Name = "Form43"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDPrint_Click()
Dim vAPCode As String
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType As String
Dim vReportName As String

On Error Resume Next

If Me.LBLAPName.Caption <> "" Then
vAPCode = Me.TXTAPCode.Text

vRepID = 500
vRepType = "AC"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
    .ReportFileName = vReportName & ".rpt"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .ParameterFields(0) = "@vAPCode;" & vAPCode & ";true"
    .Action = 1
End With
Else
MsgBox "ไม่มีรหัสเจ้าหนี้นี้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
End If
End Sub

Private Sub CMDSearchAP_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vAPList As ListItem
Dim n As Integer

On Error Resume Next

vSearch = Me.TXTSearchAP.Text

vQuery = "exec dbo.USP_AP_SearchAPCode '" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    n = n + 1
    Set vAPList = ListViewAp.ListItems.Add(, , n)
    vAPList.SubItems(1) = Trim(vRecordset.Fields("code").Value)
    vAPList.SubItems(2) = Trim(vRecordset.Fields("name1").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
        
End Sub

Private Sub ListViewAp_DblClick()
On Error Resume Next

If Me.ListViewAp.ListItems.Count > 0 Then
Me.TXTAPCode.Text = Me.ListViewAp.ListItems(Me.ListViewAp.SelectedItem.Index).SubItems(1)
End If
End Sub

Private Sub ListViewAp_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next

If Me.ListViewAp.ListItems.Count > 0 Then
Me.TXTAPCode.Text = Me.ListViewAp.ListItems(Me.ListViewAp.SelectedItem.Index).SubItems(1)
End If
End Sub

Private Sub TXTAPCode_Change()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAPCode As String
Dim vAPList As ListItem
Dim n As Integer

On Error Resume Next

vAPCode = Me.TXTAPCode.Text

vQuery = "exec dbo.USP_AP_SearchAPCodeDetails '" & vAPCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.LBLAPName.Caption = Trim(vRecordset.Fields("name1").Value)
Else
    Me.LBLAPName.Caption = ""
End If
vRecordset.Close

End Sub

Private Sub TXTSearchAP_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
Call CMDSearchAP_Click
End If
End Sub
