VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Form814 
   Caption         =   "พิมพ์ใบขอเบิกสินค้าและวัตถุดิบประจำวัน"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form814.frx":0000
   ScaleHeight     =   11490
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CHKPrintAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "พิมพ์ทั้งหมด"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2880
      TabIndex        =   3
      Top             =   7200
      Width           =   1230
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   765
      Top             =   8550
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
   Begin VB.PictureBox PicPoint 
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   585
      TabIndex        =   6
      Top             =   0
      Width           =   645
   End
   Begin VB.CommandButton BTNPrint 
      Caption         =   "พิมพ์"
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
      Left            =   1260
      TabIndex        =   2
      Top             =   7020
      Width           =   1500
   End
   Begin MSComctlLib.ListView ListViewDocList 
      Height          =   4560
      Left            =   1260
      TabIndex        =   1
      Top             =   2385
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   8043
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
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ประเภทการเบิก"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ชื่อผู้ขอเบิก"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หมายเหตุ"
         Object.Width           =   7937
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPDocDate 
      Height          =   330
      Left            =   2385
      TabIndex        =   0
      Top             =   1395
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   62324737
      CurrentDate     =   40014
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการเอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1260
      TabIndex        =   5
      Top             =   2070
      Width           =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประจำวันที่ :"
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
      Left            =   1170
      TabIndex        =   4
      Top             =   1395
      Width           =   1095
   End
End
Attribute VB_Name = "Form814"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BTNPrint_Click()
Dim i As Integer
Dim vDocdate As String
Dim vDocNo As String

On Error Resume Next

vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year
If Me.CHKPrintAll.Value = 1 Then
Call PrintIssueAll(vDocdate)
Else

For i = 1 To Me.ListViewDocList.ListItems.Count
If Me.ListViewDocList.ListItems(i).Checked = True Then
vDocNo = Me.ListViewDocList.ListItems(i).SubItems(2)
Call PrintIssue(vDocNo)
End If
Next i

End If
End Sub

Public Sub PrintIssue(vDocNo As String)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error Resume Next

vRepID = 470
vRepType = "IV"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
    
End Sub

Public Sub PrintIssueAll(vDocdate As String)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error Resume Next

vRepID = 471
vRepType = "IV"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vDocDate;" & vDocdate & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
    
End Sub

Private Sub CHKPrintAll_Click()
Dim i As Integer

On Error Resume Next

If Me.CHKPrintAll.Value = 0 Then
   For i = 1 To Me.ListViewDocList.ListItems.Count
   Me.ListViewDocList.ListItems(i).Checked = False
   Next i
Else
   For i = 1 To Me.ListViewDocList.ListItems.Count
   Me.ListViewDocList.ListItems(i).Checked = True
   Next i
End If
End Sub

Private Sub DTPDocDate_Change()
Call GetDocIssue
End Sub

Public Sub GetDocIssue()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocdate As String
Dim i As Integer
Dim vListItem As ListItem

On Error Resume Next

Me.ListViewDocList.ListItems.Clear
vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year

vQuery = "exec dbo.USP_NP_IssueDocumentByDocDate '" & vDocdate & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    For i = 1 To vRecordset.RecordCount
            Set vListItem = Me.ListViewDocList.ListItems.Add(, , i)
            vListItem.SubItems(1) = Trim(vRecordset.Fields("header").Value)
            vListItem.SubItems(2) = Trim(vRecordset.Fields("docno").Value)
            vListItem.SubItems(3) = Trim(vRecordset.Fields("personcode").Value) & "/" & Trim(vRecordset.Fields("personname").Value)
            vListItem.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
    vRecordset.MoveNext
Next i

Else
MsgBox "ไม่มีรายการเอกสาร ณ วันที่ค้นหา กรุณาตรวจสอบ", vbCritical, "Send Error Message"
End If
vRecordset.Close

End Sub

Private Sub Form_Load()
Me.DTPDocDate.Value = Now
Call SetListViewColor(ListViewDocList, PicPoint, vbWhite, vbLightGreen)
Call GetDocIssue
End Sub

