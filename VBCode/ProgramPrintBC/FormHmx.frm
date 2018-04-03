VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form98 
   Caption         =   "ตรวจสอบความครบถ้วนของเอกสาร"
   ClientHeight    =   7980
   ClientLeft      =   3405
   ClientTop       =   1530
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormHmx.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal102 
      Left            =   6435
      Top             =   7110
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.ComboBox CMBGroupReport 
      Height          =   315
      Left            =   7110
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   630
      Width           =   3165
   End
   Begin MSComCtl2.DTPicker DTP102 
      Height          =   330
      Left            =   7110
      TabIndex        =   17
      Top             =   1980
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   51314689
      CurrentDate     =   38899
   End
   Begin MSComCtl2.DTPicker DTP101 
      Height          =   330
      Left            =   7110
      TabIndex        =   16
      Top             =   1530
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   51314689
      CurrentDate     =   38899
   End
   Begin VB.ComboBox CMBReportType 
      Height          =   315
      Left            =   7110
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1080
      Width           =   3165
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   4200
      Top             =   7275
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
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7125
      TabIndex        =   4
      Top             =   2475
      Width           =   3165
   End
   Begin VB.CommandButton CMD103 
      Caption         =   "ดูข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   525
      TabIndex        =   3
      Top             =   2475
      Width           =   1365
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "พิมพ์รายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2100
      TabIndex        =   7
      Top             =   6750
      Width           =   1365
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "บันทึกข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   525
      TabIndex        =   6
      Top             =   6750
      Width           =   1365
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3615
      Left            =   525
      TabIndex        =   5
      Top             =   2925
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับที่"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "หมายเหตุ 1"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "หมายเหตุ 2"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ผู้สร้างเอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "หมายเหตุ"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "วันที่เอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ARCode"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "PersonCode"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   315
      Left            =   1875
      TabIndex        =   2
      Top             =   1950
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51314689
      CurrentDate     =   38652
   End
   Begin VB.ComboBox CMB102 
      Height          =   315
      Left            =   1875
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1500
      Width           =   2790
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   1875
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1050
      Width           =   2790
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทการดูรายงาน :"
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
      Left            =   5265
      TabIndex        =   18
      Top             =   630
      Width           =   1770
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่สิ้นสุด :"
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
      Left            =   5715
      TabIndex        =   15
      Top             =   1980
      Width           =   1320
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เริ่ม :"
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
      Left            =   6210
      TabIndex        =   14
      Top             =   1530
      Width           =   825
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรายงาน :"
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
      Left            =   4995
      TabIndex        =   12
      Top             =   1080
      Width           =   2040
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "คีย์ข้อมูลเข้า"
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
      Height          =   315
      Left            =   5925
      TabIndex        =   11
      Top             =   2475
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ตรวจสอบ"
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
      Left            =   525
      TabIndex        =   10
      Top             =   1950
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "หัวเอกสาร"
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
      Left            =   525
      TabIndex        =   9
      Top             =   1500
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทเอกสาร"
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
      Height          =   300
      Left            =   525
      TabIndex        =   8
      Top             =   1050
      Width           =   1170
   End
End
Attribute VB_Name = "Form98"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------
'moo
'-----------------------
Private Sub CMB101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocdate As Date

On Error GoTo ErrDescription

If CMB101.ListIndex = 0 Then
    CMB102.Clear
    ListView101.ListItems.Clear
    vQuery = "select * from bcnp.dbo.vw_ck_arinvoice order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMB101.ListIndex = 1 Then
    CMB102.Clear
    ListView101.ListItems.Clear
    vQuery = "select * from bcnp.dbo.vw_ck_purchaseorder order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMB101.ListIndex = 2 Then
    CMB102.Clear
    ListView101.ListItems.Clear
    vQuery = "select * from bcnp.dbo.vw_ck_apinvoice order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMB101.ListIndex = 3 Then
    CMB102.Clear
    ListView101.ListItems.Clear
    vQuery = "select * from bcnp.dbo.vw_ck_stktransfer order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMB101.ListIndex = 4 Then
    CMB102.Clear
    ListView101.ListItems.Clear
    vQuery = "select * from bcnp.dbo.vw_ck_stkissue order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMB101.ListIndex = 5 Then
    CMB102.Clear
    ListView101.ListItems.Clear
    vQuery = "select * from bcnp.dbo.vw_ck_receipt order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMB101.ListIndex = 6 Then
    CMB102.Clear
    ListView101.ListItems.Clear
    vQuery = "select * from bcnp.dbo.vw_ck_ardeposit order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMB101.ListIndex = 7 Then
    CMB102.Clear
    ListView101.ListItems.Clear
    vQuery = "select * from bcnp.dbo.vw_ck_creditnote order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMB101.ListIndex = 8 Then
    CMB102.Clear
    ListView101.ListItems.Clear
    vQuery = "select * from bcnp.dbo.vw_ck_debitnote order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMB101.ListIndex = 9 Then
    CMB102.Clear
    vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
    ListView101.ListItems.Clear
    vQuery = "exec bcnp.dbo.usp_CK_SearchReceiptSlip '" & vDocdate & "'"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
ElseIf CMB101.ListIndex = 10 Then
    CMB102.Clear
    ListView101.ListItems.Clear
    vQuery = "select * from bcnp.dbo.vw_ck_stkadjust order by header"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMB102_Click()
'On Error GoTo ErrDescription

vCheckButton = 0
Call InsertData
Call SaveData

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub



Private Sub CMBGroupReport_Click()
If CMBGroupReport.ListIndex = 1 Then
  CMBReportType.Enabled = False
    DTP101.Enabled = False
    DTP102.Enabled = False
Else
  CMBReportType.Enabled = True
  If CMBReportType.ListIndex = 1 Or CMBReportType.ListIndex = 3 Then
    DTP101.Enabled = True
    DTP102.Enabled = True
  Else
    DTP101.Enabled = False
    DTP102.Enabled = False
  End If
End If
End Sub

Private Sub CMBReportType_Click()
If CMBReportType.ListIndex = 1 Or CMBReportType.ListIndex = 3 Then
  DTP101.Enabled = True
  DTP102.Enabled = True
Else
  DTP101.Enabled = False
  DTP102.Enabled = False
End If
End Sub

Private Sub CMD101_Click()
On Error GoTo ErrDescription

vCheckButton = 1
Call SaveData

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
Dim vDocdate As Date
Dim vDocType As Integer
Dim vHeader As String
Dim vGenDate As String
Dim vDay As String
Dim vReportType As Integer
Dim vDocGroup  As Integer
Dim vBegDate As String
Dim vEndDate As String
Dim vParameter As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If CMB101.Text <> "" Then

If CMBGroupReport.ListIndex = 1 Then
  vDocType = CMB101.ListIndex
  If Len(DTPicker101.Day) = 1 Then
      vDay = Trim(0 & DTPicker101.Day)
  Else
      vDay = Trim(DTPicker101.Day)
  End If
  vGenDate = Right(DTPicker101.Year, 2) + 43 & DTPicker101.Month & vDay
  vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
  If CMB101.ListIndex <> 9 Then
      vHeader = Trim(CMB102.Text) & vGenDate
  Else
      vHeader = Trim(CMB102.Text)
  End If
  
  If CMB101.Text <> "" And vHeader <> "" Then
  
  vRepID = 288
  vRepType = "CK"
  vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
      'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 288 and reptype = 'CK' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          vReportName = Trim(vRecordset.Fields("reportname").Value)
      End If
      vRecordset.Close
  
      With Crystal101
      .ReportFileName = vReportName & ".rpt"
      .ParameterFields(0) = "@vDocType;" & vDocType & ";true"
      .ParameterFields(1) = "@vRunNumber;" & vHeader & ";true"
      .Destination = crptToWindow
      .WindowState = crptMaximized
      .Action = 1
      End With
  Else
      MsgBox "กรอกข้อมูลไม่ครบตามความต้องการของ รายงาน กรุณาตรวจสอบ", vbInformation, "ข้อความแจ้งเตือน"
      Exit Sub
  End If
  
ElseIf CMBGroupReport.ListIndex = 0 Then

  Select Case CMBReportType.ListIndex
  Case 0:
    vReportType = 0
    vHeader = Trim(CMB102.Text)
  Case 1:
    vReportType = 1
    vHeader = Trim(CMB102.Text)
  Case 2:
    vReportType = 2
    vHeader = ""
  Case 3:
    vReportType = 3
    vHeader = ""
  End Select
  
  Select Case CMB101.ListIndex
  Case 0:
    vParameter = 0
    vRepID = 318
  Case 1:
    vParameter = 1
    vRepID = 319
  Case 2:
    vParameter = 1
    vRepID = 319
  Case 3:
    vParameter = 2
    vRepID = 320
  Case 4:
    vParameter = 2
    vRepID = 320
  Case 5:
    vParameter = 0
    vRepID = 318
  Case 6:
    vParameter = 0
    vRepID = 318
  Case 7:
    vParameter = 0
    vRepID = 318
  Case 8:
    vParameter = 0
    vRepID = 318
  Case 9:
    vParameter = 2
    vRepID = 320
  Case 10:
    vParameter = 2
    vRepID = 320
  End Select
  vRepType = "CK"
  vDocGroup = CMB101.ListIndex

  If CMBReportType.ListIndex = 1 Or CMBReportType.ListIndex = 3 Then
    vBegDate = Trim(Day(DTP101) & "/" & Month(DTP101) & "/" & Year(DTP101))
    vEndDate = Trim(Day(DTP102) & "/" & Month(DTP102) & "/" & Year(DTP102))
  Else
    vBegDate = Trim(Day(DTPicker101) & "/" & Month(DTPicker101) & "/" & Year(DTPicker101))
    vEndDate = Trim(Day(DTPicker101) & "/" & Month(DTPicker101) & "/" & Year(DTPicker101))
  End If
  
  vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
  'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          vReportName = Trim(vRecordset.Fields("reportname").Value)
  End If
  vRecordset.Close
  
      With Crystal102
      .ReportFileName = vReportName & ".rpt"
      .ParameterFields(0) = "@vReportType;" & vReportType & ";true"
      .ParameterFields(1) = "@vHeader;" & vHeader & ";true"
      .ParameterFields(2) = "@vBegDate;" & vBegDate & ";true"
      .ParameterFields(3) = "@vEndDate;" & vEndDate & ";true"
      .ParameterFields(4) = "@vDocGroup;" & vDocGroup & ";true"
      .ParameterFields(5) = "@vParameter;" & vParameter & ";true"
      .Destination = crptToWindow
      .WindowState = crptMaximized
      .Action = 1
      End With
      
  '--------------------------------------------------------------------------
  ElseIf CMBGroupReport.ListIndex = 2 Then

  Select Case CMBReportType.ListIndex
  Case 0:
    vReportType = 0
    vHeader = Trim(CMB102.Text)
  Case 1:
    vReportType = 1
    vHeader = Trim(CMB102.Text)
  Case 2:
    vReportType = 2
    vHeader = ""
  Case 3:
    vReportType = 3
    vHeader = ""
  End Select
  
  Select Case CMB101.ListIndex
  Case 0:
    vParameter = 0
  Case 1:
    vParameter = 1
  Case 2:
    vParameter = 1
  Case 3:
    vParameter = 2
  Case 4:
    vParameter = 2
  Case 5:
    vParameter = 0
  Case 6:
    vParameter = 0
  Case 7:
    vParameter = 0
  Case 8:
    vParameter = 0
  Case 9:
    vParameter = 2
  Case 10:
    vParameter = 2
  End Select
  vRepID = 321
  vRepType = "CK"
  vDocGroup = CMB101.ListIndex

  If CMBReportType.ListIndex = 1 Or CMBReportType.ListIndex = 3 Then
    vBegDate = Trim(Day(DTP101) & "/" & Month(DTP101) & "/" & Year(DTP101))
    vEndDate = Trim(Day(DTP102) & "/" & Month(DTP102) & "/" & Year(DTP102))
  Else
    vBegDate = Trim(Day(DTPicker101) & "/" & Month(DTPicker101) & "/" & Year(DTPicker101))
    vEndDate = Trim(Day(DTPicker101) & "/" & Month(DTPicker101) & "/" & Year(DTPicker101))
  End If
  
  vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
  'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          vReportName = Trim(vRecordset.Fields("reportname").Value)
  End If
  vRecordset.Close
  
      With Crystal102
      .ReportFileName = vReportName & ".rpt"
      .ParameterFields(0) = "@vReportType;" & vReportType & ";true"
      .ParameterFields(1) = "@vHeader;" & vHeader & ";true"
      .ParameterFields(2) = "@vBegDate;" & vBegDate & ";true"
      .ParameterFields(3) = "@vEndDate;" & vEndDate & ";true"
      .ParameterFields(4) = "@vDocGroup;" & vDocGroup & ";true"
      .ParameterFields(5) = "@vParameter;" & vParameter & ";true"
      .Destination = crptToWindow
      .WindowState = crptMaximized
      .Action = 1
      End With
End If
Else
  MsgBox "ต้องเลือกประเภทเอกสารก่อนเป็นอย่างน้อย", vbCritical, "Sent Massege"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD103_Click()
'On Error GoTo ErrDescription

vCheckButton = 0
Call InsertData
Call SaveData

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub DTPicker101_Change()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocdate As Date

'On Error GoTo ErrDescription

If CMB101.ListIndex = 9 Then
CMB102.Clear
vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
ListView101.ListItems.Clear
vQuery = "exec bcnp.dbo.usp_CK_SearchReceiptSlip '" & vDocdate & "'"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB102.AddItem Trim(vRecordset.Fields("header").Value)
        CMB102.Text = Trim(vRecordset.Fields("header").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End If

vCheckButton = 0
Call InsertData
Call SaveData

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Form_Load()
DTPicker101.Value = Now
CMB101.AddItem Trim("บิลขาย")
CMB101.AddItem Trim("ใบสั่งซื้อ")
CMB101.AddItem Trim("ใบรับเข้าสินค้า")
CMB101.AddItem Trim("ใบโอนย้ายสินค้า")
CMB101.AddItem Trim("ใบเบิกจ่ายสินค้า")
CMB101.AddItem Trim("ใบเสร็จรับชำระ")
CMB101.AddItem Trim("ใบมัดจำ")
CMB101.AddItem Trim("ใบลดหนี้")
CMB101.AddItem Trim("ใบเพิ่มหนี้")
CMB101.AddItem Trim("ใบจ่ายสินค้า")
CMB101.AddItem Trim("ใบปรับปรุงสินค้า")
CMBReportType.AddItem Trim("ตามหัวเอกสารประจำวันที่")
CMBReportType.AddItem Trim("ตามหัวเอกสารช่วงวันที่")
CMBReportType.AddItem Trim("ตามกลุ่มเอกสารประจำวันที่")
CMBReportType.AddItem Trim("ตามกลุ่มเอกสารช่วงวันที่")
CMBReportType.Text = Trim("ตามหัวเอกสารประจำวันที่")
CMBGroupReport.AddItem Trim("รายงานเอกสารไม่ครบ")
CMBGroupReport.AddItem Trim("รายงานตรวจสอบเอกสาร")
CMBGroupReport.AddItem Trim("รายงานสรุปเอกสารไม่ครบ")
CMBGroupReport.Text = Trim("รายงานเอกสารไม่ครบ")
DTP101 = Now
DTP102 = Now
End Sub

Public Sub CheckDocument()
Dim vRecordset As New Recordset
Dim vQuery As String
Dim i As Integer
Dim vHeader As String
Dim vDocdate As Date
Dim vModule As Integer
Dim vCheckExist As Integer
Dim vDocNo As String

'vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
'vHeader = Trim(CMB102.Text)
'vModule = CMB101.ListIndex

'For i = 1 To ListView101.ListItems.Count
 '   vDocno = ListView101.ListItems.Item(i).SubItems(1)
  '  vQuery = "set dateformat dmy"
   ' gConnection.Execute vQuery
    'vQuery = "select existstatus from npmaster.dbo.TB_CK_Document where doctype = " & vModule & " and docno = '" & vDocno & "' and invoicedate = '" & vDocdate & "' "
    'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     '   vCheckExist = Trim(vRecordset.Fields("existstatus").Value)
    'Else
     '   vCheckExist = 0
    'End If
    'vRecordset.Close
    'If vCheckExist = 1 Then
     '   ListView101.ListItems(i).Checked = True
    'Else
     '   ListView101.ListItems.Item(i).Checked = False
    'End If
'Next i

End Sub

Private Sub ListView101_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vMydescription As String

vMydescription = InputBox("กรอกหมายเหตุ", "หมายเหตุของรายการเอกสาร", Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5)))
ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5) = vMydescription
End Sub


Public Sub SaveData()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRunNumber As String
Dim vDocType As Integer
Dim vDocNo As String
Dim vExistStatus As Integer
Dim vMydescription As String
Dim vChecker As String
Dim vHeader As String
Dim vExist As Integer
Dim i As Integer
Dim vNumberNo As String
Dim vCheckDate As Date
Dim vCount As Integer
Dim vDate As String
Dim vAutoNumber As String
Dim vCheckExistStatus As Integer
Dim vCheckMydescription As String
Dim vDay As String
Dim vARCode As String
Dim vPersonCode As String
Dim vCheckUser As String
Dim vCheckPersonCode As String
Dim vCheckARCode As String

If CMB101.Text <> "" Then
    If CMB102.Text <> "" Then
        If ListView101.ListItems.Count <> 0 Then
            vCheckDate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
            If Len(DTPicker101.Day) = 1 Then
                vDay = Trim(0 & DTPicker101.Day)
            Else
                vDay = Trim(DTPicker101.Day)
            End If
            vDate = Right(DTPicker101.Year, 2) + 43 & DTPicker101.Month & vDay
            vDocType = CMB101.ListIndex
            vHeader = Trim(CMB102.Text)
            If CMB101.ListIndex <> 9 Then
            vNumberNo = vHeader & vDate
            Else
            vNumberNo = vHeader
            End If
            vQuery = "set dateformat dmy"
            gConnection.Execute vQuery
            vQuery = "select  distinct runnumber  from npmaster.dbo.TB_CK_Document where  runnumber like '%" & vHeader & "%' and doctype = " & vDocType & " and runnumber = '" & vNumberNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vCount = 1
            Else
                vCount = 0
            End If
            vRecordset.Close
            For i = 1 To ListView101.ListItems.Count
                vRunNumber = vNumberNo
                vDocNo = Trim(ListView101.ListItems.Item(i).SubItems(1))
                vARCode = Trim(ListView101.ListItems.Item(i).SubItems(7))
                vPersonCode = Trim(ListView101.ListItems.Item(i).SubItems(8))
                If ListView101.ListItems.Item(i).Checked = True Then
                    vExistStatus = 1
                Else
                    vExistStatus = 0
                End If
                vMydescription = Trim(ListView101.ListItems.Item(i).SubItems(5))
                
                If CMB101.ListIndex <> 9 Then
                    vQuery = "select docno,existstatus,mydescription from npmaster.dbo.tb_ck_document where runnumber = '" & vRunNumber & "' and docno = '" & vDocNo & "' and doctype = " & vDocType & " " 'and docdate = '" & vCheckDate & "' "
                    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                        vExist = 1
                        vCheckExistStatus = Trim(vRecordset.Fields("existstatus").Value)
                        vCheckMydescription = Trim(vRecordset.Fields("mydescription").Value)
                    Else
                        vExist = 0
                    End If
                    vRecordset.Close
                Else
                    vQuery = "select docno,existstatus,mydescription,checker,code,personcode  from npmaster.dbo.tb_ck_document where runnumber = '" & vRunNumber & "' and docno = '" & vDocNo & "' " ' and doctype = " & vDocType & " "
                    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                        vExist = 1
                        vCheckExistStatus = Trim(vRecordset.Fields("existstatus").Value)
                        vCheckMydescription = Trim(vRecordset.Fields("mydescription").Value)
                        vCheckUser = Trim(vRecordset.Fields("checker").Value)
                        vCheckARCode = Trim(vRecordset.Fields("code").Value)
                        vCheckPersonCode = Trim(vRecordset.Fields("personcode").Value)
                    Else
                        vExist = 0
                    End If
                    vRecordset.Close
                End If

                If vExist = 0 Then
                vQuery = "exec bcnp.dbo.usp_CK_InsertCheckDocumentLogs '" & vRunNumber & "'," & vDocType & ",'" & vDocNo & "','" & vCheckDate & "'," & vExistStatus & ",'" & vMydescription & "','" & vUserID & "' ,'" & vARCode & "' ,'" & vPersonCode & "' "
                gConnection.Execute vQuery
                ElseIf (vExistStatus <> vCheckExistStatus) Or vMydescription <> vCheckMydescription Or vCheckUser <> vUserID Or vCheckARCode <> vARCode Or vCheckPersonCode <> vPersonCode Then
                vQuery = "exec bcnp.dbo.usp_CK_UpdateCheckStatusDocument '" & vRunNumber & "'," & vDocType & ",'" & vDocNo & "','" & vCheckDate & "','" & vMydescription & "'," & vExistStatus & ",'" & vUserID & "','" & vARCode & "','" & vPersonCode & "' "
                gConnection.Execute vQuery
                End If
            Next i
            If vCheckButton = 1 Then
                ListView101.ListItems.Clear
            End If
        Else
            If vCheckButton = 1 Then
                MsgBox "ไม่มีรายการเอกสารให้บันทึกข้อมูล", vbInformation, "ข้อความแจ้งเตือน"
            End If
            Exit Sub
        End If
    Else
        MsgBox "ไม่ได้เลือกหัวเอกสาร", vbInformation, "ข้อความแจ้งเตือน"
        Exit Sub
    End If
Else
    MsgBox "ไม่ได้เลือกประเภทเอกสาร", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If
End Sub

Public Sub InsertData()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vHeader As String
Dim vDocdate As Date
Dim vListDocno As ListItem
Dim i As Integer

If CMB101.ListIndex = 0 Then
    If CMB102.Text <> "" Then
        i = 1
        ListView101.ListItems.Clear
        vHeader = Trim(CMB102.Text)
        vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec bcnp.dbo.usp_ck_arinvoice '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    ListView101.ColumnHeaders(3).Text = Trim("ชื่อลูกค้า")
                    ListView101.ColumnHeaders(4).Text = Trim("ยอดเงิน")
                    Set vListDocno = ListView101.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("netdebtamount").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(6) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("arcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("salecode").Value)
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.Checked = True
                    Else
                        vListDocno.Checked = False
                    End If
                If vRecordset.Fields("iscancel").Value = 1 Then
                    ListView101.ListItems.Item(i).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
                End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMB101.ListIndex = 1 Then
    If CMB102.Text <> "" Then
        i = 1
        ListView101.ListItems.Clear
        vHeader = Trim(CMB102.Text)
        vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec bcnp.dbo.usp_ck_purchaseorder '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    ListView101.ColumnHeaders(3).Text = Trim("ชื่อเจ้าหนี้")
                    ListView101.ColumnHeaders(4).Text = Trim("ยอดเงิน")
                    Set vListDocno = ListView101.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = Trim(vRecordset.Fields("apname").Value)
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("netamount").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(6) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("apcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("creatorcode").Value)
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.Checked = True
                    Else
                        vListDocno.Checked = False
                    End If
                If vRecordset.Fields("iscancel").Value = 1 Then
                    ListView101.ListItems.Item(i).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
                End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMB101.ListIndex = 2 Then
    If CMB102.Text <> "" Then
        i = 1
        ListView101.ListItems.Clear
        vHeader = Trim(CMB102.Text)
        vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec bcnp.dbo.usp_ck_apinvoice '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    ListView101.ColumnHeaders(3).Text = Trim("ชื่อเจ้าหนี้")
                    ListView101.ColumnHeaders(4).Text = Trim("ยอดเงิน")
                    Set vListDocno = ListView101.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = Trim(vRecordset.Fields("apname").Value)
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("netdebtamount").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(6) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("apcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("creatorcode").Value)
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.Checked = True
                    Else
                        vListDocno.Checked = False
                    End If
                If vRecordset.Fields("iscancel").Value = 1 Then
                    ListView101.ListItems.Item(i).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
                End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMB101.ListIndex = 3 Then
    If CMB102.Text <> "" Then
        i = 1
        ListView101.ListItems.Clear
        vHeader = Trim(CMB102.Text)
        vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec bcnp.dbo.usp_ck_stktransfer '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListView101.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(3) = ""
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(6) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(7) = ""
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("creatorcode").Value)
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.Checked = True
                    Else
                        vListDocno.Checked = False
                    End If
                If vRecordset.Fields("iscancel").Value = 1 Then
                    ListView101.ListItems.Item(i).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMB101.ListIndex = 4 Then
    If CMB102.Text <> "" Then
        i = 1
        ListView101.ListItems.Clear
        vHeader = Trim(CMB102.Text)
        vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec bcnp.dbo.usp_ck_stkissue'" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListView101.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(3) = ""
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(6) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(7) = ""
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("creatorcode").Value)
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.Checked = True
                    Else
                        vListDocno.Checked = False
                    End If
                If vRecordset.Fields("iscancel").Value = 1 Then
                    ListView101.ListItems.Item(i).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMB101.ListIndex = 5 Then
    If CMB102.Text <> "" Then
        i = 1
        ListView101.ListItems.Clear
        vHeader = Trim(CMB102.Text)
        vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec bcnp.dbo.usp_ck_receipt '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListView101.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(3) = ""
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(6) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(7) = ""
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("creatorcode").Value)
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.Checked = True
                    Else
                        vListDocno.Checked = False
                    End If
                If vRecordset.Fields("iscancel").Value = 1 Then
                    ListView101.ListItems.Item(i).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    ElseIf CMB101.ListIndex = 6 Then
    If CMB102.Text <> "" Then
        i = 1
        ListView101.ListItems.Clear
        vHeader = Trim(CMB102.Text)
        vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec bcnp.dbo.usp_ck_ardeposit'" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListView101.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(3) = ""
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(6) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("arcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("creatorcode").Value)
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.Checked = True
                    Else
                        vListDocno.Checked = False
                    End If
                If vRecordset.Fields("iscancel").Value = 1 Then
                    ListView101.ListItems.Item(i).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    ElseIf CMB101.ListIndex = 7 Then
    If CMB102.Text <> "" Then
        i = 1
        ListView101.ListItems.Clear
        vHeader = Trim(CMB102.Text)
        vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec bcnp.dbo.usp_ck_arcreditnote '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListView101.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(3) = ""
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(6) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("arcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("creatorcode").Value)
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.Checked = True
                    Else
                        vListDocno.Checked = False
                    End If
                If vRecordset.Fields("iscancel").Value = 1 Then
                    ListView101.ListItems.Item(i).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
    ElseIf CMB101.ListIndex = 8 Then
    If CMB102.Text <> "" Then
        i = 1
        ListView101.ListItems.Clear
        vHeader = Trim(CMB102.Text)
        vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec bcnp.dbo.usp_ck_debitnote '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListView101.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
                    vListDocno.SubItems(3) = ""
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(6) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("arcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("creatorcode").Value)
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.Checked = True
                    Else
                        vListDocno.Checked = False
                    End If
                If vRecordset.Fields("iscancel").Value = 1 Then
                    ListView101.ListItems.Item(i).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMB101.ListIndex = 9 Then
    If CMB102.Text <> "" Then
        i = 1
        ListView101.ListItems.Clear
        vHeader = Trim(CMB102.Text)
        vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec bcnp.dbo.usp_CK_ReceiptSlip '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    ListView101.ColumnHeaders(2).Text = Trim("Running Number")
                    ListView101.ColumnHeaders(3).Text = Trim("เลขที่บิล")
                    ListView101.ColumnHeaders(4).Text = Trim("คลัง")
                    ListView101.ColumnHeaders(5).Text = Trim("คนพิมพ์เอกสาร")
                    Set vListDocno = ListView101.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = Trim(vRecordset.Fields("invoiceno").Value)
                    vListDocno.SubItems(3) = Trim(vRecordset.Fields("whcode").Value)
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("userprint").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(6) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(7) = Trim(vRecordset.Fields("arcode").Value)
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("salecode").Value)
                    If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                        vListDocno.Checked = True
                    Else
                        vListDocno.Checked = False
                    End If

            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
ElseIf CMB101.ListIndex = 10 Then
    If CMB102.Text <> "" Then
        i = 1
        ListView101.ListItems.Clear
        vHeader = Trim(CMB102.Text)
        vDocdate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec bcnp.dbo.usp_ck_stkadjust '" & vDocdate & "','" & vHeader & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListDocno = ListView101.ListItems.Add(, , i)
                    vListDocno.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDocno.SubItems(2) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDocno.SubItems(3) = ""
                    vListDocno.SubItems(4) = Trim(vRecordset.Fields("creatorcode").Value)
                    vListDocno.SubItems(5) = Trim(vRecordset.Fields("description1").Value)
                    vListDocno.SubItems(6) = Trim(vRecordset.Fields("docdate").Value)
                    vListDocno.SubItems(7) = ""
                    vListDocno.SubItems(8) = Trim(vRecordset.Fields("creatorcode").Value)
                If vRecordset.Fields("iscancel").Value = 1 Then
                    ListView101.ListItems.Item(i).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                    ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                End If
                If Trim(vRecordset.Fields("existstatus").Value) = 1 Then
                    vListDocno.Checked = True
                Else
                    vListDocno.Checked = False
                End If
            vRecordset.MoveNext
            i = i + 1
            Wend
        End If
        vRecordset.Close
    End If
End If
    'Call CheckDocument
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocNo As String
Dim i As Integer
Dim vCheckDate As Date
Dim vModule As Integer
Dim vDocNo As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    vCheckDocNo = UCase(Trim(Text101.Text))
    vCheckDate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
    vModule = CMB101.ListIndex
    
    For i = 1 To ListView101.ListItems.Count
        vDocNo = UCase(Trim(ListView101.ListItems.Item(i).SubItems(1)))
        If vCheckDocNo = vDocNo Then
            ListView101.ListItems.Item(i).Checked = True
        End If
    Next i
    Text101.Text = ""
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
