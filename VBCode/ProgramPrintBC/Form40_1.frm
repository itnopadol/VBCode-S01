VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form40_1 
   Caption         =   "หน้าพิมพ์รายงานสรุปยอดขายผู้จำหน่าย"
   ClientHeight    =   8355
   ClientLeft      =   1815
   ClientTop       =   810
   ClientWidth     =   12000
   Icon            =   "Form40_1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form40_1.frx":08CA
   ScaleHeight     =   8355
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport40_11 
      Left            =   720
      Top             =   7440
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
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMD40_11 
      Caption         =   "พิมพ์"
      Height          =   690
      Left            =   3450
      TabIndex        =   5
      Top             =   4725
      Width           =   1365
   End
   Begin MSComctlLib.ListView ListView40_11 
      Height          =   5490
      Left            =   5850
      TabIndex        =   4
      Top             =   1350
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   9684
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสเจ้าหนี้"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อเจ้าหนี้"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTP40_12 
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   3900
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   688
      _Version        =   393216
      Format          =   66584577
      CurrentDate     =   38024
   End
   Begin MSComCtl2.DTPicker DTP40_11 
      Height          =   390
      Left            =   2100
      TabIndex        =   2
      Top             =   3300
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   688
      _Version        =   393216
      Format          =   66584577
      CurrentDate     =   38024
   End
   Begin VB.TextBox TXT40_12 
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
      Height          =   390
      Left            =   2100
      TabIndex        =   1
      Top             =   2475
      Width           =   2715
   End
   Begin VB.TextBox TXT40_11 
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
      Height          =   390
      Left            =   2100
      TabIndex        =   0
      Top             =   1950
      Width           =   2715
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์รายงานสรุปยอดขายผู้จำหน่าย"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   540
      Left            =   2550
      TabIndex        =   11
      Top             =   300
      Width           =   7440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   375
      X2              =   5475
      Y1              =   6825
      Y2              =   6825
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   375
      X2              =   5475
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   5475
      X2              =   5475
      Y1              =   1425
      Y2              =   6825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   375
      X2              =   375
      Y1              =   1425
      Y2              =   6825
   End
   Begin VB.Label LBL40_12 
      BackStyle       =   0  'Transparent
      Caption         =   "รายชื่อเจ้าหนี้"
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
      Left            =   5850
      TabIndex        =   10
      Top             =   1050
      Width           =   3090
   End
   Begin VB.Label LBL40_15 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1350
      TabIndex        =   9
      Top             =   3900
      Width           =   765
   End
   Begin VB.Label LBL40_14 
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1350
      TabIndex        =   8
      Top             =   3300
      Width           =   765
   End
   Begin VB.Label LBL40_13 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงรหัสเจ้าหนี้"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   900
      TabIndex        =   7
      Top             =   2475
      Width           =   1065
   End
   Begin VB.Label LBL40_11 
      BackStyle       =   0  'Transparent
      Caption         =   "จากรหัสเจ้าหนี้"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   900
      TabIndex        =   6
      Top             =   1950
      Width           =   1215
   End
End
Attribute VB_Name = "Form40_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD40_11_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim Date1 As Date, Date2 As Date
Dim vApCode1 As String, vApCode2 As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If TXT40_11.Text <> "" And TXT40_12.Text <> "" Then
    Date1 = DTP40_11.Day & "/" & DTP40_11.Month & "/" & DTP40_11.Year
    Date2 = DTP40_12.Day & "/" & DTP40_12.Month & "/" & DTP40_12.Year
    vApCode1 = Trim(TXT40_11.Text)
    vApCode2 = Trim(TXT40_12.Text)
    
    
    vRepID = 27
    vRepType = "PO"
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from dbo.bcreportname where repid = 27 and reptype = 'PO' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            With CrystalReport40_11
                .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                .ParameterFields(0) = "@APCODE1;" & vApCode1 & ";true"
                .ParameterFields(1) = "@APCODE2;" & vApCode2 & ";true"
                .ParameterFields(2) = "@StartDate;" & Date1 & ";true"
                .ParameterFields(3) = "@EndDate;" & Date2 & ";true"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .Action = 1
            End With
    End If
    vRecordset.Close
Else
    MsgBox "กรุณาใส่เงื่อนไขให้ครบด้วยครับ", vbInformation + vbCritical, " ข้อความเตือน"
End If

TXT40_11.Text = ""
TXT40_12.Text = ""

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAPCodeItems As ListItem

On Error GoTo ErrDescription

DTP40_11 = Now
DTP40_12 = Now

vQuery = "select distinct code,name1 from BCAP order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set vAPCodeItems = ListView40_11.ListItems.Add(, , Trim(vRecordset.Fields("code").Value))
    vAPCodeItems.SubItems(1) = Trim(vRecordset.Fields("name1").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
