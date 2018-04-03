VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form395 
   Caption         =   "รายงาน เกี่ยวกับเอกสารขาย"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form395.frx":0000
   ScaleHeight     =   8640
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   2640
      Top             =   4080
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
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์เอกสาร"
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
      Left            =   3975
      TabIndex        =   2
      Top             =   1950
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   315
      Left            =   2925
      TabIndex        =   1
      Top             =   1350
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   556
      _Version        =   393216
      Format          =   69599233
      CurrentDate     =   38560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่"
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
      Left            =   2250
      TabIndex        =   0
      Top             =   1350
      Width           =   615
   End
End
Attribute VB_Name = "Form395"
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
Dim vDate As Date

vRepID = 252
vRepType = "SO"

vDate = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@date1;" & vDate & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End Sub

Private Sub Form_Load()
DTPicker101.Value = Now
End Sub
