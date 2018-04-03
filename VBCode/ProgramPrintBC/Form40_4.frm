VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form40_4 
   Caption         =   "รายงานมัดจำจ่าย เจ้าหนี้"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form40_4.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport40_4 
      Left            =   1800
      Top             =   5880
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
   Begin VB.CommandButton CMD40_41 
      Caption         =   "ดูรายงาน"
      Height          =   540
      Left            =   4350
      TabIndex        =   7
      Top             =   3600
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTP40_42 
      Height          =   390
      Left            =   3525
      TabIndex        =   5
      Top             =   2925
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   66584577
      CurrentDate     =   38075
   End
   Begin MSComCtl2.DTPicker DTP40_41 
      Height          =   390
      Left            =   3525
      TabIndex        =   4
      Top             =   2325
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   66584577
      CurrentDate     =   38075
   End
   Begin VB.TextBox TXT40_41 
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
      Left            =   3525
      TabIndex        =   3
      Top             =   1725
      Width           =   2190
   End
   Begin VB.Label LBL40_44 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
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
      Height          =   315
      Left            =   2625
      TabIndex        =   6
      Top             =   2925
      Width           =   840
   End
   Begin VB.Label LBL40_43 
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่"
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
      Height          =   390
      Left            =   2625
      TabIndex        =   2
      Top             =   2325
      Width           =   765
   End
   Begin VB.Label LBL40_42 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสเจ้าหนี้"
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
      Height          =   390
      Left            =   2625
      TabIndex        =   1
      Top             =   1725
      Width           =   840
   End
   Begin VB.Label LBL40_41 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงาน มัดจำจ่าย แยกตาม เจ้าหนี้"
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
      Left            =   2625
      TabIndex        =   0
      Top             =   225
      Width           =   7440
   End
End
Attribute VB_Name = "Form40_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD40_41_Click()
Dim vQuery As String, vApCode As String
Dim vRecordset As New ADODB.Recordset
Dim Date1 As Date, Date2 As Date
Dim vRepType As String, vReportName As String
Dim vRepID As Integer

On Error GoTo ErrDescription

vApCode = Trim(TXT40_41.Text)
Date1 = DTP40_41.Day & "/" & DTP40_41.Month & "/" & DTP40_41.Year
Date2 = DTP40_42.Day & "/" & DTP40_42.Month & "/" & DTP40_42.Year
vRepID = 87
vRepType = "AP"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)

    With CrystalReport40_4
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@APCode;" & vApCode & ";true"
    .ParameterFields(1) = "@StartDate;" & Date1 & ";true"
    .ParameterFields(2) = "@EndDate;" & Date2 & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    
    End With
End If

vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
DTP40_41 = Now
DTP40_42 = Now
End Sub
