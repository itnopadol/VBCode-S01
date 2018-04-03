VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form3_81 
   Caption         =   "หน้ารายงานยอดขายของพนักงานขาย"
   ClientHeight    =   8205
   ClientLeft      =   2130
   ClientTop       =   645
   ClientWidth     =   12000
   Icon            =   "Form3_81.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form3_81.frx":08CA
   ScaleHeight     =   8205
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport3_81 
      Left            =   720
      Top             =   6120
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   5775
      TabIndex        =   6
      Top             =   1425
      Width           =   3165
      Begin VB.OptionButton Opt3_83 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ยอดขายตามจุด POS"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   225
         TabIndex        =   9
         Top             =   1200
         Width           =   1965
      End
      Begin VB.OptionButton Opt3_82 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "รายงานวิเคราะห์ยอดขาย"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   225
         TabIndex        =   8
         Top             =   675
         Width           =   2415
      End
      Begin VB.OptionButton Opt3_81 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "รายงานยอดขายรวมมัดจำ"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   225
         TabIndex        =   7
         Top             =   150
         Value           =   -1  'True
         Width           =   2265
      End
   End
   Begin VB.CommandButton CMD3_81 
      Caption         =   "พิมพ์"
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
      Left            =   3375
      TabIndex        =   2
      Top             =   2775
      Width           =   1140
   End
   Begin MSComCtl2.DTPicker DTP3_82 
      Height          =   315
      Left            =   2250
      TabIndex        =   1
      Top             =   2100
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   556
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
      Format          =   69795841
      CurrentDate     =   38027
   End
   Begin MSComCtl2.DTPicker DTP3_81 
      Height          =   315
      Left            =   2250
      TabIndex        =   0
      Top             =   1425
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   556
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
      Format          =   69795841
      CurrentDate     =   38027
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงานยอดขายพนักงาน"
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
      TabIndex        =   5
      Top             =   300
      Width           =   7290
   End
   Begin VB.Label LBL3_82 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
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
      Left            =   1575
      TabIndex        =   4
      Top             =   2100
      Width           =   690
   End
   Begin VB.Label LBL3_81 
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่"
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
      Left            =   1500
      TabIndex        =   3
      Top             =   1425
      Width           =   765
   End
End
Attribute VB_Name = "Form3_81"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD3_81_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim Date1 As Date, Date2 As Date
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

Date1 = DTP3_81.Day & "/" & DTP3_81.Month & "/" & DTP3_81.Year
Date2 = DTP3_82.Day & "/" & DTP3_82.Month & "/" & DTP3_82.Year

If Opt3_81.Value = True Then
vRepID = 29
vRepType = "SO"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = 29 and reptype = 'SO' "
ElseIf Opt3_82.Value = True Then
vRepID = 47
vRepType = "SO"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = 47 and reptype = 'SO' "
ElseIf Opt3_83.Value = True Then
vRepID = 211
vRepType = "Sale"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = 211 and reptype = 'Sale' "
End If

If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport3_81
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@StartDate;" & Date1 & " ;true"
        .ParameterFields(1) = "@EndDate;" & Date2 & " ;true"
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
DTP3_81 = Now
DTP3_82 = Now
End Sub
