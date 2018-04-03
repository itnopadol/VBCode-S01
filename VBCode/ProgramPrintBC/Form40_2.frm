VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form40_2 
   Caption         =   "รายงานสรุปค่าใช้จ่ายประจำวัน"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   Icon            =   "Form40_2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form40_2.frx":08CA
   ScaleHeight     =   8385
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport40_21 
      Left            =   2040
      Top             =   5280
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
   Begin VB.CommandButton CMD40_21 
      Caption         =   "พิมพ์รายงาน"
      Height          =   540
      Left            =   3600
      TabIndex        =   2
      Top             =   2550
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTP40_21 
      Height          =   315
      Left            =   2550
      TabIndex        =   1
      Top             =   1500
      Width           =   2415
      _ExtentX        =   4260
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
      Format          =   66584577
      CurrentDate     =   38041
   End
   Begin VB.Label LBL40_22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่"
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
      Left            =   1050
      TabIndex        =   3
      Top             =   1500
      Width           =   1365
   End
   Begin VB.Label LBL40_21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงานสรุปค่าใช้จ่ายประจำวัน"
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
      Height          =   465
      Left            =   2550
      TabIndex        =   0
      Top             =   300
      Width           =   7440
   End
End
Attribute VB_Name = "Form40_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD40_21_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim Date1 As Date
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription
Date1 = DTP40_21.Day & "/" & DTP40_21.Month & "/" & DTP40_21.Year

vRepID = 49
vRepType = "PM"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = '49' and reptype = 'PM' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport40_21
            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
            .ParameterFields(0) = "@Date1;" & Date1 & ";true"
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
DTP40_21 = Now
End Sub
