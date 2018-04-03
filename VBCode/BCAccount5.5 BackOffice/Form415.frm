VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form415 
   Caption         =   "รายงาน ยอดเจ้าหนี้ประจำเดือน"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form415.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1305
      Top             =   5490
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
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "ดูรายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5220
      TabIndex        =   2
      Top             =   2700
      Width           =   1185
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   330
      Left            =   4500
      TabIndex        =   1
      Top             =   1755
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   393216
      Format          =   63700993
      CurrentDate     =   38812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ดูรายงาน ณ วันที่"
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
      Left            =   3015
      TabIndex        =   0
      Top             =   1755
      Width           =   1500
   End
End
Attribute VB_Name = "Form415"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vDate As Date

On Error GoTo Errdescription

vQuery = "select reportname from bcvat.dbo.bcreportname where repid = 25 and reptype = 'AP' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

vDate = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
With Crystal101
    .ReportFileName = Trim(vReportName & ".rpt")
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .ParameterFields(0) = "@AtDate;" & vDate & ";true"
    .Action = 1
End With

Errdescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_Load()
DTPicker101 = Now
End Sub
