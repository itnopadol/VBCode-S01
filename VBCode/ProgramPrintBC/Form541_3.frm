VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Form541_3 
   Caption         =   "รายงาน รับเงินประจำวันฝั่งรับชำระ"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form541_3.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1260
      Top             =   5670
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
      Caption         =   "พิมพ์รายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2925
      TabIndex        =   2
      Top             =   2655
      Width           =   1680
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   330
      Left            =   2925
      TabIndex        =   1
      Top             =   1800
      Width           =   1680
      _ExtentX        =   2963
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
      Format          =   62652417
      CurrentDate     =   40317
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานประจำวันที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   675
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
End
Attribute VB_Name = "Form541_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDPrint_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType As String
Dim vDate As Date

On Error GoTo ErrDescription

vRepID = 481
vRepType = "RE"
vDate = DTPDate.Day & "/" & DTPDate.Month & "/" & DTPDate.Year

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal101
    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
    .ParameterFields(0) = "@vDocDate;" & vDate & ";true"
    .WindowState = crptMaximized
    .Destination = crptToWindow
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
