VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Form44 
   Caption         =   "รายงาน การทำใบวางบิลประจำวัน ฝั่งเจ้าหนี้"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form44.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDPrint 
      Caption         =   "พิมพ์รายงาน"
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
      Left            =   2340
      TabIndex        =   2
      Top             =   2115
      Width           =   1815
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   855
      Top             =   6390
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
   Begin MSComCtl2.DTPicker DTPDocDate 
      Height          =   420
      Left            =   2340
      TabIndex        =   1
      Top             =   1485
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20774913
      CurrentDate     =   40611
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประจำวันที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1530
      Width           =   1500
   End
End
Attribute VB_Name = "Form44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDPrint_Click()
Dim vDocDate As String
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType As String
Dim vReportName As String

On Error Resume Next

vDocDate = Day(Me.DTPDocDate.Value) & "/" & Month(Me.DTPDocDate.Value) & "/" & Year(Me.DTPDocDate.Value)

vRepID = 501
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
    .ParameterFields(0) = "@vDocDate;" & vDocDate & ";true"
    .ParameterFields(1) = "@vUserID;" & vUserID & ";true"
    .Action = 1
End With

End Sub

Private Sub Form_Load()
Me.DTPDocDate.Value = Now
End Sub
