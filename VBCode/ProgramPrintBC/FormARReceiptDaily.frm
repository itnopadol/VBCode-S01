VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FormARReceiptRemainDaily 
   Caption         =   "รายงาน เอกสารขายที่ค้างรับชำระจากใบวางบิล"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormARReceiptDaily.frx":0000
   ScaleHeight     =   11490
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1125
      Top             =   6660
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
      Height          =   780
      Left            =   3285
      TabIndex        =   2
      Top             =   2385
      Width           =   1860
   End
   Begin MSComCtl2.DTPicker DTPDocDate 
      Height          =   465
      Left            =   3015
      TabIndex        =   1
      Top             =   1575
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   820
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
      CurrentDate     =   40806
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประจำวันที่ :"
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
      Height          =   420
      Left            =   855
      TabIndex        =   0
      Top             =   1620
      Width           =   2040
   End
End
Attribute VB_Name = "FormARReceiptRemainDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.DTPDocDate.Value = Now
End Sub


Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocdate As String
Dim vRepID  As Integer
Dim vRepType  As String


On Error Resume Next

vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year

vRepID = 505
vRepType = "AR"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal101
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocDate;" & vDocdate & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
    

End Sub


