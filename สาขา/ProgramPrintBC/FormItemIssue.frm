VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FormItemIssue 
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormItemIssue.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1305
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
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
      Height          =   735
      Left            =   5445
      TabIndex        =   4
      Top             =   3465
      Width           =   2040
   End
   Begin MSComCtl2.DTPicker DTPDate2 
      Height          =   465
      Left            =   5445
      TabIndex        =   1
      Top             =   2520
      Width           =   2040
      _ExtentX        =   3598
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
      Format          =   61145089
      CurrentDate     =   40973
   End
   Begin MSComCtl2.DTPicker DTPDate1 
      Height          =   465
      Left            =   5445
      TabIndex        =   0
      Top             =   1800
      Width           =   2040
      _ExtentX        =   3598
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
      Format          =   61145089
      CurrentDate     =   40973
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่ :"
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
      Height          =   375
      Left            =   3285
      TabIndex        =   3
      Top             =   2565
      Width           =   2085
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่ :"
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
      Left            =   3285
      TabIndex        =   2
      Top             =   1845
      Width           =   2085
   End
End
Attribute VB_Name = "FormItemIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error Resume Next
Me.DTPDate1.Value = Now
Me.DTPDate2.Value = Now
End Sub

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vFromDate As String
Dim vToDate As String
Dim vRepID As Integer
Dim vRepType As String

On Error Resume Next

vFromDate = Day(Me.DTPDate1.Value) & "/" & Month(Me.DTPDate1.Value) & "/" & Year(Me.DTPDate1.Value)
vToDate = Day(Me.DTPDate2.Value) & "/" & Month(Me.DTPDate2.Value) & "/" & Year(Me.DTPDate2.Value)

vRepID = 475
vRepType = "IS"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal101
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocDate1;" & vFromDate & ";true"
        .ParameterFields(1) = "@vDocDate2;" & vToDate & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
End Sub
