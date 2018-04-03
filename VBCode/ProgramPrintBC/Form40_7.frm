VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Form40_7 
   Caption         =   "รายงาน การตั้งหนี้เอกสารซื้อสินค้า"
   ClientHeight    =   8070
   ClientLeft      =   4440
   ClientTop       =   2595
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form40_7.frx":0000
   ScaleHeight     =   8070
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPDocDate2 
      Height          =   435
      Left            =   6540
      TabIndex        =   4
      Top             =   1560
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   767
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
      Format          =   72613889
      CurrentDate     =   40883
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   2025
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
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6540
      TabIndex        =   2
      Top             =   2220
      Width           =   2040
   End
   Begin MSComCtl2.DTPicker DTPDocDate1 
      Height          =   420
      Left            =   4005
      TabIndex        =   1
      Top             =   1575
      Width           =   2040
      _ExtentX        =   3598
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
      Format          =   72613889
      CurrentDate     =   40821
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึง"
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
      Height          =   345
      Left            =   6150
      TabIndex        =   3
      Top             =   1620
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่ทำเอกสาร :"
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
      Left            =   1260
      TabIndex        =   0
      Top             =   1620
      Width           =   2535
   End
End
Attribute VB_Name = "Form40_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.DTPDocDate1.Value = Now
Me.DTPDocDate2.Value = Now
End Sub


Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocdate1 As String
Dim vDocdate2 As String
Dim vRepID  As Integer
Dim vRepType  As String


On Error Resume Next

vDocdate1 = Me.DTPDocDate1.Day & "/" & Me.DTPDocDate1.Month & "/" & Me.DTPDocDate1.Year
vDocdate2 = Me.DTPDocDate2.Day & "/" & Me.DTPDocDate2.Month & "/" & Me.DTPDocDate2.Year

vRepID = 508
vRepType = "AP"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal101
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocDate1;" & vDocdate1 & ";true"
        .ParameterFields(1) = "@vDocDate2;" & vDocdate2 & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
    

End Sub

