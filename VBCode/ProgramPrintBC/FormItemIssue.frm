VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FormItemIssue 
   Caption         =   "�������§ҹ ���ԡ�Թ��һ�Ш��ѹ"
   ClientHeight    =   11010
   ClientLeft      =   2055
   ClientTop       =   450
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormItemIssue.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDPrint 
      Caption         =   "�������§ҹ"
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
      Left            =   4770
      TabIndex        =   4
      Top             =   2820
      Width           =   1845
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   960
      Top             =   6390
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
   Begin MSComCtl2.DTPicker DTPDate2 
      Height          =   375
      Left            =   4770
      TabIndex        =   2
      Top             =   2190
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   661
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
      Format          =   72613889
      CurrentDate     =   40852
   End
   Begin MSComCtl2.DTPicker DTPDate1 
      Height          =   375
      Left            =   4770
      TabIndex        =   1
      Top             =   1590
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   661
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
      Format          =   72613889
      CurrentDate     =   40852
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�֧�ѹ��� :"
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
      Left            =   3210
      TabIndex        =   3
      Top             =   2190
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�ҡ�ѹ��� :"
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
      Left            =   3270
      TabIndex        =   0
      Top             =   1590
      Width           =   1365
   End
End
Attribute VB_Name = "FormItemIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

vRepID = 510
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

Private Sub Form_Load()
On Error Resume Next
Me.DTPDate1.Value = Now
Me.DTPDate2.Value = Now
End Sub

