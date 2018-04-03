VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form51_1 
   Caption         =   "พิมพ์ รายงานใบรับวางบิลของพนักงานงานเก็บเงิน ตามช่วงวันที่"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form51_1.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDPrint 
      Caption         =   "พิมพ์รายงาน"
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
      Left            =   6525
      TabIndex        =   6
      Top             =   3690
      Width           =   1185
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   450
      Top             =   6300
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
   Begin VB.ComboBox CMBKeepMen 
      Height          =   315
      Left            =   5085
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
      Width           =   2625
   End
   Begin MSComCtl2.DTPicker DTP102 
      Height          =   330
      Left            =   5085
      TabIndex        =   4
      Top             =   2970
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   39387
   End
   Begin MSComCtl2.DTPicker DTP101 
      Height          =   330
      Left            =   5085
      TabIndex        =   3
      Top             =   2475
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   39387
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่ :"
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
      Height          =   330
      Left            =   3600
      TabIndex        =   2
      Top             =   3015
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่ :"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   2520
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อพนักงานเก็บเงิน :"
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
      Height          =   375
      Left            =   2970
      TabIndex        =   0
      Top             =   1800
      Width           =   1995
   End
End
Attribute VB_Name = "Form51_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vReportName As String
Dim vKeepMen As String
Dim vDocDate1 As String
Dim vDocDate2 As String
Dim StrCount As Integer

On Error GoTo ErrDescription

If Me.CMBKeepMen.Text <> "" Then
   StrCount = InStr(Trim(CMBKeepMen.Text), "/")
   vKeepMen = Trim(Left(CMBKeepMen.Text, StrCount - 1))
End If
vDocDate1 = Me.DTP101.Day & "/" & Me.DTP101.Month & "/" & Me.DTP101.Year
vDocDate2 = Me.DTP102.Day & "/" & Me.DTP102.Month & "/" & Me.DTP102.Year

vQuery = "exec dbo.USP_NP_SelectReportName 369,'AR' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vReportName = vRecordset.Fields("reportname").Value
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vKeepMoneyCode;" & vKeepMen & ";true "
.ParameterFields(1) = "@vDocDate1;" & vDocDate1 & ";true"
.ParameterFields(2) = "@vDocDate2;" & vDocDate2 & ";true "
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Me.DTP101.Value = Now
Me.DTP102.Value = Now

Call GetKeepMen
End Sub

Public Sub GetKeepMen()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

Me.CMBKeepMen.Clear
vQuery = "select * from dbo.vw_NP_KeepMoneyMenName "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   Me.CMBKeepMen.AddItem (vRecordset.Fields("keepmencodename").Value)
   vRecordset.MoveNext
   Wend
End If
vRecordset.Close

End Sub
