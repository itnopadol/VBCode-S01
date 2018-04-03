VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form51_1 
   Caption         =   "พิมพ์ รายงานใบรับวางบิลของพนักงานงานเก็บเงิน ตามช่วงวันที่"
   ClientHeight    =   9000
   ClientLeft      =   4080
   ClientTop       =   645
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form51_1.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal102 
      Left            =   900
      Top             =   7560
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
   Begin VB.TextBox TXTDocNo2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5085
      TabIndex        =   10
      Top             =   4410
      Width           =   2400
   End
   Begin VB.TextBox TXTDocNo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5085
      TabIndex        =   9
      Top             =   3735
      Width           =   2400
   End
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
      Height          =   645
      Left            =   5085
      TabIndex        =   6
      Top             =   5085
      Width           =   1680
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5085
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
      Width           =   3165
   End
   Begin MSComCtl2.DTPicker DTP102 
      Height          =   330
      Left            =   5085
      TabIndex        =   4
      Top             =   3105
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
      Format          =   61079553
      CurrentDate     =   39387
   End
   Begin MSComCtl2.DTPicker DTP101 
      Height          =   330
      Left            =   5085
      TabIndex        =   3
      Top             =   2475
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
      Format          =   61079553
      CurrentDate     =   39387
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงเลขที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3510
      TabIndex        =   8
      Top             =   4455
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากเลขที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3555
      TabIndex        =   7
      Top             =   3780
      Width           =   1410
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Top             =   3105
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Size            =   9.75
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
Dim vDocdate1 As String
Dim vDocdate2 As String
Dim StrCount As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vDocNo1 As String
Dim vDocNo2 As String

On Error GoTo ErrDescription

If Me.CMBKeepMen.Text <> "" Then
   StrCount = InStr(Trim(CMBKeepMen.Text), "/")
   vKeepMen = Trim(Left(CMBKeepMen.Text, StrCount - 1))
End If
vDocdate1 = Me.DTP101.Day & "/" & Me.DTP101.Month & "/" & Me.DTP101.Year
vDocdate2 = Me.DTP102.Day & "/" & Me.DTP102.Month & "/" & Me.DTP102.Year

vDocNo1 = Me.TXTDocNo1.Text
vDocNo2 = Me.TXTDocNo2.Text

If vDocNo2 = "" Then
vDocNo2 = vDocNo1
End If

If vKeepMen = "" And vDocNo1 = "" And vDocNo2 = "" Then
vRepID = 488
ElseIf vKeepMen <> "" And vDocNo1 = "" And vDocNo2 = "" Then
vRepID = 369
ElseIf vDocNo1 <> "" And vDocNo2 <> "" Then
vRepID = 489
End If

vRepType = "AR"

If vRepID = 488 Or vRepID = 369 Then
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vReportName = vRecordset.Fields("reportname").Value
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vKeepMoneyCode;" & vKeepMen & ";true "
.ParameterFields(1) = "@vDocDate1;" & vDocdate1 & ";true"
.ParameterFields(2) = "@vDocDate2;" & vDocdate2 & ";true "
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

ElseIf vRepID = 489 Then

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vReportName = vRecordset.Fields("reportname").Value
End If
vRecordset.Close

With Crystal102
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vKeepMoneyCode;" & vKeepMen & ";true "
.ParameterFields(1) = "@vDocDate1;" & vDocdate1 & ";true"
.ParameterFields(2) = "@vDocDate2;" & vDocdate2 & ";true "
.ParameterFields(3) = "@vBegNo;" & vDocNo1 & ";true"
.ParameterFields(4) = "@vEndNo;" & vDocNo2 & ";true "
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

End If

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
