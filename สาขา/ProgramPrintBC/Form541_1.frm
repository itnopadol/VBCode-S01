VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form541_1 
   Caption         =   "หน้าดูรายงานชำระหนี้ประจำวัน"
   ClientHeight    =   8235
   ClientLeft      =   1815
   ClientTop       =   885
   ClientWidth     =   12000
   Icon            =   "Form541_1.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form541_1.frx":08CA
   ScaleHeight     =   8235
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport541_11 
      Left            =   3360
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame FRM541_11 
      BackColor       =   &H8000000E&
      Height          =   1890
      Left            =   2175
      TabIndex        =   3
      Top             =   975
      Width           =   2865
      Begin VB.OptionButton Opt541_14 
         BackColor       =   &H80000009&
         Caption         =   "โชว์เฉพาะ RC"
         Height          =   240
         Left            =   150
         TabIndex        =   8
         Top             =   1425
         Width           =   1965
      End
      Begin VB.OptionButton Opt541_13 
         BackColor       =   &H8000000E&
         Caption         =   "โชว์เฉพาะ RD"
         Height          =   240
         Left            =   150
         TabIndex        =   6
         Top             =   1035
         Width           =   2190
      End
      Begin VB.OptionButton Opt541_12 
         BackColor       =   &H8000000E&
         Caption         =   "ไม่เอา RD"
         Height          =   240
         Left            =   150
         TabIndex        =   5
         Top             =   675
         Width           =   2190
      End
      Begin VB.OptionButton Opt541_11 
         BackColor       =   &H8000000E&
         Caption         =   "โชว์ทั้งหมด"
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   2190
      End
   End
   Begin VB.CommandButton CMD541_11 
      Caption         =   "พิมพ์"
      Height          =   615
      Left            =   3525
      TabIndex        =   1
      Top             =   4275
      Width           =   1515
   End
   Begin MSComCtl2.DTPicker DTP541_11 
      Height          =   390
      Left            =   2175
      TabIndex        =   0
      Top             =   3150
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   688
      _Version        =   393216
      Format          =   69599233
      CurrentDate     =   38027
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงานชำระหนี้ประจำวัน"
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
      Height          =   615
      Left            =   2475
      TabIndex        =   9
      Top             =   225
      Width           =   7515
   End
   Begin VB.Label LBL541_11 
      BackStyle       =   0  'Transparent
      Caption         =   "เงื่อนไขการแสดง"
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
      Left            =   900
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label LBL541_12 
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
      Left            =   1500
      TabIndex        =   2
      Top             =   3150
      Width           =   690
   End
End
Attribute VB_Name = "Form541_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD541_11_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vType As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vDate As Date

On Error GoTo ErrDescription

If Opt541_11.Value = True Then
    Call PrintReportCashReceipt
ElseIf Opt541_12.Value = True Then
    Call PrintReportCashReceipt
ElseIf Opt541_13.Value = True Then
    Call PrintReportCashReceipt
ElseIf Opt541_14.Value = True Then
    Call PrintReportCashReceipt_RC
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub


Private Sub Form_Load()
DTP541_11 = Now
End Sub

Public Function PrintReportCashReceipt_RC()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vType As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vDate As Date

On Error GoTo ErrDescription

vType = 4
vRepID = 144
vRepType = "RE"
vDate = DTP541_11.Day & "/" & DTP541_11.Month & "/" & DTP541_11.Year

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "' and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport541_11
    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
    .ParameterFields(0) = "@Date;" & vDate & ";true"
    .ParameterFields(1) = "@Type;" & vType & ";true"
    .WindowState = crptMaximized
    .Destination = crptToWindow
    .Action = 1
End With
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Function
End If
End Function

Public Function PrintReportCashReceipt()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vType As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vDate As Date

On Error GoTo ErrDescription

If Opt541_11.Value = True Then
    vType = 1
ElseIf Opt541_12.Value = True Then
    vType = 2
ElseIf Opt541_13.Value = True Then
    vType = 3
ElseIf Opt541_14.Value = True Then
    vType = 4
End If

vRepID = 33
vRepType = "RE"
vDate = DTP541_11.Day & "/" & DTP541_11.Month & "/" & DTP541_11.Year

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "' and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport541_11
    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
    .ParameterFields(0) = "@Date;" & vDate & ";true"
    .ParameterFields(1) = "@Type;" & vType & ";true"
    .WindowState = crptMaximized
    .Destination = crptToWindow
    .Action = 1
End With
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Function
End If
End Function

