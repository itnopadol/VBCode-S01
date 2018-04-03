VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Form541_1 
   Caption         =   "หน้าดูรายงานชำระหนี้ประจำวัน"
   ClientHeight    =   8235
   ClientLeft      =   4380
   ClientTop       =   1605
   ClientWidth     =   12000
   Icon            =   "Form541_1.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form541_1.frx":08CA
   ScaleHeight     =   8235
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport541_12 
      Left            =   1740
      Top             =   6510
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
   Begin VB.CheckBox CBKeepMen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ดูตาม พนักงานเก็บเงิน มีเฉพาะ เงื่อนไข โชว์ทั้งหมด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3060
      TabIndex        =   14
      Top             =   3270
      Width           =   7545
   End
   Begin VB.ComboBox CMBKeepMoney2 
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
      Left            =   3060
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4320
      Width           =   7545
   End
   Begin VB.ComboBox CMBKeepMoney1 
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
      Left            =   3060
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3690
      Width           =   7545
   End
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
      Left            =   3060
      TabIndex        =   3
      Top             =   1260
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
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   2190
      End
   End
   Begin VB.CommandButton CMD541_11 
      Caption         =   "พิมพ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   5715
      Width           =   1515
   End
   Begin MSComCtl2.DTPicker DTP541_11 
      Height          =   390
      Left            =   3060
      TabIndex        =   0
      Top             =   4920
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   688
      _Version        =   393216
      Format          =   75366401
      CurrentDate     =   38027
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงพนักงานติดตามหนี้ :"
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
      Left            =   540
      TabIndex        =   13
      Top             =   4290
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากพนักงานติดตามหนี้ :"
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
      Left            =   510
      TabIndex        =   10
      Top             =   3720
      Width           =   2445
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
      Left            =   3060
      TabIndex        =   9
      Top             =   390
      Width           =   7515
   End
   Begin VB.Label LBL541_11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เงื่อนไขการแสดง"
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
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   1470
      Width           =   1725
   End
   Begin VB.Label LBL541_12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่"
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
      Height          =   315
      Left            =   2100
      TabIndex        =   2
      Top             =   4920
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
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String


DTP541_11 = Now

vQuery = "select  distinct isnull(b.keepmoneymencode,'') as keepcode,isnull(a.name,'') as keepname from dbo.bcsale a inner join dbo.bcar b on a.code = isnull(b.keepmoneymencode,'') and a.activestatus = 1  order by isnull(b.keepmoneymencode,'')"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Me.CMBKeepMoney1.AddItem Trim(vRecordset.Fields("keepcode").Value & "//" & vRecordset.Fields("keepname").Value)
        Me.CMBKeepMoney2.AddItem Trim(vRecordset.Fields("keepcode").Value & "//" & vRecordset.Fields("keepname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

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
Dim vKeepMen1 As String
Dim vKeepMen2 As String

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


If Me.CBKeepMen.Value = 0 Then
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
End If

If Me.CBKeepMen.Value = 1 And Me.CMBKeepMoney1.Text <> "" Then

    vKeepMen1 = Left(Me.CMBKeepMoney1.Text, InStr(1, Me.CMBKeepMoney1.Text, "//") - 1)

    If Me.CMBKeepMoney2.Text <> "" Then
        vKeepMen2 = Left(Me.CMBKeepMoney2.Text, InStr(1, Me.CMBKeepMoney2.Text, "//") - 1)
    Else
        vKeepMen2 = Left(Me.CMBKeepMoney1.Text, InStr(1, Me.CMBKeepMoney1.Text, "//") - 1)
    End If

If Me.CBKeepMen.Value = 1 Then
    vRepID = 514
    vRepType = "RE"
    vDate = DTP541_11.Day & "/" & DTP541_11.Month & "/" & DTP541_11.Year
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "' and reptype = '" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        With CrystalReport541_12
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@Date;" & vDate & ";true"
        .ParameterFields(1) = "@Type;" & vType & ";true"
        .ParameterFields(2) = "@KeepMen1;" & vKeepMen1 & ";true"
        .ParameterFields(3) = "@KeepMen2;" & vKeepMen2 & ";true"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With
    End If
    vRecordset.Close
End If

ElseIf Me.CBKeepMen.Value = 1 And Me.CMBKeepMoney1.Text = "" Then
    MsgBox "กรุณาเลือก ข้อมูลของพนักงานติดตามหนี้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
    Me.CMBKeepMoney1.SetFocus
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Function
End If
End Function

