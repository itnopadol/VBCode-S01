VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form94 
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form94.frx":0000
   ScaleHeight     =   8385
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalGLBalance 
      Left            =   1920
      Top             =   6240
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
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "พิมพ์รายงาน"
      Height          =   540
      Left            =   4575
      TabIndex        =   7
      Top             =   3975
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   390
      Left            =   3225
      TabIndex        =   3
      Top             =   3150
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   688
      _Version        =   393216
      Format          =   69599233
      CurrentDate     =   38159
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   390
      Left            =   3225
      TabIndex        =   2
      Top             =   2550
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   688
      _Version        =   393216
      Format          =   69599233
      CurrentDate     =   38159
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3225
      TabIndex        =   1
      Top             =   1800
      Width           =   3465
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
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
      Left            =   2325
      TabIndex        =   6
      Top             =   3150
      Width           =   540
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่"
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
      Left            =   2325
      TabIndex        =   5
      Top             =   2550
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสบัญชี"
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
      Left            =   2325
      TabIndex        =   4
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงาน แยกประเภท"
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
      Height          =   465
      Left            =   2625
      TabIndex        =   0
      Top             =   225
      Width           =   7365
   End
End
Attribute VB_Name = "Form94"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepType As String, vReportName As String
Dim vRepID As Integer, StrCount As Integer
Dim vAccount As String
Dim Date1 As Date, Date2 As Date

On Error GoTo ErrDescription
StrCount = InStr(Trim(Combo1.Text), "=")
vAccount = Trim(Left(Combo1.Text, StrCount - 1))
Date1 = DTP1.Day & "/" & DTP1.Month & "/" & DTP1.Year
Date2 = DTP2.Day & "/" & DTP2.Month & "/" & DTP2.Year
vRepID = 149
vRepType = "GL"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    With CrystalGLBalance
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@Account;" & vAccount & ";true"
        .ParameterFields(1) = "@StartDate;" & Date1 & ";true"
        .ParameterFields(2) = "@EndDate;" & Date2 & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
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

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

vQuery = "select distinct code,name1 from bcchartofaccount order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Combo1.AddItem Trim(vRecordset.Fields("code").Value) & " = " & Trim(vRecordset.Fields("name1").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub

