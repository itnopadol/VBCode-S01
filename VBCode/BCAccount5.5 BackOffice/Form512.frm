VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form512 
   Caption         =   "รายงาน เคลื่อนไหวลูกหนี้ ตามช่วงเวลา"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form512.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1665
      Top             =   6030
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
   Begin VB.ComboBox CMB102 
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Text            =   "BCVAT48"
      Top             =   1275
      Width           =   2265
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์รายงาน"
      Height          =   465
      Left            =   3825
      TabIndex        =   3
      Top             =   4050
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker DTP102 
      Height          =   315
      Left            =   2475
      TabIndex        =   2
      Top             =   3375
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      Format          =   63700993
      CurrentDate     =   38435
   End
   Begin MSComCtl2.DTPicker DTP101 
      Height          =   315
      Left            =   2475
      TabIndex        =   1
      Top             =   2775
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      Format          =   63700993
      CurrentDate     =   38435
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   2475
      TabIndex        =   0
      Top             =   2100
      Width           =   2790
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกบริษัท"
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
      Left            =   1725
      TabIndex        =   8
      Top             =   1275
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Left            =   1500
      TabIndex        =   6
      Top             =   3375
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Left            =   1200
      TabIndex        =   5
      Top             =   2775
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า"
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
      Left            =   1725
      TabIndex        =   4
      Top             =   2100
      Width           =   540
   End
End
Attribute VB_Name = "Form512"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode1 As String
Dim vDate1 As Date, vDate2 As Date
Dim vReportName As String
Dim StrCount As Integer, vRepID As Integer

On Error GoTo Errdescription

If CMB102.Text = "" Then
    MsgBox "กรุณาเลือก บริษัทในการดูรายงานด้วยครับ"
    Exit Sub
Else
    If CMB102.Text = Trim("BCVAT47") Then
        vRepID = 15
    ElseIf CMB102.Text = Trim("BCVAT48") Then
        vRepID = 16
    End If
    
    StrCount = InStr(Trim(CMB101.Text), "/")
    vARCode1 = Trim(Left(CMB101.Text, StrCount - 1))
    vDate1 = DTP101.Day & "/" & DTP101.Month & "/" & DTP101.Year
    vDate2 = DTP102.Day & "/" & DTP102.Month & "/" & DTP102.Year
    
    vQuery = "select reportname from bcvat.dbo.bcreportname where repid = " & vRepID & "  and reptype = 'AR' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
    
    With Crystal101
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@ARCode;" & vARCode1 & ";true"
    .ParameterFields(2) = "@Start;" & vDate1 & ";true"
    .ParameterFields(3) = "@End;" & vDate2 & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    End With
End If

Errdescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

DTP101 = Now
DTP102 = Now

CMB102.AddItem Trim("BCVAT48")
CMB102.AddItem Trim("BCVAT47")

vQuery = "select code+'/'+name1 as arname from dbo.bcar order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB101.AddItem Trim(vRecordset.Fields("arname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close


End Sub
