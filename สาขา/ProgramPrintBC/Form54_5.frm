VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form54_5 
   Caption         =   "Form1"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form54_5.frx":0000
   ScaleHeight     =   8325
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport54_5 
      Left            =   2520
      Top             =   6120
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
      Left            =   4350
      TabIndex        =   8
      Top             =   3900
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   390
      Left            =   3000
      TabIndex        =   3
      Top             =   3225
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   688
      _Version        =   393216
      Format          =   69795841
      CurrentDate     =   38302
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   390
      Left            =   3000
      TabIndex        =   2
      Top             =   2700
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   688
      _Version        =   393216
      Format          =   69795841
      CurrentDate     =   38302
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   3000
      TabIndex        =   1
      Top             =   1950
      Width           =   2790
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
      Left            =   3000
      TabIndex        =   0
      Top             =   1425
      Width           =   2790
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงาน เช็คตามลูกค้า"
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
      Left            =   2475
      TabIndex        =   9
      Top             =   225
      Width           =   7365
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
      Height          =   390
      Left            =   2250
      TabIndex        =   7
      Top             =   3225
      Width           =   690
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
      Height          =   390
      Left            =   2250
      TabIndex        =   6
      Top             =   2700
      Width           =   690
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงลูกหนี้"
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
      Left            =   2250
      TabIndex        =   5
      Top             =   1950
      Width           =   690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "จากลูกหนี้"
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
      Left            =   2250
      TabIndex        =   4
      Top             =   1425
      Width           =   765
   End
End
Attribute VB_Name = "Form54_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode1 As String, vARCode2 As String
Dim vDate1 As Date, vDate2 As Date
Dim vReportName As String
Dim StrCount As Integer, StrCount1 As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

StrCount = InStr(Trim(Combo1.Text), "/")
StrCount1 = InStr(Trim(Combo2.Text), "/")
vARCode1 = Trim(Left(Combo1.Text, StrCount - 1))
vARCode2 = Trim(Left(Combo2.Text, StrCount1 - 1))
vDate1 = DTPicker1.Day & "/" & DTPicker1.Month & "/" & DTPicker1.Year + 543
vDate2 = DTPicker2.Day & "/" & DTPicker2.Month & "/" & DTPicker2.Year + 543

vRepID = 187
vRepType = "AR"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 187 and reptype = 'AR' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With CrystalReport54_5
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@ARCode1;" & vARCode1 & ";true"
.ParameterFields(1) = "@ARCode2;" & vARCode2 & ";true"
.ParameterFields(2) = "@Date1;" & vDate1 & ";true"
.ParameterFields(3) = "@Date2;" & vDate2 & ";true"
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
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

DTPicker1 = Now
DTPicker2 = Now

vQuery = "select code+'/'+name1 as arname from bcnp.dbo.bcar order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Combo1.AddItem Trim(vRecordset.Fields("arname").Value)
        Combo2.AddItem Trim(vRecordset.Fields("arname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub
