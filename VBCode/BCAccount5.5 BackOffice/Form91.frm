VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form91 
   Caption         =   "รายงานแยกประเภท"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form91.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalGLBalance 
      Left            =   2700
      Top             =   5760
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
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3975
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
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
      Left            =   3975
      TabIndex        =   7
      Top             =   1650
      Width           =   3390
   End
   Begin VB.CommandButton Command1 
      Caption         =   "พิมพ์รายงาน"
      Height          =   465
      Left            =   4875
      TabIndex        =   2
      Top             =   3675
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   390
      Left            =   3975
      TabIndex        =   1
      Top             =   2925
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   688
      _Version        =   393216
      Format          =   63700993
      CurrentDate     =   38101
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   390
      Left            =   3975
      TabIndex        =   0
      Top             =   2325
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   688
      _Version        =   393216
      Format          =   63700993
      CurrentDate     =   38101
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกบริษัท"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Top             =   1200
      Width           =   1290
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงานแยกประเภท"
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
      TabIndex        =   6
      Top             =   300
      Width           =   7215
   End
   Begin VB.Label Label3 
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
      Left            =   3075
      TabIndex        =   5
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label Label2 
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
      Left            =   3075
      TabIndex        =   4
      Top             =   2325
      Width           =   690
   End
   Begin VB.Label Label1 
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
      Height          =   390
      Left            =   3075
      TabIndex        =   3
      Top             =   1650
      Width           =   765
   End
End
Attribute VB_Name = "Form91"
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

On Error GoTo Errdescription
StrCount = InStr(Trim(Combo1.Text), "=")
vAccount = Trim(Left(Combo1.Text, StrCount - 1))
Date1 = DTP1.Day & "/" & DTP1.Month & "/" & DTP1.Year
Date2 = DTP2.Day & "/" & DTP2.Month & "/" & DTP2.Year
If vAccount = Trim("202201") Or vAccount = Trim("2132 - 30 - 0") Then
    vRepID = 7
ElseIf Combo2.Text = Trim("NPVAT46") Then
    vRepID = 0
ElseIf Combo2.Text = Trim("NPVAT47") Then
    vRepID = 5
ElseIf Combo2.Text = Trim("NPVAT48") Then
    vRepID = 23
ElseIf Combo2.Text = Trim("NPVAT") Then
    vRepID = 30
End If
vRepType = "GL"
vQuery = "select reportname from bcvat.dbo.bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
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

vQuery = "select distinct code,name1 from dbo.bcchartofaccount order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Combo1.AddItem Trim(vRecordset.Fields("code").Value) & " = " & Trim(vRecordset.Fields("name1").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

Combo2.AddItem Trim("NPVAT")
Combo2.AddItem Trim("NPVAT46")
Combo2.AddItem Trim("NPVAT47")
Combo2.AddItem Trim("NPVAT48")

Me.DTP1.Value = Now
Me.DTP2.Value = Now

End Sub
