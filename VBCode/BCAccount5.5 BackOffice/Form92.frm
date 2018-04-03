VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form92 
   Caption         =   "รายงาน สมุดรายวัน"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form92.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalGLBook 
      Left            =   1170
      Top             =   6525
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
   Begin VB.ComboBox CMBSelectData 
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
      Left            =   3150
      TabIndex        =   9
      Top             =   975
      Width           =   3240
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
      Left            =   3150
      TabIndex        =   8
      Top             =   1800
      Width           =   3240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "พิมพ์รายงาน"
      Height          =   615
      Left            =   4950
      TabIndex        =   3
      Top             =   4350
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   390
      Left            =   3150
      TabIndex        =   2
      Top             =   3600
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   688
      _Version        =   393216
      Format          =   68026369
      CurrentDate     =   38101
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   390
      Left            =   3150
      TabIndex        =   1
      Top             =   3075
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   688
      _Version        =   393216
      Format          =   68026369
      CurrentDate     =   38101
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3150
      TabIndex        =   0
      Top             =   2325
      Width           =   3240
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกดูข้อมูล"
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
      Left            =   2175
      TabIndex        =   10
      Top             =   975
      Width           =   840
   End
   Begin VB.Label Label5 
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
      Left            =   1800
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   1800
      TabIndex        =   6
      Top             =   3075
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "กลุ่มเอกสาร"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   2325
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   2250
      TabIndex        =   4
      Top             =   1800
      Width           =   765
   End
End
Attribute VB_Name = "Form92"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Dim vData As String

On Error GoTo Errdescription

vData = Trim(CMBSelectData.Text)
If vData = "ข้อมูลหลังร้าน46" Then
Call PrintGLBook46
ElseIf vData = "ข้อมูลหลังร้าน47" Then
Call PrintGLBook47
ElseIf vData = "ข้อมูลหลังร้าน48" Then
Call PrintGLBook48
ElseIf vData = "ข้อมูลหลังร้าน" Then
Call PrintGLBook
Else
MsgBox "กรุณาเลือกข้อมูลที่จะดูรายงานด้วยนะครับ", vbInformation, "ข้อความแจ้งเตือน"
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

vQuery = "select distinct code,name from bcglbook order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Combo1.AddItem Trim(vRecordset.Fields("code").Value) & " = " & Trim(vRecordset.Fields("name").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

CMBSelectData.AddItem "ข้อมูลหลังร้าน"
CMBSelectData.AddItem "ข้อมูลหลังร้าน46"
CMBSelectData.AddItem "ข้อมูลหลังร้าน47"
CMBSelectData.AddItem "ข้อมูลหลังร้าน48"

Me.DTP1.Value = Now
Me.DTP2.Value = Now
End Sub

Public Sub PrintGLBook46()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepType As String, vReportName As String
Dim vRepID As Integer, StrCount As Integer
Dim vAccount As String
Dim Date1 As Date, Date2 As Date
Dim DocNo1 As String, DocNo2 As String

StrCount = InStr(Trim(Combo1.Text), "=")
vAccount = Trim(Left(Combo1.Text, StrCount - 1))
Date1 = DTP1.Day & "/" & DTP1.Month & "/" & DTP1.Year
Date2 = DTP2.Day & "/" & DTP2.Month & "/" & DTP2.Year
vRepType = "GL"
vRepID = 6
DocNo1 = Trim(Text2.Text)
vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    With CrystalGLBook
        .ReportFileName = vReportName & ".rpt"
        
        .ParameterFields(0) = "@BookCode;" & vAccount & ";true"
        .ParameterFields(1) = "@FDate;" & Date1 & ";true"
        .ParameterFields(2) = "@LDate;" & Date2 & ";true"
        .ParameterFields(3) = "@DocNO;" & DocNo1 & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
End Sub

Public Sub PrintGLBook47()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepType As String, vReportName As String
Dim vRepID As Integer, StrCount As Integer
Dim vAccount As String
Dim Date1 As Date, Date2 As Date
Dim DocNo1 As String, DocNo2 As String

StrCount = InStr(Trim(Combo1.Text), "=")
vAccount = Trim(Left(Combo1.Text, StrCount - 1))
Date1 = DTP1.Day & "/" & DTP1.Month & "/" & DTP1.Year
Date2 = DTP2.Day & "/" & DTP2.Month & "/" & DTP2.Year
vRepType = "GL"
vRepID = 19
DocNo1 = Trim(Text2.Text)
vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    With CrystalGLBook
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@BookCode;" & vAccount & ";true"
        .ParameterFields(1) = "@FDate;" & Date1 & ";true"
        .ParameterFields(2) = "@LDate;" & Date2 & ";true"
        .ParameterFields(3) = "@DocNO;" & DocNo1 & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
End Sub

Public Sub PrintGLBook48()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepType As String, vReportName As String
Dim vRepID As Integer, StrCount As Integer
Dim vAccount As String
Dim Date1 As Date, Date2 As Date
Dim DocNo1 As String, DocNo2 As String

StrCount = InStr(Trim(Combo1.Text), "=")
vAccount = Trim(Left(Combo1.Text, StrCount - 1))
Date1 = DTP1.Day & "/" & DTP1.Month & "/" & DTP1.Year
Date2 = DTP2.Day & "/" & DTP2.Month & "/" & DTP2.Year
vRepType = "GL"
vRepID = 24
DocNo1 = Trim(Text2.Text)
vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    With CrystalGLBook
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@BookCode;" & vAccount & ";true"
        .ParameterFields(1) = "@FDate;" & Date1 & ";true"
        .ParameterFields(2) = "@LDate;" & Date2 & ";true"
        .ParameterFields(3) = "@DocNO;" & DocNo1 & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
End Sub


Public Sub PrintGLBook()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepType As String, vReportName As String
Dim vRepID As Integer, StrCount As Integer
Dim vAccount As String
Dim Date1 As Date, Date2 As Date
Dim DocNo1 As String, DocNo2 As String

StrCount = InStr(Trim(Combo1.Text), "=")
vAccount = Trim(Left(Combo1.Text, StrCount - 1))
Date1 = DTP1.Day & "/" & DTP1.Month & "/" & DTP1.Year
Date2 = DTP2.Day & "/" & DTP2.Month & "/" & DTP2.Year
vRepType = "GL"
vRepID = 24
DocNo1 = Trim(Text2.Text)
vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    With CrystalGLBook
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@BookCode;" & vAccount & ";true"
        .ParameterFields(1) = "@FDate;" & Date1 & ";true"
        .ParameterFields(2) = "@LDate;" & Date2 & ";true"
        .ParameterFields(3) = "@DocNO;" & DocNo1 & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
End Sub




