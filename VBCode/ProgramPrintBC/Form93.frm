VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form93 
   Caption         =   "รายงานงบทดลอง"
   ClientHeight    =   8385
   ClientLeft      =   5235
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form93.frx":0000
   ScaleHeight     =   8385
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport93 
      Left            =   1320
      Top             =   6000
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
   Begin VB.OptionButton Option3 
      BackColor       =   &H8000000E&
      Caption         =   "ยอด Net"
      Height          =   315
      Left            =   900
      TabIndex        =   8
      Top             =   2475
      Width           =   2040
   End
   Begin VB.ComboBox CMBYear 
      Height          =   315
      Left            =   4500
      TabIndex        =   7
      Top             =   2100
      Width           =   1740
   End
   Begin VB.ComboBox CMBMonth 
      Height          =   315
      Left            =   4500
      TabIndex        =   6
      Top             =   1500
      Width           =   1740
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H8000000E&
      Caption         =   "แบบสรุปตามปี"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   900
      TabIndex        =   5
      Top             =   2025
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "แบบรายละเอียด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   900
      TabIndex        =   4
      Top             =   1500
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.CommandButton CMD931 
      Caption         =   "ดูรายงาน"
      Height          =   540
      Left            =   4950
      TabIndex        =   0
      Top             =   2850
      Width           =   1365
   End
   Begin VB.Label LBL934 
      BackStyle       =   0  'Transparent
      Caption         =   "ปี"
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
      Left            =   4275
      TabIndex        =   3
      Top             =   2100
      Width           =   165
   End
   Begin VB.Label LBL933 
      BackStyle       =   0  'Transparent
      Caption         =   "งวดที่"
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
      Left            =   3975
      TabIndex        =   2
      Top             =   1500
      Width           =   465
   End
   Begin VB.Label LBL931 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงานงบทดลอง"
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
      Height          =   540
      Left            =   2475
      TabIndex        =   1
      Top             =   300
      Width           =   7515
   End
End
Attribute VB_Name = "Form93"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD931_Click()
On Error GoTo ErrDescription

If Option1.Value = True Then
    Call PrintGLTrailBalance_Details
ElseIf Option2.Value = True Then
    Call PrintGLTrailBalance1
    Call PrintGLTrailBalance2
ElseIf Option3.Value = True Then
    Call PrintGLTrailBalanceNET1
    Call PrintGLTrailBalanceNET2
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vQuery As String, vReportName As String
Dim vRecordset As New ADODB.Recordset

CMBMonth.AddItem Trim("1")
CMBMonth.AddItem Trim("2")
CMBMonth.AddItem Trim("3")
CMBMonth.AddItem Trim("4")
CMBMonth.AddItem Trim("5")
CMBMonth.AddItem Trim("6")
CMBMonth.AddItem Trim("7")
CMBMonth.AddItem Trim("8")
CMBMonth.AddItem Trim("9")
CMBMonth.AddItem Trim("10")
CMBMonth.AddItem Trim("11")
CMBMonth.AddItem Trim("12")

CMBYear.AddItem Trim("2004")
CMBYear.AddItem Trim("2005")
CMBYear.AddItem Trim("2006")
CMBYear.AddItem Trim("2007")

End Sub


Public Function PrintGLTrailBalance_Details()
Dim vQuery As String, vReportName As String
Dim vRecordset As New ADODB.Recordset
Dim vMonth As String, vYear As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If CMBYear.Text = "2004" Then
    vRepID = 247
ElseIf CMBYear.Text = "2005" Then
    vRepID = 145
ElseIf CMBYear.Text = "2006" Then
    vRepID = 145
ElseIf CMBYear.Text = "2007" Then
    vRepID = 360
End If

vMonth = Trim(CMBMonth.Text)
vYear = Trim(CMBYear.Text)
vRepType = "GL"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)

    With CrystalReport93
            .ReportFileName = vReportName & ".rpt"
            .ParameterFields(0) = "@month;" & vMonth & ";true"
            .ParameterFields(1) = "@year;" & vYear & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
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
Public Function PrintGLTrailBalance1()
Dim vQuery As String, vReportName As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If CMBYear.Text = "2004" Then
    vRepID = 248
ElseIf CMBYear.Text = "2005" Then
    vRepID = 146
ElseIf CMBYear.Text = "2006" Then
    vRepID = 146
ElseIf CMBYear.Text = "2007" Then
    vRepID = 146
End If

vYear = Trim(CMBYear.Text)
vRepType = "GL"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    With CrystalReport93
            .ReportFileName = vReportName & ".rpt"
            .ParameterFields(0) = "@year;" & vYear & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
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
Public Function PrintGLTrailBalance2()
Dim vQuery As String, vReportName As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If CMBYear.Text = "2004" Then
    vRepID = 249
ElseIf CMBYear.Text = "2005" Then
    vRepID = 147
ElseIf CMBYear.Text = "2006" Then
    vRepID = 147
ElseIf CMBYear.Text = "2007" Then
    vRepID = 147
End If

vYear = Trim(CMBYear.Text)
vRepType = "GL"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    With CrystalReport93
            .ReportFileName = vReportName & ".rpt"
            .ParameterFields(0) = "@year;" & vYear & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
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

Public Function PrintGLTrailBalanceNET1()
Dim vQuery As String, vReportName As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If CMBYear.Text = "2004" Then
    vRepID = 250
ElseIf CMBYear.Text = "2005" Then
    vRepID = 162
ElseIf CMBYear.Text = "2006" Then
    vRepID = 162
End If

vYear = Trim(CMBYear.Text)
vRepType = "GL"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    With CrystalReport93
            .ReportFileName = vReportName & ".rpt"
            .ParameterFields(0) = "@year;" & vYear & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
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
Public Function PrintGLTrailBalanceNET2()
Dim vQuery As String, vReportName As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If CMBYear.Text = "2004" Then
    vRepID = 251
ElseIf CMBYear.Text = "2005" Then
    vRepID = 163
ElseIf CMBYear.Text = "2006" Then
    vRepID = 163
End If

vYear = Trim(CMBYear.Text)
vRepType = "GL"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    With CrystalReport93
            .ReportFileName = vReportName & ".rpt"
            .ParameterFields(0) = "@year;" & vYear & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
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

