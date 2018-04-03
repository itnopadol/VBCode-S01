VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form93 
   Caption         =   "รายงาน งบทดลอง"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form93.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport93 
      Left            =   1125
      Top             =   6705
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
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1485
      Top             =   5310
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
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "แสดงยอดคงเหลือสิ้นงวด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   810
      TabIndex        =   9
      Top             =   1935
      Width           =   2040
   End
   Begin VB.ComboBox CMBCompany 
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
      Left            =   5085
      TabIndex        =   8
      Top             =   1200
      Width           =   2340
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "แบบสรุปแยกเดือนทั้งปี"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   825
      TabIndex        =   7
      Top             =   1575
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "แบบละเอียด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   825
      TabIndex        =   6
      Top             =   1200
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.CommandButton CMD931 
      Caption         =   "พิมพ์รายงาน"
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
      Left            =   5985
      TabIndex        =   4
      Top             =   3300
      Width           =   1440
   End
   Begin VB.ComboBox CMBYear 
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
      Left            =   5085
      TabIndex        =   1
      Top             =   2625
      Width           =   2340
   End
   Begin VB.ComboBox CMBMonth 
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
      Left            =   5085
      TabIndex        =   0
      Top             =   1950
      Width           =   2340
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงาน งบทดลอง"
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
      Left            =   2625
      TabIndex        =   5
      Top             =   300
      Width           =   7365
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ของปี :"
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
      Left            =   4410
      TabIndex        =   3
      Top             =   2610
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "งวดที่ :"
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
      Left            =   4410
      TabIndex        =   2
      Top             =   1950
      Width           =   615
   End
End
Attribute VB_Name = "Form93"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMBCompany_Click()
Dim vQuery As String, vReportName As String
Dim vRecordset As New ADODB.Recordset
Dim vCompany As String


vCompany = Trim(CMBCompany.Text)

CMBMonth.Clear
vQuery = "select distinct month(docdate) as Month1  from " & vCompany & ".dbo.bctrans order by month(docdate) "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMBMonth.AddItem Trim(vRecordset.Fields("Month1").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

CMBYear.Clear
vQuery = "select distinct year(docdate) as Year1  from " & vCompany & ".dbo.bctrans order by year(docdate) "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMBYear.AddItem Trim(vRecordset.Fields("Year1").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub

Private Sub CMD931_Click()
Dim vQuery As String, vReportName As String
Dim vRecordset As New ADODB.Recordset
Dim vMonth As String, vYear As String, vCom As String

On Error GoTo Errdescription

If CMBMonth.Text <> "" And CMBYear.Text <> "" And CMBCompany.Text <> "" Then
vMonth = Trim(CMBMonth.Text)
vYear = Trim(CMBYear.Text)
vCom = Trim(CMBCompany.Text)

If vCom = "BCVAT" Then
   If Option1.Value = True Then
      vQuery = "select reportname from bcreportname where repid = '8' and reptype = 'GL' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          vReportName = Trim(vRecordset.Fields("reportname").Value)
      
          With Crystal101
                  .ReportFileName = vReportName & ".rpt"
                  .ParameterFields(0) = "@month;" & vMonth & ";true"
                  .ParameterFields(1) = "@year;" & vYear & ";true"
                  .Destination = crptToWindow
                  .WindowState = crptMaximized
                  .Action = 1
          End With
      End If
      vRecordset.Close
   ElseIf Option2.Value = True Then
      vQuery = "select reportname from bcreportname where repid = '9' and reptype = 'GL' "
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
   
      vQuery = "select reportname from bcreportname where repid = '10' and reptype = 'GL' "
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
   ElseIf Option3.Value = True Then
      vQuery = "select reportname from bcreportname where repid = '29' and reptype = 'GL' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          vReportName = Trim(vRecordset.Fields("reportname").Value)
      
          With Crystal101
                  .ReportFileName = vReportName & ".rpt"
                  .ParameterFields(0) = "@month;" & vMonth & ";true"
                  .ParameterFields(1) = "@year;" & vYear & ";true"
                  .Destination = crptToWindow
                  .WindowState = crptMaximized
                  .Action = 1
          End With
      End If
      vRecordset.Close
   End If

ElseIf vCom = "BCVAT46A" Then
If Option1.Value = True Then
vQuery = "select reportname from bcreportname where repid = '11' and reptype = 'GL' "
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
Else
vQuery = "select reportname from bcreportname where repid = '12' and reptype = 'GL' "
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

vQuery = "select reportname from bcreportname where repid = '13' and reptype = 'GL' "
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

End If
ElseIf vCom = "BCVAT47" Then
If Option1.Value = True Then
vQuery = "select reportname from bcreportname where repid = '14' and reptype = 'GL' "
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
ElseIf Option2.Value = True Then
vQuery = "select reportname from bcreportname where repid = '15' and reptype = 'GL' "
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

vQuery = "select reportname from bcreportname where repid = '16' and reptype = 'GL' "
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
Else
vQuery = "select reportname from bcreportname where repid = '17' and reptype = 'GL' "
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

vQuery = "select reportname from bcreportname where repid = '18' and reptype = 'GL' "
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

End If
ElseIf vCom = "BCVAT48" Then
If Option1.Value = True Then
vQuery = "select reportname from bcreportname where repid = '28' and reptype = 'GL' "
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
End If
End If
Else
  MsgBox "กรุณากรอกข้อมูลเกี่ยวกับฐานข้อมูลและเดือนปี ให้ครบด้วย", vbCritical, "Send Error Message"
End If

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vQuery As String, vReportName As String
Dim vRecordset As New ADODB.Recordset
Dim vCompany As String

CMBCompany.AddItem Trim("BCVAT")
CMBCompany.AddItem Trim("BCVAT46A")
CMBCompany.AddItem Trim("BCVAT47")
CMBCompany.AddItem Trim("BCVAT48")

CMBCompany.Text = Trim("BCVAT")
vCompany = Trim(CMBCompany.Text)

CMBMonth.Clear
vQuery = "select distinct month(docdate) as Month1  from " & vCompany & ".dbo.bctrans order by month(docdate) "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMBMonth.AddItem Trim(vRecordset.Fields("Month1").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

CMBYear.Clear
vQuery = "select distinct year(docdate) as Year1  from " & vCompany & ".dbo.bctrans order by year(docdate) "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMBYear.AddItem Trim(vRecordset.Fields("Year1").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub
