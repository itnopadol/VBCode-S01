VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form92 
   Caption         =   "รายงานสรุปค่าใช้จ่ายประจำปี"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   Icon            =   "Form92.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form92.frx":08CA
   ScaleHeight     =   8340
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport92 
      Left            =   5160
      Top             =   6720
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
   Begin Crystal.CrystalReport Crystal1 
      Left            =   2760
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
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1440
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
   Begin VB.ComboBox CMBMonth2 
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
      Left            =   5700
      TabIndex        =   9
      Top             =   2925
      Width           =   3240
   End
   Begin VB.ComboBox CMBMonth1 
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
      Left            =   5700
      TabIndex        =   8
      Top             =   2400
      Width           =   3240
   End
   Begin VB.ComboBox CMB923 
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
      Left            =   5700
      TabIndex        =   6
      Top             =   3675
      Width           =   3240
   End
   Begin VB.ComboBox CMB922 
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
      Left            =   5700
      TabIndex        =   4
      Text            =   "สรุปค่าใช้จ่ายประจำปีแบบรวม"
      Top             =   1500
      Width           =   3240
   End
   Begin VB.CommandButton CMD921 
      Caption         =   "พิมพ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2550
      TabIndex        =   2
      Top             =   2550
      Width           =   1440
   End
   Begin VB.ComboBox CMB921 
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
      Left            =   2100
      TabIndex        =   1
      Top             =   1500
      Width           =   1890
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงเดือน"
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
      Left            =   4950
      TabIndex        =   11
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "จากเดือน"
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
      Left            =   4950
      TabIndex        =   10
      Top             =   2475
      Width           =   690
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "แผนก"
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
      Left            =   5100
      TabIndex        =   7
      Top             =   3675
      Width           =   465
   End
   Begin VB.Label LBL922 
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทการดูรายงาน"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   1500
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงานสรุปค่าใช้จ่ายประจำปี"
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
      TabIndex        =   3
      Top             =   300
      Width           =   7365
   End
   Begin VB.Label LBL921 
      BackStyle       =   0  'Transparent
      Caption         =   "ประจำปี"
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
      Left            =   1425
      TabIndex        =   0
      Top             =   1500
      Width           =   690
   End
End
Attribute VB_Name = "Form92"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD921_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As Integer, vMonth1 As Integer, vMonth2 As Integer
Dim vRepType As String
Dim vRepID As String, vReportName As String
Dim vDetails As String

On Error GoTo ErrDescription

    If CMB922.Text = "สรุปค่าใช้จ่ายประจำปีแบบรวม" Then
        Call PrintPayMent_Total
    ElseIf CMB922.Text = "สรุปค่าใช้จ่ายประจำปีแบบแยก" Then
        Call PrintPayMent_Depart
    Else
        Call PrintPayMent_Depart2
    End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

CMB922.AddItem "สรุปค่าใช้จ่ายประจำปีแบบรวม"
CMB922.AddItem "สรุปค่าใช้จ่ายประจำปีแบบแยก"
CMB922.AddItem "สรุปค่าใช้จ่ายประจำปีแบบแยกแผนก"

'vQuery = "select distinct year(docdate) as PaymentYear from BCPAYMENT order by paymentyear desc "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vRecordset.MoveFirst
  '  While Not vRecordset.EOF
        CMB921.AddItem Trim("2004")
        CMB921.AddItem Trim("2005")
        CMB921.AddItem Trim("2006")
        CMB921.AddItem Trim("2007")
        CMB921.AddItem Trim("2008")
   '     vRecordset.MoveNext
    'Wend
'End If
'vRecordset.Close

vQuery = "select code+'    '+name as name from BCdepartment order by code  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB923.AddItem Trim(vRecordset.Fields("name").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

vQuery = "select distinct month(docdate) as PaymentMonth from BCPAYMENT order by paymentMonth  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMBMonth1.AddItem Trim(vRecordset.Fields("PaymentMonth").Value)
        CMBMonth2.AddItem Trim(vRecordset.Fields("PaymentMonth").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub

Public Sub PrintPayMent_Total()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As Integer, vMonth1 As Integer, vMonth2 As Integer
Dim vRepType As String
Dim vRepID As String, vReportName As String
Dim vDetails As String

On Error GoTo ErrDescription


If CMB921.Text = "2004" Then
    vRepID = 244
ElseIf CMB921.Text = "2005" Then
    vRepID = 354
ElseIf CMB921.Text = "2006" Then
    vRepID = 355
ElseIf CMB921.Text = "2007" Then
    vRepID = 77
ElseIf CMB921.Text = "2008" Then
    vRepID = 77
End If

    vRepType = "GL"
    vMonth1 = Trim(CMBMonth1.Text)
    vMonth2 = Trim(CMBMonth2.Text)
    vYear = CMB921.Text

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = '" & vRepID & "' and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With CrystalReport92
    .ReportFileName = Trim(vReportName) & ".rpt"
    .ParameterFields(0) = "@year;" & vYear & ";true"
    .ParameterFields(1) = "@Mb;" & vMonth1 & ";true"
    .ParameterFields(2) = "@Me;" & vMonth2 & ";true"
    .WindowState = crptMaximized
    .Destination = crptToWindow
    .Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub PrintPayMent_Depart()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As Integer, vMonth1 As Integer, vMonth2 As Integer
Dim vRepType As String
Dim vRepID As String, vReportName As String
Dim vDetails As String

On Error GoTo ErrDescription
    
    If CMB921.Text = "2004" Then
    vRepID = 245
    ElseIf CMB921.Text = "2005" Then
    vRepID = 356
    ElseIf CMB921.Text = "2006" Then
    vRepID = 357
    ElseIf CMB921.Text = "2007" Then
    vRepID = 84
    ElseIf CMB921.Text = "2008" Then
    vRepID = 84
    End If
    
    vRepType = "GL"
    vDetails = Trim(Left(CMB923.Text, 2))
    vYear = CMB921.Text

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = '" & vRepID & "' and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
    .ReportFileName = Trim(vReportName) & ".rpt"
    .ParameterFields(0) = "@year;" & vYear & ";true"
    .ParameterFields(1) = "@DepartCode;" & vDetails & ";true"
    .WindowState = crptMaximized
    .Destination = crptToWindow
    .Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub PrintPayMent_Depart2()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vFDate As Integer, vLDate As Integer, vFYear As Integer
Dim vRepType As String
Dim vRepID As String, vReportName As String
Dim vDetails As String

On Error GoTo ErrDescription

    If CMB921.Text = "2004" Then
        vRepID = 246
    ElseIf CMB921.Text = "2005" Then
        vRepID = 358
    ElseIf CMB921.Text = "2006" Then
       vRepID = 359
    ElseIf CMB921.Text = "2007" Then
        vRepID = 168
    ElseIf CMB921.Text = "2008" Then
        vRepID = 168
    End If

    vRepType = "GL"
    vFDate = CMBMonth1.Text
    vLDate = CMBMonth2.Text
    vFYear = CMB921.Text

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = '" & vRepID & "' and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal1
    .ReportFileName = Trim(vReportName) & ".rpt"
    .ParameterFields(0) = "@FDate;" & vFDate & ";true"
    .ParameterFields(1) = "@LDate;" & vLDate & ";true"
    .ParameterFields(2) = "@FYear;" & vFYear & ";true"
    .WindowState = crptMaximized
    .Destination = crptToWindow
    .Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


