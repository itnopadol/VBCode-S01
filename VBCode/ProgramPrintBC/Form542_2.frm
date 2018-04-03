VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form542_2 
   Caption         =   "หน้าพิมพ์รายงานยอดลูกหนี้ประจำเดือนตามรหัสประเภทลูกค้า"
   ClientHeight    =   8355
   ClientLeft      =   2355
   ClientTop       =   720
   ClientWidth     =   12000
   Icon            =   "Form542_2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form542_2.frx":08CA
   ScaleHeight     =   8355
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport542_23 
      Left            =   4185
      Top             =   6660
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
   Begin VB.CheckBox CKPressMen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ตามพนักงานเร่งรัด :"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   765
      TabIndex        =   13
      Top             =   3600
      Width           =   1725
   End
   Begin Crystal.CrystalReport CrystalReport542_22 
      Left            =   2880
      Top             =   6975
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
   Begin VB.ComboBox CMBPressMen 
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
      Left            =   2610
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3600
      Width           =   4065
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รวมปีที่ยังไม่ยกยอด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4725
      TabIndex        =   11
      Top             =   4140
      Width           =   1950
   End
   Begin VB.ComboBox CMB101 
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
      Left            =   2610
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1980
      Width           =   2850
   End
   Begin Crystal.CrystalReport CrystalReport542_21 
      Left            =   1680
      Top             =   6600
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
   Begin VB.CommandButton CMD542_21 
      Caption         =   "พิมพ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5580
      TabIndex        =   8
      Top             =   4995
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView542_21 
      Height          =   5805
      Left            =   6750
      TabIndex        =   3
      Top             =   1980
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   10239
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสประเภทลูกค้า"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อประเภทลูกค้า"
         Object.Width           =   5468
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTP542_21 
      Height          =   345
      Left            =   4725
      TabIndex        =   2
      Top             =   4590
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
      _Version        =   393216
      Format          =   60817409
      CurrentDate     =   38027
   End
   Begin VB.ComboBox CMB542_22 
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
      Left            =   2595
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3060
      Width           =   4065
   End
   Begin VB.ComboBox CMB542_21 
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
      Left            =   2610
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2520
      Width           =   4065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   1215
      TabIndex        =   10
      Top             =   1980
      Width           =   1275
   End
   Begin VB.Label LBL542_24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์รายงานยอดลูกหนี้ประจำเดือนตามรหัสประเภทลูกค้า"
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
      Left            =   2550
      TabIndex        =   7
      Top             =   300
      Width           =   7440
   End
   Begin VB.Label LBL542_23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3960
      TabIndex        =   6
      Top             =   4635
      Width           =   705
   End
   Begin VB.Label LBL542_22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงประเภทลูกหนี้"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1170
      TabIndex        =   5
      Top             =   3060
      Width           =   1290
   End
   Begin VB.Label LBL542_21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากประเภทลูกหนี้"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   270
      TabIndex        =   4
      Top             =   2520
      Width           =   2220
   End
End
Attribute VB_Name = "Form542_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD542_21_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vTypeCode1 As String, vTypeCode2 As String
Dim vDate1 As Date
Dim vRepID As Integer
Dim vRepType As String
Dim vIsPresent As Integer
Dim vCheckMonth1 As Integer
Dim vPressMenCode As String

'On Error GoTo ErrDescription

If CKPressMen.Value = 1 And CMB101.ListIndex <> 0 Then
  Call ReportOfPressMen
  Exit Sub
End If

If CMB101.ListIndex = 0 Then
  Call ReportOfARPeriodByPressMen
ElseIf CMB101.ListIndex = 1 Then
  vRepID = 300
ElseIf CMB101.ListIndex = 2 Then
  vRepID = 32
End If

vRepType = "AR"
vTypeCode1 = Trim(CMB542_21.Text)
vTypeCode2 = Trim(CMB542_22.Text)
vDate1 = DTP542_21.Day & "/" & DTP542_21.Month & "/" & DTP542_21.Year

  If Me.CMBPressMen.Text <> "" Then
    vPressMenCode = Right(Me.CMBPressMen.Text, Len(Me.CMBPressMen.Text) - InStr(Me.CMBPressMen.Text, "/"))
  Else
    vPressMenCode = ""
  End If
  

If Check1.Value = 1 Then
    vIsPresent = 0
Else
    vIsPresent = 1
End If


vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport542_21
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@TYPECODE1;" & vTypeCode1 & " ;true"
        .ParameterFields(1) = "@TYPECODE2;" & vTypeCode2 & " ;true"
        .ParameterFields(2) = "@AtDate;" & vDate1 & ";true"
        .ParameterFields(3) = "@IsPresent;" & vIsPresent & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close

'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description
  '  Exit Sub
'End If
End Sub


Public Sub ReportOfARPeriodByPressMen()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vTypeCode1 As String, vTypeCode2 As String
Dim vDate1 As Date
Dim vRepID As Integer
Dim vRepType As String
Dim vIsPresent As Integer
Dim vCheckMonth1 As Integer
Dim vPressMenCode As String

On Error GoTo ErrDescription


If DTP542_21.Month < 7 Then
  vRepID = 336
ElseIf DTP542_21.Month > 6 Then
  vRepID = 337
End If

vPressMenCode = Right(Me.CMBPressMen.Text, Len(Me.CMBPressMen.Text) - InStr(Me.CMBPressMen.Text, "/"))

vRepType = "AR"
vTypeCode1 = Trim(CMB542_21.Text)
vTypeCode2 = Trim(CMB542_22.Text)
vDate1 = DTP542_21.Day & "/" & DTP542_21.Month & "/" & DTP542_21.Year

If Check1.Value = 1 Then
    vIsPresent = 0
Else
    vIsPresent = 1
End If
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport542_23
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@TYPECODE1;" & vTypeCode1 & " ;true"
        .ParameterFields(1) = "@TYPECODE2;" & vTypeCode2 & " ;true"
        .ParameterFields(2) = "@AtDate;" & vDate1 & ";true"
        .ParameterFields(3) = "@IsPresent;" & vIsPresent & ";true"
        .ParameterFields(4) = "@pressmencode;" & vPressMenCode & ";true"
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


Public Sub ReportOfPressMen()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vTypeCode1 As String, vTypeCode2 As String
Dim vDate1 As Date
Dim vRepID As Integer
Dim vRepType As String
Dim vIsPresent As Integer
Dim vCheckMonth1 As Integer
Dim vPressMenCode As String

On Error GoTo ErrDescription

If DTP542_21.Month < 7 Then
  vRepID = 338
  vPressMenCode = Right(Me.CMBPressMen.Text, Len(Me.CMBPressMen.Text) - InStr(Me.CMBPressMen.Text, "/"))
ElseIf DTP542_21.Month > 6 Then
  vRepID = 339
  vPressMenCode = Right(Me.CMBPressMen.Text, Len(Me.CMBPressMen.Text) - InStr(Me.CMBPressMen.Text, "/"))
End If

vRepType = "AR"
vTypeCode1 = Trim(CMB542_21.Text)
vTypeCode2 = Trim(CMB542_22.Text)
vDate1 = DTP542_21.Day & "/" & DTP542_21.Month & "/" & DTP542_21.Year

If Check1.Value = 1 Then
    vIsPresent = 0
Else
    vIsPresent = 1
End If
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport542_22
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@TYPECODE1;" & vTypeCode1 & " ;true"
        .ParameterFields(1) = "@TYPECODE2;" & vTypeCode2 & " ;true"
        .ParameterFields(2) = "@AtDate;" & vDate1 & ";true"
        .ParameterFields(3) = "@IsPresent;" & vIsPresent & ";true"
        .ParameterFields(4) = "@vPressMenCode;" & vPressMenCode & ";true"
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
Dim vGroupARItems As ListItem

On Error GoTo ErrDescription

DTP542_21 = Now

vQuery = "select distinct code,name from bccusttype  order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB542_21.AddItem Trim(vRecordset.Fields("code").Value)
        CMB542_22.AddItem Trim(vRecordset.Fields("code").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

vQuery = "select distinct code,name from bccusttype  order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set vGroupARItems = ListView542_21.ListItems.Add(, , Trim(vRecordset.Fields("code").Value))
    vGroupARItems.SubItems(1) = Trim(vRecordset.Fields("name").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

vQuery = "select  distinct  isnull(pressmencode,'0000') as pressmencode ,isnull(name,'ไม่มีพนักงานเร่งรัด') as pressmenname from  dbo.bcar a left join dbo.bcsale b on a.pressmencode = b.code order by pressmenname"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMBPressMen.AddItem Trim(vRecordset.Fields("pressmenname").Value & "/" & vRecordset.Fields("pressmencode").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

CMB101.AddItem Trim("รายงานทั่วไป")
CMB101.AddItem Trim("รายงานสำหรับสินเชื่อ")
CMB101.AddItem Trim("รายงานทั่วไป รวมทุกเดือน")

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If

End Sub
