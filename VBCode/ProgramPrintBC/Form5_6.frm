VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form5_6 
   Caption         =   "พิมพ์เอกสารใบตรวจสอบข้อมูลลูกค้าและใบระเบียนลูกค้า"
   ClientHeight    =   9000
   ClientLeft      =   4470
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form5_6.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PIC101 
      BackColor       =   &H00C0C0C0&
      Height          =   9015
      Left            =   0
      ScaleHeight     =   8955
      ScaleWidth      =   11970
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   12030
      Begin VB.CommandButton CMDSearch 
         Height          =   285
         Left            =   5625
         Picture         =   "Form5_6.frx":9673
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1485
         Width           =   375
      End
      Begin VB.CommandButton CMDExit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ออก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6030
         Width           =   1140
      End
      Begin VB.CommandButton CMDSelect 
         BackColor       =   &H00C0C0C0&
         Caption         =   "เลือก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8010
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6030
         Width           =   1140
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   3840
         Left            =   1350
         TabIndex        =   12
         Top             =   1890
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   6773
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสลูกค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   12347
         EndProperty
      End
      Begin VB.TextBox TextSearchARCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2565
         TabIndex        =   11
         Top             =   1485
         Width           =   3030
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ค้นหารหัสลูกค้า :"
         Height          =   285
         Left            =   1350
         TabIndex        =   10
         Top             =   1485
         Width           =   1320
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   0
         Picture         =   "Form5_6.frx":9A40
         Top             =   0
         Width           =   2160
      End
   End
   Begin VB.ComboBox CMBPressMen2 
      Enabled         =   0   'False
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
      Left            =   3150
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4140
      Width           =   6180
   End
   Begin VB.ComboBox CMBPressMen1 
      Enabled         =   0   'False
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
      Left            =   3150
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3690
      Width           =   6180
   End
   Begin VB.ComboBox CMBArGroup4 
      Enabled         =   0   'False
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
      Left            =   3150
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3105
      Width           =   6180
   End
   Begin VB.ComboBox CMBArGroup3 
      Enabled         =   0   'False
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
      Left            =   3150
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2655
      Width           =   6180
   End
   Begin VB.ComboBox CMBArGroup2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3150
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2115
      Width           =   6180
   End
   Begin VB.ComboBox CMBArGroup1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3150
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1710
      Width           =   6180
   End
   Begin Crystal.CrystalReport Crystal102 
      Left            =   1710
      Top             =   6120
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
   Begin Crystal.CrystalReport Crystal101 
      Left            =   855
      Top             =   6390
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
   Begin VB.CommandButton CMDPrint 
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
      Height          =   420
      Left            =   8190
      TabIndex        =   6
      Top             =   5850
      Width           =   1185
   End
   Begin VB.CommandButton CMD103 
      Enabled         =   0   'False
      Height          =   330
      Left            =   5310
      Picture         =   "Form5_6.frx":AEA2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4725
      Width           =   375
   End
   Begin VB.ComboBox CMBSelectReportType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3150
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1125
      Width           =   4200
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงพนักงานเร่งรัด :"
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
      Height          =   285
      Left            =   900
      TabIndex        =   25
      Top             =   4140
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากพนักงานเร่งรัด :"
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
      Height          =   285
      Left            =   810
      TabIndex        =   24
      Top             =   3690
      Width           =   2265
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงกลุ่มลูกหนี้ :"
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
      Height          =   240
      Left            =   1845
      TabIndex        =   23
      Top             =   3105
      Width           =   1230
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากกลุ่มลูกหนี้ :"
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
      Height          =   240
      Left            =   1215
      TabIndex        =   22
      Top             =   2655
      Width           =   1860
   End
   Begin VB.Label LBLArCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3150
      TabIndex        =   17
      Top             =   4725
      Width           =   2130
   End
   Begin VB.Label LBLARName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   555
      Left            =   3150
      TabIndex        =   16
      Top             =   5175
      Width           =   6225
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกค้า :"
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
      Height          =   330
      Left            =   1710
      TabIndex        =   4
      Top             =   4725
      Width           =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงประเภทลูกค้า :"
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
      Left            =   1665
      TabIndex        =   3
      Top             =   2115
      Width           =   1410
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากประเภทลูกค้า :"
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
      Height          =   330
      Left            =   1665
      TabIndex        =   2
      Top             =   1710
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทเอกสาร :"
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
      Left            =   1620
      TabIndex        =   1
      Top             =   1125
      Width           =   1455
   End
End
Attribute VB_Name = "Form5_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String

Private Sub CMBSelectReportType_Click()
If Me.CMBSelectReportType.ListIndex = 0 Then
   Me.CMD103.Enabled = True
   Me.CMBArGroup1.Enabled = False
   Me.CMBArGroup2.Enabled = False
   Me.CMBArGroup3.Enabled = False
   Me.CMBArGroup4.Enabled = False
   Me.CMBPressMen1.Enabled = False
   Me.CMBPressMen2.Enabled = False
ElseIf Me.CMBSelectReportType.ListIndex = 1 Then
   Me.CMD103.Enabled = False
   Me.CMBArGroup1.Enabled = True
   Me.CMBArGroup2.Enabled = True
   Me.CMBArGroup3.Enabled = False
   Me.CMBArGroup4.Enabled = False
   Me.CMBPressMen1.Enabled = False
   Me.CMBPressMen2.Enabled = False
ElseIf Me.CMBSelectReportType.ListIndex = 2 Then
   Me.CMD103.Enabled = False
   Me.CMBArGroup1.Enabled = False
   Me.CMBArGroup2.Enabled = False
   Me.CMBArGroup3.Enabled = True
   Me.CMBArGroup4.Enabled = True
   Me.CMBPressMen1.Enabled = False
   Me.CMBPressMen2.Enabled = False
ElseIf Me.CMBSelectReportType.ListIndex = 3 Then
   Me.CMD103.Enabled = False
   Me.CMBArGroup1.Enabled = False
   Me.CMBArGroup2.Enabled = False
   Me.CMBArGroup3.Enabled = False
   Me.CMBArGroup4.Enabled = False
   Me.CMBPressMen1.Enabled = True
   Me.CMBPressMen2.Enabled = True
End If
End Sub

Private Sub CMD103_Click()
Me.Pic101.Visible = True
Me.TextSearchARCode.SetFocus
End Sub

Private Sub CMDExit_Click()
Pic101.Visible = False
End Sub

Private Sub CMDPrint_Click()
On Error GoTo ErrDescription

If Me.CMBSelectReportType.ListIndex = 0 And Me.LBLArCode.Caption <> "" Then
   Call PrintArProfile
ElseIf Me.CMBSelectReportType.ListIndex = 1 And Me.CMBArGroup1.Text <> "" And Me.CMBArGroup2.Text <> "" Then
   Call PrintCheckARComplete
ElseIf Me.CMBSelectReportType.ListIndex = 2 And Me.CMBArGroup3.Text <> "" And Me.CMBArGroup4.Text <> "" Then
   Call PrintCheckARComplete_ARDebt
ElseIf Me.CMBSelectReportType.ListIndex = 3 And Me.CMBPressMen1.Text <> "" And Me.CMBPressMen2.Text <> "" Then
   Call PrintCheckARComplete_PressMen
Else
   MsgBox "กรุณากรอกข้อมูลที่จะดูรายให้ตรงกับรายงานที่จะดู หรือกรอกให้ครบ", vbCritical, "Send Error Message"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Public Sub PrintArProfile()
Dim vRecordset As New ADODB.Recordset
Dim vARCode As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vARCode = Trim(Me.LBLArCode.Caption)
vRepID = 364
vRepType = "AR"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 364 and reptype = 'AR' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@ARCode;" & vARCode & ";true"
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

Public Sub PrintCheckARComplete()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARGroup1 As String
Dim vARGroup2 As String
Dim vReportName As String
Dim StrCount As Integer
Dim StrCount1 As Integer
Dim vFromGroup As String
Dim vToGroup As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vFromGroup = Me.CMBArGroup1.Text
vToGroup = Me.CMBArGroup2.Text
StrCount = InStr(Trim(Me.CMBArGroup1.Text), "/")
StrCount1 = InStr(Trim(Me.CMBArGroup2.Text), "/")
vARGroup1 = Trim(Left(Me.CMBArGroup1.Text, StrCount - 1))
vARGroup2 = Trim(Left(Me.CMBArGroup2.Text, StrCount1 - 1))

vRepID = 365
vRepType = "AR"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 365 and reptype = 'AR' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal102
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "pFromArGroup;" & vARGroup1 & ";true"
.ParameterFields(1) = "pToArGroup;" & vARGroup2 & ";true"
.Formulas(0) = "vFromGroup='" & vFromGroup & "' "
.Formulas(1) = "vToGroup='" & vToGroup & "' "
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

Public Sub PrintCheckARComplete_ARDebt()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARGroup1 As String
Dim vARGroup2 As String
Dim vReportName As String
Dim StrCount As Integer
Dim StrCount1 As Integer
Dim vFromGroup As String
Dim vToGroup As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vFromGroup = Me.CMBArGroup3.Text
vToGroup = Me.CMBArGroup4.Text
StrCount = InStr(Trim(Me.CMBArGroup3.Text), "/")
StrCount1 = InStr(Trim(Me.CMBArGroup4.Text), "/")
vARGroup1 = Trim(Left(Me.CMBArGroup3.Text, StrCount - 1))
vARGroup2 = Trim(Left(Me.CMBArGroup4.Text, StrCount1 - 1))

vRepID = 413
vRepType = "AR"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal102
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "pFromArGroup;" & vARGroup1 & ";true"
.ParameterFields(1) = "pToArGroup;" & vARGroup2 & ";true"
.Formulas(0) = "vFromGroup='" & vFromGroup & "' "
.Formulas(1) = "vToGroup='" & vToGroup & "' "
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

Public Sub PrintCheckARComplete_PressMen()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARGroup1 As String
Dim vARGroup2 As String
Dim vReportName As String
Dim StrCount As Integer
Dim StrCount1 As Integer
Dim vFromGroup As String
Dim vToGroup As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vFromGroup = Me.CMBPressMen1.Text
vToGroup = Me.CMBPressMen2.Text
StrCount = InStr(Trim(Me.CMBPressMen1.Text), "/")
StrCount1 = InStr(Trim(Me.CMBPressMen2.Text), "/")
vARGroup1 = Trim(Left(Me.CMBPressMen1.Text, StrCount - 1))
vARGroup2 = Trim(Left(Me.CMBPressMen2.Text, StrCount1 - 1))

vRepID = 414
vRepType = "AR"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal102
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "pFromArGroup;" & vARGroup1 & ";true"
.ParameterFields(1) = "pToArGroup;" & vARGroup2 & ";true"
.Formulas(0) = "vFromGroup='" & vFromGroup & "' "
.Formulas(1) = "vToGroup='" & vToGroup & "' "
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

Private Sub CMDSearch_Click()
Dim vSearch As String
Dim vListAR As ListItem
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If TextSearchARCode.Text <> "" Then
  vSearch = TextSearchARCode.Text
  vQuery = "exec dbo.USP_MP_SearchArCode 1,'" & vSearch & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  Me.ListView101.ListItems.Clear
  vRecordset.MoveFirst
  While Not vRecordset.EOF
  Set vListAR = Me.ListView101.ListItems.Add(, , vRecordset.Fields("code").Value)
  vListAR.SubItems(1) = Trim(vRecordset.Fields("arname").Value)
  vRecordset.MoveNext
  Wend
  Me.ListView101.SetFocus
  Else
  Me.ListView101.ListItems.Clear
  Me.TextSearchARCode.SetFocus
  End If
  vRecordset.Close
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDSelect_Click()
Dim vIndex As Integer

On Error GoTo ErrDescription

If Me.ListView101.ListItems.Count > 0 Then
  vIndex = Me.ListView101.SelectedItem.Index
  Me.LBLArCode.Caption = Me.ListView101.ListItems(vIndex).Text
  Me.LBLARName.Caption = Me.ListView101.ListItems(vIndex).SubItems(1)
  Me.Pic101.Visible = False
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Me.CMBSelectReportType.AddItem ("ใบตรวจสอบข้อมูลลูกค้า")
Me.CMBSelectReportType.AddItem ("ใบระเบียนข้อมูลลูกค้า ตามประเภทลูกหนี้")
Me.CMBSelectReportType.AddItem ("ใบระเบียนข้อมูลลูกค้า ตามกลุ่มลูกหนี้")
Me.CMBSelectReportType.AddItem ("ใบระเบียนข้อมูลลูกค้า ตามพนักงานเร่งรัด")
Call GetARGroup
Call GetARDebtCode
Call GetPressMen
End Sub

Public Sub GetPressMen()
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

vQuery = "select  pressmencode+'/'+name as pressmen  from dbo.BCAR a inner join dbo.bcsale b on a.pressmencode = b.code where a.activestatus = 1 and b.activestatus = 1 group by pressmencode,name order by pressmencode"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   While Not vRecordset.EOF
      Me.CMBPressMen1.AddItem (vRecordset.Fields("pressmen"))
      Me.CMBPressMen2.AddItem (vRecordset.Fields("pressmen"))
      vRecordset.MoveNext
   Wend
End If
vRecordset.Close

End Sub

Public Sub GetARGroup()
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

vQuery = "select code+'/'+name as custtype from dbo.BCCustType order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   While Not vRecordset.EOF
      Me.CMBArGroup1.AddItem (vRecordset.Fields("custtype"))
      Me.CMBArGroup2.AddItem (vRecordset.Fields("custtype"))
      vRecordset.MoveNext
   Wend
End If
vRecordset.Close

End Sub

Public Sub GetARDebtCode()
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

vQuery = "select code+'/'+name as debtgroup from dbo.BCARDebtGroup order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   While Not vRecordset.EOF
      Me.CMBArGroup3.AddItem (vRecordset.Fields("debtgroup"))
      Me.CMBArGroup4.AddItem (vRecordset.Fields("debtgroup"))
      vRecordset.MoveNext
   Wend
End If
vRecordset.Close

End Sub

Private Sub ListView101_DblClick()
Dim vIndex As Integer

On Error GoTo ErrDescription

If Me.ListView101.ListItems.Count > 0 Then
  vIndex = Me.ListView101.SelectedItem.Index
  Me.LBLArCode.Caption = Me.ListView101.ListItems(vIndex).Text
  Me.LBLARName.Caption = Me.ListView101.ListItems(vIndex).SubItems(1)
  Me.Pic101.Visible = False
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
   If Me.ListView101.ListItems.Count > 0 Then
     vIndex = Me.ListView101.SelectedItem.Index
     Me.LBLArCode.Caption = Me.ListView101.ListItems(vIndex).Text
     Me.LBLARName.Caption = Me.ListView101.ListItems(vIndex).SubItems(1)
     Me.Pic101.Visible = False
   End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub TextSearchARCode_KeyPress(KeyAscii As Integer)
Dim vSearch As String
Dim vListAR As ListItem
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If KeyAscii = 13 Then
   If TextSearchARCode.Text <> "" Then
     vSearch = TextSearchARCode.Text
     vQuery = "exec dbo.USP_MP_SearchArCode 1,'" & vSearch & "' "
     If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     Me.ListView101.ListItems.Clear
     vRecordset.MoveFirst
     While Not vRecordset.EOF
     Set vListAR = Me.ListView101.ListItems.Add(, , vRecordset.Fields("code").Value)
     vListAR.SubItems(1) = Trim(vRecordset.Fields("arname").Value)
     vRecordset.MoveNext
     Wend
     Me.ListView101.SetFocus
     Else
     Me.ListView101.ListItems.Clear
     Me.TextSearchARCode.SetFocus
     End If
     vRecordset.Close
   End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub
