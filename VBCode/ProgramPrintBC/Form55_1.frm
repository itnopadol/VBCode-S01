VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Form55_1 
   Caption         =   "¾ÔÁ¾ìÃÒÂ§Ò¹áÊ´§Ë¹Õé¤§¤éÒ§¢Í§ÅÙ¡¤éÒ"
   ClientHeight    =   8280
   ClientLeft      =   4635
   ClientTop       =   1620
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form55_1.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDARDebt_CR 
      Caption         =   "¾ÔÁ¾ìÃÒÂ§Ò¹ Ë¹éÒÃéÒ¹"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2940
      TabIndex        =   12
      Top             =   6780
      Width           =   1755
   End
   Begin Crystal.CrystalReport Crystal102 
      Left            =   1290
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowPrintBtn=   0   'False
      WindowShowExportBtn=   0   'False
   End
   Begin VB.CommandButton CMDReqApproveDebt 
      Caption         =   "¾ÔÁ¾ìãº¢ÍÍ¹ØÁÑµÔà¾ÔèÁÇ§à§Ô¹"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9090
      TabIndex        =   11
      Top             =   6270
      Width           =   2325
   End
   Begin VB.CommandButton CMDPrintLoseDebt 
      Caption         =   "¾ÔÁ¾ìãº¢ÍµÑ´Ë¹ÕéÊÙ­"
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
      Left            =   7020
      TabIndex        =   10
      Top             =   6255
      Width           =   1770
   End
   Begin VB.CommandButton CMDPrintAccuse 
      Caption         =   "¾ÔÁ¾ìãº¿éÍ§ÃéÍ§"
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
      Left            =   4980
      TabIndex        =   9
      Top             =   6240
      Width           =   1770
   End
   Begin VB.CommandButton CMDSearchARCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      Picture         =   "Form55_1.frx":9673
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1755
      Width           =   330
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   450
      Top             =   7245
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
      Caption         =   "¾ÔÁ¾ìÃÒÂ§Ò¹"
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
      Left            =   2940
      TabIndex        =   3
      Top             =   6240
      Width           =   1755
   End
   Begin VB.TextBox TextSearchARCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2925
      TabIndex        =   0
      Top             =   1755
      Width           =   3525
   End
   Begin MSComctlLib.ListView ListViewARCode 
      Height          =   2760
      Left            =   2025
      TabIndex        =   2
      Top             =   2160
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4868
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
         Text            =   "ÃËÑÊÅÙ¡¤éÒ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ª×èÍÅÙ¡¤éÒ"
         Object.Width           =   10760
      EndProperty
   End
   Begin VB.Label LBLARCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2925
      TabIndex        =   8
      Top             =   5355
      Width           =   3165
   End
   Begin VB.Label LBLARName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2925
      TabIndex        =   7
      Top             =   5715
      Width           =   8475
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "¤Ó·Õè¤é¹ËÒ :"
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
      Left            =   2025
      TabIndex        =   6
      Top             =   1755
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "¤é¹ËÒÃËÑÊÅÙ¡¤éÒ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   2025
      TabIndex        =   5
      Top             =   1170
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ÃËÑÊÅÙ¡¤éÒ :"
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
      Left            =   2025
      TabIndex        =   4
      Top             =   5355
      Width           =   1140
   End
End
Attribute VB_Name = "Form55_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String

Private Sub CMDARDebt_CR_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAutoNumber As Integer
Dim vGenNumber As String
Dim vYear As String
Dim vMaxNumber As Integer
Dim vDocdate As Date
Dim vARCode As String
Dim vRepType As String
Dim vRepID As Integer

On Error GoTo ErrDescription

If vDepartment = "CR" Or vDepartment = "WS" Or vDepartment = "IT" Or vDepartment = "MC" Then
If Me.LBLArCode.Caption <> "" And Me.LBLARName.Caption <> "" Then

vARCode = Trim(LBLArCode.Caption)
vRepID = 350
vRepType = "AR"
 'vQuery = "select reportname from dbo.bcreportname where repid = 350 and reptype = 'AR' "
 
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal102
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vArCode;" & vARCode & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close

Me.LBLArCode.Caption = ""
Me.LBLARName.Caption = ""
Me.ListViewARCode.ListItems.Clear
Me.TextSearchARCode.Text = ""
Me.TextSearchARCode.SetFocus

Else
  MsgBox "¡ÃØ³Ò ¡ÃÍ¡ÃËÑÊÊÔ¹¤éÒ·ÕèµéÍ§¡ÒÃ¨Ð¾ÔÁ¾ìàÍ¡ÊÒÃ´éÇÂ", vbCritical, "Send Error"
End If

End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDPrint_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAutoNumber As Integer
Dim vGenNumber As String
Dim vYear As String
Dim vMaxNumber As Integer
Dim vDocdate As Date
Dim vARCode As String
Dim vRepType As String
Dim vRepID As Integer

On Error GoTo ErrDescription

If vDepartment = "AC" Or vDepartment = "CD" Or vDepartment = "IT" Then
If Me.LBLArCode.Caption <> "" And Me.LBLARName.Caption <> "" Then

vARCode = Trim(LBLArCode.Caption)
vRepID = 350
vRepType = "AR"
 'vQuery = "select reportname from dbo.bcreportname where repid = 350 and reptype = 'AR' "
 
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal101
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vArCode;" & vARCode & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close

Me.LBLArCode.Caption = ""
Me.LBLARName.Caption = ""
Me.ListViewARCode.ListItems.Clear
Me.TextSearchARCode.Text = ""
Me.TextSearchARCode.SetFocus

Else
  MsgBox "¡ÃØ³Ò ¡ÃÍ¡ÃËÑÊÊÔ¹¤éÒ·ÕèµéÍ§¡ÒÃ¨Ð¾ÔÁ¾ìàÍ¡ÊÒÃ´éÇÂ", vbCritical, "Send Error"
End If

End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDPrintAccuse_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAutoNumber As Integer
Dim vGenNumber As String
Dim vYear As String
Dim vMaxNumber As Integer
Dim vDocdate As Date
Dim vARCode As String
Dim vRepType As String
Dim vRepID As Integer

On Error GoTo ErrDescription

If vDepartment = "AC" Or vDepartment = "CD" Or vDepartment = "IT" Then

If Me.LBLArCode.Caption <> "" And Me.LBLARName.Caption <> "" Then

vARCode = Trim(LBLArCode.Caption)
vRepID = 466
vRepType = "AR"
 'vQuery = "select reportname from dbo.bcreportname where repid = 350 and reptype = 'AR' "
 
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal101
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vArCode;" & vARCode & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close

Me.LBLArCode.Caption = ""
Me.LBLARName.Caption = ""
Me.ListViewARCode.ListItems.Clear
Me.TextSearchARCode.Text = ""
Me.TextSearchARCode.SetFocus

Else
  MsgBox "¡ÃØ³Ò ¡ÃÍ¡ÃËÑÊÊÔ¹¤éÒ·ÕèµéÍ§¡ÒÃ¨Ð¾ÔÁ¾ìàÍ¡ÊÒÃ´éÇÂ", vbCritical, "Send Error"
End If

End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDPrintLoseDebt_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAutoNumber As Integer
Dim vGenNumber As String
Dim vYear As String
Dim vMaxNumber As Integer
Dim vDocdate As Date
Dim vARCode As String
Dim vRepType As String
Dim vRepID As Integer

On Error GoTo ErrDescription

If vDepartment = "AC" Or vDepartment = "CD" Or vDepartment = "IT" Then
If Me.LBLArCode.Caption <> "" And Me.LBLARName.Caption <> "" Then

vARCode = Trim(LBLArCode.Caption)
vRepID = 469
vRepType = "AR"
 'vQuery = "select reportname from dbo.bcreportname where repid = 350 and reptype = 'AR' "
 
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal101
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vArCode;" & vARCode & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close

Me.LBLArCode.Caption = ""
Me.LBLARName.Caption = ""
Me.ListViewARCode.ListItems.Clear
Me.TextSearchARCode.Text = ""
Me.TextSearchARCode.SetFocus

Else
  MsgBox "¡ÃØ³Ò ¡ÃÍ¡ÃËÑÊÊÔ¹¤éÒ·ÕèµéÍ§¡ÒÃ¨Ð¾ÔÁ¾ìàÍ¡ÊÒÃ´éÇÂ", vbCritical, "Send Error"
End If

End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDReqApproveDebt_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAutoNumber As Integer
Dim vGenNumber As String
Dim vYear As String
Dim vMaxNumber As Integer
Dim vDocdate As Date
Dim vARCode As String
Dim vRepType As String
Dim vRepID As Integer

On Error GoTo ErrDescription

If vDepartment = "AC" Or vDepartment = "CD" Or vDepartment = "IT" Then
If Me.LBLArCode.Caption <> "" And Me.LBLARName.Caption <> "" Then

vARCode = Trim(LBLArCode.Caption)
vRepID = 512
vRepType = "AR"
 
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal101
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vArCode;" & vARCode & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close

Me.LBLArCode.Caption = ""
Me.LBLARName.Caption = ""
Me.ListViewARCode.ListItems.Clear
Me.TextSearchARCode.Text = ""
Me.TextSearchARCode.SetFocus

Else
  MsgBox "¡ÃØ³Ò ¡ÃÍ¡ÃËÑÊÊÔ¹¤éÒ·ÕèµéÍ§¡ÒÃ¨Ð¾ÔÁ¾ìàÍ¡ÊÒÃ´éÇÂ", vbCritical, "Send Error"
End If

End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDSearchARCode_Click()
Dim vSearch As String
Dim vListAR As ListItem
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If TextSearchARCode.Text <> "" Then
  vSearch = TextSearchARCode.Text
  vQuery = "exec dbo.USP_MP_SearchArCode 1,'" & vSearch & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  Me.ListViewARCode.ListItems.Clear
  vRecordset.MoveFirst
  While Not vRecordset.EOF
  Set vListAR = Me.ListViewARCode.ListItems.Add(, , vRecordset.Fields("code").Value)
  vListAR.SubItems(1) = Trim(vRecordset.Fields("arname").Value)
  vRecordset.MoveNext
  Wend
  Me.ListViewARCode.SetFocus
  Else
  Me.ListViewARCode.ListItems.Clear
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

Private Sub ListViewARCode_DblClick()
Dim vIndex As Integer

On Error GoTo ErrDescription

If Me.ListViewARCode.ListItems.Count > 0 Then
  vIndex = Me.ListViewARCode.SelectedItem.Index
  Me.LBLArCode.Caption = Me.ListViewARCode.ListItems(vIndex).Text
  Me.LBLARName.Caption = Me.ListViewARCode.ListItems(vIndex).SubItems(1)
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListViewARCode_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
 If Me.ListViewARCode.ListItems.Count > 0 Then
   vIndex = Me.ListViewARCode.SelectedItem.Index
   Me.LBLArCode.Caption = Me.ListViewARCode.ListItems(vIndex).Text
   Me.LBLARName.Caption = Me.ListViewARCode.ListItems(vIndex).SubItems(1)
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
    Me.ListViewARCode.ListItems.Clear
    Me.LBLArCode.Caption = ""
    Me.LBLARName.Caption = ""
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set vListAR = Me.ListViewARCode.ListItems.Add(, , vRecordset.Fields("code").Value)
    vListAR.SubItems(1) = Trim(vRecordset.Fields("arname").Value)
    vRecordset.MoveNext
    Wend
    Me.ListViewARCode.SetFocus
    Else
    Me.ListViewARCode.ListItems.Clear
    Me.LBLArCode.Caption = ""
    Me.LBLARName.Caption = ""
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
