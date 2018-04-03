VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FormARDebtMovementByMonth 
   Caption         =   "รายงาน ลูกหนี้"
   ClientHeight    =   11010
   ClientLeft      =   3315
   ClientTop       =   75
   ClientWidth     =   15375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormARDebtMovementByMonth.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal104 
      Left            =   4005
      Top             =   7695
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
   Begin Crystal.CrystalReport Crystal103 
      Left            =   1935
      Top             =   7695
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
   Begin VB.CommandButton CMDPrintKeepCode 
      Caption         =   "ดูตามผู้ติดตามหนี้"
      Height          =   780
      Left            =   9135
      TabIndex        =   15
      Top             =   4860
      Width           =   1770
   End
   Begin VB.CommandButton CMDPrintCustType 
      Caption         =   "ดูตามประเภท"
      Height          =   780
      Left            =   7065
      TabIndex        =   14
      Top             =   4860
      Width           =   1725
   End
   Begin VB.ComboBox CMBArGroup2 
      Height          =   360
      Left            =   9135
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2160
      Width           =   4560
   End
   Begin VB.ComboBox CMBArGroup1 
      Height          =   360
      Left            =   2925
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2205
      Width           =   4560
   End
   Begin VB.ComboBox CMBKeepMoney2 
      Height          =   360
      Left            =   9135
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2970
      Width           =   4560
   End
   Begin VB.ComboBox CMBKeepMoney1 
      Height          =   360
      Left            =   2925
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2970
      Width           =   4560
   End
   Begin Crystal.CrystalReport Crystal102 
      Left            =   2610
      Top             =   6435
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
   Begin VB.ComboBox CMBCustType2 
      Height          =   360
      Left            =   9135
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   4560
   End
   Begin VB.ComboBox CMBCustType1 
      Height          =   360
      Left            =   2925
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   4560
   End
   Begin VB.CommandButton CMDPrintArGroup 
      Caption         =   "ดูตามกลุ่ม"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4995
      TabIndex        =   1
      Top             =   4860
      Width           =   1725
   End
   Begin VB.CommandButton CMDPrintAll 
      Caption         =   "ดูทั้งหมด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   2970
      TabIndex        =   0
      Top             =   4860
      Width           =   1725
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1440
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
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงกลุ่มลูกหนี้ :"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7830
      TabIndex        =   12
      Top             =   2205
      Width           =   1230
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากกลุ่มลูกหนี้ :"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1395
      TabIndex        =   11
      Top             =   2250
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงรหัสผู้ติดตามหนี้ :"
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   7560
      TabIndex        =   9
      Top             =   2970
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากรหัสผู้ติดตามหนี้ :"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1125
      TabIndex        =   8
      Top             =   2970
      Width           =   1725
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงประเภทลูกหนี้ :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7695
      TabIndex        =   4
      Top             =   1485
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากประเภทลูกหนี้ :"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1125
      TabIndex        =   3
      Top             =   1485
      Width           =   1725
   End
End
Attribute VB_Name = "FormARDebtMovementByMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CMDPrintAll_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepType As String
Dim vRepID As Integer

On Error GoTo ErrDescription


vRepID = 491
vRepType = "AR"
 
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal101
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
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

Private Sub CMDPrintArGroup_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARGroup1 As String
Dim vARGroup2 As String
Dim vRepType As String
Dim vRepID As Integer

On Error GoTo ErrDescription

If Me.CMBArGroup1.Text <> "" Then
    vARGroup1 = Left(Me.CMBArGroup1.Text, InStr(1, Me.CMBArGroup1.Text, "//") - 1)
    If Me.CMBArGroup2.Text <> "" Then
    vARGroup2 = Left(Me.CMBArGroup2.Text, InStr(1, Me.CMBArGroup2.Text, "//") - 1)
    Else
     vARGroup2 = Left(Me.CMBArGroup1.Text, InStr(1, Me.CMBArGroup1.Text, "//") - 1)
    End If
    
    vRepType = "AR"
    vRepID = 493
     
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        With Crystal102
            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
            .ParameterFields(0) = "@vArGroup1;" & vARGroup1 & ";true"
            .ParameterFields(1) = "@vArGroup2;" & vARGroup2 & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With
    End If
    vRecordset.Close
    
    vRepID = 494
    
    If Month(Now) > 6 Then
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        With Crystal102
            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
            .ParameterFields(0) = "@vArGroup1;" & vARGroup1 & ";true"
            .ParameterFields(1) = "@vArGroup2;" & vARGroup2 & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With
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

Private Sub CMDPrintCustType_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCustType1 As String
Dim vCustType2 As String
Dim vRepType As String
Dim vRepID As Integer

On Error GoTo ErrDescription

If Me.CMBCustType1.Text <> "" Then
    vCustType1 = Left(Me.CMBCustType1.Text, InStr(1, Me.CMBCustType1.Text, "//") - 1)
    If Me.CMBArGroup2.Text <> "" Then
    vCustType2 = Left(Me.CMBCustType2.Text, InStr(1, Me.CMBCustType2.Text, "//") - 1)
    Else
     vCustType2 = Left(Me.CMBCustType1.Text, InStr(1, Me.CMBCustType1.Text, "//") - 1)
    End If
    
    vRepType = "AR"
    vRepID = 495
     
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        With Crystal103
            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
            .ParameterFields(0) = "@vCustType1;" & vCustType1 & ";true"
            .ParameterFields(1) = "@vCustType2;" & vCustType2 & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With
    End If
    vRecordset.Close
    
    vRepID = 496
    
    If Month(Now) > 6 Then
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        With Crystal103
            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
            .ParameterFields(0) = "@vCustType1;" & vCustType1 & ";true"
            .ParameterFields(1) = "@vCustType2;" & vCustType2 & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With
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

Private Sub CMDPrintKeepCode_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vKeepMoney1 As String
Dim vKeepMoney2 As String
Dim vRepType As String
Dim vRepID As Integer

On Error GoTo ErrDescription

If Me.CMBKeepMoney1.Text <> "" Then
    vKeepMoney1 = Left(Me.CMBKeepMoney1.Text, InStr(1, Me.CMBKeepMoney1.Text, "//") - 1)
    If Me.CMBKeepMoney2.Text <> "" Then
    vKeepMoney2 = Left(Me.CMBKeepMoney2.Text, InStr(1, Me.CMBKeepMoney2.Text, "//") - 1)
    Else
     vKeepMoney2 = Left(Me.CMBKeepMoney1.Text, InStr(1, Me.CMBKeepMoney1.Text, "//") - 1)
    End If
    
    vRepType = "AR"
    vRepID = 497
     
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        With Crystal104
            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
            .ParameterFields(0) = "@vKeepMenCode1;" & vKeepMoney1 & ";true"
            .ParameterFields(1) = "@vKeepMenCode2;" & vKeepMoney2 & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With
    End If
    vRecordset.Close
    
    vRepID = 498
    
    If Month(Now) > 6 Then
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        With Crystal104
            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
            .ParameterFields(0) = "@vKeepMenCode1;" & vKeepMoney1 & ";true"
            .ParameterFields(1) = "@vKeepMenCode2;" & vKeepMoney2 & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With
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

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vGroupARItems As ListItem

On Error GoTo ErrDescription

vQuery = "select distinct code,name from dbo.bccusttype  order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Me.CMBCustType1.AddItem Trim(vRecordset.Fields("code").Value & "//" & vRecordset.Fields("name").Value)
        Me.CMBCustType2.AddItem Trim(vRecordset.Fields("code").Value & "//" & vRecordset.Fields("name").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close


vQuery = "select  distinct isnull(b.keepmoneymencode,'') as keepcode,isnull(a.name,'') as keepname from dbo.bcsale a inner join dbo.bcar b on a.code = isnull(b.keepmoneymencode,'') and a.activestatus = 1  order by isnull(b.keepmoneymencode,'')"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Me.CMBKeepMoney1.AddItem Trim(vRecordset.Fields("keepcode").Value & "//" & vRecordset.Fields("keepname").Value)
        Me.CMBKeepMoney2.AddItem Trim(vRecordset.Fields("keepcode").Value & "//" & vRecordset.Fields("keepname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

vQuery = "select code,name from bcargroup order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Me.CMBArGroup1.AddItem Trim(vRecordset.Fields("code").Value & "//" & vRecordset.Fields("name").Value)
        Me.CMBArGroup2.AddItem Trim(vRecordset.Fields("code").Value & "//" & vRecordset.Fields("name").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub
