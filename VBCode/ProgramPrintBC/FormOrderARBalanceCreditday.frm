VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FormOrderARBalanceCreditday 
   Caption         =   "รายงาน ยอดลูกหนี้คงค้างตามเครดิต"
   ClientHeight    =   11010
   ClientLeft      =   2445
   ClientTop       =   255
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
   Picture         =   "FormOrderARBalanceCreditday.frx":0000
   ScaleHeight     =   41609.42
   ScaleMode       =   0  'User
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CMBPressMen2 
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
      Left            =   3000
      TabIndex        =   17
      Top             =   5340
      Width           =   4575
   End
   Begin VB.ComboBox CMBPressMen1 
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
      Left            =   3000
      TabIndex        =   16
      Top             =   4680
      Width           =   4575
   End
   Begin VB.CommandButton CMDListviewClose 
      Caption         =   "ปิด หน้าลูกหนี้"
      Height          =   375
      Left            =   12510
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   1320
   End
   Begin MSComctlLib.ListView ListViewSearchAr 
      Height          =   1515
      Left            =   5445
      TabIndex        =   14
      Top             =   2790
      Visible         =   0   'False
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   2672
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสลูกค้า"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.CommandButton CMDSearchAr 
      Height          =   375
      Left            =   4950
      TabIndex        =   13
      Top             =   2205
      Width           =   420
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   270
      Top             =   8010
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
      Height          =   735
      Left            =   3015
      TabIndex        =   12
      Top             =   6000
      Width           =   1860
   End
   Begin VB.TextBox TXTToAr 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3015
      TabIndex        =   11
      Top             =   2790
      Width           =   1860
   End
   Begin VB.TextBox TXTFromAr 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3015
      TabIndex        =   8
      Top             =   2205
      Width           =   1860
   End
   Begin MSComCtl2.DTPicker DTPToDate 
      Height          =   375
      Left            =   3015
      TabIndex        =   7
      Top             =   4005
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16449537
      CurrentDate     =   40525
   End
   Begin MSComCtl2.DTPicker DTPFromDate 
      Height          =   420
      Left            =   3015
      TabIndex        =   6
      Top             =   3375
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16449537
      CurrentDate     =   40525
   End
   Begin VB.ComboBox CMBCreditDay 
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
      Left            =   3015
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1665
      Width           =   1230
   End
   Begin VB.Label Label7 
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
      Height          =   345
      Left            =   720
      TabIndex        =   19
      Top             =   5340
      Width           =   2205
   End
   Begin VB.Label Label6 
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
      Height          =   375
      Left            =   810
      TabIndex        =   18
      Top             =   4710
      Width           =   2115
   End
   Begin VB.Label LBLToAr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5445
      TabIndex        =   10
      Top             =   2790
      Width           =   8385
   End
   Begin VB.Label LBLFromAr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5445
      TabIndex        =   9
      Top             =   2205
      Width           =   8385
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่ครบกำหนด :"
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
      Left            =   1125
      TabIndex        =   5
      Top             =   4050
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่ครบกำหนด :"
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
      Left            =   1170
      TabIndex        =   4
      Top             =   3420
      Width           =   1770
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงรหัสลูกค้า :"
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
      Height          =   330
      Left            =   1125
      TabIndex        =   3
      Top             =   2835
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากรหัสลูกค้า :"
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
      Height          =   330
      Left            =   1125
      TabIndex        =   2
      Top             =   2250
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวนวันเครดิต :"
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
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1665
      Width           =   2220
   End
End
Attribute VB_Name = "FormOrderARBalanceCreditday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDListviewClose_Click()
Me.ListViewSearchAr.Visible = False
Me.CMDListviewClose.Visible = False
End Sub

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCreditDay As Integer
Dim vFromAr As String
Dim vToAr As String
Dim vFromDate As String
Dim vToDate As String
Dim vRepID As Integer
Dim vRepType As String
Dim vPressMenCode1 As String
Dim vPressMenCode2 As String

On Error Resume Next

If Me.TXTFromAr.Text <> "" Then
    vCreditDay = Me.CMBCreditDay.Text
    vFromAr = Me.TXTFromAr.Text
    
    If Me.TXTToAr.Text = "" Then
    vToAr = Me.TXTFromAr.Text
    Else
    vToAr = Me.TXTToAr.Text
    End If

    vFromDate = Me.DTPFromDate.Value
    vToDate = Me.DTPToDate.Value
    
    If Me.CMBPressMen1.Text <> "" Then
        vPressMenCode1 = Left(Me.CMBPressMen1.Text, InStr(Me.CMBPressMen1.Text, "/") - 1)
    Else
        vPressMenCode1 = ""
    End If
    
    
    If Me.CMBPressMen2.Text <> "" Then
        vPressMenCode2 = Left(Me.CMBPressMen2.Text, InStr(Me.CMBPressMen2.Text, "/") - 1)
    Else
        vPressMenCode2 = ""
    End If
    
    If vPressMenCode2 = "" And vPressMenCode1 <> "" Then
        vPressMenCode2 = vPressMenCode1
    End If
    
    vRepID = 490
    vRepType = "AR"
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        With Crystal101
            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
            .ParameterFields(0) = "@vCreditDay;" & vCreditDay & ";true"
            .ParameterFields(1) = "@vFromAr;" & vFromAr & ";true"
            .ParameterFields(2) = "@vToAr;" & vToAr & ";true"
            .ParameterFields(3) = "@vFromDueDate;" & vFromDate & ";true"
            .ParameterFields(4) = "@vToDueDate;" & vToDate & ";true"
            .ParameterFields(5) = "@vPressMenCode1;" & vPressMenCode1 & ";true"
            .ParameterFields(6) = "@vPressMenCode2;" & vPressMenCode2 & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With
    End If
    vRecordset.Close
    
End If

End Sub

Private Sub CMDSearchAR_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListviewAR As ListItem
Dim i As Double

On Error Resume Next

If Me.TXTFromAr.Text = "" Or Me.TXTToAr.Text = "" Then
    If Me.ListViewSearchAr.ListItems.Count > 0 Then
    Me.ListViewSearchAr.Visible = True
    Me.CMDListviewClose.Visible = True
    Else
    Me.ListViewSearchAr.Visible = True
    Me.CMDListviewClose.Visible = True
    vQuery = "select code,name1 as arname  from dbo.bcar where activestatus = 1 order by code "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        i = 1
        While Not vRecordset.EOF
        Set vListviewAR = ListViewSearchAr.ListItems.Add(, , i)
        vListviewAR.SubItems(1) = Trim(vRecordset.Fields("code").Value)
        vListviewAR.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
        i = i + 1
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
    End If
End If
End Sub

Private Sub Form_Load()
Call AddCreditDay
Call GetPressMen
Me.DTPFromDate.Value = Now
Me.DTPToDate.Value = Now
Me.CMBCreditDay.ListIndex = 0
End Sub

Public Sub AddCreditDay()
Me.CMBCreditDay.AddItem (0)
Me.CMBCreditDay.AddItem (1)
Me.CMBCreditDay.AddItem (2)
Me.CMBCreditDay.AddItem (3)
Me.CMBCreditDay.AddItem (4)
Me.CMBCreditDay.AddItem (5)
Me.CMBCreditDay.AddItem (6)
Me.CMBCreditDay.AddItem (7)
Me.CMBCreditDay.AddItem (8)
Me.CMBCreditDay.AddItem (9)
Me.CMBCreditDay.AddItem (10)
Me.CMBCreditDay.AddItem (11)
Me.CMBCreditDay.AddItem (12)
Me.CMBCreditDay.AddItem (13)
Me.CMBCreditDay.AddItem (14)
Me.CMBCreditDay.AddItem (15)
End Sub

Public Sub GetPressMen()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

vQuery = "select  * from (select  distinct pressmencode+'/'+name as pressmen  from dbo.BCAR a inner join dbo.bcsale b on a.pressmencode = b.code where a.activestatus = 1 and b.activestatus = 1 union select 'N/A' as pressmen) as  pm order by pressmen "
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

Private Sub ListViewSearchAR_DblClick()
Dim vIndex As Double

On Error Resume Next

If Me.ListViewSearchAr.ListItems.Count > 0 Then
    vIndex = Me.ListViewSearchAr.SelectedItem.Index
    
    If Me.TXTFromAr.Text = "" Then
        Me.TXTFromAr.Text = Me.ListViewSearchAr.ListItems(vIndex).SubItems(1)
        Me.LBLFromAr.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(2)
        Me.ListViewSearchAr.Visible = False
        Me.CMDListviewClose.Visible = False
        Exit Sub
    End If
    
    If Me.TXTFromAr.Text <> "" And Me.TXTToAr.Text = "" Then
        Me.TXTToAr.Text = Me.ListViewSearchAr.ListItems(vIndex).SubItems(1)
        Me.LBLToAr.Caption = Me.ListViewSearchAr.ListItems(vIndex).SubItems(2)
        Me.ListViewSearchAr.Visible = False
        Me.CMDListviewClose.Visible = False
        Exit Sub
    End If

End If
End Sub

Private Sub TXTFromAr_Change()
If Me.TXTFromAr.Text = "" Then
Me.LBLFromAr.Caption = ""
End If
End Sub

Private Sub TXTToAr_Change()
If Me.TXTToAr.Text = "" Then
Me.LBLToAr.Caption = ""
End If
End Sub
