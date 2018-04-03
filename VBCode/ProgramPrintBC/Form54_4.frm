VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Form54_4 
   Caption         =   "หน้าเคลื่อนไหวลูกหนี้"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form54_4.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   12000
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICSearchAr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   2790
      ScaleHeight     =   6015
      ScaleWidth      =   8670
      TabIndex        =   9
      Top             =   1620
      Visible         =   0   'False
      Width           =   8700
      Begin VB.CommandButton CMDSelect 
         Caption         =   "เลือก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6210
         TabIndex        =   15
         Top             =   5265
         Width           =   1140
      End
      Begin VB.CommandButton CMDClose 
         Caption         =   "ปิด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7425
         TabIndex        =   14
         Top             =   5265
         Width           =   1140
      End
      Begin VB.CommandButton CMDSearchAr 
         Height          =   375
         Left            =   6210
         Picture         =   "Form54_4.frx":9673
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox TXTSearchAr 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   270
         Width           =   5100
      End
      Begin MSComctlLib.ListView ListViewAr 
         Height          =   4380
         Left            =   90
         TabIndex        =   10
         Top             =   810
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   7726
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
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสลูกค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   9701
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ค้นหาลูกค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   11
         Top             =   315
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1800
      Top             =   6750
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
   Begin MSComCtl2.DTPicker DTPDocDate 
      Height          =   420
      Left            =   2790
      TabIndex        =   0
      Top             =   1665
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   65273857
      CurrentDate     =   40676
   End
   Begin VB.CommandButton CMDPrintReport 
      Caption         =   "รายงาน เคลื่อนไหว"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4590
      TabIndex        =   5
      Top             =   3690
      Width           =   1590
   End
   Begin VB.CommandButton BTNSearchAr 
      Height          =   420
      Left            =   5310
      Picture         =   "Form54_4.frx":9AC6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2205
      Width           =   375
   End
   Begin VB.TextBox TXT54_41 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2790
      TabIndex        =   1
      Top             =   2205
      Width           =   2490
   End
   Begin Crystal.CrystalReport CrystalReport54_41 
      Left            =   2430
      Top             =   6750
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
   Begin VB.CommandButton CMD54_41 
      Caption         =   "ดูรายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2790
      TabIndex        =   4
      Top             =   3690
      Width           =   1620
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1215
      TabIndex        =   8
      Top             =   1710
      Width           =   1455
   End
   Begin VB.Label LBLArName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   2790
      TabIndex        =   3
      Top             =   2745
      Width           =   7215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงานเคลื่อนไหวลูกหนี้"
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
      Left            =   3555
      TabIndex        =   7
      Top             =   225
      Width           =   7365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกหนี้ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   720
      TabIndex        =   6
      Top             =   2250
      Width           =   1950
   End
End
Attribute VB_Name = "Form54_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BTNSearchAr_Click()
Me.PICSearchAr.Visible = True
Me.TXTSearchAr.SetFocus
End Sub

Private Sub CMD54_41_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vARCode As String, vRepType As String, vReportName As String
Dim vRepID As Integer

On Error Resume Next

vARCode = Trim(TXT54_41.Text)
If vARCode <> "" Then
    vRepType = "AR"
    vRepID = 86
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcreportname where repid = " & vRepID & "  and reptype = '" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
        With CrystalReport54_41
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@arcode;" & vARCode & ";true"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
        End With
    
    End If
    vRecordset.Close
Else
    MsgBox "กรุณาใส่ข้อมูลดูรายงานให้ครบด้วยครับ", vbInformation + vbCritical, "ข้อความเตือน"
End If
End Sub

Private Sub CMDClose_Click()
Me.PICSearchAr.Visible = False
Me.TXT54_41.SetFocus
End Sub

Private Sub CMDPrintReport_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vARCode As String, vRepType As String, vReportName As String
Dim vRepID As Integer
Dim vDocdate As String

On Error Resume Next

vARCode = Trim(TXT54_41.Text)

If vARCode <> "" And Me.LBLARName.Caption <> "" Then

    vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year
    vRepType = "AR"
    vRepID = 502
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
        With Crystal101
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@vArCode;" & vARCode & ";true"
        .ParameterFields(1) = "@vDocDate;" & vDocdate & ";true"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
        End With
    
    End If
    vRecordset.Close
Else
    MsgBox "กรุณาใส่ข้อมูลลูกค้า ที่จะดูรายงานด้วยครับ", vbInformation + vbCritical, "ข้อความเตือน"
    Me.TXT54_41.SetFocus
End If
End Sub

Private Sub CMDSearchAR_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItem As ListItem
Dim n As Double
Dim vSearch As String

On Error Resume Next

vSearch = Me.TXTSearchAr.Text

Me.ListViewAr.ListItems.Clear

vQuery = "exec dbo.USP_NP_SearchCustomer '" & vSearch & "'"
If OpenDataBase(gConnection, vRecordset, vQuery) > 0 Then
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   n = n + 1
   Set vListItem = Me.ListViewAr.ListItems.Add(, , n)
   vListItem.SubItems(1) = vRecordset.Fields("code").Value
   vListItem.SubItems(2) = vRecordset.Fields("name1").Value
   vRecordset.MoveNext
   Wend
Else
Me.ListViewAr.ListItems.Clear
Me.TXTSearchAr.SetFocus
End If
vRecordset.Close

Me.ListViewAr.Visible = True
Me.ListViewAr.SetFocus
End Sub

Private Sub CMDSelect_Click()
On Error Resume Next

If Me.ListViewAr.ListItems.Count > 0 Then
Me.TXT54_41.Text = Me.ListViewAr.ListItems(Me.ListViewAr.SelectedItem.Index).SubItems(1)
Me.PICSearchAr.Visible = False
Me.TXT54_41.SetFocus
End If
End Sub

Private Sub ListViewAr_DblClick()
On Error Resume Next

If Me.ListViewAr.ListItems.Count > 0 Then
Me.TXT54_41.Text = Me.ListViewAr.ListItems(Me.ListViewAr.SelectedItem.Index).SubItems(1)
Me.PICSearchAr.Visible = False
Me.TXT54_41.SetFocus
End If
End Sub

Private Sub ListViewAr_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
    If Me.ListViewAr.ListItems.Count > 0 Then
    Me.TXT54_41.Text = Me.ListViewAr.ListItems(Me.ListViewAr.SelectedItem.Index).SubItems(1)
    Me.PICSearchAr.Visible = False
    Me.TXT54_41.SetFocus
    End If
End If
End Sub

Private Sub TXT54_41_Change()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode As String

On Error Resume Next

vARCode = Me.TXT54_41.Text

vQuery = "exec dbo.USP_NP_SearchArCode '" & vARCode & "'"
If OpenDataBase(gConnection, vRecordset, vQuery) > 0 Then
   Me.LBLARName.Caption = vRecordset.Fields("name1").Value
Else
Me.LBLARName.Caption = ""
Me.TXT54_41.SetFocus
End If
vRecordset.Close

End Sub

Private Sub TXT54_41_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode As String

On Error Resume Next

If KeyAscii = 13 Then
    vARCode = Me.TXT54_41.Text
    
    vQuery = "exec dbo.USP_NP_SearchArCode '" & vARCode & "'"
    If OpenDataBase(gConnection, vRecordset, vQuery) > 0 Then
       Me.LBLARName.Caption = vRecordset.Fields("name1").Value
    Else
    Me.LBLARName.Caption = ""
    Me.TXT54_41.SetFocus
    End If
    vRecordset.Close
End If
End Sub

Private Sub TXTSearchAr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CMDSearchAR_Click
End If
End Sub
