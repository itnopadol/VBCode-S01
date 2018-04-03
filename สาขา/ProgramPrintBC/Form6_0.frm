VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form6_0 
   Caption         =   "หน้าพิมพ์เช็ค"
   ClientHeight    =   8205
   ClientLeft      =   2280
   ClientTop       =   570
   ClientWidth     =   12000
   Icon            =   "Form6_0.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form6_0.frx":08CA
   ScaleHeight     =   8205
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport6_01 
      Left            =   1680
      Top             =   6240
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
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComctlLib.ListView ListView6_01 
      Height          =   2115
      Left            =   5925
      TabIndex        =   5
      Top             =   1425
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   3731
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ชื่อฟอร์ม"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.CommandButton CMD6_01 
      Caption         =   "พิมพ์"
      Height          =   690
      Left            =   3675
      TabIndex        =   4
      Top             =   3450
      Width           =   1665
   End
   Begin VB.TextBox TXT6_02 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   2475
      TabIndex        =   1
      Top             =   2250
      Width           =   2865
   End
   Begin VB.TextBox TXT6_01 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2475
      TabIndex        =   0
      Top             =   1425
      Width           =   2865
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์เช็ค"
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
      TabIndex        =   6
      Top             =   300
      Width           =   7290
   End
   Begin VB.Label LBL6_02 
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   1575
      TabIndex        =   3
      Top             =   2250
      Width           =   840
   End
   Begin VB.Label LBL6_01 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เช็ค"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   1575
      TabIndex        =   2
      Top             =   1425
      Width           =   915
   End
End
Attribute VB_Name = "Form6_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD6_01_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCheckNumber As String
Dim vBankCode As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If TXT6_01.Text <> "" And TXT6_02.Text <> "" Then
    vCheckNumber = Trim(TXT6_01.Text)
    
    If TXT6_02.Text = "พิมพ์เช็คธนาคารกรุงเทพฯ" Then
        vRepType = "CH"
        vRepID = 41
    ElseIf TXT6_02.Text = "พิมพ์เช็คธนาคารเอเชีย" Then
                vRepType = "CH"
                vRepID = 42
    End If
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport6_01
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@ChqNumber;" & vCheckNumber & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
Else
    MsgBox "กรุณาใส่เงื่อนไขการพิมพ์เช็คให้ครบด้วยครับ", vbInformation + vbCritical, "ข้อความเตือน"
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
Dim vTypeDoc As String
Dim BDListItems As ListItem

On Error GoTo ErrDescription

ListView6_01.ListItems.Clear
vTypeDoc = "CH"
        vQuery = "select Name from npmaster.dbo.NPForm where ModuleID = '" & vTypeDoc & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set BDListItems = ListView6_01.ListItems.Add(, , Trim(vRecordset.Fields("Name").Value))
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
        
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Private Sub ListView6_01_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT6_02.Text = Item
End Sub

