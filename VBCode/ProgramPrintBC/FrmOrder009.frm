VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrder009 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form009 ค้นหา เส้นทาง"
   ClientHeight    =   5265
   ClientLeft      =   3795
   ClientTop       =   2205
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmOrder009.frx":0000
   ScaleHeight     =   5265
   ScaleWidth      =   8460
   Begin VB.CommandButton CMD103 
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
      Height          =   465
      Left            =   7245
      TabIndex        =   4
      Top             =   4680
      Width           =   1005
   End
   Begin VB.CommandButton CMD102 
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
      Height          =   465
      Left            =   6030
      TabIndex        =   3
      Top             =   4680
      Width           =   1005
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3165
      Left            =   180
      TabIndex        =   2
      Top             =   1395
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   5583
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อที่ 1"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "คำอธิบาย"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ชื่อที่ 2"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   3465
      Picture         =   "FrmOrder009.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "ค้นหา ตามเงื่อนไขที่กรอก"
      Top             =   945
      Width           =   330
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1305
      TabIndex        =   0
      Top             =   945
      Width           =   2130
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาเส้นทาง"
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
      Left            =   180
      TabIndex        =   5
      Top             =   945
      Width           =   1050
   End
End
Attribute VB_Name = "FrmOrder009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem

On Error Resume Next

ListView101.ListItems.Clear
vSearch = Trim(Text101.Text)
vQuery = "exec bcnp.dbo.USP_DO_RouteSearch '" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set vSearchList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("id").Value))
    vSearchList.SubItems(1) = Trim(vRecordset.Fields("name1").Value)
    vSearchList.SubItems(2) = Trim(vRecordset.Fields("mydescription").Value)
    vSearchList.SubItems(3) = Trim(vRecordset.Fields("name2").Value)
    vRecordset.MoveNext
    Wend
Else
    MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
End If
vRecordset.Close
End Sub

Private Sub CMD102_Click()
On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    If vRouteModule = 1 Then
        Form312.Text306.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        Form312.Text307.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        Form312.Text308.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        Form312.Text309.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
    Else
        FrmOrder204.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        FrmOrder204.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        FrmOrder204.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        FrmOrder204.Text104.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        vCheckRouteOpen = 1
    End If
    Unload FrmOrder009
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Unload FrmOrder009
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmProgramPrint.Enabled = True
End Sub

Private Sub ListView101_DblClick()
On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    If vRouteModule = 1 Then
        Form312.Text306.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        Form312.Text307.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        Form312.Text308.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        Form312.Text309.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
    Else
        FrmOrder204.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        FrmOrder204.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        FrmOrder204.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        FrmOrder204.Text104.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        vCheckRouteOpen = 1
    End If
    Unload FrmOrder009
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If ListView101.ListItems.Count <> 0 Then
        If vRouteModule = 1 Then
            Form312.Text306.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
            Form312.Text307.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
            Form312.Text308.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
            Form312.Text309.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        Else
            FrmOrder204.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
            FrmOrder204.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
            FrmOrder204.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
            FrmOrder204.Text104.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
            vCheckRouteOpen = 1
        End If
        Unload FrmOrder009
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem

On Error Resume Next

If KeyAscii = 13 Then
    ListView101.ListItems.Clear
    vSearch = Trim(Text101.Text)
    vQuery = "exec bcnp.dbo.USP_DO_RouteSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("id").Value))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("name1").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("mydescription").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("name2").Value)
        vRecordset.MoveNext
        Wend
    Else
        MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
    End If
    vRecordset.Close
End If
End Sub


