VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmPicker 
   Caption         =   "�ѹ�֡�����ż��Ѵ�Թ���"
   ClientHeight    =   8895
   ClientLeft      =   5130
   ClientTop       =   1410
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmPicker.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   12000
   Begin Crystal.CrystalReport Crystal101 
      Left            =   450
      Top             =   7515
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
   Begin VB.Frame Frame101 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8970
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   11985
      Begin VB.CommandButton CMD102 
         BackColor       =   &H00C0C0C0&
         Caption         =   "�Դ"
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
         Left            =   10755
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6750
         Width           =   1005
      End
      Begin VB.CommandButton CMD101 
         BackColor       =   &H00C0C0C0&
         Caption         =   "���͡"
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
         Left            =   9630
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6750
         Width           =   1005
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   6360
         Left            =   2820
         TabIndex        =   10
         Top             =   330
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   11218
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "�������"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "���;�ѡ�ҹ�Ѵ�Թ���"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ἱ�"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "���ʾ�ѡ�ҹ"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   45
         Picture         =   "FrmPicker.frx":72FB
         Top             =   135
         Width           =   2160
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "���͡��ѡ�ҹ�Ѵ�Թ���"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   19
         Top             =   945
         Width           =   1860
      End
   End
   Begin VB.TextBox TXTPicker 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   2205
      TabIndex        =   0
      Top             =   2160
      Width           =   3615
   End
   Begin MSComctlLib.ListView ListView102 
      Height          =   3660
      Left            =   90
      TabIndex        =   13
      Top             =   3240
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   6456
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�ӴѺ"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�����Թ���"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�����Թ���"
         Object.Width           =   7761
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "�ӹǹ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "˹���"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�����"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Family"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.CommandButton CMD103 
      Height          =   285
      Left            =   5895
      Picture         =   "FrmPicker.frx":875D
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   330
   End
   Begin VB.CommandButton CMDCancel 
      Caption         =   "¡��ԡ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6795
      TabIndex        =   2
      Top             =   2565
      Width           =   780
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "��ŧ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5895
      TabIndex        =   1
      Top             =   2565
      Width           =   780
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�ѹ����� :"
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
      Left            =   4230
      TabIndex        =   21
      Top             =   1080
      Width           =   1050
   End
   Begin VB.Label LBLDocDate 
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
      Height          =   285
      Left            =   5355
      TabIndex        =   20
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��¡���Թ���"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   90
      TabIndex        =   18
      Top             =   2970
      Width           =   1185
   End
   Begin VB.Label LBLRefNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9225
      TabIndex        =   17
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ţ����͡�����ҧ�ԧ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   7470
      TabIndex        =   16
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Label LBLCustName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2205
      TabIndex        =   15
      Top             =   1440
      Width           =   8880
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����١��� :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   990
      TabIndex        =   14
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Label LBLID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2205
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�͡��êش��� :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   990
      TabIndex        =   6
      Top             =   1800
      Width           =   1185
   End
   Begin VB.Label LBLDocno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2205
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ͼ��Ѵ�Թ��� :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   990
      TabIndex        =   4
      Top             =   2160
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ţ����͡��� :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   990
      TabIndex        =   3
      Top             =   1080
      Width           =   1185
   End
End
Attribute VB_Name = "FrmPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String
Dim vDocno As String
Dim vRefDocNo As String
Dim vPicker As String
Dim vSaleOrderNo As String
Dim vSaleCode As String
Dim vTimeID As Integer
Dim vShelfGroup  As String
Dim vWHCode As String
Dim vItemSelect As Integer


Private Sub CMD101_Click()
Dim i As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count > 0 Then
For i = 1 To ListView101.ListItems.Count
       If ListView101.ListItems.Item(i).Checked = True Then
       vItemSelect = i
       Exit For
       End If
Next i

If vItemSelect > 0 Then
   TXTPicker.Text = Trim(ListView101.ListItems.Item(vItemSelect).SubItems(2)) & "/" & Trim(ListView101.ListItems.Item(vItemSelect).Text)
End If
   Frame101.Visible = False
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Frame101.Visible = False
End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListPicker As ListItem
Dim vConnectionString As String
Dim conn As New ADODB.Connection

On Error Resume Next

vConnectionString = "Provider = SQLOLEDB.1;Data Source = Nebula;Initial Catalog = BPLUS4;User ID =VBUSER;PassWord = 132"
conn.Open vConnectionString
ListView101.ListItems.Clear

'If DatePart("w", Now) <> 1 Then
vQuery = "exec bcnp.dbo.USP_HR_PickerZone4 " & vSelectZoneID & ""
'Else
   'If vSelectZoneID = 2 Then
      'vQuery = "exec bcnp.dbo.USP_HR_PickerZone 4 "
   'End If
'End If

vRecordset.Open vQuery, conn, adOpenDynamic, adLockOptimistic
    If Not vRecordset.EOF Then
    vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set vListPicker = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("nickname").Value))
            vListPicker.SubItems(1) = Trim(vRecordset.Fields("picker").Value)
            vListPicker.SubItems(2) = Trim(vRecordset.Fields("dept_thaidesc").Value)
            vListPicker.SubItems(3) = Trim(vRecordset.Fields("prs_no").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
    Frame101.Visible = True
    Me.ListView101.SetFocus
End Sub

Private Sub CMD103_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 And TXTPicker.Text <> "" Then
  Call CMDOK_Click
End If

If KeyCode = 27 Then
  Call CMDCancel_Click
End If
End Sub

Private Sub CMDCancel_Click()
On Error Resume Next

Unload FrmPicker
Call FrmQueue.StartTime
FrmQueue.Text101.Text = ""
FrmQueue.Text101.SetFocus
FrmQueue.ListView101.SelectedItem.Checked = False

End Sub

Private Sub CMDCancel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 And TXTPicker.Text <> "" Then
  Call CMDOK_Click
End If

If KeyCode = 27 Then
  Call CMDCancel_Click
End If
End Sub

Private Sub CMDOK_Click()
Dim vRecordset As New ADODB.Recordset
Dim vShelfGroup As String
Dim vDocDate As String


If TXTPicker.Text <> "" Then
  vDocno = Trim(LBLDocno.Caption)
  vPicker = Trim(TXTPicker.Text)
  vTimeID = LBLID.Caption
  vDocDate = Me.LBLDocDate.Caption
  
  'On Error GoTo ErrDescription
  
  'vQuery = "begin tran"
  'vConnection.Execute (vQuery)
  
  vQuery = "exec dbo.USP_NP_UpdatePrintStatusQueueManagement1  '" & vDocno & "','" & vDocDate & "','" & vPicker & "',1," & vTimeID & ",0 "
  vConnection.Execute (vQuery)
  
  
  'vQuery = "commit tran"
  'vConnection.Execute (vQuery)
  
'ErrDescription:
'If Err.Description <> "" Then
  'vQuery = "rollback tran"
  'vConnection.Execute vQuery
  'MsgBox Err.Description
  'Call FrmQueue.StartTime
  'Exit Sub
  'End If
  Call FrmQueue.StartTime
  'FrmQueue.ListView101.ListItems.Remove (vIndexBegin)
  Call RefreshQueueBegin
  Call RefreshQueuePicking

  FrmQueue.Text101.SetFocus
Else
  MsgBox "�ѧ������͡�����ż��Ѵ�Թ���", vbCritical, "��ͤ�����͹"
End If

'vQuery = "exec dbo.USP_NP_CheckShelfPrintPicking '" & vDocNo & "','" & vTimeID & "' "
'If OpenDataBase(qConnection, vRecordset, vQuery) <> 0 Then
 ' vShelfGroup = Trim(vRecordset.Fields("shelfgroup").Value)
  'vSaleOrderNo = Trim(vRecordset.Fields("saleorderno").Value)
  'vSaleCode = Trim(vRecordset.Fields("saleman").Value)
  'vTimeID = Trim(vRecordset.Fields("timeid").Value)
  'vRefDocNo = Trim(vRecordset.Fields("refdocno").Value)
'End If
'vRecordset.Close
  
'Select Case vShelfGroup
'Case UCase("A"):
 '     Call PrintPicking_A
'Case UCase("B"):
 '     Call PrintPicking_B
'Case UCase("C"):
 '     Call PrintPicking_C
'Case UCase("M"):
 '     Call PrintPicking_M
'Case UCase("H"):
 '     Call PrintPicking_H
'Case UCase("D"):
 '     Call PrintPicking_D
'Case UCase("E"):
 '     Call PrintPicking_E
'Case UCase("Y"):
 '     Call PrintPicking_Y
'Case UCase("O"):
 '     Call PrintPicking_O
'End Select

FrmQueue.Text101.Text = ""
FrmQueue.Text101.SetFocus
FrmQueue.Enabled = True
Unload FrmPicker
End Sub

Private Sub CMDOK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 And TXTPicker.Text <> "" Then
  Call CMDOK_Click
End If

If KeyCode = 27 Then
  Call CMDCancel_Click
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 And TXTPicker.Text <> "" Then
  Call CMDOK_Click
End If

If KeyCode = 27 Then
  Call CMDCancel_Click
End If
End Sub

Private Sub Form_Load()
Call FrmQueue.StopTime
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call FrmQueue.StartTime
Unload FrmPicker
End Sub


Private Sub ListView101_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
Dim vMemCheck As Integer

On Error Resume Next

For i = 1 To Me.ListView101.ListItems.Count
   If Me.ListView101.ListItems.Item(i).Checked = True Then
      vMemCheck = vMemCheck + 1
   End If
Next i

If vMemCheck > 1 Then
   MsgBox "���͡���Ѻ�Դ�ͺ㹡�èѴ�Թ�������§ 1 ����ҹ��", vbCritical, "Send Information Message"
   
For i = 1 To Me.ListView101.ListItems.Count
   If Me.ListView101.ListItems.Item(i).Checked = True Then
      Me.ListView101.ListItems.Item(i).Checked = False
   End If
Next i
Me.ListView101.SetFocus
End If

End Sub

Private Sub ListView102_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 And TXTPicker.Text <> "" Then
  Call CMDOK_Click
End If

If KeyCode = 27 Then
  Call CMDCancel_Click
End If
End Sub

Private Sub Text101_Change()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListPicker As ListItem
Dim vSearch As String
Dim vConnectionString As String
Dim conn As New ADODB.Connection

'vConnectionString = "Provider = SQLOLEDB.1;Data Source = Nebula;Initial Catalog = BPLUS4;User ID =VBUSER;PassWord = 132"
'conn.Open vConnectionString
'ListView101.ListItems.Clear
'vSearch = Text101.Text
'If vSearch = "" Then
'vQuery = "select  *  from bcnp.dbo.vw_HR_Checker"
'Else
'vQuery = "select  *  from bcnp.dbo.vw_HR_Checker where picker like '%'+'" & vSearch & "'+'%' "
'End If
'vRecordset.Open vQuery, conn, adOpenDynamic, adLockOptimistic
    'If Not vRecordset.EOF Then
    'vRecordset.MoveFirst
     '   While Not vRecordset.EOF
      '      Set vListPicker = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("picker").Value))
       '     vListPicker.SubItems(1) = Trim(vRecordset.Fields("nickname").Value)
        'vRecordset.MoveNext
        'Wend
    'End If
    'vRecordset.Close
End Sub

Private Sub TXTPicker_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 And TXTPicker.Text <> "" Then
  Call CMDOK_Click
End If

If KeyCode = 27 Then
  Call CMDCancel_Click
End If

End Sub

Private Sub TXTPicker_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call CMDOK_Click
End If
End Sub

Public Sub PrintPicking_A()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

If vDocno <> "" Then
  vShelfGroup = Trim(UCase("A"))
  vWHCode = Trim("010")
  
  If vTimeID = 1 Then
    vQuery = "exec dbo.USP_NP_InsertPickingDataLogs '" & vSaleOrderNo & "','" & vRefDocNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vUserID & "','" & vSaleCode & "'"
    vConnection.Execute vQuery
  Else
    vQuery = "exec dbo.USP_NP_UpdateCountOfPrintPicking '" & vUserID & "','" & vRefDocNo & "','" & vShelfGroup & "' "
    vConnection.Execute vQuery
  End If
          
  vRepType = "SO"
  
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  vRepID = 323
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  vRepID = 322
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
End If

End Sub

Public Sub PrintPicking_B()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

If vDocno <> "" Then
  vShelfGroup = Trim(UCase("B"))
  vWHCode = Trim("010")
  
  If vTimeID = 1 Then
    vQuery = "exec dbo.USP_NP_InsertPickingDataLogs '" & vSaleOrderNo & "','" & vRefDocNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vUserID & "','" & vSaleCode & "' "
    vConnection.Execute vQuery
  Else
    vQuery = "exec dbo.USP_NP_UpdateCountOfPrintPicking '" & vUserID & "','" & vRefDocNo & "','" & vShelfGroup & "' "
    vConnection.Execute vQuery
  End If
          
  vRepType = "SO"
  
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  vRepID = 323
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  vRepID = 322
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
End If
End Sub


Public Sub PrintPicking_D()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

If vDocno <> "" Then
  vShelfGroup = Trim(UCase("D"))
  vWHCode = Trim("010")
  
  If vTimeID = 1 Then
    vQuery = "exec dbo.USP_NP_InsertPickingDataLogs '" & vSaleOrderNo & "','" & vRefDocNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vUserID & "','" & vSaleCode & "' "
    vConnection.Execute vQuery
  Else
    vQuery = "exec dbo.USP_NP_UpdateCountOfPrintPicking '" & vUserID & "','" & vRefDocNo & "','" & vShelfGroup & "' "
    vConnection.Execute vQuery
  End If
          
  vRepType = "SO"
  
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  vRepID = 323
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  vRepID = 322
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
End If
End Sub

Public Sub PrintPicking_E()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

If vDocno <> "" Then
  vShelfGroup = Trim(UCase("E"))
  vWHCode = Trim("020")
  
  If vTimeID = 1 Then
    vQuery = "exec dbo.USP_NP_InsertPickingDataLogs '" & vSaleOrderNo & "','" & vRefDocNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vUserID & "','" & vSaleCode & "' "
    vConnection.Execute vQuery
  Else
    vQuery = "exec dbo.USP_NP_UpdateCountOfPrintPicking '" & vUserID & "','" & vRefDocNo & "','" & vShelfGroup & "' "
    vConnection.Execute vQuery
  End If
          
  vRepType = "SO"
  
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  vRepID = 323
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  vRepID = 322
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
End If
End Sub

Public Sub PrintPicking_C()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

If vDocno <> "" Then
  vShelfGroup = Trim(UCase("C"))
  vWHCode = Trim("015")
  
  If vTimeID = 1 Then
    vQuery = "exec dbo.USP_NP_InsertPickingDataLogs '" & vSaleOrderNo & "','" & vRefDocNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vUserID & "','" & vSaleCode & "' "
    vConnection.Execute vQuery
  Else
    vQuery = "exec dbo.USP_NP_UpdateCountOfPrintPicking '" & vUserID & "','" & vRefDocNo & "','" & vShelfGroup & "' "
    vConnection.Execute vQuery
  End If
          
  vRepType = "SO"
  
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  vRepID = 323
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  vRepID = 322
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
End If
End Sub

Public Sub PrintPicking_M()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

If vDocno <> "" Then
  vShelfGroup = Trim(UCase("M"))
  vWHCode = Trim("014")
  
  If vTimeID = 1 Then
    vQuery = "exec dbo.USP_NP_InsertPickingDataLogs '" & vSaleOrderNo & "','" & vRefDocNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vUserID & "','" & vSaleCode & "' "
    vConnection.Execute vQuery
  Else
    vQuery = "exec dbo.USP_NP_UpdateCountOfPrintPicking '" & vUserID & "','" & vRefDocNo & "','" & vShelfGroup & "' "
    vConnection.Execute vQuery
  End If
          
  vRepType = "SO"
  
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  vRepID = 323
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  vRepID = 322
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
End If
End Sub

Public Sub PrintPicking_Y()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

If vDocno <> "" Then
  vShelfGroup = Trim(UCase("Y"))
  vWHCode = Trim("016")
  
  If vTimeID = 1 Then
    vQuery = "exec dbo.USP_NP_InsertPickingDataLogs '" & vSaleOrderNo & "','" & vRefDocNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vUserID & "','" & vSaleCode & "' "
    vConnection.Execute vQuery
  Else
    vQuery = "exec dbo.USP_NP_UpdateCountOfPrintPicking '" & vUserID & "','" & vRefDocNo & "','" & vShelfGroup & "' "
    vConnection.Execute vQuery
  End If
          
  vRepType = "SO"
  
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  vRepID = 323
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  vRepID = 322
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
End If
End Sub

Public Sub PrintPicking_H()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

If vDocno <> "" Then
  vShelfGroup = Trim(UCase("H"))
  vWHCode = Trim("014")
  
  If vTimeID = 1 Then
    vQuery = "exec dbo.USP_NP_InsertPickingDataLogs '" & vSaleOrderNo & "','" & vRefDocNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vUserID & "','" & vSaleCode & "' "
    vConnection.Execute vQuery
  Else
    vQuery = "exec dbo.USP_NP_UpdateCountOfPrintPicking '" & vUserID & "','" & vRefDocNo & "','" & vShelfGroup & "' "
    vConnection.Execute vQuery
  End If
          
  vRepType = "SO"
  
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  vRepID = 323
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  vRepID = 322
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
  .ParameterFields(2) = "@vShelfGroup;" & vShelfGroup & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
End If
End Sub


Public Sub PrintPicking_O()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vDocDate As Date

If vDocno <> "" Then
  vShelfGroup = Trim(UCase("O"))
  vWHCode = Trim("014")
  vRepType = "SO"
  vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  vRepID = 324
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vDocDate;" & vDocDate & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  vRepID = 325
  vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  With FrmQueue.Crystal101
  .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
  .ParameterFields(0) = "@vDocNo;" & vSaleOrderNo & ";true"
  .ParameterFields(1) = "@vDocDate;" & vDocDate & ";true"
  .Destination = crptToPrinter
  .WindowState = crptMaximized
  .Action = 1
  End With
  End If
  vRecordset.Close
End If
End Sub
