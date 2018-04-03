VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmPicker 
   Caption         =   "บันทึกข้อมูลผู้จัดสินค้า"
   ClientHeight    =   9495
   ClientLeft      =   2685
   ClientTop       =   990
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmPicker.frx":0000
   ScaleHeight     =   9495
   ScaleWidth      =   15180
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
      Height          =   9690
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   15225
      Begin VB.CommandButton CMD102 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ปิด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   11565
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7425
         Width           =   1275
      End
      Begin VB.CommandButton CMD101 
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
         Height          =   555
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7425
         Width           =   1275
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   5820
         Left            =   2340
         TabIndex        =   10
         Top             =   1395
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   10266
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
            Text            =   "ชื่อเล่น"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อพนักงานจัดสินค้า"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "แผนก"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "รหัสพนักงาน"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   45
         Picture         =   "FrmPicker.frx":9673
         Top             =   135
         Width           =   2160
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "เลือกพนักงานจัดสินค้า"
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
         Left            =   2340
         TabIndex        =   19
         Top             =   1035
         Width           =   1860
      End
   End
   Begin VB.PictureBox PICPoint 
      Height          =   240
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   285
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
      Left            =   1530
      TabIndex        =   0
      Top             =   2340
      Width           =   3615
   End
   Begin MSComctlLib.ListView ListView102 
      Height          =   4740
      Left            =   315
      TabIndex        =   13
      Top             =   3420
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   8361
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "จำนวน"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หน่วย"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ชั้นเก็บ"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.CommandButton CMD103 
      Height          =   285
      Left            =   5220
      Picture         =   "FrmPicker.frx":AAD5
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2340
      Width           =   330
   End
   Begin VB.CommandButton CMDCancel 
      Caption         =   "ยกเลิก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6615
      TabIndex        =   3
      Top             =   2745
      Width           =   1320
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "ตกลง"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5220
      TabIndex        =   2
      Top             =   2745
      Width           =   1320
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่คิว :"
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
      Left            =   3555
      TabIndex        =   21
      Top             =   1260
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
      Left            =   4680
      TabIndex        =   20
      Top             =   1260
      Width           =   1860
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "รายการสินค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   315
      TabIndex        =   18
      Top             =   3105
      Width           =   1185
   End
   Begin VB.Label LBLRefNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8550
      TabIndex        =   17
      Top             =   1260
      Width           =   1860
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสารอ้างอิง :"
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
      Left            =   6795
      TabIndex        =   16
      Top             =   1260
      Width           =   1680
   End
   Begin VB.Label LBLCustName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1530
      TabIndex        =   15
      Top             =   1620
      Width           =   8880
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า :"
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
      Left            =   315
      TabIndex        =   14
      Top             =   1620
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
      Left            =   1530
      TabIndex        =   8
      Top             =   1980
      Width           =   915
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "เอกสารชุดที่ :"
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
      Left            =   315
      TabIndex        =   7
      Top             =   1980
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
      Left            =   1530
      TabIndex        =   6
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อผู้จัดสินค้า :"
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
      Left            =   315
      TabIndex        =   5
      Top             =   2340
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
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
      Left            =   315
      TabIndex        =   4
      Top             =   1260
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
   TXTPicker.Text = Trim(ListView101.ListItems.Item(vItemSelect).SubItems(3)) & "/" & Trim(ListView101.ListItems.Item(vItemSelect).Text)
End If
   Frame101.Visible = False
End If

Me.CMDOK.SetFocus

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
Dim vZoneID As String

'On Error Resume Next

ListView101.ListItems.Clear

If vSelectZoneID = 1 Then
vZoneID = "A"
ElseIf vSelectZoneID = 2 Then
vZoneID = "B"
ElseIf vSelectZoneID = 3 Then
vZoneID = "C"
End If

If DatePart("w", Now) <> 1 Then
   vQuery = "exec dbo.USP_NP_SearchPickerByZone '" & vZoneID & "' "
Else
   If vSelectZoneID = 2 Then
      vQuery = "exec dbo.USP_NP_SearchPickerByZone4 "
      Else
       vQuery = "exec dbo.USP_NP_SearchPickerByZone '" & vZoneID & "' "
   End If
End If

If OpenDataBase2(vConnection, vRecordset, vQuery) <> 0 Then
    While Not vRecordset.EOF
            Set vListPicker = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("nickname").Value))
            vListPicker.SubItems(1) = Right(Trim(vRecordset.Fields("name1").Value), Len(Trim(vRecordset.Fields("name1").Value)) - InStr(Trim(vRecordset.Fields("name1").Value), "/"))
            vListPicker.SubItems(2) = "" 'Trim(vRecordset.Fields("dutycode").Value)
            vListPicker.SubItems(3) = Left(Trim(vRecordset.Fields("name1").Value), InStr(Trim(vRecordset.Fields("name1").Value), "/") - 1)
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
  
  vQuery = "exec dbo.USP_NP_UpdateQueStatusDetails  '" & vDocno & "','" & vDocDate & "','" & vPicker & "',1," & vTimeID & ",0"
  vConnection.Execute (vQuery)
  
  Call FrmQueue.StartTime
  Call RefreshQueueBegin
  Call RefreshQueuePicking

  FrmQueue.Text101.SetFocus
Else
  MsgBox "ยังไม่ได้กรอกข้อมูลผู้จัดสินค้า", vbCritical, "ข้อความเตือน"
End If


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
Call SetListViewColor(ListView102, PICPoint, vbWhite, vbLightGreen)
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
   MsgBox "เลือกคนรับผิดชอบในการจัดสินค้าได้เพียง 1 คนเท่านั้น", vbCritical, "Send Information Message"
   
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

Public Sub SetListViewColor(pCtrlListView As ListView, pCtrlPictureBox As PictureBox, Color1 As Long, Color2 As Long)

On Error GoTo SetListViewColor_Error

    Dim iLineHeight As Long
    Dim iBarHeight  As Long
    Dim lBarWidth   As Long
    Dim lColor1     As Long
    Dim lColor2     As Long
 
    lColor1 = Color1
    lColor2 = Color2
    
    If pCtrlListView.View = lvwReport Then
        pCtrlListView.Picture = LoadPicture("")
        pCtrlListView.Refresh
        pCtrlPictureBox.Cls
        
        pCtrlPictureBox.AutoRedraw = True
        pCtrlPictureBox.BorderStyle = vbBSNone
        pCtrlPictureBox.ScaleMode = vbTwips
        pCtrlPictureBox.Visible = False
        
        pCtrlListView.PictureAlignment = lvwTile
        pCtrlPictureBox.Font = pCtrlListView.Font
        pCtrlPictureBox.Top = pCtrlListView.Top
        pCtrlPictureBox.Font = pCtrlListView.Font
        With pCtrlPictureBox.Font
            .Size = pCtrlListView.Font.Size '+ 2.75
            .Bold = pCtrlListView.Font.Bold
            .Charset = pCtrlListView.Font.Charset
            .Italic = pCtrlListView.Font.Italic
            .Name = pCtrlListView.Font.Name
            .Strikethrough = pCtrlListView.Font.Strikethrough
            .Underline = pCtrlListView.Font.Underline
            .Weight = pCtrlListView.Font.Weight
        End With
        pCtrlPictureBox.Refresh
        iLineHeight = pCtrlPictureBox.TextHeight("W") + Screen.TwipsPerPixelY
    
        iBarHeight = (iLineHeight * 1)
        lBarWidth = pCtrlListView.Width
    
        pCtrlPictureBox.Height = iBarHeight * 2
        pCtrlPictureBox.Width = lBarWidth
    
        pCtrlPictureBox.Line (0, 0)-(lBarWidth, iBarHeight), lColor1, BF
        pCtrlPictureBox.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), lColor2, BF
    
        pCtrlPictureBox.AutoSize = True
        pCtrlListView.Picture = pCtrlPictureBox.Image
    Else
        pCtrlListView.Picture = LoadPicture("")
    End If
    
    pCtrlListView.Refresh
    Exit Sub
SetListViewColor_Error:
    pCtrlListView.Picture = LoadPicture("")
    pCtrlListView.Refresh
End Sub

