VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrder007 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form007 ค้นหาผู้รับสินค้า"
   ClientHeight    =   5235
   ClientLeft      =   3795
   ClientTop       =   2205
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmOrder007.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   8415
   Begin VB.CommandButton CMD104 
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
      Left            =   7200
      TabIndex        =   5
      Top             =   4635
      Width           =   1050
   End
   Begin VB.CommandButton CMD103 
      Caption         =   "เพิ่มข้อมูล"
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
      Left            =   4770
      TabIndex        =   4
      Top             =   4635
      Width           =   1050
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
      Height          =   420
      Left            =   5985
      TabIndex        =   3
      Top             =   4635
      Width           =   1050
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2985
      Left            =   180
      TabIndex        =   2
      Top             =   1530
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   5265
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "คำนำหน้าชื่อ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "นามสกุล"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เบอร์บ้าน"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "เบอร์มือถือ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "หมายเหตุ"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   3870
      Picture         =   "FrmOrder007.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "ค้นหา ตามเงื่อนไขที่กรอก"
      Top             =   1080
      Width           =   330
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1575
      TabIndex        =   0
      Top             =   1080
      Width           =   2265
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหา ผู้รับสินค้า"
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
      Left            =   225
      TabIndex        =   6
      Top             =   1125
      Width           =   1500
   End
End
Attribute VB_Name = "FrmOrder007"
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
    vQuery = "exec bcnp.dbo.USP_DO_ReceiveSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("id").Value))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("titlename").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("firstname").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("surname").Value)
        vSearchList.SubItems(4) = Trim(vRecordset.Fields("homephone").Value)
        vSearchList.SubItems(5) = Trim(vRecordset.Fields("mobilephone").Value)
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
    If vReceiveModule = 1 Then
        'Form312.Text201.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        'Form312.Text202.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        'Form312.Text203.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        'Form312.Text204.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        'Form312.Text205.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
        'Form312.Text206.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5))
    ElseIf vReceiveModule = 2 Then
        'FrmOrder202.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        'FrmOrder202.CMB101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        'FrmOrder202.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        'FrmOrder202.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        'FrmOrder202.Text104.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(6))
        'FrmOrder202.MaskEdBox101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
        'FrmOrder202.MaskEdBox102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5))
        vCheckReceiveOpen = 1
        vCheckChageDataReceive = 0
    End If
    Unload FrmOrder007
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
FrmOrder202.Show
vCheckAddReceiver = 1
MDIFrmProgramPrint.Enabled = True
End Sub

Private Sub CMD104_Click()
Unload FrmOrder007
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmProgramPrint.Enabled = True
End Sub

Private Sub ListView101_DblClick()
On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    If vReceiveModule = 1 Then
        Form312.Text201.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        Form312.Text202.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        Form312.Text203.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        'Form312.Text204.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        'Form312.Text205.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
        'Form312.Text206.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5))
    ElseIf vReceiveModule = 2 Then
        'FrmOrder202.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        'FrmOrder202.CMB101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        'FrmOrder202.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        'FrmOrder202.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        'FrmOrder202.Text104.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(6))
        'FrmOrder202.MaskEdBox101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
        'FrmOrder202.MaskEdBox102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5))
        vCheckReceiveOpen = 1
        vCheckChageDataReceive = 0
    End If
    Unload FrmOrder007
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
        If vReceiveModule = 1 Then
            'Form312.Text201.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
            'Form312.Text202.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
            'Form312.Text203.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
            'Form312.Text204.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
            'Form312.Text205.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
            'Form312.Text206.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5))
        ElseIf vReceiveModule = 2 Then
            'FrmOrder202.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
            'FrmOrder202.CMB101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
            'FrmOrder202.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
            'FrmOrder202.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
            'FrmOrder202.Text104.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(6))
            'FrmOrder202.MaskEdBox101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
            'FrmOrder202.MaskEdBox102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5))
            vCheckReceiveOpen = 1
            vCheckChageDataReceive = 0
        End If
        Unload FrmOrder007
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
    vQuery = "exec bcnp.dbo.USP_DO_ReceiveSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("id").Value))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("titlename").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("firstname").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("surname").Value)
        vSearchList.SubItems(4) = Trim(vRecordset.Fields("homephone").Value)
        vSearchList.SubItems(5) = Trim(vRecordset.Fields("mobilephone").Value)
        vSearchList.SubItems(6) = Trim(vRecordset.Fields("mydescription").Value)
        vRecordset.MoveNext
        Wend
    Else
        MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
    End If
    vRecordset.Close
End If

End Sub

