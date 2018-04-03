VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrder008 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form008 ค้นหาสถานที่ขนส่ง"
   ClientHeight    =   5280
   ClientLeft      =   3795
   ClientTop       =   2205
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmOrder008.frx":0000
   ScaleHeight     =   5280
   ScaleWidth      =   8415
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
      Height          =   420
      Left            =   7245
      TabIndex        =   5
      Top             =   4545
      Width           =   960
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
      Left            =   6075
      TabIndex        =   3
      Top             =   4545
      Width           =   960
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2985
      Left            =   180
      TabIndex        =   2
      Top             =   1440
      Width           =   8025
      _ExtentX        =   14155
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ตำบล"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "อำเภอ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "จังหวัด"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "คำอธิบาย"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   4365
      Picture         =   "FrmOrder008.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "ค้นหา ตามเงื่อนไขที่กรอก"
      Top             =   1035
      Width           =   330
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1620
      TabIndex        =   0
      Top             =   1035
      Width           =   2715
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาสถานที่ขนส่ง"
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
      TabIndex        =   4
      Top             =   1035
      Width           =   1500
   End
End
Attribute VB_Name = "FrmOrder008"
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
vQuery = "exec bcnp.dbo.USP_DO_PlaceSearch '" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set vSearchList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("id").Value))
    vSearchList.SubItems(1) = Trim(vRecordset.Fields("district").Value)
    vSearchList.SubItems(2) = Trim(vRecordset.Fields("amphur").Value)
    vSearchList.SubItems(3) = Trim(vRecordset.Fields("province").Value)
    vSearchList.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
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
    If vPlaceModule = 1 Then
        Form312.Text301.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        Form312.Text302.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        Form312.Text303.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        Form312.Text304.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
    Else
        FrmOrder203.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        FrmOrder203.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        FrmOrder203.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
        FrmOrder203.CMB101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        FrmOrder203.CMB102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        vCheckPlaceOpen = 1
    End If
    Unload FrmOrder008
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Unload FrmOrder008
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmProgramPrint.Enabled = True
End Sub

Private Sub ListView101_DblClick()
On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    If vPlaceModule = 1 Then
        Form312.Text301.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        Form312.Text302.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        Form312.Text303.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        Form312.Text304.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
    Else
        FrmOrder203.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        FrmOrder203.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        FrmOrder203.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
        FrmOrder203.CMB101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        FrmOrder203.CMB102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        vCheckPlaceOpen = 1
    End If
    Unload FrmOrder008
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
        If vPlaceModule = 1 Then
            Form312.Text301.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
            Form312.Text302.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
            Form312.Text303.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
            Form312.Text304.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        Else
            FrmOrder203.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
            FrmOrder203.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
            FrmOrder203.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
            FrmOrder203.CMB101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
            FrmOrder203.CMB102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
            vCheckPlaceOpen = 1
        End If
        Unload FrmOrder008
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
    vQuery = "exec bcnp.dbo.USP_DO_PlaceSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("id").Value))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("district").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("amphur").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("province").Value)
        vSearchList.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
        vRecordset.MoveNext
        Wend
    Else
        MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
    End If
    vRecordset.Close
End If
End Sub

