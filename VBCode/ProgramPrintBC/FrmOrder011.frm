VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrder011 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form011 ค้นหารถขนส่ง"
   ClientHeight    =   5100
   ClientLeft      =   3870
   ClientTop       =   2280
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmOrder011.frx":0000
   ScaleHeight     =   5100
   ScaleWidth      =   8385
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
      Left            =   7200
      TabIndex        =   5
      Top             =   4455
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
      Left            =   5895
      TabIndex        =   3
      Top             =   4455
      Width           =   1005
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2895
      Left            =   180
      TabIndex        =   2
      Top             =   1440
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   5106
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขรถ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ทะเบียนรถ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ชื่อรถ 1"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ชื่อรถ 2"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "เลขตัวรถ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "เลขเครื่อง"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ลักษณะมาตรฐาน"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "วันที่ซื้อ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "เลขที่กรมธรรม์"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "อัตราเฉลี่ย/ลิตร"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "กว้าง"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "ยาว"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "สูง"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "คำอธิบาย"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   3690
      Picture         =   "FrmOrder011.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "ค้นหา ตามเงื่อนไขที่กรอก"
      Top             =   990
      Width           =   330
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1305
      TabIndex        =   0
      Top             =   990
      Width           =   2310
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหารถขนส่ง"
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
      Left            =   180
      TabIndex        =   4
      Top             =   1035
      Width           =   1230
   End
End
Attribute VB_Name = "FrmOrder011"
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
    vQuery = "exec bcnp.dbo.USP_DO_VehicalSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("id").Value))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("carno").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("CarLicence").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("name1").Value)
        vSearchList.SubItems(4) = Trim(vRecordset.Fields("name2").Value)
        vSearchList.SubItems(5) = Trim(vRecordset.Fields("bodynumber").Value)
        vSearchList.SubItems(6) = Trim(vRecordset.Fields("enginenumber").Value)
        vSearchList.SubItems(7) = Trim(vRecordset.Fields("standardtype").Value)
        vSearchList.SubItems(8) = Trim(vRecordset.Fields("datebuy").Value)
        vSearchList.SubItems(9) = Trim(vRecordset.Fields("insurancenumber").Value)
        vSearchList.SubItems(10) = Trim(vRecordset.Fields("distancerate").Value)
        vSearchList.SubItems(11) = Trim(vRecordset.Fields("wide").Value)
        vSearchList.SubItems(12) = Trim(vRecordset.Fields("long").Value)
        vSearchList.SubItems(13) = Trim(vRecordset.Fields("high").Value)
        vSearchList.SubItems(14) = Trim(vRecordset.Fields("mydescription").Value)
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
    If vVehicalModule = 1 Then
        FormDelivery.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        FormDelivery.Text104.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        FormDelivery.Text105.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
    Else
        FrmOrder205.Text101 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        FrmOrder205.Text102 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        FrmOrder205.Text103 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        FrmOrder205.Text104 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        FrmOrder205.Text105 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
        FrmOrder205.Text106 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5))
        FrmOrder205.Text107 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(6))
        FrmOrder205.Text108 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(7))
        FrmOrder205.Text109 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(9))
        FrmOrder205.Text110 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(10))
        FrmOrder205.Text111 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(11))
        FrmOrder205.Text112 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(12))
        FrmOrder205.Text113 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(13))
        FrmOrder205.Text114 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(14))
        FrmOrder205.DTPicker101 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(8))
        vCheckVehicalOpen = 1
    End If
    Unload FrmOrder011
End If
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Unload FrmOrder011
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmProgramPrint.Enabled = True
End Sub

Private Sub ListView101_DblClick()
On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    If vVehicalModule = 1 Then
        FormDelivery.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        FormDelivery.Text104.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        FormDelivery.Text105.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
    Else
        FrmOrder205.Text101 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
        FrmOrder205.Text102 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        FrmOrder205.Text103 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        FrmOrder205.Text104 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        FrmOrder205.Text105 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
        FrmOrder205.Text106 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5))
        FrmOrder205.Text107 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(6))
        FrmOrder205.Text108 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(7))
        FrmOrder205.Text109 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(9))
        FrmOrder205.Text110 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(10))
        FrmOrder205.Text111 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(11))
        FrmOrder205.Text112 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(12))
        FrmOrder205.Text113 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(13))
        FrmOrder205.Text114 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(14))
        FrmOrder205.DTPicker101 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(8))
        vCheckVehicalOpen = 1
    End If
    Unload FrmOrder011
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    If KeyAscii = 13 Then
        If vVehicalModule = 1 Then
            FormDelivery.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
            FormDelivery.Text104.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
            FormDelivery.Text105.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        Else
            FrmOrder205.Text101 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).Text)
            FrmOrder205.Text102 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
            FrmOrder205.Text103 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
            FrmOrder205.Text104 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
            FrmOrder205.Text105 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
            FrmOrder205.Text106 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5))
            FrmOrder205.Text107 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(6))
            FrmOrder205.Text108 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(7))
            FrmOrder205.Text109 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(9))
            FrmOrder205.Text110 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(10))
            FrmOrder205.Text111 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(11))
            FrmOrder205.Text112 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(12))
            FrmOrder205.Text113 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(13))
            FrmOrder205.Text114 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(14))
            FrmOrder205.DTPicker101 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(8))
            vCheckVehicalOpen = 1
        End If
        Unload FrmOrder011
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
    vQuery = "exec bcnp.dbo.USP_DO_VehicalSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("id").Value))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("carno").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("CarLicence").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("name1").Value)
        vSearchList.SubItems(4) = Trim(vRecordset.Fields("name2").Value)
        vSearchList.SubItems(5) = Trim(vRecordset.Fields("bodynumber").Value)
        vSearchList.SubItems(6) = Trim(vRecordset.Fields("enginenumber").Value)
        vSearchList.SubItems(7) = Trim(vRecordset.Fields("standardtype").Value)
        vSearchList.SubItems(8) = Trim(vRecordset.Fields("datebuy").Value)
        vSearchList.SubItems(9) = Trim(vRecordset.Fields("insurancenumber").Value)
        vSearchList.SubItems(10) = Trim(vRecordset.Fields("distancerate").Value)
        vSearchList.SubItems(11) = Trim(vRecordset.Fields("wide").Value)
        vSearchList.SubItems(12) = Trim(vRecordset.Fields("long").Value)
        vSearchList.SubItems(13) = Trim(vRecordset.Fields("high").Value)
        vSearchList.SubItems(14) = Trim(vRecordset.Fields("mydescription").Value)
        vRecordset.MoveNext
        Wend
    Else
        MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
    End If
    vRecordset.Close
End If
End Sub

