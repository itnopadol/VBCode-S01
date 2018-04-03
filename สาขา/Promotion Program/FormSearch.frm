VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSearchPromotion 
   Caption         =   "ค้นหา ทะเบียนโปรโมชั่น"
   ClientHeight    =   5430
   ClientLeft      =   4500
   ClientTop       =   2115
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8820
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   300
      Width           =   2715
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   4140
      Left            =   375
      TabIndex        =   0
      Top             =   825
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   7303
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
         Text            =   "รหัสโปรโมชั่น"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อโปรโมชั่น"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันเริ่ม"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันสิ้นสุด"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "คำอธิบาย"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "ชื่อโปรโมชั่น"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   375
      TabIndex        =   2
      Top             =   300
      Width           =   1065
   End
End
Attribute VB_Name = "FormSearchPromotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListPromotion As ListItem

On Error GoTo ErrDescription

ListView101.ListItems.Clear
vQuery = "execute USP_PM_Find ''"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Set vListPromotion = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("pmcode").Value))
        vListPromotion.SubItems(1) = Trim(vRecordset.Fields("pmname").Value)
        vListPromotion.SubItems(2) = Trim(vRecordset.Fields("datestart").Value)
        vListPromotion.SubItems(3) = Trim(vRecordset.Fields("dateend").Value)
        If Not IsNull(Trim(vRecordset.Fields("mydescription").Value)) Then
        vListPromotion.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
        Else
         vListPromotion.SubItems(4) = ""
        End If
        
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

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Enabled = True
End Sub

Private Sub ListView101_DblClick()
Dim vIndex As Integer

On Error GoTo ErrDescription

vIndex = ListView101.SelectedItem.Index
vCheckJob = 0
Form101.Command103.Enabled = True
Form101.Check101.Enabled = True
MDIForm1.Enabled = True
Form101.Text101.Text = Trim(ListView101.ListItems.Item(vIndex).Text)
Form101.Text102.Text = Trim(ListView101.ListItems.Item(vIndex).SubItems(1))
Form101.DTPicker101.Value = Trim(ListView101.ListItems.Item(vIndex).SubItems(2))
Form101.DTPicker102.Value = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
Form101.Text103.Text = Trim(ListView101.ListItems.Item(vIndex).SubItems(4))
Unload FormSearchPromotion

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo ErrDescription

'vCheckJob = 0
'Form101.Command103.Enabled = True
'Form101.Check101.Enabled = True
'MDIForm1.Enabled = True
'Form101.Text101.Text = Trim(ListView101.ListItems.Item(Item.Index).Text)
'Form101.Text102.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(1))
'Form101.DTPicker101.Value = Trim(ListView101.ListItems.Item(Item.Index).SubItems(2))
'Form101.DTPicker102.Value = Trim(ListView101.ListItems.Item(Item.Index).SubItems(3))
'Form101.Text103.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(4))
'Unload FormSearchPromotion

'ErrDescription:
'if Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListPromotion As ListItem
Dim vSearchPromotion As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Text101.Text <> "" Then
        ListView101.ListItems.Clear
        vSearchPromotion = Trim(Text101.Text)
        vQuery = "execute USP_PM_Find '" & vSearchPromotion & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                Set vListPromotion = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("pmcode").Value))
                vListPromotion.SubItems(1) = Trim(vRecordset.Fields("pmname").Value)
                vListPromotion.SubItems(2) = Trim(vRecordset.Fields("datestart").Value)
                vListPromotion.SubItems(3) = Trim(vRecordset.Fields("dateend").Value)
                vListPromotion.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
    Else
        MsgBox "กรุณาใส่ข้อความค้นหา โปรโมชั่นด้วยนะครับ"
    End If
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
