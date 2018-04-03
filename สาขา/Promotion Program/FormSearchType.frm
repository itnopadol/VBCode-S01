VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSearchType 
   Caption         =   "ค้นหาประเภทของสินค้าโปรโมชั่น"
   ClientHeight    =   4380
   ClientLeft      =   4875
   ClientTop       =   2490
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   8385
   Begin MSComctlLib.ListView ListView101 
      Height          =   3465
      Left            =   150
      TabIndex        =   0
      Top             =   225
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   6112
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสประเภท"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อประเภทโปรโมชั่น"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "คำอธิบาย"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "FormSearchType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vTypeList As ListItem

On Error GoTo ErrDescription

ListView101.ListItems.Clear
vQuery = "exec USP_PM_FindType"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set vTypeList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("code").Value))
    vTypeList.SubItems(1) = Trim(vRecordset.Fields("name1").Value)
    vTypeList.SubItems(2) = Trim(vRecordset.Fields("mydescription").Value)
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

    On Error Resume Next
    
    vIndex = ListView101.SelectedItem.Index
    Form201.ItemDetail107.Text = ListView101.ListItems.Item(vIndex).SubItems(1)
    MDIForm1.Enabled = True
    Form201.ItemDetail106.SetFocus
    Unload FormSearchType
End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'On Error Resume Next
    
    'Form201.ItemDetail107.Text = ListView101.ListItems.Item(Item.Index).SubItems(1)
    'MDIForm1.Enabled = True
    'Form201.ItemDetail106.SetFocus
    'Unload FormSearchType
End Sub
