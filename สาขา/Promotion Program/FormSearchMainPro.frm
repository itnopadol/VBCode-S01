VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSearchSecMan 
   Caption         =   "ค้นหา Section Manager"
   ClientHeight    =   3825
   ClientLeft      =   5055
   ClientTop       =   2115
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   8400
   Begin MSComctlLib.ListView ListView101 
      Height          =   2940
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   5186
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
         Text            =   "รหัส Section"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อ Section"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อ SectionName"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "FormSearchSecMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSecManList As ListItem

On Error GoTo ErrDescription

ListView101.ListItems.Clear
vQuery = "exec USP_PM_FindSecMan"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set vSecManList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("secmancode").Value))
    vSecManList.SubItems(1) = Trim(vRecordset.Fields("secmanname").Value)
    vSecManList.SubItems(2) = Trim(vRecordset.Fields("secname").Value)
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

    vIndex = ListView101.SelectedItem.Index
    If vMemCommand = 1 Then
    Form201.Text103.Text = ListView101.ListItems.Item(vIndex).Text
    ElseIf vMemCommand = 2 Then
    Form103.Text102.Text = ListView101.ListItems.Item(vIndex).Text
    Call SelectItemPromo
    Form103.Text104.Text = ""
    ElseIf vMemCommand = 3 Then
    Form401.Text101.Text = ListView101.ListItems.Item(vIndex).Text
    End If
    MDIForm1.Enabled = True
    Unload FormSearchSecMan
End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'If vMemCommand = 1 Then
    'Form201.Text103.Text = ListView101.ListItems.Item(Item.Index).SubItems(2)
    'ElseIf vMemCommand = 2 Then
    'Form103.Text102.Text = ListView101.ListItems.Item(Item.Index).SubItems(2)
    'Call SelectItemPromo
    'ElseIf vMemCommand = 3 Then
    'Form401.Text101.Text = ListView101.ListItems.Item(Item.Index).SubItems(2)
    'End If
    'MDIForm1.Enabled = True
    'Unload FormSearchSecMan
End Sub

