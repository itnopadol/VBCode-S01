VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSearchItem 
   Caption         =   "ค้นหาสินค้า"
   ClientHeight    =   5430
   ClientLeft      =   4680
   ClientTop       =   1935
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8820
   Begin MSComctlLib.ListView ListView101 
      Height          =   3540
      Left            =   525
      TabIndex        =   2
      Top             =   675
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   6244
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
         Text            =   "รหัสสินค้า"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อสินค้า"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ราคาปกติ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "หน่วยนับ"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1500
      TabIndex        =   1
      Top             =   225
      Width           =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "ค้นหาสินค้า"
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
      Left            =   525
      TabIndex        =   0
      Top             =   300
      Width           =   915
   End
End
Attribute VB_Name = "FormSearchItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSortResult As Integer

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Enabled = True
End Sub

Private Sub ListView101_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo ErrDescription

ListView101.Sorted = True
ListView101.SortKey = ColumnHeader.Index - 1
If vSortResult = 0 Then
    ListView101.SortOrder = lvwAscending
    vSortResult = 1
Else
    ListView101.SortOrder = lvwDescending
    vSortResult = 0
End If



ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListView101_DblClick()
Dim vIndex As Integer

On Error Resume Next

    vIndex = ListView101.SelectedItem.Index
    MDIForm1.Enabled = True
    Form201.Show
    Form201.ItemDetail101 = Trim(ListView101.ListItems.Item(vIndex).Text)
    Form201.ItemDetail102 = Trim(ListView101.ListItems.Item(vIndex).SubItems(1))
    Form201.ItemDetail104 = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
    Form201.ItemDetail103 = Trim(ListView101.ListItems.Item(vIndex).SubItems(2))
    Unload FormSearchItem
    Form201.ItemDetail105.SetFocus
    Form201.CHK103.Value = 0
    Form201.Check101.Value = 0
    Form201.ItemDetail106.Text = ""
    Form201.ItemDetail107.Text = ""

End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer

On Error Resume Next

If KeyAscii = 13 Then
    vIndex = ListView101.SelectedItem.Index
    MDIForm1.Enabled = True
    Form201.Show
    Form201.ItemDetail101 = Trim(ListView101.ListItems.Item(vIndex).Text)
    Form201.ItemDetail102 = Trim(ListView101.ListItems.Item(vIndex).SubItems(1))
    Form201.ItemDetail104 = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
    Form201.ItemDetail103 = Trim(ListView101.ListItems.Item(vIndex).SubItems(2))
    Unload FormSearchItem
    Form201.ItemDetail105.SetFocus
    Form201.CHK103.Value = 0
    Form201.Check101.Value = 0
    Form201.ItemDetail106.Text = ""
    Form201.ItemDetail107.Text = ""
End If
End Sub

Private Sub Text101_GotFocus()
'Dim vRecordset As New ADODB.Recordset
'Dim vQuery As String
'Dim vSearch As String
'Dim vListItem As ListItem

 '  If Text101.Text <> "" Then
  '      vSearch = Trim(Text101.Text)
   '     vQuery = "execute USP_PM_FindItem '" & vSearch & "' "
    '    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
     '       vRecordset.MoveFirst
      '      While Not vRecordset.EOF
       '     Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
        '    vListItem.SubItems(1) = Trim(vRecordset.Fields("itemname").Value)
         '   vListItem.SubItems(2) = Trim(vRecordset.Fields("saleprice1").Value)
          ''  vListItem.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
            'vRecordset.MoveNext
            'Wend
        'End If
        'vRecordset.Close
    'End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vListItem As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Text101.Text <> "" Then
        vSearch = Trim(Text101.Text)
        vQuery = "execute USP_PM_FindItem '" & vSearch & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
            vListItem.SubItems(1) = Trim(vRecordset.Fields("itemname").Value)
            vListItem.SubItems(2) = Trim(vRecordset.Fields("saleprice1").Value)
            vListItem.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
    End If
    ListView101.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

