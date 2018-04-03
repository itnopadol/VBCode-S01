VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSearchType 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
   Begin MSComctlLib.ListView ListView101 
      Height          =   3240
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   5715
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
         Text            =   "รหัสประเภทโปรโมชั่น"
         Object.Width           =   39688
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อภาษาอังกฤษ"
         Object.Width           =   52917
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อภาษาไทย"
         Object.Width           =   52917
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "คำอธิบาย"
         Object.Width           =   79375
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
    vTypeList.SubItems(1) = Trim(vRecordset.Fields("nameeng").Value)
    vTypeList.SubItems(2) = Trim(vRecordset.Fields("namethai").Value)
    vTypeList.SubItems(3) = Trim(vRecordset.Fields("mydescription").Value)
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
Form201.Enabled = True
End Sub
