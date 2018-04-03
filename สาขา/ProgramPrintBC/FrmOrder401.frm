VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrder401 
   Caption         =   "อนุมัติ ใบจัดคิวสินค้า"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmOrder401.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
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
      Left            =   9135
      TabIndex        =   5
      Top             =   5940
      Width           =   1005
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "อนุมัติ"
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
      Left            =   7695
      TabIndex        =   4
      Top             =   5940
      Width           =   1005
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3930
      Left            =   585
      TabIndex        =   3
      Top             =   1800
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   6932
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบจัดคิว"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่ใบจัดคิว"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่นัดส่ง"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "เวลานัดส่ง"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ค่าขนส่ง"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ผู้ทำรายการ"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton CMD101 
      Height          =   375
      Left            =   4140
      Picture         =   "FrmOrder401.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1035
      Width           =   375
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1755
      TabIndex        =   1
      Top             =   1035
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาใบจัดคิว"
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
      Left            =   585
      TabIndex        =   0
      Top             =   1080
      Width           =   1230
   End
End
Attribute VB_Name = "FrmOrder401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocno As String
Dim vSearchQueueList As ListItem
Dim i As Integer
    
On Error GoTo ErrDescription

    ListView101.ListItems.Clear
    vDocno = Trim(Text101.Text)
    vQuery = "exec bcnp.dbo.USP_DO_SearchQueueConfirm '" & vDocno & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchQueueList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        vSearchQueueList.SubItems(1) = Trim(vRecordset.Fields("docdate").Value)
        vSearchQueueList.SubItems(2) = Trim(vRecordset.Fields("duedate").Value)
        vSearchQueueList.SubItems(3) = Trim(vRecordset.Fields("duetime").Value)
        vSearchQueueList.SubItems(4) = Trim(vRecordset.Fields("hoardamount").Value)
        vSearchQueueList.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
        i = i + 1
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocno As String
Dim i As Integer
Dim vCheckCountOld As Integer
Dim vNewCount As Integer
Dim j As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    For i = 1 To ListView101.ListItems.Count
        If ListView101.ListItems.Item(i).Checked = True Then
            vDocno = Trim(ListView101.ListItems.Item(i).Text)
            vQuery = "exec bcnp.dbo.USP_DO_ConfirmQueue '" & vDocno & "' "
            gConnection.Execute vQuery
        End If
    Next i
End If
Call CMD101_Click

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Unload FrmOrder401
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocno As String
Dim vSearchQueueList As ListItem
Dim i As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    ListView101.ListItems.Clear
    vDocno = Trim(Text101.Text)
    vQuery = "exec bcnp.dbo.USP_DO_SearchQueueConfirm '" & vDocno & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchQueueList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        vSearchQueueList.SubItems(1) = Trim(vRecordset.Fields("docdate").Value)
        vSearchQueueList.SubItems(2) = Trim(vRecordset.Fields("duedate").Value)
        vSearchQueueList.SubItems(3) = Trim(vRecordset.Fields("duetime").Value)
        vSearchQueueList.SubItems(4) = Trim(vRecordset.Fields("hoardamount").Value)
        vSearchQueueList.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
        i = i + 1
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
       
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

