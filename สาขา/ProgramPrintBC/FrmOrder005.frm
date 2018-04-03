VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrder005 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form005 ระดับความสำคัญ"
   ClientHeight    =   5100
   ClientLeft      =   3795
   ClientTop       =   2205
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmOrder005.frx":0000
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
      Height          =   420
      Left            =   7425
      TabIndex        =   4
      Top             =   4365
      Width           =   870
   End
   Begin VB.CommandButton CMD101 
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
      Left            =   6300
      TabIndex        =   3
      Top             =   4365
      Width           =   870
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2805
      Left            =   90
      TabIndex        =   2
      Top             =   1440
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4948
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
         Text            =   "ลำดับ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID (รหัส)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ระดับความสำคัญ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "คำอธิบาย"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton CMD102 
      Height          =   330
      Left            =   3870
      Picture         =   "FrmOrder005.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "ค้นหา ตามเงื่อนไขที่กรอก"
      Top             =   990
      Width           =   330
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1890
      TabIndex        =   0
      Top             =   990
      Width           =   1950
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหา ระดับความสำคัญ"
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
      Left            =   90
      TabIndex        =   5
      Top             =   1035
      Width           =   1995
   End
End
Attribute VB_Name = "FrmOrder005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As Integer
Dim vIsCancel As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    If vPriorityModule = 1 Then
        Form312.CMB102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
    ElseIf vPriorityModule = 2 Then
        vID = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        vQuery = "select id,iscancel from npmaster.dbo.TB_DO_Priority where id = " & vID & " "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vIsCancel = Trim(vRecordset.Fields("iscancel").Value)
        End If
        vRecordset.Close
        If vIsCancel = 0 Then
            FrmOrder201.Image101.Visible = True
            FrmOrder201.Image102.Visible = False
        Else
            FrmOrder201.Image101.Visible = False
            FrmOrder201.Image102.Visible = True
        End If
        FrmOrder201.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        FrmOrder201.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        FrmOrder201.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
    End If
    Unload FrmOrder005
    vCheckPriorityOpen = 1
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem
Dim i As Integer
    
On Error Resume Next

    ListView101.ListItems.Clear
    vSearch = Trim(Text101.Text)
    vQuery = "exec bcnp.dbo.USP_DO_PrioritySearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(i))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("id").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("priority").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("mydescription").Value)
        vRecordset.MoveNext
        i = i + 1
        Wend
        Else
            MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
        End If
    vRecordset.Close

End Sub

Private Sub CMD103_Click()
Unload FrmOrder005
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmProgramPrint.Enabled = True
End Sub

Private Sub ListView101_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As Integer
Dim vIsCancel As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    If vPriorityModule = 1 Then
        Form312.CMB102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
    ElseIf vPriorityModule = 2 Then
        vID = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        vQuery = "select id,iscancel from npmaster.dbo.TB_DO_Priority where id = " & vID & " "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vIsCancel = Trim(vRecordset.Fields("iscancel").Value)
        End If
        vRecordset.Close
        If vIsCancel = 0 Then
            FrmOrder201.Image101.Visible = True
            FrmOrder201.Image102.Visible = False
        Else
            FrmOrder201.Image101.Visible = False
            FrmOrder201.Image102.Visible = True
        End If
        FrmOrder201.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        FrmOrder201.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        FrmOrder201.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
    End If
    Unload FrmOrder005
    vCheckPriorityOpen = 1
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As Integer
Dim vIsCancel As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If ListView101.ListItems.Count <> 0 Then
        If vPriorityModule = 1 Then
            Form312.CMB102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
        ElseIf vPriorityModule = 2 Then
            vID = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
            vQuery = "select id,iscancel from npmaster.dbo.TB_DO_Priority where id = " & vID & " "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vIsCancel = Trim(vRecordset.Fields("iscancel").Value)
            End If
            vRecordset.Close
            If vIsCancel = 0 Then
                FrmOrder201.Image101.Visible = True
                FrmOrder201.Image102.Visible = False
            Else
                FrmOrder201.Image101.Visible = False
                FrmOrder201.Image102.Visible = True
            End If
            FrmOrder201.Text101.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
            FrmOrder201.Text102.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(2))
            FrmOrder201.Text103.Text = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        End If
        Unload FrmOrder005
        vCheckPriorityOpen = 1
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
Dim i As Integer

On Error Resume Next

If KeyAscii = 13 Then
    ListView101.ListItems.Clear
    vSearch = Trim(Text101.Text)
    vQuery = "exec bcnp.dbo.USP_DO_PrioritySearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(i))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("id").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("priority").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("mydescription").Value)
        vRecordset.MoveNext
        i = i + 1
        Wend
    Else
            MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
    End If
    vRecordset.Close
End If
End Sub

