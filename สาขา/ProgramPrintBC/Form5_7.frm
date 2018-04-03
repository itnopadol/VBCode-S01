VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5_7 
   Caption         =   "เพิ่มทะเบียนกลุ่มวงเงินลูกหนี้"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form5_7.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDClearScreen 
      Caption         =   "ล้างหน้าจอ"
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
      Left            =   9810
      TabIndex        =   11
      Top             =   3600
      Width           =   1140
   End
   Begin MSComctlLib.ListView ListViewARGroup 
      Height          =   2670
      Left            =   990
      TabIndex        =   4
      Top             =   4455
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   4710
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสกลุ่มวงเงิน"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อกลุ่มวงเงิน"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "หมายเหตุ"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ผู้บันทึก"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton CMDClose 
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
      Left            =   2295
      TabIndex        =   6
      Top             =   7245
      Width           =   1140
   End
   Begin VB.CommandButton CMDDelete 
      Caption         =   "ลบ"
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
      Left            =   990
      TabIndex        =   5
      ToolTipText     =   "ต้องการลบรายการใดให้ Click เลือกรายการนั้นก่อนกดปุ่ม ลบ"
      Top             =   7245
      Width           =   1140
   End
   Begin VB.CommandButton CMDSave 
      Caption         =   "บันทึก"
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
      Left            =   8415
      TabIndex        =   3
      Top             =   3600
      Width           =   1140
   End
   Begin VB.TextBox TextMydescription 
      Appearance      =   0  'Flat
      Height          =   1050
      Left            =   3555
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2385
      Width           =   7395
   End
   Begin VB.TextBox TextName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3555
      TabIndex        =   1
      Top             =   1755
      Width           =   7395
   End
   Begin VB.TextBox TextCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3555
      TabIndex        =   0
      Top             =   1215
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการ กลุ่มวงเงินลูกหนี้"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   990
      TabIndex        =   10
      Top             =   4095
      Width           =   1950
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ :"
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
      Left            =   1485
      TabIndex        =   9
      Top             =   2340
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อกลุ่มวงเงินลูกหนี้ :"
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
      Left            =   1530
      TabIndex        =   8
      Top             =   1755
      Width           =   1950
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสกลุ่มวงเงินลูกหนี้ :"
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
      Left            =   1530
      TabIndex        =   7
      Top             =   1215
      Width           =   1950
   End
End
Attribute VB_Name = "Form5_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDClearScreen_Click()
On Error GoTo ErrDescription

   Me.TextCode.Text = ""
   Me.TextName.Text = ""
   Me.TextMydescription.Text = ""
   Me.TextCode.SetFocus
   
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDDelete_Click()
Dim vQuery As String
Dim vCode As String
Dim vIndex As Integer
Dim vAnswer As Integer

On Error GoTo ErrDescription

If Me.ListViewARGroup.ListItems.Count > 0 Then
   vIndex = Me.ListViewARGroup.SelectedItem.Index
   vCode = Me.ListViewARGroup.ListItems(vIndex).Text
   
   vAnswer = MsgBox("คุณต้องลบ กลุ่มวงเงินลูกหนี้ รหัส " & vCode & " นี้ใช่หรือไม่", vbYesNo, "Send Question Message ?")
   If vAnswer = 6 Then
      vQuery = "exec dbo.USP_BC_DeleteBCARCreditGroup '" & vCode & "' "
      gConnection.Execute (vQuery)
      Call GetARCreditGroup
      Me.TextCode.SetFocus
   End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDSave_Click()
Dim vQuery As String
Dim vCode As String
Dim vName As String
Dim vMydescription As String

On Error GoTo ErrDescription

If Me.TextCode.Text <> "" And Me.TextName.Text <> "" Then
   vCode = Me.TextCode.Text
   vName = Me.TextName.Text
   vMydescription = Me.TextMydescription.Text
   
   vQuery = "exec dbo.USP_BC_InsertBCARCreditGroup '" & vCode & "','" & vName & "','" & vMydescription & "','" & vUserID & "' "
   gConnection.Execute (vQuery)
   
   Me.TextCode.Text = ""
   Me.TextName.Text = ""
   Me.TextMydescription.Text = ""
   Me.TextCode.SetFocus
   
   Call GetARCreditGroup
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Call GetARCreditGroup
End Sub

Public Sub GetARCreditGroup()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem

On Error Resume Next

vQuery = "exec dbo.USP_BC_BCARCreditGroupData '' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   Me.ListViewARGroup.ListItems.Clear
   vRecordset.MoveFirst
   While Not vRecordset.EOF
   Set vListItem = Me.ListViewARGroup.ListItems.Add(, , vRecordset.Fields("code").Value)
   vListItem.SubItems(1) = vRecordset.Fields("name").Value
   vListItem.SubItems(2) = vRecordset.Fields("mydescription").Value
   vListItem.SubItems(3) = vRecordset.Fields("creatorcode").Value
   vRecordset.MoveNext
   Wend
End If
vRecordset.Close

End Sub

Private Sub ListViewARGroup_DblClick()
Dim vIndex  As Integer

On Error GoTo ErrDescription

If Me.ListViewARGroup.ListItems.Count > 0 Then
   vIndex = Me.ListViewARGroup.SelectedItem.Index
   Me.TextCode.Text = Me.ListViewARGroup.ListItems(vIndex).Text
   Me.TextName.Text = Me.ListViewARGroup.ListItems(vIndex).SubItems(1)
   Me.TextMydescription.Text = Me.ListViewARGroup.ListItems(vIndex).SubItems(2)
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListViewARGroup_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vQuery As String
Dim vCode As String
Dim vIndex As Integer
Dim vAnswer As Integer

On Error GoTo ErrDescription

If KeyCode = 46 Then
   If Me.ListViewARGroup.ListItems.Count > 0 Then
      vIndex = Me.ListViewARGroup.SelectedItem.Index
      vCode = Me.ListViewARGroup.ListItems(vIndex).Text
      
      vAnswer = MsgBox("คุณต้องลบ กลุ่มวงเงินลูกหนี้ รหัส " & vCode & " นี้ใช่หรือไม่", vbYesNo, "Send Question Message ?")
      If vAnswer = 6 Then
         vQuery = "exec dbo.USP_BC_DeleteBCARCreditGroup '" & vCode & "' "
         gConnection.Execute (vQuery)
         Call GetARCreditGroup
         Me.TextCode.SetFocus
      End If
   End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListViewARGroup_KeyPress(KeyAscii As Integer)
Dim vIndex  As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
   If Me.ListViewARGroup.ListItems.Count > 0 Then
      vIndex = Me.ListViewARGroup.SelectedItem.Index
      Me.TextCode.Text = Me.ListViewARGroup.ListItems(vIndex).Text
      Me.TextName.Text = Me.ListViewARGroup.ListItems(vIndex).SubItems(1)
      Me.TextMydescription.Text = Me.ListViewARGroup.ListItems(vIndex).SubItems(2)
   End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub TextCode_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCode As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
   Me.TextName.Text = ""
   Me.TextMydescription.Text = ""
   If Me.TextCode.Text <> "" Then
      vCode = Me.TextCode.Text
      vQuery = "exec dbo.USP_BC_BCARCreditGroupData '" & vCode & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         Me.TextName.Text = vRecordset.Fields("name")
         Me.TextMydescription.Text = vRecordset.Fields("mydescription")
      End If
      vRecordset.Close
   End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub TextName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.TextMydescription.SetFocus
End If
End Sub
