VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form002 
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDDetails 
      Caption         =   "ดูรายละเอียด"
      Height          =   555
      Left            =   765
      TabIndex        =   7
      Top             =   1170
      Width           =   1095
   End
   Begin VB.CommandButton CMDExit 
      Caption         =   "ออก"
      Height          =   555
      Left            =   10125
      TabIndex        =   6
      Top             =   7650
      Width           =   1050
   End
   Begin VB.CommandButton CMDSave 
      Caption         =   "บันทึก"
      Height          =   555
      Left            =   9000
      TabIndex        =   5
      Top             =   7650
      Width           =   1005
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   5235
      Left            =   810
      TabIndex        =   4
      Top             =   1890
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   9234
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสหน้าใช้งาน"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสฟอร์ม"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อฟอร์ม"
         Object.Width           =   10583
      EndProperty
   End
   Begin VB.ComboBox CMBLevelID 
      Height          =   315
      Left            =   8370
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   495
      Width           =   2805
   End
   Begin VB.ComboBox CMBDepartment 
      Height          =   315
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   495
      Width           =   5235
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ระดับผู้ใช้งาน :"
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
      Left            =   7020
      TabIndex        =   2
      Top             =   495
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "แผนก :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   855
      TabIndex        =   0
      Top             =   495
      Width           =   1050
   End
End
Attribute VB_Name = "Form002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDDetails_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vLIstPageID As ListItem
Dim vDepartmentCode As String
Dim vLevelID As Integer

If CMBDepartment.Text <> "" And CMBLevelID.Text <> "" Then
  vDepartmentCode = Left(CMBDepartment.Text, InStr(CMBDepartment.Text, "//") - 1)
  vLevelID = Left(CMBLevelID.Text, InStr(CMBLevelID.Text, "//") - 1)
  
  ListView101.ListItems.Clear
  vQuery = "exec dbo.USP_NP_CheckAuthorizePromotion '" & vDepartmentCode & "'," & vLevelID & " "
  If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
  vRecordset.MoveFirst
  While Not vRecordset.EOF
  Set vLIstPageID = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("pageid").Value))
  vLIstPageID.SubItems(1) = Trim(vRecordset.Fields("pagename").Value)
  vLIstPageID.SubItems(2) = Trim(vRecordset.Fields("pagedescription").Value)
  If Trim(vRecordset.Fields("pagestatus").Value) = 1 Then
     vLIstPageID.Checked = True
  Else
  vLIstPageID.Checked = False
  End If
  vRecordset.MoveNext
  Wend
  End If
  vRecordset.Close
  End If
End Sub

Private Sub CMDSave_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDepartmentCode As String
Dim vLevelID As Integer
Dim vPrgID As String
Dim vPageID As String
Dim vPageStatus As Integer
Dim i As Integer

If ListView101.ListItems.Count > 0 Then
For i = 1 To ListView101.ListItems.Count
  vDepartmentCode = Left(CMBDepartment.Text, InStr(CMBDepartment.Text, "//") - 1)
  vLevelID = Left(CMBLevelID.Text, InStr(CMBLevelID.Text, "//") - 1)
  vPrgID = "02"
  vPageID = ListView101.ListItems.Item(i).Text
  If ListView101.ListItems(i).Checked = True Then
  vPageStatus = 1
  Else
  vPageStatus = 0
  End If

vQuery = "exec dbo.USP_NP_InsertAuthorizePromotion '" & vDepartmentCode & "'," & vLevelID & ",'" & vPrgID & "','" & vPageID & "'," & vPageStatus & " "
gConnection.Execute vQuery
Next i
ListView101.ListItems.Clear

End If
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

vQuery = "exec dbo.USP_NP_BPlusDepartment"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst
While Not vRecordset.EOF
CMBDepartment.AddItem Trim(vRecordset.Fields("department").Value) & "//" & Trim(vRecordset.Fields("dept_thaidesc").Value)
vRecordset.MoveNext
Wend
End If
vRecordset.Close

CMBLevelID.AddItem Trim("1//พนักงานทั่วไป")
CMBLevelID.AddItem Trim("5//Section Manager")
CMBLevelID.AddItem Trim("9//ผู้บริหารและหัวหน้า")
CMBLevelID.Text = Trim("1//พนักงานทั่วไป")

End Sub
