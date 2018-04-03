VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form104 
   Caption         =   "ลบ User ค้าง"
   ClientHeight    =   8700
   ClientLeft      =   6615
   ClientTop       =   1680
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   14385
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Left            =   3960
      TabIndex        =   0
      Top             =   765
      Width           =   4470
      Begin VB.CommandButton CMD101 
         Caption         =   "ฟื้นฟูข้อมูล"
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
         Left            =   2970
         TabIndex        =   2
         Top             =   4095
         Width           =   1230
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   3210
         Left            =   270
         TabIndex        =   1
         Top             =   630
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   5662
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LogIn Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "LogIn Date Time"
            Object.Width           =   4145
         EndProperty
      End
   End
End
Attribute VB_Name = "Form104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vUserList As ListItem

ListView101.ListItems.Clear
vQuery = "select * from npmaster.dbo.TB_CK_UserActivateProgram where jobid = '02' order by userid "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set vUserList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("userid").Value))
    vUserList.SubItems(1) = Trim(vRecordset.Fields("activedatetime").Value)
    vRecordset.MoveNext
    Wend
    End If
vRecordset.Close
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vUserList As ListItem

ListView101.ListItems.Clear
vQuery = "select * from npmaster.dbo.TB_CK_UserActivateProgram where jobid = '02' order by userid "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set vUserList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("userid").Value))
    vUserList.SubItems(1) = Trim(vRecordset.Fields("activedatetime").Value)
    vRecordset.MoveNext
    Wend
    End If
vRecordset.Close
End Sub

Private Sub ListView101_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vUserID As String

If ListView101.ListItems.Count <> 0 Then
    If KeyCode = 46 Then
        i = ListView101.SelectedItem.Index
        vUserID = Trim(ListView101.ListItems.Item(i))
        ListView101.ListItems.Remove (i)
        vQuery = "delete npmaster.dbo.TB_CK_UserActivateProgram where userid = '" & vUserID & "' "
        gConnection.Execute vQuery
    End If
End If
End Sub
