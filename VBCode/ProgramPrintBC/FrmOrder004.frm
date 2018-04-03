VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrder004 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form004 ค้นหาเอกสารทำใบจัดคิว"
   ClientHeight    =   5295
   ClientLeft      =   3795
   ClientTop       =   2220
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmOrder004.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   8445
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
      Left            =   7290
      TabIndex        =   4
      Top             =   4410
      Width           =   915
   End
   Begin VB.CommandButton CMD102 
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
      Left            =   6165
      TabIndex        =   3
      Top             =   4410
      Width           =   915
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   3420
      Picture         =   "FrmOrder004.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "ค้นหา ตามเงื่อนไขที่กรอก"
      Top             =   945
      Width           =   330
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1305
      TabIndex        =   0
      Top             =   945
      Width           =   2040
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2895
      Left            =   270
      TabIndex        =   2
      Top             =   1350
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5106
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่เอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "รหัสลูกค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ค่าขนส่ง"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาเอกสาร"
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
      Left            =   270
      TabIndex        =   5
      Top             =   990
      Width           =   1185
   End
End
Attribute VB_Name = "FrmOrder004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vTextSearch As String
Dim vSearchList As ListItem
Dim i As Integer
Dim vIsPOS As Integer

On Error Resume Next

If Text101.Text <> "" Then
    ListView101.ListItems.Clear
    vTextSearch = Trim(Text101.Text)
    i = 1
    If Form312.Check101.Value = 1 Then
        vIsPOS = 1
    Else
        If Form312.CMB101.Text = Trim("ใบสั่งขาย/จอง") Then
            vIsPOS = 0
        ElseIf Form312.CMB101.Text = Trim("บิลขาย") Then
            vIsPOS = 1
        End If
    End If
    vQuery = "exec bcnp.dbo.usp_do_searchrefheader " & vIsPOS & ",'" & vTextSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(i))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("arcode").Value)
        vSearchList.SubItems(4) = Trim(vRecordset.Fields("arname").Value)
        vSearchList.SubItems(5) = Trim(vRecordset.Fields("HoardAmount").Value)
        vRecordset.MoveNext
        i = i + 1
        Wend
    Else
            MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
    End If
    vRecordset.Close
End If

End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vTextSearch As String
Dim vIsPOS As Integer
Dim vSearchList As ListItem
Dim i As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
i = 1
    vTextSearch = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
    If Form312.Check101.Value = 1 Then
        vIsPOS = 1
    Else
        vIsPOS = 0
    End If
    
    Form312.ListView101.ListItems.Clear
    Form312.Text103.Text = ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5)
    vQuery = "exec bcnp.dbo.USP_DO_SearchRef " & vIsPOS & ",'" & vTextSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set vSearchList = Form312.ListView101.ListItems.Add(, , Trim(i))
            vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
            vSearchList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
            vSearchList.SubItems(3) = Trim(vRecordset.Fields("itemname").Value)
            vSearchList.SubItems(4) = Trim(vRecordset.Fields("qty").Value)
            vSearchList.SubItems(5) = Trim(vRecordset.Fields("doremainqty").Value)
            vSearchList.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
            vSearchList.SubItems(7) = Trim(vRecordset.Fields("headid").Value)
            vSearchList.SubItems(8) = Trim(vRecordset.Fields("detailid").Value)
            vSearchList.SubItems(9) = Trim(0)
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
    Unload FrmOrder004
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Unload FrmOrder004
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmProgramPrint.Enabled = True
End Sub

Private Sub ListView101_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vTextSearch As String
Dim vIsPOS As Integer
Dim vSearchList As ListItem
Dim i As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
i = 1
    vTextSearch = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
    If Form312.Check101.Value = 1 Then
        vIsPOS = 1
    Else
        If Form312.CMB101.Text = Trim("ใบสั่งขาย/จอง") Then
            vIsPOS = 0
        ElseIf Form312.CMB101.Text = Trim("บิลขาย") Then
            vIsPOS = 1
        End If
    End If
    Form312.ListView101.ListItems.Clear
    Form312.Text103.Text = ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5)
    vQuery = "exec bcnp.dbo.USP_DO_SearchRef " & vIsPOS & ",'" & vTextSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set vSearchList = Form312.ListView101.ListItems.Add(, , Trim(i))
            vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
            vSearchList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
            vSearchList.SubItems(3) = Trim(vRecordset.Fields("itemname").Value)
            vSearchList.SubItems(4) = Trim(vRecordset.Fields("qty").Value)
            vSearchList.SubItems(5) = Trim(vRecordset.Fields("doremainqty").Value)
            vSearchList.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
            vSearchList.SubItems(7) = Trim(vRecordset.Fields("headid").Value)
            vSearchList.SubItems(8) = Trim(vRecordset.Fields("detailid").Value)
            vSearchList.SubItems(9) = Trim(0)
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
    Unload FrmOrder004
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
Dim vTextSearch As String
Dim vIsPOS As Integer
Dim vSearchList As ListItem
Dim i As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If ListView101.ListItems.Count <> 0 Then
        i = 1
    vTextSearch = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
    If Form312.Check101.Value = 1 Then
        vIsPOS = 1
    Else
        If Form312.CMB101.Text = Trim("ใบสั่งขาย/จอง") Then
            vIsPOS = 0
        ElseIf Form312.CMB101.Text = Trim("บิลขาย") Then
            vIsPOS = 1
        End If
    End If
    
    Form312.ListView101.ListItems.Clear
    Form312.Text103.Text = ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5)
    vQuery = "exec bcnp.dbo.USP_DO_SearchRef " & vIsPOS & ",'" & vTextSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set vSearchList = Form312.ListView101.ListItems.Add(, , Trim(i))
            vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
            vSearchList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
            vSearchList.SubItems(3) = Trim(vRecordset.Fields("itemname").Value)
            vSearchList.SubItems(4) = Trim(vRecordset.Fields("qty").Value)
            vSearchList.SubItems(5) = Trim(vRecordset.Fields("doremainqty").Value)
            vSearchList.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
            vSearchList.SubItems(7) = Trim(vRecordset.Fields("headid").Value)
            vSearchList.SubItems(8) = Trim(vRecordset.Fields("detailid").Value)
            vSearchList.SubItems(9) = Trim(0)
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
    Unload FrmOrder004
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
Dim vTextSearch As String
Dim vSearchList As ListItem
Dim i As Integer
Dim vIsPOS As Integer

On Error Resume Next

If KeyAscii = 13 Then
    If Text101.Text <> "" Then
        ListView101.ListItems.Clear
        vTextSearch = Trim(Text101.Text)
        i = 1
        If Form312.Check101.Value = 1 Then
            vIsPOS = 1
        Else
            If Form312.CMB101.Text = Trim("ใบสั่งขาย/จอง") Then
                vIsPOS = 0
            ElseIf Form312.CMB101.Text = Trim("บิลขาย") Then
                vIsPOS = 1
            End If
        End If
        vQuery = "exec bcnp.dbo.usp_do_searchrefheader " & vIsPOS & ",'" & vTextSearch & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set vSearchList = ListView101.ListItems.Add(, , Trim(i))
            vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
            vSearchList.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
            vSearchList.SubItems(3) = Trim(vRecordset.Fields("arcode").Value)
            vSearchList.SubItems(4) = Trim(vRecordset.Fields("arname").Value)
            vSearchList.SubItems(5) = Trim(vRecordset.Fields("HoardAmount").Value)
            vRecordset.MoveNext
            i = i + 1
            Wend
        Else
            MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
        End If
        vRecordset.Close
    End If
End If
End Sub

