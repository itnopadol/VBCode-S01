VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrder006 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form006 ค้นหาเอกสารจัดคิว"
   ClientHeight    =   7155
   ClientLeft      =   2235
   ClientTop       =   975
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmOrder006.frx":0000
   ScaleHeight     =   7155
   ScaleWidth      =   10920
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
      Left            =   9630
      TabIndex        =   4
      Top             =   5940
      Width           =   960
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
      Left            =   8415
      TabIndex        =   3
      Top             =   5940
      Width           =   960
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   4185
      Picture         =   "FrmOrder006.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "ค้นหา ตามเงื่อนไขที่กรอก"
      Top             =   990
      Width           =   330
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1800
      TabIndex        =   0
      Top             =   990
      Width           =   2355
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   4155
      Left            =   270
      TabIndex        =   2
      Top             =   1485
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   7329
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่เอกสาร"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันที่ครบกำหนด"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ระดับความสำคัญ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "เบอร์บ้าน"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "เบอร์มือถือ"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาเอกสารเลขที่"
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
      Height          =   375
      Left            =   270
      TabIndex        =   5
      Top             =   1035
      Width           =   1545
   End
End
Attribute VB_Name = "FrmOrder006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem
Dim i As Integer

On Error Resume Next

    ListView101.ListItems.Clear
    vSearch = Trim(Text101.Text)
    vQuery = "exec bcnp.dbo.USP_DO_QueueHeaderSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(i))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("duedate").Value)
        vSearchList.SubItems(4) = Trim(vRecordset.Fields("priority").Value)
        vSearchList.SubItems(5) = Trim(vRecordset.Fields("receivename").Value)
        vSearchList.SubItems(6) = Trim(vRecordset.Fields("receivetelhome").Value)
        vSearchList.SubItems(7) = Trim(vRecordset.Fields("receivetelmobile").Value)
        If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
            ListView101.ListItems(i).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
        End If
        vRecordset.MoveNext
        i = i + 1
        Wend
    Else
        MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
    End If
    vRecordset.Close
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem
Dim i As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    Form312.ListView101.ListItems.Clear
    vIsOpen1 = 1
    vSearch = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
    vQuery = "exec bcnp.dbo.USP_DO_QueueSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
            Form312.Image101.Visible = False
            Form312.Image102.Visible = True
            Form312.Image103.Visible = False
        Else
            If Trim(vRecordset.Fields("isconfirm").Value) = 0 Then
                Form312.Image101.Visible = True
                Form312.Image102.Visible = False
                Form312.Image103.Visible = False
            Else
                Form312.Image101.Visible = False
                Form312.Image102.Visible = False
                Form312.Image103.Visible = True
            End If
        End If
        Form312.DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
        Form312.DTPicker102.Value = Trim(vRecordset.Fields("duedate").Value)
        Form312.Text101.Text = Trim(vRecordset.Fields("id").Value)
        Form312.Text102.Text = Trim(vRecordset.Fields("docno").Value)
        Form312.Text103.Text = Trim(vRecordset.Fields("hoardamount").Value)
        Form312.CMB102.Text = Trim(vRecordset.Fields("priority").Value)
        Form312.Text105.Text = Trim(vRecordset.Fields("distance").Value)
        Form312.Text106.Text = Trim(vRecordset.Fields("transportlocation").Value)
        Form312.Text107.Text = Trim(vRecordset.Fields("mappart").Value)
        Form312.Text108.Text = Trim(vRecordset.Fields("mydescription").Value)
        Form312.MaskEdBox101.Text = Trim(vRecordset.Fields("duetime").Value)
        Form312.Text201.Text = Trim(vRecordset.Fields("receivename").Value)
        Form312.Text202.Text = Trim(vRecordset.Fields("receivetelhome").Value)
        Form312.Text203.Text = Trim(vRecordset.Fields("receivetelmobile").Value)
        Form312.Text301.Text = Trim(vRecordset.Fields("placeid").Value)
        Form312.Text302.Text = Trim(vRecordset.Fields("district").Value)
        Form312.Text303.Text = Trim(vRecordset.Fields("amphur").Value)
        Form312.Text304.Text = Trim(vRecordset.Fields("province").Value)
        Form312.Text305.Text = Trim(vRecordset.Fields("placemydescription").Value)
        Form312.Text306.Text = Trim(vRecordset.Fields("routeid").Value)
        Form312.Text307.Text = Trim(vRecordset.Fields("name1").Value)
        Form312.Text308.Text = Trim(vRecordset.Fields("name2").Value)
        Form312.Text309.Text = Trim(vRecordset.Fields("routemydescription").Value)
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = Form312.ListView101.ListItems.Add(, , Trim(i))
            vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
            vSearchList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
            vSearchList.SubItems(3) = Trim(vRecordset.Fields("itemname").Value)
            vSearchList.SubItems(4) = Trim(vRecordset.Fields("qty").Value)
            vSearchList.SubItems(5) = Trim(vRecordset.Fields("remainqty").Value)
            vSearchList.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
            vSearchList.SubItems(7) = Trim(vRecordset.Fields("refid").Value)
            vSearchList.SubItems(8) = Trim(vRecordset.Fields("refsubid").Value)
            vSearchList.SubItems(9) = Trim(vRecordset.Fields("iscancelsub").Value)
            If Trim(vRecordset.Fields("iscancelsub").Value) = 1 Then
                Form312.ListView101.ListItems(i).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H000000FF"
            End If
            
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
    Unload FrmOrder006
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Unload FrmOrder006
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Form001.Order101.Enabled = True '--------------------------------
Form312.Enabled = True
End Sub

Private Sub ListView101_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem
Dim i As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    Form312.ListView101.ListItems.Clear
    vIsOpen1 = 1
    vSearch = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
    vQuery = "exec bcnp.dbo.USP_DO_QueueSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
            Form312.Image101.Visible = False
            Form312.Image102.Visible = True
            Form312.Image103.Visible = False
        Else
            If Trim(vRecordset.Fields("isconfirm").Value) = 0 Then
                Form312.Image101.Visible = True
                Form312.Image102.Visible = False
                Form312.Image103.Visible = False
            Else
                Form312.Image101.Visible = False
                Form312.Image102.Visible = False
                Form312.Image103.Visible = True
            End If
        End If
        Form312.DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
        Form312.DTPicker102.Value = Trim(vRecordset.Fields("duedate").Value)
        Form312.Text101.Text = Trim(vRecordset.Fields("id").Value)
        Form312.Text102.Text = Trim(vRecordset.Fields("docno").Value)
        Form312.Text103.Text = Trim(vRecordset.Fields("hoardamount").Value)
        Form312.CMB102.Text = Trim(vRecordset.Fields("priority").Value)
        Form312.Text105.Text = Trim(vRecordset.Fields("distance").Value)
        Form312.Text106.Text = Trim(vRecordset.Fields("transportlocation").Value)
        Form312.Text107.Text = Trim(vRecordset.Fields("mappart").Value)
        Form312.Text108.Text = Trim(vRecordset.Fields("mydescription").Value)
        Form312.MaskEdBox101.Text = Trim(vRecordset.Fields("duetime").Value)
        Form312.Text201.Text = Trim(vRecordset.Fields("receivename").Value)
        Form312.Text202.Text = Trim(vRecordset.Fields("receivetelhome").Value)
        Form312.Text203.Text = Trim(vRecordset.Fields("receivetelmobile").Value)
        Form312.Text301.Text = Trim(vRecordset.Fields("placeid").Value)
        Form312.Text302.Text = Trim(vRecordset.Fields("district").Value)
        Form312.Text303.Text = Trim(vRecordset.Fields("amphur").Value)
        Form312.Text304.Text = Trim(vRecordset.Fields("province").Value)
        Form312.Text305.Text = Trim(vRecordset.Fields("placemydescription").Value)
        Form312.Text306.Text = Trim(vRecordset.Fields("routeid").Value)
        Form312.Text307.Text = Trim(vRecordset.Fields("name1").Value)
        Form312.Text308.Text = Trim(vRecordset.Fields("name2").Value)
        Form312.Text309.Text = Trim(vRecordset.Fields("routemydescription").Value)
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = Form312.ListView101.ListItems.Add(, , Trim(i))
            vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
            vSearchList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
            vSearchList.SubItems(3) = Trim(vRecordset.Fields("itemname").Value)
            vSearchList.SubItems(4) = Trim(vRecordset.Fields("qty").Value)
            vSearchList.SubItems(5) = Trim(vRecordset.Fields("remainqty").Value)
            vSearchList.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
            vSearchList.SubItems(7) = Trim(vRecordset.Fields("refid").Value)
            vSearchList.SubItems(8) = Trim(vRecordset.Fields("refsubid").Value)
            vSearchList.SubItems(9) = Trim(vRecordset.Fields("iscancelsub").Value)
            If Trim(vRecordset.Fields("iscancelsub").Value) = 1 Then
                Form312.ListView101.ListItems(i).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H000000FF"
            End If
            
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
    Form312.CMD006.Enabled = True
    Unload FrmOrder006
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
Dim vSearch As String
Dim vSearchList As ListItem
Dim i As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If ListView101.ListItems.Count <> 0 Then
    Form312.ListView101.ListItems.Clear
    vIsOpen1 = 1
    vSearch = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
    vQuery = "exec bcnp.dbo.USP_DO_QueueSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
            Form312.Image101.Visible = False
            Form312.Image102.Visible = True
            Form312.Image103.Visible = False
        Else
            If Trim(vRecordset.Fields("isconfirm").Value) = 0 Then
                Form312.Image101.Visible = True
                Form312.Image102.Visible = False
                Form312.Image103.Visible = False
            Else
                Form312.Image101.Visible = False
                Form312.Image102.Visible = False
                Form312.Image103.Visible = True
            End If
        End If
        Form312.DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
        Form312.DTPicker102.Value = Trim(vRecordset.Fields("duedate").Value)
        Form312.Text101.Text = Trim(vRecordset.Fields("id").Value)
        Form312.Text102.Text = Trim(vRecordset.Fields("docno").Value)
        Form312.Text103.Text = Trim(vRecordset.Fields("hoardamount").Value)
        Form312.CMB102.Text = Trim(vRecordset.Fields("priority").Value)
        Form312.Text105.Text = Trim(vRecordset.Fields("distance").Value)
        Form312.Text106.Text = Trim(vRecordset.Fields("transportlocation").Value)
        Form312.Text107.Text = Trim(vRecordset.Fields("mappart").Value)
        Form312.Text108.Text = Trim(vRecordset.Fields("mydescription").Value)
        Form312.MaskEdBox101.Text = Trim(vRecordset.Fields("duetime").Value)
        Form312.Text201.Text = Trim(vRecordset.Fields("receivename").Value)
        Form312.Text202.Text = Trim(vRecordset.Fields("receivetelhome").Value)
        Form312.Text203.Text = Trim(vRecordset.Fields("receivetelmobile").Value)
        Form312.Text301.Text = Trim(vRecordset.Fields("placeid").Value)
        Form312.Text302.Text = Trim(vRecordset.Fields("district").Value)
        Form312.Text303.Text = Trim(vRecordset.Fields("amphur").Value)
        Form312.Text304.Text = Trim(vRecordset.Fields("province").Value)
        Form312.Text305.Text = Trim(vRecordset.Fields("placemydescription").Value)
        Form312.Text306.Text = Trim(vRecordset.Fields("routeid").Value)
        Form312.Text307.Text = Trim(vRecordset.Fields("name1").Value)
        Form312.Text308.Text = Trim(vRecordset.Fields("name2").Value)
        Form312.Text309.Text = Trim(vRecordset.Fields("routemydescription").Value)
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = Form312.ListView101.ListItems.Add(, , Trim(i))
            vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
            vSearchList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
            vSearchList.SubItems(3) = Trim(vRecordset.Fields("itemname").Value)
            vSearchList.SubItems(4) = Trim(vRecordset.Fields("qty").Value)
            vSearchList.SubItems(5) = Trim(vRecordset.Fields("remainqty").Value)
            vSearchList.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
            vSearchList.SubItems(7) = Trim(vRecordset.Fields("refid").Value)
            vSearchList.SubItems(8) = Trim(vRecordset.Fields("refsubid").Value)
            vSearchList.SubItems(9) = Trim(vRecordset.Fields("iscancelsub").Value)
            If Trim(vRecordset.Fields("iscancelsub").Value) = 1 Then
                Form312.ListView101.ListItems(i).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
                Form312.ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H000000FF"
            End If
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
    Unload FrmOrder006
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
    vQuery = "exec bcnp.dbo.USP_DO_QueueHeaderSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(i))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("duedate").Value)
        vSearchList.SubItems(4) = Trim(vRecordset.Fields("priority").Value)
        vSearchList.SubItems(5) = Trim(vRecordset.Fields("receivename").Value)
        vSearchList.SubItems(6) = Trim(vRecordset.Fields("receivetelhome").Value)
        vSearchList.SubItems(7) = Trim(vRecordset.Fields("receivetelmobile").Value)
        If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
            ListView101.ListItems(i).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
            ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
        End If
        vRecordset.MoveNext
        i = i + 1
        Wend
    Else
        MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
    End If
    vRecordset.Close
End If
End Sub


