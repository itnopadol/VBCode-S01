VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form105 
   Caption         =   "เช็คบาร์โค้ดตรวจสอบที่เก็บ"
   ClientHeight    =   7980
   ClientLeft      =   2055
   ClientTop       =   645
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form105.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   10920
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICReceiveAddShelf 
      Height          =   7980
      Left            =   -90
      Picture         =   "Form105.frx":72FB
      ScaleHeight     =   7920
      ScaleWidth      =   10980
      TabIndex        =   16
      Top             =   -45
      Width           =   11040
      Begin VB.PictureBox PICItemAddShelf 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   270
         ScaleHeight     =   1425
         ScaleWidth      =   10470
         TabIndex        =   29
         Top             =   4095
         Width           =   10500
         Begin VB.Label Label11 
            Caption         =   "Label11"
            Height          =   420
            Left            =   1890
            TabIndex        =   31
            Top             =   225
            Width           =   2265
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            Height          =   330
            Left            =   225
            TabIndex        =   30
            Top             =   180
            Width           =   1320
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   555
         Left            =   9135
         TabIndex        =   28
         Top             =   6435
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   555
         Left            =   7155
         TabIndex        =   27
         Top             =   6345
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   330
         Left            =   360
         TabIndex        =   26
         Top             =   2250
         Width           =   375
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   5985
         TabIndex        =   22
         Text            =   "Combo3"
         Top             =   2295
         Width           =   1770
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   4005
         TabIndex        =   21
         Text            =   "Combo2"
         Top             =   2295
         Width           =   1140
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1935
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   2250
         Width           =   1140
      End
      Begin VB.TextBox TXTDocNo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1665
         TabIndex        =   19
         Top             =   1260
         Width           =   2535
      End
      Begin MSComctlLib.ListView ListViewItem 
         Height          =   2985
         Left            =   270
         TabIndex        =   17
         Top             =   3015
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   5265
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2850
         Left            =   45
         TabIndex        =   32
         Top             =   4455
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   5027
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
         Appearance      =   0
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "บาร์โค้ด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "จำนวนที่ต้องการ"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ชื่อสินค้า"
            Object.Width           =   7585
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วยนับ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "คลัง"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ชั้นเก็บ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "FamilyGroup"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ZoneID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "PickZone"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   420
         Left            =   5355
         TabIndex        =   25
         Top             =   2340
         Width           =   510
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   285
         Left            =   3330
         TabIndex        =   24
         Top             =   2385
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   285
         Left            =   765
         TabIndex        =   23
         Top             =   2295
         Width           =   915
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบรับเข้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   495
         TabIndex        =   18
         Top             =   1305
         Width           =   1095
      End
   End
   Begin VB.CommandButton CMDAddROShelf 
      BackColor       =   &H00C0C0C0&
      Caption         =   "บันทึกที่เก็บตามใบรับ"
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
      Left            =   7110
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6345
      Width           =   1860
   End
   Begin VB.ComboBox CMBZone 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1290
      Width           =   1290
   End
   Begin VB.CommandButton CMD104 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ตรวจสอบที่เก็บ"
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
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6345
      Width           =   1680
   End
   Begin VB.CommandButton Cmd103 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ลบรายการ"
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
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6345
      Width           =   1665
   End
   Begin VB.TextBox Text103 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   5580
      TabIndex        =   4
      Top             =   2055
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Cmd102 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Update Data"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1395
      Width           =   2415
   End
   Begin VB.ComboBox Cmb101 
      Height          =   315
      ItemData        =   "Form105.frx":1096E
      Left            =   1800
      List            =   "Form105.frx":10970
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   900
      Width           =   1290
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton Cmd101 
      BackColor       =   &H00C0C0C0&
      Caption         =   "เก็บข้อมูล"
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6345
      Width           =   1665
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3690
      Left            =   1800
      TabIndex        =   5
      Top             =   2475
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   6509
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสบาร์โค้ด"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อสินค้า"
         Object.Width           =   7144
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "หน่วยนับ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ที่เก็บ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "รหัสสินค้า"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "คลัง"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "โซน"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   2055
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "โซนชั้นเก็บ"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   14
      Top             =   1305
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อผู้เช็คที่เก็บ"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   4275
      TabIndex        =   12
      Top             =   2070
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "คลังที่ยิง"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   1125
      TabIndex        =   10
      Top             =   945
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสบาร์โค้ด"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   900
      TabIndex        =   9
      Top             =   2055
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ที่เก็บ"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   1200
      TabIndex        =   8
      Top             =   1680
      Width           =   540
   End
End
Attribute VB_Name = "Form105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vItemClick1 As Integer
Private Sub Cmb101_Change()
CMBZone.SetFocus
End Sub

Private Sub Cmb101_Click()
CMBZone.SetFocus
End Sub

Private Sub Cmb101_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CMBZone.SetFocus
End If
End Sub

Private Sub CMBZone_Change()
Text101.SetFocus
End Sub

Private Sub CMBZone_Click()
Text101.SetFocus
End Sub

Private Sub CMBZone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text101.SetFocus
End If
End Sub

Private Sub Cmd101_Click()

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then

            If vConnect = 1 Then
                    If ListView101.ListItems.Item(1).SubItems(1) <> "" Then
                                If Text103.Text = "" Then
                                    MsgBox "กรุณาใส่ชื่อผู้ทำการเช็คบาร์โค้ดสินค้าด้วยนะครับ และกดปุ่มเก็บข้อมูลใหม่"
                                    Text103.SetFocus
                                    Exit Sub
                                Else
                                    vUserID = Trim(Text103.Text)
                                End If
                    Else
                        MsgBox "กรุณา กดปุ่ม Update Data อีกครั้งครับ"
                        Exit Sub
                    End If
            End If
            
            If Text101.Text <> "" Then
               Call InsertToLogs
               
            Else
            MsgBox "กรุณาใส่ข้อมูลชั้นเก็บด้วยนะครับ"
            Exit Sub
            End If
            ListView101.ListItems.Clear
            Text101.Text = ""
            Text103.Text = ""
            Text101.SetFocus
Else
    MsgBox "ไม่มีรายการให้เก็บข้อมูล"
    Text101.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer

On Error GoTo ErrDescription

If vConnect = 1 Then
Call InitializeDatabase
    If Cmb101.ListCount = 0 Then
        Call WHCode
        Call ZoneCode
    End If
    If Cmb101.Text <> "" Then
        Text103.Visible = True
        Label4.Visible = True
        For i = 1 To ListView101.ListItems.Count
        vQuery = "select * from bcnp.dbo.vw_MB_CheckPrintLabel where barcode = '" & ListView101.ListItems.Item(i).Text & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            ListView101.ListItems(i).ListSubItems(1).Text = Trim(vRecordset.Fields("name1").Value)
            ListView101.ListItems(i).ListSubItems(2).Text = Trim(vRecordset.Fields("unitcode").Value)
             
        Else
            ListView101.ListItems(i).ListSubItems(1).Text = Trim("ไม่มีรหัสสินค้า ")
            MsgBox "ไม่มี บาร์โค้ด รหัส " & ListView101.ListItems.Item(i).Text & " ในฐานข้อมูล กรุณาแก้ไขรหัสบาร์โค้ดด้วยนะครับ"
        End If
        vRecordset.Close
        Next i
    Else
        MsgBox "กรุณาเลือก คลังก่อนนะครับ และ กดปุ่ม UpDate Data อีกรอบนะครับ"
        Exit Sub
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Dim vItemClick As Integer

If Me.ListView101.ListItems.Count > 0 Then
vItemClick = vItemClick1
ListView101.ListItems.Remove (vItemClick)
End If
End Sub

Private Sub CMD104_Click()
Form107.Show
Form107.SetFocus
End Sub

Private Sub Form_Load()
If vConnect = 0 Then
Call WHCode
Call ZoneCode
Else
CMD102.Enabled = True
End If
End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
vItemClick1 = Item.Index
End Sub

Private Sub Text101_LostFocus()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckShelfExist As Integer
Dim vWHCode As String
Dim vShelfCode As String

On Error GoTo ErrDescription

If Me.Text101.Text <> "" Then
   vWHCode = Me.Cmb101.Text
   vShelfCode = Me.Text101.Text
   
   vQuery = "select isnull(count(code),0) as vCount from  Npmaster.dbo.TB_RC_Shelf where whcode = '" & vWHCode & "' and code = '" & vShelfCode & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckShelfExist = vRecordset.Fields("vcount").Value
   End If
   vRecordset.Close
   If vCheckShelfExist = 0 And Text101.Text <> "" Then
      MsgBox "ไม่พบทะเบียนที่เก็บที่ต้องการบันทึก กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      Me.Text101.SetFocus
   Else
      Me.Text101.Text = UCase(Me.Text101.Text)
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text102_GotFocus()
If Text101.Text = "" Then
    MsgBox "กรุณาใส่ข้อมูลชั้นเก็บด้วยครับ"
    Text101.SetFocus
End If
End Sub

Private Sub Text102_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vBarCode As String
Dim ListItem As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If vConnect = 0 Then
        Call InsertToGrid
        Text102.Text = ""
        Text102.SetFocus
    Else
        Call InsertToGridUnConnect
        Text102.Text = ""
        Text102.SetFocus
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub


Private Sub Text101_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If vConnect = 0 Then
            If Cmb101.Text <> "" Then
                Text102.SetFocus
            Else
                MsgBox "เลือกคลังที่ยิงบาร์โค้ดด้วยนะครับ"
                Cmb101.SetFocus
                Exit Sub
            End If
      Else
            Text102.SetFocus
      End If
End If
End Sub


Public Sub WHCode()
Dim vRecordset  As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

vQuery = "select distinct whcode from bcstklocation  where whcode <> '' and  whcode is not null and whcode <> '-'  order by whcode desc"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Cmb101.AddItem Trim(vRecordset.Fields("whcode").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub

Public Sub ZoneCode()
Dim vRecordset  As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

'vQuery = "select distinct shelfcode from bcstklocation  where shelfcode <> '' and  shelfcode is not null and shelfcode <> '-'  order by shelfcode"
vQuery = "exec dbo.USP_NP_SearchShelf"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMBZone.AddItem Trim(vRecordset.Fields("code").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub

Public Sub InsertToLogs()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vTemp1 As String, vTemp2 As String, vTemp3 As String, vTemp4 As String, vTemp5 As String, vTemp6 As String, vTemp7 As String, vTemp8 As String, vTemp9 As String, vTemp10 As String

On Error GoTo ErrDescription

If Me.ListView101.ListItems.Count > 0 Then
   For i = 1 To ListView101.ListItems.Count
           vTemp1 = Trim(ListView101.ListItems.Item(i).Text)
           vTemp2 = Trim(ListView101.ListItems.Item(i).SubItems(1))
           vTemp3 = Trim(ListView101.ListItems.Item(i).SubItems(5))
           vTemp4 = Trim(ListView101.ListItems.Item(i).SubItems(3))
           vTemp5 = Now
           vTemp6 = Trim(vUserID)
           vTemp7 = Trim(ListView101.ListItems.Item(i).SubItems(2))
           vTemp8 = Trim(ListView101.ListItems.Item(i).SubItems(4))
           vTemp9 = "หน้ายิงบาร์โค้ดตรวจสอบชั้นเก็บ"
           vTemp10 = Trim(ListView101.ListItems.Item(i).SubItems(6))
           
           vQuery = "exec dbo.USP_NP_InsertScanItemShelfCode  '" & vTemp8 & "','" & vTemp1 & "','" & vTemp2 & "','" & vTemp7 & "','" & vTemp3 & "','" & vTemp10 & "','" & vTemp4 & "','" & vTemp6 & "','" & vTemp9 & "' "
           gConnection.Execute vQuery
   Next i
   MsgBox "โปรแกรมทำการเก็บข้อมูล คลัง " & vTemp3 & " ตามชั้นเก็บที่กำหนดไว้ เรียบร้อยแล้วครับ"
Else
   MsgBox "ไม่มีรายการสินค้าที่จะบันทึกที่เก็บ"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub InsertToGrid()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vBarCode As String
Dim ListItem As ListItem

If Me.Cmb101.Text <> "" Then
   If Me.CMBZone.Text <> "" Then
      vBarCode = Trim(Text102.Text)
      vQuery = "exec dbo.USP_MB_ScanBarCode '" & vBarCode & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          Set ListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("barcode").Value))
          ListItem.SubItems(1) = Trim(vRecordset.Fields("itemname").Value)
          ListItem.SubItems(2) = Trim(vRecordset.Fields("unitcode").Value)
          ListItem.SubItems(3) = Trim(Text101.Text)
          ListItem.SubItems(4) = Trim(vRecordset.Fields("itemcode").Value)
          ListItem.SubItems(5) = Trim(Cmb101.Text)
          ListItem.SubItems(6) = Trim(CMBZone.Text)
      Else
          MsgBox "ไม่มีรหัสสินค้า " & vBarCode & " นี้อยู่ในระบบ ", vbInformation, "โปรดตรวจสอบ"
          Set ListItem = ListView101.ListItems.Add(, , vBarCode)
          ListItem.SubItems(1) = Trim("หาชื่อไม่เจอ")
          ListItem.SubItems(2) = Trim("หน่วยไม่เจอ")
          ListItem.SubItems(3) = Trim(Text101.Text)
          ListItem.SubItems(4) = Trim("หารหัสไม่เจอ")
          ListItem.SubItems(5) = Trim(Cmb101.Text)
          ListItem.SubItems(6) = Trim(CMBZone.Text)
      End If
      vRecordset.Close
   Else
      MsgBox "ต้องระบุโซนสินค้าในการระบุที่เก็บสินค้าด้วย", vbCritical, "Send Error Message"
   End If
Else
   MsgBox "ต้องระบุคลังสินค้าในการระบุที่เก็บสินค้าด้วย", vbCritical, "Send Error Message"
End If

End Sub

Public Sub InsertToGridUnConnect()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vBarCode As String
Dim ListItem As ListItem


vBarCode = Trim(Text102.Text)
Set ListItem = ListView101.ListItems.Add(, , Trim(vBarCode))
ListItem.SubItems(3) = Trim(Text101.Text)


End Sub

