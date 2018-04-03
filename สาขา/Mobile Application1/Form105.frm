VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form105 
   Caption         =   "เช็คบาร์โค้ดตรวจสอบที่เก็บ"
   ClientHeight    =   9090
   ClientLeft      =   2055
   ClientTop       =   645
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form105.frx":0000
   ScaleHeight     =   9090
   ScaleWidth      =   14385
   WindowState     =   2  'Maximized
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
      TabIndex        =   12
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
      TabIndex        =   6
      Top             =   6345
      Width           =   1665
   End
   Begin VB.TextBox Text103 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1800
      TabIndex        =   3
      Top             =   2025
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
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   945
      Width           =   1380
   End
   Begin VB.ComboBox Cmb101 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Text            =   "S02"
      Top             =   900
      Width           =   1290
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   1275
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
      TabIndex        =   5
      Top             =   6345
      Width           =   1665
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3690
      Left            =   1800
      TabIndex        =   4
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
      NumItems        =   6
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
         Text            =   "Zone"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   1650
      Width           =   2415
   End
   Begin VB.Label LBLZone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4365
      TabIndex        =   13
      Top             =   2025
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อผู้เช็คที่เก็บ"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   600
      TabIndex        =   11
      Top             =   2025
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   1650
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ที่เก็บ"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   1200
      TabIndex        =   7
      Top             =   1275
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
Text102.SetFocus
End Sub

Private Sub Cmb101_Click()
Text101.SetFocus
End Sub

Private Sub Cmb101_KeyPress(KeyAscii As Integer)
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
            Me.LBLZone.Caption = ""
            Exit Sub
            End If
            ListView101.ListItems.Clear
            Text101.Text = ""
            Text103.Text = ""
            LBLZone.Caption = ""
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
   
   vQuery = "exec dbo.USP_NP_SearchShelfZone '" & vWHCode & "' , '" & vShelfCode & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckShelfExist = vRecordset.Fields("vcount").Value
      Me.LBLZone.Caption = vRecordset.Fields("zoneid").Value
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
    LBLZone.Caption = ""
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

vQuery = "select distinct whcode from bcstklocation  where whcode <> '' and  whcode is not null and whcode <> '-'  order by whcode"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Cmb101.AddItem Trim(vRecordset.Fields("whcode").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub

Public Sub InsertToLogs()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vTemp1 As String, vTemp2 As String, vTemp3 As String, vTemp4 As String, vTemp5 As String, vTemp6 As String, vTemp7 As String, vTemp8 As String, vTemp9 As String

On Error GoTo ErrDescription

If Me.ListView101.ListItems.Count > 0 Then
   For i = 1 To ListView101.ListItems.Count
           vTemp1 = Trim(ListView101.ListItems.Item(i).Text)
           vTemp2 = Trim(ListView101.ListItems.Item(i).SubItems(1))
           vTemp3 = Trim(Cmb101.Text)
           vTemp4 = Trim(ListView101.ListItems.Item(i).SubItems(3))
           vTemp5 = Trim(ListView101.ListItems.Item(i).SubItems(5))
           vTemp6 = Trim(vUserID)
           vTemp7 = Trim(ListView101.ListItems.Item(i).SubItems(2))
           vTemp8 = Trim(ListView101.ListItems.Item(i).SubItems(4))
           vTemp9 = "หน้ายิงบาร์โค้ดตรวจสอบชั้นเก็บ"
           
           vQuery = "exec dbo.USP_NP_InsertScanItemShelfCode  '" & vTemp8 & "','" & vTemp1 & "','" & vTemp2 & "','" & vTemp7 & "','" & vTemp3 & "','" & vTemp5 & "','" & vTemp4 & "','" & vTemp6 & "','" & vTemp9 & "' "
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

If Me.LBLZone.Caption <> "" Then
vBarCode = Trim(Text102.Text)
vQuery = "exec dbo.USP_MB_ScanBarCode '" & vBarCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Set ListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("barcode").Value))
    ListItem.SubItems(1) = Trim(vRecordset.Fields("itemname").Value)
    ListItem.SubItems(2) = Trim(vRecordset.Fields("unitcode").Value)
    ListItem.SubItems(3) = Trim(Text101.Text)
    ListItem.SubItems(4) = Trim(vRecordset.Fields("itemcode").Value)
    ListItem.SubItems(5) = Me.LBLZone.Caption
Else
    MsgBox "ไม่มีรหัสสินค้า " & vBarCode & " นี้อยู่ในระบบ ", vbInformation, "โปรดตรวจสอบ"
    Set ListItem = ListView101.ListItems.Add(, , vBarCode)
    ListItem.SubItems(1) = Trim("หาชื่อไม่เจอ")
    ListItem.SubItems(2) = Trim("หน่วยไม่เจอ")
    ListItem.SubItems(3) = Trim(Text101.Text)
    ListItem.SubItems(4) = Trim("หารหัสไม่เจอ")
    ListItem.SubItems(5) = Trim("หาโซนไม่เจอ")
End If
vRecordset.Close
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

