VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form55 
   Caption         =   "เพิ่มทะเบียนสถานที่ขนส่ง"
   ClientHeight    =   8280
   ClientLeft      =   3405
   ClientTop       =   1530
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form55.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text106 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2805
      TabIndex        =   5
      Top             =   2550
      Width           =   3090
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "บันทึกข้อมูล"
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
      Left            =   8955
      TabIndex        =   10
      Top             =   6600
      Width           =   1365
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   7380
      TabIndex        =   6
      Top             =   2550
      Width           =   2940
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "เพิ่มรายการ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1605
      TabIndex        =   8
      Top             =   3600
      Width           =   1065
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2415
      Left            =   1605
      TabIndex        =   9
      Top             =   4050
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   4260
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
         Text            =   "รหัสสถานที่"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "สถานที่ขนส่ง"
         Object.Width           =   10231
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "รหัสผู้ติดต่อ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "เบอร์โทร"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "รหัสขนส่ง"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "คำอธิบาย"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.TextBox Text107 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2805
      TabIndex        =   7
      Top             =   3000
      Width           =   7515
   End
   Begin VB.TextBox Text104 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6330
      TabIndex        =   3
      Top             =   1575
      Width           =   3990
   End
   Begin VB.TextBox Text105 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2805
      TabIndex        =   4
      Top             =   2100
      Width           =   7515
   End
   Begin VB.TextBox Text103 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2805
      TabIndex        =   2
      Top             =   1575
      Width           =   2415
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   5430
      TabIndex        =   1
      Top             =   1050
      Width           =   4890
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2805
      TabIndex        =   0
      Top             =   1050
      Width           =   1740
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เบอร์โทร"
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
      Height          =   315
      Left            =   1830
      TabIndex        =   18
      Top             =   2550
      Width           =   840
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
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
      Height          =   315
      Left            =   1905
      TabIndex        =   17
      Top             =   3000
      Width           =   840
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสผู้ติดต่อ"
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
      Height          =   315
      Left            =   5280
      TabIndex        =   16
      Top             =   1575
      Width           =   1065
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกรหัสขนส่ง"
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
      Height          =   315
      Left            =   5955
      TabIndex        =   15
      Top             =   2550
      Width           =   1290
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สถานที่ขนส่ง"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1605
      TabIndex        =   14
      Top             =   2100
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสสถานที่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1605
      TabIndex        =   13
      Top             =   1575
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   4605
      TabIndex        =   12
      Top             =   1050
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1605
      TabIndex        =   11
      Top             =   1050
      Width           =   1065
   End
End
Attribute VB_Name = "Form55"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDeliveryCode As String
Dim vDeliveryName As String
Dim vAddListviewAddr As ListItem
Dim vContactName As String
Dim vTransport As String
Dim vMydescription As String
Dim vTel As String

On Error GoTo ErrDescription

If Text101.Text <> "" And Text102.Text <> "" And Text103.Text <> "" And Text105.Text <> "" Then
    vDeliveryCode = Trim(Text103.Text)
    vDeliveryName = Trim(Text105.Text)
    vContactName = Trim(Text104.Text)
    vTransport = Trim(CMB101.Text)
    vTel = Trim(Text106.Text)
    vMydescription = Trim(Text107.Text)
    Set vAddListviewAddr = ListView101.ListItems.Add(, , vDeliveryCode)
    vAddListviewAddr.SubItems(1) = vDeliveryName
    vAddListviewAddr.SubItems(2) = vContactName
    vAddListviewAddr.SubItems(3) = vTel
    vAddListviewAddr.SubItems(4) = vTransport
    vAddListviewAddr.SubItems(5) = vMydescription
    Text103.Text = ""
    Text104.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    CMB101.Text = ""
Else
    MsgBox "กรุณากรอกข้อมูลในช่อง ที่มีตัวหนังสือสีแดงให้ครบด้วย", vbInformation, "ข้อความแจ้งเตือน"
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
Dim vDeliveryCode As String
Dim vDeliveryName As String
Dim vAddListviewAddr As ListItem
Dim vContactName As String
Dim vTransport As String
Dim vMydescription As String
Dim vARCode As String
Dim i  As Integer
Dim vExist As Integer
Dim vTel As String

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    For i = 1 To ListView101.ListItems.Count
        vARCode = Trim(Text101.Text)
        vDeliveryCode = ListView101.ListItems.Item(i).Text
        vQuery = "select code from bcnp.dbo.bcdeliveryaddr where arcode = '" & vARCode & "' and code = '" & vDeliveryCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vExist = 1
        Else
            vExist = 0
        End If
        vRecordset.Close
        If vExist = 0 Then
            vDeliveryName = Trim(ListView101.ListItems.Item(i).SubItems(1))
            vContactName = Trim(ListView101.ListItems.Item(i).SubItems(2))
            vTel = Trim(ListView101.ListItems.Item(i).SubItems(3))
            If ListView101.ListItems.Item(i).SubItems(4) <> "" Then
                vTransport = Left(Trim(Trim(ListView101.ListItems.Item(i).SubItems(4))), InStr((Trim(Trim(ListView101.ListItems.Item(i).SubItems(4)))), "-") - 1)
            Else
                vTransport = ""
            End If
            vMydescription = Trim(ListView101.ListItems.Item(i).SubItems(5))
            vQuery = "exec bcnp.dbo.usp_AR_DeliveryAddr '" & vARCode & "','" & vDeliveryCode & "','" & vDeliveryName & "','" & vTel & "','" & vTransport & "','" & vContactName & "','" & vMydescription & "' "
            gConnection.Execute vQuery
        End If
    Next i
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text104.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text107.Text = ""
    CMB101.Text = ""
    ListView101.ListItems.Clear
    MsgBox "ได้ทำการบันทึกข้อมูลเรียบร้อยแล้ว", vbInformation, "ข้อความแจ้งเตือน"
Else
    MsgBox "ไม่มีรายการสถานที่ขนส่งให้บันทึก", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

CMB101.Clear
vQuery = "select code+'-'+name as name1 from BCTransport order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB101.AddItem Trim(vRecordset.Fields("name1").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub

Private Sub ListView101_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode As String
Dim vDeliveryCode As String
Dim vExist As Integer
Dim vAnswer As Integer

On Error GoTo ErrDescription

If KeyCode = 46 Then
    vARCode = Trim(Text101.Text)
    vDeliveryCode = ListView101.SelectedItem.Text
    
    vQuery = "select code from bcnp.dbo.bcdeliveryaddr where arcode = '" & vARCode & "' and code = '" & vDeliveryCode & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vExist = 1
    Else
        vExist = 0
    End If
    vRecordset.Close
    
    vAnswer = MsgBox("คุณต้องการลบข้อมูลสถานที่ขนส่งรหัส " & vDeliveryCode & " นี้ใช่หรือไม่", vbYesNo, "ข้อความสอบถาม")
    
    If vAnswer = 6 Then
        vQuery = "delete bcnp.dbo.bcdeliveryaddr where arcode = '" & vARCode & "' and code = '" & vDeliveryCode & "' "
        gConnection.Execute vQuery
        ListView101.ListItems.Remove (ListView101.SelectedItem.Index)
    ElseIf vAnswer = 7 Then
       ListView101.ListItems.Remove (ListView101.SelectedItem.Index)
    End If
    MsgBox "ลบข้อมูลเรียบร้อยแล้วครับ"
End If
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text101_Change()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode As String
Dim vListDelivery As ListItem

On Error GoTo ErrDescription

    If Text101.Text <> "" Then
        vARCode = Trim(Text101.Text)
        vQuery = "select name1 from bcnp.dbo.bcar where code = '" & vARCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            Text102.Text = Trim(vRecordset.Fields("name1").Value)
        Else
           Text102.Text = ""
            ListView101.ListItems.Clear
            Exit Sub
        End If
        vRecordset.Close
        vQuery = "select  *  from bcnp.dbo.vw_ar_BCDeliveryAddr where arcode = '" & vARCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set vListDelivery = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("code").Value))
            vListDelivery.SubItems(1) = Trim(vRecordset.Fields("address").Value)
            vListDelivery.SubItems(2) = Trim(vRecordset.Fields("contact").Value)
            vListDelivery.SubItems(3) = Trim(vRecordset.Fields("tel").Value)
            vListDelivery.SubItems(4) = Trim(vRecordset.Fields("transportname").Value)
            vListDelivery.SubItems(5) = Trim(vRecordset.Fields("mydescription").Value)
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
    'Else
        'MsgBox "กรุณากรอกรหัสลูกค้าด้วย"
    End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode As String
Dim vListDelivery As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Text101.Text <> "" Then
        vARCode = Trim(Text101.Text)
        vQuery = "select name1 from bcnp.dbo.bcar where code = '" & vARCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            Text102.Text = Trim(vRecordset.Fields("name1").Value)
        Else
            MsgBox "ไม่มีทะเบียนรหัสลูกค้า '" & vARCode & "' ในระบบ กรุณาตรวจสอบด้วย"
            Exit Sub
        End If
        vRecordset.Close
        ListView101.ListItems.Clear
        vQuery = "select  *  from bcnp.dbo.vw_ar_BCDeliveryAddr where arcode = '" & vARCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set vListDelivery = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("code").Value))
            vListDelivery.SubItems(1) = Trim(vRecordset.Fields("address").Value)
            vListDelivery.SubItems(2) = Trim(vRecordset.Fields("contact").Value)
            vListDelivery.SubItems(3) = Trim(vRecordset.Fields("tel").Value)
            vListDelivery.SubItems(4) = Trim(vRecordset.Fields("transportname").Value)
            vListDelivery.SubItems(5) = Trim(vRecordset.Fields("mydescription").Value)
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
    Else
        MsgBox "กรุณากรอกรหัสลูกค้าด้วย"
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
