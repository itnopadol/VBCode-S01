VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form104 
   Caption         =   "เช็คบาร์โค้ดทำใบขอโอน"
   ClientHeight    =   7995
   ClientLeft      =   2655
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form104.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   10950
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   3600
      Top             =   7245
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Cmd108 
      Caption         =   "เพิ่มข้อมูล"
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
      Left            =   7200
      TabIndex        =   12
      Top             =   5775
      Width           =   1665
   End
   Begin VB.CommandButton Cmd107 
      Caption         =   "ลบข้อมูลในGrid"
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
      Left            =   5025
      TabIndex        =   11
      Top             =   5775
      Width           =   1665
   End
   Begin VB.ComboBox Cmb102 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1875
      TabIndex        =   5
      Top             =   2325
      Width           =   2790
   End
   Begin VB.CommandButton Cmd106 
      Caption         =   "ดูข้อมูลเก่า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   675
      TabIndex        =   4
      Top             =   2325
      Width           =   1140
   End
   Begin VB.CommandButton Cmd105 
      Caption         =   "UpDate Data"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5175
      TabIndex        =   6
      Top             =   600
      Width           =   1665
   End
   Begin VB.ComboBox Cmb101 
      Height          =   315
      Left            =   1200
      TabIndex        =   13
      Top             =   6300
      Width           =   3315
   End
   Begin VB.TextBox Text105 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   5550
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   6300
      Width           =   2115
   End
   Begin VB.CommandButton Cmd104 
      Caption         =   "ทำใบขอโอน"
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
      Left            =   675
      TabIndex        =   14
      Top             =   6750
      Width           =   1665
   End
   Begin VB.TextBox Text106 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   8925
      TabIndex        =   16
      Top             =   6300
      Width           =   1890
   End
   Begin VB.TextBox Text104 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   1425
      Width           =   1065
   End
   Begin VB.TextBox Text103 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1875
      TabIndex        =   1
      Top             =   1425
      Width           =   990
   End
   Begin VB.CommandButton Cmd103 
      Caption         =   "พิมพ์ใบขอโอน"
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
      Left            =   8970
      TabIndex        =   17
      Top             =   6825
      Width           =   1845
   End
   Begin VB.CommandButton Cmd102 
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
      Height          =   390
      Left            =   2850
      TabIndex        =   10
      Top             =   5775
      Width           =   1665
   End
   Begin VB.CommandButton Cmd101 
      Caption         =   "แก้ไขรายการ"
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
      Left            =   675
      TabIndex        =   9
      Top             =   5775
      Width           =   1665
   End
   Begin MSComctlLib.ListView ListView102 
      Height          =   1665
      Left            =   5175
      TabIndex        =   7
      Top             =   975
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   2937
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "คลัง"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชั้นเก็บ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "จำนวนสินค้าคงเหลือ"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "หน่วยนับ"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1875
      TabIndex        =   3
      Top             =   1875
      Width           =   2790
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1875
      TabIndex        =   0
      Top             =   975
      Width           =   2790
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2940
      Left            =   675
      TabIndex        =   8
      Top             =   2700
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   5186
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
         Text            =   "บาร์โค้ด"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "คลังที่ขอโอน"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เข้าคลัง"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "จำนวนขอโอน"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "หน่วยนับ"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "แผนก"
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
      Left            =   675
      TabIndex        =   24
      Top             =   6300
      Width           =   465
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "คำอธิบาย"
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
      Left            =   4725
      TabIndex        =   23
      Top             =   6300
      Width           =   765
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบขอโอน"
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
      Left            =   7800
      TabIndex        =   22
      Top             =   6300
      Width           =   1065
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "เข้าคลัง"
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
      Left            =   2970
      TabIndex        =   21
      Top             =   1425
      Width           =   690
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "โอนจากคลัง"
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
      Left            =   675
      TabIndex        =   20
      Top             =   1500
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวนที่ขอโอน"
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
      Left            =   675
      TabIndex        =   19
      Top             =   1875
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "บาร์โค้ด"
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
      Left            =   675
      TabIndex        =   18
      Top             =   975
      Width           =   1290
   End
End
Attribute VB_Name = "Form104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCheck As Integer
Dim vCheck1 As Integer

Private Sub Cmb102_Click()
ListView101.ListItems.Clear
Call GetTransferNo
Text106.Text = Cmb102.Text
End Sub

Private Sub Cmd101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String, vItemCode As String
Dim vIsConfirm As Integer, vAverageCost As Integer
Dim vFromWH As String, vTOWH As String, vQty As String
Dim vCountQty As Integer, vLine As Integer

On Error GoTo ErrDescription
'แก้ไขจำนวนที่ต้องการพิมพ์
If vCheck <> 0 Then
If Text106.Text = "" Then
    ListView101.ListItems(vCheck).ListSubItems(3).Text = Trim(Text103.Text)
    ListView101.ListItems(vCheck).ListSubItems(4).Text = Trim(Text104.Text)
    ListView101.ListItems(vCheck).ListSubItems(5).Text = Trim(Text102.Text)
    vCheck = 0
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text104.Text = ""
    Text101.SetFocus
Else
    vDocNo = Trim(Text106.Text)
    vQuery = "select isconfirm from bcnp.dbo.bcstktransfer where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
    End If
    vRecordset.Close
    If vIsConfirm = 0 Then
        vItemCode = Trim(Trim(ListView101.ListItems.Item(vCheck).Text))
        vFromWH = Trim(Text103.Text)
        vTOWH = Trim(Text104.Text)
        vQty = Trim(Text102.Text)
        vLine = vCheck - 1
        
        vQuery = "select averagecost from bcnp.dbo.bcitem  where code = '" & ListView101.ListItems.Item(vCheck).SubItems(1) & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
             If Not IsNull(vRecordset.Fields("averagecost").Value) Then
                vAverageCost = Trim(vRecordset.Fields("averagecost").Value)
             Else
                vAverageCost = 0
             End If
        End If
        vRecordset.Close
        'อัพเดทข้อมูลที่ได้แก้ไขไป
        vQuery = "Update bcnp.dbo.bcstktransfsub2 set fromwh = '" & vFromWH & "' ,towh = '" & vTOWH & "', " _
                        & " qty = " & vQty & " , sumofcost = " & vAverageCost & "* " & vQty & "  " _
                        & " where docno = '" & vDocNo & "' and itemcode ='" & vItemCode & "' and linenumber = " & vLine & " "
        gConnection.Execute vQuery
        
        ListView101.ListItems.Item(vCheck).SubItems(3) = Trim(Text103.Text)
        ListView101.ListItems.Item(vCheck).SubItems(4) = Trim(Text104.Text)
        ListView101.ListItems.Item(vCheck).SubItems(5) = Trim(Text102.Text)
        
        vQuery = "select sum(qty) as countqty from bcnp.dbo.bcstktransfsub2 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCountQty = Trim(vRecordset.Fields("countqty").Value)
        End If
        vRecordset.Close
        vQuery = "update bcnp.dbo.bcstktransfer2 set sumofqty = " & vCountQty & " where docno = '" & vDocNo & "' "
        gConnection.Execute vQuery
        
    Else
        MsgBox "ไม่สามารถบันทึกได้ เนื่องจากเอกสารเลขที่ " & vDocNo & "  ถูกอ้างอิงไปแล้ว"
    End If
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text104.Text = ""
    Text101.SetFocus
    vCheck = 0
End If
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
Dim vDocNo As String
Dim vItemCode As String
Dim vCountQty As Integer, vLine As Integer

On Error GoTo ErrDescription
'ลบรายการที่ยิงไว้
If vCheck <> 0 Then
    If Text106.Text = "" Then
        If ListView101.ListItems.Count <> 0 Then
            ListView101.ListItems.Remove (vCheck)
            Text101.Text = ""
            Text102.Text = ""
            Text103.Text = ""
            Text104.Text = ""
            Text101.SetFocus
        End If
    Else
        vLine = vCheck - 1
        vDocNo = Trim(Text106.Text)
        vItemCode = Trim(ListView101.ListItems.Item(vCheck).SubItems(1))
        vQuery = "Delete bcnp.dbo.bcstktransfsub2 where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "'  "
        gConnection.Execute vQuery
        
        vQuery = "select isnull(sum(qty),0) as countqty from bcnp.dbo.bcstktransfsub2 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCountQty = Trim(vRecordset.Fields("countqty").Value)
        End If
        vRecordset.Close
        vQuery = "update bcnp.dbo.bcstktransfer2 set sumofqty = " & vCountQty & " where docno = '" & vDocNo & "' "
        gConnection.Execute vQuery
        ListView101.ListItems.Remove (vCheck)
        
        '---------------------------------
    End If
End If

Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
Text104.Text = ""
Text101.SetFocus

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD103_Click()
If Text106.Text <> "" Then
    Call PrintTransferDoc
    ListView101.ListItems.Clear
    Cmb101.Text = ""
    Text105.Text = ""
    Text106.Text = ""
End If
End Sub

Private Sub CMD104_Click()

On Error GoTo ErrDescription
If Text106.Text = "" Then
        If ListView101.ListItems.Count <> 0 Then
            If ListView101.ListItems.Item(1).SubItems(2) <> "" Then
                    If Cmb101.Text <> "" Then
                        Call GenTransfer
                        ListView101.ListItems.Clear
                        Cmb101.Text = ""
                        Text105.Text = ""
                    Else
                        MsgBox "กรุณาเลือกแผนกที่ทำใบขอโอนสินค้าด้วยนะครับ ถ้าให้ดีใส่ข้อความหมายเหตุที่ คำอธิบายด้วยนะครับ"
                    End If
            Else
                MsgBox "กรุณา กดปุ่ม Update Data เพื่อทำการลิงค์ข้อมูลด้วยนะครับ"
            End If
        End If
Else
    MsgBox "ไม่สามารถทำใบขอโอนได้เนื่องจาก เอกสารเลขที่ " & Cmb102.Text & " ได้บันทึกไปแล้ว"
End If
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD105_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim ListItem As ListItem

If vConnect = 1 Then
    Call InitializeDatabase
    Call Department
End If
For i = 1 To ListView101.ListItems.Count
    vQuery = "select * from bcnp.dbo. vw_IV_ProgStockChecking where barcode = '" & Trim(ListView101.ListItems.Item(i).Text) & "' and whcode = '" & ListView101.ListItems.Item(i).SubItems(3) & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        ListView101.ListItems.Item(i).SubItems(1) = Trim(vRecordset.Fields("code").Value)
        ListView101.ListItems.Item(i).SubItems(2) = Trim(vRecordset.Fields("name1").Value)
        ListView101.ListItems.Item(i).SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
    End If
    vRecordset.Close
Next i


End Sub

Private Sub Cmd106_Click()
If vConnect = 1 Then
    Call InitializeDatabase
End If
Cmb102.Enabled = True
Cmb102.Clear
Call TransferDocuments
ListView101.ListItems.Clear
Text106.Text = ""
End Sub

Private Sub CMD107_Click()
ListView101.ListItems.Clear
Text106.Text = ""
Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
Text104.Text = ""
ListView102.ListItems.Clear
Text101.SetFocus
Cmb102.Text = ""
Cmb102.Enabled = False
End Sub

Private Sub CMD108_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String, vItemCode As String, vFromWH As String
Dim vDocDate As String, vTOWH As String
Dim vQty As Integer, vAverageCost As Integer
Dim vUnitCode As String, vLine As Integer, vCountQty As Integer

On Error GoTo ErrDescription

If Text103.Text <> "" And Text104.Text <> "" And Text102.Text <> "" Then
vDocNo = Trim(Text106.Text)
vItemCode = Trim(Text101.Text)
vFromWH = Trim(Text103.Text)
vTOWH = Trim(Text104.Text)
vQty = Trim(Text102.Text)
vQuery = "select docdate from bcnp.dbo.bcstktransfer2 where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocDate = Format(Trim(vRecordset.Fields("docdate").Value), "mm/dd/yyyy")
End If
vRecordset.Close

vQuery = "select averagecost,defsaleunitcode from bcnp.dbo.bcitem  where code = '" & vItemCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     vAverageCost = Trim(vRecordset.Fields("averagecost").Value)
     vUnitCode = Trim(vRecordset.Fields("defsaleunitcode").Value)
End If
vRecordset.Close
        
vQuery = "select max(linenumber) as linenumber from bcnp.dbo.bcstktransfsub2 where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vLine = Trim(vRecordset.Fields("linenumber").Value) + 1
End If
vRecordset.Close

vQuery = "insert into bcnp.dbo.bcstktransfsub2 (docno,docdate,itemcode,fromwh,fromshelf,towh,toshelf,qty,sumofcost,unitcode,linenumber)" _
                & " values('" & vDocNo & "','" & vDocDate & "','" & vItemCode & "','" & vFromWH & "','-' ,'" & vTOWH & "','-'," & vQty & "," & vQty & "*" & vAverageCost & ",'" & vUnitCode & "'," & vLine & ")"
gConnection.Execute vQuery

vQuery = "select sum(qty) as countqty from bcnp.dbo.bcstktransfsub2 where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCountQty = Trim(vRecordset.Fields("countqty").Value)
End If
vRecordset.Close
vQuery = "update bcnp.dbo.bcstktransfer2 set sumofqty = " & vCountQty & " where docno = '" & vDocNo & "' "
gConnection.Execute vQuery

ListView101.ListItems.Clear
Call GetTransferNo

Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
Text104.Text = ""
Text101.SetFocus
Else
    MsgBox "คุณยังไม่ได้กรอกข้อมูลที่จะเพิ่มเลยนะครับ  กรุณากรอกให้ครบด้วยครับ"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
If vConnect = 0 Then
    Call Department
Else
    CMD105.Enabled = True
End If
End Sub


Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
vCheck = Item.Index
Text101 = ListView101.ListItems(vCheck).Text
Text102 = ListView101.ListItems(vCheck).SubItems(5)
Text103 = ListView101.ListItems(vCheck).SubItems(3)
Text104 = ListView101.ListItems(vCheck).SubItems(4)
End Sub


Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vBarCode As String
Dim ListItem As ListItem

On Error GoTo ErrDescription
If KeyAscii = 13 Then
    If vConnect = 0 Then
            vBarCode = Trim(Text101.Text)
            ListView102.ListItems.Clear
            vQuery = "select  a.itemcode,b.barcode,a.whcode,a.shelfcode,isnull(qty,0) as qty,a.unitcode " _
                            & "  from    bcstklocation a left    join bcbarcodemaster b on a.itemcode = b.itemcode " _
                            & "  where b.barcode = '" & vBarCode & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set ListItem = ListView102.ListItems.Add(, , Trim(vRecordset.Fields("whcode").Value))
                ListItem.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
                ListItem.SubItems(2) = Trim(vRecordset.Fields("qty").Value)
                ListItem.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
                vRecordset.MoveNext
                Wend
                Text103.SetFocus
            Else
                MsgBox "ไม่มีรหัสบาร์โค้ด " & vBarCode & " รหัสนี้ครับ กรุณาตรวจสอบด้วยครับ"
                ListView102.ListItems.Clear
                Text101.Text = ""
                Text101.SetFocus
            End If
            vRecordset.Close
    Else
            Text103.SetFocus
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text102_KeyPress(KeyAscii As Integer)

On Error GoTo ErrDescription
If KeyAscii = 13 Then
    If Text106.Text = "" Then
        If vConnect = 0 Then
            Call InsertToGrid
            ListView102.ListItems.Clear
        Else
            InsertToGrid_UnConnect
        End If
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Private Sub Text103_KeyPress(KeyAscii As Integer)

On Error GoTo ErrDescription
If KeyAscii = 13 Then
    Text104.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text104_KeyPress(KeyAscii As Integer)

On Error GoTo ErrDescription
If KeyAscii = 13 Then
    Text102.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub InsertToGrid()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListItem As ListItem
Dim vBarCode

On Error GoTo ErrDescription

vBarCode = Trim(Text101.Text)
vQuery = "select * from bcnp.dbo. vw_IV_ProgStockChecking where barcode = '" & vBarCode & "' and whcode = '" & Text103.Text & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Set ListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("barcode").Value))
    ListItem.SubItems(1) = Trim(vRecordset.Fields("code").Value)
    ListItem.SubItems(2) = Trim(vRecordset.Fields("name1").Value)
    ListItem.SubItems(3) = Trim(Text103.Text)
    ListItem.SubItems(4) = Trim(Text104.Text)
    ListItem.SubItems(5) = Trim(Text102.Text)
    ListItem.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
Else
    MsgBox "ไม่มี บาร์โค้ด รหัส " & vBarCode & " ในฐานข้อมูล กรุณาแก้ไขรหัสบาร์โค้ดด้วยนะครับ"
End If
vRecordset.Close

Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
Text104.Text = ""
Text101.SetFocus

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub GenTransfer()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vBarCode As String
Dim vTemp1, vTemp2, vTemp3, vTemp4, vTemp5, vTemp6, vTemp7, vTemp8 As String
Dim vTemp9, vTemp10, vTemp11, vTemp12, vTemp13, vTemp14 As String
Dim i As Integer, vLine As Integer
Dim vToWHCode As String
Dim vDocTransfer As String, vDocTransfer1 As String, vAverageCost As Integer
Dim vDescription As String
Dim vDepartment As String
Dim vCountQty As Integer
Dim vDocDate As String, vCreateDate As String
 
On Error GoTo ErrDescription
vQuery = "select header,autonumber from npmaster.dbo.NP_Generate_DocNo where headertype = 3"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocTransfer = Trim(vRecordset.Fields("header").Value) & Format(vRecordset.Fields("autonumber").Value, "000000")
End If
vRecordset.Close
vDescription = Trim(Text105.Text)
vDepartment = Left(Cmb101.Text, InStr(1, Cmb101.Text, "-") - 1)
vDocDate = Format(Date, "mm/dd/yyyy")
vCountQty = 0
vCreateDate = Format(Date, "mm/dd/yyyy") & " " & Time
vLine = 0

vQuery = "insert into bcnp.dbo.bcstktransfer2 (docno,isconfirm,docdate,departcode,mydescription,sumofqty,billstatus,iscancel,iscompletesave,creatorcode,createdatetime) " _
        & " values('" & vDocTransfer & "',0,'" & vDocDate & "','" & vDepartment & "','" & vDescription & "'," & vCountQty & ",0,0,1,'somrod','" & vCreateDate & "')"
gConnection.Execute vQuery

For i = 1 To ListView101.ListItems.Count
    If ListView101.ListItems.Item(i).Checked = True Then
    
        vQuery = "select averagecost from bcnp.dbo.bcitem  where code = '" & ListView101.ListItems.Item(i).SubItems(1) & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vAverageCost = Trim(vRecordset.Fields("averagecost").Value)
        End If
        vRecordset.Close
        
        vTemp1 = Trim(vDocTransfer)
        vTemp2 = Trim(ListView101.ListItems(i).SubItems(1))
        vTemp3 = vDocDate
        vTemp4 = Trim(ListView101.ListItems(i).SubItems(3))
        vTemp5 = "-"
        vTemp6 = Trim(ListView101.ListItems(i).SubItems(4))
        vTemp7 = "-"
        vTemp8 = Trim(ListView101.ListItems.Item(i).SubItems(5))
        vTemp9 = Trim(ListView101.ListItems.Item(i).SubItems(5)) * vAverageCost
        vTemp10 = Trim(ListView101.ListItems.Item(i).SubItems(6))
        vCountQty = vCountQty + vTemp8
        vTemp11 = vLine

        vQuery = "insert into bcnp.dbo.bcstktransfsub2 (docno,itemcode,docdate,fromwh,fromshelf,towh,toshelf,qty,sumofcost,unitcode,linenumber) " _
                        & " values('" & vTemp1 & "','" & vTemp2 & "','" & vTemp3 & "','" & vTemp4 & "','" & vTemp5 & "','" & vTemp6 & "','" & vTemp7 & "'," & vTemp8 & "," & vTemp9 & ",'" & vTemp10 & "'," & vTemp11 & ")"
        gConnection.Execute vQuery
    End If
    vLine = vLine + 1
Next i

Text106.Text = vDocTransfer
vQuery = "Update bcnp.dbo.bcstktransfer2 set sumofqty = " & vCountQty & " where docno = '" & vDocTransfer & "' "
gConnection.Execute vQuery

MsgBox "ทำใบขอโอนเรียบร้อยแล้ว ได้เลขที่ใบขอโอน " & vDocTransfer & " "
vQuery = "Update npmaster.dbo.NP_Generate_DocNo " _
                            & " set autonumber = autonumber +1 " _
                            & " where headertype = 3 "
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Public Sub Department()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

vQuery = "select  code+'-'+name as name from bcnp.dbo.BCDepartment "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Cmb101.AddItem Trim(vRecordset.Fields("name").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub

Public Sub InsertToGrid_UnConnect()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListItem As ListItem
Dim vBarCode

On Error GoTo ErrDescription

    vBarCode = Trim(Text101.Text)
    Set ListItem = ListView101.ListItems.Add(, , Trim(Text101.Text))
    ListItem.SubItems(3) = Trim(Text103.Text)
    ListItem.SubItems(4) = Trim(Text104.Text)
    ListItem.SubItems(5) = Trim(Text102.Text)

    
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text104.Text = ""
    Text101.SetFocus

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub TransferDocuments()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim DocListItem As ListItem

vQuery = "select  docno from bcstktransfer2 where docno like 'mt%' order by docno"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Cmb102.AddItem Trim(vRecordset.Fields("docno").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub


Public Sub GetTransferNo()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim ListItem As ListItem

vDocNo = Trim(Cmb102.Text)
vQuery = "select itemcode,b.name1,fromwh,towh,qty,unitcode from bcstktransfsub2 a left join bcitem b on a.itemcode = b.code " _
                & " where a.docno = '" & vDocNo & "' order by linenumber"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set ListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
    ListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
    If IsNull(vRecordset.Fields("name1")) Then
    ListItem.SubItems(2) = ""
    Else
    ListItem.SubItems(2) = Trim(vRecordset.Fields("name1").Value)
    End If
    ListItem.SubItems(3) = Trim(vRecordset.Fields("fromwh").Value)
    ListItem.SubItems(4) = Trim(vRecordset.Fields("towh").Value)
    ListItem.SubItems(5) = Trim(vRecordset.Fields("qty").Value)
    ListItem.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub


Public Sub PrintTransferDoc()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vReportName As String
Dim vTransferNo As String

On Error Resume Next

vTransferNo = Trim(Text106.Text)
vQuery = "select reportname from bcnp.dbo.bcreportname where repid = '167' and reptype = 'TF' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

    With Crystal101
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@DocNo;" & vTransferNo & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    End With
End Sub
