VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form101 
   Caption         =   "เช็คบาร์โค้ดพิมพ์ป้ายราคา"
   ClientHeight    =   7980
   ClientLeft      =   1635
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form101.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   11850
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   8100
      Top             =   6840
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
   Begin VB.CheckBox Check102 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "พิมพ์ป้ายราคาสินค้าโชว์"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6150
      TabIndex        =   13
      Top             =   1800
      Width           =   2340
   End
   Begin VB.CheckBox Check101 
      BackColor       =   &H8000000E&
      Caption         =   "กำหนดจำนวน = 1"
      Height          =   240
      Left            =   1425
      TabIndex        =   1
      Top             =   900
      Width           =   1665
   End
   Begin VB.CommandButton Cmd105 
      Caption         =   "ลบข้อมูล"
      Height          =   390
      Left            =   5625
      TabIndex        =   8
      Top             =   6150
      Width           =   1665
   End
   Begin VB.ComboBox Cmb101 
      Height          =   315
      Left            =   4275
      TabIndex        =   2
      Text            =   "014"
      Top             =   1275
      Width           =   1365
   End
   Begin VB.CommandButton Cmd104 
      Caption         =   "UpDate Data"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4275
      TabIndex        =   4
      Top             =   1725
      Width           =   1365
   End
   Begin VB.CommandButton Cmd103 
      Caption         =   "พิมพ์ป้ายราคา"
      Height          =   390
      Left            =   1425
      TabIndex        =   9
      Top             =   6750
      Width           =   1665
   End
   Begin VB.CommandButton Cmd102 
      Caption         =   "ลบรายการ"
      Height          =   390
      Left            =   3525
      TabIndex        =   7
      Top             =   6150
      Width           =   1665
   End
   Begin VB.CommandButton Cmd101 
      Caption         =   "แก้ไขจำนวน"
      Height          =   390
      Left            =   1425
      TabIndex        =   6
      Top             =   6150
      Width           =   1665
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1425
      TabIndex        =   3
      Top             =   1725
      Width           =   2190
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1425
      MaxLength       =   13
      TabIndex        =   0
      Top             =   1275
      Width           =   2190
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3840
      Left            =   1440
      TabIndex        =   5
      Top             =   2160
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   6773
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
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อสินค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "ราคาสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "ราคาสินค้าปกติ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "จำนวนที่พิมพ์"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "หน่วยนับ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "คลัง"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   3900
      TabIndex        =   12
      Top             =   1275
      Width           =   315
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวนที่พิมพ์"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   375
      TabIndex        =   11
      Top             =   1725
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "บาร์โค้ด"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   375
      TabIndex        =   10
      Top             =   1275
      Width           =   990
   End
End
Attribute VB_Name = "Form101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCheck As Integer

Private Sub Check1_Click()
Text101.SetFocus
End Sub

Private Sub Check101_Click()
Text101.SetFocus
End Sub

Private Sub Cmd101_Click()
On Error Resume Next
'แก้ไขจำนวนที่ต้องการพิมพ์
    ListView101.ListItems(vCheck).ListSubItems(4).Text = Trim(Text102.Text)
    Text101.Text = ""
    vCheck = 0
    Text102.Text = ""
    Text101.SetFocus
End Sub

Private Sub CMD102_Click()
On Error Resume Next
'ลบรายการที่ยิงไว้
ListView101.ListItems.Remove (vCheck)
vCheck = 0
Text101.Text = ""
Text102.Text = ""
Text101.SetFocus
End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim i As Integer
Dim vItemCode As String
Dim vUnitCode As String

On Error Resume Next
'พิมพ์ป้ายราคา
'Update ChagePrice
If ListView101.ListItems.Count <> 0 Then
    For i = 1 To ListView101.ListItems.Count
        vItemCode = Trim(ListView101.ListItems.Item(i).SubItems(6))
        vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(5))
        vQuery = "exec usp_IV_UpdatePrintUpdateChangePrice '" & vItemCode & "','" & vUnitCode & "' "
        gConnection.Execute vQuery
    Next i
    Call InsertToReport_Temp
    vQuery = "select * from dbo.NP_LABEL_SETUP where id = 76 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("pathname").Value)
    End If
    vRecordset.Close
    If vReportName <> "" Then
            With Crystal101
            .ReportFileName = Trim(vReportName) & ".rpt"
            .ParameterFields(0) = "@vUserID;" & vUserID & ";true"
            .WindowState = crptMaximized
            .Destination = crptToWindow
            .Action = 1
            End With
    Else
            MsgBox "คุณยังไม่ได้เลือกฟอร์มป้ายที่จะพิมพ์  ให้คลิ๊กในชื่อฟอร์มที่จะพิมพ์"
    End If
    
    vQuery = "exec dbo.USP_NP_DeleteDataPrintLabel '" & vUserID & "' "
    gConnection.Execute vQuery
End If
End Sub

Private Sub CMD104_Click()
Dim i As Integer
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription
'กดปุ่มนี้ เพื่อทำการ Link ข้อมูลที่ยิงมากับฐานข้อมูลว่ามีข้อมูลจริงไหม
If vConnect = 1 Then
Call InitializeDatabase
Call WHCode
End If
For i = 1 To ListView101.ListItems.Count
vQuery = "select * from bcnp.dbo.vw_MB_CheckPrintLabel where barcode = '" & ListView101.ListItems.Item(i).Text & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    ListView101.ListItems(i).ListSubItems(1).Text = Trim(vRecordset.Fields("name1").Value)
    ListView101.ListItems(i).ListSubItems(2).Text = Trim(vRecordset.Fields("saleprice1").Value)
    ListView101.ListItems(i).ListSubItems(3).Text = Trim(vRecordset.Fields("priceerect").Value)
     
Else
    ListView101.ListItems(i).ListSubItems(1).Text = Trim("ไม่มีรหัสสินค้า ")
    MsgBox "ไม่มี บาร์โค้ด รหัส " & ListView101.ListItems.Item(i).Text & " ในฐานข้อมูล กรุณาแก้ไขรหัสบาร์โค้ดด้วยนะครับ"
End If
vRecordset.Close
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub CMD105_Click()
'ลบข้อมูลที่ยิงได้ทั้งหมด
ListView101.ListItems.Clear
ListView101.SetFocus
End Sub

Private Sub Form_Load()
'ดึงข้อมูลคลังสินค้า
If vConnect = 0 Then
    Call WHCode
Else
    CMD104.Enabled = True
End If
End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
vCheck = Item.Index
Text101 = ListView101.ListItems(vCheck).Text
Text102 = ListView101.ListItems(vCheck).SubItems(4)

End Sub



Private Sub Text101_KeyPress(KeyAscii As Integer)
On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Check101.Value = 0 Then
        Text102.SetFocus
    Else
        Call Text102_KeyPress(KeyAscii)
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
    If vCheck = 0 Then
        If vConnect = 0 Then
            Call InsertToGrid
        Else
            Call InsertToGrid_UnConnect
        End If
    Else
        MsgBox "กรุณา กดปุ่ม แก้ไขจำนวน หรือ ปุ่มลบรายการ"
    End If
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
Dim vNum As String

On Error GoTo ErrDescription

vBarCode = Trim(Text101.Text)
If Check101.Value = 1 Then
    vNum = 1
Else
    vNum = Text102.Text
End If
vQuery = "select * from bcnp.dbo.vw_MB_CheckPrintLabel where barcode = '" & vBarCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Set ListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("barcode").Value))
    ListItem.SubItems(1) = Trim(vRecordset.Fields("name1").Value)
    ListItem.SubItems(2) = Trim(vRecordset.Fields("saleprice1").Value)
    ListItem.SubItems(3) = Trim(vRecordset.Fields("priceerect").Value)
    ListItem.SubItems(4) = Trim(vNum)
    ListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
    ListItem.SubItems(6) = Trim(vRecordset.Fields("itemcode").Value)
Else
    MsgBox "ไม่มี บาร์โค้ด รหัส " & vBarCode & " ในฐานข้อมูล กรุณาแก้ไขรหัสบาร์โค้ดด้วยนะครับ"
End If
vRecordset.Close

Text101.Text = ""
Text102.Text = ""
Text101.SetFocus

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub InsertToGrid_UnConnect()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListItem As ListItem
Dim vBarCode
Dim vNum As String

On Error GoTo ErrDescription

vBarCode = Trim(Text101.Text)
If Check101.Value = 1 Then
    vNum = 1
Else
    vNum = Text102.Text
End If
    Set ListItem = ListView101.ListItems.Add(, , Trim(Text101.Text))
    ListItem.SubItems(4) = Trim(vNum)
Text101.Text = ""
Text102.Text = ""
Text101.SetFocus

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub InsertToReport_Temp()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer, j As Integer, vQty As Integer
Dim vPrint As Integer
Dim tmpBarCode, tmpItemCode, tmpItemName, tmpWHCode, tmpShelfCode As String
Dim tmpUnitCode, tmpCurQTY, tmpPrice, tmpPriceErect, tmpPrintQTY, tmpUser As String
Dim vBarCode As String
Dim vWHCode As String
Dim vUnitCode As String

On Error GoTo ErrDescription

For i = 1 To ListView101.ListItems.Count
        vPrint = ListView101.ListItems(i).ListSubItems(4).Text
        vBarCode = ListView101.ListItems(i).Text
        vUnitCode = ListView101.ListItems(i).ListSubItems(5).Text
        If Check102.Value <> 1 Then
            vQuery = "exec dbo.USP_IV_ProgStockChecking_PrintLabel '" & vBarCode & "','" & vUnitCode & "' "
        Else
                vQuery = "exec dbo.USP_IV_ProgStockChecking1 '" & vBarCode & "','" & vUnitCode & "' "
        End If
        
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                tmpBarCode = Trim(vRecordset.Fields("barcode").Value)
                tmpItemCode = Trim(vRecordset.Fields("code").Value)
                tmpItemName = Trim(vRecordset.Fields("name1").Value)
                tmpWHCode = Trim(vRecordset.Fields("whcode").Value)
                tmpShelfCode = Trim(vRecordset.Fields("shelfcode1").Value)
                tmpUnitCode = Trim(vRecordset.Fields("unitcode").Value)
                tmpCurQTY = Trim(vRecordset.Fields("qty").Value)
                tmpPrice = Trim(vRecordset.Fields("saleprice1").Value)
                If IsNull(vRecordset.Fields("priceerect").Value) Then
                    tmpPriceErect = 0
                Else
                    tmpPriceErect = CheckDegit(Trim(vRecordset.Fields("priceerect").Value))
                End If
                tmpPrintQTY = 1
                tmpUser = vUserID
                For j = 1 To vPrint
                vQuery = "exec dbo.USP_MB_InsertPrintLabelTemp '" & tmpItemCode & "','" & tmpBarCode & "','" & tmpItemName & "'," & tmpPrintQTY & ", " _
                                & " " & tmpPrice & "," & tmpPriceErect & ",'" & tmpUnitCode & "','" & tmpUser & "','" & tmpWHCode & "','" & tmpShelfCode & "'," & tmpCurQTY & " "
                gConnection.Execute vQuery
        Next j
        Else
        MsgBox "สินค้ารหัสบาร์โค้ด " & ListView101.ListItems(i).Text & " ไม่มีในคลัง " & Trim(Cmb101.Text) & " กรุณาตรวจสอบด้วยนะครับ เพราะจะไม่สามารถพิมพ์ป้ายราคาได้"
        End If
        vRecordset.Close
        
Next i

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
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
