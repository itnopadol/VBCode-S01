VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form103 
   Caption         =   "เช็คบาร์โค้ดตรวจสอบราคา"
   ClientHeight    =   7995
   ClientLeft      =   2265
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form103.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   10965
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1035
      Top             =   6345
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
   Begin VB.CommandButton Cmd105 
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
      Height          =   390
      Left            =   9300
      TabIndex        =   13
      Top             =   6075
      Width           =   1590
   End
   Begin VB.CheckBox CHK101 
      BackColor       =   &H80000009&
      Caption         =   "ไม่ยิงราคา"
      Height          =   240
      Left            =   5625
      TabIndex        =   12
      Top             =   525
      Width           =   1065
   End
   Begin MSComctlLib.ListView ListView102 
      Height          =   1740
      Left            =   5625
      TabIndex        =   2
      Top             =   900
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   3069
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
         Text            =   "ประเภทการขาย"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ประเภทการขนส่ง"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ราคา"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "หน่วยขาย"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.CommandButton Cmd104 
      Caption         =   "ลบรายการใน Grid"
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
      Left            =   6975
      TabIndex        =   7
      Top             =   6075
      Width           =   1590
   End
   Begin VB.ComboBox Cmb101 
      Height          =   315
      Left            =   4575
      TabIndex        =   1
      Top             =   900
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton Cmd103 
      Caption         =   "พิมพ์ป้ายราคา"
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
      Left            =   2250
      TabIndex        =   8
      Top             =   6600
      Width           =   1590
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
      Left            =   4575
      TabIndex        =   6
      Top             =   6075
      Width           =   1590
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
      Left            =   2250
      TabIndex        =   5
      Top             =   6075
      Width           =   1590
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3165
      Left            =   2250
      TabIndex        =   4
      Top             =   2775
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   5583
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อสินค้า"
         Object.Width           =   4498
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "ราคาที่ยิงได้"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "ราคาจริง"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "ผลต่างของราคา"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "หน่วยนับ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2325
      TabIndex        =   3
      Top             =   1350
      Width           =   1740
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2325
      TabIndex        =   0
      Top             =   900
      Width           =   1740
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "คลัง"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   4200
      TabIndex        =   11
      Top             =   900
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ราคา"
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
      Left            =   1350
      TabIndex        =   10
      Top             =   1350
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสบาร์โค้ด"
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
      Left            =   1275
      TabIndex        =   9
      Top             =   900
      Width           =   915
   End
End
Attribute VB_Name = "Form103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCheck As Integer


Public Sub CheckPriceLevel()
Dim vBarCode As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListItem As ListItem

On Error GoTo ErrDescription
vBarCode = Trim(Text101.Text)
            ListView102.ListItems.Clear
            vQuery = "execute bcnp.dbo.usp_IV_PrgCheckPrice '" & vBarCode & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set ListItem = ListView102.ListItems.Add(, , Trim(vRecordset.Fields("saletype").Value))
                ListItem.SubItems(1) = Trim(vRecordset.Fields("transporttype").Value)
                ListItem.SubItems(2) = Trim(vRecordset.Fields("saleprice1").Value)
                ListItem.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
                vRecordset.MoveNext
                Wend
            Else
                ListView102.ListItems.Clear
                Text101.Text = ""
                Text101.SetFocus
            End If
            vRecordset.Close
            
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CHK101_Click()
Text101.SetFocus
End Sub

Private Sub Cmd101_Click()
On Error Resume Next
    ListView101.ListItems(vCheck).ListSubItems(2).Text = Trim(Text102.Text)
    ListView101.ListItems(vCheck).ListSubItems(3).Text = Trim(Text102.Text) - ListView101.ListItems(vCheck).ListSubItems(1).Text
    Text101.Text = ""
    Text102.Text = ""
    vCheck = 0
    Text101.SetFocus
End Sub

Private Sub CMD102_Click()
On Error Resume Next
ListView101.ListItems.Remove (vCheck)
Text101.Text = ""
Text102.Text = ""
ListView102.ListItems.Clear
vCheck = 0
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
    vQuery = "select * into bcnp.dbo.Report_Temp from bcnp.dbo.np_label_temp where useduser = '' "
    gConnection.Execute vQuery
    Call InsertToReport_Temp
    vReportName = Trim("V:\Reports\LABEL_BC\PLNMA439")
    If vReportName <> "" Then
            With Crystal101
            .ReportFileName = Trim(vReportName) & ".rpt"
            .WindowState = crptMaximized
            .Destination = crptToWindow
            .Action = 1
            End With
    Else
            MsgBox "คุณยังไม่ได้เลือกฟอร์มป้ายที่จะพิมพ์  ให้คลิ๊กในชื่อฟอร์มที่จะพิมพ์"
    End If
    
    vQuery = "drop table bcnp.dbo.Report_Temp"
    gConnection.Execute vQuery
    vQuery = "delete bcnp.dbo.np_label_temp where useduser = 'somrod'"
    gConnection.Execute vQuery
End If
End Sub

Private Sub CMD104_Click()
ListView101.ListItems.Clear
ListView102.ListItems.Clear
Text101.Text = ""
Text102.Text = ""
Text101.SetFocus
End Sub

Private Sub CMD105_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vItemCode As String
Dim vBarCode As String
Dim vItemName As String
Dim vUnitCode As String
Dim vLabelPrice As Currency
Dim vSalePrice1 As Currency
Dim vDiffOfPrice As Currency
Dim vRunNumber As Integer
Dim vDocNo As String
Dim vFormatDocNo As String
Dim i As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count > 0 Then
    vQuery = "select      isnull(cast(right(max(labid),4) as int),0) as number  from    npmaster.dbo.TB_CK_LabelPrice"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRunNumber = Trim(vRecordset.Fields("Number").Value)
    End If
    vRecordset.Close
    vFormatDocNo = Format(vRunNumber + 1, "0000")
    vDocNo = "Mobile-" & vFormatDocNo
    For i = 1 To ListView101.ListItems.Count
        vItemCode = Trim(ListView101.ListItems.Item(i).Text)
        vBarCode = Trim(ListView101.ListItems.Item(i).Text)
        vItemName = Trim(ListView101.ListItems.Item(i).SubItems(1))
        vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(5))
        vLabelPrice = Trim(ListView101.ListItems.Item(i).SubItems(2))
        vSalePrice1 = Trim(ListView101.ListItems.Item(i).SubItems(3))
        vDiffOfPrice = Trim(ListView101.ListItems.Item(i).SubItems(4))
        vQuery = "exec bcnp.dbo.usp_CK_InsertDataCheckPrice '" & vDocNo & "','" & vItemCode & "','" & vBarCode & "','" & vItemName & "','" & vUnitCode & "'," & vLabelPrice & "," & vSalePrice1 & "," & vDiffOfPrice & ",'" & vUserID & "'"
        gConnection.Execute vQuery
    Next i
    ListView101.ListItems.Clear
    ListView102.ListItems.Clear
    MsgBox "บันทึกข้อมูลการตรวจสอบราคาสินค้า ได้เอกสารเลขที่ " & vDocNo & " "
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
'Call WHCode
End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ErrDescription

vCheck = Item.Index
Text101 = ListView101.ListItems(vCheck).Text
Text102 = ListView101.ListItems(vCheck).SubItems(2)
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CheckPriceLevel
    If CHK101.Value = 1 Then
        Call InsertToGrid
        Text101.SetFocus
    Else
        Text102.SetFocus
    End If

End If
End Sub

'Public Sub WHCode()
'Dim vRecordset  As New ADODB.Recordset
'Dim vQuery As String
'
'On Error Resume Next
'
'vQuery = "select distinct whcode from bcstklocation  where whcode <> '' and  whcode is not null and whcode <> '-'  order by whcode"
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vRecordset.MoveFirst
  '  While Not vRecordset.EOF
   '     Cmb101.AddItem Trim(vRecordset.Fields("whcode").Value)
    '    vRecordset.MoveNext
    'Wend
'End If
'vRecordset.Close

'End Sub


Public Sub InsertToGrid()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListItem As ListItem
Dim vBarCode

On Error GoTo ErrDescription

vBarCode = Trim(Text101.Text)
If CHK101.Value = 1 Then
    Text102.Text = 0
End If
vQuery = "select  * from bcnp.dbo.vw_CK_LabelPrice where barcode = '" & vBarCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Set ListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("barcode").Value))
    ListItem.SubItems(1) = Trim(vRecordset.Fields("name1").Value)
    ListItem.SubItems(2) = Trim(Text102.Text)
    ListItem.SubItems(3) = Trim(vRecordset.Fields("saleprice1").Value)
    ListItem.SubItems(4) = Trim(Text102.Text) - Trim(vRecordset.Fields("saleprice1").Value)
    ListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
    ListItem.SubItems(6) = Trim(vRecordset.Fields("code").Value)
    If Trim(Text102.Text) - Trim(vRecordset.Fields("saleprice1").Value) <> 0 Then
    ListView101.ListItems.Item(ListView101.ListItems.Count).Checked = True
    End If
Else
    MsgBox "ไม่มี บาร์โค้ด รหัส " & vBarCode & " ที่เป็นสินค้าคลัง 014  กรุณาแก้ไขรหัสบาร์โค้ดด้วยนะครับ"
End If
vRecordset.Close

Text101.Text = ""
Text102.Text = ""
ListView102.ListItems.Clear
Text101.SetFocus

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Private Sub Text102_KeyPress(KeyAscii As Integer)

On Error GoTo ErrDescription

If KeyAscii = 13 Then '1 เช็ค การกด ตัว Enter
    If vCheck = 0 Then
    If vConnect = 0 Then '2 เช็คว่าโปรแกรมทำการติดต่อ Connection หรือไม่
            Call InsertToGrid ' เอาข้อมูลสินค้าเข้า Grid
    End If '2
    Else
        MsgBox "กรุณา กดปุ่มแก้ไขรายการ หรือ ปุ่มลบรายการ"
    End If
End If '1

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

On Error Resume Next

For i = 1 To ListView101.ListItems.Count
    If ListView101.ListItems.Item(i).Checked = True Then
        vPrint = 1
        vQuery = "select * from bcnp.dbo.vw_IV_PrgCheckPriceList " _
                        & " where barcode = '" & ListView101.ListItems(i).Text & "' and whcode = '014' "
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
                    tmpPriceErect = Trim(vRecordset.Fields("priceerect").Value)
                End If
                tmpPrintQTY = 1
                tmpUser = "somrod"
                For j = 1 To vPrint
                vQuery = "insert  into     bcnp.dbo.Report_Temp " _
                                & " (ItemCode,BarCode,Name1,QTY,Price, " _
                                & " SPrice,UnitCode,UsedUser,Category_ID,WHCode,ShelfCode,ONHand) " _
                                & " values  ('" & tmpItemCode & "','" & tmpBarCode & "','" & tmpItemName & "'," & tmpPrintQTY & ", " _
                                & " " & tmpPrice & "," & tmpPriceErect & ",'" & tmpUnitCode & "','" & tmpUser & "','','" & tmpWHCode & "','" & tmpShelfCode & "'," & tmpCurQTY & ")"
                gConnection.Execute vQuery
        Next j
        Else
        MsgBox "สินค้ารหัสบาร์โค้ด " & ListView101.ListItems(i).Text & " ไม่มีข้อมูลคลังที่เก็บที่เป็นคลัง  014  กรุณาตรวจสอบด้วยนะครับ เพราะจะไม่สามารถพิมพ์ป้ายราคาได้"
        End If
        vRecordset.Close
    End If
Next i
End Sub

