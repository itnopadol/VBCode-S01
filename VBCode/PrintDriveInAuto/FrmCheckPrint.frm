VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form FrmCheckPrint 
   BackColor       =   &H00FF0000&
   Caption         =   "พิมพ์เอกสารอัตโนมัติ"
   ClientHeight    =   7395
   ClientLeft      =   8490
   ClientTop       =   2070
   ClientWidth     =   5595
   Icon            =   "FrmCheckPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   5595
   Begin Crystal.CrystalReport CrystalDT 
      Left            =   495
      Top             =   6750
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   1725
      Left            =   45
      TabIndex        =   4
      Top             =   4545
      Visible         =   0   'False
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   3043
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "JobID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เอกสาร"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "อ้างถึง"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "คลัง"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ชั้นเก็บ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Family"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ZoneID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "PickZone"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "SendTime"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "PrinterName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "IsPrinted"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "PrintBy"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "SendTime"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton CMDExpand 
      Caption         =   "ย่อ/ขยาย"
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
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   5505
   End
   Begin MSComctlLib.ListView ListViewPrintList 
      Height          =   2985
      Left            =   45
      TabIndex        =   1
      Top             =   1440
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   5265
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เอกสาร"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "อ้างถึง"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Family"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ZoneID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "PickZone"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Printer"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "พิมพ์เมื่อ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ชื่อผู้พิมพ์"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ส่งพิมพ์เมื่อ"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Timer TimerStart 
      Enabled         =   0   'False
      Left            =   2295
      Top             =   1530
   End
   Begin VB.Timer TimerNow 
      Enabled         =   0   'False
      Interval        =   65507
      Left            =   1845
      Top             =   1530
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3703
      Left            =   945
      Top             =   1530
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3303
      Left            =   495
      Top             =   1530
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3003
      Left            =   45
      Top             =   1530
   End
   Begin VB.Timer TimerCheckDocument 
      Interval        =   30000
      Left            =   1395
      Top             =   1530
   End
   Begin VB.PictureBox Crystal102 
      Height          =   480
      Left            =   495
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   2205
      Width           =   1200
   End
   Begin VB.PictureBox Crystal103 
      Height          =   480
      Left            =   945
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   2205
      Width           =   1200
   End
   Begin VB.PictureBox Crystal101 
      Height          =   480
      Left            =   45
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   2205
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการ พิมพ์เอกสารอัตโนมัติ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   45
      TabIndex        =   3
      Top             =   1125
      Width           =   5505
   End
   Begin VB.Label LBLStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "กำลังทำงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   5505
   End
   Begin VB.Image IM2 
      Height          =   480
      Left            =   5220
      Picture         =   "FrmCheckPrint.frx":1272
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IM3 
      Height          =   480
      Left            =   5220
      Picture         =   "FrmCheckPrint.frx":24E4
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IM1 
      Height          =   480
      Left            =   5220
      Picture         =   "FrmCheckPrint.frx":3756
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "FrmCheckPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDocno As String
Dim i As Integer

Private Sub CMDExpand_Click()

If Me.Height = 7200 Then
  Me.Height = 1600
Else
  Me.Height = 7200
End If

End Sub

Private Sub Form_Load()
Call InitializeDataBase
Me.ListViewPrintList.ListItems.Clear
End Sub

Private Sub Timer1_Timer()
Me.IM1.Visible = True
Me.IM2.Visible = False
Me.IM3.Visible = False
End Sub

Private Sub Timer2_Timer()
Me.IM1.Visible = False
Me.IM2.Visible = True
Me.IM3.Visible = False
End Sub

Private Sub Timer3_Timer()
Me.IM1.Visible = False
Me.IM2.Visible = False
Me.IM3.Visible = True
End Sub

Private Sub TimerCheckDocument_Timer()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItem As ListItem
Dim vListItemDoc As ListItem
Dim i As Integer
Dim vCount As Integer
Dim vJobID As Integer
Dim vIsPrinted As Integer

Dim vGetJobID As Integer
Dim vGetDocNo As String
Dim vGetRefNo As String
Dim vGetWHCode As String
Dim vGetShelfCode As String
Dim vGetZoneID As String
Dim vGetFamilyCode As String
Dim vGetPickZone As String
Dim vGetPrinterName As String
Dim vGetUserPrint As String
Dim vGetSendTime As String

''On Error Resume Next
    
vQuery = "exec dbo.USP_NP_SearchPrintTermal"
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  While Not vRecordset.EOF
      Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("jobid").Value))
      vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("refno").Value)
      vListItem.SubItems(3) = Trim(vRecordset.Fields("whcode").Value)
      vListItem.SubItems(4) = Trim(vRecordset.Fields("shelfcode").Value)
      vListItem.SubItems(5) = Trim(vRecordset.Fields("familycode").Value)
      vListItem.SubItems(6) = Trim(vRecordset.Fields("zoneid").Value)
      vListItem.SubItems(7) = Trim(vRecordset.Fields("pickzone").Value)
      vListItem.SubItems(8) = Trim(vRecordset.Fields("sendtime").Value)
      vListItem.SubItems(9) = Trim(vRecordset.Fields("printername").Value)
      vListItem.SubItems(10) = Trim(vRecordset.Fields("isprint").Value)
      vListItem.SubItems(11) = Trim(vRecordset.Fields("userprint").Value)
      vListItem.SubItems(12) = Trim(vRecordset.Fields("sendtime").Value)
      
      i = i + 1
  vRecordset.MoveNext
  Wend
End If
vRecordset.Close


For vCount = 1 To ListView101.ListItems.Count
    vGetJobID = Trim(ListView101.ListItems.Item(vCount).Text)
    vGetDocNo = Trim(ListView101.ListItems.Item(vCount).SubItems(1))
    vGetRefNo = Trim(ListView101.ListItems.Item(vCount).SubItems(2))
    vGetWHCode = Trim(ListView101.ListItems.Item(vCount).SubItems(3))
    vGetShelfCode = Trim(ListView101.ListItems.Item(vCount).SubItems(4))
    vGetFamilyCode = Trim(ListView101.ListItems.Item(vCount).SubItems(5))
    vGetZoneID = Trim(ListView101.ListItems.Item(vCount).SubItems(6))
    vGetPickZone = Trim(ListView101.ListItems.Item(vCount).SubItems(7))
    vGetPrinterName = Trim(ListView101.ListItems.Item(vCount).SubItems(9))
    vIsPrinted = Trim(ListView101.ListItems.Item(vCount).SubItems(10))
    vGetUserPrint = Trim(ListView101.ListItems.Item(vCount).SubItems(11))
    vGetSendTime = Trim(ListView101.ListItems.Item(vCount).SubItems(12))
    
    
    If vGetDocNo <> "" And vGetJobID = 1 And vMemIsPrintError = 0 Then
    
      vMemIsPrintError = 1
    
      vQuery = "exec dbo.USP_NP_UpdatePrintTermal " & vGetJobID & ",'" & vGetDocNo & "','" & vGetWHCode & "','" & vGetShelfCode & "','" & vGetFamilyCode & "','" & vGetZoneID & "','" & vGetPickZone & "' "
      vConnection.Execute (vQuery)

      Call PrintPickingSlipFromSaleOrder(vGetDocNo, vGetRefNo, vGetWHCode, vGetShelfCode, vGetZoneID, vGetFamilyCode, vGetPickZone, vCount, vGetPrinterName)
      
  ElseIf vGetDocNo <> "" And vGetJobID = 2 And vMemIsPrintError = 0 Then

      vMemIsPrintError = 1

      vQuery = "exec dbo.USP_NP_UpdatePrintTermal " & vGetJobID & ",'" & vGetDocNo & "','" & vGetWHCode & "','" & vGetShelfCode & "','" & vGetFamilyCode & "','" & vGetZoneID & "','" & vGetPickZone & "' "
      vConnection.Execute (vQuery)
    
      Call PrintPickingSlipFromRequest(vGetDocNo, vGetZoneID, vGetFamilyCode, vGetPickZone, vGetPrinterName)
      
      If vGetWHCode = "S1-B" Then
        Call PrintPickingSlipFromRequestDriveThru(vGetDocNo, vGetZoneID, vGetFamilyCode, vGetPickZone)
      End If
      
  ElseIf vGetDocNo <> "" And vGetJobID = 4 And vMemIsPrintError = 0 Then

      vMemIsPrintError = 1

      vQuery = "exec dbo.USP_NP_UpdatePrintTermal " & vGetJobID & ",'" & vGetDocNo & "','" & vGetWHCode & "','" & vGetShelfCode & "','" & vGetFamilyCode & "','" & vGetZoneID & "','" & vGetPickZone & "' "
      vConnection.Execute (vQuery)
    
      Call PrintPickingSlipFromRequestDriveThru(vGetDocNo, vGetZoneID, vGetFamilyCode, vGetPickZone)
      
ElseIf vGetDocNo <> "" And vGetJobID = 3 And vMemIsPrintError = 0 Then

      vMemIsPrintError = 1

      vQuery = "exec dbo.USP_NP_UpdatePrintTermal " & vGetJobID & ",'" & vGetDocNo & "','" & vGetWHCode & "','" & vGetShelfCode & "','" & vGetFamilyCode & "','" & vGetZoneID & "','" & vGetPickZone & "' "
      vConnection.Execute (vQuery)
      
      'vQuery = "exec dbo.USP_NP_SearchGroupPicking1 3,'" & vGetDocNo & "',''"
    
      Call PrintPickupSlipDriveThru(vGetDocNo, vGetZoneID, vGetFamilyCode, vGetPickZone, vGetPrinterName)
              
              
    End If
    
    
    If Me.ListViewPrintList.ListItems.Count = 0 Then
      i = 1
    Else
      i = Me.ListViewPrintList.ListItems.Count + 1
    End If
    
    
    Set vListItem = ListViewPrintList.ListItems.Add(, , i)
    vListItem.SubItems(1) = vGetDocNo
    vListItem.SubItems(2) = vGetRefNo
    vListItem.SubItems(3) = vGetFamilyCode
    vListItem.SubItems(4) = vGetZoneID
    vListItem.SubItems(5) = vGetPickZone
    vListItem.SubItems(6) = vGetPrinterName
    vListItem.SubItems(7) = Now
    vListItem.SubItems(8) = vGetUserPrint
    vListItem.SubItems(9) = vGetSendTime
    
Next vCount


Me.ListView101.ListItems.Clear
      
End Sub





Public Sub PrintPickingSlipFromSaleOrder(vSaleOrder As String, vQueueNo As String, vWHCode As String, vShelfGroup As String, vZoneID As String, vFamilyGroup As String, vPickZoneGroup As String, vCount As Integer, vPrinterName As String)
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vQuery As String
Dim vDocno As String
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer
Dim vItemName As String
Dim vSoStatus As Integer
Dim vSelectPicked As Integer
Dim vGroupDocNo As String
Dim vPrinterID As Integer

Dim strComputer As String


''On Error Resume Next


For Each printerObj In Printers
If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
Set Printer = printerObj
Set printerObj = Nothing
Exit For
End If
Next

'===========================================================================

'strComputer = "."
'Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
'Set colItems = objWMIService.ExecQuery("Select * from Win32_Printer ", , 48)
'For Each objItem In colItems

''MsgBox UCase(objItem.Name)

'If UCase(vPrinterName) = UCase(objItem.Name) And objItem.PrinterStatus = 2 Then
'
 '     vQuery = "exec dbo.USP_NP_UpdateCancelPrintTermal 1,'" & vSaleOrder & "','" & vWHCode & "','" & vShelfGroup & "','" & vFamilyGroup & "','" & vZoneID & "','" & vPickZoneGroup & "' "
  '    vConnection.Execute (vQuery)
   '   Exit Sub

'End If

'Next


'===========================================================================


vQuery = "exec dbo.USP_SO_PickingQueueFreedom3 '" & vSaleOrder & "','" & vQueueNo & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "','" & vFamilyGroup & "','" & vPickZoneGroup & "'," & vCount & " "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then

    vSoStatus = vRecordset.Fields("sostatus").Value
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 10
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print Trim("_______________________________________________________________________________________")


    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 12
    Printer.FontBold = True
    Printer.CurrentX = 0
    Printer.CurrentY = 200
    Printer.Print vRecordset.Fields("printmydesc").Value
    
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 18
    Printer.FontBold = True
    Printer.CurrentX = 2200
    Printer.CurrentY = 120
    Printer.Print Trim("คลัง  :   ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value)

    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 50
    Printer.FontBold = True
    Printer.CurrentX = 1550
    Printer.CurrentY = 0
    Printer.Print Trim(vRecordset.Fields("queueno").Value)
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 18
    Printer.FontBold = False
    Printer.CurrentX = 0
    Printer.CurrentY = 1000
    Printer.Print "ID." & Trim(vRecordset.Fields("runningno").Value)
    
    
    Printer.Font.Name = "3 of 9 Barcode"
    Printer.Font.Size = 40
    Printer.FontBold = False
    Printer.CurrentX = 1000
    Printer.CurrentY = 1000
    Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"

    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 10
    Printer.CurrentX = 0
    Printer.CurrentY = 1450
    Printer.Print Trim("_______________________________________________________________________________________")
        
    If Trim(vRecordset.Fields("iscopy").Value) = 0 Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 12
      Printer.CurrentX = 1400
      Printer.CurrentY = 1650
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า")
    Else
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 12
      Printer.CurrentX = 1400
      Printer.CurrentY = 1650
      Printer.Print Trim("ทดแทนใบจัดสินค้า")
    
    End If

    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = 1900
    Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 12
    Printer.CurrentX = 1500
    Printer.CurrentY = 1900
    Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = 2150
    Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 12
    Printer.CurrentX = 0
    Printer.CurrentY = 2400
    Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
          
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 12
    Printer.CurrentX = 0
    Printer.CurrentY = 2650
    Printer.Print Trim("ลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = 2900
    Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = 3150
    
    If vRecordset.Fields("isconditionsend").Value = 0 Then
          Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("รับเอง")
    Else
          Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("ส่งให้")
    End If
                

    If Trim(vRecordset.Fields("carlicense").Value) <> "" Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 16
      Printer.CurrentX = 1400
      Printer.CurrentY = 3400
      Printer.FontBold = True
      Printer.FontUnderline = True
      Printer.Print Trim("ทะเบียนรถขนส่ง : ") & Trim(vRecordset.Fields("carlicense").Value)
    End If
          
    If vSoStatus = 0 Then
        If Trim(vRecordset.Fields("carlicense").Value) <> "" Then
          Printer.Font.Name = "AngsanaUPC"
          Printer.Font.Size = 14
          Printer.FontBold = True
          Printer.FontUnderline = False
          Printer.CurrentX = 0
          Printer.CurrentY = 3800
          Printer.Print Trim("เวลารับของ : ") & Trim(vRecordset.Fields("requesttime").Value)
       Else
          Printer.Font.Name = "AngsanaUPC"
          Printer.Font.Size = 14
          Printer.FontBold = True
          Printer.FontUnderline = False
          Printer.CurrentX = 0
          Printer.CurrentY = 3400
          Printer.Print Trim("เวลารับของ : ") & Trim(vRecordset.Fields("requesttime").Value)
       End If
    ElseIf vSoStatus <> 0 Then
      If Trim(vRecordset.Fields("carlicense").Value) <> "" Then
          Printer.Font.Name = "AngsanaUPC"
          Printer.Font.Size = 14
          Printer.FontBold = True
          Printer.FontUnderline = False
          Printer.CurrentX = 0
          Printer.CurrentY = 3800
          Printer.Print Trim("วันที่ครบกำหนดรับของ : ") & Trim(vRecordset.Fields("duedate").Value)
       Else
          Printer.Font.Name = "AngsanaUPC"
          Printer.Font.Size = 14
          Printer.FontBold = True
          Printer.FontUnderline = False
          Printer.CurrentX = 0
          Printer.CurrentY = 3400
          Printer.Print Trim("วันที่ครบกำหนดรับของ : ") & Trim(vRecordset.Fields("duedate").Value)
       
       End If
    End If
          
    vRecordset.MoveFirst
    vLineX = 50
    vLineY = 50
    
    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 10
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 30
    Printer.FontBold = False
    Printer.FontUnderline = False
    Printer.Print Trim("-----------------------------------------------------------------------------------------------")
    n = 1
    While Not vRecordset.EOF
        Printer.Font.Size = 18
        Printer.FontBold = True
        
        If Len(Trim(vRecordset.Fields("shelfcode1").Value)) < 6 Then
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "#" & n & ". " & "ที่เก็บ : " & Trim(vRecordset.Fields("shelfcode1").Value)
        Else
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "#" & n & ". " & "ที่เก็บ : " & Left(Trim(vRecordset.Fields("shelfcode1").Value), 1) & " " & Mid(Trim(vRecordset.Fields("shelfcode1").Value), 2, 2) & " " & Right(Trim(vRecordset.Fields("shelfcode1").Value), 3)
        End If
           
         Printer.Font.Size = 11
         Printer.FontBold = False
                                    
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "รหัส : " & Trim(vRecordset.Fields("itemcode").Value) & "                                      " & "หน่วยนับหลัก : " & Trim(vRecordset.Fields("StkUnitCode").Value)
          
          
          vItemName = Trim(vRecordset.Fields("itemname").Value) & Trim(vRecordset.Fields("descriptionline"))
          If Len(vItemName) <= 55 Then
             Printer.CurrentX = 0
             Printer.CurrentY = Printer.CurrentY
             Printer.Print "ชื่อ :" & vItemName

          Else
             Printer.CurrentX = 0
             Printer.CurrentY = Printer.CurrentY
             Printer.Print "ชื่อ :" & Left(vItemName, 55)
             
             Printer.CurrentX = 0
             Printer.CurrentY = Printer.CurrentY
             Printer.Print Right(vItemName, Len(vItemName) - 55)
          End If
                    
          Printer.Font.Size = 11
          Printer.CurrentX = Printer.CurrentX
          Printer.CurrentY = Printer.CurrentY
          Printer.FontBold = True
          Printer.Print "Total/รวม" & "    |    " & "OnHand" & "    |     " & "Order/ต้องการ" & "      |     " & "Pickup/จัดได้"
          
          Printer.Font.Size = 12
          Printer.CurrentX = Printer.CurrentX
          Printer.CurrentY = Printer.CurrentY
          Printer.FontBold = True
          Printer.Print "  " & Format(Trim(vRecordset.Fields("StkWHCode").Value), "##,##0.00") & "                " & Format(Trim(vRecordset.Fields("qtylocation").Value), "##,##0.00") & "             " & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "  " & "(" & Trim(vRecordset.Fields("unitcode").Value) & ")" & "             " & "| ________ |"
                      
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY - 80
          Printer.FontBold = False
          Printer.Print Trim("-----------------------------------------------------------------------------------------------")
          
        vRecordset.MoveNext
        n = n + 1
    Wend
  End If
  vRecordset.Close
    
    
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 400
    Printer.Print Trim("_______________________________________________________________________________________________")
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + vLineY
    Printer.Print "               ผู้จัดสินค้า                                                 Checker"
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 150
    Printer.Print "         |_____________|                                    |______________|"
          
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    
    Printer.Font.Size = 10
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("วันที่พิมพ์ :") & Now & "          " & vPrinterName
    Printer.EndDoc
    
    vMemIsPrintError = 0
End Sub

Public Sub PrintPickingSlipFromRequest(vDocno As String, vZoneGroup As String, vFamilyGroup As String, vPickZoneGroup As String, vPrinterName As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocdate As Date
'Dim vPrinterName As String
Dim printerObj As Printer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim vPrinterID As Integer

Dim strComputer As String

''On Error Resume Next


For Each printerObj In Printers
If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
Set Printer = printerObj
Set printerObj = Nothing
Exit For
End If
Next

    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vQuery = "exec dbo.USP_SO_PickingQueuePos3 '" & vDocno & "','" & vDocdate & "' ,'" & vZoneGroup & "','" & vFamilyGroup & "','" & vPickZoneGroup & "' "
    If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    If Not vRecordset.EOF Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 12
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 200
      Printer.Print vRecordset.Fields("printmydesc").Value
      
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 18
      Printer.FontBold = True
      Printer.CurrentX = 2200
      Printer.CurrentY = 120
      Printer.Print Trim("คลัง  :   ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("MBShelfCode").Value)
    

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.FontBold = True
      Printer.CurrentX = 1700
      Printer.CurrentY = 0
      Printer.Print Trim(vRecordset.Fields("queueno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 18
      Printer.FontBold = False
      Printer.CurrentX = 0
      Printer.CurrentY = 1000
      Printer.Print "ID." & Trim(vRecordset.Fields("runningno").Value)
          
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 1200
      Printer.CurrentY = 1000
      Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("_______________________________________________________________________________________")
    
      If Trim(vRecordset.Fields("iscopy").Value) = 0 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 11
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ต้นฉบับใบจัดสินค้า")
      Else
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 11
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ทดแทนใบจัดสินค้า")
      
      End If
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 1900
      Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 1900
      Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 12
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value)
            
      vRecordset.MoveFirst
      vLineX = 50
      vLineY = 50
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY - 80
      Printer.Print Trim("-----------------------------------------------------------------------------------------------")
      n = 1
      While Not vRecordset.EOF
          Printer.Font.Size = 11
          
          'Printer.CurrentX = 0
          'Printer.CurrentY = Printer.CurrentY
          'Printer.Print "ขายชั้นเก็บ :" & Trim(vRecordset.Fields("MBShelfCode").Value) & "       " & Trim("OnHand: ") & Trim(vRecordset.Fields("qtyonhand").Value) & "       " & Trim("รวมคลัง : ") & "  " & Trim(vRecordset.Fields("stkwhcode").Value) & "    " & Trim(vRecordset.Fields("unitcode").Value)
                                      
          Printer.Font.Size = 18
          Printer.FontBold = True
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "#" & n & ". " & " ที่เก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)
          
          'Printer.Font.Name = "3 of 9 Barcode"
          'Printer.Font.Size = 20
          'Printer.FontBold = False
          'Printer.CurrentX = 200
          'Printer.CurrentY = Printer.CurrentY
          'Printer.Print "*" & Trim(vRecordset.Fields("barcode").Value) & "*"
      
          Printer.Font.Name = "AngsanaUPC"
          Printer.Font.Size = 11
          Printer.FontBold = False
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "รหัส :" & Trim(vRecordset.Fields("itemcode").Value) & "                                      " & "หน่วยนับหลัก : " & Trim(vRecordset.Fields("UnitCode").Value)
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ชื่อ :" & Trim(vRecordset.Fields("itemname").Value)
          
          'Printer.CurrentX = Printer.CurrentX + 15
          'Printer.CurrentY = Printer.CurrentY + 50
          'Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")
                    
          Printer.Font.Size = 11
          Printer.CurrentX = Printer.CurrentX
          Printer.CurrentY = Printer.CurrentY
          Printer.FontBold = True
          Printer.Print "Total/รวม" & "    |    " & "OnHand" & "    |     " & "Order/ต้องการ" & "      |     " & "Pickup/จัดได้"
          
          Printer.Font.Size = 12
          Printer.CurrentX = Printer.CurrentX
          Printer.CurrentY = Printer.CurrentY
          Printer.FontBold = True
          Printer.Print "  " & Format(Trim(vRecordset.Fields("stkwhcode").Value), "##,##0") & "                " & Format(Trim(vRecordset.Fields("qtyonhand").Value), "##,##0") & "             " & Format(Trim(vRecordset.Fields("qty").Value), "##,##0") & "  " & "(" & Trim(vRecordset.Fields("unitcode").Value) & ")" & "             " & "| ________ |"
                    
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY - 80
          Printer.Print Trim("-----------------------------------------------------------------------------------------------")
          
      vRecordset.MoveNext
      n = n + 1
      Wend
    End If
    End If
    vRecordset.Close
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 400
    Printer.Print Trim("_______________________________________________________________________________________________")
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + vLineY
    Printer.Print "               ผู้จัดสินค้า                                                 Checker"
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 150
    Printer.Print "         |_____________|                                    |______________|"
    
          
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    
    Printer.CurrentX = Printer.CurrentX + 2000
    Printer.Print Trim("วันที่พิมพ์ :") & Now & "          " & vPrinterName

    Printer.EndDoc
    
    vMemIsPrintError = 0
End Sub



Public Sub PrintPickupSlipDriveThru(vDocno As String, vZoneGroup As String, vFamilyGroup As String, vPickZoneGroup As String, vPrinterName As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocdate As Date
Dim printerObj As Printer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim vPrinterID As Integer

Dim strComputer As String

''On Error Resume Next


    For Each printerObj In Printers
    If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    Set Printer = printerObj
    Set printerObj = Nothing
    Exit For
    End If
    Next

    vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vQuery = "exec dbo.USP_NP_PrintDriveInDetails '" & vDocno & "','" & vPickZoneGroup & "' "
    If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    If Not vRecordset.EOF Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 12
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 200
      Printer.Print "PickupSlip-DriveThru"
      
    Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 18
      Printer.FontBold = True
      Printer.CurrentX = 1200
      Printer.CurrentY = 120
      Printer.Print Trim("ทะเบียนรถ :   ") & Trim(vRecordset.Fields("refno").Value)
      
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 18
      Printer.FontBold = True
      Printer.CurrentX = 2200
      Printer.CurrentY = 120
      Printer.Print Trim("คลัง  :   ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("ShelfCode").Value)
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 18
      Printer.FontBold = False
      Printer.CurrentX = 0
      Printer.CurrentY = 1000
      Printer.Print "ID." & Trim(vRecordset.Fields("runningno").Value)
          
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 1200
      Printer.CurrentY = 1000
      Printer.Print "*" & Trim(vRecordset.Fields("docno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("_______________________________________________________________________________________")
    
      If Trim(vRecordset.Fields("iscopy").Value) = 0 Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 11
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ต้นฉบับPickup-DriveThru")
      Else
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 11
        Printer.CurrentX = 1500
        Printer.CurrentY = 1650
        Printer.Print Trim("ทดแทนPickup-DriveThru")
      
      End If
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 1900
      Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 12
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("arname").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value)
            
      vRecordset.MoveFirst
      vLineX = 50
      vLineY = 50
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY - 80
      Printer.Print Trim("-----------------------------------------------------------------------------------------------")
      n = 1
      While Not vRecordset.EOF
          Printer.Font.Size = 11
          
          'Printer.CurrentX = 0
          'Printer.CurrentY = Printer.CurrentY
          'Printer.Print "ขายชั้นเก็บ :" & Trim(vRecordset.Fields("MBShelfCode").Value) & "       " & Trim("OnHand: ") & Trim(vRecordset.Fields("qtyonhand").Value) & "       " & Trim("รวมคลัง : ") & "  " & Trim(vRecordset.Fields("stkwhcode").Value) & "    " & Trim(vRecordset.Fields("unitcode").Value)
                                      
          Printer.Font.Size = 18
          Printer.FontBold = True
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "#" & n & ". " & " ที่เก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)
          
          Printer.Font.Name = "3 of 9 Barcode"
          Printer.Font.Size = 20
          Printer.FontBold = False
          Printer.CurrentX = 200
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "*" & Trim(vRecordset.Fields("itemcode").Value) & "*"
      
          Printer.Font.Name = "AngsanaUPC"
          Printer.Font.Size = 11
          Printer.FontBold = False
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "รหัส :" & Trim(vRecordset.Fields("itemcode").Value) & "                                      " & "หน่วยนับหลัก : " & Trim(vRecordset.Fields("UnitCode").Value)
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ชื่อ :" & Trim(vRecordset.Fields("itemname").Value)
          
          'Printer.CurrentX = Printer.CurrentX + 15
          'Printer.CurrentY = Printer.CurrentY + 50
          'Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")
                    
          Printer.Font.Size = 11
          Printer.CurrentX = Printer.CurrentX
          Printer.CurrentY = Printer.CurrentY
          Printer.FontBold = True
          Printer.Print "Total/รวม" & "    |    " & "OnHand" & "    |     " & "Order/ต้องการ" & "      |     " & "Pickup/จัดได้"
          
          Printer.Font.Size = 12
          Printer.CurrentX = Printer.CurrentX
          Printer.CurrentY = Printer.CurrentY
          Printer.FontBold = True
          Printer.Print "  " & Format(Trim(vRecordset.Fields("stkwhcode").Value), "##,##0") & "                " & Format(Trim(vRecordset.Fields("qtylocation").Value), "##,##0") & "             " & Format(Trim(vRecordset.Fields("qty").Value), "##,##0") & "  " & "(" & Trim(vRecordset.Fields("unitcode").Value) & ")" & "             " & "| ________ |"
                    
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY - 80
          Printer.Print Trim("-----------------------------------------------------------------------------------------------")
          
      vRecordset.MoveNext
      n = n + 1
      Wend
    End If
    End If
    vRecordset.Close
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 400
    Printer.Print Trim("_______________________________________________________________________________________________")
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + vLineY
    Printer.Print "               ผู้จัดสินค้า                                                 Checker"
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 150
    Printer.Print "         |_____________|                                    |______________|"
    
          
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    
    Printer.CurrentX = Printer.CurrentX + 2000
    Printer.Print Trim("วันที่พิมพ์ :") & Now & "          " & vPrinterName

    Printer.EndDoc
    
    vMemIsPrintError = 0
End Sub


Public Sub PrintPickingSlipFromRequestDriveThru(vDocno As String, vZoneGroup As String, vFamilyGroup As String, vPickZoneGroup As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType As String
Dim vDocdate As String

On Error Resume Next

vRepID = 567
vRepType = "DT"

vDocdate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)

'MsgBox Day(Now)

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    With CrystalDT
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocno;" & vDocno & ";true"
        .ParameterFields(1) = "@vDocdate;" & vDocdate & ";true"
        .ParameterFields(2) = "@vZoneID;" & vZoneGroup & ";true"
        .ParameterFields(3) = "@vFamilyGroup;" & vFamilyGroup & ";true"
        .ParameterFields(4) = "@vPickZone;" & vPickZoneGroup & ";true"
        .Destination = crptToPrinter
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
End Sub




