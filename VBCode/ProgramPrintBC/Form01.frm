VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form01 
   Caption         =   "หน้าพิมพ์ทดแทนเอกสารทั่วไป"
   ClientHeight    =   9000
   ClientLeft      =   2430
   ClientTop       =   1185
   ClientWidth     =   12000
   Icon            =   "Form01.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form01.frx":08CA
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListViewCondition 
      Height          =   2850
      Left            =   540
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5027
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เงื่อนไข"
         Object.Width           =   16581
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   285
      Left            =   10395
      TabIndex        =   20
      Top             =   7875
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   285
      Left            =   10125
      TabIndex        =   19
      Top             =   7875
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Pic101 
      Height          =   2895
      Left            =   540
      ScaleHeight     =   2835
      ScaleWidth      =   9675
      TabIndex        =   12
      Top             =   1395
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CommandButton CMD103 
         Caption         =   "ยกเลิก"
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
         Left            =   5310
         TabIndex        =   18
         Top             =   1395
         Width           =   960
      End
      Begin VB.CommandButton CMD102 
         Caption         =   "ตกลง"
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
         Left            =   4140
         TabIndex        =   17
         Top             =   1395
         Width           =   960
      End
      Begin VB.TextBox Text101 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1485
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   1395
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "รหัสผ่าน :"
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
         Left            =   630
         TabIndex        =   15
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label LBL101 
         Caption         =   "xxxx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1080
         TabIndex        =   14
         Top             =   45
         Width           =   7170
      End
      Begin VB.Label Label1 
         Caption         =   "คำอธิบาย :"
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
         Height          =   330
         Left            =   225
         TabIndex        =   13
         Top             =   45
         Width           =   1680
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   11475
      Top             =   7740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CrystalReport011 
      Left            =   11430
      Top             =   6525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000009&
      Caption         =   "ใบขอโอนขนาด A4"
      Height          =   390
      Left            =   1725
      TabIndex        =   4
      Top             =   3150
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.CheckBox UnShowDiscount 
      BackColor       =   &H80000009&
      Caption         =   "กรณี ใบสั่งขายไม่โชว์ส่วนลด"
      Height          =   390
      Left            =   1725
      TabIndex        =   3
      Top             =   2625
      Width           =   2865
   End
   Begin VB.CommandButton CMD011 
      Caption         =   "พิมพ์"
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
      Left            =   3375
      TabIndex        =   5
      Top             =   3750
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView011 
      Height          =   2115
      Left            =   6150
      TabIndex        =   1
      Top             =   1425
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   3731
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ฟอร์มที่พิมพ์"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.TextBox TXT012 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1725
      TabIndex        =   2
      Top             =   2100
      Width           =   2865
   End
   Begin VB.TextBox TXT011 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1725
      TabIndex        =   0
      Top             =   1425
      Width           =   2865
   End
   Begin VB.Label LBLQuotation 
      BackStyle       =   0  'Transparent
      Caption         =   "เลือก เงื่อนไขใบเสนอราคาขายโครงการ"
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
      Left            =   540
      TabIndex        =   22
      Top             =   4410
      Visible         =   0   'False
      Width           =   3210
   End
   Begin VB.Label LBL017 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์ทดแทนเอกสารต่าง ๆ ทดแทน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   540
      Left            =   2550
      TabIndex        =   11
      Top             =   300
      Width           =   9285
   End
   Begin VB.Label LBL011 
      BackStyle       =   0  'Transparent
      Caption         =   "กรุณากดปุ่ม Enter อีกครั้งนะครับ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   1725
      TabIndex        =   10
      Top             =   1125
      Width           =   2790
   End
   Begin VB.Label LBL013 
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์"
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
      Left            =   600
      TabIndex        =   9
      Top             =   2100
      Width           =   1065
   End
   Begin VB.Label LBL012 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
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
      Left            =   600
      TabIndex        =   8
      Top             =   1425
      Width           =   990
   End
   Begin VB.Label LBL016 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   6120
      TabIndex        =   7
      Top             =   4005
      Width           =   4650
   End
   Begin VB.Label LBL015 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   6120
      TabIndex        =   6
      Top             =   3645
      Width           =   4650
   End
End
Attribute VB_Name = "Form01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCheckValue As Boolean
Dim vCheckValue1 As Boolean
Dim vKeyword As String
Dim vCheckKeyword As String
Dim vCheckPic101 As Integer

'---------------------------

Private Sub CMD011_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String, vDocTypeID As String, vGroupDoc As String
Dim vShelfGroup(10) As String
Dim i As Integer, vCount As Integer, vPrint As Integer, vBillStatus As Integer
Dim n As Integer, vBillType As Integer, vSend As Integer
Dim vIsConfirmPrint As Integer
Dim vOverDue As Integer
Dim vCheckAVLShelf As Integer
Dim vCheckAVLRemain As Integer
Dim vCheckBillType As Integer

On Error Resume Next


If TXT011.Text <> "" And TXT012.Text <> "" Then
    vDocNo = Trim(TXT011.Text)
    vPrintNo = vDocNo
    vPrintForm = Trim(TXT012.Text)
    Call GetComputerandUser
    vQuery = "select doctypeid, groupdoc,printed from npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        vDocTypeID = Trim(vRecordset.Fields("doctypeid").Value)
        vPrint = Trim(vRecordset.Fields("printed").Value)
    End If
    vRecordset.Close
    '-----------------------------------------------------------------------------------------
   If vPrint = 1 Then
    If vDocTypeID = "SO" Then
    '-----------------------------------------------------------
    vQuery = "select billstatus ,billtype,isconditionsend  from bcnp.dbo.bcsaleorder where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vBillStatus = Trim(vRecordset.Fields("billstatus").Value)
        vBillType = Trim(vRecordset.Fields("billtype").Value)
        vSend = Trim(vRecordset.Fields("isconditionsend").Value)
    End If
    vRecordset.Close
    If vBillStatus = 0 Then
            If TXT012.Text = Trim("พิมพ์ใบสั่งขาย+พิมพ์ใบจัดสินค้า") Then
                Call SaleOrder
                MsgBox "ต้องการพิมพ์ทดแทนใบหยิบสินค้า ต้องเข้าหน้าพิมพ์ทดแทนใบหยิบสินค้า", vbInformation, "Send Information"
            ElseIf TXT012.Text = Trim("พิมพ์ใบสั่งขาย") Then
                Call SaleOrder
            ElseIf TXT012.Text = Trim("พิมพ์ใบสั่งขาย+พิมพ์ใบจัดคิวสินค้า") Then
                Call SaleOrder
                vQuery = "select isconditionsend from bcnp.dbo.bcsaleorder where docno = '" & vDocNo & "' "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vSend = Trim(vRecordset.Fields("isconditionsend").Value)
                End If
                vRecordset.Close
                If vSend = 1 Then
                    Call SaleOrder_Delivery
                Else
                    MsgBox "เอกสารขาย ที่ประเภทเป็น ลูกค้ารับเอง ไม่สามารถพิมพ์ใบจัดคิวส่งสินค้าได้", vbCritical, "Send Massage"
                End If
            ElseIf TXT012.Text = Trim("พิมพ์ใบจัดคิวสินค้า") Then
                vQuery = "select isconditionsend from bcnp.dbo.bcsaleorder where docno = '" & vDocNo & "' "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vSend = Trim(vRecordset.Fields("isconditionsend").Value)
                End If
                vRecordset.Close
                If vSend = 1 Then
                    Call SaleOrder_Delivery
                Else
                    MsgBox "เอกสารขาย ที่ประเภทเป็น ลูกค้ารับเอง ไม่สามารถพิมพ์ใบจัดคิวส่งสินค้าได้", vbCritical, "Send Massage"
                End If
            ElseIf TXT012.Text = Trim("พิมพ์ใบสั่งจองสินค้า") Then
             vQuery = "exec dbo.usp_so_CheckConfirmPrint '" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vIsConfirmPrint = Trim(vRecordset.Fields("isconfirmprint").Value)
            End If
            vRecordset.Close
            
            If vIsConfirmPrint = 0 Then
            vQuery = "exec dbo.usp_so_CheckOverdue '" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vOverDue = Trim(vRecordset.Fields("doccount").Value)
            End If
            vRecordset.Close
            End If
            If vOverDue = 0 And vIsConfirmPrint = 0 Then
                vQuery = "exec dbo.usp_so_SearchKeyword '01' "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vKeyword = Trim(vRecordset.Fields("keyword").Value)
                End If
                vRecordset.Close
                If vCheckPic101 = 0 Then
                LBL101.Caption = Trim("กรุณาใส่รหัสผ่าน เพราะใบสั่งจองเลขที่ " & vDocNo & " วันที่ครบกำหนดเกิน 15 วัน")
                Pic101.Visible = True
                Text101.SetFocus
                Exit Sub
                End If
                If vKeyword <> vCheckKeyword Then
                    If vCheckKeyword <> "" Then
                    MsgBox "รหัสผ่านไม่ถูกต้อง", vbCritical, "Send Error"
                    End If
                    vCheckPic101 = 0
                    Exit Sub
                End If
            End If
                
                vQuery = "select isnull(billtype,0) as billtype from dbo.bcsaleorder where docno = '" & vDocNo & "' "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vCheckBillType = vRecordset.Fields("billtype").Value
                End If
                vRecordset.Close
                
                If vCheckBillType = 0 Then
                
                   vQuery = "exec dbo.USP_SO_CheckQTYReserve '" & vDocNo & "' "
                   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                       vRecordset.MoveFirst
                       While Not vRecordset.EOF
                       MsgBox "สินค้ารหัส " & vRecordset.Fields("itemcode").Value & " มียอดในชั้นเก็บ AVL ไม่พอขาย ต้องทำเอกสาร BackOrder เพื่อสั่งซื้อสินค้าเพิ่ม", vbCritical, "Send Error Message"
                       vRecordset.MoveNext
                       Wend
                       vCheckAVLRemain = 1
                   End If
                   vRecordset.Close
                '--------------------------------------------------------------------------------------------
                   If vCheckAVLRemain = 0 Then
                   Call SaleOrder_Reserve
                   Else
                   Exit Sub
                   End If
                '--------------------------------------------------------------------------------------------
                Else
                Call SaleOrder_Reserve
                End If
    
            ElseIf TXT012.Text = Trim("พิมพ์ใบจัดสินค้า") Then
                  MsgBox "ต้องการพิมพ์ทดแทนใบหยิบสินค้า ต้องเข้าหน้าพิมพ์ทดแทนใบหยิบสินค้า", vbInformation, "Send Information"
            End If
            
            Else
            MsgBox "ไม่สามารถพิมพ์ทดแทนได้ เนื่องจากได้ออกบิลไปแล้ว", vbInformation, "ข้อความเตือน"
            End If
            '---------------------------------------------------------------------------------------
    ElseIf vDocTypeID = "QT" Then
            If TXT012.Text = "พิมพ์ใบเสนอราคา" Then
                Call PrintQuotation
                ElseIf TXT012.Text = "พิมพ์ใบเสนอราคา ขายโครงการ" Then
                Call PrintQuotationWholeSale
            End If
    ElseIf vDocTypeID = "RQ" Then
        If TXT012.Text = "พิมพ์ใบเสนอซื้อสินค้า" Then
            Call PrintStockRequest
        End If
    ElseIf vDocTypeID = "TF" Then
            If TXT012.Text = "พิมพ์ใบขอโอนสินค้า" Then
                    If Check1.Value = 0 Then
                        Call PrintStockTransfer
                    Else
                        Call PrintStockTransfer_A4
                    End If
            End If
    ElseIf vDocTypeID = "BO" Then
    If TXT012.Text = "พิมพ์ใบBackOrder" Then
        Call PrintBackOrder
    End If
    ElseIf vDocTypeID = "BD" Then
            Call PrintBankDeposit
    ElseIf vDocTypeID = "CI" Then
        If TXT012.Text = "พิมพ์ใบนำฝากเงินสดธนาคารกรุงเทพ" Or TXT012.Text = "พิมพ์ใบนำฝากเงินสดธนาคารเอเซีย" Then
            Call PrintBankDeposit_Cash
        End If
    ElseIf vDocTypeID = "PO" Then
       If TXT012.Text = "พิมพ์ใบสั่งซื้อ" Then
               Call PrintPO
       ElseIf TXT012.Text = "พิมพ์ใบอนุมัติค่าใช้จ่าย(สำหรับบุคคล)" Then
               Call PrintPOExpense
       End If
    End If
 Else
 MsgBox "คุณไม่สามารถพิมพ์ทดแทนได้ เนื่องจากคุณยังไม่ได้พิมพ์เอกสารตัวจริงที่หน้าพิมพ์เอกสาร", vbInformation, "ข้อความเตือน"
  Me.ListViewCondition.Visible = False
 Me.LBLQuotation.Visible = False
End If
Else
 MsgBox "ไม่มีเลขที่เอกสารที่จะพิมพ์ หรือ ฟอร์มที่จะพิมพ์", vbInformation + vbCritical, "ข้อความเตือน"
 Me.ListViewCondition.Visible = False
 Me.LBLQuotation.Visible = False
End If
TXT011.Text = ""
TXT012.Text = ""
LBL016.Caption = ""
LBL015.Caption = ""
vCheckPic101 = 0
ListView011.ListItems.Clear

End Sub

Private Sub CMD102_Click()
vCheckKeyword = Trim(Text101.Text)
vCheckPic101 = 1
Pic101.Visible = False
Call CMD011_Click
Text101.Text = ""
End Sub

Private Sub CMD103_Click()
vCheckKeyword = Trim(Text101.Text)
vCheckPic101 = 0
Pic101.Visible = False
Text101.Text = ""
End Sub

Private Sub Command1_Click()
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer

vPrinterName = Trim("EPSON TM-T88IIR Partial cut")
For Each printerObj In Printers
  If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    Set Printer = printerObj
    Set printerObj = Nothing
    Exit For
  End If
Next

    vQuery = "exec dbo.USP_SO_PickingQueue 'SCV4908-2497','015','c' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 14
      Printer.FontBold = True
      Printer.CurrentX = 1500
      Printer.CurrentY = 250
      Printer.Print "คิวที่ :" & Trim(vRecordset.Fields("queueno").Value)
      
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 800
      Printer.CurrentY = 550
      Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1000
      Printer.Print Trim("_______________________________________________________________________________________")
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 1500
      Printer.CurrentY = 1200
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า")
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 2200
      Printer.CurrentY = 1400
      Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 2200
      Printer.CurrentY = 1600
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1800
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 2000
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 2200
      Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      If vRecordset.Fields("isconditionsend").Value = 0 Then
            Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("รับเอง")
      Else
            Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("ส่งให้")
      End If
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 2600
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value)
      
      vRecordset.MoveFirst
      vLineX = 50
      vLineY = 50
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY - 50
      Printer.Print Trim("-----------------------------------------------------------------------------------------------")
      n = 1
      While Not vRecordset.EOF
          Printer.Font.Size = 10
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode1").Value) & "       " & Trim("คงเหลือตามคลัง : ") & Trim(vRecordset.Fields("qtylocation").Value) & "  " & Trim(vRecordset.Fields("unitcode").Value)
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value)
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ชื่อสินค้า :" & Trim(vRecordset.Fields("itemname").Value)
          
          Printer.CurrentX = Printer.CurrentX + 15
          Printer.CurrentY = Printer.CurrentY + 100
          Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY - 50
          Printer.Print Trim("-----------------------------------------------------------------------------------------------")
          
      vRecordset.MoveNext
      n = n + 1
      Wend
    End If
    vRecordset.Close
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 400
    Printer.Print Trim("_______________________________________________________________________________________________")
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + vLineY
    Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 150
    Printer.Print "         _____________                                    ______________"
          
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    

      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY
      Printer.Print Trim("V:\Reports\RP_BC_PickingQueue.rpt")
      Printer.CurrentX = Printer.CurrentX + 2000
      Printer.Print Trim("วันที่พิมพ์ :") & Now
           
    Printer.EndDoc
End Sub

Private Sub Command2_Click()
Dim vPrinterName As String
Dim printerObj As Printer


vPrinterName = Trim("EPSON TM-T88II(R) Reduce35")
For Each printerObj In Printers
  If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    Set Printer = printerObj
    Set printerObj = Nothing
    Exit For
  End If
Next

      Printer.FontName = "3 of 9 Barcode"
      Printer.Font.Size = 50
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print "สสส :"
      Printer.EndDoc
      
      Printer.FontName = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.CurrentX = 0
      Printer.CurrentY = 700
      Printer.Print "กกก"
      
      Printer.FontName = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.CurrentX = 0
      Printer.CurrentY = 1500
      Printer.Print "สสส :"
      
      Printer.FontName = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.CurrentX = 0
      Printer.CurrentY = 2000
      Printer.Print "สสส :"
      Printer.EndDoc

End Sub

Private Sub ListView011_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT012.Text = Item
End Sub

Private Sub TXT011_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String, vDocTypeID As String, vGroupDoc As String
Dim vCreatorcode As String
Dim vCreatedatetime As Date
Dim FormListItems As ListItem
Dim vPrint As Integer

On Error GoTo ErrDescription

ListView011.ListItems.Clear
If KeyAscii = 13 Then
    vDocNo = Trim(TXT011.Text)
    vQuery = "select docno,Lastprinteduser,lastprintdatetime,doctypeid,groupdoc,printed from npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCreatorcode = Trim(vRecordset.Fields("Lastprinteduser").Value)
        vCreatedatetime = Trim(vRecordset.Fields("lastprintdatetime").Value)
        vDocTypeID = Trim(vRecordset.Fields("doctypeid").Value)
        vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        vPrint = Trim(vRecordset.Fields("printed").Value)
    Else
        Call CheckDocument
        Exit Sub
    End If
    vRecordset.Close
    
    If vDocTypeID = "QT" Then
      Me.ListViewCondition.Visible = True
      Me.LBLQuotation.Visible = True
      Call GetQuotationCondition
    Else
      Me.ListViewCondition.Visible = False
      Me.LBLQuotation.Visible = False
    End If

    
    If vDocTypeID = "TF" Then
        Check1.Visible = True
    End If
    LBL015.Caption = "ผู้ที่ทำเอกสาร คือ   " & vCreatorcode
    LBL016.Caption = "วันที่ทำเอกสาร คือ  " & vCreatedatetime

    vQuery = "select name from npmaster.dbo.npform where moduleid = '" & vDocTypeID & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set FormListItems = ListView011.ListItems.Add(, , Trim(vRecordset.Fields("name").Value))
            vRecordset.MoveNext
        Wend
    vRecordset.Close
  
    End If

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Function CheckDocument()
Dim vRecordset  As New ADODB.Recordset
Dim vQuery As String
Dim vDocGroup As String, vDocument As String, vGroupDoc1 As String
Dim vTable As String, vCheckDoc As String, vTypeDoc As String, vDocGroup1 As String

On Error GoTo ErrDescription
vDocGroup1 = Left(Trim(TXT011.Text), 3)
vDocument = Trim(UCase(TXT011.Text))
vQuery = "select upper('" & vDocGroup1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup = Trim(vRecordset.Fields("vDocGroup").Value)
End If
vRecordset.Close
'--------------------------------------------------------------------------------------------

If vDocGroup = "SHV" Or vDocGroup = "SHN" Or vDocGroup = "SCV" Or vDocGroup = "SCN" Or vDocGroup = "SVD" Or _
    vDocGroup = "SVN" Or vDocGroup = "SVM" Or vDocGroup = "SAB" Or vDocGroup = "ROV" Or vDocGroup = "RON" Then
    vTable = "BCNP.DBO.BCSALEORDER"
    vTypeDoc = "SO"
    Me.ListViewCondition.Visible = False
    Me.LBLQuotation.Visible = False
End If
If vDocGroup = "BHV" Or vDocGroup = "BHN" Or vDocGroup = "BCV" Or vDocGroup = "BCN" Or vDocGroup = "BVD" Or _
    vDocGroup = "BVN" Or vDocGroup = "BAB" Or vDocGroup = "BVM" Then
    vTable = "BCNP.DBO.BCQUOTATION"
    vTypeDoc = "BO"
    Me.ListViewCondition.Visible = False
    Me.LBLQuotation.Visible = False
End If
If vDocGroup = "QHV" Or vDocGroup = "QHN" Or vDocGroup = "QCV" Or vDocGroup = "QVM" Or vDocGroup = "QCN" Or vDocGroup = "QVD" Or _
    vDocGroup = "QVN" Or vDocGroup = "QAB" Then
    vTable = "BCNP.DBO.BCQUOTATION"
    vTypeDoc = "QT"
    Call GetQuotationCondition
    Me.ListViewCondition.Visible = True
    Me.LBLQuotation.Visible = True
End If

If vDocGroup = "BOA" Or vDocGroup = "BOT" Or vDocGroup = "BBT" Or vDocGroup = "BBL" Then
    vTable = "BCNP.DBO.BCCHQINDEPOSIT"
    vTypeDoc = "BD"
    Me.ListViewCondition.Visible = False
    Me.LBLQuotation.Visible = False
End If

If Left(vDocGroup, 2) = "IF" Then
    vTable = "BCNP.DBO.BCSTKTRANSFER2"
    vTypeDoc = "TF"
    Me.ListViewCondition.Visible = False
    Me.LBLQuotation.Visible = False
End If

If Left(vDocGroup, 2) = "PR" Then
    vTable = "BCNP.DBO.BCSTKRequest"
    vTypeDoc = "RQ"
    Me.ListViewCondition.Visible = False
    Me.LBLQuotation.Visible = False
End If

If vDocGroup = "POV" Or vDocGroup = "PON" Or vDocGroup = "POC" Or vDocGroup = "POE" Then
    vTable = "BCNP.DBO.BCPURCHASEORDER"
    vTypeDoc = "PO"
    Me.ListViewCondition.Visible = False
    Me.LBLQuotation.Visible = False
End If

If vTable <> "" Then
vQuery = "select docno from " & vTable & " where docno = '" & vDocument & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDoc = vRecordset.Fields("Docno").Value
End If
vRecordset.Close
'-------------------------------------------------------------------------------------------------
End If

vGroupDoc1 = Left(vDocGroup, 2)


If vCheckDoc = "" Or IsNull(vCheckDoc) Then
    MsgBox "ไม่มีเอกสารนี้ในระบบ", vbCritical, "ข้อความเตือน"
Else
'--------------------------------------------------------------------------
    If vTypeDoc = "PO" Then
        vQuery = "INSERT INTO NPMaster.dbo.NPPrintLogs(DocNo,DocTypeID,GroupDOC,Code,Printed,LastPrintedUser,LastPrintDateTime)" & _
                     "select docno,'" & vTypeDoc & "' as DocTypeID,'" & vDocGroup & "' as GroupDoc,apcode,1 as Printed,'" & vUserID & "' as UserID ,getdate() From " & vTable & " where   docno = '" & vDocument & "'  "
    ElseIf vTypeDoc = "BD" Then
        vQuery = "INSERT INTO NPMaster.dbo.NPPrintLogs(DocNo,DocTypeID,GroupDOC,Printed,LastPrintedUser,LastPrintDateTime)" & _
                     "select docno,'" & vTypeDoc & "' as DocTypeID,'" & vDocGroup & "' as GroupDoc,1 as Printed,'" & vUserID & "' as UserID ,getdate() From " & vTable & " where   docno = '" & vDocument & "'  "
    ElseIf vTypeDoc = "TF" Then
     vQuery = "INSERT INTO NPMaster.dbo.NPPrintLogs(DocNo,DocTypeID,GroupDOC,Code,Printed,LastPrintedUser,LastPrintDateTime)" & _
                     "select docno,'" & vTypeDoc & "' as DocTypeID,'" & vGroupDoc1 & "' as GroupDoc,'' as arcode,1 as Printed,'" & vUserID & "' as UserID ,getdate() From " & vTable & " where   docno = '" & vDocument & "'  "
    ElseIf vTypeDoc = "PR" Then
     vQuery = "INSERT INTO NPMaster.dbo.NPPrintLogs(DocNo,DocTypeID,GroupDOC,Code,Printed,LastPrintedUser,LastPrintDateTime)" & _
                     "select docno,'" & vTypeDoc & "' as DocTypeID,'" & vGroupDoc1 & "' as GroupDoc,'' as arcode,1 as Printed,'" & vUserID & "' as UserID ,getdate() From " & vTable & " where   docno = '" & vDocument & "'  "
    ElseIf vTypeDoc = "BO" Then
     vQuery = "INSERT INTO NPMaster.dbo.NPPrintLogs(DocNo,DocTypeID,GroupDOC,Code,Printed,LastPrintedUser,LastPrintDateTime)" & _
                     "select docno,'" & vTypeDoc & "' as DocTypeID,'" & vDocGroup & "' as GroupDoc,'' as arcode,1 as Printed,'" & vUserID & "' as UserID ,getdate() From " & vTable & " where   docno = '" & vDocument & "'  "
    ElseIf vTypeDoc = "SO" Then
     vQuery = "INSERT INTO NPMaster.dbo.NPPrintLogs(DocNo,DocTypeID,GroupDOC,Code,Printed,LastPrintedUser,LastPrintDateTime)" & _
                     "select docno,'" & vTypeDoc & "' as DocTypeID,'" & vDocGroup & "' as GroupDoc,'' as arcode,1 as Printed,'" & vUserID & "' as UserID ,getdate() From " & vTable & " where   docno = '" & vDocument & "'  "
    Else
    vQuery = "INSERT INTO NPMaster.dbo.NPPrintLogs(DocNo,DocTypeID,GroupDOC,Code,Printed,LastPrintedUser,LastPrintDateTime)" & _
                     "select docno,'" & vTypeDoc & "' as DocTypeID,'" & vDocGroup & "' as GroupDoc,'' as arcode,1 as Printed,'" & vUserID & "' as UserID ,getdate() From " & vTable & " where   docno = '" & vDocument & "'  "
    End If
    gConnection.Execute vQuery
    MsgBox "กรุณากดปุ่ม Enter อีกครั้งนะครับ", vbCritical, "ข้อความแจ้งให้ทราบ"
    '-------------------------------------------------------------------------------------------
End If

vQuery = "Delete npmaster.dbo.npprintserver where docno = '" & vDocument & "' "
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If

End Function

Public Sub PrintPO()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT011.Text))
    vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' and typenumber = 5"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 'เคยพิมพ์แล้ว
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
        Else
            vCheckDocNo = 0 'ยังไม่ได้พิมพ์
        End If
        vRecordset.Close
        
        If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
            vHeaderType = 15
            vNamePrint = Trim(vUserID)
            vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                vHeader = Trim(vRecordset.Fields("header").Value)
                vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
            End If
            vRecordset.Close
            vDocuments = vDocNumber & vHeader & "-" & vPicking
            
            vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
            gConnection.Execute vQuery
            
            vTypeNumber = 5
            vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vTypeNumber = 5
            vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
            vQuery = "exec bcnp.dbo.USP_NP_InsertLogPrintRunningRes '" & vDocNo & "','" & vUserID & "' ," & vTypeNumber & " "
            gConnection.Execute vQuery
    End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน

        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = UCase(Trim(vRecordset.Fields("groupdoc").Value))
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "PO"
                            If vGroupDoc = "POV" Or vGroupDoc = "pov" Or vGroupDoc = "POC" Or vGroupDoc = "poc" Then
                                If Trim(TXT012.Text) <> Trim("พิมพ์ใบสั่งซื้อหลายบรรทัด") Then
                                    vRepID = 276
                                Else
                                    vRepID = 309
                                End If
                            ElseIf vGroupDoc = "POE" Or vGroupDoc = "poe" Then
                                vRepID = 352
                            ElseIf vGroupDoc = "PON" Or vGroupDoc = "pon" Then
                                vRepID = 277
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport1
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintPOExpense()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT011.Text))
    vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' and typenumber = 5"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 'เคยพิมพ์แล้ว
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
        Else
            vCheckDocNo = 0 'ยังไม่ได้พิมพ์
        End If
        vRecordset.Close
        
        If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
            vHeaderType = 15
            vNamePrint = Trim(vUserID)
            vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                vHeader = Trim(vRecordset.Fields("header").Value)
                vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
            End If
            vRecordset.Close
            vDocuments = vDocNumber & vHeader & "-" & vPicking
            
            vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
            gConnection.Execute vQuery
            
            vTypeNumber = 5
            vQuery = "exec dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vTypeNumber = 5
            vQuery = "exec dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
            vQuery = "exec dbo.USP_NP_InsertLogPrintRunningRes '" & vDocNo & "','" & vUserID & "' ," & vTypeNumber & " "
            gConnection.Execute vQuery
    End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน

        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = UCase(Trim(vRecordset.Fields("groupdoc").Value))
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "PO"

                            If vGroupDoc = "POE" Or vGroupDoc = "poe" Then
                                vRepID = 419
                            Else
                               MsgBox "เลขที่เอกสารไม่สามารถพิมพ์ฟอร์มนี้ได้ กรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                               Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport1
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub
Public Sub PrintPOCheck()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer

On Error Resume Next

    vDocNo = UCase(Trim(TXT011.Text))
    vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' and typenumber = 6"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 'เคยพิมพ์แล้ว
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
        Else
            vCheckDocNo = 0 'ยังไม่ได้พิมพ์
        End If
        vRecordset.Close
        
        If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
            vHeaderType = 22
            vNamePrint = Trim(vUserID)
            vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                vHeader = Trim(vRecordset.Fields("header").Value)
                vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
            End If
            vRecordset.Close
            vDocuments = vDocNumber & vHeader & "-" & vPicking
            
            vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
            gConnection.Execute vQuery
            
            vTypeNumber = 6
            vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vTypeNumber = 6
            vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
            vQuery = "exec bcnp.dbo.USP_NP_InsertLogPrintRunningRes '" & vDocNo & "','" & vUserID & "' ," & vTypeNumber & " "
            gConnection.Execute vQuery
            
    End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
        
        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "PO"
                            If vGroupDoc = "POV" Or vGroupDoc = "pov" Or vGroupDoc = "POC" Or vGroupDoc = "poc" Or vGroupDoc = "POE" Or vGroupDoc = "poe" Then
                                vRepID = 278
                            ElseIf vGroupDoc = "PON" Or vGroupDoc = "pon" Then
                                vRepID = 278
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport1
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description
'End If
End Sub

Public Sub PrintQuotation()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT011.Text))
vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDocNo = 1 'เคยพิมพ์แล้ว
    vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
Else
    vCheckDocNo = 0 'ยังไม่ได้พิมพ์
End If
vRecordset.Close

If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
    vQuery = "select docno,billtype from bcnp.dbo.BCQuotation where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckBillType = Trim(vRecordset.Fields("billtype").Value)
    End If
    vRecordset.Close
    If vCheckBillType = 0 Then
        vHeaderType = 16
    ElseIf vCheckBillType = 1 Then
        vHeaderType = 17
    End If
    
    
    vNamePrint = Trim(vUserID)
    vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
        vHeader = Trim(vRecordset.Fields("header").Value)
        vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
    End If
    vRecordset.Close
    vDocuments = vDocNumber & vHeader & "-" & vPicking
    
    vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
    gConnection.Execute vQuery
    
    vTypeNumber = 3
    vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
    gConnection.Execute vQuery
Else
    vDocuments = vRunNumber
    vTypeNumber = 3
    vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
    gConnection.Execute vQuery
    vQuery = "exec bcnp.dbo.USP_NP_InsertLogPrintRunningRes '" & vDocNo & "','" & vUserID & "' ," & vTypeNumber & " "
    gConnection.Execute vQuery
End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน

        vDocNo = Trim(TXT011.Text)
        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "QT"
                            If vGroupDoc = "QHV" Or vGroupDoc = "qhv" Then
                                vRepID = 279
                            ElseIf vGroupDoc = "QHN" Or vGroupDoc = "qhn" Then
                                vRepID = 280
                            ElseIf vGroupDoc = "QCV" Or vGroupDoc = "qcv" Then
                                vRepID = 279
                            ElseIf vGroupDoc = "QCN" Or vGroupDoc = "qcn" Then
                                vRepID = 280
                            ElseIf vGroupDoc = "QVD" Or vGroupDoc = "qvd" Then
                                vRepID = 279
                            ElseIf vGroupDoc = "QVM" Or vGroupDoc = "qvm" Then
                                vRepID = 279
                            ElseIf vGroupDoc = "QVN" Or vGroupDoc = "qvn" Then
                                vRepID = 280
                            ElseIf vGroupDoc = "QAB" Or vGroupDoc = "qab" Then
                                vRepID = 279
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport1
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                      
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintQuotationWholeSale()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer
Dim vCondition1 As String
Dim vCondition2 As String
Dim vCondition3 As String
Dim vCondition4 As String
Dim vCondition5 As String
Dim vIndexCondition1 As Integer
Dim vIndexCondition2 As Integer
Dim vIndexCondition3 As Integer
Dim vIndexCondition4 As Integer
Dim vIndexCondition5 As Integer
Dim i As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT011.Text))
vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDocNo = 1 'เคยพิมพ์แล้ว
    vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
Else
    vCheckDocNo = 0 'ยังไม่ได้พิมพ์
End If
vRecordset.Close

If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
    vQuery = "select docno,billtype from bcnp.dbo.BCQuotation where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckBillType = Trim(vRecordset.Fields("billtype").Value)
    End If
    vRecordset.Close
    If vCheckBillType = 0 Then
        vHeaderType = 16
    ElseIf vCheckBillType = 1 Then
        vHeaderType = 17
    End If
    
    
    vNamePrint = Trim(vUserID)
    vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
        vHeader = Trim(vRecordset.Fields("header").Value)
        vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
    End If
    vRecordset.Close
    vDocuments = vDocNumber & vHeader & "-" & vPicking
    
    vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
    gConnection.Execute vQuery
    
    vTypeNumber = 3
    vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
    gConnection.Execute vQuery
Else
    vDocuments = vRunNumber
    vTypeNumber = 3
    vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
    gConnection.Execute vQuery
    vQuery = "exec bcnp.dbo.USP_NP_InsertLogPrintRunningRes '" & vDocNo & "','" & vUserID & "' ," & vTypeNumber & " "
    gConnection.Execute vQuery
End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน

        vDocNo = Trim(TXT011.Text)
        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
        vRepType = "QT"
        If UCase(vGroupDoc) = "QHV" Or vGroupDoc = "qhv" Then
            vRepID = 333
        ElseIf UCase(vGroupDoc) = "QHN" Or vGroupDoc = "qhn" Then
            vRepID = 334
        ElseIf UCase(vGroupDoc) = "QCV" Or vGroupDoc = "qcv" Then
            vRepID = 333
        ElseIf UCase(vGroupDoc) = "QCN" Or vGroupDoc = "qcn" Then
            vRepID = 334
        ElseIf UCase(vGroupDoc) = "QVD" Or vGroupDoc = "qvd" Then
            vRepID = 333
        ElseIf UCase(vGroupDoc) = "QVM" Or vGroupDoc = "qvm" Then
            vRepID = 333
        ElseIf UCase(vGroupDoc) = "QVN" Or vGroupDoc = "qvn" Then
            vRepID = 334
        ElseIf UCase(vGroupDoc) = "QAB" Or vGroupDoc = "qab" Then
            vRepID = 333
        Else
        MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
        Exit Sub
        End If
                            
                            
For i = 1 To Me.ListViewCondition.ListItems.Count
    If Me.ListViewCondition.ListItems(i).Checked = True And vCondition1 = "" Then
          vCondition1 = Me.ListViewCondition.ListItems(i).Text
          vIndexCondition1 = i
    i = Me.ListViewCondition.ListItems.Count
    Else
    vIndexCondition1 = Me.ListViewCondition.ListItems.Count
    End If
Next i

If vIndexCondition1 <> Me.ListViewCondition.ListItems.Count And vIndexCondition1 <> 0 Then
For i = vIndexCondition1 + 1 To Me.ListViewCondition.ListItems.Count
    If Me.ListViewCondition.ListItems(i).Checked = True And vCondition2 = "" Then
          vCondition2 = Me.ListViewCondition.ListItems(i).Text
          vIndexCondition2 = i
    i = Me.ListViewCondition.ListItems.Count
    Else
    vIndexCondition2 = Me.ListViewCondition.ListItems.Count
    End If
Next i
End If

If vIndexCondition2 <> Me.ListViewCondition.ListItems.Count And vIndexCondition2 <> 0 Then
For i = vIndexCondition2 + 1 To Me.ListViewCondition.ListItems.Count
    If Me.ListViewCondition.ListItems(i).Checked = True And vCondition3 = "" Then
          vCondition3 = Me.ListViewCondition.ListItems(i).Text
          vIndexCondition3 = i
    i = Me.ListViewCondition.ListItems.Count
    Else
    vIndexCondition3 = Me.ListViewCondition.ListItems.Count
    End If
Next i
End If

If vIndexCondition3 <> Me.ListViewCondition.ListItems.Count And vIndexCondition3 <> 0 Then
For i = vIndexCondition3 + 1 To Me.ListViewCondition.ListItems.Count
    If Me.ListViewCondition.ListItems(i).Checked = True And vCondition4 = "" Then
          vCondition4 = Me.ListViewCondition.ListItems(i).Text
          vIndexCondition4 = i
    i = Me.ListViewCondition.ListItems.Count
    Else
    vIndexCondition4 = Me.ListViewCondition.ListItems.Count
    End If
Next i
End If

If vIndexCondition4 <> Me.ListViewCondition.ListItems.Count And vIndexCondition4 <> 0 Then
For i = vIndexCondition4 + 1 To Me.ListViewCondition.ListItems.Count
    If Me.ListViewCondition.ListItems(i).Checked = True And vCondition5 = "" Then
          vCondition5 = Me.ListViewCondition.ListItems(i).Text
    i = Me.ListViewCondition.ListItems.Count
    End If
Next i
End If

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport1
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Formulas(0) = "Condition1='" & vCondition1 & "' "
        .Formulas(1) = "Condition2='" & vCondition2 & "' "
        .Formulas(2) = "Condition3='" & vCondition3 & "' "
        .Formulas(3) = "Condition4='" & vCondition4 & "' "
        .Formulas(4) = "Condition5='" & vCondition5 & "' "
        .Action = 1
    End With
End If
vRecordset.Close

Me.ListViewCondition.Visible = False
Me.LBLQuotation.Visible = False
'---------------------------------------------------------------------------------------------------
                      
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintBackOrder()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT011.Text))
vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 'เคยพิมพ์แล้ว
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
        Else
            vCheckDocNo = 0 'ยังไม่ได้พิมพ์
        End If
        vRecordset.Close
        
        If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
            vQuery = "select docno,billtype from bcnp.dbo.BCQuotation where docno = '" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckBillType = Trim(vRecordset.Fields("billtype").Value)
            End If
            vRecordset.Close
            If vCheckBillType = 0 Then
                vHeaderType = 18
            ElseIf vCheckBillType = 1 Then
                vHeaderType = 19
            End If
            
            vNamePrint = Trim(vUserID)
            vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                vHeader = Trim(vRecordset.Fields("header").Value)
                vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
            End If
            vRecordset.Close
            vDocuments = vDocNumber & vHeader & "-" & vPicking
            
            vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
            gConnection.Execute vQuery
            
            vTypeNumber = 4
            vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vTypeNumber = 4
            vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
            vQuery = "exec bcnp.dbo.USP_NP_InsertLogPrintRunningRes '" & vDocNo & "','" & vUserID & "' ," & vTypeNumber & " "
            gConnection.Execute vQuery
    End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน

        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "BO"
                            If vGroupDoc = "BHV" Or vGroupDoc = "bhv" Then
                                vRepID = 274
                            ElseIf vGroupDoc = "BHN" Or vGroupDoc = "bhn" Then
                                vRepID = 275
                            ElseIf vGroupDoc = "BCV" Or vGroupDoc = "bcv" Then
                                vRepID = 274
                            ElseIf vGroupDoc = "BCN" Or vGroupDoc = "bcn" Then
                                vRepID = 275
                            ElseIf vGroupDoc = "BVD" Or vGroupDoc = "bvd" Then
                                vRepID = 274
                            ElseIf vGroupDoc = "BVM" Or vGroupDoc = "bvm" Then
                                vRepID = 274
                            ElseIf vGroupDoc = "BVN" Or vGroupDoc = "bvn" Then
                                vRepID = 275
                            ElseIf vGroupDoc = "BAB" Or vGroupDoc = "bab" Then
                                vRepID = 274
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport1
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                            
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub


Public Sub PrintSaleOrder011()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String, vWHCode As String
        Dim vGroupDoc As String
        Dim vPrint As Integer
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
          
        vDocNo = Trim(TXT011.Text)


        '---------------------------------------------------------------------------------------------------------
        
                            vRepType = "SO"
                            vRepID = 80
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '-------------------------------------------------------------------------------------------------------

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintSaleOrder011Copy()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String, vWHCode As String
        Dim vGroupDoc As String
        Dim vPrint As Integer
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT011.Text)
        

        '---------------------------------------------------------------------------------------------------------
        
                            vRepType = "SO"
                            vRepID = 81
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '-------------------------------------------------------------------------------------------------------

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub


Public Sub PrintSaleOrderOthers()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String, vWHCode As String
        Dim vGroupDoc As String
        Dim vPrint As Integer
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT011.Text)
        

        '---------------------------------------------------------------------------------------------------------
        
                            vRepType = "SO"
                                vRepID = 82
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '-------------------------------------------------------------------------------------------------------

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub


Public Sub PrintSaleOrderOthersCopy()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String, vWHCode As String
        Dim vGroupDoc As String
        Dim vPrint As Integer
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT011.Text)
        
 
        '---------------------------------------------------------------------------------------------------------
        
                            vRepType = "SO"
                                vRepID = 83
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '-------------------------------------------------------------------------------------------------------

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub SaleOrder()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vReportName As String
Dim vDocNo As String, vWHCode As String
Dim vGroupDoc As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vCheckDocNo As Integer
Dim vHeaderType As Integer
Dim vCheckBillType As Integer
Dim vRunNumber As String
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT011.Text))
vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
End If
vRecordset.Close

If UCase(vGroupDoc) = "SHV" Or UCase(vGroupDoc) = "SHN" Or UCase(vGroupDoc) = "SCV" Or UCase(vGroupDoc) = "SCN" Or UCase(vGroupDoc) = "SVD" Or UCase(vGroupDoc) = "SVN" Or UCase(vGroupDoc) = "SAB" Or UCase(vGroupDoc) = "SVE" Or UCase(vGroupDoc) = "SVM" Then
        vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 'เคยพิมพ์แล้ว
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
        Else
            vCheckDocNo = 0 'ยังไม่ได้พิมพ์
        End If
        vRecordset.Close
        
        If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
            vQuery = "select docno,billtype from bcnp.dbo.bcsaleorder where docno = '" & vDocNo & "' and sostatus = 0 "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckBillType = Trim(vRecordset.Fields("billtype").Value)
            End If
            vRecordset.Close
            If vCheckBillType = 0 Then
                vHeaderType = 13
            ElseIf vCheckBillType = 1 Then
                vHeaderType = 14
            End If
            
            vNamePrint = Trim(vUserID)
            vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                vHeader = Trim(vRecordset.Fields("header").Value)
                vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
            End If
            vRecordset.Close
            vDocuments = vDocNumber & vHeader & "-" & vPicking
            
            vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
            gConnection.Execute vQuery
            
            vTypeNumber = 1
            vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vTypeNumber = 1
            vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
            vQuery = "exec bcnp.dbo.USP_NP_InsertLogPrintRunningRes '" & vDocNo & "','" & vUserID & "' ," & vTypeNumber & " "
            gConnection.Execute vQuery
        End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
                        vRepType = "SO"
                            If vGroupDoc = "SHV" Or vGroupDoc = "shv" Or vGroupDoc = "SCV" Or vGroupDoc = "scv" Or vGroupDoc = "SVM" Or vGroupDoc = "svm" Or vGroupDoc = "SVD" Or vGroupDoc = "svd" Or vGroupDoc = "SAB" Or vGroupDoc = "sab" Or vGroupDoc = "SVE" Or vGroupDoc = "sve" Then
                                If UnShowDiscount.Value = False Then
                                vRepID = 281
                                Else
                                vRepID = 286
                                End If
                            ElseIf vGroupDoc = "SHN" Or vGroupDoc = "shn" Or vGroupDoc = "SCN" Or vGroupDoc = "scn" Or vGroupDoc = "SVN" Or vGroupDoc = "svn" Then
                                If UnShowDiscount.Value = False Then
                                vRepID = 282
                                Else
                                vRepID = 287
                                End If
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
        '---------------------------------------------------------------------------------------------------------
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport1
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '------------------------------------------------------------------------------------------------------
Else
    MsgBox "ไม่สามารถพิมพ์ใบสั่งขายได้ เลือกพิมพ์เอกสารผิด กรุณาเลือกใหม่"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If


End Sub


Public Sub SaleOrder_Delivery()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String, vWHCode As String
        Dim vGroupDoc As String
        Dim vPrint As Integer
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = UCase(Trim(TXT011.Text))
        
        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
               ' vPrint = Trim(vRecordset.Fields("printed").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------------------
        vRepType = "SO"
        vRepID = 76
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport1
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '------------------------------------------------------------------------------------------------------

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintBankDeposit()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT011.Text)

        vQuery = "select doctypeid,groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("doctypeid").Value)

        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        If vGroupDoc = "BD" Then
            vRepType = "BD"
            If Trim(TXT012.Text) = "พิมพ์ใบนำฝากธนาคารกรุงเทพฯ" Then
                vRepID = 39
            ElseIf Trim(TXT012.Text) = "พิมพ์ใบนำฝากธนาคารเอเชีย" Then
                vRepID = 40
            ElseIf Trim(TXT012.Text) = "พิมพ์ใบนำฝากธนาคารไทยพาณิชย์" Then
                vRepID = 229
            ElseIf Trim(TXT012.Text) = "พิมพ์ใบนำฝากธนาคารกรุงศรีอยุธยา" Then
                vRepID = 230
            ElseIf Trim(TXT012.Text) = "พิมพ์ใบนำฝากธนาคารกรุงไทย" Then
                vRepID = 228
            ElseIf Trim(TXT012.Text) = "พิมพ์ใบนำฝากธนาคารกสิกรไทย" Then
                vRepID = 231
            ElseIf Trim(TXT012.Text) = "พิมพ์ใบนำฝากธนาคารทหารไทย" Then
                vRepID = 332
            Else
                MsgBox "เลือกฟอร์มพิมพ์ใบนำฝากไม่ถูกต้อง"
            End If
        End If
        
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            With CrystalReport011
                .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .Action = 1
            End With
        End If
        vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                         
                            
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub


Public Sub PrintPicking_A()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCount As Integer
        Dim vCheckPrintA As Integer
        Dim vPicking As String
        Dim vHeader As String
        Dim vDocuments As String
        Dim vNamePrint As String
        Dim vDocNumber As String
        
        On Error GoTo ErrDescription
        
        vDocNo = UCase(Trim(TXT011.Text))
        vNamePrint = vUserID
        
        vQuery = "select saleorderno from npmaster.dbo. NP_PickingSlip_Logs where saleorderno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckPrintA = 1
        Else
            vCheckPrintA = 0
        End If
        vRecordset.Close
        
        If vCheckPrintA = 1 Then
        vQuery = "Update npmaster.dbo. NP_PickingSlip_Logs set lastuserprint = '" & vUserID & "' , lastprintdatetime = getdate() ,Lastprintcount = LastPrintCount + 1 where saleorderno = '" & vDocNo & "' and shelfgroup ='A' "
        gConnection.Execute vQuery
        Else
        '-----------------------------------------------------------------------------------------------------------------
        vQuery = "exec dbo.USP_NP_SearchNewDocNo 27 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
            vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
        End If
        vRecordset.Close
        vDocuments = vDocNumber & vHeader & "-" & vPicking
        
        vQuery = "exec dbo.USP_NP_UpdateNewDocNo 27 "
        gConnection.Execute vQuery
        
        vQuery = "insert into npmaster.dbo.np_pickingslip_logs (saleorderno,pickingno,pickingdate,whcode,shelfgroup,userprint,salecode1,lastprintcount) " _
                        & " values('" & vDocNo & "','" & vDocuments & "',getdate(),'010','A','" & vUserID & "','" & vNamePrint & "',1) "
        gConnection.Execute vQuery
        '--------------------------------------------------------------------------------------
        End If
        
        vRepType = "SO"
        vRepID = 107
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

        'vQuery = "select lastprintcount,shelfgroup  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         '   vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        'End If
        'vRecordset.Close
        
        'vCount = vCount + 1
        'vQuery = "Update npmaster.dbo.NP_PickingSlip_Logs  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' ,lastprintdatetime = getdate() where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        'gConnection.Execute vQuery
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintPicking_A_Copy()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT011.Text)
        vRepType = "SO"
        vRepID = 108
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintPicking_B()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCheckPrintB As Integer
        Dim vPicking As String
        Dim vHeader As String
        Dim vDocuments As String
        Dim vNamePrint As String
        Dim vDocNumber As String
        
        On Error GoTo ErrDescription
        
        vDocNo = UCase(Trim(TXT011.Text))
        vNamePrint = vUserID
        vQuery = "select saleorderno from npmaster.dbo. NP_PickingSlip_Logs where saleorderno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckPrintB = 1
        Else
            vCheckPrintB = 0
        End If
        vRecordset.Close
         
        If vCheckPrintB = 1 Then
        vQuery = "Update npmaster.dbo. NP_PickingSlip_Logs set lastuserprint = '" & vUserID & "' , lastprintdatetime = getdate() ,Lastprintcount = LastPrintCount + 1 where saleorderno = '" & vDocNo & "' and shelfgroup ='B' "
        gConnection.Execute vQuery
        Else
        vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 27 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
            vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
        End If
        vRecordset.Close
        vDocuments = vDocNumber & vHeader & "-" & vPicking
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 9 "
        gConnection.Execute vQuery
        
        vQuery = "insert into npmaster.dbo.np_pickingslip_logs (saleorderno,pickingno,pickingdate,whcode,shelfgroup,userprint,salecode1,lastprintcount) " _
                        & " values('" & vDocNo & "','" & vDocuments & "',getdate(),'010','B','" & vUserID & "','" & vNamePrint & "',1) "
        gConnection.Execute vQuery
        End If
            
        vRepType = "SO"
        vRepID = 109
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintPicking_B_Copy()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT011.Text)
        vRepType = "SO"
        vRepID = 110
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintPicking_C()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCheckPrintC As Integer
        Dim vPicking As String
        Dim vHeader As String
        Dim vDocuments As String
        Dim vNamePrint As String
        Dim vDocNumber As String
        
        On Error GoTo ErrDescription
        
        vDocNo = UCase(Trim(TXT011.Text))
        vNamePrint = vUserID
        vQuery = "select saleorderno from npmaster.dbo. NP_PickingSlip_Logs where saleorderno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckPrintC = 1
        Else
            vCheckPrintC = 0
        End If
        vRecordset.Close
        
        If vCheckPrintC = 1 Then
        vQuery = "Update npmaster.dbo. NP_PickingSlip_Logs set lastuserprint = '" & vUserID & "' , lastprintdatetime = getdate() ,Lastprintcount = LastPrintCount + 1 where saleorderno = '" & vDocNo & "' and shelfgroup ='C' "
        gConnection.Execute vQuery
        Else
        vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 27 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
            vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
        End If
        vRecordset.Close
        vDocuments = vDocNumber & vHeader & "-" & vPicking
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 9 "
        gConnection.Execute vQuery
        
        vQuery = "insert into npmaster.dbo.np_pickingslip_logs (saleorderno,pickingno,pickingdate,whcode,shelfgroup,userprint,salecode1,lastprintcount) " _
                        & " values('" & vDocNo & "','" & vDocuments & "',getdate(),'015','C','" & vUserID & "','" & vNamePrint & "',1) "
        gConnection.Execute vQuery
        End If
        
        vRepType = "SO"
        vRepID = 111
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintPicking_C_Copy()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT011.Text)
        vRepType = "SO"
        vRepID = 112
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintPicking_D()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCheckPrintD As Integer
        Dim vPicking As String
        Dim vHeader As String
        Dim vDocuments As String
        Dim vNamePrint As String
        Dim vDocNumber As String
        
        On Error GoTo ErrDescription
        
        vDocNo = UCase(Trim(TXT011.Text))
        vNamePrint = vUserID
        vQuery = "select saleorderno from npmaster.dbo. NP_PickingSlip_Logs where saleorderno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckPrintD = 1
        Else
            vCheckPrintD = 0
        End If
        vRecordset.Close

        If vCheckPrintD = 1 Then
        vQuery = "Update npmaster.dbo. NP_PickingSlip_Logs set lastuserprint = '" & vUserID & "' , lastprintdatetime = getdate() ,Lastprintcount = LastPrintCount + 1 where saleorderno = '" & vDocNo & "' and shelfgroup ='D' "
        gConnection.Execute vQuery
        Else
        vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 27 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
            vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
        End If
        vRecordset.Close
        vDocuments = vDocNumber & vHeader & "-" & vPicking
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 9 "
        gConnection.Execute vQuery
        
        vQuery = "insert into npmaster.dbo.np_pickingslip_logs (saleorderno,pickingno,pickingdate,whcode,shelfgroup,userprint,salecode1,lastprintcount) " _
                        & " values('" & vDocNo & "','" & vDocuments & "',getdate(),'010','D','" & vUserID & "','" & vNamePrint & "',1) "
        gConnection.Execute vQuery
        End If
        
        vRepType = "SO"
        vRepID = 113
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintPicking_D_Copy()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT011.Text)
        vRepType = "SO"
        vRepID = 114
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintStockTransfer()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT011.Text)
        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "TF"
                            vRepID = 155
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                      
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub
Public Sub PrintBankDeposit_Cash()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String, vReportType As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT011.Text)
        vReportType = Trim(TXT012.Text)
        

        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
                            
                            
                            vRepType = "BD"
                            If vReportType = "พิมพ์ใบนำฝากเงินสดธนาคารกรุงเทพ" Then
                                vRepID = 122
                            ElseIf vReportType = "พิมพ์ใบนำฝากเงินสดธนาคารเอเซีย" Then
                                vRepID = 123
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                         
                            
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintStockTransfer_A4()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT011.Text)
        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "TF"
                            vRepID = 167
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                            Check1.Visible = False
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintPicking_M()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCount As Integer
        Dim vCheckPrintM As Integer
        Dim vPicking As String
        Dim vHeader As String
        Dim vDocuments As String
        Dim vNamePrint As String
        Dim vDocNumber As String
        
        On Error GoTo ErrDescription
        
        vDocNo = UCase(Trim(TXT011.Text))
        vNamePrint = vUserID
        vQuery = "select saleorderno from npmaster.dbo. NP_PickingSlip_Logs where saleorderno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckPrintM = 1
        Else
            vCheckPrintM = 0
        End If
        vRecordset.Close

        If vCheckPrintM = 1 Then
        vQuery = "Update npmaster.dbo. NP_PickingSlip_Logs set lastuserprint = '" & vUserID & "' , lastprintdatetime = getdate() ,Lastprintcount = LastPrintCount + 1 where saleorderno = '" & vDocNo & "' and shelfgroup ='M' "
        gConnection.Execute vQuery
        Else
        vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 27 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
            vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
        End If
        vRecordset.Close
        vDocuments = vDocNumber & vHeader & "-" & vPicking
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 9 "
        gConnection.Execute vQuery
        
        vQuery = "insert into npmaster.dbo.np_pickingslip_logs (saleorderno,pickingno,pickingdate,whcode,shelfgroup,userprint,salecode1,lastprintcount) " _
                        & " values('" & vDocNo & "','" & vDocuments & "',getdate(),'014','M','" & vUserID & "','" & vNamePrint & "',1) "
        gConnection.Execute vQuery
        End If
        
        vRepType = "SO"
        vRepID = 175
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

        'vQuery = "select lastprintcount,shelfgroup  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         '   vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        'End If
        'vRecordset.Close
        
        'vCount = vCount + 1
        'vQuery = "Update npmaster.dbo.NP_PickingSlip_Logs  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' ,lastprintdatetime = getdate() where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        'gConnection.Execute vQuery
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintPicking_M_Copy()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT011.Text)
        vRepType = "SO"
        vRepID = 176
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub


Public Sub CheckValue1()
Dim i As Integer
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String

vDocNo = Trim(TXT011.Text)
vQuery = "select typecode from bcnp.dbo.vw_IV_PackingSlip where docno = '" & vDocNo & "' and shelfgroup1 = 'M' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckValue = True
Else
    vCheckValue = False
End If
vRecordset.Close


End Sub

Public Sub PrintPicking_M_OutLet()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCount As Integer
        Dim vCheckPrintMOutLet As Integer
        Dim vPicking As String
        Dim vHeader As String
        Dim vDocuments As String
        Dim vNamePrint As String
        Dim vDocNumber As String
        
        On Error GoTo ErrDescription
        
        vDocNo = UCase(Trim(TXT011.Text))
        vNamePrint = vUserID
        vQuery = "select saleorderno from npmaster.dbo. NP_PickingSlip_Logs where saleorderno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckPrintMOutLet = 1
        Else
            vCheckPrintMOutLet = 0
        End If
        vRecordset.Close

        If vCheckPrintMOutLet = 1 Then
        vQuery = "Update npmaster.dbo. NP_PickingSlip_Logs set lastuserprint = '" & vUserID & "' , lastprintdatetime = getdate() ,Lastprintcount = LastPrintCount + 1 where saleorderno = '" & vDocNo & "' and shelfgroup ='M' "
        gConnection.Execute vQuery
        Else
        vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 27 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
            vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
        End If
        vRecordset.Close
        vDocuments = vDocNumber & vHeader & "-" & vPicking
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 9 "
        gConnection.Execute vQuery
        
        vQuery = "insert into npmaster.dbo.np_pickingslip_logs (saleorderno,pickingno,pickingdate,whcode,shelfgroup,userprint,salecode1,lastprintcount) " _
                        & " values('" & vDocNo & "','" & vDocuments & "',getdate(),'014','M','" & vUserID & "','" & vNamePrint & "',1) "
        gConnection.Execute vQuery
        End If
        
        vRepType = "SO"
        vRepID = 184
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintPicking_M_OutLet_Copy()

        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT011.Text)
        vRepType = "SO"
        vRepID = 186
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub CheckValueHMX()
Dim i As Integer
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String

vDocNo = Trim(TXT011.Text)
vQuery = "select itemcode from bcnp.dbo.bcsaleordersub  where docno = '" & vDocNo & "' and whcode = '014' and  typecode not in (select itemtype from npmaster.dbo.NP_ItemOutLet) "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckValue1 = True
Else
    vCheckValue1 = False
End If
vRecordset.Close

End Sub

Public Sub PrintPicking_HMX()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCount As Integer
        Dim vCheckPrintHMX As Integer
        Dim vPicking As String
        Dim vHeader As String
        Dim vDocuments As String
        Dim vNamePrint As String
        Dim vDocNumber As String
        
        On Error GoTo ErrDescription
        
        vDocNo = UCase(Trim(TXT011.Text))
        vNamePrint = vUserID
        vQuery = "select saleorderno from npmaster.dbo. NP_PickingSlip_Logs where saleorderno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckPrintHMX = 1
        Else
            vCheckPrintHMX = 0
        End If
        vRecordset.Close

        If vCheckPrintHMX = 1 Then
        vQuery = "Update npmaster.dbo. NP_PickingSlip_Logs set lastuserprint = '" & vUserID & "' , lastprintdatetime = getdate() ,Lastprintcount = LastPrintCount + 1 where saleorderno = '" & vDocNo & "' and shelfgroup ='M' "
        gConnection.Execute vQuery
        Else
        vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 27 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
            vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
        End If
        vRecordset.Close
        vDocuments = vDocNumber & vHeader & "-" & vPicking
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 9 "
        gConnection.Execute vQuery
        
        vQuery = "insert into npmaster.dbo.np_pickingslip_logs (saleorderno,pickingno,pickingdate,whcode,shelfgroup,userprint,salecode1,lastprintcount) " _
                        & " values('" & vDocNo & "','" & vDocuments & "',getdate(),'014','H','" & vUserID & "','" & vNamePrint & "',1) "
        gConnection.Execute vQuery
        End If
        
        vRepType = "SO"
        vRepID = 200
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

        'vQuery = "select lastprintcount,shelfgroup  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         '   vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        'End If
        'vRecordset.Close
        
        'vCount = vCount + 1
        'vQuery = "Update npmaster.dbo.NP_PickingSlip_Logs  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' ,lastprintdatetime = getdate() where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        'gConnection.Execute vQuery
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintPicking_HMX_Copy()

        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT011.Text)
        vRepType = "SO"
        vRepID = 201
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintPicking_Y()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCount As Integer
        Dim vCheckPrintY As Integer
        Dim vPicking As String
        Dim vHeader As String
        Dim vDocuments As String
        Dim vNamePrint As String
        Dim vDocNumber As String
        
        On Error GoTo ErrDescription
        
        vDocNo = UCase(Trim(TXT011.Text))
        vNamePrint = vUserID
        vQuery = "select saleorderno from npmaster.dbo. NP_PickingSlip_Logs where saleorderno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckPrintY = 1
        Else
            vCheckPrintY = 0
        End If
        vRecordset.Close

        If vCheckPrintY = 1 Then
        vQuery = "Update npmaster.dbo. NP_PickingSlip_Logs set lastuserprint = '" & vUserID & "' , lastprintdatetime = getdate() ,Lastprintcount = LastPrintCount + 1 where saleorderno = '" & vDocNo & "' and shelfgroup ='A' "
        gConnection.Execute vQuery
        Else
        vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 27 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
            vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
        End If
        vRecordset.Close
        vDocuments = vDocNumber & vHeader & "-" & vPicking
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 9 "
        gConnection.Execute vQuery
        
        vQuery = "insert into npmaster.dbo.np_pickingslip_logs (saleorderno,pickingno,pickingdate,whcode,shelfgroup,userprint,salecode1,lastprintcount) " _
                        & " values('" & vDocNo & "','" & vDocuments & "',getdate(),'016','Y','" & vUserID & "','" & vNamePrint & "',1) "
        gConnection.Execute vQuery
        End If
        
        vRepType = "SO"
        vRepID = 222
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

        'vQuery = "select lastprintcount,shelfgroup  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         '   vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        'End If
        'vRecordset.Close
        
        'vCount = vCount + 1
        'vQuery = "Update npmaster.dbo.NP_PickingSlip_Logs  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' ,lastprintdatetime = getdate() where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        'gConnection.Execute vQuery
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintPicking_Y_Copy()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT011.Text)
        vRepType = "SO"
        vRepID = 223
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub SaleOrder_Reserve()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vReportName As String
Dim vDocNo As String, vWHCode As String
Dim vGroupDoc As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vSoStatus As Integer
Dim vCheckDocNo As String
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vIsConfirmPrint As Integer
Dim vOverDue As Integer
Dim vCheckKey As String
Dim vKeyword As String

On Error GoTo ErrDescription
        
vDocNo = UCase(Trim(TXT011.Text))
vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
End If
vRecordset.Close
If UCase(vGroupDoc) = "ROV" Or UCase(vGroupDoc) = "RON" Then
        vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 'เคยพิมพ์แล้ว
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
        Else
            vCheckDocNo = 0 'ยังไม่ได้พิมพ์
        End If
        vRecordset.Close
        
        If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
            vQuery = "select docno,billtype from bcnp.dbo.bcsaleorder where docno = '" & vDocNo & "' and sostatus =1"
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckBillType = Trim(vRecordset.Fields("billtype").Value)
            End If
            vRecordset.Close
            If vCheckBillType = 0 Then
                vHeaderType = 20
            ElseIf vCheckBillType = 1 Then
                vHeaderType = 21
            End If
            
            vNamePrint = Trim(vUserID)
            vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                vHeader = Trim(vRecordset.Fields("header").Value)
                vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
            End If
            vRecordset.Close
            vDocuments = vDocNumber & vHeader & "-" & vPicking
            
            vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
            gConnection.Execute vQuery
            
            vTypeNumber = 2
            vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vTypeNumber = 2
            vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
            vQuery = "exec bcnp.dbo.USP_NP_InsertLogPrintRunningRes '" & vDocNo & "','" & vUserID & "' ," & vTypeNumber & " "
            gConnection.Execute vQuery
        End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
        
        vQuery = "select sostatus from bcnp.dbo.bcsaleorder where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vSoStatus = vRecordset.Fields("sostatus").Value
        End If
        vRecordset.Close
        
        If vSoStatus = 1 Then
        '---------------------------------------------------------------------------------------------------------
        vRepType = "SO"
        If vGroupDoc = "ROV" Or vGroupDoc = "ROV" Then
            vRepID = 284
        ElseIf vGroupDoc = "RON" Or vGroupDoc = "RON" Then
            vRepID = 285
        End If
        
       vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
       'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport1
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '------------------------------------------------------------------------------------------------------
    vQuery = "exec dbo.usp_so_UpdateIsConfirmPrint '" & vDocNo & "' "
    gConnection.Execute vQuery
        Else
            MsgBox "เอกสารเลขที่ " & vDocNo & " ไม่ใช่เอกสารใบจองสินค้า กรุณาตรวจสอบ"
        End If
Else
    MsgBox "ไม่สามารถพิมพ์ใบสั่งจองได้ เลือกพิมพ์เอกสารผิด กรุณาเลือกพิมพ์เอกสารใหม่"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintPicking_HMX_CustReceive()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCount As Integer
        Dim vCheckPrintHMX As Integer
        Dim vPicking As String
        Dim vHeader As String
        Dim vDocuments As String
        Dim vNamePrint As String
        Dim vDocNumber As String
        
        On Error GoTo ErrDescription
        
        vDocNo = UCase(Trim(TXT011.Text))
        vNamePrint = vUserID
        vQuery = "select saleorderno from npmaster.dbo. NP_PickingSlip_Logs where saleorderno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckPrintHMX = 1
        Else
            vCheckPrintHMX = 0
        End If
        vRecordset.Close

        If vCheckPrintHMX = 1 Then
        vQuery = "Update npmaster.dbo. NP_PickingSlip_Logs set lastuserprint = '" & vUserID & "' , lastprintdatetime = getdate() ,Lastprintcount = LastPrintCount + 1 where saleorderno = '" & vDocNo & "' and shelfgroup ='M' "
        gConnection.Execute vQuery
        Else
        vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 27 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
            vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
        End If
        vRecordset.Close
        vDocuments = vDocNumber & vHeader & "-" & vPicking
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 9 "
        gConnection.Execute vQuery
        
        vQuery = "insert into npmaster.dbo.np_pickingslip_logs (saleorderno,pickingno,pickingdate,whcode,shelfgroup,userprint,salecode1,lastprintcount) " _
                        & " values('" & vDocNo & "','" & vDocuments & "',getdate(),'014','H','" & vUserID & "','" & vNamePrint & "',1) "
        gConnection.Execute vQuery
        End If
        
        vRepType = "SO"
        vRepID = 258
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport011
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub

Public Sub PrintStockRequest()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer
Dim vCheckHeader As String

'On Error GoTo ErrDescription

    vDocNo = UCase(Trim(TXT011.Text))
    vCheckHeader = Left(vDocNo, 3)
    vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckDocNo = 1 'เคยพิมพ์แล้ว
        vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
    Else
        vCheckDocNo = 0 'ยังไม่ได้พิมพ์
    End If
    vRecordset.Close
    vHeaderType = 25
    vNamePrint = Trim(vUserID)
    If vCheckDocNo = 0 Then
    vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
        vHeader = Trim(vRecordset.Fields("header").Value)
        vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
    End If
    vRecordset.Close
    vDocuments = vDocNumber & vHeader & "-" & vPicking

    vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
    gConnection.Execute vQuery
    
    vTypeNumber = 6
    vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
    gConnection.Execute vQuery
Else
    vDocuments = vRunNumber
    vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
    gConnection.Execute vQuery
End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
        
        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '--------------------------------------------------------------------------------------------
        If vCheckHeader = "PRE" Then
          vRepID = 351
        Else
          vRepID = 293
        End If
        
        vRepType = "RQ"
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            With CrystalReport1
                .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .Formulas(0) = "computername='" & vComputerName1 & "' "
                .Formulas(1) = "username='" & vUserName1 & "' "
                .Action = 1
            End With
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------------
                                                                                  
'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description
'End If
End Sub

Public Sub GetQuotationCondition()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim QTItemforms As ListItem

Me.ListViewCondition.ListItems.Clear
vQuery = "select * from npmaster.dbo.TB_NP_QuotationTextCondition order by roworder"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set QTItemforms = ListViewCondition.ListItems.Add(, , Trim(vRecordset.Fields("textcondition").Value))
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub

