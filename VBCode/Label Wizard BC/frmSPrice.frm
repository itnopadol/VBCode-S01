VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form frmSPrice 
   Caption         =   "sFormSPrice : Special Price"
   ClientHeight    =   5565
   ClientLeft      =   5190
   ClientTop       =   3330
   ClientWidth     =   11085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   11085
   Begin Crystal.CrystalReport Crystal101 
      Left            =   8700
      Top             =   4575
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
   Begin VB.Frame Frame1 
      Caption         =   "รูปแบบการพิมพ์"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   4800
      Width           =   4815
      Begin VB.OptionButton Option2 
         Caption         =   "พิมพ์ออกเครื่องพิมพ์"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ตัวอย่างก่อนพิมพ์"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   4980
      Width           =   1215
   End
   Begin MSComctlLib.ListView LV_SPrice 
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   7646
      View            =   3
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item Number"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Barcode"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "รายชื่อสินค้า"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ราคาพิเศษ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ราคาปกติ"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmSPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim tmpPathName As String

Private Sub cmdCancel_Click()
        frmWizard.Enabled = True
        Unload Me
End Sub

Private Sub cmdOK_Click()
        Dim strSQL, strSQL2 As String
        Dim iCount As Integer
        Dim strUpdate As String
        ' Update sPrice
        ConnectSQL
        For iCount = 1 To LV_SPrice.ListItems.count
                strSQL = "Select * From NP_Label_Temp Where barcode = '" & Trim(LV_SPrice.ListItems(iCount).SubItems(1)) & "' AND UsedUser = '" & strUsername & "'"
                If Rs2.State = adStateOpen Then Rs2.Close
                Rs2.Open strSQL, ConnSQL, adOpenKeyset, adLockOptimistic
                If Not Rs2.EOF Then
                         ' MsgBox "Row Many = " & Rs2.RecordCount
                         ' strUpdate = "UPDATE NP_Label_Temp SET SPrice = "
                        'Rs2!SPrice = Trim(Me.LV_SPrice.ListItems(iCount).SubItems(4))
                        'Rs2.Update
                End If
                Rs2.Close
        Next iCount
        
        ' Printing Process
        
        ' Declaration
        Dim tmpPathName, strCreate_Temp, strDel_Temp As String
        
        ' Get Path Name
        tmpPathName = Trim(frmWizard.LV_Report.SelectedItem.SubItems(1)) & ".rpt"
        
        ' สร้าง Report_Temp
            On Error Resume Next
            strCreate_Temp = "select * into dbo.Report_Temp From NP_LABEL_TEMP where UsedUser = 'Null'"
            ConnSQL.Execute strCreate_Temp
            
        ' ลบข้อมูลที่อยู่ใน Report_Temp
            strDel_Temp = "Delete From Report_Temp"
            ConnSQL.Execute strDel_Temp

        ' Dump Data to Report_Temp
        strSQL = "Select * From NP_Label_Temp Where UsedUser = '" & strUsername & "'"
        Rs3.Open strSQL, ConnSQL, adOpenDynamic, adLockOptimistic
        If Not Rs3.EOF Then
                Rs3.MoveFirst
                While Not Rs3.EOF
                    If Int(Rs3!QTY) > 0 Then
                    'strSQL2 = "Insert Into dbo.Report_Temp(ItemCode, barcode, NAME1, NAME2, QTY, PriceLevel, Price, UnitCode, UsedUser, Category_ID, WHCode, ShelfCode, VENDR_ID, remark, SPrice,RemainOutQTY,RemainInQTY) " _
                                        & " values('" & Trim(Rs3!Itemcode) & "','" & Trim(Rs3!barcode) & "','" & Trim(Rs3!NAME1) & "', '" & Trim(Rs3!NAME2) & "', 1,'" & Trim(Rs3!PriceLevel) & "','" & Trim(Rs3!Price) & "','" & Trim(Rs3!UnitCode) & "'," _
                                        & "'" & Trim(Rs3!UsedUser) & "','" & Trim(Rs3!Category_ID) & "','" & Trim(Rs3!WHCode) & "', '" & Trim(Rs3!ShelfCode) & "','" & Trim(Rs3!VENDR_ID) & "', '" & Trim(Rs3!Remark) & "','" & Trim(Rs3!SPrice) & "' ,'" & Rs3!RemainOutQTY & "','" & Rs3!RemainInQTY & "') "
                            strSQL2 = "Insert Into dbo.Report_Temp(ItemCode, barcode, NAME1, NAME2, QTY, PriceLevel, Price, UnitCode, UsedUser, Category_ID, WHCode, ShelfCode, VENDR_ID, remark, SPrice) " _
                                        & " values('" & Trim(Rs3!Itemcode) & "','" & Trim(Rs3!barcode) & "','" & Trim(Rs3!Name1) & "', '" & Trim(Rs3!Name2) & "', 1,'" & Trim(Rs3!PriceLevel) & "','" & Trim(Rs3!price) & "','" & Trim(Rs3!Unitcode) & "'," _
                                        & "'" & Trim(Rs3!UsedUser) & "','" & Trim(Rs3!Category_ID) & "','" & Trim(Rs3!whcode) & "', '" & Trim(Rs3!ShelfCode) & "','" & Trim(Rs3!VENDR_ID) & "', '" & Trim(Rs3!Remark) & "','" & Trim(Rs3!SPrice) & "' ) "
                            For iCount = 1 To Rs3!QTY
                                    ConnSQL.Execute strSQL2
                            Next iCount
                    End If
                    Rs3.MoveNext
                Wend
        End If
        Rs3.Close
        
        ' เชื่อมต่อกับ Crystal Report
        With Crystal101
                    .ReportFileName = tmpPathName
                    .WindowState = crptMaximized
                    .Connect = "uid=VBUser;pwd=132"
                        If Option1.Value = True Then
                                .Destination = crptToWindow
                        Else
                                .Destination = crptToPrinter
                        End If
                        .Action = 1
        End With

        ' ลบ NP_Label_Temp where useduser = strUserName
        strSQL = "Delete From NP_Label_Temp Where Useduser = '" & strUsername & "'"
        ConnSQL.Execute strSQL
        
        ' ลบ Table Report_Temp
        strSQL = "Drop Table Report_Temp"
        ConnSQL.Execute strSQL
End Sub

Private Sub Form_Load()
        ' Clear ค่าใน List
        Me.LV_SPrice.ListItems.Clear
        
        ' Set Option1 Focus
        Option1.Value = True
        
        ' Declaration
        Dim ListX As ListItem
        Dim iCount As Integer
        
        For iCount = 1 To frmWizard.ListResult.ListItems.count
            If frmWizard.ListResult.ListItems(iCount).Checked = True Then
                ' Add Detail To LV_SPrice
                Set ListX = LV_SPrice.ListItems.Add(, , frmWizard.ListResult.ListItems(iCount).Text)            ' Item Number
                ListX.SubItems(1) = frmWizard.ListResult.ListItems(iCount).SubItems(2)                                   ' Barcode
                ListX.SubItems(2) = frmWizard.ListResult.ListItems(iCount).SubItems(3)                                  ' ItemDescription
                ListX.SubItems(3) = frmWizard.ListResult.ListItems(iCount).SubItems(5)                                  ' ราคาพิเศษ
                ListX.SubItems(4) = frmWizard.ListResult.ListItems(iCount).SubItems(7)
                ' Add SPrice To LV_SPrice SubItem(4)
                'ListX.SubItems(4) = Trim(SPrice(iCount))                                                                                            ' ราคาปกติ
            End If
        Next
End Sub

Private Sub LV_SPrice_DblClick()
        ' Add Data in field Selected to frmEditSPrice
        frmEditSPrice.lbItemNumber = Trim(Me.LV_SPrice.SelectedItem.Text)
        frmEditSPrice.lbItemDesc = Me.LV_SPrice.SelectedItem.SubItems(2)
        frmEditSPrice.lbSPrice = Me.LV_SPrice.SelectedItem.SubItems(3)
        frmEditSPrice.txtPrice = Me.LV_SPrice.SelectedItem.SubItems(4)
        
        ' เก็บค่า Index LV_SPrice เวลาบันทึกกลับ
        frmEditSPrice.txtIndex = Me.LV_SPrice.SelectedItem.Index
        
        frmEditSPrice.Show
        frmEditSPrice.txtPrice.SetFocus
End Sub
