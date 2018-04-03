VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormTaxPurchase 
   Caption         =   "บันทึก ภาษีซื้อ"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormTaxPurchase.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Opt104 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "บันทึกซื้อสินค้า,บริการ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8370
      TabIndex        =   31
      Top             =   675
      Width           =   2760
   End
   Begin VB.OptionButton Opt103 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "หน้าบันทึก ลดหนี้"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5265
      TabIndex        =   30
      Top             =   1485
      Width           =   2715
   End
   Begin VB.PictureBox Pic101 
      BackColor       =   &H00C0FFC0&
      Height          =   5145
      Left            =   630
      ScaleHeight     =   5085
      ScaleWidth      =   10665
      TabIndex        =   18
      Top             =   2070
      Visible         =   0   'False
      Width           =   10725
      Begin VB.CommandButton CMD103 
         Caption         =   "ยกเลิก"
         Height          =   375
         Left            =   9495
         TabIndex        =   14
         Top             =   4365
         Width           =   915
      End
      Begin VB.CommandButton CMD102 
         Caption         =   "ตกลง"
         Height          =   375
         Left            =   8460
         TabIndex        =   13
         Top             =   4365
         Width           =   915
      End
      Begin VB.CommandButton CMD101 
         Height          =   330
         Left            =   225
         Picture         =   "FormTaxPurchase.frx":72FB
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1890
         Width           =   915
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   1950
         Left            =   225
         TabIndex        =   12
         Top             =   2340
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   3440
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่ใบกำกับ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "วันที่ใบกำกับ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "คำอธิบาย"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "อัตรา"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "ยอดก่อนภาษี"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "ยอดภาษี"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "ยอดยกเว้นภาษี"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "กลุ่มภาษี"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "วันที่ยื่นภาษี"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "รูปแบบ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "BookCode"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Source"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "SaveForm "
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "APCode"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Docdate"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker102 
         Height          =   285
         Left            =   5175
         TabIndex        =   9
         Top             =   1485
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   58916865
         CurrentDate     =   38828
      End
      Begin MSComCtl2.DTPicker DTPicker101 
         Height          =   285
         Left            =   8955
         TabIndex        =   2
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   58916865
         CurrentDate     =   38828
      End
      Begin VB.TextBox Text105 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox CMB101 
         Height          =   315
         Left            =   8955
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text108 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         TabIndex        =   8
         Top             =   1485
         Width           =   1455
      End
      Begin VB.TextBox Text107 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8955
         TabIndex        =   7
         Top             =   1035
         Width           =   1455
      End
      Begin VB.TextBox Text106 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5175
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text104 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8955
         TabIndex        =   4
         Top             =   630
         Width           =   1455
      End
      Begin VB.TextBox Text103 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         TabIndex        =   3
         Top             =   675
         Width           =   5280
      End
      Begin VB.TextBox Text102 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5175
         TabIndex        =   1
         Top             =   270
         Width           =   1455
      End
      Begin VB.TextBox Text101 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1350
         TabIndex        =   0
         ToolTipText     =   "กรอกเลขที่เอกสารแล้ว กด Enter"
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รูปแบบ :"
         Height          =   285
         Left            =   7920
         TabIndex        =   29
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ยื่นภาษี :"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4050
         TabIndex        =   28
         Top             =   1485
         Width           =   1050
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "กลุ่มภาษี :"
         Height          =   285
         Left            =   45
         TabIndex        =   27
         Top             =   1485
         Width           =   1230
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ยอดยกเว้นภาษี :"
         Height          =   285
         Left            =   7605
         TabIndex        =   26
         Top             =   1035
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ยอดภาษี :"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   4185
         TabIndex        =   25
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ยอดก่อนภาษี :"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   225
         TabIndex        =   24
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "อัตรา :"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   8055
         TabIndex        =   23
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "คำอธิบาย :"
         Height          =   240
         Left            =   450
         TabIndex        =   22
         Top             =   675
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ใบกำกับภาษี :"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7560
         TabIndex        =   21
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบกำกับ :"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3960
         TabIndex        =   20
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่เอกสาร :"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   225
         TabIndex        =   19
         Top             =   270
         Width           =   1050
      End
   End
   Begin VB.OptionButton Opt102 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "หน้าบันทึก จ่ายเงินอื่น ๆ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5265
      TabIndex        =   16
      Top             =   1080
      Width           =   2715
   End
   Begin VB.OptionButton Opt101 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "หน้าบันทึก เอกสารตั้งเจ้าหนี้อื่น ๆ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5265
      TabIndex        =   15
      Top             =   675
      Width           =   2715
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   11340
      X2              =   11340
      Y1              =   495
      Y2              =   1935
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   5040
      X2              =   11340
      Y1              =   1935
      Y2              =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   5040
      X2              =   11340
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   5040
      X2              =   5040
      Y1              =   495
      Y2              =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลือก หน้าเอกสาร  :"
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
      Height          =   285
      Left            =   3330
      TabIndex        =   17
      Top             =   1080
      Width           =   1635
   End
   Begin VB.Menu Temp 
      Caption         =   "Taxno :"
   End
   Begin VB.Menu Menu1 
      Caption         =   ""
      Begin VB.Menu mEdit 
         Caption         =   "แก้ไขรายการ"
      End
      Begin VB.Menu mDelete 
         Caption         =   "ลบรายการ"
      End
   End
End
Attribute VB_Name = "FormTaxPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdit As Integer
Dim vLineEdit As Integer

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItem As ListItem
Dim i As Integer
Dim vTaxNo As String
Dim vCheckTax As String
Dim vSaveFrom As Integer
Dim vBookCode As String
Dim vSource As Integer
Dim vDocType As Integer
Dim vDocdate As String
Dim vApCode As String
Dim vDocNo As String

On Error GoTo ErrDescription

If Text101.Text <> "" And Text102.Text <> "" And Text105.Text <> "" And Text106.Text <> "" Then
    vDocNo = Trim(Text101.Text)
    vTaxNo = Trim(Text102.Text)
    If vEdit = 0 Then
        For i = 1 To ListView101.ListItems.Count
            vCheckTax = Trim(ListView101.ListItems.Item(i).Text)
            If UCase(vTaxNo) = UCase(vCheckTax) Then
                MsgBox "มีเลขที่ใบกำกับภาษีนี้ ใช้อยู่แล้วในเอกสารเลขที่นี้", vbCritical, "Send Error"
                Text102.SetFocus
                Exit Sub
            End If
        Next i
'--------------------------------------------------------------------
    If Opt101.Value = True Then
        vSource = 31
        vSaveFrom = 6
        vDocType = 1
    ElseIf Opt102.Value = True Then
        vSource = 20
        vSaveFrom = 3
        vDocType = 2
    ElseIf Opt104.Value = True Then
        vSource = 2
        vSaveFrom = 1
        vDocType = 4
    End If
'----------------------------------------------------------------------
        If Opt103.Value = False Then
            Set vListItem = ListView101.ListItems.Add(, , Trim(Text102.Text))
            vListItem.SubItems(1) = CDate(Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year))
            vListItem.SubItems(2) = Trim(Text103.Text)
            vListItem.SubItems(3) = Format(Trim(Text104.Text), "##,##0.00")
            vListItem.SubItems(4) = Format(Trim(Text105.Text), "##,##0.00")
            vListItem.SubItems(5) = Format(Trim(Text106.Text), "##,##0.00")
            vListItem.SubItems(6) = Format(Trim(Text107.Text), "##,##0.00")
            vListItem.SubItems(7) = Trim(Text108.Text)
            vListItem.SubItems(8) = CDate(Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year))
            Select Case CMB101.Text
            Case Trim("ขอคืนภาษีได้")
                vListItem.SubItems(9) = Trim("0")
            Case Trim("ขอคืนภาษีไม่ได้")
            vListItem.SubItems(9) = Trim("1")
            End Select
            vQuery = "exec dbo.USP_AP_SearchSourceModule " & vDocType & ",'" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vBookCode = Trim(vRecordset.Fields("bookcode").Value)
                vApCode = Trim(vRecordset.Fields("apcode").Value)
                vDocdate = CDate(Day(CDate(Trim(vRecordset.Fields("docdate").Value))) & "/" & Month(CDate(Trim(vRecordset.Fields("docdate").Value))) & "/" & Year(CDate(Trim(vRecordset.Fields("docdate").Value))))
            End If
            vRecordset.Close
            vListItem.SubItems(10) = vBookCode
            vListItem.SubItems(11) = vSource
            vListItem.SubItems(12) = vSaveFrom
            vListItem.SubItems(13) = vApCode
            vListItem.SubItems(14) = vDocdate
        End If
    ElseIf vEdit = 1 Then
        ListView101.ListItems.Item(vLineEdit).Text = Trim(Text102.Text)
        ListView101.ListItems.Item(vLineEdit).SubItems(1) = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
        ListView101.ListItems.Item(vLineEdit).SubItems(2) = Trim(Text103.Text)
        ListView101.ListItems.Item(vLineEdit).SubItems(3) = Format(Trim(Text104.Text), "##,##0.00")
        ListView101.ListItems.Item(vLineEdit).SubItems(4) = Format(Trim(Text105.Text), "##,##0.00")
        ListView101.ListItems.Item(vLineEdit).SubItems(5) = Format(Trim(Text106.Text), "##,##0.00")
        ListView101.ListItems.Item(vLineEdit).SubItems(6) = Format(Trim(Text107.Text), "##,##0.00")
        ListView101.ListItems.Item(vLineEdit).SubItems(7) = Trim(Text108.Text)
        ListView101.ListItems.Item(vLineEdit).SubItems(8) = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
        Select Case CMB101.Text
        Case Trim("ขอคืนภาษีได้")
            ListView101.ListItems.Item(vLineEdit).SubItems(9) = Trim("0")
        Case Trim("ขอคืนภาษีไม่ได้")
            ListView101.ListItems.Item(vLineEdit).SubItems(9) = Trim("1")
        End Select

        vLineEdit = 0
        vEdit = 0
    End If
    Text102.Text = ""
    Text103.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text107.Text = ""
    Text108.Text = ""
    DTPicker101 = Now
    DTPicker102 = Now
    Text102.SetFocus
    Text104.Text = Format("7", "##,##0.00")
    Text107.Text = Format("0", "##,##0.00")
    CMB101.Text = CMB101.List(0)
Else
    MsgBox "ต้องกรอกข้อมูลในช่องที่มีข้อความสีน้ำเงินให้ครบด้วย", vbCritical, "Send Error"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocType As Integer
Dim i As Integer
Dim vSaveFrom As Integer
Dim vDocNo As String
Dim vBookCode As String
Dim vSource As Integer
Dim vDocdate As String
Dim vTaxDate As String
Dim vTaxDate2 As String
Dim vTaxNo As String
Dim vApCode As String
Dim vShortTaxDesc As String
Dim vTaxRate As Currency
Dim vBeforeTaxAmount As Currency
Dim vTaxAmount As Currency
Dim vExceptTaxAmount As Currency
Dim vLineNumber As Integer
Dim vReturnTax As Integer
Dim vTaxGroup As String
Dim vIsCancel As Integer
Dim vVatRecordset As New ADODB.Recordset
Dim vCount As Integer
Dim vExistDocNo As Integer
Dim vProcess As Integer


If ListView101.ListItems.Count <> 0 Then
    vDocNo = UCase(Trim(Text101.Text))
    If Opt101.Value = True Then
        vDocType = 1
    ElseIf Opt102.Value = True Then
        vDocType = 2
    ElseIf Opt103.Value = True Then
        vDocType = 3
    ElseIf Opt104.Value = True Then
        vDocType = 4
    End If
    vIsCancel = 0
    
    On Error GoTo ErrDescription
    
    vQuery = "begin tran"
    gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_AP_DeleteInputTax '" & vDocNo & "' "
    gConnection.Execute vQuery
    
    For i = 1 To ListView101.ListItems.Count
        vLineNumber = i - 1
        vTaxDate = Trim(ListView101.ListItems.Item(i).SubItems(1))
        vTaxDate2 = Trim(ListView101.ListItems.Item(i).SubItems(8))
        vTaxNo = Trim(ListView101.ListItems.Item(i).Text)
        vShortTaxDesc = Trim(ListView101.ListItems.Item(i).SubItems(2))
        vTaxRate = CCur(Trim(ListView101.ListItems.Item(i).SubItems(3)))
        vBeforeTaxAmount = CCur(Trim(ListView101.ListItems.Item(i).SubItems(4)))
        vTaxAmount = CCur(Trim(ListView101.ListItems.Item(i).SubItems(5)))
        vExceptTaxAmount = Trim(ListView101.ListItems.Item(i).SubItems(6))
        vReturnTax = Trim(ListView101.ListItems.Item(i).SubItems(9))
        vTaxGroup = Trim(ListView101.ListItems.Item(i).SubItems(7))
        vBookCode = Trim(ListView101.ListItems.Item(i).SubItems(10))
        vSource = Trim(ListView101.ListItems.Item(i).SubItems(11))
        vSaveFrom = Trim(ListView101.ListItems.Item(i).SubItems(12))
        vApCode = Trim(ListView101.ListItems.Item(i).SubItems(13))
        vDocdate = Trim(ListView101.ListItems.Item(i).SubItems(14))
                
        vQuery = "exec dbo.USP_AP_InsertInputTax " & vSaveFrom & ",'" & vDocNo & "','" & vBookCode & "', " _
        & " " & vSource & ",'" & vDocdate & "','" & vTaxDate & "','" & vTaxDate2 & "','" & vTaxNo & "','" & vApCode & "', " _
        & " '" & vShortTaxDesc & "'," & vTaxRate & "," & vBeforeTaxAmount & "," & vTaxAmount & "," & vExceptTaxAmount & ", " _
        & " " & vLineNumber & "," & vReturnTax & ",'" & vTaxGroup & "'," & vIsCancel & ",'" & vUserID & "' "
        gConnection.Execute vQuery
    Next i
    
    
    vQuery = "exec dbo.USP_AP_InsertTaxPurchaseLogs '" & vDocNo & "','" & vDocType & "','" & vUserID & "' "
    gConnection.Execute vQuery
    

    
    
'ปรับข้อมูลเลขที่ใบกำกับภาษีหลังร้าน
    
    'If vDocType = 1 Then
     ' vQuery = "select count(docno) as vCount from solar.bcvat.dbo.BCAPOTHERDEBT where docno = '" & vDocNo & "' "
      'If OpenDataBaseBCVat(vVatConnection, vVatRecordset, vQuery) <> 0 Then
       ' vCount = vVatRecordset.Fields("vcount").Value
      'End If
      'vVatRecordset.Close
    'ElseIf vDocType = 2 Then
     ' vQuery = "select count(docno) as vCount from solar.bcvat.dbo.BCOTHEREXPENSE where docno = '" & vDocNo & "' "
      'If OpenDataBaseBCVat(vVatConnection, vVatRecordset, vQuery) <> 0 Then
       ' vCount = vVatRecordset.Fields("vcount").Value
      'End If
      'vVatRecordset.Close
    'ElseIf vDocType = 3 Then
     ' vQuery = "select count(docno) as vCount from solar.bcvat.dbo.BCSTKREFUND where docno = '" & vDocNo & "' "
      'If OpenDataBaseBCVat(vVatConnection, vVatRecordset, vQuery) <> 0 Then
       ' vCount = vVatRecordset.Fields("vcount").Value
      'End If
      'vVatRecordset.Close
    'ElseIf vDocType = 4 Then
     ' vQuery = "select count(docno) as vCount from solar.bcvat.dbo.BCAPINVOICE where docno = '" & vDocNo & "' "
      'If OpenDataBaseBCVat(vVatConnection, vVatRecordset, vQuery) <> 0 Then
       ' vCount = vVatRecordset.Fields("vcount").Value
      'End If
      'vVatRecordset.Close
    'End If
    
    If vCount > 0 Then
      
    'vQuery = "delete solar.bcvat.dbo.bcinputtax where docno = '" & vDocNo & "' "
    'vVatConnection.Execute vQuery


    'vQuery = "select count(docno) as vCount from solar.bcvat.dbo.BCSTKREFUND where docno = '" & vDocNo & "' "
    'If OpenDataBaseBCVat(vVatConnection, vVatRecordset, vQuery) <> 0 Then
     ' vExistDocNo = vVatRecordset.Fields("vcount").Value
    'End If
    'vVatRecordset.Close
          
    If vExistDocNo > 0 Then
      vProcess = 2
    End If
    
    If vExistDocNo = 0 Then
      vProcess = 1
    End If
        
    For i = 1 To ListView101.ListItems.Count
        vLineNumber = i - 1
        vTaxDate = Trim(ListView101.ListItems.Item(i).SubItems(1))
        vTaxDate2 = Trim(ListView101.ListItems.Item(i).SubItems(8))
        vTaxNo = UCase(Trim(ListView101.ListItems.Item(i).Text))
        vShortTaxDesc = Trim(ListView101.ListItems.Item(i).SubItems(2))
        vTaxRate = CCur(Trim(ListView101.ListItems.Item(i).SubItems(3)))
        vBeforeTaxAmount = CCur(Trim(ListView101.ListItems.Item(i).SubItems(4)))
        vTaxAmount = CCur(Trim(ListView101.ListItems.Item(i).SubItems(5)))
        vExceptTaxAmount = Trim(ListView101.ListItems.Item(i).SubItems(6))
        vReturnTax = Trim(ListView101.ListItems.Item(i).SubItems(9))
        vTaxGroup = Trim(ListView101.ListItems.Item(i).SubItems(7))
        vBookCode = Trim(ListView101.ListItems.Item(i).SubItems(10))
        vSource = Trim(ListView101.ListItems.Item(i).SubItems(11))
        vSaveFrom = Trim(ListView101.ListItems.Item(i).SubItems(12))
        vApCode = Trim(ListView101.ListItems.Item(i).SubItems(13))
        vDocdate = Trim(ListView101.ListItems.Item(i).SubItems(14))
                
        'vQuery = "set dateformat dmy"
        'vVatConnection.Execute vQuery
        
        vQuery = "insert into solar.bcvat.dbo.BCInputTax (SaveFrom,DocNo,BookCode,Source,DocDate,TaxDate,TaxDate2," _
        & " TaxNo,ApCode,ShortTaxDesc,TaxRate,Process,BeforeTaxAmount,TaxAmount,ExceptTaxAmount,LineNumber,ReturnTax,TaxGroup,CreatorCode, " _
        & " CreateDateTime,IsCancel)   values (" & vSaveFrom & ",'" & vDocNo & "','" & vBookCode & "', " _
        & " " & vSource & ",'" & vDocdate & "','" & vTaxDate & "','" & vTaxDate2 & "','" & vTaxNo & "','" & vApCode & "', " _
        & " '" & vShortTaxDesc & "'," & vTaxRate & "," & vProcess & "," & vBeforeTaxAmount & "," & vTaxAmount & "," & vExceptTaxAmount & ", " _
        & " " & vLineNumber & "," & vReturnTax & ",'" & vTaxGroup & "','" & vUserID & "',getdate(),0) "
        'vVatConnection.Execute vQuery
    Next i
    
    If vExistDocNo > 0 Then
    vQuery = "update solar.bcvat.dbo.BCSTKRefund set taxno = '" & vTaxNo & "' where docno = '" & vDocNo & "' "
    'vVatConnection.Execute vQuery
    End If
    
    Else
      MsgBox "การเปลี่ยนข้อมูลเลขที่ใบกำกับภาษีด้านหลังร้านไม่สามารถทำได้ เนื่องจากยังไม่มีข้อมูลของเอกสารเลขที่ " & vDocNo & " นี้ในฐานข้อมูลหลังร้าน", vbCritical, "Send Error"
    End If

    MsgBox "บันทึกข้อมูล ภาษีซื้อ ของเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว ", vbInformation, "Send Message"
    
    vQuery = "commit tran"
    gConnection.Execute vQuery
    
    DTPicker101 = Now
    DTPicker102 = Now
    Opt101.Value = True
    Opt102.Value = False
    CMB101.AddItem Trim("ขอคืนภาษีได้")
    CMB101.AddItem Trim("ขอคืนภาษีไม่ได้")
    CMB101.Text = CMB101.List(0)
    Text104.Text = Format("7", "##,##0.00")
    Text107.Text = Format("0", "##,##0.00")
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text108.Text = ""
    ListView101.ListItems.Clear
    Text102.Enabled = False
Else
    vDocNo = Trim(Text101.Text)
    vQuery = "exec dbo.USP_AP_DeleteInputTax '" & vDocNo & "' "
    gConnection.Execute vQuery
    MsgBox "บันทึกข้อมูล ภาษีซื้อ ของเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว ", vbInformation, "Send Message"
    DTPicker101 = Now
    DTPicker102 = Now
    Opt101.Value = True
    Opt102.Value = False
    CMB101.AddItem Trim("ขอคืนภาษีได้")
    CMB101.AddItem Trim("ขอคืนภาษีไม่ได้")
    CMB101.Text = CMB101.List(0)
    Text104.Text = Format("7", "##,##0.00")
    Text107.Text = Format("0", "##,##0.00")
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text108.Text = ""
    ListView101.ListItems.Clear
    Text102.Enabled = True
    Text104.Enabled = True
    Text105.Enabled = True
    Text106.Enabled = True
    Text107.Enabled = True
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    vQuery = "rollback tran"
    gConnection.Execute vQuery
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
On Error Resume Next

DTPicker101 = Now
DTPicker102 = Now
Opt101.Value = True
Opt102.Value = False
CMB101.AddItem Trim("ขอคืนภาษีได้")
CMB101.AddItem Trim("ขอคืนภาษีไม่ได้")
CMB101.Text = CMB101.List(0)
Text104.Text = Format("7", "##,##0.00")
Text107.Text = Format("0", "##,##0.00")
Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
Text105.Text = ""
Text106.Text = ""
Text108.Text = ""
ListView101.ListItems.Clear
Text102.Enabled = True
Text104.Enabled = True
Text105.Enabled = True
Text106.Enabled = True
Text107.Enabled = True
Text101.SetFocus
End Sub

Private Sub Form_Load()
'InitializeDataBaseVat
DTPicker101 = Now
DTPicker102 = Now
Opt101.Value = True
Opt102.Value = False
CMB101.AddItem Trim("ขอคืนภาษีได้")
CMB101.AddItem Trim("ขอคืนภาษีไม่ได้")
CMB101.Text = CMB101.List(0)
Text104.Text = Format("7", "##,##0.00")
Text107.Text = Format("0", "##,##0.00")
End Sub

Private Sub ListView101_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim vAnswer As Integer
Dim vDocNo As String

On Error GoTo ErrDescription

If KeyCode = 46 Then
    i = ListView101.SelectedItem.Index
    vDocNo = Trim(ListView101.SelectedItem.Text)
    vAnswer = MsgBox("ต้องการลบ เลขที่ใบกำกับภาษีเลขที่ " & vDocNo & " ใช่หรือไม่", vbYesNo, "Send Question ?")
    If vAnswer = 6 Then
        ListView101.ListItems.Remove (i)
    Else
        Exit Sub
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub ListView101_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next

If ListView101.ListItems.Count <> 0 Then
    If Button = 2 Then
        Popup_Menu Menu1
    End If
End If
End Sub


Private Sub mDelete_Click()
Dim i As Integer
Dim vAnswer As Integer
Dim vDocNo As String

On Error GoTo ErrDescription

i = ListView101.SelectedItem.Index
vDocNo = Trim(ListView101.SelectedItem.Text)
vAnswer = MsgBox("ต้องการลบ เลขที่ใบกำกับภาษีเลขที่ " & vDocNo & " ใช่หรือไม่", vbYesNo, "Send Question ?")
If vAnswer = 6 Then
    ListView101.ListItems.Remove (i)
Else
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub mEdit_Click()
On Error Resume Next

vEdit = 1
vLineEdit = ListView101.SelectedItem.Index
Text102.Text = ListView101.ListItems.Item(vLineEdit).Text
Text103.Text = ListView101.ListItems.Item(vLineEdit).SubItems(2)
DTPicker101.Value = CDate(ListView101.ListItems.Item(vLineEdit).SubItems(1))
Text104.Text = Format(ListView101.ListItems.Item(vLineEdit).SubItems(3), "##,##0.00")
Text105.Text = Format(ListView101.ListItems.Item(vLineEdit).SubItems(4), "##,##0.00")
Text106.Text = Format(ListView101.ListItems.Item(vLineEdit).SubItems(5), "##,##0.00")
Text107.Text = Format(ListView101.ListItems.Item(vLineEdit).SubItems(6), "##,##0.00")
Text108.Text = ListView101.ListItems.Item(vLineEdit).SubItems(7)
DTPicker102.Value = CDate(ListView101.ListItems.Item(vLineEdit).SubItems(8))
Select Case ListView101.ListItems.Item(vLineEdit).SubItems(7)
Case 0
    CMB101.Text = CMB101.List(0)
Case 1
    CMB101.Text = CMB101.List(1)
End Select

Text104.Enabled = False
Text105.Enabled = False
Text106.Enabled = False
Text107.Enabled = False

End Sub

Private Sub Opt101_Click()
On Error Resume Next

If Opt101.Value = True Then
    Pic101.Visible = True
    Text101.SetFocus
    DTPicker101 = Now
    DTPicker102 = Now
    CMB101.AddItem Trim("ขอคืนภาษีได้")
    CMB101.AddItem Trim("ขอคืนภาษีไม่ได้")
    CMB101.Text = CMB101.List(0)
    Text104.Text = Format("7", "##,##0.00")
    Text107.Text = Format("0", "##,##0.00")
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text108.Text = ""
    ListView101.ListItems.Clear
    Text102.Enabled = False
End If
End Sub

Private Sub Opt102_Click()
On Error Resume Next
If Opt102.Value = True Then
    Pic101.Visible = True
    Text101.SetFocus
    DTPicker101 = Now
    DTPicker102 = Now
    CMB101.AddItem Trim("ขอคืนภาษีได้")
    CMB101.AddItem Trim("ขอคืนภาษีไม่ได้")
    CMB101.Text = CMB101.List(0)
    Text104.Text = Format("7", "##,##0.00")
    Text107.Text = Format("0", "##,##0.00")
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text108.Text = ""
    ListView101.ListItems.Clear
    Text102.Enabled = False
End If
End Sub

Private Sub Opt103_Click()
On Error Resume Next
If Opt103.Value = True Then
    Pic101.Visible = True
    Text101.SetFocus
    DTPicker101 = Now
    DTPicker102 = Now
    CMB101.AddItem Trim("ขอคืนภาษีได้")
    CMB101.AddItem Trim("ขอคืนภาษีไม่ได้")
    CMB101.Text = CMB101.List(0)
    Text104.Text = Format("7", "##,##0.00")
    Text107.Text = Format("0", "##,##0.00")
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text108.Text = ""
    ListView101.ListItems.Clear
    Text102.Enabled = False
End If
End Sub

Private Sub Opt104_Click()
On Error Resume Next
If Opt104.Value = True Then
    Pic101.Visible = True
    Text101.SetFocus
    DTPicker101 = Now
    DTPicker102 = Now
    CMB101.AddItem Trim("ขอคืนภาษีได้")
    CMB101.AddItem Trim("ขอคืนภาษีไม่ได้")
    CMB101.Text = CMB101.List(0)
    Text104.Text = Format("7", "##,##0.00")
    Text107.Text = Format("0", "##,##0.00")
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text108.Text = ""
    ListView101.ListItems.Clear
    Text102.Enabled = False
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vDocType As Integer
Dim vCount As Integer
Dim vListItem As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Opt101.Value = True Then
        vDocType = 1
    ElseIf Opt102.Value = True Then
        vDocType = 2
    ElseIf Opt103.Value = True Then
        vDocType = 3
    ElseIf Opt104.Value = True Then
        vDocType = 4
    End If
    vDocNo = Trim(Text101.Text)
    vQuery = "exec dbo.USP_AP_SearchDocNo " & vDocType & ",'" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCount = Trim(vRecordset.Fields("vcount").Value)
    End If
    vRecordset.Close
    If vCount = 0 Then
        MsgBox "ไม่มีเลขที่เอกสาร " & Text101.Text & " ในระบบ โปรดตรวจสอบ", vbCritical, "Send Error"
        Text101.Text = ""
        ListView101.ListItems.Clear
        Text101.SetFocus
        Text102.Enabled = False
        Exit Sub
    Else
        Text102.Enabled = True
            ListView101.ListItems.Clear
            vQuery = "exec dbo.USP_AP_SelectInputTaxDetails '" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("taxno").Value))
            vListItem.SubItems(1) = CDate(Trim(vRecordset.Fields("taxdate").Value))
            vListItem.SubItems(2) = Trim(vRecordset.Fields("shorttaxdesc").Value)
            vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("taxrate").Value), "##,##0.00")
            vListItem.SubItems(4) = Format(Trim(vRecordset.Fields("beforetaxamount").Value), "##,##0.00")
            vListItem.SubItems(5) = Format(Trim(vRecordset.Fields("taxamount").Value), "##,##0.00")
            vListItem.SubItems(6) = Format(Trim(vRecordset.Fields("excepttaxamount").Value), "##,##0.00")
            vListItem.SubItems(7) = Trim(vRecordset.Fields("taxgroup").Value)
            vListItem.SubItems(8) = CDate(Trim(vRecordset.Fields("taxdate2").Value))
            vListItem.SubItems(9) = Trim(vRecordset.Fields("returntax").Value)
            vListItem.SubItems(10) = Trim(vRecordset.Fields("bookcode").Value)
            vListItem.SubItems(11) = Trim(vRecordset.Fields("source").Value)
            vListItem.SubItems(12) = Trim(vRecordset.Fields("savefrom").Value)
            vListItem.SubItems(13) = Trim(vRecordset.Fields("apcode").Value)
            vListItem.SubItems(14) = Trim(vRecordset.Fields("docdate").Value)
            vRecordset.MoveNext
            Wend
            End If
            vRecordset.Close
            Text104.Enabled = True
            Text105.Enabled = True
            Text106.Enabled = True
            Text107.Enabled = True
        Text102.SetFocus
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Text104_LostFocus()
If Trim(Text105.Text) <> "" Then
    Text106.Text = Format((CCur(Text104.Text) * CCur(Text105.Text)) / 100, "##,##0.00")
End If
End Sub

Private Sub Text105_LostFocus()
Dim vCheckText As String

On Error GoTo ErrDescription

vCheckText = Trim(Text105.Text)
Call CheckNumber(vCheckText)
If vCheckValueNumber = False Then
    Text105.Text = ""
    MsgBox "ต้องใส่ตัวเลขเท่านั้น", vbCritical, "Send Error"
Else
    Text105.Text = Format(Text105.Text, "##,##0.00")
    Text106.Text = Format((CCur(Text104.Text) * CCur(Text105.Text)) / 100, "##,##0.00")
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub
Private Sub Text106_LostFocus()
Dim vCheckText As String

On Error GoTo ErrDescription

vCheckText = Trim(Text106.Text)
Call CheckNumber(vCheckText)
If vCheckValueNumber = False Then
    Text106.Text = ""
    MsgBox "ต้องใส่ตัวเลขเท่านั้น", vbCritical, "Send Error"
Else
    Text106.Text = Format(Text106.Text, "##,##0.00")
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Text107_LostFocus()
Dim vCheckText As String

On Error GoTo ErrDescription

vCheckText = Trim(Text107.Text)
Call CheckNumber(vCheckText)
If vCheckValueNumber = False Then
    Text107.Text = ""
    MsgBox "ต้องใส่ตัวเลขเท่านั้น", vbCritical, "Send Error"
Else
    Text107.Text = Format(Text107.Text, "##,##0.00")
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Popup_Menu(m As Menu)
    Menu1.Visible = True
    PopupMenu m, 2
    Menu1.Visible = False
End Sub
