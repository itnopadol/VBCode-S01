VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form106 
   Caption         =   "ยิงบาร์โค้ด ทำใบหยิบสินค้า"
   ClientHeight    =   8760
   ClientLeft      =   2655
   ClientTop       =   450
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form106.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   13260
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CMBSaleCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   1080
      Width           =   3075
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   3690
      Top             =   7065
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
   Begin MSComctlLib.ListView ListView102 
      Height          =   1545
      Left            =   6795
      TabIndex        =   13
      Top             =   1485
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   2725
      View            =   3
      LabelEdit       =   1
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "คลัง"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชั้นเก็บ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "จำนวนคงเหลือ"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "หน่วยนับ"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.TextBox Text106 
      Appearance      =   0  'Flat
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
      Left            =   1755
      TabIndex        =   3
      Top             =   6660
      Width           =   1860
   End
   Begin VB.CheckBox Check101 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ไม่ใส่จำนวน"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4860
      TabIndex        =   5
      Top             =   1485
      Width           =   1680
   End
   Begin VB.TextBox Text105 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1485
      TabIndex        =   1
      Top             =   3105
      Width           =   1455
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "พิมพ์"
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
      Height          =   420
      Left            =   2655
      TabIndex        =   4
      Top             =   7065
      Width           =   960
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2850
      Left            =   675
      TabIndex        =   2
      Top             =   3645
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
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1485
      TabIndex        =   0
      Top             =   1485
      Width           =   1950
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ระบุพนักงานขาย :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   -405
      TabIndex        =   27
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label LBLPickZone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6795
      TabIndex        =   25
      Top             =   3150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label LBLZoneID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4860
      TabIndex        =   24
      Top             =   1170
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label LBLFamilyCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4860
      TabIndex        =   23
      Top             =   3150
      Width           =   1680
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FamilyCode :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3150
      TabIndex        =   22
      Top             =   3150
      Width           =   1680
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "คลัง :"
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
      Left            =   585
      TabIndex        =   21
      Top             =   2700
      Width           =   870
   End
   Begin VB.Label Text108 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1485
      TabIndex        =   20
      ToolTipText     =   "เลือกโดย Click ที่รายการจำนวนคงเหลือตามคลัง"
      Top             =   2700
      Width           =   1455
   End
   Begin VB.Label Text107 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4860
      TabIndex        =   19
      ToolTipText     =   "เลือกโดย Click ที่รายการจำนวนคงเหลือตามคลัง"
      Top             =   2700
      Width           =   1680
   End
   Begin VB.Label Text103 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1485
      TabIndex        =   18
      Top             =   2295
      Width           =   5055
   End
   Begin VB.Label Text104 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4860
      TabIndex        =   17
      Top             =   1890
      Width           =   1680
   End
   Begin VB.Label Text102 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1485
      TabIndex        =   16
      Top             =   1890
      Width           =   1950
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชั้นเก็บ :"
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
      Left            =   3870
      TabIndex        =   15
      Top             =   2700
      Width           =   960
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Click เลือกชั้นเก็บที่จะขาย"
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
      Left            =   6795
      TabIndex        =   14
      Top             =   1215
      Width           =   4425
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F5 = บันทึกข้อมูลและพิมพ์เอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6795
      TabIndex        =   12
      Top             =   3150
      Width           =   4965
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
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
      Left            =   675
      TabIndex        =   11
      Top             =   6705
      Width           =   1140
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ต้องการจำนวน :"
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
      Left            =   -135
      TabIndex        =   10
      Top             =   3150
      Width           =   1590
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "หน่วยนับ :"
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
      Left            =   3915
      TabIndex        =   9
      Top             =   1890
      Width           =   915
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อสินค้า :"
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
      Left            =   630
      TabIndex        =   8
      Top             =   2295
      Width           =   825
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสสินค้า :"
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
      Left            =   270
      TabIndex        =   7
      Top             =   1890
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "บาร์โค้ด :"
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
      Left            =   270
      TabIndex        =   6
      Top             =   1485
      Width           =   1185
   End
End
Attribute VB_Name = "Form106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSelectLineItem As Integer

Private Sub Check101_Click()
Text101.SetFocus
End Sub

Public Sub SaveData()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vBarCode As String
Dim vItemCode As String
Dim vItemName As String
Dim vUnitCode As String
Dim i As Integer
Dim vMaxNumber  As String
Dim vDocNo As String
Dim vQty As Double
Dim vRefNo As String
Dim vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vDocDate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vARCode As String
Dim vNamePrint As String
Dim vQueueID As Integer
Dim vHead As String
Dim vJobID As String
Dim vModuleID As String
Dim vCompanyName As String
Dim vReportID As Integer
Dim vReportType As String
Dim vPrintStatus As Integer
Dim vCount  As Integer
Dim vQueue As String
Dim vFamilyCode As String
Dim vPickZone As String
Dim vSaleCode As String


On Error GoTo ErrDescription

If Me.CMBSaleCode.Text = "" Then
MsgBox "กรุณากรอก รหัสพนักงานขายก่อนบันทึกข้อมูลหยิบสินค้า กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Me.CMBSaleCode.SetFocus
Exit Sub
End If

If ListView101.ListItems.Count > 0 Then
    vQuery = "exec dbo.USP_MB_SearchRunningNumber 26"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vMaxNumber = Trim(vRecordset.Fields("autonumber").Value)
        vHeader = Trim(vRecordset.Fields("header").Value)
        vHead = Trim(vRecordset.Fields("docnumber").Value)
    End If
    vRecordset.Close
    
    vDocNo = vMaxNumber
    vRefNo = vHead & vHeader & "-" & Format(vMaxNumber, "0000")
    vNamePrint = Trim(vUserID)
    'vWHCode = Trim("S01")
    vDocType = 2
    vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = "1"
    vJobID = "01"
    vModuleID = "MB"
    vCompanyName = ""
    vReportID = 313
    vReportType = Trim("INV")
    vPrintStatus = 0
    vSaleCode = Left(Trim(CMBSaleCode.Text), InStr(Trim(CMBSaleCode.Text), "-") - 1)
    
    vQuery = "exec dbo.USP_MB_SearchMobileDocument '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCount = Trim(vRecordset.Fields("vCount").Value)
    End If
    vRecordset.Close
    
    If vCount > 0 Then
      MsgBox "มีเลขที่เอกสาร " & vDocNo & " นี้อยู่แล้ว กรุณา กด F5 ใหม่อีกครั้งเพื่อรันเลขที่เอกสารเลขที่ใหม่", vbCritical, "Send Error"
      Exit Sub
    End If
    
    vQuery = "exec dbo.USP_MB_InsertMobileDocumentMaster  '" & vDocNo & "','" & vDocDate & "','" & vRefNo & "','" & vSaleCode & "' "
    gConnection.Execute vQuery
        
    For i = 1 To ListView101.ListItems.Count
        vWHCode = Trim(ListView101.ListItems.Item(i).SubItems(5))
        vShelfCode = Trim(ListView101.ListItems.Item(i).SubItems(6))
        vBarCode = Trim(ListView101.ListItems.Item(i).Text)
        vItemCode = Trim(ListView101.ListItems.Item(i).SubItems(1))
        vItemName = Trim(ListView101.ListItems.Item(i).SubItems(3))
        vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(4))
        vQty = CCur(Format(Trim(ListView101.ListItems.Item(i).SubItems(2)), "##,##0.00"))
        vFamilyCode = Trim(ListView101.ListItems.Item(i).SubItems(7))
        vZoneID = Trim(ListView101.ListItems.Item(i).SubItems(8))
        vPickZone = Trim(ListView101.ListItems.Item(i).SubItems(9))
        
        vQuery = "exec bcnp.dbo.usp_MB_InsertMobileDocument3 '" & vDocNo & "','" & vWHCode & "','" & vShelfCode & "','" & vItemCode & "','" & vBarCode & "','" & vItemName & "'," & vQty & ",'" & vUnitCode & "','" & vUserID & "','" & vRefNo & "','" & vZoneID & "','" & vFamilyCode & "','" & vPickZone & "' "
        gConnection.Execute vQuery
    Next i

    vQuery = "exec bcnp.dbo.USP_MB_UpdateRunningNumber 26"
    gConnection.Execute vQuery
        
    
    Dim vCountShelf As Integer
    Dim n As Integer
    Dim vWHGroupID As String
    Dim vShelfGroupID As String
    Dim vWHGroup() As String
    Dim vShelfGroup() As String
    Dim vFamilyGroup() As String
    Dim vZoneGroup() As String
    Dim vPickZoneGroup() As String
    
    vQuery = "exec dbo.USP_MB_SearchGroupOfShelf2 '" & vRefNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vCountShelf = vRecordset.RecordCount
       ReDim vFamilyGroup(vCountShelf) As String
       ReDim vZoneGroup(vCountShelf) As String
       ReDim vShelfGroup(vCountShelf) As String
       ReDim vWHGroup(vCountShelf) As String
       ReDim vPickZoneGroup(vCountShelf) As String
       
       n = 1
       vRecordset.MoveFirst
       While Not vRecordset.EOF
       vZoneGroup(n) = vRecordset.Fields("zoneid").Value
       vFamilyGroup(n) = vRecordset.Fields("familygroup").Value
       vWHGroup(n) = vRecordset.Fields("whcode").Value
       vShelfGroup(n) = vRecordset.Fields("shelfcode").Value
       vPickZoneGroup(n) = vRecordset.Fields("pickzone").Value
       n = n + 1
       vRecordset.MoveNext
       Wend
    End If
    vRecordset.Close
    
     If vCountShelf > 0 Then
     Dim j As Integer
     
        For j = 1 To vCountShelf
                
        vQuery = "exec dbo.USP_MB_SearchRunningNumber 27"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vQueueID = Trim(vRecordset.Fields("autonumber").Value)
            vQueue = Trim(vRecordset.Fields("autonumber").Value)
        End If
        vRecordset.Close

        vQuery = "begin tran"
        gConnection.Execute vQuery
                
        vQuery = "exec dbo.USP_NP_InsertDataQueueManagement3 '" & vQueueID & "','" & vDocDate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vRefNo & "','" & vDocNo & "','" & vWHCode & "','" & vShelfGroup(j) & "','" & vZoneGroup(j) & "','" & vFamilyGroup(j) & "','" & vPickZoneGroup(j) & "',1,0"
        gConnection.Execute vQuery
              
        'vQuery = "exec dbo.USP_NP_InsertNPPrintQueue '" & vJobID & "','" & vZoneID & "','" & vModuleID & "','" & vCompanyName & "','" & vQueueID & "'," & vReportID & ",'" & vReportType & "'," & vPrintStatus & ",'" & vUserID & "' "
        'gConnection.Execute vQuery
    
        MsgBox "บันทึกข้อมูลเรียบร้อยแล้วครับ ได้คิวที่ " & vQueueID & "  กรุณาติดตามคิวให้ลูกค้าด้วย", vbInformation, "ข้อความแจ้งเตือน"
                
        vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
        gConnection.Execute vQuery
        

        
        'Call PrintHeader(vQueue, vShelfGroup)
        'Call PrintDetails(vRefNo, vZoneGroup(j), vFamilyGroup(j), vPickZoneGroup(j))
        
        
        vJobID = 2
        vQuery = "exec dbo.USP_NP_InsertPrintTermal " & vJobID & ",'" & vRefNo & "','" & vQueueID & "','" & vWHCode & "','" & vShelfGroup(j) & "','" & vFamilyGroup(j) & "','" & vZoneGroup(j) & "','" & vPickZoneGroup(j) & "','" & vUserID & "' "
        gConnection.Execute vQuery
        
        'If vWHCode = "S1-B" Then
         '   vJobID = 4
            
          '  vQuery = "exec dbo.USP_NP_InsertPrintTermal " & vJobID & ",'" & vRefNo & "','" & vQueueID & "','" & vWHCode & "','" & vShelfGroup(j) & "','" & vFamilyGroup(j) & "','" & vZoneGroup(j) & "','" & vPickZoneGroup(j) & "','" & vUserID & "' "
           ' gConnection.Execute vQuery
        'End If
        
        vQuery = "commit tran"
        gConnection.Execute vQuery
     
     Next j
     ListView101.ListItems.Clear
     End If

End If

ErrDescription:
If Err.Description <> "" Then
  MsgBox Err.Description
  vQuery = "rollback tran"
  gConnection.Execute vQuery
    If Err.Number = -2147217873 Then
        vQuery = "exec dbo.USP_MB_UpdateRunningNumber 27"
        gConnection.Execute vQuery
    End If
  Exit Sub
End If

End Sub

Private Sub CMBSaleCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    If ListView101.ListItems.Count > 0 Then
        Call SaveData
    End If
End If

If KeyCode = 8 Then
   Call ClearData
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vDocDate As Date
Dim vZoneID As String
Dim vDocType As Integer
Dim vShelfGroup As String
Dim vExistPicking As String
Dim vCheckPicking As Integer
Dim vCountPicking As Integer
Dim vNamePrint As String
Dim vWHCode As String
Dim vARCode As String
Dim vRefNo As String

On Error GoTo ErrDescription

If Text106.Text <> "" Then
     
    vDocNo = Trim(Text106.Text)
    vNamePrint = Trim(vUserID)
    vZoneID = Trim("02")
    vWHCode = Trim("014")
    vDocType = 2
    vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vARCode = "1"
    vShelfGroup = UCase(Trim("PKA"))
    
    
    vQuery = "exec dbo.USP_NP_SearchPickingRequest '" & vDocNo & "','" & vShelfGroup & "','" & vDocDate & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckPicking = 1
      vExistPicking = Trim(vRecordset.Fields("docno").Value)
      vCountPicking = Trim(vRecordset.Fields("countpicking").Value) + 1
      vRefNo = Trim(vRecordset.Fields("saleorderno").Value)
    Else
      vCheckPicking = 0
      vCountPicking = 1
      vRefNo = ""
    End If
    vRecordset.Close
                                                                                                               
    'vQuery = "exec dbo.USP_NP_InsertDataQueueManagement '" & vDocno & "','" & vDocDate & "'," & vDocType & ",'" & vARCode & "','" & vNamePrint & "','" & vRefNo & "','" & vDocno & "','" & vWHCode & "','" & vShelfGroup & "','" & vZoneID & "'," & vCountPicking & ",0 "
    'gConnection.Execute vQuery
    
    
    'Call PrintHeader(vDocNo, vShelfGroup)
    'Call PrintDetails(vDocNo, vShelfGroup)
        
    MsgBox "เลขที่ใบหยิบได้เข้าคิวการหยิบสินค้าเรียบร้อยแล้ว", vbInformation, "ข้อความแจ้งเตือน"
    
    Text106.Text = ""
    CMD102.Enabled = False
    Text101.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD102_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    If ListView101.ListItems.Count > 0 Then
        Call SaveData
    End If
End If
End Sub

Private Sub Form_Load()

ListView101.ColumnHeaders(3).Alignment = lvwColumnRight
ListView102.ColumnHeaders(2).Alignment = lvwColumnRight
Call GetSaleCode
End Sub

Private Sub ListView101_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If ListView101.ListItems.Count <> 0 Then
    If KeyCode = 46 Then
        i = ListView101.SelectedItem.Index
        ListView101.ListItems.Remove (i)
    End If
End If

If KeyCode = 116 Then
    If ListView101.ListItems.Count > 0 Then
        Call SaveData
    End If
End If
End Sub


Public Sub GetSaleCode()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset


CMBSaleCode.Clear
vQuery = "select * from vw_NP_SaleUserID "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMBSaleCode.AddItem Trim(vRecordset.Fields("salename").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub

Private Sub ListView102_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vZoneID As String
Dim vFamilyGroup As String
Dim vPickZone As String

Dim vItemCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vUnitCode As String
Dim i As Integer
Dim vCheckItemCode As String
Dim vCheckWHCode As String
Dim vCheckShelfCode As String

On Error GoTo ErrDescription

If Me.ListView102.ListItems.Count > 0 Then
   vItemCode = Me.Text102.Caption
   vUnitCode = Me.Text104.Caption
   vWHCode = Me.ListView102.ListItems(Me.ListView102.SelectedItem.Index).Text
   vShelfCode = Me.ListView102.ListItems(Me.ListView102.SelectedItem.Index).SubItems(1)
   
   'vQuery = "exec dbo.usp_MB_CheckBarcodePickZone '" & vItemCode & "','" & vShelfCode & "' "
   vQuery = "exec dbo.usp_MB_CheckBarcodePickZone_New '" & vItemCode & "','" & vUnitCode & "','" & vWHCode & "','" & vShelfCode & "' "
           If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vFamilyGroup = Trim(vRecordset.Fields("familygroup").Value)
            vZoneID = Trim(vRecordset.Fields("zoneid").Value)
            vPickZone = Trim(vRecordset.Fields("pickzone").Value)
        End If
        vRecordset.Close

        Me.LBLZoneID.Caption = vZoneID
        Me.LBLFamilyCode.Caption = vFamilyGroup
        Me.LBLPickZone.Caption = vPickZone
   
   If Left(UCase(vWHCode), 2) = "S2" Then
        MsgBox "คลังที่สามารถขายได้ ต้องเป็นคลัง S1 เท่านั้น กรุณาตรวจสอบ", vbCritical, "Send Error Message"
        Me.ListView102.SetFocus
        Exit Sub
   End If
   
   If vWHCode <> "S1-A" And vWHCode <> "S1-B" And vWHCode <> "S1-SHW" And vWHCode <> "S1-SPO" Then
      MsgBox "ชั้นเก็บที่สามารถขายได้ ต้องเป็นคลังในส่วนของ S1-A,S1-B,S1-SHW,S1-SPO เท่านั้น ส่วนยอดที่มีอยู่ชั้นเก็บอื่น ถ้าต้องการขายต้องโอนเข้าคลัง  S1-A,S1-SHW,S1-SPO ก่อน", vbCritical, "Send Message Error"
      Me.ListView102.SetFocus
      Me.Text107.Caption = ""
      Me.Text108.Caption = ""
   Else
      Me.Text107.Caption = vShelfCode
      Me.Text108.Caption = vWHCode
      
      For i = 1 To Me.ListView101.ListItems.Count
             vCheckItemCode = Me.ListView101.ListItems(i).SubItems(1)
             vCheckWHCode = Me.ListView101.ListItems(i).SubItems(5)
             vCheckShelfCode = Me.ListView101.ListItems(i).SubItems(6)

             If vItemCode = vCheckItemCode And vWHCode = vCheckWHCode And vShelfCode = vCheckShelfCode Then
                Me.Text105.Text = Format(Me.ListView101.ListItems(i).SubItems(2), "##,##0")
                vSelectLineItem = i
                GoTo Line1
             End If
      Next i
Line1:
      Me.Text105.SetFocus
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListView102_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    If ListView101.ListItems.Count > 0 Then
        Call SaveData
    End If
End If
End Sub

Private Sub ListView102_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vZoneID As String
Dim vFamilyGroup As String
Dim vPickZone As String

Dim vItemCode As String
Dim vWHCode As String
Dim vShelfCode As String
Dim i As Integer
Dim vCheckItemCode As String
Dim vCheckWHCode As String
Dim vCheckShelfCode As String

Dim vUnitCode As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
If Me.ListView102.ListItems.Count > 0 Then
   vItemCode = Me.Text102.Caption
   vWHCode = Me.ListView102.ListItems(Me.ListView102.SelectedItem.Index).Text
   vShelfCode = Me.ListView102.ListItems(Me.ListView102.SelectedItem.Index).SubItems(1)

   vUnitCode = Me.Text104.Caption

   
    vQuery = "exec dbo.usp_MB_CheckBarcodePickZone_New '" & vItemCode & "','" & vUnitCode & "','" & vWHCode & "','" & vShelfCode & "' "
         If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          vFamilyGroup = Trim(vRecordset.Fields("familygroup").Value)
          vZoneID = Trim(vRecordset.Fields("zoneid").Value)
          vPickZone = Trim(vRecordset.Fields("pickzone").Value)
      End If
      vRecordset.Close
    
      Me.LBLZoneID.Caption = vZoneID
      Me.LBLFamilyCode.Caption = vFamilyGroup
      Me.LBLPickZone.Caption = vPickZone
        
   If vWHCode <> "S1-A" And vWHCode <> "S1-B" And vWHCode <> "S1-SHW" And vWHCode <> "S1-SPO" Then
      MsgBox "คลังที่สามารถขายได้ ต้องเป็นคลังในส่วนของ S1-A,S1-SPO และ S1-SHW  เท่านั้น ส่วนยอดที่มีอยู่ชั้นเก็บอื่น ถ้าต้องการขายต้องโอนเข้าคลัง S1-A,S1-SPO และ S1-SHW  ก่อน", vbCritical, "Send Message Error"
      Me.ListView102.SetFocus
      Me.Text107.Caption = ""
      Me.Text108.Caption = ""
   Else
      Me.Text107.Caption = vShelfCode
      Me.Text108.Caption = vWHCode
      
      For i = 1 To Me.ListView101.ListItems.Count
             vCheckItemCode = Me.ListView101.ListItems(i).SubItems(1)
             vCheckWHCode = Me.ListView101.ListItems(i).SubItems(5)
             vCheckShelfCode = Me.ListView101.ListItems(i).SubItems(6)

             If vItemCode = vCheckItemCode And vWHCode = vCheckWHCode And vShelfCode = vCheckShelfCode Then
                Me.Text105.Text = Format(Me.ListView101.ListItems(i).SubItems(2), "##,##0")
                vSelectLineItem = i
                GoTo Line1
             End If
      Next i
Line1:
      Me.Text105.SetFocus
   End If
End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text101_Change()
If Text101.Text = "" Then
Call ClearData
End If
End Sub

Private Sub Text101_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    If ListView101.ListItems.Count > 0 Then
        Call SaveData
    End If
End If

If KeyCode = 8 Then
   Call ClearData
End If

End Sub

Public Sub ClearData()
Text102.Caption = ""
Text103.Caption = ""
Text104.Caption = ""
Text105.Text = ""
Text107.Caption = ""
Text108.Caption = ""
Text101.SetFocus
ListView102.ListItems.Clear
vSelectLineItem = 0
End Sub
Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vBarCode As String
Dim vItemCode As String
Dim vItemName As String
Dim vUnitCode As String
Dim vZoneID As String
Dim vFamilyGroup As String
Dim vPickZone As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Text101.Text <> "" Then
        vBarCode = Trim(Text101.Text)
        'vQuery = "exec bcnp.dbo.usp_MB_CheckBarcode '" & vBarCode & "' "
        'vQuery = "exec bcnp.dbo.usp_MB_ItemBarcodePickZone '" & vBarCode & "' "
        vQuery = "exec dbo.usp_MB_ItemBarcodePickZoneUnitCode '" & vBarCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vItemCode = Trim(vRecordset.Fields("itemcode").Value)
            vItemName = Trim(vRecordset.Fields("itemname").Value)
            vUnitCode = Trim(vRecordset.Fields("unitcode").Value)
            vFamilyGroup = Trim(vRecordset.Fields("familygroup").Value)
            vZoneID = Trim(vRecordset.Fields("zoneid").Value)
            vPickZone = Trim(vRecordset.Fields("pickzone").Value)
        Else
            MsgBox "ไม่มีรหัสบาร์โค้ด " & vBarCode & "  นี้ หรือไม่ถูกต้องตามกระบวนการ กรุณาตรวจสอบ", vbCritical, "Send Error"
            Call ClearData
            Text101.Text = ""
            Text101.SetFocus
            Exit Sub
        End If
        vRecordset.Close
        ListView102.ListItems.Clear
        Call vCheckStockOnHand
        Text102.Caption = vItemCode
        Text103.Caption = vItemName
        Text104.Caption = vUnitCode
        Me.LBLZoneID.Caption = vZoneID
        Me.LBLFamilyCode.Caption = vFamilyGroup
        Me.LBLPickZone.Caption = vPickZone
        
        If Check101.Value = 0 Then
            ListView102.SetFocus
        Else
            Call InsertToGrid
        End If
    Else
        MsgBox "ข้อมูลไม่ครบถ้วน หารหัสสินค้าไม่เจอ", vbCritical, "Send Error"
        Call ClearData
        Exit Sub
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text102_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    If ListView101.ListItems.Count > 0 Then
        Call SaveData
    End If
End If
End Sub

Private Sub Text105_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    If ListView101.ListItems.Count > 0 Then
        Call SaveData
    End If
End If
End Sub

Private Sub Text105_KeyPress(KeyAscii As Integer)
On Error GoTo ErrDescription

If KeyAscii = 13 Then
    Call InsertToGrid
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub InsertToGrid()
Dim vBarCode As String
Dim vItemCode As String
Dim vItemName As String
Dim vUnitCode As String
Dim vQty As Double
Dim vListBarCode As ListItem
Dim vCheck014 As Integer
Dim vCheckWareHouse As String
Dim vWHCode As String
Dim vCheckQty As Double
Dim vSelectShelf As String
Dim vCheckSelectShelf As String
Dim i As Integer
Dim vZoneID As String
Dim vFamilyGroup As String
Dim vPickZone As String
Dim vAnswer As Integer

If Text101.Text <> "" And Text102.Caption <> "" And Text103.Caption <> "" And Text104.Caption <> "" And Me.Text107.Caption <> "" And Me.Text108.Caption <> "" Then
    vBarCode = Trim(Text101.Text)
    vItemCode = Trim(Text102.Caption)
    vItemName = Trim(Text103.Caption)
    vUnitCode = Trim(Text104.Caption)
    If Check101.Value = 0 Then
        vQty = CCur(Trim(Text105.Text))
    Else
        vQty = 0
    End If
    vZoneID = Me.LBLZoneID.Caption
    vFamilyGroup = Me.LBLFamilyCode.Caption
    vPickZone = Me.LBLPickZone.Caption
    
    If Check101.Value = 0 Then
    vWHCode = Me.Text108.Caption
    vSelectShelf = Me.Text107.Caption
        For i = 1 To ListView102.ListItems.Count
            vCheckWareHouse = Trim(ListView102.ListItems.Item(i).Text)
            vCheckSelectShelf = Trim(ListView102.ListItems.Item(i).SubItems(1))
            If vCheckWareHouse = vWHCode And vCheckSelectShelf = vSelectShelf Then
                vCheckQty = Trim(ListView102.ListItems.Item(i).SubItems(2))
                GoTo Line1
            End If
        Next i
        
Line1:
        If vQty = 0 Then
            MsgBox "ต้องกรอก จำนวนที่ต้องการขายมากกว่า 0 เสมอ", vbCritical, "Send Error Message"
            Exit Sub
        End If
        If vQty > vCheckQty Then
            vAnswer = MsgBox("จำนวนสินค้าที่ต้องการขาย มากกว่า จำนวนสินค้าคงเหลือของคลัง S01  และชั้นเก็บ " & vSelectShelf & " คุณต้องการขายตามจำนวนนี้หรือไม่ ?", vbYesNo, "Send Message")
            If vAnswer = 7 Then
                Me.Text105.SetFocus
                Exit Sub
            End If
        End If
    End If
    If vSelectLineItem = 0 Then
    Set vListBarCode = ListView101.ListItems.Add(, , Trim(vBarCode))
    vListBarCode.SubItems(1) = vItemCode
    vListBarCode.SubItems(2) = Format(vQty, "##,##0.00")
    vListBarCode.SubItems(3) = vItemName
    vListBarCode.SubItems(4) = vUnitCode
    vListBarCode.SubItems(5) = vWHCode
    vListBarCode.SubItems(6) = vSelectShelf
    vListBarCode.SubItems(7) = vFamilyGroup
    vListBarCode.SubItems(8) = vZoneID
    vListBarCode.SubItems(9) = vPickZone
    Else
        ListView101.ListItems(vSelectLineItem).SubItems(2) = Format(vQty, "##,##0.00")
    End If
    
    Text101.Text = ""
    Text102.Caption = ""
    Text103.Caption = ""
    Text104.Caption = ""
    Text105.Text = ""
    Text107.Caption = ""
    Text108.Caption = ""
    Me.LBLFamilyCode.Caption = ""
    Me.LBLPickZone.Caption = ""
    Me.LBLZoneID.Caption = ""
    Text101.SetFocus
    ListView102.ListItems.Clear
    vSelectLineItem = 0
Else
   MsgBox "ต้องกรอกข้อมูลสินค้าให้ครบ ดังนี้ บาร์โค้ด คลัง และชั้นเก็บที่ขาย จำนวนที่ต้องการขาย ถึงจะบันทึกลงตารางได้", vbCritical, "Send Error Message"
End If

End Sub

Private Sub Text106_Change()
On Error Resume Next
CMD102.Enabled = True
End Sub

Private Sub Text106_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 116 Then
    If ListView101.ListItems.Count > 0 Then
        Call SaveData
        'Call Cmd102_Click
    End If
End If
End Sub

Public Sub vCheckStockOnHand()
Dim vBarCode As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim ListItem As ListItem

On Error Resume Next

vBarCode = Trim(Text101.Text)
vQuery = "set dateformat dmy"
gConnection.Execute vQuery
vQuery = "exec dbo.USP_MB_ShowQTYOnHand '" & vBarCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set ListItem = ListView102.ListItems.Add(, , Trim(vRecordset.Fields("whcode").Value))
    ListItem.SubItems(1) = Trim(vRecordset.Fields("shelfcode").Value)
    ListItem.SubItems(2) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
    ListItem.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub

Public Sub PrintHeader1()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

If Text106.Text <> "" Then
    vDocNo = Trim(Text106.Text)
    vRepID = 310
    vRepType = "MB"
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
    With Crystal101
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@vDocno;" & vDocNo & ";true"
    .Destination = crptToPrinter
    .WindowState = crptMaximized
    .Action = 1
    End With
End If
End Sub

Public Sub PrintDetails1()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

vDocNo = Trim(Text106.Text)
vRepID = 289
vRepType = "MB"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close
With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vDocno;" & vDocNo & ";true"
.Destination = crptToPrinter
.WindowState = crptMaximized
.Action = 1
End With

End Sub

Public Sub PrintHeader(vDocNo As String, vShelfGroup As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocDate As Date
Dim vPrinterName As String
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

On Error Resume Next

'vPrinterName = Trim("\\diy01\TM-Mobile")
vPrinterName = Trim("TM_Moo")
For Each printerObj In Printers
  If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
    Set Printer = printerObj
    Set printerObj = Nothing
    Exit For
  End If
Next


vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vQuery = "exec dbo.USP_SO_PickingQueuePos '" & vDocNo & "','" & vDocDate & "','" & vShelfGroup & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    If Not vRecordset.EOF Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.FontBold = True
      Printer.CurrentX = 1800
      Printer.CurrentY = 0
      Printer.Print Trim(vRecordset.Fields("queueno").Value)
      
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 1400
      Printer.CurrentY = 1000
      Printer.Print "*" & Trim(vRecordset.Fields("queueno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("_______________________________________________________________________________________")
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า")
      
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
      Printer.CurrentX = 2200
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value)
      

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 3150
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value)
End If
End If
vRecordset.Close

    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 100
    Printer.Print Trim("______________________________________________________________________________________________")
    
    Printer.CurrentX = Printer.CurrentX + 2000
    Printer.Print Trim("วันที่พิมพ์ :") & Now

      Printer.EndDoc
End Sub

Public Sub PrintDetails(vDocNo As String, vZoneGroup As String, vFamilyGroup As String, vPickZoneGroup As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocDate As Date
Dim vPrinterName As String
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

On Error Resume Next

If vPickZoneGroup = "01" Then
vPrinterID = 0
End If

If vPickZoneGroup = "02" Then
vPrinterID = 1
End If

If vPickZoneGroup = "03" Then
vPrinterID = 2
End If

vQuery = "exec dbo.USP_NP_SearchCheckPrinter " & vPrinterID & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vPrinterName = Trim(vRecordset.Fields("printername").Value)
End If
vRecordset.Close



For Each printerObj In Printers
If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
Set Printer = printerObj
Set printerObj = Nothing
Exit For
End If
Next

    vDocDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
    vQuery = "exec dbo.USP_SO_PickingQueuePos3 '" & vDocNo & "','" & vDocDate & "' ,'" & vZoneGroup & "','" & vFamilyGroup & "','" & vPickZoneGroup & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    If Not vRecordset.EOF Then
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.FontBold = True
      Printer.CurrentX = 1700
      Printer.CurrentY = 0
      Printer.Print Trim(vRecordset.Fields("queueno").Value)
      
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
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 1500
      Printer.CurrentY = 1650
      Printer.Print Trim("ต้นฉบับใบจัดสินค้า")
      
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
      Printer.CurrentX = 2200
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 12
      Printer.CurrentX = 0
      Printer.CurrentY = 3150
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value)
      
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
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ขายชั้นเก็บ :" & Trim(vRecordset.Fields("MBShelfCode").Value) & "       " & Trim("OnHand: ") & Trim(vRecordset.Fields("qtyonhand").Value) & "       " & Trim("รวมคลัง : ") & "  " & Trim(vRecordset.Fields("stkwhcode").Value) & "    " & Trim(vRecordset.Fields("unitcode").Value)
                                      
          Printer.Font.Size = 18
          Printer.FontBold = True
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)
          
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
          Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value)
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ชื่อสินค้า :" & Trim(vRecordset.Fields("itemname").Value)
          
          Printer.CurrentX = Printer.CurrentX + 15
          Printer.CurrentY = Printer.CurrentY + 50
          Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")
          
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
    Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 150
    Printer.Print "         _____________                                    ______________"
          
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    
    Printer.CurrentX = Printer.CurrentX + 2000
    Printer.Print Trim("วันที่พิมพ์ :") & Now & "          " & vPrinterName

    Printer.EndDoc
End Sub

