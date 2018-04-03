VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form101 
   Caption         =   "กำหนดทะเบียนโปรโมชั่น (สำหรับ Admin)"
   ClientHeight    =   8985
   ClientLeft      =   2985
   ClientTop       =   450
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   959
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "กำหนดทะเบียนโปรโมชั่น"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7170
      Left            =   300
      TabIndex        =   8
      Top             =   270
      Width           =   13800
      Begin VB.CommandButton CMDExpertCard 
         Caption         =   "Expert Card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10350
         TabIndex        =   23
         Top             =   5850
         Width           =   1230
      End
      Begin VB.ComboBox CMBPromoType 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1665
         TabIndex        =   21
         Text            =   "CMBPromoType"
         Top             =   1035
         Width           =   2805
      End
      Begin VB.CommandButton CMDGenBarcode 
         Caption         =   "GenBarcode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9000
         TabIndex        =   20
         Top             =   5850
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton Command108 
         Caption         =   "Gen Assets"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7425
         TabIndex        =   19
         Top             =   5850
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command107 
         Caption         =   "Gen Coupon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   18
         Top             =   5850
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CheckBox Check102 
         Caption         =   "บันทึกแล้วไม่ล้างหน้าจอ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1650
         TabIndex        =   17
         Top             =   5565
         Width           =   3615
      End
      Begin VB.CommandButton Command106 
         Caption         =   "พิมพ์เอกสาร"
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
         Left            =   3150
         TabIndex        =   16
         Top             =   5865
         Width           =   1215
      End
      Begin VB.CommandButton Command105 
         Caption         =   "เคลียร์หน้าจอ"
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
         Left            =   4650
         TabIndex        =   15
         Top             =   5850
         Width           =   1215
      End
      Begin VB.CommandButton Command104 
         Height          =   315
         Left            =   3750
         Picture         =   "Promo1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   525
         Width           =   315
      End
      Begin VB.CommandButton Command101 
         Height          =   315
         Left            =   4125
         Picture         =   "Promo1.frx":0357
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   525
         Width           =   315
      End
      Begin VB.CommandButton Command103 
         Caption         =   "บันทึก"
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
         Left            =   1650
         TabIndex        =   6
         Top             =   5865
         Width           =   1215
      End
      Begin VB.CheckBox Check101 
         Caption         =   "ยกเลิก (ทำเครื่องหมาย = ยกเลิก)"
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
         Height          =   315
         Left            =   1650
         TabIndex        =   14
         Top             =   5190
         Width           =   2865
      End
      Begin VB.TextBox Text103 
         Appearance      =   0  'Flat
         Height          =   1890
         Left            =   1650
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   3165
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker DTPicker102 
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Top             =   2640
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   38504
      End
      Begin MSComCtl2.DTPicker DTPicker101 
         Height          =   315
         Left            =   1650
         TabIndex        =   3
         Top             =   2115
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   -2147483628
         Format          =   16646145
         CurrentDate     =   38504
      End
      Begin VB.TextBox Text102 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Top             =   1590
         Width           =   4215
      End
      Begin VB.TextBox Text101 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   7
         Top             =   525
         Width           =   2040
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "ประเภทโปรโมชั่น "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   22
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "รายละเอียด เพิ่มเติม"
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
         Left            =   75
         TabIndex        =   13
         Top             =   3165
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "วันที่สิ้นสุด"
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
         TabIndex        =   12
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "วันที่เริ่ม"
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
         Left            =   750
         TabIndex        =   11
         Top             =   2115
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ชื่อโปรโมชั่น"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   10
         Top             =   1590
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "รหัสโปรโมชั่น"
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
         Left            =   525
         TabIndex        =   9
         Top             =   525
         Width           =   1065
      End
   End
End
Attribute VB_Name = "Form101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNEwDocno As String

Private Sub CMDExpertCard_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer, j As Integer
Dim vExpertCard As String

On Error GoTo ErrDescription
 
'For i = 16150 To 16249
For i = 16350 To 17000
vExpertCard = "A" & Format(i, "00000")

 
'vQuery = "Insert into npmaster.dbo.TB_NP_ExpertCard(ExpertCard)" _
 '               & "  select '" & vExpertCard & "' as vExpertCard"
  vQuery = "insert into npmaster.dbo.TB_MB_Card(Code,Type,UsedStatus,IsCancel,Mydescription,DateStamp,IsRedcard) " _
                    & "values ('" & vExpertCard & "',0,0,0,'Stiker Card',getdate(),0)"
                
                
gConnection.Execute vQuery

Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDGenBarcode_Click()
Dim vQuery As String
Dim i As Integer
Dim vRecordset As New Recordset
Dim vItemCode As String
Dim vBarCode As String
Dim vName As String
Dim vQty As Integer
Dim vBarCode1 As String
Dim vTop As Integer
Dim n As Integer
Dim vPrice As Double


'vQuery = "select * ,(select top 1 id from npmaster.dbo.tb_barcodecharnpaiboon order by barcode order by id desc) as topid from npmaster.dbo.tb_barcodecharnpaiboon order by barcode"
'vQuery = "select * ,isnull((select top 1 id from npmaster.dbo.TB_TempBarCodeCharnPaiboon order by id desc),0) as topid from npmaster.dbo.TB_BarCodeCharnPaiboon order by itemcode"
'If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
'vRecordset.MoveFirst
 '   vTop = vRecordset.Fields("topid").Value
  '  n = vTop
'While Not vRecordset.EOF

    vItemCode = "8851750030715"
    vBarCode = "8851750030715SS"
    vName = "ราวแขวนผ้าอลูมิเนียมกลม75ซม.PixoFS-08-ขาว"
    vQty = 120
    vPrice = 39
    vBarCode1 = Left(vBarCode, 1) + " " + Right(vBarCode, 6)
    For i = 1 To vQty
    n = n + 1
    vQuery = "insert into npmaster.dbo.TB_TempBarCodeCharnPaiboon (itemcode,itemname,barcode,price,id) values ('" & vItemCode & "','" & vName & "','" & vBarCode & "'," & vPrice & "," & n & ")"
    gConnection.Execute vQuery
    Next i
    'vRecordset.MoveNext
   'Wend
'End If
'vRecordset.Close

End Sub

Private Sub Command108_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer, j As Integer
Dim vItemCode As String, vBarCode As String

On Error GoTo ErrDescription

'vQuery = "select * into dbo.Report_Temp3 From NP_LABEL_TEMP where UsedUser = 'Null'"
'gConnection.Execute vQuery
 
For i = 1 To 500
vItemCode = "61" & Format(i, "00000")
vBarCode = vItemCode
 
vQuery = "Insert into bcnp.dbo.Report_Temp3 (itemcode,barcode,type)" _
                & "  select '" & vItemCode & "' as Itemcode,'" & vBarCode & "' as Barcode,1 as Type"
gConnection.Execute vQuery
vQuery = "Insert into bcnp.dbo.Report_Temp3 (itemcode,barcode,type)" _
                & "  select '" & vItemCode & "' as Itemcode,'" & vBarCode & "' as Barcode,1 as Type"
gConnection.Execute vQuery

Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Command101_Click()
MDIForm1.Enabled = False
FormSearchPromotion.Show
End Sub

Private Sub Command103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vPromoname As String
Dim vStartPromo As Date, vEndPromo As Date
Dim vDescriptionPromo As String
Dim vIsCancel As Integer
Dim vCheckNow As Date
Dim vPromoType As Integer

On Error GoTo ErrDescription

If Me.CMBPromoType.Text = "" Then
MsgBox "กรุณาเลือก ประเภทโปรโมชั่น กรุณาตรวจสอบ", vbCritical, "Send Message Error"
Me.CMBPromoType.SetFocus
Exit Sub
End If

If Text102.Text <> "" Then
    vPromoname = Trim(Text102.Text)
    vDescriptionPromo = Trim(Text103.Text)
    vStartPromo = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
    vEndPromo = DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year
    vCheckNow = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    vIsCancel = 0
    vPromoType = Me.CMBPromoType.ListIndex
    
    If vPromoType < 0 Then
    MsgBox "กรุณาเลือก ประเภทโปรโมชั่นที่มีอยู่ไม่สามารถกรอกเองได้ กรุณาตรวจสอบ", vbCritical, "Send Message Error"
    Me.CMBPromoType.SetFocus
    Exit Sub
    End If
    If vStartPromo <= vEndPromo Then
    If vEndPromo >= vCheckNow Then
    If vCheckJob <> 0 Then
    vQuery = "execute USP_PM_MasterNewDocNo"
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vNEwDocno = Trim(vRecordset.Fields("NewDocNo").Value)
    End If
    vRecordset.Close
    Else
        vNEwDocno = Trim(Text101.Text)
    End If
    vQuery = "exec  USP_PM_InsertOrUpdateMaster '" & vCheckJob & "','" & vNEwDocno & "','" & vPromoname & "','" & vStartPromo & "','" & vEndPromo & "','" & vDescriptionPromo & "'," & vIsCancel & ",'" & vUserID & "'," & vPromoType & " "
    gConnection.Execute (vQuery)
    
    
'conn2.Open "Provider=SQLOLEDB.1;Data Source=nebula;Initial Catalog=npmaster;User ID=vbuser;Password=132"
'rs.Open "TB_Images", conn2, adOpenDynamic, adLockOptimistic
'rs.AddNew
'vStream.Type = adTypeBinary
'vStream.Open
'vQuery = "select pmcode from npmaster.dbo.TB_PM_PromotionMaster  where pmcode = '" & vNewDocno & "' "
'If OpenDataBase1(gConnection, vRecordset1, vQuery) <> 0 Then
'vPMCode = Trim(vRecordset1.Fields("pmcode").Value)
'End If
'vRecordset1.Close

'vPicture = "V:\Reports\Promotion\" & vPMCode & ".jpg"
'vStream.LoadFromFile vPicture
'rs.Fields("Images").Value = vStream.Read
'rs.Fields("Prs_No").Value = vPMCode
'rs.Update
'rs.Close
'conn2.Close
                        
    If vCheckJob <> 0 Then
    MsgBox "ได้เอกสารเลขที่ " & vNEwDocno & " "
    Else
    MsgBox "เอกสารเลขที่ " & vNEwDocno & " ได้ทำการปรับปรุงเรียบร้อยแล้ว"
    End If
        If Check102.Value = 0 Then
            Text101.Text = ""
            Text102.Text = ""
            Text103.Text = ""
            DTPicker101 = Now
            DTPicker102 = Now
            Me.CMBPromoType.Text = ""
        Else
            Command103.Enabled = True
            Text101.Text = ""
            DTPicker101 = Now
            DTPicker102 = Now
        End If
        Check102.Value = 0
    Else
        MsgBox "วันที่เริ่มโปรโมชั่น ไม่สามารถเริ่มได้ ณ วันที่น้อยกว่าวันที่ทำเอกสาร"
        Exit Sub
    End If
    Else
        MsgBox "วันที่เริ่มโปรโมชั่น ควรจะน้อยกว่าวันที่สิ้นสุดโปรโมชั่น "
        Exit Sub
    End If
Else
    MsgBox "กรุณาใส่ ชื่อโปรโมชั่น ด้วยครับ"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub Command104_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

vCheckJob = 1
Command103.Enabled = True
vQuery = "execute USP_PM_MasterNewDocNo"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    Text101.Text = Trim(vRecordset.Fields("NewDocNo").Value)
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Command105_Click()
Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
DTPicker101.Value = Now
DTPicker102.Value = Now
End Sub

Private Sub Command106_Click()
MsgBox "ยังพิมพ์ไม่ได้ครับ"
End Sub

Private Sub Command107_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCode As String
Dim vName As String
Dim vCouVal As Currency
Dim vBegDate As Date
Dim vEndDate As Date
Dim i As Double

'Call InitializeDatabaseBranch


'อย่าตั้งชื่อยาว พิมพ์บาร์อ่านไม่ออก

For i = 1 To 730
vCode = "S257-T" & Trim(Format(i, "0000"))
vName = Trim("TOA2557")
vCouVal = 40
vBegDate = "22/03/2014"
vEndDate = "31/03/2014"

'โอนเข้า POS ส่วนกลาง
vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCode & "','" & vName & "','" & vBegDate & "','" & vEndDate & "'," & vCouVal & " "
gConnection.Execute vQuery

'โอนเข้า POS S01
If Left(vCode, 2) = "S1" Then
vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponS01 '" & vCode & "','" & vName & "','" & vBegDate & "','" & vEndDate & "'," & vCouVal & " "
gConnection.Execute vQuery
End If

'โอนเข้า POS S02
If Left(vCode, 2) = "S2" Then
vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponS02 '" & vCode & "','" & vName & "','" & vBegDate & "','" & vEndDate & "'," & vCouVal & " "
gConnection.Execute vQuery
End If


'vQuery = "exec bcnp.dbo.usp_np_InsertPosCouponTemp '" & vCode & "','" & vName & "','" & vBegDate & "','" & vEndDate & "'," & vCouVal & " "
'gConnection.Execute vQuery
Next i

'For i = 2001 To 3000
'vCode = Trim("C10-" & Format(i, "0000"))
'vName = Trim("คูปองเงินสด ตรุษจีน 2008-16022008")
'vCouVal = 10
'vBegDate = "31/01/2008"
'vEndDate = "23/03/2008"
'vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCode & "','" & vName & "','" & vBegDate & "','" & vEndDate & "'," & vCouVal & " "
'gConnection.Execute vQuery
'Next i

End Sub

Private Sub Form_Load()
If vUserID = "somrod" Then
    Command107.Visible = True
    Command108.Visible = True
    CMDGenBarcode.Visible = True
End If

Call vGetPromoType
vCheckJob = 1
DTPicker101 = Now
DTPicker102 = Now
End Sub

Public Sub vGetPromoType()
Me.CMBPromoType.AddItem ("ปกติ")
Me.CMBPromoType.AddItem ("Clearance Sale")
Me.CMBPromoType.AddItem ("Loss Leader")
End Sub
