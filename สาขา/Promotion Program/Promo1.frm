VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form101 
   Caption         =   "กำหนดทะเบียนโปรโมชั่น (สำหรับ Admin)"
   ClientHeight    =   8985
   ClientLeft      =   1545
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
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
      Height          =   6090
      Left            =   300
      TabIndex        =   8
      Top             =   300
      Width           =   11415
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
         Top             =   5310
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
         Top             =   5310
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
         Top             =   5310
         Visible         =   0   'False
         Width           =   1230
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7515
         Top             =   4005
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
         Top             =   5025
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
         Top             =   5325
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
         Top             =   5310
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
         Top             =   5325
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
         Top             =   4650
         Width           =   2865
      End
      Begin VB.TextBox Text103 
         Appearance      =   0  'Flat
         Height          =   1890
         Left            =   1650
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2625
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker DTPicker102 
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Top             =   2100
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61210625
         CurrentDate     =   38504
      End
      Begin MSComCtl2.DTPicker DTPicker101 
         Height          =   315
         Left            =   1650
         TabIndex        =   3
         Top             =   1575
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   -2147483628
         Format          =   61210625
         CurrentDate     =   38504
      End
      Begin VB.TextBox Text102 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Top             =   1050
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
         Top             =   2625
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
         Top             =   2100
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
         Top             =   1575
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
         Top             =   1050
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

Private Sub CMDGenBarcode_Click()
Dim vQuery As String
Dim i As Integer
Dim vRecordset As New Recordset
Dim vBarCode As String
Dim vName As String
Dim vQty As Integer
Dim vBarCode1 As String
Dim vTop As Integer
Dim n As Integer


'vQuery = "select * ,(select top 1 id from npmaster.dbo.tb_barcodecharnpaiboon order by barcode order by id desc) as topid from npmaster.dbo.tb_barcodecharnpaiboon order by barcode"
vQuery = "select * ,isnull((select top 1 id from npmaster.dbo.tb_barcodecharnpaiboonlineitem order by id desc),0) as topid from npmaster.dbo.tb_barcodecharnpaiboon order by barcode"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst
    vTop = vRecordset.Fields("topid").Value
    n = vTop
While Not vRecordset.EOF
    vBarCode = Trim(vRecordset.Fields("barcode").Value)
    vName = Trim(vRecordset.Fields("name1").Value)
    vQty = Trim(vRecordset.Fields("qty").Value)
    vBarCode1 = Left(vBarCode, 1) + " " + Left(Right(vBarCode, 12), 6) + " " + Right(vBarCode, 6)
    For i = 1 To vQty
    n = n + 1
    vQuery = "insert into npmaster.dbo.TB_BarcodeCharnPaiBoonLineItem (barcode,name1,barcode1,id) values ('" & vBarCode & "','" & vName & "','" & vBarCode1 & "'," & n & ")"
    gConnection.Execute vQuery
    Next i
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub

Private Sub Command108_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim i As Integer, j As Integer
Dim vItemCode As String, vBarCode As String

On Error GoTo ErrDescription

vQuery = "select * into dbo.Report_Temp3 From NP_LABEL_TEMP where UsedUser = 'Null'"
gConnection.Execute vQuery
 
For i = 1 To 500
vItemCode = "51" & Format(i, "00000")
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

On Error GoTo ErrDescription

If Text102.Text <> "" Then
    vPromoname = Trim(Text102.Text)
    vDescriptionPromo = Trim(Text103.Text)
    vStartPromo = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
    vEndPromo = DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year
    vCheckNow = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    vIsCancel = 0
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
    vQuery = "exec  USP_PM_InsertOrUpdateMaster '" & vCheckJob & "','" & vNEwDocno & "','" & vPromoname & "','" & vStartPromo & "','" & vEndPromo & "','" & vDescriptionPromo & "'," & vIsCancel & ",'" & vUserID & "' "
    gConnection.Execute (vQuery)
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
Dim i As Integer

For i = 456 To 705
vCode = "CNY2" & Trim(Format(i, "000"))
vName = Trim("Chinese New Year 2556")
vCouVal = 10
vBegDate = "01/02/2013"
vEndDate = "28/02/2013"
vQuery = "exec bcnp.dbo.usp_np_InsertPosCoupon '" & vCode & "','" & vName & "','" & vBegDate & "','" & vEndDate & "'," & vCouVal & " "
gConnection.Execute vQuery
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
End If

vCheckJob = 1
DTPicker101 = Now
DTPicker102 = Now
End Sub
