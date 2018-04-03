VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form102 
   Caption         =   "เพิ่มและพิมพ์คูปอง"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   0
      Top             =   5715
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
   Begin VB.Frame Frame2 
      Caption         =   "พิมพ์คูปอง"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   810
      TabIndex        =   8
      Top             =   4365
      Width           =   10545
      Begin VB.TextBox Text104 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   7245
         TabIndex        =   17
         Top             =   2205
         Width           =   3120
      End
      Begin VB.CommandButton CMD104 
         Caption         =   "พิมพ์คูปอง"
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
         Left            =   9270
         TabIndex        =   15
         Top             =   2610
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   1815
         Left            =   2385
         TabIndex        =   14
         Top             =   270
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   3201
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสคูปอง"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อฟอร์มคูปอง"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ที่อยู่คูปอง"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "ชื่อคูปองที่พิมพ์ :"
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
         Left            =   5355
         TabIndex        =   16
         Top             =   2205
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "รายการคูปอง :"
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
         Left            =   945
         TabIndex        =   13
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "เพิ่มทะเบียนคูปอง"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   810
      TabIndex        =   7
      Top             =   405
      Width           =   10545
      Begin VB.CheckBox Check101 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ยกเลิก คูปอง"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   5580
         TabIndex        =   21
         Top             =   3150
         Width           =   1365
      End
      Begin VB.CommandButton CMD103 
         Caption         =   "ล้างหน้าจอ"
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
         Left            =   4320
         TabIndex        =   20
         Top             =   3150
         Width           =   1095
      End
      Begin VB.TextBox Text103 
         Appearance      =   0  'Flat
         Height          =   870
         Left            =   2385
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2205
         Width           =   4560
      End
      Begin VB.CommandButton CMD101 
         Height          =   330
         Left            =   5040
         Picture         =   "Form402.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   540
         Width           =   330
      End
      Begin VB.ListBox List101 
         Appearance      =   0  'Flat
         Height          =   1980
         Left            =   2385
         TabIndex        =   18
         Top             =   945
         Visible         =   0   'False
         Width           =   2985
      End
      Begin MSComCtl2.DTPicker DTPicker102 
         Height          =   330
         Left            =   2385
         TabIndex        =   4
         Top             =   1800
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         Format          =   20774913
         CurrentDate     =   38709
      End
      Begin MSComCtl2.DTPicker DTPicker101 
         Height          =   330
         Left            =   2385
         TabIndex        =   3
         Top             =   1350
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         Format          =   20774913
         CurrentDate     =   38709
      End
      Begin VB.TextBox Text102 
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
         Left            =   2385
         MaxLength       =   2
         TabIndex        =   2
         Top             =   945
         Width           =   1500
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
         Left            =   2385
         TabIndex        =   0
         Top             =   540
         Width           =   2625
      End
      Begin VB.CommandButton CMD102 
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
         Height          =   420
         Left            =   3060
         TabIndex        =   6
         Top             =   3150
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "หมายเหตุ แสดงรายละเอียดคูปอง:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   90
         TabIndex        =   19
         Top             =   2205
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "วันสิ้นสุดคูปอง :"
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
         Left            =   540
         TabIndex        =   12
         Top             =   1800
         Width           =   1725
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "วันเริ่มใช้คูปอง :"
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
         TabIndex        =   11
         Top             =   1350
         Width           =   1635
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "หัวเลขที่คูปอง :"
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
         TabIndex        =   10
         Top             =   945
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ชื่อคูปอง :"
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
         Left            =   675
         TabIndex        =   9
         Top             =   540
         Width           =   1590
      End
   End
End
Attribute VB_Name = "Form102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer

List101.Clear
List101.Visible = True
vQuery = "select * from npmaster.dbo.tb_pm_couponmaster where iscancel = 0 order by id "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        List101.AddItem (Trim(vRecordset.Fields("couponname").Value))
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As Integer
Dim vHeaderNo As String
Dim vCouponName As String
Dim vStartDate As Date
Dim vEndDate As Date
Dim vMydescription As String
Dim vCheckHeader As Integer
Dim vIsCancel As String

If Text101.Text <> "" And Text102.Text <> "" Then
    vStartDate = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
    vEndDate = Trim(DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year)
    vMydescription = Trim(Text103.Text)
    If vStartDate > vEndDate Then
        MsgBox "วันที่เริ่มใช้คูปองต้องน้อยกว่า วันที่หมดอายุของคูปอง", vbCritical, "Send Error"
        Exit Sub
    End If
        
    If vIsOpen1 = 0 Then
        vQuery = "select isnull(max(id),0)+1 as MaxID from npmaster.dbo.tb_pm_couponmaster"
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vID = Trim(vRecordset.Fields("maxid").Value)
        End If
        vRecordset.Close
        vHeaderNo = UCase(Trim(Text102.Text))
        vCouponName = Trim(Text101.Text)
        
        vQuery = "select headerno from npmaster.dbo.tb_pm_couponmaster  where headerno = '" & vHeaderNo & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckHeader = 1
        Else
            vCheckHeader = 0
        End If
        vRecordset.Close
        If vCheckHeader = 0 Then
            vQuery = "exec dbo.USP_PM_InsertCouponMaster 0," & vID & ",'" & vHeaderNo & "','" & vCouponName & "','" & vStartDate & "','" & vEndDate & "','0','" & vMydescription & "','" & vUserID & "' "
            gConnection.Execute (vQuery)
        Else
            MsgBox "กรุณาตรวจสอบ หัวคูปอง เพราะมีอยู่แล้ว ", vbCritical, "Send Error"
            Exit Sub
        End If
        MsgBox "บันทึกข้อมูลเรียบร้อยแล้ว", vbInformation, "Send Message"
    Else
        vCouponName = Trim(Text101.Text)
        vQuery = "select ID as CPID from npmaster.dbo.tb_pm_couponmaster where couponname = '" & vCouponName & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vID = Trim(vRecordset.Fields("cpid").Value)
        End If
        vRecordset.Close
        vHeaderNo = UCase(Trim(Text102.Text))
    
        vQuery = "select headerno from npmaster.dbo.tb_pm_couponmaster  where headerno = '" & vHeaderNo & "' and id <> " & vID & " "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckHeader = 1
        Else
            vCheckHeader = 0
        End If
        vRecordset.Close
        
        If Check101.Value = 1 Then
            vIsCancel = 1
        Else
            vIsCancel = 0
        End If
        
        If vCheckHeader = 0 Then
            vQuery = "exec dbo.USP_PM_InsertCouponMaster 1," & vID & ",'" & vHeaderNo & "','" & vCouponName & "','" & vStartDate & "','" & vEndDate & "','" & vIsCancel & "','" & vMydescription & "','" & vUserID & "' "
            gConnection.Execute (vQuery)
        Else
            MsgBox "กรุณาตรวจสอบ หัวคูปอง เพราะมีอยู่แล้ว ", vbCritical, "Send Error"
            Exit Sub
        End If
        MsgBox "ปรับปรุงข้อมูลเรียบร้อยแล้ว", vbInformation, "Send Message"
    End If
    
    vIsOpen1 = 0
    Text101.Enabled = True
    DTPicker101.Value = Now
    DTPicker102.Value = Now
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    ListView101.ListItems.Clear
    Text104.Text = ""
End If
End Sub

Private Sub CMD103_Click()
vIsOpen1 = 0
Text101.Enabled = True
DTPicker101.Value = Now
DTPicker102.Value = Now
Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
ListView101.ListItems.Clear
Text104.Text = ""
End Sub

Private Sub CMD104_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCouponName As String
Dim vCPID As Integer

If Text104.Text <> "" Then
    vCPID = ListView101.ListItems.Item(ListView101.SelectedItem.Index)
    vQuery = "select pathname from npmaster.dbo.TB_PM_CouponName where cpid = " & vCPID & " "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vCouponName = Trim(vRecordset.Fields("pathname").Value)
    End If
    vRecordset.Close
    
    With Crystal101
    .ReportFileName = Trim(vCouponName)
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    End With
    
Else
    MsgBox "กรุณาเลือก ชื่อของคูปองที่อยู่ในตารางด้วย", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Private Sub Form_Load()
DTPicker101.Value = Now
DTPicker102.Value = Now
End Sub

Private Sub List101_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As Integer
Dim vCouponName As String
Dim vCouponList As ListItem

List101.Visible = False
Text101.Enabled = False
vIsOpen1 = 1
Text101.Text = List101.Text
vCouponName = Trim(Text101.Text)
vQuery = "select ID,HeaderNo,CouponName,StartDate,EndDate,IsCancel,isnull(MyDescription,'') as MyDescription from npmaster.dbo.tb_pm_couponmaster where couponname = '" & vCouponName & "' and iscancel = '0' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    Text102.Text = Trim(vRecordset.Fields("headerno").Value)
    DTPicker101.Value = Trim(vRecordset.Fields("startdate").Value)
    DTPicker102.Value = Trim(vRecordset.Fields("enddate").Value)
    vID = Trim(vRecordset.Fields("id").Value)
    Text103.Text = Trim(vRecordset.Fields("mydescription").Value)
Else
    MsgBox "ไม่มีข้อมูลคูปองนี้อยู่ในระบบ กรุณาตรวจสอบ", vbInformation, "Send Information"
End If
vRecordset.Close

vQuery = "exec dbo.USP_PM_SelectCoupon " & vID & " "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    ListView101.ListItems.Clear
    While Not vRecordset.EOF
    Set vCouponList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("cpid").Value))
    vCouponList.SubItems(1) = Trim(vRecordset.Fields("name").Value)
    vCouponList.SubItems(2) = Trim(vRecordset.Fields("pathname").Value)
    vRecordset.MoveNext
    Wend
Else
MsgBox "ยังไม่ได้สร้างฟอร์ม กรุณาแจ้งแผนกคอมพิวเตอร์สร้างฟอร์มด้วย", vbInformation, "Send Information"
End If
vRecordset.Close

End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
If ListView101.ListItems.Count <> 0 Then
    Text104.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(1))
End If
End Sub
