VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmOrder205 
   Caption         =   "Form205 กำหนด รถขนส่ง"
   ClientHeight    =   7905
   ClientLeft      =   1245
   ClientTop       =   1140
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "FrmOrder205.frx":0000
   ScaleHeight     =   7905
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
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
      Left            =   3195
      TabIndex        =   32
      Top             =   1395
      Width           =   1410
   End
   Begin VB.CommandButton CMD105 
      Height          =   330
      Left            =   4995
      Picture         =   "FrmOrder205.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "ออก"
      Top             =   7200
      Width           =   330
   End
   Begin VB.TextBox Text114 
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
      Height          =   1185
      Left            =   3195
      TabIndex        =   13
      Top             =   5850
      Width           =   5325
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   330
      Left            =   3195
      TabIndex        =   7
      Top             =   4230
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   174325761
      CurrentDate     =   38696
   End
   Begin VB.CommandButton CMD104 
      Height          =   330
      Left            =   4545
      Picture         =   "FrmOrder205.frx":7665
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "ยกเลิก ข้อมูล"
      Top             =   7200
      Width           =   330
   End
   Begin VB.CommandButton CMD103 
      Height          =   330
      Left            =   4095
      Picture         =   "FrmOrder205.frx":9803
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "ค้นหา ข้อมูล"
      Top             =   7200
      Width           =   330
   End
   Begin VB.CommandButton CMD102 
      Height          =   330
      Left            =   3645
      Picture         =   "FrmOrder205.frx":9BD0
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "บันทึกและปรับปรุงข้อมูล"
      Top             =   7200
      Width           =   330
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   3195
      Picture         =   "FrmOrder205.frx":9EF7
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "เคลียร์ หน้าจอ"
      Top             =   7200
      Width           =   330
   End
   Begin VB.TextBox Text113 
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
      Left            =   5805
      TabIndex        =   12
      Top             =   5445
      Width           =   1185
   End
   Begin VB.TextBox Text112 
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
      Left            =   4500
      TabIndex        =   11
      Top             =   5445
      Width           =   1185
   End
   Begin VB.TextBox Text111 
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
      Left            =   3195
      TabIndex        =   10
      Top             =   5445
      Width           =   1185
   End
   Begin VB.TextBox Text110 
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
      Left            =   3195
      TabIndex        =   9
      Top             =   5040
      Width           =   2355
   End
   Begin VB.TextBox Text109 
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
      Left            =   3195
      TabIndex        =   8
      Top             =   4635
      Width           =   2355
   End
   Begin VB.TextBox Text108 
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
      Left            =   3195
      TabIndex        =   6
      Top             =   3825
      Width           =   2355
   End
   Begin VB.TextBox Text107 
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
      Left            =   3195
      TabIndex        =   5
      Top             =   3420
      Width           =   2355
   End
   Begin VB.TextBox Text106 
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
      Left            =   3195
      TabIndex        =   4
      Top             =   3015
      Width           =   2355
   End
   Begin VB.TextBox Text105 
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
      Left            =   3195
      TabIndex        =   3
      Top             =   2610
      Width           =   4335
   End
   Begin VB.TextBox Text104 
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
      Left            =   3195
      TabIndex        =   2
      Top             =   2205
      Width           =   4335
   End
   Begin VB.TextBox Text103 
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
      Left            =   3195
      TabIndex        =   1
      Top             =   1800
      Width           =   2355
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
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
      Height          =   330
      Left            =   3195
      TabIndex        =   0
      Top             =   990
      Width           =   1410
   End
   Begin VB.Image Image102 
      Height          =   300
      Left            =   3195
      Picture         =   "FrmOrder205.frx":A2DC
      Top             =   540
      Width           =   570
   End
   Begin VB.Image Image101 
      Height          =   300
      Left            =   3195
      Picture         =   "FrmOrder205.frx":A818
      Top             =   540
      Width           =   570
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเลขรถ :"
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
      Left            =   1440
      TabIndex        =   31
      Top             =   1395
      Width           =   1635
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ :"
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
      Height          =   375
      Left            =   1350
      TabIndex        =   30
      Top             =   5850
      Width           =   1770
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "กว้าง x ยาว x สูง :"
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
      Left            =   1575
      TabIndex        =   29
      Top             =   5445
      Width           =   1545
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "อัตราเฉลี่ย กม./ลิตร :"
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
      Left            =   1260
      TabIndex        =   28
      Top             =   5040
      Width           =   1860
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่กรมธรรม์ :"
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
      Left            =   1305
      TabIndex        =   27
      Top             =   4635
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ซื้อ :"
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
      Height          =   330
      Left            =   1530
      TabIndex        =   26
      Top             =   4230
      Width           =   1590
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ลักษณะมาตรฐาน :"
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
      Left            =   1575
      TabIndex        =   25
      Top             =   3825
      Width           =   1545
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขเครื่อง :"
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
      Left            =   1530
      TabIndex        =   24
      Top             =   3420
      Width           =   1590
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขตัวถังรถ :"
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
      Left            =   1665
      TabIndex        =   23
      Top             =   3015
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อรถ 2 :"
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
      Height          =   330
      Left            =   1350
      TabIndex        =   22
      Top             =   2610
      Width           =   1770
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID :"
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
      Height          =   330
      Left            =   2295
      TabIndex        =   21
      Top             =   990
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ทะเบียนรถ :"
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
      Height          =   330
      Left            =   1665
      TabIndex        =   20
      Top             =   1755
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อรถ 1 :"
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
      Height          =   330
      Left            =   1845
      TabIndex        =   19
      Top             =   2205
      Width           =   1275
   End
End
Attribute VB_Name = "FrmOrder205"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
On Error GoTo ErrDescription

Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
Text104.Text = ""
Text105.Text = ""
Text106.Text = ""
Text107.Text = ""
Text108.Text = ""
Text109.Text = ""
Text110.Text = ""
Text111.Text = ""
Text112.Text = ""
Text113.Text = ""
Text114.Text = ""
DTPicker101 = Now
vCheckVehicalOpen = 0
Text102.SetFocus

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As String
Dim vCarLicence As String
Dim vName1 As String
Dim vName2 As String
Dim vBodyNumber As String
Dim vEngineNumber As String
Dim vStandardType As String
Dim vDateBuy As Date
Dim vInsuranceNumber As String
Dim vDistanceRate As Currency
Dim vWide As Currency
Dim vLong As Currency
Dim vHigh As Currency
Dim vActiveStatus As String
Dim vMydescription As String
Dim vCarNo As String
    
On Error GoTo ErrDescription

If Text102.Text <> "" Then
    vCarNo = Trim(Text102.Text)
    vCarLicence = Trim(Text103.Text)
    vName1 = Trim(Text104.Text)
    vName2 = Trim(Text105.Text)
    vBodyNumber = Trim(Text106.Text)
    vEngineNumber = Trim(Text107.Text)
    vStandardType = Trim(Text108.Text)
    vDateBuy = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
    vInsuranceNumber = Trim(Text109.Text)
    vDistanceRate = Trim(Text110.Text)
    vWide = Trim(Text111.Text)
    vLong = Trim(Text112.Text)
    vHigh = Trim(Text113.Text)
    vActiveStatus = 1
    vMydescription = Trim(Text114.Text)
    
    If vCheckVehicalOpen = 0 Then
        Call CheckIsVehical
        If vCheckIsVehical = 1 Then
            MsgBox "สถานที่ขนส่งนี้มีอยู่แล้วกรุณาตรวจสอบ", vbCritical, "Send Error"
            Exit Sub
        End If
        vID = "Null"
    Else
        vID = Trim(Text101.Text)
    End If
    vQuery = "exec bcnp.dbo.USP_DO_VehicalUpdate " & vID & ",'" & vCarLicence & "','" & vName1 & "'," _
                        & " '" & vName2 & "','" & vBodyNumber & "','" & vEngineNumber & "', '" & vStandardType & "', " _
                        & " '" & vDateBuy & "','" & vInsuranceNumber & "'," & vDistanceRate & "," & vWide & ", " _
                        & " " & vLong & "," & vHigh & ", '" & vActiveStatus & "','" & vMydescription & "','" & vUserID & "','" & vCarNo & "' "
    gConnection.Execute vQuery
    If vCheckVehicalOpen = 0 Then
        MsgBox "บันทึกข้อมูลรถขนส่งเรียบร้อยแล้ว", vbInformation, "Send Message"
    ElseIf vCheckVehicalOpen = 1 Then
        MsgBox "บันทึกการแก้ไขข้อมูลรถขนส่งเรียบร้อยแล้ว", vbInformation, "Send Message"
    End If

        Text101.Text = ""
        Text102.Text = ""
        Text103.Text = ""
        Text104.Text = ""
        Text105.Text = ""
        Text106.Text = ""
        Text107.Text = ""
        Text108.Text = ""
        Text109.Text = ""
        Text110.Text = ""
        Text111.Text = ""
        Text112.Text = ""
        Text113.Text = ""
        Text114.Text = ""
        DTPicker101 = Now
        Image101.Visible = True
        Image102.Visible = False
        vCheckVehicalOpen = 0
        Text102.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
On Error GoTo ErrDescription

MDIFrmProgramPrint.Enabled = False
FrmOrder011.Show
vVehicalModule = 2

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Public Sub CheckIsVehical()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vLicence As String

On Error GoTo ErrDescription

vLicence = Trim(Text102.Text)
vQuery = "select id from npmaster.dbo.TB_DO_Vehical Where carlicence = '" & vLicence & "'  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckIsVehical = 1
Else
    vCheckIsVehical = 0
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD104_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As Integer
Dim vAnswer As Integer

On Error GoTo ErrDescription

If Text101.Text <> "" Then
    If vCheckVehicalOpen = 1 Then
    vAnswer = MsgBox("คุณต้องการยกเลิก เส้นทางขนส่งนี้ใช่หรือไม่", vbYesNo, "Question Respond")
        If vAnswer = 6 Then
            vID = Trim(Text101.Text)
            vQuery = "Update npmaster.dbo.TB_DO_Vehical set activestatus = 0 where id = " & vID & " "
            gConnection.Execute vQuery
            Text101.Text = ""
            Text102.Text = ""
            Text103.Text = ""
            Text104.Text = ""
            Text105.Text = ""
            Text106.Text = ""
            Text107.Text = ""
            Text108.Text = ""
            Text109.Text = ""
            Text110.Text = ""
            Text111.Text = ""
            Text112.Text = ""
            Text113.Text = ""
            Text114.Text = ""
            DTPicker101 = Now
            Image101.Visible = True
            Image102.Visible = False
            vCheckVehicalOpen = 0
            Text102.SetFocus
        Else
            Exit Sub
        End If
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD105_Click()
Unload FrmOrder205
End Sub

Private Sub Form_Load()
    Image101.Visible = True
    Image102.Visible = False
    DTPicker101 = Now
End Sub


