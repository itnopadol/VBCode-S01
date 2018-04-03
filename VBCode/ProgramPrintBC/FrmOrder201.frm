VERSION 5.00
Begin VB.Form FrmOrder201 
   Caption         =   "Form201 กำหนด ระดับความสำคัญ"
   ClientHeight    =   8055
   ClientLeft      =   2145
   ClientTop       =   1035
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmOrder201.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD105 
      Height          =   330
      Left            =   4320
      Picture         =   "FrmOrder201.frx":9673
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "ออก หน้าทะเบียนระดับความสำคัญ"
      Top             =   4050
      Width           =   330
   End
   Begin VB.CommandButton CMD104 
      Height          =   330
      Left            =   3870
      Picture         =   "FrmOrder201.frx":99DD
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "ยกเลิกข้อมูล"
      Top             =   4050
      Width           =   330
   End
   Begin VB.CommandButton CMD103 
      Height          =   330
      Left            =   3420
      Picture         =   "FrmOrder201.frx":BB7B
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "ค้นหาข้อมูล"
      Top             =   4050
      Width           =   330
   End
   Begin VB.CommandButton CMD102 
      Height          =   330
      Left            =   2970
      Picture         =   "FrmOrder201.frx":BF48
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "บันทึกและปรับปรุงข้อมูล"
      Top             =   4050
      Width           =   330
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   2520
      Picture         =   "FrmOrder201.frx":C26F
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "เคลียร์หน้าจอ"
      Top             =   4050
      Width           =   330
   End
   Begin VB.TextBox Text103 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2295
      Width           =   5730
   End
   Begin VB.TextBox Text102 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1755
      Width           =   2445
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1215
      Width           =   1185
   End
   Begin VB.Image Image102 
      Height          =   300
      Left            =   2520
      Picture         =   "FrmOrder201.frx":C654
      Top             =   720
      Width           =   570
   End
   Begin VB.Image Image101 
      Height          =   300
      Left            =   2520
      Picture         =   "FrmOrder201.frx":CB90
      Top             =   720
      Width           =   570
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รายละเอียด :"
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
      Left            =   900
      TabIndex        =   10
      Top             =   2295
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ระดับความสำคัญ :"
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
      Left            =   900
      TabIndex        =   9
      Top             =   1755
      Width           =   1545
   End
   Begin VB.Label Label1 
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
      Left            =   945
      TabIndex        =   8
      Top             =   1215
      Width           =   1500
   End
End
Attribute VB_Name = "FrmOrder201"
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
Image101.Visible = True
Image102.Visible = False
vCheckPriorityOpen = 0
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
Dim vPriority As String
Dim vMydescription As String
Dim vIsCancel As String

On Error GoTo ErrDescription

If Text102.Text <> "" Then
    vPriority = Trim(Text102.Text)
    vMydescription = Trim(Text103.Text)
    Call CheckIsCancel
    vIsCancel = vCheckIsCancel1
    vQuery = "exec bcnp.dbo.USP_DO_PriorityUpdate '" & vPriority & "','" & vMydescription & "','" & vIsCancel & "' "
    gConnection.Execute vQuery
    
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Image101.Visible = True
    Image102.Visible = False
    vCheckPriorityOpen = 0
    
    MsgBox "บันทึกข้อมูลเรียบร้อยแล้วครับ", vbInformation, "Send Information"
        
Else
    MsgBox "ต้องใส่ระดับความสำคัญถึงจะบันทึกข้อมูลได้", vbInformation, "Send Information"
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
FrmOrder005.Show
vPriorityModule = 2

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
    If vCheckPriorityOpen = 1 Then
        vAnswer = MsgBox("คุณต้องการยกเลิก ระดับความสำคัญนี้ใช่หรือไม่", vbYesNo, "Question Respond")
        If vAnswer = 6 Then
            vID = Trim(Text101.Text)
            vQuery = "Update npmaster.dbo.TB_DO_Priority set iscancel = 1 where id = " & vID & " "
            gConnection.Execute vQuery
            Text101.Text = ""
            Text102.Text = ""
            Text103.Text = ""
            Image101.Visible = True
            Image102.Visible = False
            vCheckPriorityOpen = 0
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
Unload FrmOrder201
End Sub

Private Sub Form_Load()
Image101.Visible = True
Image102.Visible = False
End Sub

Public Sub CheckIsCancel()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vPriority As String

On Error GoTo ErrDescription

vPriority = Trim(Text102.Text)
vQuery = "select id,iscancel from npmaster.dbo.TB_DO_Priority Where Priority = '" & vPriority & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckIsCancel1 = Trim(vRecordset.Fields("iscancel").Value)
Else
    vCheckIsCancel1 = 0
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

