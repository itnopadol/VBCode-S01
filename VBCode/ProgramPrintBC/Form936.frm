VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form936 
   Caption         =   "����§ҹ Run Number �͡��õ�ҧ �"
   ClientHeight    =   8160
   ClientLeft      =   4845
   ClientTop       =   255
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form936.frx":0000
   ScaleHeight     =   8160
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1920
      Top             =   6000
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
   Begin VB.ComboBox CMB102 
      Height          =   315
      Left            =   3225
      TabIndex        =   8
      Top             =   1125
      Width           =   3540
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "�������§ҹ"
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
      Left            =   4050
      TabIndex        =   6
      Top             =   3450
      Width           =   1290
   End
   Begin MSComCtl2.DTPicker DTPicker102 
      Height          =   315
      Left            =   3225
      TabIndex        =   5
      Top             =   2700
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61800449
      CurrentDate     =   38628
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   315
      Left            =   3225
      TabIndex        =   3
      Top             =   2175
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61800449
      CurrentDate     =   38628
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   3225
      TabIndex        =   0
      Top             =   1575
      Width           =   2115
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "���͡��������§ҹ"
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
      Left            =   1500
      TabIndex        =   7
      Top             =   1125
      Width           =   1590
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�֧�ѹ���"
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
      Left            =   2475
      TabIndex        =   4
      Top             =   2700
      Width           =   690
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "� �ѹ���"
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
      Height          =   315
      Left            =   2475
      TabIndex        =   2
      Top             =   2175
      Width           =   690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "���͡�������͡���"
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
      Left            =   1500
      TabIndex        =   1
      Top             =   1575
      Width           =   1665
   End
End
Attribute VB_Name = "Form936"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vDocType As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vDocdate1 As Date
Dim vDocdate2 As Date

On Error GoTo ErrDescription

If CMB101.Text <> "" And CMB102.Text <> "" Then
    
    If CMB102.Text = Trim("��§ҹ ��þ�����͡��õ鹩�Ѻ") Then
        vRepID = 271
    ElseIf CMB102.Text = Trim("��§ҹ ��þ�����͡��÷�᷹") Then
        vRepID = 273
    End If
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = " & vRepID & " and reptype = 'NP' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
    
    Select Case Trim(CMB101.Text)
        Case Trim("Back Order")
            vDocType = 4
        Case Trim("��ʹ��Ҥ�")
            vDocType = 3
        Case Trim("���觢��")
            vDocType = 1
        Case Trim("���觨ͧ")
            vDocType = 2
        Case Trim("���觫���")
            vDocType = 5
        Case Trim("㺵�Ǩ�Ѻ�Թ���")
            vDocType = 6
    End Select
    
    vDocdate1 = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
    vDocdate2 = DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year
    
    
    With Crystal101
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@StartDate;" & vDocdate1 & ";true"
    .ParameterFields(1) = "@EndDate;" & vDocdate1 & ";true"
    .ParameterFields(2) = "@DocType;" & vDocType & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    End With
    
Else
    MsgBox "��س����͡ ��������§ҹ ��л������͡������١��ͧ", vbInformation, "��ͤ�����"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
DTPicker101 = Now
DTPicker102 = Now
CMB101.AddItem Trim("Back Order")
CMB101.AddItem Trim("��ʹ��Ҥ�")
CMB101.AddItem Trim("���觢��")
CMB101.AddItem Trim("���觨ͧ")
CMB101.AddItem Trim("���觫���")
CMB101.AddItem Trim("㺵�Ǩ�Ѻ�Թ���")

CMB102.AddItem Trim("��§ҹ ��þ�����͡��õ鹩�Ѻ")
CMB102.AddItem Trim("��§ҹ ��þ�����͡��÷�᷹")
End Sub
