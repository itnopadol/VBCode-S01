VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form991 
   Caption         =   "¡��ԡ���͹��ѵ��͡��õ�ҧ �"
   ClientHeight    =   11010
   ClientLeft      =   3120
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form991.frx":0000
   ScaleHeight     =   11490
   ScaleMode       =   0  'User
   ScaleWidth      =   15392.4
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICCompany 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   15585
      Left            =   0
      ScaleHeight     =   15555
      ScaleWidth      =   15285
      TabIndex        =   12
      Top             =   0
      Width           =   15315
      Begin VB.CommandButton CMDCompany 
         BackColor       =   &H00808080&
         Caption         =   "���͡"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   8055
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3870
         Width           =   2265
      End
      Begin VB.ComboBox CMBCompany 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Form991.frx":9673
         Left            =   6300
         List            =   "Form991.frx":967D
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2745
         Width           =   4065
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   10995
         Left            =   -45
         ScaleHeight     =   10965
         ScaleWidth      =   3090
         TabIndex        =   13
         Top             =   -45
         Width           =   3120
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "���͡ ����ѷ�ӧҹ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4050
         TabIndex        =   14
         Top             =   2745
         Width           =   2265
      End
   End
   Begin VB.PictureBox PICXPC 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   13110
      Left            =   0
      ScaleHeight     =   13080
      ScaleWidth      =   15240
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   15270
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   45
         ScaleHeight     =   975
         ScaleWidth      =   14970
         TabIndex        =   31
         Top             =   45
         Width           =   15000
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "XPC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   135
            TabIndex        =   32
            Top             =   270
            Width           =   2445
         End
      End
      Begin VB.CommandButton CMDBack2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "���͡����ѷ"
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
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   7875
         Width           =   1680
      End
      Begin VB.CommandButton CMDConfirm_XPC 
         BackColor       =   &H00C0C0C0&
         Caption         =   "͹��ѵ�"
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
         Left            =   1845
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   7875
         Width           =   1680
      End
      Begin VB.CommandButton CMDSave_XPC 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¡��ԡ͹��ѵ�"
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
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   7875
         Width           =   1680
      End
      Begin MSComctlLib.ProgressBar PGBUpdate_XPC 
         Height          =   240
         Left            =   90
         TabIndex        =   27
         Top             =   7605
         Width           =   13200
         _ExtentX        =   23283
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ListView ListViewDocNo_XPC 
         Height          =   5325
         Left            =   45
         TabIndex        =   26
         Top             =   2205
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   9393
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "�ӴѺ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "�Ţ����͡���"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "�ѹ����͡���"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "��ҧ�ԧ1"
            Object.Width           =   11465
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "�����˵�"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "������ҧ�͡���"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.CheckBox CHKAll_XPC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���͡������"
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
         Left            =   45
         TabIndex        =   25
         Top             =   1800
         Width           =   1590
      End
      Begin VB.ComboBox CMBDocType_XPC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10755
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1080
         Width           =   2715
      End
      Begin VB.ComboBox CMBModule_XPC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5625
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1080
         Width           =   3345
      End
      Begin MSComCtl2.DTPicker DTPDocDate_XPC 
         Height          =   465
         Left            =   1845
         TabIndex        =   20
         Top             =   1080
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   820
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   70975489
         CurrentDate     =   42149
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "�������͡��� :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9180
         TabIndex        =   23
         Top             =   1170
         Width           =   2130
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "����������� :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   4275
         TabIndex        =   21
         Top             =   1170
         Width           =   1590
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "�͡��û�Ш��ѹ��� :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   45
         TabIndex        =   19
         Top             =   1170
         Width           =   1950
      End
   End
   Begin VB.CommandButton CMDBack1 
      Caption         =   "���͡����ѷ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3465
      TabIndex        =   18
      Top             =   7920
      Width           =   1320
   End
   Begin VB.CommandButton CMDConfirm 
      Caption         =   "͹��ѵ�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2070
      TabIndex        =   11
      Top             =   7920
      Width           =   1320
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   585
      Top             =   10485
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form991.frx":9690
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form991.frx":BAE2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox CHKAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "���͡������"
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
      Left            =   675
      TabIndex        =   10
      Top             =   1980
      Width           =   13965
   End
   Begin MSComctlLib.ProgressBar PGBUpdate 
      Height          =   240
      Left            =   675
      TabIndex        =   9
      Top             =   7605
      Visible         =   0   'False
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CMDSave 
      Caption         =   "¡��ԡ͹��ѵ�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   675
      TabIndex        =   8
      Top             =   7920
      Width           =   1320
   End
   Begin VB.PictureBox PicPoint 
      Height          =   195
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   270
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin MSComCtl2.DTPicker DTPDocDate 
      Height          =   375
      Left            =   2475
      TabIndex        =   6
      Top             =   1350
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   70975489
      CurrentDate     =   40004
   End
   Begin MSComctlLib.ListView ListViewDocNo 
      Height          =   5190
      Left            =   675
      TabIndex        =   4
      Top             =   2340
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   9155
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�ӴѺ"
         Object.Width           =   2620
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�Ţ����͡���"
         Object.Width           =   4366
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�ѹ����͡���"
         Object.Width           =   2620
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "��ҧ�ԧ1"
         Object.Width           =   11352
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�����˵�"
         Object.Width           =   10479
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "������ҧ�͡���"
         Object.Width           =   3493
      EndProperty
   End
   Begin VB.ComboBox CMBModule 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5850
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1350
      Width           =   2310
   End
   Begin VB.ComboBox CMBDocType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10035
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1350
      Width           =   4560
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�͡��û�Ш��ѹ��� :"
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
      Left            =   675
      TabIndex        =   5
      Top             =   1350
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "����������� :"
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
      Height          =   375
      Left            =   4185
      TabIndex        =   2
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�������͡��� :"
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
      Height          =   330
      Left            =   8415
      TabIndex        =   0
      Top             =   1350
      Width           =   1545
   End
End
Attribute VB_Name = "Form991"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CHKAll_Click()
Dim i As Integer

If Me.CHKAll.Value = 1 Then
For i = 1 To Me.ListViewDocNo.ListItems.Count
Me.ListViewDocNo.ListItems(i).Checked = True
Next i
End If

If Me.CHKAll.Value = 0 Then
For i = 1 To Me.ListViewDocNo.ListItems.Count
Me.ListViewDocNo.ListItems(i).Checked = False
Next i
End If
End Sub

Private Sub CMBDocType_Click()
Call CheckData
End Sub

Private Sub CMBDocType_XPC_Click()
Call CheckData_XPC
End Sub

Private Sub CMBModule_Click()
If Me.CMBModule.ListIndex = 0 Then
Call Buy
ElseIf Me.CMBModule.ListIndex = 1 Then
Call Sale
ElseIf Me.CMBModule.ListIndex = 2 Then
Call Vendor
ElseIf Me.CMBModule.ListIndex = 3 Then
Call Customer
ElseIf Me.CMBModule.ListIndex = 4 Then
Call ItemStock
End If

Call CheckData
End Sub

Private Sub CMBModule_XPC_Click()
If Me.CMBModule_XPC.ListIndex = 0 Then
Call Buy
ElseIf Me.CMBModule_XPC.ListIndex = 1 Then
Call Sale
ElseIf Me.CMBModule_XPC.ListIndex = 2 Then
Call Vendor
ElseIf Me.CMBModule_XPC.ListIndex = 3 Then
Call Customer
ElseIf Me.CMBModule_XPC.ListIndex = 4 Then
Call ItemStock
End If

Call CheckData_XPC
End Sub

Private Sub CMDBack1_Click()
Me.PICCompany.Visible = True
Me.PICXPC.Visible = False
Me.CMBCompany.SetFocus
End Sub


Private Sub CMDBack2_Click()
Me.PICCompany.Visible = True
Me.PICXPC.Visible = False
Me.CMBCompany.SetFocus
End Sub

Private Sub CMDCompany_Click()
Me.PICCompany.Visible = False
If Me.CMBCompany.ListIndex = 0 Then
    Me.PICXPC.Visible = False
ElseIf Me.CMBCompany.ListIndex = 1 Then
    Me.PICXPC.Visible = True
    
    Call InitializeDataBaseXPC
End If

End Sub

Private Sub CMDConfirm_Click()
Dim vAnswer As Integer
Dim vDocNo As String
Dim i As Integer
Dim vQuery As String
Dim vCountSelect As Integer
Dim vType As Integer

On Error GoTo ErrDescription


If Me.ListViewDocNo.ListItems.Count > 0 Then
   vAnswer = MsgBox("�س��ͧ��� ͹��ѵ��͡��÷�����͡������������ ?", vbYesNo, "Send Message Question ?")
   
   If vAnswer = 6 Then
   
   For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
         vCountSelect = vCountSelect + 1
      End If
   Next i
   
   If vCountSelect = 0 Then
      MsgBox "�ѧ��������͡�͡��÷���͹��ѵ� ��سҵ�Ǩ�ͺ", vbCritical, "Send Error Message"
      Me.ListViewDocNo.SetFocus
      Exit Sub
   End If
   
   Me.PGBUpdate.Visible = True
   Me.PGBUpdate.Min = 0
   Me.PGBUpdate.Max = vCountSelect
   vType = Me.CMBDocType.ListIndex
   
   If Me.CMBModule.ListIndex = 0 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleBuy " & vType & ",'" & vDocNo & "',1 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   If Me.CMBModule.ListIndex = 1 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleSale " & vType & ",'" & vDocNo & "',1 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   If Me.CMBModule.ListIndex = 2 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleVendor " & vType & ",'" & vDocNo & "',1 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
    If Me.CMBModule.ListIndex = 3 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleCustomer " & vType & ",'" & vDocNo & "',1 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   If Me.CMBModule.ListIndex = 4 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleStock " & vType & ",'" & vDocNo & "',1 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   
   Me.ListViewDocNo.ListItems.Clear
   Me.PGBUpdate.Value = 0
   Me.PGBUpdate.Visible = False
   MsgBox "͹��ѵ��͡��÷�����͡��� ���º�������� ��سҵ�Ǩ�ͺ", vbInformation, "Send Information Message"
   
   Me.CMBModule.ListIndex = 0
   Me.DTPDocDate.Value = Now
   Me.CHKAll.Value = 0
   Me.CMBDocType.SetFocus
   End If

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDConfirm_XPC_Click()
Dim vAnswer As Integer
Dim vDocNo As String
Dim i As Integer
Dim vQuery As String
Dim vCountSelect As Integer
Dim vType As Integer

On Error GoTo ErrDescription


If Me.ListViewDocNo_XPC.ListItems.Count > 0 Then
   vAnswer = MsgBox("�س��ͧ��� ͹��ѵ��͡��÷�����͡������������ ?", vbYesNo, "Send Message Question ?")
   
   If vAnswer = 6 Then
   
   For i = 1 To Me.ListViewDocNo_XPC.ListItems.Count
      If Me.ListViewDocNo_XPC.ListItems(i).Checked = True Then
         vCountSelect = vCountSelect + 1
      End If
   Next i
   
   If vCountSelect = 0 Then
      MsgBox "�ѧ��������͡�͡��÷���͹��ѵ� ��سҵ�Ǩ�ͺ", vbCritical, "Send Error Message"
      Me.ListViewDocNo_XPC.SetFocus
      Exit Sub
   End If
   
   Me.PGBUpdate_XPC.Visible = True
   Me.PGBUpdate_XPC.Min = 0
   Me.PGBUpdate_XPC.Max = vCountSelect
   vType = Me.CMBDocType_XPC.ListIndex
   
   If Me.CMBModule_XPC.ListIndex = 0 Then
      For i = 1 To Me.ListViewDocNo_XPC.ListItems.Count
      If Me.ListViewDocNo_XPC.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo_XPC.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleBuy " & vType & ",'" & vDocNo & "',1 "
      vXPCConnection.Execute vQuery
      Me.PGBUpdate_XPC.Value = Me.PGBUpdate_XPC + 1
      End If
      Next i
   End If
   
   If Me.CMBModule_XPC.ListIndex = 1 Then
      For i = 1 To Me.ListViewDocNo_XPC.ListItems.Count
      If Me.ListViewDocNo_XPC.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo_XPC.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleSale " & vType & ",'" & vDocNo & "',1 "
      vXPCConnection.Execute vQuery
      Me.PGBUpdate_XPC.Value = Me.PGBUpdate_XPC + 1
      End If
      Next i
   End If
   
   If Me.CMBModule_XPC.ListIndex = 4 Then
      For i = 1 To Me.ListViewDocNo_XPC.ListItems.Count
      If Me.ListViewDocNo_XPC.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo_XPC.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleStock " & vType & ",'" & vDocNo & "',1 "
      vXPCConnection.Execute vQuery
      Me.PGBUpdate_XPC.Value = Me.PGBUpdate_XPC + 1
      End If
      Next i
   End If
   
   
   Me.ListViewDocNo_XPC.ListItems.Clear
   Me.PGBUpdate_XPC.Value = 0
   Me.PGBUpdate_XPC.Visible = False
   MsgBox "͹��ѵ��͡��÷�����͡��� ���º�������� ��سҵ�Ǩ�ͺ", vbInformation, "Send Information Message"
   
   Me.CMBModule_XPC.ListIndex = 0
   Me.DTPDocDate_XPC.Value = Now
   Me.CHKAll_XPC.Value = 0
   Me.CMBDocType_XPC.SetFocus
   End If

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDSave_Click()
Dim vAnswer As Integer
Dim vDocNo As String
Dim i As Integer
Dim vQuery As String
Dim vCountSelect As Integer
Dim vType As Integer

On Error GoTo ErrDescription


If Me.ListViewDocNo.ListItems.Count > 0 Then
   vAnswer = MsgBox("�س��ͧ��� ¡��ԡ���͹��ѵԢͧ�͡��÷�����͡������������ ?", vbYesNo, "Send Message Question ?")
   
   If vAnswer = 6 Then
   
   For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
         vCountSelect = vCountSelect + 1
      End If
   Next i
   
   If vCountSelect = 0 Then
      MsgBox "�ѧ��������͡�͡��÷��� ¡��ԡ���͹��ѵ� ��سҵ�Ǩ�ͺ", vbCritical, "Send Error Message"
      Me.ListViewDocNo.SetFocus
      Exit Sub
   End If
   
   Me.PGBUpdate.Visible = True
   Me.PGBUpdate.Min = 0
   Me.PGBUpdate.Max = vCountSelect
   
   vType = Me.CMBDocType.ListIndex
   
   If Me.CMBModule.ListIndex = 0 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleBuy " & vType & ",'" & vDocNo & "',0 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   If Me.CMBModule.ListIndex = 1 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleSale " & vType & ",'" & vDocNo & "',0 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
    If Me.CMBModule.ListIndex = 2 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleVendor " & vType & ",'" & vDocNo & "',0 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   
    If Me.CMBModule.ListIndex = 3 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleCustomer " & vType & ",'" & vDocNo & "',0 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   If Me.CMBModule.ListIndex = 4 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleStock " & vType & ",'" & vDocNo & "',0 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   Me.ListViewDocNo.ListItems.Clear
   Me.PGBUpdate.Value = 0
   Me.PGBUpdate.Visible = False
   MsgBox "¡��ԡ ���͹��ѵ��͡��÷�����͡��� ���º�������� ��سҵ�Ǩ�ͺ", vbInformation, "Send Information Message"
   
   Me.CMBModule.ListIndex = 0
   Me.DTPDocDate.Value = Now
   Me.CHKAll.Value = 0
   Me.CMBDocType.SetFocus
   End If

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub CMDSave_XPC_Click()
Dim vAnswer As Integer
Dim vDocNo As String
Dim i As Integer
Dim vQuery As String
Dim vCountSelect As Integer
Dim vType As Integer

On Error GoTo ErrDescription


If Me.ListViewDocNo_XPC.ListItems.Count > 0 Then
   vAnswer = MsgBox("�س��ͧ��� ¡��ԡ���͹��ѵԢͧ�͡��÷�����͡������������ ?", vbYesNo, "Send Message Question ?")
   
   If vAnswer = 6 Then
   
   For i = 1 To Me.ListViewDocNo_XPC.ListItems.Count
      If Me.ListViewDocNo_XPC.ListItems(i).Checked = True Then
         vCountSelect = vCountSelect + 1
      End If
   Next i
   
   If vCountSelect = 0 Then
      MsgBox "�ѧ��������͡�͡��÷��� ¡��ԡ���͹��ѵ� ��سҵ�Ǩ�ͺ", vbCritical, "Send Error Message"
      Me.ListViewDocNo_XPC.SetFocus
      Exit Sub
   End If
   
   Me.PGBUpdate_XPC.Visible = True
   Me.PGBUpdate_XPC.Min = 0
   Me.PGBUpdate_XPC.Max = vCountSelect
   
   vType = Me.CMBDocType_XPC.ListIndex
   
   If Me.CMBModule_XPC.ListIndex = 0 Then
      For i = 1 To Me.ListViewDocNo_XPC.ListItems.Count
      If Me.ListViewDocNo_XPC.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo_XPC.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleBuy " & vType & ",'" & vDocNo & "',0 "
      vXPCConnection.Execute vQuery
      Me.PGBUpdate_XPC.Value = Me.PGBUpdate_XPC + 1
      End If
      Next i
   End If
   
   If Me.CMBModule_XPC.ListIndex = 1 Then
      For i = 1 To Me.ListViewDocNo_XPC.ListItems.Count
      If Me.ListViewDocNo_XPC.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo_XPC.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleSale " & vType & ",'" & vDocNo & "',0 "
      vXPCConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   If Me.CMBModule_XPC.ListIndex = 4 Then
      For i = 1 To Me.ListViewDocNo_XPC.ListItems.Count
      If Me.ListViewDocNo_XPC.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo_XPC.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleStock " & vType & ",'" & vDocNo & "',0 "
      vXPCConnection.Execute vQuery
      Me.PGBUpdate_XPC.Value = Me.PGBUpdate_XPC + 1
      End If
      Next i
   End If
   
   Me.ListViewDocNo_XPC.ListItems.Clear
   Me.PGBUpdate_XPC.Value = 0
   Me.PGBUpdate_XPC.Visible = False
   MsgBox "¡��ԡ ���͹��ѵ��͡��÷�����͡��� ���º�������� ��سҵ�Ǩ�ͺ", vbInformation, "Send Information Message"
   
   Me.CMBModule_XPC.ListIndex = 0
   Me.DTPDocDate_XPC.Value = Now
   Me.CHKAll_XPC.Value = 0
   Me.CMBDocType_XPC.SetFocus
   End If

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub DTPDocDate_Change()
Call CheckData
End Sub

Private Sub DTPDocDate_XPC_Change()
Call CheckData_XPC
End Sub

Private Sub Form_Load()
Me.DTPDocDate.Value = Now
Me.DTPDocDate_XPC.Value = Now
Call SetListViewColor(ListViewDocNo, PicPoint, vbWhite, vbLightGreen)
Call CreateModule
End Sub

Public Sub CreateModule()
Me.CMBModule.AddItem ("1.�Ѵ����")
Me.CMBModule.AddItem ("2.�Ѵ���")
Me.CMBModule.AddItem ("3.���˹��")
Me.CMBModule.AddItem ("4.�١˹��")
Me.CMBModule.AddItem ("5.�Թ��Ҥ���ѧ")
Me.CMBModule.AddItem ("6.����и�Ҥ��")

Me.CMBModule_XPC.AddItem ("1.�Ѵ����")
Me.CMBModule_XPC.AddItem ("2.�Ѵ���")
Me.CMBModule_XPC.AddItem ("3.���˹��")
Me.CMBModule_XPC.AddItem ("4.�١˹��")
Me.CMBModule_XPC.AddItem ("5.�Թ��Ҥ���ѧ")
Me.CMBModule_XPC.AddItem ("6.����и�Ҥ��")

Me.CMBModule_XPC.ListIndex = 0
End Sub


Public Sub Buy()
Me.CMBDocType.Clear
Me.CMBDocType.AddItem ("�͡��� 1.��ʹͫ����Թ���")
Me.CMBDocType.AddItem ("�͡��� 2.���觫����Թ���")
Me.CMBDocType.AddItem ("�͡��� 3.㺨����Թ�Ѵ��")
Me.CMBDocType.AddItem ("�͡��� 4.㺨����Թ��ǧ˹��")
Me.CMBDocType.AddItem ("�͡��� 5.��Ѻ�Թ��Ҩҡ��ë���")
Me.CMBDocType.AddItem ("�͡��� 6.㺵��˹��ҡ��ë���")
Me.CMBDocType.AddItem ("�͡��� 7.��觤׹�Թ���")
Me.CMBDocType.AddItem ("�͡��� 8.�Ŵ˹��")
Me.CMBDocType.AddItem ("�͡��� 9.㺫����Թ�����к�ԡ��")
Me.CMBDocType.AddItem ("�͡��� 10.��觤׹�Թ���/Ŵ˹��")
Me.CMBDocType.AddItem ("�͡��� 11.�����˹��/�����Թ������˹��")

Me.CMBDocType_XPC.Clear
Me.CMBDocType_XPC.AddItem ("�͡��� 1.��ʹͫ����Թ���")
Me.CMBDocType_XPC.AddItem ("�͡��� 2.���觫����Թ���")
Me.CMBDocType_XPC.AddItem ("�͡��� 3.㺨����Թ�Ѵ��")
Me.CMBDocType_XPC.AddItem ("�͡��� 4.㺨����Թ��ǧ˹��")
Me.CMBDocType_XPC.AddItem ("�͡��� 5.��Ѻ�Թ��Ҩҡ��ë���")
Me.CMBDocType_XPC.AddItem ("�͡��� 6.㺵��˹��ҡ��ë���")
Me.CMBDocType_XPC.AddItem ("�͡��� 7.��觤׹�Թ���")
Me.CMBDocType_XPC.AddItem ("�͡��� 8.�Ŵ˹��")
Me.CMBDocType_XPC.AddItem ("�͡��� 9.㺫����Թ�����к�ԡ��")
Me.CMBDocType_XPC.AddItem ("�͡��� 10.��觤׹�Թ���/Ŵ˹��")
Me.CMBDocType_XPC.AddItem ("�͡��� 11.�����˹��/�����Թ������˹��")
End Sub

Public Sub Sale()
Me.CMBDocType.Clear
Me.CMBDocType.AddItem ("�͡��� 1.��ʹ��Ҥ�")
Me.CMBDocType.AddItem ("�͡��� 2.���觢�¤�ҧ��(BackOrder)")
Me.CMBDocType.AddItem ("�͡��� 3.���觨ͧ")
Me.CMBDocType.AddItem ("�͡��� 4.���觢��")
Me.CMBDocType.AddItem ("�͡��� 5.��Ѻ�Թ�Ѵ��")
Me.CMBDocType.AddItem ("�͡��� 6.��Ѻ�Թ��ǧ˹��")
Me.CMBDocType.AddItem ("�͡��� 7.㺤׹�Թ�Ѻ��ǧ˹��")
Me.CMBDocType.AddItem ("�͡��� 8.����Թ���,��ԡ��")
Me.CMBDocType.AddItem ("�͡��� 9.����Թ��� POS")
Me.CMBDocType.AddItem ("�͡��� 10.��Ѻ�׹�Թ���/Ŵ˹��")
Me.CMBDocType.AddItem ("�͡��� 11.�����˹��/�����Թ���(�١���)")

Me.CMBDocType_XPC.Clear
Me.CMBDocType_XPC.AddItem ("�͡��� 1.��ʹ��Ҥ�")
Me.CMBDocType_XPC.AddItem ("�͡��� 2.���觢�¤�ҧ��(BackOrder)")
Me.CMBDocType_XPC.AddItem ("�͡��� 3.���觨ͧ")
Me.CMBDocType_XPC.AddItem ("�͡��� 4.���觢��")
Me.CMBDocType_XPC.AddItem ("�͡��� 5.��Ѻ�Թ�Ѵ��")
Me.CMBDocType_XPC.AddItem ("�͡��� 6.��Ѻ�Թ��ǧ˹��")
Me.CMBDocType_XPC.AddItem ("�͡��� 7.㺤׹�Թ�Ѻ��ǧ˹��")
Me.CMBDocType_XPC.AddItem ("�͡��� 8.����Թ���,��ԡ��")
Me.CMBDocType_XPC.AddItem ("�͡��� 9.����Թ��� POS")
Me.CMBDocType_XPC.AddItem ("�͡��� 10.��Ѻ�׹�Թ���/Ŵ˹��")
Me.CMBDocType_XPC.AddItem ("�͡��� 11.�����˹��/�����Թ���(�١���)")
End Sub

Public Sub Vendor()
Me.CMBDocType.Clear
Me.CMBDocType.AddItem ("�͡��� 1.���˹��¡��")
Me.CMBDocType.AddItem ("�͡��� 2.������˹����� �")
Me.CMBDocType.AddItem ("�͡��� 3.��Ѻ�ҧ���")
Me.CMBDocType.AddItem ("�͡��� 4.㺨��ª���˹��")
Me.CMBDocType.AddItem ("�͡��� 5.�Ѵ˹���٭(���˹��)")

Me.CMBDocType_XPC.Clear
Me.CMBDocType_XPC.AddItem ("�͡��� 1.���˹��¡��")
Me.CMBDocType_XPC.AddItem ("�͡��� 2.������˹����� �")
Me.CMBDocType_XPC.AddItem ("�͡��� 3.��Ѻ�ҧ���")
Me.CMBDocType_XPC.AddItem ("�͡��� 4.㺨��ª���˹��")
Me.CMBDocType_XPC.AddItem ("�͡��� 5.�Ѵ˹���٭(���˹��)")
End Sub

Public Sub Customer()
Me.CMBDocType.Clear
Me.CMBDocType.AddItem ("�͡��� 1.�١˹��¡�ҵ鹻�")
Me.CMBDocType.AddItem ("�͡��� 2.����١˹����� �")
Me.CMBDocType.AddItem ("�͡��� 3.��ҧ���")
Me.CMBDocType.AddItem ("�͡��� 4.��ҧ����ѵ��ѵ�")
Me.CMBDocType.AddItem ("�͡��� 5.����稪��Ǥ���")
Me.CMBDocType.AddItem ("�͡��� 6.������Ѻ�Թ/�Ѻ����˹��")
Me.CMBDocType.AddItem ("�͡��� 7.�Ѵ˹���٭(�١˹��)")

Me.CMBDocType_XPC.Clear
Me.CMBDocType_XPC.AddItem ("�͡��� 1.�١˹��¡�ҵ鹻�")
Me.CMBDocType_XPC.AddItem ("�͡��� 2.����١˹����� �")
Me.CMBDocType_XPC.AddItem ("�͡��� 3.��ҧ���")
Me.CMBDocType_XPC.AddItem ("�͡��� 4.��ҧ����ѵ��ѵ�")
Me.CMBDocType_XPC.AddItem ("�͡��� 5.����稪��Ǥ���")
Me.CMBDocType_XPC.AddItem ("�͡��� 6.������Ѻ�Թ/�Ѻ����˹��")
Me.CMBDocType_XPC.AddItem ("�͡��� 7.�Ѵ˹���٭(�١˹��)")
End Sub

Public Sub ItemStock()
Me.CMBDocType.Clear
Me.CMBDocType.AddItem ("�͡��� 1.�Թ���¡��")
Me.CMBDocType.AddItem ("�͡��� 2.㺢��ԡ���Թ���,�ѵ�شԺ")
Me.CMBDocType.AddItem ("�͡��� 3.��ԡ���Թ���,�ѵ�شԺ")
Me.CMBDocType.AddItem ("�͡��� 4.��Ѻ�׹�Թ���,�ѵ�شԺ")
Me.CMBDocType.AddItem ("�͡��� 5.��Ѻ�Թ���������ٻ")
Me.CMBDocType.AddItem ("�͡��� 6.㺢��͹�Թ���")
Me.CMBDocType.AddItem ("�͡��� 7.��͹�Թ��������ҧ��ѧ")
Me.CMBDocType.AddItem ("�͡��� 8.㺻�Ѻ��ا�Թ�����ѧ��Ǩ�Ѻ")

Me.CMBDocType_XPC.Clear
Me.CMBDocType_XPC.AddItem ("�͡��� 1.�Թ���¡��")
Me.CMBDocType_XPC.AddItem ("�͡��� 2.㺢��ԡ���Թ���,�ѵ�شԺ")
Me.CMBDocType_XPC.AddItem ("�͡��� 3.��ԡ���Թ���,�ѵ�شԺ")
Me.CMBDocType_XPC.AddItem ("�͡��� 4.��Ѻ�׹�Թ���,�ѵ�شԺ")
Me.CMBDocType_XPC.AddItem ("�͡��� 5.��Ѻ�Թ���������ٻ")
Me.CMBDocType_XPC.AddItem ("�͡��� 6.㺢��͹�Թ���")
Me.CMBDocType_XPC.AddItem ("�͡��� 7.��͹�Թ��������ҧ��ѧ")
Me.CMBDocType_XPC.AddItem ("�͡��� 8.㺻�Ѻ��ا�Թ�����ѧ��Ǩ�Ѻ")
End Sub

Public Sub CheckData()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Dim vListDoc As ListItem
Dim i As Integer
Dim vDocdate As String
Dim vType As Integer
Dim vMemIsConfirm As Integer

On Error GoTo ErrDescription

If Me.CMBModule.Text <> "" And Me.CMBDocType.Text <> "" Then

 If Me.CMBModule.ListIndex = 0 Then
   Me.ListViewDocNo.ListItems.Clear
   vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year
   vType = Me.CMBDocType.ListIndex
   
   vQuery = "exec dbo.USP_NP_CancelConfirmModuleBuy " & vType & ",'" & vDocdate & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
               vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
               If vMemIsConfirm = 0 Then
               Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 1)
               Else
               Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 2)
               End If
               vListDoc.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
               vListDoc.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
               vListDoc.SubItems(3) = Trim(vRecordset.Fields("code").Value) & "/" & Trim(vRecordset.Fields("name1").Value)
               vListDoc.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
               vListDoc.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
       vRecordset.MoveNext
    Next i
    End If
    vRecordset.Close
   End If
   
   If Me.CMBModule.ListIndex = 1 Then
    Me.ListViewDocNo.ListItems.Clear
    vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year
    vType = Me.CMBDocType.ListIndex
    
    vQuery = "exec dbo.USP_NP_CancelConfirmModuleSale " & vType & ",'" & vDocdate & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        For i = 1 To vRecordset.RecordCount
                vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
                If vMemIsConfirm = 0 Then
                Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 1)
                Else
                Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 2)
                End If
                vListDoc.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                vListDoc.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
                vListDoc.SubItems(3) = Trim(vRecordset.Fields("code").Value) & "/" & Trim(vRecordset.Fields("name1").Value)
                vListDoc.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
                vListDoc.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
        vRecordset.MoveNext
    Next i
    End If
    vRecordset.Close
   End If
   
   If Me.CMBModule.ListIndex = 2 Then
        Me.ListViewDocNo.ListItems.Clear
        vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year
        vType = Me.CMBDocType.ListIndex
        
        vQuery = "exec dbo.USP_NP_CancelConfirmModuleVendor " & vType & ",'" & vDocdate & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            For i = 1 To vRecordset.RecordCount
                    vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
                    If vMemIsConfirm = 0 Then
                    Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 1)
                    Else
                    Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 2)
                    End If
                    vListDoc.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDoc.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
                    vListDoc.SubItems(3) = Trim(vRecordset.Fields("code").Value) & "/" & Trim(vRecordset.Fields("name1").Value)
                    vListDoc.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDoc.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
            vRecordset.MoveNext
        Next i
        End If
        vRecordset.Close
   End If
   
   If Me.CMBModule.ListIndex = 3 Then
    Me.ListViewDocNo.ListItems.Clear
        vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year
        vType = Me.CMBDocType.ListIndex
        
        vQuery = "exec dbo.USP_NP_CancelConfirmModuleCustomer " & vType & ",'" & vDocdate & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            For i = 1 To vRecordset.RecordCount
                    vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
                    If vMemIsConfirm = 0 Then
                    Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 1)
                    Else
                    Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 2)
                    End If
                    vListDoc.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                    vListDoc.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
                    vListDoc.SubItems(3) = Trim(vRecordset.Fields("code").Value) & "/" & Trim(vRecordset.Fields("name1").Value)
                    vListDoc.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
                    vListDoc.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
            vRecordset.MoveNext
        Next i
        End If
        vRecordset.Close
   End If
   
   If Me.CMBModule.ListIndex = 4 Then
   Me.ListViewDocNo.ListItems.Clear
   vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year
   vType = Me.CMBDocType.ListIndex
   
   vQuery = "exec dbo.USP_NP_CancelConfirmModuleStock " & vType & ",'" & vDocdate & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
               vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
               If vMemIsConfirm = 0 Then
               Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 1)
               Else
               Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 2)
               End If
               vListDoc.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
               vListDoc.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
               vListDoc.SubItems(3) = Trim(vRecordset.Fields("code").Value) & "/" & Trim(vRecordset.Fields("name1").Value)
               vListDoc.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
               vListDoc.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
       vRecordset.MoveNext
   Next i
   End If
   vRecordset.Close
   End If
   
Else
   Me.ListViewDocNo.ListItems.Clear
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub CheckData_XPC()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Dim vListDoc As ListItem
Dim i As Integer
Dim vDocdate As String
Dim vType As Integer
Dim vMemIsConfirm As Integer

On Error GoTo ErrDescription

If Me.CMBModule_XPC.Text <> "" And Me.CMBDocType_XPC.Text <> "" Then

 If Me.CMBModule_XPC.ListIndex = 0 Then
   Me.ListViewDocNo_XPC.ListItems.Clear
   vDocdate = Me.DTPDocDate_XPC.Day & "/" & Me.DTPDocDate_XPC.Month & "/" & Me.DTPDocDate_XPC.Year
   vType = Me.CMBDocType_XPC.ListIndex
   
   vQuery = "exec dbo.USP_NP_CancelConfirmModuleBuy " & vType & ",'" & vDocdate & "' "
   If OpenDataBaseXPC(vXPCConnection, vRecordset, vQuery) <> 0 Then
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
               vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
               If vMemIsConfirm = 0 Then
               Set vListDoc = Me.ListViewDocNo_XPC.ListItems.Add(, , i, , 1)
               Else
               Set vListDoc = Me.ListViewDocNo_XPC.ListItems.Add(, , i, , 2)
               End If
               vListDoc.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
               vListDoc.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
               vListDoc.SubItems(3) = Trim(vRecordset.Fields("code").Value) & "/" & Trim(vRecordset.Fields("name1").Value)
               vListDoc.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
               vListDoc.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
       vRecordset.MoveNext
   Next i
   End If
   vRecordset.Close
   End If
   
   If Me.CMBModule_XPC.ListIndex = 1 Then
   Me.ListViewDocNo_XPC.ListItems.Clear
   vDocdate = Me.DTPDocDate_XPC.Day & "/" & Me.DTPDocDate_XPC.Month & "/" & Me.DTPDocDate_XPC.Year
   vType = Me.CMBDocType_XPC.ListIndex
   
   vQuery = "exec dbo.USP_NP_CancelConfirmModuleSale " & vType & ",'" & vDocdate & "' "
   If OpenDataBaseXPC(vXPCConnection, vRecordset, vQuery) <> 0 Then
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
               vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
               If vMemIsConfirm = 0 Then
               Set vListDoc = Me.ListViewDocNo_XPC.ListItems.Add(, , i, , 1)
               Else
               Set vListDoc = Me.ListViewDocNo_XPC.ListItems.Add(, , i, , 2)
               End If
               vListDoc.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
               vListDoc.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
               vListDoc.SubItems(3) = Trim(vRecordset.Fields("code").Value) & "/" & Trim(vRecordset.Fields("name1").Value)
               vListDoc.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
               vListDoc.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
       vRecordset.MoveNext
   Next i
   End If
   vRecordset.Close
   End If
   
   If Me.CMBModule_XPC.ListIndex = 2 Then
   
   End If
   
   If Me.CMBModule_XPC.ListIndex = 3 Then
   
   End If
   
   If Me.CMBModule_XPC.ListIndex = 4 Then
   Me.ListViewDocNo_XPC.ListItems.Clear
   vDocdate = Me.DTPDocDate_XPC.Day & "/" & Me.DTPDocDate_XPC.Month & "/" & Me.DTPDocDate_XPC.Year
   vType = Me.CMBDocType_XPC.ListIndex
   
   vQuery = "exec dbo.USP_NP_CancelConfirmModuleStock " & vType & ",'" & vDocdate & "' "
   If OpenDataBaseXPC(vXPCConnection, vRecordset, vQuery) <> 0 Then
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
               vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
               If vMemIsConfirm = 0 Then
               Set vListDoc = Me.ListViewDocNo_XPC.ListItems.Add(, , i, , 1)
               Else
               Set vListDoc = Me.ListViewDocNo_XPC.ListItems.Add(, , i, , 2)
               End If
               vListDoc.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
               vListDoc.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
               vListDoc.SubItems(3) = Trim(vRecordset.Fields("code").Value) & "/" & Trim(vRecordset.Fields("name1").Value)
               vListDoc.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
               vListDoc.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
       vRecordset.MoveNext
   Next i
   End If
   vRecordset.Close
   End If
   
Else
   Me.ListViewDocNo_XPC.ListItems.Clear
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

