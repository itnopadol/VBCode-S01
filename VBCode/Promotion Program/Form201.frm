VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form201 
   Caption         =   "���ҧ��ʹ��Թ���"
   ClientHeight    =   9000
   ClientLeft      =   2745
   ClientTop       =   450
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   959
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   2925
      Top             =   8370
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox Crystal1011 
      Height          =   480
      Left            =   90
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   55
      Top             =   8055
      Width           =   1200
   End
   Begin VB.CommandButton CMD110 
      Caption         =   "������͡���"
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
      Left            =   10965
      TabIndex        =   36
      Top             =   8100
      Width           =   1440
   End
   Begin VB.CommandButton CMD109 
      Caption         =   "������˹�Ҩ�"
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
      Left            =   12840
      TabIndex        =   35
      Top             =   8100
      Width           =   1440
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6465
      Left            =   90
      TabIndex        =   23
      Top             =   1440
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   11404
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   529
      TabCaption(0)   =   "��������´�Թ���"
      TabPicture(0)   =   "Form201.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LBLHotPrice"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label9"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label10"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LBLRetailCom"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "LBLCampaignCom"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "LBLFormComDate"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "LBLToComDate"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label12"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "LBLWholeSaleCom"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "ItemDetail101"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "ItemDetail105"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "ItemDetail102"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ItemDetail103"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "ItemDetail104"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "ItemDetail106"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "ListView101"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "CHK102"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "CMD104"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "CMD105"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "CHK103"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Check101"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "ItemDetail107"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "CMD111"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "CMD108"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "CMBHotPrice"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      Begin VB.ComboBox CMBHotPrice 
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
         ItemData        =   "Form201.frx":001C
         Left            =   11520
         List            =   "Form201.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1485
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.CommandButton CMD108 
         Caption         =   "ź��¡��"
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
         Height          =   465
         Left            =   12510
         TabIndex        =   39
         ToolTipText     =   "������� ������Ѻź��¡���Թ���㹵��ҧ��ҧ��"
         Top             =   5850
         Width           =   1440
      End
      Begin VB.CommandButton CMD111 
         Height          =   315
         Left            =   6705
         Picture         =   "Form201.frx":0034
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "������� ������Ѻ������������蹢ͧ�Թ���"
         Top             =   1470
         Width           =   315
      End
      Begin VB.TextBox ItemDetail107 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4230
         TabIndex        =   12
         Top             =   1470
         Width           =   2430
      End
      Begin VB.CheckBox Check101 
         Caption         =   "IsBrochure"
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
         Left            =   1080
         TabIndex        =   11
         Top             =   1485
         Width           =   1215
      End
      Begin VB.CheckBox CHK103 
         Caption         =   "Ŵ�� %"
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
         Left            =   7005
         TabIndex        =   9
         Top             =   1065
         Width           =   1065
      End
      Begin VB.CommandButton CMD105 
         Caption         =   "�����Թ���"
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
         Left            =   225
         Picture         =   "Form201.frx":0401
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "������� ������Ѻ�����Թ��ҷ���Թ��ҵ�����������Ѿഷ"
         Top             =   1890
         Width           =   1140
      End
      Begin VB.CommandButton CMD104 
         Height          =   315
         Left            =   3030
         Picture         =   "Form201.frx":071B
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "������� ������Ѻ���������Թ������ͪ����Թ���"
         Top             =   660
         Width           =   315
      End
      Begin VB.CheckBox CHK102 
         Caption         =   "��Ҫԡ"
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
         Left            =   1080
         TabIndex        =   5
         Top             =   315
         Width           =   1890
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   3315
         Left            =   225
         TabIndex        =   16
         Top             =   2385
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   5847
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "�����Թ���"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "�����Թ���"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "�Ҥһ���"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "�Ҥ��������"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ŵ�Ҥ� (�ҷ)"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "��ǹŴ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "˹��¹Ѻ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "�����˵�"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "���������Ŵ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "��Ҫԡ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "IsBrochure"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "IsCancel"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "PromotoinType"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "PromotoinTypeCode"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.TextBox ItemDetail106 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7875
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1470
         Width           =   2685
      End
      Begin VB.TextBox ItemDetail104 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4230
         TabIndex        =   8
         Top             =   1065
         Width           =   2415
      End
      Begin VB.TextBox ItemDetail103 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   1065
         Width           =   1890
      End
      Begin VB.TextBox ItemDetail102 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4230
         TabIndex        =   30
         Top             =   660
         Width           =   9750
      End
      Begin VB.TextBox ItemDetail105 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8955
         TabIndex        =   10
         Top             =   1035
         Width           =   1605
      End
      Begin VB.TextBox ItemDetail101 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   660
         Width           =   1890
      End
      Begin VB.Label LBLWholeSaleCom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5130
         TabIndex        =   52
         Top             =   1935
         Width           =   1140
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "��Ҥ���Թ����"
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
         Left            =   3825
         TabIndex        =   51
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label LBLToComDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   12645
         TabIndex        =   50
         Top             =   1935
         Width           =   1320
      End
      Begin VB.Label LBLFormComDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   10530
         TabIndex        =   49
         Top             =   1935
         Width           =   1320
      End
      Begin VB.Label LBLCampaignCom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7290
         TabIndex        =   48
         Top             =   1935
         Width           =   2445
      End
      Begin VB.Label LBLRetailCom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2610
         TabIndex        =   47
         Top             =   1935
         Width           =   1140
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "��໭"
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
         Left            =   6480
         TabIndex        =   46
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   11700
         TabIndex        =   45
         Top             =   1980
         Width           =   870
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "�ҡ�ѹ���"
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
         Left            =   9675
         TabIndex        =   44
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label Label8 
         Caption         =   "��Ҥ���Թʴ"
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
         Left            =   1485
         TabIndex        =   43
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label LBLHotPrice 
         Alignment       =   1  'Right Justify
         Caption         =   "HotPrice "
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
         Left            =   10215
         TabIndex        =   42
         Top             =   1485
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label5 
         Caption         =   "�������������"
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
         Left            =   2835
         TabIndex        =   37
         Top             =   1485
         Width           =   1290
      End
      Begin VB.Label Label7 
         Caption         =   "�����˵�"
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
         Left            =   7080
         TabIndex        =   29
         Top             =   1485
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "��ǹŴ"
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
         Left            =   8280
         TabIndex        =   28
         Top             =   1065
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "�Ҥһ���"
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
         Left            =   180
         TabIndex        =   27
         Top             =   1065
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "˹��¹Ѻ"
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
         Left            =   3480
         TabIndex        =   26
         Top             =   1065
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "�����Թ���"
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
         Left            =   180
         TabIndex        =   25
         Top             =   675
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "�����Թ���"
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
         Left            =   3510
         TabIndex        =   24
         Top             =   675
         Width           =   690
      End
   End
   Begin VB.CommandButton CMD107 
      Caption         =   "�ѹ�֡"
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
      Left            =   9090
      TabIndex        =   22
      ToolTipText     =   "������� ������Ѻ�ѹ�֡����Ѿഷ"
      Top             =   8100
      Width           =   1440
   End
   Begin VB.Frame Frame2 
      Caption         =   "��������´��ʹ��Թ���"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   75
      TabIndex        =   17
      Top             =   75
      Width           =   14250
      Begin VB.CommandButton CMD112 
         Height          =   315
         Left            =   4575
         Picture         =   "Form201.frx":0AE8
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "������� ������Ѻź�͡��� (��Ѻ IScancel = 1����)"
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton CMD106 
         Height          =   315
         Left            =   4200
         Picture         =   "Form201.frx":0F5A
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "������� ������Ѻ�����͡����ʹ��Թ����������"
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton CMD103 
         Height          =   315
         Left            =   10200
         Picture         =   "Form201.frx":1327
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "������� ������Ѻ���� Section Manager"
         Top             =   825
         Width           =   315
      End
      Begin VB.CommandButton CMD102 
         Height          =   315
         Left            =   10200
         Picture         =   "Form201.frx":16F4
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "������� ������Ѻ���ҷ���¹�������"
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox Text103 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6825
         TabIndex        =   4
         Top             =   825
         Width           =   3315
      End
      Begin VB.TextBox Text102 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6825
         TabIndex        =   3
         Top             =   300
         Width           =   3315
      End
      Begin MSComCtl2.DTPicker DTPicker101 
         Height          =   315
         Left            =   2250
         TabIndex        =   2
         Top             =   825
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20512769
         CurrentDate     =   38504
      End
      Begin VB.CommandButton CMD101 
         Height          =   315
         Left            =   3825
         Picture         =   "Form201.frx":1AC1
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "������� ������Ѻ���ҧ�Ţ����͡����Ţ�������"
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox Text101 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   2250
         TabIndex        =   0
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label LBLPromoStopDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11880
         TabIndex        =   54
         Top             =   810
         Width           =   1500
      End
      Begin VB.Label LBLPromoStartDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11880
         TabIndex        =   53
         Top             =   315
         Width           =   1500
      End
      Begin VB.Label LBLUserID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   10575
         TabIndex        =   40
         Top             =   810
         Width           =   1140
      End
      Begin VB.Image Image101 
         Height          =   300
         Left            =   150
         Picture         =   "Form201.frx":1E18
         ToolTipText     =   "�ʴ�ʶҹ��͡��� N : New CF : Confirm CC : Cancel"
         Top             =   225
         Width           =   570
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Sec. Manager "
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
         Left            =   5250
         TabIndex        =   21
         Top             =   825
         Width           =   1440
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "�������"
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
         Left            =   5700
         TabIndex        =   20
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label16 
         Caption         =   "�ѹ����͡���"
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
         Left            =   1125
         TabIndex        =   19
         Top             =   825
         Width           =   990
      End
      Begin VB.Label Label15 
         Caption         =   "�Ţ����͡���"
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
         Left            =   1125
         TabIndex        =   18
         Top             =   300
         Width           =   990
      End
      Begin VB.Image Image102 
         Height          =   300
         Left            =   150
         Picture         =   "Form201.frx":224A
         Top             =   225
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Image Image103 
         Height          =   300
         Left            =   150
         Picture         =   "Form201.frx":26F3
         Top             =   225
         Visible         =   0   'False
         Width           =   570
      End
   End
End
Attribute VB_Name = "Form201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNEwDocno As String
Dim vCheckRequestOnListview As Integer
Dim vIndexItemUpdate As Integer
Dim vSortResult As Integer

Private Sub CHK102_Click()
Dim vDiscount As Currency

On Error Resume Next

If CHK102.Value = 1 Then
    ItemDetail105.Enabled = False
    CHK103.Enabled = False
    vDiscount = (ItemDetail103.Text - (ItemDetail103.Text * vMemberDiscount) / 100)
    ItemDetail105.Text = vDiscount
Else
    ItemDetail105.Enabled = True
    CHK103.Enabled = True
    CHK103.Value = 0
    ItemDetail105.Text = 0
End If
ItemDetail101.SetFocus


End Sub

Private Sub CHK103_Click()
On Error Resume Next
ItemDetail105.SetFocus
End Sub

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

vCheckJob1 = 1
vQuery = "execute USP_PM_RequestNewDocNo"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vNEwDocno = Trim(vRecordset.Fields("newdocno").Value)
End If
vRecordset.Close
Text101.Text = vNEwDocno
Me.LBLUserID.Caption = ""

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD102_Click()
MDIForm1.Enabled = False
vMemCommand = 1
FormSearchMainPromotion.Show
End Sub

Private Sub CMD103_Click()
MDIForm1.Enabled = False
vMemCommand = 1
FormSearchSecMan.Show
End Sub

Private Sub CMD104_Click()
MDIForm1.Enabled = False
FormSearchItem.Show
FormSearchItem.Text101.Text = Trim(Form201.ItemDetail101.Text)
End Sub

Private Sub CMD105_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItemCode As ListItem
Dim vDiscount As Currency
Dim vTypeDiscount As Integer
Dim vCountText As String
Dim vPromoPrice As Currency
Dim vIsMember As Integer
Dim vMydescription As String
Dim vUnitCode As String
Dim vDocno As String
Dim vIsCancel As Integer
Dim vIsConfirm As Integer
Dim vIsBrochure As Integer
Dim vPromotionType As String
Dim vPromotionTypeCode As String
Dim vCheckDocno As String
Dim vCheckPMCode As String
Dim vCheckItemDuplicate As Integer
Dim vCheckItemInRequest As Integer
Dim vCheckDuplicatePromotion As Integer
Dim vCheckDuplicateDocno As String
Dim vDiscount1 As Currency

On Error GoTo ErrDescription

    If Text101.Text <> "" And Text102.Text <> "" Then
        If ItemDetail107.Text <> "" Then
        vDocno = Trim(Text101.Text)
        vQuery = "select  *  from npmaster.dbo.tb_pm_request  where docno = '" & vDocno & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vIsCancel = Trim(vRecordset.Fields("iscancel").Value)
            vIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
        Else
            vIsCancel = 0
            vIsConfirm = 0
        End If
        vRecordset.Close
        Else
        MsgBox "��س����͡ �������Թ���������蹴��¤�Ѻ"
        CMD111.SetFocus
        Exit Sub
        End If
        If vIsCancel <> 1 And vIsConfirm <> 2 Then
        If vCheckRequestOnListview = 0 Then
            '----------------------------------------------------------------------------
            vCheckPMCode = Left(Trim(Text102.Text), InStr((Text102.Text), "/") - 1)
            vQuery = "select itemcode from npmaster.dbo.TB_PM_TempCheckItemDuplicateLine where itemcode = '" & Trim(ItemDetail101.Text) & "' and docno = '" & vDocno & "' and pmcode = '" & vCheckPMCode & "' and unitcode = '" & Trim(ItemDetail104.Text) & "'  "
            If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckItemDuplicate = 1 '******************************************************
            Else
                vCheckItemDuplicate = 0
            End If
            vRecordset.Close
            
            vQuery = "select itemcode from npmaster.dbo.TB_PM_requestsub  where itemcode = '" & Trim(ItemDetail101.Text) & "' and docno = '" & vDocno & "' and unitcode = '" & Trim(ItemDetail104.Text) & "'  "
            If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckItemInRequest = 1 '*******************************************************
            Else
                vCheckItemInRequest = 0
            End If
            vRecordset.Close
            
            vQuery = "exec USP_PM_ItemDuplicate  '" & Left(Trim(Text102.Text), InStr(Trim(Text102.Text), "/") - 1) & "', '" & Trim(ItemDetail101.Text) & "' , '" & Trim(ItemDetail104.Text) & "'  "
            If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckDuplicatePromotion = Trim(vRecordset.Fields("isduplicate").Value)
                vCheckDuplicateDocno = Trim(vRecordset.Fields("duplicate").Value)
            End If
            vRecordset.Close
            If vCheckDuplicatePromotion = 0 Then
        '--------------------------------------------------------------------------------
        If vCheckItemDuplicate = 0 And vCheckItemInRequest = 0 Then
        
            If CHK102.Value = 1 Then
            vIsMember = 1
            vTypeDiscount = 2
            vDiscount = (ItemDetail103.Text - ItemDetail105.Text)
            vCountText = Trim("Member")
            vPromoPrice = ItemDetail105.Text
            If Check101.Value = 1 Then
                vIsBrochure = 1
            Else
                vIsBrochure = 0
            End If
            Else
            If CHK103.Value = 1 Then
            vTypeDiscount = 1
            vDiscount = ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
            vCountText = Trim(ItemDetail105.Text & "%")
            vPromoPrice = Trim(ItemDetail103.Text) - ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
            Else
            vTypeDiscount = 0
            If Trim(ItemDetail105.Text) <> "" Then
                vDiscount = Trim(ItemDetail105.Text)
                vPromoPrice = Trim(ItemDetail103.Text) - Trim(ItemDetail105.Text)
            Else
                vDiscount = 0
                vPromoPrice = Trim(ItemDetail103.Text) - 0
            End If
            vCountText = Trim(ItemDetail105.Text)
            If Check101.Value = 1 Then
                vIsBrochure = 1
            Else
                vIsBrochure = 0
            End If
            End If
            vIsMember = 0
            End If
            vMydescription = Trim(ItemDetail106.Text)
            vUnitCode = Trim(ItemDetail104.Text)
            vPromotionType = Trim(ItemDetail107.Text)
            vPromotionTypeCode = Left(Trim(ItemDetail107.Text), 2)
            
            Set vListItemCode = ListView101.ListItems.Add(, , Trim(ItemDetail101.Text))
            vListItemCode.SubItems(1) = Trim(ItemDetail102.Text)
            vListItemCode.SubItems(2) = Trim(ItemDetail103.Text)
            vListItemCode.SubItems(3) = vPromoPrice
            vListItemCode.SubItems(4) = vDiscount 'vCountText
            vListItemCode.SubItems(5) = vCountText 'vTypeDiscount
            vListItemCode.SubItems(6) = vUnitCode 'vCountText 'vDisCount
            vListItemCode.SubItems(7) = vMydescription 'vUnitCode
            vListItemCode.SubItems(8) = vTypeDiscount 'vIsMember
            vListItemCode.SubItems(9) = vIsMember 'vMydescription
            vListItemCode.SubItems(10) = vIsBrochure
            vListItemCode.SubItems(11) = vIsCancel
            vListItemCode.SubItems(12) = vPromotionType
            vListItemCode.SubItems(13) = vPromotionTypeCode
            vListItemCode.Checked = True
'            ---------------------------------------------------------------------------------------------
            vCheckDocno = Trim(Text101.Text)
            vCheckPMCode = Left(Trim(Text102.Text), InStr((Text102.Text), "/") - 1)
            vQuery = "execute USP_PM_InsertCheckDuplicatItemLine '" & Trim(ItemDetail101.Text) & "','" & vCheckDocno & "','" & vCheckPMCode & "','" & vUserID & "','" & Trim(ItemDetail104.Text) & "' "
            gConnection.Execute vQuery
            Else
                MsgBox "�������ö�����Թ������ǡѹ ��͡���㺹����"
                ItemDetail101.Text = ""
                ItemDetail102.Text = ""
                ItemDetail103.Text = ""
                ItemDetail104.Text = ""
                ItemDetail105.Text = ""
                ItemDetail106.Text = ""
                ItemDetail107.Text = ""
                CHK102.Value = 0
                CHK103.Value = 0
                Check101.Value = 0
                Exit Sub
            End If
            '----------------------------------------------------------------------------------------------------
            Else
            MsgBox "�Թ��� ���� " & Trim(ItemDetail101.Text) & " ��ӡѹ�������蹹�� �Ѻ�Ţ����ʹ��Թ��� " & vCheckDuplicateDocno & "  ��سҵ�Ǩ�ͺ"
            End If
    Else

        If UCase(vUserID) = UCase(Trim(Me.LBLUserID.Caption)) Then
            ItemDetail101.Enabled = False
                If CHK102.Value = 1 Then
                vIsMember = 1
                    vTypeDiscount = 2
                    vDiscount = ItemDetail103.Text - ItemDetail105.Text
                    vCountText = Trim("Member")
                    vPromoPrice = ItemDetail105.Text
                    If Check101.Value = 1 Then
                        vIsBrochure = 1
                    Else
                        vIsBrochure = 0
                    End If
                    vPromotionType = Trim(ItemDetail107.Text)
                    vPromotionTypeCode = Left(Trim(ItemDetail107.Text), 2)
            Else
                If CHK103.Value = 1 Then
                    vTypeDiscount = 1
                    vDiscount = ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
                    vCountText = Trim(ItemDetail105.Text & "%")
                    vPromoPrice = Trim(ItemDetail103.Text) - ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
                Else
                    vTypeDiscount = 0
                    If Trim(ItemDetail105.Text) <> "" Then
                        vDiscount = Trim(ItemDetail105.Text)
                        vPromoPrice = Trim(ItemDetail103.Text) - Trim(ItemDetail105.Text)
                    Else
                        vDiscount = 0
                        vPromoPrice = Trim(ItemDetail103.Text) - 0
                    End If
                    vCountText = Trim(ItemDetail105.Text)
    
                End If
                    If Check101.Value = 1 Then
                        vIsBrochure = 1
                    Else
                        vIsBrochure = 0
                    End If
                vPromotionType = Trim(ItemDetail107.Text)
                vPromotionTypeCode = Left(Trim(ItemDetail107.Text), 2)
                vIsMember = 0
                If ListView101.ListItems.Item(vIndexItemUpdate).SubItems(11) = 1 Then
                    vIsCancel = 1
                Else
                    vIsCancel = 0
                End If
            End If
                vMydescription = Trim(ItemDetail106.Text)
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(3) = vPromoPrice
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(4) = vDiscount 'vCountText
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(5) = vCountText 'vTypeDiscount
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(7) = vMydescription 'vCountText 'vDisCount
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(8) = vTypeDiscount 'vIsMember
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(9) = vIsMember 'vMydescription
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(10) = vIsBrochure
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(11) = vIsCancel
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(12) = vPromotionType
                ListView101.ListItems.Item(vIndexItemUpdate).SubItems(13) = vPromotionTypeCode
                ListView101.ListItems.Item(vIndexItemUpdate).Checked = True
            Else
                MsgBox "�������ö������������Թ����� ���ͧ�ҡ User �Ѻ Section Manager ���ç�ѹ"
            End If
    End If
        On Error Resume Next
        ItemDetail101.Text = ""
        ItemDetail102.Text = ""
        ItemDetail103.Text = ""
        ItemDetail104.Text = ""
        ItemDetail105.Text = ""
        ItemDetail106.Text = ""
        ItemDetail107.Text = ""
        ItemDetail105.Enabled = True
        ItemDetail101.Enabled = True
        Check101.Value = 0
        CHK103.Enabled = True
        ItemDetail101.SetFocus
        CHK103.Value = 0
        CHK102.Value = 0
        vCheckRequestOnListview = 0
    Else
        If vIsCancel = 1 Then
            MsgBox "�͡����Ţ��� " & vDocno & " ��¡��ԡ����� �������ö��䢢�������"
        ElseIf vIsConfirm = 2 Then
            MsgBox "�͡����Ţ��� " & vDocno & " ��͹��ѵ������ �������ö��䢢�������"
        End If
        On Error Resume Next
        ItemDetail101.Text = ""
        ItemDetail102.Text = ""
        ItemDetail103.Text = ""
        ItemDetail104.Text = ""
        ItemDetail105.Text = ""
        ItemDetail106.Text = ""
        ItemDetail107.Text = ""
        ItemDetail105.Enabled = True
        ItemDetail101.Enabled = True
        Check101.Value = 0
        CHK103.Enabled = True
        ItemDetail101.SetFocus
        CHK103.Value = 0
        CHK102.Value = 0
    End If
Else
    MsgBox "��������Թ���� ���ҧ��ҧ��ҧ ��ͧ����Ţ����͡����������������蹡�͹�Ф�Ѻ"
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD106_Click()
MDIForm1.Enabled = False
ItemDetail101.Enabled = True
FormSearchReqPromo.Show
End Sub

Private Sub CMD107_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vPromoname As String
Dim vStartPromo As Date
Dim vIsCancel As String
Dim vCountItem As Integer
Dim vSecName As String
Dim vPromotionCode As String
Dim i As Integer
Dim vError As Integer
Dim vIsCompleteSave As Integer
Dim vItemCode As String
Dim vItemName As String
Dim vUnitCode As String
Dim vNormalPrice As Currency
Dim vFromQTY As Long, vToQty As Long
Dim vDiscount As Currency
Dim vDiscountWord As String
Dim vDiscountType As String
Dim vPromotionPrice As Currency
Dim vMydescription As String
Dim vLineNumber As Integer
Dim vIsBrochure As String
Dim vIsMember As String
Dim vIsConfirm As Integer
Dim vPromotionType As String
Dim vPromotionTypeCode As String
Dim vItemIsCancel As Integer
Dim vCheckDeleteDocno As String
Dim vCheckDuplicatePromotion As Integer
Dim vCheckDuplicateDocno As String
Dim vHotPrice As String

On Error GoTo ErrDescription

Me.CMD107.Enabled = False

If Trim(Text101.Text) <> "" And ListView101.ListItems.Count <> 0 Then
vCountItem = ListView101.ListItems.Count
If vCheckJob1 = 1 Then
    If vCountItem > 0 Then
        If Text101.Text <> "" Then
            If Text102.Text <> "" Or Text103.Text <> "" Then
                vStartPromo = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
                vSecName = Trim(Text103.Text)
                vPromotionCode = Left(Trim(Text102.Text), InStr(Trim(Text102.Text), "/") - 1)
                If vCheckJob1 <> 0 Then
                    vQuery = "execute USP_PM_RequestNewDocNo"
                    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                        vNEwDocno = Trim(vRecordset.Fields("newdocno").Value)
                    End If
                    vRecordset.Close
                Else
                    vNEwDocno = Trim(Text101.Text)
                End If
                
                vQuery = "execute USP_PM_InsertRequest " & vCheckJob1 & ",'" & vNEwDocno & "','" & vStartPromo & "','" & vSecName & "','" & vPromotionCode & "','" & vUserID & "' "
                gConnection.Execute vQuery
                
                For i = 1 To ListView101.ListItems.Count
                vError = 0
                If i = ListView101.ListItems.Count Then
                    vIsCompleteSave = 1
                Else
                    vIsCompleteSave = 0
                End If
                vItemCode = Trim(ListView101.ListItems.Item(i).Text)
                vIsCompleteSave = 1
                vError = 0
                vItemName = Trim(ListView101.ListItems.Item(i).SubItems(1))
                vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(6))
                vNormalPrice = Trim(ListView101.ListItems.Item(i).SubItems(2))
                vFromQTY = 1
                vToQty = 99999
                If Trim(ListView101.ListItems.Item(i).SubItems(8)) <> 2 Then
                    vDiscount = Trim(ListView101.ListItems.Item(i).SubItems(4))
                Else
                    vDiscount = 0
                End If
                vDiscountType = Trim(ListView101.ListItems.Item(i).SubItems(8))
                vDiscountWord = Trim(ListView101.ListItems.Item(i).SubItems(5))
                vPromotionPrice = Trim(ListView101.ListItems.Item(i).SubItems(3))
                vMydescription = Trim(ListView101.ListItems.Item(i).SubItems(7))
                vLineNumber = i - 1
                vIsBrochure = Trim(ListView101.ListItems.Item(i).SubItems(10))
                If Trim(ListView101.ListItems.Item(i).SubItems(9)) = 0 Then
                    vIsMember = 0
                Else
                    vIsMember = 1
                End If
                vPromotionType = Trim(ListView101.ListItems.Item(i).SubItems(12))
                vPromotionTypeCode = Trim(ListView101.ListItems.Item(i).SubItems(13))
                
                If vPromotionTypeCode = "11" Then
                vHotPrice = "S02"
                Else
                vHotPrice = ""
                End If
                
                If ListView101.ListItems.Item(i).Checked = True Then
                    vItemIsCancel = 0
                Else
                    vItemIsCancel = 0
                End If
                
                vQuery = "exec USP_PM_ItemDuplicate  '" & Trim(vPromotionCode) & "', '" & Trim(vItemCode) & "','" & vUnitCode & "'  "
                If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                    vCheckDuplicatePromotion = Trim(vRecordset.Fields("isduplicate").Value)
                    vCheckDuplicateDocno = Trim(vRecordset.Fields("duplicate").Value)
                End If
                vRecordset.Close
                If vCheckDuplicatePromotion = 0 Then
                    vQuery = "execute USP_PM_InsertRequestSub " & vError & "," & vIsCompleteSave & ",'" & vNEwDocno & "','" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vNormalPrice & "," & vFromQTY & "," & vToQty & "," & vDiscount & ",'" & vDiscountType & "','" & vDiscountWord & "'," & vPromotionPrice & ",'" & vMydescription & "','" & vItemIsCancel & "'," & vLineNumber & ",'" & vIsBrochure & "','" & vIsMember & "' ,'" & vPromotionTypeCode & "','" & vHotPrice & "' "
                    gConnection.Execute vQuery
                Else
                    vQuery = "execute USP_PM_InsertRequestSub 1," & vIsCompleteSave & ",'" & vNEwDocno & "','" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vNormalPrice & "," & vFromQTY & "," & vToQty & "," & vDiscount & ",'" & vDiscountType & "','" & vDiscountWord & "'," & vPromotionPrice & ",'" & vMydescription & "','" & vItemIsCancel & "'," & vLineNumber & ",'" & vIsBrochure & "','" & vIsMember & "' ,'" & vPromotionTypeCode & "','" & vHotPrice & "'  "
                    gConnection.Execute vQuery
                    MsgBox "��¡���Թ��� ���� " & vItemCode & " �������蹹������������ ��͡����Ţ��� " & vCheckDuplicateDocno & " ��سҵ�Ǩ�ͺ", vbCritical, "Send Error"
                    Me.CMD107.Enabled = True
                    Exit Sub
                End If
                
                Next i
                MsgBox "���͡����Ţ���  " & vNEwDocno & " "
                Call InitializeSendEmail
                vQuery = "execute USP_PM_DeliverySendMail '" & vNEwDocno & "' "
                vGetConnect.Execute vQuery
                
                vCheckDeleteDocno = Trim(Text101.Text)
                vQuery = "USP_PM_DeleteCheckDuplicatItemLine '" & vCheckDeleteDocno & "','" & vPromotionCode & "','" & vUserID & "' "
                gConnection.Execute vQuery
                Me.CMD107.Enabled = True
                
            Else
                MsgBox "��س����͡ ������������� ��� Section Manager ���¤�Ѻ"
                Me.CMD107.Enabled = True
                Exit Sub
            End If
        Else
            MsgBox "��س� ���������ҧ�Ţ����͡��ô��¤�Ѻ"
            Me.CMD107.Enabled = True
            Exit Sub
        End If
    Else
        MsgBox "������Թ��ҷ��зӡ���ʹͷ��������"
        Me.CMD107.Enabled = True
        Exit Sub
    End If
ElseIf vCheckJob1 = 0 Then
    vCheckJob1 = 0
    vNEwDocno = Trim(Text101.Text)
    vStartPromo = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
    vSecName = Trim(Text103.Text)
    vPromotionCode = Left(Trim(Text102.Text), InStr(Trim(Text102.Text), "/") - 1)
    
    vQuery = "select  *  from npmaster.dbo.tb_pm_request where docno = '" & vNEwDocno & "' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vIsCancel = Trim(vRecordset.Fields("iscancel").Value)
        vIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
    End If
    vRecordset.Close
    If vIsCancel = 0 And vIsConfirm = 0 Then
        vQuery = "execute USP_PM_InsertRequest " & vCheckJob1 & ",'" & vNEwDocno & "','" & vStartPromo & "','" & vSecName & "','" & vPromotionCode & "','" & vUserID & "' "
        gConnection.Execute vQuery
        If ListView101.ListItems.Count <> 0 Then
        For i = 1 To ListView101.ListItems.Count
        vError = 0
        If i = ListView101.ListItems.Count Then
            vIsCompleteSave = 1
        Else
            vIsCompleteSave = 0
        End If
        vItemCode = Trim(ListView101.ListItems.Item(i).Text)
        vIsCompleteSave = 1
        vError = 0
        vItemName = Trim(ListView101.ListItems.Item(i).SubItems(1))
        vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(6))
        vNormalPrice = Trim(ListView101.ListItems.Item(i).SubItems(2))
        vFromQTY = 1
        vToQty = 99999
        vDiscount = Trim(ListView101.ListItems.Item(i).SubItems(4))
        vDiscountType = Trim(ListView101.ListItems.Item(i).SubItems(8))
        vDiscountWord = Trim(ListView101.ListItems.Item(i).SubItems(5))
        vPromotionPrice = Trim(ListView101.ListItems.Item(i).SubItems(3))
        vMydescription = Trim(ListView101.ListItems.Item(i).SubItems(7))
        vLineNumber = i - 1
        If Trim(ListView101.ListItems.Item(i).SubItems(10)) = 0 Then
        vIsBrochure = 0
        Else
        vIsBrochure = 1
        End If
        vPromotionType = Trim(ListView101.ListItems.Item(i).SubItems(12))
        vPromotionTypeCode = Trim(ListView101.ListItems.Item(i).SubItems(13))
        
        If vPromotionTypeCode = "11" Then
            vHotPrice = "S02"
        Else
            vHotPrice = ""
        End If
                
        If Trim(ListView101.ListItems.Item(i).SubItems(9)) = 0 Then
            vIsMember = 0
        Else
            vIsMember = 1
        End If
        If Trim(ListView101.ListItems.Item(i).SubItems(11)) = 0 Then
            vItemIsCancel = 0
        Else
            vItemIsCancel = 1
        End If
        
        vQuery = "execute USP_PM_InsertRequestSub " & vError & "," & vIsCompleteSave & ",'" & vNEwDocno & "','" & vItemCode & "','" & vItemName & "','" & vUnitCode & "'," & vNormalPrice & "," & vFromQTY & "," & vToQty & "," & vDiscount & ",'" & vDiscountType & "','" & vDiscountWord & "'," & vPromotionPrice & ",'" & vMydescription & "','" & vItemIsCancel & "'," & vLineNumber & ",'" & vIsBrochure & "','" & vIsMember & "','" & vPromotionTypeCode & "','" & vHotPrice & "' "
        gConnection.Execute vQuery
        Next i
        
        'If vCheckDuplicatePromotion = 0 Then
        MsgBox "��Ѻ��ا�͡����Ţ���  " & vNEwDocno & " ���º�������Ǥ�Ѻ"
        'Call InitializeSendEmail
        'vQuery = "execute USP_PM_DeliverySendMail '" & vNewDocno & "' "
        'vGetConnect.Execute vQuery
        'End If
        
        vCheckDeleteDocno = Trim(Text101.Text)
        vQuery = "USP_PM_DeleteCheckDuplicatItemLine '" & vCheckDeleteDocno & "','" & vPromotionCode & "','" & vUserID & "' "
        gConnection.Execute vQuery
                
        Me.CMD107.Enabled = True

        End If
    End If
End If
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Me.LBLUserID.Caption = ""
    ItemDetail101.Text = ""
    ItemDetail102.Text = ""
    ItemDetail103.Text = ""
    ItemDetail104.Text = ""
    ItemDetail105.Text = ""
    ItemDetail106.Text = ""
    ItemDetail107.Text = ""
    CHK102.Value = 0
    CHK103.Value = 0
    Check101.Value = 0
    ListView101.ListItems.Clear
    Image101.Visible = True
    Image102.Visible = False
    Image103.Visible = False
    Me.CMD107.Enabled = True
Else
    MsgBox "�͡����ʹ��Թ��� �������¡���Թ������ҧ���� 1 ��¡�ö֧�кѹ�֡�����Ѿഷ��"
    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    DTPicker101.Value = Now
    Me.CMD107.Enabled = True
End If

ErrDescription:
If Err.Description <> "" Then
    If Err.Number = "-2147217873" Then
        MsgBox Err.Description
        MsgBox "�����Թ��ҷ���ʹ�������蹫�ӡѹ ��سҵ�Ǩ�ͺ"
        Me.CMD107.Enabled = True
    Else
        MsgBox Err.Description
        Me.CMD107.Enabled = True
    End If
Exit Sub
End If
End Sub

Private Sub CMD108_Click()
Dim i As Integer
Dim vCheckDelete As Integer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocno As String
Dim vCheckPromo As String
Dim vCheckItemCode As String

On Error Resume Next

If ListView101.ListItems.Count <> 0 Then
    i = ListView101.ListItems.Item(vIndexItemUpdate).Index
    vCheckDelete = MsgBox("��ͧ���ź��¡���Թ���  " & ListView101.ListItems.Item(vIndexItemUpdate).Text & "  ����������", vbYesNo, "���׹�ѹ㹡��ź��¡���Թ���")
    If vCheckDelete = 6 Then
        vCheckDocno = Trim(Text101.Text)
        vCheckPromo = Left(Trim(Text102.Text), InStr(Trim(Text102.Text), "/") - 1)
        vCheckItemCode = Trim(ListView101.ListItems.Item(vIndexItemUpdate).Text)
        vQuery = "USP_PM_DeleteCheckDuplicatItemLine '" & vCheckDocno & "','" & vCheckPromo & "','" & vUserID & "' ,'" & vCheckItemCode & "' "
        gConnection.Execute vQuery
        
        ItemDetail101.Enabled = True
        ListView101.ListItems.Remove (i)
        ItemDetail101.Text = ""
        ItemDetail102.Text = ""
        ItemDetail103.Text = ""
        ItemDetail104.Text = ""
        ItemDetail105.Text = ""
        ItemDetail106.Text = ""
        ItemDetail107.Text = ""
        CHK102.Value = 0
        CHK103.Value = 0
        Check101.Value = 0
        ItemDetail101.SetFocus
        CMD108.Enabled = False
    Else
        ItemDetail101.Enabled = True
        ItemDetail101.Text = ""
        ItemDetail102.Text = ""
        ItemDetail103.Text = ""
        ItemDetail104.Text = ""
        ItemDetail105.Text = ""
        ItemDetail106.Text = ""
        ItemDetail107.Text = ""
        CHK102.Value = 0
        CHK103.Value = 0
        Check101.Value = 0
        ItemDetail101.SetFocus
        CMD108.Enabled = False
        Exit Sub
    End If
Else
    MsgBox "������Թ������ź��¡��"
End If

End Sub

Private Sub CMD109_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocno As String
Dim vCheckPromo As String

On Error GoTo ErrDescription

If Trim(Text101.Text) <> "" Then
vCheckDocno = Trim(Text101.Text)
vCheckPromo = Left(Trim(Text102.Text), InStr(Trim(Text102.Text), "/") - 1)
vQuery = "USP_PM_DeleteCheckDuplicatItemLine '" & vCheckDocno & "','" & vCheckPromo & "','" & vUserID & "'  "
gConnection.Execute vQuery
End If
        
Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
Me.LBLUserID.Caption = ""
DTPicker101.Value = Now
ItemDetail101.Text = ""
ItemDetail102.Text = ""
ItemDetail103.Text = ""
ItemDetail104.Text = ""
ItemDetail105.Text = ""
ItemDetail106.Text = ""
ItemDetail107.Text = ""
ItemDetail101.Enabled = True
ListView101.ListItems.Clear
Image101.Visible = True
Image102.Visible = False
Image103.Visible = False
CMD108.Enabled = False
CHK102.Value = 0
CHK103.Value = 0
Check101.Value = 0
vCheckRequestOnListview = 0

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD110_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vDocno As String

On Error GoTo ErrDescription

If vCheckStatusPrint = 0 Then
    vDocno = Trim(Text101.Text)
    If vDocno <> "" Then
    vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 269 and reptype = 'PM' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
    
    With Crystal101
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@vDocno;" & vDocno & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    End With
    End If
Else
    MsgBox "�������ö������͡��÷�� ��Ǩ�ͺ���� ���� ͹��ѵ����� ���� ¡��ԡ �� �óյ�ͧ��èо����Դ���Ἱ�����������", vbInformation, "��ͤ�������͹"
    Exit Sub
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD111_Click()
MDIForm1.Enabled = False
FormSearchType.Show
End Sub

Private Sub CMD112_Click()
Dim vQuery  As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno As String
Dim vAnswer As Integer
Dim vCheckDocno As Integer
Dim vCheckUser As String
Dim vIsConfirm As Integer

On Error GoTo ErrDescription

If Text103.Text <> "" Then
    vCheckUser = UCase(Trim(Text103.Text))
    If UCase(vUserID) = vCheckUser Then
        If Trim(Text101.Text) <> "" Then
            vDocno = Trim(Text101.Text)
            vQuery = "select docno,isconfirm from npmaster.dbo.tb_pm_request where docno = '" & vDocno & "' "
            If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
                vCheckDocno = 1
                vIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
            Else
                vCheckDocno = 0
            End If
            vRecordset.Close
            If vIsConfirm <> 2 Then
            If vCheckDocno = 1 Then
            vAnswer = MsgBox("�س��ͧ���ź�͡����Ţ��� " & vDocno & " ������������", vbYesNo, "�Ӷ���׹�ѹ���ź")
            
            If vAnswer = 6 Then
                If vDocno <> "" Then
                    vQuery = "exec USP_PM_DeletePMRequest  '" & vDocno & "' "
                    gConnection.Execute vQuery
                    MsgBox "��ӡ��ź�͡����Ţ��� " & vDocno & " ���º�������� "
                End If
                Text101.Text = ""
                Text102.Text = ""
                Text103.Text = ""
                Me.LBLUserID.Caption = ""
                DTPicker101.Value = Now
                ListView101.ListItems.Clear
                Else
                Exit Sub
            End If
            Else
                MsgBox "�͡����Ţ��� " & vDocno & " �ѧ�����"
            End If
            Else
                MsgBox "�͡����Ţ��� " & vDocno & " ��͹��ѵ������������öź�͡�����"
            End If
        End If
    Else
        MsgBox "�س������Է���㹡��ź�͡����Ţ�����"
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

'Private Sub CMD113_Click()
'Dim i As Integer
'Dim vCheckIsCancel As Integer
'
'On Error Resume Next

'If ListView101.ListItems.Count <> 0 Then
 '   i = ListView101.ListItems.Item(vIndexItemUpdate).Index
  '  vCheckIsCancel = MsgBox("��ͧ���¡��ԡ��¡���Թ���  " & ListView101.ListItems.Item(vIndexItemUpdate).Text & "  ����������", vbYesNo, "���׹�ѹ㹡��ź��¡���Թ���")
   ' If vCheckIsCancel = 6 Then
    '    ItemDetail101.Enabled = True
     '   ItemDetail101.SetFocus
      '  ListView101.ListItems.Item(i).SubItems(11) = 1
       ' ItemDetail101.Enabled = True
        'ItemDetail101.Text = ""
        'ItemDetail102.Text = ""
        'itemDetail103.Text = ""
        'ItemDetail104.Text = ""
        'ItemDetail105.Text = ""
        'ItemDetail106.Text = ""
        'ItemDetail107.Text = ""
        'CHK102.Value = 0
        'CHK103.Value = 0
        'Check101.Value = 0
        'ItemDetail101.SetFocus
        'CMD108.Enabled = False
    'Else
     '   Exit Sub
    'End If
'Else
 '   MsgBox "������Թ������¡��ԡ��¡��"
'End If

'End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vMemberDisc As String

DTPicker101 = Now
vCheckJob1 = 1
vCheckRequestOnListview = 0
CHK102.Caption = "��Ҫԡ"
vQuery = "select memberdisc  from bpsconfig"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vMemberDisc = Left(Trim(vRecordset.Fields("memberdisc").Value), InStr(Trim(vRecordset.Fields("memberdisc").Value), "%") - 1)
End If
vRecordset.Close
CHK102.Caption = CHK102.Caption & "    " & "( Ŵ " & vMemberDisc & "% )"
vMemberDiscount = vMemberDisc

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

vQuery = "delete npmaster.dbo. TB_PM_TempCheckItemDuplicateLine where userid = '" & vUserID & "'  "
gConnection.Execute vQuery

End Sub

Private Sub ItemDetail101_GotFocus()
If Text102.Text = "" Or Text103.Text = "" Then
    MsgBox "��س����͡ ��¡��������� ��� ���� Section Manager ���¤�Ѻ"
    Exit Sub
End If
End Sub

Private Sub ItemDetail101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vRetailCom As Double
Dim vWholeSaleCom As Double
Dim vPrice As Double


On Error Resume Next

If KeyAscii = 13 Then
    vSearch = Trim(ItemDetail101.Text)
    vCheckRequestOnListview = 0
    If vSearch <> "" Then
        vQuery = "execute USP_PM_FindItemFix '" & vSearch & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        
            vPrice = Trim(vRecordset.Fields("saleprice1").Value)
            
            ItemDetail102.Text = Trim(vRecordset.Fields("itemname").Value)
            ItemDetail104.Text = Trim(vRecordset.Fields("unitcode").Value)
            ItemDetail103.Text = Format(vPrice, "##,##0.00")
                        
            If Trim(vRecordset.Fields("comname").Value) <> "" Then
              vRetailCom = Trim(vRecordset.Fields("retailcom").Value)
              vWholeSaleCom = Trim(vRecordset.Fields("wholesalecom").Value)
              Me.LBLRetailCom.Caption = Format(vRetailCom, "##,##0.00")
              Me.LBLWholeSaleCom.Caption = Format(vWholeSaleCom, "##,##0.00")
              
              
              Me.LBLCampaignCom.Caption = Trim(vRecordset.Fields("comname").Value)
              Me.LBLFormComDate.Caption = Trim(vRecordset.Fields("combegindate").Value)
              Me.LBLToComDate.Caption = Trim(vRecordset.Fields("comenddate").Value)
            End If
        Else
        
            ItemDetail102.Text = ""
            ItemDetail104.Text = ""
            ItemDetail103.Text = ""
            Me.LBLRetailCom.Caption = ""
            Me.LBLWholeSaleCom.Caption = ""
            Me.LBLCampaignCom.Caption = ""
            Me.LBLFormComDate.Caption = ""
            Me.LBLToComDate.Caption = ""
            
            MDIForm1.Enabled = False
            FormSearchItem.Text101.Text = vSearch
            FormSearchItem.Show
        End If
        vRecordset.Close
    End If
    If CHK102.Value = 1 Then
        ItemDetail105.Text = (Trim(ItemDetail103.Text) - (Trim(ItemDetail103.Text) * vMemberDiscount) / 100)
        ItemDetail106.SetFocus
    Else
        ItemDetail105.SetFocus
    End If
End If
End Sub

Private Sub ItemDetail101_LostFocus()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vRetailCom As Double
Dim vWholeSaleCom As Double
Dim vPrice As Double


On Error GoTo ErrDescription

    vSearch = Trim(ItemDetail101.Text)
    If vSearch <> "" Then
        vQuery = "execute USP_PM_FindItemFix '" & vSearch & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        
            vPrice = Trim(vRecordset.Fields("saleprice1").Value)
            
            ItemDetail102.Text = Trim(vRecordset.Fields("itemname").Value)
            ItemDetail104.Text = Trim(vRecordset.Fields("unitcode").Value)
            ItemDetail103.Text = Format(vPrice, "##,##0.00")
            
            If Trim(vRecordset.Fields("comname").Value) <> "" Then
              vRetailCom = Trim(vRecordset.Fields("retailcom").Value)
              vWholeSaleCom = Trim(vRecordset.Fields("wholesalecom").Value)
              Me.LBLRetailCom.Caption = Format(vRetailCom, "##,##0.00")
              Me.LBLWholeSaleCom.Caption = Format(vWholeSaleCom, "##,##0.00")
              
              Me.LBLCampaignCom.Caption = Trim(vRecordset.Fields("comname").Value)
              Me.LBLFormComDate.Caption = Trim(vRecordset.Fields("combegindate").Value)
              Me.LBLToComDate.Caption = Trim(vRecordset.Fields("comenddate").Value)
            End If
        Else
        
            ItemDetail102.Text = ""
            ItemDetail104.Text = ""
            ItemDetail103.Text = ""
            Me.LBLRetailCom.Caption = ""
            Me.LBLWholeSaleCom.Caption = ""
            Me.LBLCampaignCom.Caption = ""
            Me.LBLFormComDate.Caption = ""
            Me.LBLToComDate.Caption = ""
            MDIForm1.Enabled = False
            FormSearchItem.Show
            FormSearchItem.Text101.Text = vSearch
        End If
        vRecordset.Close
    End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ItemDetail105_KeyPress(KeyAscii As Integer)
Dim vItemCount As Integer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDisCountLine As Integer
Dim vCheckItemDisCount As String
Dim vUnitCode As String

On Error Resume Next

If KeyAscii = 13 Then
    If CHK103.Value = 1 Then
        vItemCount = Trim(ItemDetail103.Text) - ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
    ElseIf CHK103.Value = 0 Then
        If Trim(ItemDetail105.Text) = "" Then
            vItemCount = Trim(ItemDetail103.Text)
        Else
            vItemCount = Trim(ItemDetail103.Text) - Trim(ItemDetail105.Text)
        End If
    End If
    
    vCheckItemDisCount = Trim(ItemDetail101.Text)
    vUnitCode = Trim(ItemDetail104.Text)
    vQuery = "exec USP_PM_LowerCost  '" & vCheckItemDisCount & "'," & vItemCount & ",'" & vUnitCode & "' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vDisCountLine = Trim(vRecordset.Fields("lowercost"))
    End If
    vRecordset.Close
    
    If vItemCount < 0 Then
    MsgBox "Ŵ�Ҥ��ҡ���� �Ҥһ���"
    ItemDetail105.SetFocus
    Else
    CMD111.SetFocus
    End If
    
    If vDisCountLine = 1 Then
        MsgBox "�Թ��� ���� " & vCheckItemDisCount & " �ʹ��Ҥ�������蹢Ҵ�ع "
    End If
End If
End Sub

Private Sub ItemDetail105_LostFocus()
Dim vItemCount As Integer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDisCountLine As Integer
Dim vCheckItemDisCount As String
Dim vUnitCode As String

On Error Resume Next

    If CHK103.Value = 1 Then
        vItemCount = Trim(ItemDetail103.Text) - ((Trim(ItemDetail103.Text) * Trim(ItemDetail105.Text)) / 100)
    ElseIf CHK103.Value = 0 Then
        If Trim(ItemDetail105.Text) = "" Then
            vItemCount = Trim(ItemDetail103.Text)
        Else
            vItemCount = Trim(ItemDetail103.Text) - Trim(ItemDetail105.Text)
        End If
    End If
    
    vCheckItemDisCount = Trim(ItemDetail101.Text)
    vUnitCode = Trim(ItemDetail104.Text)
    vQuery = "exec USP_PM_LowerCost  '" & vCheckItemDisCount & "'," & vItemCount & " ,'" & vUnitCode & "' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vDisCountLine = Trim(vRecordset.Fields("lowercost"))
    End If
    vRecordset.Close
    
    If vItemCount < 0 Then
    MsgBox "Ŵ�Ҥ��ҡ���� �Ҥһ���"
    ItemDetail105.SetFocus
    Else
    If vDisCountLine = 1 Then
        MsgBox "�Թ��� ���� " & vCheckItemDisCount & " �ʹ��Ҥ�������蹢Ҵ�ع "
    End If
    
    ItemDetail106.SetFocus
    End If

End Sub

Private Sub ItemDetail106_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    CMD105.SetFocus
End If
End Sub

Private Sub ItemDetail107_Change()
Dim vPointCheck As Integer
Dim vItemType As String
Dim vItemType1 As String
Dim vItemType2 As String

If Me.ItemDetail107.Text <> "" Then
    vItemType2 = Me.ItemDetail107.Text
    vPointCheck = InStr(1, vItemType2, "/")
    
    vItemType1 = Left(vItemType2, vPointCheck - 1)
    
    If vItemType1 = "11" Then
    Me.LBLHotPrice.Visible = True
    Me.CMBHotPrice.Visible = True
    Me.CMBHotPrice.ListIndex = 1
    Me.CMBHotPrice.Enabled = False
    Else
    Me.LBLHotPrice.Visible = False
    Me.CMBHotPrice.Visible = False
    End If
Else
    Me.LBLHotPrice.Visible = False
    Me.CMBHotPrice.Visible = False
End If

End Sub

Private Sub ListView101_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo ErrDescription

ListView101.Sorted = True
ListView101.SortKey = ColumnHeader.Index - 1
If vSortResult = 0 Then
    ListView101.SortOrder = lvwAscending
    vSortResult = 1
Else
    ListView101.SortOrder = lvwDescending
    vSortResult = 0
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListView101_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
i = Item.Index
'"&H80000008&"
    If ListView101.ListItems.Item(Item.Index).Checked = False Then
        ListView101.ListItems(i).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(10).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(11).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(12).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).ListSubItems(13).ForeColor = "&H000000FF"
        ListView101.ListItems.Item(i).SubItems(11) = "1"
    Else
        ListView101.ListItems(i).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(10).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(11).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(12).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).ListSubItems(13).ForeColor = "&H80000008"
        ListView101.ListItems.Item(i).SubItems(11) = "0"
    End If

End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckIsConfirmDoc  As Integer
Dim vCheckIsCancelDoc  As Integer
Dim vDocnoCheck As String

On Error Resume Next

If ListView101.ListItems.Item(Item.Index).Checked = True Then
If ListView101.ListItems.Count <> 0 And Trim(Text101.Text) <> "" Then
vDocnoCheck = Trim(Text101.Text)
vQuery = "select isconfirm ,iscancel from npmaster.dbo.tb_pm_request where docno = '" & vDocnoCheck & "' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckIsConfirmDoc = Trim(vRecordset.Fields("isconfirm"))
    vCheckIsCancelDoc = Trim(vRecordset.Fields("iscancel"))
End If
vRecordset.Close

If vCheckIsConfirmDoc <> 2 Or vCheckIsCancelDoc = 1 Then
    vCheckRequestOnListview = 1
    CMD108.Enabled = True
    ItemDetail101.Enabled = False
    vIndexItemUpdate = Item.Index
    If Trim(ListView101.ListItems.Item(Item.Index).SubItems(9)) <> 0 Then
        CHK102.Value = 1
    Else
        CHK102.Value = 0
    End If
    
    If Trim(ListView101.ListItems.Item(Item.Index).SubItems(10)) <> 0 Then
        Check101.Value = 1
    Else
        Check101.Value = 0
    End If
    
    ItemDetail101.Text = Trim(ListView101.ListItems.Item(Item.Index).Text)
    ItemDetail102.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(1))
    ItemDetail103.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(2))
    ItemDetail104.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(6))
    If Trim(ListView101.ListItems.Item(Item.Index).SubItems(8)) = 1 Then
        CHK103.Value = 1
        ItemDetail105.Text = Left(Trim(ListView101.ListItems.Item(Item.Index).SubItems(5)), Len(Trim(ListView101.ListItems.Item(Item.Index).SubItems(5))) - 1)
    ElseIf Trim(ListView101.ListItems.Item(Item.Index).SubItems(8)) = 2 Then
        ItemDetail105.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(3))
        ItemDetail105.Enabled = False
    Else
        ItemDetail105.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(5))
        CHK103.Value = 0
    End If
    ItemDetail106.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(7))
    ItemDetail107.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(12))
End If
End If
Else
    ItemDetail101.Text = ""
    ItemDetail102.Text = ""
    ItemDetail103.Text = ""
    ItemDetail104.Text = ""
    ItemDetail105.Text = ""
    ItemDetail106.Text = ""
    ItemDetail107.Text = ""
    CHK102.Value = 0
    CHK103.Value = 0
    Check101.Value = 0
End If
End Sub

Private Sub Text4_Change()

End Sub
