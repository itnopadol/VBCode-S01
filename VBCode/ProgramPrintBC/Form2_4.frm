VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form2_4 
   Caption         =   "�����㺵�Ǩ�Ѻ�Թ�����Ҥ�ѧ"
   ClientHeight    =   8115
   ClientLeft      =   5520
   ClientTop       =   2175
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form2_4.frx":0000
   ScaleHeight     =   8115
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.OptionButton OPTConcrete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "㺵�Ǩ�Ѻ��ѧ PASS-�͹��յ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3555
      TabIndex        =   5
      Top             =   3240
      Width           =   3840
   End
   Begin VB.OptionButton OPTPass 
      BackColor       =   &H00FFFFFF&
      Caption         =   "㺵�Ǩ�Ѻ��ѧ PASS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3555
      TabIndex        =   4
      Top             =   2655
      Width           =   3840
   End
   Begin VB.OptionButton OPTNormal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "㺵�Ǩ�Ѻ�����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3555
      TabIndex        =   3
      Top             =   2070
      Value           =   -1  'True
      Width           =   3840
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   675
      Top             =   6615
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
   Begin VB.CommandButton CMDPrint 
      Caption         =   "������͡���"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   5355
      TabIndex        =   2
      Top             =   3825
      Width           =   2040
   End
   Begin VB.TextBox TBPurchaseNo 
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
      Left            =   3555
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "���͡�������͹����� :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   900
      TabIndex        =   6
      Top             =   2115
      Width           =   2445
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�Ţ����͡������觫��� :"
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
      Left            =   1575
      TabIndex        =   0
      Top             =   1440
      Width           =   1950
   End
End
Attribute VB_Name = "Form2_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer

On Error Resume Next

    vDocNo = UCase(Trim(Me.TBPurchaseNo.Text))
    Me.TBPurchaseNo.Text = vDocNo
    
    vQuery = "select docno,Lastprinteduser,lastprintdatetime,doctypeid,groupdoc,printed from npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        vPrint = Trim(vRecordset.Fields("printed").Value)
    Else
        MsgBox "����բ������Ţ����͡��� ��سҵ�Ǩ�ͺ", vbCritical, "Send Error Message"
        Exit Sub
    End If
    vRecordset.Close
    
    
    vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' and typenumber = 6"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 '�¾��������
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
    Else
            vCheckDocNo = 0 '�ѧ���������
    End If
    vRecordset.Close
        
        If vCheckDocNo = 0 Then '����Ҿ��������á���� ��᷹
            vHeaderType = 22
            vNamePrint = Trim(vUserID)
            vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                vHeader = Trim(vRecordset.Fields("header").Value)
                vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
            End If
            vRecordset.Close
            vDocuments = vDocNumber & vHeader & "-" & vPicking
            
            vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
            gConnection.Execute vQuery
            
            vTypeNumber = 6
            vQuery = "exec dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vTypeNumber = 6
            vQuery = "exec dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
            vQuery = "exec dbo.USP_NP_InsertLogPrintRunningRes '" & vDocNo & "','" & vUserID & "' ," & vTypeNumber & " "
            gConnection.Execute vQuery
            
    End If '����Ҿ��������á���� ��᷹
        
        '---------------------------------------------------------------------------------------------
        
        Dim vCountPass As Integer
            
        vQuery = "exec dbo.USP_PO_CheckItemPass  '" & vDocNo & "'"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCountPass = Trim(vRecordset.Fields("countpass").Value)
        End If
        vRecordset.Close
        
         vRepType = "PO"
         
        If vCountPass = 0 Then
            vRepID = 415
        ElseIf vCountPass > 0 And Me.OPTConcrete.Value = True Then
            vRepID = 584
        ElseIf vCountPass > 0 And Me.OPTPass.Value = True Then
            vRepID = 583
        End If
        
        If vRepID = 0 Then
            MsgBox "���͡�����㺵�Ǩ�Ѻ�Դ������ ��سҵ�Ǩ�ͺ", vbCritical, "Send Error Message"
            Exit Sub
        End If
        
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            With Crystal101
                .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .Formulas(0) = "computername='" & vComputerName1 & "' "
                .Formulas(1) = "username='" & vUserName1 & "' "
                .Action = 1
            End With
        End If
        vRecordset.Close
        
        Me.TBPurchaseNo.Text = ""
        Me.TBPurchaseNo.SetFocus

End Sub


Private Sub TBPurchaseNo_LostFocus()
Dim vDocNo As String

If Me.TBPurchaseNo.Text <> "" Then
    vDocNo = UCase(Trim(Me.TBPurchaseNo.Text))
    Me.TBPurchaseNo.Text = vDocNo
 End If
End Sub