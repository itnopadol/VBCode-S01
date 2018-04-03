VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FormPrintReserveInvoice 
   Caption         =   "พิมพ์ทดแทนเอกสารด้านขาย"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormPrintReserveInvoice.frx":0000
   ScaleHeight     =   10440
   ScaleMode       =   0  'User
   ScaleWidth      =   48797.31
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR106 
      Left            =   3465
      Top             =   6210
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
   Begin VB.OptionButton Option1 
      Caption         =   "ไม่แสดงส่วนลดรายตัว"
      Height          =   330
      Left            =   6615
      TabIndex        =   6
      Top             =   5220
      Value           =   -1  'True
      Width           =   4695
   End
   Begin Crystal.CrystalReport CR105 
      Left            =   2610
      Top             =   6255
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
   Begin Crystal.CrystalReport Crystal102 
      Left            =   2160
      Top             =   6255
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
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1710
      Top             =   6255
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
   Begin Crystal.CrystalReport CrystalReport022 
      Left            =   1260
      Top             =   6255
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
   Begin Crystal.CrystalReport CrystalReport021 
      Left            =   810
      Top             =   6255
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
   Begin VB.CommandButton CMD021 
      Caption         =   "พิมพ์เอกสาร"
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
      Left            =   6615
      TabIndex        =   5
      Top             =   5895
      Width           =   1770
   End
   Begin VB.TextBox TXT022 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2205
      TabIndex        =   4
      Top             =   2340
      Width           =   4020
   End
   Begin VB.TextBox TXT021 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2205
      TabIndex        =   3
      Top             =   1530
      Width           =   2805
   End
   Begin MSComctlLib.ListView ListView021 
      Height          =   3435
      Left            =   6615
      TabIndex        =   2
      Top             =   1530
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ฟอร์มที่จะพิมพ์"
         Object.Width           =   7584
      EndProperty
   End
   Begin VB.Label LBLDocNo 
      Height          =   330
      Left            =   2205
      TabIndex        =   10
      Top             =   1170
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการชื่อเอกสารที่จะพิมพ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   6615
      TabIndex        =   9
      Top             =   1080
      Width           =   3300
   End
   Begin VB.Label LBL024 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2205
      TabIndex        =   8
      Top             =   4095
      Width           =   4020
   End
   Begin VB.Label LBL023 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2205
      TabIndex        =   7
      Top             =   3240
      Width           =   4020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์ :"
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
      Left            =   810
      TabIndex        =   1
      Top             =   2430
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
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
      Left            =   855
      TabIndex        =   0
      Top             =   1575
      Width           =   1275
   End
End
Attribute VB_Name = "FormPrintReserveInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD021_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String, vDocTypeID As String, vGroupDoc As String
Dim i As Integer
Dim vWHCode(10) As String
Dim vTableName As String
Dim vIsCompleteSave As Integer
Dim vCheckPOS As String

On Error GoTo ErrDescription

If TXT021.Text <> "" And TXT022.Text <> "" Then
    vDocNo = Trim(TXT021.Text)
    
    vCheckPOS = Left(vDocNo, 7)
    
    If UCase(vCheckPOS) = "S02-PXT" Then
    vQuery = "exec dbo.usp_np_SearchTaxNoPrintInvoice_PXT '" & vDocNo & "' "
    Else
    vQuery = "select * from npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
    End If
    
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
    vDocTypeID = Trim(vRecordset.Fields("doctypeid").Value)
    End If
    vRecordset.Close
    
  If Right(Left(vDocNo, 3), 1) = "2" Then
  
    If vDocTypeID = "INV" Then
        vDocTypeID = "S02-INV"
    End If
    
    If vDocTypeID = "RT" Then
        vDocTypeID = "S02-RT"
    End If
    
    If vDocTypeID = "DB" Then
        vDocTypeID = "S02-DB"
    End If
  
    If vDocTypeID = "DP" Then
        vDocTypeID = "S02-DP"
    End If
  End If
    

  If vDocTypeID = "S02-INV" Then

            vTableName = "BCARInvoice"
            
            vQuery = "exec dbo.USP_NP_SearchIsCompleteSave '" & vTableName & "' ,'" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vIsCompleteSave = vRecordset.Fields("iscompletesave").Value
            End If
            vRecordset.Close
            
            If vIsCompleteSave = 0 Then
            MsgBox "เอกสารยังบันทึกไม่สมบูรณ์ ยังไม่สามารถพิมพ์ได้ กรุณารอสักครู่", vbCritical, "Send Error Message"
            Me.CMD021.SetFocus
            Exit Sub
            End If
        
        If TXT022.Text = "พิมพ์ใบเสร็จของสาขา" Or TXT022.Text = "พิมพ์ใบเสร็จ" Then
                Call PrintInvoice
            End If
    ElseIf vDocTypeID = "S02-DP" Then
    
            If TXT022.Text = "พิมพ์ใบมัดจำของสาขา" Or TXT022.Text = "พิมพ์ใบมัดจำ" Then
                Call PrintDeposit
            End If

    ElseIf vDocTypeID = "S02-DB" Then

            vTableName = "BCDebitNote1"
            
            vQuery = "exec dbo.USP_NP_SearchIsCompleteSave '" & vTableName & "' ,'" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vIsCompleteSave = vRecordset.Fields("iscompletesave").Value
            End If
            vRecordset.Close
            
            If vIsCompleteSave = 0 Then
            MsgBox "เอกสารยังบันทึกไม่สมบูรณ์ ยังไม่สามารถพิมพ์ได้ กรุณารอสักครู่", vbCritical, "Send Error Message"
            Me.CMD021.SetFocus
            Exit Sub
            End If

            If TXT022.Text = "พิมพ์ใบเพิ่มหนี้ของสาขา" Or TXT022.Text = "พิมพ์ใบเพิ่มหนี้" Then
                Call PrintDebit
            End If
    ElseIf vDocTypeID = "S02-PXT" Then

            If TXT022.Text = "พิมพ์ใบกำกับภาษีของสาขา" Then
                Call PrintPXT
            End If
    ElseIf vDocTypeID = "S02-RT" Then
    
            vTableName = "BCCreditNote"
            
            vQuery = "exec dbo.USP_NP_SearchIsCompleteSave '" & vTableName & "' ,'" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vIsCompleteSave = vRecordset.Fields("iscompletesave").Value
            End If
            vRecordset.Close
            
            If vIsCompleteSave = 0 Then
            MsgBox "เอกสารยังบันทึกไม่สมบูรณ์ ยังไม่สามารถพิมพ์ได้ กรุณารอสักครู่", vbCritical, "Send Error Message"
            Me.CMD021.SetFocus
            Exit Sub
            End If

            If TXT022.Text = "พิมพ์ใบลดหนี้ของสาขา" Or TXT022.Text = "พิมพ์ใบลดหนี้" Then
                Call PrintReturn
            End If
 Else
 MsgBox "ไม่มีเลขที่เอกสารที่จะพิมพ์", vbInformation + vbCritical, "ข้อความเตือน"
End If
Else
MsgBox "คุณใส่ข้อมูลการพิมพ์ไม่ครบ", vbCritical, "ข้อความเตือน"
End If
TXT021.Text = ""
TXT022.Text = ""
Me.LBLDocNo.Caption = ""
ListView021.ListItems.Clear


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub Form_Load()
On Error Resume Next

'Me.TXT021.SetFocus
End Sub

Private Sub ListView021_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT022.Text = Item
End Sub

Private Sub TXT021_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String, vDocTypeID As String, vGroupDoc As String
Dim vCreatorcode As String
Dim vCreatedatetime As Date
Dim FormListItems As ListItem

Dim vCheckPOS As String

On Error GoTo ErrDescription

ListView021.ListItems.Clear
If KeyAscii = 13 Then
    vDocNo = Trim(TXT021.Text)
    
    vCheckPOS = Left(vDocNo, 7)
    
    If UCase(vCheckPOS) = "S02-PXT" Then
    vQuery = "exec dbo.usp_np_SearchTaxNoPrintInvoice_PXT '" & vDocNo & "' "
    Else
    vQuery = "select docno,Lastprinteduser,lastprintdatetime,doctypeid,groupdoc from npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
    End If
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCreatorcode = Trim(vRecordset.Fields("Lastprinteduser").Value)
        vCreatedatetime = Trim(vRecordset.Fields("lastprintdatetime").Value)
        vDocTypeID = Trim(vRecordset.Fields("doctypeid").Value)
        vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        Me.LBLDocNo.Caption = Trim(vRecordset.Fields("docno").Value)
    Else
        Call CheckDocument
        Exit Sub
    End If
    vRecordset.Close
    
    Option1.Visible = True

    LBL023.Caption = "ผู้ที่ทำเอกสาร คือ   " & vCreatorcode
    LBL024.Caption = "วันที่ทำเอกสาร คือ  " & vCreatedatetime

    vQuery = "select name from npmaster.dbo.npform where moduleid = '" & vDocTypeID & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set FormListItems = ListView021.ListItems.Add(, , Trim(vRecordset.Fields("name").Value))
            vRecordset.MoveNext
        Wend
    vRecordset.Close
  
    End If

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Function CheckDocument()
Dim vRecordset  As New ADODB.Recordset
Dim vQuery As String
Dim vDocGroup As String, vDocument As String
Dim vTable As String, vCheckDoc As String, vTypeDoc As String, vDocGroup1 As String
Dim vMemDocNo As String

On Error GoTo ErrDescription

vMemDocNo = TXT021.Text
vDocGroup1 = UCase(Left(Right(vMemDocNo, Len(vMemDocNo) - InStr(vMemDocNo, "-")), 3))
vDocument = Trim(TXT021.Text)
vQuery = "select upper('" & vDocGroup1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup = Trim(vRecordset.Fields("vDocGroup").Value)
End If
vRecordset.Close
'--------------------------------------------------------------------------------------------

If vDocGroup = "IHV" Or vDocGroup = "IHN" Or vDocGroup = "ICV" Or vDocGroup = "ICN" Or vDocGroup = "IVD" Or _
    vDocGroup = "IVN" Or vDocGroup = "IAB" Or vDocGroup = "IVM" Then
    vTable = "BCNP.DBO.BCARINVOICE"
    vTypeDoc = "S02-INV"
End If
'----------------------------------

If vDocGroup = "IDV" Or vDocGroup = "IDN" Then
    vTable = "BCNP.DBO.BCARDEPOSIT"
    vTypeDoc = "S02-DP"
End If

If vDocGroup = "DCV" Or vDocGroup = "DCN" Or vDocGroup = "DHV" Or vDocGroup = "DHN" Then
    vTable = "BCNP.DBO.BCDEBITNOTE1"
    vTypeDoc = "S02-DB"
End If

If vDocGroup = "RCV" Or vDocGroup = "RCN" Or vDocGroup = "RDV" Or vDocGroup = "RDN" Or vDocGroup = "RXT" Or vDocGroup = "RHV" _
Or vDocGroup = "RHN" Or vDocGroup = "RVD" Or vDocGroup = "RVN" Or vDocGroup = "RAB" And vDocGroup <> "RVM" Then
    vTable = "BCNP.DBO.BCCREDITNOTE"
    vTypeDoc = "S02-RT"
End If


If vTable <> "" Then
vQuery = "select docno from " & vTable & " where docno = '" & vDocument & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDoc = vRecordset.Fields("Docno").Value
End If
vRecordset.Close
'-------------------------------------------------------------------------------------------------

End If
If vCheckDoc = "" Or IsNull(vCheckDoc) Then
    MsgBox "ไม่มีเอกสารนี้ในระบบ", vbCritical, "ข้อความเตือน"
Else
    vQuery = "INSERT INTO NPMaster.dbo.NPPrintLogs(DocNo,DocTypeID,GroupDOC,Code,Printed,LastPrintedUser,LastPrintDateTime)" & _
                     "select docno,'" & vTypeDoc & "' as DocTypeID,'" & vDocGroup & "' as GroupDoc,arcode,1 as Printed,'" & vUserID & "' as UserID ,getdate() From " & vTable & " where   docno = '" & vDocument & "'  "
    gConnection.Execute vQuery
    MsgBox "กรุณากดปุ่ม Enter อีกครั้งนะครับ", vbCritical, "ข้อความแจ้งให้ทราบ"
End If

vQuery = "Delete npmaster.dbo.npprintserver where docno = '" & vDocument & "' "
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If

End Function

Public Sub PrintDeposit()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT021.Text)
        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "S02-DP"
                            If UCase(vGroupDoc) = "IDV" Then
                                vRepID = 549
                            ElseIf UCase(vGroupDoc) = "IDM" Then
                                vRepID = 549
                            ElseIf UCase(vGroupDoc) = "IDN" Then
                                vRepID = 550
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport021
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true'"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------

                                                     
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintPXT()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String

Dim vRepID As Integer
Dim vRepType As String
Dim vCredit As Integer
Dim vDueDate As String
Dim vCheck As Integer


vDocNo = Me.LBLDocNo.Caption  'Trim(TXT021.Text)
vCredit = 0
        
vRepType = "S02-PXT"
vRepID = 562

 vDueDate = ""
 vCheck = 1
 vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     With Me.CR105
         .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
         .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
         .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
         .Formulas(0) = "CreditCondition='" & vDueDate & "' "
         .Destination = crptToWindow
         .WindowState = crptMaximized
         .Action = 1
     End With
 End If
 vRecordset.Close

End Sub

Public Sub PrintInvoice()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vDocNo As String
        Dim vGroupDoc As String
        Dim vReportName As String
        Dim vName As String
        Dim vPrint As Integer
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCheck As Integer
        Dim vDueDate As String
        Dim vCredit As Integer

        vDocNo = Trim(TXT021.Text)
        vCredit = 0

        On Error GoTo PrintError

        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        
        vRepType = "S02-INV"
        
            If vGroupDoc = "IHV" Or vGroupDoc = "ihv" Then
                If Option1.Value = False Then
                    vRepID = 537
                Else
                    vRepID = 538
                End If
            ElseIf vGroupDoc = "IHN" Or vGroupDoc = "ihn" Then
                If Option1.Value = False Then
                    vRepID = 551
                Else
                    vRepID = 552
                End If
            ElseIf vGroupDoc = "ICV" Or vGroupDoc = "icv" Then
            
                vCredit = 1
                 If Option1.Value = False Then
                     vRepID = 539
                 ElseIf Option1.Value = True Then
                     vRepID = 540
                 End If
            
            ElseIf vGroupDoc = "ICN" Or vGroupDoc = "icn" Then
                vCredit = 1
                 If Option1.Value = False Then
                     vRepID = 553
                 ElseIf Option1.Value = True Then
                     vRepID = 554
                 End If
                 
            ElseIf vGroupDoc = "IAB" Or vGroupDoc = "iab" Then
                If Option1.Value = False Then
                    vRepID = 541
                Else
                    vRepID = 542
                End If
            Else
            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
            Exit Sub
            End If
            
            If vCredit = 1 Then
              vQuery = "exec USP_AR_SearchDueDateInvoice '" & vDocNo & "' "
              If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vDueDate = Trim(vRecordset.Fields("conditpaycode").Value)
              End If
              vRecordset.Close
            Else
              vDueDate = ""
            End If
            
            vCheck = 0
            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                With Crystal101
                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                    .Formulas(0) = "CreditCondition='" & vDueDate & "' "
                    .Destination = crptToWindow
                    .WindowState = crptMaximized
                    .Action = 1
                End With
            End If
            vRecordset.Close
            
            Option1.Value = False
            Option1.Visible = False
            Label3.Visible = False
            
            '---------- ----------- ---------- ------------ ----------- ---------- ------------ ---------- --------- ----------

PrintError:
                If Err.Description <> "" Then
                MsgBox Err.Description
                End If
End Sub


Public Sub PrintItem()
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

vDocNo = Trim(TXT021.Text)
vRepID = 56
vRepType = "INV"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport021
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintReturn()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT021.Text)
        
        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "S02-RT"
                            If UCase(vGroupDoc) = "RHV" Or UCase(vGroupDoc) = "RXT" Or UCase(vGroupDoc) = "RAB" Or UCase(vGroupDoc) = "RDM" Then
                                vRepID = 543
                            ElseIf UCase(vGroupDoc) = "RHN" Or UCase(vGroupDoc) = "RDN" Then
                                vRepID = 545
                            ElseIf UCase(vGroupDoc) = "RCV" Or UCase(vGroupDoc) = "RVD" Or UCase(vGroupDoc) = "RVM" Or UCase(vGroupDoc) = "RDV" Then
                                vRepID = 544
                            ElseIf UCase(vGroupDoc) = "RCN" Or UCase(vGroupDoc) = "RVN" Then
                                vRepID = 546
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport021
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                 
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintDebit()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT021.Text)

        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "S02-DB"
                            If vGroupDoc = "DHV" Or vGroupDoc = "dhv" Then
                                vRepID = 547
                            ElseIf vGroupDoc = "DHN" Or vGroupDoc = "dhn" Then
                                vRepID = 548
                            ElseIf vGroupDoc = "DCV" Or vGroupDoc = "dcv" Then
                                vRepID = 547
                            ElseIf vGroupDoc = "DCN" Or vGroupDoc = "dcn" Then
                                vRepID = 548
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport021
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
               
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub


Public Sub PrintCashReceipt()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vChqStatus As String

On Error GoTo ErrDescription

        vChqStatus = 0

        vDocNo = Trim(TXT021.Text)

        vQuery = "select groupdoc from  npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "RE"
                            If vGroupDoc = "RE" Or vGroupDoc = "re" Then
                                vRepID = 38
                            ElseIf vGroupDoc = "RN" Or vGroupDoc = "rn" Then
                                vRepID = 79
                            ElseIf vGroupDoc = "RD" Or vGroupDoc = "rd" Then
                                vRepID = 38
                            ElseIf vGroupDoc = "RC" Or vGroupDoc = "rc" Then
                                vRepID = 79
                            ElseIf vGroupDoc = "RQ" Or vGroupDoc = "rq" Then
                                vRepID = 79
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With Crystal102
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "ChqStatus;" & vChqStatus & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
                            Label2.Visible = False
                            '---------------------------------------------------------------------------------------------------
                            
                             If vPrint = 0 Then
                                            vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
                                            gConnection.Execute vQuery
                                    End If
                            
                            
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintPayBill()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT021.Text)

        vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "BI"
                            If vGroupDoc = "BI" Or vGroupDoc = "bi" Then
                                vRepID = 46
                            ElseIf vGroupDoc = "BN" Or vGroupDoc = "bn" Then
                                vRepID = 46
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport021
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                       
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub




