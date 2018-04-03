VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form02 
   Caption         =   "หน้าพิมพ์เอกสารสำคัญ"
   ClientHeight    =   8370
   ClientLeft      =   2655
   ClientTop       =   645
   ClientWidth     =   12000
   Icon            =   "Form02.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form02.frx":08CA
   ScaleHeight     =   8370
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal103 
      Left            =   8160
      Top             =   5880
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
   Begin VB.CheckBox CHKUnShowItemAmount 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ไม่แสดง มูลค่าสินค้า"
      Height          =   330
      Left            =   1800
      TabIndex        =   13
      Top             =   4905
      Width           =   2580
   End
   Begin Crystal.CrystalReport Crystal102 
      Left            =   4275
      Top             =   6930
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
      Left            =   6030
      Top             =   6525
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
      Left            =   2565
      Top             =   7380
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
      BackColor       =   &H80000009&
      Caption         =   "ไม่แสดงส่วนลดรายตัว"
      Height          =   390
      Left            =   1800
      TabIndex        =   10
      Top             =   5325
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.OptionButton Opt021 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "โชว์เช็ค"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1800
      TabIndex        =   9
      Top             =   5850
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.CommandButton CMD021 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      TabIndex        =   3
      Top             =   2880
      Width           =   1200
   End
   Begin MSComctlLib.ListView ListView021 
      Height          =   3330
      Left            =   5850
      TabIndex        =   1
      Top             =   1500
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5874
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ฟอร์มที่พิมพ์"
         Object.Width           =   7267
      EndProperty
   End
   Begin VB.TextBox TXT022 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1875
      TabIndex        =   2
      Top             =   2250
      Width           =   2565
   End
   Begin VB.TextBox TXT021 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1875
      TabIndex        =   0
      Top             =   1500
      Width           =   2565
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "เกี่ยวกับบิลขาย"
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
      Left            =   525
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "เกี่ยวกับรับชำระ"
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
      Left            =   525
      TabIndex        =   11
      Top             =   5925
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์เอกสารสำคัญทดแทน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   540
      Left            =   2550
      TabIndex        =   8
      Top             =   375
      Width           =   7440
   End
   Begin VB.Label LBL024 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   825
      TabIndex        =   7
      Top             =   5025
      Width           =   4890
   End
   Begin VB.Label LBL023 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   825
      TabIndex        =   6
      Top             =   4500
      Width           =   4890
   End
   Begin VB.Label LBL022 
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์"
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
      Height          =   390
      Left            =   675
      TabIndex        =   5
      Top             =   2250
      Width           =   1065
   End
   Begin VB.Label LBL021 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
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
      Height          =   390
      Left            =   675
      TabIndex        =   4
      Top             =   1500
      Width           =   1065
   End
End
Attribute VB_Name = "Form02"
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

On Error GoTo ErrDescription

If TXT021.Text <> "" And TXT022.Text <> "" Then
    vDocNo = Trim(TXT021.Text)
    
    vQuery = "select * from npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
    vDocTypeID = Trim(vRecordset.Fields("doctypeid").Value)
    End If
    vRecordset.Close

  If vDocTypeID = "INV" Then

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
        
            If TXT022.Text = "พิมพ์ใบเสร็จ+ใบจ่ายสินค้า" Then
                Call PrintInvoice
            ElseIf TXT022.Text = "พิมพ์ใบเสร็จ" Then
                Call PrintInvoice
            ElseIf TXT022.Text = "พิมพ์ใบเสร็จ ไดรฟ์ทรู" Then
                Call PrintInvoice_DriveThru
            ElseIf Trim(TXT022.Text) = Trim("พิมพ์ใบเสร็จ เฉพาะกิจ") Then
               Call PrintInvoice_NoVat
            End If
    ElseIf vDocTypeID = "DP" Then
    
            If TXT022.Text = "พิมพ์ใบมัดจำ" Then
                Call PrintDeposit
            End If
    ElseIf vDocTypeID = "BI" Then
    
            vTableName = "BCPayBill"
            
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

            If TXT022.Text = "พิมพ์ใบวางบิล" Then
                Call PrintPayBill
            End If
    ElseIf vDocTypeID = "RE" Then

            vTableName = "BCReceipt1"
            
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

            If TXT022.Text = "พิมพ์ใบเสร็จรับชำระ" Then
                Call PrintCashReceipt
            End If
    ElseIf vDocTypeID = "DB" Then

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

            If TXT022.Text = "พิมพ์ใบเพิ่มหนี้" Then
                Call PrintDebit
            End If
    ElseIf vDocTypeID = "RT" Then
    
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

            If TXT022.Text = "พิมพ์ใบลดหนี้" Then
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
LBL023.Caption = ""
LBL024.Caption = ""
ListView021.ListItems.Clear
Opt021.Value = False
Opt021.Visible = False

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

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

On Error GoTo ErrDescription

ListView021.ListItems.Clear
If KeyAscii = 13 Then
    vDocNo = Trim(TXT021.Text)
    vQuery = "select docno,Lastprinteduser,lastprintdatetime,doctypeid,groupdoc from npmaster.dbo.vw_pc_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCreatorcode = Trim(vRecordset.Fields("Lastprinteduser").Value)
        vCreatedatetime = Trim(vRecordset.Fields("lastprintdatetime").Value)
        vDocTypeID = Trim(vRecordset.Fields("doctypeid").Value)
        vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
    Else
        Call CheckDocument
        Exit Sub
    End If
    vRecordset.Close
    
    If vDocTypeID = "RE" Then
        Opt021.Visible = True
        Label2.Visible = True
        Option1.Visible = False
        Label3.Visible = False
    ElseIf vDocTypeID = "INV" Then
        Option1.Visible = True
        Label3.Visible = True
        Opt021.Visible = False
        Label2.Visible = False
    End If
    
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
vDocGroup1 = UCase(Left(Right(vMemDocNo, Len(vMemDocNo) - InStr(vMemDocNo, "-")), 3)) 'Left(Trim(TXT021.Text), 3)
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
    vTypeDoc = "INV"
End If
'----------------------------------
If vDocGroup <> "RCV" And vDocGroup <> "RCN" And vDocGroup <> "RDV" And vDocGroup <> "RDN" And vDocGroup <> "RVM" Then
        If UCase(Left(Right(vDocGroup, Len(vDocGroup) - InStr(vDocGroup, "-")), 2)) = "RE" Or UCase(Left(Right(vDocGroup, Len(vDocGroup) - InStr(vDocGroup, "-")), 2)) = "RN" Or UCase(Left(Right(vDocGroup, Len(vDocGroup) - InStr(vDocGroup, "-")), 2)) = "RD" Or _
        UCase(Left(Right(vDocGroup, Len(vDocGroup) - InStr(vDocGroup, "-")), 2)) = "RC" Or UCase(Left(Right(vDocGroup, Len(vDocGroup) - InStr(vDocGroup, "-")), 2)) = "RQ" Then
            vTable = "BCNP.DBO.BCRECEIPT1"
            vTypeDoc = "RE"
            vDocGroup = UCase(Left(Right(vDocGroup, Len(vDocGroup) - InStr(vDocGroup, "-")), 2)) 'Left(vDocGroup, 2)
        End If
Else
            vTable = "BCNP.DBO.BCCREDITNOTE"
            vTypeDoc = "RT"
End If
'----------------------------------------------
If vDocGroup = "IDV" Or vDocGroup = "IDN" Then
    vTable = "BCNP.DBO.BCARDEPOSIT"
    vTypeDoc = "DP"
End If

If vDocGroup = "DCV" Or vDocGroup = "DCN" Or vDocGroup = "DHV" Or vDocGroup = "DHN" Then
    vTable = "BCNP.DBO.BCDEBITNOTE1"
    vTypeDoc = "DB"
End If

If vDocGroup = "RCV" Or vDocGroup = "RCN" Or vDocGroup = "RDV" Or vDocGroup = "RDN" Or vDocGroup = "RHV" _
Or vDocGroup = "RHN" Or vDocGroup = "RVD" Or vDocGroup = "RVN" Or vDocGroup = "RAB" And vDocGroup <> "RVM" Then
    vTable = "BCNP.DBO.BCCREDITNOTE"
    vTypeDoc = "RT"
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
        
                            vRepType = "DP"
                            If UCase(vGroupDoc) = "IDV" Then
                                vRepID = 340
                            ElseIf UCase(vGroupDoc) = "IDM" Then
                                vRepID = 66
                            ElseIf UCase(vGroupDoc) = "IDN" Then
                                vRepID = 67
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
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
        vRepType = "INV"
            If vGroupDoc = "IHV" Or vGroupDoc = "ihv" Then
                If Option1.Value = False Then
                    vRepID = 50
                Else
                    vRepID = 150
                End If
            ElseIf vGroupDoc = "IHN" Or vGroupDoc = "ihn" Then
                If Option1.Value = False Then
                    vRepID = 51
                Else
                    vRepID = 151
                End If
            ElseIf vGroupDoc = "ICV" Or vGroupDoc = "icv" Then
            
                vCredit = 1
                 If Option1.Value = False And Me.CHKUnShowItemAmount.Value = 0 Then
                     vRepID = 52
                 ElseIf Option1.Value = True And Me.CHKUnShowItemAmount.Value = 0 Then
                     vRepID = 152
                 ElseIf Me.CHKUnShowItemAmount.Value = 1 Then
                     vRepID = 507
                 End If
            
            ElseIf vGroupDoc = "ICN" Or vGroupDoc = "icn" Then
            
                vCredit = 1
                 If Option1.Value = False And Me.CHKUnShowItemAmount.Value = 0 Then
                     vRepID = 53
                 ElseIf Option1.Value = True And Me.CHKUnShowItemAmount.Value = 0 Then
                     vRepID = 153
                 ElseIf Me.CHKUnShowItemAmount.Value = 1 Then
                     vRepID = 507
                 End If
                 
            ElseIf vGroupDoc = "IVD" Or vGroupDoc = "ivd" Then
            MsgBox "ห้ามออกหัวเอกสารเก็บเงินปลายทาง", vbInformation + vbCritical, "ข้อความเตือน"
                'vCredit = 1
                'If Option1.Value = False Then
                 '   vRepID = 52
                'Else
                 '   vRepID = 152
                'End If
            ElseIf vGroupDoc = "IVM" Or vGroupDoc = "ivm" Then
            MsgBox "ห้ามออกหัวเอกสารเก็บเงินปลายทาง", vbInformation + vbCritical, "ข้อความเตือน"
                'vCredit = 1
                'If Option1.Value = False Then
                 '   vRepID = 52
                'Else
                 '   vRepID = 152
                'End If
            ElseIf vGroupDoc = "IVN" Or vGroupDoc = "ivn" Then
            MsgBox "ห้ามออกหัวเอกสารเก็บเงินปลายทาง", vbInformation + vbCritical, "ข้อความเตือน"
                'vCredit = 1
                'If Option1.Value = False Then
                 '   vRepID = 53
                'Else
                 '   vRepID = 153
                'End If
            ElseIf vGroupDoc = "IAB" Or vGroupDoc = "iab" Then
                If Option1.Value = False Then
                    vRepID = 54
                Else
                    vRepID = 154
                End If
            ElseIf vGroupDoc = "IVE" Or vGroupDoc = "ive" Then
                vRepID = 55
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
            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
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


Public Sub PrintInvoice_NoVat()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vDocNo As String
        Dim vGroupDoc As String
        Dim vReportName As String
        Dim vName As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCheck As Integer
        Dim vDueDate As String
        Dim vCredit As Integer

        vCredit = 0
        vDocNo = Trim(TXT021.Text)
        On Error GoTo PrintError
         vQuery = "select groupdoc from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
         If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                 vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
         End If
         vRecordset.Close
         
         vRepType = "INV"
         vRepID = 582
                 
 
          If vCredit = 1 Then
            vQuery = "exec dbo.USP_AR_SearchDueDateInvoice '" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
              vDueDate = Trim(vRecordset.Fields("conditpaycode").Value)
            End If
            vRecordset.Close
            
          Else
            vDueDate = ""
          End If
        
          vCheck = 1
          vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
          If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            With Crystal103
                .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
                .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .Action = 1
            End With
          End If
          vRecordset.Close
                            
                            '---------------------------------------------------------------------------------------------------------
PrintError:
                If Err.Description <> "" Then
                MsgBox Err.Description
                End If


End Sub



Public Sub PrintInvoice_DriveThru()
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
        vRepType = "INV"
            If vGroupDoc = "IHV" Or vGroupDoc = "ihv" Then
                If Option1.Value = False Then
                    vRepID = 571
                Else
                    vRepID = 572
                End If
            ElseIf vGroupDoc = "IHN" Or vGroupDoc = "ihn" Then
                If Option1.Value = False Then
                    vRepID = 573
                Else
                    vRepID = 574
                End If
            ElseIf vGroupDoc = "ICV" Or vGroupDoc = "icv" Then
        
                 If Option1.Value = False Then
                     vRepID = 576
                 ElseIf Option1.Value = True Then
                     vRepID = 577
                 End If
            
            ElseIf vGroupDoc = "ICN" Or vGroupDoc = "icn" Then
            
                 If Option1.Value = False Then
                     vRepID = 578
                 ElseIf Option1.Value = True Then
                     vRepID = 579
                 End If
                 
            ElseIf vGroupDoc = "IVD" Or vGroupDoc = "ivd" Then
                MsgBox "ห้ามออกหัวเอกสารเก็บเงินปลายทาง", vbInformation + vbCritical, "ข้อความเตือน"
            ElseIf vGroupDoc = "IVM" Or vGroupDoc = "ivm" Then
                MsgBox "ห้ามออกหัวเอกสารเก็บเงินปลายทาง", vbInformation + vbCritical, "ข้อความเตือน"
            ElseIf vGroupDoc = "IVN" Or vGroupDoc = "ivn" Then
                MsgBox "ห้ามออกหัวเอกสารเก็บเงินปลายทาง", vbInformation + vbCritical, "ข้อความเตือน"
            ElseIf vGroupDoc = "IAB" Or vGroupDoc = "iab" Then
                If Option1.Value = False Then
                    vRepID = 54
                Else
                    vRepID = 154
                End If
            ElseIf vGroupDoc = "IVE" Or vGroupDoc = "ive" Then
                vRepID = 55
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
            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
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
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
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
        
                            vRepType = "RT"
                            If UCase(vGroupDoc) = "RHV" Or UCase(vGroupDoc) = "RXT" Or UCase(vGroupDoc) = "RAB" Or UCase(vGroupDoc) = "RDM" Then
                                vRepID = 64
                            ElseIf UCase(vGroupDoc) = "RHN" Or UCase(vGroupDoc) = "RDN" Then
                                vRepID = 65
                            ElseIf UCase(vGroupDoc) = "RCV" Or UCase(vGroupDoc) = "RVD" Or UCase(vGroupDoc) = "RVM" Or UCase(vGroupDoc) = "RDV" Then
                                vRepID = 135
                            ElseIf UCase(vGroupDoc) = "RCN" Or UCase(vGroupDoc) = "RVN" Then
                                vRepID = 136
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
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
        
                            vRepType = "DB"
                            If vGroupDoc = "DHV" Or vGroupDoc = "dhv" Then
                                vRepID = 62
                            ElseIf vGroupDoc = "DHN" Or vGroupDoc = "dhn" Then
                                vRepID = 63
                            ElseIf vGroupDoc = "DCV" Or vGroupDoc = "dcv" Then
                                vRepID = 476
                            ElseIf vGroupDoc = "DCN" Or vGroupDoc = "dcn" Then
                                vRepID = 477
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
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

        If Opt021.Value = False Then
            vChqStatus = 1
        Else
            vChqStatus = 0
        End If
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
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
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
                            
                            Opt021.Value = False
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
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
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


