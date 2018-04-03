VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form310 
   Caption         =   "หน้าพิมพ์ใบจ่ายสินค้า"
   ClientHeight    =   8340
   ClientLeft      =   3675
   ClientTop       =   1560
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form310.frx":0000
   ScaleHeight     =   8340
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CMBWHCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2475
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2160
      Width           =   1770
   End
   Begin Crystal.CrystalReport CrystalReport310 
      Left            =   1320
      Top             =   6255
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
   Begin VB.CommandButton CMD3101 
      Caption         =   "พิมพ์ทดแทน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2790
      TabIndex        =   3
      Top             =   3600
      Width           =   1440
   End
   Begin VB.ComboBox CMB3101 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2475
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2850
      Width           =   1770
   End
   Begin VB.TextBox TXT3101 
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
      Height          =   390
      Left            =   2475
      TabIndex        =   0
      Top             =   1425
      Width           =   2940
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "คลังที่พิมพ์ :"
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
      Height          =   240
      Left            =   720
      TabIndex        =   7
      Top             =   2160
      Width           =   1590
   End
   Begin VB.Label LBL3103 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชั้นเก็บที่พิมพ์ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   495
      TabIndex        =   6
      Top             =   2880
      Width           =   1830
   End
   Begin VB.Label LBL3102 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเหลือง :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   585
      TabIndex        =   5
      Top             =   1485
      Width           =   1725
   End
   Begin VB.Label LBL3101 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "หน้าพิมพ์ใบจ่ายสินค้า (ใบเหลือง)"
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
      Left            =   2475
      TabIndex        =   4
      Top             =   225
      Width           =   8490
   End
End
Attribute VB_Name = "Form310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMBWHCode_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vWHCode As String

vQuery = "exec dbo.USP_INV_SearchWHCodePrintSlip"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
Me.CMB3101.Clear
vRecordset.MoveFirst
While Not vRecordset.EOF
CMB3101.AddItem Trim(vRecordset.Fields("shelfcode").Value)
vRecordset.MoveNext
Wend
Else
Me.CMB3101.Clear
End If
vRecordset.Close
End Sub

Private Sub CMD3101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vShelfCode As String
Dim vDocNo As String
Dim vWHCode As String
Dim vIsCompleteSave As Integer


On Error GoTo ErrDescription

vShelfCode = Trim(CMB3101.Text)
vDocNo = Trim(TXT3101.Text)
vWHCode = Trim(Me.CMBWHCode.Text)

If vDocNo <> "" Then

vQuery = "select  docno,isnull(iscompletesave,0) as iscompletesave  from dbo.bcarinvoice where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  vIsCompleteSave = vRecordset.Fields("iscompletesave").Value
End If
vRecordset.Close

If vIsCompleteSave = 0 Then
  MsgBox "เอกสารยังบันทึกไม่สมบูรณ์ ยังไม่สามารถพิมพ์ได้ กรุณารอสักครู่", vbCritical, "Send Error Message"
  Me.CMD3101.SetFocus
  Exit Sub
End If

Call PrintBillPayItemZone(vDocNo, vWHCode, vShelfCode)

Else
   MsgBox "กรุณาใส่เลขที่เอกสารด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
   Me.TXT3101.SetFocus
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If

End Sub


Public Sub PrintBillPayItemZone(vDocNo As String, vWHCode As String, vShelfCode As String)
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer


On Error GoTo ErrDescription


vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close

vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 483
vRepType = "INV"
Else
vRepID = 485
vRepType = "INV"
End If
vCheck = 0

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
.ParameterFields(2) = "@vShelfCode;" & vShelfCode & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem_AVL(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        
'If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Then
 '  vDocGroup1 = Left(Right(TXT011.Text, Len(TXT011.Text) - InStr(TXT011.Text, "-")), 3) 'Left(Trim(TXT011.Text), 3)
'Else
 '  vDocGroup1 = Left(Trim(vDocNo), 3)
'End If

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
vRepID = 426
vRepType = "INV"
Else
vRepID = 424
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_PRO(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        
'If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Then
 '  vDocGroup1 = Left(Right(TXT011.Text, Len(TXT011.Text) - InStr(TXT011.Text, "-")), 3) 'Left(Trim(TXT011.Text), 3)
'Else
 '  vDocGroup1 = Left(Trim(vDocNo), 3)
'End If

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
vRepID = 473
vRepType = "INV"
Else
vRepID = 474
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem_BK1(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
vRepID = 430
vRepType = "INV"
Else
vRepID = 428
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem_BK2(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
vRepID = 434
vRepType = "INV"
Else
vRepID = 432
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_BK3(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If
vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
vRepID = 438
vRepType = "INV"
Else
vRepID = 436
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem_SPO(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
vRepID = 442
vRepType = "INV"
Else
vRepID = 440
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_DMG(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        
'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
vRepID = 458
vRepType = "INV"
Else
vRepID = 457
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_VND(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        
'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
vRepID = 462
vRepType = "INV"
Else
vRepID = 461
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_OFS(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
vRepID = 446
vRepType = "INV"
Else
vRepID = 445
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_SHW(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
vRepID = 450
vRepType = "INV"
Else
vRepID = 449
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_RSV(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocno1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocno1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
vRepID = 454
vRepType = "INV"
Else
vRepID = 453
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Private Sub TXT3101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String, vDocno1 As String
Dim i  As Integer
Dim vCount As Integer

On Error GoTo ErrDescription

CMB3101.Clear

If KeyAscii = 13 Then
      vDocNo = Trim(TXT3101.Text)
      vQuery = "exec dbo.USP_INV_SearchShelfPrintSlip'" & vDocNo & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         Me.CMB3101.Clear
         Me.CMBWHCode.Clear
         vRecordset.MoveFirst
         While Not vRecordset.EOF
         vDocno1 = Trim(vRecordset.Fields("docno").Value)
         CMB3101.AddItem Trim(vRecordset.Fields("shelfcode").Value)
         CMBWHCode.AddItem Trim(vRecordset.Fields("whcode").Value)
         vRecordset.MoveNext
         Wend
      Else
        MsgBox "ไม่มีเอกสารเลขที่  " & vDocNo & " ในระบบ", vbInformation + vbCritical, "ข้อความเตือน"
        Exit Sub
      End If
      vRecordset.Close
                    
      MsgBox "เลือกคลังที่จะพิมพ์ใบจ่ายสินค้า (ใบเหลือง)", vbInformation + vbCritical, "ข้อความเตือน"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub


Public Sub PrintItem010()
Dim vDocNo As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCheckItemNotLocation015 As Integer
Dim vCheckItemLocation015 As String

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT3101.Text))

vQuery = "exec dbo.USP_INV_CheckItemNotLocation015 '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  vCheckItemNotLocation015 = vRecordset.Fields("vcount").Value
End If
vRecordset.Close

vQuery = "exec dbo.USP_INV_CheckItemLocation015 '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  vCheckItemLocation015 = vRecordset.Fields("vcount").Value
End If
vRecordset.Close


If vCheckItemNotLocation015 > 0 Then
  Call PrintItem010_Zone010
End If
If vCheckItemLocation015 > 0 Then
  Call PrintItem010_Zone015
End If
        
        
        'vWHCode = Trim(CMB3101.Text)
        'vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocno & "'  and whcode = '" & vWHCode & "' "
        'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         '   vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        'End If
        'vRecordset.Close
        
        'vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocno & "'  and whcode = '" & vWHCode & "' "
        'gConnection.Execute vQuery
        
        'If vCheckNumber = "" Then
               '  vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
              '  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
             '       vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            '        vHeader = Trim(vRecordset.Fields("header").Value)
           '     End If
          '      vRecordset.Close
         '       vGenerateNumber = vHeader & "-" & vAutoNumber
        '
              '  vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
             '   gConnection.Execute vQuery
            '
           '     vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
          '      & " values ('" & vDocno & "','" & vGenerateNumber & "',getdate(),'010','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
         '       gConnection.Execute vQuery
        
        'End If
        
        'vDocno1 = UCase(Left(vDocno, 3))
        'vDocGroup1 = vDocno1
       ' If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
      '  vRepID = 126
     '   vRepType = "INV"
    '    Else
   '     vRepID = 93
  '      vRepType = "INV"
 '       End If
'vCheck = 1
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                   '         If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                  '              With CrystalReport310
                 '                   .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                '                    .ParameterFields(0) = "@DocNo;" & vDocno & ";true"
               '                     .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
              '                      .Destination = crptToWindow
             '                       .WindowState = crptMaximized
            '                        .Action = 1
           '                     End With
          '                  End If
         '                   vRecordset.Close
                            
        'vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocno & "' and whcode = '" & vWHCode & "' "
        'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         '   vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        'End If
        'vRecordset.Close
        
        'vCount = vCount + 1
        'vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocno & "' and whcode = '" & vWHCode & "'"
        'gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem010_Zone010()
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCount As Integer
Dim vWHCode As String
Dim vWHCode1 As Integer
Dim vCheckNumber As String, vDocGroup1 As String
Dim vAutoNumber As String, vDocno1, vHeader As String
Dim vGenerateNumber  As String
Dim vCheck As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT3101.Text))
vWHCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
   vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
      vHeader = Trim(vRecordset.Fields("header").Value)
  End If
  vRecordset.Close
  vGenerateNumber = vHeader & "-" & vAutoNumber

  vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
  gConnection.Execute vQuery
  
  vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
  & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'010','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
  gConnection.Execute vQuery
End If

vDocno1 = UCase(Left(vDocNo, 3))
vDocGroup1 = vDocno1
vRepType = "INV"
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 346
Else
vRepID = 342
End If
vCheck = 1
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                        With CrystalReport310
                            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                            .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                            .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                            .Destination = crptToWindow
                            .WindowState = crptMaximized
                            .Action = 1
                        End With
                    End If
                    vRecordset.Close
                    
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
gConnection.Execute vQuery
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem010_Zone015()
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCount As Integer
Dim vWHCode As String
Dim vWHCode1 As Integer
Dim vCheckNumber As String, vDocGroup1 As String
Dim vAutoNumber As String, vDocno1, vHeader As String
Dim vGenerateNumber  As String
Dim vCheck As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT3101.Text))
vWHCode = Trim("010C")
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
   vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
      vHeader = Trim(vRecordset.Fields("header").Value)
  End If
  vRecordset.Close
  vGenerateNumber = vHeader & "-" & vAutoNumber

  vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
  gConnection.Execute vQuery
  
  vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
  & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'010C','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
  gConnection.Execute vQuery
End If

vDocno1 = UCase(Left(vDocNo, 3))
vDocGroup1 = vDocno1
vRepType = "INV"
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 348
Else
vRepID = 344
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                        With CrystalReport310
                            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                            .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                            .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                            .Destination = crptToWindow
                            .WindowState = crptMaximized
                            .Action = 1
                        End With
                    End If
                    vRecordset.Close
                    
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
gConnection.Execute vQuery
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub


Public Sub PrintItem011()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocno1 As String
        Dim vAutoNumber As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'011','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
         vDocno1 = UCase(Left(vDocNo, 3))
         vDocGroup1 = vDocno1
        
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 128
        vRepType = "INV"
        Else
        vRepID = 94
        vRepType = "INV"
        End If

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem012()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocno1 As String
        Dim vAutoNumber As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'012','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocno1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocno1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 130
        vRepType = "INV"
        Else
        vRepID = 95
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem015()
   Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocno1 As String
        Dim vAutoNumber As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'015','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocno1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocno1

        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 132
        vRepType = "INV"
        Else
        vRepID = 96
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' ,PrintedtimeStamp = getdate() where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem097()
   Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocno1 As String
        Dim vAutoNumber As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'097','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocno1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocno1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 134
        vRepType = "INV"
        Else
        vRepID = 98
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem014()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocGroup1 As String
        Dim vAutoNumber As String, vDocno1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'014','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocno1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocno1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 172
        vRepType = "INV"
        Else
        vRepID = 171
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem020()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocGroup1 As String
        Dim vAutoNumber As String, vDocno1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'020','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocno1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocno1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 196
        vRepType = "INV"
        Else
        vRepID = 195
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem016()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocGroup1 As String
        Dim vAutoNumber As String, vDocno1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'016','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocno1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocno1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 219
        vRepType = "INV"
        Else
        vRepID = 218
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub


Public Sub PrintItem070()
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCount As Integer
Dim vWHCode As String
Dim vWHCode1 As Integer
Dim vCheckNumber As String, vDocGroup1 As String
Dim vAutoNumber As String, vDocno1 As String
Dim vGenerateNumber, vHeader As String
Dim vCheck As Integer

        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'070','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocno1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocno1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 299
        vRepType = "INV"
        Else
        vRepID = 297
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


