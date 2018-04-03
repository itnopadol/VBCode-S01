VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form411 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form411.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4350
      Top             =   3975
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   19
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "พิมพ์รายงาน"
      Height          =   540
      Left            =   3825
      TabIndex        =   1
      Top             =   2025
      Width           =   1440
   End
   Begin VB.TextBox TXT101 
      Height          =   390
      Left            =   2625
      TabIndex        =   0
      Top             =   1350
      Width           =   2640
   End
End
Attribute VB_Name = "Form411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
        Dim vReportName As String
        Dim vDocno As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
        Dim vCheckDate As Date
        Dim vGenerateNumber, vHeader As String
      
   
        
    vQuery = "select * from npmaster.dbo.NP_Generate_DocNo where headertype = 8"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vCheckDate = Trim(vRecordset.Fields("dateupdate").Value)
    End If
    vRecordset.Close
    
    If Year(vCheckDate) <> Year(Now) Or Month(vCheckDate) <> Month(Now) Or Day(vCheckDate) <> Day(Now) Then
    Call GenHeadDocument
    vQuery = "Update npmaster.dbo.NP_Generate_DocNo set header = '" & vGenDocNo & "', dateupdate = getdate() where headertype = 8 "
    gConnection.Execute vQuery
    End If

        vQuery = "select autonumber,header  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
        End If
        vRecordset.Close
        
        vGenerateNumber = vHeader & "-" & vAutoNumber
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 "
        gConnection.Execute vQuery
        
End Sub
