'Imports System
'Imports System.Net.DNS
'Imports System.Management
'Imports System.Security
'Imports System.Security.Principal.WindowsIdentity
'Imports System.Net
'Imports System.IO
'Imports System.Data
'Imports System.Data.SqlClient
'Imports System.Data.SqlTypes
'Imports System.Drawing
'Imports System.ComponentModel
'Imports System.Windows.Forms
'Imports vb6 = Microsoft.VisualBasic
'Imports Microsoft.Win32
'Imports System.Diagnostics
'Imports System.Collections.Generic
'Imports System.Text
'Imports System.Globalization


'Public Class FormNPTransfer

'    Dim vQuery As String
'    Dim vStrExecute As String
'    Dim hostname As String
'    Dim ipaddress As String
'    Dim h As System.Net.IPHostEntry = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName)

'    Dim mc As System.Management.ManagementClass
'    Dim mo As ManagementObject

'    Declare Function SendARP Lib "iphlpapi.dll" Alias "SendARP" (ByVal DestIP As Int32, ByVal SrcIP As Int32, ByVal pMacAddr() As Byte, ByRef PhyAddrLen As Int32) As Int32
'    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


'    Dim ds As DataSet
'    Dim da As SqlDataAdapter
'    Dim dt As DataTable

'    Dim cmd As SqlCommand
'    Dim vReadQuery As SqlDataReader
'    Dim vIsConnect As Integer

'    Dim i1 As Integer
'    Dim i2 As Integer

'    Dim ret As Long, mp3file As String

'    Dim ds1 As DataSet
'    Dim ds2 As DataSet
'    Dim ds3 As DataSet
'    Dim ds4 As DataSet
'    Dim ds5 As DataSet
'    Dim ds6 As DataSet
'    Dim ds7 As DataSet
'    Dim ds8 As DataSet
'    Dim ds9 As DataSet
'    Dim ds10 As DataSet

'    Dim da1 As SqlDataAdapter
'    Dim da2 As SqlDataAdapter
'    Dim da3 As SqlDataAdapter
'    Dim da4 As SqlDataAdapter
'    Dim da5 As SqlDataAdapter
'    Dim da6 As SqlDataAdapter
'    Dim da7 As SqlDataAdapter
'    Dim da8 As SqlDataAdapter
'    Dim da9 As SqlDataAdapter
'    Dim da10 As SqlDataAdapter

'    Dim dt1 As DataTable
'    Dim dt2 As DataTable
'    Dim dt3 As DataTable
'    Dim dt4 As DataTable
'    Dim dt5 As DataTable
'    Dim dt6 As DataTable
'    Dim dt7 As DataTable
'    Dim dt8 As DataTable
'    Dim dt9 As DataTable
'    Dim dt10 As DataTable


'    Dim thread1 As System.Threading.Thread
'    Dim thread2 As System.Threading.Thread
'    Dim thread3 As System.Threading.Thread
'    Dim thread4 As System.Threading.Thread
'    Dim thread5 As System.Threading.Thread

'    Dim thread6 As System.Threading.Thread
'    Dim thread7 As System.Threading.Thread
'    Dim thread8 As System.Threading.Thread
'    Dim thread9 As System.Threading.Thread
'    Dim thread10 As System.Threading.Thread


'    Dim vTransID As String
'    Dim vTransARInvoice As String
'    Dim vTransAPInvoice As String
'    Dim vTransItem As String
'    Dim vTransReceipt As String


'    Dim vMemIsSend As Integer

'    Dim vMemCountSend As Integer
'    Dim vIsSendComplete As Integer
'    Dim vBranchID As Integer
'    Dim vMemSending As Integer


'    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
'        On Error Resume Next

'        NotifyIcon1.Visible = False
'        Me.Visible = True
'        Me.WindowState = FormWindowState.Normal
'    End Sub

'    Private Sub BTNMinimize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNMinimize.Click
'        On Error Resume Next

'        Me.WindowState = 1
'        If (Me.WindowState = FormWindowState.Minimized) Then
'            Me.Visible = False
'            NotifyIcon1.Visible = True
'        End If
'    End Sub

'    Private Sub BTNCloseProgram_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNCloseProgram.Click
'        Dim vAnswer As Integer

'        On Error Resume Next

'        vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่ การออกโปรแกรมต้องไม่อยู่ระหว่างการโอนมิฉะนั้นออกไม่ได้", MsgBoxStyle.YesNo, "Send Question Message")

'        If vAnswer = 6 Then
'            If vIsprocess = 0 Then
'                Application.Exit()
'            End If
'        End If
'    End Sub

'    Private Sub FormTransfer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'        On Error Resume Next

'        Dim c As Process = Process.GetCurrentProcess()
'        Dim p As Process

'        For Each p In Process.GetProcessesByName(c.ProcessName)
'            If p.Id <> c.Id Then
'                If p.MainModule.FileName = c.MainModule.FileName Then
'                    Application.Exit()
'                End If
'            End If

'        Next p


'        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("th-TH")

'        Dim keyName As String = Registry.CurrentUser.ToString() & "\Control Panel\International"
'        Dim valueName As String = "sShortDate"
'        Dim s As String = Registry.GetValue(keyName, valueName, String.Empty).ToString()
'        Registry.SetValue(keyName, valueName, "dd/MM/yyyy")

'        Me.WindowState = 1
'        If (Me.WindowState = FormWindowState.Minimized) Then
'            Me.Visible = False
'            NotifyIcon1.Visible = True
'        End If


'        Me.CheckForIllegalCrossThreadCalls = False
'    End Sub

'    Private Sub TConnect_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TConnect.Tick
'        On Error Resume Next

'        thread1 = New System.Threading.Thread(AddressOf ConnectDatabase)
'        thread1.Start()

'        Me.TConnect.Enabled = False
'    End Sub

'    Public Sub ConnectDatabase()

'        On Error GoTo ErrMyDescription

'        vConnStrNEBULA = "Persist Security Info = False;User ID='sa';Password='[ibdkifu';Max Pool Size = 10000;Min Pool Size = 5;Data Source = 'nebula';Initial Catalog = 'bcnp';Pooling=false"
'        vConnNEBULA = New SqlConnection(vConnStrNEBULA)
'        vConnNEBULA.Open()

'        vIsConnect = 1

'        'Me.TGetDocNo.Enabled = True
'        'Me.TConnect.Enabled = False
'        vConnNEBULA.Close()


'ErrMyDescription:
'        If Err.Description = "" Then
'            vIsConnect = 1
'            'Me.TGetDocNo.Enabled = True
'            'Me.TConnect.Enabled = False
'        Else
'            vIsConnect = 0
'        End If
'    End Sub

'    Private Sub TGetDocNo_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TGetDocNo.Tick
'        'On Error Resume Next

'        If vIsprocess = 0 Then
'            thread2 = New System.Threading.Thread(AddressOf GetDocNotTransfer)
'            thread2.Start()
'        End If

'        'Me.TGetDocNo.Enabled = False
'    End Sub

'    Public Sub GetDocTransfer()
'        Dim n, i As Integer
'        Dim vListDoc = New ListViewItem
'        Dim vDepNo, vCreator, vCreateTime As String
'        Dim vOldDep, vNewDep As Double

'        'On Error Resume Next

'        If vIsConnect = 0 Then
'            Exit Sub
'        End If

'        vIsSendComplete = 0

'        Call vConnctDataBaseNP()

'        Me.ListViewDocNo.Items.Clear()

'        vQuery = "exec dbo.USP_DP_SearchDepNotTransfer 1"
'        da = New SqlDataAdapter(vQuery, vConnNEBULA)
'        ds1 = New DataSet
'        da.Fill(ds1, "DocNo")
'        dt1 = ds1.Tables("DocNo")
'        If dt1.Rows.Count > 0 Then
'            For i = 0 To dt1.Rows.Count - 1
'                n = n + 1
'                vDepNo = dt1.Rows(i).Item("depositno")
'                vCreator = dt1.Rows(i).Item("userchange")
'                vCreateTime = dt1.Rows(i).Item("trfdate")
'                vNewDep = dt1.Rows(i).Item("billbalance_new")
'                vOldDep = dt1.Rows(i).Item("billbalance_old")

'                vListDoc = Me.ListViewDocNo.Items.Add(n)
'                vListDoc.SubItems.Add(0).Text = vDepNo
'                vListDoc.SubItems.Add(1).Text = vCreateTime
'                vListDoc.SubItems.Add(2).Text = Format(vOldDep, "##,00.00#")
'                vListDoc.SubItems.Add(3).Text = Format(vNewDep, "##,00.00#")
'            Next
'        End If

'        vConnNEBULA.Close()

'    End Sub

'    Public Sub GetDocNotTransfer()
'        Dim i As Integer
'        Dim vListDoc = New ListViewItem
'        Dim vDepNo As String


'        'On Error Resume Next

'        If vIsConnect = 0 Then
'            Exit Sub
'        End If

'        vIsSendComplete = 0

'        Call vConnctDataBaseNP()

'        Me.ListViewDocNo.Items.Clear()

'        vIsprocess = 1

'        vQuery = "exec dbo.USP_DP_SearchDepNotTransfer 0"
'        da = New SqlDataAdapter(vQuery, vConnNEBULA)
'        ds2 = New DataSet
'        da.Fill(ds1, "DocNo")
'        dt2 = ds1.Tables("DocNo")
'        If dt2.Rows.Count > 0 Then
'            For i = 0 To dt2.Rows.Count - 1

'                vDepNo = dt2.Rows(i).Item("depositno")

'                vStrExecute = "exec dbo.USP_TF_TriggerDeposit '" & vDepNo & "'"
'                cmd = New SqlCommand(vStrExecute, vConnNEBULA)
'                cmd.ExecuteNonQuery()

'            Next
'        End If

'        vConnNEBULA.Close()

'        vIsprocess = 0

'    End Sub

'    Private Sub TGetTransfer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TGetTransfer.Tick
'        On Error Resume Next

'        thread3 = New System.Threading.Thread(AddressOf GetDocTransfer)
'        thread3.Start()

'        GetDocTransfer()

'        'Me.TGetTransfer.Enabled = False
'    End Sub
'End Class
